using System.Buffers;
using System.IO;
using System.Security.Cryptography;
using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace EmlToPst.Services;

internal sealed record ExpectedMessage(string SourceFile, AttachmentPayload[] Payloads, DateTimeOffset? ReceivedDate);

internal sealed record PstVerificationResult(int AttachmentFailures, int DateFailures);

internal static class PstVerification
{
    private const string AttachmentStage = "PST_VERIFY";
    private const string DateStage = "PST_DATE_VERIFY";

    public static AttachmentPayload[] Capture(EmailDocument document, CancellationToken cancellationToken = default)
    {
        if (document.Attachments.Count == 0)
        {
            return [];
        }

        var payloads = new List<AttachmentPayload>();
        byte[] buffer = ArrayPool<byte>.Shared.Rent(64 * 1024);
        try
        {
            foreach (EmailAttachment attachment in document.Attachments)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (attachment.EmbeddedDocument is not null || attachment.LinkedPath is not null
                    || attachment.MapiAttachMethod is not (null or 1) || attachment.StructuredStorageStreams.Count > 0
                    || attachment.ContentType is "text/calendar" or "application/ics" or "text/vcard" or "text/x-vcard" or "text/directory")
                {
                    continue;
                }

                if (attachment.Length > 0 && attachment.Content is null && attachment.ContentSource is null)
                {
                    throw new InvalidDataException($"첨부파일의 데이터가 없습니다. (크기 {attachment.Length:N0}바이트)");
                }

                using Stream content = attachment.OpenContentStream();
                using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
                long length = 0;
                int read;
                while ((read = content.Read(buffer, 0, buffer.Length)) != 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    hash.AppendData(buffer, 0, read);
                    length += read;
                }

                payloads.Add(new AttachmentPayload(length, Convert.ToHexString(hash.GetHashAndReset())));
            }
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
        }

        // Compare payloads independently of the attachment order chosen by the PST reader.
        return payloads.OrderBy(payload => payload.Bytes)
            .ThenBy(payload => payload.Sha256, StringComparer.Ordinal).ToArray();
    }

    public static PstVerificationResult Verify(string pstPath, Dictionary<string, ExpectedMessage> pending, ImportLog log)
    {
        int expectedAttachments = pending.Values.Count(message => message.Payloads.Length > 0);
        int expectedDates = pending.Values.Count(message => message.ReceivedDate.HasValue);
        int checkedAttachments = 0;
        int checkedDates = 0;
        int attachmentFailures = 0;
        int dateFailures = 0;
        long attachmentBudget = Math.Max(4L * 1024 * 1024 * 1024, new FileInfo(pstPath).Length);
        var readerOptions = new EmailStoreReaderOptions(maxTotalAttachmentBytes: attachmentBudget);
        var itemOptions = new EmailStoreItemReadOptions(EmailStoreItemReadParts.AttachmentContent,
            preferStreamingAttachmentContent: true);

        using var store = EmailStoreSession.Open(pstPath, readerOptions);
        foreach (EmailStoreItemReference reference in store.EnumerateItems())
        {
            if (!pending.Remove(reference.Id, out ExpectedMessage? expected))
            {
                continue;
            }

            if (expected.ReceivedDate is DateTimeOffset expectedDate)
            {
                checkedDates++;
                if (!VerifyDate(store, reference, expected.SourceFile, expectedDate, log))
                {
                    dateFailures++;
                }
            }

            if (expected.Payloads.Length > 0)
            {
                checkedAttachments++;
                if (!VerifyAttachments(store, reference, itemOptions, expected, log))
                {
                    attachmentFailures++;
                }
            }

            if (pending.Count == 0)
            {
                break;
            }
        }

        foreach ((string id, ExpectedMessage missing) in pending)
        {
            if (missing.Payloads.Length > 0)
            {
                attachmentFailures++;
                log.Warn(AttachmentStage, $"item missing | file={missing.SourceFile} | item={id} | expected={Describe(missing.Payloads)}");
            }

            if (missing.ReceivedDate.HasValue)
            {
                dateFailures++;
                log.Warn(DateStage, $"item missing | file={missing.SourceFile} | item={id} | expected={missing.ReceivedDate:O}");
            }
        }

        log.Info(AttachmentStage, $"completed | expectedMessages={expectedAttachments} | checkedMessages={checkedAttachments} | failedMessages={attachmentFailures}");
        log.Info(DateStage, $"completed | expectedMessages={expectedDates} | checkedMessages={checkedDates} | failedMessages={dateFailures}");
        return new PstVerificationResult(attachmentFailures, dateFailures);
    }

    private static bool VerifyDate(EmailStoreSession store, EmailStoreItemReference reference, string file,
        DateTimeOffset expected, ImportLog log)
    {
        try
        {
            DateTimeOffset? actual = store.ReadSummary(reference).ReceivedAt;
            string message = $"file={file} | item={reference.Id} | expected={expected:O} | actual={actual?.ToString("O") ?? "(none)"}";
            if (actual == expected)
            {
                log.Info(DateStage, $"matched | {message}");
                return true;
            }

            log.Warn(DateStage, $"received date mismatch | {message}");
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            log.Error(DateStage, $"received date read failed | file={file} | item={reference.Id}", ex);
        }

        return false;
    }

    private static bool VerifyAttachments(EmailStoreSession store, EmailStoreItemReference reference,
        EmailStoreItemReadOptions options, ExpectedMessage expected, ImportLog log)
    {
        try
        {
            AttachmentPayload[] actual = Capture(store.ReadItem(reference, options).Document);
            string message = $"file={expected.SourceFile} | item={reference.Id}"
                + $" | expected={Describe(expected.Payloads)} | actual={Describe(actual)}";
            if (expected.Payloads.SequenceEqual(actual))
            {
                log.Info(AttachmentStage, $"matched | {message}");
                return true;
            }

            log.Warn(AttachmentStage, $"attachment count, size or SHA-256 mismatch | {message}");
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            log.Error(AttachmentStage, $"attachment read failed | file={expected.SourceFile} | item={reference.Id}", ex);
        }

        return false;
    }

    public static string Describe(AttachmentPayload[] payloads)
    {
        long largest = payloads.Length == 0 ? 0 : payloads[^1].Bytes;
        return $"{payloads.Length} attachments, {payloads.Sum(payload => payload.Bytes)} bytes, largest={largest} bytes";
    }
}
