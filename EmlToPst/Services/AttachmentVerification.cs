using System.Buffers;
using System.IO;
using System.Security.Cryptography;
using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace EmlToPst.Services;

internal sealed record AttachmentPayload(long Bytes, string Sha256);

internal sealed record ExpectedAttachments(string SourceFile, AttachmentPayload[] Payloads);

internal static class AttachmentVerification
{
    private const string Stage = "PST_VERIFY";

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

    public static int Verify(string pstPath, Dictionary<string, ExpectedAttachments> pending, ImportLog log)
    {
        int checkedMessages = 0;
        int mismatches = 0;
        int expectedMessages = pending.Count;
        long attachmentBudget = Math.Max(4L * 1024 * 1024 * 1024, new FileInfo(pstPath).Length);
        var readerOptions = new EmailStoreReaderOptions(maxTotalAttachmentBytes: attachmentBudget);
        var itemOptions = new EmailStoreItemReadOptions(EmailStoreItemReadParts.AttachmentContent,
            preferStreamingAttachmentContent: true);

        using var store = EmailStoreSession.Open(pstPath, readerOptions);
        foreach (EmailStoreItemReference reference in store.EnumerateItems())
        {
            if (!pending.Remove(reference.Id, out ExpectedAttachments? expected))
            {
                continue;
            }

            checkedMessages++;
            try
            {
                AttachmentPayload[] actual = Capture(store.ReadItem(reference, itemOptions).Document);
                bool same = expected.Payloads.SequenceEqual(actual);
                string message = $"file={expected.SourceFile} | item={reference.Id}"
                    + $" | expected={Describe(expected.Payloads)} | actual={Describe(actual)}";
                if (same)
                {
                    log.Info(Stage, $"matched | {message}");
                }
                else
                {
                    mismatches++;
                    log.Warn(Stage, $"attachment count, size or SHA-256 mismatch | {message}");
                }
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                mismatches++;
                log.Error(Stage, $"attachment read failed | file={expected.SourceFile} | item={reference.Id}", ex);
            }

            if (pending.Count == 0)
            {
                break;
            }
        }

        foreach ((string id, ExpectedAttachments missing) in pending)
        {
            log.Warn(Stage, $"item missing | file={missing.SourceFile} | item={id} | expected={Describe(missing.Payloads)}");
        }

        mismatches += pending.Count;
        log.Info(Stage, $"completed | expectedMessages={expectedMessages} | checkedMessages={checkedMessages} | failedMessages={mismatches}");
        return mismatches;
    }

    public static string Describe(AttachmentPayload[] payloads)
    {
        long largest = payloads.Length == 0 ? 0 : payloads[^1].Bytes;
        return $"{payloads.Length} attachments, {payloads.Sum(payload => payload.Bytes)} bytes, largest={largest} bytes";
    }
}
