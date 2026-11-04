using System.Diagnostics;
using System.IO;
using System.Windows.Threading;
using EmlToPst.Models;
using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace EmlToPst.Services;

internal sealed class PstImportService
{
    public async Task<ImportResult> ImportAsync(ImportOptions options, IProgress<ImportProgress> progress,
        ImportLog log, CancellationToken cancellationToken)
    {
        var elapsed = Stopwatch.StartNew();
        log.Info("START", $"source={options.SourceFolder} | output={options.OutputPath}"
            + $" | recursive={options.IncludeSubfolders} | attachToOutlook={options.AttachToOutlook}"
            + " | pstEngine=OfficeIMO.Email 3.5.3");

        try
        {
            ImportResult result = await Task.Run(
                () => WritePst(options, progress, log, cancellationToken), cancellationToken);

            if (options.AttachToOutlook && result.PstPath is string path && !result.Cancelled
                && !result.AttachmentVerificationFailed && !result.ReceivedDateVerificationFailed)
            {
                progress.Report(new ImportProgress(result.Imported + result.Failures.Count, result.Total,
                    result.Imported, result.Failures.Count, "완성된 PST를 Outlook에 연결 중"));
                await Dispatcher.Yield(DispatcherPriority.Background);
                if (!cancellationToken.IsCancellationRequested)
                {
                    result = AttachToOutlook(result, path, log);
                }
                else
                {
                    log.Info("OUTLOOK_ATTACH", "skipped: cancelled");
                }
            }
            else
            {
                string reason = (options.AttachToOutlook, result.AttachmentVerificationFailed, result.ReceivedDateVerificationFailed) switch
                {
                    (false, _, _) => "option disabled",
                    (_, true, _) => "attachment verification failed",
                    (_, _, true) => "received date verification failed",
                    _ => "no eligible saved PST"
                };
                log.Info("OUTLOOK_ATTACH", $"skipped: {reason}");
            }

            result = result with { Cancelled = result.Cancelled || cancellationToken.IsCancellationRequested };
            int saved = result.PstPath is null ? 0 : result.Imported;
            log.Info("END", $"total={result.Total} | written={result.Imported} | saved={saved}"
                + $" | failed={result.Failures.Count} | unprocessed={result.Total - result.Imported - result.Failures.Count}"
                + $" | cancelled={result.Cancelled} | attached={result.AttachedToOutlook}"
                + $" | attachmentVerificationFailed={result.AttachmentVerificationFailed}"
                + $" | receivedDateVerificationFailed={result.ReceivedDateVerificationFailed}"
                + $" | elapsedMs={elapsed.ElapsedMilliseconds} | pst={result.PstPath ?? "(none)"}");
            return result;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            log.Info("END", $"cancelled=true | saved=0 | elapsedMs={elapsed.ElapsedMilliseconds}");
            throw;
        }
        catch (Exception ex)
        {
            log.Error("END", $"saved=0 | elapsedMs={elapsed.ElapsedMilliseconds}", ex);
            throw;
        }
    }

    private static ImportResult WritePst(ImportOptions options, IProgress<ImportProgress> progress,
        ImportLog log, CancellationToken cancellationToken)
    {
        string stage = "VALIDATE_PATHS";
        string source;
        string output;
        string[] files;
        try
        {
            source = Path.GetFullPath(options.SourceFolder);
            output = Path.GetFullPath(options.OutputPath);
            ValidatePaths(source, output);

            stage = "SCAN";
            log.Info(stage, "started");
            files = FindEmlFiles(source, options.IncludeSubfolders, cancellationToken);
            log.Info(stage, $"completed | files={files.Length}");
            cancellationToken.ThrowIfCancellationRequested();
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            log.Error(stage, "failed", ex);
            throw;
        }

        if (files.Length == 0)
        {
            return new ImportResult(0, 0, [], false, null, false, []);
        }

        var reader = new EmailDocumentReader();
        var failures = new List<ImportFailure>();
        var warnings = new List<string>();
        var expectedMessages = new Dictionary<string, ExpectedMessage>(StringComparer.OrdinalIgnoreCase);
        var uiUpdate = Stopwatch.StartNew();
        int imported = 0;
        int receivedDates = 0;
        bool cancelled = false;
        bool attachmentVerificationFailed = false;
        bool receivedDateVerificationFailed = false;
        string? savedPath = null;
        string? currentFile = null;
        stage = "PST_CREATE";

        try
        {
            progress.Report(new ImportProgress(0, files.Length, 0, 0, "PST 준비 중"));
            log.Info(stage, $"started | output={output}");
            using var writer = EmailStorePstWriter.Create(output,
                new EmailStorePstWriterOptions(displayName: Path.GetFileNameWithoutExtension(output)));
            string root = writer.AddFolder("EML 가져오기", containerClass: "IPF.Note");
            var folderIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { [""] = root };

            try
            {
                for (int index = 0; index < files.Length; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    string eml = files[index];
                    currentFile = Path.GetRelativePath(source, eml);
                    if (index == 0 || uiUpdate.ElapsedMilliseconds >= 100)
                    {
                        progress.Report(new ImportProgress(index, files.Length, imported, failures.Count, currentFile));
                        uiUpdate.Restart();
                    }

                    stage = "EML_READ";
                    log.Info(stage, $"{index + 1}/{files.Length} | file={currentFile}");
                    try
                    {
                        using (EmailReadResult message = ReadEml(reader, eml, cancellationToken))
                        {
                            foreach (var diagnostic in message.Diagnostics)
                            {
                                log.Warn(stage, $"file={currentFile} | {diagnostic.Severity} | {diagnostic.Code}");
                            }

                            if (message.HasErrors)
                            {
                                string codes = string.Join(", ", message.Diagnostics.Select(item => item.Code).Distinct());
                                throw new InvalidDataException($"EML을 읽는 중 오류가 발생했습니다. ({codes})");
                            }

                            string receivedSource = EmlTimestampMapper.Apply(message.Document);
                            DateTimeOffset? receivedDate = message.Document.ReceivedDate;
                            log.Info("EML_DATES", $"file={currentFile} | receivedSource={receivedSource}"
                                + $" | received={receivedDate?.ToString("O") ?? "(none)"}"
                                + $" | sent={message.Document.Date?.ToString("O") ?? "(none)"}");
                            AttachmentPayload[] payloads = PstVerification.Capture(message.Document, cancellationToken);
                            log.Info("EML_ATTACHMENTS", $"file={currentFile} | emlBytes={message.BytesRead}"
                                + $" | {PstVerification.Describe(payloads)}");
                            cancellationToken.ThrowIfCancellationRequested();
                            stage = "PST_ADD_ITEM";
                            string relativeFolder = Path.GetDirectoryName(currentFile) ?? "";
                            string folderId = EnsureFolder(writer, folderIds, relativeFolder);
                            // Once a write starts, finish it before observing cancellation.
                            string itemId = writer.AddItem(folderId, message.Document);
                            if (receivedDate.HasValue)
                            {
                                receivedDates++;
                            }

                            if (payloads.Length > 0 || receivedDate.HasValue)
                            {
                                expectedMessages.Add(itemId, new ExpectedMessage(currentFile, payloads, receivedDate));
                            }
                        }

                        imported++;
                        log.Info(stage, $"written={imported} | failed={failures.Count} | file={currentFile}");
                    }
                    catch (Exception ex) when (stage == "EML_READ" && ex is not (OperationCanceledException or OutOfMemoryException))
                    {
                        failures.Add(new ImportFailure(currentFile, DescribeError(ex)));
                        log.Error(stage, $"skipped | file={currentFile}", ex);
                    }

                    currentFile = null;
                }
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                cancelled = true;
                log.Info("CANCEL", $"written={imported} | failed={failures.Count} | file={currentFile ?? "(none)"}");
            }

            if (imported > 0)
            {
                currentFile = null;
                stage = "PST_FINALIZE";
                progress.Report(new ImportProgress(imported + failures.Count, files.Length,
                    imported, failures.Count, "PST 파일 저장 중"));
                log.Info(stage, $"started | items={imported}");
                EmailStorePstWriteReport report = writer.Complete();
                savedPath = output;
                log.Info(stage, $"committed | bytes={report.BytesWritten} | folders={report.FolderCount}"
                    + $" | items={report.ItemCount} | output={output}");
                foreach (var diagnostic in report.Diagnostics)
                {
                    log.Warn(stage, $"{diagnostic.Severity} | {diagnostic.Code}");
                }

                if (report.HasErrors || report.HasDataLoss || report.DiagnosticsTruncated)
                {
                    warnings.Add("일부 메시지 속성 또는 첨부 형식에 대한 변환 진단이 있습니다. 로그를 확인해 주세요.");
                }

                if (imported > receivedDates)
                {
                    warnings.Add($"메일 {imported - receivedDates:N0}건에서 원본 수신 일자를 확인할 수 없었습니다. 로그의 EML_DATES를 확인해 주세요.");
                }

                if (expectedMessages.Count > 0)
                {
                    progress.Report(new ImportProgress(imported + failures.Count, files.Length,
                        imported, failures.Count, "PST 첨부파일 및 수신 일자 검증 중"));
                    int attachmentMessages = expectedMessages.Values.Count(message => message.Payloads.Length > 0);
                    log.Info("PST_VERIFY", $"started | messages={attachmentMessages}");
                    log.Info("PST_DATE_VERIFY", $"started | messages={receivedDates} | sourceDateMissing={imported - receivedDates}");
                    try
                    {
                        PstVerificationResult verification = PstVerification.Verify(output, expectedMessages, log);
                        if (verification.AttachmentFailures > 0)
                        {
                            attachmentVerificationFailed = true;
                            warnings.Add($"메일 {verification.AttachmentFailures:N0}건의 첨부파일 검증에 실패했습니다. 로그의 PST_VERIFY를 확인해 주세요.");
                        }

                        if (verification.DateFailures > 0)
                        {
                            receivedDateVerificationFailed = true;
                            warnings.Add($"메일 {verification.DateFailures:N0}건의 수신 일자 검증에 실패했습니다. 로그의 PST_DATE_VERIFY를 확인해 주세요.");
                        }
                    }
                    catch (Exception ex)
                    {
                        attachmentVerificationFailed = attachmentMessages > 0;
                        receivedDateVerificationFailed = receivedDates > 0;
                        warnings.Add($"저장된 PST를 다시 읽어 검증하지 못했습니다. {DescribeError(ex)}");
                        log.Error("PST_VERIFY", "verification stopped", ex);
                    }
                }
                else
                {
                    log.Info("PST_VERIFY", "skipped: no file attachment payloads");
                    log.Info("PST_DATE_VERIFY", $"skipped: no source received dates | sourceDateMissing={imported}");
                }
            }
            else
            {
                log.Info("PST_FINALIZE", "skipped: no messages written");
            }
        }
        catch (Exception ex)
        {
            if (stage == "PST_ADD_ITEM" && currentFile is not null)
            {
                failures.Add(new ImportFailure(currentFile, DescribeError(ex)));
            }

            warnings.Add($"변환이 중단되었습니다. {DescribeError(ex)}");
            log.Error(stage, $"failed | file={currentFile ?? "(none)"} | written={imported}", ex);
        }

        return new ImportResult(files.Length, imported, failures, cancelled, savedPath, false, warnings,
            attachmentVerificationFailed, receivedDateVerificationFailed);
    }

    private static EmailReadResult ReadEml(EmailDocumentReader reader, string path, CancellationToken cancellationToken)
    {
        using var input = File.OpenRead(path);
        if (input.Length == 0)
        {
            throw new InvalidDataException("비어 있는 EML 파일입니다.");
        }

        return reader.ReadStreaming(input, path, cancellationToken);
    }

    private static string EnsureFolder(EmailStorePstWriter writer, Dictionary<string, string> folderIds, string relativeFolder)
    {
        if (folderIds.TryGetValue(relativeFolder, out string? folderId))
        {
            return folderId;
        }

        string current = "";
        string parentId = folderIds[current];
        foreach (string name in relativeFolder.Split(Path.DirectorySeparatorChar))
        {
            current = Path.Combine(current, name);
            if (!folderIds.TryGetValue(current, out string? childId))
            {
                childId = writer.AddFolder(name, parentId, "IPF.Note");
                folderIds.Add(current, childId);
            }

            parentId = childId;
        }

        return parentId;
    }

    private static ImportResult AttachToOutlook(ImportResult result, string path, ImportLog log)
    {
        log.Info("OUTLOOK_ATTACH", $"started | pst={path}");
        try
        {
            OutlookPstAttacher.Attach(path);
            log.Info("OUTLOOK_ATTACH", "completed");
            return result with { AttachedToOutlook = true };
        }
        catch (Exception ex)
        {
            log.Error("OUTLOOK_ATTACH", "failed after PST was saved", ex);
            return result with
            {
                Warnings = [.. result.Warnings, $"PST는 저장되었습니다. Outlook 연결에 실패했습니다. {DescribeError(ex)}"]
            };
        }
    }

    internal static string DescribeError(Exception exception) =>
        $"{exception.Message} [0x{exception.HResult:X8}]";

    private static void ValidatePaths(string source, string output)
    {
        if (!Directory.Exists(source))
        {
            throw new DirectoryNotFoundException("원본 폴더를 찾을 수 없습니다.");
        }

        if (!Path.GetExtension(output).Equals(".pst", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("결과 파일의 확장자는 .pst여야 합니다.");
        }

        if (!Directory.Exists(Path.GetDirectoryName(output)))
        {
            throw new DirectoryNotFoundException("PST를 저장할 폴더를 찾을 수 없습니다.");
        }

        if (File.Exists(output) || Directory.Exists(output))
        {
            throw new IOException("같은 이름의 파일 또는 폴더가 있습니다. 새 PST 파일명을 선택해 주세요.");
        }
    }

    private static string[] FindEmlFiles(string source, bool recursive, CancellationToken cancellationToken)
    {
        var options = new EnumerationOptions
        {
            RecurseSubdirectories = recursive,
            IgnoreInaccessible = false,
            AttributesToSkip = FileAttributes.ReparsePoint
        };
        var files = new List<string>();
        foreach (string file in Directory.EnumerateFiles(source, "*", options))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (Path.GetExtension(file).Equals(".eml", StringComparison.OrdinalIgnoreCase))
            {
                files.Add(file);
            }
        }

        return files.OrderBy(file => Path.GetDirectoryName(file), StringComparer.OrdinalIgnoreCase)
            .ThenBy(file => file, StringComparer.OrdinalIgnoreCase).ToArray();
    }
}
