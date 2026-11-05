using System.Diagnostics;
using System.IO;
using EmlToPst.Models;
using MailKit;
using MailKit.Net.Imap;

namespace EmlToPst.Services;

internal sealed class ImapBackupService
{
    private readonly Func<IImapBackupSource> factory;
    private readonly Func<TimeSpan, CancellationToken, Task>? retryDelay;

    public ImapBackupService(Func<IImapBackupSource>? factory = null,
        Func<TimeSpan, CancellationToken, Task>? retryDelay = null)
    {
        this.factory = factory ?? (() => new MailKitBackupSource());
        this.retryDelay = retryDelay;
    }

    public async Task<IReadOnlyList<BackupMailbox>> ListMailboxesAsync(ImapConnectionOptions options,
        string password, ImportLog log, CancellationToken token)
    {
        using var protection = log.ProtectSecret(password);
        await using var runner = new ImapConnectionRunner(options.Normalize(), password, log, factory, retryDelay);
        return await runner.RunAsync(null, null, (source, _) => source.ListMailboxesAsync(token), token).ConfigureAwait(false);
    }

    public Task<BackupResult> BackupAsync(ImapBackupOptions options, string password,
        IProgress<BackupProgress> progress, ImportLog log, CancellationToken token) =>
        Task.Run(() => BackupCoreAsync(options, password, progress, log, token), token);

    private async Task<BackupResult> BackupCoreAsync(ImapBackupOptions options, string password,
        IProgress<BackupProgress> progress, ImportLog log, CancellationToken token)
    {
        using var protection = log.ProtectSecret(password);
        ImapConnectionOptions connection = options.Connection.Normalize();
        if (options.Mailboxes.Count == 0) throw new ArgumentException("백업할 메일함을 선택해 주세요.");
        var layout = new BackupLayout(options.DestinationRoot, connection);
        Directory.CreateDirectory(layout.MessageRoot);
        // Also protects checkpoints from a second process using the same destination/account.
        using var folderLock = AcquireLock(layout.AccountRoot);
        using var state = new BackupStateStore(layout.DatabasePath);
        state.ValidateAccount(connection);
        await using var runner = new ImapConnectionRunner(connection, password, log, factory, retryDelay);
        var clock = Stopwatch.StartNew();
        var failures = new List<BackupFailure>();
        var plans = new List<(BackupMailbox Mailbox, MailboxSnapshot Snapshot)>();
        int total = 0, downloaded = 0, skipped = 0, failed = 0, missing = 0;
        bool cancelled = false;
        string current = "메일 목록 조회 중";
        void Report(long bytes = 0, long? expected = null) => progress.Report(new BackupProgress(
            total, downloaded, skipped, failed, missing, current, bytes, expected));

        log.Info("BACKUP_START", $"host={connection.Host} | port={connection.Port} | account={connection.Account}"
            + $" | selectedFolders={options.Mailboxes.Count} | root={layout.AccountRoot}");
        try
        {
            foreach (BackupMailbox mailbox in options.Mailboxes.DistinctBy(m => m.FullName, StringComparer.Ordinal))
            {
                token.ThrowIfCancellationRequested();
                current = $"메일 목록 조회: {mailbox.FullName}";
                Report();
                try
                {
                    MailboxSnapshot snapshot = await runner.RunAsync(mailbox.FullName, null, async (source, validity) =>
                        new MailboxSnapshot(validity, await source.GetUidsAsync(token).ConfigureAwait(false)), token).ConfigureAwait(false);
                    uint? previous = state.RecordFolder(mailbox, snapshot.UidValidity);
                    if (previous.HasValue && previous.Value != snapshot.UidValidity)
                        log.Warn("IMAP_UIDVALIDITY", $"mailbox={mailbox.FullName} | old={previous} | new={snapshot.UidValidity} | oldFilesPreserved=true");
                    plans.Add((mailbox, snapshot));
                    total = checked(total + snapshot.Uids.Count);
                    log.Info("IMAP_SNAPSHOT", $"mailbox={mailbox.FullName} | validity={snapshot.UidValidity} | messages={snapshot.Uids.Count}");
                }
                catch (Exception ex) when (IsFolderError(ex))
                {
                    failures.Add(new BackupFailure(mailbox.FullName, "", log.Redact(ex.Message)));
                    log.Error("IMAP_FOLDER", $"snapshot failed | mailbox={mailbox.FullName}", ex);
                }
            }

            foreach (var plan in plans)
            {
                BackupMailbox mailbox = plan.Mailbox;
                uint validity = plan.Snapshot.UidValidity;
                // Fix the UID set at the start. Newly arriving messages are picked up by the next run.
                foreach (uint[] batch in plan.Snapshot.Uids.Chunk(100))
                {
                    token.ThrowIfCancellationRequested();
                    IReadOnlyList<RemoteMessage> summaries;
                    current = $"메일 정보 조회: {mailbox.FullName}";
                    Report();
                    try
                    {
                        summaries = await runner.RunAsync(mailbox.FullName, validity,
                            (source, _) => source.FetchMetadataAsync(batch, token), token).ConfigureAwait(false);
                    }
                    catch (Exception ex) when (IsFolderError(ex))
                    {
                        failures.Add(new BackupFailure(mailbox.FullName, "", log.Redact(ex.Message)));
                        log.Error("IMAP_FOLDER", $"metadata failed; remaining messages unprocessed | mailbox={mailbox.FullName}", ex);
                        break;
                    }

                    var metadata = summaries.ToDictionary(m => m.Uid);
                    bool folderChanged = false;
                    foreach (uint uid in batch)
                    {
                        token.ThrowIfCancellationRequested();
                        current = $"{mailbox.FullName} / UID {uid}";
                        Report();
                        if (!metadata.TryGetValue(uid, out RemoteMessage? message))
                        {
                            missing++;
                            failures.Add(new BackupFailure(mailbox.FullName, uid.ToString(), "목록 조회 후 서버에서 삭제되거나 이동한 메일입니다."));
                            log.Warn("IMAP_MISSING", $"mailbox={mailbox.FullName} | uid={uid} | stage=metadata");
                            Report();
                            continue;
                        }
                        try
                        {
                            string path = layout.MessagePath(mailbox, validity, uid);
                            StoredBackupMessage? stored = state.Find(mailbox.FullName, validity, uid);
                            if (stored is not null && (!message.Size.HasValue || message.Size == stored.StoredBytes)
                                && await RecoverOrSkipAsync(stored, path, state, log, token).ConfigureAwait(false))
                            {
                                state.RefreshMetadata(mailbox, validity, message);
                                skipped++;
                                log.Info("IMAP_SKIP", $"mailbox={mailbox.FullName} | uid={uid} | sha256Verified=true");
                            }
                            else
                            {
                                await DownloadOneAsync(runner, mailbox, validity, message, path, layout, state,
                                    (bytes, expected) => Report(bytes, expected), log, token).ConfigureAwait(false);
                                downloaded++;
                            }
                        }
                        catch (RemoteMessageMissingException ex)
                        {
                            missing++;
                            state.Fail(mailbox.FullName, validity, uid, ex.Message);
                            failures.Add(new BackupFailure(mailbox.FullName, uid.ToString(), ex.Message));
                            log.Warn("IMAP_MISSING", $"mailbox={mailbox.FullName} | uid={uid} | stage=download");
                        }
                        catch (UidValidityChangedException ex)
                        {
                            failures.Add(new BackupFailure(mailbox.FullName, uid.ToString(), ex.Message));
                            log.Error("IMAP_UIDVALIDITY", "stopped this folder; remaining messages unprocessed", ex);
                            folderChanged = true;
                            break;
                        }
                        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ImapCommandException)
                        {
                            failed++;
                            string reason = log.Redact(ex.Message);
                            state.Fail(mailbox.FullName, validity, uid, reason);
                            failures.Add(new BackupFailure(mailbox.FullName, uid.ToString(), reason));
                            log.Error("IMAP_MESSAGE_FAIL", $"mailbox={mailbox.FullName} | uid={uid}", ex);
                        }
                        Report();
                    }
                    if (folderChanged) break;
                }
            }
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
            cancelled = true;
            log.Info("BACKUP_CANCEL", "cancelled; completed messages and staged checkpoints preserved");
        }
        catch (Exception ex)
        {
            log.Error("BACKUP_ABORT", $"total={total} | downloaded={downloaded} | skipped={skipped} | failed={failed}"
                + $" | missing={missing} | unprocessed={total - downloaded - skipped - failed - missing} | current={current}", ex);
            throw;
        }

        cancelled |= token.IsCancellationRequested;
        current = cancelled ? "백업 중단됨" : "백업 처리 종료";
        Report();
        log.Info("BACKUP_END", $"total={total} | downloaded={downloaded} | skipped={skipped} | failed={failed}"
            + $" | missing={missing} | unprocessed={total - downloaded - skipped - failed - missing} | folderErrors={failures.Count(f => f.Uid.Length == 0)}"
            + $" | cancelled={cancelled} | elapsedMs={clock.ElapsedMilliseconds}");
        return new BackupResult(layout.MessageRoot, total, downloaded, skipped, failed, missing, cancelled, failures);
    }

    private static async Task<bool> RecoverOrSkipAsync(StoredBackupMessage stored, string path,
        BackupStateStore state, ImportLog log, CancellationToken token)
    {
        if (stored.Status is not ("complete" or "staged")) return false;
        if (await BackupFileIntegrity.MatchesAsync(path, stored, token).ConfigureAwait(false))
        {
            if (stored.Status != "complete")
            {
                state.Complete(stored.Mailbox, stored.UidValidity, stored.Uid);
                log.Info("IMAP_RECOVER", $"uid={stored.Uid} | recovered=finalFile");
            }
            return true;
        }
        string temporary = path + ".part";
        if (stored.Status == "staged" && await BackupFileIntegrity.MatchesAsync(temporary, stored, token).ConfigureAwait(false))
        {
            PreserveExisting(path, log);
            File.Move(temporary, path);
            state.Complete(stored.Mailbox, stored.UidValidity, stored.Uid);
            log.Info("IMAP_RECOVER", $"uid={stored.Uid} | recovered=stagedFile");
            return true;
        }
        return false;
    }

    private static async Task DownloadOneAsync(ImapConnectionRunner runner, BackupMailbox mailbox, uint validity,
        RemoteMessage message, string path, BackupLayout layout, BackupStateStore state, Action<long, long?> report,
        ImportLog log, CancellationToken token)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        PreserveExisting(path, log);
        state.BeginMessage(mailbox, validity, message, layout.RelativePath(path));
        string temporary = path + ".part";
        bool staged = false;
        var update = Stopwatch.StartNew();
        try
        {
            var actual = await runner.RunAsync(mailbox.FullName, validity, async (source, _) =>
            {
                using (var output = new FileStream(temporary, FileMode.Create, FileAccess.Write, FileShare.None,
                    64 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan))
                {
                    await source.DownloadAsync(message.Uid, output, (bytes, expected) =>
                    {
                        if (update.ElapsedMilliseconds < 150) return;
                        update.Restart();
                        report(bytes, expected ?? message.Size);
                    }, token).ConfigureAwait(false);
                    await output.FlushAsync(token).ConfigureAwait(false);
                    output.Flush(flushToDisk: true);
                }
                var inspected = await BackupFileIntegrity.InspectAsync(temporary, token).ConfigureAwait(false);
                if (inspected.Bytes == 0 || (message.Size.HasValue && inspected.Bytes != message.Size.Value))
                    throw new IOException($"EML 크기 검증 실패. UID={message.Uid}, expected={message.Size}, actual={inspected.Bytes}");
                return inspected;
            }, token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            // No cancellation between the durable staging record, atomic rename and completion record.
            state.Stage(mailbox.FullName, validity, message.Uid, actual.Bytes, actual.Hash);
            staged = true;
            File.Move(temporary, path);
            state.Complete(mailbox.FullName, validity, message.Uid);
            log.Info("IMAP_SAVED", $"mailbox={mailbox.FullName} | validity={validity} | uid={message.Uid}"
                + $" | bytes={actual.Bytes} | sha256={actual.Hash} | internalDate={message.InternalDate:O} | flags={message.Flags}");
        }
        finally
        {
            if (!staged && File.Exists(temporary))
            {
                try { File.Delete(temporary); }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                { log.Error("IMAP_PART_CLEANUP", $"file={temporary}", ex); }
            }
        }
    }

    private static void PreserveExisting(string path, ImportLog log)
    {
        if (!File.Exists(path)) return;
        string preserved = path + $".previous_{DateTimeOffset.UtcNow:yyyyMMddHHmmss}_{Guid.NewGuid():N}";
        File.Move(path, preserved);
        log.Warn("IMAP_PRESERVED", $"existing unverified file preserved | file={preserved}");
    }

    private static FileStream AcquireLock(string accountRoot)
    {
        try { return new FileStream(Path.Combine(accountRoot, ".backup.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
        catch (IOException ex) { throw new IOException("백업 폴더를 사용할 수 없습니다. 같은 계정의 다른 백업 작업과 폴더 접근 권한을 확인해 주세요.", ex); }
    }

    private static bool IsFolderError(Exception ex) => ex is FolderNotFoundException or ImapCommandException
        or UidValidityChangedException;
}
