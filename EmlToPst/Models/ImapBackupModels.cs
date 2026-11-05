using MailKit;

namespace EmlToPst.Models;

internal sealed record ImapConnectionOptions(string Host, int Port, string Account)
{
    public ImapConnectionOptions Normalize()
    {
        string host = Host.Trim();
        string account = Account.Trim();
        if (host.Length == 0 || host.Contains("://", StringComparison.Ordinal) || host.Contains('/'))
            throw new ArgumentException("서버에는 URL이 아닌 호스트 이름을 입력해 주세요.");
        if (Port is < 1 or > 65535) throw new ArgumentException("포트는 1~65535 범위여야 합니다.");
        if (account.Length == 0) throw new ArgumentException("IMAP 로그인 ID를 입력해 주세요.");
        return new ImapConnectionOptions(host.ToLowerInvariant(), Port, account);
    }
}

internal sealed record BackupMailbox(string FullName, string[] Parts);
internal sealed record ImapBackupOptions(ImapConnectionOptions Connection, string DestinationRoot,
    IReadOnlyList<BackupMailbox> Mailboxes);
internal sealed record MailboxSnapshot(uint UidValidity, IReadOnlyList<uint> Uids);
internal sealed record RemoteMessage(uint Uid, long? Size, DateTimeOffset? InternalDate,
    DateTimeOffset? SentDate, MessageFlags Flags, string[] Keywords);
internal sealed record BackupFailure(string Mailbox, string Uid, string Reason);
internal sealed record BackupProgress(int Total, int Downloaded, int Skipped, int Failed, int Missing,
    string Current, long Bytes = 0, long? ExpectedBytes = null)
{
    public int Processed => Downloaded + Skipped + Failed + Missing;
}
internal sealed record BackupResult(string MessageRoot, int Total, int Downloaded, int Skipped, int Failed,
    int Missing, bool Cancelled, IReadOnlyList<BackupFailure> Failures);
internal sealed record StoredBackupMessage(string Mailbox, uint UidValidity, uint Uid, string[] FolderParts,
    string RelativePath, string Status, long? ExpectedBytes, long? StoredBytes, string? Sha256,
    DateTimeOffset? InternalDate, DateTimeOffset? SentDate, MessageFlags Flags, string[] Keywords);

internal sealed class UidValidityChangedException(string mailbox) : Exception(
    $"메일함 식별 정보(UIDVALIDITY)가 변경되었습니다. 기존 파일은 보존됩니다. 다시 백업해 주세요. ({mailbox})") { }
internal sealed class RemoteMessageMissingException(uint uid) : Exception($"서버에서 메일을 찾을 수 없습니다. UID={uid}") { }
internal sealed class ImapTransferFailedException(Exception inner) : Exception(
    "연결 또는 파일 전송 재시도가 모두 실패하여 백업을 중단했습니다. 완료된 메일은 보존됩니다. " + inner.Message, inner) { }
