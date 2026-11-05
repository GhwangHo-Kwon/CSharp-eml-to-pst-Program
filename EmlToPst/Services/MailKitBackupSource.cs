using System.IO;
using EmlToPst.Models;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;

namespace EmlToPst.Services;

internal sealed class MailKitBackupSource : IImapBackupSource
{
    private readonly DiskImapClient client = new();
    private IMailFolder? current;

    public async Task ConnectAsync(ImapConnectionOptions options, string password, CancellationToken token)
    {
        client.Timeout = 180_000;
        // Keep standard certificate/host validation. Never accept all certificates.
        await client.ConnectAsync(options.Host, options.Port, SecureSocketOptions.SslOnConnect, token).ConfigureAwait(false);
        await client.AuthenticateAsync(options.Account, password, token).ConfigureAwait(false);
    }

    public async Task<IReadOnlyList<BackupMailbox>> ListMailboxesAsync(CancellationToken token)
    {
        var result = new List<BackupMailbox>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        var pending = new Stack<IMailFolder>();
        pending.Push(client.Inbox);
        foreach (FolderNamespace ns in client.PersonalNamespaces) pending.Push(client.GetFolder(ns));
        while (pending.TryPop(out IMailFolder? folder))
        {
            token.ThrowIfCancellationRequested();
            if (!visited.Add(folder.FullName)) continue;
            if (folder.FullName.Length > 0 && (folder.Attributes & (FolderAttributes.NoSelect | FolderAttributes.NonExistent)) == 0)
            {
                string[] parts = folder.DirectorySeparator == '\0' ? [folder.FullName]
                    : folder.FullName.Split(folder.DirectorySeparator, StringSplitOptions.RemoveEmptyEntries);
                result.Add(new BackupMailbox(folder.FullName, parts));
            }
            if ((folder.Attributes & FolderAttributes.NoInferiors) != 0) continue;
            foreach (IMailFolder child in await folder.GetSubfoldersAsync(false, token).ConfigureAwait(false))
                pending.Push(child);
        }
        return result.OrderBy(f => f.FullName, StringComparer.Ordinal).ToArray();
    }

    public async Task<uint> OpenMailboxAsync(string name, CancellationToken token)
    {
        if (current is null || !current.IsOpen || current.FullName != name)
        {
            current = await client.GetFolderAsync(name, token).ConfigureAwait(false);
            await current.OpenAsync(FolderAccess.ReadOnly, token).ConfigureAwait(false);
        }
        return current.UidValidity;
    }

    public async Task<IReadOnlyList<uint>> GetUidsAsync(CancellationToken token) =>
        (await Folder.SearchAsync(SearchQuery.All, token).ConfigureAwait(false)).Select(u => u.Id).ToArray();

    public async Task<IReadOnlyList<RemoteMessage>> FetchMetadataAsync(IReadOnlyList<uint> uids, CancellationToken token)
    {
        var ids = uids.Select(id => new UniqueId(id)).ToArray();
        var items = MessageSummaryItems.UniqueId | MessageSummaryItems.Size | MessageSummaryItems.InternalDate
            | MessageSummaryItems.Flags | MessageSummaryItems.Envelope;
        var summaries = await Folder.FetchAsync(ids, items, token).ConfigureAwait(false);
        var requested = uids.ToHashSet();
        return summaries.Where(s => requested.Contains(s.UniqueId.Id)).Select(s => new RemoteMessage(
            s.UniqueId.Id, s.Size is uint size ? size : null, s.InternalDate, s.Envelope?.Date,
            s.Flags ?? MessageFlags.None, s.Keywords?.ToArray() ?? [])).ToArray();
    }

    public async Task DownloadAsync(uint uid, Stream destination, Action<long, long?> progress, CancellationToken token)
    {
        try
        {
            using Stream input = await Folder.GetStreamAsync(new UniqueId(uid), token, new TransferProgress(progress)).ConfigureAwait(false);
            await input.CopyToAsync(destination, 64 * 1024, token).ConfigureAwait(false);
        }
        catch (MessageNotFoundException) { throw new RemoteMessageMissingException(uid); }
    }

    private IMailFolder Folder => current ?? throw new InvalidOperationException("메일함을 먼저 열어야 합니다.");

    public async ValueTask DisposeAsync()
    {
        try
        {
            if (client.IsConnected)
            {
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                await client.DisconnectAsync(true, timeout.Token).ConfigureAwait(false);
            }
        }
        catch (Exception ex) when (ex is IOException or OperationCanceledException or ImapProtocolException or ImapCommandException) { }
        finally { client.Dispose(); }
    }

    private sealed class TransferProgress(Action<long, long?> callback) : ITransferProgress
    {
        public void Report(long bytesTransferred, long totalSize) => callback(bytesTransferred, totalSize);
        public void Report(long bytesTransferred) => callback(bytesTransferred, null);
    }

    // GetStreamAsync normally uses an in-memory backing stream. Use temporary files so a large
    // attachment does not require a second, message-sized allocation in managed memory.
    private sealed class DiskImapClient : ImapClient
    {
        protected override ImapFolder CreateImapFolder(ImapFolderConstructorArgs args) => new DiskImapFolder(args);
    }

    private sealed class DiskImapFolder(ImapFolderConstructorArgs args) : ImapFolder(args)
    {
        protected override Stream CreateStream(UniqueId? uid, string section, int offset, int length)
        {
            string folder = Path.Combine(Path.GetTempPath(), "EmlToPst", "Imap");
            Directory.CreateDirectory(folder);
            return new FileStream(Path.Combine(folder, Guid.NewGuid().ToString("N") + ".tmp"),
                FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 64 * 1024,
                FileOptions.Asynchronous | FileOptions.DeleteOnClose | FileOptions.SequentialScan);
        }
    }
}
