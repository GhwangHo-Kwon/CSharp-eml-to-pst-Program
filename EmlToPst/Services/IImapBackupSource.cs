using System.IO;
using EmlToPst.Models;

namespace EmlToPst.Services;

// Keeps network access separate from checkpoint/file handling and enables offline regression tests.
internal interface IImapBackupSource : IAsyncDisposable
{
    Task ConnectAsync(ImapConnectionOptions options, string password, CancellationToken token);
    Task<IReadOnlyList<BackupMailbox>> ListMailboxesAsync(CancellationToken token);
    Task<uint> OpenMailboxAsync(string name, CancellationToken token);
    Task<IReadOnlyList<uint>> GetUidsAsync(CancellationToken token);
    Task<IReadOnlyList<RemoteMessage>> FetchMetadataAsync(IReadOnlyList<uint> uids, CancellationToken token);
    Task DownloadAsync(uint uid, Stream destination, Action<long, long?> progress, CancellationToken token);
}
