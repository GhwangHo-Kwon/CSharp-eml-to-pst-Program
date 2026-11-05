using System.IO;
using EmlToPst.Models;

namespace EmlToPst.Services;

// Locates the account DB above each EML, including when several accounts are imported at once.
internal sealed class BackupMetadataCatalog(ImportLog log) : IDisposable
{
    private readonly Dictionary<string, string?> directoryRoots = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, BackupStateStore> stores = new(StringComparer.OrdinalIgnoreCase);

    public StoredBackupMessage? Find(string emlPath, CancellationToken token)
    {
        string? root = FindRoot(Path.GetDirectoryName(Path.GetFullPath(emlPath))!);
        if (root is null) return null;
        if (!stores.TryGetValue(root, out BackupStateStore? store))
        {
            store = new BackupStateStore(Path.Combine(root, "backup.sqlite3"), readOnly: true);
            stores.Add(root, store);
            log.Info("IMAP_METADATA", $"opened | root={root}");
        }
        string relative = Path.GetRelativePath(root, emlPath).Replace('\\', '/');
        StoredBackupMessage? metadata = store.FindByPath(relative);
        if (metadata is null)
            throw new InvalidDataException("백업 DB에 이 EML의 기록이 없습니다. 같은 폴더에서 IMAP 백업을 다시 실행해 주세요.");
        if (metadata.Status != "complete" || !BackupFileIntegrity.Matches(emlPath, metadata, token))
            throw new InvalidDataException("백업 완료/무결성 검증에 실패한 EML입니다. IMAP 백업을 다시 실행해 주세요.");
        return metadata;
    }

    private string? FindRoot(string directory)
    {
        var visited = new List<string>();
        string? found = null;
        for (string? current = directory; current is not null; current = Directory.GetParent(current)?.FullName)
        {
            if (directoryRoots.TryGetValue(current, out found)) break;
            visited.Add(current);
            if (File.Exists(Path.Combine(current, "backup.sqlite3"))) { found = current; break; }
        }
        foreach (string path in visited) directoryRoots[path] = found;
        return found;
    }

    public void Dispose()
    {
        foreach (BackupStateStore store in stores.Values) store.Dispose();
    }
}
