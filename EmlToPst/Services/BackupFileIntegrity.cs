using System.IO;
using System.Security.Cryptography;
using EmlToPst.Models;

namespace EmlToPst.Services;

internal static class BackupFileIntegrity
{
    public static async Task<(long Bytes, string Hash)> InspectAsync(string path, CancellationToken token)
    {
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 64 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        byte[] hash = await SHA256.HashDataAsync(input, token).ConfigureAwait(false);
        return (input.Length, Convert.ToHexString(hash));
    }

    public static async Task<bool> MatchesAsync(string path, StoredBackupMessage stored, CancellationToken token)
    {
        if (stored.StoredBytes is not > 0 || string.IsNullOrEmpty(stored.Sha256) || !File.Exists(path)) return false;
        if (new FileInfo(path).Length != stored.StoredBytes) return false;
        var actual = await InspectAsync(path, token).ConfigureAwait(false);
        return actual.Bytes == stored.StoredBytes && actual.Hash.Equals(stored.Sha256, StringComparison.OrdinalIgnoreCase);
    }

    public static bool Matches(string path, StoredBackupMessage stored, CancellationToken token) =>
        MatchesAsync(path, stored, token).GetAwaiter().GetResult();
}
