using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using EmlToPst.Models;

namespace EmlToPst.Services;

internal sealed class BackupLayout
{
    public BackupLayout(string destinationRoot, ImapConnectionOptions connection)
    {
        connection = connection.Normalize();
        string identity = JsonSerializer.Serialize(new[] { connection.Host, connection.Port.ToString(CultureInfo.InvariantCulture), connection.Account });
        AccountRoot = Path.Combine(Path.GetFullPath(destinationRoot), SafeSegment(connection.Account) + "_" + Hash(identity)[..16]);
        MessageRoot = Path.Combine(AccountRoot, "Messages");
    }

    public string AccountRoot { get; }
    public string MessageRoot { get; }
    public string DatabasePath => Path.Combine(AccountRoot, "backup.sqlite3");

    public string MessagePath(BackupMailbox mailbox, uint validity, uint uid)
    {
        string folder = MessageRoot;
        foreach (string part in mailbox.Parts)
            folder = Path.Combine(folder, SafeSegment(part) + "_" + Hash(part)[..8]);
        return Path.Combine(folder, $"{validity:D10}_{uid:D10}.eml");
    }

    public string RelativePath(string fullPath) => Path.GetRelativePath(AccountRoot, fullPath).Replace('\\', '/');

    public static string ResolveStoredPath(string accountRoot, string relativePath)
    {
        string full = Path.GetFullPath(Path.Combine(accountRoot, relativePath.Replace('/', Path.DirectorySeparatorChar)));
        string prefix = Path.GetFullPath(accountRoot).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!full.StartsWith(prefix, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
            throw new InvalidDataException("백업 파일 경로가 백업 폴더 범위를 벗어납니다.");
        return full;
    }

    private static string SafeSegment(string input)
    {
        string value = new(input.Select(ch => ch < 32 || "<>:\"/\\|?*".Contains(ch) ? '_' : ch).Take(42).ToArray());
        value = value.Trim(' ', '.');
        if (value.Length == 0) value = "Folder";
        return value;
    }

    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
}
