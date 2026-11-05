using System.IO;
using System.Text;

namespace EmlToPst.Services;

// One instance per process. UI, IMAP and PST workers share the same writer.
internal sealed class ImportLog : IDisposable
{
    private readonly object sync = new();
    private readonly Dictionary<string, int> secrets = new(StringComparer.Ordinal);
    private StreamWriter? writer;
    private string? problem;

    public ImportLog(string? preferredDirectory = null)
    {
        string[] directories = [preferredDirectory ?? Path.Combine(AppContext.BaseDirectory, "Logs"),
            Path.Combine(Environment.CurrentDirectory, "EmlToPst", "Logs")];
        foreach (string directory in directories.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                Directory.CreateDirectory(directory);
                string path = Path.Combine(directory,
                    $"EmlToPst_{DateTimeOffset.Now:yyyyMMdd_HHmmss_fff}.log");
                writer = new StreamWriter(new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read),
                    new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)) { AutoFlush = true };
                FilePath = path;
                if (!directory.Equals(directories[0], StringComparison.OrdinalIgnoreCase))
                    Write("WARN", "LOG_PATH", "실행 폴더에 로그를 쓸 수 없어 사용자 로컬 폴더에 기록합니다.");
                return;
            }
            catch (Exception ex) when (IsFileError(ex))
            {
                problem = $"로그 파일을 만들지 못했습니다. {ex.Message}";
            }
        }
    }

    public string? FilePath { get; }
    public event Action? StateChanged;
    public string? Problem { get { lock (sync) return writer is not null ? null : problem; } }
    public void Info(string stage, string message) => Write("INFO", stage, message);
    public void Warn(string stage, string message) => Write("WARN", stage, message);
    public void Error(string stage, string message, Exception exception) => Write("ERROR", stage, $"{message} | {exception}");

    // Server error text can contain credentials. Redact them for the whole operation.
    public IDisposable ProtectSecret(string secret)
    {
        lock (sync)
        {
            if (secret.Length > 0) secrets[secret] = secrets.GetValueOrDefault(secret) + 1;
        }
        return new SecretScope(this, secret);
    }

    public string Redact(string value)
    {
        lock (sync)
        {
            foreach (string secret in secrets.Keys.OrderByDescending(s => s.Length))
                value = value.Replace(secret, "[REDACTED]", StringComparison.Ordinal);
            return value;
        }
    }

    private void Write(string level, string stage, string message)
    {
        bool changed = false;
        lock (sync)
        {
            try
            {
                writer?.WriteLine($"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff zzz} [{level}] [{stage}] "
                    + Redact(message).ReplaceLineEndings(" | "));
            }
            catch (Exception ex) when (IsFileError(ex))
            {
                problem = $"로그 기록이 중단되었습니다. {ex.Message}";
                CloseWriter();
                changed = true;
            }
        }
        if (changed)
        {
            try { StateChanged?.Invoke(); }
            catch (InvalidOperationException) { /* The UI dispatcher may already be shutting down. */ }
        }
    }

    private void CloseWriter()
    {
        StreamWriter? stream = writer;
        writer = null;
        try { stream?.Dispose(); }
        catch (Exception ex) when (IsFileError(ex)) { problem ??= ex.Message; }
    }

    public void Dispose()
    {
        lock (sync) { CloseWriter(); secrets.Clear(); }
    }

    private static bool IsFileError(Exception ex) => ex is IOException or UnauthorizedAccessException
        or ArgumentException or NotSupportedException or System.Security.SecurityException;

    private sealed class SecretScope(ImportLog owner, string secret) : IDisposable
    {
        private bool disposed;
        public void Dispose()
        {
            lock (owner.sync)
            {
                if (disposed || secret.Length == 0) return;
                disposed = true;
                if (!owner.secrets.TryGetValue(secret, out int count)) return;
                if (count <= 1) owner.secrets.Remove(secret); else owner.secrets[secret] = count - 1;
            }
        }
    }
}
