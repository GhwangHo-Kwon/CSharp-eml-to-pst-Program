using System.IO;
using System.Text;

namespace EmlToPst.Services;

internal sealed class ImportLog : IDisposable
{
    private StreamWriter? writer;

    public ImportLog()
    {
        try
        {
            string localData = Environment.CurrentDirectory;
            if (localData.Length == 0)
            {
                throw new DirectoryNotFoundException("Windows의 로컬 앱 데이터 폴더를 찾을 수 없습니다.");
            }

            string folder = Path.Combine(localData, "EmlToPst", "Logs");
            Directory.CreateDirectory(folder);
            string path = Path.Combine(folder, $"EmlToPst_{DateTimeOffset.Now:yyyyMMdd_HHmmss}.log");
            writer = new StreamWriter(
                new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read),
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            writer.AutoFlush = true;
            FilePath = path;
        }
        catch (Exception ex) when (IsFileError(ex))
        {
            Problem = $"로그 파일을 만들지 못했습니다. {ex.Message}";
            Dispose();
        }
    }

    public string? FilePath { get; }
    public string? Problem { get; private set; }

    public void Info(string stage, string message) => Write("INFO", stage, message);

    public void Warn(string stage, string message) => Write("WARN", stage, message);

    public void Error(string stage, string message, Exception exception)
    {
        Write("ERROR", stage,
            $"{message} | {exception.GetType().Name} [0x{exception.HResult:X8}] {exception.Message} | {exception.StackTrace}");
    }

    public void Dispose()
    {
        StreamWriter? stream = writer;
        writer = null;
        try
        {
            stream?.Dispose();
        }
        catch (Exception ex) when (IsFileError(ex))
        {
            Problem ??= $"로그 파일을 닫는 중 오류가 발생했습니다. {ex.Message}";
        }
    }

    private void Write(string level, string stage, string message)
    {
        try
        {
            writer?.WriteLine($"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff zzz} [{level}] [{stage}] "
                + message.ReplaceLineEndings(" | "));
        }
        catch (Exception ex) when (IsFileError(ex))
        {
            Problem = $"로그 기록이 중단되었습니다. {ex.Message}";
            Dispose();
        }
    }

    private static bool IsFileError(Exception exception) =>
        exception is IOException or UnauthorizedAccessException or ArgumentException;
}
