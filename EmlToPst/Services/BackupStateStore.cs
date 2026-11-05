using System.Globalization;
using System.IO;
using System.Text.Json;
using EmlToPst.Models;
using MailKit;
using Microsoft.Data.Sqlite;

namespace EmlToPst.Services;

internal sealed class BackupStateStore : IDisposable
{
    private readonly SqliteConnection connection;

    public BackupStateStore(string path, bool readOnly = false)
    {
        connection = new SqliteConnection(new SqliteConnectionStringBuilder
        {
            DataSource = path,
            Mode = readOnly ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate,
            Pooling = false,
            DefaultTimeout = 10
        }.ToString());
        try
        {
            connection.Open();
            int existingVersion;
            using (var check = Command("PRAGMA user_version;"))
                existingVersion = Convert.ToInt32(check.ExecuteScalar(), CultureInfo.InvariantCulture);
            if (existingVersion is not (0 or 1) || (readOnly && existingVersion != 1))
                throw new InvalidDataException("지원하지 않는 백업 DB 버전입니다.");
            if (!readOnly)
            {
                Execute("PRAGMA journal_mode=DELETE; PRAGMA synchronous=FULL;");
                Execute(Schema);
            }
            using var version = Command("PRAGMA user_version;");
            if (Convert.ToInt32(version.ExecuteScalar(), CultureInfo.InvariantCulture) != 1)
                throw new InvalidDataException("지원하지 않는 백업 DB 버전입니다.");
        }
        catch { connection.Dispose(); throw; }
    }

    public void ValidateAccount(ImapConnectionOptions options)
    {
        using var read = Command("SELECT Host, Port, Account FROM BackupInfo WHERE Id=1;");
        using var reader = read.ExecuteReader();
        if (reader.Read())
        {
            if (reader.GetString(0) != options.Host || reader.GetInt32(1) != options.Port || reader.GetString(2) != options.Account)
                throw new InvalidDataException("백업 DB의 계정 정보가 현재 접속 정보와 다릅니다.");
            return;
        }
        reader.Close();
        using var insert = Command("INSERT INTO BackupInfo(Id,Host,Port,Account) VALUES(1,$host,$port,$account);");
        Add(insert, "$host", options.Host); Add(insert, "$port", options.Port); Add(insert, "$account", options.Account);
        insert.ExecuteNonQuery();
    }

    public uint? RecordFolder(BackupMailbox mailbox, uint validity)
    {
        using var read = Command("SELECT UidValidity FROM Folders WHERE Mailbox=$mailbox;");
        Add(read, "$mailbox", mailbox.FullName);
        object? previous = read.ExecuteScalar();
        using var write = Command("""
            INSERT INTO Folders(Mailbox,UidValidity,FolderParts) VALUES($mailbox,$validity,$parts)
            ON CONFLICT(Mailbox) DO UPDATE SET UidValidity=excluded.UidValidity,FolderParts=excluded.FolderParts;
            """);
        Add(write, "$mailbox", mailbox.FullName); Add(write, "$validity", (long)validity);
        Add(write, "$parts", JsonSerializer.Serialize(mailbox.Parts));
        write.ExecuteNonQuery();
        return previous is null or DBNull ? null : checked((uint)Convert.ToInt64(previous, CultureInfo.InvariantCulture));
    }

    public StoredBackupMessage? Find(string mailbox, uint validity, uint uid)
    {
        using var command = Command("SELECT " + Columns + " FROM Messages WHERE Mailbox=$mailbox AND UidValidity=$validity AND Uid=$uid;");
        Key(command, mailbox, validity, uid);
        using var reader = command.ExecuteReader();
        return reader.Read() ? ReadMessage(reader) : null;
    }

    public StoredBackupMessage? FindByPath(string relativePath)
    {
        using var command = Command("SELECT " + Columns + " FROM Messages WHERE RelativePath=$path;");
        Add(command, "$path", relativePath);
        using var reader = command.ExecuteReader();
        return reader.Read() ? ReadMessage(reader) : null;
    }

    public void BeginMessage(BackupMailbox mailbox, uint validity, RemoteMessage message, string relativePath)
    {
        using var command = Command("""
            INSERT INTO Messages(Mailbox,UidValidity,Uid,FolderParts,RelativePath,Status,ExpectedBytes,
                InternalDate,SentDate,Flags,Keywords,Attempts,UpdatedAt)
            VALUES($mailbox,$validity,$uid,$parts,$path,'pending',$size,$internal,$sent,$flags,$keywords,1,$now)
            ON CONFLICT(Mailbox,UidValidity,Uid) DO UPDATE SET
                FolderParts=excluded.FolderParts,RelativePath=excluded.RelativePath,Status='pending',
                ExpectedBytes=excluded.ExpectedBytes,StoredBytes=NULL,Sha256=NULL,
                InternalDate=excluded.InternalDate,SentDate=excluded.SentDate,Flags=excluded.Flags,
                Keywords=excluded.Keywords,Attempts=Messages.Attempts+1,LastError=NULL,UpdatedAt=excluded.UpdatedAt;
            """);
        Key(command, mailbox.FullName, validity, message.Uid);
        Add(command, "$parts", JsonSerializer.Serialize(mailbox.Parts)); Add(command, "$path", relativePath);
        Metadata(command, message);
        Add(command, "$now", DateTimeOffset.UtcNow.ToString("O"));
        command.ExecuteNonQuery();
    }

    public void RefreshMetadata(BackupMailbox mailbox, uint validity, RemoteMessage message)
    {
        using var command = Command("""
            UPDATE Messages SET ExpectedBytes=$size,InternalDate=$internal,SentDate=$sent,
                Flags=$flags,Keywords=$keywords,UpdatedAt=$now
            WHERE Mailbox=$mailbox AND UidValidity=$validity AND Uid=$uid;
            """);
        Key(command, mailbox.FullName, validity, message.Uid); Metadata(command, message);
        Add(command, "$now", DateTimeOffset.UtcNow.ToString("O"));
        command.ExecuteNonQuery();
    }

    public void Stage(string mailbox, uint validity, uint uid, long bytes, string hash)
    {
        using var command = Command("""
            UPDATE Messages SET Status='staged',StoredBytes=$bytes,Sha256=$hash,UpdatedAt=$now
            WHERE Mailbox=$mailbox AND UidValidity=$validity AND Uid=$uid;
            """);
        Key(command, mailbox, validity, uid); Add(command, "$bytes", bytes); Add(command, "$hash", hash);
        Add(command, "$now", DateTimeOffset.UtcNow.ToString("O"));
        if (command.ExecuteNonQuery() != 1) throw new InvalidDataException("저장할 메일의 백업 기록이 없습니다.");
    }

    public void Complete(string mailbox, uint validity, uint uid)
    {
        using var command = Command("""
            UPDATE Messages SET Status='complete',LastError=NULL,UpdatedAt=$now
            WHERE Mailbox=$mailbox AND UidValidity=$validity AND Uid=$uid AND Sha256 IS NOT NULL;
            """);
        Key(command, mailbox, validity, uid); Add(command, "$now", DateTimeOffset.UtcNow.ToString("O"));
        if (command.ExecuteNonQuery() != 1) throw new InvalidDataException("완료 처리할 메일의 검증 정보가 없습니다.");
    }

    public void Fail(string mailbox, uint validity, uint uid, string error)
    {
        using var command = Command("""
            UPDATE Messages SET Status=CASE WHEN Status='staged' THEN 'staged' ELSE 'failed' END,
                LastError=$error,UpdatedAt=$now
            WHERE Mailbox=$mailbox AND UidValidity=$validity AND Uid=$uid;
            """);
        Key(command, mailbox, validity, uid); Add(command, "$error", error);
        Add(command, "$now", DateTimeOffset.UtcNow.ToString("O")); command.ExecuteNonQuery();
    }

    private static void Metadata(SqliteCommand command, RemoteMessage message)
    {
        Add(command, "$size", message.Size); Add(command, "$internal", message.InternalDate?.ToString("O"));
        Add(command, "$sent", message.SentDate?.ToString("O")); Add(command, "$flags", (long)message.Flags);
        Add(command, "$keywords", JsonSerializer.Serialize(message.Keywords));
    }

    private static StoredBackupMessage ReadMessage(SqliteDataReader r) => new(
        r.GetString(0), checked((uint)r.GetInt64(1)), checked((uint)r.GetInt64(2)),
        JsonSerializer.Deserialize<string[]>(r.GetString(3)) ?? throw new InvalidDataException("메일함 경로가 없습니다."),
        r.GetString(4), r.GetString(5), r.IsDBNull(6) ? null : r.GetInt64(6), r.IsDBNull(7) ? null : r.GetInt64(7),
        r.IsDBNull(8) ? null : r.GetString(8), ParseDate(r, 9), ParseDate(r, 10), (MessageFlags)r.GetInt64(11),
        JsonSerializer.Deserialize<string[]>(r.GetString(12)) ?? []);

    private static DateTimeOffset? ParseDate(SqliteDataReader r, int index) => r.IsDBNull(index) ? null
        : DateTimeOffset.ParseExact(r.GetString(index), "O", CultureInfo.InvariantCulture);

    private SqliteCommand Command(string text) { var c = connection.CreateCommand(); c.CommandText = text; return c; }
    private void Execute(string text) { using var c = Command(text); c.ExecuteNonQuery(); }
    private static void Add(SqliteCommand c, string name, object? value) => c.Parameters.AddWithValue(name, value ?? DBNull.Value);
    private static void Key(SqliteCommand c, string mailbox, uint validity, uint uid)
    {
        Add(c, "$mailbox", mailbox); Add(c, "$validity", (long)validity); Add(c, "$uid", (long)uid);
    }
    public void Dispose() => connection.Dispose();

    private const string Columns = "Mailbox,UidValidity,Uid,FolderParts,RelativePath,Status,ExpectedBytes,StoredBytes,Sha256,InternalDate,SentDate,Flags,Keywords";
    internal const string Schema = """
        CREATE TABLE IF NOT EXISTS BackupInfo(Id INTEGER PRIMARY KEY CHECK(Id=1),Host TEXT NOT NULL,Port INTEGER NOT NULL,Account TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS Folders(Mailbox TEXT PRIMARY KEY,UidValidity INTEGER NOT NULL,FolderParts TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS Messages(
            Mailbox TEXT NOT NULL,UidValidity INTEGER NOT NULL,Uid INTEGER NOT NULL,
            FolderParts TEXT NOT NULL,RelativePath TEXT NOT NULL UNIQUE,Status TEXT NOT NULL,
            ExpectedBytes INTEGER,StoredBytes INTEGER,Sha256 TEXT,InternalDate TEXT,SentDate TEXT,
            Flags INTEGER NOT NULL,Keywords TEXT NOT NULL,Attempts INTEGER NOT NULL DEFAULT 0,
            LastError TEXT,UpdatedAt TEXT NOT NULL,PRIMARY KEY(Mailbox,UidValidity,Uid));
        PRAGMA user_version=1;
        """;
}
