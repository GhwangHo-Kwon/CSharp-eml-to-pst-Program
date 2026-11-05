using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using EmlToPst.Models;
using EmlToPst.Services;

namespace EmlToPst.ViewModels;

internal sealed class BackupViewModel : ObservableObject
{
    private readonly ImapBackupService service;
    private readonly FileDialogService dialogs;
    private readonly ImportLog log;
    private readonly OperationCoordinator coordinator;
    private CancellationTokenSource? cancellation;
    private ImapConnectionOptions? loadedConnection;
    private string host = "Host Address 입력", port = "993", account = "", password = "";
    private string destination = "", status = "로그인 정보를 입력한 뒤 메일함을 조회해 주세요.";
    private string current = "", notes = "", messageRoot = "";
    private int processed, total = 1;
    private bool isBusy, indeterminate;
    private IReadOnlyList<BackupFailure> failures = [];

    public BackupViewModel(ImapBackupService service, FileDialogService dialogs, ImportLog log, OperationCoordinator coordinator)
    {
        this.service = service; this.dialogs = dialogs; this.log = log; this.coordinator = coordinator;
        ConnectCommand = new RelayCommand(async () => await LoadMailboxesAsync(), () => IsIdle && HasCredentials);
        SelectDestinationCommand = new RelayCommand(SelectDestination, () => IsIdle);
        StartCommand = new RelayCommand(async () => await StartAsync(), () => IsIdle && HasCredentials
            && loadedConnection is not null && Mailboxes.Any(m => m.IsSelected) && Destination.Length > 0);
        CancelCommand = new RelayCommand(RequestCancel, () => isBusy && cancellation is { IsCancellationRequested: false });
        SelectAllCommand = new RelayCommand(() => SelectAll(true), () => IsIdle && Mailboxes.Count > 0);
        SelectNoneCommand = new RelayCommand(() => SelectAll(false), () => IsIdle && Mailboxes.Count > 0);
        UseForPstCommand = new RelayCommand(() => UseForPst?.Invoke(MessageRoot),
            () => IsIdle && MessageRoot.Length > 0 && Directory.Exists(MessageRoot));
        coordinator.Changed += Refresh;
    }

    public ObservableCollection<MailboxSelection> Mailboxes { get; } = [];
    public RelayCommand ConnectCommand { get; }
    public RelayCommand SelectDestinationCommand { get; }
    public RelayCommand StartCommand { get; }
    public RelayCommand CancelCommand { get; }
    public RelayCommand SelectAllCommand { get; }
    public RelayCommand SelectNoneCommand { get; }
    public RelayCommand UseForPstCommand { get; }
    public event Action<string>? UseForPst;

    public string Host { get => host; set { if (Set(ref host, value)) InvalidateConnection(); } }
    public string Port { get => port; set { if (Set(ref port, value)) InvalidateConnection(); } }
    public string Account { get => account; set { if (Set(ref account, value)) InvalidateConnection(); } }
    public string Destination { get => destination; private set { Set(ref destination, value); Refresh(); } }
    public string Status { get => status; private set => Set(ref status, value); }
    public string Current { get => current; private set => Set(ref current, value); }
    public string Notes { get => notes; private set => Set(ref notes, value); }
    public string MessageRoot { get => messageRoot; private set { Set(ref messageRoot, value); UseForPstCommand.Refresh(); } }
    public int Processed { get => processed; private set => Set(ref processed, value); }
    public int Total { get => total; private set => Set(ref total, Math.Max(1, value)); }
    public bool IsIndeterminate { get => indeterminate; private set => Set(ref indeterminate, value); }
    public IReadOnlyList<BackupFailure> Failures { get => failures; private set => Set(ref failures, value); }
    public bool IsIdle => coordinator.CanStart;
    public string SelectionText => $"메일함 {Mailboxes.Count:N0}개 중 {Mailboxes.Count(m => m.IsSelected):N0}개 선택";
    private bool HasCredentials => Host.Trim().Length > 0 && Account.Trim().Length > 0 && password.Length > 0;

    public void SetPassword(string value) { password = value; Refresh(); }
    public void ClearCredentials() { password = ""; }

    private void InvalidateConnection()
    {
        loadedConnection = null;
        foreach (MailboxSelection row in Mailboxes) row.PropertyChanged -= SelectionChanged;
        Mailboxes.Clear();
        Refresh();
    }

    private void SelectionChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e) => Refresh();

    private void Refresh()
    {
        Notify(nameof(IsIdle)); Notify(nameof(SelectionText));
        ConnectCommand.Refresh(); SelectDestinationCommand.Refresh(); StartCommand.Refresh(); CancelCommand.Refresh();
        SelectAllCommand.Refresh(); SelectNoneCommand.Refresh(); UseForPstCommand.Refresh();
    }

    private ImapConnectionOptions Connection()
    {
        if (!int.TryParse(Port, NumberStyles.None, CultureInfo.InvariantCulture, out int number))
            throw new ArgumentException("포트 번호를 숫자로 입력해 주세요.");
        return new ImapConnectionOptions(Host, number, Account).Normalize();
    }

    private async Task LoadMailboxesAsync() => await RunAsync("IMAP_QUERY", async (token, secret) =>
    {
        Status = "서버에 연결하여 메일함을 조회하는 중입니다.";
        ImapConnectionOptions connection = Connection();
        IReadOnlyList<BackupMailbox> folders = await service.ListMailboxesAsync(connection, secret, log, token);
        foreach (MailboxSelection row in Mailboxes) row.PropertyChanged -= SelectionChanged;
        Mailboxes.Clear();
        foreach (BackupMailbox folder in folders)
        {
            var row = new MailboxSelection(folder);
            row.PropertyChanged += SelectionChanged;
            Mailboxes.Add(row);
        }
        loadedConnection = connection;
        Status = $"연결 성공 · 메일함 {folders.Count:N0}개를 조회했습니다.";
        Notes = "백업할 메일함과 저장 폴더를 선택해 주세요. 비밀번호는 파일에 저장하지 않습니다.";
        log.Info("IMAP_MAILBOXES", $"folders={folders.Count}");
        Refresh();
    });

    private async Task StartAsync() => await RunAsync("IMAP_BACKUP", async (token, secret) =>
    {
        ImapConnectionOptions connection = Connection();
        if (connection != loadedConnection) throw new InvalidOperationException("접속 정보가 변경되었습니다. 메일함을 다시 조회해 주세요.");
        BackupMailbox[] selected = Mailboxes.Where(m => m.IsSelected).Select(m => m.Mailbox).ToArray();
        var options = new ImapBackupOptions(connection, Destination, selected);
        MessageRoot = new BackupLayout(Destination, connection).MessageRoot;
        Failures = []; Processed = 0; Total = 1; Notes = "";
        Status = "백업할 메일 목록을 조회하는 중입니다.";
        CancellationTokenSource? operation = cancellation;
        var progress = new Progress<BackupProgress>(update =>
        {
            if (ReferenceEquals(cancellation, operation) && isBusy) UpdateProgress(update);
        });
        BackupResult result = await service.BackupAsync(options, secret, progress, log, token);
        Failures = result.Failures;
        Processed = result.Downloaded + result.Skipped + result.Failed + result.Missing; Total = result.Total;
        string outcome = result.Cancelled ? "백업 중단됨" : result.Failures.Count > 0 ? "일부 메일 확인 필요" : "백업 완료";
        Status = $"{outcome} · 새 저장 {result.Downloaded:N0} · 기존 백업 {result.Skipped:N0} · 실패 {result.Failed:N0}"
            + $" · 서버에서 사라짐 {result.Missing:N0} · 미처리 {Math.Max(0, result.Total - Processed):N0}";
        Notes = "같은 계정과 저장 폴더로 다시 시작하면 완료된 메일은 검증 후 건너뜁니다. "
            + "서버에서 삭제한 메일도 기존 백업에서는 보존합니다.";
    });

    private async Task RunAsync(string stage, Func<CancellationToken, string, Task> operation)
    {
        if (!coordinator.TryBegin()) return;
        using var run = new CancellationTokenSource();
        string secret = password;
        using var protection = log.ProtectSecret(secret);
        cancellation = run; isBusy = true;
        IsIndeterminate = true; Current = ""; Notes = "";
        Refresh();
        log.Info(stage, "operation started");
        try { await operation(run.Token, secret); }
        catch (OperationCanceledException) when (run.IsCancellationRequested)
        { Status = "작업이 중단되었습니다. 저장 완료된 메일은 보존됩니다."; log.Info(stage, "cancelled"); }
        catch (Exception ex)
        {
            Status = "작업을 완료하지 못했습니다.";
            Notes = log.Redact(ex.Message);
            log.Error(stage, "operation failed", ex);
        }
        finally
        {
            if (log.Problem is string problem) Notes += Environment.NewLine + problem;
            Current = ""; IsIndeterminate = false; isBusy = false; cancellation = null;
            log.Info(stage, "operation ended");
            coordinator.End(); Refresh();
        }
    }

    private void UpdateProgress(BackupProgress update)
    {
        Total = update.Total; Processed = update.Processed;
        IsIndeterminate = update.Total == 0;
        string counts = $"처리 {update.Processed:N0}/{update.Total:N0} · 새 저장 {update.Downloaded:N0} · 기존 백업 {update.Skipped:N0}"
            + $" · 실패 {update.Failed:N0} · 서버에서 사라짐 {update.Missing:N0}";
        Status = cancellation?.IsCancellationRequested == true ? "중단 요청 중 · " + counts : counts;
        Current = update.Current;
        if (update.Bytes > 0)
            Current += update.ExpectedBytes is long expected ? $" · {update.Bytes / 1048576d:N1}/{expected / 1048576d:N1} MB"
                : $" · {update.Bytes / 1048576d:N1} MB";
    }

    private void SelectDestination()
    {
        if (dialogs.SelectBackupFolder() is not string folder) return;
        Destination = folder;
        MessageRoot = "";
        log.Info("UI_BACKUP_DESTINATION", $"folder={folder}");
    }

    private void SelectAll(bool selected)
    {
        foreach (MailboxSelection row in Mailboxes) row.IsSelected = selected;
        log.Info("UI_BACKUP_SELECTION", $"selected={Mailboxes.Count(m => m.IsSelected)} | total={Mailboxes.Count}");
    }

    public void RequestCancel()
    {
        if (cancellation is null || cancellation.IsCancellationRequested) return;
        log.Info("UI_BACKUP_CANCEL", "cancel requested");
        cancellation.Cancel();
        Status = "중단 요청 중입니다. 현재 파일과 백업 기록을 정리합니다.";
        Refresh();
    }
}
