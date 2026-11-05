using System.ComponentModel;
using System.Diagnostics;
using EmlToPst.Services;

namespace EmlToPst.ViewModels;

internal sealed class ApplicationViewModel : ObservableObject
{
    private readonly ImportLog log;
    private readonly OperationCoordinator coordinator;
    private int selectedTab;

    public ApplicationViewModel(ImportLog log, FileDialogService dialogs)
    {
        this.log = log;
        coordinator = new OperationCoordinator();
        Pst = new MainViewModel(new PstImportService(), dialogs, log, coordinator);
        Backup = new BackupViewModel(new ImapBackupService(), dialogs, log, coordinator);
        Backup.UseForPst += folder => { Pst.UseBackupFolder(folder); SelectedTab = 1; };
        OpenLogCommand = new RelayCommand(OpenLog, () => log.FilePath is not null);
        coordinator.Changed += () => { Notify(nameof(SessionStatus)); Notify(nameof(LogProblem)); };
        log.StateChanged += () => System.Windows.Application.Current.Dispatcher.BeginInvoke(
            new Action(() => Notify(nameof(LogProblem))));
    }

    public MainViewModel Pst { get; }
    public BackupViewModel Backup { get; }
    public RelayCommand OpenLogCommand { get; }
    public string LogPath => log.FilePath ?? "로그 파일 생성 실패";
    public string LogProblem => log.Problem ?? "";
    public string SessionStatus => coordinator.IsClosing ? "작업을 정리한 뒤 종료합니다." : "실행 시작부터 종료까지 현재 로그에 기록합니다.";
    public bool IsBusy => coordinator.IsBusy;
    public int SelectedTab
    {
        get => selectedTab;
        set { if (Set(ref selectedTab, value)) log.Info("UI_TAB", $"tab={(value == 0 ? "IMAP" : "PST")}"); }
    }

    public async Task PrepareCloseAsync()
    {
        coordinator.BeginClosing();
        Backup.RequestCancel();
        Pst.RequestCancel();
        await coordinator.WhenIdle;
        Backup.ClearCredentials();
    }

    private void OpenLog()
    {
        if (log.FilePath is not string path) return;
        try
        {
            log.Info("UI_OPEN_LOG", "open requested");
            using Process? process = Process.Start(new ProcessStartInfo("notepad.exe") { ArgumentList = { path } });
        }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException)
        { log.Error("UI_OPEN_LOG", "failed", ex); }
    }
}
