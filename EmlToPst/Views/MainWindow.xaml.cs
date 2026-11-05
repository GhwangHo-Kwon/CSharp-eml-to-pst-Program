using System.ComponentModel;
using System.Windows;
using EmlToPst.ViewModels;
using EmlToPst.Services;

namespace EmlToPst.Views;

public partial class MainWindow : Window
{
    private readonly ImportLog log;
    private bool waitingForClose;
    private bool allowClose;

    internal MainWindow(ImportLog log)
    {
        this.log = log;
        InitializeComponent();
    }

    protected override async void OnClosing(CancelEventArgs e)
    {
        base.OnClosing(e);
        if (e.Cancel || allowClose) return;
        log.Info("UI_CLOSE_REQUEST", "window close requested");
        if (DataContext is ApplicationViewModel viewModel)
        {
            e.Cancel = true;
            if (waitingForClose) return;
            waitingForClose = true;
            try
            {
                if (viewModel.IsBusy) log.Info("APP_CLOSE_WAIT", "cancel active operation and wait for file cleanup");
                await viewModel.PrepareCloseAsync();
                // Leave the current Closing event before calling Close again.
                await System.Windows.Threading.Dispatcher.Yield();
                allowClose = true;
                Close();
            }
            catch (Exception ex)
            {
                log.Error("APP_CLOSE_ERROR", "cleanup failed", ex);
                waitingForClose = false;
            }
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        BackupPage.ClearPassword();
        if (DataContext is ApplicationViewModel viewModel) viewModel.Backup.ClearCredentials();
        log.Info("APP_WINDOW_CLOSED", "main window closed");
        base.OnClosed(e);
    }
}
