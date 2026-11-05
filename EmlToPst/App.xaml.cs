using System.Windows;
using System.Diagnostics;
using System.Windows.Threading;
using EmlToPst.Services;
using EmlToPst.ViewModels;
using EmlToPst.Views;

namespace EmlToPst;

public partial class App : Application
{
    private readonly ImportLog log = new();
    private readonly Stopwatch lifetime = Stopwatch.StartNew();

    public App()
    {
        log.Info("APP_START", $"pid={Environment.ProcessId} | version={typeof(App).Assembly.GetName().Version}"
            + $" | os={Environment.OSVersion} | runtime={Environment.Version} | baseDirectory={AppContext.BaseDirectory}");
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        SessionEnding += (_, e) => log.Warn("APP_SESSION_END", $"reason={e.ReasonSessionEnding}");
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        try
        {
            var window = new MainWindow(log);
            window.DataContext = new ApplicationViewModel(log, new FileDialogService(window));
            MainWindow = window;
            window.Show();
            log.Info("APP_READY", "main window shown");
            if (log.Problem is string problem)
                MessageBox.Show(window, problem, "로그 기록 오류", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            log.Error("APP_STARTUP_ERROR", "startup failed", ex);
            MessageBox.Show($"프로그램을 시작하지 못했습니다.\n{ex.Message}\n로그: {log.FilePath}",
                "시작 오류", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(-1);
        }
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e) =>
        log.Error("APP_UI_EXCEPTION", "unhandled dispatcher exception", e.Exception);

    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception exception)
            log.Error("APP_UNHANDLED", $"terminating={e.IsTerminating}", exception);
        else log.Warn("APP_UNHANDLED", $"terminating={e.IsTerminating} | {e.ExceptionObject}");
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e) =>
        log.Error("APP_TASK_EXCEPTION", "unobserved task exception", e.Exception);

    protected override void OnExit(ExitEventArgs e)
    {
        try { base.OnExit(e); }
        finally
        {
            DispatcherUnhandledException -= OnDispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
            TaskScheduler.UnobservedTaskException -= OnUnobservedTaskException;
            log.Info("APP_EXIT", $"exitCode={e.ApplicationExitCode} | elapsedMs={lifetime.ElapsedMilliseconds}");
            log.Dispose();
        }
    }
}
