using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using EmlToPst.Models;
using EmlToPst.Services;

namespace EmlToPst.ViewModels;

internal sealed class MainViewModel : INotifyPropertyChanged
{
    private readonly PstImportService importer;
    private readonly FileDialogService dialogs;
    private CancellationTokenSource? cancellation;
    private bool isBusy;
    private string sourceFolder = "";
    private string outputPath = "";
    private string status = "원본 폴더와 새 PST 파일의 저장 위치를 선택해 주세요.";
    private string currentFile = "";
    private string savedPath = "";
    private string logPath = "";
    private string notes = "";
    private int processed;
    private int progressMaximum = 1;
    private bool isIndeterminate;
    private IReadOnlyList<ImportFailure> failures = [];

    public MainViewModel(PstImportService importer, FileDialogService dialogs)
    {
        this.importer = importer;
        this.dialogs = dialogs;
        SelectSourceCommand = new RelayCommand(SelectSource, () => !IsBusy);
        SelectOutputCommand = new RelayCommand(SelectOutput, () => !IsBusy);
        StartCommand = new RelayCommand(async () => await StartAsync(),
            () => !IsBusy && SourceFolder.Length > 0 && OutputPath.Length > 0);
        CancelCommand = new RelayCommand(Cancel, () => IsBusy && cancellation?.IsCancellationRequested == false);
        OpenLogCommand = new RelayCommand(OpenLog, () => LogPath.Length > 0);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public RelayCommand SelectSourceCommand { get; }
    public RelayCommand SelectOutputCommand { get; }
    public RelayCommand StartCommand { get; }
    public RelayCommand CancelCommand { get; }
    public RelayCommand OpenLogCommand { get; }

    public string SourceFolder { get => sourceFolder; private set => Set(ref sourceFolder, value); }
    public string OutputPath { get => outputPath; private set => Set(ref outputPath, value); }
    public string Status { get => status; private set => Set(ref status, value); }
    public string CurrentFile { get => currentFile; private set => Set(ref currentFile, value); }
    public string SavedPath { get => savedPath; private set => Set(ref savedPath, value); }
    public string LogPath { get => logPath; private set => Set(ref logPath, value); }
    public string Notes { get => notes; private set => Set(ref notes, value); }
    public int Processed { get => processed; private set => Set(ref processed, value); }
    public int ProgressMaximum { get => progressMaximum; private set => Set(ref progressMaximum, value); }
    public bool IsIndeterminate { get => isIndeterminate; private set => Set(ref isIndeterminate, value); }
    public IReadOnlyList<ImportFailure> Failures { get => failures; private set => Set(ref failures, value); }

    public bool IncludeSubfolders { get; set; } = true;
    public bool AttachToOutlook { get; set; }
    public bool IsIdle => !IsBusy;

    public bool IsBusy
    {
        get => isBusy;
        private set
        {
            Set(ref isBusy, value);
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsIdle)));
            SelectSourceCommand.Refresh();
            SelectOutputCommand.Refresh();
            StartCommand.Refresh();
            CancelCommand.Refresh();
        }
    }

    private void SelectSource()
    {
        if (dialogs.SelectSourceFolder() is string path)
        {
            SourceFolder = path;
            StartCommand.Refresh();
        }
    }

    private void SelectOutput()
    {
        if (dialogs.SelectOutputPath(OutputPath) is string path)
        {
            OutputPath = path;
            StartCommand.Refresh();
        }
    }

    private async Task StartAsync()
    {
        using var run = new CancellationTokenSource();
        cancellation = run;
        IsBusy = true;
        Failures = [];
        SavedPath = "";
        LogPath = "";
        OpenLogCommand.Refresh();
        Notes = "";
        CurrentFile = "";
        Processed = 0;
        ProgressMaximum = 1;
        IsIndeterminate = true;
        Status = "EML 파일을 찾는 중입니다.";

        var log = new ImportLog();
        try
        {
            LogPath = log.FilePath ?? "";
            OpenLogCommand.Refresh();
            var options = new ImportOptions(SourceFolder, OutputPath, IncludeSubfolders, AttachToOutlook);
            var progress = new Progress<ImportProgress>(UpdateProgress);
            ImportResult result = await importer.ImportAsync(options, progress, log, run.Token);
            ShowResult(result);
        }
        catch (OperationCanceledException) when (run.IsCancellationRequested)
        {
            Status = "변환이 취소되었습니다.";
        }
        catch (Exception ex)
        {
            Status = "변환을 시작하지 못했습니다.";
            Notes = PstImportService.DescribeError(ex);
        }
        finally
        {
            log.Dispose();
            if (log.Problem is string problem)
            {
                AppendNote(problem);
            }

            CurrentFile = "";
            IsIndeterminate = false;
            cancellation = null;
            IsBusy = false;
        }
    }

    private void UpdateProgress(ImportProgress update)
    {
        if (!IsBusy)
        {
            return;
        }

        IsIndeterminate = false;
        ProgressMaximum = update.Total;
        Processed = update.Processed;
        CurrentFile = update.CurrentFile;
        string counts = $"처리 {update.Processed:N0} / {update.Total:N0}건  ·  기록 {update.Imported:N0}건  ·  실패 {update.Failed:N0}건";
        Status = cancellation?.IsCancellationRequested == true ? $"취소 요청 중  ·  {counts}" : counts;
    }

    private void ShowResult(ImportResult result)
    {
        Failures = result.Failures;
        SavedPath = result.PstPath ?? "";
        Notes = string.Join(Environment.NewLine, result.Warnings);
        if (result.Total == 0)
        {
            Status = "선택한 범위에 EML 파일이 없습니다.";
            return;
        }

        ProgressMaximum = result.Total;
        Processed = result.Imported + result.Failures.Count;
        string outcome = result switch
        {
            { Cancelled: true } => "취소됨",
            { PstPath: null } => "저장된 PST 없음",
            _ when result.Imported < result.Total => "부분 저장",
            _ when result.Warnings.Count > 0 => "저장됨 · 안내 확인",
            _ => "처리 완료"
        };
        string attachment = result.AttachedToOutlook ? "  ·  Outlook에 연결됨" : "";
        int saved = result.PstPath is null ? 0 : result.Imported;
        Status = $"{outcome}  ·  저장 {saved:N0}건  ·  실패 {result.Failures.Count:N0}건"
            + $"  ·  미처리 {result.Total - Processed:N0}건{attachment}";
    }

    private void Cancel()
    {
        cancellation?.Cancel();
        Status = "취소 요청 중입니다. 현재 파일 처리 후 PST를 정리합니다.";
        CancelCommand.Refresh();
    }

    private void OpenLog()
    {
        try
        {
            using Process? process = Process.Start(new ProcessStartInfo("notepad.exe")
            {
                ArgumentList = { LogPath }
            });
        }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException)
        {
            AppendNote($"로그 파일을 열지 못했습니다. {PstImportService.DescribeError(ex)}");
        }
    }

    private void AppendNote(string message)
    {
        Notes = Notes.Length == 0 ? message : Notes + Environment.NewLine + message;
    }

    private void Set<T>(ref T field, T value, [CallerMemberName] string? property = null)
    {
        if (!EqualityComparer<T>.Default.Equals(field, value))
        {
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
        }
    }
}
