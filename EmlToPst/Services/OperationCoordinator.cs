namespace EmlToPst.Services;

// Accessed on the WPF dispatcher. Prevents backup/import overlap and new work during shutdown.
internal sealed class OperationCoordinator
{
    private TaskCompletionSource? idle;
    public bool IsBusy => idle is not null;
    public bool IsClosing { get; private set; }
    public bool CanStart => !IsBusy && !IsClosing;
    public Task WhenIdle => idle?.Task ?? Task.CompletedTask;
    public event Action? Changed;

    public bool TryBegin()
    {
        if (!CanStart) return false;
        idle = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        Changed?.Invoke();
        return true;
    }

    public void End()
    {
        TaskCompletionSource? completion = idle;
        idle = null;
        Changed?.Invoke();
        completion?.TrySetResult();
    }

    public void BeginClosing()
    {
        IsClosing = true;
        Changed?.Invoke();
    }
}
