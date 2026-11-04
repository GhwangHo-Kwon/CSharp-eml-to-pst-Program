using System.Windows.Input;

namespace EmlToPst.ViewModels;

internal sealed class RelayCommand(Action execute, Func<bool> canExecute) : ICommand
{
    public event EventHandler? CanExecuteChanged;

    public bool CanExecute(object? parameter) => canExecute();

    public void Execute(object? parameter) => execute();

    public void Refresh() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}
