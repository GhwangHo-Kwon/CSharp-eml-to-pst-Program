using System.Windows;
using System.Windows.Controls;
using EmlToPst.ViewModels;

namespace EmlToPst.Views;

public partial class BackupView : UserControl
{
    public BackupView() => InitializeComponent();
    private void PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (DataContext is BackupViewModel viewModel) viewModel.SetPassword(ImapPassword.Password);
    }
    internal void ClearPassword() => ImapPassword.Clear();
}
