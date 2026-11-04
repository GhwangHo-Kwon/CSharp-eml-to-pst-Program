using System.Windows;
using EmlToPst.Services;
using EmlToPst.ViewModels;
using EmlToPst.Views;

namespace EmlToPst;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var window = new MainWindow();
        window.DataContext = new MainViewModel(
            new PstImportService(),
            new FileDialogService(window));
        MainWindow = window;
        window.Show();
    }
}
