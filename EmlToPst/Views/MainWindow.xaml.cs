using System.ComponentModel;
using System.Windows;
using EmlToPst.ViewModels;

namespace EmlToPst.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (DataContext is MainViewModel { IsBusy: true })
        {
            e.Cancel = true;
            MessageBox.Show(this, "변환을 취소한 뒤 PST 정리가 끝나면 창을 닫을 수 있습니다.",
                "변환 진행 중", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        base.OnClosing(e);
    }
}
