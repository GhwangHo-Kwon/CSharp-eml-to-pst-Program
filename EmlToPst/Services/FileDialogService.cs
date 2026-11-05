using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace EmlToPst.Services;

internal sealed class FileDialogService(Window owner)
{
    public string? SelectBackupFolder()
    {
        var dialog = new OpenFolderDialog { Title = "웹메일 백업 저장 폴더 선택" };
        return dialog.ShowDialog(owner) == true ? dialog.FolderName : null;
    }

    public string? SelectSourceFolder()
    {
        var dialog = new OpenFolderDialog { Title = "EML 원본 폴더 선택" };
        return dialog.ShowDialog(owner) == true ? dialog.FolderName : null;
    }

    public string? SelectOutputPath(string currentPath)
    {
        var dialog = new SaveFileDialog
        {
            Title = "새 Outlook 데이터 파일 저장",
            Filter = "Outlook 데이터 파일 (*.pst)|*.pst",
            DefaultExt = ".pst",
            AddExtension = true,
            FileName = string.IsNullOrWhiteSpace(currentPath) ? "EML_Backup.pst" : Path.GetFileName(currentPath),
            CheckPathExists = true,
            OverwritePrompt = false
        };

        return dialog.ShowDialog(owner) == true ? dialog.FileName : null;
    }
}
