using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace PdfMergeTool.Services;

internal static class SavePathPromptService
{
    public static string? ResolveOutputPath(
        Window owner,
        string proposedPath,
        string title,
        string filter = "PDF 파일 (*.pdf)|*.pdf")
    {
        var currentPath = Path.GetFullPath(proposedPath);
        while (File.Exists(currentPath))
        {
            var result = MessageBox.Show(
                owner,
                $"이미 같은 이름의 파일이 있습니다.\n\n{currentPath}\n\n덮어쓰겠습니까?",
                title,
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                return currentPath;
            }

            if (result == MessageBoxResult.Cancel)
            {
                return null;
            }

            var dialog = CreateSaveDialog(currentPath, title, filter);
            if (dialog.ShowDialog(owner) != true)
            {
                return null;
            }

            currentPath = Path.GetFullPath(dialog.FileName);
        }

        return currentPath;
    }

    public static SaveFileDialog CreateSaveDialog(string path, string title, string filter = "PDF 파일 (*.pdf)|*.pdf")
    {
        return new SaveFileDialog
        {
            Filter = filter,
            DefaultExt = ".pdf",
            AddExtension = true,
            OverwritePrompt = false,
            InitialDirectory = Path.GetDirectoryName(path) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            FileName = Path.GetFileName(path),
            Title = title
        };
    }
}
