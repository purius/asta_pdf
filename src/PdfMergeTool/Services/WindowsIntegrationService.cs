using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace PdfMergeTool.Services;

internal static class WindowsIntegrationService
{
    private const string ProductId = "PdfMergeTool";
    private const string PdfProgId = "PdfMergeTool.Pdf";
    private static readonly string[] PdfContextMenuPaths =
    [
        @"Software\Classes\SystemFileAssociations\.pdf\shell\PdfMergeTool",
        @"Software\Classes\.pdf\shell\PdfMergeTool",
        @$"Software\Classes\{PdfProgId}\shell\PdfMergeTool"
    ];

    public static string GetCurrentPdfDefaultApp()
    {
        using var userChoiceKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice");
        var userChoice = userChoiceKey?.GetValue("ProgId") as string;
        if (!string.IsNullOrWhiteSpace(userChoice))
        {
            return userChoice;
        }

        using var classesKey = Registry.CurrentUser.OpenSubKey(@"Software\Classes\.pdf");
        return classesKey?.GetValue(null) as string ?? "확인 안 됨";
    }

    public static bool IsPdfDefaultAppRegistered()
    {
        return string.Equals(GetCurrentPdfDefaultApp(), PdfProgId, StringComparison.OrdinalIgnoreCase);
    }

    public static void OpenDefaultAppsSettings()
    {
        Process.Start(new ProcessStartInfo("ms-settings:defaultapps")
        {
            UseShellExecute = true
        });
    }

    public static void OpenFolder(string folder)
    {
        Directory.CreateDirectory(folder);
        Process.Start(new ProcessStartInfo(folder)
        {
            UseShellExecute = true
        });
    }

    public static void RegisterPdfContextMenu()
    {
        var appExe = Environment.ProcessPath ?? System.Reflection.Assembly.GetExecutingAssembly().Location;
        foreach (var contextMenuPath in PdfContextMenuPaths)
        {
            using var verbKey = Registry.CurrentUser.CreateSubKey(contextMenuPath);
            verbKey.SetValue(null, "PDF 통합...");
            verbKey.SetValue("Icon", appExe);
            verbKey.SetValue("MultiSelectModel", "Player");

            using var commandKey = verbKey.CreateSubKey("command");
            commandKey.SetValue(null, $"\"{appExe}\" --merge \"%1\"");
        }
    }

    public static void RemovePdfContextMenu()
    {
        foreach (var contextMenuPath in PdfContextMenuPaths)
        {
            Registry.CurrentUser.DeleteSubKeyTree(contextMenuPath, throwOnMissingSubKey: false);
        }
    }

    public static bool IsPdfContextMenuRegistered()
    {
        return PdfContextMenuPaths.Any(contextMenuPath =>
        {
            using var key = Registry.CurrentUser.OpenSubKey($@"{contextMenuPath}\command");
            return key is not null;
        });
    }

    public static int CleanTempFiles()
    {
        var deleted = 0;
        deleted += DeleteTempFiles("PdfMergeTool-*.pdf");

        var droppedFilesFolder = Path.Combine(Path.GetTempPath(), "PdfMergeTool-DroppedFiles");
        if (Directory.Exists(droppedFilesFolder))
        {
            try
            {
                Directory.Delete(droppedFilesFolder, recursive: true);
                deleted++;
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "드롭 임시 폴더를 삭제하지 못했습니다.");
            }
        }

        return deleted;
    }

    private static int DeleteTempFiles(string searchPattern)
    {
        var deleted = 0;
        foreach (var file in Directory.EnumerateFiles(Path.GetTempPath(), searchPattern))
        {
            try
            {
                File.Delete(file);
                deleted++;
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, $"임시 파일을 삭제하지 못했습니다: {file}");
            }
        }

        return deleted;
    }
}
