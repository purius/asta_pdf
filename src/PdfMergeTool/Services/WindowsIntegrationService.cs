using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace PdfMergeTool.Services;

internal static class WindowsIntegrationService
{
    private const string ProductId = "PdfMergeTool";
    private const string PdfProgId = "PdfMergeTool.Pdf";
    private static readonly string[] PdfMergeContextMenuPaths =
    [
        @"Software\Classes\SystemFileAssociations\.pdf\shell\PdfMergeTool",
        @"Software\Classes\.pdf\shell\PdfMergeTool",
        @$"Software\Classes\{PdfProgId}\shell\PdfMergeTool"
    ];

    private static readonly string[] PdfSplitContextMenuPaths =
    [
        @"Software\Classes\SystemFileAssociations\.pdf\shell\PdfMergeToolSplit",
        @"Software\Classes\.pdf\shell\PdfMergeToolSplit",
        @$"Software\Classes\{PdfProgId}\shell\PdfMergeToolSplit"
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
        foreach (var contextMenuPath in PdfMergeContextMenuPaths)
        {
            using var verbKey = Registry.CurrentUser.CreateSubKey(contextMenuPath);
            verbKey.SetValue(null, "PDF 통합...");
            verbKey.SetValue("Icon", appExe);
            verbKey.SetValue("MultiSelectModel", "Player");

            using var commandKey = verbKey.CreateSubKey("command");
            commandKey.SetValue(null, $"\"{appExe}\" --merge \"%1\"");
        }

        foreach (var contextMenuPath in PdfSplitContextMenuPaths)
        {
            using var splitKey = Registry.CurrentUser.CreateSubKey(contextMenuPath);
            splitKey.DeleteValue(string.Empty, throwOnMissingValue: false);
            splitKey.DeleteValue("SubCommands", throwOnMissingValue: false);
            splitKey.DeleteSubKeyTree("SplitByPage", throwOnMissingSubKey: false);
            splitKey.DeleteSubKeyTree("SplitByInterval", throwOnMissingSubKey: false);
            splitKey.DeleteSubKeyTree("SplitByParity", throwOnMissingSubKey: false);
            splitKey.SetValue("MUIVerb", "PDF 분리");
            splitKey.SetValue("Icon", appExe);
            splitKey.SetValue("MultiSelectModel", "Player");

            using var extendedCommandsKey = splitKey.CreateSubKey("ExtendedSubCommandsKey");
            using var shellKey = extendedCommandsKey.CreateSubKey("shell");
            RegisterSplitCommand(shellKey, "SplitByPage", "페이지별 분할", "--split-pages", appExe);
            RegisterSplitCommand(shellKey, "SplitByInterval", "N페이지마다 분리", "--split-interval", appExe);
            RegisterSplitCommand(shellKey, "SplitByParity", "홀수/짝수 분리", "--split-parity", appExe);
        }
    }

    public static void RefreshPdfContextMenuRegistration()
    {
        if (HasAnyPdfContextMenuRegistration())
        {
            RegisterPdfContextMenu();
        }
    }

    public static void RemovePdfContextMenu()
    {
        foreach (var contextMenuPath in PdfMergeContextMenuPaths.Concat(PdfSplitContextMenuPaths))
        {
            Registry.CurrentUser.DeleteSubKeyTree(contextMenuPath, throwOnMissingSubKey: false);
        }
    }

    public static bool IsPdfContextMenuRegistered()
    {
        return PdfMergeContextMenuPaths.All(HasCommand) &&
               PdfSplitContextMenuPaths.All(HasSplitMenu);
    }

    private static void RegisterSplitCommand(
        RegistryKey shellKey,
        string verb,
        string label,
        string argument,
        string appExe)
    {
        using var verbKey = shellKey.CreateSubKey(verb);
        verbKey.DeleteValue(string.Empty, throwOnMissingValue: false);
        verbKey.SetValue("MUIVerb", label);
        verbKey.SetValue("MultiSelectModel", "Player");

        using var commandKey = verbKey.CreateSubKey("command");
        commandKey.SetValue(null, $"\"{appExe}\" {argument} \"%1\"");
    }

    private static bool HasAnyPdfContextMenuRegistration()
    {
        foreach (var contextMenuPath in PdfMergeContextMenuPaths.Concat(PdfSplitContextMenuPaths))
        {
            using var key = Registry.CurrentUser.OpenSubKey(contextMenuPath);
            if (key is not null)
            {
                return true;
            }
        }

        return false;
    }

    private static bool HasCommand(string contextMenuPath)
    {
        using var commandKey = Registry.CurrentUser.OpenSubKey($"{contextMenuPath}\\command");
        return commandKey?.GetValue(null) is string command && !string.IsNullOrWhiteSpace(command);
    }

    private static bool HasSplitMenu(string contextMenuPath)
    {
        using var splitKey = Registry.CurrentUser.OpenSubKey(contextMenuPath);
        if (splitKey?.GetValue("MUIVerb") is not string { Length: > 0 })
        {
            return false;
        }

        return HasSplitCommand(splitKey, "SplitByPage") &&
               HasSplitCommand(splitKey, "SplitByInterval") &&
               HasSplitCommand(splitKey, "SplitByParity");
    }

    private static bool HasSplitCommand(RegistryKey splitKey, string verb)
    {
        using var commandKey = splitKey.OpenSubKey($"ExtendedSubCommandsKey\\shell\\{verb}\\command");
        return commandKey?.GetValue(null) is string command && !string.IsNullOrWhiteSpace(command);
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
