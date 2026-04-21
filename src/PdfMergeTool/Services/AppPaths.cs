using System.IO;

namespace PdfMergeTool.Services;

internal static class AppPaths
{
    public static string AppDataDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "PdfMergeTool");

    public static string LogsDirectory { get; } = Path.Combine(AppDataDirectory, "Logs");

    public static string ViewerRuntimeDirectory { get; } = Path.Combine(AppDataDirectory, "ViewerRuntime");

    public static string FallbackCacheDirectory { get; } = Path.Combine(ViewerRuntimeDirectory, "FallbackCache");

    public static string RecoveredPdfDirectory { get; } = Path.Combine(ViewerRuntimeDirectory, "RecoveredPdf");

    public static string SettingsPath { get; } = Path.Combine(AppDataDirectory, "settings.json");

    public static string CurrentLogPath { get; } = Path.Combine(LogsDirectory, $"PdfMergeTool-{DateTime.Now:yyyyMMdd}.log");

    public static void EnsureDirectories()
    {
        Directory.CreateDirectory(AppDataDirectory);
        Directory.CreateDirectory(LogsDirectory);
        Directory.CreateDirectory(ViewerRuntimeDirectory);
        Directory.CreateDirectory(FallbackCacheDirectory);
        Directory.CreateDirectory(RecoveredPdfDirectory);
    }
}
