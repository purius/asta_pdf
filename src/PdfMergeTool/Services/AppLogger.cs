using System.IO;

namespace PdfMergeTool.Services;

internal static class AppLogger
{
    private static readonly object SyncRoot = new();

    public static void Info(string message) => Write("INFO", message);

    public static void Error(Exception exception, string message)
    {
        Write("ERROR", $"{message}{Environment.NewLine}{exception}");
    }

    private static void Write(string level, string message)
    {
        try
        {
            AppPaths.EnsureDirectories();
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}";
            lock (SyncRoot)
            {
                File.AppendAllText(AppPaths.CurrentLogPath, line);
            }
        }
        catch
        {
            // Logging must never break the app.
        }
    }
}
