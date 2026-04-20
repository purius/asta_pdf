using System.IO;
using System.Text.Json;

namespace PdfMergeTool.Services;

public sealed class AppSettings
{
    private const int DefaultRecentFileLimit = 10;
    private const string SettingsMutexName = "PdfMergeTool.Settings";
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public int RecentFileLimit { get; set; } = DefaultRecentFileLimit;
    public List<string> RecentFiles { get; set; } = [];
    public string ReorderedSuffix { get; set; } = "_재정렬";
    public string MergedSuffix { get; set; } = "_통합";
    public string A4FitSuffix { get; set; } = "_A4맞춤";
    public string A4OptimizedSuffix { get; set; } = "_A4최적화";
    public double? WindowLeft { get; set; }
    public double? WindowTop { get; set; }
    public double? WindowWidth { get; set; }
    public double? WindowHeight { get; set; }
    public bool AutoCleanTempFiles { get; set; } = true;

    public static AppSettings Load()
    {
        try
        {
            AppPaths.EnsureDirectories();
            if (!File.Exists(AppPaths.SettingsPath))
            {
                return new AppSettings();
            }

            using var mutex = WaitForSettingsMutex();
            var json = File.ReadAllText(AppPaths.SettingsPath);
            return JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, "설정 파일을 읽지 못했습니다.");
            return new AppSettings();
        }
    }

    public void Save()
    {
        try
        {
            AppPaths.EnsureDirectories();
            using var mutex = WaitForSettingsMutex();
            var tempPath = Path.Combine(AppPaths.AppDataDirectory, $"settings-{Guid.NewGuid():N}.tmp");
            File.WriteAllText(tempPath, JsonSerializer.Serialize(this, JsonOptions));
            File.Move(tempPath, AppPaths.SettingsPath, overwrite: true);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, "설정 파일을 저장하지 못했습니다.");
        }
    }

    public void AddRecentFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path) ||
            !File.Exists(path) ||
            !string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var fullPath = Path.GetFullPath(path);
        RecentFiles.RemoveAll(item => string.Equals(item, fullPath, StringComparison.OrdinalIgnoreCase));
        RecentFiles.Insert(0, fullPath);
        TrimRecentFiles();
    }

    public void RemoveMissingRecentFiles()
    {
        RecentFiles = RecentFiles
            .Where(File.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(Math.Clamp(RecentFileLimit, 1, 30))
            .ToList();
    }

    public void ClearRecentFiles()
    {
        RecentFiles.Clear();
    }

    private void TrimRecentFiles()
    {
        RecentFileLimit = Math.Clamp(RecentFileLimit, 1, 30);
        RecentFiles = RecentFiles
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(RecentFileLimit)
            .ToList();
    }

    private static MutexReleaser? WaitForSettingsMutex()
    {
        var mutex = new Mutex(false, SettingsMutexName);
        try
        {
            if (mutex.WaitOne(TimeSpan.FromSeconds(3)))
            {
                return new MutexReleaser(mutex);
            }
        }
        catch (AbandonedMutexException)
        {
            return new MutexReleaser(mutex);
        }

        mutex.Dispose();
        return null;
    }

    private sealed class MutexReleaser(Mutex mutex) : IDisposable
    {
        public void Dispose()
        {
            mutex.ReleaseMutex();
            mutex.Dispose();
        }
    }
}
