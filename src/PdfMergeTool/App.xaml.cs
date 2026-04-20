using System.IO;
using System.Windows;
using PdfMergeTool.Services;

namespace PdfMergeTool;

public partial class App : Application
{
    private SingleInstanceService? _mergeSingleInstance;

    private void OnStartup(object sender, StartupEventArgs e)
    {
        AppPaths.EnsureDirectories();
        AppLogger.Info("앱을 시작합니다.");
        DispatcherUnhandledException += (_, args) =>
        {
            AppLogger.Error(args.Exception, "처리되지 않은 UI 예외가 발생했습니다.");
        };
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            if (args.ExceptionObject is Exception ex)
            {
                AppLogger.Error(ex, "처리되지 않은 앱 예외가 발생했습니다.");
            }
        };

        var settings = AppSettings.Load();
        if (settings.AutoCleanTempFiles)
        {
            WindowsIntegrationService.CleanTempFiles();
        }

        var args = e.Args.ToList();
        var shouldOpenMergeWindow = args.Any(arg => string.Equals(arg, "--merge", StringComparison.OrdinalIgnoreCase));
        var paths = args
            .Where(arg => !arg.StartsWith("--", StringComparison.Ordinal))
            .Where(File.Exists)
            .Where(arg => string.Equals(Path.GetExtension(arg), ".pdf", StringComparison.OrdinalIgnoreCase))
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (shouldOpenMergeWindow)
        {
            if (ForwardMergeRequestToPrimary(paths))
            {
                return;
            }

            MergeWindow? window = null;
            var pendingMergePaths = new List<string>();
            _mergeSingleInstance?.StartServer((receivedPaths, _) =>
            {
                Dispatcher.BeginInvoke(() =>
                {
                    if (window is null)
                    {
                        pendingMergePaths.AddRange(receivedPaths);
                        return;
                    }

                    window.AddFiles(receivedPaths);
                    if (window.WindowState == WindowState.Minimized)
                    {
                        window.WindowState = WindowState.Normal;
                    }

                    window.Activate();
                });
            });
            window = new MergeWindow(paths.Concat(pendingMergePaths));
            window.Show();
            return;
        }

        if (paths.Count == 0)
        {
            var window = new MainWindow([], false);
            window.Show();
            return;
        }

        foreach (var path in paths)
        {
            var window = new MainWindow([path], false);
            window.Show();
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _mergeSingleInstance?.Dispose();
        base.OnExit(e);
    }

    private bool ForwardMergeRequestToPrimary(IReadOnlyList<string> paths)
    {
        _mergeSingleInstance = SingleInstanceService.Create();
        if (_mergeSingleInstance.IsPrimary)
        {
            AppLogger.Info("PDF 통합 기본 인스턴스로 실행합니다.");
            return false;
        }

        AppLogger.Info($"기존 PDF 통합 창으로 파일 {paths.Count}개 전달을 시도합니다.");
        var sent = SingleInstanceService
            .SendToPrimaryAsync(paths, openMergeWindow: true, CancellationToken.None)
            .GetAwaiter()
            .GetResult();

        _mergeSingleInstance.Dispose();
        _mergeSingleInstance = null;

        if (!sent)
        {
            AppLogger.Info("기존 PDF 통합 창으로 전달하지 못해 새 창으로 실행합니다.");
            return false;
        }

        AppLogger.Info("기존 PDF 통합 창으로 전달을 완료하고 보조 인스턴스를 종료합니다.");
        Shutdown(0);
        Environment.Exit(0);
        System.Diagnostics.Process.GetCurrentProcess().Kill();
        return true;
    }
}
