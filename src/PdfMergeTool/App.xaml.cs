using System.IO;
using System.Windows;
using PdfMergeTool.Services;

namespace PdfMergeTool;

public partial class App : Application
{
    private SingleInstanceService? _mergeSingleInstance;

    private async void OnStartup(object sender, StartupEventArgs e)
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
        RefreshPdfContextMenuRegistration();

        var args = e.Args.ToList();
        var explorerCommand = PdfExplorerCommandParser.Parse(args);
        var shouldOpenMergeWindow = args.Any(arg => string.Equals(arg, "--merge", StringComparison.OrdinalIgnoreCase));
        var paths = args
            .Where(arg => !arg.StartsWith("--", StringComparison.Ordinal))
            .Where(File.Exists)
            .Where(arg => string.Equals(Path.GetExtension(arg), ".pdf", StringComparison.OrdinalIgnoreCase))
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (explorerCommand != PdfExplorerCommand.None && paths.Count > 0)
        {
            StartDeferredStartupWork(settings);
            await RunExplorerSplitAsync(explorerCommand, paths);
            Shutdown();
            return;
        }

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
            StartDeferredStartupWork(settings);
            return;
        }

        if (paths.Count == 0)
        {
            var window = new MainWindow([], false);
            window.Show();
            StartDeferredStartupWork(settings);
            return;
        }

        foreach (var path in paths)
        {
            var window = new MainWindow([path], false);
            window.Show();
        }

        StartDeferredStartupWork(settings);
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

    private static void StartDeferredStartupWork(AppSettings settings)
    {
        _ = Task.Run(() =>
        {
            try
            {
                if (settings.AutoCleanTempFiles)
                {
                    WindowsIntegrationService.CleanTempFiles();
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "Deferred startup work failed.");
            }
        });
    }

    private static void RefreshPdfContextMenuRegistration()
    {
        try
        {
            WindowsIntegrationService.RefreshPdfContextMenuRegistration();
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, "PDF 우클릭 메뉴 등록을 갱신하지 못했습니다.");
        }
    }

    private static async Task RunExplorerSplitAsync(PdfExplorerCommand command, IReadOnlyList<string> paths)
    {
        var splitMode = command switch
        {
            PdfExplorerCommand.PageSplit => PdfSplitMode.PageByPage,
            PdfExplorerCommand.IntervalSplit => PdfSplitMode.Interval,
            PdfExplorerCommand.ParitySplit => PdfSplitMode.Parity,
            _ => throw new ArgumentOutOfRangeException(nameof(command))
        };
        var interval = 1;
        if (splitMode == PdfSplitMode.Interval)
        {
            var intervalWindow = new SplitIntervalWindow();
            if (intervalWindow.ShowDialog() != true)
            {
                return;
            }

            interval = intervalWindow.Interval;
        }

        var pdfService = new PdfMergeService();
        var successes = new List<string>();
        var failures = new List<string>();

        foreach (var path in paths)
        {
            try
            {
                var pageCount = pdfService.GetPageCount(path);
                var plan = splitMode switch
                {
                    PdfSplitMode.PageByPage => PdfSplitPlanner.CreateIntervalPlan(path, pageCount, 1),
                    PdfSplitMode.Interval => PdfSplitPlanner.CreateIntervalPlan(path, pageCount, interval),
                    PdfSplitMode.Parity => PdfSplitPlanner.CreateParityPlan(path, pageCount),
                    _ => throw new ArgumentOutOfRangeException(nameof(splitMode))
                };
                var results = await pdfService.ExecuteSplitPlanAsync(plan, CancellationToken.None);
                successes.Add($"{Path.GetFileName(path)}: {results.Count}개 파일\n{plan.OutputFolder}");
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, $"PDF 분리 실패: {path}");
                failures.Add($"{Path.GetFileName(path)}: {ex.Message}");
            }
        }

        var messageParts = new List<string>();
        if (successes.Count > 0)
        {
            messageParts.Add($"완료 {successes.Count}개\n{string.Join("\n\n", successes)}");
        }

        if (failures.Count > 0)
        {
            messageParts.Add($"실패 {failures.Count}개\n{string.Join("\n", failures)}");
        }

        MessageBox.Show(
            string.Join("\n\n", messageParts),
            "PDF 분리",
            MessageBoxButton.OK,
            failures.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
    }
}
