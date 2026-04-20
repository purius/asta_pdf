using System.IO;
using System.Text.Json;

namespace PdfMergeTool.Services;

public sealed class SingleInstanceService : IDisposable
{
    private const string MutexName = "PdfMergeTool.SingleInstance";
    private const string QueueEventName = "PdfMergeTool.SingleInstance.MergeQueue";

    private readonly Mutex _mutex;
    private readonly bool _ownsMutex;
    private readonly CancellationTokenSource _cancellation = new();

    private SingleInstanceService(Mutex mutex, bool ownsMutex)
    {
        _mutex = mutex;
        _ownsMutex = ownsMutex;
    }

    public bool IsPrimary => _ownsMutex;

    public static SingleInstanceService Create()
    {
        var mutex = new Mutex(true, MutexName, out var ownsMutex);
        return new SingleInstanceService(mutex, ownsMutex);
    }

    public static Task<bool> SendToPrimaryAsync(IReadOnlyList<string> paths, bool openMergeWindow, CancellationToken cancellationToken)
    {
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            Directory.CreateDirectory(GetQueueDirectory());
            var message = JsonSerializer.Serialize(new SingleInstanceMessage(paths, openMergeWindow));
            var queuePath = Path.Combine(GetQueueDirectory(), $"{DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}-{Guid.NewGuid():N}.json");
            File.WriteAllText(queuePath, message);

            try
            {
                using var queueEvent = EventWaitHandle.OpenExisting(QueueEventName);
                queueEvent.Set();
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                // The primary process will still pick the file up on its next poll.
            }

            return Task.FromResult(true);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, "기존 PDF 통합 창으로 파일 목록을 전달하지 못했습니다.");
            return Task.FromResult(false);
        }
    }

    public void StartServer(Action<IReadOnlyList<string>, bool> onPathsReceived)
    {
        if (!_ownsMutex)
        {
            return;
        }

        _ = Task.Run(async () =>
        {
            Directory.CreateDirectory(GetQueueDirectory());
            using var queueEvent = new EventWaitHandle(false, EventResetMode.AutoReset, QueueEventName);
            while (!_cancellation.IsCancellationRequested)
            {
                try
                {
                    DrainQueue(onPathsReceived);
                    await Task.Run(() => queueEvent.WaitOne(TimeSpan.FromMilliseconds(500)), _cancellation.Token);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch
                {
                    // Keep the pipe alive even if a secondary process exits mid-message.
                }
            }
        }, _cancellation.Token);
    }

    private static void DrainQueue(Action<IReadOnlyList<string>, bool> onPathsReceived)
    {
        foreach (var path in Directory.EnumerateFiles(GetQueueDirectory(), "*.json").OrderBy(File.GetCreationTimeUtc))
        {
            try
            {
                var json = File.ReadAllText(path);
                File.Delete(path);
                var message = JsonSerializer.Deserialize<SingleInstanceMessage>(json);
                if (message is { Paths.Count: > 0 })
                {
                    AppLogger.Info($"기존 PDF 통합 창으로 파일 {message.Paths.Count}개를 전달받았습니다.");
                    onPathsReceived(message.Paths, message.OpenMergeWindow);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex, "PDF 통합 대기열 파일을 처리하지 못했습니다.");
            }
        }
    }

    private static string GetQueueDirectory()
    {
        return Path.Combine(AppPaths.AppDataDirectory, "MergeQueue");
    }

    public void Dispose()
    {
        _cancellation.Cancel();
        _cancellation.Dispose();

        if (_ownsMutex)
        {
            _mutex.ReleaseMutex();
        }

        _mutex.Dispose();
    }
}

public sealed record SingleInstanceMessage(IReadOnlyList<string> Paths, bool OpenMergeWindow);
