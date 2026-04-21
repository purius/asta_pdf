using System.IO;
using System.Globalization;
using System.Printing;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using PdfMergeTool.Services;

namespace PdfMergeTool;

public partial class MainWindow : Window
{
    private const string ViewerHost = "pdfviewer.local";
    private const string ViewerCacheHost = "pdfcache.local";
    private const int A4ImageMaxWidthPixels = 2480;
    private const int A4ImageMaxHeightPixels = 3508;
    private const int A4OptimizedMaxWidthPixels = 2200;
    private const int A4OptimizedMaxHeightPixels = 3112;
    private const string PageTransferClipboardFormat = "PdfMergeTool.Pages.v1";
    private static readonly JsonSerializerOptions PageTransferJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly AppSettings _settings = AppSettings.Load();
    private readonly PdfMergeService _pdfService = new();
    private readonly PdfFallbackRenderService _fallbackRenderService = new();
    private bool _viewerReady;
    private string? _currentPdfPath;
    private string? _referencePdfPath;
    private string? _pendingPdfPath;
    private string? _pendingReferencePdfPath;
    private bool _pendingDirtyAfterLoad;
    private IReadOnlyList<int> _pageOrder = [];
    private IReadOnlyDictionary<int, int> _pageRotations = new Dictionary<int, int>();
    private IReadOnlyList<int> _selectedPages = [];
    private int? _activePage;
    private bool _isDirty;
    private MergeWindow? _mergeWindow;
    private TaskCompletionSource<bool>? _printReadyCompletion;
    private TaskCompletionSource<IReadOnlyList<PdfImagePage>>? _a4ExportCompletion;
    private readonly List<ExportedA4Page> _a4ExportedPages = [];
    private int _a4ExportExpectedPages;
    private readonly List<NativeFileDropTarget> _viewerDropTargets = [];
    private int? _lastLoggedPageCount;
    private string? _servedPdfLinkPath;
    private string? _recoveredPdfPath;
    private bool _viewerLoadRecoveryAttempted;
    private bool _fallbackModeActive;
    private string? _fallbackSessionId;

    public MainWindow(IEnumerable<string> initialFiles, bool openMergeWindow)
    {
        InitializeComponent();
        ApplyWindowSettings();
        AddHandler(DragDrop.PreviewDragOverEvent, new DragEventHandler(OnWindowDragOver), true);
        AddHandler(DragDrop.PreviewDropEvent, new DragEventHandler(OnWindowDrop), true);
        Loaded += OnLoaded;
        Closing += OnClosing;
        OpenFiles(initialFiles, openMergeWindow);
    }

    private sealed record ExportedA4Page(int Index, PdfImagePage Image);
    private sealed record PageTransferPayload(string SourcePath, List<PdfPageTransform> Pages, bool Cut);
    private sealed record ExternalPagesDropMessage(string SourcePath, List<PdfPageTransform> Pages, int InsertionIndex);
    private sealed record ExternalFilesDropMessage(List<string> Paths, int InsertionIndex);

    public void OpenFiles(IEnumerable<string> paths, bool openMergeWindow = false)
    {
        var pdfPaths = paths
            .Where(path => File.Exists(path))
            .Where(path => string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (pdfPaths.Count == 0)
        {
            if (openMergeWindow)
            {
                OpenMergeWindow();
            }

            return;
        }

        if (openMergeWindow)
        {
            LoadPdf(pdfPaths[0]);
            OpenMergeWindow(pdfPaths);
            return;
        }

        if (string.IsNullOrWhiteSpace(_currentPdfPath))
        {
            LoadPdf(pdfPaths[0]);
            pdfPaths = pdfPaths.Skip(1).ToList();
        }

        foreach (var path in pdfPaths)
        {
            var window = new MainWindow([path], false);
            window.Show();
        }
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        UpdatePdfContextMenuOption();
        UpdateRecentFilesMenu();
        await InitializeViewerAsync();
    }

    private async Task InitializeViewerAsync()
    {
        await PdfViewer.EnsureCoreWebView2Async();
        PdfViewer.AllowExternalDrop = false;
        PdfViewer.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        PdfViewer.CoreWebView2.Settings.IsZoomControlEnabled = false;
        RegisterNativeViewerDropTargets();
        PdfViewer.CoreWebView2.NavigationCompleted += (_, _) => RegisterNativeViewerDropTargets();
        PdfViewer.CoreWebView2.WebMessageReceived += async (_, args) =>
        {
            using var document = JsonDocument.Parse(args.WebMessageAsJson);
            if (!document.RootElement.TryGetProperty("type", out var typeProperty))
            {
                return;
            }

            var type = typeProperty.GetString();
            if (type == "viewerReady")
            {
                _viewerReady = true;
                if (string.IsNullOrWhiteSpace(_pendingPdfPath))
                {
                    return;
                }

                var pending = _pendingPdfPath;
                var pendingReference = _pendingReferencePdfPath;
                var pendingDirty = _pendingDirtyAfterLoad;
                _pendingPdfPath = null;
                _pendingReferencePdfPath = null;
                _pendingDirtyAfterLoad = false;
                _referencePdfPath = pendingReference ?? pending;
                await SendPdfToViewerAsync(pending, pendingDirty);
                return;
            }

            if (type == "printReady")
            {
                var ready = document.RootElement.TryGetProperty("ready", out var readyElement) &&
                            readyElement.GetBoolean();
                _printReadyCompletion?.TrySetResult(ready);
                return;
            }

            if (type == "addBlankA4Page")
            {
                OnAddBlankA4PageClick(this, new RoutedEventArgs());
                return;
            }

            if (type == "pasteClipboardImage")
            {
                OnPasteClipboardImageClick(this, new RoutedEventArgs());
                return;
            }

            if (type == "fitAllPagesToA4")
            {
                OnFitAllPagesToA4Click(this, new RoutedEventArgs());
                return;
            }

            if (type == "optimizeA4FileSize")
            {
                OnOptimizeA4FileSizeClick(this, new RoutedEventArgs());
                return;
            }

            if (type == "copySelectedPages")
            {
                await CopySelectedPagesToClipboardAsync(cut: false);
                return;
            }

            if (type == "cutSelectedPages")
            {
                await CopySelectedPagesToClipboardAsync(cut: true);
                return;
            }

            if (type == "pasteTransferredPages")
            {
                await PastePagesOrImageAsync();
                return;
            }

            if (type == "insertExternalPages")
            {
                await InsertExternalPagesAsync(document.RootElement);
                return;
            }

            if (type == "insertExternalFiles")
            {
                await InsertExternalFilesAsync(document.RootElement);
                return;
            }

            if (type == "a4PageImage")
            {
                ReceiveA4PageImage(document.RootElement);
                return;
            }

            if (type == "a4ExportComplete")
            {
                CompleteA4PageImageExport(document.RootElement);
                return;
            }

            if (type == "viewerDiagnostic")
            {
                LogViewerDiagnostic(document.RootElement);
                return;
            }

            if (type == "fallbackRenderRequest")
            {
                await HandleFallbackRenderRequestAsync(document.RootElement);
                return;
            }

            if (type == "viewerLoadFailed")
            {
                var message = document.RootElement.TryGetProperty("message", out var messageElement)
                    ? messageElement.GetString()
                    : "알 수 없는 오류";
                await HandleViewerLoadFailedAsync(message);
                return;
            }

            if (type == "pageOrderChanged" &&
                document.RootElement.TryGetProperty("pageOrder", out var pageOrderElement))
            {
                _pageOrder = pageOrderElement
                    .EnumerateArray()
                    .Select(element => element.GetInt32())
                    .ToList();
                _selectedPages = document.RootElement.TryGetProperty("selectedPages", out var selectedElement)
                    ? selectedElement.EnumerateArray().Select(element => element.GetInt32()).ToList()
                    : [];
                _pageRotations = document.RootElement.TryGetProperty("rotations", out var rotationsElement) &&
                                 rotationsElement.ValueKind == JsonValueKind.Object
                    ? ReadPageRotations(rotationsElement)
                    : new Dictionary<int, int>();
                _activePage = document.RootElement.TryGetProperty("activePage", out var activeElement) &&
                              activeElement.ValueKind == JsonValueKind.Number
                    ? activeElement.GetInt32()
                    : null;
                _isDirty = document.RootElement.TryGetProperty("isDirty", out var dirtyElement) &&
                           dirtyElement.GetBoolean();
                if (_pageOrder.Count > 0 && _lastLoggedPageCount != _pageOrder.Count)
                {
                    _lastLoggedPageCount = _pageOrder.Count;
                    AppLogger.Info($"PDF 로드 완료: {_pageOrder.Count}페이지, {_referencePdfPath ?? _currentPdfPath}");
                }

                UpdateWindowTitle();
            }
        };

        var viewerFolder = Path.Combine(AppContext.BaseDirectory, "Assets", "PdfViewer");
        PdfViewer.CoreWebView2.SetVirtualHostNameToFolderMapping(
            ViewerHost,
            viewerFolder,
            CoreWebView2HostResourceAccessKind.Allow);
        PdfViewer.CoreWebView2.SetVirtualHostNameToFolderMapping(
            ViewerCacheHost,
            AppPaths.ViewerRuntimeDirectory,
            CoreWebView2HostResourceAccessKind.Allow);
        PdfViewer.CoreWebView2.Navigate($"https://{ViewerHost}/viewer.html");
    }

    private async void LoadPdf(string path, string? referencePath = null, bool dirtyAfterLoad = false)
    {
        CleanupCompatibilityState();
        _currentPdfPath = path;
        _referencePdfPath = referencePath ?? path;
        _pageOrder = [];
        _pageRotations = new Dictionary<int, int>();
        _selectedPages = [];
        _activePage = null;
        _isDirty = dirtyAfterLoad;
        _lastLoggedPageCount = null;
        CurrentFileText.Text = IsSamePath(path, _referencePdfPath)
            ? path
            : $"{_referencePdfPath} (편집 중)";
        ViewerLoading.Visibility = Visibility.Collapsed;
        UpdateWindowTitle();
        if (!dirtyAfterLoad)
        {
            _settings.AddRecentFile(path);
            _settings.Save();
            UpdateRecentFilesMenu();
        }

        if (!_viewerReady)
        {
            _pendingPdfPath = path;
            _pendingReferencePdfPath = _referencePdfPath;
            _pendingDirtyAfterLoad = dirtyAfterLoad;
            return;
        }

        await SendPdfToViewerAsync(path, dirtyAfterLoad);
    }

    private void UpdateWindowTitle()
    {
        var name = string.IsNullOrWhiteSpace(_currentPdfPath)
            ? "PDF 뷰어"
            : $"{Path.GetFileName(_referencePdfPath ?? _currentPdfPath)} - PDF 뷰어";
        Title = _isDirty ? $"* {name}" : name;
    }

    private static bool IsSamePath(string? left, string? right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
        {
            return false;
        }

        return string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsPdfFile(string path)
    {
        return string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSupportedImageFile(string path)
    {
        return Path.GetExtension(path).ToLowerInvariant() is ".jpg" or ".jpeg" or ".png" or ".bmp" or ".gif" or ".tif" or ".tiff";
    }

    private static bool IsSupportedInsertionFile(string path)
    {
        return IsPdfFile(path) || IsSupportedImageFile(path);
    }

    private Task SendPdfToViewerAsync(string path, bool dirtyAfterLoad = false)
    {
        var servedPath = PrepareServedPdfPath(path);
        var servedFileName = Path.GetFileName(servedPath);
        var pdfUrl = $"https://{ViewerHost}/ServedPdf/{Uri.EscapeDataString(servedFileName)}?v={Guid.NewGuid():N}";
        var fileLength = new FileInfo(path).Length;
        var message = JsonSerializer.Serialize(new
        {
            type = "loadPdf",
            url = pdfUrl,
            isDirty = dirtyAfterLoad,
            sourcePath = _referencePdfPath ?? path,
            largeDocumentHint = fileLength >= 80L * 1024 * 1024,
            fileLength
        });
        PdfViewer.CoreWebView2.PostWebMessageAsJson(message);
        AppLogger.Info($"PDF 뷰어 로드 요청: {path}, {fileLength:N0} bytes");
        return Task.CompletedTask;
    }

    private async Task HandleViewerLoadFailedAsync(string? message)
    {
        var failureMessage = string.IsNullOrWhiteSpace(message) ? "알 수 없는 오류" : message;
        AppLogger.Info($"PDF 뷰어 로드 실패: {_currentPdfPath}, {failureMessage}");

        if (!_viewerLoadRecoveryAttempted &&
            !string.IsNullOrWhiteSpace(_currentPdfPath) &&
            File.Exists(_currentPdfPath))
        {
            _viewerLoadRecoveryAttempted = true;
            var recoveredPath = Path.Combine(
                AppPaths.RecoveredPdfDirectory,
                $"recovered-{Guid.NewGuid():N}.pdf");

            try
            {
                await _pdfService.NormalizeForViewingAsync(_currentPdfPath, recoveredPath, CancellationToken.None);
                if (File.Exists(recoveredPath) && new FileInfo(recoveredPath).Length > 0)
                {
                    CleanupRecoveredPdf();
                    _recoveredPdfPath = recoveredPath;
                    _currentPdfPath = recoveredPath;
                    AppLogger.Info($"PDF 자동 복구 후 재로드: {_referencePdfPath} -> {recoveredPath}");
                    await SendPdfToViewerAsync(recoveredPath, _isDirty);
                    return;
                }
            }
            catch (Exception ex)
            {
                TryDeleteFile(recoveredPath);
                AppLogger.Error(ex, $"PDF 자동 복구 실패: {_currentPdfPath}");
            }
        }

        if (await TryOpenFallbackViewerAsync(failureMessage))
        {
            return;
        }

        MessageBox.Show(
            this,
            $"PDF를 열지 못했습니다.\n\n{failureMessage}",
            "PDF 열기 실패",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
    }

    private async Task<bool> TryOpenFallbackViewerAsync(string failureMessage)
    {
        if (string.IsNullOrWhiteSpace(_currentPdfPath) || !File.Exists(_currentPdfPath))
        {
            return false;
        }

        try
        {
            CloseFallbackSession();
            var session = _fallbackRenderService.OpenDocument(_currentPdfPath);
            _fallbackSessionId = session.SessionId;
            _fallbackModeActive = true;
            AppLogger.Info($"PDFium fallback 뷰어 사용: {_referencePdfPath ?? _currentPdfPath}, {session.Pages.Count} pages, reason={failureMessage}");

            var message = JsonSerializer.Serialize(new
            {
                type = "loadFallbackDocument",
                sessionId = session.SessionId,
                isDirty = _isDirty,
                sourcePath = _referencePdfPath ?? _currentPdfPath,
                largeDocumentMode = session.LargeDocumentMode,
                pages = session.Pages.Select(page => new
                {
                    page.Number,
                    page.Width,
                    page.Height
                })
            });
            PdfViewer.CoreWebView2.PostWebMessageAsJson(message);
            await Task.CompletedTask;
            return true;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, $"PDFium fallback 뷰어 열기 실패: {_currentPdfPath}");
            CloseFallbackSession();
            return false;
        }
    }

    private async Task HandleFallbackRenderRequestAsync(JsonElement root)
    {
        if (!root.TryGetProperty("requestId", out var requestIdElement) ||
            !root.TryGetProperty("sessionId", out var sessionIdElement) ||
            !root.TryGetProperty("pageNumber", out var pageNumberElement))
        {
            return;
        }

        var requestId = requestIdElement.GetString() ?? string.Empty;
        var sessionId = sessionIdElement.GetString() ?? string.Empty;
        var role = root.TryGetProperty("role", out var roleElement)
            ? roleElement.GetString()
            : "main";
        var thumbnail = string.Equals(role, "thumb", StringComparison.OrdinalIgnoreCase);
        var targetWidth = root.TryGetProperty("targetWidth", out var targetWidthElement)
            ? targetWidthElement.GetInt32()
            : thumbnail ? 96 : 900;
        var rotation = root.TryGetProperty("rotation", out var rotationElement)
            ? rotationElement.GetInt32()
            : 0;

        try
        {
            var rendered = await _fallbackRenderService.RenderPageAsync(
                sessionId,
                pageNumberElement.GetInt32(),
                targetWidth,
                rotation,
                thumbnail,
                CancellationToken.None);
            var message = JsonSerializer.Serialize(new
            {
                type = "fallbackPageRendered",
                requestId,
                success = true,
                pageNumber = rendered.PageNumber,
                role = thumbnail ? "thumb" : "main",
                url = ToCacheUrl(rendered.ImagePath),
                rendered.SourceWidth,
                rendered.SourceHeight
            });
            PdfViewer.CoreWebView2.PostWebMessageAsJson(message);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, $"PDFium fallback 페이지 렌더링 실패: session={sessionId}, page={pageNumberElement.GetInt32()}");
            var message = JsonSerializer.Serialize(new
            {
                type = "fallbackPageRendered",
                requestId,
                success = false,
                error = ex.Message
            });
            PdfViewer.CoreWebView2.PostWebMessageAsJson(message);
        }
    }

    private static string ToCacheUrl(string path)
    {
        var relative = Path.GetRelativePath(AppPaths.ViewerRuntimeDirectory, path)
            .Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .Select(Uri.EscapeDataString);
        return $"https://{ViewerCacheHost}/{string.Join("/", relative)}?v={Guid.NewGuid():N}";
    }

    private void LogViewerDiagnostic(JsonElement root)
    {
        var level = root.TryGetProperty("level", out var levelElement)
            ? levelElement.GetString()
            : "info";
        var message = root.TryGetProperty("message", out var messageElement)
            ? messageElement.GetString()
            : null;
        var details = root.TryGetProperty("details", out var detailsElement)
            ? detailsElement.ToString()
            : string.Empty;
        AppLogger.Info($"PDF 뷰어 진단[{level}]: {message} {details}");
    }

    private void CleanupCompatibilityState()
    {
        _viewerLoadRecoveryAttempted = false;
        _fallbackModeActive = false;
        CloseFallbackSession();
        CleanupRecoveredPdf();
    }

    private void CloseFallbackSession()
    {
        _fallbackRenderService.CloseSession(_fallbackSessionId);
        _fallbackSessionId = null;
    }

    private void CleanupRecoveredPdf()
    {
        TryDeleteFile(_recoveredPdfPath);
        _recoveredPdfPath = null;
    }

    private static void TryDeleteFile(string? path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Temporary compatibility files are best-effort cleanup.
        }
    }

    private string PrepareServedPdfPath(string sourcePath)
    {
        CleanupServedPdfLink();

        var viewerFolder = Path.Combine(AppContext.BaseDirectory, "Assets", "PdfViewer");
        var servedFolder = Path.Combine(viewerFolder, "ServedPdf");
        Directory.CreateDirectory(servedFolder);
        CleanupServedPdfFolder(servedFolder);

        var servedPath = Path.Combine(servedFolder, $"current-{Guid.NewGuid():N}.pdf");
        if (!TryCreateHardLink(servedPath, sourcePath))
        {
            File.Copy(sourcePath, servedPath, overwrite: true);
        }

        _servedPdfLinkPath = servedPath;
        return servedPath;
    }

    private void CleanupServedPdfLink()
    {
        if (string.IsNullOrWhiteSpace(_servedPdfLinkPath))
        {
            return;
        }

        try
        {
            if (File.Exists(_servedPdfLinkPath))
            {
                File.Delete(_servedPdfLinkPath);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, $"PDF 뷰어 임시 링크를 삭제하지 못했습니다: {_servedPdfLinkPath}");
        }
        finally
        {
            _servedPdfLinkPath = null;
        }
    }

    private static void CleanupServedPdfFolder(string servedFolder)
    {
        try
        {
            foreach (var file in Directory.EnumerateFiles(servedFolder, "current-*.pdf"))
            {
                try
                {
                    if (File.GetLastWriteTimeUtc(file) < DateTime.UtcNow.AddHours(-12))
                    {
                        File.Delete(file);
                    }
                }
                catch
                {
                    // A currently loading viewer may still hold its own link.
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, $"PDF 뷰어 임시 링크 폴더를 정리하지 못했습니다: {servedFolder}");
        }
    }

    private static bool TryCreateHardLink(string linkPath, string sourcePath)
    {
        try
        {
            return CreateHardLink(linkPath, sourcePath, IntPtr.Zero);
        }
        catch
        {
            return false;
        }
    }

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CreateHardLink(string lpFileName, string lpExistingFileName, IntPtr lpSecurityAttributes);

    private static IReadOnlyDictionary<int, int> ReadPageRotations(JsonElement rotationsElement)
    {
        var rotations = new Dictionary<int, int>();
        foreach (var property in rotationsElement.EnumerateObject())
        {
            try
            {
                var pageNumber = Convert.ToInt32(property.Name, CultureInfo.InvariantCulture);
                rotations[pageNumber] = property.Value.GetInt32();
            }
            catch
            {
                // Ignore malformed viewer state keys.
            }
        }

        return rotations;
    }

    private void OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_isDirty)
        {
            var result = MessageBox.Show(
                this,
                "저장하지 않은 페이지 변경사항이 있습니다.\n저장하지 않고 닫을까요?",
                "PDF 뷰어",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
            {
                e.Cancel = true;
                return;
            }
        }

        SaveWindowSettings();
        DisposeNativeViewerDropTargets();
        CleanupServedPdfLink();
        CleanupCompatibilityState();
        _fallbackRenderService.Dispose();
    }

    private void ApplyWindowSettings()
    {
        if (_settings.WindowWidth is >= 960 and <= 5000)
        {
            Width = _settings.WindowWidth.Value;
        }

        if (_settings.WindowHeight is >= 620 and <= 5000)
        {
            Height = _settings.WindowHeight.Value;
        }

        if (_settings.WindowLeft is { } left &&
            _settings.WindowTop is { } top &&
            left > -2000 &&
            top > -2000 &&
            left < SystemParameters.VirtualScreenWidth &&
            top < SystemParameters.VirtualScreenHeight)
        {
            Left = left;
            Top = top;
            WindowStartupLocation = WindowStartupLocation.Manual;
        }
    }

    private void SaveWindowSettings()
    {
        if (WindowState == WindowState.Minimized)
        {
            return;
        }

        _settings.WindowLeft = RestoreBounds.Left;
        _settings.WindowTop = RestoreBounds.Top;
        _settings.WindowWidth = RestoreBounds.Width;
        _settings.WindowHeight = RestoreBounds.Height;
        _settings.Save();
    }

    private void RegisterNativeViewerDropTargets()
    {
        DisposeNativeViewerDropTargets();
        foreach (var hwnd in NativeFileDropTarget.GetWindowAndDescendants(PdfViewer.Handle))
        {
            try
            {
                var dropTarget = new NativeFileDropTarget(
                    hwnd,
                    (paths, screenPoint) => Dispatcher.BeginInvoke(() => HandleNativeFileDragOver(paths, screenPoint)),
                    (paths, screenPoint) => Dispatcher.BeginInvoke(() => HandleNativeFileDrop(paths, screenPoint)),
                    (payload, screenPoint) => Dispatcher.BeginInvoke(() => SendNativePageTransferMessage("nativePageTransferDragOver", payload, screenPoint)),
                    () => Dispatcher.BeginInvoke(SendNativeDropLeaveMessage),
                    (payload, screenPoint) => Dispatcher.BeginInvoke(() => SendNativePageTransferMessage("nativePageTransferDrop", payload, screenPoint)));
                dropTarget.Register();
                _viewerDropTargets.Add(dropTarget);
            }
            catch
            {
                // Some WebView2 child windows may reject OLE drop registration; other child HWNDs still cover the viewer.
            }
        }
    }

    private void DisposeNativeViewerDropTargets()
    {
        foreach (var dropTarget in _viewerDropTargets)
        {
            dropTarget.Dispose();
        }

        _viewerDropTargets.Clear();
    }

    private void SendNativePageTransferMessage(string type, string payload, Point screenPoint)
    {
        if (!_viewerReady || PdfViewer.CoreWebView2 is null)
        {
            return;
        }

        var clientPoint = PdfViewer.PointFromScreen(screenPoint);
        var message = JsonSerializer.Serialize(new
        {
            type,
            payload,
            clientX = clientPoint.X,
            clientY = clientPoint.Y
        });
        PdfViewer.CoreWebView2.PostWebMessageAsJson(message);
    }

    private void SendNativeFileDropMessage(string type, IReadOnlyList<string> paths, Point screenPoint)
    {
        if (!_viewerReady || PdfViewer.CoreWebView2 is null)
        {
            return;
        }

        var clientPoint = PdfViewer.PointFromScreen(screenPoint);
        var message = JsonSerializer.Serialize(new
        {
            type,
            paths,
            clientX = clientPoint.X,
            clientY = clientPoint.Y
        });
        PdfViewer.CoreWebView2.PostWebMessageAsJson(message);
    }

    private void SendNativeDropLeaveMessage()
    {
        if (!_viewerReady || PdfViewer.CoreWebView2 is null)
        {
            return;
        }

        PdfViewer.CoreWebView2.PostWebMessageAsJson("""
            {"type":"nativePageTransferDragLeave"}
            """);
    }

    private void HandleNativeFileDragOver(IReadOnlyList<string> paths, Point screenPoint)
    {
        if (CanInsertDroppedFilesIntoCurrentDocument())
        {
            SendNativeFileDropMessage("nativeFileDragOver", paths, screenPoint);
        }
    }

    private void HandleNativeFileDrop(IReadOnlyList<string> paths, Point screenPoint)
    {
        if (CanInsertDroppedFilesIntoCurrentDocument())
        {
            SendNativeFileDropMessage("nativeFileDrop", paths, screenPoint);
            return;
        }

        OpenFiles(paths.Where(IsPdfFile));
    }

    private bool CanInsertDroppedFilesIntoCurrentDocument()
    {
        return _viewerReady &&
               PdfViewer.CoreWebView2 is not null &&
               !string.IsNullOrWhiteSpace(_currentPdfPath) &&
               _pageOrder.Count > 0;
    }

    private void SendViewerCommand(string command, object? options = null)
    {
        if (!_viewerReady || PdfViewer.CoreWebView2 is null)
        {
            return;
        }

        var message = JsonSerializer.Serialize(new
        {
            type = "command",
            command,
            options
        });
        PdfViewer.CoreWebView2.PostWebMessageAsJson(message);
    }

    private void ReceiveA4PageImage(JsonElement root)
    {
        if (_a4ExportCompletion is null)
        {
            return;
        }

        try
        {
            var index = root.GetProperty("index").GetInt32();
            _a4ExportExpectedPages = root.GetProperty("total").GetInt32();
            var width = root.GetProperty("width").GetInt32();
            var height = root.GetProperty("height").GetInt32();
            var base64 = root.GetProperty("base64").GetString() ?? string.Empty;
            var commaIndex = base64.IndexOf(',', StringComparison.Ordinal);
            if (commaIndex >= 0)
            {
                base64 = base64[(commaIndex + 1)..];
            }

            _a4ExportedPages.RemoveAll(page => page.Index == index);
            _a4ExportedPages.Add(new ExportedA4Page(
                index,
                new PdfImagePage(Convert.FromBase64String(base64), width, height)));
        }
        catch (Exception ex)
        {
            _a4ExportCompletion.TrySetException(ex);
        }
    }

    private void CompleteA4PageImageExport(JsonElement root)
    {
        if (_a4ExportCompletion is null)
        {
            return;
        }

        var success = !root.TryGetProperty("success", out var successElement) || successElement.GetBoolean();
        if (!success)
        {
            var message = root.TryGetProperty("message", out var messageElement)
                ? messageElement.GetString()
                : null;
            _a4ExportCompletion.TrySetException(new InvalidOperationException(message ?? "A4 페이지 변환에 실패했습니다."));
            return;
        }

        if (_a4ExportExpectedPages > 0 && _a4ExportedPages.Count != _a4ExportExpectedPages)
        {
            _a4ExportCompletion.TrySetException(new InvalidOperationException("일부 페이지 이미지를 받지 못했습니다."));
            return;
        }

        _a4ExportCompletion.TrySetResult(
            _a4ExportedPages
                .OrderBy(page => page.Index)
                .Select(page => page.Image)
                .ToList());
    }

    private async Task<IReadOnlyList<PdfImagePage>> ExportCurrentPagesAsA4ImagesAsync(bool optimizeSize = false)
    {
        if (_fallbackModeActive && !string.IsNullOrWhiteSpace(_fallbackSessionId))
        {
            AppLogger.Info($"PDFium fallback A4 이미지 내보내기: {_referencePdfPath ?? _currentPdfPath}");
            return await _fallbackRenderService.ExportPagesAsImagesAsync(
                _fallbackSessionId,
                _pageOrder,
                _pageRotations,
                optimizeSize ? A4OptimizedMaxWidthPixels : A4ImageMaxWidthPixels,
                optimizeSize ? A4OptimizedMaxHeightPixels : A4ImageMaxHeightPixels,
                optimizeSize ? 86 : 92,
                CancellationToken.None);
        }

        _a4ExportedPages.Clear();
        _a4ExportExpectedPages = 0;
        _a4ExportCompletion = new TaskCompletionSource<IReadOnlyList<PdfImagePage>>(TaskCreationOptions.RunContinuationsAsynchronously);

        try
        {
            SendViewerCommand("exportA4PageImages", optimizeSize
                ? new
                {
                    quality = 0.86,
                    maxWidth = A4OptimizedMaxWidthPixels,
                    maxHeight = A4OptimizedMaxHeightPixels,
                    statusText = "A4 기준 용량 최적화 중..."
                }
                : new
                {
                    quality = 0.92,
                    maxWidth = A4ImageMaxWidthPixels,
                    maxHeight = A4ImageMaxHeightPixels,
                    statusText = "A4 맞춤 이미지 생성 중..."
                });
            var completed = await Task.WhenAny(_a4ExportCompletion.Task, Task.Delay(TimeSpan.FromMinutes(5)));
            if (completed != _a4ExportCompletion.Task)
            {
                throw new TimeoutException("A4 페이지 변환 시간이 초과되었습니다.");
            }

            return await _a4ExportCompletion.Task;
        }
        finally
        {
            _a4ExportCompletion = null;
            _a4ExportedPages.Clear();
            _a4ExportExpectedPages = 0;
        }
    }

    private void OpenMergeWindow(IEnumerable<string>? initialFiles = null)
    {
        var paths = initialFiles?.ToList() ?? new List<string>();
        if (paths.Count == 0 && !string.IsNullOrWhiteSpace(_currentPdfPath))
        {
            paths.Add(_currentPdfPath);
        }

        if (_mergeWindow is { IsVisible: true })
        {
            _mergeWindow.AddFiles(paths);
            _mergeWindow.Activate();
            return;
        }

        _mergeWindow = new MergeWindow(paths);
        if (IsVisible)
        {
            _mergeWindow.Owner = this;
        }

        _mergeWindow.Closed += (_, _) => _mergeWindow = null;
        _mergeWindow.Show();
    }

    private void UpdatePdfContextMenuOption()
    {
        PdfContextMenuToggleItem.IsChecked = WindowsIntegrationService.IsPdfContextMenuRegistered();
    }

    private void UpdateRecentFilesMenu()
    {
        _settings.RemoveMissingRecentFiles();
        RecentFilesMenuItem.Items.Clear();

        if (_settings.RecentFiles.Count == 0)
        {
            RecentFilesMenuItem.Items.Add(new MenuItem
            {
                Header = "최근 파일 없음",
                IsEnabled = false
            });
            return;
        }

        foreach (var path in _settings.RecentFiles)
        {
            var item = new MenuItem
            {
                Header = Path.GetFileName(path),
                ToolTip = path,
                Tag = path
            };
            item.Click += OnRecentFileClick;
            RecentFilesMenuItem.Items.Add(item);
        }

        RecentFilesMenuItem.Items.Add(new Separator());
        var clearItem = new MenuItem { Header = "최근 파일 지우기" };
        clearItem.Click += OnClearRecentFilesClick;
        RecentFilesMenuItem.Items.Add(clearItem);
    }

    private void OnTogglePdfContextMenuClick(object sender, RoutedEventArgs e)
    {
        try
        {
            if (PdfContextMenuToggleItem.IsChecked)
            {
                WindowsIntegrationService.RegisterPdfContextMenu();
            }
            else
            {
                WindowsIntegrationService.RemovePdfContextMenu();
            }

            UpdatePdfContextMenuOption();
        }
        catch (Exception ex)
        {
            UpdatePdfContextMenuOption();
            MessageBox.Show(this, ex.Message, "우클릭 메뉴 설정 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OnRecentFileClick(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem { Tag: string path })
        {
            OpenFiles([path]);
        }
    }

    private void OnClearRecentFilesClick(object sender, RoutedEventArgs e)
    {
        _settings.ClearRecentFiles();
        _settings.Save();
        UpdateRecentFilesMenu();
    }

    private async void OnCheckForUpdatesClick(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem menuItem)
        {
            menuItem.IsEnabled = false;
        }

        try
        {
            var update = await UpdateService.CheckForUpdatesAsync();
            if (update.IsUpdateAvailable)
            {
                var result = MessageBox.Show(
                    this,
                    $"새 버전 {update.LatestVersionText}이 있습니다.\n현재 버전: {update.CurrentVersionText}\n\n다운로드 페이지를 열까요?",
                    "업데이트 확인",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    UpdateService.OpenUpdatePage(update);
                }

                return;
            }

            MessageBox.Show(
                this,
                $"현재 최신 버전을 사용 중입니다.\n현재 버전: {update.CurrentVersionText}",
                "업데이트 확인",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                this,
                $"업데이트 정보를 확인하지 못했습니다.\n인터넷 연결을 확인한 뒤 다시 시도해주세요.\n\n{ex.Message}",
                "업데이트 확인 실패",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            if (sender is MenuItem clickedMenuItem)
            {
                clickedMenuItem.IsEnabled = true;
            }
        }
    }

    private void OnOpenSettingsClick(object sender, RoutedEventArgs e)
    {
        var settingsWindow = new SettingsWindow(_settings)
        {
            Owner = this
        };

        if (settingsWindow.ShowDialog() == true)
        {
            UpdatePdfContextMenuOption();
            UpdateRecentFilesMenu();
        }
    }

    private void OnOpenPdfClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "PDF 파일 (*.pdf)|*.pdf",
            Multiselect = true,
            Title = "PDF 열기"
        };

        if (dialog.ShowDialog(this) == true)
        {
            OpenFiles(dialog.FileNames);
        }
    }

    private void OnOpenMergeClick(object sender, RoutedEventArgs e)
    {
        OpenMergeWindow();
    }

    private void OnRefreshViewerClick(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(_currentPdfPath))
        {
            LoadPdf(_currentPdfPath);
        }
    }

    private async void OnSavePageOrderClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_currentPdfPath) || _pageOrder.Count == 0)
        {
            MessageBox.Show(this, "먼저 PDF를 열어주세요.", "PDF 저장", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var referencePath = _referencePdfPath ?? _currentPdfPath;
        var outputPath = SavePathPromptService.ResolveOutputPath(this, referencePath, "PDF 저장");
        if (outputPath is null)
        {
            return;
        }

        try
        {
            var result = await _pdfService.SaveTransformedPagesAsync(_currentPdfPath, GetCurrentPageTransforms(), outputPath, CancellationToken.None);
            var message = string.IsNullOrWhiteSpace(result.WarningMessage)
                ? $"저장 완료:\n{result.OutputPath}"
                : $"저장 완료:\n{result.OutputPath}\n\n참고: {result.WarningMessage}";

            _isDirty = false;
            UpdateWindowTitle();
            SendViewerCommand("markClean");
            LoadPdf(result.OutputPath);
            MessageBox.Show(this, message, "PDF 저장", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "PDF 저장 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void OnExtractSelectedPagesClick(object sender, RoutedEventArgs e)
    {
        if (!EnsurePdfLoaded("PDF 추출"))
        {
            return;
        }

        if (_selectedPages.Count == 0)
        {
            MessageBox.Show(this, "추출할 썸네일 페이지를 선택하세요.", "PDF 추출", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var referencePath = _referencePdfPath ?? _currentPdfPath!;
        var folder = Path.GetDirectoryName(referencePath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var defaultPath = Path.Combine(folder, $"{Path.GetFileNameWithoutExtension(referencePath)}_추출.pdf");
        var dialog = SavePathPromptService.CreateSaveDialog(defaultPath, "선택 페이지 추출");

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var outputPath = SavePathPromptService.ResolveOutputPath(this, dialog.FileName, "PDF 추출");
        if (outputPath is null)
        {
            return;
        }

        try
        {
            var selected = _pageOrder
                .Where(page => _selectedPages.Contains(page))
                .Select(page => new PdfPageTransform(page, GetRotation(page)))
                .ToList();
            var result = await _pdfService.SaveTransformedPagesAsync(_currentPdfPath!, selected, outputPath, CancellationToken.None);
            MessageBox.Show(this, $"저장 완료:\n{result.OutputPath}", "PDF 추출", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "PDF 추출 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void OnSplitPagesClick(object sender, RoutedEventArgs e)
    {
        if (!EnsurePdfLoaded("PDF 분할"))
        {
            return;
        }

        var referencePath = _referencePdfPath ?? _currentPdfPath!;
        var folder = Path.GetDirectoryName(referencePath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var outputFolder = Path.Combine(folder, $"{Path.GetFileNameWithoutExtension(referencePath)}_분할");

        try
        {
            var results = await _pdfService.SplitPagesAsync(
                _currentPdfPath!,
                GetCurrentPageTransforms(),
                outputFolder,
                Path.GetFileNameWithoutExtension(referencePath) ?? "분할",
                CancellationToken.None);

            MessageBox.Show(this, $"{results.Count}개 파일로 분할했습니다.\n{outputFolder}", "PDF 분할", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "PDF 분할 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void OnAddBlankA4PageClick(object sender, RoutedEventArgs e)
    {
        if (!EnsurePdfLoaded("빈 페이지 추가"))
        {
            return;
        }

        var insertPath = CreateTempPdfPath("blank");
        try
        {
            _pdfService.CreateBlankA4Pdf(insertPath);
            await InsertGeneratedPdfPageAsync(insertPath, "빈 A4 페이지 추가");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "빈 페이지 추가 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            TryDeleteTempFile(insertPath);
        }
    }

    private async void OnPasteClipboardImageClick(object sender, RoutedEventArgs e)
    {
        if (!EnsurePdfLoaded("이미지 붙여넣기"))
        {
            return;
        }

        if (!Clipboard.ContainsImage())
        {
            MessageBox.Show(this, "클립보드에 이미지가 없습니다.", "이미지 붙여넣기", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var image = Clipboard.GetImage();
        if (image is null)
        {
            MessageBox.Show(this, "클립보드 이미지를 읽을 수 없습니다.", "이미지 붙여넣기", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var insertPath = CreateTempPdfPath("clipboard-image");
        try
        {
            var encodedImage = EncodeJpegForA4(image);
            _pdfService.CreateImageA4Pdf(
                insertPath,
                encodedImage.JpegBytes,
                encodedImage.Width,
                encodedImage.Height);
            await InsertGeneratedPdfPageAsync(insertPath, "클립보드 이미지 붙여넣기");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "이미지 붙여넣기 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            TryDeleteTempFile(insertPath);
        }
    }

    private async Task PastePagesOrImageAsync()
    {
        if (Clipboard.ContainsData(PageTransferClipboardFormat))
        {
            await PasteTransferredPagesFromClipboardAsync();
            return;
        }

        OnPasteClipboardImageClick(this, new RoutedEventArgs());
    }

    private async Task CopySelectedPagesToClipboardAsync(bool cut)
    {
        if (!EnsurePdfLoaded(cut ? "페이지 잘라내기" : "페이지 복사"))
        {
            return;
        }

        var pages = GetSelectedPageTransforms();
        if (pages.Count == 0)
        {
            MessageBox.Show(this, "복사할 썸네일 페이지를 선택하세요.", cut ? "페이지 잘라내기" : "페이지 복사", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (cut && pages.Count >= _pageOrder.Count)
        {
            MessageBox.Show(this, "모든 페이지는 잘라낼 수 없습니다. 복사를 사용하거나 새 PDF로 저장하세요.", "페이지 잘라내기", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var tempPath = CreateTempPdfPath(cut ? "cut-pages" : "copy-pages");
        try
        {
            await _pdfService.SaveTransformedPagesAsync(_currentPdfPath!, pages, tempPath, CancellationToken.None);
            var payload = new PageTransferPayload(
                tempPath,
                [new PdfPageTransform(1, 0)],
                cut);
            var data = new DataObject();
            data.SetData(PageTransferClipboardFormat, JsonSerializer.Serialize(payload, PageTransferJsonOptions));

            var fileDrop = new System.Collections.Specialized.StringCollection
            {
                tempPath
            };
            data.SetFileDropList(fileDrop);
            Clipboard.SetDataObject(data, copy: true);

            if (cut)
            {
                SendViewerCommand("deleteSelectedPages");
            }

            CurrentFileText.Text = cut
                ? $"{pages.Count}개 페이지를 잘라냈습니다."
                : $"{pages.Count}개 페이지를 복사했습니다.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, cut ? "페이지 잘라내기 실패" : "페이지 복사 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task PasteTransferredPagesFromClipboardAsync()
    {
        if (!EnsurePdfLoaded("페이지 붙여넣기"))
        {
            return;
        }

        try
        {
            var json = Clipboard.GetData(PageTransferClipboardFormat) as string;
            var payload = string.IsNullOrWhiteSpace(json)
                ? null
                : JsonSerializer.Deserialize<PageTransferPayload>(json, PageTransferJsonOptions);
            if (payload is null || !File.Exists(payload.SourcePath))
            {
                MessageBox.Show(this, "붙여넣을 페이지 데이터를 찾을 수 없습니다.", "페이지 붙여넣기", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await InsertPreparedPdfPagesAsync(payload.SourcePath, GetInsertionIndexAfterSelection(), "페이지 붙여넣기", showMessage: false);
            CurrentFileText.Text = payload.Cut ? "페이지를 이동했습니다." : "페이지를 붙여넣었습니다.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "페이지 붙여넣기 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task InsertExternalPagesAsync(JsonElement root)
    {
        if (!EnsurePdfLoaded("페이지 드롭"))
        {
            return;
        }

        try
        {
            var message = JsonSerializer.Deserialize<ExternalPagesDropMessage>(root.GetRawText(), PageTransferJsonOptions)
                ?? throw new InvalidOperationException("드롭한 페이지 정보를 읽을 수 없습니다.");
            if (!File.Exists(message.SourcePath))
            {
                throw new FileNotFoundException("원본 PDF 파일을 찾을 수 없습니다.", message.SourcePath);
            }

            if (message.Pages.Count == 0)
            {
                throw new InvalidOperationException("드롭한 페이지가 없습니다.");
            }

            var selectedPath = CreateTempPdfPath("drag-pages");
            try
            {
                await _pdfService.SaveTransformedPagesAsync(message.SourcePath, message.Pages, selectedPath, CancellationToken.None);
                await InsertPreparedPdfPagesAsync(
                    selectedPath,
                    Math.Clamp(message.InsertionIndex, 0, _pageOrder.Count),
                    "페이지 드롭",
                    showMessage: false);
                CurrentFileText.Text = $"{message.Pages.Count}개 페이지를 드롭해서 복사했습니다.";
            }
            finally
            {
                TryDeleteTempFile(selectedPath);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "페이지 드롭 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task InsertExternalFilesAsync(JsonElement root)
    {
        if (!EnsurePdfLoaded("파일 삽입"))
        {
            return;
        }

        var tempPaths = new List<string>();
        try
        {
            var message = JsonSerializer.Deserialize<ExternalFilesDropMessage>(root.GetRawText(), PageTransferJsonOptions)
                ?? throw new InvalidOperationException("드롭한 파일 정보를 읽을 수 없습니다.");

            var paths = message.Paths
                .Where(File.Exists)
                .Where(IsSupportedInsertionFile)
                .Select(Path.GetFullPath)
                .ToList();

            if (paths.Count == 0)
            {
                throw new InvalidOperationException("삽입할 수 있는 PDF 또는 이미지 파일이 없습니다.");
            }

            var pdfParts = new List<string>();
            foreach (var path in paths)
            {
                if (IsPdfFile(path))
                {
                    pdfParts.Add(path);
                    continue;
                }

                var imagePdfPath = CreateTempPdfPath("dropped-image");
                var image = LoadBitmapFromFile(path);
                var encodedImage = EncodeJpegForA4(image);
                _pdfService.CreateImageA4Pdf(
                    imagePdfPath,
                    encodedImage.JpegBytes,
                    encodedImage.Width,
                    encodedImage.Height);
                tempPaths.Add(imagePdfPath);
                pdfParts.Add(imagePdfPath);
            }

            var insertPath = pdfParts.Count == 1
                ? pdfParts[0]
                : CreateTempPdfPath("dropped-files");

            if (pdfParts.Count > 1)
            {
                tempPaths.Add(insertPath);
                await _pdfService.CombinePdfFilesAsync(pdfParts, insertPath, CancellationToken.None);
            }

            await InsertPreparedPdfPagesAsync(
                insertPath,
                Math.Clamp(message.InsertionIndex, 0, _pageOrder.Count),
                "파일 삽입",
                showMessage: false);
            CurrentFileText.Text = $"{paths.Count}개 파일을 현재 PDF에 삽입했습니다.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "파일 삽입 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            foreach (var tempPath in tempPaths)
            {
                TryDeleteTempFile(tempPath);
            }
        }
    }

    private async void OnFitAllPagesToA4Click(object sender, RoutedEventArgs e)
    {
        if (!EnsurePdfLoaded("A4 맞춤"))
        {
            return;
        }

        var referencePath = _referencePdfPath ?? _currentPdfPath!;
        var folder = Path.GetDirectoryName(referencePath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var defaultPath = Path.Combine(folder, $"{Path.GetFileNameWithoutExtension(referencePath)}{_settings.A4FitSuffix}.pdf");
        var dialog = SavePathPromptService.CreateSaveDialog(defaultPath, "모든 페이지를 A4로 맞춤");

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var outputPath = SavePathPromptService.ResolveOutputPath(this, dialog.FileName, "A4 맞춤");
        if (outputPath is null)
        {
            return;
        }

        try
        {
            ViewerLoading.Visibility = Visibility.Visible;
            CurrentFileText.Text = "A4 페이지로 변환 중입니다...";
            var images = await ExportCurrentPagesAsA4ImagesAsync();
            var result = await _pdfService.CreateA4ImagePagesPdfAsync(images, outputPath, CancellationToken.None);
            LoadPdf(result.OutputPath);
            MessageBox.Show(this, $"저장 완료:\n{result.OutputPath}", "A4 맞춤", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ViewerLoading.Visibility = Visibility.Collapsed;
            CurrentFileText.Text = _currentPdfPath ?? "PDF를 열거나 이 창에 끌어다 놓으세요.";
            MessageBox.Show(this, ex.Message, "A4 맞춤 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void OnOptimizeA4FileSizeClick(object sender, RoutedEventArgs e)
    {
        if (!EnsurePdfLoaded("A4 기준 용량 최적화"))
        {
            return;
        }

        var referencePath = _referencePdfPath ?? _currentPdfPath!;
        var folder = Path.GetDirectoryName(referencePath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var defaultPath = Path.Combine(folder, $"{Path.GetFileNameWithoutExtension(referencePath)}{_settings.A4OptimizedSuffix}.pdf");
        var dialog = SavePathPromptService.CreateSaveDialog(defaultPath, "A4 기준 용량 최적화");

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var outputPath = SavePathPromptService.ResolveOutputPath(this, dialog.FileName, "A4 기준 용량 최적화");
        if (outputPath is null)
        {
            return;
        }

        try
        {
            ViewerLoading.Visibility = Visibility.Visible;
            CurrentFileText.Text = "A4 기준으로 용량 최적화 중입니다...";
            var originalSize = File.Exists(_currentPdfPath!) ? new FileInfo(_currentPdfPath!).Length : 0;
            var images = await ExportCurrentPagesAsA4ImagesAsync(optimizeSize: true);
            var result = await _pdfService.CreateA4ImagePagesPdfAsync(images, outputPath, CancellationToken.None);
            LoadPdf(result.OutputPath);
            var outputSize = new FileInfo(result.OutputPath).Length;
            var sizeMessage = originalSize > 0
                ? $"\n원본: {FormatFileSize(originalSize)}\n결과: {FormatFileSize(outputSize)}"
                : $"\n결과: {FormatFileSize(outputSize)}";
            MessageBox.Show(this, $"저장 완료:\n{result.OutputPath}{sizeMessage}", "A4 기준 용량 최적화", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ViewerLoading.Visibility = Visibility.Collapsed;
            CurrentFileText.Text = _currentPdfPath ?? "PDF를 열거나 이 창에 끌어다 놓으세요.";
            MessageBox.Show(this, ex.Message, "A4 기준 용량 최적화 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task InsertGeneratedPdfPageAsync(string insertPath, string title)
    {
        var insertionIndex = GetInsertionIndexAfterSelection();
        await InsertPreparedPdfPagesAsync(insertPath, insertionIndex, title, showMessage: true);
    }

    private async Task InsertPreparedPdfPagesAsync(string insertPath, int insertionIndex, string title, bool showMessage)
    {
        var outputPath = CreateTempPdfPath("edited");
        var referencePath = _referencePdfPath ?? _currentPdfPath!;

        await _pdfService.InsertPdfPagesAsync(
            _currentPdfPath!,
            GetCurrentPageTransforms(),
            insertPath,
            insertionIndex,
            outputPath,
            CancellationToken.None);

        LoadPdf(outputPath, referencePath, true);
        if (showMessage)
        {
            MessageBox.Show(this, $"{insertionIndex + 1}번째 위치에 추가했습니다.", title, MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private int GetInsertionIndexAfterSelection()
    {
        if (_pageOrder.Count == 0)
        {
            return 0;
        }

        HashSet<int> anchorPages;
        if (_selectedPages.Count > 0)
        {
            anchorPages = _selectedPages.ToHashSet();
        }
        else if (_activePage is { } activePage)
        {
            anchorPages = [activePage];
        }
        else
        {
            anchorPages = [];
        }

        if (anchorPages.Count == 0)
        {
            return _pageOrder.Count;
        }

        var insertionIndex = _pageOrder
            .Select((page, index) => new { Page = page, PositionAfter = index + 1 })
            .Where(item => anchorPages.Contains(item.Page))
            .Select(item => item.PositionAfter)
            .DefaultIfEmpty(_pageOrder.Count)
            .Max();

        return insertionIndex;
    }

    private static EncodedJpegImage EncodeJpegForA4(BitmapSource source)
    {
        var scale = Math.Min(
            1d,
            Math.Min(
                (double)A4ImageMaxWidthPixels / source.PixelWidth,
                (double)A4ImageMaxHeightPixels / source.PixelHeight));

        BitmapSource resizedSource = source;
        if (scale < 1d)
        {
            resizedSource = new TransformedBitmap(source, new ScaleTransform(scale, scale));
            resizedSource.Freeze();
        }

        BitmapSource frameSource = resizedSource.Format == PixelFormats.Bgr24
            ? resizedSource
            : new FormatConvertedBitmap(resizedSource, PixelFormats.Bgr24, null, 0);
        frameSource.Freeze();

        var encoder = new JpegBitmapEncoder
        {
            QualityLevel = 88
        };
        encoder.Frames.Add(BitmapFrame.Create(frameSource));
        using var stream = new MemoryStream();
        encoder.Save(stream);
        return new EncodedJpegImage(stream.ToArray(), frameSource.PixelWidth, frameSource.PixelHeight);
    }

    private static BitmapSource LoadBitmapFromFile(string path)
    {
        using var stream = File.OpenRead(path);
        var decoder = BitmapDecoder.Create(
            stream,
            BitmapCreateOptions.PreservePixelFormat,
            BitmapCacheOption.OnLoad);
        var frame = decoder.Frames.FirstOrDefault()
            ?? throw new InvalidOperationException("이미지 파일을 읽을 수 없습니다.");
        frame.Freeze();
        return frame;
    }

    private sealed record EncodedJpegImage(byte[] JpegBytes, int Width, int Height);

    private static string CreateTempPdfPath(string prefix)
    {
        return Path.Combine(Path.GetTempPath(), $"PdfMergeTool-{prefix}-{Guid.NewGuid():N}.pdf");
    }

    private static string FormatFileSize(long bytes)
    {
        var mb = bytes / 1024d / 1024d;
        return mb >= 1
            ? $"{mb:0.##} MB"
            : $"{bytes / 1024d:0.#} KB";
    }

    private static void TryDeleteTempFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Temporary files are best-effort cleanup.
        }
    }

    private bool EnsurePdfLoaded(string title)
    {
        if (!string.IsNullOrWhiteSpace(_currentPdfPath) && _pageOrder.Count > 0)
        {
            return true;
        }

        MessageBox.Show(this, "먼저 PDF를 열어주세요.", title, MessageBoxButton.OK, MessageBoxImage.Information);
        return false;
    }

    private IReadOnlyList<PdfPageTransform> GetCurrentPageTransforms()
    {
        return _pageOrder
            .Select(page => new PdfPageTransform(page, GetRotation(page)))
            .ToList();
    }

    private List<PdfPageTransform> GetSelectedPageTransforms()
    {
        if (_selectedPages.Count > 0)
        {
            return _pageOrder
                .Where(page => _selectedPages.Contains(page))
                .Select(page => new PdfPageTransform(page, GetRotation(page)))
                .ToList();
        }

        if (_activePage is { } activePage && _pageOrder.Contains(activePage))
        {
            return [new PdfPageTransform(activePage, GetRotation(activePage))];
        }

        return [];
    }

    private int GetRotation(int pageNumber)
    {
        return _pageRotations.TryGetValue(pageNumber, out var rotation) ? rotation : 0;
    }

    private void OnExitClick(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private async void OnPrintClick(object sender, RoutedEventArgs e)
    {
        if (!_viewerReady || PdfViewer.CoreWebView2 is null || string.IsNullOrWhiteSpace(_currentPdfPath))
        {
            MessageBox.Show(this, "먼저 PDF를 열어주세요.", "PDF 인쇄", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string? printPdfPath = null;
        try
        {
            var printDialog = new System.Windows.Controls.PrintDialog
            {
                UserPageRangeEnabled = true
            };

            if (printDialog.ShowDialog() != true)
            {
                return;
            }

            printPdfPath = await CreatePrintPdfAsync(CancellationToken.None);
            var status = await PrintPdfWithNativeViewerAsync(printPdfPath, printDialog);
            if (status != CoreWebView2PrintStatus.Succeeded)
            {
                MessageBox.Show(this, status.ToString(), "PDF 인쇄 실패", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, "PDF 인쇄 실패");
            MessageBox.Show(this, ex.Message, "PDF 인쇄 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _printReadyCompletion = null;
            TryDeletePrintPdf(printPdfPath);
        }
    }

    private async Task<string> CreatePrintPdfAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_currentPdfPath) || !File.Exists(_currentPdfPath))
        {
            throw new FileNotFoundException("인쇄할 PDF 파일을 찾을 수 없습니다.", _currentPdfPath);
        }

        if (_pageOrder.Count == 0)
        {
            throw new InvalidOperationException("인쇄할 페이지가 없습니다.");
        }

        var printPdfPath = Path.Combine(
            Path.GetTempPath(),
            $"PdfMergeTool-print-{Guid.NewGuid():N}.pdf");
        var pages = _pageOrder
            .Select(page => new PdfPageTransform(page, GetRotation(page)))
            .ToList();

        await _pdfService.SaveTransformedPagesAsync(_currentPdfPath, pages, printPdfPath, cancellationToken);
        return printPdfPath;
    }

    private async Task<CoreWebView2PrintStatus> PrintPdfWithNativeViewerAsync(
        string pdfPath,
        System.Windows.Controls.PrintDialog printDialog)
    {
        var printQueue = printDialog.PrintQueue;
        if (printQueue is null)
        {
            using var printServer = new LocalPrintServer();
            printQueue = printServer.DefaultPrintQueue;
        }

        if (printQueue is null || string.IsNullOrWhiteSpace(printQueue.FullName))
        {
            throw new InvalidOperationException("선택된 프린터를 확인할 수 없습니다.");
        }

        var printWindow = new Window
        {
            Width = 1,
            Height = 1,
            Left = -32000,
            Top = -32000,
            WindowStyle = WindowStyle.None,
            ResizeMode = ResizeMode.NoResize,
            ShowInTaskbar = false,
            ShowActivated = false,
            Opacity = 0,
            Content = new Microsoft.Web.WebView2.Wpf.WebView2()
        };

        var printViewer = (Microsoft.Web.WebView2.Wpf.WebView2)printWindow.Content;
        printWindow.Show();

        try
        {
            await printViewer.EnsureCoreWebView2Async(PdfViewer.CoreWebView2.Environment);
            printViewer.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            printViewer.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;

            await NavigateForPrintAsync(printViewer, pdfPath);

            var printSettings = printViewer.CoreWebView2.Environment.CreatePrintSettings();
            printSettings.PrinterName = printQueue.FullName;
            printSettings.ShouldPrintHeaderAndFooter = false;
            printSettings.HeaderTitle = string.Empty;
            printSettings.FooterUri = string.Empty;
            printSettings.ShouldPrintBackgrounds = true;
            printSettings.MarginTop = 0;
            printSettings.MarginBottom = 0;
            printSettings.MarginLeft = 0;
            printSettings.MarginRight = 0;

            var printTicket = printDialog.PrintTicket;
            if (printTicket?.CopyCount is { } copyCount && copyCount > 0)
            {
                printSettings.Copies = copyCount;
            }

            if (printTicket?.PageOrientation == PageOrientation.Landscape ||
                printTicket?.PageOrientation == PageOrientation.ReverseLandscape)
            {
                printSettings.Orientation = CoreWebView2PrintOrientation.Landscape;
            }
            else if (printTicket?.PageOrientation == PageOrientation.Portrait ||
                     printTicket?.PageOrientation == PageOrientation.ReversePortrait)
            {
                printSettings.Orientation = CoreWebView2PrintOrientation.Portrait;
            }

            if (printDialog.PageRangeSelection == System.Windows.Controls.PageRangeSelection.UserPages)
            {
                printSettings.PageRanges = $"{printDialog.PageRange.PageFrom}-{printDialog.PageRange.PageTo}";
            }

            return await printViewer.CoreWebView2.PrintAsync(printSettings);
        }
        finally
        {
            printWindow.Close();
        }
    }

    private static async Task NavigateForPrintAsync(Microsoft.Web.WebView2.Wpf.WebView2 printViewer, string pdfPath)
    {
        var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

        void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs args)
        {
            completion.TrySetResult(args.IsSuccess);
        }

        printViewer.CoreWebView2.NavigationCompleted += OnNavigationCompleted;
        try
        {
            printViewer.CoreWebView2.Navigate(new Uri(pdfPath).AbsoluteUri);
            var completed = await Task.WhenAny(completion.Task, Task.Delay(TimeSpan.FromSeconds(30)));
            if (completed != completion.Task || !await completion.Task)
            {
                throw new InvalidOperationException("인쇄용 PDF를 준비하지 못했습니다.");
            }

            await Task.Delay(700);
        }
        finally
        {
            printViewer.CoreWebView2.NavigationCompleted -= OnNavigationCompleted;
        }
    }

    private static void TryDeletePrintPdf(string? path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // The temporary print file is best-effort cleanup.
        }
    }

    private void OnMainZoomInClick(object sender, RoutedEventArgs e) => SendViewerCommand("mainZoomIn");

    private void OnMainZoomOutClick(object sender, RoutedEventArgs e) => SendViewerCommand("mainZoomOut");

    private void OnMainZoomResetClick(object sender, RoutedEventArgs e) => SendViewerCommand("mainZoomReset");

    private void OnFitPageClick(object sender, RoutedEventArgs e) => SendViewerCommand("fitPage");

    private void OnThumbZoomInClick(object sender, RoutedEventArgs e) => SendViewerCommand("thumbZoomIn");

    private void OnThumbZoomOutClick(object sender, RoutedEventArgs e) => SendViewerCommand("thumbZoomOut");

    private void OnThumbZoomResetClick(object sender, RoutedEventArgs e) => SendViewerCommand("thumbZoomReset");

    private void OnPrevPageClick(object sender, RoutedEventArgs e) => SendViewerCommand("prevPage");

    private void OnNextPageClick(object sender, RoutedEventArgs e) => SendViewerCommand("nextPage");

    private void OnFirstPageClick(object sender, RoutedEventArgs e) => SendViewerCommand("firstPage");

    private void OnLastPageClick(object sender, RoutedEventArgs e) => SendViewerCommand("lastPage");

    private void OnUndoClick(object sender, RoutedEventArgs e) => SendViewerCommand("undo");

    private void OnRedoClick(object sender, RoutedEventArgs e) => SendViewerCommand("redo");

    private async void OnCopySelectedPagesClick(object sender, RoutedEventArgs e) => await CopySelectedPagesToClipboardAsync(cut: false);

    private async void OnCutSelectedPagesClick(object sender, RoutedEventArgs e) => await CopySelectedPagesToClipboardAsync(cut: true);

    private async void OnPastePagesClick(object sender, RoutedEventArgs e) => await PastePagesOrImageAsync();

    private void OnDeleteSelectedPagesClick(object sender, RoutedEventArgs e) => SendViewerCommand("deleteSelectedPages");

    private void OnRotateClockwiseClick(object sender, RoutedEventArgs e) => SendViewerCommand("rotateSelectedClockwise");

    private void OnRotateCounterClockwiseClick(object sender, RoutedEventArgs e) => SendViewerCommand("rotateSelectedCounterClockwise");

    private void OnReversePageOrderClick(object sender, RoutedEventArgs e) => SendViewerCommand("reversePageOrder");

    private void OnWindowKeyDown(object sender, KeyEventArgs e)
    {
        var modifiers = Keyboard.Modifiers;
        var control = modifiers.HasFlag(ModifierKeys.Control);
        var shift = modifiers.HasFlag(ModifierKeys.Shift);

        if (control && !shift && e.Key == Key.O)
        {
            OnOpenPdfClick(sender, e);
            e.Handled = true;
        }
        else if (control && !shift && e.Key == Key.M)
        {
            OnOpenMergeClick(sender, e);
            e.Handled = true;
        }
        else if (e.Key == Key.F5)
        {
            OnRefreshViewerClick(sender, e);
            e.Handled = true;
        }
        else if (control && !shift && e.Key == Key.S)
        {
            OnSavePageOrderClick(sender, e);
            e.Handled = true;
        }
        else if (control && !shift && e.Key == Key.P)
        {
            OnPrintClick(sender, e);
            e.Handled = true;
        }
        else if (control && !shift && e.Key == Key.Z)
        {
            SendViewerCommand("undo");
            e.Handled = true;
        }
        else if (control && !shift && e.Key == Key.Y)
        {
            SendViewerCommand("redo");
            e.Handled = true;
        }
        else if (control && !shift && e.Key == Key.C)
        {
            _ = CopySelectedPagesToClipboardAsync(cut: false);
            e.Handled = true;
        }
        else if (control && !shift && e.Key == Key.X)
        {
            _ = CopySelectedPagesToClipboardAsync(cut: true);
            e.Handled = true;
        }
        else if (control && shift && e.Key == Key.N)
        {
            OnAddBlankA4PageClick(sender, e);
            e.Handled = true;
        }
        else if (control && shift && e.Key == Key.A)
        {
            OnFitAllPagesToA4Click(sender, e);
            e.Handled = true;
        }
        else if (control && shift && e.Key == Key.O)
        {
            OnOptimizeA4FileSizeClick(sender, e);
            e.Handled = true;
        }
        else if (control && !shift && e.Key == Key.V)
        {
            _ = PastePagesOrImageAsync();
            e.Handled = true;
        }
        else if (control && !shift && e.Key == Key.R)
        {
            SendViewerCommand("rotateSelectedClockwise");
            e.Handled = true;
        }
        else if (e.Key == Key.Delete)
        {
            SendViewerCommand("deleteSelectedPages");
            e.Handled = true;
        }
        else if (control && !shift && (e.Key == Key.OemPlus || e.Key == Key.Add))
        {
            SendViewerCommand("mainZoomIn");
            e.Handled = true;
        }
        else if (control && !shift && (e.Key == Key.OemMinus || e.Key == Key.Subtract))
        {
            SendViewerCommand("mainZoomOut");
            e.Handled = true;
        }
        else if (control && !shift && (e.Key == Key.D0 || e.Key == Key.NumPad0))
        {
            SendViewerCommand("mainZoomReset");
            e.Handled = true;
        }
        else if (control && !shift && (e.Key == Key.D1 || e.Key == Key.NumPad1))
        {
            SendViewerCommand("fitPage");
            e.Handled = true;
        }
        else if (control && shift && (e.Key == Key.OemPlus || e.Key == Key.Add))
        {
            SendViewerCommand("thumbZoomIn");
            e.Handled = true;
        }
        else if (control && shift && (e.Key == Key.OemMinus || e.Key == Key.Subtract))
        {
            SendViewerCommand("thumbZoomOut");
            e.Handled = true;
        }
        else if (control && shift && (e.Key == Key.D0 || e.Key == Key.NumPad0))
        {
            SendViewerCommand("thumbZoomReset");
            e.Handled = true;
        }
    }

    private void OnWindowDragOver(object sender, DragEventArgs e)
    {
        if (CanInsertDroppedFilesIntoCurrentDocument() &&
            TryGetDroppedInsertionPaths(e, out var insertionPaths))
        {
            e.Effects = DragDropEffects.Copy;
            SendNativeFileDropMessage("nativeFileDragOver", insertionPaths, PdfViewer.PointToScreen(e.GetPosition(PdfViewer)));
            e.Handled = true;
            return;
        }

        if (TryGetDroppedPdfPaths(e, out _))
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }
    }

    private void OnWindowDrop(object sender, DragEventArgs e)
    {
        if (CanInsertDroppedFilesIntoCurrentDocument() &&
            TryGetDroppedInsertionPaths(e, out var insertionPaths))
        {
            SendNativeFileDropMessage("nativeFileDrop", insertionPaths, PdfViewer.PointToScreen(e.GetPosition(PdfViewer)));
            e.Handled = true;
            return;
        }

        if (TryGetDroppedPdfPaths(e, out var paths))
        {
            OpenFiles(paths);
            e.Handled = true;
        }
    }

    private static bool TryGetDroppedPdfPaths(DragEventArgs e, out string[] paths)
    {
        paths = [];

        if (!e.Data.GetDataPresent(DataFormats.FileDrop) ||
            e.Data.GetData(DataFormats.FileDrop) is not string[] droppedPaths)
        {
            return false;
        }

        paths = droppedPaths
            .Where(path => File.Exists(path))
            .Where(IsPdfFile)
            .ToArray();

        return paths.Length > 0;
    }

    private static bool TryGetDroppedInsertionPaths(DragEventArgs e, out string[] paths)
    {
        paths = [];

        if (!e.Data.GetDataPresent(DataFormats.FileDrop) ||
            e.Data.GetData(DataFormats.FileDrop) is not string[] droppedPaths)
        {
            return false;
        }

        paths = droppedPaths
            .Where(path => File.Exists(path))
            .Where(IsSupportedInsertionFile)
            .ToArray();

        return paths.Length > 0;
    }
}
