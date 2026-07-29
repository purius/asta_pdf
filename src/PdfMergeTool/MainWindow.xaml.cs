using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Printing;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using PdfMergeTool.Services;

namespace PdfMergeTool;

public partial class MainWindow : Window
{
    private const string ViewerHost = "pdfviewer.local";
    private const string ViewerCacheHost = "pdfcache.local";
    private const string ViewerAssetFolderName = "PdfViewerOfficial";
    private const int A4ImageMaxWidthPixels = 2480;
    private const int A4ImageMaxHeightPixels = 3508;
    private const int A4OptimizedMaxWidthPixels = 2200;
    private const int A4OptimizedMaxHeightPixels = 3112;
    private const string PageTransferClipboardFormat = "PdfMergeTool.Pages.v1";
    private const string PageOrganizerDragDataFormat = "PdfMergeTool.PageOrganizerPage.v1";
    private const double DefaultPageOrganizerThumbnailHeight = 142;
    private const double MinimumPageOrganizerThumbnailHeight = 106;
    private const double MaximumPageOrganizerThumbnailHeight = 238;
    private const double PageOrganizerThumbnailWidthRatio = 0.7071067811865476;
    private const double PageOrganizerThumbnailCardExtraWidth = 16;
    private static readonly JsonSerializerOptions PageTransferJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly AppSettings _settings = AppSettings.Load();
    private readonly PdfMergeService _pdfService = new();
    private readonly PdfFallbackRenderService _fallbackRenderService = new();
    private readonly PdfFallbackRenderService _pageOrganizerRenderService = new();
    private readonly DocumentOperationCoordinator _documentOperations = new();
    private bool _viewerReady;
    private string? _currentPdfPath;
    private string? _referencePdfPath;
    private string? _workingSaveTargetPath;
    private string? _pendingPdfPath;
    private string? _pendingReferencePdfPath;
    private bool _pendingDirtyAfterLoad;
    private int? _pendingInitialPage;
    private int _pendingLoadGeneration;
    private IReadOnlyList<int> _pageOrder = [];
    private IReadOnlyDictionary<int, int> _pageRotations = new Dictionary<int, int>();
    private IReadOnlyList<int> _selectedPages = [];
    private int? _activePage;
    private bool _isDirty;
    private bool _editorDirty;
    private EditorDocumentState? _pageOrganizerState;
    private CancellationTokenSource? _pageOrganizerThumbnailCancellation;
    private PageOrganizerThumbnailScheduler? _pageOrganizerThumbnailScheduler;
    private HashSet<int> _pageOrganizerThumbnailCacheWindow = [];
    private string? _pageOrganizerThumbnailSourcePath;
    private int? _pageOrganizerThumbnailWorkerGeneration;
    private int _pageOrganizerThumbnailGeneration;
    private int _documentLoadGeneration;
    private int? _documentMutationGeneration;
    private int? _pageOrganizerDragPageNumber;
    private Point? _pageOrganizerDragStartPosition;
    private int? _pageOrganizerPendingPlainSelectionPageNumber;
    private int? _pageOrganizerDropInsertionIndex;
    private int? _pendingPageOrganizerFollowPage;
    private int _pendingPageOrganizerFollowLoadGeneration;
    private int _pageOrganizerFollowRequestId;
    private bool _pageOrganizerFollowQueued;
    private bool _isPageOrganizerDragInProgress;
    private double _pageOrganizerThumbnailHeight = DefaultPageOrganizerThumbnailHeight;
    private MergeWindow? _mergeWindow;
    private TaskCompletionSource<bool>? _printReadyCompletion;
    private readonly List<NativeFileDropTarget> _viewerDropTargets = [];
    private int? _lastLoggedPageCount;
    private string? _servedPdfLinkPath;
    private string? _recoveredPdfPath;
    private bool _viewerLoadRecoveryAttempted;
    private bool _fallbackModeActive;
    private string? _fallbackSessionId;
    private TaskCompletionSource<EditorExportState>? _editorStateCompletion;
    private string? _editorStateRequestId;
    private TaskCompletionSource<string>? _overlayPdfExportCompletion;
    private string? _overlayPdfExportRequestId;

    public ObservableCollection<PageOrganizerItem> PageOrganizerItems { get; } = [];

    public MainWindow(IEnumerable<string> initialFiles, bool openMergeWindow)
    {
        InitializeComponent();
        ApplyWindowSettings();
        PageOrganizerList.AddHandler(
            ScrollViewer.ScrollChangedEvent,
            new ScrollChangedEventHandler(OnPageOrganizerThumbnailScrollChanged));
        AddHandler(DragDrop.PreviewDragOverEvent, new DragEventHandler(OnWindowDragOver), true);
        AddHandler(DragDrop.PreviewDropEvent, new DragEventHandler(OnWindowDrop), true);
        Loaded += OnLoaded;
        Closing += OnClosing;
        OpenFiles(initialFiles, openMergeWindow);
    }

    private sealed record PageTransferPayload(string SourcePath, List<PdfPageTransform> Pages, bool Cut);
    private sealed record ExternalPagesDropMessage(string SourcePath, List<PdfPageTransform> Pages, int InsertionIndex);
    private sealed record ExternalFilesDropMessage(List<string> Paths, int InsertionIndex);
    private sealed record EditorExportState(IReadOnlyList<JsonElement> Edits, IReadOnlyList<string> UsedFontNames, bool IsDirty);

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

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _ = Dispatcher.BeginInvoke(
            new Action(async () => await InitializeStartupUiAsync()),
            DispatcherPriority.ApplicationIdle);
    }

    private async Task InitializeStartupUiAsync()
    {
        try
        {
            UpdatePdfContextMenuOption();
            UpdateRecentFilesMenu();
            await InitializeViewerAsync();
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, "Viewer startup initialization failed.");
            MessageBox.Show(this, ex.Message, "PDF 뷰어 초기화 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task InitializeViewerAsync()
    {
        await PdfViewer.EnsureCoreWebView2Async();
        PdfViewer.DefaultBackgroundColor = System.Drawing.Color.White;
        PdfViewer.AllowExternalDrop = false;
        PdfViewer.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        PdfViewer.CoreWebView2.Settings.IsZoomControlEnabled = false;
        PdfViewer.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = false;
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
                SendEditorFontsToViewer();
                if (string.IsNullOrWhiteSpace(_pendingPdfPath))
                {
                    return;
                }

                var pending = _pendingPdfPath;
                var pendingReference = _pendingReferencePdfPath;
                var pendingDirty = _pendingDirtyAfterLoad;
                var pendingInitialPage = _pendingInitialPage;
                var pendingLoadGeneration = _pendingLoadGeneration;
                _pendingPdfPath = null;
                _pendingReferencePdfPath = null;
                _pendingDirtyAfterLoad = false;
                _pendingInitialPage = null;
                _pendingLoadGeneration = 0;
                if (pendingLoadGeneration != _documentLoadGeneration)
                {
                    return;
                }

                _referencePdfPath = pendingReference ?? pending;
                await SendPdfToViewerAsync(pending, pendingLoadGeneration, pendingDirty, pendingInitialPage);
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
                if (!IsCurrentViewerLoadMessage(document.RootElement))
                {
                    return;
                }

                var message = document.RootElement.TryGetProperty("message", out var messageElement)
                    ? messageElement.GetString()
                    : "알 수 없는 오류";
                await HandleViewerLoadFailedAsync(message, _documentLoadGeneration);
                return;
            }

            if (type == "viewerFirstPageRendered")
            {
                if (IsCurrentViewerLoadMessage(document.RootElement))
                {
                    ViewerLoading.Visibility = Visibility.Collapsed;
                }

                return;
            }

            if (type == "editorStateChanged")
            {
                if (!IsCurrentViewerLoadMessage(document.RootElement))
                {
                    return;
                }

                _editorDirty = document.RootElement.TryGetProperty("isDirty", out var dirtyElement) &&
                               dirtyElement.GetBoolean();
                RefreshDirtyState();
                return;
            }

            if (type == "editorStateCollected")
            {
                if (IsCurrentViewerLoadMessage(document.RootElement))
                {
                    CompleteEditorStateCollection(document.RootElement);
                }

                return;
            }

            if (type == "overlayPdfExported")
            {
                if (IsCurrentViewerLoadMessage(document.RootElement))
                {
                    CompleteOverlayPdfExport(document.RootElement);
                }

                return;
            }

            if (type == "overlayPdfExportFailed")
            {
                if (IsCurrentViewerLoadMessage(document.RootElement))
                {
                    CompleteOverlayPdfExportFailure(document.RootElement);
                }

                return;
            }

            if (type == "activePageChanged")
            {
                if (IsCurrentViewerLoadMessage(document.RootElement))
                {
                    ReceiveActivePageChanged(document.RootElement);
                }

                return;
            }

        };

        var viewerFolder = Path.Combine(AppContext.BaseDirectory, "Assets", ViewerAssetFolderName);
        PdfViewer.CoreWebView2.SetVirtualHostNameToFolderMapping(
            ViewerHost,
            viewerFolder,
            CoreWebView2HostResourceAccessKind.Allow);
        PdfViewer.CoreWebView2.SetVirtualHostNameToFolderMapping(
            ViewerCacheHost,
            AppPaths.ViewerRuntimeDirectory,
            CoreWebView2HostResourceAccessKind.Allow);
        PdfViewer.CoreWebView2.Navigate(BuildViewerUrl());
    }

    private string BuildViewerUrl()
    {
        var optimizedPartialRendering = _settings.EnableOptimizedPartialRendering ? "true" : "false";
        return $"https://{ViewerHost}/web/viewer.html?enableoptimizedpartialrendering={optimizedPartialRendering}";
    }

    private void ReceiveActivePageChanged(JsonElement root)
    {
        if (_pageOrganizerState is null ||
            !root.TryGetProperty("activePage", out var activeElement) ||
            activeElement.ValueKind != JsonValueKind.Number)
        {
            return;
        }

        var activePage = activeElement.GetInt32();
        if (_pageOrganizerState.PageNumbers.Contains(activePage))
        {
            var nextState = _pageOrganizerState.ActivatePage(activePage);
            ApplyPageOrganizerState(nextState);
            QueueActivePageFollow(activePage);
        }
    }

    private void QueueActivePageFollow(int pageNumber)
    {
        if (IsActivePageFollowSuspended())
        {
            return;
        }

        _pendingPageOrganizerFollowPage = pageNumber;
        _pendingPageOrganizerFollowLoadGeneration = _documentLoadGeneration;
        if (_pageOrganizerFollowQueued)
        {
            return;
        }

        _pageOrganizerFollowQueued = true;
        var requestId = ++_pageOrganizerFollowRequestId;
        _ = Dispatcher.BeginInvoke(new Action(() =>
        {
            if (requestId != _pageOrganizerFollowRequestId)
            {
                return;
            }

            _pageOrganizerFollowQueued = false;
            var page = _pendingPageOrganizerFollowPage;
            var loadGeneration = _pendingPageOrganizerFollowLoadGeneration;
            _pendingPageOrganizerFollowPage = null;
            if (page is null || loadGeneration != _documentLoadGeneration || IsActivePageFollowSuspended())
            {
                return;
            }

            FollowActivePageOrganizerItem(page.Value);
        }), DispatcherPriority.Loaded);
    }

    private void ResetActivePageFollow()
    {
        _pendingPageOrganizerFollowPage = null;
        _pendingPageOrganizerFollowLoadGeneration = 0;
        _pageOrganizerFollowQueued = false;
        _pageOrganizerFollowRequestId++;
    }

    private bool IsActivePageFollowSuspended() =>
        IsDocumentMutationInProgress || _isPageOrganizerDragInProgress;

    private void FollowActivePageOrganizerItem(int pageNumber)
    {
        if (_pageOrganizerState is null)
        {
            return;
        }

        var index = _pageOrganizerState.PageNumbers.ToList().IndexOf(pageNumber);
        if (index < 0 || PageOrganizerList.ItemContainerGenerator.ContainerFromIndex(index) is not FrameworkElement container)
        {
            return;
        }

        var scrollViewer = FindVisualDescendant<ScrollViewer>(PageOrganizerList);
        if (scrollViewer is null || scrollViewer.ViewportHeight <= 0 || container.ActualHeight <= 0)
        {
            return;
        }

        try
        {
            var itemTop = container.TranslatePoint(new Point(), scrollViewer).Y;
            var targetOffset = PageOrganizerViewport.GetVerticalOffsetToReveal(
                scrollViewer.VerticalOffset,
                scrollViewer.ViewportHeight,
                itemTop,
                itemTop + container.ActualHeight,
                scrollViewer.ScrollableHeight);
            if (targetOffset is { } offset)
            {
                scrollViewer.ScrollToVerticalOffset(offset);
            }
        }
        catch (InvalidOperationException)
        {
            // The item can be regenerated while the PDF viewer is changing pages.
        }
    }

    private bool IsCurrentViewerLoadMessage(JsonElement root)
    {
        return root.TryGetProperty("loadId", out var loadIdElement) &&
               loadIdElement.ValueKind == JsonValueKind.Number &&
               loadIdElement.TryGetInt32(out var loadId) &&
               loadId == _documentLoadGeneration;
    }

    private async Task ExecuteDocumentMutationAsync(
        DocumentOperationToken operation,
        Func<DocumentOperationToken, Task> mutation)
    {
        using var lease = await _documentOperations.EnterMutationAsync(operation);
        SetDocumentMutationUiState(operation, isBusy: true);
        try
        {
            _documentOperations.ThrowIfSuperseded(operation);
            await mutation(operation);
        }
        finally
        {
            SetDocumentMutationUiState(operation, isBusy: false);
        }
    }

    private async Task RunCurrentDocumentMutationAsync(
        string operationName,
        string errorTitle,
        Func<DocumentOperationToken, Task> mutation)
    {
        if (!TryCaptureDocumentOperation(out var operation))
        {
            return;
        }
        try
        {
            await ExecuteDocumentMutationAsync(operation, mutation);
        }
        catch (OperationCanceledException) when (IsDocumentOperationSuperseded(operation))
        {
            AppLogger.Info($"Ignored {operationName} after the active document changed.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, errorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private DocumentOperationToken CaptureDocumentOperation()
    {
        return _documentOperations.Capture();
    }

    private bool TryCaptureDocumentOperation(out DocumentOperationToken operation)
    {
        if (IsDocumentMutationInProgress)
        {
            operation = default;
            return false;
        }

        operation = CaptureDocumentOperation();
        return true;
    }

    private bool IsDocumentMutationInProgress => _documentMutationGeneration is not null;

    private void SetDocumentMutationUiState(DocumentOperationToken operation, bool isBusy)
    {
        if (isBusy)
        {
            CancelPageOrganizerThumbnailRendering();
            _documentMutationGeneration = operation.Generation;
            PageOrganizerList.IsHitTestVisible = false;
            PdfViewer.IsHitTestVisible = false;
            return;
        }

        if (_documentMutationGeneration != operation.Generation)
        {
            return;
        }

        _documentMutationGeneration = null;
        PageOrganizerList.IsHitTestVisible = true;
        PdfViewer.IsHitTestVisible = true;
        ResumePageOrganizerThumbnailRenderingAfterMutation();
    }

    private bool IsDocumentOperationSuperseded(DocumentOperationToken operation)
    {
        return !_documentOperations.IsCurrent(operation);
    }

    private void CancelPendingViewerDocumentOperations()
    {
        _editorStateCompletion?.TrySetCanceled();
        _overlayPdfExportCompletion?.TrySetCanceled();
    }

    private async void LoadPdf(
        string path,
        string? referencePath = null,
        bool dirtyAfterLoad = false,
        int? initialPage = null,
        bool preserveWorkingSaveTarget = false)
    {
        var loadGeneration = _documentOperations.StartNewDocument();
        _documentLoadGeneration = loadGeneration;
        ResetActivePageFollow();
        CancelPendingViewerDocumentOperations();
        CleanupCompatibilityState();
        _pendingPdfPath = null;
        _pendingReferencePdfPath = null;
        _pendingDirtyAfterLoad = false;
        _pendingInitialPage = null;
        _pendingLoadGeneration = 0;
        if (!preserveWorkingSaveTarget)
        {
            _workingSaveTargetPath = null;
        }

        _currentPdfPath = path;
        _referencePdfPath = referencePath ?? path;
        _pageOrder = [];
        _pageRotations = new Dictionary<int, int>();
        _selectedPages = [];
        _activePage = null;
        _editorDirty = false;
        _isDirty = dirtyAfterLoad;
        _pageOrganizerState = null;
        _lastLoggedPageCount = null;
        CurrentFileText.Text = IsSamePath(path, _referencePdfPath)
            ? path
            : $"{_referencePdfPath} (편집 중)";
        ViewerLoading.Visibility = Visibility.Visible;
        await InitializePageOrganizerStateAsync(path, dirtyAfterLoad, initialPage, loadGeneration);
        if (loadGeneration != _documentLoadGeneration)
        {
            return;
        }

        RefreshDirtyState();
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
            _pendingInitialPage = initialPage;
            _pendingLoadGeneration = loadGeneration;
            return;
        }

        await SendPdfToViewerAsync(path, loadGeneration, dirtyAfterLoad, initialPage);
    }

    private async Task InitializePageOrganizerStateAsync(
        string path,
        bool dirtyAfterLoad,
        int? initialPage,
        int loadGeneration)
    {
        CancelPageOrganizerThumbnailRendering();
        PageOrganizerItems.Clear();

        try
        {
            var pageCount = await Task.Run(() => _pdfService.GetPageCount(path));
            if (loadGeneration != _documentLoadGeneration)
            {
                return;
            }

            var state = EditorDocumentState.Create(pageCount, dirtyAfterLoad);
            if (initialPage is { } pageNumber && state.PageNumbers.Contains(pageNumber))
            {
                state = state.SelectPage(pageNumber, PageSelectionMode.Replace);
            }

            ApplyPageOrganizerState(state);
            _lastLoggedPageCount = pageCount;
            AppLogger.Info($"Page Organizer initialized: {pageCount} pages, {_referencePdfPath ?? path}");
            StartPageOrganizerThumbnailRendering(path, state.PageNumbers);
        }
        catch (Exception ex)
        {
            if (loadGeneration != _documentLoadGeneration)
            {
                return;
            }

            AppLogger.Error(ex, $"Page Organizer initialization failed: {path}");
            PageOrganizerSummaryText.Text = "페이지 정보를 읽는 중 오류가 발생했습니다.";
        }
    }

    private void ApplyPageOrganizerState(
        EditorDocumentState state,
        bool navigatePreview = false,
        bool allowDuringDocumentMutation = false)
    {
        _pageOrganizerState = state;
        _pageOrder = state.PageNumbers.ToArray();
        _pageRotations = state.PageRotations.ToDictionary(pair => pair.Key, pair => pair.Value);
        _selectedPages = state.SelectedPageNumbers.ToArray();
        _activePage = state.ActivePageNumber;
        RefreshPageOrganizerItems(state);
        RefreshDirtyState();

        if (navigatePreview && state.ActivePageNumber is { } activePage)
        {
            SendViewerCommand(
                "goToPage",
                new { pageNumber = activePage },
                allowDuringDocumentMutation);
        }
    }

    private void RefreshPageOrganizerItems(EditorDocumentState state)
    {
        var existing = PageOrganizerItems.ToDictionary(item => item.PageNumber);
        var orderChanged = PageOrganizerItems.Count != state.PageNumbers.Count ||
                           !PageOrganizerItems.Select(item => item.PageNumber).SequenceEqual(state.PageNumbers);
        var selected = state.SelectedPageNumbers.ToHashSet();
        var next = new List<PageOrganizerItem>(state.PageNumbers.Count);
        for (var index = 0; index < state.PageNumbers.Count; index++)
        {
            var pageNumber = state.PageNumbers[index];
            if (!existing.TryGetValue(pageNumber, out var item))
            {
                item = new PageOrganizerItem(pageNumber);
            }

            item.Position = index + 1;
            item.Rotation = state.GetRotation(pageNumber);
            item.IsSelected = selected.Contains(pageNumber);
            item.IsActive = state.ActivePageNumber == pageNumber;
            ApplyPageOrganizerThumbnailDimensions(item);
            next.Add(item);
        }

        if (orderChanged)
        {
            PageOrganizerItems.Clear();
            foreach (var item in next)
            {
                PageOrganizerItems.Add(item);
            }
        }

        PageOrganizerSummaryText.Text = state.PageNumbers.Count == 0
            ? "페이지 없음"
            : $"{state.PageNumbers.Count} 페이지 · {state.SelectedPageNumbers.Count} 선택";

        try
        {
            PageOrganizerList.SelectedItems.Clear();
            foreach (var item in next.Where(item => item.IsSelected))
            {
                PageOrganizerList.SelectedItems.Add(item);
            }
        }
        catch (InvalidOperationException)
        {
            // The organizer can refresh before its ListBox finishes item generation.
        }
    }

    private void RefreshDirtyState()
    {
        if (_pageOrganizerState is not null)
        {
            _isDirty = _pageOrganizerState.IsDirty || _editorDirty;
        }
        else
        {
            _isDirty = _isDirty || _editorDirty;
        }

        UpdateWindowTitle();
    }

    private void StartPageOrganizerThumbnailRendering(string path, IReadOnlyList<int> pageNumbers)
    {
        CancelPageOrganizerThumbnailRendering();
        var cancellation = new CancellationTokenSource();
        _pageOrganizerThumbnailCancellation = cancellation;
        _pageOrganizerThumbnailSourcePath = path;
        var scheduler = new PageOrganizerThumbnailScheduler(pageNumbers);
        _pageOrganizerThumbnailScheduler = scheduler;
        _pageOrganizerThumbnailCacheWindow = scheduler.GetCacheWindow([]).ToHashSet();
        var generation = ++_pageOrganizerThumbnailGeneration;

        foreach (var item in PageOrganizerItems)
        {
            if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation))
            {
                return;
            }

            item.Thumbnail = null;
            item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Pending;
        }

        scheduler.Prioritize(_pageOrganizerThumbnailCacheWindow);
        EnsurePageOrganizerThumbnailWorker(generation, cancellation);
        _ = Dispatcher.BeginInvoke(
            new Action(() =>
            {
                if (IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) &&
                    ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
                {
                    RefreshPageOrganizerThumbnailViewport();
                }
            }),
            DispatcherPriority.Loaded);
    }

    private bool IsCurrentPageOrganizerThumbnailRequest(
        int generation,
        CancellationTokenSource cancellation)
    {
        return generation == _pageOrganizerThumbnailGeneration &&
               ReferenceEquals(_pageOrganizerThumbnailCancellation, cancellation) &&
               !cancellation.IsCancellationRequested;
    }

    private void EnsurePageOrganizerThumbnailWorker(
        int generation,
        CancellationTokenSource cancellation)
    {
        var path = _pageOrganizerThumbnailSourcePath;
        if (_pageOrganizerThumbnailWorkerGeneration == generation ||
            !IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
            _pageOrganizerThumbnailScheduler is null ||
            string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        _pageOrganizerThumbnailWorkerGeneration = generation;
        _ = RenderPageOrganizerThumbnailsAsync(path, generation, cancellation);
    }

    private async Task RenderPageOrganizerThumbnailsAsync(
        string path,
        int generation,
        CancellationTokenSource cancellation)
    {
        string? sessionId = null;
        var cancellationToken = cancellation.Token;
        try
        {
            var session = await Task.Run(() => _pageOrganizerRenderService.OpenDocument(path), cancellationToken);
            sessionId = session.SessionId;

            while (IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) &&
                   _pageOrganizerThumbnailScheduler is { } scheduler &&
                   scheduler.TryTakeNext(out var pageNumber))
            {
                if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
                    !ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
                {
                    return;
                }

                var item = PageOrganizerItems.FirstOrDefault(candidate => candidate.PageNumber == pageNumber);
                if (item is null)
                {
                    scheduler.Complete(pageNumber);
                    continue;
                }

                if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
                    !ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
                {
                    return;
                }

                item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Loading;
                try
                {
                    var rendered = await _pageOrganizerRenderService.RenderPageAsync(
                        sessionId,
                        pageNumber,
                        targetWidth: 132,
                        rotationDegrees: 0,
                        thumbnail: true,
                        cancellationToken: cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
                        !ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
                    {
                        return;
                    }

                    if (_pageOrganizerThumbnailCacheWindow.Contains(pageNumber))
                    {
                        var thumbnail = LoadPageOrganizerThumbnail(rendered.ImagePath);
                        if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
                            !ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
                        {
                            return;
                        }

                        item.Thumbnail = thumbnail;
                        item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Ready;
                    }
                    else
                    {
                        item.Thumbnail = null;
                        item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Evicted;
                    }

                    scheduler.Complete(pageNumber);
                }
                catch (OperationCanceledException)
                {
                    scheduler.Complete(pageNumber);
                    throw;
                }
                catch (Exception ex)
                {
                    if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
                        !ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
                    {
                        return;
                    }

                    AppLogger.Error(
                        ex,
                        $"Page Organizer thumbnail rendering failed: {pageNumber} in {path}");
                    var retryScheduled = scheduler.RegisterFailure(pageNumber);
                    if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
                        !ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
                    {
                        return;
                    }

                    item.Thumbnail = null;
                    item.ThumbnailRenderState = retryScheduled
                        ? PageOrganizerThumbnailRenderState.Pending
                        : PageOrganizerThumbnailRenderState.Failed;
                }
            }
        }
        catch (OperationCanceledException)
        {
            // A newer document replaced this organizer rendering request.
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, $"Page Organizer thumbnail rendering failed: {path}");
        }
        finally
        {
            _pageOrganizerRenderService.CloseSession(sessionId);
            if (_pageOrganizerThumbnailWorkerGeneration == generation)
            {
                _pageOrganizerThumbnailWorkerGeneration = null;
            }

            if (!ReferenceEquals(_pageOrganizerThumbnailCancellation, cancellation))
            {
                cancellation.Dispose();
            }
        }
    }

    private static BitmapImage LoadPageOrganizerThumbnail(string path)
    {
        var image = new BitmapImage();
        image.BeginInit();
        image.CacheOption = BitmapCacheOption.OnLoad;
        image.UriSource = new Uri(path, UriKind.Absolute);
        image.EndInit();
        image.Freeze();
        return image;
    }

    private void CancelPageOrganizerThumbnailRendering()
    {
        _pageOrganizerThumbnailGeneration++;
        var cancellation = _pageOrganizerThumbnailCancellation;
        _pageOrganizerThumbnailCancellation = null;
        _pageOrganizerThumbnailSourcePath = null;
        _pageOrganizerThumbnailScheduler = null;
        _pageOrganizerThumbnailCacheWindow = [];
        cancellation?.Cancel();
    }

    private void ResumePageOrganizerThumbnailRenderingAfterMutation()
    {
        if (IsDocumentMutationInProgress ||
            _pageOrganizerThumbnailScheduler is not null ||
            string.IsNullOrWhiteSpace(_currentPdfPath) ||
            _pageOrganizerState is not { PageNumbers.Count: > 0 } state)
        {
            return;
        }

        StartPageOrganizerThumbnailRendering(_currentPdfPath, state.PageNumbers);
    }

    private void OnPageOrganizerThumbnailScrollChanged(
        object sender,
        ScrollChangedEventArgs e)
    {
        if (e.VerticalChange == 0 &&
            e.ExtentHeightChange == 0 &&
            e.ViewportHeightChange == 0)
        {
            return;
        }

        RefreshPageOrganizerThumbnailViewport();
    }

    private IReadOnlyList<int> GetVisiblePageOrganizerThumbnailNumbers()
    {
        var scrollViewer = FindVisualDescendant<ScrollViewer>(PageOrganizerList);
        if (scrollViewer is null || scrollViewer.ViewportHeight <= 0)
        {
            return [];
        }

        var visiblePageNumbers = new List<int>();
        for (var index = 0; index < PageOrganizerItems.Count; index++)
        {
            if (PageOrganizerList.ItemContainerGenerator.ContainerFromIndex(index) is not FrameworkElement container ||
                container.ActualHeight <= 0)
            {
                continue;
            }

            var itemTop = container.TranslatePoint(new Point(), scrollViewer).Y;
            var itemBottom = itemTop + container.ActualHeight;
            if (itemBottom > 0 && itemTop < scrollViewer.ViewportHeight)
            {
                visiblePageNumbers.Add(PageOrganizerItems[index].PageNumber);
            }
        }

        return visiblePageNumbers;
    }

    private void RefreshPageOrganizerThumbnailViewport()
    {
        var scheduler = _pageOrganizerThumbnailScheduler;
        var cancellation = _pageOrganizerThumbnailCancellation;
        var generation = _pageOrganizerThumbnailGeneration;
        if (scheduler is null ||
            cancellation is null ||
            !IsCurrentPageOrganizerThumbnailRequest(generation, cancellation))
        {
            return;
        }

        var visiblePageNumbers = GetVisiblePageOrganizerThumbnailNumbers();
        var cacheWindow = scheduler.GetCacheWindow(visiblePageNumbers);
        if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
            !ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
        {
            return;
        }

        _pageOrganizerThumbnailCacheWindow = cacheWindow.ToHashSet();

        foreach (var item in PageOrganizerItems)
        {
            if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
                !ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
            {
                return;
            }

            if (!cacheWindow.Contains(item.PageNumber) &&
                item.Thumbnail is not null)
            {
                item.Thumbnail = null;
                item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Evicted;
            }
            else if (cacheWindow.Contains(item.PageNumber) &&
                     item.Thumbnail is null &&
                     item.ThumbnailRenderState != PageOrganizerThumbnailRenderState.Failed)
            {
                item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Pending;
                scheduler.Request(item.PageNumber, priority: true);
            }
        }

        if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
            !ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
        {
            return;
        }

        scheduler.Prioritize(cacheWindow);
        EnsurePageOrganizerThumbnailWorker(generation, cancellation);
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

    private Task SendPdfToViewerAsync(string path, int loadGeneration, bool dirtyAfterLoad = false, int? initialPage = null)
    {
        if (loadGeneration != _documentLoadGeneration || !_viewerReady || PdfViewer.CoreWebView2 is null)
        {
            return Task.CompletedTask;
        }

        var servedPath = PrepareServedPdfPath(path);
        var servedFileName = Path.GetFileName(servedPath);
        var pdfUrl = $"https://{ViewerHost}/web/ServedPdf/{Uri.EscapeDataString(servedFileName)}?v={Guid.NewGuid():N}";
        var fileLength = new FileInfo(path).Length;
        var message = JsonSerializer.Serialize(new
        {
            type = "loadPdf",
            url = pdfUrl,
            isDirty = dirtyAfterLoad,
            sourcePath = _referencePdfPath ?? path,
            largeDocumentHint = fileLength >= 80L * 1024 * 1024,
            fileLength,
            initialPage,
            loadId = loadGeneration
        });
        PdfViewer.CoreWebView2.PostWebMessageAsJson(message);
        AppLogger.Info($"PDF 뷰어 로드 요청: {path}, {fileLength:N0} bytes");
        return Task.CompletedTask;
    }

    private async Task HandleViewerLoadFailedAsync(string? message, int loadGeneration)
    {
        if (loadGeneration != _documentLoadGeneration)
        {
            return;
        }

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
                if (loadGeneration != _documentLoadGeneration)
                {
                    TryDeleteFile(recoveredPath);
                    return;
                }

                if (File.Exists(recoveredPath) && new FileInfo(recoveredPath).Length > 0)
                {
                    CleanupRecoveredPdf();
                    _recoveredPdfPath = recoveredPath;
                    _currentPdfPath = recoveredPath;
                    AppLogger.Info($"PDF 자동 복구 후 재로드: {_referencePdfPath} -> {recoveredPath}");
                    await SendPdfToViewerAsync(recoveredPath, loadGeneration, _isDirty);
                    return;
                }
            }
            catch (Exception ex)
            {
                TryDeleteFile(recoveredPath);
                AppLogger.Error(ex, $"PDF 자동 복구 실패: {_currentPdfPath}");
            }
        }

        if (await TryOpenFallbackViewerAsync(failureMessage, loadGeneration))
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

    private async Task<bool> TryOpenFallbackViewerAsync(string failureMessage, int loadGeneration)
    {
        if (loadGeneration != _documentLoadGeneration ||
            string.IsNullOrWhiteSpace(_currentPdfPath) ||
            !File.Exists(_currentPdfPath))
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
                loadId = loadGeneration,
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

        var viewerFolder = Path.Combine(AppContext.BaseDirectory, "Assets", ViewerAssetFolderName, "web");
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
        _documentLoadGeneration++;
        _documentOperations.Dispose();
        CancelPageOrganizerThumbnailRendering();
        _fallbackRenderService.Dispose();
        _pageOrganizerRenderService.Dispose();
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
                    static (_, _) => { },
                    (paths, _) => Dispatcher.BeginInvoke(() => HandleNativeFileDrop(paths)),
                    static (_, _) => { },
                    static () => { },
                    (payload, _) => Dispatcher.BeginInvoke(() => HandleNativePageTransferDrop(payload)));
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

    private void HandleNativePageTransferDrop(string payload)
    {
        if (!CanInsertDroppedFilesIntoCurrentDocument())
        {
            return;
        }

        try
        {
            var transfer = JsonSerializer.Deserialize<PageTransferPayload>(payload, PageTransferJsonOptions)
                ?? throw new InvalidOperationException("드롭한 페이지 정보를 읽을 수 없습니다.");
            _ = InsertExternalPagesAsync(new ExternalPagesDropMessage(
                transfer.SourcePath,
                transfer.Pages,
                GetInsertionIndexAfterSelection()));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "페이지 드롭 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void HandleNativeFileDrop(IReadOnlyList<string> paths)
    {
        if (CanInsertDroppedFilesIntoCurrentDocument())
        {
            _ = InsertExternalFilesAsync(new ExternalFilesDropMessage(
                paths.ToList(),
                GetInsertionIndexAfterSelection()));
            return;
        }

        OpenFiles(paths.Where(IsPdfFile));
    }

    private bool CanInsertDroppedFilesIntoCurrentDocument()
    {
        return !IsDocumentMutationInProgress &&
               !string.IsNullOrWhiteSpace(_currentPdfPath) &&
               _pageOrganizerState is { PageNumbers.Count: > 0 };
    }

    private void SendViewerCommand(string command, object? options = null, bool allowDuringDocumentMutation = false)
    {
        if ((!allowDuringDocumentMutation && IsDocumentMutationInProgress) ||
            !_viewerReady ||
            PdfViewer.CoreWebView2 is null)
        {
            return;
        }

        var message = JsonSerializer.Serialize(new
        {
            type = "command",
            command,
            options,
            loadId = _documentLoadGeneration
        });
        PdfViewer.CoreWebView2.PostWebMessageAsJson(message);
    }

    private void SendEditorFontsToViewer()
    {
        if (!_viewerReady || PdfViewer.CoreWebView2 is null)
        {
            return;
        }

        try
        {
            var fonts = WindowsFontService.GetInstalledFonts()
                .Select(font => new { font.Name })
                .ToList();
            var message = JsonSerializer.Serialize(new
            {
                type = "setEditorFonts",
                fonts
            });
            PdfViewer.CoreWebView2.PostWebMessageAsJson(message);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex, "Installed font list could not be sent to the PDF editor.");
        }
    }

    private async Task<EditorExportState> CollectEditorStateAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!_viewerReady || PdfViewer.CoreWebView2 is null)
        {
            return new EditorExportState([], [], false);
        }

        var requestId = Guid.NewGuid().ToString("N");
        var completion = new TaskCompletionSource<EditorExportState>(TaskCreationOptions.RunContinuationsAsynchronously);
        _editorStateCompletion = completion;
        _editorStateRequestId = requestId;
        var message = JsonSerializer.Serialize(new
        {
            type = "collectEditorState",
            requestId,
            loadId = _documentLoadGeneration
        });
        PdfViewer.CoreWebView2.PostWebMessageAsJson(message);

        try
        {
            var completed = await Task.WhenAny(
                completion.Task,
                Task.Delay(TimeSpan.FromSeconds(15), cancellationToken));
            if (completed != completion.Task)
            {
                cancellationToken.ThrowIfCancellationRequested();
                throw new TimeoutException("PDF editor state collection timed out.");
            }

            return await completion.Task;
        }
        finally
        {
            if (ReferenceEquals(_editorStateCompletion, completion))
            {
                _editorStateCompletion = null;
                _editorStateRequestId = null;
            }
        }
    }

    private void CompleteEditorStateCollection(JsonElement root)
    {
        var completion = _editorStateCompletion;
        if (completion is null)
        {
            return;
        }

        if (!IsExpectedRequest(root, _editorStateRequestId, "editor state collection"))
        {
            return;
        }

        try
        {
            var edits = root.TryGetProperty("edits", out var editsElement) && editsElement.ValueKind == JsonValueKind.Array
                ? editsElement.EnumerateArray().Select(element => element.Clone()).ToList()
                : [];
            var usedFontNames = root.TryGetProperty("usedFontNames", out var fontsElement) && fontsElement.ValueKind == JsonValueKind.Array
                ? fontsElement.EnumerateArray().Select(element => element.GetString() ?? string.Empty).Where(name => !string.IsNullOrWhiteSpace(name)).ToList()
                : [];
            var isDirty = root.TryGetProperty("isDirty", out var dirtyElement) && dirtyElement.GetBoolean();
            completion.TrySetResult(new EditorExportState(edits, usedFontNames, isDirty));
        }
        catch (Exception ex)
        {
            completion.TrySetException(ex);
        }
        finally
        {
            if (ReferenceEquals(_editorStateCompletion, completion))
            {
                _editorStateCompletion = null;
                _editorStateRequestId = null;
            }
        }
    }

    private async Task<string> ExportOverlayPdfAsync(
        string sourcePath,
        EditorExportState editorState,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!_viewerReady || PdfViewer.CoreWebView2 is null)
        {
            throw new InvalidOperationException("PDF viewer is not ready.");
        }

        var requestId = Guid.NewGuid().ToString("N");
        var completion = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        _overlayPdfExportCompletion = completion;
        _overlayPdfExportRequestId = requestId;
        var message = JsonSerializer.Serialize(new
        {
            type = "exportOverlayPdf",
            requestId,
            loadId = _documentLoadGeneration,
            sourceBase64 = Convert.ToBase64String(File.ReadAllBytes(sourcePath)),
            edits = editorState.Edits,
            fonts = WindowsFontService.ReadFontBase64(editorState.UsedFontNames)
        });
        PdfViewer.CoreWebView2.PostWebMessageAsJson(message);

        try
        {
            var completed = await Task.WhenAny(
                completion.Task,
                Task.Delay(TimeSpan.FromSeconds(60), cancellationToken));
            if (completed != completion.Task)
            {
                cancellationToken.ThrowIfCancellationRequested();
                throw new TimeoutException("PDF editor export timed out.");
            }

            return await completion.Task;
        }
        finally
        {
            if (ReferenceEquals(_overlayPdfExportCompletion, completion))
            {
                _overlayPdfExportCompletion = null;
                _overlayPdfExportRequestId = null;
            }
        }
    }

    private void CompleteOverlayPdfExport(JsonElement root)
    {
        var completion = _overlayPdfExportCompletion;
        if (completion is null)
        {
            return;
        }

        if (!IsExpectedRequest(root, _overlayPdfExportRequestId, "overlay PDF export"))
        {
            return;
        }

        try
        {
            var pdfBase64 = root.TryGetProperty("pdfBase64", out var pdfElement)
                ? pdfElement.GetString() ?? string.Empty
                : string.Empty;
            completion.TrySetResult(pdfBase64);
        }
        catch (Exception ex)
        {
            completion.TrySetException(ex);
        }
        finally
        {
            if (ReferenceEquals(_overlayPdfExportCompletion, completion))
            {
                _overlayPdfExportCompletion = null;
                _overlayPdfExportRequestId = null;
            }
        }
    }

    private void CompleteOverlayPdfExportFailure(JsonElement root)
    {
        var completion = _overlayPdfExportCompletion;
        if (completion is null)
        {
            return;
        }

        if (!IsExpectedRequest(root, _overlayPdfExportRequestId, "overlay PDF export failure"))
        {
            return;
        }

        try
        {
            var message = root.TryGetProperty("message", out var messageElement)
                ? messageElement.GetString()
                : null;
            completion.TrySetException(new InvalidOperationException(
                string.IsNullOrWhiteSpace(message)
                    ? "PDF editor export failed."
                    : message));
        }
        finally
        {
            if (ReferenceEquals(_overlayPdfExportCompletion, completion))
            {
                _overlayPdfExportCompletion = null;
                _overlayPdfExportRequestId = null;
            }
        }
    }

    private static bool IsExpectedRequest(JsonElement root, string? expectedRequestId, string operation)
    {
        var actualRequestId = root.TryGetProperty("requestId", out var requestElement)
            ? requestElement.GetString()
            : null;
        if (!string.IsNullOrWhiteSpace(expectedRequestId) &&
            string.Equals(actualRequestId, expectedRequestId, StringComparison.Ordinal))
        {
            return true;
        }

        AppLogger.Info($"Ignoring stale {operation} response. Expected requestId={expectedRequestId ?? "<none>"}, actual requestId={actualRequestId ?? "<none>"}.");
        return false;
    }

    private async Task<IReadOnlyList<PdfImagePage>> ExportCurrentPagesAsA4ImagesAsync(
        bool optimizeSize,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var sessionId = _fallbackModeActive && !string.IsNullOrWhiteSpace(_fallbackSessionId)
            ? _fallbackSessionId
            : null;
        var ownsSession = false;
        if (sessionId is null)
        {
            var sourcePath = _currentPdfPath ?? throw new InvalidOperationException("현재 PDF를 찾을 수 없습니다.");
            sessionId = _fallbackRenderService.OpenDocument(sourcePath).SessionId;
            ownsSession = true;
        }

        try
        {
            AppLogger.Info($"PDFium A4 이미지 내보내기: {_referencePdfPath ?? _currentPdfPath}");
            return await _fallbackRenderService.ExportPagesAsImagesAsync(
                sessionId,
                _pageOrder,
                _pageRotations,
                optimizeSize ? A4OptimizedMaxWidthPixels : A4ImageMaxWidthPixels,
                optimizeSize ? A4OptimizedMaxHeightPixels : A4ImageMaxHeightPixels,
                optimizeSize ? 86 : 92,
                cancellationToken);
        }
        finally
        {
            if (ownsSession)
            {
                _fallbackRenderService.CloseSession(sessionId);
            }
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
        var previousOptimizedPartialRendering = _settings.EnableOptimizedPartialRendering;
        var settingsWindow = new SettingsWindow(_settings)
        {
            Owner = this
        };

        if (settingsWindow.ShowDialog() == true)
        {
            UpdatePdfContextMenuOption();
            UpdateRecentFilesMenu();
            if (previousOptimizedPartialRendering != _settings.EnableOptimizedPartialRendering)
            {
                ReloadViewerAfterViewerSettingsChanged();
            }
        }
    }

    private void ReloadViewerAfterViewerSettingsChanged()
    {
        if (PdfViewer.CoreWebView2 is null)
        {
            return;
        }

        var currentPath = _currentPdfPath;
        var referencePath = _referencePdfPath;
        var dirtyAfterLoad = _isDirty;
        _viewerReady = false;
        if (!string.IsNullOrWhiteSpace(currentPath))
        {
            _pendingPdfPath = currentPath;
            _pendingReferencePdfPath = referencePath ?? currentPath;
            _pendingDirtyAfterLoad = dirtyAfterLoad;
            _pendingInitialPage = _activePage;
            _pendingLoadGeneration = _documentLoadGeneration;
        }

        PdfViewer.CoreWebView2.Navigate(BuildViewerUrl());
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
            LoadPdf(_currentPdfPath, preserveWorkingSaveTarget: true);
        }
    }

    private async void OnSavePageOrderClick(object sender, RoutedEventArgs e)
    {
        if (IsDocumentMutationInProgress)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(_currentPdfPath) || _pageOrder.Count == 0)
        {
            MessageBox.Show(this, "먼저 PDF를 열어주세요.", "PDF 저장", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var outputPath = ResolveWorkingSaveTarget();
        if (outputPath is null)
        {
            return;
        }

        if (!TryCaptureDocumentOperation(out var operation))
        {
            return;
        }
        string? transformedTempPath = null;
        string? publicationTempPath = null;
        try
        {
            await ExecuteDocumentMutationAsync(operation, async currentOperation =>
            {
                var sourcePath = _currentPdfPath ?? throw new InvalidOperationException("저장할 PDF를 찾을 수 없습니다.");
                var editorState = await CollectEditorStateAsync(currentOperation.CancellationToken);
                _documentOperations.ThrowIfSuperseded(currentOperation);

                var pageTransforms = GetCurrentPageTransforms();
                var remappedEditorState = RemapEditorStateToOutputPageOrder(editorState, pageTransforms);
                publicationTempPath = CreatePublicationTempPath(outputPath);
                var outputTarget = remappedEditorState.Edits.Count > 0 ? CreateTempPdfPath("editor-source") : publicationTempPath;
                var result = await _pdfService.SaveTransformedPagesAsync(
                    sourcePath,
                    pageTransforms,
                    outputTarget,
                    currentOperation.CancellationToken);
                _documentOperations.ThrowIfSuperseded(currentOperation);

                transformedTempPath = remappedEditorState.Edits.Count > 0 ? result.OutputPath : null;
                if (remappedEditorState.Edits.Count > 0)
                {
                    var exportedBase64 = await ExportOverlayPdfAsync(
                        result.OutputPath,
                        remappedEditorState,
                        currentOperation.CancellationToken);
                    _documentOperations.ThrowIfSuperseded(currentOperation);
                    await File.WriteAllBytesAsync(
                        publicationTempPath,
                        Convert.FromBase64String(exportedBase64),
                        currentOperation.CancellationToken);
                }

                _documentOperations.ThrowIfSuperseded(currentOperation);
                VerifySavedPdf(publicationTempPath, pageTransforms.Count);
                _documentOperations.ThrowIfSuperseded(currentOperation);
                PdfSavePublisher.Publish(publicationTempPath, outputPath);
                publicationTempPath = null;
                result = result with { OutputPath = outputPath };
                var message = string.IsNullOrWhiteSpace(result.WarningMessage)
                    ? $"저장 완료:\n{result.OutputPath}"
                    : $"저장 완료:\n{result.OutputPath}\n\n참고: {result.WarningMessage}";

                _editorDirty = false;
                if (_pageOrganizerState is not null)
                {
                    ApplyPageOrganizerState(_pageOrganizerState.MarkClean());
                }
                else
                {
                    _isDirty = false;
                    UpdateWindowTitle();
                }

                SendViewerCommand("markClean", allowDuringDocumentMutation: true);
                LoadPdf(result.OutputPath, preserveWorkingSaveTarget: true);
                MessageBox.Show(this, message, "PDF 저장", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }
        catch (OperationCanceledException) when (IsDocumentOperationSuperseded(operation))
        {
            AppLogger.Info("Ignored a save result from a document that was replaced while saving.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "PDF 저장 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            if (transformedTempPath is not null)
            {
                TryDeleteTempFile(transformedTempPath);
            }

            if (publicationTempPath is not null)
            {
                TryDeleteTempFile(publicationTempPath);
            }
        }
    }

    private string? ResolveWorkingSaveTarget()
    {
        if (!string.IsNullOrWhiteSpace(_workingSaveTargetPath))
        {
            return _workingSaveTargetPath;
        }

        var referencePath = _referencePdfPath ?? _currentPdfPath!;
        var folder = Path.GetDirectoryName(referencePath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var defaultPath = Path.Combine(folder, $"{Path.GetFileNameWithoutExtension(referencePath)}_편집본.pdf");
        var dialog = SavePathPromptService.CreateSaveDialog(defaultPath, "PDF 다른 이름으로 저장");
        if (dialog.ShowDialog(this) != true)
        {
            return null;
        }

        var outputPath = SavePathPromptService.ResolveOutputPath(this, dialog.FileName, "PDF 다른 이름으로 저장");
        if (outputPath is not null)
        {
            _workingSaveTargetPath = outputPath;
        }

        return outputPath;
    }

    private static string CreatePublicationTempPath(string outputPath)
    {
        var folder = Path.GetDirectoryName(outputPath) ?? Path.GetTempPath();
        var name = Path.GetFileNameWithoutExtension(outputPath);
        return Path.Combine(folder, $".{name}.{Guid.NewGuid():N}.pending.pdf");
    }

    private void VerifySavedPdf(string outputPath, int expectedPageCount)
    {
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("저장된 임시 PDF 파일을 찾을 수 없습니다.");
        }

        var savedPageCount = _pdfService.GetPageCount(outputPath);
        if (savedPageCount != expectedPageCount)
        {
            throw new InvalidOperationException($"저장 검증 실패: 페이지 수가 예상과 다릅니다. 예상 {expectedPageCount}, 실제 {savedPageCount}");
        }
    }

    private async void OnExtractSelectedPagesClick(object sender, RoutedEventArgs e)
    {
        if (IsDocumentMutationInProgress || !EnsurePdfLoaded("PDF 추출"))
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

        if (!TryCaptureDocumentOperation(out var operation))
        {
            return;
        }
        try
        {
            await ExecuteDocumentMutationAsync(operation, async currentOperation =>
            {
                var sourcePath = _currentPdfPath ?? throw new InvalidOperationException("현재 PDF를 찾을 수 없습니다.");
                var selected = _pageOrder
                    .Where(page => _selectedPages.Contains(page))
                    .Select(page => new PdfPageTransform(page, GetRotation(page)))
                    .ToList();
                var result = await _pdfService.SaveTransformedPagesAsync(
                    sourcePath,
                    selected,
                    outputPath,
                    currentOperation.CancellationToken);
                _documentOperations.ThrowIfSuperseded(currentOperation);
                MessageBox.Show(this, $"저장 완료:\n{result.OutputPath}", "PDF 추출", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }
        catch (OperationCanceledException) when (IsDocumentOperationSuperseded(operation))
        {
            AppLogger.Info("Ignored an extract result from a document that was replaced while extracting.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "PDF 추출 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void OnSplitPagesClick(object sender, RoutedEventArgs e)
    {
        if (!CanSplitCurrentDocument())
        {
            return;
        }

        await SplitCurrentDocumentAsync(PdfSplitMode.PageByPage, 1);
    }

    private async void OnSplitByIntervalClick(object sender, RoutedEventArgs e)
    {
        if (!CanSplitCurrentDocument())
        {
            return;
        }

        var intervalWindow = new SplitIntervalWindow { Owner = this };
        if (intervalWindow.ShowDialog() != true)
        {
            return;
        }

        await SplitCurrentDocumentAsync(PdfSplitMode.Interval, intervalWindow.Interval);
    }

    private async void OnSplitByParityClick(object sender, RoutedEventArgs e)
    {
        if (!CanSplitCurrentDocument())
        {
            return;
        }

        await SplitCurrentDocumentAsync(PdfSplitMode.Parity, 1);
    }

    private bool CanSplitCurrentDocument()
    {
        return !IsDocumentMutationInProgress && EnsurePdfLoaded("PDF 분할");
    }

    private async Task SplitCurrentDocumentAsync(PdfSplitMode splitMode, int interval)
    {
        var referencePath = _referencePdfPath ?? _currentPdfPath!;

        if (!TryCaptureDocumentOperation(out var operation))
        {
            return;
        }
        try
        {
            await ExecuteDocumentMutationAsync(operation, async currentOperation =>
            {
                var sourcePath = _currentPdfPath ?? throw new InvalidOperationException("현재 PDF를 찾을 수 없습니다.");
                var pageTransforms = GetCurrentPageTransforms();
                var plan = splitMode switch
                {
                    PdfSplitMode.PageByPage => PdfSplitPlanner.CreateIntervalPlan(sourcePath, pageTransforms, 1, referencePath),
                    PdfSplitMode.Interval => PdfSplitPlanner.CreateIntervalPlan(sourcePath, pageTransforms, interval, referencePath),
                    PdfSplitMode.Parity => PdfSplitPlanner.CreateParityPlan(sourcePath, pageTransforms, referencePath),
                    _ => throw new ArgumentOutOfRangeException(nameof(splitMode))
                };
                var results = await _pdfService.ExecuteSplitPlanAsync(plan, currentOperation.CancellationToken);
                _documentOperations.ThrowIfSuperseded(currentOperation);
                MessageBox.Show(
                    this,
                    $"{results.Count}개 파일로 {GetSplitDescription(splitMode, interval)} 분할했습니다.\n{plan.OutputFolder}",
                    "PDF 분할",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            });
        }
        catch (OperationCanceledException) when (IsDocumentOperationSuperseded(operation))
        {
            AppLogger.Info("Ignored a split result from a document that was replaced while splitting.");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "PDF 분할 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string GetSplitDescription(PdfSplitMode splitMode, int interval)
    {
        return splitMode switch
        {
            PdfSplitMode.PageByPage => "페이지별로",
            PdfSplitMode.Interval => $"{interval}페이지마다",
            PdfSplitMode.Parity => "홀수/짝수로",
            _ => throw new ArgumentOutOfRangeException(nameof(splitMode))
        };
    }

    private async void OnAddBlankA4PageClick(object sender, RoutedEventArgs e)
    {
        if (!EnsurePdfLoaded("빈 페이지 추가"))
        {
            return;
        }

        await RunCurrentDocumentMutationAsync("blank A4 page insertion", "빈 페이지 추가 실패", async operation =>
        {
            var insertPath = CreateTempPdfPath("blank");
            try
            {
                _pdfService.CreateBlankA4Pdf(insertPath);
                await InsertGeneratedPdfPageAsync(insertPath, "빈 A4 페이지 추가", operation);
            }
            finally
            {
                TryDeleteTempFile(insertPath);
            }
        });
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

        await RunCurrentDocumentMutationAsync("clipboard image insertion", "이미지 붙여넣기 실패", async operation =>
        {
            var insertPath = CreateTempPdfPath("clipboard-image");
            try
            {
                var encodedImage = EncodeJpegForA4(image);
                _pdfService.CreateImageA4Pdf(
                    insertPath,
                    encodedImage.JpegBytes,
                    encodedImage.Width,
                    encodedImage.Height);
                await InsertGeneratedPdfPageAsync(insertPath, "클립보드 이미지 붙여넣기", operation);
            }
            finally
            {
                TryDeleteTempFile(insertPath);
            }
        });
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
        if (IsDocumentMutationInProgress ||
            !EnsurePdfLoaded(cut ? "페이지 잘라내기" : "페이지 복사"))
        {
            return;
        }

        if (!TryCaptureDocumentOperation(out var operation))
        {
            return;
        }
        try
        {
            await ExecuteDocumentMutationAsync(operation, async currentOperation =>
            {
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

                var sourcePath = _currentPdfPath ?? throw new InvalidOperationException("현재 PDF를 찾을 수 없습니다.");
                var tempPath = CreateTempPdfPath(cut ? "cut-pages" : "copy-pages");
                var keepTempPathForClipboard = false;
                try
                {
                    await _pdfService.SaveTransformedPagesAsync(
                        sourcePath,
                        pages,
                        tempPath,
                        currentOperation.CancellationToken);
                    _documentOperations.ThrowIfSuperseded(currentOperation);

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
                    keepTempPathForClipboard = true;

                    if (cut)
                    {
                        ApplyPageOrganizerEdit(
                            state => state.DeleteSelectedPages(),
                            allowDuringDocumentMutation: true);
                    }

                    CurrentFileText.Text = cut
                        ? $"{pages.Count}개 페이지를 잘라냈습니다."
                        : $"{pages.Count}개 페이지를 복사했습니다.";
                }
                finally
                {
                    if (!keepTempPathForClipboard)
                    {
                        TryDeleteTempFile(tempPath);
                    }
                }
            });
        }
        catch (OperationCanceledException) when (IsDocumentOperationSuperseded(operation))
        {
            AppLogger.Info("Ignored a copy or cut result from a document that was replaced while processing.");
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

        var json = Clipboard.GetData(PageTransferClipboardFormat) as string;
        var payload = string.IsNullOrWhiteSpace(json)
            ? null
            : JsonSerializer.Deserialize<PageTransferPayload>(json, PageTransferJsonOptions);
        if (payload is null || !File.Exists(payload.SourcePath))
        {
            MessageBox.Show(this, "붙여넣을 페이지 데이터를 찾을 수 없습니다.", "페이지 붙여넣기", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var insertionIndex = GetInsertionIndexAfterSelection();
        await RunCurrentDocumentMutationAsync("page paste", "페이지 붙여넣기 실패", async operation =>
        {
            await InsertPreparedPdfPagesAsync(payload.SourcePath, insertionIndex, "페이지 붙여넣기", showMessage: false, operation);
            CurrentFileText.Text = payload.Cut ? "페이지를 이동했습니다." : "페이지를 붙여넣었습니다.";
        });
    }

    private async Task InsertExternalPagesAsync(JsonElement root)
    {
        try
        {
            var message = JsonSerializer.Deserialize<ExternalPagesDropMessage>(root.GetRawText(), PageTransferJsonOptions)
                ?? throw new InvalidOperationException("드롭한 페이지 정보를 읽을 수 없습니다.");
            await InsertExternalPagesAsync(message);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "페이지 드롭 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task InsertExternalPagesAsync(ExternalPagesDropMessage message)
    {
        if (!EnsurePdfLoaded("페이지 드롭"))
        {
            return;
        }

        if (!File.Exists(message.SourcePath))
        {
            MessageBox.Show(this, "원본 PDF 파일을 찾을 수 없습니다.", "페이지 드롭 실패", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (message.Pages.Count == 0)
        {
            MessageBox.Show(this, "드롭한 페이지가 없습니다.", "페이지 드롭 실패", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        await RunCurrentDocumentMutationAsync("external page insertion", "페이지 드롭 실패", async operation =>
        {
            var selectedPath = CreateTempPdfPath("drag-pages");
            try
            {
                await _pdfService.SaveTransformedPagesAsync(
                    message.SourcePath,
                    message.Pages,
                    selectedPath,
                    operation.CancellationToken);
                _documentOperations.ThrowIfSuperseded(operation);
                await InsertPreparedPdfPagesAsync(
                    selectedPath,
                    Math.Clamp(message.InsertionIndex, 0, _pageOrder.Count),
                    "페이지 드롭",
                    showMessage: false,
                    operation);
                CurrentFileText.Text = $"{message.Pages.Count}개 페이지를 드롭해서 복사했습니다.";
            }
            finally
            {
                TryDeleteTempFile(selectedPath);
            }
        });
    }

    private async Task InsertExternalFilesAsync(JsonElement root)
    {
        try
        {
            var message = JsonSerializer.Deserialize<ExternalFilesDropMessage>(root.GetRawText(), PageTransferJsonOptions)
                ?? throw new InvalidOperationException("드롭한 파일 정보를 읽을 수 없습니다.");
            await InsertExternalFilesAsync(message);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "파일 삽입 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task InsertExternalFilesAsync(ExternalFilesDropMessage message)
    {
        if (!EnsurePdfLoaded("파일 삽입"))
        {
            return;
        }

        await RunCurrentDocumentMutationAsync("external file insertion", "파일 삽입 실패", async operation =>
        {
            var tempPaths = new List<string>();
            try
            {
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
                    await _pdfService.CombinePdfFilesAsync(pdfParts, insertPath, operation.CancellationToken);
                    _documentOperations.ThrowIfSuperseded(operation);
                }

                await InsertPreparedPdfPagesAsync(
                    insertPath,
                    Math.Clamp(message.InsertionIndex, 0, _pageOrder.Count),
                    "파일 삽입",
                    showMessage: false,
                    operation);
                CurrentFileText.Text = $"{paths.Count}개 파일을 현재 PDF에 삽입했습니다.";
            }
            finally
            {
                foreach (var tempPath in tempPaths)
                {
                    TryDeleteTempFile(tempPath);
                }
            }
        });
    }

    private async void OnFitAllPagesToA4Click(object sender, RoutedEventArgs e)
    {
        if (IsDocumentMutationInProgress || !EnsurePdfLoaded("A4 맞춤"))
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

        if (!TryCaptureDocumentOperation(out var operation))
        {
            return;
        }
        try
        {
            await ExecuteDocumentMutationAsync(operation, async currentOperation =>
            {
                ViewerLoading.Visibility = Visibility.Visible;
                CurrentFileText.Text = "A4 페이지로 변환 중입니다...";
                var images = await ExportCurrentPagesAsA4ImagesAsync(false, currentOperation.CancellationToken);
                _documentOperations.ThrowIfSuperseded(currentOperation);
                var result = await _pdfService.CreateA4ImagePagesPdfAsync(
                    images,
                    outputPath,
                    currentOperation.CancellationToken);
                _documentOperations.ThrowIfSuperseded(currentOperation);
                LoadPdf(result.OutputPath);
                MessageBox.Show(this, $"저장 완료:\n{result.OutputPath}", "A4 맞춤", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }
        catch (OperationCanceledException) when (IsDocumentOperationSuperseded(operation))
        {
            AppLogger.Info("Ignored an A4 conversion result from a document that was replaced while converting.");
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
        if (IsDocumentMutationInProgress || !EnsurePdfLoaded("A4 기준 용량 최적화"))
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

        if (!TryCaptureDocumentOperation(out var operation))
        {
            return;
        }
        try
        {
            await ExecuteDocumentMutationAsync(operation, async currentOperation =>
            {
                var sourcePath = _currentPdfPath ?? throw new InvalidOperationException("현재 PDF를 찾을 수 없습니다.");
                ViewerLoading.Visibility = Visibility.Visible;
                CurrentFileText.Text = "A4 기준으로 용량 최적화 중입니다...";
                var originalSize = File.Exists(sourcePath) ? new FileInfo(sourcePath).Length : 0;
                var images = await ExportCurrentPagesAsA4ImagesAsync(true, currentOperation.CancellationToken);
                _documentOperations.ThrowIfSuperseded(currentOperation);
                var result = await _pdfService.CreateA4ImagePagesPdfAsync(
                    images,
                    outputPath,
                    currentOperation.CancellationToken);
                _documentOperations.ThrowIfSuperseded(currentOperation);
                LoadPdf(result.OutputPath);
                var outputSize = new FileInfo(result.OutputPath).Length;
                var sizeMessage = originalSize > 0
                    ? $"\n원본: {FormatFileSize(originalSize)}\n결과: {FormatFileSize(outputSize)}"
                    : $"\n결과: {FormatFileSize(outputSize)}";
                MessageBox.Show(this, $"저장 완료:\n{result.OutputPath}{sizeMessage}", "A4 기준 용량 최적화", MessageBoxButton.OK, MessageBoxImage.Information);
            });
        }
        catch (OperationCanceledException) when (IsDocumentOperationSuperseded(operation))
        {
            AppLogger.Info("Ignored an A4 optimization result from a document that was replaced while optimizing.");
        }
        catch (Exception ex)
        {
            ViewerLoading.Visibility = Visibility.Collapsed;
            CurrentFileText.Text = _currentPdfPath ?? "PDF를 열거나 이 창에 끌어다 놓으세요.";
            MessageBox.Show(this, ex.Message, "A4 기준 용량 최적화 실패", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task InsertGeneratedPdfPageAsync(
        string insertPath,
        string title,
        DocumentOperationToken operation)
    {
        var insertionIndex = GetInsertionIndexAfterSelection();
        await InsertPreparedPdfPagesAsync(insertPath, insertionIndex, title, showMessage: true, operation);
    }

    private async Task InsertPreparedPdfPagesAsync(
        string insertPath,
        int insertionIndex,
        string title,
        bool showMessage,
        DocumentOperationToken operation)
    {
        _documentOperations.ThrowIfSuperseded(operation);
        var outputPath = CreateTempPdfPath("edited");
        var referencePath = _referencePdfPath ?? _currentPdfPath!;
        var restorePageNumber = _activePage.GetValueOrDefault(_selectedPages.FirstOrDefault());
        var sourcePath = _currentPdfPath ?? throw new InvalidOperationException("현재 PDF를 찾을 수 없습니다.");
        var pageTransforms = GetCurrentPageTransforms();

        await _pdfService.InsertPdfPagesAsync(
            sourcePath,
            pageTransforms,
            insertPath,
            insertionIndex,
            outputPath,
            operation.CancellationToken);
        _documentOperations.ThrowIfSuperseded(operation);

        LoadPdf(outputPath, referencePath, true, restorePageNumber > 0 ? restorePageNumber : null);
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

    private static EditorExportState RemapEditorStateToOutputPageOrder(
        EditorExportState editorState,
        IReadOnlyList<PdfPageTransform> pageTransforms)
    {
        if (editorState.Edits.Count == 0)
        {
            return editorState;
        }

        var pageToOutputIndex = pageTransforms
            .Select((transform, index) => new { transform.PageNumber, OutputPage = index + 1 })
            .GroupBy(item => item.PageNumber)
            .ToDictionary(group => group.Key, group => group.First().OutputPage);

        var remappedEdits = new List<JsonElement>(editorState.Edits.Count);
        foreach (var edit in editorState.Edits)
        {
            if (!edit.TryGetProperty("page", out var pageElement) ||
                pageElement.ValueKind != JsonValueKind.Number ||
                !pageToOutputIndex.TryGetValue(pageElement.GetInt32(), out var outputPage))
            {
                continue;
            }

            var editNode = JsonNode.Parse(edit.GetRawText())?.AsObject();
            if (editNode is null)
            {
                continue;
            }

            editNode["page"] = outputPage;
            using var document = JsonDocument.Parse(editNode.ToJsonString());
            remappedEdits.Add(document.RootElement.Clone());
        }

        return editorState with { Edits = remappedEdits };
    }

    private List<PdfPageTransform> GetSelectedPageTransforms()
    {
        return _pageOrder
            .Where(page => _selectedPages.Contains(page))
            .Select(page => new PdfPageTransform(page, GetRotation(page)))
            .ToList();
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
        if (IsDocumentMutationInProgress)
        {
            return;
        }

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

    private void OnThumbZoomInClick(object sender, RoutedEventArgs e) => AdjustPageOrganizerThumbnailHeight(18);

    private void OnThumbZoomOutClick(object sender, RoutedEventArgs e) => AdjustPageOrganizerThumbnailHeight(-18);

    private void OnThumbZoomResetClick(object sender, RoutedEventArgs e) =>
        SetPageOrganizerThumbnailHeight(DefaultPageOrganizerThumbnailHeight);

    private void OnPageOrganizerZoomSliderValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) =>
        SetPageOrganizerThumbnailHeight(e.NewValue, synchronizeSlider: false);

    private void AdjustPageOrganizerThumbnailHeight(double delta)
    {
        SetPageOrganizerThumbnailHeight(_pageOrganizerThumbnailHeight + delta);
    }

    private void SetPageOrganizerThumbnailHeight(double height, bool synchronizeSlider = true)
    {
        var normalizedHeight = Math.Round(Math.Clamp(
            height,
            MinimumPageOrganizerThumbnailHeight,
            MaximumPageOrganizerThumbnailHeight));
        if (Math.Abs(_pageOrganizerThumbnailHeight - normalizedHeight) < 0.1)
        {
            return;
        }

        _pageOrganizerThumbnailHeight = normalizedHeight;
        if (synchronizeSlider &&
            PageOrganizerZoomSlider is not null &&
            Math.Abs(PageOrganizerZoomSlider.Value - normalizedHeight) >= 0.1)
        {
            PageOrganizerZoomSlider.Value = normalizedHeight;
        }

        RefreshPageOrganizerThumbnailHeight();
    }

    private void RefreshPageOrganizerThumbnailHeight()
    {
        foreach (var item in PageOrganizerItems)
        {
            ApplyPageOrganizerThumbnailDimensions(item);
        }
    }

    private void ApplyPageOrganizerThumbnailDimensions(PageOrganizerItem item)
    {
        var thumbnailWidth = Math.Round(_pageOrganizerThumbnailHeight * PageOrganizerThumbnailWidthRatio);
        item.ThumbnailHeight = _pageOrganizerThumbnailHeight;
        item.ThumbnailWidth = thumbnailWidth;
        item.ThumbnailCardWidth = thumbnailWidth + PageOrganizerThumbnailCardExtraWidth;
    }

    private void OnPrevPageClick(object sender, RoutedEventArgs e) => NavigatePageOrganizer(-1);

    private void OnNextPageClick(object sender, RoutedEventArgs e) => NavigatePageOrganizer(1);

    private void OnFirstPageClick(object sender, RoutedEventArgs e) => NavigatePageOrganizerBoundary(last: false);

    private void OnLastPageClick(object sender, RoutedEventArgs e) => NavigatePageOrganizerBoundary(last: true);

    private void OnUndoClick(object sender, RoutedEventArgs e)
    {
        if (IsDocumentMutationInProgress)
        {
            return;
        }

        if (_pageOrganizerState?.CanUndo == true)
        {
            ApplyPageOrganizerState(_pageOrganizerState.Undo(), navigatePreview: true);
            return;
        }

        SendViewerCommand("undo");
    }

    private void OnRedoClick(object sender, RoutedEventArgs e)
    {
        if (IsDocumentMutationInProgress)
        {
            return;
        }

        if (_pageOrganizerState?.CanRedo == true)
        {
            ApplyPageOrganizerState(_pageOrganizerState.Redo(), navigatePreview: true);
            return;
        }

        SendViewerCommand("redo");
    }

    private void OnEditorTextClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorText");

    private void OnEditorReplaceTextClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorReplaceText");

    private void OnEditorWhiteoutClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorWhiteout");

    private void OnEditorRedactClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorRedact");

    private void OnEditorUnderlineClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorUnderline");

    private void OnEditorStrikeoutClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorStrikeout");

    private void OnEditorRectangleClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorRectangle");

    private void OnEditorEllipseClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorEllipse");

    private void OnEditorLineClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorLine");

    private void OnEditorArrowClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorArrow");

    private void OnEditorPenClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorPen");

    private void OnEditorHighlightClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorHighlight");

    private void OnEditorStampClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorStamp");

    private void OnEditorCopySelectionClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorCopySelection");

    private void OnEditorPasteSelectionClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorPasteSelection");

    private void OnEditorDuplicateSelectionClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorDuplicateSelection");

    private void OnEditorBringForwardClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorBringForward");

    private void OnEditorSendBackwardClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorSendBackward");

    private void OnEditorBringToFrontClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorBringToFront");

    private void OnEditorSendToBackClick(object sender, RoutedEventArgs e) => SendViewerCommand("editorSendToBack");

    private async void OnCopySelectedPagesClick(object sender, RoutedEventArgs e) => await CopySelectedPagesToClipboardAsync(cut: false);

    private async void OnCutSelectedPagesClick(object sender, RoutedEventArgs e) => await CopySelectedPagesToClipboardAsync(cut: true);

    private async void OnPastePagesClick(object sender, RoutedEventArgs e) => await PastePagesOrImageAsync();

    private void OnDeleteSelectedPagesClick(object sender, RoutedEventArgs e) =>
        ApplyPageOrganizerEdit(state => state.DeleteSelectedPages());

    private void OnRotateClockwiseClick(object sender, RoutedEventArgs e) =>
        ApplyPageOrganizerEdit(state => state.RotateSelectedPages(90));

    private void OnRotateCounterClockwiseClick(object sender, RoutedEventArgs e) =>
        ApplyPageOrganizerEdit(state => state.RotateSelectedPages(-90));

    private void OnReversePageOrderClick(object sender, RoutedEventArgs e) =>
        ApplyPageOrganizerEdit(state => state.ReversePageOrder());

    private void ApplyPageOrganizerEdit(
        Func<EditorDocumentState, EditorDocumentState> edit,
        bool allowDuringDocumentMutation = false)
    {
        if ((!allowDuringDocumentMutation && IsDocumentMutationInProgress) ||
            _pageOrganizerState is null)
        {
            return;
        }

        var next = edit(_pageOrganizerState);
        if (ReferenceEquals(next, _pageOrganizerState))
        {
            return;
        }

        ApplyPageOrganizerState(
            next,
            navigatePreview: true,
            allowDuringDocumentMutation);
    }

    private void NavigatePageOrganizer(int delta)
    {
        if (IsDocumentMutationInProgress ||
            _pageOrganizerState is not { PageNumbers.Count: > 0 } state)
        {
            return;
        }

        var current = state.ActivePageNumber ?? state.SelectedPageNumbers.FirstOrDefault(state.PageNumbers[0]);
        var currentIndex = state.PageNumbers.ToList().IndexOf(current);
        var nextIndex = Math.Clamp(currentIndex + delta, 0, state.PageNumbers.Count - 1);
        ApplyPageOrganizerState(state.ActivatePage(state.PageNumbers[nextIndex]), navigatePreview: true);
    }

    private void NavigatePageOrganizerBoundary(bool last)
    {
        if (IsDocumentMutationInProgress ||
            _pageOrganizerState is not { PageNumbers.Count: > 0 } state)
        {
            return;
        }

        var pageNumber = last ? state.PageNumbers[^1] : state.PageNumbers[0];
        ApplyPageOrganizerState(state.ActivatePage(pageNumber), navigatePreview: true);
    }

    private void OnPageOrganizerCheckBoxPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (IsDocumentMutationInProgress)
        {
            e.Handled = true;
            return;
        }

        if (_pageOrganizerState is null || sender is not CheckBox { DataContext: PageOrganizerItem item })
        {
            ClearPageOrganizerPointerState();
            return;
        }

        PageOrganizerList.Focus();
        ClearPageOrganizerPointerState();
        ApplyPageOrganizerState(
            _pageOrganizerState.SelectPage(item.PageNumber, PageSelectionMode.Toggle),
            navigatePreview: true);
        e.Handled = true;
    }

    private void OnPageOrganizerThumbnailRetryClick(object sender, RoutedEventArgs e)
    {
        var scheduler = _pageOrganizerThumbnailScheduler;
        var cancellation = _pageOrganizerThumbnailCancellation;
        var generation = _pageOrganizerThumbnailGeneration;
        if (scheduler is null ||
            cancellation is null ||
            !IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
            sender is not FrameworkElement { DataContext: PageOrganizerItem item } ||
            !PageOrganizerItems.Contains(item) ||
            !scheduler.RequestManualRetry(item.PageNumber))
        {
            return;
        }

        if (!IsCurrentPageOrganizerThumbnailRequest(generation, cancellation) ||
            !ReferenceEquals(_pageOrganizerThumbnailScheduler, scheduler))
        {
            return;
        }

        item.Thumbnail = null;
        item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Pending;
        scheduler.Prioritize([item.PageNumber]);
        EnsurePageOrganizerThumbnailWorker(generation, cancellation);
        e.Handled = true;
    }

    private void OnPageOrganizerItemPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (IsDocumentMutationInProgress)
        {
            e.Handled = true;
            return;
        }

        if (_pageOrganizerState is null || e.OriginalSource is not DependencyObject source)
        {
            ClearPageOrganizerPointerState();
            return;
        }

        if (FindVisualAncestor<CheckBox>(source) is not null ||
            FindVisualAncestor<Button>(source) is not null)
        {
            return;
        }

        var item = FindPageOrganizerItem(source);
        if (item is null)
        {
            ClearPageOrganizerPointerState();
            return;
        }

        PageOrganizerList.Focus();
        var modifiers = Keyboard.Modifiers;
        var selectionMode = modifiers.HasFlag(ModifierKeys.Shift)
            ? PageSelectionMode.Range
            : modifiers.HasFlag(ModifierKeys.Control)
                ? PageSelectionMode.Toggle
                : PageSelectionMode.Replace;
        var preserveCurrentGroupForDrag = selectionMode == PageSelectionMode.Replace &&
                                          _pageOrganizerState.SelectedPageNumbers.Count > 1 &&
                                          _pageOrganizerState.SelectedPageNumbers.Contains(item.PageNumber);
        _pageOrganizerDragPageNumber = item.PageNumber;
        _pageOrganizerDragStartPosition = e.GetPosition(PageOrganizerList);
        _pageOrganizerPendingPlainSelectionPageNumber = preserveCurrentGroupForDrag
            ? item.PageNumber
            : null;
        Mouse.Capture(PageOrganizerList, CaptureMode.SubTree);
        if (!preserveCurrentGroupForDrag)
        {
            ApplyPageOrganizerState(
                _pageOrganizerState.SelectPage(item.PageNumber, selectionMode),
                navigatePreview: true);
        }

        e.Handled = true;
    }

    private void OnPageOrganizerItemPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (IsDocumentMutationInProgress)
        {
            ClearPageOrganizerPointerState();
            e.Handled = true;
            return;
        }

        if (_pageOrganizerPendingPlainSelectionPageNumber is { } pageNumber &&
            _pageOrganizerState is not null &&
            _pageOrganizerState.PageNumbers.Contains(pageNumber))
        {
            ApplyPageOrganizerState(
                _pageOrganizerState.SelectPage(pageNumber, PageSelectionMode.Replace),
                navigatePreview: true);
        }

        ClearPageOrganizerPointerState();
        e.Handled = true;
    }

    private void OnPageOrganizerItemPreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (IsDocumentMutationInProgress ||
            _pageOrganizerDragPageNumber is not { } pageNumber ||
            e.LeftButton != MouseButtonState.Pressed ||
            _pageOrganizerState is null ||
            !_pageOrganizerState.PageNumbers.Contains(pageNumber))
        {
            return;
        }

        if (_pageOrganizerDragStartPosition is not { } startPosition)
        {
            return;
        }

        var currentPosition = e.GetPosition(PageOrganizerList);
        var movedFarEnough = Math.Abs(currentPosition.X - startPosition.X) >= SystemParameters.MinimumHorizontalDragDistance ||
                               Math.Abs(currentPosition.Y - startPosition.Y) >= SystemParameters.MinimumVerticalDragDistance;
        if (!movedFarEnough)
        {
            return;
        }

        _pageOrganizerPendingPlainSelectionPageNumber = null;
        _isPageOrganizerDragInProgress = true;
        try
        {
            var data = new DataObject(PageOrganizerDragDataFormat, pageNumber);
            DragDrop.DoDragDrop(PageOrganizerList, data, DragDropEffects.Move);
        }
        finally
        {
            _isPageOrganizerDragInProgress = false;
            ClearPageOrganizerPointerState();
        }
    }

    private void OnPageOrganizerDragOver(object sender, DragEventArgs e)
    {
        var canMovePages = !IsDocumentMutationInProgress &&
                           e.Data.GetDataPresent(PageOrganizerDragDataFormat);
        e.Effects = canMovePages ? DragDropEffects.Move : DragDropEffects.None;
        if (canMovePages)
        {
            UpdatePageOrganizerDropIndicator(GetPageOrganizerInsertionIndex(e));
        }
        else
        {
            ClearPageOrganizerDropIndicator();
        }

        e.Handled = true;
    }

    private void OnPageOrganizerDragLeave(object sender, DragEventArgs e)
    {
        _ = Dispatcher.BeginInvoke(
            new Action(() =>
            {
                if (!PageOrganizerList.IsMouseOver)
                {
                    ClearPageOrganizerDropIndicator();
                }
            }),
            DispatcherPriority.Input);
    }

    private void OnPageOrganizerDrop(object sender, DragEventArgs e)
    {
        if (IsDocumentMutationInProgress ||
            _pageOrganizerState is null ||
            !e.Data.GetDataPresent(PageOrganizerDragDataFormat) ||
            e.Data.GetData(PageOrganizerDragDataFormat) is not int draggedPageNumber)
        {
            ClearPageOrganizerPointerState();
            return;
        }

        var insertionIndex = GetPageOrganizerInsertionIndex(e);
        ApplyPageOrganizerEdit(state => state.MovePageGroup(draggedPageNumber, insertionIndex));
        ClearPageOrganizerPointerState();
        e.Handled = true;
    }

    private void ClearPageOrganizerPointerState()
    {
        if (Mouse.Captured == PageOrganizerList)
        {
            Mouse.Capture(null);
        }

        _pageOrganizerDragPageNumber = null;
        _pageOrganizerDragStartPosition = null;
        _pageOrganizerPendingPlainSelectionPageNumber = null;
        ClearPageOrganizerDropIndicator();
    }

    private void UpdatePageOrganizerDropIndicator(int insertionIndex)
    {
        if (_pageOrganizerState is null || PageOrganizerItems.Count == 0)
        {
            ClearPageOrganizerDropIndicator();
            return;
        }

        var normalizedInsertionIndex = Math.Clamp(insertionIndex, 0, PageOrganizerItems.Count);
        _pageOrganizerDropInsertionIndex = normalizedInsertionIndex;
        for (var index = 0; index < PageOrganizerItems.Count; index++)
        {
            var item = PageOrganizerItems[index];
            item.IsDropBefore = normalizedInsertionIndex < PageOrganizerItems.Count &&
                                index == normalizedInsertionIndex;
            item.IsDropAfter = normalizedInsertionIndex == PageOrganizerItems.Count &&
                               index == PageOrganizerItems.Count - 1;
        }
    }

    private void ClearPageOrganizerDropIndicator()
    {
        _pageOrganizerDropInsertionIndex = null;
        foreach (var item in PageOrganizerItems)
        {
            item.IsDropBefore = false;
            item.IsDropAfter = false;
        }
    }

    private int GetPageOrganizerInsertionIndex(DragEventArgs e)
    {
        if (_pageOrganizerState is null)
        {
            return _pageOrganizerState?.PageNumbers.Count ?? 0;
        }

        var containers = new List<(int Index, FrameworkElement Container, Point Origin)>();
        for (var index = 0; index < _pageOrganizerState.PageNumbers.Count; index++)
        {
            if (PageOrganizerList.ItemContainerGenerator.ContainerFromIndex(index) is FrameworkElement container &&
                container.ActualWidth > 0 &&
                container.ActualHeight > 0)
            {
                containers.Add((index, container, container.TranslatePoint(new Point(), PageOrganizerList)));
            }
        }

        if (containers.Count == 0)
        {
            return _pageOrganizerState.PageNumbers.Count;
        }

        var dropPosition = e.GetPosition(PageOrganizerList);
        var isMultiColumn = containers.Count > 1 &&
                            Math.Abs(containers[0].Origin.Y - containers[1].Origin.Y) < 1;
        if (!isMultiColumn)
        {
            foreach (var entry in containers)
            {
                if (dropPosition.Y < entry.Origin.Y + entry.Container.ActualHeight / 2d)
                {
                    return entry.Index;
                }
            }

            return _pageOrganizerState.PageNumbers.Count;
        }

        var rows = containers
            .GroupBy(entry => Math.Round(entry.Origin.Y))
            .Select(group => group.OrderBy(entry => entry.Origin.X).ToArray())
            .ToArray();
        var row = rows
            .OrderBy(entries =>
            {
                var top = entries.Min(entry => entry.Origin.Y);
                var bottom = entries.Max(entry => entry.Origin.Y + entry.Container.ActualHeight);
                return dropPosition.Y < top
                    ? top - dropPosition.Y
                    : dropPosition.Y > bottom
                        ? dropPosition.Y - bottom
                        : 0d;
            })
            .First();

        foreach (var entry in row)
        {
            if (dropPosition.X < entry.Origin.X + entry.Container.ActualWidth / 2d)
            {
                return entry.Index;
            }
        }

        return row[^1].Index + 1;
    }

    private static PageOrganizerItem? FindPageOrganizerItem(DependencyObject source)
    {
        for (DependencyObject? current = source; current is not null; current = GetVisualParent(current))
        {
            if (current is FrameworkElement { DataContext: PageOrganizerItem item })
            {
                return item;
            }
        }

        return null;
    }

    private static T? FindVisualAncestor<T>(DependencyObject source)
        where T : DependencyObject
    {
        for (DependencyObject? current = source; current is not null; current = GetVisualParent(current))
        {
            if (current is T typed)
            {
                return typed;
            }
        }

        return null;
    }

    private static T? FindVisualDescendant<T>(DependencyObject root)
        where T : DependencyObject
    {
        for (var index = 0; index < VisualTreeHelper.GetChildrenCount(root); index++)
        {
            var child = VisualTreeHelper.GetChild(root, index);
            if (child is T typed)
            {
                return typed;
            }

            var descendant = FindVisualDescendant<T>(child);
            if (descendant is not null)
            {
                return descendant;
            }
        }

        return null;
    }

    private static DependencyObject? GetVisualParent(DependencyObject current)
    {
        try
        {
            return VisualTreeHelper.GetParent(current) ?? LogicalTreeHelper.GetParent(current);
        }
        catch (InvalidOperationException)
        {
            return LogicalTreeHelper.GetParent(current);
        }
    }

    private void OnWindowKeyDown(object sender, KeyEventArgs e)
    {
        HandleApplicationKeyDown(sender, e);
    }

    private void OnViewerPreviewKeyDown(object sender, KeyEventArgs e)
    {
        HandleApplicationKeyDown(sender, e);
    }

    private void OnPageOrganizerPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (!PageOrganizerList.IsKeyboardFocusWithin)
        {
            return;
        }

        if (TryHandlePageOrganizerNavigationKey(e))
        {
            e.Handled = true;
            return;
        }

        HandleApplicationKeyDown(sender, e);
    }

    private bool TryHandlePageOrganizerNavigationKey(KeyEventArgs e)
    {
        var navigationKey = e.Key is Key.Left or Key.Right or Key.Up or Key.Down or
            Key.PageUp or Key.PageDown or Key.Home or Key.End;
        if (!navigationKey)
        {
            return false;
        }

        if (Keyboard.Modifiers != ModifierKeys.None)
        {
            return true;
        }

        switch (e.Key)
        {
            case Key.Left:
            case Key.Up:
                NavigatePageOrganizer(-1);
                break;
            case Key.Right:
            case Key.Down:
                NavigatePageOrganizer(1);
                break;
            case Key.PageUp:
                NavigatePageOrganizer(-10);
                break;
            case Key.PageDown:
                NavigatePageOrganizer(10);
                break;
            case Key.Home:
                NavigatePageOrganizerBoundary(last: false);
                break;
            case Key.End:
                NavigatePageOrganizerBoundary(last: true);
                break;
        }

        return true;
    }

    private void HandleApplicationKeyDown(object sender, KeyEventArgs e)
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
            OnUndoClick(sender, e);
            e.Handled = true;
        }
        else if (control && !shift && e.Key == Key.Y)
        {
            OnRedoClick(sender, e);
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
            OnRotateClockwiseClick(sender, e);
            e.Handled = true;
        }
        else if (e.Key == Key.Delete)
        {
            OnDeleteSelectedPagesClick(sender, e);
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
            OnThumbZoomInClick(sender, e);
            e.Handled = true;
        }
        else if (control && shift && (e.Key == Key.OemMinus || e.Key == Key.Subtract))
        {
            OnThumbZoomOutClick(sender, e);
            e.Handled = true;
        }
        else if (control && shift && (e.Key == Key.D0 || e.Key == Key.NumPad0))
        {
            OnThumbZoomResetClick(sender, e);
            e.Handled = true;
        }
    }

    private void OnWindowDragOver(object sender, DragEventArgs e)
    {
        if (CanInsertDroppedFilesIntoCurrentDocument() &&
            TryGetDroppedInsertionPaths(e, out var insertionPaths))
        {
            e.Effects = DragDropEffects.Copy;
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
            _ = InsertExternalFilesAsync(new ExternalFilesDropMessage(
                insertionPaths.ToList(),
                GetInsertionIndexAfterSelection()));
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

public sealed class PageOrganizerItem : INotifyPropertyChanged
{
    private int _position;
    private int _rotation;
    private bool _isSelected;
    private bool _isActive;
    private bool _isDropBefore;
    private bool _isDropAfter;
    private ImageSource? _thumbnail;
    private PageOrganizerThumbnailRenderState _thumbnailRenderState;
    private double _thumbnailHeight = 142;
    private double _thumbnailWidth = 100;
    private double _thumbnailCardWidth = 116;

    public PageOrganizerItem(int pageNumber)
    {
        PageNumber = pageNumber;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public int PageNumber { get; }

    public int Position
    {
        get => _position;
        set => SetField(ref _position, value, nameof(Position), nameof(PositionLabel));
    }

    public int Rotation
    {
        get => _rotation;
        set => SetField(ref _rotation, value, nameof(Rotation), nameof(RotationLabel));
    }

    public bool IsSelected
    {
        get => _isSelected;
        set => SetField(ref _isSelected, value, nameof(IsSelected));
    }

    public bool IsActive
    {
        get => _isActive;
        set => SetField(ref _isActive, value, nameof(IsActive));
    }

    public bool IsDropBefore
    {
        get => _isDropBefore;
        set => SetField(ref _isDropBefore, value, nameof(IsDropBefore));
    }

    public bool IsDropAfter
    {
        get => _isDropAfter;
        set => SetField(ref _isDropAfter, value, nameof(IsDropAfter));
    }

    public ImageSource? Thumbnail
    {
        get => _thumbnail;
        set => SetField(ref _thumbnail, value, nameof(Thumbnail));
    }

    public PageOrganizerThumbnailRenderState ThumbnailRenderState
    {
        get => _thumbnailRenderState;
        set => SetField(
            ref _thumbnailRenderState,
            value,
            nameof(ThumbnailRenderState),
            nameof(IsThumbnailLoading),
            nameof(IsThumbnailFailed));
    }

    public bool IsThumbnailLoading =>
        ThumbnailRenderState is PageOrganizerThumbnailRenderState.Pending or
        PageOrganizerThumbnailRenderState.Loading or
        PageOrganizerThumbnailRenderState.Evicted;

    public bool IsThumbnailFailed =>
        ThumbnailRenderState == PageOrganizerThumbnailRenderState.Failed;

    public double ThumbnailHeight
    {
        get => _thumbnailHeight;
        set => SetField(ref _thumbnailHeight, value, nameof(ThumbnailHeight));
    }

    public double ThumbnailWidth
    {
        get => _thumbnailWidth;
        set => SetField(ref _thumbnailWidth, value, nameof(ThumbnailWidth));
    }

    public double ThumbnailCardWidth
    {
        get => _thumbnailCardWidth;
        set => SetField(ref _thumbnailCardWidth, value, nameof(ThumbnailCardWidth));
    }

    public string PositionLabel => $"{Position}";

    public string RotationLabel => Rotation == 0 ? string.Empty : $"{Rotation}°";

    private void SetField<T>(ref T field, T value, params string[] propertyNames)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return;
        }

        field = value;
        foreach (var propertyName in propertyNames)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
