# Lazy Thumbnail Loading Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (- [ ]) syntax for tracking.

**Goal:** Render every Page Organizer thumbnail for large PDFs without a fixed page cap, while keeping scrolling responsive, failures isolated, and memory bounded.

**Architecture:** Replace the one-shot first-96-page loop with a pure scheduler that owns pending, priority, retry, and cache-window decisions. MainWindow owns the WPF render worker, obtains visible page numbers from the existing ListBox/ScrollViewer, applies per-page visual state, and cancels stale work using the existing document-generation lifecycle. PdfFallbackRenderService remains the only native PDF rasterizer.

**Tech Stack:** .NET 8 WPF, PdfiumViewer, xUnit, Node.js contract tests, PowerShell release verification.

## Global Constraints

- Remove the fixed first-96-page rendering limit; every page must remain eligible for rendering.
- Render visible thumbnails and their nearby cache window before remaining background pages, and reprioritize after a sidebar scroll.
- Do not hard-cancel an in-flight Pdfium render merely because the user scrolls; allow it to finish or be discarded, then choose the newest priority work.
- Opening a new PDF, closing the window, or starting a replacement document operation must cancel stale thumbnail work through the existing generation and cancellation mechanism.
- A single page render failure must receive one automatic retry and never stop later pages. After a second failure, show a dedicated retry control for that page only.
- Keep only the current cache window of thumbnail ImageSource objects in memory. A distant page may be evicted and rendered again from the bounded Pdfium disk cache when it becomes relevant.
- Preserve Page Organizer ownership of Page Selection, Page Move Group, Edit History, drag/drop indicators, active-page follow, and keyboard navigation.
- A retry control must not change Page Selection or initiate a drag.
- Add no packages and do not change PDF.js structural editing ownership.
- Preserve the current responsive WrapPanel grid and thumbnail zoom behavior.

## File Structure

- Create: src/PdfMergeTool/Services/PageOrganizerThumbnailScheduler.cs - pure order, priority, retry, and cache-window policy.
- Create: test/PdfMergeTool.Tests/PageOrganizerThumbnailSchedulerTests.cs - xUnit behavior coverage for high-page priority, retry isolation, manual retry, and bounded windows.
- Modify: test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj - link the new pure service.
- Modify: src/PdfMergeTool/MainWindow.xaml.cs - cancellable WPF worker, visible-range scheduling, cache eviction, individual failure handling, and retry command.
- Modify: src/PdfMergeTool/MainWindow.xaml - loading indicator and retry icon within an existing thumbnail card.
- Modify: test/thumbnail-selection-contract.test.mjs - guard the no-cap scheduler, state, retry, and scroll paths.
- Modify: scripts/verify-page-organizer.ps1 - prevent release builds from reintroducing the first-96-page limit or removing failure/retry ownership.
- Modify: CONTEXT.md - retain the agreed thumbnail-loading vocabulary already captured during design.

---

### Task 1: Add The Pure Thumbnail Scheduler And Its Tests

**Files:**
- Create: src/PdfMergeTool/Services/PageOrganizerThumbnailScheduler.cs
- Create: test/PdfMergeTool.Tests/PageOrganizerThumbnailSchedulerTests.cs
- Modify: test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj

**Interfaces:**
- Produces: PageOrganizerThumbnailRenderState with Pending, Loading, Ready, Failed, and Evicted values.
- Produces: PageOrganizerThumbnailScheduler(IReadOnlyList<int> pageNumbers).
- Produces: Prioritize(IEnumerable<int>), Request(int, bool), TryTakeNext(out int), Complete(int), RegisterFailure(int), RequestManualRetry(int), and GetCacheWindow(IReadOnlyCollection<int>).
- Consumed by: MainWindow thumbnail worker and PageOrganizerItem state in Task 2.

- [ ] **Step 1: Write the failing scheduler tests**

Create test/PdfMergeTool.Tests/PageOrganizerThumbnailSchedulerTests.cs:

~~~csharp
using PdfMergeTool.Services;
using Xunit;

namespace PdfMergeTool.Tests;

public sealed class PageOrganizerThumbnailSchedulerTests
{
    [Fact]
    public void TryTakeNext_prioritizes_visible_high_pages_before_background_pages()
    {
        var scheduler = new PageOrganizerThumbnailScheduler(Enumerable.Range(1, 166).ToArray());
        scheduler.Prioritize([97, 98, 99]);

        Assert.True(scheduler.TryTakeNext(out var first));
        scheduler.Complete(first);
        Assert.True(scheduler.TryTakeNext(out var second));
        scheduler.Complete(second);
        Assert.True(scheduler.TryTakeNext(out var third));

        Assert.Equal(new[] { 97, 98, 99 }, new[] { first, second, third });
    }

    [Fact]
    public void RegisterFailure_retries_once_then_leaves_the_failed_page_out_of_the_queue()
    {
        var scheduler = new PageOrganizerThumbnailScheduler([1, 2, 3]);

        Assert.True(scheduler.TryTakeNext(out var firstAttempt));
        Assert.Equal(1, firstAttempt);
        Assert.True(scheduler.RegisterFailure(firstAttempt));
        Assert.True(scheduler.TryTakeNext(out var retryAttempt));
        Assert.Equal(1, retryAttempt);
        Assert.False(scheduler.RegisterFailure(retryAttempt));
        Assert.True(scheduler.TryTakeNext(out var nextPage));

        Assert.Equal(2, nextPage);
    }

    [Fact]
    public void RequestManualRetry_restores_a_page_after_its_automatic_retry_is_exhausted()
    {
        var scheduler = new PageOrganizerThumbnailScheduler([1, 2]);

        Assert.True(scheduler.TryTakeNext(out var page));
        Assert.True(scheduler.RegisterFailure(page));
        Assert.True(scheduler.TryTakeNext(out page));
        Assert.False(scheduler.RegisterFailure(page));
        Assert.True(scheduler.RequestManualRetry(page));
        Assert.True(scheduler.TryTakeNext(out var retriedPage));

        Assert.Equal(1, retriedPage);
    }

    [Fact]
    public void GetCacheWindow_keeps_a_bounded_range_around_visible_pages()
    {
        var scheduler = new PageOrganizerThumbnailScheduler(Enumerable.Range(1, 200).ToArray());

        var window = scheduler.GetCacheWindow([97, 98, 99, 100]);

        Assert.Equal(52, window.Count);
        Assert.Contains(73, window);
        Assert.Contains(124, window);
        Assert.DoesNotContain(72, window);
        Assert.DoesNotContain(125, window);
    }
}
~~~

Add this linked source entry beside the existing pure service links:

~~~xml
<Compile Include="../../src/PdfMergeTool/Services/PageOrganizerThumbnailScheduler.cs"
         Link="Services/PageOrganizerThumbnailScheduler.cs" />
~~~

- [ ] **Step 2: Run the focused test to verify it fails**

Run:

~~~bash
dotnet test test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj --filter FullyQualifiedName~PageOrganizerThumbnailSchedulerTests
~~~

Expected: FAIL because PageOrganizerThumbnailScheduler does not exist.

- [ ] **Step 3: Implement the scheduler**

Create src/PdfMergeTool/Services/PageOrganizerThumbnailScheduler.cs:

~~~csharp
namespace PdfMergeTool.Services;

public enum PageOrganizerThumbnailRenderState
{
    Pending,
    Loading,
    Ready,
    Failed,
    Evicted
}

public sealed class PageOrganizerThumbnailScheduler
{
    private const int MinimumCacheRadius = 24;
    private const int MaximumRenderAttempts = 2;
    private readonly int[] _pageNumbers;
    private readonly HashSet<int> _knownPages;
    private readonly Dictionary<int, int> _pageIndexes;
    private readonly HashSet<int> _pendingPages;
    private readonly HashSet<int> _inFlightPages = [];
    private readonly LinkedList<int> _priorityPages = [];
    private readonly HashSet<int> _priorityPageSet = [];
    private readonly Dictionary<int, int> _failureCounts = [];

    public PageOrganizerThumbnailScheduler(IReadOnlyList<int> pageNumbers)
    {
        _pageNumbers = pageNumbers.Distinct().ToArray();
        _knownPages = _pageNumbers.ToHashSet();
        _pageIndexes = _pageNumbers
            .Select((pageNumber, index) => new { pageNumber, index })
            .ToDictionary(entry => entry.pageNumber, entry => entry.index);
        _pendingPages = _pageNumbers.ToHashSet();
    }

    public void Prioritize(IEnumerable<int> pageNumbers)
    {
        foreach (var pageNumber in pageNumbers)
        {
            if (_pendingPages.Contains(pageNumber) && _priorityPageSet.Add(pageNumber))
            {
                _priorityPages.AddLast(pageNumber);
            }
        }
    }

    public bool Request(int pageNumber, bool priority)
    {
        if (!_knownPages.Contains(pageNumber) || _inFlightPages.Contains(pageNumber))
        {
            return false;
        }

        var added = _pendingPages.Add(pageNumber);
        if (priority)
        {
            Prioritize([pageNumber]);
        }

        return added;
    }

    public bool TryTakeNext(out int pageNumber)
    {
        while (_priorityPages.First is { } priority)
        {
            _priorityPages.RemoveFirst();
            _priorityPageSet.Remove(priority.Value);
            if (TryStart(priority.Value, out pageNumber))
            {
                return true;
            }
        }

        foreach (var candidate in _pageNumbers)
        {
            if (TryStart(candidate, out pageNumber))
            {
                return true;
            }
        }

        pageNumber = default;
        return false;
    }

    public void Complete(int pageNumber)
    {
        _inFlightPages.Remove(pageNumber);
        _failureCounts.Remove(pageNumber);
    }

    public bool RegisterFailure(int pageNumber)
    {
        _inFlightPages.Remove(pageNumber);
        var failures = _failureCounts.TryGetValue(pageNumber, out var existing)
            ? existing + 1
            : 1;
        _failureCounts[pageNumber] = failures;
        return failures < MaximumRenderAttempts && Request(pageNumber, priority: true);
    }

    public bool RequestManualRetry(int pageNumber)
    {
        _failureCounts.Remove(pageNumber);
        return Request(pageNumber, priority: true);
    }

    public IReadOnlySet<int> GetCacheWindow(IReadOnlyCollection<int> visiblePageNumbers)
    {
        var visibleIndexes = visiblePageNumbers
            .Where(_pageIndexes.ContainsKey)
            .Select(pageNumber => _pageIndexes[pageNumber])
            .Order()
            .ToArray();
        if (visibleIndexes.Length == 0)
        {
            return _pageNumbers.Take(MinimumCacheRadius).ToHashSet();
        }

        var radius = Math.Max(MinimumCacheRadius, visibleIndexes.Length * 2);
        var firstIndex = Math.Max(0, visibleIndexes[0] - radius);
        var lastIndex = Math.Min(_pageNumbers.Length - 1, visibleIndexes[^1] + radius);
        return _pageNumbers
            .Skip(firstIndex)
            .Take(lastIndex - firstIndex + 1)
            .ToHashSet();
    }

    private bool TryStart(int pageNumber, out int nextPageNumber)
    {
        if (_pendingPages.Remove(pageNumber))
        {
            _inFlightPages.Add(pageNumber);
            nextPageNumber = pageNumber;
            return true;
        }

        nextPageNumber = default;
        return false;
    }
}
~~~

- [ ] **Step 4: Run the focused test to verify it passes**

Run:

~~~bash
dotnet test test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj --filter FullyQualifiedName~PageOrganizerThumbnailSchedulerTests
~~~

Expected: PASS with four scheduler tests.

- [ ] **Step 5: Commit the pure scheduler**

~~~bash
git add src/PdfMergeTool/Services/PageOrganizerThumbnailScheduler.cs test/PdfMergeTool.Tests/PageOrganizerThumbnailSchedulerTests.cs test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj
git commit -m "feat: schedule lazy organizer thumbnails"
~~~

### Task 2: Integrate Cancellable Rendering, Priority, Cache Eviction, And Retry

**Files:**
- Modify: src/PdfMergeTool/MainWindow.xaml.cs:44-78, 530-590, 780-860, 1240-1255, 3270-3320, 3840-3995
- Modify: test/thumbnail-selection-contract.test.mjs

**Interfaces:**
- Consumes: PageOrganizerThumbnailScheduler and PageOrganizerThumbnailRenderState from Task 1.
- Produces: OnPageOrganizerThumbnailScrollChanged, RefreshPageOrganizerThumbnailViewport, OnPageOrganizerThumbnailRetryClick, and a render worker that removes the fixed page cap.
- Preserves: CancelPageOrganizerThumbnailRendering, DocumentOperationCoordinator, PageOrganizer selection/drag handling, and PdfFallbackRenderService.RenderPageAsync.

- [ ] **Step 1: Add failing static integration assertions**

Add these assertions after the existing Page Organizer follow assertions in test/thumbnail-selection-contract.test.mjs:

~~~javascript
assert.match(mainWindow, /PageOrganizerThumbnailScheduler/);
assert.match(mainWindow, /OnPageOrganizerThumbnailScrollChanged/);
assert.match(mainWindow, /RefreshPageOrganizerThumbnailViewport/);
assert.match(mainWindow, /OnPageOrganizerThumbnailRetryClick/);
assert.match(mainWindow, /PageOrganizerThumbnailRenderState.Failed/);
assert.doesNotMatch(mainWindow, /pageNumbers.Take(96)/);
~~~

- [ ] **Step 2: Run the contract test to verify it fails**

Run:

~~~bash
node test/thumbnail-selection-contract.test.mjs
~~~

Expected: FAIL at PageOrganizerThumbnailScheduler because the old renderer only contains pageNumbers.Take(96).

- [ ] **Step 3: Add worker state and scheduler lifecycle**

Add fields beside the current organizer thumbnail cancellation fields:

~~~csharp
private PageOrganizerThumbnailScheduler? _pageOrganizerThumbnailScheduler;
private HashSet<int> _pageOrganizerThumbnailCacheWindow = [];
private string? _pageOrganizerThumbnailSourcePath;
private int? _pageOrganizerThumbnailWorkerGeneration;
~~~

Replace the fixed-cap StartPageOrganizerThumbnailRendering and RenderPageOrganizerThumbnailsAsync flow with these ownership rules:

~~~csharp
private void StartPageOrganizerThumbnailRendering(string path, IReadOnlyList<int> pageNumbers)
{
    CancelPageOrganizerThumbnailRendering();
    var cancellation = new CancellationTokenSource();
    _pageOrganizerThumbnailCancellation = cancellation;
    _pageOrganizerThumbnailSourcePath = path;
    _pageOrganizerThumbnailScheduler = new PageOrganizerThumbnailScheduler(pageNumbers);
    _pageOrganizerThumbnailCacheWindow = _pageOrganizerThumbnailScheduler
        .GetCacheWindow([]);
    var generation = ++_pageOrganizerThumbnailGeneration;

    foreach (var item in PageOrganizerItems)
    {
        item.Thumbnail = null;
        item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Pending;
    }

    _pageOrganizerThumbnailScheduler.Prioritize(_pageOrganizerThumbnailCacheWindow);
    EnsurePageOrganizerThumbnailWorker(generation, cancellation);
    _ = Dispatcher.BeginInvoke(
        new Action(RefreshPageOrganizerThumbnailViewport),
        DispatcherPriority.Loaded);
}
~~~

Implement the worker methods exactly as follows. The worker is intentionally single-threaded because PdfFallbackRenderService already serializes access to one Pdfium session; prioritization takes effect after the current render completes.

~~~csharp
private void EnsurePageOrganizerThumbnailWorker(
    int generation,
    CancellationTokenSource cancellation)
{
    if (_pageOrganizerThumbnailWorkerGeneration == generation ||
        generation != _pageOrganizerThumbnailGeneration ||
        _pageOrganizerThumbnailScheduler is null ||
        string.IsNullOrWhiteSpace(_pageOrganizerThumbnailSourcePath) ||
        cancellation.IsCancellationRequested)
    {
        return;
    }

    _pageOrganizerThumbnailWorkerGeneration = generation;
    _ = RenderPageOrganizerThumbnailsAsync(
        _pageOrganizerThumbnailSourcePath,
        generation,
        cancellation);
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
        var session = await Task.Run(
            () => _pageOrganizerRenderService.OpenDocument(path),
            cancellationToken);
        sessionId = session.SessionId;

        while (generation == _pageOrganizerThumbnailGeneration &&
               !cancellationToken.IsCancellationRequested &&
               _pageOrganizerThumbnailScheduler is { } scheduler &&
               scheduler.TryTakeNext(out var pageNumber))
        {
            var item = PageOrganizerItems.FirstOrDefault(
                candidate => candidate.PageNumber == pageNumber);
            if (item is null)
            {
                scheduler.Complete(pageNumber);
                continue;
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
                if (generation != _pageOrganizerThumbnailGeneration)
                {
                    return;
                }

                scheduler.Complete(pageNumber);
                if (_pageOrganizerThumbnailCacheWindow.Contains(pageNumber))
                {
                    item.Thumbnail = LoadPageOrganizerThumbnail(rendered.ImagePath);
                    item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Ready;
                }
                else
                {
                    item.Thumbnail = null;
                    item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Evicted;
                }
            }
            catch (OperationCanceledException)
            {
                scheduler.Complete(pageNumber);
                throw;
            }
            catch (Exception ex)
            {
                if (generation != _pageOrganizerThumbnailGeneration ||
                    cancellationToken.IsCancellationRequested)
                {
                    return;
                }

                AppLogger.Error(
                    ex,
                    $"Page Organizer thumbnail rendering failed: {pageNumber} in {path}");
                var retryScheduled = scheduler.RegisterFailure(pageNumber);
                item.Thumbnail = null;
                item.ThumbnailRenderState = retryScheduled
                    ? PageOrganizerThumbnailRenderState.Pending
                    : PageOrganizerThumbnailRenderState.Failed;
            }
        }
    }
    catch (OperationCanceledException)
    {
        // A newer document or document mutation replaced this thumbnail request.
    }
    catch (Exception ex)
    {
        AppLogger.Error(ex, $"Page Organizer thumbnail session failed: {path}");
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
~~~

Update CancelPageOrganizerThumbnailRendering so it invalidates the existing generation, moves the current cancellation source out of the active field, calls Cancel without disposing it synchronously, and clears source, scheduler, and cache-window references. The old worker disposes its own source in its finally block only after it has observed that it is no longer active. Do not clear a newer worker-generation marker or close a session owned by a newer generation.

~~~csharp
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
~~~

At the beginning of the existing isBusy branch in SetDocumentMutationUiState, call CancelPageOrganizerThumbnailRendering before disabling input. After the existing mutation-generation guard and input re-enable path in the isBusy false branch, call this helper:

~~~csharp
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
~~~

This resumes only an unchanged document whose prior thumbnail work was cancelled. A successful replacement load initializes its own scheduler and therefore makes the helper a no-op.

- [ ] **Step 4: Prioritize visible pages and bound in-memory images**

Register the existing PageOrganizerList for ScrollViewer.ScrollChanged in the MainWindow constructor:

~~~csharp
PageOrganizerList.AddHandler(
    ScrollViewer.ScrollChangedEvent,
    new ScrollChangedEventHandler(OnPageOrganizerThumbnailScrollChanged));
~~~

Implement the event handler and visible-page reader exactly as follows:

~~~csharp
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
~~~

Implement RefreshPageOrganizerThumbnailViewport with this exact policy:

~~~csharp
private void RefreshPageOrganizerThumbnailViewport()
{
    if (_pageOrganizerThumbnailScheduler is null)
    {
        return;
    }

    var visiblePageNumbers = GetVisiblePageOrganizerThumbnailNumbers();
    var cacheWindow = _pageOrganizerThumbnailScheduler.GetCacheWindow(visiblePageNumbers);
    _pageOrganizerThumbnailCacheWindow = cacheWindow.ToHashSet();

    foreach (var item in PageOrganizerItems)
    {
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
            _pageOrganizerThumbnailScheduler.Request(item.PageNumber, priority: true);
        }
    }

    _pageOrganizerThumbnailScheduler.Prioritize(cacheWindow);
    EnsurePageOrganizerThumbnailWorker(
        _pageOrganizerThumbnailGeneration,
        _pageOrganizerThumbnailCancellation!);
}
~~~

Keep the existing document-generation guard around every item update. A new load or document mutation must not alter the old PageOrganizerItems.

- [ ] **Step 5: Add a retry command that preserves selection**

Extend PageOrganizerItem with:

~~~csharp
private PageOrganizerThumbnailRenderState _thumbnailRenderState;

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
~~~

Add this retry handler:

~~~csharp
private void OnPageOrganizerThumbnailRetryClick(object sender, RoutedEventArgs e)
{
    if (_pageOrganizerThumbnailScheduler is null ||
        sender is not FrameworkElement { DataContext: PageOrganizerItem item } ||
        !_pageOrganizerThumbnailScheduler.RequestManualRetry(item.PageNumber))
    {
        return;
    }

    item.Thumbnail = null;
    item.ThumbnailRenderState = PageOrganizerThumbnailRenderState.Pending;
    _pageOrganizerThumbnailScheduler.Prioritize([item.PageNumber]);
    EnsurePageOrganizerThumbnailWorker(
        _pageOrganizerThumbnailGeneration,
        _pageOrganizerThumbnailCancellation!);
    e.Handled = true;
}
~~~

In OnPageOrganizerItemPreviewMouseLeftButtonDown, treat a Button ancestor exactly like the existing CheckBox ancestor and return before selection/pointer logic. This keeps a retry click from changing Page Selection or beginning a drag.

- [ ] **Step 6: Run focused contract and source checks**

Run:

~~~bash
node test/thumbnail-selection-contract.test.mjs
git diff --check
~~~

Expected: Node contract passes and no whitespace errors are reported.

- [ ] **Step 7: Commit the rendering integration**

~~~bash
git add src/PdfMergeTool/MainWindow.xaml.cs test/thumbnail-selection-contract.test.mjs
git commit -m "feat: load organizer thumbnails lazily"
~~~

### Task 3: Show Loading And Retry States And Lock The Release Contract

**Files:**
- Modify: src/PdfMergeTool/MainWindow.xaml:877-1035
- Modify: scripts/verify-page-organizer.ps1:73-135

**Interfaces:**
- Consumes: PageOrganizerItem.IsThumbnailLoading, IsThumbnailFailed, and OnPageOrganizerThumbnailRetryClick from Task 2.
- Produces: visible loading/retry UI and release-time guards against restoring pageNumbers.Take(96).

- [ ] **Step 1: Add failing XAML and release-contract assertions**

Add these Node assertions to test/thumbnail-selection-contract.test.mjs:

~~~javascript
assert.match(xaml, /OnPageOrganizerThumbnailRetryClick/);
assert.match(xaml, /Binding="{Binding IsThumbnailLoading}"/);
assert.match(xaml, /Binding="{Binding IsThumbnailFailed}"/);
~~~

Extend scripts/verify-page-organizer.ps1 with these MainWindow fragments:

~~~powershell
$mainWindow -match 'pageNumbers.Take(96)' -or
$mainWindow -notmatch 'PageOrganizerThumbnailScheduler' -or
$mainWindow -notmatch 'OnPageOrganizerThumbnailScrollChanged' -or
$mainWindow -notmatch 'RefreshPageOrganizerThumbnailViewport' -or
$mainWindow -notmatch 'OnPageOrganizerThumbnailRetryClick' -or
$mainWindow -notmatch 'PageOrganizerThumbnailRenderState.Failed'
~~~

Extend the XAML condition with:

~~~powershell
$xaml -notmatch 'Click="OnPageOrganizerThumbnailRetryClick"' -or
$xaml -notmatch 'Binding="{Binding IsThumbnailLoading}"' -or
$xaml -notmatch 'Binding="{Binding IsThumbnailFailed}"'
~~~

- [ ] **Step 2: Run the Node contract to verify it fails**

Run:

~~~bash
node test/thumbnail-selection-contract.test.mjs
~~~

Expected: FAIL because the thumbnail template does not yet expose loading and retry state.

- [ ] **Step 3: Add the thumbnail visual states**

Within the existing thumbnail Border Grid in MainWindow.xaml, keep the Image binding and add these two overlays:

~~~xml
<ProgressBar Width="34"
             Height="3"
             IsIndeterminate="True"
             VerticalAlignment="Center"
             HorizontalAlignment="Center"
             IsHitTestVisible="False">
    <ProgressBar.Style>
        <Style TargetType="{x:Type ProgressBar}">
            <Setter Property="Visibility" Value="Collapsed" />
            <Style.Triggers>
                <DataTrigger Binding="{Binding IsThumbnailLoading}" Value="True">
                    <Setter Property="Visibility" Value="Visible" />
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </ProgressBar.Style>
</ProgressBar>
<Button Width="28"
        Height="28"
        Padding="0"
        ToolTip="썸네일 다시 불러오기"
        Background="#FFF7ED"
        BorderBrush="#F59E0B"
        Click="OnPageOrganizerThumbnailRetryClick">
    <Button.Style>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Visibility" Value="Collapsed" />
            <Style.Triggers>
                <DataTrigger Binding="{Binding IsThumbnailFailed}" Value="True">
                    <Setter Property="Visibility" Value="Visible" />
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </Button.Style>
    <iconPacks:PackIconMaterial Kind="Refresh"
                                Width="16"
                                Height="16"
                                Foreground="#B45309" />
</Button>
~~~

The progress indicator must be non-hit-testable. The retry button must be the only new hit-testable control and uses the existing Material icon library instead of a text command.

- [ ] **Step 4: Run all available verification**

Run:

~~~bash
git diff --check
dotnet test test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj
node test/thumbnail-selection-contract.test.mjs
pwsh scripts/verify-page-organizer.ps1
pwsh scripts/verify-stability.ps1
~~~

Expected: all commands pass. On this Mac, record the unavailable dotnet and pwsh commands and require the existing Windows release workflow to run xUnit, both PowerShell scripts, installer build, and packaged-startup smoke verification.

- [ ] **Step 5: Execute the Windows manual smoke sequence**

Use a PDF with at least 166 pages:

1. Open the document and confirm pages 97 through 166 become real thumbnails rather than permanent PDF placeholders.
2. Scroll quickly from page 1 to around page 120 and confirm the nearby page range loads before unrelated background pages.
3. Leave the sidebar idle and confirm background work progressively completes the rest of the document without blocking preview navigation or editing.
4. Confirm deterministic retry behavior through the PageOrganizerThumbnailScheduler xUnit test; when a known Pdfium-failing source page is available, also confirm the retry icon and successful manual retry without losing Page Selection.
5. Revisit a distant page after scrolling elsewhere and confirm its image is restored without accumulating every thumbnail ImageSource in memory.
6. Open a second PDF and start a document mutation while work is active; confirm no stale thumbnail is assigned to the new document and no old worker survives.
7. Recheck Ctrl/Shift selection, Page Move Group drag/drop indicators, active-page follow, keyboard navigation, undo, redo, save, and the thumbnail zoom slider.

- [ ] **Step 6: Commit UI and release guards**

~~~bash
git add src/PdfMergeTool/MainWindow.xaml scripts/verify-page-organizer.ps1 test/thumbnail-selection-contract.test.mjs CONTEXT.md
git commit -m "feat: show thumbnail loading and retry state"
~~~
