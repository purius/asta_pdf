# Page Organizer Follow And Keyboard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make Page Organizer follow the preview's active page only when needed and provide focused keyboard navigation plus existing page-operation shortcuts.

**Architecture:** Keep `EditorDocumentState` as the sole owner of selection, order, rotation, and history. Add a small pure viewport-offset calculator for unit-tested geometry, then have `MainWindow` coalesce preview updates and perform WPF visual-tree scrolling only after layout. Attach keyboard handling to the focused Page Organizer list and reuse existing navigation and page-operation methods.

**Tech Stack:** .NET 8, WPF, WebView2/PDF.js preview adapter, xUnit, Node.js contract tests, PowerShell release verification.

## Global Constraints

- Preserve the responsive multi-column `WrapPanel` and existing drag/drop insertion indicators.
- Page Organizer remains the sole authority for Page Selection, page order, rotations, and Edit History.
- Active Page Follow scrolls only when the active thumbnail is outside the visible viewport; it never centers visible items.
- Coalesce rapid preview changes to the newest active page after layout.
- Suppress Active Page Follow during a Page Move Group drag, during a Document Mutation, and for stale document generations.
- Keyboard navigation acts only while Page Organizer owns keyboard focus.
- Every arrow key moves the active page one position in current document order. `PageUp`/`PageDown` move 10 positions. `Home`/`End` move to the first/last position.
- Keyboard navigation preserves Page Selection and does not create Edit History. Do not add Shift-arrow range selection in this work.
- `Ctrl+C`, `Ctrl+X`, `Ctrl+V`, `Ctrl+Z`, `Ctrl+Y`, and `Delete` must call the existing Page Organizer operations.
- Add no package dependencies and do not change the PDF.js structural-edit ownership boundary.

## File Structure

- Create: `src/PdfMergeTool/Services/PageOrganizerViewport.cs` - pure viewport offset calculation with no WPF dependency.
- Create: `test/PdfMergeTool.Tests/PageOrganizerViewportTests.cs` - xUnit coverage for visible, above, below, and clamped viewport cases.
- Modify: `test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj` - link the new pure service into the existing non-WPF test project.
- Modify: `src/PdfMergeTool/MainWindow.xaml` - make Page Organizer focusable and attach its preview key handler.
- Modify: `src/PdfMergeTool/MainWindow.xaml.cs` - coalesced follow queue, visual-tree scroll bridge, drag guard, focus transfer, and keyboard routing.
- Modify: `test/thumbnail-selection-contract.test.mjs` - require the new WPF ownership and keyboard/follow paths.
- Modify: `scripts/verify-page-organizer.ps1` - fail release verification if the follow or focused-keyboard contract disappears.

---

### Task 1: Add A Pure Page Organizer Viewport Calculator

**Files:**
- Create: `src/PdfMergeTool/Services/PageOrganizerViewport.cs`
- Create: `test/PdfMergeTool.Tests/PageOrganizerViewportTests.cs`
- Modify: `test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj`

**Interfaces:**
- Produces: `PdfMergeTool.Services.PageOrganizerViewport.GetVerticalOffsetToReveal(double currentOffset, double viewportHeight, double itemTop, double itemBottom, double scrollableHeight)`, returning `double?`.
- Consumed by: `MainWindow.FollowActivePageOrganizerItem` in Task 2.

- [ ] **Step 1: Write the failing viewport tests**

Create `test/PdfMergeTool.Tests/PageOrganizerViewportTests.cs`:

```csharp
using PdfMergeTool.Services;
using Xunit;

namespace PdfMergeTool.Tests;

public sealed class PageOrganizerViewportTests
{
    [Fact]
    public void GetVerticalOffsetToReveal_returns_null_when_item_is_fully_visible()
    {
        var result = PageOrganizerViewport.GetVerticalOffsetToReveal(100, 300, 20, 140, 900);

        Assert.Null(result);
    }

    [Fact]
    public void GetVerticalOffsetToReveal_moves_only_the_amount_needed_below_viewport()
    {
        var result = PageOrganizerViewport.GetVerticalOffsetToReveal(100, 300, 280, 340, 900);

        Assert.Equal(140d, result);
    }

    [Fact]
    public void GetVerticalOffsetToReveal_moves_only_the_amount_needed_above_viewport()
    {
        var result = PageOrganizerViewport.GetVerticalOffsetToReveal(100, 300, -18, 82, 900);

        Assert.Equal(82d, result);
    }

    [Fact]
    public void GetVerticalOffsetToReveal_clamps_to_scrollable_height()
    {
        var result = PageOrganizerViewport.GetVerticalOffsetToReveal(900, 200, 180, 250, 920);

        Assert.Equal(920d, result);
    }
}
```

Add this linked source entry to `test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj` beside the other service links:

```xml
<Compile Include="../../src/PdfMergeTool/Services/PageOrganizerViewport.cs"
         Link="Services/PageOrganizerViewport.cs" />
```

- [ ] **Step 2: Run the focused test to verify it fails**

Run:

```bash
dotnet test test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj --filter FullyQualifiedName~PageOrganizerViewportTests
```

Expected: FAIL because `PageOrganizerViewport` does not exist.

- [ ] **Step 3: Implement the pure calculator**

Create `src/PdfMergeTool/Services/PageOrganizerViewport.cs`:

```csharp
namespace PdfMergeTool.Services;

public static class PageOrganizerViewport
{
    public static double? GetVerticalOffsetToReveal(
        double currentOffset,
        double viewportHeight,
        double itemTop,
        double itemBottom,
        double scrollableHeight)
    {
        if (viewportHeight <= 0 || itemBottom <= itemTop || scrollableHeight <= 0)
        {
            return null;
        }

        var targetOffset = itemTop < 0
            ? currentOffset + itemTop
            : itemBottom > viewportHeight
                ? currentOffset + itemBottom - viewportHeight
                : currentOffset;
        targetOffset = Math.Clamp(targetOffset, 0, scrollableHeight);

        return Math.Abs(targetOffset - currentOffset) < 0.1 ? null : targetOffset;
    }
}
```

- [ ] **Step 4: Run the focused test to verify it passes**

Run:

```bash
dotnet test test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj --filter FullyQualifiedName~PageOrganizerViewportTests
```

Expected: PASS with four viewport tests.

- [ ] **Step 5: Commit the independently testable calculator**

```bash
git add src/PdfMergeTool/Services/PageOrganizerViewport.cs test/PdfMergeTool.Tests/PageOrganizerViewportTests.cs test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj
git commit -m "test: cover organizer viewport follow calculation"
```

### Task 2: Integrate Coalesced Follow And Focused Keyboard Commands

**Files:**
- Modify: `src/PdfMergeTool/MainWindow.xaml:877-893`
- Modify: `src/PdfMergeTool/MainWindow.xaml.cs:45-86, 384-398, 498-558, 3137-3290, 3477-3607`
- Uses: `src/PdfMergeTool/Services/PageOrganizerViewport.cs`

**Interfaces:**
- Consumes: `PageOrganizerViewport.GetVerticalOffsetToReveal(...)` from Task 1.
- Produces: `QueueActivePageFollow(int pageNumber)`, `FollowActivePageOrganizerItem(int pageNumber)`, and `OnPageOrganizerPreviewKeyDown(object sender, KeyEventArgs e)` in `MainWindow`.
- Preserves: `NavigatePageOrganizer(int delta)`, `NavigatePageOrganizerBoundary(bool last)`, `CopySelectedPagesToClipboardAsync`, `PastePagesOrImageAsync`, `OnUndoClick`, `OnRedoClick`, and `OnDeleteSelectedPagesClick`.

- [ ] **Step 1: Add a failing contract assertion for the focused paths**

Add these assertions to `test/thumbnail-selection-contract.test.mjs` immediately after the current Page Organizer assertions:

```javascript
assert.match(mainWindow, /QueueActivePageFollow\(/);
assert.match(mainWindow, /FollowActivePageOrganizerItem\(/);
assert.match(mainWindow, /PageOrganizerViewport\.GetVerticalOffsetToReveal/);
assert.match(mainWindow, /OnPageOrganizerPreviewKeyDown/);
assert.match(mainWindow, /PageOrganizerList\.Focus\(\)/);
assert.match(mainWindow, /Key\.PageUp/);
assert.match(mainWindow, /Key\.PageDown/);
assert.match(mainWindow, /Key\.Home/);
assert.match(mainWindow, /Key\.End/);
assert.match(xaml, /x:Name="PageOrganizerList"[\s\S]*?Focusable="True"[\s\S]*?PreviewKeyDown="OnPageOrganizerPreviewKeyDown"/);
```

- [ ] **Step 2: Run the contract test to verify it fails**

Run:

```bash
node test/thumbnail-selection-contract.test.mjs
```

Expected: FAIL because the follow and focused-keyboard symbols are absent.

- [ ] **Step 3: Give Page Organizer focus and route its keys**

In `MainWindow.xaml`, replace the Page Organizer `Focusable="False"` setting and add the preview handler:

```xml
Focusable="True"
IsTabStop="True"
PreviewKeyDown="OnPageOrganizerPreviewKeyDown"
```

In `MainWindow.xaml.cs`, call `PageOrganizerList.Focus()` in both thumbnail click paths before applying their existing selection action. Add this handler and keep modified navigation keys from falling through to `ListBox` defaults:

```csharp
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
```

In each of the existing checkbox and thumbnail mouse-down handlers, add this one statement without changing the current selection logic:

~~~csharp
PageOrganizerList.Focus();
~~~

Place it immediately before `ClearPageOrganizerPointerState();` in `OnPageOrganizerCheckBoxPreviewMouseLeftButtonDown`, and immediately before `var modifiers = Keyboard.Modifiers;` in `OnPageOrganizerItemPreviewMouseLeftButtonDown`. Those locations ensure a real thumbnail click owns keyboard focus while clicks on empty space and guarded mutation paths remain unchanged.

This preserves selected pages because `NavigatePageOrganizer` already uses `EditorDocumentState.ActivatePage`.

- [ ] **Step 4: Add the coalesced follow queue and viewport bridge**

Add these fields beside the existing Page Organizer pointer state:

```csharp
private int? _pendingPageOrganizerFollowPage;
private int _pendingPageOrganizerFollowLoadGeneration;
private int _pageOrganizerFollowRequestId;
private bool _pageOrganizerFollowQueued;
private bool _isPageOrganizerDragInProgress;
```

At the start of `LoadPdf`, clear follow state after assigning `_documentLoadGeneration`:

```csharp
ResetActivePageFollow();
```

Add the reset method beside the queue so changing documents invalidates both a queued dispatcher callback and its pending page:

~~~csharp
private void ResetActivePageFollow()
{
    _pendingPageOrganizerFollowPage = null;
    _pendingPageOrganizerFollowLoadGeneration = 0;
    _pageOrganizerFollowQueued = false;
    _pageOrganizerFollowRequestId++;
}
~~~

Change `ReceiveActivePageChanged` to update state and schedule follow separately:

```csharp
var nextState = _pageOrganizerState.ActivatePage(activePage);
ApplyPageOrganizerState(nextState);
QueueActivePageFollow(activePage);
```

Implement the queue with a request id so an old dispatcher operation cannot cancel a newer document's request:

```csharp
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
```

Implement `FollowActivePageOrganizerItem` by locating the generated `ListBoxItem`, finding the nested `ScrollViewer`, translating item coordinates into that viewport, passing `VerticalOffset`, `ViewportHeight`, item top/bottom, and `ScrollableHeight` to Task 1's calculator, then calling `ScrollToVerticalOffset` only when the result is non-null. Add a recursive `FindVisualDescendant<T>` helper next to the existing visual-tree helpers.

Use this exact viewport bridge:

~~~csharp
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
~~~

Add this recursive visual-tree helper next to the existing visual-tree helpers:

~~~csharp
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
~~~

Use this precise suspension condition:

```csharp
private bool IsActivePageFollowSuspended() =>
    IsDocumentMutationInProgress || _isPageOrganizerDragInProgress;
```

Set `_isPageOrganizerDragInProgress = true` immediately before `DragDrop.DoDragDrop` in `OnPageOrganizerItemPreviewMouseMove` and reset it in that method's existing `finally` block before clearing pointer state.

Use this exact shape for that existing drag block:

~~~csharp
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
~~~

- [ ] **Step 5: Run the focused unit and contract checks**

Run:

```bash
dotnet test test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj
node test/thumbnail-selection-contract.test.mjs
```

Expected: all xUnit tests pass and the Node contract exits successfully.

- [ ] **Step 6: Commit the WPF integration**

```bash
git add src/PdfMergeTool/MainWindow.xaml src/PdfMergeTool/MainWindow.xaml.cs test/thumbnail-selection-contract.test.mjs
git commit -m "feat: follow active organizer page and keyboard navigation"
```

### Task 3: Enforce Release Contracts And Perform Windows Verification

**Files:**
- Modify: `scripts/verify-page-organizer.ps1:73-120`

**Interfaces:**
- Consumes: the methods and XAML event names introduced in Task 2.
- Produces: release-time verification that the Page Organizer remains the owner of follow and focused keyboard interactions.

- [ ] **Step 1: Add the PowerShell release contract**

Extend the `MainWindow` condition in `scripts/verify-page-organizer.ps1` with these required fragments:

```powershell
$mainWindow -notmatch 'QueueActivePageFollow\(' -or
$mainWindow -notmatch 'FollowActivePageOrganizerItem\(' -or
$mainWindow -notmatch 'IsActivePageFollowSuspended\(' -or
$mainWindow -notmatch 'PageOrganizerViewport\.GetVerticalOffsetToReveal' -or
$mainWindow -notmatch 'OnPageOrganizerPreviewKeyDown' -or
$mainWindow -notmatch 'PageOrganizerList\.Focus\(\)' -or
$mainWindow -notmatch 'Key\.PageUp' -or
$mainWindow -notmatch 'Key\.PageDown' -or
$mainWindow -notmatch 'Key\.Home' -or
$mainWindow -notmatch 'Key\.End'
```

Extend the XAML condition with:

```powershell
$xaml -notmatch 'x:Name="PageOrganizerList"[\s\S]*?Focusable="True"' -or
$xaml -notmatch 'PreviewKeyDown="OnPageOrganizerPreviewKeyDown"'
```

- [ ] **Step 2: Run the release contract after the completed Task 2 integration**

Run:

```powershell
pwsh scripts/verify-page-organizer.ps1
```

Expected: PASS because Task 2 has already supplied every required symbol and XAML binding.

- [ ] **Step 3: Run all local verification**

Run:

```bash
git diff --check
dotnet test test/PdfMergeTool.Tests/PdfMergeTool.Tests.csproj
node test/thumbnail-selection-contract.test.mjs
pwsh scripts/verify-page-organizer.ps1
pwsh scripts/verify-stability.ps1
```

Expected: all commands exit successfully. On a non-Windows machine without the .NET SDK, record the local limitation and require the existing Windows release workflow to run the same test project, contracts, installer build, and packaged-startup smoke test.

- [ ] **Step 4: Execute the Windows manual smoke sequence**

Use a PDF with at least 30 pages and enough sidebar width for multiple thumbnail columns:

1. Scroll the main preview until page 1 leaves the sidebar, then confirm the sidebar moves only when the active thumbnail becomes outside its viewport.
2. Rapidly wheel from an early page to a distant page and confirm the sidebar settles once on the final active thumbnail.
3. Select several pages, drag the Page Move Group, and confirm the sidebar does not shift beneath the pointer.
4. Click a thumbnail, press each arrow key, `PageUp`, `PageDown`, `Home`, and `End`, and confirm only the active highlight and preview move while the selected group remains unchanged.
5. With the selected group still retained, test `Ctrl+C`, `Ctrl+X`, `Ctrl+V`, `Ctrl+Z`, `Ctrl+Y`, and `Delete`; confirm each operation uses the same outcomes and history behavior as the corresponding menu command.
6. Open a second PDF before a queued follow can run and confirm no old-page sidebar scroll occurs.

- [ ] **Step 5: Commit the verification contract**

```bash
git add scripts/verify-page-organizer.ps1
git commit -m "test: verify organizer follow and keyboard controls"
```
