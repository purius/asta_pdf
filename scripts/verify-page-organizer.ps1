Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$mainWindowPath = Join-Path $root 'src\PdfMergeTool\MainWindow.xaml.cs'
$mainWindowXamlPath = Join-Path $root 'src\PdfMergeTool\MainWindow.xaml'
$statePath = Join-Path $root 'src\PdfMergeTool\Services\EditorDocumentState.cs'
$operationCoordinatorPath = Join-Path $root 'src\PdfMergeTool\Services\DocumentOperationCoordinator.cs'
$thumbnailSchedulerPath = Join-Path $root 'src\PdfMergeTool\Services\PageOrganizerThumbnailScheduler.cs'
$adapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\app-adapter.js'
$editorAdapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\editor-adapter.js'
$viewerScriptPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\viewer.mjs'
$overlayGeometryVerificationPath = Join-Path $root 'scripts\verify-overlay-geometry.ps1'
$overlayExportVerificationPath = Join-Path $root 'scripts\verify-pdf-lib-overlay-export.ps1'

foreach ($path in @($mainWindowPath, $mainWindowXamlPath, $statePath, $operationCoordinatorPath, $thumbnailSchedulerPath, $adapterPath, $editorAdapterPath, $viewerScriptPath, $overlayGeometryVerificationPath, $overlayExportVerificationPath)) {
    if (-not (Test-Path $path)) {
        throw "Required Page Organizer contract file is missing: $path"
    }
}

& $overlayGeometryVerificationPath
& $overlayExportVerificationPath

$mainWindow = Get-Content -Raw $mainWindowPath
$xaml = Get-Content -Raw $mainWindowXamlPath
$state = Get-Content -Raw $statePath
$operationCoordinator = Get-Content -Raw $operationCoordinatorPath
$thumbnailScheduler = Get-Content -Raw $thumbnailSchedulerPath
$adapter = Get-Content -Raw $adapterPath
$editorAdapter = Get-Content -Raw $editorAdapterPath
$viewer = Get-Content -Raw $viewerScriptPath

function Assert-ThumbnailOverlayVisibilityTrigger {
    param(
        [System.Xml.XmlElement]$Control,
        [string]$StylePropertyElementName,
        [string]$BindingName,
        [string]$ControlName
    )

    $style = $Control.SelectSingleNode(
        "./*[local-name()='$StylePropertyElementName']/*[local-name()='Style']"
    )
    if ($null -eq $style) {
        throw "$ControlName must define a local Style."
    }

    $collapsedSetter = $style.SelectSingleNode(
        './*[local-name()="Setter" and @Property="Visibility" and @Value="Collapsed"]'
    )
    if ($null -eq $collapsedSetter) {
        throw "$ControlName must default Visibility to Collapsed in its Style."
    }

    $trigger = @(
        $style.SelectNodes(
            './*[local-name()="Style.Triggers"]/*[local-name()="DataTrigger"]'
        ) | Where-Object {
            $_.GetAttribute('Binding') -eq "{Binding $BindingName}" -and
            $_.GetAttribute('Value') -eq 'True'
        }
    ) | Select-Object -First 1
    if ($null -eq $trigger) {
        throw "$ControlName must have a $BindingName DataTrigger."
    }

    $visibleSetter = $trigger.SelectSingleNode(
        './*[local-name()="Setter" and @Property="Visibility" and @Value="Visible"]'
    )
    if ($null -eq $visibleSetter) {
        throw "$ControlName must set Visibility to Visible in its $BindingName DataTrigger."
    }
}

if ($viewer -notmatch 'enableSplitMerge:\s*\{\s*value:\s*false') {
    throw 'PDF.js split/merge editing must remain disabled because Page Organizer owns structural edits.'
}

if ($adapter -match 'pageOrderChanged' -or $adapter -match 'syncPageOrderFromPagesMapper' -or $adapter -match 'undoPageEdit\(' -or $adapter -match 'eventBus\?\._on\("pagesedited"') {
    throw 'Preview adapter must not own page order, page selection, or structural undo state.'
}

if ($adapter -notmatch 'function configurePreviewOnlyViewer\(' -or
    $adapter -notmatch 'viewsManagerToggleButton' -or
    $adapter -notmatch 'function queuePdfOpen\(' -or
    $adapter -notmatch 'latestRequestedLoadId' -or
    $adapter -notmatch 'currentDocumentLoadId' -or
    $adapter -notmatch 'currentDocumentLoadId = 0;\s*window\.AstaViewerLoadId = 0;' -or
    $adapter -notmatch 'case "goToPage":\s*goToPage\(options\?\.pageNumber\);' -or
    $adapter -notmatch 'type:\s*"activePageChanged"') {
    throw 'Preview adapter must hide native page editing and serialize load-scoped preview navigation.'
}

if ($editorAdapter -notmatch 'AstaViewerLoadId') {
    throw 'Editor adapter must identify the active viewer load before reporting editor state.'
}

if ($state -notmatch 'MovePageGroup\(' -or
    $state -notmatch 'RotateSelectedPages\(' -or
    $state -notmatch 'DeleteSelectedPages\(' -or
    $state -notmatch 'ReversePageOrder\(' -or
    $state -notmatch 'EditorDocumentState Undo\(' -or
    $state -notmatch 'EditorDocumentState Redo\(' -or
    $state -notmatch 'PageSelectionMode.Range' -or
    $state -notmatch 'PageSelectionMode.Toggle') {
    throw 'Page Organizer state must own selection, page mutations, and reversible history.'
}

if ($operationCoordinator -notmatch 'StartNewDocument\(' -or
    $operationCoordinator -notmatch 'EnterMutationAsync\(' -or
    $operationCoordinator -notmatch 'ThrowIfSuperseded\(' -or
    $operationCoordinator -notmatch 'CancellationTokenSource') {
    throw 'Document operations must cancel stale work and serialize mutations for the active PDF.'
}

$pageOrganizerMouseDownButtonBypassPattern = 'private void OnPageOrganizerItemPreviewMouseLeftButtonDown\(object sender, MouseButtonEventArgs e\)[\s\S]*?FindVisualAncestor<CheckBox>\(source\) is not null\s*\|\|\s*FindVisualAncestor<Button>\(source\) is not null\s*\)\s*\{\s*return;\s*\}'
$pageOrganizerMouseUpButtonBypassPattern = 'private void OnPageOrganizerItemPreviewMouseLeftButtonUp\(object sender, MouseButtonEventArgs e\)\s*\{\s*if\s*\(\s*e\.OriginalSource is DependencyObject source\s*&&\s*FindVisualAncestor<Button>\(source\) is not null\s*\)\s*\{\s*return;\s*\}'

if ($mainWindow -notmatch 'EditorDocumentState\? _pageOrganizerState' -or
    $mainWindow -notmatch 'InitializePageOrganizerStateAsync\(' -or
    $mainWindow -notmatch 'ApplyPageOrganizerState\(' -or
    $mainWindow -notmatch 'OnPageOrganizerCheckBoxPreviewMouseLeftButtonDown' -or
    $mainWindow -notmatch 'OnPageOrganizerItemPreviewMouseLeftButtonDown' -or
    $mainWindow -notmatch $pageOrganizerMouseDownButtonBypassPattern -or
    $mainWindow -notmatch $pageOrganizerMouseUpButtonBypassPattern -or
    $mainWindow -notmatch 'OnPageOrganizerDrop' -or
    $mainWindow -notmatch '_pendingLoadGeneration' -or
    $mainWindow -notmatch 'IsCurrentViewerLoadMessage\(' -or
    $mainWindow -notmatch 'DocumentOperationCoordinator _documentOperations' -or
    $mainWindow -notmatch 'RunCurrentDocumentMutationAsync\(' -or
    $mainWindow -notmatch 'SetDocumentMutationUiState\(' -or
    $mainWindow -notmatch 'PageOrganizerList\.IsHitTestVisible = false' -or
    $mainWindow -notmatch 'PdfViewer\.IsHitTestVisible = false' -or
    $mainWindow -notmatch 'OnViewerPreviewKeyDown' -or
    $mainWindow -notmatch 'MinimumHorizontalDragDistance' -or
    $mainWindow -notmatch 'Mouse\.Capture\(PageOrganizerList, CaptureMode\.SubTree\)' -or
    $mainWindow -notmatch 'OnPageOrganizerZoomSliderValueChanged' -or
    $mainWindow -notmatch 'SetPageOrganizerThumbnailHeight\(' -or
    $mainWindow -notmatch 'ThumbnailCardWidth' -or
    $mainWindow -notmatch 'UpdatePageOrganizerDropIndicator\(GetPageOrganizerInsertionIndex\(e\)\)' -or
    $mainWindow -notmatch 'ClearPageOrganizerDropIndicator\(' -or
    $mainWindow -notmatch 'IsDropBefore' -or
    $mainWindow -notmatch 'IsDropAfter' -or
    $mainWindow -notmatch '_pageOrganizerState\.Undo\(\)' -or
    $mainWindow -notmatch '_pageOrganizerState\.Redo\(\)' -or
    $mainWindow -notmatch 'QueueActivePageFollow\(' -or
    $mainWindow -notmatch 'FollowActivePageOrganizerItem\(' -or
    $mainWindow -notmatch 'IsActivePageFollowSuspended\(' -or
    $mainWindow -notmatch 'PageOrganizerList\.ScrollIntoView\(PageOrganizerRows\[rowIndex\]\)' -or
    $mainWindow -match 'ScrollToVerticalOffset' -or
    $mainWindow -notmatch 'OnPageOrganizerPreviewKeyDown' -or
    $mainWindow -notmatch 'PageOrganizerList\.Focus\(\)' -or
    $mainWindow -notmatch 'Key\.PageUp' -or
    $mainWindow -notmatch 'Key\.PageDown' -or
    $mainWindow -notmatch 'Key\.Home' -or
    $mainWindow -notmatch 'Key\.End') {
    throw 'MainWindow must route Page Organizer selection, drag moves, and undo/redo through one app-owned state.'
}

if ($mainWindow -match 'SendViewerCommand\("thumbZoom(In|Out|Reset)"\)' -or
    $mainWindow -notmatch 'OnThumbZoomInClick\(sender, e\);' -or
    $mainWindow -notmatch '_fallbackRenderService\.OpenDocument\(sourcePath\)' -or
    $mainWindow -match 'exportA4PageImages') {
    throw 'Thumbnail zoom and A4 conversion must be owned by the WPF Page Organizer and native renderer.'
}

if ($mainWindow -match 'pageNumbers\s*\.\s*Take\s*\(\s*96\s*\)' -or
    $mainWindow -notmatch 'PageOrganizerThumbnailScheduler' -or
    $mainWindow -notmatch 'OnPageOrganizerThumbnailScrollChanged' -or
    $mainWindow -notmatch 'RefreshPageOrganizerThumbnailViewport' -or
    $mainWindow -notmatch 'OnPageOrganizerThumbnailRetryClick' -or
    $mainWindow -notmatch 'PageOrganizerThumbnailRenderState.Failed') {
    throw 'Page Organizer thumbnail scheduling, viewport refresh, and retry state must remain app-owned.'
}

if ($mainWindow -notmatch 'ObservableCollection<PageOrganizerRow> PageOrganizerRows' -or
    $mainWindow -notmatch 'Dictionary<int, PageOrganizerItem> _pageOrganizerItemsByPageNumber' -or
    $mainWindow -notmatch 'void RebuildPageOrganizerRows\(' -or
    $mainWindow -notmatch 'void OnPageOrganizerListSizeChanged\(' -or
    $mainWindow -notmatch 'scheduler\.UpdateCacheOrder\(state\.PageNumbers\)' -or
    $mainWindow -notmatch '_pageOrganizerThumbnailViewportRefreshQueued' -or
    $thumbnailScheduler -notmatch 'const int MaximumCacheWindowSize = 96' -or
    $thumbnailScheduler -notmatch 'void UpdateCacheOrder\(') {
    throw 'Page Organizer rows must virtualize outer rows and keep cache order bounded by the current display order.'
}

$visiblePageDiscovery = [regex]::Match(
    $mainWindow,
    'private IReadOnlyList<int> GetVisiblePageOrganizerThumbnailNumbers\(\)[\s\S]*?(?=    private void RefreshPageOrganizerThumbnailViewport\(\))'
)
if (-not $visiblePageDiscovery.Success -or
    $visiblePageDiscovery.Value -notmatch 'GetRealizedPageOrganizerItemContainers\(' -or
    $visiblePageDiscovery.Value -match 'PageOrganizerItems\.Count' -or
    $visiblePageDiscovery.Value -match 'ItemContainerGenerator\.ContainerFromIndex\(index\)') {
    throw 'Viewport discovery must inspect only realized organizer rows and cards.'
}

$viewportRefresh = [regex]::Match(
    $mainWindow,
    'private void RefreshPageOrganizerThumbnailViewport\(\)[\s\S]*?(?=    private void UpdateWindowTitle\(\))'
)
if (-not $viewportRefresh.Success -or
    $viewportRefresh.Value -notmatch 'previousCacheWindow\.Except\(cacheWindow\)' -or
    $viewportRefresh.Value -notmatch 'cacheWindow\.Except\(previousCacheWindow\)' -or
    $viewportRefresh.Value -match 'foreach \(var item in PageOrganizerItems\)') {
    throw 'Viewport refresh must apply cache deltas instead of scanning every organizer item.'
}

$thumbnailScrollHandler = [regex]::Match(
    $mainWindow,
    'private void OnPageOrganizerThumbnailScrollChanged\([\s\S]*?(?=    private IReadOnlyList<int> GetVisiblePageOrganizerThumbnailNumbers\(\))'
)
if (-not $thumbnailScrollHandler.Success -or
    $thumbnailScrollHandler.Value -notmatch 'QueuePageOrganizerThumbnailViewportRefresh\(' -or
    $thumbnailScrollHandler.Value -match 'RefreshPageOrganizerThumbnailViewport\(\);') {
    throw 'Organizer scroll events must coalesce viewport refresh work.'
}

$cacheWindowMethod = [regex]::Match(
    $thumbnailScheduler,
    'public IReadOnlySet<int> GetCacheWindow\(IReadOnlyCollection<int> visiblePageNumbers\)[\s\S]*?(?=    private bool TryStart\()'
)
if (-not $cacheWindowMethod.Success -or
    $cacheWindowMethod.Value -notmatch 'maximumWindowSize' -or
    $cacheWindowMethod.Value -match 'foreach \(var visibleIndex') {
    throw 'Thumbnail cache windows must stay contiguous and hard-bounded in current display order.'
}

if ($xaml -notmatch 'x:Name="PageOrganizerList"' -or
    $xaml -notmatch 'x:Name="PageOrganizerList"[\s\S]*?Focusable="True"' -or
    $xaml -notmatch 'PreviewMouseLeftButtonDown="OnPageOrganizerItemPreviewMouseLeftButtonDown"' -or
    $xaml -notmatch 'PreviewMouseLeftButtonDown="OnPageOrganizerCheckBoxPreviewMouseLeftButtonDown"' -or
    $xaml -notmatch 'PreviewMouseLeftButtonUp="OnPageOrganizerItemPreviewMouseLeftButtonUp"' -or
    $xaml -notmatch 'PreviewKeyDown="OnViewerPreviewKeyDown"' -or
    $xaml -notmatch 'PreviewKeyDown="OnPageOrganizerPreviewKeyDown"' -or
    $xaml -notmatch 'x:Name="PageOrganizerZoomSlider"' -or
    $xaml -notmatch 'ItemsSource="\{Binding PageOrganizerRows,' -or
    $xaml -notmatch 'ScrollViewer\.CanContentScroll="True"' -or
    $xaml -notmatch '<VirtualizingStackPanel[\s\S]*?VirtualizationMode="Recycling"' -or
    $xaml -notmatch '<WrapPanel Orientation="Horizontal"' -or
    $xaml -notmatch 'ScrollViewer\.HorizontalScrollBarVisibility="Disabled"' -or
    $xaml -notmatch 'DragLeave="OnPageOrganizerDragLeave"' -or
    $xaml -notmatch 'x:Name="DropBeforeIndicator"' -or
    $xaml -notmatch 'x:Name="DropAfterIndicator"' -or
    $xaml -notmatch 'Drop="OnPageOrganizerDrop"') {
    throw 'MainWindow must expose the app-owned Page Organizer panel and its direct input handlers.'
}

try {
    [xml]$xamlDocument = $xaml
} catch {
    throw "MainWindow XAML must be valid XML for thumbnail overlay verification: $($_.Exception.Message)"
}

$thumbnailDataTemplate = $xamlDocument.SelectSingleNode(
    '//*[local-name()="DataTemplate"][.//*[local-name()="Image" and @Source="{Binding Thumbnail}"]]'
)
if ($null -eq $thumbnailDataTemplate) {
    throw 'MainWindow must expose thumbnail overlays inside the DataTemplate that binds Image.Source to Thumbnail.'
}

$thumbnailLoadingProgressBar = $thumbnailDataTemplate.SelectSingleNode(
    './/*[local-name()="ProgressBar" and @IsHitTestVisible="False"]'
)
if ($null -eq $thumbnailLoadingProgressBar) {
    throw 'Thumbnail DataTemplate must contain a non-hit-testable loading ProgressBar.'
}
Assert-ThumbnailOverlayVisibilityTrigger -Control $thumbnailLoadingProgressBar -StylePropertyElementName 'ProgressBar.Style' -BindingName 'IsThumbnailLoading' -ControlName 'Thumbnail loading ProgressBar'

$thumbnailRetryButton = $thumbnailDataTemplate.SelectSingleNode(
    './/*[local-name()="Button" and @Click="OnPageOrganizerThumbnailRetryClick"]'
)
if ($null -eq $thumbnailRetryButton) {
    throw 'Thumbnail DataTemplate must contain the retry Button.'
}
if ($thumbnailRetryButton.HasAttribute('IsHitTestVisible') -and
    $thumbnailRetryButton.GetAttribute('IsHitTestVisible') -ieq 'False') {
    throw 'Thumbnail retry Button must remain interactive.'
}
Assert-ThumbnailOverlayVisibilityTrigger -Control $thumbnailRetryButton -StylePropertyElementName 'Button.Style' -BindingName 'IsThumbnailFailed' -ControlName 'Thumbnail retry Button'

$thumbnailRetryIcon = $thumbnailRetryButton.SelectSingleNode(
    './/*[local-name()="PackIconMaterial" and @Kind="Refresh"]'
)
if ($null -eq $thumbnailRetryIcon) {
    throw 'Thumbnail retry Button must contain a Refresh PackIconMaterial.'
}

Write-Output 'page organizer checks passed.'
