Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$mainWindowPath = Join-Path $root 'src\PdfMergeTool\MainWindow.xaml.cs'
$mainWindowXamlPath = Join-Path $root 'src\PdfMergeTool\MainWindow.xaml'
$statePath = Join-Path $root 'src\PdfMergeTool\Services\EditorDocumentState.cs'
$operationCoordinatorPath = Join-Path $root 'src\PdfMergeTool\Services\DocumentOperationCoordinator.cs'
$adapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\app-adapter.js'
$editorAdapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\editor-adapter.js'
$viewerScriptPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\viewer.mjs'
$overlayGeometryVerificationPath = Join-Path $root 'scripts\verify-overlay-geometry.ps1'
$overlayExportVerificationPath = Join-Path $root 'scripts\verify-pdf-lib-overlay-export.ps1'

foreach ($path in @($mainWindowPath, $mainWindowXamlPath, $statePath, $operationCoordinatorPath, $adapterPath, $editorAdapterPath, $viewerScriptPath, $overlayGeometryVerificationPath, $overlayExportVerificationPath)) {
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
$adapter = Get-Content -Raw $adapterPath
$editorAdapter = Get-Content -Raw $editorAdapterPath
$viewer = Get-Content -Raw $viewerScriptPath

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

if ($mainWindow -notmatch 'EditorDocumentState\? _pageOrganizerState' -or
    $mainWindow -notmatch 'InitializePageOrganizerStateAsync\(' -or
    $mainWindow -notmatch 'ApplyPageOrganizerState\(' -or
    $mainWindow -notmatch 'OnPageOrganizerCheckBoxPreviewMouseLeftButtonDown' -or
    $mainWindow -notmatch 'OnPageOrganizerItemPreviewMouseLeftButtonDown' -or
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
    $mainWindow -notmatch 'PageOrganizerViewport\.GetVerticalOffsetToReveal' -or
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

if ($xaml -notmatch 'x:Name="PageOrganizerList"' -or
    $xaml -notmatch 'x:Name="PageOrganizerList"[\s\S]*?Focusable="True"' -or
    $xaml -notmatch 'PreviewMouseLeftButtonDown="OnPageOrganizerItemPreviewMouseLeftButtonDown"' -or
    $xaml -notmatch 'PreviewMouseLeftButtonDown="OnPageOrganizerCheckBoxPreviewMouseLeftButtonDown"' -or
    $xaml -notmatch 'PreviewMouseLeftButtonUp="OnPageOrganizerItemPreviewMouseLeftButtonUp"' -or
    $xaml -notmatch 'PreviewKeyDown="OnViewerPreviewKeyDown"' -or
    $xaml -notmatch 'PreviewKeyDown="OnPageOrganizerPreviewKeyDown"' -or
    $xaml -notmatch 'x:Name="PageOrganizerZoomSlider"' -or
    $xaml -notmatch '<WrapPanel Orientation="Horizontal"' -or
    $xaml -notmatch 'ScrollViewer\.HorizontalScrollBarVisibility="Disabled"' -or
    $xaml -notmatch 'DragLeave="OnPageOrganizerDragLeave"' -or
    $xaml -notmatch 'x:Name="DropBeforeIndicator"' -or
    $xaml -notmatch 'x:Name="DropAfterIndicator"' -or
    $xaml -notmatch 'Drop="OnPageOrganizerDrop"') {
    throw 'MainWindow must expose the app-owned Page Organizer panel and its direct input handlers.'
}

Write-Output 'page organizer checks passed.'
