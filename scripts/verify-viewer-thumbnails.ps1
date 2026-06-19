Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$legacyViewerPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewer'
$officialViewerPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\viewer.html'
$officialAdapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\app-adapter.js'
$officialEditorAdapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\editor-adapter.js'
$officialPdfLibAdapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\pdf-lib-adapter.js'
$officialPdfLibPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\vendor\pdf-lib.min.js'
$officialFontkitPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\vendor\fontkit.umd.min.js'
$officialViewerScriptPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\viewer.mjs'
$officialBuildPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\build\pdf.mjs'
$overlayGeometryVerificationPath = Join-Path $root 'scripts\verify-overlay-geometry.ps1'
$overlayExportVerificationPath = Join-Path $root 'scripts\verify-pdf-lib-overlay-export.ps1'
$mainWindowPath = Join-Path $root 'src\PdfMergeTool\MainWindow.xaml.cs'
$mainWindowXamlPath = Join-Path $root 'src\PdfMergeTool\MainWindow.xaml'
$projectPath = Join-Path $root 'src\PdfMergeTool\PdfMergeTool.csproj'
$fontServicePath = Join-Path $root 'src\PdfMergeTool\Services\WindowsFontService.cs'
$mainWindow = Get-Content -Raw $mainWindowPath
$mainWindowXaml = Get-Content -Raw $mainWindowXamlPath
$project = Get-Content -Raw $projectPath

if (-not (Test-Path $officialViewerPath)) {
    throw 'official PDF.js viewer.html must be packaged.'
}

if (-not (Test-Path $officialAdapterPath)) {
    throw 'official viewer adapter must be packaged.'
}

if (-not (Test-Path $officialPdfLibAdapterPath)) {
    throw 'official pdf-lib editor adapter must be packaged.'
}

if (-not (Test-Path $officialEditorAdapterPath)) {
    throw 'official viewer editor overlay adapter must be packaged.'
}

if (-not (Test-Path $officialPdfLibPath)) {
    throw 'pdf-lib browser bundle must be packaged.'
}

if (-not (Test-Path $officialFontkitPath)) {
    throw 'fontkit browser bundle must be packaged for custom Windows font embedding.'
}

if (-not (Test-Path $officialBuildPath)) {
    throw 'official PDF.js build files must be packaged.'
}

if (-not (Test-Path $overlayGeometryVerificationPath)) {
    throw 'overlay geometry verification script must exist.'
}

& $overlayGeometryVerificationPath

if (-not (Test-Path $overlayExportVerificationPath)) {
    throw 'pdf-lib overlay export verification script must exist.'
}

& $overlayExportVerificationPath

$officialViewer = Get-Content -Raw $officialViewerPath
$officialAdapter = Get-Content -Raw $officialAdapterPath
$officialEditorAdapter = Get-Content -Raw $officialEditorAdapterPath
$officialPdfLibAdapter = Get-Content -Raw $officialPdfLibAdapterPath
$officialViewerScript = Get-Content -Raw $officialViewerScriptPath
$fontService = Get-Content -Raw $fontServicePath

if ($officialViewer -notmatch 'app-adapter\.js') {
    throw 'official viewer.html must load the app adapter.'
}

if ($officialViewer -notmatch 'vendor/pdf-lib\.min\.js') {
    throw 'official viewer.html must load pdf-lib.'
}

if ($officialViewer -notmatch 'vendor/fontkit\.umd\.min\.js') {
    throw 'official viewer.html must load fontkit for custom font embedding.'
}

if ($officialViewer -notmatch 'pdf-lib-adapter\.js') {
    throw 'official viewer.html must load the pdf-lib editor adapter.'
}

if ($officialViewer -notmatch 'editor-adapter\.js') {
    throw 'official viewer.html must load the editor overlay adapter.'
}

if ($officialAdapter -notmatch 'window\.PDFViewerApplication') {
    throw 'official viewer adapter must use the PDF.js viewer application API.'
}

if ($officialAdapter -notmatch 'type:\s*"viewerDiagnostic"') {
    throw 'official viewer adapter diagnostics must match the WPF viewerDiagnostic message contract.'
}

if ($officialAdapter -notmatch 'rotations:\s*pageRotations') {
    throw 'official viewer adapter must publish rotations using the existing WPF message field.'
}

if ($officialAdapter -notmatch 'exportOverlayPdf') {
    throw 'official viewer adapter must expose a host-callable pdf-lib export path.'
}

if ($officialAdapter -notmatch 'type:\s*"overlayPdfExportFailed"' -or $officialAdapter -notmatch 'requestId:\s*data\.requestId \?\? null') {
    throw 'official viewer adapter must report pdf-lib export failures back to WPF without waiting for a timeout.'
}

if ($officialAdapter -notmatch 'setEditorFonts') {
    throw 'official viewer adapter must accept the Windows font list from WPF.'
}

if ($officialAdapter -notmatch 'collectEditorState') {
    throw 'official viewer adapter must expose editor state collection before saving.'
}

if ($officialAdapter -notmatch 'thumbZoomIn' -or $officialAdapter -notmatch 'thumbZoomOut' -or $officialAdapter -notmatch 'thumbZoomReset') {
    throw 'official viewer adapter must preserve thumbnail zoom commands from the WPF toolbar.'
}

if ($officialAdapter -notmatch 'reversePageOrder') {
    throw 'official viewer adapter must preserve reverse page order command semantics.'
}

if ($officialAdapter -notmatch 'case "undo"' -or $officialAdapter -notmatch 'case "redo"') {
    throw 'official viewer adapter must preserve undo and redo commands.'
}

if ($officialAdapter -notmatch 'applyPageStatePresentation') {
    throw 'official viewer adapter must visually hide deleted pages and thumbnails while saving page state.'
}

$officialPageChanging = [regex]::Match(
    $officialAdapter,
    'eventBus\?\._on\("pagechanging",\s*event\s*=>\s*\{[\s\S]*?\n    \}\);')
if (-not $officialPageChanging.Success -or
    $officialPageChanging.Groups[0].Value -match 'postPageOrder\(') {
    throw 'official viewer pagechanging events must not republish full page order state during navigation.'
}

if ($officialAdapter -notmatch 'type:\s*"activePageChanged"[\s\S]*selectedPages:\s*\[\.\.\.selectedPages\]') {
    throw 'official viewer pagechanging events must publish active page and selection without rebuilding page order state.'
}

if ($officialAdapter -notmatch 'function beginExplicitPageNavigation\(' -or
    $officialAdapter -notmatch 'function shouldAcceptPageChange\(' -or
    $officialAdapter -notmatch 'if \(!shouldAcceptPageChange\(activePage\)\) \{[\s\S]*?return;[\s\S]*?\}') {
    throw 'official viewer adapter must ignore transient pagechanging events that do not match the latest explicit navigation target.'
}

if ($officialAdapter -notmatch 'explicitNavigationSettledUntil' -or
    $officialAdapter -notmatch 'lastAcceptedExplicitNavigationPage' -or
    $officialAdapter -notmatch 'lastAcceptedExplicitNavigationPage = Number\(pageNumber\);' -or
    $officialAdapter -notmatch 'explicitNavigationSettledUntil = Date\.now\(\) \+ 450;' -or
    $officialAdapter -notmatch 'Date\.now\(\) <= explicitNavigationSettledUntil' -or
    $officialAdapter -notmatch 'Number\(pageNumber\) !== lastAcceptedExplicitNavigationPage') {
    throw 'official viewer adapter must keep suppressing stale pagechanging events briefly after accepting an explicit navigation target.'
}

$userPageChangeIntentBlock = [regex]::Match(
    $officialAdapter,
    'function hasRecentUserPageChangeIntent\(\) \{[\s\S]*?\n  \}').Value
if (-not $userPageChangeIntentBlock -or
    $userPageChangeIntentBlock -notmatch 'Date\.now\(\) - lastUserPageChangeIntentAt <= 1200') {
    throw 'official viewer adapter must distinguish direct user scroll/key page changes from stale programmatic pagechanging events.'
}

$shouldAcceptPageChangeBlock = [regex]::Match(
    $officialAdapter,
    'function shouldAcceptPageChange\(pageNumber\) \{[\s\S]*?\n  function findNavigationTargetPage').Value
if (-not $shouldAcceptPageChangeBlock -or
    $shouldAcceptPageChangeBlock -notmatch 'if \(lastAcceptedExplicitNavigationPage && Date\.now\(\) > explicitNavigationSettledUntil\) \{[\s\S]*?explicitNavigationSettledUntil = 0;[\s\S]*?lastAcceptedExplicitNavigationPage = null;[\s\S]*?\}') {
    throw 'official viewer adapter must clear accepted explicit navigation state after the settle window expires.'
}

if ($officialAdapter -notmatch 'protectedExplicitNavigationPage' -or
    -not $shouldAcceptPageChangeBlock -or
    $shouldAcceptPageChangeBlock -notmatch 'if \(\s*protectedExplicitNavigationPage &&[\s\S]*?Number\(pageNumber\) !== protectedExplicitNavigationPage &&[\s\S]*?!hasRecentUserPageChangeIntent\(\)[\s\S]*?\) \{[\s\S]*?return false;[\s\S]*?\}') {
    throw 'official viewer adapter must keep rejecting stale non-target pagechanging events after explicit navigation until the user scrolls or keys navigation.'
}

$beginExplicitPageNavigationBlock = [regex]::Match(
    $officialAdapter,
    'function beginExplicitPageNavigation\(pageNumber\) \{[\s\S]*?\n  \}').Value
if (-not $beginExplicitPageNavigationBlock -or
    $beginExplicitPageNavigationBlock -notmatch 'protectedExplicitNavigationPage = null;') {
    throw 'starting a new explicit page navigation must clear the previous protected page so rapid thumbnail clicks can reach the newest target.'
}

if ($officialAdapter -notmatch 'function markUserPageChangeIntent\(' -or
    $officialAdapter -notmatch 'document\.addEventListener\("wheel", markUserPageChangeIntent, \{ passive: true \}\)' -or
    $officialAdapter -notmatch 'document\.addEventListener\("touchstart", markUserPageChangeIntent, \{ passive: true \}\)' -or
    $officialAdapter -notmatch 'document\.addEventListener\("keydown", markKeyboardPageChangeIntent, true\)') {
    throw 'official viewer adapter must record real user scroll and keyboard navigation before accepting non-target pagechanging events.'
}

if ($officialAdapter -notmatch 'function markNativeViewerNavigationIntent\(' -or
    $officialAdapter -notmatch '#previous, #next, #firstPage, #lastPage, #pageNumber' -or
    $officialAdapter -notmatch 'document\.addEventListener\("click", markNativeViewerNavigationIntent, true\)' -or
    $officialAdapter -notmatch 'document\.addEventListener\("change", markNativeViewerNavigationIntent, true\)' -or
    $officialAdapter -notmatch 'document\.addEventListener\("input", markNativeViewerNavigationIntent, true\)') {
    throw 'official viewer adapter must treat native PDF.js toolbar page controls as direct user navigation intent.'
}

if ($officialAdapter -notmatch 'function resetExplicitNavigationTracking\(\)' -or
    $officialAdapter -notmatch 'resetExplicitNavigationTracking\(\);[\s\S]*?window\.EditorAdapter\?\.clear\?\.\(\);' -or
    $officialAdapter -notmatch 'function rebuildPageOrder\([\s\S]*?resetExplicitNavigationTracking\(\);') {
    throw 'official viewer adapter must clear stale explicit navigation guards when loading or rebuilding a document.'
}

if ($officialAdapter -notmatch 'captureExplicitNavigationIntent' -or
    $officialAdapter -notmatch '#thumbnailsView|#thumbnailView' -or
    $officialAdapter -notmatch 'page-number') {
    throw 'official viewer adapter must capture rapid thumbnail clicks as explicit navigation targets.'
}

$captureNavigationIntentBlock = [regex]::Match(
    $officialAdapter,
    'function captureExplicitNavigationIntent\(event\) \{[\s\S]*?\n  \}').Value
if (-not $captureNavigationIntentBlock -or
    $captureNavigationIntentBlock -notmatch 'if \(event\.defaultPrevented\) return;' -or
    $captureNavigationIntentBlock -notmatch 'if \(event\.type === "pointerdown" && event\.isPrimary === false\) return;' -or
    $captureNavigationIntentBlock -notmatch 'if \("button" in event && event\.button !== 0\) return;' -or
    $captureNavigationIntentBlock -notmatch 'if \(event\.type === "click" && event\.detail === 0\) return;') {
    throw 'official viewer adapter must only capture primary, unhandled thumbnail navigation events.'
}

if ($officialAdapter -notmatch 'nativePageTransferDragOver' -or
    $officialAdapter -notmatch 'nativePageTransferDrop' -or
    $officialAdapter -notmatch 'nativeFileDragOver' -or
    $officialAdapter -notmatch 'nativeFileDrop') {
    throw 'official viewer adapter must preserve native page and file drag/drop insertion messages from WPF.'
}

if ($officialAdapter -notmatch 'type:\s*"insertExternalPages"[\s\S]*?insertionIndex' -or
    $officialAdapter -notmatch 'type:\s*"insertExternalFiles"[\s\S]*?paths[\s\S]*?insertionIndex') {
    throw 'official viewer adapter must translate native drops into existing WPF insertion messages with an insertion index.'
}

if ($officialAdapter -notmatch 'function getVisiblePageOrder\(' -or
    $officialAdapter -notmatch 'function goRelativeInPageOrder\(' -or
    $officialAdapter -notmatch 'const currentIndex = orderedPages\.indexOf\(currentPage\)' -or
    $officialAdapter -match 'function goRelative\(delta\) \{\s*goToPage\(\(getApp\(\)\?\.page \?\? 1\) \+ delta\);') {
    throw 'official viewer next/previous commands must navigate through the app pageOrder, not raw PDF page numbers.'
}

if ($officialAdapter -notmatch 'function goToPageOrderBoundary\(' -or
    $officialAdapter -notmatch 'goToPage\(orderedPages\[0\]\)' -or
    $officialAdapter -notmatch 'goToPage\(orderedPages\[orderedPages\.length - 1\]\)' -or
    $officialAdapter -match 'case "firstPage":\s*goToPage\(1\);' -or
    $officialAdapter -match 'case "lastPage":\s*goToPage\(getTotalPages\(\)\);') {
    throw 'official viewer first/last commands must navigate to the first/last page in app pageOrder.'
}

if ($officialViewerScript -notmatch 'enableSplitMerge:\s*\{\s*value:\s*true' -or
    $officialAdapter -notmatch 'function syncPageOrderFromPagesMapper\(' -or
    $officialAdapter -notmatch 'eventBus\?\._on\("pagesedited"' -or
    $officialAdapter -notmatch 'pagesMapper\.getPrevPageNumber\(index\)' -or
    $officialAdapter -notmatch 'pageStateDirty = true' -or
    $officialAdapter -match 'draggable = "true"' -or
    $officialAdapter -match 'addEventListener\("dragstart", handleThumbnailDragStart\)') {
    throw 'official viewer thumbnail drag reordering must use the built-in PDF.js split/merge pointer flow and sync pagesedited into pageOrder.'
}

if ($officialEditorAdapter -notmatch 'type:\s*"editorStateChanged"') {
    throw 'editor adapter must notify WPF when overlay edits make the document dirty.'
}

if ($officialEditorAdapter -notmatch 'type:\s*"editorStateCollected"') {
    throw 'editor adapter must return bounded overlay edit state for persistence.'
}

if ($officialEditorAdapter -notmatch 'astaEditorToolbar') {
    throw 'editor adapter must expose in-viewer text and shape editing controls.'
}

if ($officialEditorAdapter -notmatch 'data-action="toggleTools"' -or
    $officialEditorAdapter -notmatch 'asta-editor-tool-panel' -or
    $officialEditorAdapter -notmatch '#astaEditorToolbar\.expanded \.asta-editor-tool-panel' -or
    $officialEditorAdapter -notmatch 'max-height:\s*min\(70vh, 520px\)' -or
    $officialEditorAdapter -notmatch 'aria-expanded') {
    throw 'editor adapter toolbar must be collapsed into a compact menu so it does not cover the PDF page by default.'
}

if ($mainWindowXaml -notmatch 'Header="Edit Tools"' -or
    $mainWindowXaml -notmatch 'MenuItem Header="Text"[\s\S]*?Click="OnEditorTextClick"' -or
    $mainWindowXaml -notmatch 'MenuItem Header="Back"[\s\S]*?Click="OnEditorSendToBackClick"' -or
    $mainWindowXaml -match '<TextBlock Text="Text" />') {
    throw 'MainWindow editor commands must be compressed into a single toolbar menu instead of a long top button row.'
}

if ($officialEditorAdapter -notmatch 'data-mode="ellipse"' -or $officialEditorAdapter -notmatch 'data-mode="line"') {
    throw 'editor adapter must expose ellipse and line shape editing controls.'
}

if ($officialEditorAdapter -notmatch 'data-mode="arrow"') {
    throw 'editor adapter must expose arrow editing controls.'
}

if ($officialEditorAdapter -match 'data-mode="image"' -or $officialEditorAdapter -match 'data-mode="signature"') {
    throw 'editor adapter must not expose custom image or signature overlay controls when the official PDF.js tools are enabled.'
}

if ($officialEditorAdapter -notmatch 'data-mode="stamp"') {
    throw 'editor adapter must keep the custom stamp overlay control.'
}

if ($officialViewerScript -notmatch 'enableSignatureEditor:\s*\{\s*value:\s*true' -or
    $officialViewerScript -notmatch 'enableUpdatedAddImage:\s*\{\s*value:\s*true') {
    throw 'official PDF.js signature and updated add-image editors must be enabled.'
}

if ($officialViewerScript -notmatch 'enableoptimizedpartialrendering' -or
    $mainWindow -notmatch 'BuildViewerUrl\(\)' -or
    $mainWindow -notmatch 'EnableOptimizedPartialRendering') {
    throw 'optimized partial rendering must be controllable from the app settings and passed into the official viewer.'
}

if ($officialEditorAdapter -notmatch 'data-mode="pen"' -or $officialEditorAdapter -notmatch 'data-mode="highlight"' -or $officialEditorAdapter -notmatch 'function startInk\(' -or $officialEditorAdapter -notmatch 'function renderInkElement\(') {
    throw 'editor adapter must expose freehand pen and highlight annotation tools.'
}

if ($officialEditorAdapter -notmatch 'data-mode="whiteout"' -or $officialEditorAdapter -notmatch 'type:\s*"whiteout"') {
    throw 'editor adapter must expose a whiteout tool for visually covering arbitrary PDF regions.'
}

if ($officialEditorAdapter -notmatch 'function addSelectedTextHighlightEdits\(' -or $officialEditorAdapter -notmatch 'function getSelectedTextHighlightTargets\(' -or $officialEditorAdapter -notmatch 'type:\s*"textHighlight"') {
    throw 'editor adapter must support highlighting selected PDF text ranges.'
}

if ($officialEditorAdapter -notmatch 'function addSelectedTextWhiteoutEdits\(' -or $officialEditorAdapter -notmatch 'function getSelectedTextWhiteoutTargets\(' -or $officialEditorAdapter -notmatch 'Selected text whiteout is limited to one PDF page at a time') {
    throw 'editor adapter must support whiteout over selected PDF text ranges.'
}

if ($officialEditorAdapter -notmatch 'data-mode="redact"' -or $officialEditorAdapter -notmatch 'function addSelectedTextRedactEdits\(' -or $officialEditorAdapter -notmatch 'function getSelectedTextRedactTargets\(' -or $officialEditorAdapter -notmatch 'fillColor:\s*"#000000"') {
    throw 'editor adapter must support black redaction over selected PDF text ranges.'
}

if ($officialEditorAdapter -notmatch 'type:\s*"whiteout"[\s\S]*?variant:\s*edit\.variant') {
    throw 'editor adapter must preserve redaction variant when exporting whiteout edits.'
}

if ($officialEditorAdapter -notmatch 'data-mode="underline"' -or $officialEditorAdapter -notmatch 'data-mode="strikeout"' -or $officialEditorAdapter -notmatch 'function addSelectedTextLineMarkupEdits\(' -or $officialEditorAdapter -notmatch 'function getSelectedTextLineMarkupTargets\(') {
    throw 'editor adapter must support underline and strikeout markup over selected PDF text ranges.'
}

if ($officialEditorAdapter -notmatch 'data-mode="replaceText"' -or $officialEditorAdapter -notmatch 'function findTextLayerElementAt\(') {
    throw 'editor adapter must expose a PDF text-layer replacement tool.'
}

if ($officialEditorAdapter -notmatch 'function getTextReplacementMetrics\(' -or $officialEditorAdapter -notmatch 'whiteoutPadding' -or $officialEditorAdapter -notmatch 'lineHeight') {
    throw 'editor adapter must preserve text-layer replacement metrics for stable visual text edits.'
}

if ($officialEditorAdapter -notmatch 'function getSelectedTextReplacementTarget\(' -or $officialEditorAdapter -notmatch 'getRangeAt\(0\)' -or $officialEditorAdapter -notmatch 'getClientRects\(\)' -or $officialEditorAdapter -notmatch 'addSelectedTextReplacementEdit\(\)') {
    throw 'editor adapter must support replacing a selected PDF text range, not only a clicked text span.'
}

if ($officialEditorAdapter -notmatch 'contentEditable\s*=\s*"true"' -or $officialEditorAdapter -notmatch 'function beginInlineTextEdit\(' -or $officialEditorAdapter -notmatch 'function commitInlineTextEdit\(') {
    throw 'editor adapter must support direct inline editing for text, text replacement, and stamp overlays.'
}

if ($officialEditorAdapter -notmatch 'function isEditorTypingTarget\(' -or
    $officialEditorAdapter -notmatch 'target\.matches\?\.\("input, textarea, select"\)' -or
    $officialEditorAdapter -notmatch 'isEditorTypingTarget\(event\.target\)') {
    throw 'editor adapter keyboard shortcuts must not delete/copy/paste overlay objects while toolbar inputs or inline text are focused.'
}

if ($officialEditorAdapter -notmatch 'event\.defaultPrevented' -or
    $officialEditorAdapter -notmatch 'event\.isComposing' -or
    $officialEditorAdapter -notmatch 'if \(event\.defaultPrevented \|\| event\.isComposing \|\| isEditorTypingTarget\(event\.target\)\) return;') {
    throw 'editor adapter keyboard shortcuts must ignore handled and IME composition key events.'
}

if ($officialEditorAdapter -notmatch 'const onKeyDown = keyEvent => \{[\s\S]*?if \(keyEvent\.isComposing\) return;[\s\S]*?keyEvent\.key === "Enter"') {
    throw 'inline text editing must not commit or cancel while an IME composition key event is active.'
}

if ($officialEditorAdapter -notmatch 'data-role="strokeWidth"' -or $officialEditorAdapter -notmatch 'data-role="fillColor"' -or $officialEditorAdapter -notmatch 'function applySelectedProperties\(' -or $officialEditorAdapter -notmatch 'function syncToolbarFromEdit\(') {
    throw 'editor adapter must expose per-object font, size, stroke width, and fill controls.'
}

if ($officialEditorAdapter -notmatch 'function bindToolbarPropertyInput\(' -or
    $officialEditorAdapter -notmatch 'bindToolbarPropertyInput\("size"' -or
    $officialEditorAdapter -notmatch 'bindToolbarPropertyInput\("color"' -or
    $officialEditorAdapter -notmatch 'bindToolbarPropertyInput\("strokeWidth"' -or
    $officialEditorAdapter -notmatch 'bindToolbarPropertyInput\("fillColor"' -or
    $officialEditorAdapter -notmatch 'bindToolbarPropertyInput\("opacity"' -or
    $officialEditorAdapter -notmatch 'bindToolbarPropertyInput\("rotation"' -or
    $officialEditorAdapter -notmatch 'addEventListener\("input", guardedHandler\)' -or
    $officialEditorAdapter -notmatch 'addEventListener\("change", guardedHandler\)') {
    throw 'editor toolbar property controls must apply changes immediately while the user types or picks colors.'
}

if ($officialEditorAdapter -notmatch 'let lastAppliedValue = control\.value;' -or
    $officialEditorAdapter -notmatch 'const guardedHandler = event => \{[\s\S]*?if \(nextValue === lastAppliedValue\) return;[\s\S]*?lastAppliedValue = nextValue;[\s\S]*?handler\(event\);[\s\S]*?\};') {
    throw 'editor toolbar property controls must ignore duplicate input/change events for the same value.'
}

if ($officialEditorAdapter -notmatch 'control\.dataset\.lastAppliedValue = lastAppliedValue;' -or
    $officialEditorAdapter -notmatch 'lastAppliedValue = control\.dataset\.lastAppliedValue \?\? lastAppliedValue;' -or
    $officialEditorAdapter -notmatch 'function setToolbarValue\(' -or
    $officialEditorAdapter -notmatch 'control\.dataset\.lastAppliedValue = control\.value;' -or
    $officialEditorAdapter -notmatch 'setToolbarValue\(fontInput, edit\.fontName \|\| state\.fontName\)') {
    throw 'editor toolbar cached values must stay synchronized when selecting a different overlay edit.'
}

if ($officialEditorAdapter -notmatch 'asta-editor-resize-handle' -or $officialEditorAdapter -notmatch 'function startResize\(') {
    throw 'editor adapter must support resizing selected overlay objects.'
}

$startDragBlock = [regex]::Match($officialEditorAdapter, 'function startDrag\(event, editId\) \{[\s\S]*?\n  \}').Value

if (-not $startDragBlock -or
    $startDragBlock -notmatch 'let historyRecorded = false;' -or
    $startDragBlock -notmatch 'let changed = false;' -or
    $startDragBlock -notmatch 'const recordDragHistory = \(\) => \{[\s\S]*?recordHistory\(\);[\s\S]*?\};' -or
    $startDragBlock -notmatch 'recordDragHistory\(\);[\s\S]*?changed = true;' -or
    $startDragBlock -notmatch 'if \(!changed\) return;[\s\S]*?state\.dirty = true;') {
    throw 'dragging an overlay edit must record history and dirty state only after the edit actually moves.'
}

$startResizeBlock = [regex]::Match($officialEditorAdapter, 'function startResize\(event, editId\) \{[\s\S]*?\n  \}').Value

if (-not $startResizeBlock -or
    $startResizeBlock -notmatch 'let historyRecorded = false;' -or
    $startResizeBlock -notmatch 'let changed = false;' -or
    $startResizeBlock -notmatch 'const recordResizeHistory = \(\) => \{[\s\S]*?recordHistory\(\);[\s\S]*?\};' -or
    $startResizeBlock -notmatch 'recordResizeHistory\(\);[\s\S]*?changed = true;' -or
    $startResizeBlock -notmatch 'if \(!changed\) return;[\s\S]*?state\.dirty = true;') {
    throw 'resizing an overlay edit must record history and dirty state only after the edit size actually changes.'
}

if ($officialEditorAdapter -notmatch 'function undo\(\)' -or $officialEditorAdapter -notmatch 'function redo\(\)') {
    throw 'editor adapter must support undo and redo for overlay edits.'
}

if ($officialEditorAdapter -notmatch 'function copySelected\(' -or $officialEditorAdapter -notmatch 'function pasteCopiedEdit\(' -or $officialEditorAdapter -notmatch 'data-action="duplicate"') {
    throw 'editor adapter must support copying, pasting, and duplicating selected overlay edits.'
}

if ($officialEditorAdapter -notmatch 'function changeSelectedLayerOrder\(' -or $officialEditorAdapter -notmatch 'data-action="bringForward"' -or $officialEditorAdapter -notmatch 'data-action="sendBackward"' -or $officialEditorAdapter -notmatch 'zIndex') {
    throw 'editor adapter must support layer order controls for overlapping overlay edits.'
}

if ($officialEditorAdapter -notmatch 'data-role="opacity"' -or $officialEditorAdapter -notmatch 'function normalizeOpacity\(' -or $officialEditorAdapter -notmatch 'edit\.opacity = properties\.opacity') {
    throw 'editor adapter must expose and persist per-object opacity controls.'
}

if ($officialEditorAdapter -notmatch 'data-role="rotation"' -or $officialEditorAdapter -notmatch 'function normalizeRotation\(' -or $officialEditorAdapter -notmatch 'edit\.rotate = properties\.rotate') {
    throw 'editor adapter must expose and persist per-object rotation controls.'
}

if ($officialEditorAdapter -notmatch 'function buildSelectedPropertySnapshot\(' -or
    $officialEditorAdapter -notmatch 'function arePropertySnapshotsEqual\(' -or
    $officialEditorAdapter -notmatch 'const beforeProperties = buildSelectedPropertySnapshot\(edit\);' -or
    $officialEditorAdapter -notmatch 'const afterProperties = buildSelectedPropertySnapshot\(edit, properties\);' -or
    $officialEditorAdapter -notmatch 'if \(arePropertySnapshotsEqual\(beforeProperties, afterProperties\)\) return;[\s\S]*?recordHistory\(\);') {
    throw 'editor adapter must not record history or dirty state when toolbar properties do not change the selected edit.'
}

if ($officialEditorAdapter -notmatch 'rotate:\s*normalizeRotation\(edit\.rotate, 0\)') {
    throw 'editor adapter must export normalized rotation for rotatable overlay edits.'
}

$textHighlightExportBlock = [regex]::Match($officialEditorAdapter, 'if \(edit\.type === "textHighlight"\) \{[\s\S]*?return \{\s*type:\s*"textHighlight",[\s\S]*?\};').Value

if (-not $textHighlightExportBlock -or $textHighlightExportBlock -notmatch 'rotate:\s*normalizeRotation\(edit\.rotate, 0\)') {
    throw 'editor adapter must export rotation for selected text highlight edits.'
}

$whiteoutExportBlock = [regex]::Match($officialEditorAdapter, 'if \(edit\.type === "whiteout"\) \{[\s\S]*?return \{\s*type:\s*"whiteout",[\s\S]*?\};').Value

if (-not $whiteoutExportBlock -or $whiteoutExportBlock -notmatch 'rotate:\s*normalizeRotation\(edit\.rotate, 0\)') {
    throw 'editor adapter must export rotation for whiteout and redaction edits.'
}

$lineArrowExportBlock = [regex]::Match($officialEditorAdapter, 'if \(edit\.type === "line" \|\| edit\.type === "arrow"\) \{[\s\S]*?(?=if \(edit\.type === "ink"\))').Value

if (-not $lineArrowExportBlock -or $lineArrowExportBlock -notmatch 'borderWidth:\s*\(edit\.borderWidth \|\| 2\) \* Math\.max\(scaleX, scaleY\)') {
    throw 'editor adapter must export line and arrow stroke widths using the strongest page scale.'
}

if (-not $lineArrowExportBlock -or $lineArrowExportBlock -notmatch 'rotate:\s*normalizeRotation\(edit\.rotate, 0\)') {
    throw 'editor adapter must export rotation for line and arrow edits.'
}

if ($lineArrowExportBlock -match 'borderWidth:\s*\(edit\.borderWidth \|\| 2\) \* scaleX') {
    throw 'editor adapter must not export line and arrow stroke widths using only scaleX.'
}

$shapeExportBlock = [regex]::Match($officialEditorAdapter, 'return \{\s*type:\s*"rectangle",[\s\S]*?\};').Value

if (-not $shapeExportBlock -or $shapeExportBlock -notmatch 'borderWidth:\s*\(edit\.borderWidth \|\| 2\) \* Math\.max\(scaleX, scaleY\)') {
    throw 'editor adapter must export rectangle and ellipse border widths using the strongest page scale.'
}

if ($shapeExportBlock -match 'borderWidth:\s*\(edit\.borderWidth \|\| 2\) \* scaleX') {
    throw 'editor adapter must not export rectangle and ellipse border widths using only scaleX.'
}

if ($shapeExportBlock -notmatch 'rotate:\s*normalizeRotation\(edit\.rotate, 0\)') {
    throw 'editor adapter must export rotation for rectangle and ellipse edits.'
}

$stampExportBlock = [regex]::Match($officialEditorAdapter, 'if \(edit\.type === "stamp"\) \{[\s\S]*?return \{\s*type:\s*"stamp",[\s\S]*?\};').Value

if (-not $stampExportBlock -or $stampExportBlock -notmatch 'borderWidth:\s*\(edit\.borderWidth \|\| 3\) \* Math\.max\(scaleX, scaleY\)') {
    throw 'editor adapter must export stamp border widths using the strongest page scale.'
}

if ($stampExportBlock -match 'borderWidth:\s*\(edit\.borderWidth \|\| 3\) \* scaleX') {
    throw 'editor adapter must not export stamp border widths using only scaleX.'
}

if ($officialEditorAdapter -match [char]0xfffd) {
    throw 'editor adapter UI strings must not contain replacement characters from broken encoding.'
}

if ($officialPdfLibAdapter -notmatch 'PDFLib\.PDFDocument\.load') {
    throw 'pdf-lib adapter must load the source PDF through pdf-lib.'
}

if ($officialPdfLibAdapter -notmatch 'pdfDoc\.registerFontkit') {
    throw 'pdf-lib adapter must register fontkit before embedding custom fonts.'
}

if ($officialPdfLibAdapter -notmatch 'pdfDoc\.embedFont') {
    throw 'pdf-lib adapter must embed fonts for text overlay edits.'
}

if ($officialPdfLibAdapter -notmatch 'async function drawTextOverlay\(' -or $officialPdfLibAdapter -notmatch 'String\(edit\.text \?\? ""\)\.split\(/\\r\\n\|\\r\|\\n/\)' -or $officialPdfLibAdapter -notmatch 'Number\(edit\.lineHeight\)') {
    throw 'pdf-lib adapter must persist multiline text overlay edits with line-height.'
}

if ($officialPdfLibAdapter -notmatch 'drawEllipse' -or $officialPdfLibAdapter -notmatch 'drawLine') {
    throw 'pdf-lib adapter must persist ellipse and line overlay edits.'
}

if ($officialPdfLibAdapter -notmatch 'function drawArrow\(') {
    throw 'pdf-lib adapter must persist arrow overlay edits.'
}

if ($officialPdfLibAdapter -notmatch 'function getLineEndpoints\(' -or $officialPdfLibAdapter -notmatch 'const centerY = pageHeight - yFromTop - height / 2' -or $officialPdfLibAdapter -notmatch 'const \{ start, end \} = getLineEndpoints\(edit, pageHeight\)') {
    throw 'pdf-lib adapter must persist line and arrow overlays using the same centered horizontal geometry as the editor UI.'
}

if ($officialPdfLibAdapter -notmatch 'function rotatePointAroundCenter\(' -or $officialPdfLibAdapter -notmatch 'getLineEndpoints\(edit, pageHeight\)[\s\S]*?rotatePointAroundCenter' -or $officialPdfLibAdapter -notmatch 'drawArrow\([\s\S]*?getLineEndpoints\(edit, pageHeight\)') {
    throw 'pdf-lib adapter must persist line and arrow rotation around the same center as the editor UI.'
}

if ($officialPdfLibAdapter -notmatch 'function drawArrow\([\s\S]*?const opacity = editOpacity\(edit, 1\)' -or $officialPdfLibAdapter -notmatch 'drawArrow\([\s\S]*?page\.drawLine\(\{[\s\S]*?opacity') {
    throw 'pdf-lib adapter must persist opacity for arrow overlay edits.'
}

if ($officialPdfLibAdapter -notmatch 'function embedImage\(' -or $officialPdfLibAdapter -notmatch 'page\.drawImage') {
    throw 'pdf-lib adapter must persist image and signature overlay edits.'
}

if ($officialPdfLibAdapter -notmatch 'function drawInkPath\(' -or $officialPdfLibAdapter -notmatch 'edit\.type === "ink"') {
    throw 'pdf-lib adapter must persist freehand pen and highlight annotations.'
}

if ($officialPdfLibAdapter -notmatch 'function drawInkPath\([\s\S]*?editOpacity\(edit, edit\.tool === "highlight" \? 0\.38 : 1\)') {
    throw 'pdf-lib adapter must normalize opacity for freehand pen and highlight annotations.'
}

if ($officialPdfLibAdapter -notmatch 'sort\(\(left, right\) => \(Number\(left\.zIndex\) \|\| 0\) - \(Number\(right\.zIndex\) \|\| 0\)\)') {
    throw 'pdf-lib adapter must preserve overlay layer order when saving edits.'
}

if ($officialPdfLibAdapter -notmatch 'const pageCount = pdfDoc\.getPageCount\(\)' -or
    $officialPdfLibAdapter -notmatch 'if \(pageIndex < 0 \|\| pageIndex >= pageCount\) \{[\s\S]*?continue;[\s\S]*?\}') {
    throw 'pdf-lib adapter must skip stale overlay edits whose page no longer exists in the saved PDF.'
}

if ($officialPdfLibAdapter -notmatch 'function editOpacity\(' -or $officialPdfLibAdapter -notmatch 'page\.drawImage\(image,\s*\{[\s\S]*?opacity:\s*editOpacity\(edit, 1\)' -or $officialPdfLibAdapter -notmatch 'async function drawTextOverlay\([\s\S]*?opacity:\s*editOpacity\(edit, 1\)') {
    throw 'pdf-lib adapter must persist opacity for text and image overlay edits.'
}

if ($officialPdfLibAdapter -notmatch 'function editRotation\(' -or $officialPdfLibAdapter -notmatch 'page\.drawImage\(image,\s*\{[\s\S]*?rotate:\s*editRotation\(pdfLib, edit\)' -or $officialPdfLibAdapter -notmatch 'page\.drawRectangle\(\{[\s\S]*?rotate:\s*editRotation\(pdfLib, edit\)') {
    throw 'pdf-lib adapter must persist rotation for rotatable overlay edits.'
}

if ($officialPdfLibAdapter -notmatch 'pdfLib\.degrees\(-degrees\)') {
    throw 'pdf-lib adapter must convert CSS clockwise rotation into PDF coordinates.'
}

if ($officialPdfLibAdapter -notmatch 'function getRotatedBoxOrigin\(' -or $officialPdfLibAdapter -notmatch 'function getRotatedBoxPoint\(') {
    throw 'pdf-lib adapter must rotate box-based overlay edits around the same center as the editor UI.'
}

if ($officialPdfLibAdapter -notmatch 'page\.drawImage\(image,\s*\{[\s\S]*?x:\s*imageOrigin\.x,[\s\S]*?y:\s*imageOrigin\.y') {
    throw 'pdf-lib adapter must persist image and signature rotation around the editor box center.'
}

if ($officialPdfLibAdapter -notmatch 'edit\.type === "textHighlight"') {
    throw 'pdf-lib adapter must persist selected text highlight annotations.'
}

if ($officialPdfLibAdapter -notmatch 'edit\.type === "textHighlight"[\s\S]*?opacity:\s*editOpacity\(edit, 0\.38\)') {
    throw 'pdf-lib adapter must normalize opacity for selected text highlight annotations.'
}

if ($officialPdfLibAdapter -notmatch 'edit\.type === "whiteout"') {
    throw 'pdf-lib adapter must persist visual whiteout edits.'
}

if ($officialPdfLibAdapter -notmatch 'const isRedaction = edit\.variant === "redact"' -or $officialPdfLibAdapter -notmatch 'isRedaction \? \[0, 0, 0\]' -or $officialPdfLibAdapter -notmatch 'isRedaction \? 1 : editOpacity\(edit, 1\)') {
    throw 'pdf-lib adapter must persist redaction edits as opaque black regions.'
}

if ($officialPdfLibAdapter -notmatch 'edit\.type === "stamp"') {
    throw 'pdf-lib adapter must persist stamp overlay edits.'
}

if ($officialPdfLibAdapter -notmatch 'const font = await resolveFont\(pdfDoc, pdfLib, edit, fonts\)' -or $officialPdfLibAdapter -notmatch 'edit\.fontName') {
    throw 'pdf-lib adapter must persist custom Windows fonts for stamp and overlay text edits.'
}

if ($officialPdfLibAdapter -notmatch 'edit\.type === "textReplace"' -or $officialPdfLibAdapter -notmatch 'drawTextReplacement') {
    throw 'pdf-lib adapter must persist visual replacement edits for existing PDF text.'
}

if ($officialPdfLibAdapter -notmatch 'whiteoutPadding' -or $officialPdfLibAdapter -notmatch 'lineHeight') {
    throw 'pdf-lib adapter must apply text replacement metrics when saving visual text edits.'
}

$textExportBlock = [regex]::Match($officialEditorAdapter, 'if \(edit\.type === "text" \|\| edit\.type === "textReplace"\) \{[\s\S]*?(?=\n      if \(edit\.type === "line" \|\| edit\.type === "arrow"\))').Value

if (-not $textExportBlock -or $textExportBlock -notmatch 'borderWidth:\s*\(edit\.borderWidth \|\| 0\) \* Math\.max\(scaleX, scaleY\)') {
    throw 'editor adapter must export text and replacement text border widths using the strongest page scale.'
}

if ($textExportBlock -match 'borderWidth:\s*\(edit\.borderWidth \|\| 0\) \* scaleX') {
    throw 'editor adapter must not export text and replacement text border widths using only scaleX.'
}

if ($officialViewerScript -notmatch 'defaultOptions\.defaultUrl\s*=\s*\{[\s\S]*?value:\s*""') {
    throw 'official viewer must not auto-open the PDF.js sample document.'
}

if ($mainWindow -notmatch 'ViewerAssetFolderName\s*=\s*"PdfViewerOfficial"') {
    throw 'MainWindow must route WebView2 to the official PDF.js viewer assets.'
}

if ($mainWindow -notmatch 'Navigate\(BuildViewerUrl\(\)\)' -or
    $mainWindow -notmatch 'https://\{ViewerHost\}/web/viewer\.html\?enableoptimizedpartialrendering=') {
    throw 'MainWindow must navigate to the official PDF.js viewer entry point.'
}

if ($mainWindow -notmatch '/web/ServedPdf/') {
    throw 'Served PDFs must be exposed under the official viewer web root.'
}

if ($mainWindow -notmatch 'CollectEditorStateAsync') {
    throw 'MainWindow must collect editor overlay state before saving.'
}

if ($mainWindow -notmatch 'ExportOverlayPdfAsync') {
    throw 'MainWindow must export overlay edits through pdf-lib before saving.'
}

if ($mainWindow -notmatch 'overlayPdfExportFailed' -or $mainWindow -notmatch 'CompleteOverlayPdfExportFailure') {
    throw 'MainWindow must complete failed overlay PDF exports immediately instead of timing out.'
}

if ($mainWindow -notmatch 'WindowsFontService\.ReadFontBase64') {
    throw 'MainWindow must embed only the Windows fonts used by overlay text edits.'
}

if ($mainWindow -notmatch 'OnEditorTextClick' -or $mainWindow -notmatch 'editorReplaceText' -or $mainWindow -match 'editorSignature' -or $mainWindow -match 'OnEditorImageClick' -or $mainWindow -notmatch 'editorHighlight' -or $mainWindow -notmatch 'editorPen' -or $mainWindow -notmatch 'editorWhiteout' -or $mainWindow -notmatch 'editorRedact' -or $mainWindow -notmatch 'editorUnderline' -or $mainWindow -notmatch 'editorStrikeout' -or $mainWindow -notmatch 'editorDuplicateSelection' -or $mainWindow -notmatch 'editorBringForward') {
    throw 'MainWindow must expose WPF toolbar commands for PDF editor tools.'
}

if ($officialAdapter -notmatch 'case "editorRedact":' -or
    $officialAdapter -notmatch 'setMode\?\.\("redact"\)' -or
    $officialAdapter -notmatch 'case "editorUnderline":' -or
    $officialAdapter -notmatch 'setMode\?\.\("underline"\)' -or
    $officialAdapter -notmatch 'case "editorStrikeout":' -or
    $officialAdapter -notmatch 'setMode\?\.\("strikeout"\)') {
    throw 'official viewer adapter must route WPF redaction, underline, and strikeout editor commands.'
}

if ($fontService -notmatch 'CurrentVersion\\Fonts') {
    throw 'WindowsFontService must read installed Windows fonts from the registry.'
}

if ($fontService -notmatch 'NormalizeFontNames' -or $fontService -notmatch '\.Split\(' -or $fontService -notmatch '" & "') {
    throw 'WindowsFontService must expose bundled registry font family names such as Korean UI font collections as individual choices.'
}

if ($fontService -notmatch 'Convert\.ToBase64String\(File\.ReadAllBytes') {
    throw 'WindowsFontService must provide font bytes for pdf-lib embedding.'
}

if ($project -notmatch 'Assets\\PdfViewerOfficial\\\*\*\\\*') {
    throw 'PdfViewerOfficial assets must be copied to build and publish output.'
}

if ($project -match 'Assets\\PdfViewer\\\*\*\\\*') {
    throw 'Legacy custom PdfViewer assets must not be copied to build or publish output after the official viewer transition.'
}

if (Test-Path $legacyViewerPath) {
    throw 'Legacy custom PdfViewer assets must be removed after the official viewer has shipped as the active viewer.'
}

Write-Output 'viewer thumbnail rendering checks passed.'
