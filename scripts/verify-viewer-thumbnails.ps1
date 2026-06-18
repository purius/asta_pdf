Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$viewerPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewer\viewer.html'
$officialViewerPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\viewer.html'
$officialAdapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\app-adapter.js'
$officialEditorAdapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\editor-adapter.js'
$officialPdfLibAdapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\pdf-lib-adapter.js'
$officialPdfLibPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\vendor\pdf-lib.min.js'
$officialFontkitPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\vendor\fontkit.umd.min.js'
$officialViewerScriptPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\viewer.mjs'
$officialBuildPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\build\pdf.mjs'
$mainWindowPath = Join-Path $root 'src\PdfMergeTool\MainWindow.xaml.cs'
$projectPath = Join-Path $root 'src\PdfMergeTool\PdfMergeTool.csproj'
$fontServicePath = Join-Path $root 'src\PdfMergeTool\Services\WindowsFontService.cs'
$viewer = Get-Content -Raw $viewerPath
$mainWindow = Get-Content -Raw $mainWindowPath
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

if ($officialEditorAdapter -notmatch 'type:\s*"editorStateChanged"') {
    throw 'editor adapter must notify WPF when overlay edits make the document dirty.'
}

if ($officialEditorAdapter -notmatch 'type:\s*"editorStateCollected"') {
    throw 'editor adapter must return bounded overlay edit state for persistence.'
}

if ($officialEditorAdapter -notmatch 'astaEditorToolbar') {
    throw 'editor adapter must expose in-viewer text and shape editing controls.'
}

if ($officialEditorAdapter -notmatch 'data-mode="ellipse"' -or $officialEditorAdapter -notmatch 'data-mode="line"') {
    throw 'editor adapter must expose ellipse and line shape editing controls.'
}

if ($officialEditorAdapter -notmatch 'data-mode="arrow"') {
    throw 'editor adapter must expose arrow editing controls.'
}

if ($officialEditorAdapter -notmatch 'data-mode="image"' -or $officialEditorAdapter -notmatch 'data-mode="stamp"' -or $officialEditorAdapter -notmatch 'data-mode="signature"') {
    throw 'editor adapter must expose image, stamp, and signature overlay controls.'
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

if ($officialEditorAdapter -notmatch 'data-role="strokeWidth"' -or $officialEditorAdapter -notmatch 'data-role="fillColor"' -or $officialEditorAdapter -notmatch 'function applySelectedProperties\(' -or $officialEditorAdapter -notmatch 'function syncToolbarFromEdit\(') {
    throw 'editor adapter must expose per-object font, size, stroke width, and fill controls.'
}

if ($officialEditorAdapter -notmatch 'asta-editor-resize-handle' -or $officialEditorAdapter -notmatch 'function startResize\(') {
    throw 'editor adapter must support resizing selected overlay objects.'
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

if ($officialPdfLibAdapter -notmatch 'function embedImage\(' -or $officialPdfLibAdapter -notmatch 'page\.drawImage') {
    throw 'pdf-lib adapter must persist image and signature overlay edits.'
}

if ($officialPdfLibAdapter -notmatch 'function drawInkPath\(' -or $officialPdfLibAdapter -notmatch 'edit\.type === "ink"') {
    throw 'pdf-lib adapter must persist freehand pen and highlight annotations.'
}

if ($officialPdfLibAdapter -notmatch 'sort\(\(left, right\) => \(Number\(left\.zIndex\) \|\| 0\) - \(Number\(right\.zIndex\) \|\| 0\)\)') {
    throw 'pdf-lib adapter must preserve overlay layer order when saving edits.'
}

if ($officialPdfLibAdapter -notmatch 'function editOpacity\(' -or $officialPdfLibAdapter -notmatch 'page\.drawImage\(image,\s*\{[\s\S]*?opacity:\s*editOpacity\(edit, 1\)' -or $officialPdfLibAdapter -notmatch 'async function drawTextOverlay\([\s\S]*?opacity:\s*editOpacity\(edit, 1\)') {
    throw 'pdf-lib adapter must persist opacity for text and image overlay edits.'
}

if ($officialPdfLibAdapter -notmatch 'edit\.type === "textHighlight"') {
    throw 'pdf-lib adapter must persist selected text highlight annotations.'
}

if ($officialPdfLibAdapter -notmatch 'edit\.type === "whiteout"') {
    throw 'pdf-lib adapter must persist visual whiteout edits.'
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

if ($officialViewerScript -notmatch 'defaultOptions\.defaultUrl\s*=\s*\{[\s\S]*?value:\s*""') {
    throw 'official viewer must not auto-open the PDF.js sample document.'
}

if ($mainWindow -notmatch 'ViewerAssetFolderName\s*=\s*"PdfViewerOfficial"') {
    throw 'MainWindow must route WebView2 to the official PDF.js viewer assets.'
}

if ($mainWindow -notmatch 'Navigate\(\$"https://\{ViewerHost\}/web/viewer\.html"\)') {
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

if ($mainWindow -notmatch 'WindowsFontService\.ReadFontBase64') {
    throw 'MainWindow must embed only the Windows fonts used by overlay text edits.'
}

if ($mainWindow -notmatch 'OnEditorTextClick' -or $mainWindow -notmatch 'editorReplaceText' -or $mainWindow -notmatch 'editorSignature' -or $mainWindow -notmatch 'editorHighlight' -or $mainWindow -notmatch 'editorPen' -or $mainWindow -notmatch 'editorWhiteout' -or $mainWindow -notmatch 'editorDuplicateSelection' -or $mainWindow -notmatch 'editorBringForward') {
    throw 'MainWindow must expose WPF toolbar commands for PDF editor tools.'
}

if ($fontService -notmatch 'CurrentVersion\\Fonts') {
    throw 'WindowsFontService must read installed Windows fonts from the registry.'
}

if ($fontService -notmatch 'Convert\.ToBase64String\(File\.ReadAllBytes') {
    throw 'WindowsFontService must provide font bytes for pdf-lib embedding.'
}

if ($project -notmatch 'Assets\\PdfViewerOfficial\\\*\*\\\*') {
    throw 'PdfViewerOfficial assets must be copied to build and publish output.'
}

if ($viewer -notmatch 'async function renderAllThumbnails\(token\)') {
    throw 'viewer.html must render thumbnails independently from the main page render window.'
}

if ($viewer -notmatch 'let thumbnailRenderKey = ''''') {
    throw 'viewer.html must track whether the thumbnail list actually needs to be re-rendered.'
}

if ($viewer -notmatch 'if \(thumbnailRenderKey === nextKey\)') {
    throw 'thumbnail rendering must skip rebuilding unchanged thumbnails during page navigation.'
}

if ($viewer -notmatch 'let lastActiveThumb = null') {
    throw 'viewer.html must track the previously active thumbnail to avoid rewriting every thumbnail on scroll.'
}

if ($viewer -notmatch 'function updateActiveThumb\(previousPage, nextPage\)') {
    throw 'viewer.html must update only changed active thumbnail nodes.'
}

if ($viewer -notmatch 'aspect-ratio:\s+var\(--thumb-aspect-ratio') {
    throw 'thumbnail media must reserve a stable aspect ratio before async rendering completes.'
}

if ($viewer -notmatch 'let isRenderingDocument = false') {
    throw 'viewer.html must suppress scroll synchronization while the main render window is rebuilding.'
}

if ($viewer -notmatch 'let programmaticScrollTargetPage = null') {
    throw 'viewer.html must distinguish programmatic page jumps from user scrolling.'
}

if ($viewer -notmatch 'function shouldIgnoreScrollEvent\(\)') {
    throw 'viewer.html must guard scroll events during programmatic navigation.'
}

if ($viewer -notmatch 'beginProgrammaticScroll\(targetPage') {
    throw 'renderDocument must mark its target scroll as programmatic before scrollIntoView.'
}

if ($viewer -notmatch 'let lastUserScrollAt = 0') {
    throw 'viewer.html must track recent direct user scrolling separately from page navigation.'
}

if ($viewer -notmatch 'function markUserScrollIntent\(\)') {
    throw 'viewer.html must record user scroll intent before allowing render-window switches.'
}

if ($viewer -notmatch 'function hasRecentUserScrollIntent\(\)') {
    throw 'viewer.html must gate render-window switches on recent user scroll intent.'
}

if ($viewer -notmatch 'function shouldSyncActivePageFromScroll\(\)') {
    throw 'viewer.html must separately gate scroll-driven active page synchronization.'
}

if ($viewer -notmatch 'if \(!shouldSyncActivePageFromScroll\(\)\) return;') {
    throw 'scroll events must not update the active page after programmatic page jumps.'
}

if ($viewer -notmatch 'let pageNavigationToken = 0') {
    throw 'viewer.html must track the newest explicit page navigation request.'
}

if ($viewer -notmatch 'function beginPageNavigation\(pageNumber\)') {
    throw 'viewer.html must assign a token to each explicit page navigation request.'
}

if ($viewer -notmatch 'function isCurrentPageNavigation\(token\)') {
    throw 'viewer.html must ignore stale page navigation completions.'
}

if ($viewer -notmatch 'async function renderDocument\(keepPage = true, targetPageOverride = null, navigationToken = null\)') {
    throw 'renderDocument must accept a navigation token so stale renders cannot scroll old pages.'
}

if ($viewer -notmatch 'if \(navigationToken !== null && !isCurrentPageNavigation\(navigationToken\)\) return;') {
    throw 'renderDocument must stop before activating or scrolling a stale navigation target.'
}

if ($viewer -notmatch 'const navigationToken = beginPageNavigation\(page\)') {
    throw 'goToPage must create a navigation token before async rendering.'
}

if ($viewer -notmatch 'async function goToPage\(pageNumber, smooth = false\)') {
    throw 'page navigation must default to immediate scrolling to avoid smooth-scroll feedback loops.'
}

if ($viewer -match "behavior: smooth \\? 'smooth' : 'auto'") {
    throw 'page navigation must not use smooth scrolling for document page jumps.'
}

if ($viewer -notmatch 'for \(let index = 0; index < pageOrder\.length; index \+= 1\)') {
    throw 'thumbnail rendering must iterate over every page in pageOrder.'
}

$clearRenderedPagesMatch = [regex]::Match(
    $viewer,
    'function clearRenderedPages\(\) \{(?<body>[\s\S]*?)\n    \}')
if (-not $clearRenderedPagesMatch.Success) {
    throw 'viewer.html must define clearRenderedPages.'
}

if ($clearRenderedPagesMatch.Groups['body'].Value -match 'thumbs\.replaceChildren\(\);') {
    throw 'main page clearing must not remove already-rendered thumbnails.'
}

if ($clearRenderedPagesMatch.Groups['body'].Value -match 'activePage\s*=\s*1;') {
    throw 'main page clearing must not reset the active thumbnail to page 1.'
}

if ($viewer -match 'thumbs\.appendChild\(thumb\);[\s\S]{0,900}await page\.render\(\{ canvasContext: pageContext') {
    throw 'main PDF page rendering must not create thumbnails only for the visible render window.'
}

if ($viewer -match 'thumbs\.appendChild\(thumb\);[\s\S]{0,900}const mainResult = await requestFallbackRender') {
    throw 'fallback main page rendering must not create thumbnails only for the visible render window.'
}

Write-Output 'viewer thumbnail rendering checks passed.'
