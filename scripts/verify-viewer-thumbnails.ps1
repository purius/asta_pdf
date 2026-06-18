Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$viewerPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewer\viewer.html'
$officialViewerPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\viewer.html'
$officialAdapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\app-adapter.js'
$officialPdfLibAdapterPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\pdf-lib-adapter.js'
$officialPdfLibPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\vendor\pdf-lib.min.js'
$officialFontkitPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\vendor\fontkit.umd.min.js'
$officialViewerScriptPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\web\viewer.mjs'
$officialBuildPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewerOfficial\build\pdf.mjs'
$mainWindowPath = Join-Path $root 'src\PdfMergeTool\MainWindow.xaml.cs'
$projectPath = Join-Path $root 'src\PdfMergeTool\PdfMergeTool.csproj'
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
$officialPdfLibAdapter = Get-Content -Raw $officialPdfLibAdapterPath
$officialViewerScript = Get-Content -Raw $officialViewerScriptPath

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

if ($officialPdfLibAdapter -notmatch 'PDFLib\.PDFDocument\.load') {
    throw 'pdf-lib adapter must load the source PDF through pdf-lib.'
}

if ($officialPdfLibAdapter -notmatch 'pdfDoc\.registerFontkit') {
    throw 'pdf-lib adapter must register fontkit before embedding custom fonts.'
}

if ($officialPdfLibAdapter -notmatch 'pdfDoc\.embedFont') {
    throw 'pdf-lib adapter must embed fonts for text overlay edits.'
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
