Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$viewerPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewer\viewer.html'
$viewer = Get-Content -Raw $viewerPath

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
