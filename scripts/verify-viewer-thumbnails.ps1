Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$viewerPath = Join-Path $root 'src\PdfMergeTool\Assets\PdfViewer\viewer.html'
$viewer = Get-Content -Raw $viewerPath

if ($viewer -notmatch 'async function renderAllThumbnails\(token\)') {
    throw 'viewer.html must render thumbnails independently from the main page render window.'
}

if ($viewer -notmatch 'for \(let index = 0; index < pageOrder\.length; index \+= 1\)') {
    throw 'thumbnail rendering must iterate over every page in pageOrder.'
}

if ($viewer -match 'thumbs\.appendChild\(thumb\);[\s\S]{0,900}await page\.render\(\{ canvasContext: pageContext') {
    throw 'main PDF page rendering must not create thumbnails only for the visible render window.'
}

if ($viewer -match 'thumbs\.appendChild\(thumb\);[\s\S]{0,900}const mainResult = await requestFallbackRender') {
    throw 'fallback main page rendering must not create thumbnails only for the visible render window.'
}

Write-Output 'viewer thumbnail rendering checks passed.'
