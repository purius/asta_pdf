Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$fallbackPath = Join-Path $root 'src\PdfMergeTool\Services\PdfFallbackRenderService.cs'
$updatePath = Join-Path $root 'src\PdfMergeTool\Services\UpdateService.cs'
$mainWindowPath = Join-Path $root 'src\PdfMergeTool\MainWindow.xaml.cs'
$installerBuildPath = Join-Path $root 'scripts\build-installer.ps1'
$fallback = Get-Content -Raw $fallbackPath
$update = Get-Content -Raw $updatePath
$mainWindow = Get-Content -Raw $mainWindowPath
$installerBuild = Get-Content -Raw $installerBuildPath

if ($fallback -notmatch 'private const int MaxRenderCacheEntries') {
    throw 'PdfFallbackRenderService must cap per-session render cache entries.'
}

if ($fallback -notmatch 'TrimRenderCache\(session\);') {
    throw 'PdfFallbackRenderService must trim old cached fallback render files after rendering.'
}

if ($fallback -notmatch 'await Task\.Run\(\(\) =>\s*\{[\s\S]*?session\.Document\.Render') {
    throw 'Fallback PDF rendering must run off the UI continuation thread.'
}

$exportMethod = [regex]::Match(
    $fallback,
    'public async Task<IReadOnlyList<PdfImagePage>> ExportPagesAsImagesAsync\([\s\S]*?\n    \}')
if (-not $exportMethod.Success -or
    $exportMethod.Groups[0].Value -notmatch 'await session\.RenderGate\.WaitAsync\(cancellationToken\)') {
    throw 'Fallback A4 export must share the session render gate with page rendering.'
}

if ($update -notmatch 'private static readonly HttpClient HttpClient') {
    throw 'UpdateService must reuse a static HttpClient instead of creating one per check.'
}

if ($update -notmatch 'Timeout = TimeSpan\.FromSeconds\(10\)') {
    throw 'UpdateService must use a bounded timeout for update checks.'
}

if ($update -notmatch 'Process\.Start\(new ProcessStartInfo\(result\.InstallerUrl\.AbsoluteUri\)') {
    throw 'UpdateService must open the installer asset URL, not only the release page.'
}

if ($update -notmatch 'BuildInstallerDownloadUrl\(latestTag\)' -or
    $update -notmatch 'releases/download/\{Uri\.EscapeDataString\(tag\.Trim\(\)\)\}/\{InstallerAssetName\}') {
    throw 'UpdateService must fall back to a tag-based installer download URL when release assets are delayed.'
}

if ($mainWindow -notmatch '_editorStateRequestId' -or
    $mainWindow -notmatch '_overlayPdfExportRequestId' -or
    $mainWindow -notmatch 'IsExpectedRequest\(root,\s*_editorStateRequestId' -or
    $mainWindow -notmatch 'IsExpectedRequest\(root,\s*_overlayPdfExportRequestId') {
    throw 'WebView async editor responses must be matched by requestId before completing pending WPF tasks.'
}

if ($installerBuild -notmatch 'Invoke-InnoCompilerWithRetry' -or
    $installerBuild -notmatch 'Start-Sleep -Seconds' -or
    $installerBuild -notmatch 'attempt -lt') {
    throw 'Installer build must retry Inno Setup after transient file lock failures.'
}

if ($installerBuild -notmatch '2>&1' -or
    $installerBuild -notmatch 'Compile aborted' -or
    $installerBuild -notmatch 'Error in') {
    throw 'Installer build must treat Inno Setup error output as a failed attempt even when an installer file exists.'
}

Write-Output 'stability checks passed.'
