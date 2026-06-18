Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$fallbackPath = Join-Path $root 'src\PdfMergeTool\Services\PdfFallbackRenderService.cs'
$updatePath = Join-Path $root 'src\PdfMergeTool\Services\UpdateService.cs'
$mainWindowPath = Join-Path $root 'src\PdfMergeTool\MainWindow.xaml.cs'
$installerBuildPath = Join-Path $root 'scripts\build-installer.ps1'
$buildScriptPath = Join-Path $root 'scripts\build.ps1'
$publishScriptPath = Join-Path $root 'scripts\publish.ps1'
$releaseWorkflowPath = Join-Path $root '.github\workflows\release.yml'
$fallback = Get-Content -Raw $fallbackPath
$update = Get-Content -Raw $updatePath
$mainWindow = Get-Content -Raw $mainWindowPath
$installerBuild = Get-Content -Raw $installerBuildPath
$buildScript = Get-Content -Raw $buildScriptPath
$publishScript = Get-Content -Raw $publishScriptPath
$releaseWorkflow = Get-Content -Raw $releaseWorkflowPath

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

if ($mainWindow -notmatch 'RemapEditorStateToOutputPageOrder' -or
    $mainWindow -notmatch 'TryGetProperty\("page"' -or
    $mainWindow -notmatch 'pageToOutputIndex' -or
    $mainWindow -notmatch 'remappedEditorState' -or
    $mainWindow -notmatch 'remappedEditorState\.Edits\.Count > 0 \? CreateTempPdfPath\("editor-source"\) : outputPath') {
    throw 'Overlay edits must be remapped from original PDF page numbers to saved output page order before export.'
}

if ($mainWindow -notmatch 'string\? transformedTempPath = null;' -or
    $mainWindow -notmatch 'finally[\s\S]*?TryDeleteTempFile\(transformedTempPath\)') {
    throw 'Overlay editor-source temporary PDFs must be cleaned up in a finally block after save attempts.'
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

if ($installerBuild -notmatch 'previousErrorActionPreference' -or
    $installerBuild -notmatch "\`$ErrorActionPreference = 'Continue'") {
    throw 'Installer build must capture Inno Setup stderr without letting PowerShell native-command errors bypass retry handling.'
}

if ($buildScript -notmatch 'AstaPdf\.DotNetBuild' -or
    $buildScript -notmatch '\.WaitOne\(' -or
    $buildScript -notmatch 'finally[\s\S]*?ReleaseMutex\(\)' -or
    $publishScript -notmatch 'AstaPdf\.DotNetBuild' -or
    $publishScript -notmatch '\.WaitOne\(' -or
    $publishScript -notmatch 'finally[\s\S]*?ReleaseMutex\(\)') {
    throw 'Build and publish scripts must serialize dotnet restore/build/publish operations with the same mutex.'
}

if ($releaseWorkflow -notmatch 'verify-viewer-thumbnails\.ps1' -or
    $releaseWorkflow -notmatch 'verify-stability\.ps1') {
    throw 'Release workflow must run viewer thumbnail and stability verifications before publishing an installer.'
}

if ($releaseWorkflow -notmatch 'node --check src/PdfMergeTool/Assets/PdfViewerOfficial/web/app-adapter\.js' -or
    $releaseWorkflow -notmatch 'node --check src/PdfMergeTool/Assets/PdfViewerOfficial/web/editor-adapter\.js' -or
    $releaseWorkflow -notmatch 'node --check src/PdfMergeTool/Assets/PdfViewerOfficial/web/pdf-lib-adapter\.js') {
    throw 'Release workflow must syntax-check official viewer adapter JavaScript before publishing an installer.'
}

if ($releaseWorkflow -match '(?s)name: Build installer.*?verify-viewer-thumbnails\.ps1') {
    throw 'Release workflow must run verification steps before building the installer.'
}

Write-Output 'stability checks passed.'
