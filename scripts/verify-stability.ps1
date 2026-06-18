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
$releaseVerificationPath = Join-Path $root 'scripts\verify-latest-release.ps1'
$officialViewerPlanPath = Join-Path $root 'docs\superpowers\plans\2026-06-18-pdfjs-official-viewer-editor.md'
$officialViewerSpecPath = Join-Path $root 'docs\superpowers\specs\2026-06-18-pdfjs-official-viewer-editor-design.md'
$fallback = Get-Content -Raw $fallbackPath
$update = Get-Content -Raw $updatePath
$mainWindow = [System.IO.File]::ReadAllText($mainWindowPath, [System.Text.UTF8Encoding]::new($false, $true))
$installerBuild = Get-Content -Raw $installerBuildPath
$buildScript = Get-Content -Raw $buildScriptPath
$publishScript = Get-Content -Raw $publishScriptPath
$releaseWorkflow = Get-Content -Raw $releaseWorkflowPath
$releaseVerification = if (Test-Path $releaseVerificationPath) { Get-Content -Raw $releaseVerificationPath } else { '' }
$officialViewerPlan = Get-Content -Raw $officialViewerPlanPath
$officialViewerSpec = Get-Content -Raw $officialViewerSpecPath

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

if ($update -notmatch 'CheckForUpdatesViaReleaseRedirectAsync' -or
    $update -notmatch 'https://github\.com/purius/asta_pdf/releases/latest' -or
    $update -notmatch 'HttpCompletionOption\.ResponseHeadersRead' -or
    $update -notmatch 'TryGetReleaseTagFromUrl' -or
    $update -notmatch 'catch \(HttpRequestException ex\)') {
    throw 'UpdateService must fall back to the GitHub releases/latest redirect when the API is rate-limited or unavailable.'
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

if ([regex]::IsMatch($mainWindow, '\p{IsCJKUnifiedIdeographs}') -or $mainWindow.Contains([char]0xfffd)) {
    throw 'MainWindow user-facing strings must not contain mojibake or replacement characters.'
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

if (-not (Test-Path $releaseVerificationPath) -or
    $releaseVerification -notmatch 'api\.github\.com/repos/purius/asta_pdf/releases/latest' -or
    $releaseVerification -notmatch 'PdfMergeToolSetup\.exe' -or
    $releaseVerification -notmatch 'src\\PdfMergeTool\\PdfMergeTool\.csproj' -or
    $releaseVerification -notmatch 'Assert-HeadOk') {
    throw 'Latest release verification must confirm the project version, GitHub latest release metadata, and installer download URL.'
}

if ($releaseVerification -notmatch 'Test-ExpectedTagReleaseFallback' -or
    $releaseVerification -notmatch 'releases/tag/\$expectedTag' -or
    $releaseVerification -notmatch 'releases/download/\$expectedTag/\$installerAssetName' -or
    $releaseVerification -notmatch 'GitHub API verification failed after') {
    throw 'Latest release verification must fall back to expected tag release and installer URLs when GitHub API rate limits are exhausted.'
}

if ($releaseVerification -notmatch '\[int\]\$RetryCount' -or
    $releaseVerification -notmatch '\[int\]\$RetryDelaySeconds' -or
    $releaseVerification -notmatch 'for \(\$attempt = 1; \$attempt -le \$RetryCount; \$attempt\+\+\)' -or
    $releaseVerification -notmatch 'Start-Sleep -Seconds \$RetryDelaySeconds' -or
    $releaseVerification -notmatch '\$attempt -eq \$RetryCount') {
    throw 'Latest release verification must retry while GitHub release metadata and asset URLs settle.'
}

if ($releaseWorkflow -notmatch 'Verify published latest release' -or
    $releaseWorkflow -notmatch 'verify-latest-release\.ps1 -ExpectedVersion \$\{\{ github\.ref_name \}\}' -or
    $releaseWorkflow -notmatch '-RetryCount 60' -or
    $releaseWorkflow -notmatch '-RetryDelaySeconds 10') {
    throw 'Release workflow must verify the published latest release and installer URL after creating the GitHub release.'
}

if ($officialViewerPlan -match '- \[ \]' -or
    $officialViewerPlan -match 'Preserve the old custom viewer as rollback' -or
    $officialViewerPlan -notmatch 'Status: shipped in v1\.0\.' -or
    $officialViewerPlan -notmatch 'Legacy custom viewer removed') {
    throw 'Official viewer implementation plan must reflect the shipped official viewer/editor state and removed legacy viewer.'
}

if ($officialViewerSpec -match 'remains temporarily as a rollback reference' -or
    $officialViewerSpec -match 'without removing the old viewer assets' -or
    $officialViewerSpec -notmatch 'Current State' -or
    $officialViewerSpec -notmatch 'legacy custom viewer assets have been removed') {
    throw 'Official viewer design spec must reflect the current shipped architecture, not the old rollback phase.'
}

Write-Output 'stability checks passed.'
