$ErrorActionPreference = 'Stop'

$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$dotnet = Join-Path $root '.tools\dotnet\dotnet.exe'
$dist = Join-Path $root 'dist'
$zip = Join-Path $dist 'PdfMergeTool-win-x64.zip'
$installerProject = Join-Path $root 'src\PdfMergeTool.Installer\PdfMergeTool.Installer.csproj'
$installerProjectDir = Split-Path -Parent $installerProject
$payloadDir = Join-Path $root 'src\PdfMergeTool.Installer\Payload'
$payload = Join-Path $payloadDir 'PdfMergeTool-win-x64.zip'
$publishDir = Join-Path $dist 'PdfMergeToolSetup-publish'
$setupExe = Join-Path $publishDir 'PdfMergeToolSetup.exe'
$asciiTarget = Join-Path $dist 'PdfMergeToolSetup.exe'

& (Join-Path $PSScriptRoot 'package.ps1')

New-Item -ItemType Directory -Path $payloadDir -Force | Out-Null
Copy-Item -LiteralPath $zip -Destination $payload -Force

Remove-Item -LiteralPath (Join-Path $installerProjectDir 'bin') -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $installerProjectDir 'obj') -Recurse -Force -ErrorAction SilentlyContinue

if (Test-Path $publishDir) {
    Remove-Item -LiteralPath $publishDir -Recurse -Force
}

& $dotnet publish $installerProject `
    --configuration Release `
    --runtime win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:PublishTrimmed=false `
    -p:SatelliteResourceLanguages=ko `
    --output $publishDir

if (-not (Test-Path $setupExe)) {
    throw "Installer publish failed. Missing $setupExe"
}

Copy-Item -LiteralPath $setupExe -Destination $asciiTarget -Force
Remove-Item -LiteralPath $payload -Force

Write-Host "Installer: $asciiTarget"
