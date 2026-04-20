$ErrorActionPreference = 'Stop'

$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$dotnet = Join-Path $root '.tools\dotnet\dotnet.exe'
$project = Join-Path $root 'src\PdfMergeTool\PdfMergeTool.csproj'
$publishDir = Join-Path $root 'dist\PdfMergeTool'
$qpdfSource = Join-Path $root '.tools\qpdf\qpdf-12.3.2-msvc64\bin'
$qpdfTarget = Join-Path $publishDir 'tools\qpdf'

& (Join-Path $PSScriptRoot 'restore-tools.ps1')

if (Test-Path $publishDir) {
    Remove-Item -LiteralPath $publishDir -Recurse -Force
}

& $dotnet publish $project `
    --configuration Release `
    --runtime win-x64 `
    --self-contained true `
    -p:PublishSingleFile=false `
    -p:PublishReadyToRun=false `
    -p:SatelliteResourceLanguages=ko `
    --output $publishDir

$allowedCultureDirs = @('ko')
Get-ChildItem -LiteralPath $publishDir -Directory | Where-Object {
    $_.Name -match '^[a-z]{2}(-[A-Z][A-Za-z]+)?$' -and $allowedCultureDirs -notcontains $_.Name
} | Remove-Item -Recurse -Force

New-Item -ItemType Directory -Path $qpdfTarget -Force | Out-Null
Copy-Item -Path (Join-Path $qpdfSource '*') -Destination $qpdfTarget -Recurse -Force

$exe = Join-Path $publishDir 'PdfMergeTool.exe'
if (-not (Test-Path $exe)) {
    throw "Publish failed. Missing $exe"
}

Write-Host "Published: $exe"
Write-Host "Bundled qpdf: $(Join-Path $qpdfTarget 'qpdf.exe')"
