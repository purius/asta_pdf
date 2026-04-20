$ErrorActionPreference = 'Stop'

& (Join-Path $PSScriptRoot 'publish.ps1')

$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$dist = Join-Path $root 'dist'
$source = Join-Path $dist 'PdfMergeTool'
$zip = Join-Path $dist 'PdfMergeTool-win-x64.zip'

if (Test-Path $zip) {
    Remove-Item -LiteralPath $zip -Force
}

Compress-Archive -Path (Join-Path $source '*') -DestinationPath $zip -Force
Write-Host "Package: $zip"
