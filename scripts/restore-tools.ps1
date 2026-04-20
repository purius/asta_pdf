$ErrorActionPreference = 'Stop'

$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$tools = Join-Path $root '.tools'
$dotnetDir = Join-Path $tools 'dotnet'
$dotnetZip = Join-Path $tools 'dotnet-sdk-8.0.420-win-x64.zip'
$qpdfDir = Join-Path $tools 'qpdf'
$qpdfZip = Join-Path $tools 'qpdf-12.3.2-msvc64.zip'

New-Item -ItemType Directory -Path $tools -Force | Out-Null

if (-not (Test-Path (Join-Path $dotnetDir 'sdk\8.0.420\Sdks\Microsoft.NET.Sdk\Sdk\Sdk.props'))) {
    if (Test-Path $dotnetDir) {
        Remove-Item -LiteralPath $dotnetDir -Recurse -Force
    }

    New-Item -ItemType Directory -Path $dotnetDir -Force | Out-Null

    if (-not (Test-Path $dotnetZip)) {
        Invoke-WebRequest -Uri 'https://builds.dotnet.microsoft.com/dotnet/Sdk/8.0.420/dotnet-sdk-8.0.420-win-x64.zip' -OutFile $dotnetZip
    }

    tar -xf $dotnetZip -C $dotnetDir
}

if (-not (Test-Path (Join-Path $qpdfDir 'qpdf-12.3.2-msvc64\bin\qpdf.exe'))) {
    if (Test-Path $qpdfDir) {
        Remove-Item -LiteralPath $qpdfDir -Recurse -Force
    }

    New-Item -ItemType Directory -Path $qpdfDir -Force | Out-Null

    if (-not (Test-Path $qpdfZip)) {
        Invoke-WebRequest -Uri 'https://github.com/qpdf/qpdf/releases/download/v12.3.2/qpdf-12.3.2-msvc64.zip' -OutFile $qpdfZip
    }

    tar -xf $qpdfZip -C $qpdfDir
}

& (Join-Path $dotnetDir 'dotnet.exe') --info
& (Join-Path $qpdfDir 'qpdf-12.3.2-msvc64\bin\qpdf.exe') --version
