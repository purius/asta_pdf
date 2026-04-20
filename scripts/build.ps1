$ErrorActionPreference = 'Stop'

$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$dotnet = Join-Path $root '.tools\dotnet\dotnet.exe'
$project = Join-Path $root 'src\PdfMergeTool\PdfMergeTool.csproj'

if (-not (Test-Path $dotnet)) {
    & (Join-Path $PSScriptRoot 'restore-tools.ps1')
}

& $dotnet build $project
