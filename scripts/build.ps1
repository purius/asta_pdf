$ErrorActionPreference = 'Stop'

$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$dotnet = Join-Path $root '.tools\dotnet\dotnet.exe'
$project = Join-Path $root 'src\PdfMergeTool\PdfMergeTool.csproj'
$mutex = [System.Threading.Mutex]::new($false, 'AstaPdf.DotNetBuild')
$mutexAcquired = $false

try {
    $mutexAcquired = $mutex.WaitOne([TimeSpan]::FromMinutes(10))
    if (-not $mutexAcquired) {
        throw 'Timed out waiting for another dotnet build or publish operation to finish.'
    }

    if (-not (Test-Path $dotnet)) {
        & (Join-Path $PSScriptRoot 'restore-tools.ps1')
    }

    & $dotnet build $project
}
finally {
    if ($mutexAcquired) {
        $mutex.ReleaseMutex()
    }

    $mutex.Dispose()
}
