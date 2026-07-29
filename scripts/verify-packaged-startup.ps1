param(
    [string]$ApplicationPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\PdfMergeTool\PdfMergeTool.exe')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $ApplicationPath)) {
    throw "Packaged application is missing: $ApplicationPath"
}

$process = $null
try {
    $process = Start-Process -FilePath $ApplicationPath -PassThru
    Start-Sleep -Seconds 10
    $process.Refresh()
    if ($process.HasExited) {
        throw "Packaged application exited during startup smoke test. Exit code: $($process.ExitCode)"
    }

    Write-Output 'packaged application startup smoke test passed.'
}
finally {
    if ($null -ne $process) {
        $process.Refresh()
        if (-not $process.HasExited) {
            Stop-Process -Id $process.Id -Force
            $process.WaitForExit()
        }
    }
}
