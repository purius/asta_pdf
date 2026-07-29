param(
    [string]$ApplicationPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\PdfMergeTool\PdfMergeTool.exe')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-StartupFailureDiagnostics {
    $logsDirectory = Join-Path $env:LOCALAPPDATA 'PdfMergeTool\Logs'
    if (-not (Test-Path $logsDirectory)) {
        Write-Host "PdfMergeTool startup logs were not found: $logsDirectory"
        return
    }

    Write-Host '--- PdfMergeTool startup logs ---'
    Get-ChildItem -Path $logsDirectory -File |
        Sort-Object LastWriteTime |
        Select-Object -Last 2 |
        ForEach-Object {
            Write-Host "--- $($_.FullName) ---"
            Get-Content -LiteralPath $_.FullName -Tail 200
        }
}

if (-not (Test-Path $ApplicationPath)) {
    throw "Packaged application is missing: $ApplicationPath"
}

$process = $null
try {
    $process = Start-Process -FilePath $ApplicationPath -PassThru
    Start-Sleep -Seconds 10
    $process.Refresh()
    if ($process.HasExited) {
        Write-StartupFailureDiagnostics
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
