$ErrorActionPreference = 'Stop'

$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$dist = Join-Path $root 'dist'
$project = Join-Path $root 'src\PdfMergeTool\PdfMergeTool.csproj'
$iss = Join-Path $root 'installer\PdfMergeTool.iss'
$setupExe = Join-Path $dist 'PdfMergeToolSetup.exe'

function Get-InnoCompiler {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'),
        'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
        'C:\Program Files\Inno Setup 6\ISCC.exe'
    )

    $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($winget) {
        Write-Host 'Inno Setup compiler not found. Installing Inno Setup with winget...'
        & $winget.Source install --id JRSoftware.InnoSetup --exact --silent --accept-package-agreements --accept-source-agreements
        foreach ($candidate in $candidates) {
            if (Test-Path $candidate) {
                return $candidate
            }
        }
    }

    throw 'Inno Setup compiler ISCC.exe was not found. Install Inno Setup 6 and run this script again.'
}

function Invoke-InnoCompilerWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CompilerPath,
        [Parameter(Mandatory = $true)]
        [string[]]$CompilerArgs,
        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $maxAttempts = 3
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        if (Test-Path $OutputPath) {
            Remove-Item -LiteralPath $OutputPath -Force
        }

        & $CompilerPath @CompilerArgs
        if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputPath)) {
            return
        }

        if ($attempt -lt $maxAttempts) {
            Write-Warning "Inno Setup failed on attempt $attempt of $maxAttempts. Retrying after transient file handles settle..."
            Start-Sleep -Seconds (2 * $attempt)
            continue
        }

        if ($LASTEXITCODE -ne 0) {
            throw "Inno Setup failed with exit code $LASTEXITCODE after $maxAttempts attempts."
        }
    }
}

$version = ([xml](Get-Content $project)).Project.PropertyGroup.Version
if ([string]::IsNullOrWhiteSpace($version)) {
    throw "Missing Version in $project"
}

& (Join-Path $PSScriptRoot 'publish.ps1')

if (Test-Path $setupExe) {
    Remove-Item -LiteralPath $setupExe -Force
}

$iscc = Get-InnoCompiler
Invoke-InnoCompilerWithRetry `
    -CompilerPath $iscc `
    -CompilerArgs @(
        "/DRootDir=$root",
        "/DSourceDir=$(Join-Path $dist 'PdfMergeTool')",
        "/DOutputDir=$dist",
        "/DAppVersion=$version",
        $iss
    ) `
    -OutputPath $setupExe

if (-not (Test-Path $setupExe)) {
    throw "Installer build failed. Missing $setupExe"
}

Write-Host "Installer: $setupExe"
