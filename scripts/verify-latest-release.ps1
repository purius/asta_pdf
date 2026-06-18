param(
    [string]$ExpectedVersion,
    [switch]$SkipDownloadHeadCheck,
    [int]$RetryCount = 1,
    [int]$RetryDelaySeconds = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $root 'src\PdfMergeTool\PdfMergeTool.csproj'
$latestReleaseApiUrl = 'https://api.github.com/repos/purius/asta_pdf/releases/latest'
$installerAssetName = 'PdfMergeToolSetup.exe'

function Get-ProjectVersion {
    $project = [xml](Get-Content -Raw $projectPath)
    $version = $project.Project.PropertyGroup.Version
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Missing Version in $projectPath"
    }

    return $version.Trim()
}

function Normalize-VersionTag {
    param([Parameter(Mandatory = $true)][string]$Value)
    return $Value.Trim().TrimStart('v', 'V')
}

function Assert-HeadOk {
    param([Parameter(Mandatory = $true)][string]$Url)

    $previousErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $output = curl.exe -L -I $Url 2>&1
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    $status = ($output | Select-String -Pattern '^HTTP/.* (200|302|403|404)' | Select-Object -Last 1).Line
    if ($status -notmatch ' 200 ') {
        throw "Expected 200 OK for $Url, got: $status"
    }
}

$headers = @{
    'User-Agent' = 'PdfMergeTool-release-verifier'
    'Accept' = 'application/vnd.github+json'
}

if ([string]::IsNullOrWhiteSpace($ExpectedVersion)) {
    $ExpectedVersion = Get-ProjectVersion
}

if ($RetryCount -lt 1) {
    throw 'RetryCount must be at least 1.'
}

if ($RetryDelaySeconds -lt 0) {
    throw 'RetryDelaySeconds must be zero or greater.'
}

$lastError = $null
for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
    try {
        $release = Invoke-RestMethod -Uri $latestReleaseApiUrl -Headers $headers
        $latestTag = [string]$release.tag_name
        $latestVersion = Normalize-VersionTag $latestTag
        $expectedNormalized = Normalize-VersionTag $ExpectedVersion

        if ($latestVersion -ne $expectedNormalized) {
            throw "GitHub latest release is $latestTag, but project version is $ExpectedVersion."
        }

        $installerAsset = @($release.assets) |
            Where-Object { $_.name -ieq $installerAssetName } |
            Select-Object -First 1

        if (-not $installerAsset) {
            throw "Latest release $latestTag does not include $installerAssetName."
        }

        $downloadUrl = [string]$installerAsset.browser_download_url
        if ([string]::IsNullOrWhiteSpace($downloadUrl)) {
            throw "Latest release $latestTag has no browser_download_url for $installerAssetName."
        }

        if (-not $SkipDownloadHeadCheck) {
            Assert-HeadOk $downloadUrl
        }

        Write-Output "latest release verified: $latestTag $downloadUrl"
        exit 0
    }
    catch {
        $lastError = $_
        if ($attempt -eq $RetryCount) {
            throw
        }

        Write-Output "latest release verification attempt $attempt of $RetryCount failed: $($_.Exception.Message)"
        Start-Sleep -Seconds $RetryDelaySeconds
    }
}

throw $lastError
