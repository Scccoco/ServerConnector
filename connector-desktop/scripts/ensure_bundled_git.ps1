param(
    [string]$TargetDir = "",
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($TargetDir)) {
    throw "TargetDir is required"
}

$targetBin = Join-Path $TargetDir 'bin/git.exe'
$targetCmd = Join-Path $TargetDir 'cmd/git.exe'
$markerPath = Join-Path $TargetDir '.structura-bundle-version'
$expectedMarker = 'MinGit-2.55.0.3-64-bit.zip sha256:f48e2d2dc74a24454adc6d8fd0ac25bf9c2386f19cfb06202b9465aaad4f9f05'
if (-not $Force -and
    ((Test-Path $targetBin) -or (Test-Path $targetCmd)) -and
    (Test-Path $markerPath) -and
    ((Get-Content -LiteralPath $markerPath -Raw).Trim() -eq $expectedMarker)) {
    Write-Host "Bundled git already exists at $TargetDir"
    exit 0
}

if (Test-Path $TargetDir) {
    Remove-Item -Path $TargetDir -Recurse -Force
}
New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null

$assetName = 'MinGit-2.55.0.3-64-bit.zip'
$assetUrl = "https://github.com/git-for-windows/git/releases/download/v2.55.0.windows.3/$assetName"
$expectedSha256 = 'f48e2d2dc74a24454adc6d8fd0ac25bf9c2386f19cfb06202b9465aaad4f9f05'

$zipPath = Join-Path ([System.IO.Path]::GetTempPath()) ("mingit_" + [Guid]::NewGuid().ToString('N') + ".zip")
Write-Host "Downloading pinned $assetName"
Invoke-WebRequest -Uri $assetUrl -OutFile $zipPath -TimeoutSec 120

try {
    $actualSha256 = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actualSha256 -ne $expectedSha256) {
        throw "MinGit SHA-256 mismatch: expected $expectedSha256, got $actualSha256"
    }
    Expand-Archive -Path $zipPath -DestinationPath $TargetDir -Force
}
finally {
    if (Test-Path $zipPath) {
        Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
    }
}

if (-not ((Test-Path $targetBin) -or (Test-Path $targetCmd))) {
    throw "Bundled git extracted, but git.exe not found in bin/cmd"
}

$expectedMarker | Set-Content -LiteralPath $markerPath -Encoding ASCII
Write-Host "Bundled git prepared at $TargetDir"
