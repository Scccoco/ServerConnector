# Consistent PostgreSQL backup for Connector.
#
# Reads CONNECTOR_DB_URL from C:\Connector\runtime\.env, creates a compressed
# custom-format pg_dump, verifies it with pg_restore --list, then atomically
# publishes the .dump file under C:\Connector\backup\postgres.

param(
    [string]$Reason = 'scheduled'
)

$ErrorActionPreference = 'Stop'

$runtimeDir = 'C:\Connector\runtime'
$envFile = Join-Path $runtimeDir '.env'
$pgDump = 'C:\pgsql\pgsql\bin\pg_dump.exe'
$pgRestore = 'C:\pgsql\pgsql\bin\pg_restore.exe'
$backupDir = 'C:\Connector\backup\postgres'
$logDir = Join-Path $runtimeDir 'logs'
$logFile = Join-Path $logDir 'backup_postgres.log'

function Import-RuntimeEnv([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Runtime env file not found: $Path"
    }

    Get-Content -LiteralPath $Path | ForEach-Object {
        $line = $_.Trim()
        if (-not $line -or $line.StartsWith('#')) {
            return
        }
        $eq = $line.IndexOf('=')
        if ($eq -le 0) {
            return
        }
        $key = $line.Substring(0, $eq).Trim()
        $value = $line.Substring($eq + 1).Trim()
        if (($value.StartsWith('"') -and $value.EndsWith('"')) -or
            ($value.StartsWith("'") -and $value.EndsWith("'"))) {
            $value = $value.Substring(1, $value.Length - 2)
        }
        Set-Item -Path "Env:$key" -Value $value
    }
}

if (-not (Test-Path -LiteralPath $pgDump)) {
    throw "pg_dump not found: $pgDump"
}
if (-not (Test-Path -LiteralPath $pgRestore)) {
    throw "pg_restore not found: $pgRestore"
}
if (-not (Test-Path -LiteralPath $backupDir)) {
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
}
if (-not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

Import-RuntimeEnv -Path $envFile
$dbUrl = if ($null -eq $env:CONNECTOR_DB_URL) { '' } else { $env:CONNECTOR_DB_URL.Trim() }
if (-not $dbUrl.StartsWith('postgresql://', [StringComparison]::OrdinalIgnoreCase) -and
    -not $dbUrl.StartsWith('postgres://', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'CONNECTOR_DB_URL is missing or is not PostgreSQL'
}

$uri = [Uri]$dbUrl
$separator = $uri.UserInfo.IndexOf(':')
if ($separator -le 0) {
    throw 'CONNECTOR_DB_URL does not contain username/password'
}
$username = [Uri]::UnescapeDataString($uri.UserInfo.Substring(0, $separator))
$password = [Uri]::UnescapeDataString($uri.UserInfo.Substring($separator + 1))
$database = [Uri]::UnescapeDataString($uri.AbsolutePath.TrimStart('/'))
$port = if ($uri.Port -gt 0) { $uri.Port } else { 5432 }

$reasonValue = if ($null -eq $Reason) { 'backup' } else { $Reason.Trim() }
$safeReason = [regex]::Replace($reasonValue, '[^0-9A-Za-z._-]+', '-').Trim('-')
if (-not $safeReason) {
    $safeReason = 'backup'
}
$timestamp = Get-Date -Format 'yyyy-MM-ddTHH-mm-ss'
$finalPath = Join-Path $backupDir "connector_prod.$safeReason.$timestamp.dump"
$partialPath = "$finalPath.partial"
$startedAt = Get-Date
$previousPgPassword = $env:PGPASSWORD

try {
    $env:PGPASSWORD = $password
    "[$($startedAt.ToString('s'))] pg_dump start reason=$safeReason database=$database" |
        Add-Content -LiteralPath $logFile -Encoding UTF8

    & $pgDump `
        '--host' $uri.Host `
        '--port' $port `
        '--username' $username `
        '--dbname' $database `
        '--format' 'custom' `
        '--compress' '9' `
        '--no-owner' `
        '--no-privileges' `
        '--file' $partialPath
    if ($LASTEXITCODE -ne 0) {
        throw "pg_dump exited with code $LASTEXITCODE"
    }
    if (-not (Test-Path -LiteralPath $partialPath) -or (Get-Item -LiteralPath $partialPath).Length -le 0) {
        throw 'pg_dump produced an empty file'
    }

    & $pgRestore '--list' $partialPath | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "pg_restore verification exited with code $LASTEXITCODE"
    }

    Move-Item -LiteralPath $partialPath -Destination $finalPath -Force
    $size = (Get-Item -LiteralPath $finalPath).Length
    $duration = [math]::Round((New-TimeSpan -Start $startedAt -End (Get-Date)).TotalSeconds, 1)
    "[$((Get-Date).ToString('s'))] pg_dump success file=$finalPath bytes=$size duration_seconds=$duration" |
        Add-Content -LiteralPath $logFile -Encoding UTF8
    Write-Output $finalPath
}
catch {
    Remove-Item -LiteralPath $partialPath -Force -ErrorAction SilentlyContinue
    "[$((Get-Date).ToString('s'))] pg_dump failed: $($_.Exception.Message)" |
        Add-Content -LiteralPath $logFile -Encoding UTF8
    throw
}
finally {
    if ($null -eq $previousPgPassword) {
        Remove-Item Env:PGPASSWORD -ErrorAction SilentlyContinue
    } else {
        $env:PGPASSWORD = $previousPgPassword
    }
}
