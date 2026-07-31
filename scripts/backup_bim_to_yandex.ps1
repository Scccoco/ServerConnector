# Daily encrypted off-site backup to Yandex.Disk via rclone.
#
# The destination is an rclone crypt remote. Encryption is required not only
# for confidentiality: Yandex content filtering rejects several legitimate
# BIM dependency DLLs with HTTP 409 even under unique names. Encrypting the
# payload makes the backup byte-complete while rclone still restores the
# original directory structure and file contents.
#
# Sources:
#   C:\BIM_Models                     -> yandex_crypt:bim/current
#   C:\Connector\backup\postgres      -> yandex_crypt:connector-db/current
#
# Replaced/deleted objects are moved to archive/<yyyy-MM-dd>/ and retained for
# 90 days. The Scheduled Task ConnectorPostgresBackup runs at 02:30, before
# this task starts at 03:00.

param(
    [string]$RemoteName = 'yandex_crypt',
    [ValidateRange(1, 64)][int]$Transfers = 16,
    [ValidateRange(1, 100)][int]$TpsLimit = 20
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$rclone        = 'C:\Tools\rclone\rclone.exe'
$rcloneConfig  = 'C:\Connector\runtime\rclone.conf'
$bimSource     = 'C:\BIM_Models'
$postgresSource = 'C:\Connector\backup\postgres'
$today         = Get-Date -Format 'yyyy-MM-dd'
$logFile       = 'C:\Connector\runtime\logs\backup_yandex.log'
$statusFile    = 'C:\Connector\runtime\last_yandex_backup.json'
$logDir        = Split-Path $logFile -Parent
$retentionDays = 90

function Write-BackupLog([string]$Message) {
    $Message | Add-Content -LiteralPath $logFile -Encoding UTF8
}

function Invoke-EncryptedSync(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$DestinationBase,
    [string[]]$ExtraArgs = @()
) {
    if (-not (Test-Path -LiteralPath $Source)) {
        throw "Backup source not found: $Source"
    }

    $args = @(
        'sync', $Source, "${RemoteName}:${DestinationBase}/current",
        '--config', $rcloneConfig,
        '--backup-dir', "${RemoteName}:${DestinationBase}/archive/$today",
        '--transfers', $Transfers,
        '--checkers', ([Math]::Max(16, $Transfers * 2)),
        '--tpslimit', $TpsLimit,
        '--tpslimit-burst', $TpsLimit,
        '--contimeout', '30s',
        '--timeout', '10m',
        '--retries', '5',
        '--retries-sleep', '15s',
        '--low-level-retries', '20',
        '--log-file', $logFile,
        '--log-level', 'INFO',
        '--stats', '1m'
    ) + $ExtraArgs

    Write-BackupLog "=== SYNC START source=$Source destination=${RemoteName}:${DestinationBase}/current ==="
    & $rclone @args
    $exitCode = $LASTEXITCODE
    Write-BackupLog "=== SYNC EXIT $exitCode source=$Source ==="
    if ($exitCode -ne 0) {
        throw "rclone sync failed for $Source with exit $exitCode"
    }
}

function Invoke-RemoteRetention([Parameter(Mandatory=$true)][string]$DestinationBase) {
    $args = @(
        'delete', "${RemoteName}:${DestinationBase}/archive",
        '--config', $rcloneConfig,
        '--min-age', "${retentionDays}d",
        '--rmdirs',
        '--tpslimit', $TpsLimit,
        '--tpslimit-burst', $TpsLimit,
        '--retries', '3',
        '--low-level-retries', '10',
        '--log-file', $logFile,
        '--log-level', 'INFO'
    )
    & $rclone @args
    $exitCode = $LASTEXITCODE
    Write-BackupLog "=== RETENTION EXIT $exitCode destination=$DestinationBase ==="
    return $exitCode
}

if (-not (Test-Path -LiteralPath $rclone)) {
    throw "rclone not installed at $rclone"
}
if (-not (Test-Path -LiteralPath $rcloneConfig)) {
    throw "rclone config not found at $rcloneConfig"
}
if (-not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

$configuredRemotes = & $rclone listremotes --config $rcloneConfig
if ($LASTEXITCODE -ne 0 -or $configuredRemotes -notcontains "${RemoteName}:") {
    throw "Encrypted rclone remote '${RemoteName}:' is not configured"
}

if ((Test-Path -LiteralPath $logFile) -and ((Get-Item -LiteralPath $logFile).Length -gt 50MB)) {
    Move-Item -LiteralPath $logFile -Destination "$logFile.$(Get-Date -Format 'yyyyMMdd-HHmmss')" -Force
}

$startedAt = Get-Date
Write-BackupLog "=== BACKUP START $($startedAt.ToString('s')) remote=$RemoteName ==="

try {
    # The database dump is small and operationally critical. Upload it first
    # so an unrelated large BIM-file failure cannot postpone its off-site copy.
    Invoke-EncryptedSync -Source $postgresSource -DestinationBase 'connector-db'

    Invoke-EncryptedSync -Source $bimSource -DestinationBase 'bim' -ExtraArgs @(
        '--exclude', '*.tmp',
        '--exclude', 'Thumbs.db',
        '--exclude', '~$*',
        '--exclude', 'desktop.ini'
    )

    # Retention failures are recorded but do not invalidate a successful new
    # backup. Keeping old versions longer is safer than marking fresh data bad.
    $bimRetentionExit = Invoke-RemoteRetention -DestinationBase 'bim'
    $dbRetentionExit = Invoke-RemoteRetention -DestinationBase 'connector-db'

    $finishedAt = Get-Date
    $status = [ordered]@{
        ok = $true
        started_at = $startedAt.ToUniversalTime().ToString('o')
        finished_at = $finishedAt.ToUniversalTime().ToString('o')
        duration_minutes = [math]::Round((New-TimeSpan -Start $startedAt -End $finishedAt).TotalMinutes, 1)
        remote = $RemoteName
        retention_ok = ($bimRetentionExit -eq 0 -and $dbRetentionExit -eq 0)
    }
    $status | ConvertTo-Json | Set-Content -LiteralPath $statusFile -Encoding UTF8
    Write-BackupLog "=== BACKUP DONE total=$($status.duration_minutes) min retention_ok=$($status.retention_ok) ==="
    exit 0
}
catch {
    $finishedAt = Get-Date
    $message = $_.Exception.Message
    $status = [ordered]@{
        ok = $false
        started_at = $startedAt.ToUniversalTime().ToString('o')
        finished_at = $finishedAt.ToUniversalTime().ToString('o')
        duration_minutes = [math]::Round((New-TimeSpan -Start $startedAt -End $finishedAt).TotalMinutes, 1)
        remote = $RemoteName
        error = $message
    }
    $status | ConvertTo-Json | Set-Content -LiteralPath $statusFile -Encoding UTF8
    Write-BackupLog "=== BACKUP FAILED: $message ==="
    Write-Error $message
    exit 1
}
