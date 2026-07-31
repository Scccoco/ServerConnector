# Идёмпотентная настройка rclone remote 'yandex_disk' для Yandex.Disk.
#
# Запускается на VPS один раз после получения OAuth-токена (см. doc/YANDEX_BACKUP_RU.md).
# Токен передаётся обязательным параметром -Token (строка JSON, в одинарных кавычках).
#
# Создаёт C:\Connector\runtime\rclone.conf — конфиг хранится в runtime\
# (не в src\), потому что содержит секрет.

param(
    [Parameter(Mandatory=$true)][string]$Token,
    [string]$CryptPassword = '',
    [string]$CryptSalt = ''
)

$ErrorActionPreference = 'Stop'

$rclone       = 'C:\Tools\rclone\rclone.exe'
$rcloneConfig = 'C:\Connector\runtime\rclone.conf'
$remoteName   = 'yandex_disk'
$cryptRemoteName = 'yandex_crypt'
$cryptRoot = 'yandex_disk:Structura/BIM_Backup_Encrypted'

if (-not (Test-Path $rclone)) { throw "rclone not installed at $rclone" }

# Validate token is JSON-ish (must contain access_token key)
if ($Token -notmatch '"access_token"') {
    throw "Token does not look like rclone OAuth JSON (no 'access_token' field)"
}

# rclone config create yandex_disk yandex token <json>
$args = @(
    'config', 'create', $remoteName, 'yandex',
    'token', $Token,
    '--config', $rcloneConfig,
    '--non-interactive'
)
& $rclone @args
if ($LASTEXITCODE -ne 0) { throw "rclone config create failed: $LASTEXITCODE" }

$configuredRemotes = & $rclone listremotes --config $rcloneConfig
if ($LASTEXITCODE -ne 0) { throw "rclone listremotes failed: $LASTEXITCODE" }

if ($configuredRemotes -notcontains "${cryptRemoteName}:") {
    if ([string]::IsNullOrWhiteSpace($CryptPassword)) {
        $bytes = New-Object byte[] 32
        [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
        $CryptPassword = [Convert]::ToBase64String($bytes)
    }
    if ([string]::IsNullOrWhiteSpace($CryptSalt)) {
        $bytes = New-Object byte[] 32
        [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
        $CryptSalt = [Convert]::ToBase64String($bytes)
    }

    & $rclone config create $cryptRemoteName crypt `
        remote $cryptRoot `
        password $CryptPassword `
        password2 $CryptSalt `
        filename_encryption standard `
        directory_name_encryption true `
        --config $rcloneConfig `
        --non-interactive | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "rclone crypt config create failed: $LASTEXITCODE" }
} else {
    Write-Host "Encrypted remote ${cryptRemoteName}: already exists; encryption keys were preserved."
}

Write-Host ''
Write-Host '=== Test connection ==='
& $rclone --config $rcloneConfig lsd "${remoteName}:" --max-depth 1
if ($LASTEXITCODE -ne 0) { throw "rclone lsd failed: $LASTEXITCODE" }
& $rclone --config $rcloneConfig mkdir "${cryptRemoteName}:"
if ($LASTEXITCODE -ne 0) { throw "rclone crypt mkdir failed: $LASTEXITCODE" }
& $rclone --config $rcloneConfig lsd "${cryptRemoteName}:" --max-depth 1
if ($LASTEXITCODE -ne 0) { throw "rclone crypt lsd failed: $LASTEXITCODE" }

Write-Host ''
Write-Host "RCLONE_CONFIG_OK at $rcloneConfig"
Write-Warning "Back up the complete rclone.conf securely. Without its crypt password/password2 values, encrypted backups cannot be restored."
