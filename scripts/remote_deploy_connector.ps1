param(
    [Parameter(Mandatory = $true)]
    [string]$AdminApiKey,
    [Parameter(Mandatory = $false)]
    [string]$AllowedIpsCsv = "185.13.177.230,91.219.23.141"
)

$ErrorActionPreference = 'Stop'

$root = 'C:\Connector'
$serverDir = Join-Path $root 'server'
$sourceServerDir = 'C:\Users\opwork_admin\connector\server'

New-Item -Path $root -ItemType Directory -Force | Out-Null
if (Test-Path $serverDir) {
    Remove-Item -Path $serverDir -Recurse -Force
}
Copy-Item -Path $sourceServerDir -Destination $serverDir -Recurse -Force

# Create config.json (do not store plaintext device tokens in repo)
$cfg = @{
    api_bind = '0.0.0.0'
    api_port = 8080
    admin_api_key = $AdminApiKey
    managed_ports = @(3389, 1238, 445)
    allowed_stale_minutes = 15
    devices = @{}
}
$cfg | ConvertTo-Json -Depth 6 | Set-Content -Path (Join-Path $serverDir 'config.json') -Encoding UTF8

# Prepare Python runtime
$py = $null
try { $py = (Get-Command py -ErrorAction Stop).Source } catch {}
if (-not $py) {
    try { $py = (Get-Command python -ErrorAction Stop).Source } catch {}
}
if (-not $py) {
    throw 'Python is not installed on the server. Install Python 3.11+ and rerun.'
}

$venvDir = Join-Path $serverDir '.venv'
if (-not (Test-Path $venvDir)) {
    & $py -3 -m venv $venvDir
}

$venvPython = Join-Path $venvDir 'Scripts\python.exe'
& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r (Join-Path $serverDir 'requirements.txt')

# Firewall for admin UI/API 8080
$allowed = ($AllowedIpsCsv.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ','
$ruleName = 'Connector API 8080'
if (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue) {
    Set-NetFirewallRule -DisplayName $ruleName -RemoteAddress $allowed | Out-Null
} else {
    New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort 8080 -RemoteAddress $allowed -Profile Any | Out-Null
}

# Register auto-start task
$taskName = 'ConnectorApi'
$taskCmd = "cmd /c cd /d $serverDir && `"$venvPython`" -m uvicorn app:app --host 0.0.0.0 --port 8080"

schtasks /Delete /TN $taskName /F | Out-Null 2>$null
schtasks /Create /SC ONSTART /TN $taskName /TR $taskCmd /RU SYSTEM /RL HIGHEST /F | Out-Null
schtasks /Run /TN $taskName | Out-Null

Start-Sleep -Seconds 3

Write-Output "DEPLOY_ROOT=$serverDir"
Get-NetFirewallRule -DisplayName $ruleName | Get-NetFirewallAddressFilter | Select-Object InstanceID, RemoteAddress
Get-NetTCPConnection -LocalPort 8080 -State Listen -ErrorAction SilentlyContinue | Select-Object LocalAddress, LocalPort, OwningProcess
