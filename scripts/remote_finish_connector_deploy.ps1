$ErrorActionPreference = 'Stop'

$serverDir = 'C:\Connector\server'
$venvPython = 'C:\Connector\server\.venv\Scripts\python.exe'

$ruleName = 'Connector API 8080'
if (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue) {
    Set-NetFirewallRule -DisplayName $ruleName -RemoteAddress Any | Out-Null
} else {
    New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort 8080 -RemoteAddress Any -Profile Any | Out-Null
}

$taskName = 'ConnectorApi'
$taskCmd = "cmd /c cd /d $serverDir && `"$venvPython`" -m uvicorn app:app --host 0.0.0.0 --port 8080"

schtasks /Delete /TN $taskName /F | Out-Null 2>$null
schtasks /Create /SC ONSTART /TN $taskName /TR $taskCmd /RU SYSTEM /RL HIGHEST /F | Out-Null
schtasks /Run /TN $taskName | Out-Null

Start-Sleep -Seconds 4

Write-Output 'CONNECTOR_DEPLOY_DONE'
Get-NetFirewallRule -DisplayName $ruleName | Get-NetFirewallAddressFilter | Select-Object InstanceID, RemoteAddress
Get-NetTCPConnection -LocalPort 8080 -State Listen -ErrorAction SilentlyContinue | Select-Object LocalAddress, LocalPort, OwningProcess
