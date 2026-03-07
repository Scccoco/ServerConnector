$ErrorActionPreference = 'Stop'

$serverUrl = 'http://62.113.36.107:8080'
$adminKey = 'XXBKHWgBICU28oLIIj0hXFV2orcUSkg3AqS5YpC4dnQ'
$deviceId = 'pc-' + $env:COMPUTERNAME.ToLower()
$issuedTo = 'Lagom | Local PC'

$body = @{
    device_id = $deviceId
    issued_to = $issuedTo
} | ConvertTo-Json

$resp = Invoke-RestMethod -Method Post -Uri ($serverUrl + '/admin/tokens') -Headers @{
    'X-Admin-Key' = $adminKey
    'X-Admin-Actor' = 'OpenWork'
} -ContentType 'application/json' -Body $body

if (-not $resp.token) {
    throw 'Token was not returned by API'
}

& 'E:\00_Cursor\18_Server\connector\agent\install_agent.ps1' -ServerUrl $serverUrl -DeviceId $deviceId -DeviceToken $resp.token -HeartbeatSeconds 60

Start-Sleep -Seconds 3

Write-Output 'Autostart mode: HKCU Run'
Get-Content (Join-Path $env:LOCALAPPDATA 'ConnectorAgent\agent.json')
