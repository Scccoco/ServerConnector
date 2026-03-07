$ErrorActionPreference = 'Stop'

$path = 'C:\Connector\server\config.json'
$cfg = Get-Content $path -Raw | ConvertFrom-Json
$cfg.admin_api_key = $cfg.admin_api_key.Trim("'")
$cfg | ConvertTo-Json -Depth 6 | Set-Content $path -Encoding UTF8

Write-Output 'KEY_FIXED'
