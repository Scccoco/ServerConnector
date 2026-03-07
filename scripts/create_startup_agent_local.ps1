$ErrorActionPreference = 'Stop'

$installDir = 'C:\ProgramData\ConnectorAgent'
$startupDir = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
$cmdPath = Join-Path $startupDir 'ConnectorAgent.cmd'
$pyPath = 'C:\Windows\py.exe'
$logPath = 'C:\ProgramData\ConnectorAgent\agent.log'

New-Item -ItemType Directory -Path $startupDir -Force | Out-Null

$body = @"
@echo off
cd /d "$installDir"
"$pyPath" -3 "$installDir\agent.py" >> "$logPath" 2>&1
"@

Set-Content -Path $cmdPath -Value $body -Encoding ASCII
Start-Process -WindowStyle Hidden -FilePath 'cmd.exe' -ArgumentList '/c', $cmdPath

Write-Output "Startup launcher: $cmdPath"
