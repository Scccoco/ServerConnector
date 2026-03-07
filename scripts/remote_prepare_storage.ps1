$ErrorActionPreference = 'Stop'

$modelRoot = if (Test-Path 'D:\') { 'D:\BIM_Models' } else { 'C:\BIM_Models' }

New-Item -Path $modelRoot -ItemType Directory -Force | Out-Null
New-Item -Path (Join-Path $modelRoot 'Tekla') -ItemType Directory -Force | Out-Null
New-Item -Path (Join-Path $modelRoot 'Revit') -ItemType Directory -Force | Out-Null

if (-not (Get-SmbShare -Name 'BIM_Models' -ErrorAction SilentlyContinue)) {
    New-SmbShare -Name 'BIM_Models' -Path $modelRoot -FullAccess 'Administrators' -ChangeAccess 'opwork_admin' | Out-Null
}

icacls $modelRoot /grant 'Administrators:(OI)(CI)F' 'opwork_admin:(OI)(CI)M' /T | Out-Null

if (-not (Get-NetFirewallRule -DisplayName 'Tekla MultiUser 1238 TCP' -ErrorAction SilentlyContinue)) {
    New-NetFirewallRule -DisplayName 'Tekla MultiUser 1238 TCP' -Direction Inbound -Action Allow -Protocol TCP -LocalPort 1238 | Out-Null
}

Write-Output "MODEL_ROOT=$modelRoot"
Get-SmbShare -Name 'BIM_Models' | Select-Object Name, Path
Get-NetFirewallRule -DisplayName 'Tekla MultiUser 1238 TCP' | Select-Object DisplayName, Enabled, Direction, Action
