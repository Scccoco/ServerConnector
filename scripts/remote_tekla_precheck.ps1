$ErrorActionPreference = 'Stop'

Write-Output '=== SYSTEM ==='
Get-ComputerInfo | Select-Object WindowsProductName, WindowsVersion, OsHardwareAbstractionLayer

Write-Output '=== STORAGE ==='
$modelRoot = if (Test-Path 'D:\BIM_Models') { 'D:\BIM_Models' } elseif (Test-Path 'C:\BIM_Models') { 'C:\BIM_Models' } else { '<missing>' }
Write-Output "MODEL_ROOT=$modelRoot"
if ($modelRoot -ne '<missing>') {
    Get-ChildItem -Path $modelRoot -Force | Select-Object Name, FullName
}

Write-Output '=== SMB SHARE ==='
Get-SmbShare -Name 'BIM_Models' -ErrorAction SilentlyContinue | Select-Object Name, Path, CurrentUsers

Write-Output '=== FIREWALL 1238 ==='
Get-NetFirewallRule -DisplayName 'Tekla MultiUser 1238 TCP' -ErrorAction SilentlyContinue | Select-Object DisplayName, Enabled, Direction, Action

Write-Output '=== LISTENER 1238 ==='
Get-NetTCPConnection -LocalPort 1238 -State Listen -ErrorAction SilentlyContinue | Select-Object LocalAddress, LocalPort, OwningProcess

Write-Output '=== TEKLA SERVICES ==='
Get-Service | Where-Object { $_.DisplayName -match 'Tekla|Multiuser|Multi-user' -or $_.Name -match 'tekla|multi' } | Select-Object Name, DisplayName, Status, StartType
