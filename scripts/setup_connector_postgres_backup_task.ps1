# Creates or updates the daily PostgreSQL backup task.
# It runs before ConnectorYandexBackup so the fresh verified dump is included
# in the encrypted off-site sync.

$ErrorActionPreference = 'Stop'

$taskName = 'ConnectorPostgresBackup'
$script = 'C:\Connector\src\scripts\backup_connector_postgres.ps1'

if (-not (Test-Path -LiteralPath $script)) {
    throw "PostgreSQL backup script not found: $script"
}

$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -Reason scheduled"
$trigger = New-ScheduledTaskTrigger -Daily -At '02:30'
$principal = New-ScheduledTaskPrincipal `
    -UserId 'SYSTEM' `
    -LogonType ServiceAccount `
    -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 10) `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Force | Out-Null

Write-Host "Task $taskName created/updated."
Get-ScheduledTask -TaskName $taskName | Select-Object TaskName, State
Get-ScheduledTaskInfo -TaskName $taskName | Select-Object LastRunTime, LastTaskResult, NextRunTime
