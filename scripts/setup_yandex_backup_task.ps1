# Создаёт/обновляет Scheduled Task ConnectorYandexBackup.
#
# Запускается один раз вручную на VPS после того как настроен rclone-config
# (OAuth) и проверен initial backup. Идемпотентен — можно перезапускать.
#
# Расписание: ежедневно в 03:00 MSK (Windows time = UTC+3 на VPS, проверить
# что timezone правильный).

$ErrorActionPreference = 'Stop'

$taskName = 'ConnectorYandexBackup'
$script   = 'C:\Connector\src\scripts\backup_bim_to_yandex.ps1'

if (-not (Test-Path $script)) { throw "Backup script not found: $script" }

$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script`""
$trigger = New-ScheduledTaskTrigger -Daily -At '03:00'
$principal = New-ScheduledTaskPrincipal `
    -UserId 'SYSTEM' `
    -LogonType ServiceAccount `
    -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 15) `
    -ExecutionTimeLimit (New-TimeSpan -Hours 6)

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Force | Out-Null

Write-Host "Task $taskName created/updated:"
Get-ScheduledTask -TaskName $taskName | Select-Object TaskName, State
Get-ScheduledTaskInfo -TaskName $taskName | Select-Object LastRunTime, LastTaskResult, NextRunTime
