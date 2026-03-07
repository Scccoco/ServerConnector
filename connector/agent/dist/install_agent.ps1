param(
    [Parameter(Mandatory = $false)]
    [string]$ServerUrl = 'http://62.113.36.107:8080',
    [Parameter(Mandatory = $false)]
    [string]$DeviceId = ('pc-' + $env:COMPUTERNAME.ToLower()),
    [Parameter(Mandatory = $false)]
    [string]$DeviceToken,
    [Parameter(Mandatory = $false)]
    [int]$HeartbeatSeconds = 60,
    [Parameter(Mandatory = $false)]
    [switch]$InstallPythonIfMissing
)

$ErrorActionPreference = 'Stop'

$installDir = Join-Path $env:LOCALAPPDATA 'ConnectorAgent'
$runKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$runValueName = 'ConnectorAgent'

function Resolve-PythonCommand {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        return 'py -3'
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        return 'python'
    }
    return $null
}

$pyCmd = Resolve-PythonCommand
if (-not $pyCmd -and $InstallPythonIfMissing) {
    $tmp = Join-Path $env:TEMP 'python-3.11.9-amd64.exe'
    Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe' -OutFile $tmp
    Start-Process -FilePath $tmp -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1 Include_test=0 SimpleInstall=1' -Wait
    $pyCmd = Resolve-PythonCommand
}

if (-not $pyCmd) {
    throw 'Python not found. Install Python 3.11+ first or use -InstallPythonIfMissing.'
}

New-Item -ItemType Directory -Path $installDir -Force | Out-Null
Copy-Item -Path "$PSScriptRoot\agent.py" -Destination "$installDir\agent.py" -Force

if (-not $DeviceToken) {
    $DeviceToken = Read-Host 'Enter device token'
}

$secure = ConvertTo-SecureString $DeviceToken -AsPlainText -Force
$secure | ConvertFrom-SecureString | Set-Content -Path "$installDir\device_token.sec" -Encoding ASCII

$cfg = @{
    server_url = $ServerUrl
    device_id = $DeviceId
    token_encrypted_path = "$installDir/device_token.sec"
    heartbeat_seconds = $HeartbeatSeconds
}
$cfg | ConvertTo-Json -Depth 4 | Set-Content -Path "$installDir\agent.json" -Encoding UTF8

$pyExe = (Get-Command py -ErrorAction SilentlyContinue).Source
if (-not $pyExe) {
    $pyExe = (Get-Command python -ErrorAction SilentlyContinue).Source
}
if (-not $pyExe) {
    throw 'Python executable path not found.'
}

$runCmdPath = Join-Path $installDir 'run_agent.cmd'
$runCmdBody = "@echo off`r`ncd /d `"$installDir`"`r`n`"$pyExe`" -3 `"$installDir\agent.py`"`r`n"
Set-Content -Path $runCmdPath -Value $runCmdBody -Encoding ASCII

New-Item -Path $runKeyPath -Force | Out-Null
Set-ItemProperty -Path $runKeyPath -Name $runValueName -Value ('"' + $runCmdPath + '"')
Start-Process -WindowStyle Hidden -FilePath 'cmd.exe' -ArgumentList '/c', $runCmdPath

$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'Connector Agent.lnk'
$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = 'explorer.exe'
$shortcut.Arguments = $installDir
$shortcut.WorkingDirectory = $installDir
$shortcut.Description = 'Open Connector Agent folder'
$shortcut.Save()

Write-Output "Installed: $installDir"
Write-Output "Autostart: HKCU Run ($runValueName)"
Write-Output "Shortcut: $shortcutPath"
