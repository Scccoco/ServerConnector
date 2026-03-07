$ErrorActionPreference = 'SilentlyContinue'

$runKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$runValueName = 'ConnectorAgent'
$installDirs = @(
    (Join-Path $env:LOCALAPPDATA 'ConnectorAgent'),
    'C:\ProgramData\ConnectorAgent'
)

Remove-ItemProperty -Path $runKeyPath -Name $runValueName -ErrorAction SilentlyContinue

Get-CimInstance Win32_Process |
    Where-Object { $_.Name -eq 'python.exe' -and $_.CommandLine -like '*ConnectorAgent*agent.py*' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

foreach ($installDir in $installDirs) {
    if (Test-Path $installDir) {
        Remove-Item -Path $installDir -Recurse -Force
    }
}

Write-Output 'Connector removed.'
