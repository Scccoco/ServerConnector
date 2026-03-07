$ErrorActionPreference = 'Stop'

$cmd = '"C:\Windows\py.exe" -3 "C:\ProgramData\ConnectorAgent\agent.py"'

reg add 'HKCU\Software\Microsoft\Windows\CurrentVersion\Run' /v ConnectorAgent /t REG_SZ /d $cmd /f | Out-Null

Start-Process -WindowStyle Hidden -FilePath 'C:\Windows\py.exe' -ArgumentList '-3', 'C:\ProgramData\ConnectorAgent\agent.py' -RedirectStandardOutput 'C:\ProgramData\ConnectorAgent\agent.log' -RedirectStandardError 'C:\ProgramData\ConnectorAgent\agent.err.log'

Write-Output 'autostart-configured'
