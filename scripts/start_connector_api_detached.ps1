$ErrorActionPreference = 'Stop'

$wd = 'C:\Connector\server'
$py = 'C:\Connector\server\.venv\Scripts\python.exe'
$runnerLog = 'C:\Connector\server\connector_runner.log'
$outLog = 'C:\Connector\server\uvicorn.out.log'
$errLog = 'C:\Connector\server\uvicorn.err.log'

"[$(Get-Date -Format s)] runner start" | Add-Content -Path $runnerLog

# Stop existing uvicorn app:app instances to avoid duplicates
Get-CimInstance Win32_Process |
    Where-Object { $_.Name -eq 'python.exe' -and $_.CommandLine -like '*uvicorn app:app*' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

Start-Process -WindowStyle Hidden -WorkingDirectory $wd -FilePath $py -ArgumentList '-m','uvicorn','app:app','--host','0.0.0.0','--port','8080' -RedirectStandardOutput $outLog -RedirectStandardError $errLog

Start-Sleep -Seconds 2
$p = Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'python.exe' -and $_.CommandLine -like '*uvicorn app:app*' } | Select-Object -First 1
if ($p) {
    "[$(Get-Date -Format s)] runner success pid=$($p.ProcessId)" | Add-Content -Path $runnerLog
} else {
    "[$(Get-Date -Format s)] runner failed to start uvicorn" | Add-Content -Path $runnerLog
}
