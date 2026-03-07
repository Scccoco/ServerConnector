$ErrorActionPreference = 'Stop'

$toolsDir = 'C:\Tools\nssm'
$zipPath = 'C:\Users\opwork_admin\Downloads\nssm-2.24.zip'
$url = 'https://nssm.cc/release/nssm-2.24.zip'

New-Item -Path $toolsDir -ItemType Directory -Force | Out-Null
Invoke-WebRequest -Uri $url -OutFile $zipPath
Expand-Archive -Path $zipPath -DestinationPath 'C:\Tools' -Force

$nssm = 'C:\Tools\nssm-2.24\win64\nssm.exe'
if (-not (Test-Path $nssm)) {
    throw "nssm not found at $nssm"
}

$py = 'C:\Connector\server\.venv\Scripts\python.exe'
$args = '-m uvicorn app:app --host 0.0.0.0 --port 8080'

& $nssm stop ConnectorApi | Out-Null 2>$null
& $nssm remove ConnectorApi confirm | Out-Null 2>$null

& $nssm install ConnectorApi $py $args
& $nssm set ConnectorApi AppDirectory 'C:\Connector\server'
& $nssm set ConnectorApi Start SERVICE_AUTO_START
& $nssm set ConnectorApi AppStdout 'C:\Connector\server\nssm.out.log'
& $nssm set ConnectorApi AppStderr 'C:\Connector\server\nssm.err.log'
& $nssm start ConnectorApi

Start-Sleep -Seconds 2

Get-Service -Name ConnectorApi | Select-Object Name, Status, StartType
Get-NetTCPConnection -LocalPort 8080 -State Listen -ErrorAction SilentlyContinue | Select-Object LocalAddress, LocalPort, OwningProcess
