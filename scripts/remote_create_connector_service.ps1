$ErrorActionPreference = 'Stop'

$name = 'ConnectorApiSvc'
$display = 'Connector API Service'
$bin = 'cmd /c cd /d C:\Connector\server && C:\Connector\server\.venv\Scripts\python.exe -m uvicorn app:app --host 0.0.0.0 --port 8080'

if (Get-Service -Name $name -ErrorAction SilentlyContinue) {
    sc.exe stop $name | Out-Null
    Start-Sleep -Seconds 1
    sc.exe delete $name | Out-Null
    Start-Sleep -Seconds 1
}

New-Service -Name $name -DisplayName $display -BinaryPathName $bin -StartupType Automatic
Start-Service -Name $name

Get-Service -Name $name | Select-Object Name, Status, StartType
Get-NetTCPConnection -LocalPort 8080 -State Listen -ErrorAction SilentlyContinue | Select-Object LocalAddress, LocalPort, OwningProcess
