$ErrorActionPreference = 'Stop'

$url = 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe'
$dst = 'C:\Users\opwork_admin\Downloads\python-3.11.9-amd64.exe'

Invoke-WebRequest -Uri $url -OutFile $dst
Start-Process -FilePath $dst -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1 Include_test=0 SimpleInstall=1' -Wait

Write-Output 'PYTHON_INSTALL_DONE'
