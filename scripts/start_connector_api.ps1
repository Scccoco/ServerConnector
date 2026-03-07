$ErrorActionPreference = 'Stop'
Set-Location 'C:\Connector\server'
& 'C:\Connector\server\.venv\Scripts\python.exe' -m uvicorn app:app --host 0.0.0.0 --port 8080
