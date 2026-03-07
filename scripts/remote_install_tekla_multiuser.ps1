$ErrorActionPreference = 'Continue'

$installer = 'C:\Users\Administrator\Downloads\TeklaStructuresMultiuserServer250.exe'
if (-not (Test-Path $installer)) {
    throw "Installer not found: $installer"
}

$argSets = @(
    '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-',
    '/quiet /norestart',
    '/S',
    '/s /v"/qn /norestart"'
)

$installed = $false
foreach ($args in $argSets) {
    Write-Output "TRY_ARGS=$args"
    try {
        $p = Start-Process -FilePath $installer -ArgumentList $args -PassThru -Wait -WindowStyle Hidden
        Write-Output "EXIT_CODE=$($p.ExitCode)"
    } catch {
        Write-Output "START_FAILED=$($_.Exception.Message)"
    }

    Start-Sleep -Seconds 3

    $svc = Get-Service | Where-Object { $_.DisplayName -match 'Tekla|Multiuser|Multi-user' -or $_.Name -match 'tekla|multi' }
    if ($svc) {
        $installed = $true
        break
    }
}

$svc = Get-Service | Where-Object { $_.DisplayName -match 'Tekla|Multiuser|Multi-user' -or $_.Name -match 'tekla|multi' }

if ($svc) {
    foreach ($s in $svc) {
        if ($s.Status -ne 'Running') {
            try { Start-Service -Name $s.Name } catch {}
        }
        try { Set-Service -Name $s.Name -StartupType Automatic } catch {}
    }
}

Write-Output '=== SERVICES ==='
$svc | Select-Object Name, DisplayName, Status, StartType

Write-Output '=== LISTENER_1238 ==='
Get-NetTCPConnection -LocalPort 1238 -State Listen -ErrorAction SilentlyContinue | Select-Object LocalAddress, LocalPort, OwningProcess

Write-Output "INSTALLED_FLAG=$installed"
