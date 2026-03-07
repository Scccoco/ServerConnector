param(
    [Parameter(Mandatory = $true)]
    [string]$DeviceId,
    [Parameter(Mandatory = $true)]
    [string]$RemoteIp,
    [Parameter(Mandatory = $true)]
    [string]$PortsCsv
)

$ErrorActionPreference = 'Stop'
$ports = $PortsCsv.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }

foreach ($port in $ports) {
    $ruleName = "Connector Allow $DeviceId TCP $port"
    $existing = Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue

    if ($existing) {
        Set-NetFirewallRule -DisplayName $ruleName -RemoteAddress $RemoteIp | Out-Null
    } else {
        New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort $port -RemoteAddress $RemoteIp -Profile Any | Out-Null
    }
}

Write-Output ("updated:{0}:{1}:{2}" -f $DeviceId, $RemoteIp, $PortsCsv)
