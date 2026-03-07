param(
    [Parameter(Mandatory = $true)]
    [string]$AllowedIp
)

$ErrorActionPreference = 'Stop'

function Set-RuleRemoteAddress {
    param([Parameter(Mandatory = $true)]$Rules, [string]$Ip)
    foreach ($rule in $Rules) {
        Set-NetFirewallRule -Name $rule.Name -RemoteAddress $Ip | Out-Null
    }
}

# SSH custom rule
$sshRule = Get-NetFirewallRule -Name 'sshd-22' -ErrorAction SilentlyContinue
if ($sshRule) { Set-RuleRemoteAddress -Rules @($sshRule) -Ip $AllowedIp }

# Tekla custom rule
$teklaRule = Get-NetFirewallRule -DisplayName 'Tekla MultiUser 1238 TCP' -ErrorAction SilentlyContinue
if ($teklaRule) { Set-RuleRemoteAddress -Rules @($teklaRule) -Ip $AllowedIp }

# RDP rules (enabled inbound rules only)
$rdpRules = Get-NetFirewallRule -DisplayGroup 'Remote Desktop' |
    Where-Object { $_.Enabled -eq 'True' -and $_.Direction -eq 'Inbound' }
if ($rdpRules) { Set-RuleRemoteAddress -Rules $rdpRules -Ip $AllowedIp }

# SMB rules on port 445 only (enabled inbound rules)
$smbRules = Get-NetFirewallRule -DisplayGroup 'File and Printer Sharing' |
    Where-Object { $_.Enabled -eq 'True' -and $_.Direction -eq 'Inbound' } |
    Where-Object {
        $pf = Get-NetFirewallPortFilter -AssociatedNetFirewallRule $_ -ErrorAction SilentlyContinue
        $pf -and $pf.LocalPort -eq '445' -and $pf.Protocol -eq 'TCP'
    }
if ($smbRules) { Set-RuleRemoteAddress -Rules $smbRules -Ip $AllowedIp }

Write-Output "ALLOWED_IP=$AllowedIp"
Write-Output '--- SSH ---'
Get-NetFirewallRule -Name 'sshd-22' -ErrorAction SilentlyContinue | Get-NetFirewallAddressFilter | Select-Object InstanceID, RemoteAddress
Write-Output '--- TEKLA ---'
Get-NetFirewallRule -DisplayName 'Tekla MultiUser 1238 TCP' -ErrorAction SilentlyContinue | Get-NetFirewallAddressFilter | Select-Object InstanceID, RemoteAddress
Write-Output '--- RDP ---'
Get-NetFirewallRule -DisplayGroup 'Remote Desktop' | Where-Object { $_.Enabled -eq 'True' -and $_.Direction -eq 'Inbound' } | Get-NetFirewallAddressFilter | Select-Object InstanceID, RemoteAddress
Write-Output '--- SMB 445 ---'
Get-NetFirewallRule -DisplayGroup 'File and Printer Sharing' |
    Where-Object { $_.Enabled -eq 'True' -and $_.Direction -eq 'Inbound' } |
    Where-Object {
        $pf = Get-NetFirewallPortFilter -AssociatedNetFirewallRule $_ -ErrorAction SilentlyContinue
        $pf -and $pf.LocalPort -eq '445' -and $pf.Protocol -eq 'TCP'
    } |
    Get-NetFirewallAddressFilter |
    Select-Object InstanceID, RemoteAddress
