$ErrorActionPreference = 'Stop'

Write-Output '=== SERVICES ==='
Get-Service -Name LanmanServer, LanmanWorkstation | Select-Object Name, Status, StartType

Write-Output '=== SMB SHARE ==='
Get-SmbShare -Name 'BIM_Models' | Select-Object Name, Path, CurrentUsers

Write-Output '=== FIREWALL SMB RULES ==='
Get-NetFirewallRule -DisplayGroup 'File and Printer Sharing' |
    Where-Object { $_.Direction -eq 'Inbound' } |
    ForEach-Object {
        $pf = Get-NetFirewallPortFilter -AssociatedNetFirewallRule $_ -ErrorAction SilentlyContinue
        $af = Get-NetFirewallAddressFilter -AssociatedNetFirewallRule $_ -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            Name = $_.Name
            DisplayName = $_.DisplayName
            Enabled = $_.Enabled
            Action = $_.Action
            Profile = $_.Profile
            Protocol = $pf.Protocol
            LocalPort = $pf.LocalPort
            RemoteAddress = $af.RemoteAddress
        }
    } |
    Where-Object { $_.Protocol -eq 'TCP' -and $_.LocalPort -eq '445' } |
    Sort-Object Enabled -Descending |
    Format-Table -AutoSize

Write-Output '=== CUSTOM SMB ALLOW RULE ==='
Get-NetFirewallRule -DisplayName 'SMB 445 allow 91.219.23.141' -ErrorAction SilentlyContinue |
    Get-NetFirewallAddressFilter |
    Select-Object InstanceID, RemoteAddress

Write-Output '=== LISTENER 445 ==='
Get-NetTCPConnection -LocalPort 445 -State Listen -ErrorAction SilentlyContinue | Select-Object LocalAddress, LocalPort, OwningProcess
