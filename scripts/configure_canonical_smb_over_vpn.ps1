param(
    [string]$PublicInterfaceAlias = 'Ethernet',
    [string]$TunnelInterfaceAlias = 'awgserver',
    [string]$PublicIpAddress = '62.113.36.107',
    [string]$VpnSubnet = '10.77.123.0/24',
    [string]$FirewallRuleName = 'AmneziaWG SMB to VPN subnet'
)

$ErrorActionPreference = 'Stop'

$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this script from an elevated PowerShell session.'
}

$publicInterface = Get-NetIPInterface -InterfaceAlias $PublicInterfaceAlias -AddressFamily IPv4 -ErrorAction Stop
$tunnelInterface = Get-NetIPInterface -InterfaceAlias $TunnelInterfaceAlias -AddressFamily IPv4 -ErrorAction Stop

$publicAddress = Get-NetIPAddress `
    -InterfaceIndex $publicInterface.InterfaceIndex `
    -AddressFamily IPv4 `
    -IPAddress $PublicIpAddress `
    -ErrorAction SilentlyContinue

if (-not $publicAddress) {
    throw "IPv4 address $PublicIpAddress is not assigned to interface $PublicInterfaceAlias."
}

# A client routes the server's public /32 into AmneziaWG so existing UNC paths
# stay unchanged. Windows must therefore accept a packet for an address owned
# by Ethernet when it arrives on awgserver, and send the response back through
# awgserver even though the response source address belongs to Ethernet.
foreach ($interfaceAlias in @($PublicInterfaceAlias, $TunnelInterfaceAlias)) {
    Set-NetIPInterface `
        -InterfaceAlias $interfaceAlias `
        -AddressFamily IPv4 `
        -WeakHostSend Enabled `
        -WeakHostReceive Enabled
}

$firewallRule = Get-NetFirewallRule -DisplayName $FirewallRuleName -ErrorAction SilentlyContinue
if ($firewallRule) {
    $firewallRule |
        Set-NetFirewallRule -Enabled True -Direction Inbound -Action Allow -Profile Any |
        Out-Null
    $firewallRule |
        Get-NetFirewallPortFilter |
        Set-NetFirewallPortFilter -Protocol TCP -LocalPort 445 |
        Out-Null
    $firewallRule |
        Get-NetFirewallAddressFilter |
        Set-NetFirewallAddressFilter -RemoteAddress $VpnSubnet |
        Out-Null
} else {
    New-NetFirewallRule `
        -DisplayName $FirewallRuleName `
        -Direction Inbound `
        -Action Allow `
        -Protocol TCP `
        -LocalPort 445 `
        -RemoteAddress $VpnSubnet `
        -Profile Any |
        Out-Null
}

$activeInterfaces = @(
    Get-NetIPInterface -PolicyStore ActiveStore -AddressFamily IPv4 |
        Where-Object InterfaceAlias -in @($PublicInterfaceAlias, $TunnelInterfaceAlias)
)
$persistentInterfaces = @(
    Get-NetIPInterface -PolicyStore PersistentStore -AddressFamily IPv4 |
        Where-Object InterfaceAlias -in @($PublicInterfaceAlias, $TunnelInterfaceAlias)
)

foreach ($interfaceAlias in @($PublicInterfaceAlias, $TunnelInterfaceAlias)) {
    $active = $activeInterfaces | Where-Object InterfaceAlias -eq $interfaceAlias
    $persistent = $persistentInterfaces | Where-Object InterfaceAlias -eq $interfaceAlias
    if ($active.WeakHostSend -ne 'Enabled' -or
        $active.WeakHostReceive -ne 'Enabled' -or
        $persistent.WeakHostSend -ne 'Enabled' -or
        $persistent.WeakHostReceive -ne 'Enabled') {
        throw "Weak-host mode was not persisted for interface $interfaceAlias."
    }
}

[pscustomobject]@{
    PublicInterface = $PublicInterfaceAlias
    TunnelInterface = $TunnelInterfaceAlias
    PublicIp = $PublicIpAddress
    VpnSubnet = $VpnSubnet
    WeakHostActive = 'Enabled'
    WeakHostPersistent = 'Enabled'
    FirewallRule = $FirewallRuleName
}
