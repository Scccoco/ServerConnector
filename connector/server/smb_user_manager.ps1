param(
    [Parameter(Mandatory = $true)]
    [string]$UserName,
    [Parameter(Mandatory = $true)]
    [string]$Password,
    [Parameter(Mandatory = $true)]
    [string]$ShareName,
    [Parameter(Mandatory = $true)]
    [string]$SharePath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $SharePath)) {
    $share = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
    if ($share -and (Test-Path $share.Path)) {
        $SharePath = $share.Path
    } else {
        throw "Share path not found: $SharePath"
    }
}

$securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$existingUser = Get-LocalUser -Name $UserName -ErrorAction SilentlyContinue

if ($existingUser) {
    Set-LocalUser -Name $UserName -Password $securePassword
    Set-LocalUser -Name $UserName -PasswordNeverExpires $true
    Enable-LocalUser -Name $UserName
} else {
    New-LocalUser -Name $UserName -Password $securePassword -PasswordNeverExpires -UserMayNotChangePassword -AccountNeverExpires | Out-Null
}

$account = "$env:COMPUTERNAME\$UserName"

Grant-SmbShareAccess -Name $ShareName -AccountName $account -AccessRight Change -Force | Out-Null
& icacls $SharePath /grant "${account}:(OI)(CI)M" /T /C | Out-Null

Write-Output ("ok:{0}:{1}" -f $account, $ShareName)
