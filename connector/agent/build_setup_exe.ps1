$ErrorActionPreference = 'Stop'

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$distDir = Join-Path $baseDir 'dist'
$outExe = Join-Path $baseDir 'ConnectorAgentSetup.exe'
$sedPath = Join-Path $distDir 'connector_agent.sed'

New-Item -ItemType Directory -Path $distDir -Force | Out-Null

Copy-Item -Path (Join-Path $baseDir 'install_agent.ps1') -Destination (Join-Path $distDir 'install_agent.ps1') -Force
Copy-Item -Path (Join-Path $baseDir 'install_wizard.ps1') -Destination (Join-Path $distDir 'install_wizard.ps1') -Force
Copy-Item -Path (Join-Path $baseDir 'agent.py') -Destination (Join-Path $distDir 'agent.py') -Force

$targetEscaped = $outExe
$sourceEscaped = $distDir + '\\'

$sed = @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=1
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=
TargetName=$targetEscaped
FriendlyName=Connector Agent Setup
AppLaunched=install_connector.cmd
PostInstallCmd=<None>
AdminQuietInstCmd=install_connector.cmd
UserQuietInstCmd=install_connector.cmd
SourceFiles=SourceFiles
[SourceFiles]
SourceFiles0=$sourceEscaped
[SourceFiles0]
%FILE0%=
%FILE1%=
%FILE2%=
%FILE3%=
[Strings]
FILE0=install_connector.cmd
FILE1=install_agent.ps1
FILE2=agent.py
FILE3=install_wizard.ps1
"@

Set-Content -Path $sedPath -Value $sed -Encoding ASCII

Start-Process -FilePath 'iexpress.exe' -ArgumentList '/N', '/Q', $sedPath -Wait

if (-not (Test-Path $outExe)) {
    throw "Build failed: $outExe not found"
}

Write-Output "Built: $outExe"
