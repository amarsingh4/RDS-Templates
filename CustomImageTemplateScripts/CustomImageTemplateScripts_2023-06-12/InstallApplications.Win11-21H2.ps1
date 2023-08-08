$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$aibrepoUrl = "https://nxpwvdprodemeaaibreposa.blob.core.windows.net/aib-repository/"
$aibrepoSas = "?sv=2019-12-12&ss=b&srt=sco&sp=rl&se=2030-12-03T22:26:24Z&st=2020-12-03T14:26:24Z&spr=https&sig=Io6bNagCRnw3WyvfGOJWriTJlQ%2BIXg1BC%2FraTxPggQs%3D"

$avdOptContainerUrl = "https://nxpwvdprodemeawvdoptssa.blob.core.windows.net/wvd-opts-scripts"
$avdOptContainerSasToken = "?sp=r&st=2022-10-28T07:38:56Z&se=2032-10-28T15:38:56Z&spr=https&sv=2021-06-08&sr=c&sig=misyCEdYX0wqc0LGvVFKdF9gaw0dNkLBqo4JFFE5ZNg%3D"

########################## IMPORTANT ##########################
######### Change when creating new image definition!! #########
$avdOptZipFileName = "AVDOptimizer-Win11-21H2.zip"
########################## IMPORTANT ##########################

if ($null -eq (Get-Item -Path "c:\buildArtifacts" -ErrorAction SilentlyContinue)) {
    New-Item -Path "c:\buildArtifacts" -ItemType Directory -Force
}

Set-Location "c:\buildArtifacts"
$logFileLocation = "c:\buildArtifacts\AppsInstallation.log"
$logTranscriptFileLocation = "c:\buildArtifacts\InstallationTranscript.log"

try {
    Stop-Transcript
}
catch {}

Start-Transcript -Path $logTranscriptFileLocation -Append

#elevate script
If (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
    Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
    Exit
}
#
#region Change Sysprep mode
#########################################
$fileLocation = "c:\DeprovisioningScript.ps1"
Write-Host ("Locating file '{0}'" -f $fileLocation)
if (!(Test-Path -Path $fileLocation)) {
    throw ("File not found")
}

Write-Host "Reading file..."
$fileContent = Get-Content -Path $fileLocation -Raw
$fileContent = $fileContent -replace "Sysprep.exe /oobe /generalize /quiet /quit", "Sysprep.exe /oobe /generalize /quiet /quit /mode:vm"

Write-Host "Updating file..."
$fileContent | Set-Content -Path $fileLocation -Force
Write-Host "Done"
#endregion

#region Deactivation of windows firewall
#########################################
Write-Host "Disabling Firewall"
Write-Host (Get-Date).ToString("o")
Set-NetFirewallProfile -Profile Domain, Public, Private -Enabled False
#endregion

#region FSlogix
#########################################
Write-Host "Installing FSlogix"
Write-Host (Get-Date).ToString("o")

$FSlogixurl = "https://aka.ms/fslogix_download"
$FSlogixAppInstaller = "c:\buildArtifacts\FSLogixAppsSetup\x64\Release\FSLogixAppsSetup.exe"
$FSlogixInstallerzip = "c:\buildArtifacts\FSLogixAppsSetup.zip"

("FSlogix full download URL = '{0}'" -f $FSlogixurl) | Out-File $logFileLocation -Append
("FSlogix download location = '{0}'" -f $FSlogixInstallerzip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $FSlogixurl -OutFile $FSlogixInstallerzip -UseBasicParsing

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -LiteralPath $FSlogixInstallerzip -DestinationPath "c:\buildArtifacts\FSLogixAppsSetup\"
("Extraction finished.") | Out-File $logFileLocation -Append

("Starting installer...") | Out-File $logFileLocation -Append
$FSlogix_install_status = Start-Process $FSlogixAppInstaller -ArgumentList @('/install', '/quiet', '/norestart') -wait -PassThru
("Installer finished with returncode '{0}'" -f $FSlogix_install_statuss.ExitCode) | Out-File $logFileLocation -Append

#########################################
#endregion

#region StickyNotes
#########################################
Write-Host "Installing StickyNotes"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}StickyNotes.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\StickyNotes.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

#dependencies

$DependenciesStickyVCLibs = "c:\buildArtifacts\Microsoft.VCLibs.140.00_14.0.30035.0_x64__8wekyb3d8bbwe.appx"
$DependenciesStickyNetFrame = "c:\buildArtifacts\Microsoft.NET.Native.Framework.2.2_2.2.29512.0_x64__8wekyb3d8bbwe.Appx"
$DependenciesStickyNetRun = "c:\buildArtifacts\Microsoft.NET.Native.Runtime.2.2_2.2.28604.0_x64__8wekyb3d8bbwe.Appx"
$NonMSIXAppInstaller = "c:\buildArtifacts\Microsoft.MicrosoftStickyNotes_4.1.4.0_neutral___8wekyb3d8bbwe.Msixbundle"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Add-AppPackage $DependenciesStickyVCLibs
$NonMSIXApp_install_status = Add-AppPackage $DependenciesStickyNetFrame
$NonMSIXApp_install_status = Add-AppPackage $DependenciesStickyNetRun
$NonMSIXApp_install_status = Add-AppPackage $NonMSIXAppInstaller
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append
#endregion

#region Winget Installation
#########################################
Write-Host "Winget Installation"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}WinGet.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\WinGet.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

#########################################
#endregion

#region TeradataODBC
#########################################
Write-Host "Installing TeradataODBC"
Write-Host (Get-Date).ToString("o")

$NonMSIXAppUrl = ("{0}TeradataODBC__windows_indep.17.10.08.00.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\TeradataODBC__windows_indep.17.10.08.00.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

("Starting installer...") | Out-File $logFileLocation -Append
$TeradataODBC_install_status = Start-Process -FilePath "c:\buildArtifacts\TeradataODBC\silent_install.bat" -ArgumentList @('"ODBC" "C:\Program Files"') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $TeradataODBC_install_status.ExitCode) | Out-File $logFileLocation -Append
#endregion

#region OracleODBC
#########################################
Write-Host "Installing OracleODBC"
Write-Host (Get-Date).ToString("o")

$NonMSIXAppUrl = ("{0}instantclient_19_11.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\instantclient_19_11.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\Program Files\"
("Extraction finished.") | Out-File $logFileLocation -Append

("Starting installer...") | Out-File $logFileLocation -Append
Set-Location "C:\Program Files\instantclient_19_11"
$OracleODBC_install_status = Start-Process -FilePath "odbc_install.exe" -Wait -Passthru
("Installer finished with returncode '{0}'" -f $OracleODBC_install_status.ExitCode) | Out-File $logFileLocation -Append
#endregion

#region Office Templates
#########################################
Write-Host "Applying Office Templates"
Write-Host (Get-Date).ToString("o")

$NonMSIXAppUrl = ("{0}NXP Office Templates.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\NXP Office Templates.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "C:\Program Files\"
("Extraction finished.") | Out-File $logFileLocation -Append
#endregion

#region Collabnet-SubversionClient
#########################################
Write-Host "Installing Collabnet-SubversionClient"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}Collabnet_SubversionClientR01_1.8.5_v1.0.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\Collabnet_SubversionClientR01_1.8.5_v1.0.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\Collabnet_SubversionClientR01_1.8.5_v1.0\Package\Resource\CollabNetSubversionClient-185-61-x64-R01-B01.msi"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('ALLUSERS="1"', '/quiet', '/norestart') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append
#endregion

#region Nunit
#########################################
Write-Host "Installling Nunit"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}Nunit-34-EN-61-R01-B01.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\Nunit-34-EN-61-R01-B01.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\Nunit-34-EN-61-R01-B01\Install\NUnit.3.4.0.msi"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('TRANSFORMS="C:\buildArtifacts\Nunit-34-EN-61-R01-B01\Install\Nunit-34-EN-61-R01-B01.mst"', '/quiet', '/norestart') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append
#########################################
#endregion

#region PDFCreator
#########################################
Write-Host "Installing PDFforge PDFCreator"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}PDFforge_PDFCreator_3.4.0_v1.2.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\PDFforge_PDFCreator_3.4.0_v1.2.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\PDFforge_PDFCreator_3.4.0_v1.2\Package\Resource\PDFforge_PDFCreator_3.4.0_v1.0.msi"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/quiet', '/norestart') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append
#########################################
#endregion

#region Segger_Jlink
#########################################
Write-Host "Installing SEGGER JLink"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}SEGGER_JLink_V610n_v1.0.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\SEGGER_JLink_V610n_v1.0.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\SEGGER_JLink_V610n_v1.0\Package\Deploy-Application.exe"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('Install', 'Silent') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#########################################
#endregion

#region NI_LabVIEW
#########################################
Write-Host "Installing NI_LabVIEW"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}NI_LabVIEW2018SP1f3Patch_2018.0_v1.0.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\NI_LabVIEW2018SP1f3Patch_2018.0_v1.0.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\Windows\System32\wscript.exe"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @("C:\buildArtifacts\NI_LabVIEW2018SP1f3Patch_2018.0_v1.0\Package\NI_Labview2018_SP1_f3_v1.0.vbs") -Passthru
do {
    Start-Sleep -Seconds 1
    $processWscript = Get-Process -Name "wscript" -ErrorAction SilentlyContinue
} while ($null -ne $processWscript)
$processToKill = Get-Process -Name "nierserver" -ErrorAction SilentlyContinue
if ($null -ne $processToKill) {
    $processToKill | Stop-Process -Force
}
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#########################################
#endregion

#region OpenSSL
#########################################
Write-Host "Installing OpenSSL"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}OpenSSLInstallerTeam_OpenSSLX64_1.1.1b_v1.0.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\OpenSSLInstallerTeam_OpenSSLX64_1.1.1b_v1.0.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\OpenSSLInstallerTeam_OpenSSLX64_1.1.1b_v1.0\Package\Setup.EXE"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#########################################
#endregion

#region JMPPro
#########################################
Write-Host "Installing JMPPro"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}SAS_JMPPro_15.2_v1.0.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\SAS_JMPPro_15.2_v1.0.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\SAS_JMPPro_15.2_v1.0\Package\Deploy-Application.exe"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('Install', 'Silent') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#########################################
#endregion

#region OxygenXML
#########################################
Write-Host "Installing OxygenXMLAuthor"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}SyncROSoft_OxygenXMLAuthor_21.1_v1.0.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\SyncROSoft_OxygenXMLAuthor_21.1_v1.0.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\SyncROSoft_OxygenXMLAuthor_21.1_v1.0\Package\Setup.EXE"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#########################################
#endregion
#region OxygenXML 24.1
#########################################
Write-Host "Installing OxygenXMLAuthor 24.1"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}SyncROSoft_OxygenXMLAuthor_24.1_v1.1.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\SyncROSoft_OxygenXMLAuthor_24.1_v1.1.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\SyncROSoft_OxygenXMLAuthor_24.1_v1.1\Package\silent_install.bat"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#########################################
#endregion
#region ToadforOracle
#########################################
Write-Host "Installing ToadforOracle"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}ToadforOracle-11-R01-B01.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\ToadforOracle-11-R01-B01.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

Set-Location "C:\buildArtifacts\ToadforOracle-11-R01-B01\Installation"
#MSI 1
Write-Host "Installing Toad Prereq #1"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppInstaller = "C:\buildArtifacts\ToadforOracle-11-R01-B01\Installation\Toad_Standalone\11_0_0_116\Toad for Oracle 11.msi"
("Starting installer 1...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('TRANSFORMS="ToadforOracle-11-R01-B01.mst"', 'ALLUSERS=1', '/q', 'INSTALLDIR="C:\Program Files (x86)\Quest Software\Toad for Oracle 11\"') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#MSI 2
Write-Host "Installing Toad Prereq #2"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppInstaller = "C:\buildArtifacts\ToadforOracle-11-R01-B01\Installation\QuestSQLOptimizer_Oracle\8_5_0_2033\QuestSQLOptimizer_Oracle_8.5.0.2033.msi"
("Starting installer 2...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/q', 'INSTALLDIR="C:\Program Files (x86)\Quest Software\Quest SQL Optimizer for Oracle\"') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

# #MSI 3
# Write-Host "Installing Toad Prereq #3"
# Write-Host (Get-Date).ToString("o")
# $NonMSIXAppInstaller = "C:\buildArtifacts\ToadforOracle-11-R01-B01\Installation\Toad_Data_Modeler\4_1_5_8\ToadDataModeler_4.1.5.8.msi"
# ("Starting installer 3...") | Out-File $logFileLocation -Append
# #$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/q', 'INSTALLDIR="C:\Program Files (x86)\Quest Software\Toad Data Modeler 4.1"', 'ADDLOCAL=Complete,XPMANIFEST,SHORTCUTUNINSTALL,SHORTCUTDESKTOP,SHORTCUTSTARTMENU,SUPPORTDATABASES,UNIVERSAL_DATABASE,ORACLE_9,ORACLE_10G,ORACLE_11G_R1,ORACLE_11G_R2,ORACLE,ORACLE_11G') -Wait -Passthru
# $NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/q', 'INSTALLDIR="C:\Program Files (x86)\Quest Software\Toad Data Modeler 4.1"') -Wait -Passthru
# ("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#.NET 3.5
Write-Host "Installing .NET 3.5"
Write-Host (Get-Date).ToString("o")
("Enable-WindowsOptionalFeature .NET 3.5 ...") | Out-File $logFileLocation -Append
Enable-WindowsOptionalFeature -Online -FeatureName "NetFx3"
("Enable-WindowsOptionalFeature .NET 3.5 finished") | Out-File $logFileLocation -Append

#MSI 4
Write-Host "Installing Toad Prereq #3"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppInstaller = "C:\buildArtifacts\ToadforOracle-11-R01-B01\Installation\Toad_Data_Analysis\3_0_1_1734\ToadforDataAnalysts_3.0.1.1734.msi"
("Starting installer 4...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/q', 'INSTALLDIR="C:\Program Files (x86)\Quest Software\Toad for Data Analysts 3.0\"') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#########################################
#endregion

#region Visio
#########################################
Write-Host "Installing Visio Std 2016"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}VisioStd2016_64_MSI.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\VisioStd2016_64_MSI.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "setup.exe"

("Starting Visio download...") | Out-File $logFileLocation -Append
Set-Location "C:\buildArtifacts\VisioStd2016_64_MSI"
$Downloadstatus = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/download', '"InstallInternet_UpdateInternet_EN.xml"') -Wait -Passthru
Set-Location "C:\buildArtifacts\"
("Download finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

("Starting Visio install...") | Out-File $logFileLocation -Append
Set-Location "C:\buildArtifacts\VisioStd2016_64_MSI"
$Downloadstatus = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/configure', '"InstallInternet_UpdateInternet_EN.xml"') -Wait -Passthru
Set-Location "C:\buildArtifacts\"
("Install finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

("Activating Visio...") | Out-File $logFileLocation -Append
Set-Location "C:\Program Files\Microsoft Office\Office16"
cscript ospp.vbs /inpkey:9222V-NWBTY-49FXV-2CMHK-HCF26

#########################################
#endregion

#region SAP PreReq
#########################################
Write-Host "Installing SAP PreReq"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}SAP-PreReq.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\SAP-PreReq.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\SAP-PreReq\vstor_redist.exe"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/q', '/norestart') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#########################################
#endregion

#region SAP GUI
#########################################
Write-Host "Installing SAP GUI"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}SAPGUI770_new_MSoffice.exe{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstaller = "c:\buildArtifacts\SAPGUI770_new_MSoffice.exe"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstaller) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstaller -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/install', '/silent', '/force') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

#########################################
#endregion

#region Remove Included MSIX apps
#########################################
Write-Host "Removing MSIX from the Image"
Write-Host (Get-Date).ToString("o")
("Starting removal") | Out-File $logFileLocation -Append
$appname = @(
    "*BingWeather*"
    "*ZuneMusic*"
    "*ZuneVideo*"
    "*YourPhone*"
    "*XboxGameOverlay*"
    "*People*"
    "*microsoftskydrive*"
    "*XboxGamingOverlay*"
    "*XboxApp*"
    "*Getstarted*"
    "*SkypeApp*"
    "*OneNote*"
    "*OfficeHub*"
    "*GetHelp*"
    "*OneConnect*"
    "*Whiteboard*"
    "*WindowsStore*"
    "*MicrosoftSolitaireCollection*"
    "*MixedReality*"
    "*Messaging*"
    "*WindowsFeedbackHub*"
    "*549981C3F5F10*"
)

ForEach ($app in $appname) {
    Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like $app } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
    Get-AppxPackage -Name $app | Remove-AppxPackage -ErrorAction SilentlyContinue
    Write-Host Removal of $app Finished
}
("Removal finished.") | Out-File $logFileLocation -Append
#########################################
#endregion

#region Data Provider for Teradata Driver
#########################################
Write-Host "Installing Teradata Data Provider Driver"
Write-Host (Get-Date).ToString("o")

$NonMSIXAppUrl = ("{0}tddriver.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\TDNetDP.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\TDNetDP\.NET Data Provider for Teradata.msi"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/quiet', '/norestart') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append
#########################################
#endregion

#region Optional Windows features
#########################################
#RSAT tools
Write-Host "Adding RSAT"
Write-Host (Get-Date).ToString("o")
$logFileLocation = "c:\buildArtifacts\OptionalWindowsfeatures.log"
("Starting RSAT installation...") | Out-File $logFileLocation -Append
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
Add-WindowsCapability -Online -Name "Rsat.Dns.Tools~~~~0.0.1.0"
Add-WindowsCapability -Online -Name "Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0"
("Finished RSAT installation...") | Out-File $logFileLocation -Append
#Media Player
("Starting WindowsMediaPlayer installation...") | Out-File $logFileLocation -Append
Add-WindowsCapability -Online -Name "Media.WindowsMediaPlayer~~~~0.0.12.0"
("Finished WindowsMediaPlayer installation...") | Out-File $logFileLocation -Append
#endregion

#region LAPS
#########################################
Write-Host "Adding LAPS"
Write-Host (Get-Date).ToString("o")
$logFileLocation = "c:\buildArtifacts\LAPS.log"
$LapsUrl = ("{0}LAPS.x64.msi{1}" -f $aibrepoUrl, $aibrepoSas)
$LapsMsi = "c:\buildArtifacts\LAPS.x64.msi"
("Starting LAPS installation...") | Out-File $logFileLocation -Append

("Starting download LAPS...") | Out-File $logFileLocation -Append
Write-Host ("Starting download LAPS...")
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($LapsUrl, $LapsMsi)
("Download finished.") | Out-File $logFileLocation -Append
Write-Host ("Download finished.")

("Starting installer...") | Out-File $logFileLocation -Append
$Laps_install_status = Start-Process -FilePath $LapsMsi -ArgumentList @('/quiet') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $Laps_install_status.ExitCode) | Out-File $logFileLocation -Append
Write-Host ("Installer finished with returncode '{0}'" -f $Laps_install_status.ExitCode)

("Finished LAPS installation...") | Out-File $logFileLocation -Append
#endregion

#region Open JDK
#########################################
Write-Host "Adding Open JDK"
Write-Host (Get-Date).ToString("o")
$logFileLocation = "c:\buildArtifacts\OpenJDK.log"
$OpenJDKUrl = ("{0}java-1.8.0-openjdk-1.8.0.275-1.b01.dev.redhat.windows.x86_64.msi{1}" -f $aibrepoUrl, $aibrepoSas)
$OpenJDKMsi = "c:\buildArtifacts\java-1.8.0-openjdk-1.8.0.275-1.b01.dev.redhat.windows.x86_64.msi"
("Starting OpenJDK installation...") | Out-File $logFileLocation -Append

("Starting download OpenJDK...") | Out-File $logFileLocation -Append
Write-Host ("Starting download OpenJDK...")
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($OpenJDKUrl, $OpenJDKMsi)
("Download finished.") | Out-File $logFileLocation -Append
Write-Host ("Download finished.")

("Starting installer...") | Out-File $logFileLocation -Append
$OpenJDK_install_status = Start-Process -FilePath $OpenJDKMsi -ArgumentList @('/quiet') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $OpenJDK_install_status.ExitCode) | Out-File $logFileLocation -Append
Write-Host ("Installer finished with returncode '{0}'" -f $OpenJDK_install_status.ExitCode)

("Finished OpenJDK installation...") | Out-File $logFileLocation -Append
#endregion

#region MySQLWorkbench
#########################################
Write-Host "Adding MySQLWorkbench"
Write-Host (Get-Date).ToString("o")
$logFileLocation = "c:\buildArtifacts\MySQLWorkbench.log"
$MySQLWorkbenchUrl = ("{0}mysql-workbench-community-8.0.28-winx64.msi{1}" -f $aibrepoUrl, $aibrepoSas)
$MySQLWorkbenchMsi = "c:\buildArtifacts\mysql-workbench-community-8.0.28-winx64.msi"
("Starting MySQLWorkbench installation...") | Out-File $logFileLocation -Append

("Starting download MySQLWorkbench...") | Out-File $logFileLocation -Append
Write-Host ("Starting download MySQLWorkbench...")
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($MySQLWorkbenchUrl, $MySQLWorkbenchMsi)
("Download finished.") | Out-File $logFileLocation -Append
Write-Host ("Download finished.")

("Starting installer...") | Out-File $logFileLocation -Append
$MySQLWorkbench_install_status = Start-Process -FilePath $MySQLWorkbenchMsi -ArgumentList @('/qn /norestart') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $MySQLWorkbench_install_status.ExitCode) | Out-File $logFileLocation -Append
Write-Host ("Installer finished with returncode '{0}'" -f $MySQLWorkbench_install_status.ExitCode)

("Finished MySQLWorkbench installation...") | Out-File $logFileLocation -Append
#endregion

#region CiscoCLI
#########################################
Write-Host "Adding CiscoCLI"
Write-Host (Get-Date).ToString("o")
$logFileLocation = "c:\buildArtifacts\CiscoCLI.log"
$CiscoCLIUrl = ("{0}Cisco-CLI-Analyzer.3-6-8.x64.msi{1}" -f $aibrepoUrl, $aibrepoSas)
$CiscoCLIMsi = "c:\buildArtifacts\Cisco-CLI-Analyzer.3-6-8.x64.msi"
("Starting CiscoCLI installation...") | Out-File $logFileLocation -Append

("Starting download CiscoCLI...") | Out-File $logFileLocation -Append
Write-Host ("Starting download CiscoCLI...")
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($CiscoCLIUrl, $CiscoCLIMsi)
("Download finished.") | Out-File $logFileLocation -Append
Write-Host ("Download finished.")

("Starting installer...") | Out-File $logFileLocation -Append
$CiscoCLI_install_status = Start-Process -FilePath $CiscoCLIMsi -ArgumentList @('/qn /norestart') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $CiscoCLI_install_status.ExitCode) | Out-File $logFileLocation -Append
Write-Host ("Installer finished with returncode '{0}'" -f $CiscoCLI_install_status.ExitCode)

("Finished CiscoCLI installation...") | Out-File $logFileLocation -Append
#endregion

#region RealVNC
#########################################
Write-Host "Adding RealVNC"
Write-Host (Get-Date).ToString("o")
$logFileLocation = "c:\buildArtifacts\RealVNC.log"
$RealVNCUrl = ("{0}VNC-Viewer-6.21.1109-Windows.exe{1}" -f $aibrepoUrl, $aibrepoSas)
$RealVNCMsi = "c:\buildArtifacts\VNC-Viewer-6.21.1109-Windows.exe"
("Starting RealVNC installation...") | Out-File $logFileLocation -Append

("Starting download RealVNC...") | Out-File $logFileLocation -Append
Write-Host ("Starting download RealVNC...")
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($RealVNCUrl, $RealVNCMsi)
("Download finished.") | Out-File $logFileLocation -Append
Write-Host ("Download finished.")

("Starting installer...") | Out-File $logFileLocation -Append
$RealVNC_install_status = Start-Process -FilePath $RealVNCMsi -ArgumentList @('/qn REBOOT=ReallySuppress') -Wait -Passthru
("Installer finished with returncode '{0}'" -f $RealVNC_install_status.ExitCode) | Out-File $logFileLocation -Append
Write-Host ("Installer finished with returncode '{0}'" -f $RealVNC_install_status.ExitCode)

("Finished RealVNC installation...") | Out-File $logFileLocation -Append
#endregion

#region Python 3.7
#########################################
Write-Host "Installing Python 3.7"
Write-Host (Get-Date).ToString("o")

$NonMSIXAppUrl = ("{0}python-3.7.3-amd64.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\python-3.7.3-amd64.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\python-3.7.3-amd64.exe"

("Starting installer...") | Out-File $logFileLocation -Append
$Python3_7_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/quiet /log "python3_7.log" InstallAllUsers=1 PrependPath=1') -Wait -Passthru

("Installer finished with returncode '{0}'" -f $Python3_7_install_status.ExitCode) | Out-File $logFileLocation -Append
#endregion

#region Python 3.8
#########################################
Write-Host "Installing Python 3.8"
Write-Host (Get-Date).ToString("o")

$NonMSIXAppUrl = ("{0}python-3.8.2-amd64.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\python-3.8.2-amd64.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\python-3.8.2-amd64.exe"

("Starting installer...") | Out-File $logFileLocation -Append
$Python3_8_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/quiet /log "python3_8.log" InstallAllUsers=1 PrependPath=1') -Wait -Passthru

("Installer finished with returncode '{0}'" -f $Python3_8_install_status.ExitCode) | Out-File $logFileLocation -Append
#endregion

#region Python 3.10
#########################################
Write-Host "Installing Python 3.10"
Write-Host (Get-Date).ToString("o")

$NonMSIXAppUrl = ("{0}python-3.10.8-amd64.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\python-3.10.8-amd64.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\python-3.10.8-amd64.exe"

("Starting installer...") | Out-File $logFileLocation -Append
$Python3_10_install_status = Start-Process -FilePath $NonMSIXAppInstaller -ArgumentList @('/quiet /log "python3_10.log" InstallAllUsers=1 PrependPath=1') -Wait -Passthru

("Installer finished with returncode '{0}'" -f $Python3_10_install_status.ExitCode) | Out-File $logFileLocation -Append
#endregion


#region MSWebRTC
#########################################
Write-Host "Installing WebRTCs"
Write-Host (Get-Date).ToString("o")

$WVDWebRTCurl = "https://aka.ms/msrdcwebrtcsvc/msi"
$WebRTCInstaller = "c:\buildArtifacts\MsRdcWebRTCSvc_HostSetup.msi"

("MSWebRTC full download URL = '{0}'" -f $WVDWebRTCurl) | Out-File $logFileLocation -Append
("MSWebRTC download location = '{0}'" -f $WebRTCInstaller) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $WVDWebRTCurl -OutFile $WebRTCInstaller -UseBasicParsing

("Starting installer...") | Out-File $logFileLocation -Append
$WebRTC_install_status = Start-Process $WebRTCInstaller -ArgumentList @('/q' , '/n') -wait -PassThru

("Installer finished with returncode '{0}'" -f $WebRTC_install_statuss.ExitCode) | Out-File $logFileLocation -Append
#########################################
#endregion

#region SalesforceCLI
#########################################
Write-Host "Installing Salesforce CLI"
Write-Host (Get-Date).ToString("o")


$SalesforceCLIurl = "https://developer.salesforce.com/media/salesforce-cli/sfdx/channels/stable/sfdx-x64.exe"
$SalesforceCLIInstaller = "c:\buildArtifacts\sfdx.exe"


("SalesforceCLI full download URL = '{0}'" -f $SalesforceCLIurl) | Out-File $logFileLocation -Append
("SalesforceCLI download location = '{0}'" -f $SalesforceCLIInstaller) | Out-File $logFileLocation -Append


("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $SalesforceCLIurl -OutFile $SalesforceCLIInstaller -UseBasicParsing

("Starting installer...") | Out-File $logFileLocation -Append
$SalesforceCLI_install_status = Start-Process $SalesforceCLIInstaller -ArgumentList @('/S') -wait -PassThru

$pathToAppend = 'C:\Program Files\sfdx\bin'
$systemPathOriginal = [Environment]::GetEnvironmentVariable('PATH', [System.EnvironmentVariableTarget]::Machine)
[Environment]::SetEnvironmentVariable('PATH', "${systemPathOriginal};${pathToAppend}", [System.EnvironmentVariableTarget]::Machine)

("Installer finished with returncode '{0}'" -f $SalesforceCLI_install_statuss.ExitCode) | Out-File $logFileLocation -Append
#endregion

<#region QuickAssist
#########################################
Write-Host "Quick Assist"
Write-Host (Get-Date).ToString("o")

$NonMSIXAppUrl = ("{0}QuickAssist.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\QuickAssist.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\MicrosoftCorporationII.QuickAssist_2022.825.2016.0_neutral___8wekyb3d8bbwe.AppxBundle"

("Uninstalling QuickAssist Older Version ") | Out-File $logFileLocation -Append
Remove-WindowsCapability -Online -Name 'App.Support.QuickAssist~~~~0.0.1.0' -ErrorAction 'SilentlyContinue'
("QuickAssist Older Version Uninstalled ") | Out-File $logFileLocation -Append

("Starting New Quick Assist installer...") | Out-File $logFileLocation -Append
$QuickAssist_install_status = Add-AppxProvisionedPackage -online -SkipLicense -PackagePath $NonMSIXAppInstaller

("Installer finished with returncode '{0}'" -f $QuickAssist_install_status.ExitCode) | Out-File $logFileLocation -Append

#endregion#>

#region TIBCO BW 6.7
#########################################
Write-Host "Installing TIBCO TRA"
Write-Host (Get-Date).ToString("o")

$NonMSIXAppUrl = ("{0}TIB_BW_6.7.0_win_x86_64.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\TIB_BW_6.7.0_win_x86_64.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

    
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "c:\buildArtifacts\TIB_BW_6.7.0_win_x86_64\silent_install.bat"

("Starting installer...") | Out-File $logFileLocation -Append
$TIBCO_BW67_install_status = Start-Process -FilePath $NonMSIXAppInstaller -Wait -Passthru

#Setting up permissions for BW
$Path = "C:\Program Files\tibco_bw67"
$acl = Get-Acl $path
$Accessrule = New-Object System.Security.AccessControl.FileSystemAccessRule("NT AUTHORITY\Authenticated Users", "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")
$acl.SetAccessRule($AccessRule)
$acl | Set-Acl $path

$Path = "C:\Program Files\tibco_bw67\studio\4.0\eclipse\TIBCOBusinessStudio.ini"
$acl = Get-Acl $path
$acl.SetAccessRuleProtection($false,$true)
Set-Acl $path -AclObject $acl

$Path = "C:\Program Files\tibco_bw67\studio\4.0\eclipse\configuration\config.ini"
$acl = Get-Acl $path
$acl.SetAccessRuleProtection($false,$true)
Set-Acl $path -AclObject $acl

("Installer finished with returncode '{0}'" -f $TIBCO_BW67_install_status.ExitCode) | Out-File $logFileLocation -Append
#endregion

#region WinShuttle
#########################################
Write-Host "Installing WinShuttle"
Write-Host (Get-Date).ToString("o")
$NonMSIXAppUrl = ("{0}Winshuttle_Studio_20.0209.zip{1}" -f $aibrepoUrl, $aibrepoSas)
$NonMSIXAppInstallerZip = "c:\buildArtifacts\Winshuttle_Studio_20.0209.zip"

("NonMSIXApp full download URL = '{0}'" -f $NonMSIXAppUrl) | Out-File $logFileLocation -Append
("NonMSIXApp download location = '{0}'" -f $NonMSIXAppInstallerZip) | Out-File $logFileLocation -Append

("Starting download...") | Out-File $logFileLocation -Append
Invoke-WebRequest -Uri $NonMSIXAppUrl -OutFile $NonMSIXAppInstallerZip -UseBasicParsing
("Download finished.") | Out-File $logFileLocation -Append

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $NonMSIXAppInstallerZip -DestinationPath "c:\buildArtifacts\"
("Extraction finished.") | Out-File $logFileLocation -Append

$NonMSIXAppInstaller = "C:\buildArtifacts\Winshuttle_Studio_20.0209\Package\silent_install.bat"

("Starting installer...") | Out-File $logFileLocation -Append
$NonMSIXApp_install_status = Start-Process -FilePath $NonMSIXAppInstaller -Wait -Passthru
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append
#endregion

#region Winget Apps
#Warning - 7zip prereq for Language packs installation
#########################################
Write-Host "Installing Apps via Winget"
Write-Host (Get-Date).ToString("o")
("Starting installer...") | Out-File $logFileLocation -Append

Set-Location "c:\buildArtifacts\WinGet"

$installapps = @(
    "7zip.7zip"
    "Google.Chrome"
    "Citrix.Workspace"
    "Mozilla.Firefox"
    "Adobe.Acrobat.Reader.64-bit"
    "PuTTY.PuTTY"
    "ScooterSoftware.BeyondCompare4"
    "JetBrains.PyCharm.Community"
    "Microsoft.VisualStudioCode"
    "DominikReichl.KeePass"
    "Microsoft.PowerBI"
    "Git.Git"
    "Cisco.Jabber"
    "Amazon.AWSCLI"
    "Microsoft.Teams"
    "Microsoft.OneDrive"
    "WinSCP.WinSCP"
    "GnuPG.Gpg4win"
)
ForEach ($app in $installapps) {
    .\AppInstallerCLI.exe install --id $app --scope machine --silent --accept-package-agreements --accept-source-agreements
}
#.\AppInstallerCLI.exe install --id "Python.Python.3" --version 3.7.3150.0 --silent --accept-package-agreements --accept-source-agreements
#.\AppInstallerCLI.exe install --id "Python.Python.3" --version 3.8.2150.0 --silent --accept-package-agreements --accept-source-agreements
.\AppInstallerCLI.exe install --id "TortoiseSVN.TortoiseSVN" --version 1.14.0 --silent --accept-package-agreements --accept-source-agreements
.\AppInstallerCLI.exe install --id "Microsoft.AzureCLI" --silent --accept-package-agreements --accept-source-agreements
.\AppInstallerCLI.exe install --id "Microsoft.Powershell" --silent --accept-package-agreements --accept-source-agreements
("Installer finished with returncode '{0}'" -f $NonMSIXApp_install_status.ExitCode) | Out-File $logFileLocation -Append

Set-Location "c:\buildArtifacts"

#########################################
#endregion

#region Language Packs
# #########################################
# Write-Host "Adding additional Language Packs"
# Write-Host (Get-Date).ToString("o")
# if ($null -eq (Get-Item -Path "c:\buildArtifacts\LP" -ErrorAction SilentlyContinue)) {
#     New-Item -Path "c:\buildArtifacts\LP" -ItemType Directory -Force
# }

# $wvdLanguagePackScriptUrl = "<<wvdLanguagePackScriptUrl>>"
# $logFileLocation = "c:\buildArtifacts\wvdLanguagePack.log"
# $wvdLanguagePackScript = "c:\buildArtifacts\LP\Install-LanguagePacks.ps1"

# ("Starting download script...") | Out-File $logFileLocation -Append
# Write-Host ("Starting download script...")
# $wc = New-Object System.Net.WebClient
# $wc.DownloadFile($wvdLanguagePackScriptUrl, $wvdLanguagePackScript)
# ("Download finished.") | Out-File $logFileLocation -Append
# Write-Host ("Download finished.")

# Set-Location "c:\buildArtifacts\LP"
# ("Starting wvdLanguagePack...") | Out-File $logFileLocation -Append
# Write-Host ("Starting wvdLanguagePack...")
# .\Install-LanguagePacks.ps1
# ("wvdLanguagePack finished...") | Out-File $logFileLocation -Append
# Write-Host ("wvdLanguagePack finished...")
#endregion

#region Office365 Language Packs
#########################################
Write-Host "Adding Office365 Language Packs"
Write-Host (Get-Date).ToString("o")
if ($null -eq (Get-Item -Path "c:\buildArtifacts\O365LP" -ErrorAction SilentlyContinue)) {
    New-Item -Path "c:\buildArtifacts\O365LP" -ItemType Directory -Force
}
Set-Location "c:\buildArtifacts\O365LP"

$office365LanguagePackUrl = "<<office365LanguagePackUrl>>"
$logFileLocation = "c:\buildArtifacts\office365LanguagePack.log"
$office365LanguagePackZip = "c:\buildArtifacts\O365LP\Office365LP.zip"

("Starting download pack...") | Out-File $logFileLocation -Append
Write-Host ("Starting download pack...")
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($office365LanguagePackUrl, $office365LanguagePackZip)
("Download finished.") | Out-File $logFileLocation -Append
Write-Host ("Download finished.")

("Starting extraction...") | Out-File $logFileLocation -Append
Expand-Archive -Path $office365LanguagePackZip -DestinationPath "c:\buildArtifacts\O365LP\"
("Extraction finished.") | Out-File $logFileLocation -Append

Set-Location "C:\buildArtifacts\O365LP\O365"
("Starting office365LanguagePack...") | Out-File $logFileLocation -Append
Write-Host ("Starting office365LanguagePack...")
.\O365LPInstaller.ps1
("office365LanguagePack finished...") | Out-File $logFileLocation -Append
Write-Host ("office365LanguagePack finished...")
#endregion

# region Uninstall SilverLight  Apps
#########################################
Write-Host "Un-installing Microsoft SilverLight"
Write-Host (Get-Date).ToString("o")
("Starting Uninstallation...") | Out-File $logFileLocation -Append

Set-Location "c:\buildArtifacts\WinGet"

.\AppInstallerCLI.exe uninstall "Microsoft Silverlight" --silent --accept-source-agreements

("Silverlight Uninstallation Finished...") | Out-File $logFileLocation -Append

#########################################
#endregion

#region Cleanup shortcuts
#########################################
Write-Host "Removing shortcuts"
Write-Host (Get-Date).ToString("o")

Remove-Item -Path "C:\Program Files\Oxygen XML Author 21\lib\xproc\calabash\lib\log4j-core-2.1.jar" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:public\Desktop\Toad for Data Analysts 3.0.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:public\Desktop\Toad Data Modeler 4.1.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:public\Desktop\Acrobat Reader DC.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:public\Desktop\Beyond Compare 4.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:public\Desktop\KeePass.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\.NET Data Provider for Teradata 16.20.9" -Force -ErrorAction SilentlyContinue -Recurse
Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Java Development Kit" -Force -ErrorAction SilentlyContinue -Recurse
Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OpenSSL" -Force -ErrorAction SilentlyContinue -Recurse
Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Silverlight" -Force -ErrorAction SilentlyContinue -Recurse

#########################################
#endregion
#region AVD Optimizer
Write-Host "AVD Optimizer"
Set-Location "c:\buildArtifacts"
$AVDOptimizerUrl = ("{0}/{1}{2}" -f $avdOptContainerUrl, $avdOptZipFileName, $avdOptContainerSasToken)
$AVDOptimizerZip = "c:\buildArtifacts\AVDOptimizer.zip"

Write-Host ("AVD Optimizer full download URL = '{0}'" -f $AVDOptimizerUrl) 
Write-host ("AVD Optimizer download location = '{0}'" -f $AVDOptimizerZip)

Write-Host ("Starting download...")
Invoke-WebRequest -Uri $AVDOptimizerUrl -OutFile $AVDOptimizerZip -UseBasicParsing
Write-Host ("Download finished.")

Write-Host ("Starting extraction...")
Expand-Archive -Path $AVDOptimizerZip -DestinationPath "c:\buildArtifacts\"
Write-Host ("Extraction finished.")

Write-Host ("Running AVD Optimizer...")
$AvdoptimizerDirectory = Get-ChildItem -Filter "AVDOptimizer*" -Directory
Set-Location $AvdoptimizerDirectory.FullName
.\Windows_VDOT.ps1 -Optimizations All -Verbose -AcceptEula
Write-Host ("Execution finished.")

Set-Location "c:\buildArtifacts"
#endregion

try {
    & del "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\System Tools\Command Prompt.lnk"
    & del "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\System Tools\Control Panel.lnk"
    & del "C:\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\System Tools\Administrative Tools.lnk"
}
catch {}

Set-Location "c:\buildArtifacts"

try {
    $ErrorActionPreference = "SilentlyContinue"
    Remove-Item * -Include *.* -Exclude *.log -Recurse -Force -ErrorAction SilentlyContinue
}
catch {}

Write-Host "All done!"
try {
    Stop-Transcript
}
catch {}

Exit(0)

