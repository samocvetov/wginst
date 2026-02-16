Start-Transcript -Path "$env:TEMP\oobe_automation.log" -Force
$TypeDef=@"
using System;
using System.Runtime.InteropServices;
namespace Api{public class Kernel32{[DllImport("kernel32.dll",CharSet=CharSet.Auto,SetLastError=true)]public static extern int OOBEComplete(ref bool bIsOOBEComplete);}}
"@
try{Add-Type -TypeDefinition $TypeDef -Language CSharp -ErrorAction Stop}catch{Stop-Transcript;exit 1}
$IsOOBEComplete=$false
[Api.Kernel32]::OOBEComplete([ref]$IsOOBEComplete)|Out-Null
if($IsOOBEComplete){Stop-Transcript;exit 0}
$Username="sysop"
$AnswerFile="$env:TEMP\UnattendOOBE.xml"
$UnattendXml=@"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
<settings pass="oobeSystem">
<component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
<OOBE>
<HideEULAPage>true</HideEULAPage>
<HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
<HideOnlineAccountScreens>true</HideOnlineAccountScreens>
<HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
<NetworkLocation>Work</NetworkLocation>
<ProtectYourPC>3</ProtectYourPC>
</OOBE>
<UserAccounts><LocalAccounts>
<LocalAccount wcm:action="add"><Name>$Username</Name><Group>Administrators</Group></LocalAccount>
</LocalAccounts></UserAccounts>
</component></settings></unattend>
"@
$UnattendXml|Out-File -FilePath $AnswerFile -Encoding utf8 -Force
$SysprepPath="$env:SystemRoot\System32\Sysprep\sysprep.exe"
if(Test-Path $SysprepPath){Start-Process -FilePath $SysprepPath -ArgumentList "/reboot /oobe /unattend:$AnswerFile" -Wait}
Stop-Transcript
