if(-not([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs;exit}

winget install --id Microsoft.OfficeDeploymentTool -e --accept-source-agreements --accept-package-agreements

$odtPath="$env:ProgramFiles\Microsoft Office\Office Deployment Tool"

@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume">
      <Language ID="ru-ru"/>
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE"/>
  <Property Name="AUTOACTIVATE" Value="1"/>
</Configuration>
"@ | Out-File "$env:TEMP\config.xml" -Encoding UTF8

Start-Process "$odtPath\setup.exe" -ArgumentList "/configure `"$env:TEMP\config.xml`"" -Wait

Write-Host "Office LTSC 2024 installation finished!!!"
