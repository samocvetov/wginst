if(-not([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs;exit}

$Path="C:\ODT"
if(!(Test-Path $Path)){New-Item $Path -ItemType Directory|Out-Null}

Invoke-WebRequest -Uri "https://download.microsoft.com/download/office-deployment-tool.exe" -OutFile "$Path\odt.exe"
Start-Process "$Path\odt.exe" -ArgumentList "/quiet /extract:$Path" -Wait

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
"@ | Out-File "$Path\config.xml" -Encoding UTF8

Start-Process "$Path\setup.exe" -ArgumentList "/configure config.xml" -WorkingDirectory $Path -Wait

Write-Host "Office LTSC 2024 installation complete."
