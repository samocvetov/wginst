if(-not([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs;exit}

winget install --id Microsoft.OfficeDeploymentTool -e --accept-source-agreements --accept-package-agreements

$setup = Get-ChildItem "C:\Program Files" -Recurse -Filter setup.exe -ErrorAction SilentlyContinue | 
Where-Object { $_.FullName -like "*OfficeDeploymentTool*" } | Select-Object -First 1

if(!$setup){Write-Host "ODT setup.exe not found."; exit}

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

Start-Process $setup.FullName -ArgumentList "/configure `"$env:TEMP\config.xml`"" -Wait

Write-Host "Office LTSC 2024 installation finished!@"
