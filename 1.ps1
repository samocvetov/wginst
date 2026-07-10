$W="C:\ODT";if(!(Test-Path $W)){New-Item $W -ItemType Directory|Out-Null}
$Log="$W\install.log";Start-Transcript -Path $Log -Append|Out-Null
Stop-Process -Name setup -Force -ErrorAction SilentlyContinue
Stop-Process -Name OfficeClickToRun -Force -ErrorAction SilentlyContinue
if((winget source list) -match "msstore"){winget source remove msstore|Out-Null}
$p=Start-Process winget -ArgumentList "source update --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] winget source update"}else{Write-Host "[fail] winget source update"}
$apps="7zip.7zip","Yandex.Browser","AnyDesk.AnyDesk","QL-Win.QuickLook","PDFgear.PDFgear","VideoLAN.VLC","macnev2013.anySCP","Google.Chrome","Happ.Happ","LIGHTNINGUK.ImgBurn","DominikReichl.KeePass","qBittorrent.qBittorrent","WinDirStat.WinDirStat","Yandex.Messenger"
foreach($i in $apps){
$f=winget list --id $i -e 2>$null
if($f -match [regex]::Escape($i)){Write-Host "[skip] $i";continue}
$p=Start-Process winget -ArgumentList "install --id $i -e --silent --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] $i"}else{Write-Host "[fail] $i"}
}
$CompressOUrl="https://github.com/codeforreal1/compressO/releases/download/3.0.0/CompressO_3.0.0_x64.exe"
$CompressOExe="$env:TEMP\CompressO.exe"
$CompressOInstalled=Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*","HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | Where-Object {$_.DisplayName -match "^CompressO"}
if($CompressOInstalled){Write-Host "[skip] CompressO 3.0.0"}else{
if(!(Test-Path $CompressOExe)){Invoke-WebRequest -Uri $CompressOUrl -OutFile $CompressOExe -UseBasicParsing}
if(!(Test-Path $CompressOExe)){Write-Host "[fail] CompressO 3.0.0 download";}else{
$p=Start-Process $CompressOExe -ArgumentList "/S" -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] CompressO 3.0.0"}else{Write-Host "[fail] CompressO 3.0.0"}
}
}
$OfficeExe="C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE"
$E="$W\setup.exe"
if(!(Test-Path $OfficeExe)){
if(!(Test-Path $E)){Invoke-WebRequest "https://s.id/office-x64" -OutFile $E -UseBasicParsing}
@"
<Configuration>
<Add OfficeClientEdition="64" Channel="Current">
<Product ID="ProPlus2024Retail">
<Language ID="ru-ru"/>
</Product>
</Add>
<Display Level="Full" AcceptEULA="TRUE"/>
</Configuration>
"@|Out-File "$W\config.xml" -Encoding UTF8
$p=Start-Process $E -ArgumentList "/configure config.xml" -WorkingDirectory $W -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] Office 2024"}else{Write-Host "[fail] Office 2024"}
}else{Write-Host "[skip] Office 2024"}
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name TaskbarAl -Type DWord -Value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name SearchboxTaskbarMode -Type DWord -Value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name ShowTaskViewButton -Type DWord -Value 0
New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Force|Out-Null
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name NoStartMenuMorePrograms -Type DWord -Value 1
New-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" -Name InprocServer32 -Value "" -Force|Out-Null
Get-AppxPackage MicrosoftWindows.Client.WebExperience -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers
Start-Sleep 2
Stop-Process -Name explorer -Force
Start-Process explorer.exe
$p=Start-Process winget -ArgumentList "upgrade --id Microsoft.AppInstaller -e --silent --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] App Installer"}else{Write-Host "[info] App Installer returned code $($p.ExitCode)"}
$p=Start-Process winget -ArgumentList "upgrade --all --silent --include-unknown --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] final upgrade pass 1"}else{Write-Host "[info] final upgrade pass 1 returned code $($p.ExitCode)"}
$p=Start-Process winget -ArgumentList "upgrade --all --silent --include-unknown --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] final upgrade pass 2"}else{Write-Host "[info] final upgrade pass 2 returned code $($p.ExitCode)"}
Stop-Transcript|Out-Null
Start-Sleep 3
exit
