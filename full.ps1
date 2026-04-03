$W="C:\ODT";if(!(Test-Path $W)){New-Item $W -ItemType Directory|Out-Null}
$Log="$W\install.log";Start-Transcript -Path $Log -Append|Out-Null
Stop-Process -Name setup -Force -ErrorAction SilentlyContinue
Stop-Process -Name OfficeClickToRun -Force -ErrorAction SilentlyContinue
if((winget source list) -match "msstore"){winget source remove msstore|Out-Null}
$p=Start-Process winget -ArgumentList "source update --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] winget source update"}else{Write-Host "[fail] winget source update"}
$apps="7zip.7zip","Google.Chrome.EXE","Yandex.Browser","RustDesk.RustDesk","AnyDesk.AnyDesk","QL-Win.QuickLook","PDFgear.PDFgear","VideoLAN.VLC","AdrienAllard.FileConverter"
foreach($i in $apps){
$f=winget list --id $i -e 2>$null
if($f -match $i){Write-Host "[skip] $i";continue}
$p=Start-Process winget -ArgumentList "install --id $i -e --silent --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] $i"}else{Write-Host "[fail] $i"}
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
Get-AppxPackage MicrosoftWindows.Client.WebExperience -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers
Start-Sleep 2
Stop-Process -Name explorer -Force
$p=Start-Process winget -ArgumentList "upgrade --all --silent --include-unknown --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] final upgrade"}else{Write-Host "[info] final upgrade returned code $($p.ExitCode)"}
$p=Start-Process winget -ArgumentList "upgrade --all --silent --include-unknown --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] final upgrade"}else{Write-Host "[info] final upgrade returned code $($p.ExitCode)"}
Stop-Transcript|Out-Null
Start-Sleep 3
Start-Process powershell -Verb RunAs -ArgumentList "-Command `"& ([ScriptBlock]::Create((curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win | Out-String))) /Z-WindowsESUOffice`""
exit
