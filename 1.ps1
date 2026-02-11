Clear-Host;Write-Host "=== WINGET AUTO-INSTALLER ==="
winget source reset --force|Out-Null
Start-Process wsreset.exe -Wait
winget source update --accept-source-agreements|Out-Null
winget upgrade --all --silent --include-unknown --accept-source-agreements --accept-package-agreements
$apps="7zip.7zip","Google.Chrome.EXE","Yandex.Browser","RustDesk.RustDesk","AnyDesk.AnyDesk","QL-Win.QuickLook","PDFgear.PDFgear","VideoLAN.VLC","AdrienAllard.FileConverter"
foreach($i in $apps){$f=winget list --id $i -e --accept-source-agreements 2>$null;if($f -match $i){Write-Host "[skip] $i";continue};Write-Host "Installing $i..." -NoNewline;$p=Start-Process winget -ArgumentList "install --id $i -e --silent --force --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru;if($p.ExitCode -eq 0){Write-Host " [ok]"}else{Write-Host " [fail]"}}
if(!(Test-Path "C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE")){$W="C:\ODT";if(!(Test-Path $W)){New-Item $W -ItemType Directory|Out-Null};$U="https://s.id/office-x64";$E="$W\setup.exe";Write-Host "Downloading Office Deployment Tool...";Invoke-WebRequest $U -OutFile $E -UseBasicParsing;@"
<Configuration>
<Add OfficeClientEdition="64" Channel="Current">
<Product ID="O365ProPlusRetail">
<Language ID="ru-ru"/>
</Product>
</Add>
<Display Level="None" AcceptEULA="TRUE"/>
</Configuration>
"@|Out-File "$W\config.xml" -Encoding UTF8;Write-Host "Downloading Office...";Start-Process $E -ArgumentList "/download config.xml" -WorkingDirectory $W -Wait;Write-Host "Installing Office...";Start-Process $E -ArgumentList "/configure config.xml" -WorkingDirectory $W -Wait;Write-Host "Microsoft 365 installation finished!"}else{Write-Host "[skip] Microsoft 365 already installed"}
Write-Host "Done";Start-Sleep 3
