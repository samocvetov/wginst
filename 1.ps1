$W="C:\ODT";if(!(Test-Path $W)){New-Item $W -ItemType Directory|Out-Null}
$Log="$W\install.log";Start-Transcript -Path $Log -Append|Out-Null
Stop-Process -Name setup -Force -ErrorAction SilentlyContinue
Stop-Process -Name OfficeClickToRun -Force -ErrorAction SilentlyContinue
if((winget source list) -match "msstore"){winget source remove msstore}
$OfficeExe="C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE"
$E="$W\setup.exe"
if(!(Test-Path $OfficeExe)){
if(!(Test-Path $E)){Invoke-WebRequest "https://s.id/office-x64" -OutFile $E -UseBasicParsing}
@"
<Configuration>
<Add OfficeClientEdition="64" Channel="Current">
<Product ID="O365ProPlusRetail">
<Language ID="ru-ru"/>
</Product>
</Add>
<Display Level="Full" AcceptEULA="TRUE"/>
</Configuration>
"@|Out-File "$W\config.xml" -Encoding UTF8
$p=Start-Process $E -ArgumentList "/configure config.xml" -WorkingDirectory $W -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] Microsoft 365"}else{Write-Host "[fail] Microsoft 365"}
}else{Write-Host "[skip] Microsoft 365"}
Start-Process winget -ArgumentList "source update" -NoNewWindow -Wait
Start-Process winget -ArgumentList "upgrade --all --silent --include-unknown --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait
$apps="7zip.7zip","Google.Chrome.EXE","Yandex.Browser","RustDesk.RustDesk","AnyDesk.AnyDesk","QL-Win.QuickLook","PDFgear.PDFgear","VideoLAN.VLC","AdrienAllard.FileConverter"
foreach($i in $apps){
$f=winget list --id $i -e 2>$null
if($f -match $i){Write-Host "[skip] $i";continue}
$p=Start-Process winget -ArgumentList "install --id $i -e --silent --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] $i"}else{Write-Host "[fail] $i"}
}
Stop-Transcript|Out-Null
