$W="C:\ODT";if(!(Test-Path $W)){New-Item $W -ItemType Directory|Out-Null}
$Log="$W\install.log";Start-Transcript -Path $Log -Append|Out-Null
if(!(Test-Path "C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE")){
$E="$W\setup.exe";Invoke-WebRequest "https://s.id/office-x64" -OutFile $E -UseBasicParsing
@"
<Configuration>
<Add OfficeClientEdition="64" Channel="Current">
<Product ID="O365ProPlusRetail">
<Language ID="ru-ru"/>
</Product>
</Add>
<Display Level="None" AcceptEULA="TRUE"/>
</Configuration>
"@|Out-File "$W\config.xml" -Encoding UTF8
Start-Process $E -ArgumentList "/download config.xml" -WorkingDirectory $W -Wait
$p=Start-Process $E -ArgumentList "/configure config.xml" -WorkingDirectory $W -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] Microsoft 365"}else{Write-Host "[fail] Microsoft 365"}
}else{Write-Host "[skip] Microsoft 365"}
winget source update --accept-source-agreements
winget upgrade --all --silent --include-unknown --accept-source-agreements --accept-package-agreements
$apps="7zip.7zip","Google.Chrome.EXE","Yandex.Browser","RustDesk.RustDesk","AnyDesk.AnyDesk","QL-Win.QuickLook","PDFgear.PDFgear","VideoLAN.VLC","AdrienAllard.FileConverter"
foreach($i in $apps){
$f=winget list --id $i -e --accept-source-agreements 2>$null
if($f -match $i){Write-Host "[skip] $i";continue}
$p=Start-Process winget -ArgumentList "install --id $i -e --silent --force --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] $i"}else{Write-Host "[fail] $i"}
}
Stop-Transcript|Out-Null
