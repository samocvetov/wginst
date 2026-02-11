$W="C:\ODT";if(!(Test-Path $W)){New-Item $W -ItemType Directory|Out-Null}
$Log="$W\install.log";Start-Transcript -Path $Log -Append|Out-Null
Stop-Process -Name setup -Force -ErrorAction SilentlyContinue
Stop-Process -Name OfficeClickToRun -Force -ErrorAction SilentlyContinue
if((winget source list) -match "msstore"){winget source remove msstore|Out-Null}
Start-Process winget -ArgumentList "source update --disable-interactivity --nowarn" -NoNewWindow -Wait|Out-Null
Write-Host "[ok] winget source update"
$apps="7zip.7zip","Google.Chrome.EXE","Yandex.Browser","RustDesk.RustDesk","AnyDesk.AnyDesk","QL-Win.QuickLook","PDFgear.PDFgear","VideoLAN.VLC","AdrienAllard.FileConverter"
$max=3
$jobs=@()
function Get-InstallerType($id){
$t=winget show --id $id -e 2>$null
if($t -match "Installer Type:\s+msi"){return "msi"}
if($t -match "Installer Type:\s+exe"){return "exe"}
return "exe"
}
foreach($i in $apps){
$f=winget list --id $i -e 2>$null
if($f -match $i){Write-Host "[skip] $i";continue}
$type=Get-InstallerType $i
if($type -eq "msi"){
$p=Start-Process winget -ArgumentList "install --id $i -e --silent --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] $i"}else{Write-Host "[fail] $i"}
}else{
while(($jobs | Where-Object {$_.State -eq "Running"}).Count -ge $max){
Start-Sleep 1
}
$jobs+=Start-Job -ArgumentList $i -ScriptBlock {
param($app)
winget install --id $app -e --silent --accept-package-agreements --disable-interactivity --nowarn
return @{App=$app;Code=$LASTEXITCODE}
}
}
}
foreach($j in $jobs){
$r=Receive-Job -Job $j -Wait
if($r.Code -eq 0){Write-Host "[ok] $($r.App)"}else{Write-Host "[fail] $($r.App)"}
Remove-Job $j
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
$p=Start-Process winget -ArgumentList "upgrade --all --silent --include-unknown --accept-package-agreements --disable-interactivity --nowarn" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host "[ok] final upgrade"}else{Write-Host "[fail] final upgrade"}
Stop-Transcript|Out-Null
if($p.ExitCode -eq 0){
Start-Process powershell -ArgumentList "-NoExit -Command winget upgrade --all --silent --include-unknown --accept-package-agreements --disable-interactivity --nowarn"
exit
}
