Clear-Host;Write-Host "=== WINGET AUTO-INSTALLER 1 ==="
winget upgrade --all --silent --include-unknown --accept-source-agreements --accept-package-agreements
winget source update --accept-source-agreements|Out-Null
winget list --accept-source-agreements --accept-package-agreements|Out-Null
$apps="7zip.7zip","Google.Chrome.EXE","Yandex.Browser","RustDesk.RustDesk","AnyDesk.AnyDesk","QL-Win.QuickLook","PDFgear.PDFgear","VideoLAN.VLC","AdrienAllard.FileConverter"
foreach($i in $apps){
$found=winget list --id $i -e --accept-source-agreements 2>$null
if($found -match $i){Write-Host "[skip] $i";continue}
Write-Host "Installing $i..." -NoNewline
$p=Start-Process winget -ArgumentList "install --id $i -e --silent --force --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru
if($p.ExitCode -eq 0){Write-Host " [ok]"}else{Write-Host " [fail]"}}
Write-Host "Done";Start-Sleep 3
