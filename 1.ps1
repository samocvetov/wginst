Write-Host "=== WINGET AUTO-INSTALLER ==="
$c="--accept-source-agreements --accept-package-agreements --source winget"
winget source update --accept-source-agreements|Out-Null
winget list $c|Out-Null
Write-Host "Checking updates..."
$u=winget upgrade $c
if($LASTEXITCODE -eq 0 -and $u){Write-Host $u;if((Read-Host "Update all? [y/n]")-eq"y"){winget upgrade --all --silent --include-unknown $c}}else{Write-Host "System is up to date."}
$a="7zip.7zip","Google.Chrome.EXE","Yandex.Browser","RustDesk.RustDesk","AnyDesk.AnyDesk","QL-Win.QuickLook","PDFgear.PDFgear","VideoLAN.VLC","AdrienAllard.FileConverter"
foreach($i in $a){winget list --id $i -e $c|Out-Null;if($LASTEXITCODE -ne 0){Write-Host "Installing $i..." -NoNewline;winget install --id $i -e --silent --force $c|Out-Null;if($LASTEXITCODE -eq 0){Write-Host " [OK]"}else{Write-Host " [FAIL]"}}}
Write-Host "Done";Start-Sleep 3
