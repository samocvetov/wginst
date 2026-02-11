Write-Host "=== WINGET AUTO-INSTALLER ==="
$apps = @("7zip.7zip","Google.Chrome","Yandex.Browser","RustDesk.RustDesk","AnyDesk.AnyDesk","QL-Win.QuickLook","PDFgear.PDFgear","VideoLAN.VLC","AdrienAllard.FileConverter")

$inst = winget list --accept-source-agreements | Out-String
foreach ($a in $apps) {
    if ($inst -notlike "*$a*") {
        Write-Host "Installing $a..." -NoNewline
        $p = Start-Process winget -Args "install --id $a -e --silent --accept-source-agreements --accept-package-agreements --force" -NoNewWindow -Wait -PassThru
        if ($p.ExitCode -eq 0) { Write-Host " [OK]" } else { Write-Host " [FAIL]" }
    }
}

Write-Host "`nChecking updates..."
winget upgrade --accept-source-agreements
if ((Read-Host "Update all? [y/n]") -eq 'y') { 
    winget upgrade --all --silent --accept-source-agreements --accept-package-agreements --include-unknown 
}
Write-Host "Done"; Start-Sleep 3
