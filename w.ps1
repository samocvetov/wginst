$apps = @(
    "7zip.7zip", "Notepad++.Notepad++", "RustDesk.RustDesk", "AnyDesk.AnyDesk", 
    "VideoLAN.VLC", "PDFgear.PDFgear", "Google.Chrome.EXE", "Telegram.TelegramDesktop", 
    "9NKSQGP7F2NH", "Zoom.Zoom", "Yandex.Browser", "Yandex.Messenger", 
    "AdrienAllard.FileConverter", "alexx2000.DoubleCommander", "WinDirStat.WinDirStat"
)

Write-Host "--- Updating System ---" -ForegroundColor Cyan
winget upgrade --all --silent --accept-source-agreements --accept-package-agreements | Out-Null

Write-Host "--- Installing Packages ---" -ForegroundColor Cyan

foreach ($app in $apps) {
    # Скрываем технический вывод winget
    $null = winget install --id $app --silent --accept-source-agreements --accept-package-agreements
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[ OK ] $app" -ForegroundColor Green
    } elseif ($LASTEXITCODE -eq -1978335189) {
        Write-Host "[SKIP] $app (already installed)" -ForegroundColor Gray
    } else {
        Write-Host "[FAIL] $app (Code: $LASTEXITCODE)" -ForegroundColor Red
    }
}

Write-Host "`nDone!" -ForegroundColor Cyan

# Self-cleanup
if (Test-Path $PSCommandPath) { Remove-Item $PSCommandPath -Force }
