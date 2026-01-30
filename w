$apps = @("7zip.7zip", "Notepad++.Notepad++", "RustDesk.RustDesk", "AnyDesk.AnyDesk", "VideoLAN.VLC", "PDFgear.PDFgear", "Google.Chrome.EXE", "Telegram.TelegramDesktop", "9NKSQGP7F2NH", "Zoom.Zoom", "Yandex.Browser", "Yandex.Messenger", "AdrienAllard.FileConverter", "alexx2000.DoubleCommander", "WinDirStat.WinDirStat")

Write-Host "Обновление системы..." -ForegroundColor Cyan
winget upgrade --all --silent --accept-source-agreements --accept-package-agreements | Out-Null

Write-Host "Установка программ:" -ForegroundColor Cyan

foreach ($app in $apps) {
    # Пытаемся установить, подавляя весь вывод самого winget
    $process = winget install --id $app --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[ OK ] $app" -ForegroundColor Green
    } elseif ($LASTEXITCODE -eq -1978335189) {
        Write-Host "[SKIP] $app (уже есть)" -ForegroundColor Gray
    } else {
        Write-Host "[FAIL] $app (Код: $LASTEXITCODE)" -ForegroundColor Red
    }
}

Write-Host "`nГотово!" -ForegroundColor Cyan

# Самоудаление (если файл существует)
if (Test-Path $PSCommandPath) { Remove-Item $PSCommandPath -Force }
