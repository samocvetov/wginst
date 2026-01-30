[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$apps = @(
    "7zip.7zip",
    "Notepad++.Notepad++",
    "RustDesk.RustDesk",
    "VideoLAN.VLC",
    "PDFgear.PDFgear",
    "Google.Chrome.EXE",
    "Telegram.TelegramDesktop",
    "9NKSQGP7F2NH",             # WhatsApp (Microsoft Store ID)
    "Zoom.Zoom",
    "Yandex.Browser",
    "Yandex.Disk",
    "Yandex.Messenger",
    "Yandex.Music",
    "AdrienAllard.FileConverter",
    "alexx2000.DoubleCommander",
    "WinDirStat.WinDirStat"
)

Write-Host "--- Запуск автоматической установки программ ---" -ForegroundColor Cyan

# Проверка наличия winget
if (!(Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "ОШИБКА: winget не установлен. Установите 'App Installer' из Microsoft Store." -ForegroundColor Red
    exit
}

foreach ($app in $apps) {
    Write-Host "Пробую установить: $app..." -ForegroundColor Yellow
    
    # --accept-source-agreements и --accept-package-agreements позволяют пропустить подтверждение лицензий
    winget install --id $app --silent --accept-source-agreements --accept-package-agreements

    if ($LASTEXITCODE -eq 0) {
        Write-Host "Успешно: $app" -ForegroundColor Green
    } elseif ($LASTEXITCODE -eq -1978335189) {
        Write-Host "Пропущено: $app уже установлена." -ForegroundColor Gray
    } else {
        Write-Host "Ошибка при установке $app (Код: $LASTEXITCODE)" -ForegroundColor Red
    }
}

Write-Host "--- Установка завершена! ---" -ForegroundColor Cyan
