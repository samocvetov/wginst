# Исправляем кодировку для корректного отображения русского языка
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$apps = @(
    "7zip.7zip",
    "Notepad++.Notepad++",
    "RustDesk.RustDesk",
    "VideoLAN.VLC",
    "PDFgear.PDFgear",
    "Google.Chrome.EXE",
    "Telegram.TelegramDesktop",
    "9NKSQGP7F2NH",
    "Zoom.Zoom",
    "Yandex.Browser",
    "Yandex.Disk",
    "Yandex.Messenger",
    "Yandex.Music",
    "AdrienAllard.FileConverter",
    "alexx2000.DoubleCommander",
    "WinDirStat.WinDirStat"
)

Write-Host "`n--- Запуск автоматической установки программ ---" -ForegroundColor Cyan

if (!(Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "ОШИБКА: winget не найден!" -ForegroundColor Red
    exit
}

foreach ($app in $apps) {
    Write-Host "`n[+] Работаю с: $app" -ForegroundColor Yellow
    
    # Пытаемся установить. Если уже есть — winget сам это скажет.
    winget install --id $app --silent --accept-source-agreements --accept-package-agreements

    if ($LASTEXITCODE -eq 0) {
        Write-Host "Готово: $app" -ForegroundColor Green
    } elseif ($LASTEXITCODE -eq -1978335189) {
        Write-Host "Результат: $app уже установлена актуальная версия." -ForegroundColor Gray
    }
}

Write-Host "`n--- Все задачи выполнены! ---" -ForegroundColor Cyan
