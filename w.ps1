# Список программ для установки
$appsToInstall = @(
    "7zip.7zip", "Notepad++.Notepad++", "RustDesk.RustDesk", "AnyDesk.AnyDesk", 
    "VideoLAN.VLC", "PDFgear.PDFgear", "Google.Chrome", "Telegram.TelegramDesktop", 
    "Zoom.Zoom", "Yandex.Browser", "Yandex.Messenger", "AdrienAllard.FileConverter", 
    "alexx2000.DoubleCommander", "WinDirStat.WinDirStat", "Piriform.Recuva", 
    "PowerSoftware.AnyBurn", "qBittorrent.qBittorrent", 
    "9NKSQGP7F2NH", "XPDDT99J9GKB5C"
)

Write-Host "`n--- Проверка доступных обновлений ---" -ForegroundColor Cyan
# Получаем список обновляемых пакетов (пропускаем заголовок и разделители)
$updates = winget upgrade | Select-String -Pattern '^\S+' | Select-Object -Skip 2

if ($updates) {
    foreach ($line in $updates) {
        # Извлекаем имя и ID (обычно первая и вторая колонки)
        $fields = $line.ToString() -split '\s{2,}'
        $name = $fields[0].Trim()
        $id = $fields[1].Trim()

        if ($id -and $id -ne "ID") {
            $confirmUpdate = Read-Host "Обновить $name ($id)? [y/n]"
            if ($confirmUpdate -eq 'y') {
                Write-Host "Обновление $id..." -ForegroundColor Yellow
                winget upgrade --id $id --silent --accept-source-agreements --accept-package-agreements
            } else {
                Write-Host "Пропущено." -ForegroundColor Gray
            }
        }
    }
} else {
    Write-Host "Все программы уже обновлены." -ForegroundColor Green
}

Write-Host "`n--- Установка новых пакетов ---" -ForegroundColor Cyan

foreach ($app in $appsToInstall) {
    $confirmation = Read-Host "Установить $app? [y/n]"
    
    if ($confirmation -eq 'y') {
        Write-Host "Установка $app..." -NoNewline -ForegroundColor White
        
        $process = Start-Process winget -ArgumentList "install --id $app --silent --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru
        $exitCode = $process.ExitCode
        
        if ($exitCode -eq 0) {
            Write-Host "`r[ OK ] $app                          " -ForegroundColor Green
        } elseif ($exitCode -eq -1978335189) {
            Write-Host "`r[SKIP] $app (Уже установлена)        " -ForegroundColor Gray
        } else {
            Write-Host "`r[FAIL] $app (Ошибка: $exitCode)      " -ForegroundColor Red
        }
    } else {
        Write-Host "[Пропуск] $app" -ForegroundColor Yellow
    }
}

Write-Host "`nВсе задачи выполнены!" -ForegroundColor Cyan

# Самоудаление
if (Test-Path $PSCommandPath) { Remove-Item $PSCommandPath -Force }
