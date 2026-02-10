$appsToInstall = @(
    "7zip.7zip", "Notepad++.Notepad++", "RustDesk.RustDesk", "AnyDesk.AnyDesk", 
    "VideoLAN.VLC", "PDFgear.PDFgear", "Google.Chrome", "Telegram.TelegramDesktop", 
    "Zoom.Zoom", "Yandex.Browser", "Yandex.Messenger", "AdrienAllard.FileConverter", 
    "alexx2000.DoubleCommander", "WinDirStat.WinDirStat", "Piriform.Recuva", 
    "PowerSoftware.AnyBurn", "qBittorrent.qBittorrent", 
    "9NKSQGP7F2NH", "XPDDT99J9GKB5C"
)

Write-Host "`n--- Checking for available updates ---" -ForegroundColor Cyan
$updateRaw = winget upgrade --accept-source-agreements
$updates = $updateRaw | Select-String -Pattern '^\S+' | Select-Object -Skip 2

$foundUpdates = $false
foreach ($line in $updates) {
    $fields = $line.ToString() -split '\s{2,}'
    if ($fields.Count -gt 1) {
        $name = $fields[0].Trim()
        $id = $fields[1].Trim()
        if ($id -and $id -ne "ID" -and $id -ne "Name") {
            $foundUpdates = $true
            $confirmUpdate = Read-Host "Update available for $name ($id). Apply? [y/n]"
            if ($confirmUpdate -eq 'y') {
                Write-Host "Updating $id..." -ForegroundColor Yellow
                winget upgrade --id $id --silent --accept-source-agreements --accept-package-agreements
            }
        }
    }
}

if (-not $foundUpdates) {
    Write-Host "No updates required." -ForegroundColor Green
}

Write-Host "`n--- Installing new packages ---" -ForegroundColor Cyan
# Получаем список уже установленных программ один раз, чтобы не дергать winget в цикле
$installedList = winget list --accept-source-agreements | Out-String

foreach ($app in $appsToInstall) {
    # Если ID программы уже есть в списке установленных — просто скипаем без вопросов
    if ($installedList -like "*$app*") {
        Write-Host "[SKIP] $app (Already installed)" -ForegroundColor Gray
        continue
    }

    # Теперь имя точно будет видно
    $prompt = "Install " + $app + "? [y/n]"
    $confirmation = Read-Host $prompt
    
    if ($confirmation -eq 'y') {
        Write-Host "Processing $app..." -NoNewline -ForegroundColor White
        $process = Start-Process winget -ArgumentList "install --id $app --silent --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -eq 0) {
            Write-Host "`r[ OK ] $app                           " -ForegroundColor Green
        } else {
            Write-Host "`r[FAIL] $app (Error: $($process.ExitCode))" -ForegroundColor Red
        }
    }
}

Write-Host "`nDone!" -ForegroundColor Cyan
if (Test-Path $PSCommandPath) { Remove-Item $PSCommandPath -Force }
