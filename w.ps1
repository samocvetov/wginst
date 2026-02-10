$appsToInstall = @(
    "7zip.7zip", "Notepad++.Notepad++", "RustDesk.RustDesk", "AnyDesk.AnyDesk", 
    "VideoLAN.VLC", "PDFgear.PDFgear", "Google.Chrome", "Telegram.TelegramDesktop", 
    "Zoom.Zoom", "Yandex.Browser", "Yandex.Messenger", "AdrienAllard.FileConverter", 
    "alexx2000.DoubleCommander", "WinDirStat.WinDirStat", "Piriform.Recuva", 
    "PowerSoftware.AnyBurn", "qBittorrent.qBittorrent", 
    "9NKSQGP7F2NH", "XPDDT99J9GKB5C"
)

$friendlyNames = @{
    "9NKSQGP7F2NH" = "WhatsApp"
    "XPDDT99J9GKB5C" = "Netflix"
}

Write-Host "`n--- Checking for available updates ---" -ForegroundColor Cyan
$updateRaw = winget upgrade --accept-source-agreements
$updates = $updateRaw | Select-String -Pattern '^\S+' | Select-Object -Skip 2

$foundUpdates = $false
foreach ($line in $updates) {
    $fields = $line.ToString() -split '\s{2,}'
    if ($fields.Count -gt 1) {
        $name = $fields[0].Trim()
        $id = $fields[1].Trim()
        
        # Фильтруем мусор и заголовки
        if ($id -and $id -ne "ID" -and $id -ne "Name" -and $id -notlike "*---*") {
            $foundUpdates = $true
            $confirmUpdate = Read-Host "Update available for $name ($id). Apply? [y/n]"
            if ($confirmUpdate -eq 'y') {
                Write-Host "Updating $id..." -ForegroundColor Yellow
                # Используем --exact, чтобы winget не путался в похожих названиях
                winget upgrade --id "$id" --exact --silent --force --accept-source-agreements --accept-package-agreements
            }
        }
    }
}

if (-not $foundUpdates) {
    Write-Host "No updates required." -ForegroundColor Green
}

Write-Host "`n--- Installing new packages ---" -ForegroundColor Cyan
$installedList = winget list --accept-source-agreements | Out-String

foreach ($app in $appsToInstall) {
    if ($installedList -like "*$app*") {
        Write-Host "[SKIP] $app (Already installed)" -ForegroundColor Gray
        continue
    }

    $displayName = if ($friendlyNames.ContainsKey($app)) { $friendlyNames[$app] } else { $app }
    $prompt = "Install " + $displayName + "? [y/n]"
    $confirmation = Read-Host $prompt
    
    if ($confirmation -eq 'y') {
        Write-Host "Processing $displayName..." -NoNewline -ForegroundColor White
        $process = Start-Process winget -ArgumentList "install --id $app --silent --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -eq 0) {
            Write-Host "`r[ OK ] $displayName                          " -ForegroundColor Green
        } else {
            Write-Host "`r[FAIL] $displayName (Error: $($process.ExitCode))" -ForegroundColor Red
        }
    }
}

Write-Host "`nDone!" -ForegroundColor Cyan
if (Test-Path $PSCommandPath) { Remove-Item $PSCommandPath -Force }
