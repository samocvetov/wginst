$apps = @(
    "7zip.7zip",
    "Notepad++.Notepad++",
    "RustDesk.RustDesk",
    "VideoLAN.VLC",
    "PDFgear.PDFgear",
    "Google.Chrome.EXE",
    "Telegram.TelegramDesktop",
    "9NKSQGP7F2NH", # WhatsApp
    "Zoom.Zoom",
    "Yandex.Browser",
    "Yandex.Disk",
    "Yandex.Messenger",
    "Yandex.Music",
    "AdrienAllard.FileConverter",
    "alexx2000.DoubleCommander",
    "WinDirStat.WinDirStat"
)

Write-Host "--- SYSTEM UPDATE STARTING ---" -ForegroundColor Cyan
winget upgrade --all --silent --accept-source-agreements --accept-package-agreements

Write-Host "`n--- INSTALLING PACKAGES ---" -ForegroundColor Cyan

foreach ($app in $apps) {
    Write-Host "Processing: $app..." -ForegroundColor Yellow
    winget install --id $app --silent --accept-source-agreements --accept-package-agreements

    if ($LASTEXITCODE -eq 0) {
        Write-Host "[SUCCESS] $app" -ForegroundColor Green
    } elseif ($LASTEXITCODE -eq -1978335189) {
        Write-Host "[SKIP] $app already installed" -ForegroundColor Gray
    } else {
        Write-Host "[ERROR] $app (Code: $LASTEXITCODE)" -ForegroundColor Red
    }
}

Write-Host "`n--- ALL TASKS COMPLETED ---" -ForegroundColor Cyan

# Self-cleanup
if (Test-Path "$PSScriptRoot\setup.ps1") {
    Remove-Item "$PSScriptRoot\setup.ps1" -Force
}
