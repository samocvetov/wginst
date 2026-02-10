$appsToInstall = @(
    "7zip.7zip", "Notepad++.Notepad++", "RustDesk.RustDesk", "AnyDesk.AnyDesk", "VideoLAN.VLC", 
    "PDFgear.PDFgear", "Google.Chrome", "Telegram.TelegramDesktop", "Zoom.Zoom",
    "Yandex.Browser", "Yandex.Messenger", "AdrienAllard.FileConverter", "alexx2000.DoubleCommander",
    "WinDirStat.WinDirStat", "Piriform.Recuva", "DominikReichl.KeePass",
    "ventoy.ventoy", "Termius.Termius", "WireGuard.WireGuard", "Mikrotik.Winbox",
    "REALiX.HWiNFO", "CPUID.CPU-Z", "TechPowerUp.GPU-Z", "angryziber.AngryIPScanner",
    "9NKSQGP7F2NH", "9NV4BS3L1H4S", "XPDDT99J9GKB5C"
)$appsToInstall = @(
    "7zip.7zip", "Notepad++.Notepad++", "RustDesk.RustDesk", "AnyDesk.AnyDesk", "VideoLAN.VLC", 
    "PDFgear.PDFgear", "Google.Chrome", "Telegram.TelegramDesktop", "Zoom.Zoom",
    "Yandex.Browser", "Yandex.Messenger", "AdrienAllard.FileConverter", "alexx2000.DoubleCommander",
    "WinDirStat.WinDirStat", "Piriform.Recuva", "DominikReichl.KeePass",
    "ventoy.ventoy", "Termius.Termius", "WireGuard.WireGuard", "Mikrotik.Winbox",
    "REALiX.HWiNFO", "CPUID.CPU-Z", "TechPowerUp.GPU-Z", "angryziber.AngryIPScanner",
    "9NKSQGP7F2NH", "9NV4BS3L1H4S", "XPDDT99J9GKB5C"
)

$friendlyNames = @{
    "9NKSQGP7F2NH" = "WhatsApp"
    "9NV4BS3L1H4S" = "QuickLook"
    "XPDDT99J9GKB5C" = "Samsung Magician"
}

# --- ИСПРАВЛЕННАЯ ФУНКЦИЯ (ДОБАВЛЕН IconLocation) ---
function Add-WingetShortcut {
    param (
        [string]$AppId
    )
    
    $StartMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
    $WingetLinksPath = "$env:LOCALAPPDATA\Microsoft\WinGet\Links"
    
    # 1. Определяем имя для поиска
    if ($AppId -match "\.") {
        $searchName = $AppId.Split('.')[-1]
    } else {
        return 
    }
    
    # 2. Ищем exe файл
    $targetExe = Get-ChildItem -Path $WingetLinksPath -Filter "*$searchName*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($targetExe) {
        $shortcutName = $targetExe.BaseName 
        $shortcutPath = "$StartMenuPath\$shortcutName.lnk"
        
        # Удаляем старый ярлык, если он был без иконки, чтобы пересоздать правильно
        if (Test-Path $shortcutPath) {
            Remove-Item $shortcutPath -Force
        }

        try {
            $WScript = New-Object -ComObject WScript.Shell
            $Shortcut = $WScript.CreateShortcut($shortcutPath)
            
            $Shortcut.TargetPath = $targetExe.FullName
            $Shortcut.WorkingDirectory = $targetExe.DirectoryName
            
            # --- ВАЖНО: Явно указываем путь к иконке ---
            # ",0" означает "взять первую иконку из файла"
            $Shortcut.IconLocation = "$($targetExe.FullName),0"
            # -------------------------------------------
            
            $Shortcut.Save()
            Write-Host "   [+] Created shortcut with icon: $shortcutName" -ForegroundColor DarkGray
        } catch {
            Write-Host "   [!] Failed to create shortcut" -ForegroundColor DarkGray
        }
    }
}
# --------------------------------

Write-Host "`n--- Checking for available updates ---" -ForegroundColor Cyan
$updateRaw = winget upgrade --accept-source-agreements
$lines = $updateRaw | Select-String -Pattern '^\S+' | Select-Object -Skip 2

$foundUpdates = $false
foreach ($line in $lines) {
    $columns = $line.ToString() -split '\s{2,}'
    if ($columns.Count -ge 2) {
        $name = $columns[0].Trim()
        $id = $columns[1].Trim()

        if ($id -and $id -ne "ID" -and $id -ne "Name" -and $id -notlike "---*") {
            $foundUpdates = $true
            if ($id -match "\s") { $id = ($id -split "\s")[0] }

            $confirmUpdate = Read-Host "Update available for $name ($id). Apply? [y/n]"
            if ($confirmUpdate -eq 'y') {
                Write-Host "Updating $id..." -ForegroundColor Yellow
                winget upgrade --id "$id" --silent --force --accept-source-agreements --accept-package-agreements
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
        # Пересоздаем ярлык даже для установленных, чтобы починить иконку
        Add-WingetShortcut -AppId $app
        continue
    }

    $displayName = if ($friendlyNames.ContainsKey($app)) { $friendlyNames[$app] } else { $app }
    $prompt = "Install " + $displayName + "? [y/n]"
    $confirmation = Read-Host $prompt
    
    if ($confirmation -eq 'y') {
        Write-Host "Processing $displayName..." -NoNewline -ForegroundColor White
        $process = Start-Process winget -ArgumentList "install --id $app --silent --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -eq 0) {
            Write-Host "`r[ OK ] $displayName                       " -ForegroundColor Green
            Add-WingetShortcut -AppId $app
        } else {
            Write-Host "`r[FAIL] $displayName (Error: $($process.ExitCode))" -ForegroundColor Red
        }
    }
}

Write-Host "`nDone!" -ForegroundColor Cyan
Start-Sleep -Seconds 3

$friendlyNames = @{
    "9NKSQGP7F2NH" = "WhatsApp"
    "9NV4BS3L1H4S" = "QuickLook"
    "XPDDT99J9GKB5C" = "Samsung Magician"
}

# --- ФУНКЦИЯ СОЗДАНИЯ ЯРЛЫКОВ (ПО ИМЕНИ ФАЙЛА) ---
function Add-WingetShortcut {
    param (
        [string]$AppId
    )
    
    # Эта функция работает только для портативных утилит, которые Winget кладет в папку Links.
    # Обычные установщики (Chrome, Zoom) создают ярлыки сами.
    
    # Пути
    $StartMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
    $WingetLinksPath = "$env:LOCALAPPDATA\Microsoft\WinGet\Links"
    
    # Пытаемся угадать имя файла. Берем часть ID после последней точки.
    # Например: из "Mikrotik.Winbox" берем "Winbox"
    if ($AppId -match "\.") {
        $searchName = $AppId.Split('.')[-1]
    } else {
        return # Если это ID из Store (без точек), пропускаем, они сами ставятся в пуск
    }
    
    # Ищем любой .exe файл, содержащий это имя
    $targetExe = Get-ChildItem -Path $WingetLinksPath -Filter "*$searchName*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($targetExe) {
        # Имя ярлыка берем из имени самого файла (без .exe)
        # Например: winbox64.exe -> ярлык будет называться "winbox64"
        $shortcutName = $targetExe.BaseName 
        $shortcutPath = "$StartMenuPath\$shortcutName.lnk"
        
        if (-not (Test-Path $shortcutPath)) {
            try {
                $WScript = New-Object -ComObject WScript.Shell
                $Shortcut = $WScript.CreateShortcut($shortcutPath)
                $Shortcut.TargetPath = $targetExe.FullName
                $Shortcut.WorkingDirectory = $targetExe.DirectoryName
                $Shortcut.Save()
                
                # Иконка подтянется автоматически из TargetPath
                Write-Host "   [+] Created shortcut: $shortcutName" -ForegroundColor DarkGray
            } catch {
                Write-Host "   [!] Failed to create shortcut" -ForegroundColor DarkGray
            }
        }
    }
}
# --------------------------------

Write-Host "`n--- Checking for available updates ---" -ForegroundColor Cyan
$updateRaw = winget upgrade --accept-source-agreements
$lines = $updateRaw | Select-String -Pattern '^\S+' | Select-Object -Skip 2

$foundUpdates = $false
foreach ($line in $lines) {
    $columns = $line.ToString() -split '\s{2,}'
    if ($columns.Count -ge 2) {
        $name = $columns[0].Trim()
        $id = $columns[1].Trim()

        if ($id -and $id -ne "ID" -and $id -ne "Name" -and $id -notlike "---*") {
            $foundUpdates = $true
            if ($id -match "\s") { $id = ($id -split "\s")[0] }

            $confirmUpdate = Read-Host "Update available for $name ($id). Apply? [y/n]"
            if ($confirmUpdate -eq 'y') {
                Write-Host "Updating $id..." -ForegroundColor Yellow
                winget upgrade --id "$id" --silent --force --accept-source-agreements --accept-package-agreements
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
        # Проверяем, нужен ли ярлык, даже если приложение уже установлено
        Add-WingetShortcut -AppId $app
        continue
    }

    $displayName = if ($friendlyNames.ContainsKey($app)) { $friendlyNames[$app] } else { $app }
    $prompt = "Install " + $displayName + "? [y/n]"
    $confirmation = Read-Host $prompt
    
    if ($confirmation -eq 'y') {
        Write-Host "Processing $displayName..." -NoNewline -ForegroundColor White
        $process = Start-Process winget -ArgumentList "install --id $app --silent --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru
        if ($process.ExitCode -eq 0) {
            Write-Host "`r[ OK ] $displayName                       " -ForegroundColor Green
            
            # --- Создаем ярлык после успешной установки ---
            Add-WingetShortcut -AppId $app
            
        } else {
            Write-Host "`r[FAIL] $displayName (Error: $($process.ExitCode))" -ForegroundColor Red
        }
    }
}

Write-Host "`nDone!" -ForegroundColor Cyan
Start-Sleep -Seconds 3
