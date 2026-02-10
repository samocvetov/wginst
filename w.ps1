# --- НАСТРОЙКИ СКРИПТА ---
$ScriptVersion = "6.2.15"

# Очищаем экран и выводим заголовок
Clear-Host
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "    WINGET AUTO-INSTALLER  |  v$ScriptVersion    " -ForegroundColor Yellow
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

$appsToInstall = @(
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

# --- ФУНКЦИЯ СОЗДАНИЯ ЯРЛЫКОВ ---
function Add-WingetShortcut {
    param (
        [string]$AppId
    )
    
    $StartMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
    $WingetLinksPath = "$env:LOCALAPPDATA\Microsoft\WinGet\Links"
    $WingetPackagesPath = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages"
    
    # СЛОВАРЬ ИСКЛЮЧЕНИЙ: Точное имя exe файла для поиска
    $exeOverrides = @{
        "ventoy.ventoy" = "Ventoy2Disk.exe"
    }

    # 1. Определяем имя файла для поиска
    if ($exeOverrides.ContainsKey($AppId)) {
        $searchFileName = $exeOverrides[$AppId]
    }
    elseif ($AppId -match "\.") {
        $cleanName = $AppId.Split('.')[-1]
        $searchFileName = "*$cleanName*.exe"
    } 
    else {
        return 
    }
    
    # ПЕРЕМЕННАЯ ДЛЯ ХРАНЕНИЯ НАЙДЕННОГО ФАЙЛА
    $targetFile = $null

    # 2. ПОПЫТКА №1: Ищем в папке Links (быстрый способ для portable)
    $targetFile = Get-ChildItem -Path $WingetLinksPath -Filter $searchFileName -ErrorAction SilentlyContinue | Select-Object -First 1

    # 3. ПОПЫТКА №2: (ФОЛЛБЭК) Если в Links нет, ищем в папке установки Packages
    if (-not $targetFile) {
        # Ищем папку пакета (по ID)
        $packageDir = Get-ChildItem -Path $WingetPackagesPath -Filter "${AppId}*" -Directory -ErrorAction SilentlyContinue | 
                      Sort-Object LastWriteTime -Descending | Select-Object -First 1
        
        if ($packageDir) {
            # Ищем exe внутри папки пакета (рекурсивно)
            $targetFile = Get-ChildItem -Path $packageDir.FullName -Filter $searchFileName -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        }
    }

    # 4. СОЗДАНИЕ ЯРЛЫКА
    if ($targetFile) {
        $shortcutName = $targetFile.BaseName
        $shortcutPath = "$StartMenuPath\$shortcutName.lnk"
        $realPath = $targetFile.FullName

        # Обработка симлинков
        try {
            if ($targetFile.LinkType -eq 'SymbolicLink') {
                $target = $targetFile.Target
                if (-not [System.IO.Path]::IsPathRooted($target)) {
                    $target = Join-Path $targetFile.DirectoryName $target
                }
                $realPath = (Get-Item $target).FullName
            }
        } catch {}

        # Удаляем старый ярлык
        if (Test-Path $shortcutPath) { Remove-Item $shortcutPath -Force }

        try {
            $WScript = New-Object -ComObject WScript.Shell
            $Shortcut = $WScript.CreateShortcut($shortcutPath)
            
            $Shortcut.TargetPath = $targetFile.FullName
            $Shortcut.WorkingDirectory = $targetFile.DirectoryName
            $Shortcut.IconLocation = "$realPath,0"
            
            $Shortcut.Save()
            Write-Host "   [+] Shortcut created: $shortcutName" -ForegroundColor DarkGray
        } catch {
            Write-Host "   [!] Failed to create shortcut" -ForegroundColor Red
        }
    }
}
# --------------------------------

Write-Host "--- Checking for available updates ---" -ForegroundColor Cyan
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
    $alreadyInstalled = $installedList -like "*$app*"
    
    if ($alreadyInstalled) {
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
            Write-Host "`r[ OK ] $displayName                       " -ForegroundColor Green
            Add-WingetShortcut -AppId $app
        } else {
            Write-Host "`r[FAIL] $displayName (Error: $($process.ExitCode))" -ForegroundColor Red
        }
    }
}

Write-Host "`nDone!" -ForegroundColor Cyan
Start-Sleep -Seconds 3
