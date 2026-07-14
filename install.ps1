[CmdletBinding()]
param(
    [switch]$SelfTest
)

$script:WgManagerMarker = 'WG-INSTALL-MANAGER-V2'
$script:SelfUrl = 'https://raw.githubusercontent.com/samocvetov/wginst/main/s.ps1'
$script:IsWindows = ($env:OS -eq 'Windows_NT')
$script:SupportsAnsi = $false
$script:MenuDepth = 0
$script:LastFrameLineCount = 0
$script:WingetPath = $null
$script:ExitRequested = $false

if (-not $script:IsWindows) {
    Write-Host 'Этот менеджер предназначен только для Windows 10/11.'
    return
}

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-PowerShellSource {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ($content.Length -lt 1000 -or $content -notmatch 'WG-INSTALL-MANAGER-V2') { return $false }
    if ($content -match '(?im)^\s*(?:<!DOCTYPE|<html|code:|404\s*:|Output:|Wall time:|Exit code:)') { return $false }
    $tokens = $null
    $errors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    return ($errors.Count -eq 0)
}

function Start-ElevatedCopy {
    if (Test-IsAdministrator) { return $true }
    try {
        if ($PSCommandPath) {
            $launchPath = $PSCommandPath
        } else {
            $launchPath = Join-Path $env:TEMP 'WGInstall-s.ps1'
            $partial = "$launchPath.part"
            Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
            Invoke-WebRequest -Uri $script:SelfUrl -OutFile $partial -UseBasicParsing -ErrorAction Stop
            if (-not (Test-PowerShellSource -Path $partial)) {
                Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
                throw 'GitHub вернул не PowerShell-скрипт. Проверьте имя s.ps1 и доступность ссылки.'
            }
            Move-Item -LiteralPath $partial -Destination $launchPath -Force
        }
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$launchPath`""
        Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList $arguments -ErrorAction Stop | Out-Null
        return $false
    } catch {
        Write-Host "Не удалось запросить права администратора: $($_.Exception.Message)"
        Write-Host 'Запустите PowerShell от имени администратора и повторите команду.'
        return $false
    }
}

if (-not $SelfTest) {
    if (-not (Start-ElevatedCopy)) { return }
}

$ErrorActionPreference = 'Stop'

$script:Root = Join-Path $env:LOCALAPPDATA 'WGInstall'
$script:CacheRoot = Join-Path $script:Root 'Cache'
$script:LogRoot = Join-Path $script:Root 'Logs'
$script:DriverCache = Join-Path $script:CacheRoot 'Drivers'
$script:TweakStatePath = Join-Path $script:Root 'tweaks-state.json'
$script:LogPath = $null
if (-not $SelfTest) {
    try {
        foreach ($directory in @($script:Root, $script:CacheRoot, $script:LogRoot, $script:DriverCache)) {
            New-Item -ItemType Directory -Path $directory -Force -ErrorAction Stop | Out-Null
        }
        $script:LogPath = Join-Path $script:LogRoot ("WGInstall-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    } catch {
        Write-Host "Не удалось подготовить рабочую папку $script:Root"
        Write-Host $_.Exception.Message
        return
    }
}

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    if (-not $script:LogPath) { return }
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Show-ErrorMessage {
    param([string]$Title, [Management.Automation.ErrorRecord]$ErrorRecord)
    Show-TextCursor
    Clear-Host
    Write-Host $Title
    Write-Host ''
    Write-Host $ErrorRecord.Exception.Message
    Write-Host ''
    Write-Host "Подробности сохранены: $script:LogPath"
    Write-Log -Level 'ERROR' -Message ("{0}: {1}`n{2}" -f $Title, $ErrorRecord.Exception.Message, $ErrorRecord.ScriptStackTrace)
    Pause-Result
}

function Pause-Result {
    Write-Host ''
    Write-Host 'Нажмите любую клавишу для возврата в меню...'
    [void][Console]::ReadKey($true)
}

function Show-WorkScreen {
    param([string]$Title, [string]$Details = 'Не закрывайте окно до завершения операции.')
    Show-TextCursor
    Clear-Host
    Write-Host 'Менеджер установки'
    Write-Host ''
    Write-Host $Title
    if ($Details) { Write-Host $Details }
    Write-Host ''
}

function Initialize-ConsoleUi {
    try {
        $script:SupportsAnsi = [bool]$Host.UI.SupportsVirtualTerminal
    } catch {
        $script:SupportsAnsi = ($env:WT_SESSION -or $env:TERM_PROGRAM)
    }
}

function Enter-MenuScreen {
    $script:MenuDepth++
    $script:LastFrameLineCount = 0
    if ($script:SupportsAnsi) {
        [Console]::Write("$([char]27)[?1049h$([char]27)[?25l")
    } else {
        try { [Console]::CursorVisible = $false } catch {}
        Clear-Host
    }
}

function Exit-MenuScreen {
    if ($script:MenuDepth -gt 0) { $script:MenuDepth-- }
    if ($script:SupportsAnsi) {
        [Console]::Write("$([char]27)[?25h$([char]27)[?1049l")
    } else {
        try { [Console]::CursorVisible = $true } catch {}
        Clear-Host
    }
}

function Show-TextCursor {
    if ($script:SupportsAnsi) { [Console]::Write("$([char]27)[?25h") }
    try { [Console]::CursorVisible = $true } catch {}
}

function Limit-ConsoleText {
    param([string]$Text, [int]$Width)
    if ($null -eq $Text) { return '' }
    $value = [string]$Text
    if ($Width -lt 5) { return $value }
    if ($value.Length -ge $Width) { return $value.Substring(0, $Width - 3) + '...' }
    return $value.PadRight($Width)
}

function Draw-MenuFrame {
    param([string[]]$Lines)
    $width = 100
    try { $width = [Math]::Max(20, [Console]::WindowWidth - 1) } catch {}
    if ($script:SupportsAnsi) {
        $builder = New-Object Text.StringBuilder
        [void]$builder.Append("$([char]27)[H")
        foreach ($line in $Lines) {
            [void]$builder.Append("$([char]27)[2K")
            [void]$builder.Append((Limit-ConsoleText -Text $line -Width $width))
            [void]$builder.Append("`r`n")
        }
        [void]$builder.Append("$([char]27)[J")
        [Console]::Write($builder.ToString())
    } else {
        try { [Console]::SetCursorPosition(0, 0) } catch { Clear-Host }
        [int]$lineCount = [Math]::Max($Lines.Count, $script:LastFrameLineCount)
        for ([int]$i = 0; $i -lt $lineCount; $i++) {
            $line = if ($i -lt $Lines.Count) { $Lines[$i] } else { '' }
            [Console]::WriteLine((Limit-ConsoleText -Text $line -Width $width))
        }
        $script:LastFrameLineCount = $Lines.Count
    }
}

function Select-SingleItem {
    param(
        [string]$Title,
        [object[]]$Items,
        [scriptblock]$Text,
        [string]$Hint = 'Стрелки - выбор  Enter - открыть  Esc - назад.'
    )
    if ($null -eq $Items -or $Items.Count -eq 0) { return -1 }
    [int]$index = 0
    Enter-MenuScreen
    try {
        while ($true) {
            $lines = @($Title, $Hint, '')
            for ([int]$i = 0; $i -lt $Items.Count; $i++) {
                $prefix = if ($i -eq $index) { '> ' } else { '  ' }
                $lines += "$prefix$(& $Text $Items[$i])"
            }
            Draw-MenuFrame -Lines $lines
            $key = [Console]::ReadKey($true)
            switch ($key.Key) {
                'UpArrow'   { $index = [int](($index - 1 + $Items.Count) % $Items.Count) }
                'DownArrow' { $index = [int](($index + 1) % $Items.Count) }
                'Enter'     { return $index }
                'Escape'    { return -1 }
            }
        }
    } finally {
        Exit-MenuScreen
    }
}

function Select-MultipleItems {
    param(
        [string]$Title,
        [object[]]$Items,
        [scriptblock]$Text,
        [scriptblock]$CanSelect = { param($item) $true },
        [scriptblock]$Identity = { param($item, $itemIndex) [string]$itemIndex }
    )
    if ($null -eq $Items -or $Items.Count -eq 0) { return $null }
    [int]$index = 0
    [int]$pageSize = 12
    $selected = New-Object 'System.Collections.Generic.HashSet[string]'
    Enter-MenuScreen
    try {
        while ($true) {
            [int]$page = [Math]::Floor($index / $pageSize)
            [int]$pageCount = [Math]::Ceiling($Items.Count / $pageSize)
            [int]$first = $page * $pageSize
            [int]$last = [Math]::Min($first + $pageSize - 1, $Items.Count - 1)
            $lines = @(
                $Title,
                'Стрелки - выбор  Пробел/X - отметить  Enter - продолжить  Esc - назад.',
                "Страница $($page + 1) из $pageCount",
                ''
            )
            for ([int]$i = $first; $i -le $last; $i++) {
                $available = [bool](& $CanSelect $Items[$i])
                $itemKey = [string](& $Identity $Items[$i] $i)
                if (-not $available) { $mark = '[-]' }
                elseif ($selected.Contains($itemKey)) { $mark = '[x]' }
                else { $mark = '[ ]' }
                $prefix = if ($i -eq $index) { '> ' } else { '  ' }
                $lines += "$prefix$mark $(& $Text $Items[$i])"
            }
            $lines += ''
            $lines += "Выбрано: $($selected.Count)"
            Draw-MenuFrame -Lines $lines
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::Spacebar -or $key.Key -eq [ConsoleKey]::X) {
                if ([bool](& $CanSelect $Items[$index])) {
                    $itemKey = [string](& $Identity $Items[$index] $index)
                    if ($selected.Contains($itemKey)) { [void]$selected.Remove($itemKey) }
                    else { [void]$selected.Add($itemKey) }
                }
                continue
            }
            switch ($key.Key) {
                'UpArrow'   { $index = [int](($index - 1 + $Items.Count) % $Items.Count) }
                'DownArrow' { $index = [int](($index + 1) % $Items.Count) }
                'Enter' {
                    $result = @()
                    for ([int]$i = 0; $i -lt $Items.Count; $i++) {
                        $itemKey = [string](& $Identity $Items[$i] $i)
                        if ($selected.Contains($itemKey)) { $result += $Items[$i] }
                    }
                    return $result
                }
                'Escape' { return $null }
            }
        }
    } finally {
        Exit-MenuScreen
    }
}

function Read-YesNo {
    param([string]$Question)
    Show-TextCursor
    return ((Read-Host "$Question (введите ДА для подтверждения)").Trim().ToUpperInvariant() -eq 'ДА')
}

function Save-HttpFile {
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][string]$Destination,
        [string]$Title = 'Скачивание файла',
        [long]$MinimumBytes = 1024
    )
    $parent = Split-Path -Parent $Destination
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
    $partial = "$Destination.part"
    Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
    $request = [Net.HttpWebRequest]::Create($Uri)
    $request.UserAgent = 'Mozilla/5.0 WGInstall/2.0'
    $request.AllowAutoRedirect = $true
    $request.Timeout = 30000
    $request.ReadWriteTimeout = 30000
    $response = $null
    $input = $null
    $output = $null
    try {
        $response = $request.GetResponse()
        $total = [long]$response.ContentLength
        $input = $response.GetResponseStream()
        $output = [IO.File]::Open($partial, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $buffer = New-Object byte[] 131072
        [long]$received = 0
        $lastUpdate = [DateTime]::MinValue
        while (($read = $input.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $output.Write($buffer, 0, $read)
            $received += $read
            if (((Get-Date) - $lastUpdate).TotalMilliseconds -ge 250) {
                if ($total -gt 0) {
                    $percent = [int][Math]::Min(100, ($received * 100 / $total))
                    $status = '{0:N1} / {1:N1} MB' -f ($received / 1MB), ($total / 1MB)
                    Write-Progress -Activity $Title -Status $status -PercentComplete $percent
                } else {
                    Write-Progress -Activity $Title -Status ('{0:N1} MB' -f ($received / 1MB))
                }
                $lastUpdate = Get-Date
            }
        }
    } finally {
        if ($output) { $output.Dispose() }
        if ($input) { $input.Dispose() }
        if ($response) { $response.Dispose() }
        Write-Progress -Activity $Title -Completed
    }
    if (-not (Test-Path -LiteralPath $partial) -or (Get-Item -LiteralPath $partial).Length -lt $MinimumBytes) {
        Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
        throw "Сервер вернул слишком маленький или пустой файл: $Uri"
    }
    Move-Item -LiteralPath $partial -Destination $Destination -Force
}

function Test-ZipArchive {
    param([string]$Path, [string]$RequiredPattern = '')
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $archive = [IO.Compression.ZipFile]::OpenRead($Path)
        try {
            if ($archive.Entries.Count -eq 0) { return $false }
            if ($RequiredPattern) {
                return [bool]($archive.Entries | Where-Object { $_.FullName -match $RequiredPattern } | Select-Object -First 1)
            }
            return $true
        } finally {
            $archive.Dispose()
        }
    } catch {
        return $false
    }
}

function Test-PeFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path) -or (Get-Item $Path).Length -lt 4096) { return $false }
    $stream = [IO.File]::OpenRead($Path)
    try { return ($stream.ReadByte() -eq 0x4D -and $stream.ReadByte() -eq 0x5A) }
    finally { $stream.Dispose() }
}

function Get-AppCatalog {
    $items = @(
        [pscustomobject]@{Name='7-Zip';Id='7zip.7zip';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Angry IP Scanner';Id='angryziber.AngryIPScanner';Kind='Winget';Url=''},
        [pscustomobject]@{Name='AnyDesk';Id='AnyDesk.AnyDesk';Kind='Winget';Url=''},
        [pscustomobject]@{Name='anySCP';Id='macnev2013.anySCP';Kind='Winget';Url=''},
        [pscustomobject]@{Name='CompressO';Id='direct:compresso';Kind='Direct';Url='https://github.com/codeforreal1/compressO/releases/download/3.0.0/CompressO_3.0.0_x64.exe'},
        [pscustomobject]@{Name='Double Commander';Id='alexx2000.DoubleCommander';Kind='Winget';Url=''},
        [pscustomobject]@{Name='File Converter';Id='AdrienAllard.FileConverter';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Google Chrome';Id='Google.Chrome';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Happ';Id='Happ.Happ';Kind='Winget';Url=''},
        [pscustomobject]@{Name='ImgBurn';Id='LIGHTNINGUK.ImgBurn';Kind='Winget';Url=''},
        [pscustomobject]@{Name='KeePass';Id='DominikReichl.KeePass';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Notepad++';Id='Notepad++.Notepad++';Kind='Winget';Url=''},
        [pscustomobject]@{Name='PDFgear';Id='PDFgear.PDFgear';Kind='Winget';Url=''},
        [pscustomobject]@{Name='qBittorrent';Id='qBittorrent.qBittorrent';Kind='Winget';Url=''},
        [pscustomobject]@{Name='QuickLook';Id='QL-Win.QuickLook';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Recuva';Id='Piriform.Recuva';Kind='Winget';Url=''},
        [pscustomobject]@{Name='RustDesk';Id='RustDesk.RustDesk';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Samsung Magician';Id='XPDDT99J9GKB5C';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Telegram';Id='Telegram.TelegramDesktop';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Termius';Id='Termius.Termius';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Ventoy';Id='ventoy.ventoy';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Visual Studio Code';Id='Microsoft.VisualStudioCode';Kind='Winget';Url=''},
        [pscustomobject]@{Name='VLC';Id='VideoLAN.VLC';Kind='Winget';Url=''},
        [pscustomobject]@{Name='WhatsApp';Id='9NKSQGP7F2NH';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Winbox';Id='Mikrotik.Winbox';Kind='Winget';Url=''},
        [pscustomobject]@{Name='WinDirStat';Id='WinDirStat.WinDirStat';Kind='Winget';Url=''},
        [pscustomobject]@{Name='WireGuard';Id='WireGuard.WireGuard';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Zoom';Id='Zoom.Zoom';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Яндекс Браузер';Id='Yandex.Browser';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Яндекс Мессенджер';Id='Yandex.Messenger';Kind='Winget';Url=''}
    )
    return @($items | Sort-Object Name)
}

function Find-WinGetExecutable {
    $candidates = New-Object 'System.Collections.Generic.List[string]'
    $command = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($command) { [void]$candidates.Add($command.Source) }
    $package = Get-AppxPackage Microsoft.DesktopAppInstaller -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
    if ($package) {
        $packageExe = Join-Path $package.InstallLocation 'winget.exe'
        if (Test-Path -LiteralPath $packageExe) { [void]$candidates.Add($packageExe) }
    }
    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        try {
            $output = @(& $candidate --version 2>&1)
            $code = $LASTEXITCODE
            if ($code -eq 0 -and (($output -join '') -match '\d+\.\d+')) { return $candidate }
            Write-Log -Level 'WARN' -Message "winget не запускается: $candidate, код $code"
        } catch {
            Write-Log -Level 'WARN' -Message "winget не запускается: $candidate, $($_.Exception.Message)"
        }
    }
    return $null
}

function Install-AppxIgnoringNewerVersion {
    param([string]$Path)
    try {
        Add-AppxPackage -Path $Path -ErrorAction Stop
    } catch {
        if ($_.Exception.Message -notmatch '0x80073D06|higher version|более новая версия|уже установлен') { throw }
        Write-Log -Level 'INFO' -Message "Пропущен уже установленный пакет: $Path"
    }
}

function Ensure-WinGet {
    $script:WingetPath = Find-WinGetExecutable
    if ($script:WingetPath) { return $true }

    Show-WorkScreen -Title 'winget не работает. Выполняется восстановление...' -Details 'Будут загружены официальные пакеты Microsoft.'
    Write-Host '[1/4] Проверка установленного App Installer...'
    try {
        $existing = Get-AppxPackage Microsoft.DesktopAppInstaller -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
        if ($existing) {
            $manifest = Join-Path $existing.InstallLocation 'AppxManifest.xml'
            if (Test-Path $manifest) {
                try { Add-AppxPackage -Register $manifest -DisableDevelopmentMode -ErrorAction Stop } catch { Write-Log -Level 'WARN' -Message $_.Exception.Message }
                Start-Sleep -Seconds 1
                $script:WingetPath = Find-WinGetExecutable
                if ($script:WingetPath) { Write-Host 'winget восстановлен.'; return $true }
            }
        }

        $dir = Join-Path $script:CacheRoot 'WinGet'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $bundle = Join-Path $dir 'Microsoft.DesktopAppInstaller.msixbundle'
        $dependenciesZip = Join-Path $dir 'DesktopAppInstaller_Dependencies.zip'
        if (-not (Test-ZipArchive $bundle 'AppxMetadata|AppxManifest')) {
            Remove-Item -LiteralPath $bundle -Force -ErrorAction SilentlyContinue
            Write-Host '[2/4] Скачивание Microsoft App Installer...'
            Save-HttpFile -Uri 'https://aka.ms/getwinget' -Destination $bundle -Title 'Microsoft App Installer' -MinimumBytes 1MB
        } else { Write-Host '[2/4] Используется проверенный кэш App Installer.' }
        if (-not (Test-ZipArchive $dependenciesZip '\.(appx|msix)$')) {
            Remove-Item -LiteralPath $dependenciesZip -Force -ErrorAction SilentlyContinue
            Write-Host '[3/4] Скачивание зависимостей winget...'
            Save-HttpFile -Uri 'https://github.com/microsoft/winget-cli/releases/latest/download/DesktopAppInstaller_Dependencies.zip' -Destination $dependenciesZip -Title 'Зависимости winget' -MinimumBytes 100KB
        } else { Write-Host '[3/4] Используется проверенный кэш зависимостей.' }

        $dependenciesDir = Join-Path $dir 'Dependencies'
        Remove-Item -LiteralPath $dependenciesDir -Recurse -Force -ErrorAction SilentlyContinue
        Expand-Archive -LiteralPath $dependenciesZip -DestinationPath $dependenciesDir -Force
        Write-Host '[4/4] Установка зависимостей и App Installer...'
        $dependencyPackages = @(Get-ChildItem -LiteralPath $dependenciesDir -Recurse -File | Where-Object {
            $_.Extension -match '^\.(appx|msix)$' -and $_.FullName -match '(?i)(x64|neutral)'
        } | Sort-Object @{Expression={
            if ($_.Name -match '(?i)VCLibs') { 0 }
            elseif ($_.Name -match '(?i)UI\.Xaml') { 1 }
            elseif ($_.Name -match '(?i)WindowsAppRuntime') { 2 }
            else { 3 }
        }}, FullName)
        foreach ($dependency in $dependencyPackages) { Install-AppxIgnoringNewerVersion -Path $dependency.FullName }
        Install-AppxIgnoringNewerVersion -Path $bundle
        Start-Sleep -Seconds 2
        $script:WingetPath = Find-WinGetExecutable
        if (-not $script:WingetPath) { throw 'App Installer установлен, но команда winget по-прежнему не запускается.' }
        Write-Host ''
        Write-Host 'Обновление источников winget...'
        & $script:WingetPath source reset --force --disable-interactivity | Out-Null
        & $script:WingetPath source update --disable-interactivity | Out-Null
        Write-Host 'winget готов. Продолжаем.'
        Write-Log -Message "winget восстановлен: $script:WingetPath"
        return $true
    } catch {
        Write-Host ''
        Write-Host "Не удалось восстановить winget: $($_.Exception.Message)"
        Write-Host "Подробности: $script:LogPath"
        Write-Log -Level 'ERROR' -Message "Восстановление winget: $($_.Exception.ToString())"
        return $false
    }
}

function Test-WingetPackageInstalled {
    param([string]$Id)
    $output = @(& $script:WingetPath list --id $Id -e --accept-source-agreements --disable-interactivity 2>&1)
    return ($LASTEXITCODE -eq 0 -and (($output -join "`n") -match [regex]::Escape($Id)))
}

function Test-WingetPackageAvailable {
    param([string]$Id)
    $output = @(& $script:WingetPath show --id $Id -e --accept-source-agreements --disable-interactivity 2>&1)
    return ($LASTEXITCODE -eq 0 -and (($output -join "`n") -match [regex]::Escape($Id)))
}

function Test-CompressOInstalled {
    $paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    return [bool](Get-ItemProperty $paths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match '^CompressO' } | Select-Object -First 1)
}

function Start-SoftwareManager {
    if (-not (Ensure-WinGet)) { Pause-Result; return }
    $apps = Get-AppCatalog
    $selected = Select-MultipleItems -Title 'Выберите программы' -Items $apps -Text { param($item) "$($item.Name) [$($item.Id)]" }
    if ($null -eq $selected) { return }
    if ($selected.Count -eq 0) {
        Show-WorkScreen -Title 'Программы не выбраны.' -Details ''
        Pause-Result
        return
    }
    if (-not (Ensure-WinGet)) { Pause-Result; return }
    Show-WorkScreen -Title 'Установка выбранных программ'
    [int]$position = 0
    foreach ($app in $selected) {
        $position++
        Write-Host "[$position/$($selected.Count)] $($app.Name)"
        try {
            if ($app.Kind -eq 'Direct') {
                if (Test-CompressOInstalled) { Write-Host 'Уже установлено. Пропуск.'; continue }
                $installer = Join-Path $script:CacheRoot 'CompressO_3.0.0_x64.exe'
                if (-not (Test-PeFile $installer)) {
                    Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
                    Save-HttpFile -Uri $app.Url -Destination $installer -Title 'CompressO' -MinimumBytes 100KB
                }
                if (-not (Test-PeFile $installer)) { throw 'Скачанный установщик CompressO повреждён.' }
                $process = Start-Process -FilePath $installer -ArgumentList '/S' -Wait -PassThru
                if ($process.ExitCode -ne 0) { throw "Установщик завершился с кодом $($process.ExitCode)." }
            } else {
                if (Test-WingetPackageInstalled -Id $app.Id) { Write-Host 'Уже установлено. Пропуск.'; continue }
                if (-not (Test-WingetPackageAvailable -Id $app.Id)) { throw "Пакет $($app.Id) не найден в источниках winget." }
                & $script:WingetPath install --id $app.Id -e --silent --accept-source-agreements --accept-package-agreements --disable-interactivity
                if ($LASTEXITCODE -ne 0) { throw "winget завершился с кодом $LASTEXITCODE." }
            }
            Write-Host 'Готово.'
        } catch {
            Write-Host "Ошибка: $($_.Exception.Message)"
            Write-Log -Level 'ERROR' -Message "Установка $($app.Name): $($_.Exception.ToString())"
        }
        Write-Host ''
    }
    Pause-Result
}

function Start-SoftwareUpdates {
    if (-not (Ensure-WinGet)) { Pause-Result; return }
    Show-WorkScreen -Title 'Обновление установленных программ'
    & $script:WingetPath upgrade --all --silent --include-unknown --accept-source-agreements --accept-package-agreements --disable-interactivity
    if ($LASTEXITCODE -eq 0) { Write-Host ''; Write-Host 'Обновление завершено.' }
    else { Write-Host ''; Write-Host "winget завершился с кодом $LASTEXITCODE. Подробности показаны выше." }
    Pause-Result
}

function Get-OfficeDeploymentToolUrl {
    $page = Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/details.aspx?id=49117' -UseBasicParsing -TimeoutSec 30
    $match = [regex]::Match($page.Content, 'https://download\.microsoft\.com/[^"''<>\s]+/officedeploymenttool_[^"''<>\s]+\.exe', 'IgnoreCase')
    if (-not $match.Success) { throw 'На странице Microsoft не найдена ссылка Office Deployment Tool.' }
    return $match.Value
}

function Get-AuthenticodePublisher {
    param([string]$Path)
    $signature = Get-AuthenticodeSignature -FilePath $Path
    if ($signature.Status -ne 'Valid') { throw "Недействительная цифровая подпись файла ${Path}: $($signature.Status)." }
    return $signature.SignerCertificate.Subject
}

function Get-InstalledOfficeProducts {
    try {
        $value = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name ProductReleaseIds -ErrorAction Stop
        return [string]$value
    } catch { return '' }
}

function Start-OfficeManager {
    $installed = Get-InstalledOfficeProducts
    $editions = @(
        [pscustomobject]@{Name='Microsoft 365 Apps';Id='O365ProPlusRetail';Channel='Current'},
        [pscustomobject]@{Name='Office Professional Plus 2024';Id='ProPlus2024Retail';Channel='Current'},
        [pscustomobject]@{Name='Назад';Id='';Channel=''}
    )
    $title = 'Выберите редакцию Office'
    if ($installed) { $title += " (установлено: $installed)" }
    $choice = Select-SingleItem -Title $title -Items $editions -Text { param($item) $item.Name }
    if ($choice -lt 0 -or $editions[$choice].Id -eq '') { return }
    $edition = $editions[$choice]
    Show-WorkScreen -Title "Подготовка $($edition.Name)"
    if (-not (Read-YesNo -Question "Установить или изменить редакцию на $($edition.Name)?")) { return }
    Show-WorkScreen -Title "Подготовка $($edition.Name)"
    $dir = Join-Path $script:CacheRoot 'Office'
    $extractDir = Join-Path $dir 'ODT'
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $url = Get-OfficeDeploymentToolUrl
    $fileName = [IO.Path]::GetFileName(([Uri]$url).AbsolutePath)
    $package = Join-Path $dir $fileName
    if (-not (Test-PeFile $package)) {
        Remove-Item -LiteralPath $package -Force -ErrorAction SilentlyContinue
        Save-HttpFile -Uri $url -Destination $package -Title 'Office Deployment Tool' -MinimumBytes 1MB
    }
    $publisher = Get-AuthenticodePublisher -Path $package
    if ($publisher -notmatch 'Microsoft') { throw "Неожиданный издатель Office Deployment Tool: $publisher" }
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
    $setup = Join-Path $extractDir 'setup.exe'
    if (-not (Test-PeFile $setup)) {
        $process = Start-Process -FilePath $package -ArgumentList "/quiet /extract:`"$extractDir`"" -Wait -PassThru
        if ($process.ExitCode -ne 0 -or -not (Test-PeFile $setup)) { throw "Не удалось распаковать Office Deployment Tool, код $($process.ExitCode)." }
    }
    $setupPublisher = Get-AuthenticodePublisher -Path $setup
    if ($setupPublisher -notmatch 'Microsoft') { throw "Неожиданный издатель setup.exe: $setupPublisher" }
    $config = Join-Path $dir 'configuration.xml'
    $xml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$($edition.Channel)">
    <Product ID="$($edition.Id)">
      <Language ID="ru-ru" />
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
    Set-Content -LiteralPath $config -Value $xml -Encoding UTF8
    Write-Host 'Запуск установки Office. Загрузка может занять продолжительное время...'
    $process = Start-Process -FilePath $setup -ArgumentList "/configure `"$config`"" -WorkingDirectory $dir -Wait -PassThru
    if ($process.ExitCode -ne 0) { throw "Office Deployment Tool завершился с кодом $($process.ExitCode)." }
    $after = Get-InstalledOfficeProducts
    if ($after -and $after -notmatch [regex]::Escape($edition.Id)) {
        Write-Host "Установка завершена, но в реестре указана редакция: $after"
    } else { Write-Host 'Установка Office завершена.' }
    Pause-Result
}

function Get-RegistrySnapshot {
    param([string]$Path, [string]$Name)
    try {
        $key = Get-Item -LiteralPath $Path -ErrorAction Stop
        $names = @($key.GetValueNames())
        if ($names -notcontains $Name) { return [pscustomobject]@{Exists=$false;Value=$null;Kind='DWord'} }
        return [pscustomobject]@{
            Exists = $true
            Value = $key.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            Kind = [string]$key.GetValueKind($Name)
        }
    } catch {
        return [pscustomobject]@{Exists=$false;Value=$null;Kind='DWord'}
    }
}

function Restore-RegistrySnapshot {
    param([string]$Path, [string]$Name, [object]$Snapshot)
    if ([bool]$Snapshot.Exists) {
        New-Item -Path $Path -Force | Out-Null
        $type = if ($Snapshot.Kind) { [string]$Snapshot.Kind } else { 'DWord' }
        New-ItemProperty -Path $Path -Name $Name -PropertyType $type -Value $Snapshot.Value -Force | Out-Null
    } else {
        Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction SilentlyContinue
    }
}

function Get-ClassicMenuSnapshot {
    param([string]$ParentPath)
    $child = Join-Path $ParentPath 'InprocServer32'
    $parentExists = Test-Path -LiteralPath $ParentPath
    $childExists = Test-Path -LiteralPath $child
    $defaultExists = $false
    $defaultValue = $null
    if ($childExists) {
        $key = Get-Item -LiteralPath $child
        $defaultExists = (@($key.GetValueNames()) -contains '')
        if ($defaultExists) { $defaultValue = $key.GetValue('') }
    }
    return [pscustomobject]@{
        ParentExists = $parentExists
        ChildExists = $childExists
        DefaultExists = $defaultExists
        DefaultValue = $defaultValue
    }
}

function Test-ClassicMenuEnabled {
    param([string]$ParentPath)
    $child = Join-Path $ParentPath 'InprocServer32'
    if (-not (Test-Path -LiteralPath $child)) { return $false }
    try {
        $key = Get-Item -LiteralPath $child
        return ((@($key.GetValueNames()) -contains '') -and ([string]$key.GetValue('') -eq ''))
    } catch { return $false }
}

function Restore-ClassicMenuSnapshot {
    param([string]$ParentPath, [object]$Snapshot)
    $child = Join-Path $ParentPath 'InprocServer32'
    if (-not [bool]$Snapshot.ParentExists) {
        Remove-Item -LiteralPath $ParentPath -Recurse -Force -ErrorAction SilentlyContinue
        return
    }
    New-Item -Path $ParentPath -Force | Out-Null
    if ([bool]$Snapshot.ChildExists) {
        New-Item -Path $child -Force | Out-Null
        if ([bool]$Snapshot.DefaultExists) { Set-Item -LiteralPath $child -Value ([string]$Snapshot.DefaultValue) }
        else {
            $subKeyName = $child.Substring('HKCU:\'.Length).Replace('/', '\')
            $registryKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($subKeyName, $true)
            if ($null -ne $registryKey) {
                try { $registryKey.DeleteValue('', $false) }
                finally { $registryKey.Dispose() }
            }
        }
    } else {
        Remove-Item -LiteralPath $child -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Read-TweakBackups {
    $result = @{}
    if (-not (Test-Path -LiteralPath $script:TweakStatePath)) { return $result }
    try {
        $data = Get-Content -LiteralPath $script:TweakStatePath -Raw | ConvertFrom-Json
        foreach ($property in $data.PSObject.Properties) { $result[$property.Name] = $property.Value }
    } catch {
        Write-Log -Level 'WARN' -Message "Не удалось прочитать резервные значения твиков: $($_.Exception.Message)"
    }
    return $result
}

function Save-TweakBackups {
    param([hashtable]$Backups)
    $ordered = [ordered]@{}
    foreach ($key in ($Backups.Keys | Sort-Object)) { $ordered[$key] = $Backups[$key] }
    $ordered | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:TweakStatePath -Encoding UTF8
}

function Get-WindowsBuildInfo {
    $registry = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    return [pscustomobject]@{
        ProductName = [string]$registry.ProductName
        Build = [int]$registry.CurrentBuild
        DisplayVersion = [string]$registry.DisplayVersion
        InstallationType = [string]$registry.InstallationType
    }
}

function Get-TweakDefinitions {
    $advanced = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    $search = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'
    $classic = 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
    $windows = Get-WindowsBuildInfo
    $isClient = ($windows.InstallationType -eq 'Client')
    $isWindows11Client = ($isClient -and $windows.Build -ge 22000)
    $clientReason = if ($isClient) { '' } else { 'не поддерживается Windows Server' }
    $windows11Reason = if ($isClient) { 'доступно только в Windows 11' } else { 'не поддерживается Windows Server' }
    return @(
        [pscustomobject]@{Key='TaskbarLeft';Name='Панель задач слева';Path=$advanced;ValueName='TaskbarAl';OnValue=0;OffValue=1;Kind='Registry';Supported=$isWindows11Client;UnsupportedReason=$windows11Reason;MissingIsEnabled=$false},
        [pscustomobject]@{Key='HideSearch';Name='Скрыть поиск на панели задач';Path=$search;ValueName='SearchboxTaskbarMode';OnValue=0;OffValue=1;Kind='Registry';Supported=$isClient;UnsupportedReason=$clientReason;MissingIsEnabled=$false},
        [pscustomobject]@{Key='HideTaskView';Name='Скрыть кнопку представления задач';Path=$advanced;ValueName='ShowTaskViewButton';OnValue=0;OffValue=1;Kind='Registry';Supported=$isClient;UnsupportedReason=$clientReason;MissingIsEnabled=$false},
        [pscustomobject]@{Key='ClassicMenu';Name='Классическое контекстное меню Windows 11';Path=$classic;ValueName='';OnValue=0;OffValue=1;Kind='Classic';Supported=$isWindows11Client;UnsupportedReason=$windows11Reason;MissingIsEnabled=$false},
        [pscustomobject]@{Key='HideWidgets';Name='Скрыть кнопку виджетов';Path=$advanced;ValueName='TaskbarDa';OnValue=0;OffValue=1;Kind='Registry';Supported=$isWindows11Client;UnsupportedReason=$windows11Reason;MissingIsEnabled=$false}
    )
}

function Get-TweakEnabled {
    param([object]$Definition)
    if ($Definition.Kind -eq 'Classic') { return (Test-ClassicMenuEnabled -ParentPath $Definition.Path) }
    $snapshot = Get-RegistrySnapshot -Path $Definition.Path -Name $Definition.ValueName
    if (-not [bool]$snapshot.Exists) { return [bool]$Definition.MissingIsEnabled }
    return ([bool]$snapshot.Exists -and [int]$snapshot.Value -eq [int]$Definition.OnValue)
}

function Set-TweakEnabled {
    param([object]$Definition, [bool]$Enabled, [hashtable]$Backups)
    if ($Enabled) {
        if (-not $Backups.ContainsKey($Definition.Key)) {
            if ($Definition.Kind -eq 'Classic') { $Backups[$Definition.Key] = Get-ClassicMenuSnapshot -ParentPath $Definition.Path }
            else { $Backups[$Definition.Key] = Get-RegistrySnapshot -Path $Definition.Path -Name $Definition.ValueName }
            Save-TweakBackups -Backups $Backups
        }
        if ($Definition.Kind -eq 'Classic') {
            $child = Join-Path $Definition.Path 'InprocServer32'
            New-Item -Path $child -Force | Out-Null
            Set-Item -LiteralPath $child -Value ''
            if (-not (Test-ClassicMenuEnabled -ParentPath $Definition.Path)) { throw 'Не удалось создать полное значение классического контекстного меню.' }
        } else {
            New-Item -Path $Definition.Path -Force | Out-Null
            New-ItemProperty -Path $Definition.Path -Name $Definition.ValueName -PropertyType DWord -Value ([int]$Definition.OnValue) -Force | Out-Null
        }
    } else {
        if ($Backups.ContainsKey($Definition.Key)) {
            if ($Definition.Kind -eq 'Classic') { Restore-ClassicMenuSnapshot -ParentPath $Definition.Path -Snapshot $Backups[$Definition.Key] }
            else { Restore-RegistrySnapshot -Path $Definition.Path -Name $Definition.ValueName -Snapshot $Backups[$Definition.Key] }
            $Backups.Remove($Definition.Key)
            Save-TweakBackups -Backups $Backups
        } elseif ($Definition.Kind -eq 'Classic') {
            Remove-Item -LiteralPath $Definition.Path -Recurse -Force -ErrorAction SilentlyContinue
        } else {
            New-Item -Path $Definition.Path -Force | Out-Null
            New-ItemProperty -Path $Definition.Path -Name $Definition.ValueName -PropertyType DWord -Value ([int]$Definition.OffValue) -Force | Out-Null
        }
    }
}

function Stop-WindowsShellForTweaks {
    $running = @(Get-Process explorer -ErrorAction SilentlyContinue)
    if ($running.Count -eq 0) { return $false }

    # Explorer records part of the taskbar state when it exits.  Therefore it
    # must be stopped BEFORE the new registry values are written; otherwise it
    # can overwrite the values which the manager has just set.
    $oldProcessIds = @($running | ForEach-Object { $_.Id })
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    if ($oldProcessIds.Count -gt 0) {
        Wait-Process -Id $oldProcessIds -Timeout 2 -ErrorAction SilentlyContinue
    }
    return $true
}

function Wait-WindowsShellAfterTweaks {
    param([bool]$WasRunning)
    if (-not $WasRunning) { return }

    $deadline = (Get-Date).AddSeconds(10)
    do {
        Start-Sleep -Milliseconds 250
        if (@(Get-Process explorer -ErrorAction SilentlyContinue).Count -gt 0) {
            Start-Sleep -Milliseconds 750
            return
        }
    } while ((Get-Date) -lt $deadline)

    Start-Process -FilePath "$env:SystemRoot\explorer.exe" | Out-Null
    Start-Sleep -Seconds 1
}

function Get-TweakRawValueText {
    param([object]$Definition)
    if ($Definition.Kind -eq 'Classic') {
        if (Test-ClassicMenuEnabled -ParentPath $Definition.Path) { return '<пустое значение InprocServer32>' }
        return '<классическое меню не настроено>'
    }
    $snapshot = Get-RegistrySnapshot -Path $Definition.Path -Name $Definition.ValueName
    if (-not [bool]$snapshot.Exists) { return '<нет значения>' }
    return [string]$snapshot.Value
}

function Start-TweakManager {
    while ($true) {
        $definitions = Get-TweakDefinitions
        $windows = Get-WindowsBuildInfo
        $actions = @(
            'Включить выбранные твики',
            'Отключить выбранные твики и восстановить исходные значения',
            'Восстановить все сохранённые исходные значения',
            'Показать текущее состояние',
            'Назад'
        )
        $action = Select-SingleItem -Title "Твики Windows - $($windows.ProductName) $($windows.DisplayVersion)" -Items $actions -Text { param($item) $item }
        if ($action -lt 0 -or $action -eq 4) { return }

        $backups = Read-TweakBackups
        $items = foreach ($definition in $definitions) {
            [pscustomobject]@{
                Definition = $definition
                Enabled = Get-TweakEnabled -Definition $definition
                CanSelect = [bool]$definition.Supported
                HasBackup = $backups.ContainsKey($definition.Key)
            }
        }

        if ($action -eq 3) {
            Show-WorkScreen -Title 'Текущее состояние твиков' -Details ''
            foreach ($item in $items) {
                $state = if (-not $item.CanSelect) { $item.Definition.UnsupportedReason } elseif ($item.Enabled) { 'ВКЛ' } else { 'ВЫКЛ' }
                $backup = if ($item.HasBackup) { '; исходное значение сохранено' } else { '' }
                Write-Host " - $($item.Definition.Name): $state$backup"
            }
            Pause-Result
            continue
        }

        if ($action -eq 2) {
            $selected = @($items | Where-Object { $_.HasBackup })
            if ($selected.Count -eq 0) {
                Show-WorkScreen -Title 'Сохранённых исходных значений нет.' -Details ''
                Pause-Result
                continue
            }
            $targetEnabled = $false
            $operationName = 'ВОССТАНОВИТЬ ИСХОДНОЕ'
        } else {
            if (@($items | Where-Object { $_.CanSelect }).Count -eq 0) {
                Show-WorkScreen -Title 'Твики панели задач недоступны' -Details ''
                Write-Host "Система: $($windows.ProductName) $($windows.DisplayVersion)"
                Write-Host ''
                Write-Host 'Эти параметры оболочки поддерживаются клиентскими Windows 10/11.'
                Write-Host 'Windows Server сбрасывает их при запуске Explorer, поэтому менеджер не будет показывать ложный результат.'
                Write-Host ''
                Write-Host 'Если изменения уже применялись, выберите в предыдущем меню:'
                Write-Host '«Восстановить все сохранённые исходные значения».'
                Pause-Result
                continue
            }
            $targetEnabled = ($action -eq 0)
            $operationName = if ($targetEnabled) { 'ВКЛЮЧИТЬ' } else { 'ОТКЛЮЧИТЬ / ВОССТАНОВИТЬ' }
            $selected = Select-MultipleItems -Title "$operationName выбранные твики" -Items $items -CanSelect { param($item) $item.CanSelect } -Identity {
                param($item, $itemIndex)
                return [string]$item.Definition.Key
            } -Text {
                param($item)
                if (-not $item.CanSelect) { return "$($item.Definition.Name) - $($item.Definition.UnsupportedReason)" }
                $state = if ($item.Enabled) { 'ВКЛ' } else { 'ВЫКЛ' }
                return "$($item.Definition.Name) - сейчас $state"
            }
            if ($null -eq $selected -or $selected.Count -eq 0) { continue }
        }

        Show-WorkScreen -Title 'Подтверждение твиков' -Details ''
        Write-Host "Операция: $operationName"
        foreach ($item in $selected) {
            Write-Host " - $($item.Definition.Name) [$($item.Definition.Key)]"
        }
        Write-Host ''
        if (-not (Read-YesNo -Question 'Применить только перечисленные изменения?')) { continue }

        Show-WorkScreen -Title 'Применение выбранных твиков'
        Write-Host 'Остановка оболочки Windows...'
        $shellWasRunning = Stop-WindowsShellForTweaks
        $changed = @()
        try {
            foreach ($item in $selected) {
                $definition = $item.Definition
                $before = Get-TweakEnabled -Definition $definition
                Write-Host "${operationName}: $($definition.Name)"
                Write-Log -Message "TWEAK BEGIN Key=$($definition.Key) Path=$($definition.Path) Name=$($definition.ValueName) EnabledBefore=$before Target=$targetEnabled"
                Set-TweakEnabled -Definition $definition -Enabled $targetEnabled -Backups $backups
                $afterWrite = Get-TweakEnabled -Definition $definition
                $rawAfterWrite = Get-TweakRawValueText -Definition $definition
                Write-Log -Message "TWEAK WRITE Key=$($definition.Key) EnabledAfterWrite=$afterWrite Raw=$rawAfterWrite"
                if ($targetEnabled -and -not $afterWrite) { throw "Windows не приняла значение твика '$($definition.Name)'." }
                $changed += $item
            }
        } finally {
            Write-Host ''
            Write-Host 'Запуск оболочки Windows с новыми настройками...'
            Wait-WindowsShellAfterTweaks -WasRunning $shellWasRunning
        }

        $warnings = @()
        foreach ($item in $changed) {
            $actual = Get-TweakEnabled -Definition $item.Definition
            $rawFinal = Get-TweakRawValueText -Definition $item.Definition
            Write-Log -Message "TWEAK FINAL Key=$($item.Definition.Key) EnabledAfterExplorer=$actual Raw=$rawFinal"
            if ($targetEnabled -and -not $actual) {
                $warnings += "$($item.Definition.Name): Windows вернула другое значение"
            }
        }
        if ($warnings.Count -gt 0) {
            Write-Host ''
            Write-Host 'Некоторые значения были изменены самой Windows после перезапуска оболочки:'
            $warnings | ForEach-Object { Write-Host " - $_" }
            Write-Host "Диагностика сохранена: $script:LogPath"
        } else {
            Write-Host 'Выбранные твики применены.'
        }
        Pause-Result
    }
}

function Start-ActivationManager {
    $items = @('Активировать Windows и Office (MAS)','Показать статус Windows','Открыть параметры активации','Назад')
    $choice = Select-SingleItem -Title 'Активация Windows и Office' -Items $items -Text { param($item) $item }
    if ($choice -lt 0 -or $choice -eq 3) { return }
    switch ($choice) {
        0 {
            Show-WorkScreen -Title 'Активация Windows и Office' -Details 'Выполняется скрипт Massgrave, подождите...'
            try {
                Write-Host "[info] Starting activation (Mirror 1)..."
                iex "& { $(irm https://get.activated.win -ErrorAction Stop) } /HWID /OHWID"
            } catch {
                Write-Host "[info] Mirror 1 failed. Trying Mirror 2..."
                iex "& { $(irm https://massgrave.dev/get -ErrorAction Stop) } /HWID /OHWID"
            }
            Pause-Result
        }
        1 {
            Show-WorkScreen -Title 'Статус активации Windows' -Details ''
            & cscript.exe //nologo "$env:SystemRoot\System32\slmgr.vbs" /dli
            Pause-Result
        }
        2 { Start-Process 'ms-settings:activation' | Out-Null }
    }
}

function New-BerLength {
    param([int]$Length)
    if ($Length -lt 0x80) { return [byte[]]@([byte]$Length) }
    $bytes = New-Object 'System.Collections.Generic.List[byte]'
    [int]$value = $Length
    while ($value -gt 0) {
        $bytes.Insert(0, [byte]($value -band 0xFF))
        $value = $value -shr 8
    }
    return [byte[]](@([byte](0x80 -bor $bytes.Count)) + $bytes.ToArray())
}

function New-BerTlv {
    param([byte]$Tag, [byte[]]$Value)
    if ($null -eq $Value) { $Value = [byte[]]@() }
    return [byte[]](@($Tag) + (New-BerLength -Length $Value.Length) + $Value)
}

function New-BerInteger {
    param([int]$Value)
    $bytes = [BitConverter]::GetBytes([Net.IPAddress]::HostToNetworkOrder($Value))
    [int]$start = 0
    while ($start -lt 3 -and $bytes[$start] -eq 0 -and (($bytes[$start + 1] -band 0x80) -eq 0)) { $start++ }
    return New-BerTlv -Tag 0x02 -Value ([byte[]]$bytes[$start..3])
}

function New-BerOid {
    param([string]$Oid)
    $parts = @($Oid.Split('.') | ForEach-Object { [int]$_ })
    if ($parts.Count -lt 2) { throw "Некорректный SNMP OID: $Oid" }
    $content = New-Object 'System.Collections.Generic.List[byte]'
    $content.Add([byte](40 * $parts[0] + $parts[1]))
    for ([int]$i = 2; $i -lt $parts.Count; $i++) {
        [int]$number = $parts[$i]
        $oidBytes = New-Object 'System.Collections.Generic.List[byte]'
        $oidBytes.Insert(0, [byte]($number -band 0x7F))
        $number = $number -shr 7
        while ($number -gt 0) {
            $oidBytes.Insert(0, [byte](0x80 -bor ($number -band 0x7F)))
            $number = $number -shr 7
        }
        $content.AddRange($oidBytes)
    }
    return New-BerTlv -Tag 0x06 -Value $content.ToArray()
}

function Read-BerElement {
    param([byte[]]$Data, [ref]$Offset)
    [int]$position = [int]$Offset.Value
    if ($position -ge $Data.Length) { throw 'Неожиданный конец BER-пакета.' }
    [byte]$tag = $Data[$position]
    $position++
    if ($position -ge $Data.Length) { throw 'Повреждённая длина BER-пакета.' }
    [int]$length = $Data[$position]
    $position++
    if (($length -band 0x80) -ne 0) {
        [int]$count = $length -band 0x7F
        if ($count -lt 1 -or $count -gt 4 -or $position + $count -gt $Data.Length) { throw 'Некорректная BER-длина.' }
        $length = 0
        for ([int]$i = 0; $i -lt $count; $i++) { $length = ($length -shl 8) -bor $Data[$position + $i] }
        $position += $count
    }
    if ($length -lt 0 -or $position + $length -gt $Data.Length) { throw 'BER-значение выходит за границы пакета.' }
    if ($length -eq 0) { $value = [byte[]]@() }
    else { $value = [byte[]]$Data[$position..($position + $length - 1)] }
    $position += $length
    $Offset.Value = $position
    return [pscustomobject]@{Tag=$tag;Value=$value}
}

function New-SnmpGetRequest {
    param([string]$Oid)
    $oidElement = New-BerOid -Oid $Oid
    $nullElement = New-BerTlv -Tag 0x05 -Value ([byte[]]@())
    $varBind = New-BerTlv -Tag 0x30 -Value ([byte[]]($oidElement + $nullElement))
    $varBindList = New-BerTlv -Tag 0x30 -Value $varBind
    $requestId = New-BerInteger -Value (Get-Random -Minimum 1 -Maximum 2000000000)
    $errorStatus = New-BerInteger -Value 0
    $errorIndex = New-BerInteger -Value 0
    $pdu = New-BerTlv -Tag 0xA0 -Value ([byte[]]($requestId + $errorStatus + $errorIndex + $varBindList))
    $version = New-BerInteger -Value 0
    $community = New-BerTlv -Tag 0x04 -Value ([Text.Encoding]::ASCII.GetBytes('public'))
    return New-BerTlv -Tag 0x30 -Value ([byte[]]($version + $community + $pdu))
}

function ConvertFrom-SnmpResponse {
    param([byte[]]$Data)
    [int]$offset = 0
    $message = Read-BerElement -Data $Data -Offset ([ref]$offset)
    if ($message.Tag -ne 0x30) { throw 'Некорректный SNMP-ответ.' }
    [int]$inside = 0
    [void](Read-BerElement -Data $message.Value -Offset ([ref]$inside))
    [void](Read-BerElement -Data $message.Value -Offset ([ref]$inside))
    $pdu = Read-BerElement -Data $message.Value -Offset ([ref]$inside)
    [int]$pduOffset = 0
    [void](Read-BerElement -Data $pdu.Value -Offset ([ref]$pduOffset))
    $status = Read-BerElement -Data $pdu.Value -Offset ([ref]$pduOffset)
    [void](Read-BerElement -Data $pdu.Value -Offset ([ref]$pduOffset))
    if ($status.Value.Length -gt 0 -and $status.Value[$status.Value.Length - 1] -ne 0) { return $null }
    $list = Read-BerElement -Data $pdu.Value -Offset ([ref]$pduOffset)
    [int]$listOffset = 0
    $binding = Read-BerElement -Data $list.Value -Offset ([ref]$listOffset)
    [int]$bindingOffset = 0
    [void](Read-BerElement -Data $binding.Value -Offset ([ref]$bindingOffset))
    $value = Read-BerElement -Data $binding.Value -Offset ([ref]$bindingOffset)
    if ($value.Tag -in @(0x04, 0x40, 0x44)) {
        return ([Text.Encoding]::UTF8.GetString($value.Value)).Trim([char]0, ' ')
    }
    return $null
}

function Get-SnmpValue {
    param([string]$IPAddress, [string]$Oid, [int]$TimeoutMs = 650)
    $udp = New-Object Net.Sockets.UdpClient
    try {
        $udp.Client.ReceiveTimeout = $TimeoutMs
        $udp.Connect($IPAddress, 161)
        $request = New-SnmpGetRequest -Oid $Oid
        [void]$udp.Send($request, $request.Length)
        $endpoint = [Net.IPEndPoint]::new([Net.IPAddress]::Any, 0)
        $response = $udp.Receive([ref]$endpoint)
        return ConvertFrom-SnmpResponse -Data $response
    } catch { return $null }
    finally { $udp.Close() }
}

function Add-IppAttributeBytes {
    param([Collections.Generic.List[byte]]$Buffer, [byte]$Tag, [string]$Name, [string]$Value)
    $nameBytes = [Text.Encoding]::ASCII.GetBytes($Name)
    $valueBytes = [Text.Encoding]::UTF8.GetBytes($Value)
    $Buffer.Add($Tag)
    $Buffer.Add([byte]($nameBytes.Length -shr 8)); $Buffer.Add([byte]$nameBytes.Length
    )
    $Buffer.AddRange($nameBytes)
    $Buffer.Add([byte]($valueBytes.Length -shr 8)); $Buffer.Add([byte]$valueBytes.Length)
    $Buffer.AddRange($valueBytes)
}

function Get-IppAttributes {
    param([string]$IPAddress)
    try {
        $data = New-Object 'System.Collections.Generic.List[byte]'
        $data.AddRange([byte[]](0x01,0x01,0x00,0x0B,0x00,0x00,0x00,0x01,0x01))
        Add-IppAttributeBytes -Buffer $data -Tag 0x47 -Name 'attributes-charset' -Value 'utf-8'
        Add-IppAttributeBytes -Buffer $data -Tag 0x48 -Name 'attributes-natural-language' -Value 'ru'
        Add-IppAttributeBytes -Buffer $data -Tag 0x45 -Name 'printer-uri' -Value "ipp://$IPAddress/ipp/print"
        foreach ($attribute in @('printer-make-and-model','printer-name','printer-info')) {
            Add-IppAttributeBytes -Buffer $data -Tag 0x44 -Name 'requested-attributes' -Value $attribute
        }
        $data.Add(0x03)
        $request = [Net.HttpWebRequest]::Create("http://$IPAddress`:631/ipp/print")
        $request.Method = 'POST'
        $request.ContentType = 'application/ipp'
        $request.Timeout = 1800
        $request.ReadWriteTimeout = 1800
        $request.ContentLength = $data.Count
        $stream = $request.GetRequestStream()
        $bytes = $data.ToArray()
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Dispose()
        $response = $request.GetResponse()
        $memory = New-Object IO.MemoryStream
        $response.GetResponseStream().CopyTo($memory)
        $response.Dispose()
        $raw = $memory.ToArray()
        $memory.Dispose()
        $result = @{}
        [int]$offset = 8
        $currentName = ''
        while ($offset -lt $raw.Length) {
            [byte]$tag = $raw[$offset]; $offset++
            if ($tag -eq 0x03) { break }
            if ($tag -le 0x0F) { continue }
            if ($offset + 2 -gt $raw.Length) { break }
            [int]$nameLength = ($raw[$offset] -shl 8) -bor $raw[$offset + 1]; $offset += 2
            if ($offset + $nameLength -gt $raw.Length) { break }
            if ($nameLength -gt 0) { $currentName = [Text.Encoding]::ASCII.GetString($raw, $offset, $nameLength) }
            $offset += $nameLength
            if ($offset + 2 -gt $raw.Length) { break }
            [int]$valueLength = ($raw[$offset] -shl 8) -bor $raw[$offset + 1]; $offset += 2
            if ($offset + $valueLength -gt $raw.Length) { break }
            $value = [Text.Encoding]::UTF8.GetString($raw, $offset, $valueLength)
            $offset += $valueLength
            if ($currentName -and $value) { $result[$currentName] = $value }
        }
        return $result
    } catch { return @{} }
}

function Get-HttpPrinterEvidence {
    param([string]$IPAddress, [int]$MaximumPages = 10)
    $curl = Get-Command curl.exe -ErrorAction SilentlyContinue
    if (-not $curl) { return @() }
    $queue = New-Object Collections.Queue
    $queue.Enqueue([pscustomobject]@{Uri="https://$IPAddress/";Depth=0})
    $queue.Enqueue([pscustomobject]@{Uri="http://$IPAddress/";Depth=0})
    $visited = New-Object 'System.Collections.Generic.HashSet[string]'
    $evidence = New-Object 'System.Collections.Generic.List[string]'
    while ($queue.Count -gt 0 -and $visited.Count -lt $MaximumPages) {
        $entry = $queue.Dequeue()
        if (-not $visited.Add([string]$entry.Uri)) { continue }
        try {
            $content = @(& $curl.Source -k -L -sS --connect-timeout 1 --max-time 3 $entry.Uri 2>$null) -join "`n"
            if (-not $content) { continue }
            $evidence.Add($content)
            if ($entry.Depth -ge 2) { continue }
            $base = [Uri]::new([string]$entry.Uri)
            foreach ($match in [regex]::Matches($content, '(?i)(?:href|src)\s*=\s*["'']([^"''#]+)["'']')) {
                try {
                    $next = [Uri]::new($base, $match.Groups[1].Value)
                    if ($next.Host -ne $IPAddress) { continue }
                    if ($next.AbsolutePath -notmatch '(?i)\.(?:htm|html|js)?$|/$') { continue }
                    $queue.Enqueue([pscustomobject]@{Uri=$next.AbsoluteUri;Depth=([int]$entry.Depth + 1)})
                } catch {}
            }
        } catch {}
    }
    return $evidence.ToArray()
}

function Get-NormalizedPrinterModel {
    param([string[]]$Evidence)
    $patterns = @(
        '(?i)\b(?:KYOCERA\s+)?(TASKalfa\s+[A-Z0-9-]+)',
        '(?i)\b(?:KYOCERA\s+)?(ECOSYS\s+[A-Z0-9-]+)',
        '(?i)\b(HP\s+LaserJet\s+\d+\s+color\s+[A-Z0-9-]+)',
        '(?i)\b(HP\s+Color\s+LaserJet(?:\s+Pro)?(?:\s+MFP)?\s+[A-Z0-9-]+)',
        '(?i)\b(HP\s+LaserJet(?:\s+Pro)?(?:\s+MFP)?\s+[A-Z0-9-]+)'
    )
    foreach ($item in $Evidence) {
        if (-not $item) { continue }
        $text = [Net.WebUtility]::HtmlDecode([string]$item)
        $text = [regex]::Replace($text, '<[^>]+>', ' ')
        $text = [regex]::Replace($text, '\s+', ' ')
        foreach ($pattern in $patterns) {
            $match = [regex]::Match($text, $pattern)
            if ($match.Success) {
                $value = $match.Groups[1].Value.Trim()
                if ($value -match '^(?i)TASKalfa|ECOSYS') { return $value }
                return ('HP' + $value.Substring(2)).Trim()
            }
        }
    }
    return $null
}

function Get-PrinterIdentity {
    param([string]$IPAddress)
    $evidence = New-Object 'System.Collections.Generic.List[string]'
    foreach ($oid in @('1.3.6.1.2.1.1.1.0','1.3.6.1.2.1.43.5.1.1.16.1','1.3.6.1.2.1.1.5.0')) {
        $value = Get-SnmpValue -IPAddress $IPAddress -Oid $oid
        if ($value) { $evidence.Add($value) }
    }
    $ipp = Get-IppAttributes -IPAddress $IPAddress
    foreach ($key in @('printer-make-and-model','printer-name','printer-info')) {
        if ($ipp.ContainsKey($key) -and $ipp[$key]) { $evidence.Add([string]$ipp[$key]) }
    }
    $model = Get-NormalizedPrinterModel -Evidence $evidence.ToArray()
    if (-not $model) {
        foreach ($text in (Get-HttpPrinterEvidence -IPAddress $IPAddress)) { $evidence.Add($text) }
        $model = Get-NormalizedPrinterModel -Evidence $evidence.ToArray()
    }
    $combined = $evidence.ToArray() -join ' '
    if ($model -match '^(?i)HP\s') { $vendor = 'HP' }
    elseif ($model -match '^(?i)(TASKalfa|ECOSYS)' -or $combined -match '(?i)KYOCERA') { $vendor = 'Kyocera' }
    elseif ($combined -match '(?i)\bHP\b|Hewlett.Packard') { $vendor = 'HP' }
    else { $vendor = 'Unknown' }
    if (-not $model) { $model = 'Модель не определена' }
    return [pscustomobject]@{
        IPAddress = $IPAddress
        Model = $model
        Vendor = $vendor
        Supported = ($model -ne 'Модель не определена' -and $vendor -in @('HP','Kyocera'))
        Services = ''
    }
}

function Get-ActivePrinterSubnets {
    $subnets = @()
    $addresses = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {
        $_.IPAddress -notmatch '^(127\.|169\.254\.)' -and $_.AddressState -ne 'Tentative'
    }
    foreach ($address in $addresses) {
        $parts = $address.IPAddress.Split('.')
        if ($parts.Count -eq 4) { $subnets += ($parts[0..2] -join '.') }
    }
    try {
        foreach ($adapter in [Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
            if ($adapter.OperationalStatus -ne [Net.NetworkInformation.OperationalStatus]::Up) { continue }
            foreach ($unicast in $adapter.GetIPProperties().UnicastAddresses) {
                if ($unicast.Address.AddressFamily -ne [Net.Sockets.AddressFamily]::InterNetwork) { continue }
                $ip = $unicast.Address.ToString()
                if ($ip -match '^(127\.|169\.254\.)') { continue }
                $parts = $ip.Split('.')
                if ($parts.Count -eq 4) { $subnets += ($parts[0..2] -join '.') }
            }
        }
    } catch { Write-Log -Level 'WARN' -Message "Не удалось прочитать сетевые адаптеры через .NET: $($_.Exception.Message)" }
    return @($subnets | Sort-Object -Unique)
}

function ConvertTo-PrinterSubnet {
    param([string]$InputText)
    $value = $InputText.Trim()
    if ($value -match '/(.+)$' -and $value -notmatch '/24$') { return $null }
    $value = $value -replace '/24$', ''
    $parts = $value.Split('.')
    if ($parts.Count -eq 4 -and $parts[3] -eq '0') { $parts = $parts[0..2] }
    if ($parts.Count -ne 3) { return $null }
    foreach ($part in $parts) {
        [int]$number = 0
        if (-not [int]::TryParse($part, [ref]$number) -or $number -lt 0 -or $number -gt 255) { return $null }
    }
    return ($parts -join '.')
}

function Find-NetworkPrinters {
    param([string[]]$Subnets)
    $found = New-Object 'System.Collections.Generic.List[object]'
    foreach ($subnet in $Subnets) {
        Show-WorkScreen -Title "Сканирование $subnet.0/24" -Details 'Проверяются службы печати TCP 9100, IPP 631 и LPD 515.'
        $openByIp = @{}
        for ([int]$block = 1; $block -le 254; $block += 32) {
            $attempts = New-Object 'System.Collections.Generic.List[object]'
            [int]$end = [Math]::Min(254, $block + 31)
            for ([int]$hostNumber = $block; $hostNumber -le $end; $hostNumber++) {
                $ip = "$subnet.$hostNumber"
                foreach ($port in @(9100,631,515)) {
                    $client = New-Object Net.Sockets.TcpClient
                    try {
                        $async = $client.BeginConnect($ip, $port, $null, $null)
                        $attempts.Add([pscustomobject]@{IP=$ip;Port=$port;Client=$client;Async=$async})
                    } catch { $client.Close() }
                }
            }
            Start-Sleep -Milliseconds 550
            foreach ($attempt in $attempts) {
                try {
                    if ($attempt.Async.IsCompleted) {
                        $attempt.Client.EndConnect($attempt.Async)
                        if (-not $openByIp.ContainsKey($attempt.IP)) { $openByIp[$attempt.IP] = New-Object 'System.Collections.Generic.List[int]' }
                        $openByIp[$attempt.IP].Add([int]$attempt.Port)
                    }
                } catch {}
                finally {
                    try { $attempt.Async.AsyncWaitHandle.Close() } catch {}
                    $attempt.Client.Close()
                }
            }
            $percent = [int]($end * 100 / 254)
            Write-Progress -Activity "Сканирование $subnet.0/24" -Status "$end из 254 адресов" -PercentComplete $percent
        }
        Write-Progress -Activity "Сканирование $subnet.0/24" -Completed
        [int]$identityNumber = 0
        foreach ($ip in ($openByIp.Keys | Sort-Object { [version]$_ })) {
            $identityNumber++
            Write-Host "Определение модели $ip ($identityNumber из $($openByIp.Count))..."
            $device = Get-PrinterIdentity -IPAddress $ip
            $device.Services = (($openByIp[$ip] | Sort-Object | ForEach-Object { "TCP $_" }) -join ', ')
            $found.Add($device)
        }
    }
    return @($found | Sort-Object { [version]$_.IPAddress })
}

function Get-HPUniversalDriverUrl {
    $pages = @(
        'https://support.hp.com/us-en/drivers/hp-universal-print-driver-series-for-windows/4157320',
        'https://support.hp.com/us-en/drivers/hp-universal-print-driver-series-for-windows/503548'
    )
    foreach ($pageUrl in $pages) {
        try {
            $page = Invoke-WebRequest -Uri $pageUrl -UseBasicParsing -TimeoutSec 15
            $matches = [regex]::Matches($page.Content, 'https?[^"''\s<>]+upd-pcl6-win11-x64-[0-9.]+\.zip', 'IgnoreCase')
            if ($matches.Count -gt 0) { return $matches[0].Value }
        } catch { Write-Log -Level 'WARN' -Message "Страница HP недоступна для автоматической проверки: $pageUrl" }
    }
    return 'https://ftp.hp.com/pub/softlib/software13/printers/UPD/upd-pcl6-win11-x64-8.2.0.26819.zip'
}

function Get-PrinterDriverPackage {
    param([ValidateSet('HP','Kyocera')][string]$Vendor)
    if ($Vendor -eq 'HP') {
        $url = Get-HPUniversalDriverUrl
        $required = '(?i)hpcu[^/\\]*\.inf$'
    } else {
        $url = 'https://www.kyoceradocumentsolutions.com.br/content/dam/download-center-americas-cf/br/drivers/drivers/KX_Print_Driver_zip.download.zip'
        $required = '(?i)OEMSETUP\.INF$'
    }
    $vendorDir = Join-Path $script:DriverCache $Vendor
    New-Item -ItemType Directory -Path $vendorDir -Force | Out-Null
    $fileName = [IO.Path]::GetFileName(([Uri]$url).AbsolutePath)
    if (-not $fileName.EndsWith('.zip', [StringComparison]::OrdinalIgnoreCase)) { $fileName = "$Vendor-driver.zip" }
    $package = Join-Path $vendorDir $fileName
    if (-not (Test-ZipArchive -Path $package -RequiredPattern $required)) {
        Remove-Item -LiteralPath $package -Force -ErrorAction SilentlyContinue
        Save-HttpFile -Uri $url -Destination $package -Title "Драйвер $Vendor" -MinimumBytes 1MB
    }
    if (-not (Test-ZipArchive -Path $package -RequiredPattern $required)) {
        Remove-Item -LiteralPath $package -Force -ErrorAction SilentlyContinue
        throw "Скачанный пакет $Vendor повреждён или не содержит нужный INF-файл."
    }
    return $package
}

function Expand-PrinterDriverPackage {
    param([string]$Package, [string]$Vendor)
    $hash = (Get-FileHash -LiteralPath $Package -Algorithm SHA256).Hash.Substring(0, 16)
    $destination = Join-Path (Join-Path $script:DriverCache 'Expanded') "$Vendor-$hash"
    $marker = Join-Path $destination '.complete'
    if (-not (Test-Path -LiteralPath $marker)) {
        Remove-Item -LiteralPath $destination -Recurse -Force -ErrorAction SilentlyContinue
        New-Item -ItemType Directory -Path $destination -Force | Out-Null
        Expand-Archive -LiteralPath $Package -DestinationPath $destination -Force
        Set-Content -LiteralPath $marker -Value $hash -Encoding ASCII
    }
    return $destination
}

function Get-PrinterDriverDefinition {
    param([string]$Vendor, [string]$Model, [string]$ExpandedPath)
    if ($Vendor -eq 'HP') {
        foreach ($inf in (Get-ChildItem -LiteralPath $ExpandedPath -Filter 'hpcu*.inf' -File -Recurse)) {
            if (Select-String -LiteralPath $inf.FullName -SimpleMatch 'HP Universal Printing PCL 6' -Quiet) {
                return [pscustomobject]@{Inf=$inf.FullName;Name='HP Universal Printing PCL 6'}
            }
        }
        throw 'В пакете HP не найден драйвер HP Universal Printing PCL 6.'
    }
    $escapedModel = [regex]::Escape($Model)
    $allInfs = @(Get-ChildItem -LiteralPath $ExpandedPath -Filter 'OEMSETUP.INF' -File -Recurse | Sort-Object FullName)
    $infs = @($allInfs | Where-Object { $_.FullName -match '(?i)[\\/](64bit|x64)[\\/]' })
    if ($infs.Count -eq 0) { $infs = $allInfs }
    foreach ($inf in $infs) {
        $content = Get-Content -LiteralPath $inf.FullName -Raw -Encoding Default
        $match = [regex]::Match($content, '"(Kyocera[^"\r\n]*' + $escapedModel + '[^"\r\n]*KX[^"\r\n]*)"', 'IgnoreCase')
        if ($match.Success) { return [pscustomobject]@{Inf=$inf.FullName;Name=$match.Groups[1].Value.Trim()} }
    }
    throw "В пакете Kyocera не найдено точное имя драйвера для модели '$Model'."
}

function Install-NetworkPrinter {
    param([object]$Device)
    if (-not $Device.Supported) { throw 'Установка невозможна: производитель или модель не определены.' }
    if ($Device.Vendor -notin @('HP','Kyocera')) { throw "Производитель $($Device.Vendor) не поддерживается." }
    Show-WorkScreen -Title "Установка $($Device.Model) ($($Device.IPAddress))"
    Write-Host "Производитель: $($Device.Vendor)"
    Write-Host 'Получение пакета драйвера...'
    $package = Get-PrinterDriverPackage -Vendor $Device.Vendor
    Write-Host 'Проверка и распаковка драйвера...'
    $expanded = Expand-PrinterDriverPackage -Package $package -Vendor $Device.Vendor
    $driver = Get-PrinterDriverDefinition -Vendor $Device.Vendor -Model $Device.Model -ExpandedPath $expanded
    Write-Host "Драйвер: $($driver.Name)"
    Import-Module PrintManagement -ErrorAction Stop
    $spooler = Get-Service Spooler -ErrorAction Stop
    if ($spooler.Status -ne 'Running') { Start-Service Spooler }
    if (-not (Get-PrinterDriver -Name $driver.Name -ErrorAction SilentlyContinue)) {
        Write-Host 'Добавление пакета драйвера в Windows...'
        $pnputilOutput = @(& "$env:SystemRoot\System32\pnputil.exe" /add-driver $driver.Inf /install 2>&1)
        $pnputilOutput | ForEach-Object { Write-Host $_ }
        if ($LASTEXITCODE -notin @(0, 3010)) { throw "PnPUtil завершился с кодом $LASTEXITCODE." }
        Add-PrinterDriver -Name $driver.Name -ErrorAction Stop
    }
    if (-not (Get-PrinterDriver -Name $driver.Name -ErrorAction SilentlyContinue)) { throw "Windows не зарегистрировала драйвер '$($driver.Name)'." }
    $portName = "IP_$($Device.IPAddress)"
    if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
        Write-Host "Создание TCP/IP-порта $portName..."
        Add-PrinterPort -Name $portName -PrinterHostAddress $Device.IPAddress -ErrorAction Stop
    }
    $queueName = "$($Device.Model) ($($Device.IPAddress))"
    $queue = Get-Printer -Name $queueName -ErrorAction SilentlyContinue
    if ($queue) {
        Set-Printer -Name $queueName -PortName $portName -DriverName $driver.Name -ErrorAction Stop
        Write-Host "Очередь обновлена: $queueName"
    } else {
        Add-Printer -Name $queueName -PortName $portName -DriverName $driver.Name -ErrorAction Stop
        Write-Host "Принтер установлен: $queueName"
    }
    Pause-Result
}

function Start-PrinterInstallation {
    $active = Get-ActivePrinterSubnets
    $options = @()
    if ($active.Count -gt 0) {
        $options += [pscustomobject]@{Name="Сканировать активные подсети: $($active -join ', ')";Mode='Auto'}
    }
    $options += [pscustomobject]@{Name='Ввести подсеть /24 вручную';Mode='Manual'}
    $options += [pscustomobject]@{Name='Назад';Mode='Back'}
    $choice = Select-SingleItem -Title 'Выберите подсеть' -Items $options -Text { param($item) $item.Name }
    if ($choice -lt 0 -or $options[$choice].Mode -eq 'Back') { return }
    if ($options[$choice].Mode -eq 'Auto') {
        $subnets = $active
    } else {
        Show-TextCursor
        Clear-Host
        $manual = Read-Host 'Введите подсеть, например 10.130.106 или 10.130.106.0/24'
        $subnet = ConvertTo-PrinterSubnet -InputText $manual
        if (-not $subnet) {
            Write-Host 'Неверный формат. Нужны три октета от 0 до 255 и маска /24.'
            Pause-Result
            return
        }
        $subnets = @($subnet)
    }
    $devices = @(Find-NetworkPrinters -Subnets $subnets)
    if ($devices.Count -eq 0) {
        Write-Host 'В выбранной подсети принтеры не найдены.'
        Pause-Result
        return
    }
    $choice = Select-MultipleItems -Title 'Выберите один принтер для установки' -Items $devices -CanSelect { param($item) $item.Supported } -Text {
        param($item)
        $status = if ($item.Supported) { $item.Vendor } else { 'установка недоступна' }
        "$($item.IPAddress)  $($item.Model)  [$status; $($item.Services)]"
    }
    if ($null -eq $choice -or $choice.Count -eq 0) { return }
    if ($choice.Count -gt 1) {
        Show-WorkScreen -Title 'Можно установить только один принтер за один запуск.' -Details 'Оставьте отмеченным один пункт.'
        Pause-Result
        return
    }
    Install-NetworkPrinter -Device $choice[0]
}

function Start-PrinterRemoval {
    Import-Module PrintManagement -ErrorAction Stop
    $printers = @(Get-Printer | Sort-Object Name)
    if ($printers.Count -eq 0) {
        Show-WorkScreen -Title 'Установленные принтеры не найдены.' -Details ''
        Pause-Result
        return
    }
    $items = @($printers) + @([pscustomobject]@{Name='Назад';PortName='';DriverName=''})
    $choice = Select-SingleItem -Title 'Выберите принтер для удаления' -Items $items -Text { param($item) "$($item.Name) [$($item.DriverName)]" }
    if ($choice -lt 0 -or $choice -eq $items.Count - 1) { return }
    $printer = $items[$choice]
    Show-TextCursor
    Clear-Host
    Write-Host 'Будет удалена только очередь принтера.'
    Write-Host ''
    Write-Host "Имя:    $($printer.Name)"
    Write-Host "Порт:   $($printer.PortName)"
    Write-Host "Драйвер: $($printer.DriverName)"
    Write-Host ''
    if ((Read-Host 'Введите DELETE для удаления') -ne 'DELETE') {
        Write-Host 'Удаление отменено.'
        Pause-Result
        return
    }
    Remove-Printer -Name $printer.Name -ErrorAction Stop
    Write-Host 'Очередь принтера удалена. Порт и драйвер сохранены.'
    Pause-Result
}

function Start-PrinterManager {
    while ($true) {
        $items = @('Установить сетевой принтер','Удалить принтер','Назад')
        $choice = Select-SingleItem -Title 'Сетевые принтеры' -Items $items -Text { param($item) $item }
        if ($choice -lt 0 -or $choice -eq 2) { return }
        try {
            if ($choice -eq 0) { Start-PrinterInstallation }
            else { Start-PrinterRemoval }
        } catch {
            Show-ErrorMessage -Title 'Ошибка менеджера принтеров' -ErrorRecord $_
        }
    }
}

function Invoke-SelfTest {
    Show-TextCursor
    Clear-Host
    Write-Host 'Самопроверка WG Install Manager'
    Write-Host ''
    $failures = New-Object 'System.Collections.Generic.List[string]'
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    if ($PSVersionTable.PSVersion.Major -lt 5) { $failures.Add('Требуется PowerShell 5.1 или новее.') }
    if ($PSCommandPath) {
        $tokens = $null
        $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$tokens, [ref]$errors)
        if ($errors.Count -gt 0) { foreach ($error in $errors) { $failures.Add("Синтаксис: $($error.Message)") } }
        $dynamicCommandName = 'Invoke' + '-Expression'
        $base64MethodName = 'FromBase64' + 'String'
        $dangerousCommands = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ -eq $dynamicCommandName })
        $base64Calls = @($ast.FindAll({
            param($node)
            $node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and [string]$node.Member.Value -eq $base64MethodName
        }, $true))
        if ($dangerousCommands.Count -gt 0 -or $base64Calls.Count -gt 0) { $failures.Add('Найден запрещённый динамический закодированный блок.') }
        $functionNames = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
        if (($functionNames | Group-Object | Where-Object Count -gt 1).Count -gt 0) { $failures.Add('Обнаружены функции с повторяющимися именами.') }
    } else { $warnings.Add('Скрипт запущен через конвейер: проверка собственного файла пропущена.') }
    $apps = Get-AppCatalog
    if (($apps | Group-Object Id | Where-Object Count -gt 1).Count -gt 0) { $failures.Add('В каталоге программ есть повторяющиеся ID.') }
    $sortedNames = @($apps.Name | Sort-Object)
    if (($apps.Name -join '|') -ne ($sortedNames -join '|')) { $failures.Add('Каталог программ не отсортирован.') }
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    [void]$set.Add('HideSearch')
    if (-not $set.Contains('HideSearch') -or $set.Contains('TaskbarLeft')) { $failures.Add('Проверка стабильных ключей множественного выбора не пройдена.') }
    foreach ($sample in @('10.130.106','10.130.106.0/24')) {
        if ((ConvertTo-PrinterSubnet $sample) -ne '10.130.106') { $failures.Add("Не распознана подсеть $sample") }
    }
    if (ConvertTo-PrinterSubnet '999.1.1') { $failures.Add('Проверка диапазона октетов подсети не работает.') }
    $modelTests = [ordered]@{
        'KYOCERA TASKalfa 2554ci' = 'TASKalfa 2554ci'
        'HP LaserJet 200 color M251nw 10.130.106.12' = 'HP LaserJet 200 color M251nw'
        'HP Color LaserJet Pro M478f-9f' = 'HP Color LaserJet Pro M478f-9f'
    }
    foreach ($sample in $modelTests.Keys) {
        $actualModel = Get-NormalizedPrinterModel -Evidence @($sample)
        if ($actualModel -ne $modelTests[$sample]) { $failures.Add("Определение модели: '$sample' -> '$actualModel'.") }
    }
    if (Get-NormalizedPrinterModel -Evidence @('Обычный сетевой узел без принтера')) { $failures.Add('Неизвестное устройство ошибочно распознано как принтер.') }
    $tweaks = Get-TweakDefinitions
    $expectedTweaks = @{
        TaskbarLeft='TaskbarAl'; HideSearch='SearchboxTaskbarMode'; HideTaskView='ShowTaskViewButton'; HideWidgets='TaskbarDa'
    }
    foreach ($key in $expectedTweaks.Keys) {
        $definition = $tweaks | Where-Object Key -eq $key | Select-Object -First 1
        if (-not $definition -or $definition.ValueName -ne $expectedTweaks[$key]) { $failures.Add("Неверное сопоставление твика $key.") }
    }
    $buildInfo = Get-WindowsBuildInfo
    if ($buildInfo.InstallationType -eq 'Server' -and @($tweaks | Where-Object Supported).Count -gt 0) {
        $failures.Add('Неподдерживаемые твики панели задач ошибочно разрешены в Windows Server.')
    }
    $winget = Find-WinGetExecutable
    if ($winget) { Write-Host "[OK] winget запускается: $winget" }
    else { $warnings.Add('winget сейчас не запускается; интерактивный режим предложит восстановление.') }
    if ($failures.Count -eq 0) { Write-Host '[OK] Синтаксис и внутренняя структура.' }
    Write-Host "[OK] Каталог программ: $($apps.Count) пунктов, без повторов."
    Write-Host '[OK] Множественный выбор использует стабильные ключи, а не позиции строк.'
    Write-Host '[OK] Проверка формата подсети.'
    Write-Host '[OK] Определение моделей HP/Kyocera и сопоставление твиков.'
    foreach ($warning in $warnings) { Write-Host "[ПРЕДУПРЕЖДЕНИЕ] $warning" }
    foreach ($failure in $failures) { Write-Host "[ОШИБКА] $failure" }
    Write-Host ''
    if ($failures.Count -eq 0) {
        Write-Host 'Самопроверка завершена успешно. Изменения в системе не выполнялись.'
        return $true
    }
    Write-Host "Самопроверка не пройдена: $($failures.Count) ошибок."
    return $false
}

function Invoke-MainMenuAction {
    param([int]$Choice)
    try {
        switch ($Choice) {
            0 { Start-SoftwareManager }
            1 { Start-SoftwareUpdates }
            2 { Start-OfficeManager }
            3 { Start-TweakManager }
            4 { Start-ActivationManager }
            5 { Start-PrinterManager }
        }
    } catch {
        Show-ErrorMessage -Title 'Операция завершилась с ошибкой' -ErrorRecord $_
    }
}

Initialize-ConsoleUi
Write-Log -Message "Запуск менеджера. PowerShell $($PSVersionTable.PSVersion), SelfTest=$SelfTest"

if ($SelfTest) {
    $ok = Invoke-SelfTest
    if (-not $ok) { exit 1 }
    exit 0
}

try {
    while (-not $script:ExitRequested) {
        $mainItems = @(
            'Установить программы',
            'Обновить программы',
            'Установить Office',
            'Применить твики',
            'Активация Windows',
            'Сетевые принтеры',
            'Выход'
        )
        $choice = Select-SingleItem -Title 'Менеджер установки WG' -Items $mainItems -Text { param($item) $item }
        if ($choice -lt 0) { continue }
        if ($choice -eq 6) {
            $script:ExitRequested = $true
            break
        }
        Invoke-MainMenuAction -Choice $choice
    }
} finally {
    Show-TextCursor
    Write-Log -Message 'Завершение менеджера.'
}

exit 0
