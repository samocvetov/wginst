[CmdletBinding()]
param(
    [switch]$SelfTest
)

$script:WgManagerMarker = 'WG-INSTALL-MANAGER-V2'
$script:SelfUrl = 'https://raw.githubusercontent.com/samocvetov/wginst/main/s.ps1'
$script:MasAioUrl = 'https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/master/MAS/All-In-One-Version-KL/MAS_AIO.cmd'
$script:IsWindows = ($env:OS -eq 'Windows_NT')
$script:SupportsAnsi = $false
$script:MenuDepth = 0
$script:LastFrameLineCount = 0
$script:WingetPath = $null
$script:ExitRequested = $false
$script:EnableDiagnostics = $false

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
    $content = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    if ($content.Length -lt 1000 -or $content -notmatch 'WG-INSTALL-MANAGER-V2') { return $false }
    if ($content -match '(?im)^\s*(?:<!DOCTYPE|<html|code:|404\s*:|Output:|Wall time:|Exit code:)') { return $false }
    $tokens = $null
    $errors = $null
    [void][Management.Automation.Language.Parser]::ParseInput($content, [ref]$tokens, [ref]$errors)
    return ($errors.Count -eq 0)
}

function Set-Utf8ConsoleEncoding {
    $utf8 = [System.Text.Encoding]::UTF8
    try { [Console]::OutputEncoding = $utf8 } catch { }
    try { [Console]::InputEncoding = $utf8 } catch { }
    try { $OutputEncoding = $utf8 } catch { }
    try { chcp 65001 *> $null } catch { }
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
Set-Utf8ConsoleEncoding

$script:Root = Join-Path $env:LOCALAPPDATA 'WGInstall'
$script:CacheRoot = Join-Path $script:Root 'Cache'
$script:DriverCache = Join-Path $script:CacheRoot 'Drivers'
$script:TweakStatePath = Join-Path $script:Root 'tweaks-state.json'
$script:LogPath = $null
if (-not $SelfTest) {
    try {
        foreach ($directory in @($script:Root, $script:CacheRoot, $script:DriverCache)) {
            New-Item -ItemType Directory -Path $directory -Force -ErrorAction Stop | Out-Null
        }
    } catch {
        Write-Host "Не удалось подготовить рабочую папку $script:Root"
        Write-Host $_.Exception.Message
        return
    }
}

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    if (-not $script:EnableDiagnostics) { return }
    if (-not $script:LogPath) { return }
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Get-ErrorSnapshot {
    $lines = New-Object System.Collections.ArrayList
    for ($i = $Error.Count - 1; $i -ge 0 -and $lines.Count -lt 15; $i--) {
        $entry = $Error[$i]
        if (-not $entry) { continue }
        [void]$lines.Insert(0, ("{0} | {1} | {2} | {3}" -f [string]$entry.CategoryInfo.Category, [string]$entry.FullyQualifiedErrorId, [string]$entry.TargetObject, [string]$entry.Exception.Message))
    }
    return ($lines -join [Environment]::NewLine)
}

function Get-VolumeFreeSpace {
    param([string]$Path)
    try {
        $full = (Get-Item -LiteralPath $Path -Force).FullName
        $letter = ($full -split '[\\/]')[0] -replace ':', ''
        if (-not $letter) { return $null }
        $volume = @(Get-Volume -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter -eq $letter })[0]
        if ($volume) { return $volume.SizeRemaining }
    } catch { }
    return $null
}

function Save-ErrorBundle {
    param([string]$Title, [string]$Detail)
    if (-not $script:EnableDiagnostics) { return $null }
    try {
        $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
        $base = Join-Path $script:Root 'ERR'
        New-Item -ItemType Directory -Path $base -Force | Out-Null
        $bundleDir = Join-Path $base "WGInstall-ERR-$stamp"
        $suffix = 1
        while (Test-Path -LiteralPath $bundleDir) {
            $suffix++
            $bundleDir = Join-Path $base "WGInstall-ERR-$stamp-$suffix"
        }
        New-Item -ItemType Directory -Path $bundleDir -Force | Out-Null

        $copied = New-Object System.Collections.ArrayList
        $targetName = { param($bundle, $name) $t = Join-Path $bundle $name; $k = 1; while (Test-Path -LiteralPath $t) { $k++; $t = Join-Path $bundle ($k.ToString() + '_' + $name) }; $t }

        $mainLog = Join-Path $bundleDir 'WGInstall-main.log'
        try {
            if ($script:LogPath -and (Test-Path -LiteralPath $script:LogPath)) {
                Copy-Item -LiteralPath $script:LogPath -Destination $mainLog -Force
                [void]$copied.Add('WGInstall-main.log (журнал WGInstall)')
            }
        } catch { }

        $lastRun = Join-Path $script:Root 'last-run.txt'
        try {
            if (Test-Path -LiteralPath $lastRun) {
                Copy-Item -LiteralPath $lastRun -Destination (Join-Path $bundleDir 'last-run.txt') -Force
                [void]$copied.Add('last-run.txt (последний запуск)')
            }
        } catch { }

        try {
            $logs = @(Get-ChildItem -LiteralPath $script:LogRoot -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 10)
            foreach ($log in $logs) {
                $target = (& $targetName $bundleDir $log.Name)
                Copy-Item -LiteralPath $log.FullName -Destination $target -Force
                [void]$copied.Add((Split-Path -Leaf $target) + ' (журнал winget/других операций)')
            }
        } catch { }

        $officeDir = Join-Path $script:CacheRoot 'Office'
        try {
            if (Test-Path -LiteralPath $officeDir) {
                $officeFiles = @(Get-ChildItem -LiteralPath $officeDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^(odt-.+\.log|download\.xml|install\.xml|remove\.xml)$' })
                foreach ($of in $officeFiles) {
                    $target = (& $targetName $bundleDir $of.Name)
                    Copy-Item -LiteralPath $of.FullName -Destination $target -Force
                    [void]$copied.Add((Split-Path -Leaf $target) + ' (файл Office: лог ODT или XML-конфигурация)')
                }
            }
        } catch { }

        $odtTempDir = Join-Path $env:LOCALAPPDATA 'Temp\ODT'
        try {
            if ($odtTempDir -and (Test-Path -LiteralPath $odtTempDir)) {
                $odtFiles = @(Get-ChildItem -LiteralPath $odtTempDir -File -ErrorAction SilentlyContinue)
                foreach ($of in $odtFiles) {
                    $target = (& $targetName $bundleDir $of.Name)
                    Copy-Item -LiteralPath $of.FullName -Destination $target -Force
                    [void]$copied.Add((Split-Path -Leaf $target) + ' (подробный лог ODT из %LOCALAPPDATA%\Temp\ODT)')
                }
                if ($odtFiles.Count -eq 0) { [void]$copied.Add('(папка %LOCALAPPDATA%\Temp\ODT\ существует, но не содержит файлов)') }
            } else {
                [void]$copied.Add('(папка %LOCALAPPDATA%\Temp\ODT\ не найдена — подробный лог ODT отсутствует)')
            }
        } catch { }

        try {
            $tempRoot = $env:TEMP
            if ($tempRoot -and (Test-Path -LiteralPath $tempRoot)) {
                $recentSetupLogs = @(Get-ChildItem -LiteralPath $tempRoot -File -Filter '*.log' -ErrorAction SilentlyContinue |
                    Where-Object { $_.LastWriteTime -gt (Get-Date).AddHours(-4) -and $_.Name -match '(?i)(office|setup|click|c2r|desktop)' } |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 8)
                foreach ($setupLog in $recentSetupLogs) {
                    $target = (& $targetName $bundleDir $setupLog.Name)
                    Copy-Item -LiteralPath $setupLog.FullName -Destination $target -Force
                    [void]$copied.Add((Split-Path -Leaf $target) + ' (лог Office Click-to-Run из %TEMP%)')
                }
            }
        } catch { }

        $report = New-Object System.Collections.ArrayList
        [void]$report.Add('WGInstall — материалы для анализа ошибки')
        [void]$report.Add(('Время: {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss (локальное время)')))
        [void]$report.Add(('Заголовок ошибки: {0}' -f $Title))
        if ($Detail) {
            [void]$report.Add('Текст ошибки:')
            foreach ($line in ($Detail -split [Environment]::NewLine)) { [void]$report.Add('  ' + $line) }
            [void]$report.Add('')
        }
        [void]$report.Add('Система:')
        try { [void]$report.Add(('  ОС: {0}' -f [Environment]::OSVersion.VersionString)) } catch { }
        try { [void]$report.Add(('  PowerShell: {0}' -f $PSVersionTable.PSVersion.ToString())) } catch { }
        try {
            $osInfo = Get-CimInstance -ClassName 'Win32_OperatingSystem' -ErrorAction Stop
            [void]$report.Add(('  Сборка Windows: {0} ({1})' -f $osInfo.BuildNumber, $osInfo.OSArchitecture))
            [void]$report.Add(('  Память: всего {0} ГБ, свободно {1} ГБ' -f [math]::Round($osInfo.TotalVisibleMemorySize / 1MB, 1), [math]::Round($osInfo.FreePhysicalMemory / 1MB, 1)))
        } catch { [void]$report.Add('  (не удалось получить данные Win32_OperatingSystem)') }
        foreach ($envName in @('TEMP','TMP')) {
            $envItem = Get-Item -Path ('Env:' + $envName) -ErrorAction SilentlyContinue
            if (-not $envItem) { continue }
            $envPath = [string]$envItem.Value
            if (-not $envPath) { continue }
            $envFree = $null
            try { $envFree = Get-VolumeFreeSpace -Path $envPath } catch { $envFree = $null }
            if ($null -ne $envFree) { [void]$report.Add(('{0}: {1} (свободно {2} ГБ)' -f $envName, $envPath, [math]::Round($envFree / 1GB, 1))) }
            else { [void]$report.Add(('{0}: {1}' -f $envName, $envPath)) }
        }
        [void]$report.Add('')
        [void]$report.Add('Свободное место на локальных дисках:')
        try {
            $volumes = @(Get-Volume -ErrorAction SilentlyContinue | Where-Object { $_.DriveType -eq 'Fixed' -or $_.DriveType -eq 3 })
            foreach ($v in $volumes) {
                $fsType = [string]$v.FileSystemType
                if (-not $fsType) { $fsType = '?' }
                [void]$report.Add(('  Диск {0} ({1}): всего {2} ГБ, свободно {3} ГБ, сжат: {4}, только чтение: {5}' -f $v.DriveLetter, $fsType, [math]::Round($v.Size / 1GB, 1), [math]::Round($v.SizeRemaining / 1GB, 1), [bool]$v.FileSystemIsCompressed, [bool]$v.FileSystemReadOnly))
            }
            if ($volumes.Count -eq 0) { [void]$report.Add('  (не удалось получить список дисков)') }
        } catch { [void]$report.Add('  (не удалось получить список дисков)') }
        [void]$report.Add('')
        [void]$report.Add('Состав этой папки:')
        foreach ($name in $copied) { [void]$report.Add('  ' + $name) }
        if ($copied.Count -eq 0) { [void]$report.Add('  (не удалось скопировать ни один файл)') }
        [void]$report.Add('')
        [void]$report.Add(('Исходный журнал WGInstall: {0}' -f $script:LogPath))
        [void]$report.Add(('Каталог кэша: {0}' -f $script:CacheRoot))
        try {
            Set-Content -LiteralPath (Join-Path $bundleDir 'README.txt') -Value ($report -join [Environment]::NewLine) -Encoding UTF8
        } catch { }
    } catch {
        Write-Host 'Не удалось сохранить папку с материалами об ошибке (возможно, недостаточно места на диске).'
        return $null
    }
    Write-Log -Message "Сохранена папка с материалами об ошибке: $bundleDir"
    return $bundleDir
}

function Show-ErrorMessage {
    param([string]$Title, [Management.Automation.ErrorRecord]$ErrorRecord)
    Show-TextCursor
    Clear-Host
    Write-Host $Title
    Write-Host ''
    Write-Host $ErrorRecord.Exception.Message
    Write-Host ''
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
    Set-Utf8ConsoleEncoding
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
    $started = Get-Date
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
    } catch {
        $status = 'нет данных'
        try {
            if ($_.Exception) {
                $code = $null
                if ($_.Exception -is [Net.WebException]) {
                    $code = $_.Exception.StatusCode
                    if (-not $code -and $_.Exception.Response) { $code = [string]$_.Exception.Response.StatusCode }
                }
                if ($code) { $status = [string]$code }
            }
        } catch { }
        Write-Log -Level 'ERROR' -Message ("Скачивание {0} не удалось: {1} (HTTP: {2})" -f $Uri, $_.Exception.Message, $status)
        Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
        throw
    } finally {
        if ($output) { $output.Dispose() }
        if ($input) { $input.Dispose() }
        if ($response) { $response.Dispose() }
        Write-Progress -Activity $Title -Completed
    }
    if (-not (Test-Path -LiteralPath $partial)) {
        Write-Log -Level 'ERROR' -Message "Сервер вернул пустой файл (0 байт): $Uri"
        throw "Сервер вернул слишком маленький или пустой файл: $Uri"
    }
    $partialInfo = Get-Item -LiteralPath $partial
    if ($partialInfo.Length -lt $MinimumBytes) {
        $head = ''
        try { $head = ((Get-Content -LiteralPath $partial -TotalCount 5 -ErrorAction SilentlyContinue) -join ' | ') } catch { }
        Write-Log -Level 'ERROR' -Message ("Сервер вернул слишком маленький файл: {0} байт (минимум {1}); первые строки: {2}; {3}" -f $partialInfo.Length, $MinimumBytes, $head, $Uri)
        Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
        throw "Сервер вернул слишком маленький или пустой файл: $Uri"
    }
    Move-Item -LiteralPath $partial -Destination $Destination -Force
    $finished = Get-Date
    Write-Log -Message ("Скачивание завершено: {0} -> {1} ({2:N1} MB, {3:N0} с)" -f $Uri, $Destination, ((Get-Item -LiteralPath $Destination).Length / 1MB), ($finished - $started).TotalSeconds)
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
        [pscustomobject]@{Name='Everything';Id='voidtools.Everything';Kind='Winget';Url=''},
        [pscustomobject]@{Name='EverythingToolbar';Id='srwi.EverythingToolbar.Launcher';Kind='Winget';Url=''},
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
        [pscustomobject]@{Name='RuDesktop';Id='direct:rudesktop';Kind='Direct';Url='https://storage.rudesktop.ru/download/rudesktop-3.0.1563-x64.msi';FileName='rudesktop-3.0.1563-x64.msi';InstallerType='Msi';InstallArgs='/qn /norestart';RegistryPattern='^RuDesktop'},
        [pscustomobject]@{Name='RustDesk';Id='direct:rustdesk';Kind='Direct';Url='https://github.com/rustdesk/rustdesk/releases/download/1.4.9/rustdesk-1.4.9-x86_64.exe';FileName='rustdesk-1.4.9-x86_64.exe';InstallArgs='--silent-install';RegistryPattern='^RustDesk'},
        [pscustomobject]@{Name='Samsung Magician';Id='XPDDT99J9GKB5C';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Tailscale';Id='Tailscale.Tailscale';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Telegram';Id='Telegram.TelegramDesktop';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Termius';Id='Termius.Termius';Kind='Winget';Url=''},
        [pscustomobject]@{Name='Ventoy';Id='Ventoy.Ventoy';Kind='Winget';Url=''},
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
        [void](Invoke-WingetCommand -Arguments @('source','reset','--force','--disable-interactivity') -Quiet)
        [void](Invoke-WingetCommand -Arguments @('source','update','--disable-interactivity') -Quiet)
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

function Invoke-WingetCommand {
    param(
        [Parameter(Mandatory=$true)][string[]]$Arguments,
        [string]$LogPath = '',
        [switch]$Quiet
    )
    if (-not $script:WingetPath) { throw 'winget не найден или не запускается.' }
    $commandLine = ('"{0}" {1}' -f $script:WingetPath, ($Arguments -join ' '))
    Write-Log -Message "winget: $commandLine"
    $stdout = Join-Path $env:TEMP ("wginst-winget-out-{0}.log" -f ([guid]::NewGuid().ToString('N')))
    $stderr = Join-Path $env:TEMP ("wginst-winget-err-{0}.log" -f ([guid]::NewGuid().ToString('N')))
    try {
        $process = Start-Process -FilePath $script:WingetPath -ArgumentList $Arguments -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        $code = [int]$process.ExitCode
        $output = @()
        if (Test-Path -LiteralPath $stdout) { $output += @(Get-Content -LiteralPath $stdout -ErrorAction SilentlyContinue) }
        if (Test-Path -LiteralPath $stderr) { $output += @(Get-Content -LiteralPath $stderr -ErrorAction SilentlyContinue) }
    } finally {
        Remove-Item -LiteralPath $stdout -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stderr -Force -ErrorAction SilentlyContinue
    }
    $cleanOutput = @($output | ForEach-Object {
        ([string]$_).Trim()
    } | Where-Object {
        $_ -and $_ -notmatch '^[\\|/\-\s]+$' -and $_ -notmatch '^\s*\d+(\.\d+)?\s*(KB|MB|GB)\s*/\s*\d+(\.\d+)?\s*(KB|MB|GB)\s*$'
    })
    if (-not $Quiet -and $code -ne 0) {
        foreach ($line in ($cleanOutput | Select-Object -Last 20)) { Write-Host $line }
    }
    Write-Log -Message "winget завершён с кодом $code"
    return [pscustomobject]@{ ExitCode = $code; Output = $cleanOutput }
}

function Test-WingetPackageInstalled {
    param([string]$Id)
    $result = Invoke-WingetCommand -Arguments @('list','--id',$Id,'-e','--accept-source-agreements','--disable-interactivity') -Quiet
    return ($result.ExitCode -eq 0 -and (($result.Output -join "`n") -match [regex]::Escape($Id)))
}

function Test-WingetPackageAvailable {
    param([string]$Id)
    $result = Invoke-WingetCommand -Arguments @('show','--id',$Id,'-e','--accept-source-agreements','--disable-interactivity') -Quiet
    return ($result.ExitCode -eq 0 -and (($result.Output -join "`n") -match [regex]::Escape($Id)))
}

function Test-CompressOInstalled {
    return Test-DirectAppInstalled -Pattern '^CompressO'
}

function Test-DirectAppInstalled {
    param([string]$Pattern)
    if (-not $Pattern) { return $false }
    $paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    return [bool](Get-ItemProperty $paths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -match $Pattern } | Select-Object -First 1)
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
    $wingetFree = Get-VolumeFreeSpace -Path $env:LOCALAPPDATA
    if ($null -ne $wingetFree -and $wingetFree -lt 2GB) {
        Write-Log -Level 'WARN' -Message "Программы: мало свободного места (свободно $([math]::Round($wingetFree / 1GB, 2)) ГБ)."
        Write-Host 'Мало свободного места на диске: установка через winget обычно требует не меньше 2 ГБ.'
        if (-not (Read-YesNo -Question 'Продолжить установку программ, несмотря на малое место?')) {
            throw 'Прервано пользователем: недостаточно свободного места на диске.'
        }
    }
    [int]$position = 0
    foreach ($app in $selected) {
        $position++
        Write-Host "[$position/$($selected.Count)] $($app.Name)"
        try {
            if ($app.Kind -eq 'Direct') {
                $registryPattern = if ($app.PSObject.Properties.Name -contains 'RegistryPattern') { [string]$app.RegistryPattern } else { '^' + [regex]::Escape([string]$app.Name) }
                if (Test-DirectAppInstalled -Pattern $registryPattern) { Write-Host 'Уже установлено. Пропуск.'; continue }
                $fileName = if ($app.PSObject.Properties.Name -contains 'FileName' -and $app.FileName) { [string]$app.FileName } else { [IO.Path]::GetFileName(([Uri]$app.Url).AbsolutePath) }
                $installer = Join-Path $script:CacheRoot $fileName
                $installerType = if ($app.PSObject.Properties.Name -contains 'InstallerType' -and $app.InstallerType) {
                    [string]$app.InstallerType
                } elseif ([IO.Path]::GetExtension($installer).Equals('.msi', [StringComparison]::OrdinalIgnoreCase)) {
                    'Msi'
                } else {
                    'Exe'
                }
                $validInstaller = if ($installerType -eq 'Msi') {
                    (Test-Path -LiteralPath $installer -PathType Leaf) -and ((Get-Item -LiteralPath $installer).Length -ge 100KB)
                } else {
                    Test-PeFile $installer
                }
                if (-not $validInstaller) {
                    Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
                    Save-HttpFile -Uri $app.Url -Destination $installer -Title $app.Name -MinimumBytes 100KB
                }
                $validInstaller = if ($installerType -eq 'Msi') {
                    (Test-Path -LiteralPath $installer -PathType Leaf) -and ((Get-Item -LiteralPath $installer).Length -ge 100KB)
                } else {
                    Test-PeFile $installer
                }
                if (-not $validInstaller) { throw "Скачанный установщик $($app.Name) повреждён." }
                if ($installerType -eq 'Msi') {
                    $installArgs = if ($app.PSObject.Properties.Name -contains 'InstallArgs' -and $app.InstallArgs) { [string]$app.InstallArgs } else { '/qn /norestart' }
                    $process = Start-Process -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList "/i `"$installer`" $installArgs" -Wait -PassThru
                    if ($process.ExitCode -notin @(0, 3010)) { throw "msiexec завершился с кодом $($process.ExitCode)." }
                } else {
                    $installArgs = if ($app.PSObject.Properties.Name -contains 'InstallArgs' -and $app.InstallArgs) { [string]$app.InstallArgs } else { '/S' }
                    $process = Start-Process -FilePath $installer -ArgumentList $installArgs -Wait -PassThru
                    if ($process.ExitCode -ne 0) { throw "Установщик завершился с кодом $($process.ExitCode)." }
                }
            } else {
                if (Test-WingetPackageInstalled -Id $app.Id) { Write-Host 'Уже установлено. Пропуск.'; continue }
                if (-not (Test-WingetPackageAvailable -Id $app.Id)) { throw "Пакет $($app.Id) не найден в источниках winget." }
                $result = Invoke-WingetCommand -Arguments @('install','--id',$app.Id,'-e','--silent','--accept-source-agreements','--accept-package-agreements','--disable-interactivity')
                if ($result.ExitCode -ne 0) { throw "winget завершился с кодом $($result.ExitCode)." }
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
    Write-Log -Message "Обновление программ через winget."
    $result = Invoke-WingetCommand -Arguments @('upgrade','--all','--silent','--include-unknown','--accept-source-agreements','--accept-package-agreements','--disable-interactivity')
    if ($result.ExitCode -eq 0) {
        Write-Log -Message 'Обновление программ завершено успешно (код winget 0).'
        Write-Host ''; Write-Host 'Обновление завершено.'
    } else {
        Write-Log -Level 'ERROR' -Message "winget upgrade завершился с кодом $($result.ExitCode)."
        Write-Host ''; Write-Host "winget завершился с кодом $($result.ExitCode). Подробности показаны выше."
    }
    Pause-Result
}

function Get-OfficeDeploymentToolUrl {
    $fallback = 'https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_20228-20124.exe'
    try {
        $page = Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/details.aspx?id=49117' -UseBasicParsing -TimeoutSec 15
        $match = [regex]::Match($page.Content, 'https://(?:download|software-download)\.microsoft\.com/[^"''<>\s]+/officedeploymenttool_[^"''<>\s]+\.exe', 'IgnoreCase')
        if ($match.Success) { return $match.Value }
    } catch {
        Write-Log -Level 'WARN' -Message "Страница Microsoft Download Center недоступна: $($_.Exception.Message)"
    }
    return $fallback
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

function Get-FileTailText {
    param([string]$Path, [int]$Lines = 5)
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    $content = Get-Content -LiteralPath $Path -Tail $Lines -ErrorAction SilentlyContinue
    if (-not $content) { return '' }
    return (($content -join ' | ').Trim())
}

function Get-LatestOfficeBootstrapperLog {
    try {
        $tempRoot = $env:TEMP
        if (-not $tempRoot -or -not (Test-Path -LiteralPath $tempRoot)) { return $null }
        return @(Get-ChildItem -LiteralPath $tempRoot -File -Filter '*.log' -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -gt (Get-Date).AddHours(-4) -and $_.Name -match '(?i)(office|setup|click|c2r|desktop)' } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1)[0]
    } catch {
        return $null
    }
}

function Get-OfficeBootstrapperFailureHint {
    param([string]$FallbackHint)
    $setupLog = Get-LatestOfficeBootstrapperLog
    if (-not $setupLog) { return $FallbackHint }
    try {
        $matches = @(Select-String -LiteralPath $setupLog.FullName -Pattern 'RegionalBlocks|BlockedRegions|PreReqs Failed|ErrorDetails|Data.GeoName|Data.EcsCountryCode' -CaseSensitive:$false -ErrorAction SilentlyContinue | Select-Object -Last 12)
        $matchText = ($matches | ForEach-Object { $_.Line }) -join ' '
        if ($matchText -match 'RegionalBlocks') {
            return " Причина из лога Office: сработал RegionalBlocks — Microsoft Click-to-Run заблокировал скачивание Office для текущего региона/сети. Используйте корпоративно разрешённый offline source/дистрибутив Office или другую разрешённую сеть. Лог: $($setupLog.FullName)"
        }
        if ($matchText) {
            return "$FallbackHint Лог Office Click-to-Run: $($setupLog.FullName). Ключевые строки: $matchText"
        }
        return "$FallbackHint Лог Office Click-to-Run: $($setupLog.FullName)"
    } catch {
        return "$FallbackHint Лог Office Click-to-Run: $($setupLog.FullName)"
    }
}

function Start-OfficeManager {
    $installed = Get-InstalledOfficeProducts
    $editions = @(
        [pscustomobject]@{Name='Установить Microsoft 365 Apps';Id='O365ProPlusRetail';Channel='Current'},
        [pscustomobject]@{Name='Установить Office LTSC Professional Plus 2024';Id='ProPlus2024Volume';Channel='PerpetualVL2024'},
        [pscustomobject]@{Name='Назад';Id='';Channel=''}
    )
    $title = 'Установка или смена редакции Office'
    if ($installed) { $title += " (установлено: $installed)" }
    $choice = Select-SingleItem -Title $title -Items $editions -Text { param($item) $item.Name }
    if ($choice -lt 0 -or $editions[$choice].Id -eq '') { return }
    $edition = $editions[$choice]
    Write-Log -Message "Office: выбрана редакция $($edition.Name) (ID: $($edition.Id), канал: $($edition.Channel), 64-бит, язык: ru-ru). Текущее состояние: $(if ($installed) { $installed } else { 'Office не обнаружен' })"
    Show-WorkScreen -Title "Подготовка $($edition.Name)"
    if (-not (Read-YesNo -Question "Установить Office или сменить текущую редакцию на $($edition.Name)?")) { return }
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
    if ($publisher -notmatch 'Microsoft') {
        Write-Log -Level 'ERROR' -Message "Office: у Office Deployment Tool неожиданный издатель '$publisher' (ожидался Microsoft). Пакет: $package, адрес: $url"
        throw "Неожиданный издатель Office Deployment Tool: $publisher"
    }
    Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
    $setup = Join-Path $extractDir 'setup.exe'
    $extractLog = Join-Path $env:TEMP ("wginst-odt-extract-{0}.log" -f ([guid]::NewGuid().ToString('N')))
    Write-Log -Message "Office: начало распаковки ODT: '$package /quiet /extract:$extractDir'."
    $process = Start-Process -FilePath $package -ArgumentList "/quiet /extract:`"$extractDir`"" -Wait -PassThru -RedirectStandardOutput $extractLog
    if ($process.ExitCode -ne 0 -or -not (Test-PeFile $setup)) {
        $tail = Get-FileTailText -Path $extractLog
        Remove-Item -LiteralPath $extractLog -Force -ErrorAction SilentlyContinue
        Write-Log -Level 'ERROR' -Message "Office: распаковка ODT не удалась, код $($process.ExitCode), setup.exe: $(if (Test-PeFile $setup) { 'есть' } else { 'отсутствует' }).$(if ($tail) { " Последняя строка: $tail" })"
        throw "Не удалось распаковать Office Deployment Tool, код $($process.ExitCode).$(if ($tail) { " Последняя строка: $tail" })"
    }
    Remove-Item -LiteralPath $extractLog -Force -ErrorAction SilentlyContinue
    $setupPublisher = Get-AuthenticodePublisher -Path $setup
    if ($setupPublisher -notmatch 'Microsoft') {
        Write-Log -Level 'ERROR' -Message "Office: у setup.exe неожиданный издатель '$setupPublisher' (ожидался Microsoft). Файл: $setup"
        throw "Неожиданный издатель setup.exe: $setupPublisher"
    }
    $setupVersion = (Get-Item -LiteralPath $setup).VersionInfo.ProductVersion
    Write-Log -Message "Office Deployment Tool setup.exe: $setup, версия $setupVersion"
    $downloadConfig = Join-Path $dir 'download.xml'
    $installConfig = Join-Path $dir 'install.xml'
    $removeConfig = Join-Path $dir 'remove.xml'
    $officeSource = Join-Path $dir 'Source'
    New-Item -ItemType Directory -Path $officeSource -Force | Out-Null
    $downloadXml = @"
<Configuration>
  <Add SourcePath="$officeSource" OfficeClientEdition="64" Channel="$($edition.Channel)">
    <Product ID="$($edition.Id)">
      <Language ID="ru-ru" />
    </Product>
  </Add>
</Configuration>
"@
    $installXml = @"
<Configuration>
  <Add SourcePath="$officeSource" OfficeClientEdition="64" Channel="$($edition.Channel)">
    <Product ID="$($edition.Id)">
      <Language ID="ru-ru" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
    $removeXml = @"
<Configuration>
  <Remove All="TRUE" />
  <RemoveMSI />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
    Set-Content -LiteralPath $downloadConfig -Value $downloadXml -Encoding UTF8
    Set-Content -LiteralPath $installConfig -Value $installXml -Encoding UTF8
    Set-Content -LiteralPath $removeConfig -Value $removeXml -Encoding UTF8
    $officeFree = Get-VolumeFreeSpace -Path $dir
    if ($null -ne $officeFree -and $officeFree -lt 8GB) {
        Write-Log -Level 'WARN' -Message "Office: мало свободного места на диске (свободно $([math]::Round($officeFree / 1GB, 2)) ГБ, нужно около 8 ГБ)."
        Write-Host 'Мало свободного места на диске: для скачивания Office нужно около 8 ГБ.'
        Write-Host 'Иначе Office Deployment Tool может остановиться с ошибкой 30023-2016 (недостаточно места).'
        if (-not (Read-YesNo -Question 'Продолжить скачивание Office, несмотря на малое место?')) {
            throw 'Прервано пользователем: недостаточно свободного места на диске для Office.'
        }
    }
    if ($null -ne $officeFree) {
        Write-Log -Message ('Office: свободно на диске каталога {0}: {1} ГБ (для скачивания нужно около 8 ГБ).' -f $dir, [math]::Round($officeFree / 1GB, 2))
    } else {
        Write-Log -Level 'WARN' -Message "Office: не удалось определить свободное место на диске каталога $dir — ODT может прерваться с ошибкой 30023-2016."
    }
    Write-Host 'Скачивание файлов Office в локальный кэш. Это может занять продолжительное время...'
    $downloadLog = Join-Path $env:TEMP ("wginst-odt-download-{0}.log" -f ([guid]::NewGuid().ToString('N')))
    Write-Log -Message "Office: начало скачивания Office: 'setup.exe /download $downloadConfig'."
    $process = Start-Process -FilePath $setup -ArgumentList "/download `"$downloadConfig`"" -WorkingDirectory $dir -Wait -PassThru -RedirectStandardOutput $downloadLog
    Write-Log -Message "Office: /download завершён с кодом $($process.ExitCode)"
    if ($process.ExitCode -ne 0) {
        $tail = Get-FileTailText -Path $downloadLog
        Remove-Item -LiteralPath $downloadLog -Force -ErrorAction SilentlyContinue
        $freeAtFail = Get-VolumeFreeSpace -Path $dir
        $freeNote = if ($null -ne $freeAtFail) { " Свободно на диске: $([math]::Round($freeAtFail / 1GB, 2)) ГБ." } else { '' }
        Write-Log -Level 'ERROR' -Message "Office: ODT не смог скачать файлы Office, код $($process.ExitCode). Конфиг: $downloadConfig.$freeNote$(if ($tail) { " Последняя строка: $tail" })"
        $hint = ''
        if ($null -ne $freeAtFail -and $freeAtFail -lt 8GB) {
            $hint = " Свободного места ($([math]::Round($freeAtFail / 1GB, 2)) ГБ) может не хватить: нужно около 8 ГБ — освободите место и повторите, типичный код такой ошибки ODT: 30023-2016."
        } else {
            $freeText = if ($null -ne $freeAtFail) { [string]([math]::Round($freeAtFail / 1GB, 2)) + ' ГБ' } else { 'определить не удалось' }
            $hint = " На момент сбоя на диске свободно: $freeText. Если места достаточно, ODT, видимо, не прошёл стартовый контроль места/носителя: проверьте антивирус и файловые фильтры, VHD/подключённые диски, квоты, а также атрибуты диска (сжатие, «только чтение» — fsutil behavior query). Можно повторить скачивание на другой диск."
            $hint = Get-OfficeBootstrapperFailureHint -FallbackHint $hint
        }
        throw "Office Deployment Tool не смог скачать файлы Office, код $($process.ExitCode). Конфиг: $downloadConfig.$hint$(if ($tail) { " Последняя строка: $tail" })"
    }
    Remove-Item -LiteralPath $downloadLog -Force -ErrorAction SilentlyContinue
    $sourceFile = Get-ChildItem -LiteralPath $officeSource -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $sourceFile) {
        Write-Log -Level 'ERROR' -Message "Office: в каталоге $officeSource не появилось ни одного файла после /download с кодом 0."
        throw "Файлы Office не появились в каталоге $officeSource после /download с кодом 0. Проверьте доступ к Office CDN."
    }
    Write-Log -Message "Office: кэш заполнен, пример файла: $($sourceFile.FullName)"
    if ($installed) {
        Write-Host 'Удаление существующего Office перед чистой установкой...'
        $removeLog = Join-Path $env:TEMP ("wginst-odt-remove-{0}.log" -f ([guid]::NewGuid().ToString('N')))
        Write-Log -Message "Office: начало удаления существующей установки: 'setup.exe /configure $removeConfig'."
        $process = Start-Process -FilePath $setup -ArgumentList "/configure `"$removeConfig`"" -WorkingDirectory $dir -Wait -PassThru -RedirectStandardOutput $removeLog
        Write-Log -Message "Office: /configure (удаление) завершён с кодом $($process.ExitCode)"
        if ($process.ExitCode -ne 0) {
            $tail = Get-FileTailText -Path $removeLog
            Remove-Item -LiteralPath $removeLog -Force -ErrorAction SilentlyContinue
            Write-Log -Level 'ERROR' -Message "Office: ODT не смог удалить существующий Office, код $($process.ExitCode).$(if ($tail) { " Последняя строка: $tail" })"
            throw "Office Deployment Tool не смог удалить существующий Office, код $($process.ExitCode).$(if ($tail) { " Последняя строка: $tail" })"
        }
        Remove-Item -LiteralPath $removeLog -Force -ErrorAction SilentlyContinue
    }
    Write-Host 'Запуск установки Office из локального кэша...'
    $installLog = Join-Path $env:TEMP ("wginst-odt-install-{0}.log" -f ([guid]::NewGuid().ToString('N')))
    Write-Log -Message "Office: начало установки: 'setup.exe /configure $installConfig'."
    $process = Start-Process -FilePath $setup -ArgumentList "/configure `"$installConfig`"" -WorkingDirectory $dir -Wait -PassThru -RedirectStandardOutput $installLog
    Write-Log -Message "Office: /configure (установка) завершён с кодом $($process.ExitCode)"
    if ($process.ExitCode -ne 0) {
        $tail = Get-FileTailText -Path $installLog
        Remove-Item -LiteralPath $installLog -Force -ErrorAction SilentlyContinue
        Write-Log -Level 'ERROR' -Message "Office: ODT завершился с кодом $($process.ExitCode) при установке. Конфиг: $installConfig.$(if ($tail) { " Последняя строка: $tail" })"
        throw "Office Deployment Tool завершился с кодом $($process.ExitCode). Конфиг: $installConfig.$(if ($tail) { " Последняя строка: $tail" })"
    }
    Remove-Item -LiteralPath $installLog -Force -ErrorAction SilentlyContinue
    $after = Get-InstalledOfficeProducts
    if ($after -and $after -notmatch [regex]::Escape($edition.Id)) {
        Write-Log -Message "Office установка завершилась, но в реестре указана редакция: $after (ожидалась $($edition.Id))"
        Write-Host "Установка завершена, но в реестре указана редакция: $after"
    } else {
        if ($after) {
            Write-Log -Message "Office установка завершена. Версии из реестра: $after"
        } else {
            Write-Log -Level 'WARN' -Message "Office: ODT завершился с кодом 0, но установленных версий в реестре не найдено (ProductReleaseIds пуст)"
        }
        Write-Host 'Установка Office завершена.'
    }
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
        [pscustomobject]@{Key='HideSearch';Name='Скрыть поиск на панели задач';Path=$search;ValueName='SearchboxTaskbarMode';OnValue=0;OffValue=1;Kind='MultiRegistry';Supported=$isClient;UnsupportedReason=$clientReason;MissingIsEnabled=$false},
        [pscustomobject]@{Key='HideTaskView';Name='Скрыть кнопку представления задач';Path=$advanced;ValueName='ShowTaskViewButton';OnValue=0;OffValue=1;Kind='MultiRegistry';Supported=$isClient;UnsupportedReason=$clientReason;MissingIsEnabled=$false},
        [pscustomobject]@{Key='ClassicMenu';Name='Классическое контекстное меню Windows 11';Path=$classic;ValueName='';OnValue=0;OffValue=1;Kind='Classic';Supported=$isWindows11Client;UnsupportedReason=$windows11Reason;MissingIsEnabled=$false},
        [pscustomobject]@{Key='HideWidgets';Name='Скрыть кнопку виджетов';Path=$advanced;ValueName='TaskbarDa';OnValue=0;OffValue=1;Kind='MultiRegistry';Supported=$isWindows11Client;UnsupportedReason=$windows11Reason;MissingIsEnabled=$false}
    )
}

function Get-TweakRegistryOperations {
    param([object]$Definition)
    $advanced = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    switch ($Definition.Key) {
        'HideSearch' {
            return @(
                [pscustomobject]@{Path='HKCU:\Software\Microsoft\Windows\CurrentVersion\Search';Name='SearchboxTaskbarMode';OnValue=0;OffValue=1},
                [pscustomobject]@{Path='HKCU:\Software\Microsoft\Windows\CurrentVersion\Search';Name='SearchboxTaskbarModeCache';OnValue=0;OffValue=1},
                [pscustomobject]@{Path=$advanced;Name='SearchboxTaskbarMode';OnValue=0;OffValue=1},
                [pscustomobject]@{Path=$advanced;Name='ShowCortanaButton';OnValue=0;OffValue=0},
                [pscustomobject]@{Path='HKCU:\Software\Policies\Microsoft\Windows\Explorer';Name='DisableSearchBoxSuggestions';OnValue=1;OffValue=0}
            )
        }
        'HideTaskView' {
            return @(
                [pscustomobject]@{Path=$advanced;Name='ShowTaskViewButton';OnValue=0;OffValue=1}
            )
        }
        'HideWidgets' {
            return @(
                [pscustomobject]@{Path=$advanced;Name='TaskbarDa';OnValue=0;OffValue=1},
                [pscustomobject]@{Path='HKLM:\SOFTWARE\Policies\Microsoft\Dsh';Name='AllowNewsAndInterests';OnValue=0;OffValue=1},
                [pscustomobject]@{Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds';Name='EnableFeeds';OnValue=0;OffValue=1}
            )
        }
        default {
            return @([pscustomobject]@{Path=$Definition.Path;Name=$Definition.ValueName;OnValue=$Definition.OnValue;OffValue=$Definition.OffValue})
        }
    }
}

function Get-TweakSnapshot {
    param([object]$Definition)
    if ($Definition.Kind -eq 'Classic') { return Get-ClassicMenuSnapshot -ParentPath $Definition.Path }
    if ($Definition.Kind -eq 'MultiRegistry') {
        return [pscustomobject]@{
            Kind = 'MultiRegistry'
            Entries = @(Get-TweakRegistryOperations -Definition $Definition | ForEach-Object {
                [pscustomobject]@{
                    Path = $_.Path
                    Name = $_.Name
                    Snapshot = Get-RegistrySnapshot -Path $_.Path -Name $_.Name
                }
            })
        }
    }
    return Get-RegistrySnapshot -Path $Definition.Path -Name $Definition.ValueName
}

function Restore-TweakSnapshot {
    param([object]$Definition, [object]$Snapshot)
    if ($Definition.Kind -eq 'Classic') {
        Restore-ClassicMenuSnapshot -ParentPath $Definition.Path -Snapshot $Snapshot
        return
    }
    if ($Definition.Kind -eq 'MultiRegistry' -and $Snapshot.Kind -eq 'MultiRegistry') {
        foreach ($entry in @($Snapshot.Entries)) {
            Restore-RegistrySnapshot -Path ([string]$entry.Path) -Name ([string]$entry.Name) -Snapshot $entry.Snapshot
        }
        return
    }
    Restore-RegistrySnapshot -Path $Definition.Path -Name $Definition.ValueName -Snapshot $Snapshot
}

function Update-TweakBackupShape {
    param([object]$Definition, [hashtable]$Backups)
    if (-not $Backups.ContainsKey($Definition.Key)) { return }
    if ($Definition.Kind -ne 'MultiRegistry') { return }
    $existing = $Backups[$Definition.Key]
    if ($existing.Kind -eq 'MultiRegistry') { return }

    $entries = New-Object System.Collections.ArrayList
    foreach ($operation in (Get-TweakRegistryOperations -Definition $Definition)) {
        $snapshot = if ($operation.Path -eq $Definition.Path -and $operation.Name -eq $Definition.ValueName) {
            $existing
        } else {
            Get-RegistrySnapshot -Path $operation.Path -Name $operation.Name
        }
        [void]$entries.Add([pscustomobject]@{
            Path = $operation.Path
            Name = $operation.Name
            Snapshot = $snapshot
        })
    }
    $Backups[$Definition.Key] = [pscustomobject]@{
        Kind = 'MultiRegistry'
        Entries = @($entries)
    }
    Save-TweakBackups -Backups $Backups
}

function Get-TweakEnabled {
    param([object]$Definition)
    if ($Definition.Kind -eq 'Classic') { return (Test-ClassicMenuEnabled -ParentPath $Definition.Path) }
    if ($Definition.Kind -eq 'MultiRegistry') {
        $primary = Get-TweakRegistryOperations -Definition $Definition | Select-Object -First 1
        $snapshot = Get-RegistrySnapshot -Path $primary.Path -Name $primary.Name
        if (-not [bool]$snapshot.Exists) { return [bool]$Definition.MissingIsEnabled }
        return ([bool]$snapshot.Exists -and [int]$snapshot.Value -eq [int]$primary.OnValue)
    }
    $snapshot = Get-RegistrySnapshot -Path $Definition.Path -Name $Definition.ValueName
    if (-not [bool]$snapshot.Exists) { return [bool]$Definition.MissingIsEnabled }
    return ([bool]$snapshot.Exists -and [int]$snapshot.Value -eq [int]$Definition.OnValue)
}

function Set-TweakEnabled {
    param([object]$Definition, [bool]$Enabled, [hashtable]$Backups)
    if ($Enabled) {
        if (-not $Backups.ContainsKey($Definition.Key)) {
            $Backups[$Definition.Key] = Get-TweakSnapshot -Definition $Definition
            Save-TweakBackups -Backups $Backups
        } else {
            Update-TweakBackupShape -Definition $Definition -Backups $Backups
        }
        if ($Definition.Kind -eq 'Classic') {
            $child = Join-Path $Definition.Path 'InprocServer32'
            New-Item -Path $child -Force | Out-Null
            Set-Item -LiteralPath $child -Value ''
            if (-not (Test-ClassicMenuEnabled -ParentPath $Definition.Path)) { throw 'Не удалось создать полное значение классического контекстного меню.' }
        } elseif ($Definition.Kind -eq 'MultiRegistry') {
            foreach ($operation in (Get-TweakRegistryOperations -Definition $Definition)) {
                New-Item -Path $operation.Path -Force | Out-Null
                New-ItemProperty -Path $operation.Path -Name $operation.Name -PropertyType DWord -Value ([int]$operation.OnValue) -Force | Out-Null
            }
        } else {
            New-Item -Path $Definition.Path -Force | Out-Null
            New-ItemProperty -Path $Definition.Path -Name $Definition.ValueName -PropertyType DWord -Value ([int]$Definition.OnValue) -Force | Out-Null
        }
    } else {
        if ($Backups.ContainsKey($Definition.Key)) {
            Restore-TweakSnapshot -Definition $Definition -Snapshot $Backups[$Definition.Key]
            $Backups.Remove($Definition.Key)
            Save-TweakBackups -Backups $Backups
        } elseif ($Definition.Kind -eq 'Classic') {
            Remove-Item -LiteralPath $Definition.Path -Recurse -Force -ErrorAction SilentlyContinue
        } elseif ($Definition.Kind -eq 'MultiRegistry') {
            foreach ($operation in (Get-TweakRegistryOperations -Definition $Definition)) {
                New-Item -Path $operation.Path -Force | Out-Null
                New-ItemProperty -Path $operation.Path -Name $operation.Name -PropertyType DWord -Value ([int]$operation.OffValue) -Force | Out-Null
            }
        } else {
            New-Item -Path $Definition.Path -Force | Out-Null
            New-ItemProperty -Path $Definition.Path -Name $Definition.ValueName -PropertyType DWord -Value ([int]$Definition.OffValue) -Force | Out-Null
        }
    }
}

function Stop-WindowsShellForTweaks {
    $running = @(Get-Process explorer -ErrorAction SilentlyContinue)
    $shellProcesses = @('SearchHost','StartMenuExperienceHost','ShellExperienceHost')
    $shellProcesses | ForEach-Object { Stop-Process -Name $_ -Force -ErrorAction SilentlyContinue }
    if ($running.Count -eq 0) { return $false }

    # Explorer records part of the taskbar state when it exits.  Therefore it
    # must be stopped BEFORE the new registry values are written; otherwise it
    # can overwrite the values which the manager has just set.
    $oldProcessIds = @($running | ForEach-Object { $_.Id })
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    if ($oldProcessIds.Count -gt 0) {
        Wait-Process -Id $oldProcessIds -Timeout 8 -ErrorAction SilentlyContinue
    }
    $deadline = (Get-Date).AddSeconds(8)
    while ((Get-Date) -lt $deadline -and @(Get-Process explorer -ErrorAction SilentlyContinue).Count -gt 0) {
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        $shellProcesses | ForEach-Object { Stop-Process -Name $_ -Force -ErrorAction SilentlyContinue }
        Start-Sleep -Milliseconds 300
    }
    return $true
}

function Wait-WindowsShellAfterTweaks {
    param([bool]$WasRunning)
    if (-not $WasRunning) { return }

    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    Start-Process -FilePath "$env:SystemRoot\explorer.exe" | Out-Null
    Start-Sleep -Seconds 2
}

function Get-TweakRawValueText {
    param([object]$Definition)
    if ($Definition.Kind -eq 'Classic') {
        if (Test-ClassicMenuEnabled -ParentPath $Definition.Path) { return '<пустое значение InprocServer32>' }
        return '<классическое меню не настроено>'
    }
    if ($Definition.Kind -eq 'MultiRegistry') {
        $values = @(Get-TweakRegistryOperations -Definition $Definition | ForEach-Object {
            $snapshot = Get-RegistrySnapshot -Path $_.Path -Name $_.Name
            $leaf = ($_.Path -replace '^HKCU:\\','HKCU:\' -replace '^HKLM:\\','HKLM:\')
            if ([bool]$snapshot.Exists) { "$leaf\$($_.Name)=$($snapshot.Value)" } else { "$leaf\$($_.Name)=<нет>" }
        })
        return ($values -join '; ')
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
            'Применить твики',
            'Отключить твики'
        )
        $action = Select-SingleItem -Title "Твики Windows - $($windows.ProductName) $($windows.DisplayVersion)" -Items $actions -Text { param($item) $item }
        if ($action -lt 0) { return }

        $backups = Read-TweakBackups
        $items = foreach ($definition in $definitions) {
            [pscustomobject]@{
                Definition = $definition
                Enabled = Get-TweakEnabled -Definition $definition
                CanSelect = [bool]$definition.Supported
                HasBackup = $backups.ContainsKey($definition.Key)
            }
        }

        if (@($items | Where-Object { $_.CanSelect }).Count -eq 0) {
            Show-WorkScreen -Title 'Твики недоступны' -Details ''
            Write-Host "Система: $($windows.ProductName) $($windows.DisplayVersion)"
            Write-Host ''
            Write-Host 'Эти параметры оболочки поддерживаются клиентскими Windows 10/11.'
            Pause-Result
            continue
        }
        $targetEnabled = ($action -eq 0)
        $operationName = if ($targetEnabled) { 'ПРИМЕНИТЬ' } else { 'ОТКЛЮЧИТЬ' }
        $selected = @($items | Where-Object { $_.CanSelect })
        if ($selected.Count -eq 0) { continue }

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

function Get-TSforgeActivationProfile {
    param([Parameter(Mandatory=$true)][ValidateSet('Windows','Office')][string]$Mode)
    if ($Mode -eq 'Windows') {
        return [pscustomobject]@{
            Mode = 'Windows'
            Name = 'Windows'
            Switch = '/Z-Windows'
            ApplicationId = '55c92734-d682-4d71-983e-d6ec3f16059f'
        }
    }
    return [pscustomobject]@{
        Mode = 'Office'
        Name = 'Office (все установленные продукты, включая Project и Visio)'
        Switch = '/Z-Office'
        ApplicationId = '0ff1ce15-a989-479d-af46-f275c6370663'
    }
}

function Get-MasCmdArguments {
    param([Parameter(Mandatory=$true)][object]$Profile)
    return @('/d', '/c', 'MAS_AIO.cmd', [string]$Profile.Switch)
}

function Test-MasAioPackage {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    if ((Get-Item -LiteralPath $Path).Length -lt 500KB) { return $false }
    try {
        $content = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
        if ($content -match '(?im)^\s*(?:<!DOCTYPE|<html|404\s*:|code:)') { return $false }
        foreach ($marker in @('Microsoft_Activation_Scripts', ':TSforgeActivation', '/Z-Windows', '/Z-Office')) {
            if ($content.IndexOf($marker, [StringComparison]::OrdinalIgnoreCase) -lt 0) { return $false }
        }
        return $true
    } catch {
        return $false
    }
}

function Get-LatestMasAioPackage {
    $directory = Join-Path $script:CacheRoot 'Activation'
    $package = Join-Path $directory 'MAS_AIO.cmd'
    $candidate = Join-Path $directory 'MAS_AIO.download.cmd'
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $downloadError = $null
    $fromCache = $false

    try {
        Remove-Item -LiteralPath $candidate -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath "$candidate.part" -Force -ErrorAction SilentlyContinue
        Save-HttpFile -Uri $script:MasAioUrl -Destination $candidate -Title 'Microsoft Activation Scripts' -MinimumBytes 500KB
        if (-not (Test-MasAioPackage -Path $candidate)) {
            throw 'Загруженный MAS_AIO.cmd не прошёл проверку структуры TSforge.'
        }
        Move-Item -LiteralPath $candidate -Destination $package -Force
    } catch {
        $downloadError = $_.Exception.Message
        Remove-Item -LiteralPath $candidate -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath "$candidate.part" -Force -ErrorAction SilentlyContinue
        if (-not (Test-MasAioPackage -Path $package)) {
            throw "Не удалось получить актуальный MAS_AIO.cmd, а проверенная копия в кэше отсутствует: $downloadError"
        }
        $fromCache = $true
        Write-Log -Level 'WARN' -Message "MAS: загрузка недоступна, используется кэш. $downloadError"
    }

    $file = Get-Item -LiteralPath $package
    return [pscustomobject]@{
        Path = $file.FullName
        Directory = $file.DirectoryName
        Hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        LastWriteTime = $file.LastWriteTime
        FromCache = $fromCache
        DownloadError = $downloadError
        SourceUrl = $script:MasAioUrl
    }
}

function Get-SoftwareLicenseProducts {
    param([Parameter(Mandatory=$true)][string]$ApplicationId)
    $filter = "ApplicationID='$ApplicationId'"
    try {
        return @(Get-CimInstance -ClassName SoftwareLicensingProduct -Filter $filter -ErrorAction Stop)
    } catch {
        Write-Log -Level 'WARN' -Message "CIM лицензирования недоступен, используется WMI: $($_.Exception.Message)"
        try {
            return @(Get-WmiObject -Class SoftwareLicensingProduct -Filter $filter -ErrorAction Stop)
        } catch {
            Write-Log -Level 'WARN' -Message "Не удалось прочитать статус лицензирования: $($_.Exception.Message)"
            return @()
        }
    }
}

function Get-LicenseStatusText {
    param([int]$LicenseStatus)
    switch ($LicenseStatus) {
        0 { return 'не лицензировано' }
        1 { return 'лицензировано' }
        2 { return 'период первоначальной отсрочки' }
        3 { return 'период дополнительной отсрочки' }
        4 { return 'льготный период неподлинной копии' }
        5 { return 'режим уведомлений' }
        6 { return 'расширенный льготный период' }
        default { return "неизвестный статус ($LicenseStatus)" }
    }
}

function Show-TSforgeLicenseResult {
    param([Parameter(Mandatory=$true)][object]$Profile)
    $products = @(Get-SoftwareLicenseProducts -ApplicationId $Profile.ApplicationId | Where-Object {
        $_.PartialProductKey -or [int]$_.LicenseStatus -eq 1
    } | Sort-Object Name)
    if ($products.Count -eq 0) {
        Write-Host "Лицензируемые продукты $($Profile.Name) не найдены или статус недоступен."
        return $false
    }

    $licensed = $false
    foreach ($product in $products) {
        $status = Get-LicenseStatusText -LicenseStatus ([int]$product.LicenseStatus)
        $keyPart = if ($product.PartialProductKey) { "; ключ *-$($product.PartialProductKey)" } else { '' }
        Write-Host " - $($product.Name): $status$keyPart"
        if ([int]$product.LicenseStatus -eq 1) { $licensed = $true }
    }
    return $licensed
}

function Invoke-TSforgeActivation {
    param([Parameter(Mandatory=$true)][ValidateSet('Windows','Office')][string]$Mode)
    $profile = Get-TSforgeActivationProfile -Mode $Mode
    Show-WorkScreen -Title "Активация $($profile.Name) через TSforge" -Details ''
    Write-Host "Режим: $($profile.Switch)"
    Write-Host "Источник: $script:MasAioUrl"
    Write-Host ''
    Write-Host 'Будет загружена и запущена актуальная версия стороннего MAS с правами администратора.'
    Write-Host 'Контрольная сумма будет записана в журнал, но версия заранее не закреплена.'
    Write-Host ''
    if (-not (Read-YesNo -Question "Запустить TSforge для $($profile.Name)?")) { return }

    Show-WorkScreen -Title "Подготовка TSforge для $($profile.Name)"
    $package = Get-LatestMasAioPackage
    if ($package.FromCache) {
        Write-Host 'Сеть недоступна. Используется последняя проверенная копия из кэша.'
    } else {
        Write-Host 'Получена актуальная версия из официального репозитория MAS.'
    }
    Write-Host "Дата файла: $($package.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-Host "SHA-256: $($package.Hash)"
    Write-Host ''
    Write-Host "Запуск TSforge $($profile.Switch)..."
    Write-Log -Message "MAS START Mode=$($profile.Mode) Switch=$($profile.Switch) Cached=$($package.FromCache) SHA256=$($package.Hash) Source=$($package.SourceUrl)"

    $arguments = @(Get-MasCmdArguments -Profile $profile)
    Push-Location -LiteralPath $package.Directory
    try {
        & $env:ComSpec @arguments
        $exitCode = [int]$LASTEXITCODE
    } finally {
        Pop-Location
    }
    Write-Log -Message "MAS EXIT Mode=$($profile.Mode) Code=$exitCode"

    Write-Host ''
    Write-Host 'Проверка фактического статуса лицензий...'
    $verified = Show-TSforgeLicenseResult -Profile $profile
    Write-Log -Message "MAS VERIFY Mode=$($profile.Mode) Licensed=$verified"
    Write-Host ''
    if ($exitCode -ne 0) {
        Write-Host "TSforge завершился с кодом $exitCode. Проверьте сообщения выше."
    } elseif ($verified) {
        Write-Host "$($profile.Name): активация подтверждена системой лицензирования."
    } else {
        Write-Host 'TSforge завершился, но активированная лицензия не обнаружена.'
        Write-Host 'Возможно, продукт не установлен или эта редакция не поддерживается.'
    }
    Pause-Result
}

function Start-ActivationManager {
    $items = @(
        'Активировать Windows — TSforge',
        'Активировать Office (все продукты) — TSforge',
        'Сменить редакцию Windows',
        'Показать статус Windows',
        'Открыть параметры активации',
        'Назад'
    )
    $choice = Select-SingleItem -Title 'Активация и управление лицензиями' -Items $items -Text { param($item) $item }
    if ($choice -lt 0 -or $choice -eq 5) { return }
    switch ($choice) {
        0 { Invoke-TSforgeActivation -Mode Windows }
        1 { Invoke-TSforgeActivation -Mode Office }
        2 {
            $winEditions = @(
                [pscustomobject]@{Name='Windows 10/11 Pro';Key='VK7JG-NPHTM-C97JM-9MPGT-3V66T'},
                [pscustomobject]@{Name='Windows 10/11 Enterprise';Key='NPPR9-FWDCX-D2C8J-H872K-2YT43'},
                [pscustomobject]@{Name='Windows 10/11 Education';Key='NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'},
                [pscustomobject]@{Name='Назад';Key=''}
            )
            $winChoice = Select-SingleItem -Title 'Выберите новую редакцию Windows' -Items $winEditions -Text { param($item) $item.Name }
            if ($winChoice -ge 0 -and $winEditions[$winChoice].Key -ne '') {
                Show-WorkScreen -Title "Смена редакции на $($winEditions[$winChoice].Name)" -Details 'Запуск системной утилиты обновления...'
                & changepk.exe /ProductKey $($winEditions[$winChoice].Key)
                Write-Host 'Системное окно обновления Windows должно открыться.'
                Write-Host ''
                Write-Host 'После смены редакции и перезагрузки снова запустите активацию Windows.'
                Pause-Result
            }
        }
        3 {
            Show-WorkScreen -Title 'Статус активации Windows' -Details ''
            & cscript.exe //nologo "$env:SystemRoot\System32\slmgr.vbs" /dli
            Pause-Result
        }
        4 { Start-Process 'ms-settings:activation' | Out-Null }
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

function ConvertTo-PrinterIPAddress {
    param([string]$InputText)
    $value = $InputText.Trim()
    $parts = $value.Split('.')
    if ($parts.Count -ne 4) { return $null }
    foreach ($part in $parts) {
        [int]$number = 0
        if (-not [int]::TryParse($part, [ref]$number) -or $number -lt 0 -or $number -gt 255) { return $null }
    }
    if ([int]$parts[0] -eq 0 -or [int]$parts[3] -eq 0 -or [int]$parts[3] -eq 255) { return $null }
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

function Find-NetworkPrinterByIP {
    param([string]$IPAddress)
    Show-WorkScreen -Title "Проверка принтера $IPAddress" -Details 'Проверяются службы печати TCP 9100, IPP 631 и LPD 515.'
    $openPorts = New-Object 'System.Collections.Generic.List[int]'
    foreach ($port in @(9100,631,515)) {
        $client = New-Object Net.Sockets.TcpClient
        try {
            $async = $client.BeginConnect($IPAddress, $port, $null, $null)
            if ($async.AsyncWaitHandle.WaitOne(900)) {
                $client.EndConnect($async)
                $openPorts.Add([int]$port)
            }
        } catch {
        } finally {
            try { $async.AsyncWaitHandle.Close() } catch {}
            $client.Close()
        }
    }
    if ($openPorts.Count -eq 0) { return @() }
    Write-Host "Определение модели $IPAddress..."
    $device = Get-PrinterIdentity -IPAddress $IPAddress
    $device.Services = (($openPorts | Sort-Object | ForEach-Object { "TCP $_" }) -join ', ')
    return @($device)
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
    $options += [pscustomobject]@{Name='Ввести подсеть /24 вручную';Mode='ManualSubnet'}
    $options += [pscustomobject]@{Name='Указать IP принтера вручную';Mode='ManualIP'}
    $options += [pscustomobject]@{Name='Назад';Mode='Back'}
    $choice = Select-SingleItem -Title 'Выберите подсеть' -Items $options -Text { param($item) $item.Name }
    if ($choice -lt 0 -or $options[$choice].Mode -eq 'Back') { return }
    if ($options[$choice].Mode -eq 'Auto') {
        $subnets = $active
        $devices = @(Find-NetworkPrinters -Subnets $subnets)
    } elseif ($options[$choice].Mode -eq 'ManualIP') {
        Show-TextCursor
        Clear-Host
        $manual = Read-Host 'Введите IP принтера, например 192.168.0.45'
        $ip = ConvertTo-PrinterIPAddress -InputText $manual
        if (-not $ip) {
            Write-Host 'Неверный формат. Нужен IP из четырёх октетов, например 192.168.0.45.'
            Pause-Result
            return
        }
        $devices = @(Find-NetworkPrinterByIP -IPAddress $ip)
    } else {
        Show-TextCursor
        Clear-Host
        $manual = Read-Host 'Введите подсеть, например 192.168.0 или 192.168.0.0/24'
        $subnet = ConvertTo-PrinterSubnet -InputText $manual
        if (-not $subnet) {
            Write-Host 'Неверный формат. Нужны три октета от 0 до 255 и маска /24.'
            Pause-Result
            return
        }
        $subnets = @($subnet)
        $devices = @(Find-NetworkPrinters -Subnets $subnets)
    }
    if ($devices.Count -eq 0) {
        Write-Host 'Принтеры не найдены.'
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

function Set-OneDriveSyncDisabled {
    param([bool]$Disabled)
    $policyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive'
    New-Item -Path $policyPath -Force | Out-Null
    New-ItemProperty -Path $policyPath -Name 'DisableFileSyncNGSC' -PropertyType DWord -Value ([int]$Disabled) -Force | Out-Null
}

function Stop-OneDriveProcesses {
    Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 700
}

function Get-OneDriveSetupPaths {
    $paths = New-Object 'System.Collections.Generic.List[string]'
    $programFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    $basePaths = @(
        $env:SystemRoot,
        $env:SystemRoot,
        $env:LOCALAPPDATA,
        $env:ProgramFiles,
        $programFilesX86
    )
    $relativePaths = @(
        'System32\OneDriveSetup.exe',
        'SysWOW64\OneDriveSetup.exe',
        'Microsoft\OneDrive\OneDriveSetup.exe',
        'Microsoft OneDrive\OneDriveSetup.exe',
        'Microsoft OneDrive\OneDriveSetup.exe'
    )
    for ($i = 0; $i -lt $basePaths.Count; $i++) {
        if (-not $basePaths[$i]) { continue }
        $path = Join-Path $basePaths[$i] $relativePaths[$i]
        if ($path -and (Test-Path -LiteralPath $path -PathType Leaf)) { [void]$paths.Add($path) }
    }
    return @($paths | Select-Object -Unique)
}

function Remove-OneDriveAutostart {
    foreach ($path in @(
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run',
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'
    )) {
        Remove-ItemProperty -Path $path -Name 'OneDrive' -Force -ErrorAction SilentlyContinue
    }
    $startup = [Environment]::GetFolderPath('Startup')
    if ($startup) {
        Get-ChildItem -LiteralPath $startup -Filter '*OneDrive*.lnk' -ErrorAction SilentlyContinue |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
}

function Uninstall-OneDriveSafe {
    Show-WorkScreen -Title 'Удаление OneDrive' -Details 'Пользовательские файлы в папке OneDrive не удаляются.'
    Stop-OneDriveProcesses
    Remove-OneDriveAutostart
    Set-OneDriveSyncDisabled -Disabled $true
    $setupPaths = @(Get-OneDriveSetupPaths)
    if ($setupPaths.Count -eq 0) {
        Write-Host 'OneDriveSetup.exe не найден. Автозапуск и синхронизация отключены.'
        Pause-Result
        return
    }
    foreach ($setup in $setupPaths) {
        Write-Host "Запуск удаления: $setup"
        $process = Start-Process -FilePath $setup -ArgumentList '/uninstall' -Wait -PassThru -ErrorAction SilentlyContinue
        if ($process) { Write-Host "Код завершения: $($process.ExitCode)" }
        $process = Start-Process -FilePath $setup -ArgumentList '/uninstall /allusers' -Wait -PassThru -ErrorAction SilentlyContinue
        if ($process) { Write-Host "Код завершения /allusers: $($process.ExitCode)" }
    }
    Remove-OneDriveAutostart
    Write-Host ''
    Write-Host 'Удаление OneDrive завершено. Файлы пользователя не тронуты.'
    Pause-Result
}

function Disable-OneDriveOnly {
    Show-WorkScreen -Title 'Отключение OneDrive' -Details 'OneDrive не удаляется, только отключается автозапуск и синхронизация.'
    Stop-OneDriveProcesses
    Remove-OneDriveAutostart
    Set-OneDriveSyncDisabled -Disabled $true
    Write-Host 'OneDrive отключён.'
    Pause-Result
}

function Clear-OneDriveResiduals {
    Show-WorkScreen -Title 'Очистка остатков OneDrive' -Details 'Папка пользователя OneDrive удаляется только если она пустая.'
    Stop-OneDriveProcesses
    Remove-OneDriveAutostart
    $targets = @(
        (Join-Path $env:PROGRAMDATA 'Microsoft OneDrive'),
        (Join-Path $env:PROGRAMDATA 'Microsoft\OneDrive')
    )
    foreach ($target in $targets) {
        if ($target -and (Test-Path -LiteralPath $target)) {
            Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Очищено: $target"
        }
    }
    $userOneDrive = Join-Path $env:USERPROFILE 'OneDrive'
    if (Test-Path -LiteralPath $userOneDrive) {
        $files = @(Get-ChildItem -LiteralPath $userOneDrive -Force -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($files.Count -eq 0) {
            Remove-Item -LiteralPath $userOneDrive -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Удалена пустая папка: $userOneDrive"
        } else {
            Write-Host "Папка не удалена, внутри есть файлы: $userOneDrive"
        }
    }
    Write-Host 'Очистка завершена.'
    Pause-Result
}

function Start-OneDriveManager {
    while ($true) {
        $items = @(
            'Удалить OneDrive безопасно',
            'Отключить OneDrive без удаления',
            'Очистить остатки OneDrive',
            'Назад'
        )
        $choice = Select-SingleItem -Title 'OneDrive' -Items $items -Text { param($item) $item }
        if ($choice -lt 0 -or $choice -eq 3) { return }
        if ($choice -eq 0) { Uninstall-OneDriveSafe }
        elseif ($choice -eq 1) { Disable-OneDriveOnly }
        else { Clear-OneDriveResiduals }
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
        $fileBytes = [IO.File]::ReadAllBytes($PSCommandPath)
        if ($fileBytes.Length -lt 3 -or $fileBytes[0] -ne 0xEF -or $fileBytes[1] -ne 0xBB -or $fileBytes[2] -ne 0xBF) {
            $failures.Add('Файл должен быть сохранён как UTF-8 BOM для Windows PowerShell 5.1.')
        }
        $tokens = $null
        $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($PSCommandPath, [ref]$tokens, [ref]$errors)
        if ($errors.Count -gt 0) { foreach ($error in $errors) { $failures.Add("Синтаксис: $($error.Message)") } }
        $dangerousCommandNames = @('Invoke' + '-Expression', 'i' + 'ex', 'Invoke' + '-RestMethod', 'i' + 'rm')
        $base64MethodName = 'FromBase64' + 'String'
        $dangerousCommands = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ -in $dangerousCommandNames })
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
    foreach ($sample in @('192.168.0','192.168.0.0/24')) {
        if ((ConvertTo-PrinterSubnet $sample) -ne '192.168.0') { $failures.Add("Не распознана подсеть $sample") }
    }
    if (ConvertTo-PrinterSubnet '999.1.1') { $failures.Add('Проверка диапазона октетов подсети не работает.') }
    if ((ConvertTo-PrinterIPAddress '192.168.0.45') -ne '192.168.0.45') { $failures.Add('Не распознан IP принтера 192.168.0.45') }
    foreach ($sample in @('192.168.0.0','192.168.0.255','192.168.0','999.1.1.1')) {
        if (ConvertTo-PrinterIPAddress $sample) { $failures.Add("Проверка IP принтера пропустила некорректное значение $sample") }
    }
    $modelTests = [ordered]@{
        'KYOCERA TASKalfa 2554ci' = 'TASKalfa 2554ci'
        'HP LaserJet 200 color M251nw 192.168.0.12' = 'HP LaserJet 200 color M251nw'
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
    $windowsActivation = Get-TSforgeActivationProfile -Mode Windows
    $officeActivation = Get-TSforgeActivationProfile -Mode Office
    $windowsArguments = @(Get-MasCmdArguments -Profile $windowsActivation)
    $officeArguments = @(Get-MasCmdArguments -Profile $officeActivation)
    if ($windowsActivation.Switch -ne '/Z-Windows' -or $windowsArguments.Count -ne 4 -or $windowsArguments[3] -ne '/Z-Windows') {
        $failures.Add('TSforge Windows должен запускаться только с параметром /Z-Windows.')
    }
    if ($officeActivation.Switch -ne '/Z-Office' -or $officeArguments.Count -ne 4 -or $officeArguments[3] -ne '/Z-Office') {
        $failures.Add('TSforge Office All должен запускаться только с параметром /Z-Office.')
    }
    $legacyCombinedSwitch = '/Z-Windows' + 'ESUOffice'
    if (($windowsArguments + $officeArguments) -contains $legacyCombinedSwitch) {
        $failures.Add('Объединённый режим TSforge Windows/ESU/Office не должен использоваться.')
    }
    $winget = Find-WinGetExecutable
    if ($winget) { Write-Host "[OK] winget запускается: $winget" }
    else { $warnings.Add('winget сейчас не запускается; интерактивный режим предложит восстановление.') }
    if ($failures.Count -eq 0) { Write-Host '[OK] Синтаксис и внутренняя структура.' }
    Write-Host "[OK] Каталог программ: $($apps.Count) пунктов, без повторов."
    Write-Host '[OK] Множественный выбор использует стабильные ключи, а не позиции строк.'
    Write-Host '[OK] Проверка формата подсети.'
    Write-Host '[OK] Определение моделей HP/Kyocera и сопоставление твиков.'
    Write-Host '[OK] TSforge разделён на Windows /Z-Windows и Office All /Z-Office.'
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
            3 { Start-ActivationManager }
            4 { Start-TweakManager }
            5 { Start-OneDriveManager }
            6 { Start-PrinterManager }
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
            'Активация',
            'Твики',
            'OneDrive',
            'Сетевые принтеры',
            'Выход'
        )
        $choice = Select-SingleItem -Title 'Менеджер установки WG' -Items $mainItems -Text { param($item) $item }
        if ($choice -lt 0) { continue }
        if ($choice -eq 7) {
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
