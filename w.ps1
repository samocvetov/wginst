$Ver="6.3.1"; Clear-Host; Write-Host "=== WINGET AUTO-INSTALLER v$Ver ===" -F Cyan; Write-Host ""

$apps = @("7zip.7zip", "Notepad++.Notepad++", "RustDesk.RustDesk", "AnyDesk.AnyDesk", "VideoLAN.VLC", "PDFgear.PDFgear", "Google.Chrome", "Telegram.TelegramDesktop", "Zoom.Zoom", "Yandex.Browser", "Yandex.Messenger", "AdrienAllard.FileConverter", "alexx2000.DoubleCommander", "WinDirStat.WinDirStat", "Piriform.Recuva", "DominikReichl.KeePass", "ventoy.ventoy", "Termius.Termius", "WireGuard.WireGuard", "Mikrotik.Winbox", "REALiX.HWiNFO", "CPUID.CPU-Z", "TechPowerUp.GPU-Z", "angryziber.AngryIPScanner", "9NKSQGP7F2NH", "9NV4BS3L1H4S", "XPDDT99J9GKB5C")
$fNames = @{ "9NKSQGP7F2NH"="WhatsApp"; "9NV4BS3L1H4S"="QuickLook"; "XPDDT99J9GKB5C"="Samsung Magician" }

function Add-Shortcut ($Id) {
    $exes = @{ "ventoy.ventoy"="Ventoy2Disk.exe" }; $skip = @("angryziber.AngryIPScanner")
    if ($Id -notmatch "\." -or $skip -contains $Id) { return }
    $sMenu = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"; $desk = [Environment]::GetFolderPath("Desktop")
    $links = "$env:LOCALAPPDATA\Microsoft\WinGet\Links"; $pkgs = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages"
    
    $file = if ($exes[$Id]) { $exes[$Id] } else { "*$($Id.Split('.')[-1])*.exe" }
    $target = Get-ChildItem $links -Filter $file -EA SilentlyContinue | Select -First 1
    if (!$target) {
        $dir = Get-ChildItem $pkgs -Filter "${Id}*" -Dir -EA SilentlyContinue | Sort LastWriteTime -Desc | Select -First 1
        if ($dir) { $target = Get-ChildItem $dir.FullName -Filter $file -Recurse -EA SilentlyContinue | Select -First 1 }
    }

    if ($target) {
        $name = $target.BaseName; if ($name -eq "Ventoy2Disk") { $name = "Ventoy" }
        $lnkS = "$sMenu\$name.lnk"; $lnkD = "$desk\$name.lnk"; $real = $target.FullName
        try { if ($target.LinkType -eq 'SymbolicLink') { $t = $target.Target; if (![IO.Path]::IsPathRooted($t)) { $t = Join-Path $target.DirectoryName $t }; $real = (Get-Item $t).FullName } } catch {}
        if (Test-Path $lnkS) { rm $lnkS -Force }
        try {
            $ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut($lnkS)
            $s.TargetPath = $target.FullName; $s.WorkingDirectory = $target.DirectoryName; $s.IconLocation = "$real,0"; $s.Save()
            Copy-Item $lnkS $lnkD -Force; Write-Host "   [+] Shortcut created: $name" -F DarkGray
        } catch { Write-Host "   [!] Shortcut failed" -F Red }
    }
}

Write-Host "--- Checking updates ---" -F Cyan
$raw = winget upgrade --accept-source-agreements; $lines = $raw | Select-String '^\S+' | Select -Skip 2
foreach ($l in $lines) {
    $c = $l.ToString() -split '\s{2,}'; if ($c.Count -lt 2) { continue }
    $id = $c[1].Trim(); if ($id -match "\s") { $id = $id.Split(" ")[0] }
    if ($id -and $id -ne "ID" -and $id -notlike "-*") {
        $ans = Read-Host "Update $id? [y/n]"
        if ($ans -eq 'y') { winget upgrade --id $id -h --force --accept-source-agreements --accept-package-agreements }
    }
}

Write-Host "`n--- Installing ---" -F Cyan
$inst = winget list --accept-source-agreements | Out-String
foreach ($app in $apps) {
    if ($inst -like "*$app*") { Write-Host "[SKIP] $app" -F Gray; continue }
    
    # Исправленное определение имени
    $dName = $app; if ($fNames.ContainsKey($app)) { $dName = $fNames[$app] }
    
    $ans = Read-Host "Install $dName? [y/n]"
    if ($ans -eq 'y') {
        Write-Host "Installing $dName..." -NoNewline
        $p = Start-Process winget -Args "install --id $app -h --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru
        if ($p.ExitCode -eq 0) { Write-Host "`r[ OK ] $dName           " -F Green; Add-Shortcut $app }
        else { Write-Host "`r[FAIL] $dName ($($p.ExitCode))" -F Red }
    }
}
Write-Host "`nDone!" -F Cyan; Start-Sleep 3
