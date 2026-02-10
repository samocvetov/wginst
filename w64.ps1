$Ver="6.4"; Clear-Host; Write-Host "=== WINGET AUTO-INSTALLER v$Ver ===" -F Cyan; Write-Host ""
$apps = @("7zip.7zip", "Google.Chrome", "Yandex.Browser", "DominikReichl.KeePass", "Notepad++.Notepad++", "Telegram.TelegramDesktop", "9NKSQGP7F2NH", "Yandex.Messenger", "Zoom.Zoom", "RustDesk.RustDesk", "AnyDesk.AnyDesk", "WireGuard.WireGuard", "Termius.Termius", "Mikrotik.Winbox", "angryziber.AngryIPScanner", "alexx2000.DoubleCommander", "9NV4BS3L1H4S", "PDFgear.PDFgear", "VideoLAN.VLC", "AdrienAllard.FileConverter", "XPDDT99J9GKB5C", "WinDirStat.WinDirStat", "Piriform.Recuva", "ventoy.ventoy")
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
            Copy-Item $lnkS $lnkD -Force; Write-Host "   [+] Shortcut: $name" -F DarkGray
        } catch { Write-Host "   [!] Shortcut failed" -F Red }
    }
}
Write-Host "--- Checking updates ---" -F Cyan
$raw = winget upgrade --accept-source-agreements; $lines = $raw | Select-String '^\S+' | Select -Skip 2
foreach ($l in $lines) {
    $c = $l.ToString() -split '\s{2,}'; if ($c.Count -lt 2) { continue }
    $id = $c[1].Trim(); if ($id -match "\s") { $id = $id.Split(" ")[0] }
    if ($id -and $id -ne "ID" -and $id -notlike "-*") {
        $msg = "Update " + $c[0].Trim() + " ($id)? [y/n]"
        $ans = Read-Host $msg
        if ($ans -eq 'y') { winget upgrade --id $id --silent --force --accept-source-agreements --accept-package-agreements }
    }
}
Write-Host "`n--- Installing ---" -F Cyan
$inst = winget list --accept-source-agreements | Out-String
foreach ($app in $apps) {
    $dName = $app
    if ($fNames.ContainsKey($app)) { $dName = $fNames[$app] }

    if ($inst -like "*$app*") { 
        Write-Host "[ skip ] $dName" -F Gray 
        continue 
    }
    $msg = "Install " + $dName + "? [y/n]"
    $ans = Read-Host $msg
    if ($ans -eq 'y') {
        Write-Host "Processing $dName..." -NoNewline
        $p = Start-Process winget -Args "install --id $app --silent --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru
        if ($p.ExitCode -eq 0) { Write-Host "`r[ ok ] $dName           " -F Green; Add-Shortcut $app }
        else { Write-Host "`r[FAIL] $dName ($($p.ExitCode))" -F Red }
    }
}
Write-Host "`nDone!" -F Cyan; Start-Sleep 3
