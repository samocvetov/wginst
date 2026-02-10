$apps = @("7zip.7zip", "Notepad++.Notepad++", "RustDesk.RustDesk", "AnyDesk.AnyDesk", "VideoLAN.VLC", "PDFgear.PDFgear", "Google.Chrome", "Telegram.TelegramDesktop", "Zoom.Zoom", "Yandex.Browser", "Yandex.Messenger", "AdrienAllard.FileConverter", "alexx2000.DoubleCommander", "WinDirStat.WinDirStat", "Piriform.Recuva", "PowerSoftware.AnyBurn", "qBittorrent.qBittorrent", "9NKSQGP7F2NH", "XPDDT99J9GKB5C", "Paddington.QuickLook", "DominikReichl.KeePass", "ventoy.ventoy", "Termius.Termius", "TimKosse.FileZilla.Client", "WireGuard.WireGuard", "REALiX.HWiNFO", "CPUID.CPU-Z", "TechPowerUp.GPU-Z", "angryziber.AngryIPScanner")
$friendlyNames = @{ "9NKSQGP7F2NH" = "WhatsApp"; "XPDDT99J9GKB5C" = "Netflix" }
Write-Host "[Checking Updates]" -F Cyan
$lines = winget upgrade --accept-source-agreements | Select-String -Pattern '^\S+' | Select-Object -Skip 2
foreach ($line in $lines) {
    $col = $line.ToString() -split '\s{2,}'
    if ($col.Count -ge 2 -and $col[1] -ne "ID" -and $col[1] -notlike "---*") {
        $id = ($col[1] -split "\s")[0]
        if ((Read-Host "Update $($col[0])? [y/n]") -eq 'y') {
            winget upgrade --id "$id" --silent --force --accept-source-agreements --accept-package-agreements
        }
    }
}
Write-Host "[Installing New Packages]" -F Cyan
$installed = (winget list --accept-source-agreements | Out-String)
foreach ($app in $apps) {
    if ($installed -like "*$app*") {
        Write-Host "[SKIP] $app" -F Gray
        continue
    }
    $displayName = if ($friendlyNames.ContainsKey($app)) { $friendlyNames[$app] } else { $app }
    if ((Read-Host "Install $displayName? [y/n]") -eq 'y') {
        Write-Host "Installing $displayName..." -NoNewline
        $p = Start-Process winget -ArgumentList "install --id $app --silent --accept-source-agreements --accept-package-agreements" -NoNewWindow -Wait -PassThru
        if ($p.ExitCode -eq 0) { Write-Host "`r[ OK ] $displayName" -F Green } 
        else { Write-Host "`r[FAIL] $displayName (Code: $($p.ExitCode))" -F Red }
    }
}
Write-Host "Done!" -F Cyan
if (Test-Path $PSCommandPath) { Remove-Item $PSCommandPath -Force }
