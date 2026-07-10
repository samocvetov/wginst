[CmdletBinding()]
param()

$remoteScriptUrl = 'https://raw.githubusercontent.com/samocvetov/wginst/main/print.ps1'
$startedFromWeb = [string]::IsNullOrWhiteSpace($PSCommandPath)
$scriptPath = if ($startedFromWeb) { Join-Path $env:LOCALAPPDATA 'Printer-Manager\Printer-Manager.ps1' } else { $PSCommandPath }

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host 'Requesting administrator privileges...' -ForegroundColor Yellow
    try {
        if ($startedFromWeb) {
            $arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"irm '$remoteScriptUrl' | iex`""
        } else {
            $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        }
        $process = Start-Process -FilePath 'powershell.exe' -Verb RunAs -Wait -PassThru -ArgumentList $arguments
        exit $process.ExitCode
    } catch {
        Write-Host 'Administrator privileges were not granted. The manager was not started.' -ForegroundColor Red
        exit 1
    }
}

Exit code: 0
Wall time: 0.4 seconds
Output:
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:CacheRoot = Join-Path (Split-Path -Parent $scriptPath) 'driver-cache'
$script:HPFallbackUrl = 'https://ftp.hp.com/pub/softlib/software13/printers/UPD/upd-pcl6-win11-x64-8.2.0.26819.zip'
$script:KyoceraUrl = 'https://www.kyoceradocumentsolutions.com.br/content/dam/download-center-americas-cf/br/drivers/drivers/KX_Print_Driver_zip.download.zip'

function Select-ConsoleItem {
    param([string]$Title, [object[]]$Items, [scriptblock]$Label)
    $index = 0
    while ($true) {
        Clear-Host; Write-Host $Title -ForegroundColor Cyan; Write-Host 'Use Up/Down arrows, Enter or Esc.' -ForegroundColor DarkGray
        for ($i = 0; $i -lt $Items.Count; $i++) {
            if ($i -eq $index) {
                $prefix = '> '
                $color = 'Yellow'
            } else {
                $prefix = '  '
                $color = 'White'
            }
            Write-Host $prefix -NoNewline -ForegroundColor $color
            Write-Host (& $Label $Items[$i]) -ForegroundColor $color
        }
        switch ([Console]::ReadKey($true).Key) {
            UpArrow { $index = ($index - 1 + $Items.Count) % $Items.Count }
            DownArrow { $index = ($index + 1) % $Items.Count }
            Enter { return $index }
            Escape { return -1 }
        }
    }
}

function Get-ActiveSubnets {
    Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notmatch '^(127|169\.254)\.' -and $_.PrefixLength -ge 24 } | ForEach-Object { (($_.IPAddress -split '\.')[0..2] -join '.') } | Select-Object -Unique
}

function Get-SnmpText { param([string]$IP)
    $requests=@(
        [byte[]](0x30,0x29,0x02,0x01,0x00,0x04,0x06,0x70,0x75,0x62,0x6C,0x69,0x63,0xA0,0x1C,0x02,0x04,0x01,0x02,0x03,0x04,0x02,0x01,0x00,0x02,0x01,0x00,0x30,0x0E,0x30,0x0C,0x06,0x08,0x2B,0x06,0x01,0x02,0x01,0x01,0x01,0x00,0x05,0x00),
        [byte[]](0x30,0x2C,0x02,0x01,0x00,0x04,0x06,0x70,0x75,0x62,0x6C,0x69,0x63,0xA0,0x1F,0x02,0x04,0x01,0x02,0x03,0x04,0x02,0x01,0x00,0x02,0x01,0x00,0x30,0x11,0x30,0x0F,0x06,0x0B,0x2B,0x06,0x01,0x02,0x01,0x2B,0x05,0x01,0x01,0x10,0x01,0x05,0x00)
    );$result=@();foreach($request in $requests){$udp=[Net.Sockets.UdpClient]::new();try{$udp.Client.ReceiveTimeout=800;$udp.Connect($IP,161);[void]$udp.Send($request,$request.Length);$end=[Net.IPEndPoint]::new([Net.IPAddress]::Any,0);$raw=$udp.Receive([ref]$end);$result+=(-join($raw|Where-Object{$_ -ge 32 -and $_ -le 126}|ForEach-Object{[char]$_}))}catch{}finally{$udp.Close()}};return $result
}

function Get-IppModel { param([string]$IP)
    try {
        $data=[Collections.Generic.List[byte]]::new(); $data.AddRange([byte[]](1,1,0,11,0,0,0,1,1))
        foreach($a in @(@(0x47,'attributes-charset','utf-8'),@(0x48,'attributes-natural-language','en'),@(0x45,'printer-uri',"ipp://$IP/ipp/print"),@(0x44,'requested-attributes','printer-make-and-model'))){$n=[Text.Encoding]::ASCII.GetBytes($a[1]);$v=[Text.Encoding]::ASCII.GetBytes($a[2]);$data.Add([byte]$a[0]);$data.Add([byte]($n.Length -shr 8));$data.Add([byte]$n.Length);$data.AddRange($n);$data.Add([byte]($v.Length -shr 8));$data.Add([byte]$v.Length);$data.AddRange($v)}; $data.Add(3)
        $r=[Net.WebRequest]::Create("http://$IP`:631/ipp/print");$r.Method='POST';$r.ContentType='application/ipp';$r.Timeout=1800;$r.ContentLength=$data.Count;$s=$r.GetRequestStream();$s.Write($data.ToArray(),0,$data.Count);$s.Close();$resp=$r.GetResponse();$m=[IO.MemoryStream]::new();$resp.GetResponseStream().CopyTo($m);$resp.Close(); return [Text.Encoding]::ASCII.GetString($m.ToArray())
    } catch { return $null }
}

function Get-PrinterIdentity { param([string]$IP)
    $texts=@(Get-SnmpText $IP);$ipp=Get-IppModel $IP;if($ipp){$texts+=$ipp};if(Get-Command curl.exe -ErrorAction SilentlyContinue){try{$texts+=((& curl.exe -k -s --connect-timeout 2 --max-time 3 "https://$IP/" 2>$null)-join "`n")}catch{}}
    $model='Model not identified'; foreach($text in $texts){if($text){foreach($pattern in @('(?i)\bTASKalfa\s+[A-Z0-9-]+(?:ci|i|dn|dw)?\b','(?i)\bECOSYS\s+[A-Z0-9-]+\b','(?i)\bHP\s+(?:Color\s+)?(?:LaserJet|OfficeJet|DesignJet|PageWide|Neverstop)[^<\r\n\x00-\x1f]{0,55}')){ $m=[regex]::Match($text,$pattern);if($m.Success){$model=($m.Value -replace '&nbsp;.*$','').Trim();break} };if($model -ne 'Model not identified'){break} } }
    $vendor=if($model -match '(?i)^HP\s'){'HP'}elseif($model -match '(?i)TASKalfa|ECOSYS|^CS\s'){'Kyocera'}else{'Unknown'}
    [pscustomobject]@{IPAddress=$IP;Model=$model;Vendor=$vendor}
}

function Find-NetworkPrinters { param([string[]]$Subnets)
    $found=@(); foreach($subnet in $Subnets){Write-Host "Scanning $subnet.0/24..." -ForegroundColor Cyan;$attempts=foreach($n in 1..254){$ip="$subnet.$n";foreach($port in 9100,631,515){$client=[Net.Sockets.TcpClient]::new();[pscustomobject]@{IP=$ip;Port=$port;Client=$client;Task=$client.BeginConnect($ip,$port,$null,$null)}}};Start-Sleep -Milliseconds 650;$open=foreach($a in $attempts){try{if($a.Task.IsCompleted){$a.Client.EndConnect($a.Task);$a}}catch{}finally{$a.Client.Close()}};foreach($group in ($open|Group-Object IP)){ $d=Get-PrinterIdentity $group.Name;$d|Add-Member NoteProperty PrintServices (($group.Group.Port|ForEach-Object{"TCP $_"})-join ', ');$found+=$d }};$found|Sort-Object{[version]$_.IPAddress}
}

function Get-HPDriverUrl {
    try { $page=Invoke-WebRequest -Uri 'https://support.hp.com/us-en/drivers/hp-universal-print-driver-series-for-windows/503548' -UseBasicParsing -TimeoutSec 12; $m=[regex]::Match($page.Content,'https?[^"''\s]+upd-pcl6-win11-x64-[^"''\s]+\.zip');if($m.Success){return $m.Value} } catch {}
    return $script:HPFallbackUrl
}

function Get-DriverPackage { param([ValidateSet('HP','Kyocera')]$Vendor)
    $url=if($Vendor -eq 'HP'){Get-HPDriverUrl}else{$script:KyoceraUrl};$dir=Join-Path $script:CacheRoot $Vendor;New-Item -ItemType Directory -Force -Path $dir|Out-Null;$file=Join-Path $dir ([IO.Path]::GetFileName(([Uri]$url).AbsolutePath))
    if(-not(Test-Path $file)){Write-Host "Downloading official $Vendor driver..." -ForegroundColor Cyan;Invoke-WebRequest -Uri $url -OutFile $file -UseBasicParsing}
    $bytes=[IO.File]::ReadAllBytes($file);if($bytes.Length -lt 4 -or $bytes[0] -ne 80 -or $bytes[1] -ne 75){Remove-Item $file -Force;throw "$Vendor download is not a ZIP archive. Open the official support page and retry."};return $file
}

function Install-SelectedPrinter { param([psobject]$Device)
    if($Device.Vendor -eq 'Unknown'){throw "Model for $($Device.IPAddress) was not identified. Installation was cancelled."}
    $package=Get-DriverPackage $Device.Vendor;$root=Join-Path $script:CacheRoot ("expanded-$($Device.Vendor)");New-Item -ItemType Directory -Force -Path $root|Out-Null;Expand-Archive -LiteralPath $package -DestinationPath $root -Force
    if($Device.Vendor -eq 'HP'){$inf=Get-ChildItem $root -Filter 'hpcu*.inf' -Recurse|Select-Object -First 1;$driver='HP Universal Printing PCL 6'}else{$inf=Get-ChildItem $root -Filter 'OEMSETUP.INF' -Recurse|Where-Object{$_.FullName -match '\\64bit\\'}|Select-Object -First 1;$driver="Kyocera $($Device.Model) KX"}
    if(-not $inf){throw "Required $($Device.Vendor) INF was not found in the driver package."};Write-Host "Installing $driver..." -ForegroundColor Cyan
    if(-not(Get-PrinterDriver -Name $driver -ErrorAction SilentlyContinue)){& "$env:SystemRoot\System32\pnputil.exe" /add-driver $inf.FullName /install;if($LASTEXITCODE -ne 0){throw "PnPUtil exited with code $LASTEXITCODE"};Add-PrinterDriver -Name $driver}
    $port="IP_$($Device.IPAddress)";if(-not(Get-PrinterPort -Name $port -ErrorAction SilentlyContinue)){Add-PrinterPort -Name $port -PrinterHostAddress $Device.IPAddress}
    $name="$($Device.Model) ($($Device.IPAddress))";$existing=Get-Printer -Name $name -ErrorAction SilentlyContinue;if($existing){Set-Printer -Name $name -PortName $port -DriverName $driver}else{Add-Printer -Name $name -PortName $port -DriverName $driver};Write-Host "Installed: $name" -ForegroundColor Green
}

function Start-PrinterManager {
    while($true){$choice=Select-ConsoleItem 'Printer Manager' @('Install network printer','Remove printer','Exit') {param($x)$x};if($choice -lt 0 -or $choice -eq 2){return}
        if($choice -eq 1){$items=@(Get-Printer|Sort-Object Name)+@([pscustomobject]@{Name='Back to main menu';DriverName=''});$i=Select-ConsoleItem 'Select a printer to remove' $items {param($x)"$($x.Name) [$($x.DriverName)]"};if($i -lt 0 -or $i -eq $items.Count-1){continue};$p=$items[$i];Clear-Host;Write-Host "Name: $($p.Name)`nPort: $($p.PortName)`nDriver: $($p.DriverName)" -ForegroundColor Yellow;if((Read-Host 'Type DELETE to remove this printer') -eq 'DELETE'){Remove-Printer -Name $p.Name;Write-Host 'Printer removed.' -ForegroundColor Green;[Console]::ReadKey($true)|Out-Null};continue}
        $s=Select-ConsoleItem 'Choose subnet' @('Scan active local subnets','Enter /24 subnet manually','Back to main menu') {param($x)$x};if($s -lt 0 -or $s -eq 2){continue};if($s -eq 0){$nets=@(Get-ActiveSubnets)}else{$manual=Read-Host 'Enter first three octets (example 10.130.106)';if($manual -notmatch '^(\d{1,3}\.){2}\d{1,3}$'){Write-Host 'Invalid subnet.' -ForegroundColor Yellow;[Console]::ReadKey($true)|Out-Null;continue};$nets=@($manual)};$devices=@(Find-NetworkPrinters $nets);if(-not $devices){Write-Host 'No printers found.' -ForegroundColor Yellow;[Console]::ReadKey($true)|Out-Null;continue};$i=Select-ConsoleItem 'Select a discovered printer' $devices {param($x)"$($x.IPAddress)  $($x.Model)  [$($x.Vendor)]"};if($i -lt 0){continue};try{Install-SelectedPrinter $devices[$i]}catch{Write-Host $_.Exception.Message -ForegroundColor Red};[Console]::ReadKey($true)|Out-Null
    }
}


Start-PrinterManager
