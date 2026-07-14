Windows Install "shift+F10"

```
oobe\bypassnro
start ms-cxh:localonly
```
or
```
iwr https://raw.githubusercontent.com/samocvetov/wginst/main/o.ps1 -OutFile $env:TEMP\o.ps1;Set-ExecutionPolicy Bypass -Scope Process -Force;& $env:TEMP\o.ps1
```
or
```
powershell -ep bypass -c "iwr https://raw.githubusercontent.com/samocvetov/wginst/main/o.ps1 -OutFile $env:TEMP\o.ps1;& $env:TEMP\o.ps1"
```

Winget Apps First Install & Update

```
irm s.id/smcwg | iex
```
or
```
irm https://raw.githubusercontent.com/samocvetov/wginst/main/1.ps1 | iex
```

Printer manager

```
irm https://raw.githubusercontent.com/samocvetov/wginst/main/p.ps1 | iex
```

Winget For Admins

```
curl -L https://raw.githubusercontent.com/samocvetov/wginst/main/s.ps1 -o s.ps1 && powershell -ExecutionPolicy Bypass -File s.ps1
```
```
irm https://raw.githubusercontent.com/samocvetov/wginst/main/s.ps1 | iex
```
```
irm s.id/sysop | iex
```
