Windows Install "shift+F10"

```
oobe\bypassnro
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

Winget For Admins

```
curl -L https://raw.githubusercontent.com/samocvetov/wginst/main/w.ps1 -o i.ps1 && powershell -ExecutionPolicy Bypass -File i.ps1
```
