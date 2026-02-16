Windows Install "shift+F10"

```
oobe\bypassnro
```
or
```
iwr https://raw.githubusercontent.com/samocvetov/wginst/main/oo.ps1 -OutFile $env:TEMP\oo.ps1;Set-ExecutionPolicy Bypass -Scope Process -Force;& $env:TEMP\oo.ps1
```

Winget Apps First Install & Update

```
irm https://s.id/smcwg | iex
```
or
```
irm https://raw.githubusercontent.com/samocvetov/wginst/main/1.ps1 | iex
```

Winget For Admins

```
curl -L https://raw.githubusercontent.com/samocvetov/wginst/main/w.ps1 -o i.ps1 && powershell -ExecutionPolicy Bypass -File i.ps1
```
