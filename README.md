# Windows Install "shift+F10"

```
oobe\bypassnro
```

# Windows Activate

```
irm https://get.activated.win | iex
```

# Winget Apps First Install & Update

```
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
iwr https://raw.githubusercontent.com/samocvetov/wginst/main/1.ps1 -OutFile i.ps1; & .\i.ps1
```

# Winget For Admins

```
curl -L https://raw.githubusercontent.com/samocvetov/wginst/main/w.ps1 -o i.ps1 && powershell -ExecutionPolicy Bypass -File i.ps1
```
