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
irm https://clck.ru/3Rk6CD | iex
```

```
irm https://bit.ly/3OjPQfo | iex
```

```
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
irm https://raw.githubusercontent.com/samocvetov/wginst/main/1.ps1 | iex
```

# Winget For Admins

```
curl -L https://raw.githubusercontent.com/samocvetov/wginst/main/w.ps1 -o i.ps1 && powershell -ExecutionPolicy Bypass -File i.ps1
```
