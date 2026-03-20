$ws = New-Object -ComObject WScript.Shell
$startup = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup', 'start_eprice.lnk')
$sc = $ws.CreateShortcut($startup)
$sc.TargetPath = 'D:\AI_Research\eprice\start_eprice.bat'
$sc.WorkingDirectory = 'D:\AI_Research\eprice'
$sc.Save()
Write-Host "Shortcut created at: $startup"
