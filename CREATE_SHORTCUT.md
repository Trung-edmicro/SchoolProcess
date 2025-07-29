# School Process UI - Desktop Shortcut Creator

## Tạo shortcut trên Desktop

### Windows
```powershell
# Copy và paste vào PowerShell (Run as Administrator)
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = "$DesktopPath\School Process UI.lnk"
$TargetPath = "$(Get-Location)\start_ui.bat"
$IconLocation = "$(Get-Location)\start_ui.bat"

$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $TargetPath
$Shortcut.WorkingDirectory = Get-Location
$Shortcut.IconLocation = $IconLocation
$Shortcut.Description = "School Process - UI Application"
$Shortcut.Save()

Write-Host "Shortcut created: $ShortcutPath"
```

### Hoặc tạo thủ công:
1. Right-click trên Desktop
2. New → Shortcut
3. Location: `[path_to_project]\start_ui.bat`
4. Name: `School Process UI`

---

### Linux/Mac
```bash
# Tạo script launcher
cat > ~/Desktop/school-process-ui.sh << 'EOF'
#!/bin/bash
cd "$(dirname "$0")/[path_to_project]"
python3 main_ui.py
EOF

chmod +x ~/Desktop/school-process-ui.sh
```
