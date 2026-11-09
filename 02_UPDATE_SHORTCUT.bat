@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo HCE Oracle - update the desktop shortcut
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0update_hce_shortcut.ps1"
if errorlevel 1 (
  echo Failed to update shortcut. You may still use 10_START_GUI.bat.
  pause
  exit /b 1
)
echo Desktop shortcut updated.
pause
