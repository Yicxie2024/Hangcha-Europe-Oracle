@echo off
setlocal
cd /d "%~dp0"
if errorlevel 1 (
  echo ERROR: Cannot access the application folder.
  pause
  exit /b 1
)
if not exist "%~dp0install_runtime.ps1" (
  echo ERROR: install_runtime.ps1 was not found.
  echo Please copy both files from the patch into the same application folder.
  pause
  exit /b 1
)
if not exist "%~dp0gui.py" (
  echo ERROR: gui.py was not found. Install this patch into the EXISTING HCE Oracle application folder.
  pause
  exit /b 1
)
echo HCE Oracle - One-click setup
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_runtime.ps1"
set "SETUP_RESULT=%ERRORLEVEL%"
if not "%SETUP_RESULT%"=="0" (
  echo.
  echo Installation was not completed. Read the error above.
  echo For assistance, send the installer log from the app folder to your administrator.
  pause
  exit /b %SETUP_RESULT%
)
echo.
echo Setup finished. The GUI should open automatically.
exit /b 0
