@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "automation.py" (
  echo ERROR: Missing automation.py. This package is incomplete; automation.py is missing.
  pause
  exit /b 1
)
if exist ".venv\Scripts\pythonw.exe" (
  start "" ".venv\Scripts\pythonw.exe" "%~dp0gui.py"
  exit /b 0
)
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" "%~dp0gui.py"
  pause
  exit /b %errorlevel%
)
echo ERROR: Python virtual environment was not found.
echo Run 01_INSTALL.bat from the original package first.
pause
exit /b 1
