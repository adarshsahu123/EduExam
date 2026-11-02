@echo off
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python is not installed. Install it from python.org, then try again.
  pause
  exit /b
)
python exam.py
pause
