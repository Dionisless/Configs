@echo off
chcp 65001 > nul
setlocal

echo Building digsi_readable.exe ...
pyinstaller --onefile --name digsi_readable digsi_readable.py
if errorlevel 1 (
  echo Build failed
  exit /b 1
)

echo Done: dist\digsi_readable.exe
