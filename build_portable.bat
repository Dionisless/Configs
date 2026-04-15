@echo off
chcp 65001 > nul
setlocal

echo [1/3] Build digsi_readable.exe...
pyinstaller --onefile --name digsi_readable.exe digsi_readable.py
if errorlevel 1 (
  echo Build failed
  exit /b 1
)

echo [2/3] Build optional pipeline.exe (if Node/pkg configured)...
if exist pipeline.js (
  call npm run build:pipeline
)

echo [3/3] Assemble portable bundle...
python tools\build_portable.py --output-dir portable
if errorlevel 1 (
  echo Portable bundle build failed
  exit /b 1
)

echo Done. Portable zip: portable\digsi_readable-portable.zip
