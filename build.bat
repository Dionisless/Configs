@echo off
chcp 65001 > nul
echo ============================================================
echo  Сборка compare_tool.exe
echo ============================================================
echo.

:: Проверить наличие Python
python --version 2>nul
if errorlevel 1 (
    echo ОШИБКА: Python не найден. Установите Python 3.8+ с python.org
    pause
    exit /b 1
)

echo.
echo [1/3] Установка зависимостей...
pip install python-docx pyinstaller --quiet
if errorlevel 1 (
    echo ОШИБКА при установке зависимостей.
    pause
    exit /b 1
)

echo.
echo [2/3] Сборка EXE...
pyinstaller ^
    --onefile ^
    --name compare_tool ^
    --collect-all docx ^
    --collect-all lxml ^
    --hidden-import xml.etree.ElementTree ^
    compare_tool.py

if errorlevel 1 (
    echo ОШИБКА при сборке EXE.
    pause
    exit /b 1
)

echo.
echo [3/3] Готово!
echo.
echo Файл: dist\compare_tool.exe
echo.
echo Использование:
echo   - Двойной щелчок — открыть GUI
echo   - Из командной строки: compare_tool.exe таблица.docx конфиг.xml -o отчёт.txt
echo.
pause
