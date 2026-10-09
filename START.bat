@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo === Проверка библиотек ===
python -c "import requests, pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Устанавливаю библиотеки, подождите...
    python -m pip install --upgrade pip >nul
    python -m pip install requests pandas openpyxl
)

echo === Запуск! ===
python script.py