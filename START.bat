@echo off
chcp 65001 >nul
title Excel Pro Master

echo ====================================
echo      Excel Pro Master v1.0
echo ====================================
echo.

REM Проверяем Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Python не установлен!
    echo Скачайте с python.org
    pause
    exit
)

REM Устанавливаем библиотеки если нужно
pip install -q numpy pandas scipy matplotlib seaborn xlsxwriter openpyxl 2>nul

REM Запускаем программу
python excel_pro_master.py

if errorlevel 1 pause
