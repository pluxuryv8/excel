@echo off
chcp 65001 >nul
echo.
echo ============================================
echo  Excel Analytics PRO - Windows Launcher
echo ============================================
echo.

REM Проверяем Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не установлен!
    echo Скачай: https://www.python.org/downloads/
    pause
    exit
)

echo Python найден!
echo.

REM Устанавливаем зависимости (если нужно)
echo Проверяем зависимости...
pip install -q numpy pandas scipy matplotlib xlsxwriter openpyxl >nul 2>&1

echo.
echo ============================================
echo  Выбери режим запуска:
echo ============================================
echo.
echo 1. Консольный режим (самый простой)
echo 2. Веб-интерфейс (в браузере)
echo 3. Прямой запуск (с указанием файлов)
echo.

set /p mode="Твой выбор (1/2/3): "

if "%mode%"=="1" goto console
if "%mode%"=="2" goto web
if "%mode%"=="3" goto direct

:console
echo.
echo ============================================
echo  КОНСОЛЬНЫЙ РЕЖИМ
echo ============================================
echo.
python SUPER_EASY.py
goto end

:web
echo.
echo ============================================
echo  ВЕБ-ИНТЕРФЕЙС
echo ============================================
echo.
echo Запускаю веб-сервер...
echo Откроется браузер по адресу: http://localhost:5000
echo.
echo Для остановки нажми Ctrl+C
echo.
pip install -q flask >nul 2>&1
python web_interface.py
goto end

:direct
echo.
echo ============================================
echo  ПРЯМОЙ ЗАПУСК
echo ============================================
echo.
echo Перетащи TXT файлы с данными в это окно и нажми Enter.
echo Или укажи путь к файлам через пробел.
echo.
set /p files="Файлы: "
if "%files%"=="" (
    echo Нет файлов!
    pause
    exit
)
python report.py %files%
goto end

:end
echo.
echo ============================================
echo  Готово! Результат в папке out/report_pro.xlsx
echo ============================================
pause
