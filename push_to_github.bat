@echo off
chcp 65001 >nul
color 0A

echo ╔══════════════════════════════════════════════════════╗
echo ║         GITHUB PUSH SCRIPT FOR EXCEL PRO MASTER     ║
echo ╚══════════════════════════════════════════════════════╝
echo.

REM Initialize git if not already initialized
if not exist ".git" (
    echo [*] Initializing Git repository...
    git init
    echo.
)

REM Add all files
echo [*] Adding files to Git...
git add .
echo.

REM Create commit
echo [*] Creating commit...
git commit -m "🚀 Excel Pro Master - Space Edition | Professional Statistical Analysis Tool"
echo.

REM Ask for GitHub repository URL
echo ══════════════════════════════════════════════════════
echo.
echo Enter your GitHub repository URL:
echo Example: https://github.com/yourusername/excel-pro-master.git
echo.
set /p repo_url="GitHub URL: "

if "%repo_url%"=="" (
    echo [ERROR] No URL provided!
    pause
    exit /b 1
)

REM Add remote origin
echo.
echo [*] Adding remote origin...
git remote remove origin 2>nul
git remote add origin %repo_url%
echo.

REM Push to GitHub
echo [*] Pushing to GitHub...
git branch -M main
git push -u origin main

if %errorlevel% equ 0 (
    echo.
    echo ╔══════════════════════════════════════════════════════╗
    echo ║              ✅ SUCCESSFULLY PUSHED TO GITHUB!       ║
    echo ╚══════════════════════════════════════════════════════╝
    echo.
    echo Your repository is now available at:
    echo %repo_url%
) else (
    echo.
    echo [ERROR] Failed to push to GitHub!
    echo.
    echo Possible solutions:
    echo 1. Check your internet connection
    echo 2. Make sure the repository exists on GitHub
    echo 3. Check your GitHub credentials
    echo 4. Try: git config --global credential.helper manager
)

echo.
pause
