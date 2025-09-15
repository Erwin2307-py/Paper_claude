@echo off
echo =========================================
echo    GitHub Deployment Script (Windows)
echo =========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python ist nicht installiert oder nicht im PATH!
    echo Bitte installieren Sie Python von https://python.org
    pause
    exit /b 1
)

REM Check if Git is installed
git --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Git ist nicht installiert oder nicht im PATH!
    echo Bitte installieren Sie Git von https://git-scm.com
    pause
    exit /b 1
)

echo ✅ Python und Git sind verfügbar
echo.

REM Install required Python packages
echo 📦 Installiere Python-Abhängigkeiten...
pip install requests >nul 2>&1

echo 🚀 Starte Deployment-Script...
echo.

REM Run the Python deployment script
python deploy_to_github.py

echo.
echo Deployment abgeschlossen!
pause