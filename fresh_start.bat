@echo off
chcp 65001 >nul
echo ========================================
echo   Paper Claude - Fresh GitHub Start
echo ========================================
echo.

echo SCHRITT 1: Aktuelles Git Repository loeschen...
if exist ".git" (
    rmdir /s /q ".git"
    echo Git History geloescht
)

echo.
echo SCHRITT 2: Neues Git Repository initialisieren...
git init
git config user.email "erwin.schimak@example.com"
git config user.name "Erwin Schimak"

echo.
echo SCHRITT 3: Dateien vorbereiten...
git add .
git commit -m "Initial commit: Paper Claude - Streamlit Research Application"

echo.
echo ========================================
echo   GITHUB REPOSITORY ERSTELLEN
echo ========================================
echo.
echo 1. Gehen Sie zu: https://github.com/new
echo 2. Repository Name: Paper_claude
echo 3. Description: Streamlit Research Application
echo 4. Public (fuer Streamlit Cloud)
echo 5. NICHT "Add README" ankreuzen!
echo 6. "Create repository" klicken
echo.
echo Druecken Sie eine Taste wenn Repository erstellt...
pause

echo.
echo SCHRITT 4: Code zu GitHub hochladen...
git branch -M main
git remote add origin https://github.com/Erwin2307-py/Paper_claude.git
git push -u origin main

if errorlevel 1 (
    echo.
    echo FEHLER beim Upload!
    echo Pruefen Sie die Repository-URL und Internetverbindung
    pause
    exit /b 1
)

echo.
echo ========================================
echo   ERFOLGREICH!
echo ========================================
echo.
echo Repository: https://github.com/Erwin2307-py/Paper_claude
echo.
echo STREAMLIT DEPLOYMENT:
echo 1. https://share.streamlit.io/
echo 2. "New app" klicken
echo 3. Repository: Erwin2307-py/Paper_claude
echo 4. Main file: streamlit_app.py
echo 5. Secrets konfigurieren
echo.
pause