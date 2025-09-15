@echo off
chcp 65001 >nul
echo ========================================
echo   Paper Claude - GitHub Setup
echo ========================================
echo.

echo Repository: Erwin2307-py/Paper_claude
echo.

echo 1. Git Repository initialisieren...
if not exist ".git" (
    git init
    git config user.email "erwin.schimak@example.com"
    git config user.name "Erwin Schimak"
    echo Git Repository initialisiert
) else (
    echo Git Repository bereits vorhanden
)

echo.
echo 2. Dateien zu Git hinzufuegen...
git add .
git commit -m "Initial commit: Paper Claude - Streamlit Research Application"

echo.
echo ========================================
echo   MANUELLE GITHUB ERSTELLUNG
echo ========================================
echo.
echo SCHRITT 1: GitHub Repository erstellen
echo - Gehen Sie zu: https://github.com/new
echo - Repository Name: Paper_claude
echo - Description: Streamlit Research Application
echo - Public (fuer Streamlit Cloud)
echo - NICHT "Add README" ankreuzen!
echo - "Create repository" klicken
echo.
echo SCHRITT 2: Druecken Sie eine Taste wenn Repository erstellt...
pause

echo.
echo 3. Code zu GitHub hochladen...
git branch -M main
git remote add origin https://github.com/Erwin2307-py/Paper_claude.git
git push -u origin main

if errorlevel 1 (
    echo.
    echo FEHLER beim Upload!
    echo Pruefen Sie:
    echo - Repository wurde auf GitHub erstellt
    echo - Internet-Verbindung
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   ERFOLGREICH HOCHGELADEN!
echo ========================================
echo.
echo Repository: https://github.com/Erwin2307-py/Paper_claude
echo.
echo NAECHSTE SCHRITTE - STREAMLIT DEPLOYMENT:
echo 1. Besuche https://share.streamlit.io/
echo 2. Klicke "New app"
echo 3. Repository: Erwin2307-py/Paper_claude
echo 4. Branch: main
echo 5. Main file: streamlit_app.py
echo 6. Secrets konfigurieren (siehe DEPLOYMENT.md)
echo 7. Deploy!
echo.
pause