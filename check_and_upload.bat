@echo off
chcp 65001 >nul
echo ========================================
echo   GitHub Repository Check und Upload
echo ========================================
echo.

echo Repository pruefen: Erwin2307-py/Paper_claude
echo.

echo WICHTIG: Stellen Sie sicher, dass das Repository existiert!
echo.
echo Moegliche Namen - waehlen Sie den richtigen:
echo 1. Paper_claude
echo 2. Paper-claude
echo 3. paper_claude
echo 4. search (falls Sie das bestehende nutzen wollen)
echo.

set /p repo_name="Geben Sie den EXAKTEN Repository-Namen ein: "

if "%repo_name%"=="" (
    echo Kein Name eingegeben!
    pause
    exit /b 1
)

echo.
echo Verwende Repository: Erwin2307-py/%repo_name%
echo Repository-URL: https://github.com/Erwin2307-py/%repo_name%
echo.

echo Ist dies korrekt? (j/n)
set /p confirm="> "

if /i not "%confirm%"=="j" (
    echo Abgebrochen.
    pause
    exit /b 1
)

echo.
echo Remote Repository setzen...
git remote remove origin 2>nul
git remote add origin https://github.com/Erwin2307-py/%repo_name%.git

echo.
echo Teste Verbindung zum Repository...
git ls-remote origin >nul 2>&1

if errorlevel 1 (
    echo.
    echo FEHLER: Repository nicht gefunden!
    echo.
    echo Bitte pruefen Sie:
    echo 1. Repository wurde auf GitHub erstellt
    echo 2. Repository Name ist korrekt: %repo_name%
    echo 3. Repository ist Public
    echo.
    echo Aktuelle Repository-URL:
    echo https://github.com/Erwin2307-py/%repo_name%
    echo.
    pause
    exit /b 1
)

echo Repository gefunden! Lade Code hoch...
git push -u origin main

if errorlevel 1 (
    echo.
    echo Upload fehlgeschlagen!
    echo Versuche force push...
    git push -u origin main --force
)

if errorlevel 1 (
    echo.
    echo FEHLER: Upload immer noch fehlgeschlagen!
    echo.
    echo Moegliche Loesungen:
    echo 1. Repository loeschen und neu erstellen
    echo 2. Anderen Repository-Namen waehlen
    echo 3. Repository Settings pruefen
    pause
    exit /b 1
)

echo.
echo ========================================
echo   ERFOLGREICH HOCHGELADEN!
echo ========================================
echo.
echo Repository: https://github.com/Erwin2307-py/%repo_name%
echo.
echo STREAMLIT DEPLOYMENT:
echo 1. https://share.streamlit.io/
echo 2. "New app"
echo 3. Repository: Erwin2307-py/%repo_name%
echo 4. Main file: streamlit_app.py
echo.
pause