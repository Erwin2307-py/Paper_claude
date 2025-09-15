#!/usr/bin/env python3
"""
Manual Deploy Script - Ohne GitHub API
Erstellt Git Repository und gibt Ihnen die Befehle zum manuellen Upload
"""

import os
import subprocess
import sys

GITHUB_USERNAME = "Erwin2307-py"
REPO_NAME = "Paper_claude"

def init_git_locally():
    """Git lokal initialisieren"""
    print("🔧 Initialisiere Git Repository lokal...")

    try:
        # Prüfe ob .git bereits existiert
        if os.path.exists(".git"):
            print("⚠️ Git Repository bereits initialisiert")
            return True

        # Git initialisieren
        subprocess.run(["git", "init"], check=True, cwd=".")
        print("✅ Git Repository initialisiert")

        # Git config setzen
        subprocess.run(["git", "config", "user.email", "erwin.schimak@example.com"], check=True, cwd=".")
        subprocess.run(["git", "config", "user.name", "Erwin Schimak"], check=True, cwd=".")
        print("✅ Git Konfiguration gesetzt")

        # Alle Dateien hinzufügen
        subprocess.run(["git", "add", "."], check=True, cwd=".")
        print("✅ Dateien zu Git hinzugefügt")

        # Commit erstellen
        subprocess.run([
            "git", "commit", "-m", "🚀 Initial commit: Paper Claude - Streamlit Research Application"
        ], check=True, cwd=".")
        print("✅ Initial Commit erstellt")

        return True

    except subprocess.CalledProcessError as e:
        print(f"❌ Git Fehler: {e}")
        return False

def show_manual_instructions():
    """Zeigt manuelle Anweisungen für GitHub"""
    print("\n" + "=" * 60)
    print("📋 MANUELLE GITHUB ERSTELLUNG")
    print("=" * 60)

    print("\n1️⃣ GITHUB REPOSITORY ERSTELLEN:")
    print("   • Gehen Sie zu: https://github.com/new")
    print(f"   • Repository Name: {REPO_NAME}")
    print("   • Description: Streamlit Research Application")
    print("   • ✅ Public (für Streamlit Cloud)")
    print("   • ❌ NICHT 'Add README' ankreuzen!")
    print("   • 'Create repository' klicken")

    print("\n2️⃣ LOKALEN CODE HOCHLADEN:")
    print("   Führen Sie diese Befehle in der Eingabeaufforderung aus:")
    print()
    print("   git branch -M main")
    print(f"   git remote add origin https://github.com/{GITHUB_USERNAME}/{REPO_NAME}.git")
    print("   git push -u origin main")

    print("\n3️⃣ STREAMLIT CLOUD DEPLOYMENT:")
    print("   • Gehen Sie zu: https://share.streamlit.io/")
    print("   • 'New app' klicken")
    print(f"   • Repository: {GITHUB_USERNAME}/{REPO_NAME}")
    print("   • Branch: main")
    print("   • Main file: streamlit_app.py")
    print("   • Secrets konfigurieren (siehe DEPLOYMENT.md)")

def create_batch_file():
    """Erstellt Batch-Datei für Windows"""
    batch_content = f"""@echo off
echo ========================================
echo   Git Upload zu GitHub
echo ========================================
echo.

echo Repository: {GITHUB_USERNAME}/{REPO_NAME}
echo.

echo 🔧 Setze Main Branch...
git branch -M main

echo 📡 Füge Remote Repository hinzu...
git remote add origin https://github.com/{GITHUB_USERNAME}/{REPO_NAME}.git

echo 📤 Lade Code zu GitHub hoch...
git push -u origin main

if errorlevel 1 (
    echo.
    echo ❌ Upload fehlgeschlagen!
    echo Mögliche Ursachen:
    echo - Repository wurde nicht auf GitHub erstellt
    echo - Falsche Repository-URL
    echo - Keine Internet-Verbindung
    echo.
    pause
    exit /b 1
)

echo.
echo ✅ CODE ERFOLGREICH HOCHGELADEN!
echo.
echo 📋 Nächste Schritte:
echo 1. Besuche https://share.streamlit.io/
echo 2. Klicke 'New app'
echo 3. Repository: {GITHUB_USERNAME}/{REPO_NAME}
echo 4. Main file: streamlit_app.py
echo 5. Secrets konfigurieren
echo 6. Deploy!
echo.
pause
"""

    with open("upload_to_github.bat", "w", encoding="utf-8") as f:
        f.write(batch_content)

    print(f"✅ Batch-Datei erstellt: upload_to_github.bat")

def main():
    """Hauptfunktion"""
    print("🚀 Manual GitHub Deployment")
    print("=" * 40)
    print(f"Repository: {GITHUB_USERNAME}/{REPO_NAME}")
    print("=" * 40)

    # Git lokal initialisieren
    if not init_git_locally():
        print("❌ Git Initialisierung fehlgeschlagen!")
        return False

    # Batch-Datei erstellen
    create_batch_file()

    # Manuelle Anweisungen zeigen
    show_manual_instructions()

    print("\n" + "=" * 60)
    print("🎯 ZUSAMMENFASSUNG")
    print("=" * 60)
    print("1. ✅ Git Repository lokal initialisiert")
    print("2. ✅ Batch-Datei erstellt: upload_to_github.bat")
    print("3. 📋 Manuelle Schritte angezeigt")
    print()
    print("🚀 NÄCHSTER SCHRITT:")
    print("1. Erstellen Sie das Repository auf GitHub (siehe Anweisungen oben)")
    print("2. Doppelklick auf 'upload_to_github.bat'")
    print("   ODER führen Sie die Git-Befehle manuell aus")

    return True

if __name__ == "__main__":
    main()