#!/usr/bin/env python3
"""
GitHub Repository Creation and Deployment Script
Erstellt automatisch ein neues GitHub Repository und lädt den Code hoch
"""

import os
import sys
import subprocess
import json
import requests
from pathlib import Path

class GitHubDeployer:
    def __init__(self, username="Erwin2307-py", repo_name="Paper_claude"):
        self.username = username
        self.repo_name = repo_name
        self.github_token = None
        self.repo_url = f"https://github.com/{username}/{repo_name}.git"

    def get_github_token(self):
        """GitHub Personal Access Token abfragen"""
        print("🔑 GitHub Personal Access Token benötigt!")
        print("Erstellen Sie einen Token unter: https://github.com/settings/tokens")
        print("Benötigte Berechtigung: 'repo' (Full control of private repositories)")
        print()

        token = input("GitHub Token eingeben: ").strip()
        if not token:
            print("❌ Kein Token eingegeben!")
            sys.exit(1)

        self.github_token = token
        return token

    def create_github_repo(self):
        """Erstellt neues GitHub Repository"""
        print(f"📁 Erstelle GitHub Repository: {self.username}/{self.repo_name}")

        headers = {
            "Authorization": f"token {self.github_token}",
            "Accept": "application/vnd.github.v3+json"
        }

        data = {
            "name": self.repo_name,
            "description": "Streamlit Research Application for Scientific Paper Analysis",
            "private": False,  # Öffentlich für Streamlit Community Cloud
            "has_issues": True,
            "has_projects": True,
            "has_wiki": True,
            "auto_init": False
        }

        response = requests.post(
            "https://api.github.com/user/repos",
            headers=headers,
            json=data
        )

        if response.status_code == 201:
            print("✅ Repository erfolgreich erstellt!")
            return True
        elif response.status_code == 422:
            print("⚠️ Repository existiert bereits - verwende existierendes Repository")
            return True
        else:
            print(f"❌ Fehler beim Erstellen des Repositories: {response.status_code}")
            print(f"Response: {response.text}")
            return False

    def init_git_repo(self):
        """Initialisiert Git Repository lokal"""
        print("🔧 Initialisiere Git Repository...")

        try:
            # Git initialisieren
            subprocess.run(["git", "init"], check=True, cwd=".")
            print("✅ Git Repository initialisiert")

            # Remote hinzufügen
            subprocess.run([
                "git", "remote", "add", "origin", self.repo_url
            ], check=True, cwd=".")
            print(f"✅ Remote origin hinzugefügt: {self.repo_url}")

        except subprocess.CalledProcessError as e:
            print(f"❌ Git Fehler: {e}")
            return False

        return True

    def prepare_files(self):
        """Bereitet Dateien für Deployment vor"""
        print("📝 Bereite Dateien vor...")

        # README.md erstellen
        readme_content = f"""# Streamlit Research Application

Eine umfassende Streamlit-Anwendung für wissenschaftliche Forschung und Papieranalyse.

## Features

🔍 **Paper Search**: PubMed und wissenschaftliche Datenbank-Suche
📊 **Data Analysis**: Gene und SNP Datenanalyse
📧 **Email Integration**: Automatische Email-Berichte
🤖 **AI Analysis**: OpenAI-basierte Papieranalyse
📁 **Excel Management**: Persistente Datenbankverwaltung

## Streamlit Cloud Deployment

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/{self.username}/{self.repo_name})

### Setup

1. **Fork** dieses Repository
2. **Streamlit Community Cloud** besuchen: https://share.streamlit.io/
3. **New app** erstellen und Repository auswählen
4. **Secrets** konfigurieren (siehe unten)

### Required Secrets

Füge folgende Secrets in Streamlit Cloud hinzu:

```toml
[login]
username = "dein_username"
password = "dein_passwort"

[openai]
api_key = "sk-dein_openai_key"

[email]
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "deine@email.com"
sender_password = "dein_app_passwort"
```

## Lokale Installation

```bash
git clone https://github.com/{self.username}/{self.repo_name}.git
cd {self.repo_name}
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Technologie Stack

- **Frontend**: Streamlit
- **Backend**: Python 3.8+
- **AI**: OpenAI GPT Models
- **Data**: Pandas, Excel/CSV
- **PDFs**: PyPDF2, pdfplumber
- **Web Scraping**: Selenium, Scholarly

## Mitwirken

1. Fork das Repository
2. Erstelle einen Feature Branch
3. Committe deine Änderungen
4. Push zum Branch
5. Öffne einen Pull Request

## Lizenz

MIT License - siehe [LICENSE](LICENSE) für Details.
"""

        with open("README.md", "w", encoding="utf-8") as f:
            f.write(readme_content)
        print("✅ README.md erstellt")

        # LICENSE erstellen
        license_content = """MIT License

Copyright (c) 2024 Erwin Schimak

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

        with open("LICENSE", "w", encoding="utf-8") as f:
            f.write(license_content)
        print("✅ LICENSE erstellt")

        return True

    def commit_and_push(self):
        """Committed und pushed alle Dateien"""
        print("📤 Committe und pushe Dateien...")

        try:
            # Alle Dateien hinzufügen
            subprocess.run(["git", "add", "."], check=True, cwd=".")
            print("✅ Dateien zu Git hinzugefügt")

            # Git config setzen falls nötig
            try:
                subprocess.run([
                    "git", "config", "user.email", "erwin.schimak@example.com"
                ], check=True, cwd=".")
                subprocess.run([
                    "git", "config", "user.name", "Erwin Schimak"
                ], check=True, cwd=".")
            except:
                pass  # Config bereits gesetzt

            # Commit erstellen
            subprocess.run([
                "git", "commit", "-m", "Initial commit: Streamlit Research Application"
            ], check=True, cwd=".")
            print("✅ Commit erstellt")

            # Push zum Repository
            subprocess.run([
                "git", "branch", "-M", "main"
            ], check=True, cwd=".")

            subprocess.run([
                "git", "push", "-u", "origin", "main"
            ], check=True, cwd=".")
            print("✅ Code erfolgreich gepusht!")

        except subprocess.CalledProcessError as e:
            print(f"❌ Git Fehler: {e}")
            return False

        return True

    def deploy(self):
        """Hauptfunktion für Deployment"""
        print("🚀 Starte GitHub Deployment...")
        print(f"Repository: {self.username}/{self.repo_name}")
        print("=" * 50)

        # 1. GitHub Token holen
        self.get_github_token()

        # 2. GitHub Repository erstellen
        if not self.create_github_repo():
            return False

        # 3. Git initialisieren
        if not self.init_git_repo():
            return False

        # 4. Dateien vorbereiten
        if not self.prepare_files():
            return False

        # 5. Commit und Push
        if not self.commit_and_push():
            return False

        print("\n" + "=" * 50)
        print("🎉 DEPLOYMENT ERFOLGREICH!")
        print(f"📁 Repository: https://github.com/{self.username}/{self.repo_name}")
        print(f"🚀 Streamlit Deploy: https://share.streamlit.io/")
        print("\n📋 Nächste Schritte:")
        print("1. Besuche https://share.streamlit.io/")
        print("2. Klicke 'New app'")
        print(f"3. Wähle Repository: {self.username}/{self.repo_name}")
        print("4. Main file: streamlit_app.py")
        print("5. Konfiguriere Secrets (siehe README.md)")
        print("6. Deploy!")

        return True

def main():
    """Hauptfunktion"""
    print("🔧 GitHub Deployment Script")
    print("=" * 30)

    # Repository Name anpassen falls gewünscht
    repo_name = input("Repository Name (Enter für 'streamlit-research-app'): ").strip()
    if not repo_name:
        repo_name = "streamlit-research-app"

    deployer = GitHubDeployer(repo_name=repo_name)

    # Warnung anzeigen
    print("\n⚠️ WICHTIG:")
    print("- Stellen Sie sicher, dass Git installiert ist")
    print("- Sie benötigen einen GitHub Personal Access Token")
    print("- Das Repository wird öffentlich erstellt (für Streamlit Community Cloud)")
    print("- Sensible Daten werden über .gitignore ausgeschlossen")

    confirm = input("\nFortfahren? (y/N): ").lower()
    if confirm != 'y':
        print("❌ Abgebrochen")
        return

    # Deployment starten
    success = deployer.deploy()

    if success:
        print("\n✅ Deployment abgeschlossen!")
    else:
        print("\n❌ Deployment fehlgeschlagen!")
        sys.exit(1)

if __name__ == "__main__":
    main()