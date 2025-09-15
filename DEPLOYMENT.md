# 🚀 Deployment Anleitung

Diese Anleitung erklärt, wie Sie Ihre Streamlit-Anwendung automatisch auf GitHub hochladen und in Streamlit Cloud bereitstellen.

## 🛠️ Voraussetzungen

### 1. Software installieren
- **Python 3.8+**: https://python.org
- **Git**: https://git-scm.com
- **GitHub Account**: https://github.com

### 2. GitHub Personal Access Token erstellen
1. Gehen Sie zu: https://github.com/settings/tokens
2. Klicken Sie auf "Generate new token (classic)"
3. Name: `streamlit-deployment`
4. Berechtigung auswählen: ✅ **repo** (Full control of private repositories)
5. Token generieren und **sicher speichern**!

## 🚀 Automatisches Deployment

### Option 1: Windows Batch-Datei (Empfohlen)
```bash
# Doppelklick auf:
deploy_to_github.bat
```

### Option 2: Python-Script direkt
```bash
python deploy_to_github.py
```

### Was das Script macht:
1. ✅ Erstellt neues GitHub Repository
2. ✅ Initialisiert Git lokal
3. ✅ Erstellt README.md und LICENSE
4. ✅ Committed alle Dateien
5. ✅ Pusht Code zu GitHub

## 🌐 Streamlit Cloud Deployment

### 1. Nach GitHub Upload
1. Besuchen Sie: https://share.streamlit.io/
2. Mit GitHub Account anmelden
3. "New app" klicken

### 2. App konfigurieren
- **Repository**: `Erwin2307-py/streamlit-research-app`
- **Branch**: `main`
- **Main file path**: `streamlit_app.py`

### 3. Secrets konfigurieren
Klicken Sie auf "Advanced settings" und fügen Sie folgende Secrets hinzu:

```toml
[login]
username = "ihr_username"
password = "ihr_sicheres_passwort"

[openai]
api_key = "sk-ihr_openai_api_key"

[email]
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "ihre@email.com"
sender_password = "ihr_gmail_app_passwort"

[excel]
template_path = "data/master_papers.xlsx"
auto_create_sheets = true
max_sheets = 50
```

### 4. App deployen
- "Deploy!" klicken
- Warten auf Build (ca. 2-5 Minuten)
- Ihre App ist live! 🎉

## 🔧 Manuelle Alternative

Falls das automatische Script nicht funktioniert:

```bash
# 1. Git initialisieren
git init

# 2. Remote hinzufügen
git remote add origin https://github.com/Erwin2307-py/IHR_REPO_NAME.git

# 3. Dateien hinzufügen
git add .

# 4. Commit erstellen
git commit -m "Initial commit: Streamlit Research Application"

# 5. Push zu GitHub
git branch -M main
git push -u origin main
```

## 📝 Wichtige Hinweise

### Secrets Management
- ❌ **NIEMALS** API-Keys in den Code committen
- ✅ Immer Streamlit Secrets verwenden
- ✅ `.gitignore` ist bereits konfiguriert

### Excel-Dateien
- ✅ **Excel Manager** erstellt automatisch alle benötigten Dateien
- ✅ Keine manuellen Uploads erforderlich
- ✅ Robuste Fehlerbehandlung

### API-Keys benötigt
- **OpenAI API**: https://platform.openai.com/api-keys
- **Gmail App Password**: https://support.google.com/accounts/answer/185833

## 🐛 Troubleshooting

### Git Fehler
```bash
# Falls Git nicht konfiguriert:
git config --global user.name "Ihr Name"
git config --global user.email "ihre@email.com"
```

### Python Abhängigkeiten
```bash
# Falls Packages fehlen:
pip install requests streamlit
```

### Repository existiert bereits
- Das Script überschreibt bestehende Repositories nicht
- Löschen Sie das Repository auf GitHub oder wählen Sie einen anderen Namen

## 📞 Support

Bei Problemen:
1. Prüfen Sie alle Voraussetzungen
2. Stellen Sie sicher, dass der GitHub Token korrekt ist
3. Überprüfen Sie die Internetverbindung
4. Kontaktieren Sie den Support

---

**Viel Erfolg beim Deployment! 🚀**