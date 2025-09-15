# 🔬 Paper Claude - Streamlit Research Application

Eine umfassende Streamlit-Anwendung für wissenschaftliche Forschung und Papieranalyse mit Claude AI Integration.

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/Erwin2307-py/Paper_claude)

## 🚀 Features

- 🔍 **Paper Search**: PubMed und wissenschaftliche Datenbank-Suche
- 📊 **Data Analysis**: Gene und SNP Datenanalyse
- 📧 **Email Integration**: Automatische Email-Berichte
- 🤖 **AI Analysis**: OpenAI-basierte Papieranalyse
- 📁 **Excel Management**: Persistente Datenbankverwaltung
- 🎯 **Claude AI**: Intelligente Forschungsunterstützung

## 🌐 Live Demo

**Streamlit Cloud**: [Hier klicken für Live-Demo](https://share.streamlit.io/Erwin2307-py/Paper_claude)

## 🛠️ Installation

### 1. Repository klonen
```bash
git clone https://github.com/Erwin2307-py/Paper_claude.git
cd Paper_claude
```

### 2. Dependencies installieren
```bash
pip install -r requirements.txt
```

### 3. Umgebungsvariablen setzen
Erstellen Sie eine `.env` Datei:
```
OPENAI_API_KEY=sk-your_openai_key
```

### 4. Anwendung starten
```bash
streamlit run streamlit_app.py
```

## ☁️ Streamlit Cloud Deployment

### Required Secrets
Fügen Sie in Streamlit Cloud folgende Secrets hinzu:

```toml
[login]
username = "ihr_username"
password = "ihr_passwort"

[openai]
api_key = "sk-ihr_openai_key"

[email]
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "ihre@email.com"
sender_password = "ihr_app_passwort"
```

## 📋 Module

### 🏠 **Home Dashboard**
- Übersicht über alle Funktionen
- Systemstatus und Statistiken

### 🔍 **Paper Search**
- PubMed API Integration
- Erweiterte Suchfilter
- Export zu Excel

### 📊 **Data Analysis**
- Gene Expression Analyse
- SNP Datenverarbeitung
- Statistische Auswertungen

### 🤖 **AI Analysis**
- PDF Paper Analyse
- Zusammenfassungen generieren
- Key Findings extraktion

### 📁 **Excel Manager**
- Persistente Dateiverwaltung
- Automatische Template-Erstellung
- Datenbank-Backup

### 📧 **Email Integration**
- Automatische Berichte
- Multiple Empfänger
- Excel-Anhänge

## 🔧 Technologie Stack

- **Frontend**: Streamlit
- **Backend**: Python 3.8+
- **AI**: OpenAI GPT Models
- **Data Processing**: Pandas, NumPy
- **File Handling**: openpyxl, PyPDF2
- **Web APIs**: PubMed, Scholarly
- **Email**: SMTP Integration

## 📖 Dokumentation

- **CLAUDE.md**: Development Guidelines
- **DEPLOYMENT.md**: Deployment Instructions
- **requirements.txt**: Python Dependencies

## 🤝 Contributing

1. Fork das Repository
2. Erstelle einen Feature Branch (`git checkout -b feature/amazing-feature`)
3. Commit deine Änderungen (`git commit -m 'Add amazing feature'`)
4. Push zum Branch (`git push origin feature/amazing-feature`)
5. Öffne einen Pull Request

## 📄 License

Dieses Projekt steht unter der MIT License - siehe [LICENSE](LICENSE) für Details.

## 👨‍💻 Author

**Erwin Schimak** - [@Erwin2307-py](https://github.com/Erwin2307-py)

---

🔬 **Paper Claude** - Revolutionäre Forschungstools powered by AI
