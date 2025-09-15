# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Common Commands

### Running the Application
```bash
streamlit run streamlit_app.py
```

### Installing Dependencies
```bash
pip install -r requirements.txt
```

## Architecture Overview

This is a Streamlit-based research application focused on scientific paper analysis and research tasks. The application is structured as a multi-module system with the following key components:

### Main Application (`streamlit_app.py`)
- Main Streamlit application with login functionality using Streamlit secrets
- Multi-page interface with various research modules
- Uses OpenAI API for AI-powered analysis (requires `OPENAI_API_KEY` environment variable)
- Supports both German and English interfaces with translation capabilities

### Core Modules (`modules/`)
- **`analyze_paper.py`**: PDF paper analysis using OpenAI GPT models with text extraction capabilities
- **`email_module.py`**: Email functionality with SMTP integration and Streamlit secrets configuration
- **`codewords_pubmed.py`**: PubMed search and analysis functionality
- **`online_api_filter.py`**: Online API filtering and data processing
- **`module_haystack_qa.py`**: Question-answering system using Haystack framework
- **`chonkie_scientific_analysis.py`**: Scientific text analysis using Chonkie framework
- **`labelstudio_scientific_images.py`**: Image annotation and analysis with Label Studio integration

### Key Dependencies
- Streamlit for web interface
- OpenAI API (v0.28) for text analysis
- Multiple PDF processing libraries (PyPDF2, pdfplumber, PyMuPDF)
- Scientific research tools (scholarly, selenium for web scraping)
- ML/AI frameworks (transformers, langchain, faiss-cpu, chromadb)
- Email functionality (built-in smtplib)
- Excel processing (openpyxl, pandas)

### Configuration
- Uses `.env` file for environment variables (OpenAI API key)
- Streamlit secrets for login credentials and email configuration
- Login credentials accessed via `st.secrets["login"]["username"]` and `st.secrets["login"]["password"]`

### Data Files
- `vorlage_gene.xlsx` and `vorlage_paperqa2.xlsx`: Excel templates for research data
- `genes.xlsx` and `snp.xlsx`: Gene and SNP data files
- Various font files for PDF generation

## Development Notes

- The application uses session state extensively for maintaining user state across pages
- Most modules are designed to work both standalone and as part of the main Streamlit application
- PDF processing supports multiple libraries as fallbacks for robust text extraction
- Translation support using `google_trans_new` for German-English translation
- The codebase includes backup functionality in the `Backup/` directory

## Streamlit Cloud Deployment

### Persistent File Management
- **Excel Manager** (`modules/excel_manager.py`): Ensures all required Excel files exist and are properly initialized
- **Auto-creation**: Missing Excel files (genes.xlsx, SNP data, paper templates) are automatically created
- **Robust file handling**: Files are created with sample data if missing during deployment

### Required Configuration Files
- `.streamlit/config.toml`: Streamlit configuration optimized for cloud deployment
- `.streamlit/secrets.toml`: Template for secrets (API keys, login credentials) - not committed to git
- `packages.txt`: System packages required for OCR and PDF processing
- `.gitignore`: Properly excludes sensitive files and temporary data

### Environment Variables and Secrets
- API keys can be set via environment variables or Streamlit secrets
- Login credentials managed through Streamlit secrets
- Fallback mechanisms ensure the app works even with missing configuration

### Deployment Checklist
1. Set up Streamlit secrets with API keys and login credentials
2. Ensure all Excel files are initialized via the Excel Manager
3. Test all modules work with the persistent file system
4. Verify OCR and PDF processing work with system packages