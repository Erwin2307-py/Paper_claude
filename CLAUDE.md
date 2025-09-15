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

### Development Testing
- No specific test runner configured - test modules individually
- Check API connectivity through the Online-API Filter module
- Use the Excel Manager to verify persistent data handling

## Architecture Overview

This is a Streamlit-based research application focused on scientific paper analysis and research tasks. The application is structured as a multi-module system with the following key components:

### Main Application (`streamlit_app.py`)
- Main Streamlit application with login functionality using Streamlit secrets
- Multi-page interface with sidebar navigation between research modules
- Uses OpenAI API (v0.28) for AI-powered analysis
- Supports both German and English interfaces with translation capabilities via `google_trans_new`
- Integrated chatbot functionality in right sidebar
- Robust module import system with fallbacks for missing modules

### Core Modules (`modules/`)

#### Research & Search Modules
- **`unified_paper_search.py`**: Central paper search system with multi-API support (PubMed, Europe PMC, Semantic Scholar)
- **`api_config_manager.py`**: Manages API configurations and connectivity testing for all external APIs
- **`codewords_pubmed.py`**: PubMed search and analysis functionality
- **`online_api_filter.py`**: API filtering and data processing with connection testing

#### Data Management
- **`excel_manager.py`**: Persistent Excel file management for Streamlit Cloud deployment with auto-creation of missing files
- Gene and SNP data management with predefined templates

#### Analysis & AI
- **`analyze_paper.py`**: PDF paper analysis using OpenAI GPT models with text extraction capabilities
- **`module_haystack_qa.py`**: Question-answering system using Haystack framework
- **`chonkie_scientific_analysis.py`**: Scientific text analysis using Chonkie framework
- **`labelstudio_scientific_images.py`**: Image annotation and analysis with Label Studio integration

#### Communication
- **`email_module.py`**: Email functionality with SMTP integration and Streamlit secrets configuration

### Key Dependencies
- **Streamlit**: Web interface with session state management
- **OpenAI API (v0.28)**: Text analysis and chatbot functionality
- **PDF Processing**: Multiple libraries (PyPDF2, pdfplumber, PyMuPDF) with fallback support
- **Scientific APIs**: scholarly, selenium for web scraping
- **ML/AI**: transformers, langchain, faiss-cpu, chromadb, haystack
- **Data Processing**: pandas, openpyxl for Excel handling
- **OCR**: pytesseract, easyocr, pdf2image (requires system packages)
- **Translation**: google_trans_new for German-English translation

### Configuration Architecture

#### Secrets Management
- **Primary**: Streamlit secrets (`st.secrets`) for cloud deployment
- **Fallback**: Environment variables for local development
- **Helper function**: `get_secret(category, key, fallback_env_var)` with cascading retrieval

#### Required Secrets Structure
```toml
[login]
username = "your_username"
password = "your_password"

[openai]
api_key = "sk-your_openai_key"

[email]
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "your_email@gmail.com"
sender_password = "your_app_password"
```

### Data Files Structure
- **Templates**: `vorlage_gene.xlsx`, `vorlage_paperqa2.xlsx` for research data
- **Gene Data**: `modules/genes.xlsx`, `modules/snp.xlsx` with predefined datasets
- **Fonts**: `modules/DejaVuSansCondensed.ttf` for PDF generation
- **Backup**: Full `Backup/` directory with application copies

## Development Architecture Patterns

### Module Import Strategy
- **Safe imports**: `safe_import_module()` function with error handling
- **Fallback functionality**: Integrated alternatives when external modules fail
- **Module existence checking**: `check_module_exists()` before import attempts

### Session State Management
- Extensive use of `st.session_state` for cross-page data persistence
- Centralized configuration storage in session state
- User authentication state management

### API Integration Pattern
- **Multi-source search**: Unified Paper Search coordinates multiple APIs
- **Rate limiting**: Built-in delays and retry mechanisms for Semantic Scholar
- **Connectivity testing**: Systematic API health checks via API Configuration Manager
- **Graceful degradation**: Continue with available APIs when others fail

### Error Handling Strategy
- **Robust fallbacks**: Integrated functionality when external modules unavailable
- **User-friendly messages**: Clear error states with actionable guidance
- **Progressive functionality**: Core features work even with missing configurations

## Streamlit Cloud Deployment

### System Dependencies (`packages.txt`)
```
tesseract-ocr
tesseract-ocr-deu
poppler-utils
libgl1-mesa-glx
libglib2.0-0
```

### Configuration Files
- **`.streamlit/config.toml`**: Streamlit settings with increased upload limits and custom theme
- **`.streamlit/secrets.toml`**: Template for secrets (not committed to git)
- **`requirements.txt`**: Python dependencies with specific versions for stability

### Persistent File Management
- **Excel Manager**: Auto-creates missing Excel files on startup
- **Template system**: Predefined gene lists and research templates
- **Cloud-optimized**: Files created in memory and persisted appropriately

### Deployment Workflow
1. **GitHub Integration**: Automated deployment scripts (`deploy_to_github.py`, `deploy_to_github.bat`)
2. **Secret Configuration**: Set up Streamlit Cloud secrets matching template
3. **File Initialization**: Excel Manager ensures all required files exist
4. **API Testing**: Verify external API connectivity via Online-API Filter

## Module Interaction Patterns

### Unified Paper Search System
- **Central coordinator**: `UnifiedPaperSearcher` class manages all search operations
- **Multi-API support**: PubMed, Europe PMC, Semantic Scholar with configurable sources
- **Enhanced data model**: `Paper` dataclass with comprehensive metadata including ChatGPT ratings
- **Export integration**: Direct Excel export functionality with progress tracking

### API Configuration Flow
1. User starts at Home page - sees API configuration status
2. If not configured, directed to Online-API Filter for testing
3. API Manager checks connectivity and stores results in session state
4. Unified Search becomes available only after successful API configuration

### Email Integration
- **Configuration via secrets**: SMTP settings from Streamlit secrets
- **Integrated fallback**: Built-in email functionality when external module fails
- **Search result notifications**: Automated email reports for paper search results

### Paper Excel Filler System
- **`paper_excel_filler.py`**: Core Excel automation system with Claude AI integration
- **`page_excel_filler.py`**: Standalone UI interface with comprehensive workflow
- **Claude API Integration**: Intelligent paper analysis and gene extraction
- **Multi-template support**: Works with vorlage_paperqa2.xlsx and vorlage_gene.xlsx
- **Batch processing**: Automated Excel generation from selected papers

---

## Recent Development Updates (2025-09-15)

### Paper Excel Filler Module Implementation

#### What was implemented:
1. **Core Excel Filler System** (`modules/paper_excel_filler.py`):
   - `PaperExcelFiller` class with Claude AI integration
   - Automatic Excel template copying and intelligent data filling
   - Gene extraction from paper titles and abstracts
   - Comprehensive error handling and fallback mechanisms

2. **User Interface** (`modules/page_excel_filler.py`):
   - 4-tab interface: Paper loading, Selection, Excel creation, Statistics
   - Integration with Unified Search results
   - Sample paper data for testing purposes
   - Real-time status dashboard and metrics

3. **Main App Integration**:
   - Added imports and navigation entries in `streamlit_app.py`
   - Home page promotion section with feature highlights
   - Sidebar navigation integration

#### Issues Resolved:

**Problem 1: Module Not Visible in UI**
- **Issue**: Paper Excel Filler was technically implemented but not visible in UI
- **Root Cause**: Duplicate `page_home()` functions in streamlit_app.py (lines 559 and 910)
- **Solution**: Removed duplicate function, kept version with Paper Excel Filler promotion
- **Result**: Module now properly displayed with button on home screen

**Problem 2: API Configuration Not Persistent**
- **Issue**: After configuring Online API Filter, Paper Search was not accessible
- **Root Cause**: API configuration stored only in Session State (temporary)
- **Solution**: Implemented persistent storage in `api_config.json`
- **Implementation Details**:
  - Added `_load_config_from_file()` for startup configuration loading
  - Added `_save_config_to_file()` for automatic persistence after API tests
  - Modified `force_reconfiguration()` to clear both session and file storage
- **Result**: API configuration persists across app restarts and page switches

#### Current Functionality:

**Paper Excel Filler Features:**
- ✅ **Home Dashboard Integration**: Prominent button and feature promotion
- ✅ **Sidebar Navigation**: Full module access via "📊 Paper Excel Filler"
- ✅ **Claude AI Analysis**: Intelligent paper data extraction and analysis
- ✅ **Automatic Gene Detection**: Extracts genes from paper content
- ✅ **Excel Template Support**: Works with existing vorlage_paperqa2.xlsx and vorlage_gene.xlsx
- ✅ **Batch Processing**: Multiple papers to Excel files simultaneously
- ✅ **Integration with Unified Search**: Uses papers from search results
- ✅ **Sample Data**: Built-in test papers for demonstration

**API Configuration System:**
- ✅ **Persistent Storage**: Configuration saved in `api_config.json`
- ✅ **Automatic Loading**: Restores configuration on app startup
- ✅ **Real-time Testing**: Live API connectivity verification
- ✅ **Multi-API Support**: PubMed, Europe PMC, Semantic Scholar, OpenAlex
- ✅ **Graceful Fallbacks**: Continues with available APIs when others fail

#### Files Modified/Created:

**New Files:**
- `modules/paper_excel_filler.py` - Core Excel automation system
- `modules/page_excel_filler.py` - User interface module
- `api_config.json` - Persistent API configuration storage (auto-created)

**Modified Files:**
- `streamlit_app.py` - Integration, navigation, duplicate function removal
- `modules/api_config_manager.py` - Persistent configuration system
- `CLAUDE.md` - This documentation update

#### Git Commits:
1. **Initial Implementation**: Paper Excel Filler module creation and integration
2. **Integration Fix**: Remove duplicate page_home() function - fixed UI visibility
3. **Persistent API Config**: Fixed Paper Search accessibility after API configuration

#### Usage Workflow:
1. **API Setup**: 📊 Online-API Filter → Test APIs (stored persistently)
2. **Paper Search**: 🔍 Paper Search → Find and analyze papers
3. **Excel Generation**: 📊 Paper Excel Filler → Select papers → Generate Excel files
4. **Results**: Download completed Excel files with intelligent data filling

#### Next Development Priorities:
- Manual paper input functionality in Excel Filler
- Enhanced Claude AI prompts for better analysis accuracy
- Additional Excel template support
- Advanced filtering and selection options
- Export statistics and reporting features

---