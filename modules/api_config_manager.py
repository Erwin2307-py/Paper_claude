# modules/api_config_manager.py - API Configuration Manager
import streamlit as st
import requests
import time
import pandas as pd
import json
import os
from typing import Dict, List, Tuple
from dataclasses import dataclass

@dataclass
class APIStatus:
    """API Status Information"""
    name: str
    url: str
    status: bool
    response_time: float
    error_message: str = ""
    last_checked: str = ""

class APIConfigurationManager:
    """Manages API configurations and connectivity checks"""

    def __init__(self):
        self.config_file = "api_config.json"
        self.initialize_session_state()

    def initialize_session_state(self):
        """Initialize API configuration in session state"""
        if "api_config" not in st.session_state:
            # Try to load from persistent file first
            config = self._load_config_from_file()
            st.session_state["api_config"] = config

    def _load_config_from_file(self) -> Dict:
        """Load API configuration from persistent file"""
        default_config = {
            "configured": False,
            "last_check": None,
            "available_apis": [],
            "failed_apis": [],
            "config_version": 1.0
        }

        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    saved_config = json.load(f)
                    # Merge with default config to ensure all keys exist
                    default_config.update(saved_config)
                    return default_config
        except Exception as e:
            # If file is corrupted or unreadable, use default
            pass

        return default_config

    def _save_config_to_file(self, config: Dict):
        """Save API configuration to persistent file"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            # Silently fail if cannot write file
            pass

    def check_all_apis(self) -> Dict[str, APIStatus]:
        """Test all available APIs and return status"""
        apis_to_test = [
            ("PubMed", self._check_pubmed),
            ("Europe PMC", self._check_europe_pmc),
            ("Semantic Scholar", self._check_semantic_scholar),
            ("OpenAlex", self._check_openalex),
        ]

        results = {}

        # Progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, (api_name, test_func) in enumerate(apis_to_test):
            status_text.text(f"🔍 Teste {api_name}...")
            progress = (i + 0.5) / len(apis_to_test)
            progress_bar.progress(progress)

            start_time = time.time()
            try:
                status = test_func()
                response_time = time.time() - start_time

                results[api_name.lower().replace(' ', '_')] = APIStatus(
                    name=api_name,
                    url=self._get_api_url(api_name),
                    status=status,
                    response_time=response_time,
                    last_checked=time.strftime("%H:%M:%S")
                )
            except Exception as e:
                response_time = time.time() - start_time
                results[api_name.lower().replace(' ', '_')] = APIStatus(
                    name=api_name,
                    url=self._get_api_url(api_name),
                    status=False,
                    response_time=response_time,
                    error_message=str(e),
                    last_checked=time.strftime("%H:%M:%S")
                )

            progress = (i + 1) / len(apis_to_test)
            progress_bar.progress(progress)

        status_text.text("✅ API-Tests abgeschlossen!")

        # Update session state
        available_apis = [api for api, status in results.items() if status.status]
        failed_apis = [api for api, status in results.items() if not status.status]

        # Update session state
        updated_config = {
            "last_check": time.strftime("%Y-%m-%d %H:%M:%S"),
            "available_apis": available_apis,
            "failed_apis": failed_apis,
            "configured": len(available_apis) > 0,
            "config_version": 1.0
        }
        st.session_state["api_config"].update(updated_config)

        # Persist configuration to file
        self._save_config_to_file(st.session_state["api_config"])

        return results

    def _check_pubmed(self) -> bool:
        """Test PubMed API"""
        url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
        params = {"db": "pubmed", "term": "test", "retmode": "json", "retmax": 1}
        try:
            response = requests.get(url, params=params, timeout=5)
            response.raise_for_status()
            data = response.json()
            return "esearchresult" in data
        except Exception:
            return False

    def _check_europe_pmc(self) -> bool:
        """Test Europe PMC API"""
        url = "https://www.ebi.ac.uk/europepmc/webservices/rest/search"
        params = {"query": "test", "format": "json", "pageSize": 1}
        try:
            response = requests.get(url, params=params, timeout=5)
            response.raise_for_status()
            data = response.json()
            return "resultList" in data and "result" in data["resultList"]
        except Exception:
            return False

    def _check_semantic_scholar(self) -> bool:
        """Test Semantic Scholar API"""
        url = "https://api.semanticscholar.org/graph/v1/paper/search"
        params = {"query": "test", "limit": 1, "fields": "title"}
        headers = {"User-Agent": "Paper-Claude-Research-Tool/1.0"}
        try:
            response = requests.get(url, params=params, headers=headers, timeout=10)
            if response.status_code == 429:
                return False  # Rate limited
            response.raise_for_status()
            data = response.json()
            return "data" in data
        except Exception:
            return False

    def _check_openalex(self) -> bool:
        """Test OpenAlex API"""
        url = "https://api.openalex.org/works"
        params = {"search": "test", "per_page": 1}
        try:
            response = requests.get(url, params=params, timeout=5)
            response.raise_for_status()
            data = response.json()
            return "results" in data
        except Exception:
            return False

    def _get_api_url(self, api_name: str) -> str:
        """Get base URL for API"""
        urls = {
            "PubMed": "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/",
            "Europe PMC": "https://www.ebi.ac.uk/europepmc/webservices/rest/",
            "Semantic Scholar": "https://api.semanticscholar.org/graph/v1/",
            "OpenAlex": "https://api.openalex.org/"
        }
        return urls.get(api_name, "")

    def is_configured(self) -> bool:
        """Check if API configuration is complete"""
        config = st.session_state.get("api_config", {})
        return config.get("configured", False) and len(config.get("available_apis", [])) > 0

    def get_available_apis(self) -> List[str]:
        """Get list of available APIs"""
        return st.session_state.get("api_config", {}).get("available_apis", [])

    def get_failed_apis(self) -> List[str]:
        """Get list of failed APIs"""
        return st.session_state.get("api_config", {}).get("failed_apis", [])

    def force_reconfiguration(self):
        """Force reconfiguration by clearing status"""
        st.session_state["api_config"]["configured"] = False
        st.session_state["api_config"]["available_apis"] = []
        st.session_state["api_config"]["failed_apis"] = []
        # Also clear from persistent file
        self._save_config_to_file(st.session_state["api_config"])

def create_default_settings_file():
    """Create default user settings Excel file if it doesn't exist"""
    try:
        import pandas as pd

        user_settings_data = {
            'User_Name': [
                'Standard',
                'Erwin_Genetics',
                'Erwin_Cancer',
                'Erwin_Comprehensive',
                'Demo_Basic',
                'Demo_Advanced'
            ],
            'Max_Results_Per_API': [50, 100, 75, 150, 25, 200],
            'Enable_PubMed': [True, True, True, True, True, True],
            'Enable_Europe_PMC': [True, True, True, True, False, True],
            'Enable_Semantic_Scholar': [True, True, False, True, False, True],
            'Enable_OpenAlex': [False, True, False, True, False, True],
            'ChatGPT_Analysis': [True, True, True, True, False, True],
            'Min_Citation_Count': [0, 5, 10, 0, 0, 20],
            'Max_Publication_Age_Years': [10, 5, 3, 15, 20, 2],
            'Include_Review_Papers': [True, True, False, True, True, False],
            'Include_Clinical_Trials': [True, True, True, True, False, True],
            'Language_Filter': ['en', 'en', 'en', 'en,de', 'en', 'en'],
            'Email_Notifications': [False, True, True, True, False, True],
            'Auto_Excel_Export': [False, True, False, True, False, True],
            'Search_Description': [
                'Standard settings for general searches',
                'Optimized for genetic research with high citation requirements',
                'Cancer research focused with recent papers only',
                'Comprehensive search across all databases',
                'Basic settings for demonstration purposes',
                'Advanced settings for power users'
            ]
        }

        df = pd.DataFrame(user_settings_data)
        df.to_excel('user_search_settings.xlsx', index=False, engine='openpyxl')
        return True
    except Exception as e:
        return False

def show_api_configuration_interface():
    """User Search Settings Interface"""
    st.title("⚙️ Search Settings - User Profile Selection")
    st.write("**Wählen Sie Ihre bevorzugten Sucheinstellungen** aus vorkonfigurierten Profilen oder passen Sie diese an.")

    # Load user settings from Excel
    try:
        import pandas as pd
        settings_df = pd.read_excel("user_search_settings.xlsx")
    except Exception as e:
        st.error(f"❌ **User Settings Excel nicht gefunden**: {e}")
        st.info("💡 Die Datei `user_search_settings.xlsx` wird automatisch erstellt...")

        # Create default settings if file doesn't exist
        create_default_settings_file()
        try:
            settings_df = pd.read_excel("user_search_settings.xlsx")
            st.success("✅ **Standard-Settings erstellt!** Seite wird neu geladen...")
            st.rerun()
        except:
            st.error("❌ Fehler beim Erstellen der Settings-Datei")
            return

    # Current selected settings display
    current_settings = st.session_state.get("selected_search_settings", {})
    if current_settings:
        st.success(f"✅ **Aktuelle Einstellungen**: {current_settings.get('User_Name', 'Unbekannt')}")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Max Results", current_settings.get('Max_Results_Per_API', 0))
        with col2:
            active_apis = sum([
                current_settings.get('Enable_PubMed', False),
                current_settings.get('Enable_Europe_PMC', False),
                current_settings.get('Enable_Semantic_Scholar', False),
                current_settings.get('Enable_OpenAlex', False)
            ])
            st.metric("Active APIs", active_apis)
        with col3:
            st.metric("ChatGPT Analysis", "✅" if current_settings.get('ChatGPT_Analysis', False) else "❌")
        with col4:
            st.metric("Min Citations", current_settings.get('Min_Citation_Count', 0))

    # Settings selection interface
    st.subheader("🎯 Profile Selection")

    # Display available profiles
    for idx, row in settings_df.iterrows():
        with st.expander(f"👤 **{row['User_Name']}** - {row['Search_Description']}", expanded=False):

            col_info1, col_info2 = st.columns(2)

            with col_info1:
                st.write("**📊 Search Parameters:**")
                st.write(f"• Max Results per API: {row['Max_Results_Per_API']}")
                st.write(f"• Min Citation Count: {row['Min_Citation_Count']}")
                st.write(f"• Max Age (Years): {row['Max_Publication_Age_Years']}")
                st.write(f"• Language: {row['Language_Filter']}")

            with col_info2:
                st.write("**🔗 Active APIs:**")
                apis = []
                if row['Enable_PubMed']: apis.append("PubMed")
                if row['Enable_Europe_PMC']: apis.append("Europe PMC")
                if row['Enable_Semantic_Scholar']: apis.append("Semantic Scholar")
                if row['Enable_OpenAlex']: apis.append("OpenAlex")
                for api in apis:
                    st.write(f"• ✅ {api}")

                st.write("**🤖 Features:**")
                st.write(f"• ChatGPT Analysis: {'✅' if row['ChatGPT_Analysis'] else '❌'}")
                st.write(f"• Email Notifications: {'✅' if row['Email_Notifications'] else '❌'}")
                st.write(f"• Auto Excel Export: {'✅' if row['Auto_Excel_Export'] else '❌'}")

            if st.button(f"🚀 **Use {row['User_Name']} Settings**", key=f"select_{idx}"):
                # Save selected settings to session state
                selected_settings = row.to_dict()
                st.session_state["selected_search_settings"] = selected_settings
                st.success(f"✅ **{row['User_Name']} Settings aktiviert!** Paper Search ist bereit.")
                st.rerun()

    # Information section
    with st.expander("ℹ️ Über die APIs"):
        st.markdown("""
        **🏥 PubMed**: NCBI's biomedizinische Datenbank - kostenlos, keine API-Limits

        **🌍 Europe PMC**: Europäische biomedizinische Datenbank - kostenlos, Volltext verfügbar

        **🔬 Semantic Scholar**: Interdisziplinäre Forschungsdatenbank - kostenlos, aber mit Rate Limits

        **🔗 OpenAlex**: Open-Access wissenschaftliche Datenbank - kostenlos, keine Authentifizierung erforderlich

        **Hinweis**: Mindestens eine API muss funktionieren, um die Unified Search zu verwenden.
        """)

    # Troubleshooting
    if not manager.is_configured():
        with st.expander("🛠️ Fehlerbehebung"):
            st.markdown("""
            **Häufige Probleme:**

            - **Internetverbindung**: Prüfen Sie Ihre Netzwerkverbindung
            - **Firewall**: Möglicherweise blockiert Ihre Firewall API-Aufrufe
            - **Rate Limits**: Einige APIs haben Nutzungsbeschränkungen
            - **Temporäre Ausfälle**: APIs können vorübergehend nicht verfügbar sein

            **Lösungen:**
            - Warten Sie einige Minuten und testen Sie erneut
            - Verwenden Sie verschiedene Netzwerkverbindungen
            - Kontaktieren Sie Ihren IT-Administrator bei persistenten Problemen
            """)

def show_api_test_results(results: Dict[str, APIStatus]):
    """Display detailed API test results"""
    st.subheader("📊 Testergebnisse")

    success_count = sum(1 for status in results.values() if status.status)
    total_count = len(results)

    # Summary metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("✅ Erfolgreich", success_count)
    with col2:
        st.metric("❌ Fehlgeschlagen", total_count - success_count)
    with col3:
        st.metric("📊 Erfolgsrate", f"{(success_count/total_count)*100:.0f}%")

    # Detailed results table
    result_data = []
    for api_key, status in results.items():
        result_data.append({
            "API": status.name,
            "Status": "✅ Online" if status.status else "❌ Offline",
            "Antwortzeit": f"{status.response_time:.2f}s",
            "Letzter Test": status.last_checked,
            "Fehler": status.error_message[:50] + "..." if len(status.error_message) > 50 else status.error_message
        })

    df = pd.DataFrame(result_data)
    st.dataframe(df, width=1000)

    if success_count > 0:
        st.success(f"🎉 **Konfiguration erfolgreich!** {success_count} von {total_count} APIs sind verfügbar.")
    else:
        st.error("❌ **Keine APIs verfügbar!** Prüfen Sie Ihre Internetverbindung und versuchen Sie es erneut.")

def require_api_configuration(func):
    """Decorator to require API configuration before function execution"""
    def wrapper(*args, **kwargs):
        manager = APIConfigurationManager()

        if not manager.is_configured():
            st.error("🚫 **API-Konfiguration erforderlich!**")
            st.info("👉 Gehen Sie zum **'📊 Online-API Filter'** und testen Sie die APIs, bevor Sie die Suche starten.")

            if st.button("🔧 Zur API-Konfiguration"):
                st.session_state["current_page"] = "📊 Online-API Filter"
                st.rerun()

            return None

        return func(*args, **kwargs)

    return wrapper

# Main module function
def module_online_api_filter():
    """Main API configuration module"""
    show_api_configuration_interface()