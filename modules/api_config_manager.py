# modules/api_config_manager.py - API Configuration Manager
import streamlit as st
import requests
import time
import pandas as pd
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
        self.initialize_session_state()

    def initialize_session_state(self):
        """Initialize API configuration in session state"""
        if "api_config" not in st.session_state:
            st.session_state["api_config"] = {
                "configured": False,
                "last_check": None,
                "available_apis": [],
                "failed_apis": [],
                "config_version": 1.0
            }

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

        st.session_state["api_config"].update({
            "last_check": time.strftime("%Y-%m-%d %H:%M:%S"),
            "available_apis": available_apis,
            "failed_apis": failed_apis,
            "configured": len(available_apis) > 0
        })

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

def show_api_configuration_interface():
    """Main API Configuration Interface"""
    st.title("🔧 API-Konfiguration & Verbindungstests")
    st.write("**Pflichtschritt**: APIs müssen getestet werden, bevor die Suche gestartet werden kann.")

    manager = APIConfigurationManager()

    # Current status display
    if manager.is_configured():
        col1, col2, col3 = st.columns(3)

        with col1:
            available_count = len(manager.get_available_apis())
            st.metric("✅ Verfügbare APIs", available_count)

        with col2:
            failed_count = len(manager.get_failed_apis())
            st.metric("❌ Nicht verfügbare APIs", failed_count)

        with col3:
            last_check = st.session_state["api_config"].get("last_check", "Nie")
            st.metric("🕐 Letzter Test", last_check.split(" ")[1] if " " in last_check else last_check)

        st.success("✅ **API-Konfiguration abgeschlossen** - Sie können jetzt die Unified Search verwenden!")

        # Show available APIs
        if manager.get_available_apis():
            st.subheader("🟢 Verfügbare APIs")
            for api in manager.get_available_apis():
                st.write(f"✅ {api.replace('_', ' ').title()}")

        # Show failed APIs
        if manager.get_failed_apis():
            st.subheader("🔴 Nicht verfügbare APIs")
            for api in manager.get_failed_apis():
                st.write(f"❌ {api.replace('_', ' ').title()}")
    else:
        st.warning("⚠️ **API-Konfiguration erforderlich** - Bitte testen Sie die Verbindungen!")

    # Action buttons
    col1, col2 = st.columns(2)

    with col1:
        if st.button("🔍 **APIs testen**", type="primary"):
            with st.spinner("🔍 Teste API-Verbindungen..."):
                results = manager.check_all_apis()

                # Show detailed results
                show_api_test_results(results)

    with col2:
        if st.button("🔄 Konfiguration zurücksetzen"):
            manager.force_reconfiguration()
            st.info("🔄 Konfiguration zurückgesetzt - führen Sie einen neuen Test durch")
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