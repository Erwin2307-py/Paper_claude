# modules/unified_paper_search.py - Zentrales Paper-Search-System
import streamlit as st
import requests
import xml.etree.ElementTree as ET
import pandas as pd
import datetime
import time
import re
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass

# Lokale Imports
from modules.excel_manager import initialize_excel_manager
from modules.email_module import load_email_config_from_secrets, send_paper_results_email
from modules.api_config_manager import APIConfigurationManager, require_api_configuration

# Try to import the Excel Filler - with fallback if not available
try:
    from modules.paper_excel_filler import show_paper_excel_interface, PaperExcelFiller
    EXCEL_FILLER_AVAILABLE = True
except ImportError:
    EXCEL_FILLER_AVAILABLE = False

@dataclass
class Paper:
    """Einheitliche Paper-Datenstruktur mit erweiterten Analyse-Funktionen"""
    title: str
    authors: str
    journal: str
    year: str
    abstract: str
    doi: str = ""
    pubmed_id: str = ""
    url: str = ""
    keywords: str = ""
    citations: int = 0
    relevance_score: float = 0.0
    source: str = "unknown"  # pubmed, scholar, semantic_scholar, etc.
    chatgpt_rating: float = 0.0  # ChatGPT Rating 0-10
    chatgpt_summary: str = ""  # ChatGPT Zusammenfassung
    is_downloadable: bool = False  # Kann heruntergeladen werden
    pdf_url: str = ""  # URL für PDF Download
    analyzed: bool = False  # Wurde analysiert

    def to_dict(self) -> Dict:
        """Konvertiert zu Dictionary für Excel/CSV Export"""
        return {
            'title': self.title,
            'authors': self.authors,
            'journal': self.journal,
            'year': self.year,
            'abstract': self.abstract,
            'doi': self.doi,
            'pubmed_id': self.pubmed_id,
            'url': self.url,
            'keywords': self.keywords,
            'citations': self.citations,
            'relevance_score': self.relevance_score,
            'source': self.source,
            'chatgpt_rating': self.chatgpt_rating,
            'chatgpt_summary': self.chatgpt_summary,
            'is_downloadable': self.is_downloadable,
            'pdf_url': self.pdf_url,
            'analyzed': self.analyzed,
            'date_added': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

class UnifiedPaperSearcher:
    """Zentrales Paper-Search-System mit erweiterten Analyse-Funktionen"""

    def __init__(self):
        self.excel_manager = initialize_excel_manager()
        self.email_config = load_email_config_from_secrets()
        self.api_manager = APIConfigurationManager()
        # OpenAI API Key aus Streamlit Secrets
        self.openai_api_key = self._get_openai_key()

    def _get_openai_key(self) -> Optional[str]:
        """Holt OpenAI API Key aus Streamlit Secrets"""
        try:
            if hasattr(st, 'secrets'):
                return st.secrets.get("openai", {}).get("api_key")
        except Exception:
            pass
        return None

    def search_all_sources(self, query: str, max_results: int = 50,
                          sources: List[str] = None) -> List[Paper]:
        """Sucht in allen verfügbaren Quellen mit robuster Fehlerbehandlung"""
        if sources is None:
            sources = ["pubmed", "europe_pmc"]  # Semantic Scholar als optional

        all_papers = []
        successful_sources = []
        failed_sources = []

        # Progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()

        total_sources = len(sources)
        per_source_limit = max(10, max_results // len(sources))

        for i, source in enumerate(sources):
            source_display = source.replace('_', ' ').title()
            status_text.text(f"🔍 Suche in {source_display}...")
            progress = (i + 0.5) / total_sources
            progress_bar.progress(progress)

            papers = []
            try:
                if source == "pubmed":
                    papers = self.search_pubmed(query, per_source_limit)
                elif source == "semantic_scholar":
                    papers = self.search_semantic_scholar(query, min(per_source_limit, 20))
                elif source == "europe_pmc":
                    papers = self.search_europe_pmc(query, per_source_limit)
                else:
                    st.warning(f"⚠️ Unbekannte Quelle: {source}")
                    continue

                if papers:
                    all_papers.extend(papers)
                    successful_sources.append(f"{source_display} ({len(papers)} Papers)")
                else:
                    st.info(f"ℹ️ Keine Ergebnisse von {source_display}")

            except Exception as e:
                error_msg = str(e)
                if "429" in error_msg:
                    st.warning(f"⚠️ Rate Limit bei {source_display} - übersprungen")
                elif "timeout" in error_msg.lower():
                    st.warning(f"⚠️ Timeout bei {source_display} - übersprungen")
                else:
                    st.warning(f"⚠️ Fehler bei {source_display}: {error_msg}")

                failed_sources.append(source_display)

            progress = (i + 1) / total_sources
            progress_bar.progress(progress)

        # Ergebnis-Summary
        if successful_sources:
            st.success(f"✅ Erfolgreich: {', '.join(successful_sources)}")

        if failed_sources:
            st.info(f"⚠️ Übersprungen: {', '.join(failed_sources)}")

        # Duplikate entfernen basierend auf Titel
        if all_papers:
            unique_papers = self._remove_duplicates(all_papers)
            status_text.text(f"✅ {len(unique_papers)} einzigartige Papers von {len(successful_sources)} Quellen gefunden!")
        else:
            unique_papers = []
            status_text.text("❌ Keine Papers gefunden - versuchen Sie andere Suchbegriffe")

        return unique_papers

    def search_pubmed(self, query: str, max_results: int = 20) -> List[Paper]:
        """Sucht in PubMed"""
        papers = []

        # 1. E-Search für PMIDs
        search_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
        search_params = {
            "db": "pubmed",
            "term": query,
            "retmode": "json",
            "retmax": max_results
        }

        try:
            response = requests.get(search_url, params=search_params, timeout=10)
            response.raise_for_status()
            search_data = response.json()

            pmids = search_data.get("esearchresult", {}).get("idlist", [])

            if not pmids:
                return papers

            # 2. E-Fetch für Details
            fetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
            fetch_params = {
                "db": "pubmed",
                "id": ",".join(pmids[:max_results]),
                "retmode": "xml"
            }

            response = requests.get(fetch_url, params=fetch_params, timeout=15)
            response.raise_for_status()

            # XML parsen
            root = ET.fromstring(response.text)

            for article in root.findall(".//PubmedArticle"):
                try:
                    paper = self._parse_pubmed_article(article)
                    # Zusätzliche Validierung
                    if paper and hasattr(paper, 'title') and paper.title:
                        papers.append(paper)
                except Exception as e:
                    continue  # Skip fehlerhafter Artikel

        except Exception as e:
            st.error(f"PubMed-Suche fehlgeschlagen: {str(e)}")

        return papers

    def search_semantic_scholar(self, query: str, max_results: int = 20) -> List[Paper]:
        """Sucht in Semantic Scholar mit Rate Limiting und Retry-Logic"""
        papers = []

        url = "https://api.semanticscholar.org/graph/v1/paper/search"
        params = {
            "query": query,
            "limit": min(max_results, 20),  # Limit auf 20 für bessere Rate Limits
            "fields": "title,authors,venue,year,abstract,url,citationCount,externalIds"
        }

        headers = {
            "User-Agent": "Paper-Claude-Research-Tool/1.0",
            "Accept": "application/json"
        }

        # Retry-Logic mit exponential backoff
        max_retries = 3
        base_delay = 2

        for attempt in range(max_retries):
            try:
                response = requests.get(url, params=params, headers=headers, timeout=15)

                if response.status_code == 429:
                    # Rate limit erreicht
                    retry_after = int(response.headers.get("Retry-After", base_delay * (2 ** attempt)))
                    if attempt < max_retries - 1:
                        st.warning(f"⏳ Semantic Scholar Rate Limit - Warte {retry_after}s...")
                        time.sleep(retry_after)
                        continue
                    else:
                        st.warning("⚠️ Semantic Scholar Rate Limit - überspringe diese Quelle")
                        break

                elif response.status_code == 403:
                    st.warning("⚠️ Semantic Scholar API nicht verfügbar - überspringe diese Quelle")
                    break

                response.raise_for_status()
                data = response.json()

                for paper_data in data.get("data", []):
                    try:
                        paper = self._parse_semantic_scholar_paper(paper_data)
                        # Zusätzliche Validierung
                        if paper and hasattr(paper, 'title') and paper.title:
                            papers.append(paper)
                    except Exception:
                        continue

                break  # Erfolgreicher Request

            except requests.exceptions.Timeout:
                if attempt < max_retries - 1:
                    st.warning(f"⏳ Semantic Scholar Timeout - Retry {attempt + 1}/{max_retries}")
                    time.sleep(base_delay * (2 ** attempt))
                    continue
                else:
                    st.warning("⚠️ Semantic Scholar Timeout - überspringe diese Quelle")

            except Exception as e:
                if "429" in str(e):
                    if attempt < max_retries - 1:
                        delay = base_delay * (2 ** attempt)
                        st.warning(f"⏳ Rate Limit - Warte {delay}s...")
                        time.sleep(delay)
                        continue
                    else:
                        st.warning("⚠️ Semantic Scholar Rate Limit - überspringe diese Quelle")
                        break
                else:
                    st.warning(f"⚠️ Semantic Scholar-Fehler: {str(e)} - überspringe diese Quelle")
                    break

        return papers

    def search_europe_pmc(self, query: str, max_results: int = 20) -> List[Paper]:
        """Sucht in Europe PMC"""
        papers = []

        url = "https://www.ebi.ac.uk/europepmc/webservices/rest/search"
        params = {
            "query": query,
            "format": "json",
            "pageSize": max_results
        }

        try:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            data = response.json()

            for paper_data in data.get("resultList", {}).get("result", []):
                try:
                    paper = self._parse_europe_pmc_paper(paper_data)
                    # Zusätzliche Validierung
                    if paper and hasattr(paper, 'title') and paper.title:
                        papers.append(paper)
                except Exception:
                    continue

        except Exception as e:
            st.error(f"Europe PMC-Suche fehlgeschlagen: {str(e)}")

        return papers

    def save_to_excel(self, papers: List[Paper], search_term: str) -> bool:
        """Speichert Papers in Excel-Datenbank"""
        try:
            for paper in papers:
                paper_data = paper.to_dict()
                success = self.excel_manager.add_paper_to_database(paper_data, search_term)
                if not success:
                    st.warning(f"⚠️ Fehler beim Speichern von: {paper.title[:50]}...")

            st.success(f"✅ {len(papers)} Papers in Excel-Datenbank gespeichert!")
            return True

        except Exception as e:
            st.error(f"❌ Excel-Speicherung fehlgeschlagen: {str(e)}")
            return False

    def send_email_notification(self, papers: List[Paper], search_term: str) -> bool:
        """Sendet Email-Benachrichtigung"""
        if not self.email_config or not papers:
            return False

        try:
            # Konvertiere Papers für Email
            paper_list = [paper.to_dict() for paper in papers]
            send_paper_results_email(self.email_config, search_term, paper_list)
            return True
        except Exception as e:
            st.error(f"❌ Email-Versand fehlgeschlagen: {str(e)}")
            return False

    def analyze_papers_with_chatgpt(self, papers: List[Paper], query: str) -> List[Paper]:
        """Analysiert Papers mit ChatGPT für Rating und Zusammenfassung"""
        if not self.openai_api_key:
            st.warning("⚠️ OpenAI API Key nicht konfiguriert - ChatGPT-Analyse übersprungen")
            return papers

        analyzed_papers = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, paper in enumerate(papers):
            status_text.text(f"🤖 ChatGPT analysiert Paper {i+1}/{len(papers)}: {paper.title[:50]}...")
            progress = i / len(papers)
            progress_bar.progress(progress)

            try:
                rating, summary = self._get_chatgpt_analysis(paper, query)
                paper.chatgpt_rating = rating
                paper.chatgpt_summary = summary
                paper.analyzed = True
            except Exception as e:
                st.warning(f"⚠️ ChatGPT-Analyse fehlgeschlagen für: {paper.title[:30]}... - {str(e)}")
                paper.chatgpt_rating = 0.0
                paper.chatgpt_summary = "Analyse fehlgeschlagen"

            analyzed_papers.append(paper)

        progress_bar.progress(1.0)
        status_text.text("✅ ChatGPT-Analyse abgeschlossen!")

        # Sortiere nach Rating (höchste zuerst)
        analyzed_papers.sort(key=lambda p: p.chatgpt_rating, reverse=True)
        return analyzed_papers

    def _get_chatgpt_analysis(self, paper: Paper, original_query: str) -> Tuple[float, str]:
        """Einzelne ChatGPT-Analyse für ein Paper"""
        prompt = f"""
        Analysiere das folgende wissenschaftliche Paper im Kontext der Suchanfrage: "{original_query}"

        **Paper Details:**
        Titel: {paper.title}
        Autoren: {paper.authors}
        Journal: {paper.journal} ({paper.year})
        Abstract: {paper.abstract[:1000]}...

        **Bitte bewerte das Paper auf einer Skala von 0-10 basierend auf:**
        1. Relevanz zur Suchanfrage
        2. Wissenschaftliche Qualität
        3. Methodische Stärke
        4. Praktische Anwendbarkeit

        **Antwortformat:**
        Rating: [0-10]
        Zusammenfassung: [2-3 Sätze über Haupterkenntnisse und Relevanz]
        """

        try:
            import openai
            openai.api_key = self.openai_api_key

            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Du bist ein Experte für wissenschaftliche Paper-Analyse."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=300
            )

            content = response.choices[0].message.content

            # Parse Rating und Zusammenfassung
            lines = content.split('\n')
            rating = 0.0
            summary = ""

            for line in lines:
                if line.startswith("Rating:"):
                    try:
                        rating_str = line.split("Rating:")[1].strip()
                        rating = float(re.search(r'(\d+\.?\d*)', rating_str).group(1))
                        rating = max(0, min(10, rating))  # Begrenze auf 0-10
                    except:
                        rating = 5.0
                elif line.startswith("Zusammenfassung:"):
                    summary = line.split("Zusammenfassung:")[1].strip()

            return rating, summary or "Keine Zusammenfassung verfügbar"

        except Exception as e:
            raise Exception(f"OpenAI API Error: {str(e)}")

    def check_paper_downloadability(self, papers: List[Paper]) -> List[Paper]:
        """Prüft welche Papers herunterladbar sind"""
        for paper in papers:
            # Europe PMC oft mit Volltext
            if paper.source == "europe_pmc" and paper.url:
                paper.is_downloadable = True
                paper.pdf_url = paper.url
            # PubMed PMC Links
            elif paper.pubmed_id and paper.source == "pubmed":
                pmc_url = f"https://www.ncbi.nlm.nih.gov/pmc/articles/PMC{paper.pubmed_id}/"
                paper.pdf_url = pmc_url
                paper.is_downloadable = True
            # DOI zu PDF versuchen
            elif paper.doi:
                paper.pdf_url = f"https://doi.org/{paper.doi}"
                paper.is_downloadable = True

        return papers

    def _parse_pubmed_article(self, article) -> Paper:
        """Parst PubMed XML Article zu Paper"""
        # PMID
        pmid_elem = article.find(".//PMID")
        pmid = pmid_elem.text if pmid_elem is not None else ""

        # Titel mit vollständiger Null-Prüfung
        title_elem = article.find(".//ArticleTitle")
        title = "No title"
        if title_elem is not None and title_elem.text is not None and title_elem.text.strip():
            title = title_elem.text.strip()

        # Autoren
        authors = []
        for author in article.findall(".//Author"):
            lastname = author.find("LastName")
            forename = author.find("ForeName")
            if lastname is not None and forename is not None:
                authors.append(f"{forename.text} {lastname.text}")

        authors_str = ", ".join(authors) if authors else "Unknown authors"

        # Journal
        journal_elem = article.find(".//Journal/Title")
        journal = journal_elem.text if journal_elem is not None else "Unknown journal"

        # Jahr
        year_elem = article.find(".//PubDate/Year")
        year = year_elem.text if year_elem is not None else ""

        # Abstract
        abstract_elem = article.find(".//Abstract/AbstractText")
        abstract = abstract_elem.text if abstract_elem is not None else "No abstract available"

        # DOI
        doi = ""
        for article_id in article.findall(".//ArticleId"):
            if article_id.get("IdType") == "doi":
                doi = article_id.text
                break

        return Paper(
            title=title or "No title",
            authors=authors_str or "Unknown authors",
            journal=journal or "Unknown journal",
            year=year or "",
            abstract=abstract or "No abstract available",
            doi=doi or "",
            pubmed_id=pmid or "",
            url=f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else "",
            source="pubmed"
        )

    def _parse_semantic_scholar_paper(self, data: Dict) -> Paper:
        """Parst Semantic Scholar API Response zu Paper"""
        authors = []
        for author in data.get("authors", []):
            authors.append(author.get("name", ""))

        external_ids = data.get("externalIds", {})

        return Paper(
            title=data.get("title") or "No title",
            authors=", ".join(authors) if authors else "Unknown authors",
            journal=data.get("venue") or "Unknown venue",
            year=str(data.get("year") or ""),
            abstract=data.get("abstract") or "No abstract available",
            doi=external_ids.get("DOI", ""),
            pubmed_id=external_ids.get("PubMed", ""),
            url=data.get("url", ""),
            citations=data.get("citationCount", 0),
            source="semantic_scholar"
        )

    def _parse_europe_pmc_paper(self, data: Dict) -> Paper:
        """Parst Europe PMC API Response zu Paper"""
        return Paper(
            title=data.get("title") or "No title",
            authors=data.get("authorString") or "Unknown authors",
            journal=data.get("journalTitle") or "Unknown journal",
            year=str(data.get("pubYear") or ""),
            abstract=data.get("abstractText") or "No abstract available",
            doi=data.get("doi", ""),
            pubmed_id=data.get("pmid", ""),
            url=data.get("fullTextUrlList", {}).get("fullTextUrl", [{}])[0].get("url", ""),
            source="europe_pmc"
        )

    def _remove_duplicates(self, papers: List[Paper]) -> List[Paper]:
        """Entfernt Duplikate basierend auf Titel-Ähnlichkeit"""
        unique_papers = []
        seen_titles = set()

        for paper in papers:
            # Umfassender Sicherheitscheck für None/leere Titel
            if not hasattr(paper, 'title') or paper.title is None or not isinstance(paper.title, str) or paper.title.strip() == "":
                continue

            try:
                # Normalisiere Titel für Vergleich
                normalized_title = re.sub(r'[^\w\s]', '', paper.title.lower()).strip()

                if normalized_title not in seen_titles and len(normalized_title) > 5:
                    seen_titles.add(normalized_title)
                    unique_papers.append(paper)
            except (AttributeError, TypeError) as e:
                # Überspringe fehlerhafte Papers
                continue

        return unique_papers

@require_api_configuration
def show_unified_search_interface():
    """Hauptinterface für die einzige Paper-Suche im System"""
    st.title("🔍 Paper Search - Das einzige Suchmodul")
    st.write("Durchsucht mehrere wissenschaftliche Datenbanken gleichzeitig mit erweiterten Analyse-Funktionen")

    # Initialize searcher
    searcher = UnifiedPaperSearcher()

    # Load settings from API Filter configuration
    search_settings = st.session_state.get("search_settings", {
        "default_max_results": 50,
        "default_sources": ["pubmed", "europe_pmc"],
        "enable_semantic_scholar": False,
        "auto_save_excel": True,
        "auto_send_email": False
    })

    # Show current configuration status
    with st.expander("⚙️ Aktuelle Konfiguration", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.write("**🔍 Standard-Einstellungen:**")
            st.write(f"• Max. Ergebnisse: {search_settings['default_max_results']}")
            st.write(f"• Auto Excel speichern: {'✅' if search_settings['auto_save_excel'] else '❌'}")
            st.write(f"• Auto Email versenden: {'✅' if search_settings['auto_send_email'] else '❌'}")
        with col2:
            st.write("**📚 Konfigurierte Datenbanken:**")
            for source in search_settings['default_sources']:
                st.write(f"• {source.replace('_', ' ').title()}")

        if st.button("🔧 Konfiguration ändern"):
            st.session_state["current_page"] = "📊 Online-API Filter"
            st.rerun()

    # Search interface
    col1, col2 = st.columns([3, 1])

    with col1:
        search_query = st.text_input(
            "🔍 Suchbegriff:",
            placeholder="z.B. BRCA1 breast cancer treatment",
            help="Verwenden Sie spezifische Begriffe für bessere Ergebnisse"
        )

    with col2:
        max_results = st.number_input(
            "Max. Ergebnisse:",
            min_value=10, max_value=200,
            value=search_settings["default_max_results"]
        )

    # Source selection based on configured settings
    st.subheader("📊 Datenbank-Auswahl")
    st.info("ℹ️ Basiert auf der Konfiguration im Online-API Filter. Individuelle Anpassungen für diese Suche möglich.")

    col1, col2, col3 = st.columns(3)

    configured_sources = search_settings["default_sources"]

    with col1:
        use_pubmed = st.checkbox(
            "🏥 PubMed",
            value="pubmed" in configured_sources,
            help="NCBI PubMed - Biomedizinische Literatur"
        )
    with col2:
        use_semantic = st.checkbox(
            "🔬 Semantic Scholar",
            value="semantic_scholar" in configured_sources,
            help="⚠️ Rate Limits - nur bei Bedarf aktivieren"
        )
    with col3:
        use_europe_pmc = st.checkbox(
            "🌍 Europe PMC",
            value="europe_pmc" in configured_sources,
            help="European PubMed Central - Volltext verfügbar"
        )

    # Informative Hinweise
    if use_semantic:
        st.info("ℹ️ **Semantic Scholar**: Kann aufgrund von Rate Limits langsamer sein. Bei Fehlern wird die Quelle übersprungen.")

    if not any([use_pubmed, use_semantic, use_europe_pmc]):
        st.warning("⚠️ Bitte wählen Sie mindestens eine Datenbank aus!")

    # Erweiterte Optionen mit neuen Funktionen
    with st.expander("⚙️ Erweiterte Optionen & Analyse", expanded=True):
        col1, col2, col3 = st.columns(3)

        with col1:
            st.write("**💾 Speichern & Versenden:**")
            save_to_excel = st.checkbox(
                "💾 In Excel speichern",
                value=search_settings["auto_save_excel"]
            )
            send_email = st.checkbox(
                "📧 Email-Benachrichtigung",
                value=search_settings["auto_send_email"] and bool(searcher.email_config)
            )
            if search_settings["auto_send_email"] and not searcher.email_config:
                st.warning("⚠️ Email ist aktiviert, aber nicht konfiguriert")

        with col2:
            st.write("**🤖 ChatGPT-Analyse:**")
            use_chatgpt = st.checkbox(
                "🤖 ChatGPT Rating aktivieren",
                value=bool(searcher.openai_api_key),
                disabled=not searcher.openai_api_key,
                help="Bewertet Papers mit ChatGPT (0-10) und erstellt Zusammenfassungen"
            )
            if not searcher.openai_api_key:
                st.warning("⚠️ OpenAI API Key erforderlich")

        with col3:
            st.write("**📄 Download & Analyse:**")
            check_downloads = st.checkbox(
                "📄 Download-Verfügbarkeit prüfen",
                value=True,
                help="Prüft welche Papers als PDF verfügbar sind"
            )
            online_analysis = st.checkbox(
                "🔍 Online-Analyse verfügbar",
                value=True,
                help="Ermöglicht direkte Online-Analyse der Papers"
            )

    # Search button
    if st.button("🚀 Suche starten", type="primary") and search_query:

        # Build sources list
        sources = []
        if use_pubmed:
            sources.append("pubmed")
        if use_semantic:
            sources.append("semantic_scholar")
        if use_europe_pmc:
            sources.append("europe_pmc")

        if not sources:
            st.error("❌ Bitte wählen Sie mindestens eine Datenbank aus!")
            return

        # Perform search with extended analysis
        with st.spinner("🔍 Durchsuche Datenbanken..."):
            papers = searcher.search_all_sources(search_query, max_results, sources)

        if papers:
            st.success(f"✅ {len(papers)} Papers gefunden!")

            # Check downloadability if requested
            if check_downloads:
                with st.spinner("📄 Prüfe Download-Verfügbarkeit..."):
                    papers = searcher.check_paper_downloadability(papers)

            # ChatGPT analysis if requested
            if use_chatgpt and searcher.openai_api_key:
                with st.spinner("🤖 ChatGPT analysiert Papers..."):
                    papers = searcher.analyze_papers_with_chatgpt(papers, search_query)

            # Save to Excel if requested (now with all analysis data)
            if save_to_excel:
                searcher.save_to_excel(papers, search_query)
                st.success("💾 Alle Daten (inkl. Analysen) in Excel gespeichert!")

            # Send email if requested
            if send_email:
                searcher.send_email_notification(papers, search_query)

            # Show enhanced results
            show_enhanced_paper_results(papers, search_query, use_chatgpt, check_downloads, online_analysis)

            # Excel Template Filling Integration
            if EXCEL_FILLER_AVAILABLE:
                st.markdown("---")
                with st.expander("📊 **Automatisierte Excel-Ausfüllung** - Papers in Vorlage übertragen", expanded=False):
                    st.info("🤖 **Neue Funktion:** Verwandeln Sie ausgewählte Papers automatisch in ausgefüllte Excel-Vorlagen mit Claude AI Analyse!")

                    # Filter papers that have been analyzed (ChatGPT rating > 0)
                    analyzed_papers = [p for p in papers if p.chatgpt_rating > 0 or p.analyzed] if use_chatgpt else papers

                    if analyzed_papers:
                        st.write(f"📋 **{len(analyzed_papers)} Papers verfügbar** für Excel-Ausfüllung")

                        # Show selection interface for Excel filling
                        selected_papers = []

                        # Quick selection options
                        col_sel1, col_sel2, col_sel3 = st.columns(3)

                        with col_sel1:
                            if st.button("⭐ **Top 3 Papers auswählen**"):
                                # Select top 3 by ChatGPT rating or first 3
                                if use_chatgpt:
                                    sorted_papers = sorted(analyzed_papers, key=lambda x: x.chatgpt_rating, reverse=True)
                                    selected_papers = sorted_papers[:3]
                                else:
                                    selected_papers = analyzed_papers[:3]
                                st.session_state["excel_selected_papers"] = selected_papers
                                st.success(f"✅ {len(selected_papers)} Top Papers ausgewählt!")

                        with col_sel2:
                            if st.button("🎯 **Alle bewerteten Papers**") and use_chatgpt:
                                high_rated = [p for p in analyzed_papers if p.chatgpt_rating >= 7.0]
                                if high_rated:
                                    selected_papers = high_rated
                                    st.session_state["excel_selected_papers"] = selected_papers
                                    st.success(f"✅ {len(selected_papers)} hoch bewertete Papers ausgewählt!")
                                else:
                                    st.warning("⚠️ Keine Papers mit Rating ≥7.0 gefunden")

                        with col_sel3:
                            if st.button("📋 **Manuelle Auswahl**"):
                                st.session_state["show_manual_selection"] = True

                        # Manual selection interface
                        if st.session_state.get("show_manual_selection", False):
                            st.subheader("🎯 Manuelle Paper-Auswahl")

                            paper_selections = {}
                            for i, paper in enumerate(analyzed_papers):
                                rating_text = f" (⭐ {paper.chatgpt_rating:.1f})" if paper.chatgpt_rating > 0 else ""
                                paper_selections[i] = st.checkbox(
                                    f"**{paper.title[:60]}...** - {paper.authors[:30]}...{rating_text}",
                                    key=f"paper_select_{i}"
                                )

                            if st.button("✅ **Auswahl bestätigen**"):
                                selected_papers = [analyzed_papers[i] for i, selected in paper_selections.items() if selected]
                                if selected_papers:
                                    st.session_state["excel_selected_papers"] = selected_papers
                                    st.success(f"✅ {len(selected_papers)} Papers für Excel-Ausfüllung ausgewählt!")
                                    st.session_state["show_manual_selection"] = False
                                else:
                                    st.warning("⚠️ Keine Papers ausgewählt!")

                        # Show Excel filling interface if papers are selected
                        excel_papers = st.session_state.get("excel_selected_papers", [])
                        if excel_papers:
                            st.markdown("### 📊 Ausgewählte Papers für Excel-Ausfüllung:")

                            # Show selected papers summary
                            for i, paper in enumerate(excel_papers, 1):
                                rating_display = f" - 🤖 {paper.chatgpt_rating:.1f}/10" if paper.chatgpt_rating > 0 else ""
                                st.write(f"**{i}.** {paper.title}{rating_display}")

                            # Clear selection button
                            if st.button("🗑️ Auswahl zurücksetzen"):
                                st.session_state["excel_selected_papers"] = []
                                if "show_manual_selection" in st.session_state:
                                    del st.session_state["show_manual_selection"]
                                st.rerun()

                            # Excel filling interface
                            show_paper_excel_interface(excel_papers)

                    else:
                        st.info("ℹ️ Aktivieren Sie die ChatGPT-Analyse für beste Excel-Ausfüllung Ergebnisse!")

        else:
            st.warning("⚠️ Keine Papers gefunden. Versuchen Sie andere Suchbegriffe.")

def show_enhanced_paper_results(papers: List[Paper], search_query: str, has_chatgpt: bool, has_downloads: bool, has_online_analysis: bool):
    """Zeigt erweiterte Paper-Ergebnisse mit allen Analyse-Funktionen an"""
    st.subheader(f"📋 Erweiterte Ergebnisse für: '{search_query}'")

    # Enhanced summary stats
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("📊 Total Papers", len(papers))
    with col2:
        sources = set(paper.source for paper in papers)
        st.metric("🔗 Datenbanken", len(sources))
    with col3:
        with_abstract = len([p for p in papers if p.abstract and p.abstract != "No abstract available"])
        st.metric("📄 Mit Abstract", with_abstract)
    with col4:
        downloadable = len([p for p in papers if p.is_downloadable]) if has_downloads else 0
        st.metric("📥 Downloadbar", downloadable)
    with col5:
        analyzed = len([p for p in papers if p.analyzed]) if has_chatgpt else 0
        st.metric("🤖 Analysiert", analyzed)

    # ChatGPT Rating summary if available
    if has_chatgpt and any(p.analyzed for p in papers):
        st.subheader("🏆 Top bewertete Papers (ChatGPT)")
        top_papers = [p for p in papers if p.analyzed and p.chatgpt_rating > 0][:3]
        if top_papers:
            for i, paper in enumerate(top_papers, 1):
                rating_stars = "⭐" * int(paper.chatgpt_rating)
                st.write(f"**{i}. {paper.title}** - {paper.chatgpt_rating:.1f}/10 {rating_stars}")

    # Filter options
    st.subheader("🔍 Filter & Sortierung")
    col1, col2, col3 = st.columns(3)

    with col1:
        sort_by = st.selectbox("Sortieren nach:", [
            "Standard", "ChatGPT Rating", "Zitierungen", "Jahr", "Titel"
        ])
    with col2:
        filter_downloadable = st.checkbox("Nur downloadbare Papers zeigen", value=False) if has_downloads else False
    with col3:
        min_rating = st.slider("Min. ChatGPT Rating:", 0.0, 10.0, 0.0, 0.5) if has_chatgpt else 0.0

    # Apply filters and sorting
    filtered_papers = papers.copy()

    if filter_downloadable:
        filtered_papers = [p for p in filtered_papers if p.is_downloadable]

    if has_chatgpt and min_rating > 0:
        filtered_papers = [p for p in filtered_papers if p.chatgpt_rating >= min_rating]

    # Sort papers
    if sort_by == "ChatGPT Rating" and has_chatgpt:
        filtered_papers.sort(key=lambda p: p.chatgpt_rating, reverse=True)
    elif sort_by == "Zitierungen":
        filtered_papers.sort(key=lambda p: p.citations, reverse=True)
    elif sort_by == "Jahr":
        filtered_papers.sort(key=lambda p: p.year, reverse=True)
    elif sort_by == "Titel":
        filtered_papers.sort(key=lambda p: p.title)

    st.write(f"**Anzeige: {len(filtered_papers)} von {len(papers)} Papers**")

    # Enhanced paper list
    for i, paper in enumerate(filtered_papers, 1):
        # Paper header with rating
        rating_display = ""
        if has_chatgpt and paper.analyzed:
            rating_stars = "⭐" * int(paper.chatgpt_rating)
            rating_display = f" - 🤖 {paper.chatgpt_rating:.1f}/10 {rating_stars}"

        download_icon = " 📥" if (has_downloads and paper.is_downloadable) else ""

        header = f"**{i}. {paper.title}** ({paper.source.upper()}){rating_display}{download_icon}"

        with st.expander(header):
            col1, col2 = st.columns([2, 1])

            with col1:
                st.write(f"**Autoren:** {paper.authors}")
                st.write(f"**Journal:** {paper.journal} ({paper.year})")
                if paper.abstract:
                    st.write(f"**Abstract:** {paper.abstract[:300]}...")

                # ChatGPT summary if available
                if has_chatgpt and paper.analyzed and paper.chatgpt_summary:
                    st.markdown("**🤖 ChatGPT Zusammenfassung:**")
                    st.info(paper.chatgpt_summary)

            with col2:
                if paper.doi:
                    st.write(f"**DOI:** {paper.doi}")
                if paper.pubmed_id:
                    st.write(f"**PubMed ID:** {paper.pubmed_id}")
                if paper.citations > 0:
                    st.write(f"**Zitierungen:** {paper.citations}")

                # Action buttons
                button_col1, button_col2 = st.columns(2)

                with button_col1:
                    if paper.url:
                        st.link_button("🔗 Online lesen", paper.url, width="stretch")

                with button_col2:
                    if has_downloads and paper.is_downloadable and paper.pdf_url:
                        st.link_button("📥 PDF Download", paper.pdf_url, width="stretch")

                # Online analysis button
                if has_online_analysis:
                    if st.button(f"🔬 Analysieren", key=f"analyze_{i}"):
                        analyze_paper_online(paper)

def analyze_paper_online(paper: Paper):
    """Online-Analyse eines Papers"""
    st.subheader(f"🔬 Online-Analyse: {paper.title}")

    tab1, tab2, tab3 = st.tabs(["📊 Statistiken", "🔍 Details", "🤖 AI-Insights"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Jahr", paper.year)
            st.metric("Zitierungen", paper.citations)
            st.metric("Quelle", paper.source.upper())
        with col2:
            if paper.analyzed:
                st.metric("ChatGPT Rating", f"{paper.chatgpt_rating:.1f}/10")
            st.metric("Download verfügbar", "✅" if paper.is_downloadable else "❌")

    with tab2:
        st.write("**Vollständiger Abstract:**")
        st.write(paper.abstract)
        if paper.keywords:
            st.write(f"**Keywords:** {paper.keywords}")

    with tab3:
        if paper.analyzed and paper.chatgpt_summary:
            st.write("**ChatGPT Zusammenfassung:**")
            st.info(paper.chatgpt_summary)
        else:
            st.write("Keine AI-Analyse verfügbar")

# Integration function for main app
def module_unified_search():
    """Main module function for integration"""
    show_unified_search_interface()