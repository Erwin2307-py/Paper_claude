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

@dataclass
class Paper:
    """Einheitliche Paper-Datenstruktur"""
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
            'date_added': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

class UnifiedPaperSearcher:
    """Zentrales Paper-Search-System"""

    def __init__(self):
        self.excel_manager = initialize_excel_manager()
        self.email_config = load_email_config_from_secrets()

    def search_all_sources(self, query: str, max_results: int = 50,
                          sources: List[str] = None) -> List[Paper]:
        """Sucht in allen verfügbaren Quellen"""
        if sources is None:
            sources = ["pubmed", "semantic_scholar", "europe_pmc"]

        all_papers = []

        # Progress tracking
        progress_bar = st.progress(0)
        status_text = st.empty()

        total_sources = len(sources)

        for i, source in enumerate(sources):
            status_text.text(f"🔍 Suche in {source.upper()}...")
            progress = (i + 0.5) / total_sources
            progress_bar.progress(progress)

            try:
                if source == "pubmed":
                    papers = self.search_pubmed(query, max_results // len(sources))
                elif source == "semantic_scholar":
                    papers = self.search_semantic_scholar(query, max_results // len(sources))
                elif source == "europe_pmc":
                    papers = self.search_europe_pmc(query, max_results // len(sources))
                else:
                    papers = []

                all_papers.extend(papers)

            except Exception as e:
                st.warning(f"⚠️ Fehler bei {source}: {str(e)}")

            progress = (i + 1) / total_sources
            progress_bar.progress(progress)

        # Duplikate entfernen basierend auf Titel
        unique_papers = self._remove_duplicates(all_papers)

        status_text.text(f"✅ {len(unique_papers)} einzigartige Papers gefunden!")

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
                    papers.append(paper)
                except Exception as e:
                    continue  # Skip fehlerhafter Artikel

        except Exception as e:
            st.error(f"PubMed-Suche fehlgeschlagen: {str(e)}")

        return papers

    def search_semantic_scholar(self, query: str, max_results: int = 20) -> List[Paper]:
        """Sucht in Semantic Scholar"""
        papers = []

        url = "https://api.semanticscholar.org/graph/v1/paper/search"
        params = {
            "query": query,
            "limit": max_results,
            "fields": "title,authors,venue,year,abstract,url,citationCount,externalIds"
        }

        try:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            data = response.json()

            for paper_data in data.get("data", []):
                try:
                    paper = self._parse_semantic_scholar_paper(paper_data)
                    papers.append(paper)
                except Exception:
                    continue

        except Exception as e:
            st.error(f"Semantic Scholar-Suche fehlgeschlagen: {str(e)}")

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

    def _parse_pubmed_article(self, article) -> Paper:
        """Parst PubMed XML Article zu Paper"""
        # PMID
        pmid_elem = article.find(".//PMID")
        pmid = pmid_elem.text if pmid_elem is not None else ""

        # Titel
        title_elem = article.find(".//ArticleTitle")
        title = title_elem.text if title_elem is not None else "No title"

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
            title=title,
            authors=authors_str,
            journal=journal,
            year=year,
            abstract=abstract,
            doi=doi,
            pubmed_id=pmid,
            url=f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/",
            source="pubmed"
        )

    def _parse_semantic_scholar_paper(self, data: Dict) -> Paper:
        """Parst Semantic Scholar API Response zu Paper"""
        authors = []
        for author in data.get("authors", []):
            authors.append(author.get("name", ""))

        external_ids = data.get("externalIds", {})

        return Paper(
            title=data.get("title", "No title"),
            authors=", ".join(authors),
            journal=data.get("venue", "Unknown venue"),
            year=str(data.get("year", "")),
            abstract=data.get("abstract", "No abstract available"),
            doi=external_ids.get("DOI", ""),
            pubmed_id=external_ids.get("PubMed", ""),
            url=data.get("url", ""),
            citations=data.get("citationCount", 0),
            source="semantic_scholar"
        )

    def _parse_europe_pmc_paper(self, data: Dict) -> Paper:
        """Parst Europe PMC API Response zu Paper"""
        return Paper(
            title=data.get("title", "No title"),
            authors=data.get("authorString", "Unknown authors"),
            journal=data.get("journalTitle", "Unknown journal"),
            year=str(data.get("pubYear", "")),
            abstract=data.get("abstractText", "No abstract available"),
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
            # Normalisiere Titel für Vergleich
            normalized_title = re.sub(r'[^\w\s]', '', paper.title.lower()).strip()

            if normalized_title not in seen_titles and len(normalized_title) > 5:
                seen_titles.add(normalized_title)
                unique_papers.append(paper)

        return unique_papers

def show_unified_search_interface():
    """Hauptinterface für einheitliche Paper-Suche"""
    st.title("🔍 Einheitliche Paper-Suche")
    st.write("Durchsucht mehrere wissenschaftliche Datenbanken gleichzeitig")

    # Initialize searcher
    searcher = UnifiedPaperSearcher()

    # Search interface
    col1, col2 = st.columns([3, 1])

    with col1:
        search_query = st.text_input(
            "🔍 Suchbegriff:",
            placeholder="z.B. BRCA1 breast cancer treatment",
            help="Verwenden Sie spezifische Begriffe für bessere Ergebnisse"
        )

    with col2:
        max_results = st.number_input("Max. Ergebnisse:", min_value=10, max_value=200, value=50)

    # Source selection
    st.subheader("📊 Datenbank-Auswahl")
    col1, col2, col3 = st.columns(3)

    with col1:
        use_pubmed = st.checkbox("🏥 PubMed", value=True)
    with col2:
        use_semantic = st.checkbox("🔬 Semantic Scholar", value=True)
    with col3:
        use_europe_pmc = st.checkbox("🌍 Europe PMC", value=True)

    # Options
    with st.expander("⚙️ Erweiterte Optionen"):
        col1, col2 = st.columns(2)
        with col1:
            save_to_excel = st.checkbox("💾 In Excel speichern", value=True)
        with col2:
            send_email = st.checkbox("📧 Email-Benachrichtigung", value=bool(searcher.email_config))

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

        # Perform search
        with st.spinner("🔍 Durchsuche Datenbanken..."):
            papers = searcher.search_all_sources(search_query, max_results, sources)

        if papers:
            # Display results
            st.success(f"✅ {len(papers)} Papers gefunden!")

            # Save to Excel if requested
            if save_to_excel:
                searcher.save_to_excel(papers, search_query)

            # Send email if requested
            if send_email:
                searcher.send_email_notification(papers, search_query)

            # Show results
            show_paper_results(papers, search_query)
        else:
            st.warning("⚠️ Keine Papers gefunden. Versuchen Sie andere Suchbegriffe.")

def show_paper_results(papers: List[Paper], search_query: str):
    """Zeigt Paper-Ergebnisse an"""
    st.subheader(f"📋 Ergebnisse für: '{search_query}'")

    # Summary stats
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📊 Total Papers", len(papers))
    with col2:
        sources = set(paper.source for paper in papers)
        st.metric("🔗 Datenbanken", len(sources))
    with col3:
        with_abstract = len([p for p in papers if p.abstract and p.abstract != "No abstract available"])
        st.metric("📄 Mit Abstract", with_abstract)
    with col4:
        with_doi = len([p for p in papers if p.doi])
        st.metric("🔗 Mit DOI", with_doi)

    # Paper list
    for i, paper in enumerate(papers, 1):
        with st.expander(f"**{i}. {paper.title}** ({paper.source.upper()})"):
            col1, col2 = st.columns([2, 1])

            with col1:
                st.write(f"**Autoren:** {paper.authors}")
                st.write(f"**Journal:** {paper.journal} ({paper.year})")
                if paper.abstract:
                    st.write(f"**Abstract:** {paper.abstract[:300]}...")

            with col2:
                if paper.doi:
                    st.write(f"**DOI:** {paper.doi}")
                if paper.pubmed_id:
                    st.write(f"**PubMed ID:** {paper.pubmed_id}")
                if paper.citations > 0:
                    st.write(f"**Zitierungen:** {paper.citations}")
                if paper.url:
                    st.link_button("🔗 Paper öffnen", paper.url)

# Integration function for main app
def module_unified_search():
    """Main module function for integration"""
    show_unified_search_interface()