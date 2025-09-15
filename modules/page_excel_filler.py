# modules/page_excel_filler.py - Standalone Page for Paper Excel Filler
import streamlit as st
import datetime
from typing import List, Dict, Any, Optional

# Try to import the Excel Filler functionality
try:
    from modules.paper_excel_filler import PaperExcelFiller, show_paper_excel_interface
    from modules.unified_paper_search import Paper
    EXCEL_FILLER_AVAILABLE = True
except ImportError as e:
    st.error(f"Excel Filler Module konnte nicht importiert werden: {e}")
    EXCEL_FILLER_AVAILABLE = False


def create_sample_papers() -> List[Paper]:
    """Erstellt Beispiel-Papers für Demonstrationszwecke"""
    sample_papers = [
        Paper(
            title="BRCA1 mutations and breast cancer risk in European populations",
            authors="Smith, J.A., Johnson, B.C., Wilson, D.E., Brown, K.L.",
            journal="Nature Genetics",
            year="2023",
            abstract="This comprehensive study examines the association between BRCA1 mutations and breast cancer risk across multiple European populations. We analyzed genetic data from 50,000 women and identified significant risk variants. Our findings demonstrate that certain BRCA1 mutations confer up to 8-fold increased breast cancer risk. The study provides crucial insights for personalized medicine approaches.",
            doi="10.1038/s41588-023-01234-5",
            pubmed_id="37123456",
            keywords="BRCA1, breast cancer, genetic risk, European populations, mutations",
            chatgpt_rating=9.2,
            chatgpt_summary="Excellent large-scale genetic study with robust methodology and significant clinical implications for breast cancer risk assessment.",
            source="pubmed",
            citations=142,
            is_downloadable=True,
            analyzed=True
        ),
        Paper(
            title="TP53 pathway alterations in colorectal cancer: therapeutic implications",
            authors="Martinez, R., Chen, L., Thompson, M.J., Davis, P.K.",
            journal="Cell",
            year="2023",
            abstract="The TP53 tumor suppressor pathway plays a critical role in colorectal cancer development. This study investigates TP53 mutations in 1,200 colorectal cancer samples and their therapeutic implications. We identified novel therapeutic targets and drug resistance mechanisms. Our results suggest personalized treatment strategies based on TP53 status.",
            doi="10.1016/j.cell.2023.05.012",
            pubmed_id="37234567",
            keywords="TP53, colorectal cancer, tumor suppressor, therapeutic targets",
            chatgpt_rating=8.7,
            chatgpt_summary="Important study linking TP53 mutations to therapeutic responses in colorectal cancer with clinical translation potential.",
            source="pubmed",
            citations=98,
            is_downloadable=True,
            analyzed=True
        ),
        Paper(
            title="EGFR inhibitor resistance mechanisms in lung adenocarcinoma",
            authors="Lee, H.S., Kumar, A., Patel, N., Anderson, C.R.",
            journal="Science",
            year="2023",
            abstract="Resistance to EGFR inhibitors remains a major challenge in lung cancer treatment. We performed comprehensive genomic analysis of 800 resistant tumors and identified key resistance mechanisms. Novel combination therapies are proposed to overcome resistance. This work has immediate clinical applications for improving patient outcomes.",
            doi="10.1126/science.abc1234",
            pubmed_id="37345678",
            keywords="EGFR, lung cancer, resistance mechanisms, targeted therapy",
            chatgpt_rating=9.5,
            chatgpt_summary="Groundbreaking research on EGFR inhibitor resistance with direct clinical applications and novel therapeutic strategies.",
            source="pubmed",
            citations=203,
            is_downloadable=True,
            analyzed=True
        ),
        Paper(
            title="Genetic variants in DNA repair genes and cancer susceptibility",
            authors="Zhang, W., Roberts, M., Taylor, S.L., White, J.M.",
            journal="Nature Medicine",
            year="2023",
            abstract="DNA repair mechanisms are crucial for maintaining genomic stability. This meta-analysis examines genetic variants in DNA repair genes across multiple cancer types. We analyzed data from 100,000 patients and identified 15 high-risk variants. The findings provide insights for cancer prevention strategies.",
            doi="10.1038/s41591-023-02345-6",
            pubmed_id="37456789",
            keywords="DNA repair, genetic variants, cancer susceptibility, genomic stability",
            chatgpt_rating=8.3,
            chatgpt_summary="Comprehensive meta-analysis providing valuable insights into DNA repair gene variants and cancer risk across multiple cancer types.",
            source="europe_pmc",
            citations=156,
            is_downloadable=False,
            analyzed=True
        )
    ]
    return sample_papers


def show_excel_filler_page():
    """Hauptseite für Paper Excel Filler"""
    st.title("📊 Paper Excel Filler")
    st.write("Automatisierte Excel-Ausfüllung für wissenschaftliche Papers mit Claude AI Analyse")

    if not EXCEL_FILLER_AVAILABLE:
        st.error("❌ **Excel Filler Modul nicht verfügbar!**")
        st.write("**Mögliche Ursachen:**")
        st.write("• Fehlende Module-Dependencies")
        st.write("• Import-Fehler in paper_excel_filler.py")
        st.write("• Fehlende Excel-Vorlagen")

        if st.button("🔄 Seite neu laden"):
            st.rerun()
        return

    # Initialize Excel Filler
    filler = PaperExcelFiller()

    # Status Dashboard
    st.subheader("🔧 System Status")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        api_status = "✅ Verfügbar" if filler.claude_api_key else "❌ Fehlt"
        st.metric("Claude API", api_status)

    with col2:
        # Check for Excel templates
        template_count = 0
        template_paths = ["vorlage_paperqa2.xlsx", "modules/vorlage_paperqa2.xlsx", "vorlage_gene.xlsx"]
        for path in template_paths:
            import os
            if os.path.exists(path):
                template_count += 1
        st.metric("Excel Vorlagen", template_count)

    with col3:
        # Check session state for papers
        session_papers = len(st.session_state.get("search_results", []))
        st.metric("Geladene Papers", session_papers)

    with col4:
        filled_count = len(st.session_state.get("filled_excels", []))
        st.metric("Erstellte Excel", filled_count)

    # Configuration Section
    with st.expander("⚙️ Konfiguration & Setup", expanded=False):
        col_cfg1, col_cfg2 = st.columns(2)

        with col_cfg1:
            st.write("**📋 Erforderliche Komponenten:**")
            st.write("✅ paper_excel_filler.py Modul")
            st.write("✅ unified_paper_search.py Integration")
            st.write(f"{'✅' if filler.claude_api_key else '❌'} Claude API Key")
            st.write(f"{'✅' if template_count > 0 else '❌'} Excel-Vorlagen")

        with col_cfg2:
            st.write("**🔧 Claude API Setup:**")
            st.code("""
# Streamlit Secrets (.streamlit/secrets.toml)
[claude]
api_key = "your_claude_api_key"

[anthropic]
api_key = "your_claude_api_key"
            """, language="toml")

        if not filler.claude_api_key:
            st.warning("⚠️ **Claude API Key nicht konfiguriert!** Fallback-Analyse wird verwendet.")
            st.info("💡 Konfigurieren Sie den API Key in den Streamlit Secrets für beste Ergebnisse.")

    # Main functionality tabs
    tab1, tab2, tab3, tab4 = st.tabs(["📋 Paper laden", "🎯 Paper auswählen", "📊 Excel erstellen", "📈 Statistiken"])

    with tab1:
        st.subheader("📋 Paper-Datenquelle")

        source_option = st.radio(
            "Wählen Sie Ihre Paper-Quelle:",
            [
                "🔍 Aus aktueller Suche (Unified Search)",
                "📝 Beispiel-Papers laden",
                "📁 Manuelle Paper-Eingabe"
            ]
        )

        if source_option == "🔍 Aus aktueller Suche (Unified Search)":
            # Load from session state
            search_papers = st.session_state.get("search_results", [])

            if search_papers:
                st.success(f"✅ {len(search_papers)} Papers aus der aktuellen Suche geladen!")
                st.session_state["excel_source_papers"] = search_papers

                # Show preview
                with st.expander("👀 Paper-Vorschau"):
                    for i, paper in enumerate(search_papers[:3], 1):
                        rating_text = f" - ⭐ {paper.chatgpt_rating:.1f}" if hasattr(paper, 'chatgpt_rating') and paper.chatgpt_rating > 0 else ""
                        st.write(f"**{i}.** {paper.title}{rating_text}")
                    if len(search_papers) > 3:
                        st.write(f"... und {len(search_papers) - 3} weitere Papers")

            else:
                st.info("ℹ️ Keine Papers in der aktuellen Session gefunden.")
                st.write("**Anleitung:**")
                st.write("1. Gehen Sie zu **🔍 Paper Search**")
                st.write("2. Führen Sie eine Suche durch")
                st.write("3. Aktivieren Sie **ChatGPT-Analyse** für beste Ergebnisse")
                st.write("4. Kehren Sie hier zurück")

                if st.button("🔍 Zur Paper Search"):
                    st.session_state["current_page"] = "🔍 Paper Search"
                    st.rerun()

        elif source_option == "📝 Beispiel-Papers laden":
            if st.button("🔄 Beispiel-Papers laden"):
                sample_papers = create_sample_papers()
                st.session_state["excel_source_papers"] = sample_papers
                st.success(f"✅ {len(sample_papers)} Beispiel-Papers geladen!")
                st.rerun()

            # Show loaded sample papers
            sample_papers = st.session_state.get("excel_source_papers", [])
            if sample_papers and all(hasattr(p, 'chatgpt_rating') for p in sample_papers):
                st.info("📝 Beispiel-Papers sind geladen und bereit für Excel-Erstellung!")

        elif source_option == "📁 Manuelle Paper-Eingabe":
            st.info("🚧 **Coming Soon:** Manuelle Paper-Eingabe wird in einer zukünftigen Version verfügbar sein.")
            st.write("**Aktuell verfügbare Optionen:**")
            st.write("• Paper aus Unified Search laden")
            st.write("• Beispiel-Papers für Testzwecke verwenden")

    with tab2:
        st.subheader("🎯 Paper-Auswahl für Excel-Erstellung")

        source_papers = st.session_state.get("excel_source_papers", [])

        if not source_papers:
            st.warning("⚠️ Keine Papers geladen. Wechseln Sie zum Tab 'Paper laden'.")
        else:
            st.write(f"📋 **{len(source_papers)} Papers verfügbar**")

            # Selection methods
            selection_method = st.selectbox(
                "Auswahl-Methode:",
                [
                    "🎯 Manuelle Einzelauswahl",
                    "⭐ Top-bewertete Papers (Rating ≥8.0)",
                    "🔝 Beste 3 Papers",
                    "📊 Alle Papers"
                ]
            )

            selected_papers = []

            if selection_method == "🎯 Manuelle Einzelauswahl":
                st.write("**📋 Wählen Sie Papers aus:**")
                selections = {}

                for i, paper in enumerate(source_papers):
                    rating_display = f" - ⭐ {paper.chatgpt_rating:.1f}/10" if hasattr(paper, 'chatgpt_rating') and paper.chatgpt_rating > 0 else ""
                    journal_year = f" ({paper.journal}, {paper.year})" if paper.journal and paper.year else ""

                    selections[i] = st.checkbox(
                        f"**{paper.title[:70]}...** {journal_year}{rating_display}",
                        key=f"manual_select_{i}"
                    )

                selected_papers = [source_papers[i] for i, selected in selections.items() if selected]

            elif selection_method == "⭐ Top-bewertete Papers (Rating ≥8.0)":
                high_rated = [p for p in source_papers if hasattr(p, 'chatgpt_rating') and p.chatgpt_rating >= 8.0]
                selected_papers = high_rated
                if selected_papers:
                    st.success(f"✅ {len(selected_papers)} hoch-bewertete Papers automatisch ausgewählt!")
                else:
                    st.info("ℹ️ Keine Papers mit Rating ≥8.0 gefunden.")

            elif selection_method == "🔝 Beste 3 Papers":
                if any(hasattr(p, 'chatgpt_rating') for p in source_papers):
                    sorted_papers = sorted(
                        [p for p in source_papers if hasattr(p, 'chatgpt_rating')],
                        key=lambda x: x.chatgpt_rating,
                        reverse=True
                    )
                    selected_papers = sorted_papers[:3]
                else:
                    selected_papers = source_papers[:3]
                st.success(f"✅ Top {len(selected_papers)} Papers automatisch ausgewählt!")

            elif selection_method == "📊 Alle Papers":
                selected_papers = source_papers
                st.success(f"✅ Alle {len(selected_papers)} Papers ausgewählt!")

            # Show selected papers
            if selected_papers:
                st.session_state["excel_selected_papers"] = selected_papers

                st.markdown("### 📋 Ausgewählte Papers:")
                for i, paper in enumerate(selected_papers, 1):
                    rating_text = f" - ⭐ {paper.chatgpt_rating:.1f}/10" if hasattr(paper, 'chatgpt_rating') and paper.chatgpt_rating > 0 else ""
                    st.write(f"**{i}.** {paper.title}{rating_text}")

                # Clear selection
                if st.button("🗑️ Auswahl zurücksetzen"):
                    st.session_state["excel_selected_papers"] = []
                    st.rerun()

    with tab3:
        st.subheader("📊 Excel-Dateien erstellen")

        selected_papers = st.session_state.get("excel_selected_papers", [])

        if not selected_papers:
            st.info("ℹ️ Keine Papers für Excel-Erstellung ausgewählt.")
            st.write("👉 Wechseln Sie zum Tab **'Paper auswählen'** um Papers auszuwählen.")
        else:
            # Show Excel interface
            show_paper_excel_interface(selected_papers)

    with tab4:
        st.subheader("📈 Nutzungsstatistiken")

        # Statistics from session state
        filled_excels = st.session_state.get("filled_excels", [])

        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)

        with col_stat1:
            total_papers = len(st.session_state.get("excel_source_papers", []))
            st.metric("Geladene Papers", total_papers)

        with col_stat2:
            selected_count = len(st.session_state.get("excel_selected_papers", []))
            st.metric("Ausgewählte Papers", selected_count)

        with col_stat3:
            st.metric("Erstellte Excel", len(filled_excels))

        with col_stat4:
            api_calls = st.session_state.get("claude_api_calls", 0)
            st.metric("Claude API Aufrufe", api_calls)

        # Recent activity
        if filled_excels:
            st.subheader("📋 Kürzlich erstellte Excel-Dateien")
            for excel_info in filled_excels[-5:]:  # Last 5
                created_time = excel_info.get('created_at', 'Unbekannt')
                filename = excel_info.get('filename', 'Unbekannt')
                paper_title = excel_info.get('paper_title', 'Unbekannt')[:50]

                st.write(f"📄 **{filename}** - {paper_title}... ({created_time})")

        else:
            st.info("📊 Noch keine Excel-Dateien erstellt.")

    # Footer
    st.markdown("---")
    st.markdown("### 💡 Tipps für beste Ergebnisse:")
    st.write("🎯 **Verwenden Sie ChatGPT-analysierte Papers** für detaillierteste Excel-Ausfüllung")
    st.write("🧬 **Gene werden automatisch erkannt** aus Titel und Abstract")
    st.write("📊 **Claude API** liefert die besten Analyse-Ergebnisse")
    st.write("💾 **Alle Formate bleiben erhalten** - Excel-Vorlagen werden 1:1 kopiert")


# Main page function for navigation
def page_excel_filler():
    """Main page function for Excel Filler module"""
    show_excel_filler_page()