import os
import sys
from io import BytesIO

import streamlit as st
from docx import Document

# zorg dat we modules in dezelfde map kunnen importeren
sys.path.append(os.path.dirname(__file__))

from html_converter import docx_to_html
from pptx_converter_hybrid import docx_to_pptx_hybrid

# LES-AI (optioneel): deze maakt een les-Word via AI als lesson_from_docx.py bestaat
LESSON_ANALYZER_ERROR = None
try:
    from lesson_from_docx import docx_to_vmbo_lesson_json  # deze geeft een DOCX (BytesIO) terug
    HAS_LESSON_ANALYZER = True
except Exception as e:
    HAS_LESSON_ANALYZER = False
    LESSON_ANALYZER_ERROR = str(e)


st.set_page_config(page_title="Triade DOCX Tools", page_icon="📘", layout="wide")
st.title("📘 Triade DOCX Tools")

tab1, tab2, tab3 = st.tabs([
    "💚 HTML (Stermonitor / LessonUp)",
    "🤖 PowerPoint / Les-Word",
    "📘 Werkboekje generator",
])

# ==========================================================
# TAB 1: DOCX → HTML
# ==========================================================
with tab1:
    st.subheader("DOCX → HTML Converter")
    st.caption("Zet je Word-lesstof automatisch om naar nette HTML voor Stermonitor of LessonUp.")

    uploaded_html = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="html_upload")

    if uploaded_html:
        with st.spinner("Word-bestand wordt omgezet..."):
            try:
                html_out = docx_to_html(uploaded_html)
            except Exception as e:
                st.error(f"❌ Er ging iets mis bij het omzetten naar HTML: {e}")
            else:
                st.success("✅ Klaar! HTML gegenereerd.")
                st.code(html_out, language="html")
                st.download_button(
                    "⬇️ Download HTML-bestand",
                    data=html_out,
                    file_name="les_stermonitor.html",
                    mime="text/html",
                )
    else:
        st.info("Upload een .docx-bestand om te converteren naar HTML.")


# ==========================================================
# TAB 2: DOCX → PowerPoint én (optioneel) les-Word via AI
# ==========================================================
with tab2:
    st.subheader("DOCX → PowerPoint (AI) / Les-Word")
    st.caption("Maak een PowerPoint in jouw layout, of laat AI eerst een les-Word in VMBO-stijl maken.")

    uploaded_ai = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="hybrid_upload")

    if not HAS_LESSON_ANALYZER:
        msg = "⚠️ De AI-lesmodule (`lesson_from_docx.py`) kon niet worden geladen."
        if LESSON_ANALYZER_ERROR:
            msg += f"\n\n**Details:** {LESSON_ANALYZER_ERROR}"
        st.warning(msg)

    if uploaded_ai:
        col1, col2 = st.columns(2)

        # --------- les-Word maken (als module aanwezig is) ----------
        with col1:
            st.markdown("**Les-Word laten maken (AI)**")
            st.caption("Maakt een nieuw .docx in les-format (kop, uitleg, vraag).")
            if st.button("📝 Maak les-Word"):
                if not HAS_LESSON_ANALYZER:
                    st.error("Les-generatie is niet beschikbaar.")
                else:
                    with st.spinner("Les wordt door AI opgebouwd..."):
                        try:
                            lesson_docx = docx_to_vmbo_lesson_json(uploaded_ai)
                        except Exception as e:
                            st.error(f"❌ Kon geen les-Word maken: {e}")
                        else:
                            st.success("✅ Les-Word gemaakt.")
                            st.download_button(
                                "⬇️ Download les_vmbo.docx",
                                data=lesson_docx,
                                file_name="les_vmbo.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            )

        # --------- PowerPoint maken ----------
        with col2:
            st.markdown("**PowerPoint maken in vaste layout**")
            st.caption("Gebruikt je basis PowerPoint-template.")
            if st.button("📽️ Maak PowerPoint"):
                with st.spinner("PowerPoint wordt opgebouwd..."):
                    try:
                        pptx_bytes = docx_to_pptx_hybrid(uploaded_ai)
                    except Exception as e:
                        st.error(f"❌ Kon geen PowerPoint maken: {e}")
                    else:
                        st.success("✅ Klaar! PowerPoint gegenereerd.")
                        st.download_button(
                            "⬇️ Download PowerPoint",
                            data=pptx_bytes,
                            file_name="les_ai_hybride.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )
    else:
        st.info("Upload een .docx-bestand om een PowerPoint of les-Word te maken.")


# ==========================================================
# TAB 3: Werkboekje generator
# ==========================================================
with tab3:
    st.subheader("📘 Werkboekje generator")
    st.caption(
        "Maak een werkboekje met voorpagina en daarna losse stappen. "
        "Je kunt per stap kiezen hoeveel afbeeldingen/teksten je wilt."
    )

    # --- algemene info ---
    col_a, col_b = st.columns(2)
    with col_a:
        wb_docent = st.text_input("Docent", key="wb_docent")
        wb_project = st.text_input("Project / opdracht", key="wb_project")
    with col_b:
        wb_cover = st.file_uploader("Omslag-afbeelding (optioneel)", type=["png", "jpg", "jpeg"], key="wb_cover")

    # --- init stappen in session_state ---
    if "wb_steps" not in st.session_state:
        st.session_state.wb_steps = []

    st.markdown("### Stappen")

    if st.button("➕ Nieuwe stap"):
        st.session_state.wb_steps.append(
            {
                "layout": "1 afbeelding + tekst",
            }
        )

    # stappen tonen/bewerken
    for idx, _ in enumerate(st.session_state.wb_steps):
        st.markdown(f"#### Stap {idx+1}")

        layout = st.selectbox(
            "Kies layout",
            ["1 afbeelding + tekst", "2 afbeeldingen + 2 teksten", "3 afbeeldingen + 3 teksten"],
            key=f"wb_layout_{idx}",
        )
        title = st.text_input("Titel van deze stap", key=f"wb_title_{idx}")

        # bepaal hoeveel blokken we laten zien
        if layout.startswith("1"):
            max_blocks = 1
        elif layout.startswith("2"):
            max_blocks = 2
        else:
            max_blocks = 3

        for i in range(max_blocks):
            st.markdown(f"**Blok {i+1}**")
            st.file_uploader(
                f"Afbeelding {i+1}",
                type=["png", "jpg", "jpeg"],
                key=f"wb_img_{idx}_{i}",
            )
            st.text_area(f"Tekst {i+1}", key=f"wb_txt_{idx}_{i}")

        st.divider()

    # --- werkboekje genereren als docx ---
    if st.button("📄 Maak werkboekje (DOCX)"):
        doc = Document()

        # voorpagina
        if wb_project:
            doc.add_heading(wb_project, level=0)
        if wb_docent:
            doc.add_paragraph(f"Docent: {wb_docent}")
        doc.add_paragraph(" ")

        # (optioneel: wb_cover hier invoegen met doc.add_picture if you want)

        # stappen
        for idx, _ in enumerate(st.session_state.wb_steps):
            doc.add_heading(f"Stap {idx+1}", level=1)
            titel = st.session_state.get(f"wb_title_{idx}", "")
            if titel:
                doc.add_paragraph(titel)

            layout = st.session_state.get(f"wb_layout_{idx}", "1 afbeelding + tekst")
            if layout.startswith("1"):
                max_blocks = 1
            elif layout.startswith("2"):
                max_blocks = 2
            else:
                max_blocks = 3

            for i in range(max_blocks):
                txt = st.session_state.get(f"wb_txt_{idx}_{i}", "")
                if txt:
                    doc.add_paragraph(txt)

            doc.add_paragraph("")

        out = BytesIO()
        doc.save(out)
        out.seek(0)

        st.success("✅ Werkboekje gemaakt.")
        st.download_button(
            "⬇️ Download werkboekje.docx",
            data=out,
            file_name="werkboekje.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

