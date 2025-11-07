import os
import sys
import streamlit as st

# --- zorg dat Python modules in dezelfde map kan vinden ---
sys.path.append(os.path.dirname(__file__))

# --- vaste imports ---
from html_converter import docx_to_html
from pptx_converter_hybrid import docx_to_pptx_hybrid

# --- probeer de AI-lesmodule te laden ---
LESSON_ANALYZER_ERROR = None
try:
    from lesson_from_docx import docx_to_vmbo_lesson_json
    HAS_LESSON_ANALYZER = True
except Exception as e:
    HAS_LESSON_ANALYZER = False
    LESSON_ANALYZER_ERROR = str(e)

# --- pagina setup ---
st.set_page_config(page_title="Triade DOCX Tools", page_icon="📘", layout="wide")
st.title("📘 Triade DOCX → HTML / PowerPoint")

tab1, tab2 = st.tabs(["💚 HTML (Stermonitor / LessonUp)", "🤖 PowerPoint (AI-vmbo)"])

# ==========================================================
# TAB 1: DOCX → HTML converter
# ==========================================================
with tab1:
    st.subheader("DOCX → HTML Converter")
    st.caption("Zet je Word-lesstof automatisch om naar nette HTML voor Stermonitor of LessonUp.")

    uploaded_html = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="html_upload")

    if uploaded_html:
        with st.spinner("Word-bestand wordt omgezet..."):
            try:
                html_out = docx_to_html(uploaded_html)
                st.success("✅ Klaar! HTML gegenereerd.")
                st.code(html_out, language="html")
                st.download_button(
                    "⬇️ Download HTML-bestand",
                    data=html_out,
                    file_name="les_stermonitor.html",
                    mime="text/html",
                )
            except Exception as e:
                st.error(f"❌ Er ging iets mis bij het omzetten naar HTML: {e}")
    else:
        st.info("Upload een .docx-bestand om te converteren naar HTML.")

# ==========================================================
# TAB 2: DOCX → PowerPoint (AI of direct)
# ==========================================================
with tab2:
    st.subheader("DOCX → PowerPoint (AI-vmbo)")
    st.caption("Maakt een PowerPoint in jouw layout. AI wordt gebruikt als die beschikbaar is.")

    uploaded_ai = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="hybrid_upload")

    if not HAS_LESSON_ANALYZER:
        msg = "⚠️ De AI-analyse module (`lesson_from_docx.py`) kon niet worden geladen."
        if LESSON_ANALYZER_ERROR:
            msg += f"\n\n**Details:** {LESSON_ANALYZER_ERROR}"
        st.warning(msg)

    if uploaded_ai:
        # ------------------------------------------------------
        # 1. Eerst analyseren met AI (optioneel)
        # ------------------------------------------------------
        if HAS_LESSON_ANALYZER:
            if st.button("🧠 Analyseer met AI (lesstructuur bekijken)"):
                with st.spinner("AI is de les aan het opdelen in dia’s..."):
                    try:
                        lesson_json = docx_to_vmbo_lesson_json(uploaded_ai)
                        st.success("✅ Lesstructuur gemaakt!")
                        st.json(lesson_json)
                    except Exception as e:
                        st.error(f"❌ Kon lesstructuur niet maken: {e}")

        # ------------------------------------------------------
        # 2. Altijd PowerPoint kunnen maken
        # ------------------------------------------------------
        if st.button("📽️ Maak PowerPoint", type="primary"):
            with st.spinner("PowerPoint wordt opgebouwd..."):
                try:
                    pptx_bytes = docx_to_pptx_hybrid(uploaded_ai)
                except Exception as e:
                    st.error(f"❌ Kon geen PowerPoint maken: {e}")
                else:
                    st.success("✅ Klaar! PowerPoint gegenereerd.")
                    st.download_button(
                        "⬇️ Download PowerPoint (AI-hybride)",
                        data=pptx_bytes,
                        file_name="les_ai_hybride.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
    else:
        st.info("Upload een .docx-bestand om een PowerPoint te maken.")

