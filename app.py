import os
import sys
import streamlit as st

# zorgen dat we modules in dezelfde map kunnen vinden
sys.path.append(os.path.dirname(__file__))

from html_converter import docx_to_html
from pptx_converter_hybrid import docx_to_pptx_hybrid

# probeer de les-generator (die een DOCX teruggeeft) te laden
LESSON_ANALYZER_ERROR = None
try:
    # deze functie geeft nu een DOCX (BytesIO) terug
    from lesson_from_docx import docx_to_vmbo_lesson_json
    HAS_LESSON_ANALYZER = True
except Exception as e:
    HAS_LESSON_ANALYZER = False
    LESSON_ANALYZER_ERROR = str(e)


st.set_page_config(page_title="Triade DOCX Tools", page_icon="📘", layout="wide")
st.title("📘 Triade DOCX → HTML / PowerPoint / Les-Word")

tab1, tab2 = st.tabs(["💚 HTML (Stermonitor / LessonUp)", "🤖 PowerPoint + les-Word"])


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
# TAB 2: DOCX → PowerPoint + DOCX-les
# ==========================================================
with tab2:
    st.subheader("DOCX → PowerPoint (AI) en/of nieuw les-Word")
    st.caption("Upload je Word-bestand en kies of je een PowerPoint of een les-Word wilt.")

    uploaded_ai = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="hybrid_upload")

    if not HAS_LESSON_ANALYZER:
        msg = "⚠️ De AI-les-module (`lesson_from_docx.py`) kon niet worden geladen."
        if LESSON_ANALYZER_ERROR:
            msg += f"\n\n**Details:** {LESSON_ANALYZER_ERROR}"
        st.warning(msg)

    if uploaded_ai:
        col1, col2 = st.columns(2)

        # --------- knop 1: les-Word maken (dit is jouw optie 2) ----------
        with col1:
            st.markdown("**Les-Word maken (via OpenAI, 1 call)**")
            if st.button("📝 Maak VMBO-les als Word-bestand"):
                if not HAS_LESSON_ANALYZER:
                    st.error("Les-generatie is niet beschikbaar (lesson_from_docx.py mist of heeft een fout).")
                else:
                    with st.spinner("Les wordt herschreven naar VMBO-formaat..."):
                        try:
                            # deze functie geeft nu een DOCX (BytesIO) terug
                            lesson_docx = docx_to_vmbo_lesson_json(uploaded_ai)
                            st.success("✅ Les-Word gemaakt.")
                            st.download_button(
                                "⬇️ Download les_vmbo.docx",
                                data=lesson_docx,
                                file_name="les_vmbo.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            )
                        except Exception as e:
                            st.error(f"❌ Kon geen les-Word maken: {e}")

        # --------- knop 2: PowerPoint maken (zoals je al had) ----------
        with col2:
            st.markdown("**PowerPoint maken in jouw layout**")
            if st.button("📽️ Maak PowerPoint"):
                with st.spinner("PowerPoint wordt opgebouwd..."):
                    try:
                        pptx_bytes = docx_to_pptx_hybrid(uploaded_ai)
                        st.success("✅ Klaar! PowerPoint gegenereerd.")
                        st.download_button(
                            "⬇️ Download PowerPoint",
                            data=pptx_bytes,
                            file_name="les_ai_hybride.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )
                    except Exception as e:
                        st.error(f"❌ Kon geen PowerPoint maken: {e}")
    else:
        st.info("Upload een .docx-bestand om een les-Word of PowerPoint te maken.")

