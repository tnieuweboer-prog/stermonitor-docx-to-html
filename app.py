import streamlit as st
from html_converter import docx_to_html
from pptx_converter_hybrid import docx_to_pptx_hybrid
from lesson_from_docx import docx_to_vmbo_lesson_json  # ← dit is de “eerst-analyseren” stap

st.set_page_config(page_title="Triade DOCX Tools", page_icon="📘", layout="wide")

st.title("📘 Triade DOCX → HTML / PowerPoint")

tab1, tab2 = st.tabs(["💚 HTML (Stermonitor / LessonUp)", "🤖 PowerPoint (AI-vmbo)"])


# ---------------- TAB 1: HTML Converter ----------------
with tab1:
    st.subheader("DOCX → HTML Converter")
    st.caption("Zet je Word-lesstof automatisch om naar nette HTML voor Stermonitor of LessonUp.")

    uploaded_html = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="html_upload")

    if uploaded_html:
        with st.spinner("Word-bestand wordt omgezet..."):
            html_out = docx_to_html(uploaded_html)
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


# ---------------- TAB 2: AI → eerst lesstructuur, dan PPT ----------------
with tab2:
    st.subheader("DOCX → PowerPoint (AI-vmbo)")
    st.caption("Stap 1: AI analyseert je Word-bestand. Stap 2: we maken er een PowerPoint van met jouw template.")

    uploaded_ai = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="hybrid_upload")

    # we bewaren de JSON in session_state zodat je 'm kunt hergebruiken
    if "lesson_json" not in st.session_state:
        st.session_state.lesson_json = None

    if uploaded_ai:
        col_a, col_b = st.columns(2)
        with col_a:
            analyse = st.button("🧠 1. Analyseer met AI")
        with col_b:
            maak_ppt = st.button("📽️ 2. Maak PowerPoint")

        if analyse:
            with st.spinner("AI is de les aan het opdelen in dia’s..."):
                try:
                    lesson_json = docx_to_vmbo_lesson_json(uploaded_ai)
                except Exception as e:
                    st.error(f"❌ Kon lesstructuur niet maken: {e}")
                    st.session_state.lesson_json = None
                else:
                    st.session_state.lesson_json = lesson_json
                    st.success("✅ Lesstructuur gemaakt!")
                    st.json(lesson_json)  # laat zien wat AI heeft bedacht

        if maak_ppt:
            if not st.session_state.lesson_json:
                st.warning("Eerst analyseren met AI (stap 1), daarna pas PowerPoint maken.")
            else:
                with st.spinner("PowerPoint wordt opgebouwd..."):
                    try:
                        # hier gebruiken we je bestaande converter; je kunt 'm aanpassen
                        pptx_bytes = docx_to_pptx_hybrid(uploaded_ai)
                    except Exception as e:
                        st.error(f"❌ Kon geen PowerPoint maken: {e}")
                    else:
                        st.success("✅ Klaar! PowerPoint gegenereerd.")
                        st.download_button(
                            "⬇️ Download PowerPoint (AI-vmbo)",
                            data=pptx_bytes,
                            file_name="les_ai_hybride.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )
    else:
        st.info("Upload een .docx-bestand om een AI-les te maken.")

