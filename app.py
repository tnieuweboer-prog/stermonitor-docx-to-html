import streamlit as st
from html_converter import docx_to_html
from pptx_converter_hybrid import docx_to_pptx_hybrid

st.set_page_config(page_title="Triade DOCX Tools", page_icon="📘", layout="wide")

st.title("📘 Triade DOCX → HTML / PowerPoint")

tab1, tab2 = st.tabs(["💚 HTML (Stermonitor / LessonUp)", "🤖 PowerPoint (AI-hybride)"])


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


# ---------------- TAB 2: AI-Hybride PowerPoint ----------------
with tab2:
    st.subheader("DOCX → PowerPoint (AI-Hybride)")
    st.caption(
        "Gebruik de vertrouwde layout uit je klassieke converter, maar laat AI de tekst inkorten en samenvatten."
    )

    uploaded_ai = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="hybrid_upload")

    if uploaded_ai:
        with st.spinner("PowerPoint wordt opgebouwd met AI..."):
            pptx_bytes = docx_to_pptx_hybrid(uploaded_ai)
        st.success("✅ Klaar! PowerPoint gegenereerd.")
        st.download_button(
            "⬇️ Download PowerPoint (AI-hybride)",
            data=pptx_bytes,
            file_name="les_ai_hybride.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    else:
        st.info("Upload een .docx-bestand om een AI-dia te genereren.")

try:
    pptx_bytes = docx_to_pptx_hybrid(uploaded_file)
    # laten downloaden / tonen
except Exception as e:
    st.error(f"AI kon geen les maken: {e}")

