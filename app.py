import streamlit as st
from html_converter import docx_to_html
from pptx_converter_hybrid import docx_to_pptx_hybrid

st.title("📘 Triade DOCX Tools")

tab1, tab2 = st.tabs(["💚 HTML (Stermonitor)", "🤖 PowerPoint (AI-hybride)"])

# --- HTML converter ---
with tab1:
    st.subheader("DOCX → HTML")
    uploaded_html = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="html_upload")

    if uploaded_html:
        html_out = docx_to_html(uploaded_html, platform="Stermonitor")
        st.code(html_out, language="html")
        st.download_button("⬇️ Download HTML", data=html_out, file_name="les.html", mime="text/html")
    else:
        st.info("Upload een .docx-bestand om te converteren naar HTML.")

# --- Hybride AI PowerPoint converter ---
with tab2:
    st.subheader("DOCX → PowerPoint (AI-hybride)")
    st.caption("Gebruikt het klassieke design, maar kort de tekst automatisch in.")
    uploaded_ai = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="hybrid_upload")

    if uploaded_ai:
        pptx_bytes = docx_to_pptx_hybrid(uploaded_ai)
        st.download_button("⬇️ Download PowerPoint (AI)", data=pptx_bytes,
                           file_name="les_ai_hybride.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    else:
        st.info("Upload een .docx-bestand om een AI-dia te genereren.")
