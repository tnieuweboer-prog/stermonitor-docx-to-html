# app.py
import streamlit as st
import streamlit as st
from pptx_converter_ai import docx_to_pptx_ai
from html_converter import docx_to_html
from pptx_converter import docx_to_pptx
# plus eventuele Cloudinary-setup hier of in html_converter via callback

st.title("Triade DOCX Converter")
tab1, tab2 = st.tabs(["HTML", "PowerPoint"])

uploaded = st.file_uploader("Upload docx voor AI-dia's", type=["docx"])
if uploaded:
    pptx_bytes = docx_to_pptx_ai(uploaded)
    st.download_button(
        "Download AI PowerPoint",
        data=pptx_bytes,
        file_name="les_ai.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


with tab1:
    uploaded = st.file_uploader("Upload voor HTML", type=["docx"])
    if uploaded:
        html = docx_to_html(uploaded, image_uploader=None, platform="Stermonitor")
        st.code(html, language="html")

with tab2:
    uploaded2 = st.file_uploader("Upload voor PPTX", type=["docx"])
    if uploaded2:
        pptx_bytes = docx_to_pptx(uploaded2)
        st.download_button("Download PPTX", data=pptx_bytes, file_name="lessonup.pptx")
