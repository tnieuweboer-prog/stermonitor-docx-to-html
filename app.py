import streamlit as st

from html_converter import docx_to_html
from pptx_converter import docx_to_pptx
from pptx_converter_ai import docx_to_pptx_ai

st.set_page_config(page_title="Triade DOCX tools", page_icon="📘", layout="centered")

st.title("📘 Triade DOCX tools")

tab1, tab2, tab3 = st.tabs(["💚 HTML (Stermonitor)", "💜 PowerPoint (klassiek)", "🤖 PowerPoint (AI)"])

# ---------- TAB 1: HTML ----------
with tab1:
    st.subheader("DOCX → HTML")
    platform = st.selectbox("Platform", ["Stermonitor", "LessonUp"], key="platform_html")
    uploaded_html = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="html_upload")

    if uploaded_html:
        # cloudinary config ophalen uit secrets, mag ook None zijn
        cloud_name = st.secrets.get("CLOUDINARY_CLOUD_NAME")
        api_key = st.secrets.get("CLOUDINARY_API_KEY")
        api_secret = st.secrets.get("CLOUDINARY_API_SECRET")

        html_out = docx_to_html(
            uploaded_html,
            platform=platform,
            cloud_name=cloud_name,
            api_key=api_key,
            api_secret=api_secret,
        )
        st.code(html_out, language="html")
        st.download_button("⬇️ Download HTML", data=html_out, file_name="output.html", mime="text/html")
    else:
        st.info("Upload eerst een .docx-bestand.")


# ---------- TAB 2: PPTX klassieke converter ----------
with tab2:
    st.subheader("DOCX → PowerPoint (klassiek)")
    uploaded_pptx = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="pptx_upload")

    if uploaded_pptx:
        pptx_bytes = docx_to_pptx(uploaded_pptx)
        st.download_button(
            "⬇️ Download PowerPoint",
            data=pptx_bytes,
            file_name="les_klassiek.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    else:
        st.info("Upload een .docx voor een PowerPoint.")


# ---------- TAB 3: AI PPTX ----------
with tab3:
    st.subheader("DOCX → PowerPoint (AI samenvatting per dia)")
    st.caption("Let op: gebruikt OPENAI_API_KEY uit secrets.")
    uploaded_ai = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="pptx_ai_upload")

    if uploaded_ai:
        pptx_ai_bytes = docx_to_pptx_ai(uploaded_ai)
        st.download_button(
            "⬇️ Download AI PowerPoint",
            data=pptx_ai_bytes,
            file_name="les_ai.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    else:
        st.info("Upload een .docx om een AI-dia te maken.")

with tab2:
    uploaded2 = st.file_uploader("Upload voor PPTX", type=["docx"])
    if uploaded2:
        pptx_bytes = docx_to_pptx(uploaded2)
        st.download_button("Download PPTX", data=pptx_bytes, file_name="lessonup.pptx")
