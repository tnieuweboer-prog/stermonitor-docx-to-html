import os
import sys
import io
import requests
from io import BytesIO
import streamlit as st
from docx import Document
import cloudinary
import cloudinary.uploader
import cloudinary.api

# ---------------- CONFIG / IMPORTS ----------------
sys.path.append(os.path.dirname(__file__))

from html_converter import docx_to_html
from pptx_converter_hybrid import docx_to_pptx_hybrid
from workbook_builder import build_workbook_docx_front_and_steps

# optionele AI-les
LESSON_ANALYZER_ERROR = None
try:
    from lesson_from_docx import docx_to_vmbo_lesson_json
    HAS_LESSON_ANALYZER = True
except Exception as e:
    HAS_LESSON_ANALYZER = False
    LESSON_ANALYZER_ERROR = str(e)

# cloudinary
cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
    api_key=os.getenv("CLOUDINARY_API_KEY"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET"),
    secure=True,
)

def upload_image_to_cloudinary(file_obj, folder="werkboekjes"):
    resp = cloudinary.uploader.upload(file_obj, folder=folder)
    return resp["secure_url"], resp["public_id"]

def delete_from_cloudinary(public_id):
    try:
        cloudinary.uploader.destroy(public_id)
    except Exception:
        pass

def download_image(url: str) -> bytes:
    r = requests.get(url)
    r.raise_for_status()
    return r.content


# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="Triade DOCX Tools", page_icon="📘", layout="wide")
st.title("📘 Triade DOCX Tools")

tab1, tab2, tab3 = st.tabs([
    "💚 HTML (Stermonitor / LessonUp)",
    "🤖 PowerPoint / Les-Word",
    "📘 Werkboekje generator",
])

# =========================================================
# TAB 1: HTML
# =========================================================
with tab1:
    st.subheader("DOCX → HTML Converter")
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


# =========================================================
# TAB 2: DOCX → PPTX / Les-Word
# =========================================================
with tab2:
    st.subheader("DOCX → PowerPoint (AI) / Les-Word")
    uploaded_ai = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="hybrid_upload")

    if not HAS_LESSON_ANALYZER:
        if LESSON_ANALYZER_ERROR:
            st.warning(f"⚠️ AI-lesmodule niet geladen: {LESSON_ANALYZER_ERROR}")
        else:
            st.warning("⚠️ AI-lesmodule niet geladen.")

    if uploaded_ai:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Les-Word (AI)**")
            if st.button("📝 Maak les-Word"):
                if not HAS_LESSON_ANALYZER:
                    st.error("Les-generatie niet beschikbaar.")
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

        with col2:
            st.markdown("**PowerPoint in vaste layout**")
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


# =========================================================
# TAB 3: Werkboekje
# =========================================================
with tab3:
    st.subheader("📘 Werkboekje generator")

    col_a, col_b = st.columns(2)
    with col_a:
        wb_vak = st.text_input("Vak (bijv. BWI)", value="BWI", key="wb_vak")
        wb_profieldeel = st.text_input("Profieldeel", key="wb_profieldeel")
        wb_opdracht_titel = st.text_input("Titel van opdracht", key="wb_opdracht_titel")
        wb_duur = st.text_input("Duur van opdracht", value="11 x 45 minuten", key="wb_duur")
    with col_b:
        wb_docent = st.text_input("Docent", key="wb_docent")
        wb_klas = st.text_input("Klas", key="wb_klas")
        wb_cover = st.file_uploader("Omslag-afbeelding (komt op voorblad)", type=["png", "jpg", "jpeg"], key="wb_cover")

    if "wb_steps" not in st.session_state:
        st.session_state.wb_steps = []

    st.markdown("### Stappen")
    if st.button("➕ Nieuwe stap"):
        st.session_state.wb_steps.append({"layout": "1 afbeelding + tekst"})

    for idx, _ in enumerate(st.session_state.wb_steps):
        st.markdown(f"#### Stap {idx + 1}")
        layout = st.selectbox(
            "Kies layout",
            ["1 afbeelding + tekst", "2 afbeeldingen + 2 teksten", "3 afbeeldingen + 3 teksten"],
            key=f"wb_layout_{idx}",
        )
        st.text_input("Titel van deze stap", key=f"wb_title_{idx}")
        max_blocks = int(layout[0])
        for i in range(max_blocks):
            st.markdown(f"**Blok {i+1}**")
            st.file_uploader(f"Afbeelding {i+1}", type=["png", "jpg", "jpeg"], key=f"wb_img_{idx}_{i}")
            st.text_area(f"Tekst {i+1}", key=f"wb_txt_{idx}_{i}")
        st.divider()

    if st.button("📄 Maak werkboekje (DOCX)"):
        # meta invullen
        meta = {
            "vak": wb_vak,
            "profieldeel": wb_profieldeel,
            "opdracht_nr": "1",
            "opdracht_titel": wb_opdracht_titel,
            "duur": wb_duur,
            "docent": wb_docent,
            "klas": wb_klas,
        }

        # logo
        logo_path = os.path.join("assets", "logo-triade-460px.png")
        if os.path.exists(logo_path):
            with open(logo_path, "rb") as f:
                meta["logo"] = f.read()

        # 🔴 hier gaat het om: upload direct naar bytes
        if wb_cover is not None:
            meta["cover_bytes"] = wb_cover.read()

        # stappen + afbeeldingen (via cloudinary)
        uploaded_public_ids = []
        steps = []
        for idx, _ in enumerate(st.session_state.wb_steps):
            layout = st.session_state.get(f"wb_layout_{idx}", "1 afbeelding + tekst")
            max_blocks = int(layout[0])
            text_blocks = []
            images = []
            title = st.session_state.get(f"wb_title_{idx}", "")

            for i in range(max_blocks):
                txt = st.session_state.get(f"wb_txt_{idx}_{i}", "")
                img_file = st.session_state.get(f"wb_img_{idx}_{i}")
                if txt:
                    text_blocks.append(txt)
                if img_file is not None:
                    url, pid = upload_image_to_cloudinary(img_file)
                    uploaded_public_ids.append(pid)
                    img_bytes = download_image(url)
                    images.append(img_bytes)

            steps.append({
                "title": title,
                "text_blocks": text_blocks,
                "images": images,
            })

        with st.spinner("Werkboekje wordt opgebouwd..."):
            try:
                docx_bytes = build_workbook_docx_front_and_steps(meta, steps)
            except Exception as e:
                st.error(f"❌ Fout bij maken werkboekje: {e}")
            else:
                for pid in uploaded_public_ids:
                    delete_from_cloudinary(pid)

                st.success("✅ Werkboekje gemaakt en Cloudinary opgeschoond.")
                st.download_button(
                    "⬇️ Download werkboekje.docx",
                    data=docx_bytes,
                    file_name="werkboekje.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

