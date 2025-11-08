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

# -----------------------------------------------------
# Lokale imports
# -----------------------------------------------------
sys.path.append(os.path.dirname(__file__))

from html_converter import docx_to_html
from pptx_converter_hybrid import docx_to_pptx_hybrid
from workbook_builder import build_workbook_docx_front_and_steps

# -----------------------------------------------------
# Les-analyse module (optioneel)
# -----------------------------------------------------
LESSON_ANALYZER_ERROR = None
try:
    from lesson_from_docx import docx_to_vmbo_lesson_json
    HAS_LESSON_ANALYZER = True
except Exception as e:
    HAS_LESSON_ANALYZER = False
    LESSON_ANALYZER_ERROR = str(e)

# -----------------------------------------------------
# Cloudinary config
# -----------------------------------------------------
cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
    api_key=os.getenv("CLOUDINARY_API_KEY"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET"),
    secure=True,
)

def upload_image_to_cloudinary(file_obj, folder="werkboekjes"):
    """Upload afbeelding en geef (url, public_id) terug."""
    resp = cloudinary.uploader.upload(file_obj, folder=folder)
    return resp["secure_url"], resp["public_id"]

def delete_from_cloudinary(public_id):
    """Verwijder afbeelding na gebruik."""
    try:
        cloudinary.uploader.destroy(public_id)
    except Exception:
        pass

def download_image(url: str) -> bytes:
    """Haal afbeelding op als bytes (voor docx)."""
    r = requests.get(url)
    r.raise_for_status()
    return r.content

# -----------------------------------------------------
# Streamlit setup
# -----------------------------------------------------
st.set_page_config(page_title="Triade DOCX Tools", page_icon="📘", layout="wide")
st.title("📘 Triade DOCX Tools")

tab1, tab2, tab3 = st.tabs([
    "💚 HTML (Stermonitor / LessonUp)",
    "🤖 PowerPoint / Les-Word",
    "📘 Werkboekje generator",
])

# =====================================================
# TAB 1: HTML CONVERTER
# =====================================================
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


# =====================================================
# TAB 2: DOCX → PowerPoint en Les-Word
# =====================================================
with tab2:
    st.subheader("DOCX → PowerPoint (AI) / Les-Word")
    st.caption("Maak een PowerPoint in jouw layout of een les-Word in VMBO-stijl.")

    uploaded_ai = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="hybrid_upload")

    if not HAS_LESSON_ANALYZER:
        msg = "⚠️ De AI-lesmodule (`lesson_from_docx.py`) kon niet worden geladen."
        if LESSON_ANALYZER_ERROR:
            msg += f"\n\n**Details:** {LESSON_ANALYZER_ERROR}"
        st.warning(msg)

    if uploaded_ai:
        col1, col2 = st.columns(2)

        # ---- Les-Word (AI)
        with col1:
            st.markdown("**Les-Word laten maken (AI)**")
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

        # ---- PowerPoint (AI)
        with col2:
            st.markdown("**PowerPoint maken in vaste layout**")
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


# =====================================================
# TAB 3: WERKBOEKJE GENERATOR
# =====================================================
with tab3:
    st.subheader("📘 Werkboekje generator")
    st.caption(
        "Maak een werkboekje met voorpagina, Triade-logo en daarna losse stappen. "
        "Elke stap kan 1, 2 of 3 afbeeldingen en tekstblokken hebben."
    )

    col_a, col_b = st.columns(2)
    with col_a:
        wb_vak = st.text_input("Vak (bijv. PIE / BWI)", key="wb_vak", value="BWI")
        wb_profieldeel = st.text_input("Profieldeel", key="wb_profieldeel")
        wb_opdracht_titel = st.text_input("Titel van opdracht", key="wb_opdracht_titel")
        wb_duur = st.text_input("Duur van opdracht", key="wb_duur", value="11 x 45 minuten")
    with col_b:
        wb_docent = st.text_input("Docent", key="wb_docent")
        wb_klas = st.text_input("Klas", key="wb_klas")
        wb_cover = st.file_uploader("Omslag-afbeelding (optioneel)", type=["png", "jpg", "jpeg"], key="wb_cover")

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
        title = st.text_input("Titel van deze stap", key=f"wb_title_{idx}")
        max_blocks = int(layout[0])
        for i in range(max_blocks):
            st.markdown(f"**Blok {i+1}**")
            st.file_uploader(f"Afbeelding {i+1}", type=["png", "jpg", "jpeg"], key=f"wb_img_{idx}_{i}")
            st.text_area(f"Tekst {i+1}", key=f"wb_txt_{idx}_{i}")
        st.divider()

    # -----------------------
    # Werkboekje genereren
    # -----------------------
    if st.button("📄 Maak werkboekje (DOCX)"):
        uploaded_public_ids = []

        # verzamel meta-info
        meta = {
            "vak": wb_vak,
            "profieldeel": wb_profieldeel,
            "opdracht_nr": "1",
            "opdracht_titel": wb_opdracht_titel,
            "duur": wb_duur,
            "docent": wb_docent,
            "klas": wb_klas,
        }

        # verzamel stappen
        steps = []
        for idx, _ in enumerate(st.session_state.wb_steps):
            layout = st.session_state.get(f"wb_layout_{idx}", "1 afbeelding + tekst")
            max_blocks = int(layout[0])
            text_blocks = []
            images = []
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
                "title": st.session_state.get(f"wb_title_{idx}", ""),
                "text_blocks": text_blocks,
                "images": images,
            })

        # --------------------
        # Triade-logo inladen
        # --------------------
        logo_path = os.path.join("assets", "logo-triade-460px.png")
        if not os.path.exists(logo_path):
            st.error("❌ Logo niet gevonden in assets-map.")
        else:
            with open(logo_path, "rb") as f:
                meta["logo"] = f.read()

        # maak DOCX via aparte builder
        with st.spinner("Werkboekje wordt opgebouwd..."):
            try:
                docx_bytes = build_workbook_docx_front_and_steps(meta, steps)
            except Exception as e:
                st.error(f"❌ Fout bij maken werkboekje: {e}")
            else:
                # alle cloudinary-bestanden verwijderen
                for pid in uploaded_public_ids:
                    delete_from_cloudinary(pid)

                st.success("✅ Werkboekje gemaakt en Cloudinary opgeschoond.")
                st.download_button(
                    "⬇️ Download werkboekje.docx",
                    data=docx_bytes,
                    file_name="werkboekje.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
