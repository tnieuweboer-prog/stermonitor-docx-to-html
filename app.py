import os
import streamlit as st
from html_converter import docx_to_html
from pptx_converter_hybrid import docx_to_pptx_hybrid
from workbook_builder import build_workbook_docx_front_and_steps

st.set_page_config(page_title="Triade DOCX Tools", page_icon="📘", layout="wide")

# ---------- CSS: alleen een smalle topbar ----------
st.markdown(
    """
    <style>
    .stApp {
        background: #e5f7e5;
    }
    .triade-topbar {
        width: 100%;
        background: #c9edc9;
        border-bottom: 1px solid #b2e3b2;
        padding: 1rem 1.2rem;
        display: flex;
        align-items: center;
        gap: 1rem;
    }
    .triade-logo {
        height: 46px;
        object-fit: contain;
    }
    .triade-search input {
        width: 100%;
        padding: 0.45rem 1rem;
        border-radius: 9999px;
        border: 1px solid #d8efdb;
        background: white;
        outline: none;
    }
    .triade-btn {
        background: #0fa14b;
        color: #fff;
        padding: 0.35rem 0.85rem;
        border-radius: 0.5rem;
        border: none;
        font-weight: 600;
        font-size: 0.82rem;
    }
    .triade-nav {
        display: flex;
        gap: 0.7rem;
        color: #0b4c2c;
        font-weight: 600;
        font-size: 0.8rem;
    }
    /* tabs dichter tegen topbar */
    .block-container {
        padding-top: 0.4rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- TOPBAR ----------
logo_path = os.path.join("assets", "logo-triade-460px.png")
top1, top2, top3, top4 = st.columns([1.2, 3, 1.6, 1.6])
with top1:
    if os.path.exists(logo_path):
        st.image(logo_path, width=120)
    else:
        st.markdown("### De Triade")
with top2:
    st.text_input("Zoeken", placeholder="Zoeken")
with top3:
    st.markdown(
        '<div style="display:flex;gap:0.4rem;justify-content:flex-end;">'
        '<button class="triade-btn">Triade dagen</button>'
        '<button class="triade-btn">Schoolfoto\'s</button>'
        '</div>',
        unsafe_allow_html=True,
    )
with top4:
    st.markdown(
        '<div class="triade-nav" style="justify-content:flex-end;">'
        '<span>Over De Triade</span>'
        '<span>Onderwijs</span>'
        '<span>Praktisch</span>'
        '</div>',
        unsafe_allow_html=True,
    )

# ⚠️ géén hero-blok meer hier!


# ---------- TABS ----------
tab1, tab2, tab3 = st.tabs(
    ["💚 HTML (Stermonitor/ Elodigitaal)", "🤖 PowerPoint", "📘 Werkboekjes-generator"]
)

# ---------------- TAB 1 ----------------
with tab1:
    st.subheader("DOCX → HTML Converter")
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


# ---------------- TAB 2 ----------------
with tab2:
    st.subheader("DOCX → PowerPoint (AI-hybride)")
    uploaded_ai = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"], key="hybrid_upload")

    if uploaded_ai:
        if st.button("📽️ Maak PowerPoint", type="primary"):
            with st.spinner("PowerPoint wordt opgebouwd met AI..."):
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
        st.info("Upload een .docx-bestand om een AI-dia te genereren.")


# ---------------- TAB 3 ----------------
with tab3:
    st.subheader("📘 Werkboekjes-generator")
    st.caption("Voorblad → (optioneel) materiaalstaat → daarna pagina’s die je zelf kiest.")

    # 1. Voorblad
    col1, col2 = st.columns(2)
    with col1:
        wb_opdracht_titel = st.text_input("Opdracht titel")
        wb_vak = st.text_input("Vak (bijv. BWI)", value="BWI")
        wb_profieldeel = st.text_input("Keuze/profieldeel")
    with col2:
        wb_docent = st.text_input("Docent")
        wb_duur = st.text_input("Duur van de opdracht", value="11 x 45 minuten")
        wb_cover = st.file_uploader("📸 Omslagfoto (optioneel)", type=["png", "jpg", "jpeg"])

    st.markdown("---")

    # 2. Materiaalstaat
    if "num_material_rows" not in st.session_state:
        st.session_state.num_material_rows = 1

    def add_material_row():
        st.session_state.num_material_rows += 1

    include_materiaalstaat = st.checkbox("Materiaalstaat toevoegen aan werkboekje")

    materialen = []
    if include_materiaalstaat:
        st.markdown("#### Materiaalstaat invullen")
        st.caption("Vul hieronder de materialen in.")
        headers = ["Nummer", "Aantal", "Benaming", "Lengte", "Breedte", "Dikte", "Materiaal"]
        header_cols = st.columns([1, 1, 2, 1, 1, 1, 1])
        for i, h in enumerate(headers):
            header_cols[i].markdown(f"**{h}**")

        for row_idx in range(st.session_state.num_material_rows):
            cols = st.columns([1, 1, 2, 1, 1, 1, 1])
            values = []
            for col_idx, h in enumerate(headers):
                values.append(
                    cols[col_idx].text_input(
                        label="", key=f"mat_{h}_{row_idx}", placeholder=h
                    )
                )
            materialen.append(dict(zip(headers, values)))
        st.button("➕ Voeg materiaal toe", on_click=add_material_row)

    st.markdown("---")

    # 3. Pagina's
    st.markdown("### Pagina's")

    if "wb_pages" not in st.session_state:
        st.session_state.wb_pages = []

    layout_options = [
        "Werktekening (1 grote afbeelding)",
        "1 stap: korte tekst + grote afbeelding",
        "2 stappen: tekst + afbeelding (past op 1 pagina)",
        "3 stappen: tekst + afbeelding (past op 1 pagina)",
    ]

    pages_data = []
    for idx, page in enumerate(st.session_state.wb_pages):
        st.markdown(f"#### Pagina {idx + 1}")
        layout = st.selectbox(
            "Kies layout",
            layout_options,
            index=layout_options.index(page.get("layout", layout_options[0])),
            key=f"layout_{idx}",
        )
        page_data = {"layout": layout}

        if layout == "Werktekening (1 grote afbeelding)":
            img = st.file_uploader(
                f"Afbeelding voor pagina {idx+1}", type=["png", "jpg", "jpeg"], key=f"page_img_{idx}_0"
            )
            page_data["images"] = [img.read()] if img else []
            page_data["steps"] = []
        elif layout == "1 stap: korte tekst + grote afbeelding":
            title = st.text_input(f"Titel voor pagina {idx+1}", key=f"page_title_{idx}_0")
            text = st.text_area(f"Tekst (max 4 regels)", key=f"page_text_{idx}_0", height=80)
            img = st.file_uploader(
                f"Afbeelding voor pagina {idx+1}", type=["png", "jpg", "jpeg"], key=f"page_img_{idx}_0"
            )
            page_data["steps"] = [{"title": title, "text": text}]
            page_data["images"] = [img.read()] if img else []
        elif layout == "2 stappen: tekst + afbeelding (past op 1 pagina)":
            steps_list, images_list = [], []
            for s in range(2):
                title = st.text_input(f"Titel stap {s+1} (pagina {idx+1})", key=f"page_title_{idx}_{s}")
                text = st.text_area(f"Tekst stap {s+1}", key=f"page_text_{idx}_{s}", height=80)
                img = st.file_uploader(
                    f"Afbeelding stap {s+1}", type=["png", "jpg", "jpeg"], key=f"page_img_{idx}_{s}"
                )
                steps_list.append({"title": title, "text": text})
                images_list.append(img.read() if img else None)
            page_data["steps"] = steps_list
            page_data["images"] = images_list
        elif layout == "3 stappen: tekst + afbeelding (past op 1 pagina)":
            steps_list, images_list = [], []
            for s in range(3):
                title = st.text_input(f"Titel stap {s+1} (pagina {idx+1})", key=f"page_title_{idx}_{s}")
                text = st.text_area(f"Tekst stap {s+1}", key=f"page_text_{idx}_{s}", height=80)
                img = st.file_uploader(
                    f"Afbeelding stap {s+1}", type=["png", "jpg", "jpeg"], key=f"page_img_{idx}_{s}"
                )
                steps_list.append({"title": title, "text": text})
                images_list.append(img.read() if img else None)
            page_data["steps"] = steps_list
            page_data["images"] = images_list

        pages_data.append(page_data)
        st.markdown("---")

    # knop onderaan
    if st.button("➕ Nieuwe pagina"):
        st.session_state.wb_pages.append({"layout": "Werktekening (1 grote afbeelding)"})

    st.markdown("---")

    # genereren
    if st.button("📘 Werkboekje genereren"):
        meta = {
            "opdracht_titel": wb_opdracht_titel,
            "vak": wb_vak,
            "profieldeel": wb_profieldeel,
            "docent": wb_docent,
            "duur": wb_duur,
            "include_materiaalstaat": include_materiaalstaat,
            "materialen": materialen,
        }

        # logo uit assets
        logo_path = os.path.join("assets", "logo-triade-460px.png")
        if os.path.exists(logo_path):
            with open(logo_path, "rb") as f:
                meta["logo"] = f.read()

        if wb_cover is not None:
            meta["cover_bytes"] = wb_cover.read()

        steps = []
        for page in pages_data:
            layout = page["layout"]
            if layout == "Werktekening (1 grote afbeelding)":
                img_bytes = page["images"][0] if page["images"] else None
                steps.append({"title": "Werktekening", "text_blocks": [], "images": [img_bytes] if img_bytes else []})
            elif layout == "1 stap: korte tekst + grote afbeelding":
                stp = page["steps"][0] if page["steps"] else {"title": "", "text": ""}
                img_bytes = page["images"][0] if page["images"] else None
                steps.append({
                    "title": stp.get("title", ""),
                    "text_blocks": [stp.get("text", "")] if stp.get("text") else [],
                    "images": [img_bytes] if img_bytes else [],
                })
            elif layout == "2 stappen: tekst + afbeelding (past op 1 pagina)":
                for i, stp in enumerate(page["steps"]):
                    img_bytes = page["images"][i] if i < len(page["images"]) else None
                    steps.append({
                        "title": stp.get("title", ""),
                        "text_blocks": [stp.get("text", "")] if stp.get("text") else [],
                        "images": [img_bytes] if img_bytes else [],
                    })
            elif layout == "3 stappen: tekst + afbeelding (past op 1 pagina)":
                for i, stp in enumerate(page["steps"]):
                    img_bytes = page["images"][i] if i < len(page["images"]) else None
                    steps.append({
                        "title": stp.get("title", ""),
                        "text_blocks": [stp.get("text", "")] if stp.get("text") else [],
                        "images": [img_bytes] if img_bytes else [],
                    })

        with st.spinner("Werkboekje wordt gemaakt..."):
            try:
                docx_bytes = build_workbook_docx_front_and_steps(meta, steps)
            except Exception as e:
                st.error(f"❌ Kon werkboekje niet maken: {e}")
            else:
                st.success("✅ Werkboekje klaar!")
                st.download_button(
                    "⬇️ Download werkboekje (Word)",
                    data=docx_bytes,
                    file_name="werkboekje.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

