import os
import streamlit as st
from html_converter import docx_to_html
from pptx_converter_hybrid import docx_to_pptx_hybrid
from workbook_builder import build_workbook_docx_front_and_steps

st.set_page_config(page_title="Triade DOCX Tools", page_icon="📘", layout="wide")
st.title("📘 Triade DOCX Tools")

tab1, tab2, tab3 = st.tabs([
    "💚 HTML (Stermonitor / LessonUp)",
    "🤖 PowerPoint (AI-hybride)",
    "📘 Werkboekjes-generator"
])

# =========================================================
# TAB 1
# =========================================================
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


# =========================================================
# TAB 2
# =========================================================
with tab2:
    st.subheader("DOCX → PowerPoint (AI-Hybride)")
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


# =========================================================
# TAB 3
# =========================================================
with tab3:
    st.subheader("📘 Werkboekje-generator")
    st.caption("Voorblad → (optioneel) materiaalstaat → daarna zelf pagina’s toevoegen met vaste layouts.")

    # ----------------- 1. VOORPAGINA -----------------
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

    # ----------------- 2. MATERIAALSTAAT -----------------
    if "num_material_rows" not in st.session_state:
        st.session_state.num_material_rows = 1

    def add_material_row():
        st.session_state.num_material_rows += 1

    include_materiaalstaat = st.checkbox("Materiaalstaat toevoegen aan werkboekje")

    materialen = []
    if include_materiaalstaat:
        st.markdown("#### Materiaalstaat invullen")
        st.caption("Vul de materialen in. Klik op ➕ voor een extra rij.")
        headers = ["Nummer", "Aantal", "Benaming", "Lengte", "Breedte", "Dikte", "Materiaal"]

        for row_idx in range(st.session_state.num_material_rows):
            cols = st.columns([1, 1, 2, 1, 1, 1, 1])
            values = []
            for col_idx, h in enumerate(headers):
                values.append(
                    cols[col_idx].text_input(
                        label="",
                        key=f"mat_{h}_{row_idx}",
                        placeholder=h,
                    )
                )
            materialen.append(dict(zip(headers, values)))

        st.button("➕ Voeg materiaal toe", on_click=add_material_row)

    st.markdown("---")

    # ----------------- 3. PAGINA'S / LAYOUTS -----------------
    st.markdown("### Pagina's toevoegen")

    # lijst van pagina's in session state
    if "wb_pages" not in st.session_state:
        st.session_state.wb_pages = []

    # knop om pagina toe te voegen
    if st.button("➕ Nieuwe pagina"):
        st.session_state.wb_pages.append({
            "layout": "Werktekening (1 grote afbeelding)",
        })

    layout_options = [
        "Werktekening (1 grote afbeelding)",
        "1 stap: korte tekst + grote afbeelding",
        "2 stappen: tekst + afbeelding (past op 1 pagina)",
        "3 stappen: tekst + afbeelding (past op 1 pagina)",
    ]

    # UI voor elke pagina
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

        # layout 1: werktekening
        if layout == "Werktekening (1 grote afbeelding)":
            img = st.file_uploader(f"Afbeelding (grote werktekening) voor pagina {idx+1}", type=["png", "jpg", "jpeg"], key=f"page_img_{idx}_0")
            page_data["images"] = [img.read()] if img else []
            page_data["steps"] = []  # geen tekst

        # layout 2: 1 stap + grote afbeelding
        elif layout == "1 stap: korte tekst + grote afbeelding":
            title = st.text_input(f"Titel stap (max 4 regels tekst) voor pagina {idx+1}", key=f"page_title_{idx}_0")
            text = st.text_area(f"Tekst (max 4 regels)", key=f"page_text_{idx}_0", height=80)
            img = st.file_uploader(f"Afbeelding groot voor pagina {idx+1}", type=["png", "jpg", "jpeg"], key=f"page_img_{idx}_0")
            page_data["steps"] = [{"title": title, "text": text}]
            page_data["images"] = [img.read()] if img else []

        # layout 3: 2 stappen
        elif layout == "2 stappen: tekst + afbeelding (past op 1 pagina)":
            steps_list = []
            images_list = []
            for s in range(2):
                title = st.text_input(f"Titel stap {s+1} (max 4 regels) voor pagina {idx+1}", key=f"page_title_{idx}_{s}")
                text = st.text_area(f"Tekst stap {s+1}", key=f"page_text_{idx}_{s}", height=80)
                img = st.file_uploader(f"Afbeelding stap {s+1}", type=["png", "jpg", "jpeg"], key=f"page_img_{idx}_{s}")
                steps_list.append({"title": title, "text": text})
                images_list.append(img.read() if img else None)
            page_data["steps"] = steps_list
            page_data["images"] = images_list

        # layout 4: 3 stappen
        elif layout == "3 stappen: tekst + afbeelding (past op 1 pagina)":
            steps_list = []
            images_list = []
            for s in range(3):
                title = st.text_input(f"Titel stap {s+1} voor pagina {idx+1}", key=f"page_title_{idx}_{s}")
                text = st.text_area(f"Tekst stap {s+1}", key=f"page_text_{idx}_{s}", height=80)
                img = st.file_uploader(f"Afbeelding stap {s+1}", type=["png", "jpg", "jpeg"], key=f"page_img_{idx}_{s}")
                steps_list.append({"title": title, "text": text})
                images_list.append(img.read() if img else None)
            page_data["steps"] = steps_list
            page_data["images"] = images_list

        pages_data.append(page_data)
        st.markdown("---")

    # ----------------- 4. GENEREREN -----------------
    if st.button("📘 Werkboekje genereren"):
        # meta voor voorblad + materiaalstaat
        meta = {
            "opdracht_titel": wb_opdracht_titel,
            "vak": wb_vak,
            "profieldeel": wb_profieldeel,
            "docent": wb_docent,
            "duur": wb_duur,
            "include_materiaalstaat": include_materiaalstaat,
            "materialen": materialen,
        }

        # logo automatisch uit assets
        logo_path = os.path.join("assets", "logo-triade-460px.png")
        if os.path.exists(logo_path):
            with open(logo_path, "rb") as f:
                meta["logo"] = f.read()

        if wb_cover is not None:
            meta["cover_bytes"] = wb_cover.read()

        # vertaal pages_data naar steps voor builder
        steps = []
        for page in pages_data:
            layout = page["layout"]

            if layout == "Werktekening (1 grote afbeelding)":
                # één stap zonder tekst, met afbeelding
                img_bytes = page["images"][0] if page["images"] else None
                steps.append({
                    "title": "Werktekening",
                    "text_blocks": [],
                    "images": [img_bytes] if img_bytes else [],
                })

            elif layout == "1 stap: korte tekst + grote afbeelding":
                stp = page["steps"][0] if page["steps"] else {"title": "", "text": ""}
                img_bytes = page["images"][0] if page["images"] else None
                steps.append({
                    "title": stp.get("title", ""),
                    "text_blocks": [stp.get("text", "")] if stp.get("text") else [],
                    "images": [img_bytes] if img_bytes else [],
                })

            elif layout == "2 stappen: tekst + afbeelding (past op 1 pagina)":
                for idx_s, stp in enumerate(page["steps"]):
                    img_bytes = page["images"][idx_s] if idx_s < len(page["images"]) else None
                    steps.append({
                        "title": stp.get("title", ""),
                        "text_blocks": [stp.get("text", "")] if stp.get("text") else [],
                        "images": [img_bytes] if img_bytes else [],
                    })

            elif layout == "3 stappen: tekst + afbeelding (past op 1 pagina)":
                for idx_s, stp in enumerate(page["steps"]):
                    img_bytes = page["images"][idx_s] if idx_s < len(page["images"]) else None
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


