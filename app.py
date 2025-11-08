import os
import io
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
    st.caption("Gebruik de vertrouwde layout uit je klassieke converter, maar laat AI de tekst omzetten naar VMBO-lesvorm.")

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


# ---------------- TAB 3: Werkboekjes-generator ----------------
with tab3:
    st.subheader("📘 Werkboekje-generator (met stappenplan en materiaalstaat)")
    st.caption("Maak eenvoudig een werkboekje in Word met Triade-stijl, logo, omslagfoto en optionele materiaalstaat.")

    with st.form("workbook_form"):
        col1, col2 = st.columns(2)
        with col1:
            wb_opdracht_titel = st.text_input("Opdracht titel")
            wb_vak = st.text_input("Vak (bijv. BWI)", value="BWI")
            wb_profieldeel = st.text_input("Keuze/profieldeel")
        with col2:
            wb_docent = st.text_input("Docent")
            wb_duur = st.text_input("Duur van de opdracht (bijv. 3 weken)")

        wb_cover = st.file_uploader("📸 Voeg omslagfoto toe (optioneel)", type=["png", "jpg", "jpeg"])

        st.markdown("---")
        st.markdown("### 🧱 Materiaalstaat (optioneel)")
        include_materiaalstaat = st.checkbox("Materiaalstaat toevoegen aan werkboekje")

        materialen = []
        if include_materiaalstaat:
            st.caption("Voeg materialen toe. Klik op ➕ om extra regels te maken.")
            num_items = st.number_input("Aantal materialen", min_value=1, max_value=30, value=3, step=1)
            for i in range(num_items):
                st.markdown(f"**Materiaal {i + 1}**")
                cols = st.columns(7)
                headers = ["Nummer", "Aantal", "Benaming", "Lengte", "Breedte", "Dikte", "Materiaal"]
                data = []
                for j, h in enumerate(headers):
                    data.append(cols[j].text_input(h, key=f"{h}_{i}", value=""))
                materialen.append(dict(zip(headers, data)))

        st.markdown("---")
        st.markdown("### ➕ Stappen toevoegen")
        st.caption("Voeg één of meer stappen toe voor het stappenplan.")

        num_steps = st.number_input("Aantal stappen", min_value=1, max_value=20, value=3, step=1)

        steps = []
        for i in range(num_steps):
            st.markdown(f"#### Stap {i + 1}")
            title = st.text_input(f"Titel stap {i + 1}", key=f"title_{i}")
            text = st.text_area(f"Tekst stap {i + 1}", key=f"text_{i}")
            img = st.file_uploader(f"Afbeelding voor stap {i + 1} (optioneel)", type=["png", "jpg", "jpeg"], key=f"img_{i}")

            step_data = {"title": title, "text_blocks": [text] if text else []}

            if img:
                step_data["images"] = [img.read()]
            else:
                step_data["images"] = []

            steps.append(step_data)

        generate_btn = st.form_submit_button("📘 Werkboekje genereren")

    if generate_btn:
        meta = {
            "opdracht_titel": wb_opdracht_titel,
            "vak": wb_vak,
            "profieldeel": wb_profieldeel,
            "docent": wb_docent,
            "duur": wb_duur,
            "include_materiaalstaat": include_materiaalstaat,
            "materialen": materialen,
        }

        logo_path = os.path.join("assets", "logo-triade-460px.png")
        if os.path.exists(logo_path):
            with open(logo_path, "rb") as f:
                meta["logo"] = f.read()

        if wb_cover is not None:
            meta["cover_bytes"] = wb_cover.read()

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

