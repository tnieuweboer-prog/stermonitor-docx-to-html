import streamlit as st
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import re
import base64

st.set_page_config(page_title="Stermonitor HTML Converter")

st.title("Stermonitor HTML Converter")
st.write(
    "Upload een Word (.docx) bestand. Tekst wordt omgezet naar eenvoudige HTML, afbeeldingen worden los aangeboden."
)

uploaded = st.file_uploader("Kies een Word-bestand", type=["docx"])


def extract_images(doc):
    """
    Haalt alle afbeeldingen uit het docx-document.
    Geeft een lijst van tuples terug: (bestandsnaam, bytes)
    """
    images = []
    idx = 1
    for rel in doc.part.rels.values():
        if rel.reltype == RT.IMAGE:
            image = rel.target_part.blob
            # probeer een extensie te raden
            ext = rel.target_part.partname.ext
            filename = f"image_{idx}.{ext}"
            images.append((filename, image))
            idx += 1
    return images


def docx_to_html(file):
    doc = Document(file)
    html_parts = []
    buffer = ""  # tijdelijke tekstbuffer
    image_placeholders = []  # we plaatsen hier img-tags op volgorde

    # afbeeldingen alvast ophalen
    images = extract_images(doc)
    img_counter = 0

    for para in doc.paragraphs:
        text = para.text.strip()

        # Als paragraaf leeg is, gewoon overslaan
        if not text and not para.runs:
            continue

        # Koppen blijven aparte blokken
        if para.style.name.startswith("Heading"):
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            try:
                level = int(para.style.name.split()[-1])
            except ValueError:
                level = 2
            html_parts.append(f"<h{level}>{text}</h{level}>")
            continue

        # Check of er in deze paragraaf een afbeelding-run zit
        has_image = any("graphic" in run._element.xml for run in para.runs)

        if has_image:
            # eerst evt. buffer tekst wegschrijven
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""

            # voeg een img-tag toe met placeholder-naam
            img_counter += 1
            img_name = f"IMAGE_{img_counter}"
            html_parts.append(f'<p><img src="{img_name}" alt="afbeelding {img_counter}"></p>')
            continue

        # anders: gewone tekst → buffer
        if text:
            buffer += " " + text
            # als de regel eindigt op punt, vraagteken of uitroepteken → schrijf blok weg
            if re.search(r"[.!?]$", text):
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""

    # restbuffer wegschrijven
    if buffer:
        html_parts.append(f"<p>{buffer.strip()}</p>")

    return "\n".join(html_parts), images


if uploaded:
    html_output, images = docx_to_html(uploaded)

    st.subheader("Gegenereerde HTML-code")
    st.code(html_output, language="html")

    st.download_button(
        label="Download HTML-bestand",
        data=html_output,
        file_name="ster_monitor.html",
        mime="text/html",
    )

    if images:
        st.subheader("Afbeeldingen uit het Word-bestand")
        st.write("Deze kun je apart uploaden naar je ELO / Stermonitor en de src in de HTML vervangen door de echte URL.")
        for filename, img_bytes in images:
            st.image(img_bytes, caption=filename)
            st.download_button(
                label=f"Download {filename}",
                data=img_bytes,
                file_name=filename,
                mime="image/png"  # vaak goed; Word kan ook jpg hebben maar png is veilig
            )
    else:
        st.info("Geen afbeeldingen gevonden in dit document.")
else:
    st.info("Upload hierboven een .docx-bestand om te beginnen.")


