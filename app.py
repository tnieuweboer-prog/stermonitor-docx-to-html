import streamlit as st
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import cloudinary
import cloudinary.uploader
import re

st.set_page_config(page_title="Stermonitor HTML Converter")

st.title("Stermonitor HTML Converter (met Cloudinary)")
st.write(
    "Upload een Word (.docx) bestand. Tekst wordt omgezet naar HTML, "
    "afbeeldingen worden automatisch naar Cloudinary geüpload."
)

uploaded = st.file_uploader("Kies een Word-bestand", type=["docx"])

# 🔧 Cloudinary configureren met jouw gegevens
cloudinary.config(
    cloud_name=st.secrets["CLOUDINARY_CLOUD_NAME"],
    api_key=st.secrets["CLOUDINARY_API_KEY"],
    api_secret=st.secrets["CLOUDINARY_API_SECRET"],
    secure=True
)

def extract_images(doc):
    images = []
    idx = 1
    for rel in doc.part.rels.values():
        if rel.reltype == RT.IMAGE:
            image = rel.target_part.blob
            ext = rel.target_part.partname.ext
            filename = f"image_{idx}.{ext}"
            images.append((filename, image))
            idx += 1
    return images

def upload_to_cloudinary(filename, data):
    """Upload afbeelding naar Cloudinary en geef publieke URL terug."""
    result = cloudinary.uploader.upload(
        data,
        public_id=filename.split('.')[0],
        folder="ster_monitor",  # map in je Cloudinary-account
        resource_type="image"
    )
    return result["secure_url"]

def docx_to_html(file):
    doc = Document(file)
    html_parts = []
    buffer = ""
    images = extract_images(doc)
    image_urls = []

    # Eerst alle afbeeldingen uploaden
    for filename, img_bytes in images:
        url = upload_to_cloudinary(filename, img_bytes)
        image_urls.append(url)

    img_counter = 0

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text and not para.runs:
            continue

        # Koppen
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

        # Afbeelding
        has_image = any("graphic" in run._element.xml for run in para.runs)
        if has_image:
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            if img_counter < len(image_urls):
                img_url = image_urls[img_counter]
                html_parts.append(f'<p><img src="{img_url}" alt="afbeelding {img_counter+1}"></p>')
                img_counter += 1
            continue

        # Tekst samenvoegen per zin
        if text:
            buffer += " " + text
            if re.search(r"[.!?]$", text):
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""

    if buffer:
        html_parts.append(f"<p>{buffer.strip()}</p>")

    return "\n".join(html_parts)


if uploaded:
    html_output = docx_to_html(uploaded)
    st.subheader("Gegenereerde HTML-code")
    st.code(html_output, language="html")
    st.download_button(
        label="Download HTML",
        data=html_output,
        file_name="ster_monitor.html",
        mime="text/html"
    )
else:
    st.info("Upload hierboven een .docx-bestand om te beginnen.")
