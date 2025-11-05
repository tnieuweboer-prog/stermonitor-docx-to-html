import streamlit as st
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import cloudinary
import cloudinary.uploader
import re

st.set_page_config(page_title="Stermonitor HTML Converter")

st.title("Stermonitor HTML Converter (met Cloudinary)")
st.write(
    "Upload een Word (.docx) bestand. Tekst wordt omgezet naar HTML "
    "en afbeeldingen worden automatisch naar Cloudinary geüpload "
    "met een vaste grootte van 300×300 px."
)

uploaded = st.file_uploader("Kies een Word-bestand", type=["docx"])

# ▸ Controleer of Cloudinary-secrets aanwezig zijn
required_keys = ["CLOUDINARY_CLOUD_NAME", "CLOUDINARY_API_KEY", "CLOUDINARY_API_SECRET"]
missing = [k for k in required_keys if k not in st.secrets]
if missing:
    st.warning(
        "Cloudinary is nog niet goed ingesteld. "
        "Vul in Streamlit → Edit secrets deze waarden in:\n"
        "CLOUDINARY_CLOUD_NAME, CLOUDINARY_API_KEY, CLOUDINARY_API_SECRET"
    )
else:
    # ▸ Cloudinary configureren
    cloudinary.config(
        cloud_name=st.secrets["CLOUDINARY_CLOUD_NAME"],
        api_key=st.secrets["CLOUDINARY_API_KEY"],
        api_secret=st.secrets["CLOUDINARY_API_SECRET"],
        secure=True
    )


def extract_images(doc):
    """Haal alle afbeeldingen uit het Word-document."""
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
    if missing:
        return None
    try:
        result = cloudinary.uploader.upload(
            data,
            public_id=filename.split('.')[0],
            folder="ster_monitor",
            resource_type="image"
        )
        return result["secure_url"]
    except Exception as e:
        st.error(f"Cloudinary-upload mislukt: {e}")
        return None


def docx_to_html(file):
    """Zet tekst en afbeeldingen uit docx om naar HTML met vaste afbeeldingsstijl."""
    doc = Document(file)
    html_parts = []
    buffer = ""

    images = extract_images(doc)
    image_urls = [upload_to_cloudinary(f, b) for f, b in images]
    img_counter = 0

    for para in doc.paragraphs:
        text = para.text.strip()

        # Koppen (Kop 1, 2, 3…)
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

        # Afbeelding in paragraaf
        has_image = any("graphic" in run._element.xml for run in para.runs)
        if has_image:
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            if img_counter < len(image_urls) and image_urls[img_counter]:
                img_url = image_urls[img_counter]
                html_parts.append(
                    f'<p><img src="{img_url}" alt="afbeelding {img_counter+1}" '
                    'style="width:300px;height:300px;object-fit:cover;'
                    'border:1px solid #ccc;border-radius:8px;padding:4px;"></p>'
                )
            else:
                html_parts.append("<p>[afbeelding kon niet worden geüpload]</p>")
            img_counter += 1
            continue

        # Gewone tekst → voeg samen tot einde van zin
        if text:
            buffer += " " + text
            if re.search(r"[.!?]$", text):
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""

    if buffer:
        html_parts.append(f"<p>{buffer.strip()}</p>")

    return "\n".join(html_parts)


# ▸ Streamlit-interface
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

