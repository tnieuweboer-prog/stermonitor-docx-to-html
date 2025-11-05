import streamlit as st
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import cloudinary
import cloudinary.uploader
import re

st.set_page_config(page_title="Stermonitor HTML Converter")

st.title("Stermonitor HTML Converter (met Cloudinary)")
st.write(
    "Upload een Word (.docx) bestand. "
    "Tekst wordt omgezet naar HTML, afbeeldingen worden geüpload naar Cloudinary, "
    "opsommingen (• zoals in Word) worden herkend en omgezet naar <ul><li>."
)

uploaded = st.file_uploader("Kies een Word-bestand", type=["docx"])

# ─────────────────────────────
# Cloudinary configureren
# ─────────────────────────────
required_keys = ["CLOUDINARY_CLOUD_NAME", "CLOUDINARY_API_KEY", "CLOUDINARY_API_SECRET"]
missing = [k for k in required_keys if k not in st.secrets]
if missing:
    st.warning(
        "Cloudinary is nog niet goed ingesteld. Vul in Streamlit → Edit secrets:\n"
        "CLOUDINARY_CLOUD_NAME, CLOUDINARY_API_KEY, CLOUDINARY_API_SECRET"
    )
else:
    cloudinary.config(
        cloud_name=st.secrets["CLOUDINARY_CLOUD_NAME"],
        api_key=st.secrets["CLOUDINARY_API_KEY"],
        api_secret=st.secrets["CLOUDINARY_API_SECRET"],
        secure=True,
    )
# ─────────────────────────────


def extract_images(doc):
    """Haal alle afbeeldingen uit het Word-document."""
    images = []
    idx = 1
    for rel in doc.part.rels.values():
        if rel.reltype == RT.IMAGE:
            blob = rel.target_part.blob
            ext = rel.target_part.partname.ext
            filename = f"image_{idx}.{ext}"
            images.append((filename, blob))
            idx += 1
    return images


def upload_to_cloudinary(filename, data):
    """Upload afbeelding naar Cloudinary en geef publieke URL terug."""
    if missing:
        return None
    try:
        result = cloudinary.uploader.upload(
            data,
            public_id=filename.split(".")[0],
            folder="ster_monitor",
            resource_type="image",
        )
        return result["secure_url"]
    except Exception as e:
        st.error(f"Cloudinary-upload mislukt: {e}")
        return None


# ⭐ belangrijk: echte Word-lijsten herkennen
def is_word_list_paragraph(para):
    """
    Probeert te bepalen of dit een opsomming is.
    - veel NL/ENG Word versies geven style 'List Paragraph' of 'Lijstparagraaf'
    - docx bewaart nummering in pPr/numPr
    """
    style_name = (para.style.name or "").lower()
    if "list" in style_name or "lijst" in style_name:
        return True

    # check op numPr (echte docx-numbering)
    p = para._p
    ppr = p.pPr
    if ppr is not None and ppr.numPr is not None:
        return True

    return False


def docx_to_html(file):
    doc = Document(file)
    html_parts = []
    buffer = ""
    in_list = False

    # afbeeldingen alvast uploaden
    images = extract_images(doc)
    image_urls = [upload_to_cloudinary(f, b) for (f, b) in images]
    img_idx = 0

    for para in doc.paragraphs:
        text = (para.text or "").strip()

        # 1. koppen
        if para.style.name.startswith("Heading"):
            # eerst openstaande tekst wegschrijven
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            # openstaande lijst sluiten
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            try:
                level = int(para.style.name.split()[-1])
            except ValueError:
                level = 2
            html_parts.append(f"<h{level}>{text}</h{level}>")
            continue

        # 2. afbeelding?
        has_image = any("graphic" in run._element.xml for run in para.runs)
        if has_image:
            # eerst openstaande tekst wegschrijven
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            # openstaande lijst sluiten
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            if img_idx < len(image_urls) and image_urls[img_idx]:
                img_url = image_urls[img_idx]
                html_parts.append(
                    f'<p><img src="{img_url}" alt="afbeelding {img_idx+1}" '
                    'style="width:300px;height:300px;object-fit:cover;'
                    'border:1px solid #ccc;border-radius:8px;padding:4px;"></p>'
                )
            else:
                html_parts.append("<p>[afbeelding kon niet worden geüpload]</p>")
            img_idx += 1
            continue

        # 3. Word-lijst?
        if is_word_list_paragraph(para):
            # tekstbuffer eerst weg
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            # nieuwe lijst starten als we er nog niet in zitten
            if not in_list:
                html_parts.append("<ul>")
                in_list = True

            # in Word staat de bullet niet in de text, dus text is de echte inhoud
            html_parts.append(f"<li>{text}</li>")
            continue

        # 4. gewone tekst
        if text:
            # als we in een lijst zaten en er komt gewone tekst → lijst sluiten
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            buffer += " " + text
            # als de regel eindigt op . ! ? dan schrijven we 'm weg
            if re.search(r"[.!?]$", text):
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""

    # na alle paragrafen
    if buffer:
        html_parts.append(f"<p>{buffer.strip()}</p>")
    if in_list:
        html_parts.append("</ul>")

    return "\n".join(html_parts)


# ─────────────────────────────
# UI
# ─────────────────────────────
if uploaded:
    html_output = docx_to_html(uploaded)
    st.subheader("Gegenereerde HTML-code")
    st.code(html_output, language="html")
    st.download_button(
        label="Download HTML",
        data=html_output,
        file_name="ster_monitor.html",
        mime="text/html",
    )
else:
    st.info("Upload hierboven een .docx-bestand om te beginnen.")
