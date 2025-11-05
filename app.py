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
    "Tekst wordt omgezet naar HTML, afbeeldingen worden naar Cloudinary geüpload, "
    "Word-opsommingen worden herkend en omgezet naar "
    '<ul class="browser-default">...</ul>.'
)

uploaded = st.file_uploader("Kies een Word-bestand", type=["docx"])

# ───────── Cloudinary config check ─────────
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
# ───────────────────────────────────────────


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


def is_word_list_paragraph(para):
    """
    Herken een Word-opsomming.
    We kijken naar:
    - style name (List Paragraph / Lijstparagraaf / Opsomming)
    - numPr in pPr (echte docx-numbering)
    """
    style_name = (para.style.name or "").lower()
    if (
        "list" in style_name
        or "lijst" in style_name
        or "opsom" in style_name  # vangt 'Opsommingstekens' op
    ):
        return True

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

        # 1. HEADINGS
        if para.style.name.startswith("Heading"):
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            try:
                level = int(para.style.name.split()[-1])
            except ValueError:
                level = 2
            html_parts.append(f"<h{level}>{text}</h{level}>")
            continue

        # 2. AFBEELDING
        has_image = any("graphic" in run._element.xml for run in para.runs)
        if has_image:
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
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

        # 3. OPSOMMING (Word bullet / nummering)
        if is_word_list_paragraph(para):
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            if not in_list:
                # ← hier jouw vereiste class
                html_parts.append('<ul class="browser-default">')
                in_list = True

            html_parts.append(f"<li>{text}</li>")
            continue

        # 4. GEWONE TEKST
        if text:
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            buffer += " " + text
            if re.search(r"[.!?]$", text):
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""

    # einde document
    if buffer:
        html_parts.append(f"<p>{buffer.strip()}</p>")
    if in_list:
        html_parts.append("</ul>")

    return "\n".join(html_parts)


# ───────── UI ─────────
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
