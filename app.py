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
    "Tekst wordt omgezet naar HTML, afbeeldingen worden naar Cloudinary geüpload "
    "en opsommingen worden automatisch herkend."
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


def is_list_item(text: str) -> bool:
    """Eenvoudige herkenning van opsommingen."""
    if not text:
        return False
    stripped = text.lstrip()
    return (
        stripped.startswith("- ")
        or stripped.startswith("* ")
        or stripped.startswith("• ")
    )


def docx_to_html(file):
    """Zet tekst, afbeeldingen en lijstjes uit docx om naar HTML."""
    doc = Document(file)
    html_parts = []
    buffer = ""          # voor gewone alinea's die we per zin willen bundelen
    in_list = False      # zijn we nu in een <ul> ... </ul>
    images = extract_images(doc)
    image_urls = [upload_to_cloudinary(f, b) for f, b in images]
    img_counter = 0

    for para in doc.paragraphs:
        text = para.text.strip()

        # 1. HEADINGS
        if para.style.name.startswith("Heading"):
            # openstaande buffer eerst wegschrijven
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            # openstaande lijst eerst sluiten
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            # heading genereren
            try:
                level = int(para.style.name.split()[-1])
            except ValueError:
                level = 2
            html_parts.append(f"<h{level}>{text}</h{level}>")
            continue

        # 2. AFBEELDING in deze paragraaf?
        has_image = any("graphic" in run._element.xml for run in para.runs)
        if has_image:
            # eerst buffer wegschrijven
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            # openstaande lijst sluiten
            if in_list:
                html_parts.append("</ul>")
                in_list = False

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

        # 3. LIJST-ITEM?
        if is_list_item(text):
            # buffer eerst wegschrijven
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            # als we nog niet in een lijst zaten, begin er een
            if not in_list:
                html_parts.append("<ul>")
                in_list = True
            # tekst zonder het opsommingsteken
            stripped = text.lstrip()
            if stripped[0] in "-*•":
                item_text = stripped[2:]
            else:
                item_text = stripped
            html_parts.append(f"<li>{item_text}</li>")
            continue

        # 4. GEWONE TEKST
        if text:
            # als we net in een lijst zaten en nu gewone tekst krijgen: lijst sluiten
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            buffer += " " + text
            # als tekst eindigt op punt/vraagteken/uitroepteken, schrijf de alinea weg
            if re.search(r"[.!?]$", text):
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""

    # na de loop: openstaande buffer wegschrijven
    if buffer:
        html_parts.append(f"<p>{buffer.strip()}</p>")
    # openstaande lijst sluiten
    if in_list:
        html_parts.append("</ul>")

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

