import streamlit as st
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import cloudinary
import cloudinary.uploader
import re
import io
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="DOCX → HTML & PPTX Converter")

st.title("DOCX → Stermonitor / LessonUp én PowerPoint")

# keuze voor HTML-uitvoer
platform = st.selectbox("Kies HTML-platform", ["Stermonitor", "LessonUp"])

# twee aparte uploads
uploaded_html = st.file_uploader(
    "Upload Word voor HTML (Stermonitor/LessonUp)", type=["docx"], key="html_uploader"
)
uploaded_pptx = st.file_uploader(
    "Upload Word voor PowerPoint (LessonUp-import)", type=["docx"], key="pptx_uploader"
)

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
    """Herken opsommingen uit Word."""
    style_name = (para.style.name or "").lower()
    if "list" in style_name or "lijst" in style_name or "opsom" in style_name:
        return True
    p = para._p
    ppr = p.pPr
    return ppr is not None and ppr.numPr is not None


def runs_to_html(para):
    """Zet runs om naar HTML, inclusief vetgedrukt."""
    parts = []
    for run in para.runs:
        text = run.text.strip()
        if not text:
            continue
        if run.bold:
            parts.append(f"<strong>{text}</strong>")
        else:
            parts.append(text)
    return " ".join(parts)


def docx_to_html(file, platform="Stermonitor"):
    """Zet tekst, afbeeldingen, lijstjes en vetgedrukt om naar HTML passend bij platform."""
    doc = Document(file)
    html_parts = []
    buffer = ""
    in_list = False
    first_bold_seen = False

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
                if platform == "Stermonitor":
                    html_parts.append(
                        f'<p><img src="{img_url}" alt="afbeelding {img_idx+1}" '
                        'style="width:300px;height:300px;object-fit:cover;'
                        'border:1px solid #ccc;border-radius:8px;padding:4px;"></p>'
                    )
                else:
                    html_parts.append(
                        f'<p><img src="{img_url}" alt="afbeelding {img_idx+1}"></p>'
                    )
            else:
                html_parts.append("<p>[afbeelding kon niet worden geüpload]</p>")
            img_idx += 1
            continue

        # 3. OPSOMMING
        if is_word_list_paragraph(para):
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            if not in_list:
                if platform == "Stermonitor":
                    html_parts.append('<ul class="browser-default">')
                else:
                    html_parts.append("<ul>")
                in_list = True
            html_parts.append(f"<li>{runs_to_html(para)}</li>")
            continue

        # 4. VETTE REGEL
        bold_runs = [r for r in para.runs if r.bold]
        if bold_runs and not is_word_list_paragraph(para):
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            bold_html = runs_to_html(para)
            if platform == "Stermonitor":
                if first_bold_seen:
                    html_parts.append("<br>")
            html_parts.append(f"<p>{bold_html}</p>")
            first_bold_seen = True
            continue

        # 5. GEWONE TEKST
        if text:
            if in_list:
                html_parts.append("</ul>")
                in_list = False
            line_html = runs_to_html(para)
            buffer += " " + line_html
            # als de tekst met . ! ? eindigt, doe er een paragraaf van
            if re.search(r"[.!?]$", text):
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""

    # afsluiten
    if buffer:
        html_parts.append(f"<p>{buffer.strip()}</p>")
    if in_list:
        html_parts.append("</ul>")

    return "\n".join(html_parts)


# ------------- PPTX-deel -------------


def _get_text_frame_from_slide(slide, prs):
    """Probeer een text_frame op deze slide te vinden, anders nieuwe slide met body."""
    for shape in slide.shapes:
        if hasattr(shape, "text_frame"):
            return shape.text_frame, slide

    # geen tekstvak → maak nieuwe slide met layout 'Title and Content'
    new_slide = prs.slides.add_slide(prs.slide_layouts[1])
    return new_slide.shapes.placeholders[1].text_frame, new_slide


def docx_to_pptx(doc_bytes):
    """Zet een docx grofweg om naar een PowerPoint (voor LessonUp-import)."""
    prs = Presentation()
    doc = Document(doc_bytes)

    # titel-slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Inhoud uit Word"
    if len(slide.placeholders) > 1:
        slide.placeholders[1].text = "Geconverteerd voor LessonUp"

    current_slide = prs.slides[-1]

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        if not text:
            continue

        # heading → nieuwe slide
        if para.style.name.startswith("Heading"):
            current_slide = prs.slides.add_slide(prs.slide_layouts[1])
            current_slide.shapes.title.text = text
            tf, current_slide = _get_text_frame_from_slide(current_slide, prs)
            tf.text = ""
            continue

        # afbeelding → aparte slide met titel
        has_image = any("graphic" in run._element.xml for run in para.runs)
        if has_image:
            img_slide = prs.slides.add_slide(prs.slide_layouts[5])  # title only
            img_slide.shapes.title.text = "Afbeelding"
            current_slide = img_slide
            continue

        # opsomming → bullet
        if is_word_list_paragraph(para):
            tf, current_slide = _get_text_frame_from_slide(current_slide, prs)
            p = tf.add_paragraph()
            p.text = text
            p.level = 0
            continue

        # gewone tekst → bullet
        tf, current_slide = _get_text_frame_from_slide(current_slide, prs)
        if tf.text == "":
            tf.text = text
        else:
            p = tf.add_paragraph()
            p.text = text
            p.level = 0

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio


# ───────── UI ─────────

# 1. HTML-uitvoer
if uploaded_html:
    html_output = docx_to_html(uploaded_html, platform=platform)
    st.subheader(f"HTML voor {platform}")
    st.code(html_output, language="html")
    st.download_button(
        label=f"Download HTML ({platform})",
        data=html_output,
        file_name=f"{platform.lower()}_html.html",
        mime="text/html",
    )

# 2. PowerPoint-uitvoer
if uploaded_pptx:
    pptx_bytes = docx_to_pptx(uploaded_pptx)
    st.subheader("PowerPoint voor LessonUp")
    st.download_button(
        label="Download PowerPoint (.pptx)",
        data=pptx_bytes,
        file_name="lessonup_import.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

if not uploaded_html and not uploaded_pptx:
    st.info("Upload hierboven een Word-bestand voor HTML en/of voor PowerPoint.")
