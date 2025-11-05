import re
import streamlit as st
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import cloudinary
import cloudinary.uploader

st.set_page_config(page_title="DOCX → Stermonitor / LessonUp HTML")

st.title("DOCX → HTML converter")

platform = st.selectbox("Kies platform", ["Stermonitor", "LessonUp"])
uploaded = st.file_uploader("Upload Word-bestand (.docx)", type=["docx"])

# ----- Cloudinary config -----
required_keys = ["CLOUDINARY_CLOUD_NAME", "CLOUDINARY_API_KEY", "CLOUDINARY_API_SECRET"]
missing = [k for k in required_keys if k not in st.secrets]

if missing:
    st.info("Je kunt zonder Cloudinary testen, maar afbeeldingen worden dan niet geüpload.")
else:
    cloudinary.config(
        cloud_name=st.secrets["CLOUDINARY_CLOUD_NAME"],
        api_key=st.secrets["CLOUDINARY_API_KEY"],
        api_secret=st.secrets["CLOUDINARY_API_SECRET"],
        secure=True,
    )


def extract_images(doc: Document):
    """geeft lijst van (filename, bytes) uit docx"""
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


def upload_image(filename: str, data: bytes):
    """upload naar cloudinary, geef url terug; als geen cloudinary → None"""
    if missing:
        return None
    try:
        res = cloudinary.uploader.upload(
            data,
            public_id=filename.split(".")[0],
            folder="ster_monitor",
            resource_type="image",
        )
        return res["secure_url"]
    except Exception as e:
        st.error(f"Upload mislukt voor {filename}: {e}")
        return None


def is_word_list_paragraph(p):
    name = (p.style.name or "").lower()
    if "list" in name or "lijst" in name or "opsom" in name:
        return True
    ppr = p._p.pPr
    return ppr is not None and ppr.numPr is not None


def runs_to_html(p):
    out = []
    for r in p.runs:
        txt = r.text.strip()
        if not txt:
            continue
        if r.bold:
            out.append(f"<strong>{txt}</strong>")
        else:
            out.append(txt)
    return " ".join(out)


def docx_to_html(file, platform="Stermonitor"):
    doc = Document(file)
    html_parts = []
    buffer = ""
    in_list = False
    first_bold_seen = False

    # afbeeldingen alvast
    images = extract_images(doc)
    image_urls = [upload_image(fn, b) for fn, b in images]
    img_i = 0

    for p in doc.paragraphs:
        text = (p.text or "").strip()

        # kopjes
        if p.style.name.startswith("Heading"):
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            try:
                level = int(p.style.name.split()[-1])
            except ValueError:
                level = 2
            html_parts.append(f"<h{level}>{text}</h{level}>")
            continue

        # afbeelding
        has_image = any("graphic" in r._element.xml for r in p.runs)
        if has_image:
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            url = image_urls[img_i] if img_i < len(image_urls) else None
            img_i += 1

            if url:
                if platform == "Stermonitor":
                    html_parts.append(
                        f'<p><img src="{url}" alt="afbeelding" '
                        'style="width:300px;height:300px;object-fit:cover;'
                        'border:1px solid #ccc;border-radius:8px;padding:4px;"></p>'
                    )
                else:
                    html_parts.append(f'<p><img src="{url}" alt="afbeelding"></p>')
            else:
                html_parts.append("<p>[afbeelding niet geüpload]</p>")
            continue

        # opsomming
        if is_word_list_paragraph(p):
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            if not in_list:
                html_parts.append(
                    '<ul class="browser-default">' if platform == "Stermonitor" else "<ul>"
                )
                in_list = True
            html_parts.append(f"<li>{runs_to_html(p)}</li>")
            continue

        # vette regel
        bold_runs = [r for r in p.runs if r.bold]
        if bold_runs and not is_word_list_paragraph(p):
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            if in_list:
                html_parts.append("</ul>")
                in_list = False

            bold_html = runs_to_html(p)
            if platform == "Stermonitor" and first_bold_seen:
                html_parts.append("<br>")
            html_parts.append(f"<p>{bold_html}</p>")
            first_bold_seen = True
            continue

        # gewone tekst
        if text:
            if in_list:
                html_parts.append("</ul>")
                in_list = False
            buffer += " " + runs_to_html(p)
            if re.search(r"[.!?]$", text):
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""

    if buffer:
        html_parts.append(f"<p>{buffer.strip()}</p>")
    if in_list:
        html_parts.append("</ul>")

    return "\n".join(html_parts)


if uploaded:
    html = docx_to_html(uploaded, platform)
    st.subheader(f"HTML voor {platform}")
    st.code(html, language="html")
    st.download_button(
        "Download HTML",
        data=html,
        file_name="output.html",
        mime="text/html",
    )
else:
    st.info("Upload een .docx om HTML te genereren.")



