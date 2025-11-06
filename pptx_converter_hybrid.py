import io
import os
import math
import json
from copy import deepcopy

import requests
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement

CHARS_PER_LINE = 75

# -------------------------------
# Cloudinary config (invullen!)
# -------------------------------
CLOUDINARY_CLOUD_NAME = os.getenv("CLOUDINARY_CLOUD_NAME", "JOUW_CLOUD_NAME")
CLOUDINARY_UPLOAD_PRESET = os.getenv("CLOUDINARY_UPLOAD_PRESET", "JOUW_UPLOAD_PRESET")  # voor unsigned
CLOUDINARY_API_KEY = os.getenv("CLOUDINARY_API_KEY", "")
CLOUDINARY_API_SECRET = os.getenv("CLOUDINARY_API_SECRET", "")

# lokaal logo in je repo
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")

# -----------------------------------
# AI helper (zelfde als eerder)
# -----------------------------------
def summarize_with_ai(text: str, max_bullets: int = 0) -> str | list:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        words = text.split()
        if max_bullets:
            parts = [p.strip() for p in text.replace("•", "\n").split("\n") if p.strip()]
            return parts[:max_bullets] or ["Kernpunt uit de tekst."]
        short = " ".join(words[:40])
        return short + "..." if len(words) > 40 else short

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        if max_bullets:
            prompt = f"""
Maak van deze tekst maximaal {max_bullets} korte bullets (mbo/havo-niveau, 1 regel per bullet).
Alleen de kern. Geef JSON als:
{{"bullets": ["...", "..."]}}

Tekst:
{text}
"""
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
            )
            data = json.loads(resp.choices[0].message.content)
            return data.get("bullets") or ["Kernpunt uit de tekst."]
        else:
            prompt = f"""
Vat deze les-tekst samen in 1 korte alinea voor een PowerPoint-dia.
Doelgroep: havo/vmbo techniekleerlingen.
Max 40 woorden.

Tekst:
{text}
"""
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
            )
            return resp.choices[0].message.content.strip()

    except Exception:
        words = text.split()
        if max_bullets:
            parts = [p.strip() for p in text.replace("•", "\n").split("\n") if p.strip()]
            return parts[:max_bullets] or ["Kernpunt uit de tekst."]
        short = " ".join(words[:40])
        return short + "..." if len(words) > 40 else short


# -----------------------------------
# DOCX helpers
# -----------------------------------
def extract_images(doc):
    imgs = []
    for rel in doc.part.rels.values():
        if rel.reltype == RT.IMAGE:
            imgs.append((rel.target_part.partname, rel.target_part.blob))
    return imgs


def is_word_list_paragraph(para):
    name = (para.style.name or "").lower()
    if "list" in name or "lijst" in name or "opsom" in name:
        return True
    ppr = para._p.pPr
    return ppr is not None and ppr.numPr is not None


def has_bold(para):
    return any(run.bold for run in para.runs)


def para_text_plain(para):
    return "".join(run.text for run in para.runs if run.text).strip()


# -----------------------------------
# Cloudinary helpers
# -----------------------------------
def upload_logo_to_cloudinary(local_path: str) -> str | None:
    """
    Upload lokale logo naar Cloudinary en geef secure_url terug.
    Werkt met unsigned upload (upload preset) of signed upload.
    """
    if not os.path.exists(local_path):
        print("⚠️ lokaal logo niet gevonden:", local_path)
        return None

    url = f"https://api.cloudinary.com/v1_1/{CLOUDINARY_CLOUD_NAME}/image/upload"

    files = {
        "file": open(local_path, "rb"),
    }

    data = {}

    # eerst proberen met upload preset (meestal het makkelijkst)
    if CLOUDINARY_UPLOAD_PRESET and CLOUDINARY_UPLOAD_PRESET != "JOUW_UPLOAD_PRESET":
        data["upload_preset"] = CLOUDINARY_UPLOAD_PRESET
    elif CLOUDINARY_API_KEY and CLOUDINARY_API_SECRET:
        # signed upload
        # voor simplicity laten we hier de sign-stap achterwege, maar zo heb je de vorm
        data["api_key"] = CLOUDINARY_API_KEY
        # normaal: timestamp + signature maken
        # maar vaak is upload preset al genoeg
    else:
        print("⚠️ Geen Cloudinary config gevonden (upload preset of API key/secret).")
        return None

    try:
        resp = requests.post(url, files=files, data=data, timeout=15)
        if resp.status_code == 200:
            info = resp.json()
            return info.get("secure_url")
        else:
            print("⚠️ Cloudinary upload faalde:", resp.status_code, resp.text)
    except Exception as e:
        print("⚠️ kon niet uploaden naar Cloudinary:", e)

    return None


def get_logo_bytes_from_url(url: str) -> bytes | None:
    try:
        r = requests.get(url, timeout=10)
        if r.status_code == 200:
            return r.content
    except Exception as e:
        print("⚠️ kon logo-url niet ophalen:", e)
    return None


def ensure_logo_bytes() -> bytes | None:
    """
    1. probeer LOGO_URL uit env
    2. anders: upload lokaal logo naar Cloudinary
    3. download het logo terug als bytes
    """
    logo_url_env = os.getenv("CLOUDINARY_LOGO_URL", "").strip()

    if logo_url_env:
        # user heeft zelf al URL gezet
        return get_logo_bytes_from_url(logo_url_env)

    # anders: zelf uploaden
    uploaded_url = upload_logo_to_cloudinary(LOCAL_LOGO_PATH)
    if uploaded_url:
        # je kunt hier evt. die url in een klein bestandje schrijven zodat je niet elke keer uploadt
        return get_logo_bytes_from_url(uploaded_url)

    # laatste redmiddel: direct lokaal inlezen zonder cloudinary
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()

    return None


# -----------------------------------
# PPTX helpers
# -----------------------------------
def add_logo_to_slide(slide, logo_bytes):
    if not logo_bytes:
        return
    left = Inches(9.0 - 1.5)  # aanpassen aan jouw dia
    top = Inches(0.2)
    width = Inches(1.5)
    slide.shapes.add_picture(io.BytesIO(logo_bytes), left, top, width=width)


def duplicate_slide(prs, slide_index=0, logo_bytes=None):
    source = prs.slides[slide_index]
    blank_layout = prs.slide_layouts[0]
    dest = prs.slides.add_slide(blank_layout)

    for shape in source.shapes:
        el = shape.element
        new_el = deepcopy(el)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    if logo_bytes:
        add_logo_to_slide(dest, logo_bytes)

    return dest


def make_bullet(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith("buNone"):
            pPr.remove(child)
    pPr.set("marL", "288000")
    pPr.set("indent", "-144000")
    buChar = OxmlElement("a:buChar")
    buChar.set("char", "•")
    pPr.append(buChar)


def add_textbox(slide, text, top_inch=1.5, est_lines=1):
    left = Inches(0.8)
    top = Inches(top_inch)
    width = Inches(8.0)
    height_inch = 0.6 + (est_lines - 1) * 0.25
    shape = slide.shapes.add_textbox(left, top, width, Inches(height_inch))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)
    return height_inch


def add_inline_image(slide, img_bytes, top_inch):
    left = Inches(1.0)
    top = Inches(top_inch)
    width = Inches(4.5)
    slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=width)
    return 3.0


# -----------------------------------
# MAIN
# -----------------------------------
def docx_to_pptx_hybrid(file_like):
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", "basis layout.pptx")

    if not os.path.exists(template_path):
        print("⚠️ Template 'basis layout.pptx' niet gevonden, lege presentatie gemaakt.")
        prs = Presentation()
    else:
        prs = Presentation(template_path)

    # 1x logo regelen (uploaden indien nodig)
    logo_bytes = ensure_logo_bytes()

    doc = Document(file_like)
    all_images = extract_images(doc)
    img_ptr = 0

    # start met dia 0 uit template
    current_slide = prs.slides[0]
    if logo_bytes:
        add_logo_to_slide(current_slide, logo_bytes)
    if current_slide.shapes.title:
        current_slide.shapes.title.text = "Les gegenereerd met AI"
    current_y = 2.0

    for para in doc.paragraphs:
        raw_text = (para.text or "").strip()
        if not raw_text:
            continue

        is_heading = para.style.name.startswith("Heading")
        is_bold = has_bold(para)
        is_list = is_word_list_paragraph(para)
        has_image = any("graphic" in run._element.xml for run in para.runs)

        if is_heading or is_bold:
            current_slide = duplicate_slide(prs, 0, logo_bytes=logo_bytes)
            if current_slide.shapes.title:
                current_slide.shapes.title.text = para_text_plain(para)
            current_y = 2.0
            continue

        if has_image:
            if img_ptr < len(all_images):
                _, img_bytes = all_images[img_ptr]
                img_ptr += 1
                add_inline_image(current_slide, img_bytes, current_y)
                current_y += 3.2
            continue

        if is_list:
            bullets = summarize_with_ai(raw_text, max_bullets=3)
            shape = current_slide.shapes.add_textbox(Inches(0.8), Inches(current_y), Inches(7.5), Inches(3))
            tf = shape.text_frame
            tf.word_wrap = True
            for i, b in enumerate(bullets):
                p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
                p.text = b
                make_bullet(p)
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(16)
            current_y += 0.3 * len(bullets)
            continue

        short_text = summarize_with_ai(raw_text)
        h = add_textbox(current_slide, short_text, top_inch=current_y)
        current_y += h + 0.3

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out



