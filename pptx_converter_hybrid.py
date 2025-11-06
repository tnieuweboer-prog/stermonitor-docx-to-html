import io
import os
import json
from copy import deepcopy

import requests
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.enum.shapes import MSO_SHAPE_TYPE

# -------------------------------------------------
# instellingen
# -------------------------------------------------
BASE_TEMPLATE_NAME = "basis layout.pptx"   # jouw template in /templates
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")

CHARS_PER_LINE = 75

# optioneel cloudinary
CLOUDINARY_CLOUD_NAME = os.getenv("CLOUDINARY_CLOUD_NAME", "")
CLOUDINARY_UPLOAD_PRESET = os.getenv("CLOUDINARY_UPLOAD_PRESET", "")
CLOUDINARY_LOGO_URL = os.getenv("CLOUDINARY_LOGO_URL", "")


# -------------------------------------------------
# AI helper (met fallback)
# -------------------------------------------------
def summarize_with_ai(text: str, max_bullets: int = 0) -> str | list:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        # simpele fallback
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


# -------------------------------------------------
# DOCX → blokken
# -------------------------------------------------
def has_bold(para):
    return any(run.bold for run in para.runs)


def para_text_plain(para):
    return "".join(run.text for run in para.runs if run.text).strip()


def docx_to_blocks(doc: Document):
    """
    Maakt blokken: (kop) + (alle tekst eronder) tot volgende kop.
    Kop = Heading of vet.
    """
    blocks = []
    current = None
    for para in doc.paragraphs:
        txt = (para.text or "").strip()
        if not txt:
            continue

        is_heading = para.style and para.style.name and para.style.name.startswith("Heading")
        is_bold = has_bold(para)

        if is_heading or is_bold:
            if current:
                blocks.append(current)
            current = {"title": para_text_plain(para), "body": []}
        else:
            if current is None:
                current = {"title": "Lesstof", "body": []}
            current["body"].append(txt)

    if current:
        blocks.append(current)

    return blocks


# -------------------------------------------------
# Cloudinary / logo helpers
# -------------------------------------------------
def upload_logo_to_cloudinary(local_path: str) -> str | None:
    if not CLOUDINARY_CLOUD_NAME or not os.path.exists(local_path):
        return None

    url = f"https://api.cloudinary.com/v1_1/{CLOUDINARY_CLOUD_NAME}/image/upload"
    files = {"file": open(local_path, "rb")}
    data = {}
    if CLOUDINARY_UPLOAD_PRESET:
        data["upload_preset"] = CLOUDINARY_UPLOAD_PRESET

    try:
        resp = requests.post(url, files=files, data=data, timeout=15)
        if resp.status_code == 200:
            return resp.json().get("secure_url")
    except Exception as e:
        print("⚠️ Cloudinary upload fout:", e)
    return None


def get_logo_bytes():
    # 1. als er al een url is
    if CLOUDINARY_LOGO_URL:
        try:
            r = requests.get(CLOUDINARY_LOGO_URL, timeout=10)
            if r.status_code == 200:
                return r.content
        except Exception:
            pass

    # 2. probeer te uploaden
    if os.path.exists(LOCAL_LOGO_PATH) and CLOUDINARY_CLOUD_NAME:
        up_url = upload_logo_to_cloudinary(LOCAL_LOGO_PATH)
        if up_url:
            try:
                r = requests.get(up_url, timeout=10)
                if r.status_code == 200:
                    return r.content
            except Exception:
                pass

    # 3. anders lokaal inlezen
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()

    # 4. geen logo
    return None


# -------------------------------------------------
# PPTX helpers
# -------------------------------------------------
def add_logo_to_slide(slide, logo_bytes):
    if not logo_bytes:
        return
    try:
        left = Inches(9.0 - 1.5)
        top = Inches(0.2)
        width = Inches(1.5)
        slide.shapes.add_picture(io.BytesIO(logo_bytes), left, top, width=width)
    except Exception as e:
        print("⚠️ logo niet toegevoegd:", e)


def get_or_add_title(slide, text: str):
    """
    Zet titel in bestaande title-placeholder, of maak er zelf eentje bovenaan.
    """
    if slide.shapes.title is not None:
        slide.shapes.title.text = text
        return

    # anders zelf een titelvak maken
    left = Inches(0.6)
    top = Inches(0.4)
    width = Inches(9.0)
    height = Inches(0.8)
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(28)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0, 0, 0)


def add_body_text(slide, text: str, top_inch: float = 2.0):
    left = Inches(0.8)
    top = Inches(top_inch)
    width = Inches(8.5)
    height = Inches(4.0)
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)


def duplicate_slide_no_external_pics(prs, slide_index=0, logo_bytes=None):
    source = prs.slides[slide_index]
    blank_layout = prs.slide_layouts[0]
    dest = prs.slides.add_slide(blank_layout)

    for shape in source.shapes:
        # gelinkte plaatjes uit template overslaan
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            continue
        new_el = deepcopy(shape.element)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    if logo_bytes:
        add_logo_to_slide(dest, logo_bytes)

    return dest


# -------------------------------------------------
# MAIN
# -------------------------------------------------
def docx_to_pptx_hybrid(file_like):
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)

    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        print("⚠️ template niet gevonden, lege ppt.")
        prs = Presentation()

    # docx inlezen en naar blokken
    doc = Document(file_like)
    blocks = docx_to_blocks(doc)

    # logo regelen (mag None zijn)
    logo_bytes = get_logo_bytes()

    # zorgen dat er iig 1 slide is
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])

    # eerste dia vullen
    first_slide = prs.slides[0]
    if logo_bytes:
        add_logo_to_slide(first_slide, logo_bytes)

    if blocks:
        get_or_add_title(first_slide, blocks[0]["title"])
        body_text = "\n".join(blocks[0]["body"]) if blocks[0]["body"] else ""
        if len(body_text) > 500:
            body_text = summarize_with_ai(body_text)
        add_body_text(first_slide, body_text, top_inch=2.0)
    else:
        get_or_add_title(first_slide, "Les gegenereerd met AI")

    # overige blokken → nieuwe dia’s
    for block in blocks[1:]:
        slide = duplicate_slide_no_external_pics(prs, 0, logo_bytes=logo_bytes)
        get_or_add_title(slide, block["title"])
        body_text = "\n".join(block["body"]) if block["body"] else ""
        if len(body_text) > 500:
            body_text = summarize_with_ai(body_text)
        add_body_text(slide, body_text, top_inch=2.0)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out



