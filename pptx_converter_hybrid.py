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

BASE_TEMPLATE_NAME = "basis layout.pptx"
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")

CLOUDINARY_CLOUD_NAME = os.getenv("CLOUDINARY_CLOUD_NAME", "")
CLOUDINARY_UPLOAD_PRESET = os.getenv("CLOUDINARY_UPLOAD_PRESET", "")
CLOUDINARY_LOGO_URL = os.getenv("CLOUDINARY_LOGO_URL", "")


# ---------- AI fallback ----------
def summarize_with_ai(text: str, max_bullets: int = 0) -> str | list:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        words = text.split()
        if max_bullets:
            parts = [p.strip() for p in text.replace("•", "\n").split("\n") if p.strip()]
            return parts[:max_bullets] or ["Kernpunt uit de tekst."]
        short = " ".join(words[:40])
        return short + "..." if len(words) > 40 else short
    # als je echte AI wilt gebruiken kun je dit stuk houden,
    # maar voor nu houden we de fallback simpel
    words = text.split()
    if max_bullets:
        parts = [p.strip() for p in text.replace("•", "\n").split("\n") if p.strip()]
        return parts[:max_bullets] or ["Kernpunt uit de tekst."]
    short = " ".join(words[:40])
    return short + "..." if len(words) > 40 else short


# ---------- DOCX helpers ----------
def has_bold(para):
    return any(run.bold for run in para.runs)


def is_all_caps_heading(text: str) -> bool:
    """
    Herken jouw stijl kopjes: volledig hoofdletters, kort, geen punt.
    """
    txt = text.strip()
    if not txt:
        return False
    # bv "SOORTEN KABELS", "VMVL KABEL"
    if len(txt) <= 35 and txt.upper() == txt and " " in txt:
        return True
    return False


def para_text_plain(para):
    return "".join(run.text for run in para.runs if run.text).strip()


def docx_to_blocks(doc: Document):
    """
    Maak blokken: (kop) + (tekst eronder) tot volgende kop.
    Kop = Heading, of vet, of ALL CAPS-kort.
    """
    blocks = []
    current = None

    for para in doc.paragraphs:
        txt = (para.text or "").strip()
        if not txt:
            continue

        is_heading_style = para.style and para.style.name and para.style.name.startswith("Heading")
        is_bold_para = has_bold(para)
        is_caps = is_all_caps_heading(txt)

        if is_heading_style or is_bold_para or is_caps:
            # nieuw blok
            if current:
                blocks.append(current)
            current = {"title": txt, "body": []}
        else:
            # hoort bij huidige blok
            if current is None:
                current = {"title": "Lesstof", "body": []}
            current["body"].append(txt)

    if current:
        blocks.append(current)

    return blocks


# ---------- logo helpers ----------
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
        print("⚠️ cloudinary upload fout:", e)
    return None


def get_logo_bytes():
    # 1. url uit env
    if CLOUDINARY_LOGO_URL:
        try:
            r = requests.get(CLOUDINARY_LOGO_URL, timeout=10)
            if r.status_code == 200:
                return r.content
        except Exception:
            pass

    # 2. upload lokaal logo
    if os.path.exists(LOCAL_LOGO_PATH) and CLOUDINARY_CLOUD_NAME:
        up_url = upload_logo_to_cloudinary(LOCAL_LOGO_PATH)
        if up_url:
            try:
                r = requests.get(up_url, timeout=10)
                if r.status_code == 200:
                    return r.content
            except Exception:
                pass

    # 3. lokaal inlezen
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()

    return None


# ---------- pptx helpers ----------
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
    # gebruik de bestaande titel als die er is
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


def clear_all_text_except_title(slide):
    """
    Op een gekloonde dia alle tekstvakjes leegmaken behalve de titel.
    Zo voorkom je dat oude template-tekst blijft staan.
    """
    title_shape = slide.shapes.title
    for shape in slide.shapes:
        # sla plaatjes en lijnen over
        if not hasattr(shape, "text_frame"):
            continue
        if title_shape is not None and shape == title_shape:
            continue
        # tekst leeg
        shape.text_frame.clear()


def duplicate_slide_no_external_pics(prs, slide_index=0, logo_bytes=None):
    source = prs.slides[slide_index]
    blank_layout = prs.slide_layouts[0]
    dest = prs.slides.add_slide(blank_layout)

    for shape in source.shapes:
        # gelinkte plaatjes overslaan
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            continue
        new_el = deepcopy(shape.element)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # nu oude tekst wissen (behalve titel)
    clear_all_text_except_title(dest)

    if logo_bytes:
        add_logo_to_slide(dest, logo_bytes)

    return dest


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


# ---------- MAIN ----------
def docx_to_pptx_hybrid(file_like):
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)

    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        print("⚠️ template niet gevonden, lege ppt.")
        prs = Presentation()

    doc = Document(file_like)
    blocks = docx_to_blocks(doc)  # hier zitten nu netjes INSTALLATIEKABEL, SOORTEN KABELS, ...

    logo_bytes = get_logo_bytes()

    # zorg dat er iig 1 slide is
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])

    # eerste dia
    first_slide = prs.slides[0]
    if logo_bytes:
        add_logo_to_slide(first_slide, logo_bytes)

    if blocks:
        get_or_add_title(first_slide, blocks[0]["title"])
        body = "\n".join(blocks[0]["body"]) if blocks[0]["body"] else ""
        if len(body) > 500:
            body = summarize_with_ai(body)
        add_body_text(first_slide, body, top_inch=2.0)
    else:
        get_or_add_title(first_slide, "Les gegenereerd met AI")

    # volgende blokken -> nieuwe dia
    for block in blocks[1:]:
        slide = duplicate_slide_no_external_pics(prs, 0, logo_bytes=logo_bytes)
        get_or_add_title(slide, block["title"])
        body = "\n".join(block["body"]) if block["body"] else ""
        if len(body) > 500:
            body = summarize_with_ai(body)
        add_body_text(slide, body, top_inch=2.0)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
