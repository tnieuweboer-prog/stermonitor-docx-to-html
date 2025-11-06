import io
import os
import json
import math
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
CHARS_PER_LINE = 75

BASE_TEMPLATE_NAME = "basis layout.pptx"  # jouw bestand
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")

# cloudinary (optioneel)
CLOUDINARY_CLOUD_NAME = os.getenv("CLOUDINARY_CLOUD_NAME", "")
CLOUDINARY_UPLOAD_PRESET = os.getenv("CLOUDINARY_UPLOAD_PRESET", "")
CLOUDINARY_LOGO_URL = os.getenv("CLOUDINARY_LOGO_URL", "")


# -------------------------------------------------
# AI helper (zoals eerder)
# -------------------------------------------------
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
            resp = client.chat_completions.create(
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
# DOCX helpers
# -------------------------------------------------
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


def docx_to_blocks(doc: Document):
    """
    Maak blokken: elke keer dat we een heading of vetgedrukte paragraaf zien,
    starten we een nieuw blok. Alles daarna (gewone tekst, lijsten) hoort bij dat blok.
    """
    blocks = []
    current_block = None

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        if not text:
            continue

        is_heading = para.style and para.style.name and para.style.name.startswith("Heading")
        is_bold = has_bold(para)

        if is_heading or is_bold:
            # nieuw blok starten
            if current_block:
                blocks.append(current_block)
            current_block = {
                "title": para_text_plain(para),
                "body": []
            }
        else:
            # hoort bij huidige blok
            if current_block is None:
                # tekst zonder kop → zet in algemene blok
                current_block = {"title": "Lesstof", "body": []}
            current_block["body"].append(text)

    if current_block:
        blocks.append(current_block)

    return blocks


# -------------------------------------------------
# Cloudinary helpers
# -------------------------------------------------
def upload_logo_to_cloudinary(local_path: str) -> str | None:
    if not CLOUDINARY_CLOUD_NAME:
        return None
    if not os.path.exists(local_path):
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
        else:
            print("Cloudinary upload mislukte:", resp.status_code, resp.text)
    except Exception as e:
        print("Cloudinary upload fout:", e)
    return None


def get_logo_bytes():
    # 1. als er al een url is, gebruik die
    if CLOUDINARY_LOGO_URL:
        try:
            r = requests.get(CLOUDINARY_LOGO_URL, timeout=10)
            if r.status_code == 200:
                return r.content
        except Exception:
            pass

    # 2. anders proberen lokaal up te loaden
    if os.path.exists(LOCAL_LOGO_PATH) and CLOUDINARY_CLOUD_NAME:
        url = upload_logo_to_cloudinary(LOCAL_LOGO_PATH)
        if url:
            try:
                r = requests.get(url, timeout=10)
                if r.status_code == 200:
                    return r.content
            except Exception:
                pass

    # 3. anders: gewoon lokaal inlezen (dan staat hij niet op cloudinary, maar wél embedded)
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()

    return None


# -------------------------------------------------
# PPTX helpers
# -------------------------------------------------
def add_logo_to_slide(slide, logo_bytes):
    if not logo_bytes:
        return
    left = Inches(9.0 - 1.5)  # beetje van rechts
    top = Inches(0.2)
    width = Inches(1.5)
    slide.shapes.add_picture(io.BytesIO(logo_bytes), left, top, width=width)


def duplicate_slide_no_external_pics(prs, slide_index=0, logo_bytes=None):
    """
    Kopieer dia 0, maar sla externe/gelinkte afbeeldingen over (die geven dat privacy-vierkant).
    """
    source = prs.slides[slide_index]
    blank_layout = prs.slide_layouts[0]
    dest = prs.slides.add_slide(blank_layout)

    for shape in source.shapes:
        # sommige templates hebben gelinkte images → die slaan we over
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            # skip, we voegen straks zelf het logo toe
            continue
        el = shape.element
        new_el = deepcopy(el)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    if logo_bytes:
        add_logo_to_slide(dest, logo_bytes)

    return dest


def add_body_text(slide, text, top_inch=2.0):
    """
    Voeg een tekstvak toe op een vaste plek.
    """
    left = Inches(0.8)
    top = Inches(top_inch)
    width = Inches(8.5)
    height = Inches(3.5)
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)


# -------------------------------------------------
# MAIN
# -------------------------------------------------
def docx_to_pptx_hybrid(file_like):
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)

    if not os.path.exists(template_path):
        print("⚠️ Template niet gevonden, maak lege presentatie.")
        prs = Presentation()
    else:
        prs = Presentation(template_path)

    # logo 1x regelen
    logo_bytes = get_logo_bytes()

    # docx lezen en omzetten naar blokken (kop + tekst)
    doc = Document(file_like)
    blocks = docx_to_blocks(doc)

    # eerste dia van template gebruiken als basis
    # en meteen logo erop
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])
    prs.slides[0].shapes.title.text = blocks[0]["title"] if blocks else "Les gegenereerd met AI"
    if logo_bytes:
        add_logo_to_slide(prs.slides[0], logo_bytes)
    if blocks and blocks[0]["body"]:
        body_text = "\n".join(blocks[0]["body"])
        add_body_text(prs.slides[0], body_text, top_inch=2.0)

    # overige blokken → nieuwe dia’s
    for block in blocks[1:]:
        slide = duplicate_slide_no_external_pics(prs, 0, logo_bytes=logo_bytes)
        # titel invullen
        if slide.shapes.title:
            slide.shapes.title.text = block["title"]
        # tekst invullen
        body_text = "\n".join(block["body"]) if block["body"] else ""
        # evt. nog korter maken
        if len(body_text) > 450:
            body_text = summarize_with_ai(body_text)
        add_body_text(slide, body_text, top_inch=2.0)

    # presentatie teruggeven
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out



