import io
import os
from copy import deepcopy

import requests  # kun je weghalen als je echt geen cloud wilt
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.enum.shapes import MSO_SHAPE_TYPE

# -------------------------------------------------
# configuratie
# -------------------------------------------------
BASE_TEMPLATE_NAME = "basis layout.pptx"   # staat in /templates
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")

CHARS_PER_LINE = 75


# -------------------------------------------------
# HELPER: koppen uit Word herkennen
# -------------------------------------------------
def has_bold(para):
    return any(run.bold for run in para.runs)


def is_all_caps_heading(text: str) -> bool:
    """herken dingen als 'SOORTEN KABELS', 'XMVK KABEL'"""
    txt = text.strip()
    if not txt:
        return False
    if len(txt) <= 40 and txt.upper() == txt:
        return True
    return False


def para_text_plain(para):
    return "".join(r.text for r in para.runs if r.text).strip()


def docx_to_blocks(doc: Document):
    """
    Maak: [ {title: '...', body: ['regel', ...]}, ... ]
    Kop = vet OF ALL CAPS OF Heading.
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
            # nieuwe dia
            if current:
                blocks.append(current)
            current = {"title": txt, "body": []}
        else:
            # hoort bij huidige dia
            if current is None:
                current = {"title": "Lesstof", "body": []}
            current["body"].append(txt)

    if current:
        blocks.append(current)

    return blocks


# -------------------------------------------------
# HELPER: logo (simpel, alleen lokaal)
# -------------------------------------------------
def get_logo_bytes():
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()
    return None


def add_logo_to_slide(slide, logo_bytes):
    if not logo_bytes:
        return
    left = Inches(9.0 - 1.5)   # beetje van rechts
    top = Inches(0.2)
    width = Inches(1.5)
    slide.shapes.add_picture(io.BytesIO(logo_bytes), left, top, width=width)


# -------------------------------------------------
# HELPER: dia dupliceren en daarna leegmaken
# -------------------------------------------------
def duplicate_slide_clean(prs: Presentation, slide_index: int, logo_bytes=None):
    """
    1. kloon de slide (alle vormen blijven)
    2. wis ALLE tekstframes
    3. voeg logo toe
    """
    source = prs.slides[slide_index]
    blank = prs.slide_layouts[0]
    dest = prs.slides.add_slide(blank)

    # kopieer alle shapes behalve gelinkte plaatjes
    for shape in source.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            # die uit de template zijn vaak gelinkt → overslaan
            continue
        new_el = deepcopy(shape.element)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # alle tekst leegmaken (we willen zelf de titel/body plaatsen)
    for shape in dest.shapes:
        if hasattr(shape, "text_frame"):
            shape.text_frame.clear()

    # logo opnieuw
    if logo_bytes:
        add_logo_to_slide(dest, logo_bytes)

    return dest


# -------------------------------------------------
# HELPER: titel en body op vaste plek
# -------------------------------------------------
def set_title(slide, text: str):
    # probeer eerst bestaande titel
    if slide.shapes.title is not None:
        slide.shapes.title.text = text
        return

    # anders zelf een titelvak maken
    left = Inches(0.6)
    top = Inches(0.4)
    width = Inches(9.0)
    height = Inches(0.8)
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(28)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0, 0, 0)


def set_body(slide, text: str):
    left = Inches(0.8)
    top = Inches(2.0)     # vaste plek
    width = Inches(8.5)
    height = Inches(4.0)
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
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

    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()  # fallback

    # docx → blokken
    doc = Document(file_like)
    blocks = docx_to_blocks(doc)

    # logo
    logo_bytes = get_logo
