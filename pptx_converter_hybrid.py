import io
import os
from copy import deepcopy

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
# je template moet hier staan: <project>/templates/basis layout.pptx
BASE_TEMPLATE_NAME = "basis layout.pptx"
# optioneel logo: <project>/assets/logo.png
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


# -------------------------------------------------
# simpele fallback samenvatter (geen OpenAI nodig)
# -------------------------------------------------
def summarize_text(text: str, max_chars: int = 450) -> str:
    text = text.strip()
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rsplit(" ", 1)[0] + "..."


# -------------------------------------------------
# DOCX → blokken
# -------------------------------------------------
def has_bold(para) -> bool:
    return any(run.bold for run in para.runs)


def is_all_caps_heading(text: str) -> bool:
    """
    Herken jouw stijl kopjes: korte regels in hoofdletters zoals 'XMVK KABEL'
    """
    txt = text.strip()
    if not txt:
        return False
    if len(txt) <= 40 and txt.upper() == txt:
        return True
    return False


def para_text_plain(para) -> str:
    return "".join(r.text for r in para.runs if r.text).strip()


def docx_to_blocks(doc: Document):
    """
    Maak een lijst van blokken:
    [
      {"title": "INSTALLATIEKABEL", "body": ["uitleg...", "nog een regel..."]},
      {"title": "SOORTEN KABELS", "body": [...]},
      ...
    ]
    Kop = Heading of vet of ALL CAPS.
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
            # start nieuw blok
            if current:
                blocks.append(current)
            current = {"title": txt, "body": []}
        else:
            # hoort bij huidig blok
            if current is None:
                current = {"title": "Lesstof", "body": []}
            current["body"].append(txt)

    if current:
        blocks.append(current)

    return blocks


# -------------------------------------------------
# logo helpers (alleen lokaal)
# -------------------------------------------------
def get_logo_bytes():
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()
    return None


def add_logo_to_slide(slide, logo_bytes):
    if not logo_bytes:
        return
    # positie kun je later tweaken
    left = Inches(9.0 - 1.5)  # beetje van rechts
    top = Inches(0.2)
    width = Inches(1.5)
    slide.shapes.add_picture(io.BytesIO(logo_bytes), left, top, width=width)


# -------------------------------------------------
# pptx helpers
# -------------------------------------------------
def duplicate_slide_clean(prs: Presentation, slide_index: int, logo_bytes=None):
    """
    1. kloon de slide (vormgeving)
    2. sla gelinkte plaatjes uit template over (die gaven dat privacy-kruisje)
    3. wis ALLE tekst op de nieuwe slide
    4. zet logo erop
    """
    source = prs.slides[slide_index]
    blank = prs.slide_layouts[0]
    dest = prs.slides.add_slide(blank)

    # shapes kopiëren
    for shape in source.shapes:
        # afbeeldingen uit template overslaan
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            continue
        new_el = deepcopy(shape.element)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # alle tekst leegmaken
    for shape in dest.shapes:
        if hasattr(shape, "text_frame"):
            shape.text_frame.clear()

    # logo weer toevoegen
    if logo_bytes:
        add_logo_to_slide(dest, logo_bytes)

    return dest


def set_title(slide, text: str):
    """
    Zet titel in bestaande title-placeholder, of maak zelf een titelvak.
    """
    if slide.shapes.title is not None:
        slide.shapes.title.text = text
        return

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
    """
    Plaats de tekst op een vaste plek onder de titel.
    """
    left = Inches(0.8)
    top = Inches(2.0)      # hier komt je tekst ALTIJD
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
# MAIN FUNCTIE
# -------------------------------------------------
def docx_to_pptx_hybrid(file_like):
    # 1. template laden
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)

    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()  # fallback als template mist

    # 2. docx → blokken
    doc = Document(file_like)
    blocks = docx_to_blocks(doc)

    # 3. logo (optioneel)
    logo_bytes = get_logo_bytes()

    # 4. er moet minstens 1 slide zijn
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])

    # 5. eerste slide schoonmaken en vullen
    first = prs.slides[0]

    # alle tekst op de eerste slide leegmaken (anders blijft template-tekst staan)
    for shape in first.shapes:
        if hasattr(shape, "text_frame"):
            shape.text_frame.clear()

    # logo erop
    if logo_bytes:
        add_logo_to_slide(first, logo_bytes)

    if blocks:
        # titel + body uit eerste blok
        set_title(first, blocks[0]["title"])
        body_txt = "\n".join(blocks[0]["body"]) if blocks[0]["body"] else ""
        body_txt = summarize_text(body_txt, max_chars=500)
        set_body(first, body_txt)
    else:
        set_title(first, "Les gegenereerd met AI")

    # 6. overige blokken → nieuwe dia’s
    for block in blocks[1:]:
        slide = duplicate_slide_clean(prs, 0, logo_bytes=logo_bytes)
        set_title(slide, block["title"])
        body_txt = "\n".join(block["body"]) if block["body"] else ""
        body_txt = summarize_text(body_txt, max_chars=500)
        set_body(slide, body_txt)

    # 7. teruggeven als bytes
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
