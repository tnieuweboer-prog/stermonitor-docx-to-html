import io
import os
from copy import deepcopy

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

# -------------------------------------------------
# configuratie
# -------------------------------------------------
BASE_TEMPLATE_NAME = "basis layout.pptx"      # jouw template in /templates
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


# -------------------------------------------------
# simpele samenvatter
# -------------------------------------------------
def summarize_text(text: str, max_chars: int = 500) -> str:
    text = text.strip()
    if len(text) <= max_chars:
        return text
    # netjes op woord afkappen
    return text[:max_chars].rsplit(" ", 1)[0] + "..."


# -------------------------------------------------
# DOCX → blokken: elke kop + tekst eronder
# -------------------------------------------------
def has_bold(para):
    return any(run.bold for run in para.runs)


def is_all_caps_heading(text: str) -> bool:
    text = text.strip()
    if not text:
        return False
    # jouw document gebruikt veel korte ALL CAPS koppen
    return len(text) <= 40 and text.upper() == text


def para_text_plain(para):
    return "".join(r.text for r in para.runs if r.text).strip()


def docx_to_blocks(doc: Document):
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
            if current is None:
                current = {"title": "Lesstof", "body": []}
            current["body"].append(txt)

    if current:
        blocks.append(current)
    return blocks


# -------------------------------------------------
# logo (alleen lokaal, simpel)
# -------------------------------------------------
def get_logo_bytes():
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()
    return None


def add_logo_to_slide(slide, logo_bytes):
    if not logo_bytes:
        return
    # je plaatje stond bij jou rechts onder, maar je kunt dit aanpassen
    left = Inches(9.0 - 1.5)
    top = Inches(0.2)
    width = Inches(1.5)
    slide.shapes.add_picture(io.BytesIO(logo_bytes), left, top, width=width)


# -------------------------------------------------
# posities uit dia 1 halen
# -------------------------------------------------
def get_title_and_body_positions_from_slide(slide):
    """
    We zoeken in dia 1 naar de eerste 2 shapes met echte tekst.
    1e = titel, 2e = body.
    Als we ze niet vinden, gebruiken we default posities.
    """
    text_shapes = []
    for shp in slide.shapes:
        if hasattr(shp, "text") and shp.text and shp.text.strip():
            text_shapes.append(shp)

    if len(text_shapes) >= 2:
        title_shape = text_shapes[0]
        body_shape = text_shapes[1]
        return {
            "title": {
                "left": title_shape.left,
                "top": title_shape.top,
                "width": title_shape.width,
                "height": title_shape.height,
            },
            "body": {
                "left": body_shape.left,
                "top": body_shape.top,
                "width": body_shape.width,
                "height": body_shape.height,
            },
        }

    # fallback (ongeveer wat jij had)
    return {
        "title": {
            "left": Inches(0.55),
            "top": Inches(0.85),
            "width": Inches(9.0),
            "height": Inches(0.8),
        },
        "body": {
            "left": Inches(0.55),
            "top": Inches(3.4),
            "width": Inches(11.65),
            "height": Inches(1.45),
        },
    }


# -------------------------------------------------
# dia klonen en leegmaken
# -------------------------------------------------
def duplicate_slide_clean(prs: Presentation, slide_index: int):
    source = prs.slides[slide_index]
    blank = prs.slide_layouts[0]
    dest = prs.slides.add_slide(blank)

    # alle vormen kopiëren behalve plaatjes (logo uit template was gelinkt)
    for shp in source.shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
            continue
        new_el = deepcopy(shp.element)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # alle tekst leegmaken
    for shp in dest.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()

    return dest


# -------------------------------------------------
# titel/body plaatsen op exact dezelfde plek als dia 1
# -------------------------------------------------
def place_title(slide, text: str, pos):
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tf = box.text_frame
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(28)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0, 0, 0)


def place_body(slide, text: str, pos):
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
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
    # 1. template laden
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # 2. docx → blokken
    doc = Document(file_like)
    blocks = docx_to_blocks(doc)

    # 3. logo
    logo_bytes = get_logo_bytes()

    # 4. er moet minstens 1 dia zijn
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])

    first_slide = prs.slides[0]

    # 5. posities uit dia 1 halen (die jij goed hebt gezet)
    positions = get_title_and_body_positions_from_slide(first_slide)

    # 6. eerste dia leegmaken (alle tekst weg) maar vormgeving laten staan
    for shp in first_slide.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()

    # logo erop
    if logo_bytes:
        add_logo_to_slide(first_slide, logo_bytes)

    # 7. eerste blok invullen
    if blocks:
        place_title(first_slide, blocks[0]["title"], positions["title"])
        body_text = "\n".join(blocks[0]["body"]) if blocks[0]["body"] else ""
        body_text = summarize_text(body_text, 500)
        place_body(first_slide, body_text, positions["body"])
    else:
        place_title(first_slide, "Les gegenereerd met AI", positions["title"])

    # 8. overige blokken → nieuwe dia’s met exact dezelfde posities
    for block in blocks[1:]:
        slide = duplicate_slide_clean(prs, 0)
        if logo_bytes:
            add_logo_to_slide(slide, logo_bytes)
        place_title(slide, block["title"], positions["title"])
        body_text = "\n".join(block["body"]) if block["body"] else ""
        body_text = summarize_text(body_text, 500)
        place_body(slide, body_text, positions["body"])

    # 9. teruggeven als bytes
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
