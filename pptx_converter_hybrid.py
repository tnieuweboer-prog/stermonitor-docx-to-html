import io
import os
from copy import deepcopy

from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# =========================================================
# CONFIG
# =========================================================
BASE_TEMPLATE_NAME = "basis layout.pptx"   # verwacht in /templates
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


# =========================================================
# GENERIC TEXT HELPERS
# =========================================================
def plain_para_text(para) -> str:
    return "".join(r.text for r in para.runs if r.text).strip()


def para_is_bold(para) -> bool:
    return any(r.bold for r in para.runs)


def para_is_allcaps(text: str) -> bool:
    text = text.strip()
    if not text:
        return False
    # vaak in lesdocs: korte koppen in caps
    return len(text) <= 50 and text.upper() == text


def para_is_heading_style(para) -> bool:
    return bool(para.style and para.style.name and para.style.name.lower().startswith("heading"))


def para_is_list(para) -> bool:
    # simpele lijst-detectie
    if para.style and "list" in para.style.name.lower():
        return True
    # punt/nummer aan begin
    txt = (para.text or "").lstrip()
    if txt[:2] in ("- ", "• "):
        return True
    if txt[:3].isdigit() and txt[2] == ".":
        return True
    return False


def summarize_text(text: str, max_chars: int = 350) -> str:
    text = text.strip()
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rsplit(" ", 1)[0] + "..."


# =========================================================
# 1) DOCX → STRUCTUUR
# =========================================================
def docx_to_blocks_generic(doc: Document):
    """
    Probeert van een willekeurig Word-bestand een lijst met blokken te maken:
    [
      { "title": "...", "body": [ "regel", "regel" ], "bullets": [...] },
      ...
    ]
    We herkennen:
    - 'echte' headings
    - vetgedrukte regels
    - ALL CAPS-koppen
    - lijsten
    Als er niks te herkennen is, maken we kunstmatige blokken.
    """
    blocks = []
    current = None
    block_counter = 1

    for para in doc.paragraphs:
        txt = plain_para_text(para)
        if not txt:
            continue

        is_heading = para_is_heading_style(para) or para_is_bold(para) or para_is_allcaps(txt)

        if is_heading:
            # nieuwe sectie
            if current:
                blocks.append(current)
            current = {
                "title": txt,
                "body": [],
                "bullets": [],
            }
        else:
            # bepalen of dit bullets zijn of body
            if para_is_list(para):
                if current is None:
                    current = {
                        "title": f"Onderdeel {block_counter}",
                        "body": [],
                        "bullets": [],
                    }
                    block_counter += 1
                current["bullets"].append(txt.lstrip("-• ").strip())
            else:
                if current is None:
                    current = {
                        "title": f"Onderdeel {block_counter}",
                        "body": [],
                        "bullets": [],
                    }
                    block_counter += 1
                current["body"].append(txt)

    if current:
        blocks.append(current)

    # als het document echt kaal was
    if not blocks:
        blocks = [{
            "title": "Lesstof",
            "body": ["(Geen structuur gevonden in het document.)"],
            "bullets": [],
        }]

    return blocks


# =========================================================
# 2) BLOCk → LESSONUP-STIJL
# =========================================================
def rewrite_block_to_lessonup(block: dict) -> dict:
    """
    Neemt 1 blok uit het Word-document en maakt er een vmbo-/LessonUp-dia van.
    Output:
    {
      "title": "...",
      "bullets": ["...", "...", ...],
      "check": "..."
    }
    """
    title = block.get("title", "Lesstof").strip()
    body_lines = block.get("body", [])
    list_lines = block.get("bullets", [])

    # basis-inhoud: body + bullets samen
    merged_text = " ".join(body_lines).strip()
    merged_text = summarize_text(merged_text, 400)

    bullets = []

    # 1. als er al bullets waren in Word: neem er max 3 over, maar kort
    for b in list_lines[:3]:
        bullets.append(summarize_text(b, 120))

    # 2. als er geen bullets waren, maak ze uit de body
    if not bullets:
        # hak de body op in zinnen
        parts = merged_text.replace(". ", ".\n").split("\n")
        for p in parts:
            p = p.strip(". ").strip()
            if not p:
                continue
            bullets.append(p)
            if len(bullets) >= 3:
                break

    # 3. voeg 1 “waarom” toe voor vmbo
    if len(bullets) < 4:
        if "kabel" in title.lower():
            bullets.append("Kabels zijn veiliger dan losse draden.")
        else:
            bullets.append("Dit heb je nodig tijdens het werken aan installaties.")

    # maximaal 4
    bullets = bullets[:4]

    # checkvraag maken
    if "kabel" in title.lower():
        check = "Waarom kies je hier voor deze kabel?"
    else:
        check = f"Wanneer gebruik je {title.lower()}?"

    return {
        "title": title.title(),
        "bullets": bullets,
        "check": check,
    }


# =========================================================
# 3) TEMPLATE HULP
# =========================================================
def get_logo_bytes():
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()
    return None


def add_logo(slide, logo_bytes):
    if not logo_bytes:
        return
    # positie zo laten als eerder
    left = Inches(9.0 - 1.5)
    top = Inches(0.2)
    width = Inches(1.5)
    slide.shapes.add_picture(io.BytesIO(logo_bytes), left, top, width=width)


def get_title_and_body_positions_from_slide(slide):
    """
    Probeer de posities van de eerste 2 tekstvormen uit jouw template-dia te halen.
    Als dat niet lukt: fallback waarden gebruiken.
    """
    text_shapes = [s for s in slide.shapes if hasattr(s, "text") and s.text and s.text.strip()]
    if len(text_shapes) >= 2:
        t, b = text_shapes[0], text_shapes[1]
        return {
            "title": {"left": t.left, "top": t.top, "width": t.width, "height": t.height},
            "body": {"left": b.left, "top": b.top, "width": b.width, "height": b.height},
        }
    # fallback
    return {
        "title": {"left": Inches(0.6), "top": Inches(0.8), "width": Inches(9), "height": Inches(0.8)},
        "body": {"left": Inches(0.6), "top": Inches(3.4), "width": Inches(11.5), "height": Inches(2)},
    }


def duplicate_slide_clean(prs: Presentation, slide_index: int):
    """
    Kloon een dia, maar:
    - kopieer geen gelinkte plaatjes (die geven dat privacy-kruisje)
    - wis alle tekst
    """
    source = prs.slides[slide_index]
    dest = prs.slides.add_slide(prs.slide_layouts[0])

    for shp in source.shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
            continue
        new_el = deepcopy(shp.element)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    for shp in dest.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()

    return dest


def place_title(slide, text: str, pos: dict):
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tf = box.text_frame
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(28)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0, 0, 0)


def place_lessonup_body(slide, bullets: list[str], check: str, pos: dict):
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tf = box.text_frame
    tf.word_wrap = True

    first = True
    for b in bullets:
        p = tf.add_paragraph() if not first else tf.paragraphs[0]
        p.text = "• " + b
        first = False

    if check:
        p = tf.add_paragraph()
        p.text = f"Check: {check}"
        for r in p.runs:
            r.font.bold = True


# =========================================================
# MAIN ENTRYPOINT
# =========================================================
def docx_to_pptx_hybrid(file_like):
    """
    Hoofdfunctie: neemt ELK .docx en maakt er jouw vmbo-/LessonUp-stijl powerpoint van.
    """
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)

    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # docx inlezen
    doc = Document(file_like)

    # stap 1: analyseren
    blocks = docx_to_blocks_generic(doc)

    # logo
    logo_bytes = get_logo_bytes()

    # zorg dat er minimaal 1 dia is
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])

    # posities uit dia 0
    first_slide = prs.slides[0]
    positions = get_title_and_body_positions_from_slide(first_slide)

    # eerste dia leegmaken
    for shp in first_slide.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()
    if logo_bytes:
        add_logo(first_slide, logo_bytes)

    # eerste blok vullen
    if blocks:
        lesson = rewrite_block_to_lessonup(blocks[0])
        place_title(first_slide, lesson["title"], positions["title"])
        place_lessonup_body(first_slide, lesson["bullets"], lesson["check"], positions["body"])
    else:
        place_title(first_slide, "Les gegenereerd met AI", positions["title"])

    # de rest van de blokken
    for block in blocks[1:]:
        slide = duplicate_slide_clean(prs, 0)
        if logo_bytes:
            add_logo(slide, logo_bytes)
        lesson = rewrite_block_to_lessonup(block)
        place_title(slide, lesson["title"], positions["title"])
        place_lessonup_body(slide, lesson["bullets"], lesson["check"], positions["body"])

    # teruggeven
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

