import io
import os
import json
from copy import deepcopy

from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

# =========================================================
# CONFIG
# =========================================================
BASE_TEMPLATE_NAME = "basis layout.pptx"
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


# =========================================================
# DOCX → blokken
# =========================================================
def docx_to_blocks(doc: Document):
    """
    We pakken heading/vet/ALL CAPS als kop.
    Alles eronder hoort bij die kop.
    """
    blocks = []
    current_title = None
    current_body = []

    for para in doc.paragraphs:
        txt = "".join(r.text for r in para.runs if r.text).strip()
        if not txt:
            continue

        is_heading = (
            (para.style and para.style.name and para.style.name.lower().startswith("heading"))
            or any(r.bold for r in para.runs)
            or (len(txt) <= 50 and txt.upper() == txt)
        )

        if is_heading:
            if current_title or current_body:
                blocks.append({
                    "title": current_title,
                    "body": "\n".join(current_body).strip()
                })
            current_title = txt
            current_body = []
        else:
            current_body.append(txt)

    if current_title or current_body:
        blocks.append({
            "title": current_title,
            "body": "\n".join(current_body).strip()
        })

    return blocks


# =========================================================
# AI → vmbo-lesblok
# =========================================================
def ai_vmbo_block_from_text(raw_text: str) -> dict:
    """
    Roept OpenAI aan. Als er een rate limit of iets anders misgaat,
    gooien we een duidelijke fout zodat de app een melding kan tonen.
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Geen OPENAI_API_KEY ingesteld. Zet je sleutel in de omgeving.")

    from openai import OpenAI, RateLimitError

    client = OpenAI(api_key=api_key)

    prompt = f"""
Je bent docent installatietechniek op een vmbo-school (basis/kader/GL).
Maak van de onderstaande tekst 1 dia.

Eisen:
- 1 duidelijke vmbo-titel (max 8 woorden). Niet letterlijk de eerste regel herhalen.
- 2 of 3 korte vertellende zinnen in de je-vorm.
- 1 controlevraag die past bij deze uitleg.
- Gebruik eenvoudige woorden.

Geef ALLEEN JSON in dit formaat:
{{
  "title": "goede titel",
  "text": ["zin 1", "zin 2", "zin 3"],
  "check": "vraag ..."
}}

Tekst:
{raw_text}
"""

    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
        )
    except RateLimitError:
        # hier GEEN fallback meer, dit wil je juist zien in je app
        raise RuntimeError("AI-limiet bereikt bij OpenAI. Probeer het zo nog een keer.")
    except Exception as e:
        raise RuntimeError(f"AI kon geen lesblok maken: {e}")

    # nu proberen te parsen
    try:
        data = json.loads(resp.choices[0].message.content)
    except Exception:
        raise RuntimeError("AI gaf geen geldig JSON terug.")

    title = (data.get("title") or "").strip()
    text_lines = [t.strip() for t in (data.get("text") or []) if t.strip()]
    check = (data.get("check") or "").strip()

    if not title or not text_lines:
        raise RuntimeError("AI gaf te weinig inhoud terug voor deze dia.")

    return {
        "title": title,
        "text": text_lines[:3],
        "check": check or "Leg uit waarom je dit zo moet doen."
    }


# =========================================================
# Template / PPTX helpers
# =========================================================
def get_logo_bytes():
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()
    return None


def add_logo(slide, logo_bytes):
    if not logo_bytes:
        return
    left = Inches(7.5)
    top = Inches(0.2)
    width = Inches(1.5)
    slide.shapes.add_picture(io.BytesIO(logo_bytes), left, top, width=width)


def get_positions_from_first_slide(slide):
    text_shapes = [s for s in slide.shapes if hasattr(s, "text") and s.text and s.text.strip()]
    if len(text_shapes) >= 2:
        t, b = text_shapes[0], text_shapes[1]
        return {
            "title": {"left": t.left, "top": t.top, "width": t.width, "height": t.height},
            "body": {"left": b.left, "top": b.top, "width": b.width, "height": b.height},
        }
    # fallback
    return {
        "title": {"left": Inches(0.8), "top": Inches(0.8), "width": Inches(9), "height": Inches(1)},
        "body": {"left": Inches(0.8), "top": Inches(3), "width": Inches(10), "height": Inches(3)},
    }


def duplicate_slide_clean(prs: Presentation, slide_index: int):
    src = prs.slides[slide_index]
    dest = prs.slides.add_slide(prs.slide_layouts[0])

    for shp in src.shapes:
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


def place_vmbo_text(slide, lines: list[str], check: str, pos: dict):
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tf = box.text_frame
    tf.word_wrap = True

    first = True
    for line in lines:
        p = tf.add_paragraph() if not first else tf.paragraphs[0]
        p.text = line
        first = False

    # lege regel
    tf.add_paragraph().text = ""

    if check:
        p = tf.add_paragraph()
        p.text = check
        for r in p.runs:
            r.font.bold = True
            r.font.size = Pt(16)

    # stijl
    for p in tf.paragraphs:
        for r in p.runs:
            if not r.font.size:
                r.font.name = "Arial"
                r.font.size = Pt(16)
                r.font.color.rgb = RGBColor(0, 0, 0)


# =========================================================
# MAIN
# =========================================================
def docx_to_pptx_hybrid(file_like):
    # template
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)
    prs = Presentation(template_path) if os.path.exists(template_path) else Presentation()

    # docx → blokken
    doc = Document(file_like)
    blocks = docx_to_blocks(doc)
    if not blocks:
        # niks zinnigs uit docx
        raise RuntimeError("Er zijn geen onderdelen gevonden in het Word-bestand.")

    # logo
    logo_bytes = get_logo_bytes()

    # iig 1 dia
    if not prs.slides:
        prs.slides.add_slide(prs.slide_layouts[0])
    first_slide = prs.slides[0]

    positions = get_positions_from_first_slide(first_slide)

    # dia 1 leegmaken
    for shp in first_slide.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()
    if logo_bytes:
        add_logo(first_slide, logo_bytes)

    # eerste blok → AI
    first_block = blocks[0]
    lesson = ai_vmbo_block_from_text(first_block.get("body") or first_block.get("title") or "")
    place_title(first_slide, lesson["title"], positions["title"])
    place_vmbo_text(first_slide, lesson["text"], lesson["check"], positions["body"])

    # volgende blokken
    for block in blocks[1:]:
        slide = duplicate_slide_clean(prs, 0)
        if logo_bytes:
            add_logo(slide, logo_bytes)
        lesson = ai_vmbo_block_from_text(block.get("body") or block.get("title") or "")
        place_title(slide, lesson["title"], positions["title"])
        place_vmbo_text(slide, lesson["text"], lesson["check"], positions["body"])

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out



