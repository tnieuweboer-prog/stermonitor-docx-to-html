import io
import os
import re
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
BASE_TEMPLATE_NAME = "basis layout.pptx"   # /templates/basis layout.pptx
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


# =========================================================
# HULP: zinnen uit tekst
# =========================================================
def split_into_sentences(text: str) -> list[str]:
    text = (text or "").strip()
    if not text:
        return []
    # heel simpele zin-split
    parts = re.split(r"[\.!\?]\s+", text)
    parts = [p.strip(" .?!") for p in parts if p.strip()]
    return parts


# =========================================================
# 1. AI-helper: vmbo-lesblok
# =========================================================
def ai_vmbo_block_from_text(raw_text: str) -> dict:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return fallback_vmbo_block(raw_text)

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        prompt = f"""
Je bent docent installatietechniek op een vmbo-school (basis/kader/GL).
Je krijgt een stukje technische tekst.

Maak hiervan lesmateriaal voor 1 dia.

- schrijf 1 pakkende vmbo-titel (max 6 woorden)
- schrijf 3 korte zinnen in je-vorm
- schrijf 1 inhoudelijke controlevraag
- gebruik eenvoudige woorden
- geef alleen JSON in dit formaat:
{{
  "title": "...",
  "text": ["...", "...", "..."],
  "check": "..."
}}

Tekst:
{raw_text}
"""
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
        )
        data = json.loads(resp.choices[0].message.content)

        title = (data.get("title") or "Lesonderdeel").strip()
        text_lines = data.get("text") or []
        text_lines = [t.strip() for t in text_lines if t.strip()]
        text_lines = text_lines[:3]
        check = (data.get("check") or "Wat gebeurt er als je dit niet goed doet?").strip()

        return {
            "title": title,
            "text": text_lines,
            "check": check,
        }
    except Exception:
        return fallback_vmbo_block(raw_text)


# =========================================================
# 2. Slimmere fallback (als er geen AI is)
# =========================================================
def fallback_vmbo_block(raw_text: str) -> dict:
    raw_text = (raw_text or "").strip()

    # zinnen halen
    sentences = split_into_sentences(raw_text)

    # titel maken: eerste zin inkorten
    if sentences:
        first = sentences[0]
        words = first.split()
        # bv. "Aansluiting op liggende leiding" / "Leiding niet vernauwen"
        title = " ".join(words[:6]).capitalize()
    else:
        title = "Lesonderdeel"

    # 3 vertelzinnen maken
    text_lines = []
    for s in sentences[:3]:
        # wat vriendelijker maken
        s = s.replace("men ", "je ")
        s = s.replace("Men ", "Je ")
        if not s.lower().startswith("je "):
            s = "Je " + s[0].lower() + s[1:]
        text_lines.append(s)

    if not text_lines:
        text_lines = [
            "Je leert hier een stap uit de installatie.",
            "Lees mee met je docent.",
            "Let op wat er niet mag.",
        ]

    # check-vraag maken op basis van de eerste zin
    if sentences:
        first = sentences[0].lower()
        if "niet" in first or "mag" in first or "nooit" in first:
            check = "Waarom mag je dit niet zo aansluiten?"
        elif "bocht" in first or "leiding" in first:
            check = "Wat gebeurt er als je dit verkeerd doet?"
        else:
            check = "Kun je uitleggen waarom je dit zo doet?"
    else:
        check = "Wat is hier het belangrijkste?"

    return {
        "title": title,
        "text": text_lines,
        "check": check,
    }


# =========================================================
# 3. DOCX analyseren → tekstblokken
# =========================================================
def plain_para_text(para) -> str:
    return "".join(r.text for r in para.runs if r.text).strip()


def para_is_heading_like(para) -> bool:
    txt = plain_para_text(para)
    if not txt:
        return False
    # echte heading
    if para.style and para.style.name and para.style.name.lower().startswith("heading"):
        return True
    # vet
    if any(r.bold for r in para.runs):
        return True
    # ALL CAPS kort
    if len(txt) <= 50 and txt.upper() == txt:
        return True
    return False


def docx_to_raw_blocks(doc: Document) -> list[str]:
    blocks = []
    current = []

    for para in doc.paragraphs:
        txt = plain_para_text(para)
        if not txt:
            continue

        if para_is_heading_like(para):
            if current:
                blocks.append("\n".join(current).strip())
            current = [txt]
        else:
            current.append(txt)

    if current:
        blocks.append("\n".join(current).strip())

    if not blocks:
        blocks = ["(Dit document had geen duidelijke indeling.)"]

    return blocks


# =========================================================
# 4. Template helpers
# =========================================================
def get_logo_bytes():
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()
    return None


def add_logo(slide, logo_bytes):
    if not logo_bytes:
        return
    left = Inches(9.0 - 1.5)
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
        "title": {"left": Inches(0.6), "top": Inches(0.8), "width": Inches(9), "height": Inches(0.8)},
        "body": {"left": Inches(0.6), "top": Inches(3.4), "width": Inches(11.5), "height": Inches(2)},
    }


def duplicate_slide_clean(prs: Presentation, slide_index: int):
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


def place_vmbo_text(slide, lines: list[str], check: str, pos: dict):
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tf = box.text_frame
    tf.word_wrap = True

    # vertelzinnen
    first = True
    for line in lines:
        line = line.strip()
        if not line:
            continue
        p = tf.add_paragraph() if not first else tf.paragraphs[0]
        p.text = line
        first = False

    # lege regel
    blank = tf.add_paragraph()
    blank.text = ""

    # vraag
    if check:
        p = tf.add_paragraph()
        p.text = check
        for r in p.runs:
            r.font.bold = True
            r.font.name = "Arial"
            r.font.size = Pt(16)

    # stijl voor gewone zinnen
    for p in tf.paragraphs:
        for r in p.runs:
            if not r.font.size:
                r.font.name = "Arial"
                r.font.size = Pt(16)
                r.font.color.rgb = RGBColor(0, 0, 0)


# =========================================================
# 5. MAIN
# =========================================================
def docx_to_pptx_hybrid(file_like):
    # template
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # docx → blokken
    doc = Document(file_like)
    raw_blocks = docx_to_raw_blocks(doc)

    # logo
    logo_bytes = get_logo_bytes()

    # minstens 1 dia
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])
    first_slide = prs.slides[0]

    # posities uit dia 1
    positions = get_positions_from_first_slide(first_slide)

    # dia 1 leegmaken
    for shp in first_slide.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()
    if logo_bytes:
        add_logo(first_slide, logo_bytes)

    # eerste blok
    if raw_blocks:
        lesson = ai_vmbo_block_from_text(raw_blocks[0])
        place_title(first_slide, lesson["title"], positions["title"])
        place_vmbo_text(first_slide, lesson["text"], lesson["check"], positions["body"])
    else:
        place_title(first_slide, "Les gegenereerd met AI", positions["title"])

    # overige blokken
    for block_text in raw_blocks[1:]:
        slide = duplicate_slide_clean(prs, 0)
        if logo_bytes:
            add_logo(slide, logo_bytes)
        lesson = ai_vmbo_block_from_text(block_text)
        place_title(slide, lesson["title"], positions["title"])
        place_vmbo_text(slide, lesson["text"], lesson["check"], positions["body"])

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out



