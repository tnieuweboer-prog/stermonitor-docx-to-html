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
BASE_TEMPLATE_NAME = "basis layout.pptx"
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")

# =========================================================
# 1. Zinnen opdelen
# =========================================================
def split_sentences(text):
    text = (text or "").strip()
    if not text:
        return []
    parts = re.split(r"[.!?]\s+", text)
    parts = [p.strip(" .!?") for p in parts if p.strip()]
    return parts

# =========================================================
# 2. AI helper
# =========================================================
def ai_vmbo_block_from_text(raw_text):
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return fallback_vmbo_block(raw_text)

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        prompt = f"""
Je bent docent installatietechniek op een vmbo-school (basis/kader/GL).
Je krijgt een technisch stukje tekst uit een handboek.

Maak er één duidelijke dia van voor een lespresentatie.

Regels:
- Schrijf een korte, begrijpelijke titel (max 8 woorden).
- De titel mag NIET letterlijk in de tekst herhaald worden.
- Schrijf 2–3 korte, vertellende zinnen in de je-vorm.
- Gebruik eenvoudige woorden, alsof je het uitlegt aan een leerling.
- Sluit af met één logische controlevraag bij het onderwerp.
- Geef ALLEEN geldig JSON in dit formaat:

{{
  "title": "goede zin als titel",
  "text": ["eerste uitlegzin", "tweede uitlegzin", "derde uitlegzin"],
  "check": "controlevraag"
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
        title = data.get("title", "Lesonderdeel").strip()
        text_lines = [t.strip() for t in data.get("text", []) if t.strip()]
        text_lines = text_lines[:3]
        check = data.get("check", "Wat gebeurt er als je dit verkeerd doet?").strip()

        return {"title": title, "text": text_lines, "check": check}

    except Exception:
        return fallback_vmbo_block(raw_text)

# =========================================================
# 3. Fallback zonder AI
# =========================================================
def fallback_vmbo_block(raw_text):
    text = (raw_text or "").strip()
    sentences = split_sentences(text)

    # Titel maken uit kernwoorden, niet letterlijk de eerste zin
    title = "Hoe werkt dit onderdeel?"
    if sentences:
        for s in sentences:
            if "leiding" in s or "bocht" in s or "afvoer" in s:
                title = re.sub(r"[^a-zA-Z0-9 ]", "", s).strip().capitalize()
                title = " ".join(title.split()[:7])
                break

    # 2–3 zinnen samenvatten
    text_lines = []
    for s in sentences[:3]:
        s = s.replace("Men ", "Je ").replace("men ", "je ")
        if not s.lower().startswith("je "):
            s = "Je " + s[0].lower() + s[1:]
        text_lines.append(s)

    if not text_lines:
        text_lines = [
            "Je leert hier hoe je leidingen goed aansluit.",
            "Zo kan water en lucht goed doorstromen.",
            "Dat voorkomt verstoppingen en lawaai.",
        ]

    # Vraag maken
    check = "Wat gebeurt er als je dit verkeerd doet?"
    for s in sentences:
        if "mag" in s or "niet" in s or "nooit" in s:
            check = "Waarom mag je dit niet zo doen?"
            break
        if "leiding" in s or "afvoer" in s:
            check = "Wat gebeurt er als je dit verkeerd aansluit?"
            break

    return {"title": title, "text": text_lines, "check": check}

# =========================================================
# 4. DOCX naar tekstblokken
# =========================================================
def plain_para_text(para):
    return "".join(r.text for r in para.runs if r.text).strip()

def para_is_heading_like(para):
    txt = plain_para_text(para)
    if not txt:
        return False
    if para.style and para.style.name and para.style.name.lower().startswith("heading"):
        return True
    if any(r.bold for r in para.runs):
        return True
    if len(txt) <= 50 and txt.upper() == txt:
        return True
    return False

def docx_to_raw_blocks(doc):
    blocks, current = [], []
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
        blocks = ["(Geen duidelijke indeling gevonden.)"]
    return blocks

# =========================================================
# 5. Template helpers
# =========================================================
def get_logo_bytes():
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()
    return None

def add_logo(slide, logo_bytes):
    if not logo_bytes:
        return
    left, top, width = Inches(7.5), Inches(0.2), Inches(1.5)
    slide.shapes.add_picture(io.BytesIO(logo_bytes), left, top, width=width)

def get_positions_from_first_slide(slide):
    shapes = [s for s in slide.shapes if hasattr(s, "text") and s.text.strip()]
    if len(shapes) >= 2:
        t, b = shapes[0], shapes[1]
        return {
            "title": {"left": t.left, "top": t.top, "width": t.width, "height": t.height},
            "body": {"left": b.left, "top": b.top, "width": b.width, "height": b.height},
        }
    return {
        "title": {"left": Inches(0.8), "top": Inches(0.8), "width": Inches(9), "height": Inches(1)},
        "body": {"left": Inches(0.8), "top": Inches(3), "width": Inches(10), "height": Inches(3)},
    }

def duplicate_slide_clean(prs, slide_index):
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

# =========================================================
# 6. Tekst plaatsen
# =========================================================
def place_title(slide, text, pos):
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tf = box.text_frame
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(28)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0, 0, 0)

def place_vmbo_text(slide, lines, check, pos):
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if not line.strip():
            continue
        p = tf.add_paragraph() if not first else tf.paragraphs[0]
        p.text = line
        first = False
    blank = tf.add_paragraph()
    blank.text = ""
    if check:
        p = tf.add_paragraph()
        p.text = check
        for r in p.runs:
            r.font.bold = True
            r.font.size = Pt(16)
    for p in tf.paragraphs:
        for r in p.runs:
            if not r.font.size:
                r.font.name = "Arial"
                r.font.size = Pt(16)

# =========================================================
# 7. MAIN
# =========================================================
def docx_to_pptx_hybrid(file_like):
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)
    prs = Presentation(template_path) if os.path.exists(template_path) else Presentation()
    doc = Document(file_like)
    blocks = docx_to_raw_blocks(doc)
    logo_bytes = get_logo_bytes()
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

    # eerste blok
    lesson = ai_vmbo_block_from_text(blocks[0]) if blocks else {"title": "Lesonderdeel", "text": [], "check": ""}
    place_title(first_slide, lesson["title"], positions["title"])
    place_vmbo_text(first_slide, lesson["text"], lesson["check"], positions["body"])

    # volgende dia’s
    for block_text in blocks[1:]:
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



