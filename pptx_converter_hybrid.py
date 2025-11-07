import io
import os
import json
from copy import deepcopy

from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

# -------------------------------------------------
# CONFIG
# -------------------------------------------------
BASE_TEMPLATE_NAME = "basis layout.pptx"     # zet je template hier
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


# -------------------------------------------------
# AI helper: maak vmbo-lesblok
# -------------------------------------------------
def ai_vmbo_block_from_text(raw_text: str) -> dict:
    """
    Probeert van willekeurige tekst een vmbo-lesblok te maken.
    Verwacht JSON:
    {
      "title": "...",
      "bullets": ["...", "..."],
      "check": "..."
    }
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        # fallback zonder AI
        return fallback_vmbo_block(raw_text)

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        prompt = f"""
Je bent docent installatietechniek op een vmbo (basis/kader/GL).
Je krijgt een stuk ruwe les-tekst.
Maak hiervan lesmateriaal voor 1 PowerPoint-dia / LessonUp.

Eisen:
- titel moet eenvoudig en didactisch zijn, max 6 woorden (bv. "Afvoer goed aansluiten", "Soorten kabels", "Kabel in de grond")
- maak maximaal 4 bullets
- elke bullet 1 korte zin
- taalniveau vmbo
- voeg 1 check-vraag toe die bij de tekst past

Geef alleen JSON in dit formaat:
{{
  "title": "...",
  "bullets": ["...", "..."],
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
        # sanity
        return {
            "title": data.get("title") or "Lesonderdeel",
            "bullets": data.get("bullets") or ["Belangrijk punt uit de tekst."],
            "check": data.get("check") or "Wat is hier belangrijk?",
        }
    except Exception:
        return fallback_vmbo_block(raw_text)


def fallback_vmbo_block(raw_text: str) -> dict:
    raw_text = (raw_text or "").strip()
    if not raw_text:
        return {
            "title": "Lesonderdeel",
            "bullets": ["Leg in eigen woorden uit.", "Schrijf 1 toepassing op."],
            "check": "Waarom doe je dit zo?",
        }
    # pak eerste zin als titel
    short = raw_text.split(".")[0]
    short = short.strip()
    if len(short.split()) > 6:
        short = " ".join(short.split()[:6])
    title = short or "Lesonderdeel"

    bullets = []
    for line in raw_text.replace(". ", ".\n").split("\n"):
        line = line.strip(". ").strip()
        if not line:
            continue
        bullets.append(line)
        if len(bullets) >= 3:
            break
    if not bullets:
        bullets = ["Belangrijk punt uit de tekst."]
    check = "Wat gebeurt er als je dit niet goed doet?"
    return {"title": title, "bullets": bullets, "check": check}


# -------------------------------------------------
# DOCX analyseren → blokken
# -------------------------------------------------
def plain_para_text(para) -> str:
    return "".join(r.text for r in para.runs if r.text).strip()


def para_is_heading_like(para) -> bool:
    txt = plain_para_text(para)
    if not txt:
        return False
    # echte heading-style
    if para.style and para.style.name and para.style.name.lower().startswith("heading"):
        return True
    # dikgedrukt
    if any(r.bold for r in para.runs):
        return True
    # ALL CAPS korte regel
    if len(txt) <= 50 and txt.upper() == txt:
        return True
    return False


def docx_to_raw_blocks(doc: Document):
    """
    Probeert van elk docx een lijst teksten te maken.
    Elke 'heading-achtige' paragraaf start een nieuw blok.
    Alles eronder hoort bij dat blok.
    """
    blocks = []
    current = []

    for para in doc.paragraphs:
        txt = plain_para_text(para)
        if not txt:
            continue

        if para_is_heading_like(para):
            # nieuw blok
            if current:
                blocks.append("\n".join(current).strip())
            current = [txt]
        else:
            current.append(txt)

    if current:
        blocks.append("\n".join(current).strip())

    # als er echt niets herkenbaars was
    if not blocks:
        blocks = ["(Geen structuur gevonden in het document.)"]

    return blocks


# -------------------------------------------------
# template helpers
# -------------------------------------------------
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
    """
    Kijk op dia 1 waar jij je titel en tekst hebt gezet.
    We nemen de eerste 2 tekstvormen.
    """
    text_shapes = [s for s in slide.shapes if hasattr(s, "text") and s.text and s.text.strip()]
    if len(text_shapes) >= 2:
        t, b = text_shapes[0], text_shapes[1]
        return {
            "title": {"left": t.left, "top": t.top, "width": t.width, "height": t.height},
            "body": {"left": b.left, "top": b.top, "width": b.width, "height": b.height},
        }
    # fallback-positie
    return {
        "title": {"left": Inches(0.6), "top": Inches(0.8), "width": Inches(9), "height": Inches(0.8)},
        "body": {"left": Inches(0.6), "top": Inches(3.4), "width": Inches(11.5), "height": Inches(2)},
    }


def duplicate_slide_clean(prs: Presentation, slide_index: int):
    """
    Kloon dia en wis alle tekst, sla gelinkte plaatjes over.
    """
    source = prs.slides[slide_index]
    dest = prs.slides.add_slide(prs.slide_layouts[0])

    for shp in source.shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
            continue
        new_el = deepcopy(shp.element)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # tekst leeg
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


def place_body_lessonup(slide, bullets: list[str], check: str, pos: dict):
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tf = box.text_frame
    tf.word_wrap = True

    first = True
    for b in bullets:
        p = tf.add_paragraph() if not first else tf.paragraphs[0]
        p.text = f"• {b}"
        first = False

    if check:
        p = tf.add_paragraph()
        p.text = f"Check: {check}"
        for r in p.runs:
            r.font.bold = True


# -------------------------------------------------
# MAIN
# -------------------------------------------------
def docx_to_pptx_hybrid(file_like):
    # 1. template inladen
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # 2. docx analyseren
    doc = Document(file_like)
    raw_blocks = docx_to_raw_blocks(doc)  # lijst van tekstblokken

    # 3. logo
    logo_bytes = get_logo_bytes()

    # 4. zorg dat we een eerste dia hebben
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])
    first_slide = prs.slides[0]

    # 5. posities van dia 1 pakken
    positions = get_positions_from_first_slide(first_slide)

    # 6. eerste dia leegmaken
    for shp in first_slide.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()
    if logo_bytes:
        add_logo(first_slide, logo_bytes)

    # 7. eerste blok met AI herschrijven
    if raw_blocks:
        lesson = ai_vmbo_block_from_text(raw_blocks[0])
        place_title(first_slide, lesson["title"], positions["title"])
        place_body_lessonup(first_slide, lesson["bullets"], lesson["check"], positions["body"])
    else:
        place_title(first_slide, "Les gegenereerd met AI", positions["title"])

    # 8. overige blokken → nieuwe dia’s
    for block_text in raw_blocks[1:]:
        slide = duplicate_slide_clean(prs, 0)
        if logo_bytes:
            add_logo(slide, logo_bytes)
        lesson = ai_vmbo_block_from_text(block_text)
        place_title(slide, lesson["title"], positions["title"])
        place_body_lessonup(slide, lesson["bullets"], lesson["check"], positions["body"])

    # 9. teruggeven
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


