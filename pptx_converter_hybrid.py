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


# =========================
# CONFIG
# =========================
BASE_TEMPLATE_NAME = "basis layout.pptx"  # /templates/basis layout.pptx
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


# =========================
# 1. DOCX → blokken (kop + tekst)
# =========================
def docx_to_blocks(doc: Document) -> list[dict]:
    """
    We halen de structuur uit Word:
    elke heading / vet / ALL CAPS = nieuwe dia
    alles eronder = tekst van die dia
    return: [{"title": "...", "body": "..."}, ...]
    """
    blocks = []
    current_title = None
    current_body: list[str] = []

    for para in doc.paragraphs:
        txt = "".join(r.text for r in para.runs if r.text).strip()
        if not txt:
            continue

        is_heading = (
            (para.style and para.style.name and para.style.name.lower().startswith("heading"))
            or any(r.bold for r in para.runs)
            or (len(txt) <= 50 and txt.upper() == txt)  # korte regel in CAPS
        )

        if is_heading:
            # oud blok afsluiten
            if current_title or current_body:
                blocks.append(
                    {
                        "title": current_title,
                        "body": "\n".join(current_body).strip(),
                    }
                )
            current_title = txt
            current_body = []
        else:
            current_body.append(txt)

    # laatste blok
    if current_title or current_body:
        blocks.append(
            {
                "title": current_title,
                "body": "\n".join(current_body).strip(),
            }
        )

    # echt niks?
    if not blocks:
        blocks = [{"title": "Lesonderdeel", "body": "(Geen duidelijke structuur gevonden in dit document.)"}]

    return blocks


# =========================
# 2. AI: alles in 1 keer → slides
# =========================
def ai_make_all_slides_from_blocks(blocks: list[dict]) -> list[dict]:
    """
    Stuurt ALLE blokken in één prompt naar OpenAI.
    Zo gebruiken we maar 1 API-call.
    Return-formaat: [{"title": "...", "text": [...], "check": "..."}, ...]
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Geen OPENAI_API_KEY ingesteld.")

    from openai import OpenAI, RateLimitError

    client = OpenAI(api_key=api_key)

    # blokken netjes in de prompt gieten
    parts = []
    for i, b in enumerate(blocks, start=1):
        parts.append(
            f"### Onderdeel {i}\nKop: {b.get('title') or ''}\nTekst:\n{b.get('body') or ''}\n"
        )
    joined = "\n\n".join(parts)

    prompt = f"""
Je krijgt hieronder meerdere onderdelen uit een les over installatietechniek (sanitair/riolering).
Maak hier dia's van voor een VMBO-les (basis/kader/GL).

Voor elk onderdeel:
- bedenk 1 korte, begrijpelijke titel (max 8 woorden)
- schrijf 2 of 3 korte, vertellende zinnen in de je-vorm
- schrijf 1 controlevraag die past bij de uitleg
- herhaal de titel NIET in de tekst
- gebruik eenvoudige woorden

Geef ALLEEN geldig JSON in dit formaat:

{{
  "slides": [
    {{
      "title": "…",
      "text": ["…", "…", "…"],
      "check": "…"
    }}
  ]
}}

Hier zijn de onderdelen uit het Word-document:

{joined}
"""

    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
        )
    except RateLimitError as e:
        # we vangen dit op in de caller
        raise RuntimeError("AI-limiet bereikt bij OpenAI.") from e
    except Exception as e:
        raise RuntimeError(f"AI kon geen dia's maken: {e}") from e

    try:
        data = json.loads(resp.choices[0].message.content)
    except Exception as e:
        raise RuntimeError("AI gaf geen geldig JSON terug.") from e

    slides = data.get("slides")
    if not slides:
        raise RuntimeError("AI gaf geen dia's terug.")
    return slides


# =========================
# 3. Fallback: als AI faalt → zelf dia's maken
# =========================
def fallback_slides_from_blocks(blocks: list[dict]) -> list[dict]:
    """
    Simpele backup: gebruik kop uit Word + 2-3 zinnetjes uit body + vraag.
    Dit is minder mooi dan AI, maar de gebruiker krijgt wel iets.
    """
    slides = []
    for b in blocks:
        title_raw = (b.get("title") or "Lesonderdeel").strip()
        body_raw = (b.get("body") or "").strip()

        # kop normaliseren
        title = title_raw.capitalize()

        # zinnen uit body halen
        sentences = re.split(r"[.!?]\s+", body_raw)
        sentences = [s.strip(" .!?") for s in sentences if s.strip()]

        text_lines = []
        for s in sentences[:3]:
            # klein beetje leerling-taal
            s = s.replace("Men ", "Je ").replace("men ", "je ")
            if not s.lower().startswith(("je ", "dit ", "zo ", "dan ", "als ")):
                s = "Je " + s[0].lower() + s[1:]
            text_lines.append(s)

        if not text_lines:
            text_lines = [
                "Je leert hier hoe je dit onderdeel goed uitvoert.",
                "Zo kan water en lucht goed weg.",
                "Dan krijg je geen stank.",
            ]

        lower = body_raw.lower()
        if "niet" in lower or "mag" in lower or "nooit" in lower:
            check = "Waarom mag je dit niet zo doen?"
        elif "leiding" in lower or "afvoer" in lower:
            check = "Wat gebeurt er als je dit verkeerd aansluit?"
        else:
            check = "Kun je uitleggen waarom je dit zo doet?"

        slides.append(
            {
                "title": title,
                "text": text_lines,
                "check": check,
            }
        )

    return slides


# =========================
# 4. PPTX helpers
# =========================
def get_logo_bytes():
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()
    return None


def add_logo(slide, logo_bytes):
    if not logo_bytes:
        return
    slide.shapes.add_picture(io.BytesIO(logo_bytes), Inches(7.5), Inches(0.2), width=Inches(1.5))


def get_positions_from_first_slide(slide):
    text_shapes = [s for s in slide.shapes if hasattr(s, "text") and s.text and s.text.strip()]
    if len(text_shapes) >= 2:
        t, b = text_shapes[0], text_shapes[1]
        return {
            "title": {"left": t.left, "top": t.top, "width": t.width, "height": t.height},
            "body": {"left": b.left, "top": b.top, "width": b.width, "height": b.height},
        }
    # fallback-posities
    return {
        "title": {"left": Inches(0.8), "top": Inches(0.8), "width": Inches(9), "height": Inches(1)},
        "body": {"left": Inches(0.8), "top": Inches(3), "width": Inches(10), "height": Inches(3)},
    }


def duplicate_slide_clean(prs: Presentation, slide_index: int):
    src = prs.slides[slide_index]
    dest = prs.slides.add_slide(prs.slide_layouts[0])

    for shp in src.shapes:
        # gelinkte plaatjes niet kopiëren
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
            continue
        new_el = deepcopy(shp.element)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # alle tekst wissen
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


def place_text_and_question(slide, lines: list[str], check: str, pos: dict):
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tf = box.text_frame
    tf.word_wrap = True

    first = True
    for line in lines:
        if not line:
            continue
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

    # algemene stijl
    for p in tf.paragraphs:
        for r in p.runs:
            if not r.font.size:
                r.font.name = "Arial"
                r.font.size = Pt(16)
                r.font.color.rgb = RGBColor(0, 0, 0)


# =========================
# 5. MAIN: DOCX → PPTX
# =========================
def docx_to_pptx_hybrid(file_like):
    # 1. template inladen
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # 2. docx lezen
    doc = Document(file_like)
    blocks = docx_to_blocks(doc)

    # 3. probeer ALLES in 1x naar AI
    try:
        slides_data = ai_make_all_slides_from_blocks(blocks)
    except Exception:
        # als AI faalt → toch iets maken
        slides_data = fallback_slides_from_blocks(blocks)

    # 4. logo en eerste dia
    logo_bytes = get_logo_bytes()
    if not prs.slides:
        prs.slides.add_slide(prs.slide_layouts[0])
    first_slide = prs.slides[0]

    positions = get_positions_from_first_slide(first_slide)

    # eerste dia leegmaken
    for shp in first_slide.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()
    if logo_bytes:
        add_logo(first_slide, logo_bytes)

    # 5. eerste slide vullen
    first = slides_data[0]
    place_title(first_slide, first["title"], positions["title"])
    place_text_and_question(first_slide, first.get("text", []), first.get("check", ""), positions["body"])

    # 6. overige slides
    for sd in slides_data[1:]:
        slide = duplicate_slide_clean(prs, 0)
        if logo_bytes:
            add_logo(slide, logo_bytes)
        place_title(slide, sd["title"], positions["title"])
        place_text_and_question(slide, sd.get("text", []), sd.get("check", ""), positions["body"])

    # 7. terug als bytes
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

