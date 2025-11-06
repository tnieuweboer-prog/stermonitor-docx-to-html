import io
import os
from copy import deepcopy
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

# =========================
# configuratie
# =========================
BASE_TEMPLATE_NAME = "basis layout.pptx"   # /templates/basis layout.pptx
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


# =========================
# kleine hulpfuncties
# =========================
def summarize_text(text: str, max_chars: int = 400) -> str:
    text = text.strip()
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rsplit(" ", 1)[0] + "..."


def has_bold(para) -> bool:
    return any(run.bold for run in para.runs)


def is_all_caps_heading(text: str) -> bool:
    text = text.strip()
    if not text:
        return False
    # jouw document: korte koppen in CAPS
    return len(text) <= 40 and text.upper() == text


def para_text_plain(para) -> str:
    return "".join(r.text for r in para.runs if r.text).strip()


# =========================
# 1. DOCX → blokken
# =========================
def docx_to_blocks(doc: Document):
    """
    Maak een lijst van blokken:
    [{ "title": ..., "body": [..,..] }, ...]
    Kop = heading / vet / ALL CAPS
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


# =========================
# 2. LessonUp-stijl herschrijven
# =========================
def rewrite_for_lessonup(title: str, body_lines: list[str], niveau: str = "vmbo") -> dict:
    """
    Maak van (title, lange tekst) een LessonUp-dia:
    {
      "title": "...",
      "bullets": ["...", "..."],
      "check": "..."
    }
    """
    raw_text = " ".join(body_lines).strip()
    raw_text = summarize_text(raw_text, 500)

    bullets = []

    # simpele heuristiek op basis van je voorbeeld
    # 1e bullet: wat het is
    if any(w in title.lower() for w in ["kabel", "kabels"]):
        bullets.append(f"{title.title()} wordt gebruikt in elektrische installaties.")
    else:
        bullets.append(raw_text if len(raw_text) < 120 else summarize_text(raw_text, 120))

    # 2e bullet: eigenschap
    if "grond" in raw_text.lower():
        bullets.append("Heeft extra bescherming voor in de grond.")
    elif "ymvk" in raw_text.lower():
        bullets.append("Mag je in goten leggen en bundelen.")
    elif "xmvk" in raw_text.lower():
        bullets.append("Geschikt voor lichtpunten en stopcontacten.")

    # 3e bullet: waarom
    bullets.append("Is veiliger dan losse draden.")

    # korte tekst uit body als extra bullet
    if len(body_lines) > 0:
        extra = summarize_text(body_lines[0], 90)
        if extra not in bullets:
            bullets.append(extra)

    # max 4 bullets
    bullets = bullets[:4]

    # checkvraag bouwen
    check = f"Waarom gebruik je hier {title.lower()}?" if "kabel" in title.lower() else "Waarom heb je dit nodig in een installatie?"

    return {
        "title": title.title(),
        "bullets": bullets,
        "check": check,
    }


# =========================
# 3. logo
# =========================
def get_logo_bytes():
    if os.path.exists(LOCAL_LOGO_PATH):
        with open(LOCAL_LOGO_PATH, "rb") as f:
            return f.read()
    return None


def add_logo_to_slide(slide, logo_bytes):
    if not logo_bytes:
        return
    left = Inches(9.0 - 1.5)
    top = Inches(0.2)
    width = Inches(1.5)
    slide.shapes.add_picture(io.BytesIO(logo_bytes), left, top, width=width)


# =========================
# 4. posities van dia 1 overnemen
# =========================
def get_title_and_body_positions_from_slide(slide):
    """
    Kijk op jouw template-dia waar jij titel en tekst hebt gezet.
    We nemen de eerste 2 tekstvormen.
    """
    text_shapes = [s for s in slide.shapes if hasattr(s, "text") and s.text and s.text.strip()]
    if len(text_shapes) >= 2:
        t, b = text_shapes[0], text_shapes[1]
        return {
            "title": {"left": t.left, "top": t.top, "width": t.width, "height": t.height},
            "body": {"left": b.left, "top": b.top, "width": b.width, "height": b.height},
        }
    # fallback als je template leeg is
    return {
        "title": {"left": Inches(0.6), "top": Inches(0.8), "width": Inches(9), "height": Inches(0.8)},
        "body": {"left": Inches(0.6), "top": Inches(3.4), "width": Inches(11.6), "height": Inches(1.6)},
    }


# =========================
# 5. dia klonen en leegmaken
# =========================
def duplicate_slide_clean(prs: Presentation, slide_index: int):
    src = prs.slides[slide_index]
    dest = prs.slides.add_slide(prs.slide_layouts[0])

    # kopieer alle niet-afbeelding-shapes
    for shp in src.shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
            continue
        new_el = deepcopy(shp.element)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # alle tekst leeg
    for shp in dest.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()

    return dest


# =========================
# 6. tekst plaatsen volgens posities
# =========================
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

    # eerste bullet
    first = True
    for b in bullets:
        p = tf.add_paragraph() if not first else tf.paragraphs[0]
        p.text = b
        p.level = 0
        first = False

    # checkvraag
    if check:
        p = tf.add_paragraph()
        p.text = f"Check: {check}"
        p.level = 0
        # evt. iets vetter
        for r in p.runs:
            r.font.bold = True


# =========================
# MAIN
# =========================
def docx_to_pptx_hybrid(file_like):
    # 1. template
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # 2. docx inlezen
    doc = Document(file_like)
    blocks = docx_to_blocks(doc)

    # 3. logo
    logo_bytes = get_logo_bytes()

    # 4. zorgen dat er 1 dia is
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])
    first_slide = prs.slides[0]

    # 5. posities uit dia 1 halen
    positions = get_title_and_body_positions_from_slide(first_slide)

    # 6. eerste dia leegmaken
    for shp in first_slide.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()
    if logo_bytes:
        add_logo_to_slide(first_slide, logo_bytes)

    # 7. eerste blok invullen
    if blocks:
        lesson = rewrite_for_lessonup(blocks[0]["title"], blocks[0]["body"])
        place_title(first_slide, lesson["title"], positions["title"])
        place_lessonup_body(first_slide, lesson["bullets"], lesson["check"], positions["body"])
    else:
        place_title(first_slide, "Les gegenereerd met AI", positions["title"])

    # 8. overige blokken → nieuwe dia’s
    for block in blocks[1:]:
        slide = duplicate_slide_clean(prs, 0)
        if logo_bytes:
            add_logo_to_slide(slide, logo_bytes)
        lesson = rewrite_for_lessonup(block["title"], block["body"])
        place_title(slide, lesson["title"], positions["title"])
        place_lessonup_body(slide, lesson["bullets"], lesson["check"], positions["body"])

    # 9. teruggeven
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
