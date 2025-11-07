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
BASE_TEMPLATE_NAME = "basis layout.pptx"   # verwacht in /templates
LOCAL_LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")


# =========================================================
# 1. AI-helper: maak VMBO-lesblok
# =========================================================
def ai_vmbo_block_from_text(raw_text):
    """
    Maakt van willekeurige technische tekst een vmbo-lesblok:
    {
      "title": "pakkende titel",
      "text": ["zin 1", "zin 2", "zin 3"],
      "check": "vraag ..."
    }
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return fallback_vmbo_block(raw_text)

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        prompt = f"""
Je bent docent installatietechniek op een vmbo-school (basis/kader/GL).
Je krijgt hieronder een ruwe, technische tekst.

Zet dit om naar lesmateriaal voor 1 dia.

Regels:
- Schrijf 1 korte, pakkende titel (max 6 woorden) die duidelijk is voor vmbo-leerlingen.
- Schrijf 3 korte, vertellende zinnen in de je-vorm.
- Schrijf 1 check-vraag die past bij de uitleg.
- Gebruik eenvoudige woorden.
- Geef ALLEEN geldige JSON zoals hieronder.

Voorbeeld:
{{
  "title": "Afvoer rustig aansluiten",
  "text": [
    "Je maakt de bochten niet te scherp.",
    "Zo kan het water en de lucht er goed door.",
    "Dan krijg je minder kans op verstopping."
  ],
  "check": "Waarom is een rustige bocht beter?"
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
        # als er iets mis gaat met de API, gebruik fallback
        return fallback_vmbo_block(raw_text)


def fallback_vmbo_block(raw_text):
    """Fallback als er geen AI is."""
    raw_text = (raw_text or "").strip()
    if not raw_text:
        return {
            "title": "Lesonderdeel",
            "text": [
                "Je leert hier een stap uit de installatie.",
                "Lees mee en kijk wat je moet doen.",
                "Vraag om hulp als je het niet snapt."
            ],
            "check": "Waarom doe je deze stap zo?"
        }

    # titel: eerste zin, max 6 woorden
    first_sentence = raw_text.split(".")[0].strip()
    words = first_sentence.split()
    if len(words) > 6:
        first_sentence = " ".join(words[:6])
    title = first_sentence or "Lesonderdeel"

    return {
        "title": title,
        "text": [
            "Je gaat dit onderdeel leren.",
            "Kijk goed naar de volgorde.",
            "Zo voorkom je fouten."
        ],
        "check": "Wat is hier het belangrijkste?"
    }


# =========================================================
# 2. DOCX analyseren → tekstblokken
# =========================================================
def plain_para_text(para):
    return "".join(r.text for r in para.runs if r.text).strip()


def para_is_heading_like(para):
    txt = plain_para_text(para)
    if not txt:
        return False
    # echte heading
    if para.style and para.style.name and para.style.name.lower().startswith("heading"):
        return True
    # vet?
    if any(r.bold for r in para.runs):
        return True
    # ALL CAPS en kort
    if len(txt) <= 50 and txt.upper() == txt:
        return True
    return False


def docx_to_raw_blocks(doc):
    """
    Maak een lijst met tekstblokken uit een willekeurig Word-bestand.
    Elke 'heading-achtige' paragraaf start een nieuw blok.
    """
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
# 3. template helpers
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
    """
    probeer posities van titel en tekst uit jouw template-dia te halen
    """
    text_shapes = [s for s in slide.shapes if hasattr(s, "text") and s.text and s.text.strip()]
    if len(text_shapes) >= 2:
        t, b = text_shapes[0], text_shapes[1]
        return {
            "title": {
                "left": t.left,
                "top": t.top,
                "width": t.width,
                "height": t.height,
            },
            "body": {
                "left": b.left,
                "top": b.top,
                "width": b.width,
                "height": b.height,
            },
        }
    # fallback posities
    return {
        "title": {"left": Inches(0.6), "top": Inches(0.8), "width": Inches(9), "height": Inches(0.8)},
        "body": {"left": Inches(0.6), "top": Inches(3.4), "width": Inches(11.5), "height": Inches(2.2)},
    }


def duplicate_slide_clean(prs, slide_index):
    """
    Kloon een dia en maak alle tekst leeg.
    Gelinkte plaatjes uit template slaan we over.
    """
    source = prs.slides[slide_index]
    dest = prs.slides.add_slide(prs.slide_layouts[0])

    for shp in source.shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
            continue
        new_el = deepcopy(shp.element)
        dest.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # nu alle tekst leeg
    for shp in dest.shapes:
        if hasattr(shp, "text_frame"):
            shp.text_frame.clear()

    return dest


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
    """
    Zet 3 vertelzinnen onder elkaar en dan een lege regel en dan de vraag.
    """
    box = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tf = box.text_frame
    tf.word_wrap = True

    first = True
    for line in lines:
        line = line.strip()
        if not line:
            continue
        p = tf.add_paragraph() if not first else tf.paragraphs[0]
        p.text = line
        first = False

    # lege regel tussen tekst en vraag
    blank_p = tf.add_paragraph()
    blank_p.text = ""

    # check-vraag
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
# 4. MAIN
# =========================================================
def docx_to_pptx_hybrid(file_like):
    # template inladen
    base_dir = os.path.dirname(__file__)
    template_path = os.path.join(base_dir, "templates", BASE_TEMPLATE_NAME)
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # docx lezen
    doc = Document(file_like)
    raw_blocks = docx_to_raw_blocks(doc)

    # logo
    logo_bytes = get_logo_bytes()

    # zorg dat we iig 1 dia hebben
    if len(prs.slides) == 0:
        prs.slides.add_slide(prs.slide_layouts[0])
    first_slide = prs.slides[0]

    # posities
    positions = get_positions_from_first_slide(first_slide)

    # eerste dia leegmaken
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

    # output
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out



