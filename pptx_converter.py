import io
import os
import json
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# OpenAI client
from openai import OpenAI

# als er geen key is, zetten we client op None en doen we fallback
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

TITLE_LAYOUT = 0        # title slide
TITLE_ONLY_LAYOUT = 5   # title-only slide


def extract_images(doc):
    """Alle afbeeldingen uit het docx in volgorde."""
    imgs = []
    for rel in doc.part.rels.values():
        if rel.reltype == RT.IMAGE:
            imgs.append(rel.target_part.blob)
    return imgs


def call_ai_for_slide(paragraph_text: str) -> dict:
    """
    Vraagt OpenAI om van een stukje les-tekst 1 dia te maken.
    Geeft een dict terug met: title, bullets, image_hint
    """
    if not client:
        # geen key -> simpele fallback
        return {
            "title": paragraph_text[:50] + "..." if len(paragraph_text) > 50 else paragraph_text,
            "bullets": [],
            "image_hint": ""
        }

    prompt = f"""
Je bent een docent elektrotechniek en maakt een PowerPoint-dia voor leerlingen (mbo).
Maak van de volgende tekst precies één dia.

Regels:
- Houd de titel kort.
- Maak maximaal 4 bullets.
- Gebruik korte, actieve zinnen.
- Laat herhaling weg.
- Als het een definitie/uitleg is: zet de kernpunten in bullets.
- Antwoord als JSON.

Tekst:
{paragraph_text}

Verplichte JSON:
{{
  "title": "...",
  "bullets": ["...", "..."],
  "image_hint": "kort zinnetje over welk plaatje erbij past (mag leeg)"
}}
"""
    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
        )
        data = json.loads(resp.choices[0].message.content)
        # minimale sanity
        if "title" not in data:
            data["title"] = paragraph_text[:40]
        if "bullets" not in data:
            data["bullets"] = []
        return data
    except Exception:
        # bij fout: heel eenvoudige slide
        return {
            "title": paragraph_text[:50] + "..." if len(paragraph_text) > 50 else paragraph_text,
            "bullets": [],
            "image_hint": ""
        }


def add_ai_slide(prs: Presentation, ai_obj: dict, image_bytes: bytes | None = None):
    """
    Maakt echt de dia in PowerPoint vanuit de AI-output.
    Layout:
    - Titel boven
    - Bullets links
    - Afbeelding rechts (als aanwezig)
    """
    slide = prs.slides.add_slide(prs.slide_layouts[TITLE_ONLY_LAYOUT])
    title = ai_obj.get("title") or "Lesonderdeel"
    bullets = ai_obj.get("bullets") or []
    slide.shapes.title.text = title

    # bullets links
    left = Inches(0.8)
    top = Inches(1.8)
    width = Inches(5.5) if image_bytes else Inches(7.5)
    height = Inches(4.0)
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.text = ""  # leeg beginnen

    if not bullets:
        bullets = ["Belangrijk punt uit de tekst."]

    for i, b in enumerate(bullets):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = b
        p.level = 0
        # styling
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)

    # afbeelding rechts
    if image_bytes:
        img_left = Inches(6.0)
        img_top = Inches(2.0)
        slide.shapes.add_picture(io.BytesIO(image_bytes), img_left, img_top, width=Inches(2.5))

    return slide


def docx_to_pptx_ai(file_like):
    """
    Hoofdfunctie: DOCX -> AI-dia's -> PPTX bytes
    """
    doc = Document(file_like)
    prs = Presentation()

    # verzamel ALLE afbeeldingen in het document (in volgorde)
    all_imgs = extract_images(doc)
    img_index = 0

    # optioneel: eerste dia
    first = prs.slides.add_slide(prs.slide_layouts[TITLE_LAYOUT])
    first.shapes.title.text = "Les gegenereerd uit Word"
    if len(first.placeholders) > 1:
        first.placeholders[1].text = "Automatisch samengevat"

    # we lopen gewoon alle paragrafen af
    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        # niks? skip
        if not text and not has_image:
            continue

        # AI vragen om 1 slide voor deze paragraaf
        ai_obj = call_ai_for_slide(text if text else "Afbeelding bij les")

        # als deze paragraaf ook een plaatje had, koppelen we dat mee
        img_bytes = None
        if has_image and img_index < len(all_imgs):
            img_bytes = all_imgs[img_index]
            img_index += 1

        add_ai_slide(prs, ai_obj, image_bytes=img_bytes)

    # terug naar bytes
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio
