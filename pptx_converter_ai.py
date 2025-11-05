import io
import os
import json
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

from openai import OpenAI  # zorg dat 'openai' in requirements.txt staat

TITLE_LAYOUT = 0
TITLE_ONLY_LAYOUT = 5

# haal key uit env (in Streamlit: st.secrets → wordt ook env)
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None


def extract_images(doc):
    """Alle afbeeldingen uit het docx in volgorde."""
    imgs = []
    for rel in doc.part.rels.values():
        if rel.reltype == RT.IMAGE:
            imgs.append(rel.target_part.blob)
    return imgs


def call_ai_for_slide(text: str) -> dict:
    """
    Vraag OpenAI om van 1 alinea 1 dia te maken.
    Als er geen API-key is of er gaat iets mis: gebruik fallback.
    """
    if not client:
        # fallback
        return {
            "title": text[:50] + "..." if len(text) > 50 else text,
            "bullets": [],
            "image_hint": ""
        }

    prompt = f"""
Je bent een docent elektrotechniek en je maakt een PowerPoint-dia voor mbo-studenten.
Maak van de volgende tekst precies 1 dia:

- korte titel
- maximaal 4 bullets
- simpele taal
- geen vervolg-dia
- alleen de kern

Tekst:
{text}

Geef exact deze JSON terug:
{{
  "title": "...",
  "bullets": ["...", "..."],
  "image_hint": "..."
}}
"""
    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
        )
        data = json.loads(resp.choices[0].message.content)
        if "title" not in data:
            data["title"] = text[:40]
        if "bullets" not in data:
            data["bullets"] = []
        return data
    except Exception:
        # bij error: simpele dia
        return {
            "title": text[:50] + "..." if len(text) > 50 else text,
            "bullets": [],
            "image_hint": ""
        }


def add_ai_slide(prs: Presentation, ai_obj: dict, image_bytes: bytes | None = None):
    """
    Maak de dia op basis van AI-output.
    Titel boven, bullets links, plaatje rechts (als er is).
    """
    slide = prs.slides.add_slide(prs.slide_layouts[TITLE_ONLY_LAYOUT])

    title = ai_obj.get("title") or "Lesonderdeel"
    bullets = ai_obj.get("bullets") or ["Belangrijk punt uit de tekst."]

    slide.shapes.title.text = title

    # bullets links
    left = Inches(0.8)
    top = Inches(1.8)
    width = Inches(5.5) if image_bytes else Inches(7.5)
    height = Inches(4.0)
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.text = ""

    for i, b in enumerate(bullets):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = b
        p.level = 0
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
    DOCX → AI-dia's.
    1 paragraaf = 1 dia
    afbeelding uit die paragraaf komt op dezelfde dia.
    """
    doc = Document(file_like)
    prs = Presentation()

    all_imgs = extract_images(doc)
    img_idx = 0

    # optionele openingsdia
    first = prs.slides.add_slide(prs.slide_layouts[TITLE_LAYOUT])
    first.shapes.title.text = "Les gegenereerd uit Word"
    if len(first.placeholders) > 1:
        first.placeholders[1].text = "AI-samenvattingen"

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in r._element.xml for r in para.runs)

        if not text and not has_image:
            continue

        ai_obj = call_ai_for_slide(text if text else "Afbeelding bij uitleg")

        img_bytes = None
        if has_image and img_idx < len(all_imgs):
            img_bytes = all_imgs[img_idx]
            img_idx += 1

        add_ai_slide(prs, ai_obj, image_bytes=img_bytes)

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio

