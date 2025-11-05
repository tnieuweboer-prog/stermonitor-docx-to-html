import io
import math
import os
import json
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement

TITLE_LAYOUT = 0
TITLE_ONLY_LAYOUT = 5
BLANK_LAYOUT = 6
MAX_LINES_PER_SLIDE = 12
CHARS_PER_LINE = 75
MAX_BOTTOM_INCH = 6.6


# ----------- AI helper -----------
def summarize_with_ai(text: str, max_bullets: int = 0) -> str | list:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        # lokale fallback
        words = text.split()
        if max_bullets:
            parts = [p.strip() for p in text.replace("•", "\n").split("\n") if p.strip()]
            return parts[:max_bullets] or ["Kernpunt uit de tekst."]
        short = " ".join(words[:40])
        if len(words) > 40:
            short += "..."
        return short

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        if max_bullets:
            prompt = f"""
Maak van deze tekst maximaal {max_bullets} korte bullets (mbo-niveau, 1 regel per bullet).
Alleen de kern.

Tekst:
{text}

Geef JSON: {{"bullets": ["...","..."]}}
"""
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
            )
            data = json.loads(resp.choices[0].message.content)
            return data.get("bullets") or ["Kernpunt uit de tekst."]
        else:
            prompt = f"""
Vat deze les-tekst samen in 1 korte alinea voor een PowerPoint-dia.
Doelgroep: mbo, installatietechniek.
Max 40 woorden.

Tekst:
{text}
"""
            resp = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
            )
            return resp.choices[0].message.content.strip()

    except Exception:
        words = text.split()
        if max_bullets:
            parts = [p.strip() for p in text.replace("•", "\n").split("\n") if p.strip()]
            return parts[:max_bullets] or ["Kernpunt uit de tekst."]
        short = " ".join(words[:40])
        if len(words) > 40:
            short += "..."
        return short


# ----------- DOCX helpers -----------
def extract_images(doc):
    imgs = []
    for rel in doc.part.rels.values():
        if rel.reltype == RT.IMAGE:
            imgs.append((rel.target_part.partname, rel.target_part.blob))
    return imgs


def is_word_list_paragraph(para):
    name = (para.style.name or "").lower()
    if "list" in name or "lijst" in name or "opsom" in name:
        return True
    ppr = para._p.pPr
    return ppr is not None and ppr.numPr is not None


def has_bold(para):
    return any(run.bold for run in para.runs)


def para_text_plain(para):
    return "".join(run.text for run in para.runs if run.text).strip()


def estimate_line_count(text: str) -> int:
    if not text:
        return 0
    return max(1, math.ceil(len(text) / CHARS_PER_LINE))


# ----------- PPTX helpers -----------
def make_bullet(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith("buNone"):
            pPr.remove(child)
    pPr.set("marL", "288000")
    pPr.set("indent", "-144000")
    buChar = OxmlElement("a:buChar")
    buChar.set("char", "•")
    pPr.append(buChar)


def add_textbox(slide, text, top_inch=1.0, est_lines=1):
    left = Inches(0.8)
    top = Inches(top_inch)
    width = Inches(8.0)
    height_inch = 0.6 + (est_lines - 1) * 0.25
    shape = slide.shapes.add_textbox(left, top, width, Inches(height_inch))
    tf = shape.text_frame
    tf.text = text
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)
    return height_inch


def create_title_only_slide(prs, title_text):
    slide = prs.slides.add_slide(prs.slide_layouts[TITLE_ONLY_LAYOUT])
    slide.shapes.title.text = title_text
    return slide


def create_blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[BLANK_LAYOUT])


def add_inline_image(slide, img_bytes, top_inch):
    left = Inches(1.0)
    top = Inches(top_inch)
    width = Inches(4.5)
    slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=width)
    return 3.0


# ----------- MAIN -----------
def docx_to_pptx_hybrid(file_like):
    doc = Document(file_like)
    prs = Presentation()
    all_images = extract_images(doc)
    img_ptr = 0

    current_slide = prs.slides.add_slide(prs.slide_layouts[TITLE_ONLY_LAYOUT])
    current_slide.shapes.title.text = "Les gegenereerd met AI"
    current_y = 2.0
    used_lines = 0

    for para in doc.paragraphs:
        raw_text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)
        is_heading = para.style.name.startswith("Heading")
        is_bold_title = has_bold(para)
        is_list = is_word_list_paragraph(para)

        if is_heading or is_bold_title:
            current_slide = create_title_only_slide(prs, para_text_plain(para))
            current_y = 2.0
            used_lines = 0
            continue

        if has_image:
            if img_ptr < len(all_images):
                _, img_bytes = all_images[img_ptr]
                img_ptr += 1
                add_inline_image(current_slide, img_bytes, current_y)
                current_y += 3.2
            continue

        if is_list and raw_text:
            bullets = summarize_with_ai(raw_text, max_bullets=3)
            left = Inches(0.8)
            top = Inches(current_y)
            width = Inches(7.0)
            shape = current_slide.shapes.add_textbox(left, top, width, Inches(3.0))
            tf = shape.text_frame
            tf.word_wrap = True
            for i, b in enumerate(bullets):
                p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
                p.text = b
                make_bullet(p)
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(16)
            current_y += 0.3 * len(bullets)
            continue

        if raw_text:
            short_text = summarize_with_ai(raw_text, max_bullets=0)
            h = add_textbox(current_slide, short_text, top_inch=current_y)
            current_y += h + 0.3

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


