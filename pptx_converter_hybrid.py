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

# layouts
TITLE_LAYOUT = 0
TITLE_ONLY_LAYOUT = 5
BLANK_LAYOUT = 6

MAX_LINES_PER_SLIDE = 12
CHARS_PER_LINE = 75
MAX_BOTTOM_INCH = 6.6  # onderrand


# ---------- AI helper ----------
def summarize_with_ai(text: str, max_bullets: int = 0) -> str | list:
    """
    Probeer tekst korter te maken met OpenAI.
    - als max_bullets == 0 → geef 1 korte alinea terug
    - als max_bullets > 0 → geef een lijst bullets terug
    Bij fout of geen key → simpele lokale verkorting.
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        # lokale fallback
        words = text.split()
        if max_bullets:
            # simpele split in bullets
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
Maak van deze tekst maximaal {max_bullets} hele korte bullets (mbo-niveau, 1 regel per bullet).
Alleen de kern, geen inleiding.

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
Vat deze les-tekst samen in 1 korte alinea voor op een PowerPoint-dia.
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
        # fallback
        words = text.split()
        if max_bullets:
            parts = [p.strip() for p in text.replace("•", "\n").split("\n") if p.strip()]
            return parts[:max_bullets] or ["Kernpunt uit de tekst."]
        short = " ".join(words[:40])
        if len(words) > 40:
            short += "..."
        return short


# ---------- DOCX helpers ----------
def extract_images(doc):
    imgs = []
    idx = 1
    for rel in doc.part.rels.values():
        if rel.reltype == RT.IMAGE:
            blob = rel.target_part.blob
            ext = rel.target_part.partname.ext
            filename = f"image_{idx}.{ext}"
            imgs.append((filename, blob))
            idx += 1
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


# ---------- PPTX helpers ----------
def make_bullet(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith("buNone"):
            pPr.remove(child)

    # 8 mm
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
    height_inch = min(height_inch, 4.0)
    shape = slide.shapes.add_textbox(left, top, width, Inches(height_inch))
    tf = shape.text_frame
    tf.text = text
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)
    return height_inch


def create_title_slide(prs, title="Inhoud uit Word"):
    slide = prs.slides.add_slide(prs.slide_layouts[TITLE_LAYOUT])
    slide.shapes.title.text = title
    if len(slide.placeholders) > 1:
        slide.placeholders[1].text = "Geconverteerd voor LessonUp"
    return slide


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


# ---------- MAIN (hybride) ----------
def docx_to_pptx_hybrid(file_like):
    doc = Document(file_like)
    prs = Presentation()

    all_images = extract_images(doc)
    img_ptr = 0

    current_slide = create_title_slide(prs)
    current_y = 2.0
    used_lines = 0

    current_list_tf = None
    current_list_top = 0.0
    current_list_lines = 0

    for para in doc.paragraphs:
        raw_text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        is_heading = para.style.name.startswith("Heading")
        is_bold_title = has_bold(para) and not has_image and not is_word_list_paragraph(para)
        is_list = is_word_list_paragraph(para)

        # uit lijst lopen:
        if not is_list:
            current_list_tf = None
            current_list_lines = 0

        # 1. kop / vet → nieuwe dia
        if is_heading or is_bold_title:
            current_slide = create_title_only_slide(prs, para_text_plain(para))
            current_y = 2.0
            used_lines = 0
            current_list_tf = None
            continue

        # 2. afbeelding
        if has_image:
            if img_ptr < len(all_images):
                _, img_bytes = all_images[img_ptr]
                img_ptr += 1
                if current_y + 3.0 <= MAX_BOTTOM_INCH:
                    h_used = add_inline_image(current_slide, img_bytes, current_y)
                    current_y += h_used + 0.2
                else:
                    current_slide = create_blank_slide(prs)
                    h_used = add_inline_image(current_slide, img_bytes, 1.0)
                    current_y = 1.0 + h_used + 0.2
                    used_lines = 0
            continue

        # 3. lijst → eerst samenvatten naar korte bullets
        if is_list and raw_text:
            # vraag AI om max 3 bullets
            bullets = summarize_with_ai(raw_text, max_bullets=3)

            # we zetten dit nog steeds in 1 tekstvak
            for idx, b in enumerate(bullets):
                lines_needed = estimate_line_count(b)
                if (
                    used_lines + lines_needed > MAX_LINES_PER_SLIDE
                    or current_y + 0.6 > MAX_BOTTOM_INCH
                    or (current_list_tf is None and used_lines >= MAX_LINES_PER_SLIDE)
                ):
                    current_slide = create_blank_slide(prs)
                    current_y = 1.0
                    used_lines = 0
                    current_list_tf = None
                    current_list_lines = 0

                if current_list_tf is None:
                    left = Inches(0.8)
                    top = Inches(current_y)
                    width = Inches(7.0)
                    height = Inches(4.0)
                    shape = current_slide.shapes.add_textbox(left, top, width, height)
                    tf = shape.text_frame
                    tf.clear()
                    tf.word_wrap = True
                    tf.margin_left = Inches(0.1)
                    tf.margin_right = Inches(0.1)
                    p = tf.paragraphs[0]
                    p.text = b
                    make_bullet(p)
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(16)
                    current_list_tf = tf
                    current_list_top = current_y
                    current_list_lines = lines_needed
                else:
                    p = current_list_tf.add_paragraph()
                    p.text = b
                    make_bullet(p)
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(16)
                    current_list_lines += lines_needed

                used_lines += lines_needed
                current_y = current_list_top + 0.35 * current_list_lines + 0.3

            continue

        # 4. gewone tekst → AI kort maken
        if raw_text:
            short_text = summarize_with_ai(raw_text, max_bullets=0)
            lines_needed = estimate_line_count(short_text)
            if (
                used_lines + lines_needed > MAX_LINES_PER_SLIDE
                or current_y + 0.6 > MAX_BOTTOM_INCH
            ):
                current_slide = create_blank_slide(prs)
                current_y = 1.0
                used_lines = 0

            h_used = add_textbox(current_slide, short_text, top_inch=current_y, est_lines=lines_needed)
            current_y += h_used + 0.15
            used_lines += lines_needed

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

