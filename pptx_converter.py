# pptx_converter.py
import io
import math
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Standaard layouts uit PowerPoint
TITLE_LAYOUT = 0        # titel-slide
TITLE_ONLY_LAYOUT = 5   # alleen titel
BLANK_LAYOUT = 6        # blanco

# jouw regels
MAX_LINES_PER_SLIDE = 12
CHARS_PER_LINE = 75     # schatting voor 1 regel
MAX_BOTTOM_INCH = 6.6   # tot hier willen we op de dia schrijven


def extract_images(doc):
    """Haal alle afbeeldingen uit het Word-document in volgorde."""
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
    """Herken opsommingen uit Word."""
    name = (para.style.name or "").lower()
    if "list" in name or "lijst" in name or "opsom" in name:
        return True
    ppr = para._p.pPr
    return ppr is not None and ppr.numPr is not None


def has_bold(para):
    """True als er ergens vet staat."""
    return any(run.bold for run in para.runs)


def para_text_plain(para):
    return "".join(run.text for run in para.runs if run.text).strip()


def estimate_line_count(text: str) -> int:
    """Schat hoeveel echte regels dit in PowerPoint wordt."""
    if not text:
        return 0
    return max(1, math.ceil(len(text) / CHARS_PER_LINE))


def add_textbox(slide, text, top_inch=1.0, est_lines=1):
    """
    Voeg een tekstvak toe met automatische afbreking en Arial 16.
    Hoogte schaalt mee met de tekst.
    """
    left = Inches(0.8)
    top = Inches(top_inch)
    width = Inches(8.0)

    # basis 0.6 + 0.25 per extra regel
    height_inch = 0.6 + (est_lines - 1) * 0.25
    height_inch = min(height_inch, 4.0)  # niet absurd groot

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
    """
    Voeg een afbeelding toe op de huidige dia, op de meegegeven hoogte.
    Afbeelding wordt kleiner en gecentreerd.
    """
    img_width = Inches(4.5)
    slide_width = slide.part.slide_layout.part.package.presentation.slide_width
    left = (slide_width - img_width) // 2
    top = Inches(top_inch)
    slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=img_width)
    # ongeveer 3 inch hoog voor spacing
    return 3.0


def docx_to_pptx(file_like):
    doc = Document(file_like)
    prs = Presentation()

    all_images = extract_images(doc)
    img_ptr = 0

    # begin met 1 titel-dia
    current_slide = create_title_slide(prs)
    current_y = 2.0       # waar eerste tekst komt
    used_lines = 0        # geschatte regels op deze dia

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        is_heading = para.style.name.startswith("Heading")
        is_bold_title = has_bold(para) and not has_image and not is_word_list_paragraph(para)

        # 1. kop of bold → nieuwe dia met titel
        if is_heading or is_bold_title:
            current_slide = create_title_only_slide(prs, para_text_plain(para))
            current_y = 2.0
            used_lines = 0
            continue

        # 2. afbeelding → eerst proberen op huidige dia te zetten
        if has_image:
            if img_ptr < len(all_images):
                _, img_bytes = all_images[img_ptr]
                img_ptr += 1
                # past hij nog op deze dia?
                if current_y + 3.0 <= MAX_BOTTOM_INCH:
                    used_height = add_inline_image(current_slide, img_bytes, current_y)
                    current_y += used_height + 0.2
                else:
                    # nieuwe blanco dia alleen voor deze afbeelding
                    img_slide = create_blank_slide(prs)
                    add_inline_image(img_slide, img_bytes, top_inch=1.0)
                    current_slide = img_slide
                    current_y = 4.2
                    used_lines = 0
            continue

        # 3. normale tekst / opsomming
        if text:
            # opsomming krijgt een bullet
            if is_word_list_paragraph(para):
                display_text = "• " + text
            else:
                display_text = text

            # hoeveel regels kost dit?
            lines_needed = estimate_line_count(display_text)

            # als dit niet meer past → nieuwe blanco slide
            if used_lines + lines_needed > MAX_LINES_PER_SLIDE or current_y > MAX_BOTTOM_INCH:
                current_slide = create_blank_slide(prs)
                current_y = 1.0
                used_lines = 0

            used_height = add_textbox(
                current_slide,
                display_text,
                top_inch=current_y,
                est_lines=lines_needed,
            )
            current_y += used_height + 0.15
            used_lines += lines_needed

    # export naar bytes
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
