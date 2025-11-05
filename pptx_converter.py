# pptx_converter.py
import io
import math
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# layout-indexen in standaard PowerPoint
TITLE_LAYOUT = 0        # titel-slide
TITLE_ONLY_LAYOUT = 5   # alleen titel
BLANK_LAYOUT = 6        # blanco

# maximaal "echte" regels per dia
MAX_LINES_PER_SLIDE = 12

# ruwe schatting: ongeveer zoveel tekens past 1 powerpoint-regel
CHARS_PER_LINE = 75


def extract_images(doc):
    """Haal alle afbeeldingen uit het Word-document."""
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
    """Herken Word-opsommingen (stijl + numPr)."""
    name = (para.style.name or "").lower()
    if "list" in name or "lijst" in name or "opsom" in name:
        return True
    ppr = para._p.pPr
    return ppr is not None and ppr.numPr is not None


def has_bold(para):
    """True als er ergens vet staat in deze paragraaf."""
    return any(run.bold for run in para.runs)


def para_text_plain(para):
    """Alle runs samenvoegen tot platte tekst."""
    return "".join(run.text for run in para.runs if run.text).strip()


def estimate_line_count(text: str) -> int:
    """Schat hoeveel powerpoint-regels dit ongeveer worden."""
    if not text:
        return 0
    return max(1, math.ceil(len(text) / CHARS_PER_LINE))


def add_textbox(slide, text, top_offset_inch=1.0, est_lines=1):
    """
    Maak een tekstvak op de dia met automatische afbreking.
    Hoogte wordt groter naarmate er meer tekst is.
    """
    left = Inches(0.8)
    top = Inches(top_offset_inch)
    width = Inches(8.0)

    # basis 0.6 inch + 0.25 per extra lijn
    height_inch = 0.6 + (est_lines - 1) * 0.25
    # niet belachelijk groot
    height_inch = min(height_inch, 4.0)

    shape = slide.shapes.add_textbox(left, top, width, Inches(height_inch))
    tf = shape.text_frame
    tf.text = text
    tf.word_wrap = True
    tf.auto_size = False
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)

    # stijl: Arial 16
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)

    # hoeveel verticale ruimte we echt gebruikt hebben
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
    slide = prs.slides.add_slide(prs.slide_layouts[BLANK_LAYOUT])
    return slide


def add_image_slide(prs, img_bytes, title="Afbeelding"):
    """Voeg een dia toe met een (kleinere) gecentreerde afbeelding."""
    slide = prs.slides.add_slide(prs.slide_layouts[TITLE_ONLY_LAYOUT])
    slide.shapes.title.text = title

    # kleinere afbeelding in het midden
    slide_width = prs.slide_width
    img_width = Inches(4.5)
    left = (slide_width - img_width) // 2
    top = Inches(1.4)

    slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=img_width)
    return slide


def docx_to_pptx(file_like):
    doc = Document(file_like)
    prs = Presentation()

    all_images = extract_images(doc)
    img_ptr = 0

    # startslide
    current_slide = create_title_slide(prs)
    current_y = 2.0  # waar tekstvak start
    used_lines = 0   # geschatte lijnen op deze slide

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        is_heading = para.style.name.startswith("Heading")
        is_bold_title = has_bold(para) and not has_image and not is_word_list_paragraph(para)

        # 1. kop of bold → nieuwe slide met titel
        if is_heading or is_bold_title:
            current_slide = create_title_only_slide(prs, para_text_plain(para))
            current_y = 2.0
            used_lines = 0
            continue

        # 2. afbeelding → eigen slide
        if has_image:
            if img_ptr < len(all_images):
                _, img_bytes = all_images[img_ptr]
                img_ptr += 1
                current_slide = add_image_slide(prs, img_bytes)
            else:
                # geen bytes? maak dan tenminste een slide
                current_slide = create_title_only_slide(prs, "Afbeelding")
            current_y = 3.5
            used_lines = 0
            continue

        # 3. gewone tekst / opsomming
        if text:
            # maak opsommingen wat duidelijker
            if is_word_list_paragraph(para):
                display_text = "• " + text
            else:
                display_text = text

            # schat benodigde lijnen
            lines_needed = estimate_line_count(display_text)

            # als dit niet meer op deze slide past → nieuwe blanco slide
            if used_lines + lines_needed > MAX_LINES_PER_SLIDE:
                current_slide = create_blank_slide(prs)
                current_y = 1.0
                used_lines = 0

            # tekstvak toevoegen
            height_used = add_textbox(
                current_slide,
                display_text,
                top_offset_inch=current_y,
                est_lines=lines_needed,
            )
            # volgende vak iets lager
            current_y += height_used + 0.1
            used_lines += lines_needed

    # export
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
