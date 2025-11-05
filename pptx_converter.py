import io
import math
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# standaard layouts
TITLE_LAYOUT = 0
TITLE_ONLY_LAYOUT = 5
BLANK_LAYOUT = 6

MAX_LINES_PER_SLIDE = 12
CHARS_PER_LINE = 75
MAX_BOTTOM_INCH = 6.6  # onderrand waarop we stoppen met schrijven


# ───────── hulpfuncties docx ─────────
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


# ───────── hulpfuncties pptx ─────────
def add_textbox(slide, text, top_inch=1.0, est_lines=1):
    """los tekstblok (voor gewone alinea's)"""
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
    """Afbeelding op huidige slide met vaste positie/breedte."""
    left = Inches(1.0)
    top = Inches(top_inch)
    width = Inches(4.5)
    slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=width)
    return 3.0  # ongeveer gebruikte hoogte


# ───────── hoofdconverter ─────────
def docx_to_pptx(file_like):
    doc = Document(file_like)
    prs = Presentation()

    all_images = extract_images(doc)
    img_ptr = 0

    # startslide
    current_slide = create_title_slide(prs)
    current_y = 2.0
    used_lines = 0

    # extra state voor een lopende opsomming
    current_list_tf = None   # text_frame van de lijst
    current_list_top = 0.0   # waar die lijst begint
    current_list_lines = 0   # hoeveel regels in deze lijst

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        is_heading = para.style.name.startswith("Heading")
        is_bold_title = has_bold(para) and not has_image and not is_word_list_paragraph(para)
        is_list = is_word_list_paragraph(para)

        # als we van soort wisselen (lijst -> geen lijst), lijst afsluiten
        if not is_list:
            current_list_tf = None
            current_list_lines = 0

        # 1. kop of vet → nieuwe dia met titel
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
                # past op huidige dia?
                if current_y + 3.0 <= MAX_BOTTOM_INCH:
                    h_used = add_inline_image(current_slide, img_bytes, current_y)
                    current_y += h_used + 0.2
                else:
                    # nieuwe blanco dia
                    current_slide = create_blank_slide(prs)
                    h_used = add_inline_image(current_slide, img_bytes, 1.0)
                    current_y = 1.0 + h_used + 0.2
                    used_lines = 0
            continue

        # 3. opsomming → in één blok
        if is_list and text:
            # schat regels
            lines_needed = estimate_line_count(text)

            # als deze lijst niet meer past → nieuwe blanco slide + nieuwe lijst starten
            if (
                used_lines + lines_needed > MAX_LINES_PER_SLIDE
                or current_y + 0.6 > MAX_BOTTOM_INCH
                or current_list_tf is None and used_lines >= MAX_LINES_PER_SLIDE
            ):
                current_slide = create_blank_slide(prs)
                current_y = 1.0
                used_lines = 0
                current_list_tf = None
                current_list_lines = 0

            if current_list_tf is None:
                # nieuw lijst-tekstvak maken
                left = Inches(0.8)
                top = Inches(current_y)
                width = Inches(8.0)
                height = Inches(4.0)  # ruim genoeg; we vullen hem met paragrafen
                shape = current_slide.shapes.add_textbox(left, top, width, height)
                tf = shape.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                p.text = text
                p.level = 0
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(16)
                current_list_tf = tf
                current_list_top = current_y
                current_list_lines = lines_needed
            else:
                p = current_list_tf.add_paragraph()
                p.text = text
                p.level = 0
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(16)
                current_list_lines += lines_needed

            used_lines += lines_needed
            # zet current_y net onder de lijst (0.35 per bullet)
            current_y = current_list_top + 0.35 * current_list_lines + 0.3
            continue

        # 4. gewone tekst
        if text:
            lines_needed = estimate_line_count(text)
            # past dit nog?
            if (
                used_lines + lines_needed > MAX_LINES_PER_SLIDE
                or current_y + 0.6 > MAX_BOTTOM_INCH
            ):
                current_slide = create_blank_slide(prs)
                current_y = 1.0
                used_lines = 0

            h_used = add_textbox(current_slide, text, top_inch=current_y, est_lines=lines_needed)
            current_y += h_used + 0.15
            used_lines += lines_needed

    # klaar → naar bytes
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
