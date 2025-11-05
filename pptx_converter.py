import io
import math
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

TITLE_LAYOUT = 0        # titel-slide
TITLE_ONLY_LAYOUT = 5   # alleen titel
BLANK_LAYOUT = 6        # blanco

MAX_LINES_PER_SLIDE = 12
CHARS_PER_LINE = 75
MAX_BOTTOM_INCH = 6.6   # onderrand waar we willen stoppen


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
    """
    Afbeelding op huidige slide zetten, vaste positie.
    Geen slide_width gebruiken -> geen AttributeError.
    """
    left = Inches(1.0)          # 1 inch van links
    top = Inches(top_inch)
    width = Inches(4.5)         # wat smaller
    slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=width)
    # hoogte van deze afbeelding op dia
    return 3.0


def docx_to_pptx(file_like):
    doc = Document(file_like)
    prs = Presentation()

    all_images = extract_images(doc)
    img_ptr = 0

    current_slide = create_title_slide(prs)
    current_y = 2.0
    used_lines = 0

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        is_heading = para.style.name.startswith("Heading")
        is_bold_title = has_bold(para) and not has_image and not is_word_list_paragraph(para)

        # 1. kop of vet → nieuwe dia met titel
        if is_heading or is_bold_title:
            current_slide = create_title_only_slide(prs, para_text_plain(para))
            current_y = 2.0
            used_lines = 0
            continue

        # 2. afbeelding → eerst proberen op huidige dia
        if has_image:
            if img_ptr < len(all_images):
                _, img_bytes = all_images[img_ptr]
                img_ptr += 1
                # past 'ie nog?
                if current_y + 3.0 <= MAX_BOTTOM_INCH:
                    used_h = add_inline_image(current_slide, img_bytes, current_y)
                    current_y += used_h + 0.2
                else:
                    # nieuwe blanco dia met afbeelding
                    new_slide = create_blank_slide(prs)
                    add_inline_image(new_slide, img_bytes, 1.0)
                    current_slide = new_slide
                    current_y = 4.2
                    used_lines = 0
            continue

        # 3. gewone tekst / opsomming
        if text:
            if is_word_list_paragraph(para):
                display_text = "• " + text
            else:
                display_text = text

            lines_needed = estimate_line_count(display_text)

            # past dit nog op deze dia?
            if (
                used_lines + lines_needed > MAX_LINES_PER_SLIDE
                or current_y + 0.6 > MAX_BOTTOM_INCH
            ):
                current_slide = create_blank_slide(prs)
                current_y = 1.0
                used_lines = 0

            used_h = add_textbox(
                current_slide,
                display_text,
                top_inch=current_y,
                est_lines=lines_needed,
            )
            current_y += used_h + 0.15
            used_lines += lines_needed

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
