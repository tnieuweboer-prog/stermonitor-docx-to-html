# pptx_converter.py
import io
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


TITLE_LAYOUT = 0       # titel-slide
TITLE_ONLY_LAYOUT = 5  # alleen titel, geen groot tekstvak
BLANK_LAYOUT = 6       # blanco


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
    parts = []
    for run in para.runs:
        if run.text:
            parts.append(run.text)
    return "".join(parts).strip()


def add_textbox(slide, text, top_offset_inch=2.0):
    """
    Maak een eigen tekstvak dat niet achter een placeholder verdwijnt.
    """
    left = Inches(0.8)
    top = Inches(top_offset_inch)
    width = Inches(8.0)
    height = Inches(0.8)

    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.text = text

    # tekst laten afbreken binnen het vak
    tf.word_wrap = True
    tf.auto_size = False
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)

    # stijl: Arial 16 zwart
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)

    return shape


def create_title_only_slide(prs, title_text: str):
    """
    Maak een dia met alleen een titel (layout 5), geen groot body-vak.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[TITLE_ONLY_LAYOUT])
    slide.shapes.title.text = title_text
    return slide


def docx_to_pptx(file_like):
    doc = Document(file_like)
    prs = Presentation()

    # alle afbeeldingen uit docx zodat we ze kunnen plaatsen
    all_images = extract_images(doc)
    img_ptr = 0

    # eerste slide: standaard titel
    first = prs.slides.add_slide(prs.slide_layouts[TITLE_LAYOUT])
    first.shapes.title.text = "Inhoud uit Word"
    if len(first.placeholders) > 1:
        first.placeholders[1].text = "Geconverteerd voor LessonUp"

    current_slide = first
    current_text_y = 2.0  # waar het eerste tekstvak komt op de huidige dia

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        # bepalen of dit een nieuwe dia moet zijn
        is_heading = para.style.name.startswith("Heading")
        is_bold_title = has_bold(para) and not has_image and not is_word_list_paragraph(para)

        # 1. nieuwe dia bij kop OF vet
        if is_heading or is_bold_title:
            title_text = para_text_plain(para)
            current_slide = create_title_only_slide(prs, title_text)
            current_text_y = 2.0
            continue

        # 2. afbeelding → aparte dia (ook title only)
        if has_image:
            img_slide = prs.slides.add_slide(prs.slide_layouts[TITLE_ONLY_LAYOUT])
            img_slide.shapes.title.text = "Afbeelding"
            if img_ptr < len(all_images):
                _, img_bytes = all_images[img_ptr]
                img_ptr += 1
                img_stream = io.BytesIO(img_bytes)
                img_slide.shapes.add_picture(
                    img_stream,
                    Inches(1),
                    Inches(1.2),
                    width=Inches(6),
                )
            current_slide = img_slide
            current_text_y = 3.5  # na afbeelding verder naar beneden
            continue

        # 3. opsomming → ook gewoon eigen tekstvak
        if is_word_list_paragraph(para):
            if text:
                add_textbox(current_slide, "• " + text, top_offset_inch=current_text_y)
                current_text_y += 0.7
            continue

        # 4. gewone tekst → eigen tekstvak
        if text:
            add_textbox(current_slide, text, top_offset_inch=current_text_y)
            current_text_y += 0.7

    # naar bytes
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
