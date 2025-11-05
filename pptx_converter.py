# pptx_converter.py
import io
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


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
    """Herken Word-opsommingen (lijststijlen en numPr)."""
    name = (para.style.name or "").lower()
    if "list" in name or "lijst" in name or "opsom" in name:
        return True
    ppr = para._p.pPr
    return ppr is not None and ppr.numPr is not None


def has_bold(para):
    """True als er in deze paragraaf ergens vet staat."""
    return any(run.bold for run in para.runs)


def para_text_plain(para):
    """Alle runs samenvoegen tot platte tekst."""
    parts = []
    for run in para.runs:
        if run.text:
            parts.append(run.text)
    return "".join(parts).strip()


def add_textbox(slide, text, top_offset_inch=2.0):
    """
    Maak een tekstvak op de slide met automatische woordafbreking
    en Arial 16.
    """
    left = Inches(0.8)
    top = Inches(top_offset_inch)
    width = Inches(8.0)
    height = Inches(0.8)  # basis-hoogte

    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.text = text

    # tekst laten afbreken binnen het vak
    tf.word_wrap = True
    tf.auto_size = False
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)

    # stijl
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)

    return shape


def docx_to_pptx(file_like):
    """
    Converteer een .docx naar een PowerPoint:
    - Heading OF vetgedrukt → nieuwe dia met titel
    - Afbeelding → aparte dia met afbeelding
    - Overige tekst → tekstvak(ken) onder elkaar
    """
    doc = Document(file_like)
    prs = Presentation()

    # alle images uit docx
    all_images = extract_images(doc)
    img_ptr = 0

    # eerste dia
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Inhoud uit Word"
    if len(title_slide.placeholders) > 1:
        title_slide.placeholders[1].text = "Geconverteerd voor LessonUp"

    current_slide = title_slide
    # waar komt het volgende tekstvak op deze slide
    current_text_y = 2.0  # inches vanaf boven

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        # bepaal type paragraaf
        is_heading = para.style.name.startswith("Heading")
        is_bold_title = has_bold(para) and not has_image and not is_word_list_paragraph(para)

        # 1. heading of vet → nieuwe dia
        if is_heading or is_bold_title:
            current_slide = prs.slides.add_slide(prs.slide_layouts[1])
            current_slide.shapes.title.text = para_text_plain(para)
            current_text_y = 2.0  # reset positie op nieuwe dia
            continue

        # 2. afbeelding → aparte dia
        if has_image:
            img_slide = prs.slides.add_slide(prs.slide_layouts[5])  # title only
            img_slide.shapes.title.text = "Afbeelding"
            if img_ptr < len(all_images):
                _, img_bytes = all_images[img_ptr]
                img_ptr += 1
                img_stream = io.BytesIO(img_bytes)
                # afbeelding invoegen
                img_slide.shapes.add_picture(
                    img_stream,
                    Inches(1),
                    Inches(1.2),
                    width=Inches(6),
                )
            # volgende tekst op deze slide moet lager komen
            current_slide = img_slide
            current_text_y = 3.5
            continue

        # 3. opsomming → ook in tekstvak (kan later naar bullets)
        if is_word_list_paragraph(para):
            if text:
                add_textbox(current_slide, text, top_offset_inch=current_text_y)
                current_text_y += 0.7
            continue

        # 4. gewone tekst → altijd tekstvak
        if text:
            add_textbox(current_slide, text, top_offset_inch=current_text_y)
            current_text_y += 0.7

    # opslaan naar bytes
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
