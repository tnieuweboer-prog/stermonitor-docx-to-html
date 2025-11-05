# pptx_converter.py
import io
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


def extract_images(doc):
    imgs = []
    idx = 1
    for rel in doc.part.rels.values():
        if rel.reltype == RT.IMAGE:
            imgs.append((f"image_{idx}.{rel.target_part.partname.ext}", rel.target_part.blob))
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
    """Maak een los tekstvak op de dia op een vaste plek en zet tekst erin."""
    left = Inches(0.8)
    top = Inches(top_offset_inch)
    width = Inches(8.0)
    height = Inches(0.6)
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.text = text
    # stijl
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)
    return shape


def docx_to_pptx(file_like):
    doc = Document(file_like)
    prs = Presentation()

    # alle images
    all_images = extract_images(doc)
    img_ptr = 0

    # titel-dia
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Inhoud uit Word"
    if len(title_slide.placeholders) > 1:
        title_slide.placeholders[1].text = "Geconverteerd voor LessonUp"
    current_slide = title_slide
    # we houden per slide bij waar we het volgende tekstvak onder moeten zetten
    current_text_y = 2.0  # in inches onder de titel

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        # 1. echte heading of vet → nieuwe dia
        is_heading = para.style.name.startswith("Heading")
        is_bold_title = has_bold(para) and not has_image and not is_word_list_paragraph(para)

        if is_heading or is_bold_title:
            current_slide = prs.slides.add_slide(prs.slide_layouts[1])
            current_slide.shapes.title.text = para_text_plain(para)
            current_text_y = 2.0  # reset tekstpositie voor nieuwe slide
            continue

        # 2. afbeelding → aparte dia
        if has_image:
            img_slide = prs.slides.add_slide(prs.slide_layouts[5])
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
            # volgende teksten op deze slide (onder afbeelding)
            current_slide = img_slide
            current_text_y = 3.5
            continue

        # 3. opsomming → ook tekstvak, maar we kunnen een bullet toevoegen
        if is_word_list_paragraph(para):
            tb = add_textbox(current_slide, text, top_offset_inch=current_text_y)
            current_text_y += 0.6  # volgende vak wat lager
            continue

        # 4. gewone tekst → altijd tekstvak
        if text:
            add_textbox(current_slide, text, top_offset_inch=current_text_y)
            current_text_y += 0.6

    # naar bytes
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

