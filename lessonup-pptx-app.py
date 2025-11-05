# pptx_converter.py
import io
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from docx.opc.constants import RELATIONSHIP_TYPE as RT

def extract_images(doc):
    imgs=[]
    idx=1
    for rel in doc.part.rels.values():
        if rel.reltype==RT.IMAGE:
            imgs.append((f"image_{idx}.{rel.target_part.partname.ext}", rel.target_part.blob))
            idx+=1
    return imgs

def is_word_list_paragraph(para):
    name=(para.style.name or "").lower()
    if "list" in name or "lijst" in name or "opsom" in name:
        return True
    ppr = para._p.pPr
    return ppr is not None and ppr.numPr is not None

def _get_body(slide, prs):
    for shape in slide.shapes:
        if hasattr(shape, "text_frame"):
            return shape.text_frame, slide
    new_slide = prs.slides.add_slide(prs.slide_layouts[1])
    return new_slide.shapes.placeholders[1].text_frame, new_slide

def _apply_style_to_text_frame(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0,0,0)

def docx_to_pptx(file_like):
    doc = Document(file_like)
    prs = Presentation()
    all_images = extract_images(doc)
    img_ptr = 0

    # title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Inhoud uit Word"
    if len(slide.placeholders) > 1:
        slide.placeholders[1].text = "Geconverteerd voor LessonUp"
    current_slide = slide

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in r._element.xml for r in para.runs)
        if not text and not has_image:
            continue

        if para.style.name.startswith("Heading"):
            current_slide = prs.slides.add_slide(prs.slide_layouts[1])
            current_slide.shapes.title.text = text
            # empty body
            body_tf, current_slide = _get_body(current_slide, prs)
            body_tf.text = ""
            _apply_style_to_text_frame(body_tf)
            continue

        if has_image:
            s = prs.slides.add_slide(prs.slide_layouts[5])  # title only
            s.shapes.title.text = "Afbeelding"
            if img_ptr < len(all_images):
                _, b = all_images[img_ptr]; img_ptr += 1
                s.shapes.add_picture(io.BytesIO(b), Inches(1), Inches(1.2), width=Inches(6))
            current_slide = s
            continue

        # content -> bullets
        body_tf, current_slide = _get_body(current_slide, prs)
        if is_word_list_paragraph(para):
            p = body_tf.add_paragraph()
            p.text = text
            p.level = 0
        else:
            if body_tf.text == "":
                body_tf.text = text
            else:
                p = body_tf.add_paragraph()
                p.text = text
                p.level = 0
        _apply_style_to_text_frame(body_tf)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

