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
    """True als er ergens in deze paragraaf vetgedrukte tekst staat."""
    return any(run.bold for run in para.runs)


def para_text_plain(para):
    """Alle runs samenvoegen tot één tekst (zonder HTML)."""
    parts = []
    for run in para.runs:
        t = run.text
        if t:
            parts.append(t)
    return "".join(parts).strip()


def _get_body(slide, prs):
    """Zoek een text_frame op de dia, anders maak een nieuwe dia met body."""
    for shape in slide.shapes:
        if hasattr(shape, "text_frame"):
            return shape.text_frame, slide
    new_slide = prs.slides.add_slide(prs.slide_layouts[1])
    return new_slide.shapes.placeholders[1].text_frame, new_slide


def _apply_style_to_text_frame(tf):
    """Zet alle tekst in dit tekstvak op Arial 16 zwart."""
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name = "Arial"
            r.font.size = Pt(16)
            r.font.color.rgb = RGBColor(0, 0, 0)


def docx_to_pptx(file_like):
    doc = Document(file_like)
    prs = Presentation()

    # alle images (voor dia's met afbeelding)
    all_images = extract_images(doc)
    img_ptr = 0

    # titel-dia
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Inhoud uit Word"
    if len(title_slide.placeholders) > 1:
        title_slide.placeholders[1].text = "Geconverteerd voor LessonUp"
    current_slide = title_slide

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        # 1. Echte Word-kop → nieuwe dia
        is_heading = para.style.name.startswith("Heading")

        # 2. Jouw nieuwe regel: vetgedrukt → óók nieuwe dia
        is_bold_title = has_bold(para) and not has_image and not is_word_list_paragraph(para)

        if is_heading or is_bold_title:
            # nieuwe dia met titel = tekst van deze paragraaf
            current_slide = prs.slides.add_slide(prs.slide_layouts[1])
            # titeltekst: neem de platte tekst (ook bij bold)
            current_slide.shapes.title.text = para_text_plain(para)
            # body leegmaken en stylen
            body_tf, current_slide = _get_body(current_slide, prs)
            body_tf.text = ""
            _apply_style_to_text_frame(body_tf)
            continue

        # 3. Afbeelding → aparte dia met afbeelding
        if has_image:
            img_slide = prs.slides.add_slide(prs.slide_layouts[5])  # title only
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
            continue

        # 4. Opsomming → bullet op huidige dia
        if is_word_list_paragraph(para):
            body_tf, current_slide = _get_body(current_slide, prs)
            p = body_tf.add_paragraph()
            p.text = text
            p.level = 0
            _apply_style_to_text_frame(body_tf)
            continue

        # 5. Gewone tekst → bullet op huidige dia
        if text:
            body_tf, current_slide = _get_body(current_slide, prs)
            if body_tf.text == "":
                body_tf.text = text
            else:
                p = body_tf.add_paragraph()
                p.text = text
                p.level = 0
            _apply_style_to_text_frame(body_tf)

    # naar bytes
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

