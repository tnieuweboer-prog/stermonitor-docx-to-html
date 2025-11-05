# pptx_converter.py
import io
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# layout-indexen uit de standaard PowerPoint template
TITLE_LAYOUT = 0        # titel-slide
TITLE_ONLY_LAYOUT = 5   # alleen titel, geen body
BLANK_LAYOUT = 6        # blanco dia

MAX_LINES_PER_SLIDE = 12   # jouw eis


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


def add_textbox(slide, text, top_offset_inch=1.0):
    """
    Maak een tekstvak op de dia met automatische afbreking.
    """
    left = Inches(0.8)
    top = Inches(top_offset_inch)
    width = Inches(8.0)
    height = Inches(0.8)

    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.text = text
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


def create_title_slide(prs, title_text="Inhoud uit Word"):
    slide = prs.slides.add_slide(prs.slide_layouts[TITLE_LAYOUT])
    slide.shapes.title.text = title_text
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


def docx_to_pptx(file_like):
    doc = Document(file_like)
    prs = Presentation()

    # alle afbeeldingen uit docx
    all_images = extract_images(doc)
    img_ptr = 0

    # start met een titel-slide
    current_slide = create_title_slide(prs)
    # beginhoogte voor tekst op deze slide
    current_text_y = 2.0
    # aantal tekstregels op huidige slide
    current_line_count = 0

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        has_image = any("graphic" in run._element.xml for run in para.runs)

        # bepalen of dit een "kop" moet zijn
        is_heading = para.style.name.startswith("Heading")
        is_bold_title = has_bold(para) and not has_image and not is_word_list_paragraph(para)

        # 1. nieuwe dia bij kop of vet → title-only slide
        if is_heading or is_bold_title:
            current_slide = create_title_only_slide(prs, para_text_plain(para))
            current_text_y = 2.0
            current_line_count = 0
            continue

        # 2. afbeelding → aparte title-only slide met afbeelding
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
            current_text_y = 3.5
            current_line_count = 0
            continue

        # 3. gewone/opsommingstekst → eerst checken of er nog plek is
        if text:
            # als we al 12 regels hebben → nieuwe BLANCO slide
            if current_line_count >= MAX_LINES_PER_SLIDE:
                current_slide = create_blank_slide(prs)
                current_text_y = 1.0
                current_line_count = 0

            # opsomming krijgt een puntje aan het begin
            if is_word_list_paragraph(para):
                display_text = "• " + text
            else:
                display_text = text

            add_textbox(current_slide, display_text, top_offset_inch=current_text_y)
            current_text_y += 0.7  # volgende regel iets lager
            current_line_count += 1

    # als alles verwerkt is, teruggeven als bytes
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out
