import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _p(doc, text="", bold=False, size=12, align=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.bold = bold
    if align:
        p.alignment = align
    return p


def add_cover_page(
    doc: Document,
    *,
    vak: str,
    profieldeel: str,
    opdracht_nr: str,
    opdracht_titel: str,
    duur: str,
    docent: str = "",
    klas: str = "",
    logo: bytes = None,
    cover_upload=None,
):
    # bovenste rij met logo rechts
    if logo:
        tbl = doc.add_table(rows=1, cols=2)
        left_cell, right_cell = tbl.rows[0].cells
        # links laten we leeg
        p_right = right_cell.paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p_right.add_run()
        run.add_picture(io.BytesIO(logo), width=Inches(1.0), height=Inches(1.0))
    else:
        _p(doc, "")

    _p(doc, "")

    # hoofdvelden
    _p(doc, vak, bold=True, size=20, align=WD_ALIGN_PARAGRAPH.CENTER)

    if profieldeel:
        _p(doc, f"Profieldeel: {profieldeel}", size=14, align=WD_ALIGN_PARAGRAPH.CENTER)

    _p(doc, "")
    _p(doc, f"Opdracht {opdracht_nr}:", bold=True, size=14)
    _p(doc, opdracht_titel, bold=True, size=18)

    if duur:
        _p(doc, f"Duur van de opdracht:     {duur}", size=12)

    if docent:
        _p(doc, f"Docent: {docent}", size=12)
    if klas:
        _p(doc, f"Klas: {klas}", size=12)

    _p(doc, "")

    # geüploade cover-afbeelding (via app.py doorgegeven)
    if cover_upload is not None:
        # kan een Streamlit UploadedFile zijn of al bytes
        if hasattr(cover_upload, "read"):
            cover_bytes = cover_upload.read()
        else:
            cover_bytes = cover_upload
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        r.add_picture(io.BytesIO(cover_bytes), width=Inches(4.5))

    _p(doc, "")

    # invulblokje onderaan
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Naam:"
    table.rows[1].cells[0].text = "Klas:"

    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(12)


def build_workbook_docx_front_and_steps(meta: dict, steps: list[dict]) -> io.BytesIO:
    doc = Document()

    add_cover_page(
        doc,
        vak=meta.get("vak", "BWI"),
        profieldeel=meta.get("profieldeel", ""),
        opdracht_nr=meta.get("opdracht_nr", "1"),
        opdracht_titel=meta.get("opdracht_titel", "Opdracht"),
        duur=meta.get("duur", ""),
        docent=meta.get("docent", ""),
        klas=meta.get("klas", ""),
        logo=meta.get("logo"),
        cover_upload=meta.get("cover_upload"),
    )

    # stappen achter de voorkant
    if steps:
        doc.add_page_break()

    for i, step in enumerate(steps, start=1):
        doc.add_heading(f"Stap {i}", level=1)
        title = step.get("title") or ""
        if title:
            _p(doc, title, bold=True, size=12)
        for txt in step.get("text_blocks", []):
            _p(doc, txt, size=11)
        for img_bytes in step.get("images", []):
            doc.add_picture(io.BytesIO(img_bytes), width=Inches(3.5))
            _p(doc, "")
        if i < len(steps):
            doc.add_page_break()

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


