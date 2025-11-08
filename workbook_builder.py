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
    opdracht_titel: str,
    vak: str,
    profieldeel: str,
    docent: str,
    duur: str,
    logo: bytes = None,
    cover_bytes: bytes = None,
):
    # bovenste rij met logo rechts (optioneel)
    if logo:
        tbl = doc.add_table(rows=1, cols=2)
        left_cell, right_cell = tbl.rows[0].cells
        p_right = right_cell.paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p_right.add_run()
        r.add_picture(io.BytesIO(logo), width=Inches(1.0), height=Inches(1.0))
    else:
        _p(doc, "")

    # Opdracht
    _p(doc, "Opdracht :", size=12)
    _p(doc, opdracht_titel, bold=True, size=14)
    _p(doc, "")

    # vak
    _p(doc, vak, bold=True, size=14)

    # Keuze/profieldeel
    _p(doc, "Keuze/profieldeel:", size=12)
    if profieldeel:
        _p(doc, profieldeel, size=12)

    # Docent
    if docent:
        _p(doc, f"Docent: {docent}", size=12)
    else:
        _p(doc, "Docent:", size=12)

    # Duur
    if duur:
        _p(doc, f"Duur van de opdracht:     {duur}", size=12)
    else:
        _p(doc, "Duur van de opdracht:", size=12)

    _p(doc, "")

    # 🔹 hier: optionele afbeelding op de voorpagina
    if cover_bytes:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        r.add_picture(io.BytesIO(cover_bytes), width=Inches(4.5))
        _p(doc, "")  # beetje ruimte na afbeelding

    # Naam / Klas in tabel (voor leerlingen)
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Naam:"
    table.rows[0].cells[1].text = ""
    table.rows[1].cells[0].text = "Klas:"
    table.rows[1].cells[1].text = ""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(12)

    _p(doc, "")
    _p(doc, "")


def build_workbook_docx_front_and_steps(meta: dict, steps: list[dict]) -> io.BytesIO:
    doc = Document()

    add_cover_page(
        doc,
        opdracht_titel=meta.get("opdracht_titel", ""),
        vak=meta.get("vak", "BWI"),
        profieldeel=meta.get("profieldeel", ""),
        docent=meta.get("docent", ""),
        duur=meta.get("duur", ""),
        logo=meta.get("logo"),
        cover_bytes=meta.get("cover_bytes"),  # 👈 hier komt de upload
    )

    # stappen op één pagina onder elkaar
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
        _p(doc, "")
        _p(doc, "")

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

