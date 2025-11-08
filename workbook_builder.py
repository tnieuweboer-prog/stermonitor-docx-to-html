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


def add_logo_to_header(section, logo_bytes: bytes):
    """Voegt het logo rechtsboven toe in de koptekst (100x100px)."""
    header = section.header
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run()
    run.add_picture(io.BytesIO(logo_bytes), width=Inches(1.0), height=Inches(1.0))


def add_materiaalstaat_page(doc: Document, materialen: list[dict]):
    """Voegt de materiaalstaat toe (direct na de voorpagina)."""
    doc.add_page_break()
    doc.add_heading("Materiaalstaat", level=1)

    cols = ["Nummer", "Aantal", "Benaming", "Lengte", "Breedte", "Dikte", "Materiaal"]
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = "Table Grid"

    hdr_cells = table.rows[0].cells
    for idx, col_name in enumerate(cols):
        hdr_cells[idx].text = col_name
        for p in hdr_cells[idx].paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.name = "Arial"
                r.font.size = Pt(12)

    for item in materialen:
        row = table.add_row().cells
        for j, key in enumerate(cols):
            row[j].text = item.get(key, "")

    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(11)


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
    """Bouwt de voorpagina van het werkboekje."""
    if logo:
        add_logo_to_header(doc.sections[0], logo)

    _p(doc, "Opdracht :", bold=True, size=14)

    if opdracht_titel:
        _p(doc, opdracht_titel, bold=True, size=28)
    else:
        _p(doc, " ", size=28)

    _p(doc, "")
    _p(doc, vak, bold=True, size=14)

    if profieldeel:
        _p(doc, "Keuze/profieldeel:", size=12)
        _p(doc, profieldeel, size=12)

    if docent:
        _p(doc, f"Docent: {docent}", size=12)
    else:
        _p(doc, "Docent:", size=12)

    if duur:
        _p(doc, f"Duur van de opdracht:     {duur}", size=12)
    else:
        _p(doc, "Duur van de opdracht:", size=12)

    _p(doc, "")

    if cover_bytes:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        r.add_picture(io.BytesIO(cover_bytes), width=Inches(4.5))
        _p(doc, "")

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
    """Bouwt het volledige werkboekje met:
    - Voorpagina
    - Optioneel: Materiaalstaat
    - Stappenplan
    """
    doc = Document()

    add_cover_page(
        doc,
        opdracht_titel=meta.get("opdracht_titel", ""),
        vak=meta.get("vak", "BWI"),
        profieldeel=meta.get("profieldeel", ""),
        docent=meta.get("docent", ""),
        duur=meta.get("duur", ""),
        logo=meta.get("logo"),
        cover_bytes=meta.get("cover_bytes"),
    )

    # Materiaalstaat komt direct na de voorpagina
    if meta.get("include_materiaalstaat"):
        add_materiaalstaat_page(doc, meta.get("materialen", []))

    # Dan het stappenplan
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

