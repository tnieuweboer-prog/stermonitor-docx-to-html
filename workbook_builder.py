import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _p(doc, text="", bold=False, size=12, align=None):
    """Hulpfunctie voor een nette paragraaf in Arial."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.bold = bold
    if align:
        p.alignment = align
    return p


def _add_logo_in_header(doc: Document, logo_bytes: bytes):
    """Voegt het Triade-logo (100x100 px) toe aan de koptekst, rechts uitgelijnd."""
    if not logo_bytes:
        return
    section = doc.sections[0]
    header = section.header
    paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run()
    run.add_picture(io.BytesIO(logo_bytes), width=Inches(1.0), height=Inches(1.0))


def add_cover_page(doc: Document, *, vak: str, profieldeel: str, opdracht_nr: str,
                   opdracht_titel: str, duur: str, docent: str = "", klas: str = "",
                   logo: bytes = None):
    """
    Bouwt de voorkant van het werkboekje op met vaste layout en Triade-logo in de header.
    """
    # Voeg logo toe aan header
    if logo:
        _add_logo_in_header(doc, logo)

    # witruimte boven
    _p(doc, "")

    # Vak (groot, vet, gecentreerd)
    _p(doc, vak, bold=True, size=20, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Profieldeel
    _p(doc, f"Profieldeel: {profieldeel}", size=14, align=WD_ALIGN_PARAGRAPH.CENTER)

    # lege regel
    _p(doc, "")

    # Opdracht
    _p(doc, f"Opdracht {opdracht_nr}:", bold=True, size=14)
    _p(doc, opdracht_titel, bold=True, size=18)

    # Duur
    _p(doc, f"Duur van de opdracht:     {duur}", size=12)

    # lege regel
    _p(doc, "")

    # Naam / klas tabel
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Naam:"
    hdr_cells[1].text = ""
    hdr_cells = table.rows[1].cells
    hdr_cells[0].text = "Klas:"
    hdr_cells[1].text = ""

    # lettertype forceren naar Arial
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(12)

    # extra ruimte onderaan
    _p(doc, "")
    _p(doc, "")


def build_workbook_docx_front_and_steps(meta: dict, steps: list[dict]) -> io.BytesIO:
    """
    Bouwt het volledige werkboekje (voorkant + stappenpagina's).
    meta bevat algemene info, steps bevat blokken met tekst en afbeeldingen.
    """
    doc = Document()

    # Voorpagina
    add_cover_page(
        doc,
        vak=meta.get("vak", "BWI"),
        profieldeel=meta.get("profieldeel", ""),
        opdracht_nr=meta.get("opdracht_nr", "1"),
        opdracht_titel=meta.get("opdracht_titel", "Opdracht"),
        duur=meta.get("duur", ""),
        docent=meta.get("docent", ""),
        klas=meta.get("klas", ""),
        logo=meta.get("logo", None),
    )

    # Nieuwe pagina voor stappen
    doc.add_page_break()

    # Stappenpagina's
    for i, step in enumerate(steps, start=1):
        doc.add_heading(f"Stap {i}", level=1)
        title = step.get("title") or ""
        if title:
            _p(doc, title, bold=True, size=12)
        for txt in step.get("text_blocks", []):
            _p(doc, txt, size=11)

        # Voeg eventuele afbeeldingen toe (van Cloudinary)
        for img_bytes in step.get("images", []):
            doc.add_picture(io.BytesIO(img_bytes), width=Inches(3.5))
            _p(doc, "")

        doc.add_page_break()

    # Teruggeven als bestand
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

