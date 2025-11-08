import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _p(doc, text="", bold=False, size=12, align=None):
    """Voegt een paragraaf toe in Arial."""
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
    """
    Layout van de voorkant:
    Logo in koptekst (rechts)
    Opdracht :
    <titel>
    <vak>
    Keuze/profieldeel (alleen als ingevuld)
    Docent
    Duur
    [Afbeelding]
    [Tabel Naam / Klas]
    """

    # 1️⃣ Logo in de koptekst (alleen voor de eerste sectie)
    if logo:
        add_logo_to_header(doc.sections[0], logo)

    # 2️⃣ "Opdracht :" vetgedrukt, 14 pt
    _p(doc, "Opdracht :", bold=True, size=14)

    # 3️⃣ De ingevulde titel vetgedrukt, 28 pt
    if opdracht_titel:
        _p(doc, opdracht_titel, bold=True, size=28)
    else:
        _p(doc, " ", size=28)

    _p(doc, "")

    # 4️⃣ Vak (zoals BWI)
    _p(doc, vak, bold=True, size=14)

    # 5️⃣ Alleen weergeven als profieldeel is ingevuld
    if profieldeel:
        _p(doc, "Keuze/profieldeel:", size=12)
        _p(doc, profieldeel, size=12)

    # 6️⃣ Docent
    if docent:
        _p(doc, f"Docent: {docent}", size=12)
    else:
        _p(doc, "Docent:", size=12)

    # 7️⃣ Duur
    if duur:
        _p(doc, f"Duur van de opdracht:     {duur}", size=12)
    else:
        _p(doc, "Duur van de opdracht:", size=12)

    _p(doc, "")

    # 8️⃣ Afbeelding (optioneel)
    if cover_bytes:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        r.add_picture(io.BytesIO(cover_bytes), width=Inches(4.5))
        _p(doc, "")

    # 9️⃣ Tabel Naam / Klas
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
    """Bouwt het volledige werkboekje met voorkant en stappen onder elkaar."""
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

    # Stappen beginnen op nieuwe pagina
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


