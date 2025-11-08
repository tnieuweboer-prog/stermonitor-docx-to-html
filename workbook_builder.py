import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _p(doc, text="", bold=False, size=12, align=None):
    """Maak een paragraaf in Arial."""
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
):
    """
    Voorkant zoals je laatste voorbeeld:
    Opdracht :
    <titel>

    BWI
    Keuze/profieldeel:
    Docent: ...
    Duur van de opdracht: ...

    Naam:
    Klas:
    """
    # bovenste rij met logo rechts
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

    # vak (bijv. BWI)
    _p(doc, vak, bold=True, size=14)

    # Keuze/profieldeel
    _p(doc, "Keuze/profieldeel:", size=12)
    if profieldeel:
        _p(doc, profieldeel, size=12)

    # docent
    if docent:
        _p(doc, f"Docent: {docent}", size=12)
    else:
        _p(doc, "Docent:", size=12)

    # duur
    if duur:
        _p(doc, f"Duur van de opdracht:     {duur}", size=12)
    else:
        _p(doc, "Duur van de opdracht:", size=12)

    _p(doc, "")
    _p(doc, "")

    # Naam / Klas in tabel (voor leerlingen om in te vullen)
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Naam:"
    table.rows[0].cells[1].text = ""
    table.rows[1].cells[0].text = "Klas:"
    table.rows[1].cells[1].text = ""

    # alle tekst in tabel -> Arial
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(12)

    _p(doc, "")
    _p(doc, "")
    # klaar met voorkant


def build_workbook_docx_front_and_steps(meta: dict, steps: list[dict]) -> io.BytesIO:
    """
    Bouwt werkboekje:
    - voorkant volgens layout
    - page break
    - alle stappen onder elkaar (géén page break per stap)
    """
    doc = Document()

    # voorkant
    add_cover_page(
        doc,
        opdracht_titel=meta.get("opdracht_titel", ""),
        vak=meta.get("vak", "BWI"),
        profieldeel=meta.get("profieldeel", ""),
        docent=meta.get("docent", ""),
        duur=meta.get("duur", ""),
        logo=meta.get("logo"),
    )

    # stappen beginnen op nieuwe pagina
    if steps:
        doc.add_page_break()

    for i, step in enumerate(steps, start=1):
        # titel van stap
        doc.add_heading(f"Stap {i}", level=1)

        # eventuele subtitel
        title = step.get("title") or ""
        if title:
            _p(doc, title, bold=True, size=12)

        # tekstblokken
        for txt in step.get("text_blocks", []):
            _p(doc, txt, size=11)

        # afbeeldingen (die jij via Cloudinary al als bytes aanlevert)
        for img_bytes in step.get("images", []):
            doc.add_picture(io.BytesIO(img_bytes), width=Inches(3.5))
            _p(doc, "")

        # i.p.v. page_break: gewoon een lege regel tussen twee stappen
        _p(doc, "")
        _p(doc, "")

    # teruggeven als bestand
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


