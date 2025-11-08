import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _p(doc, text="", bold=False, size=12, align=None):
    """klein hulpfunctietje om een paragraaf te maken in Arial."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.bold = bold
    if align:
        p.alignment = align
    return p


def add_cover_page(doc: Document, *, vak: str, profieldeel: str, opdracht_nr: str, opdracht_titel: str,
                   duur: str, docent: str = "", klas: str = ""):
    """
    Maakt de voorkant van het werkboekje zoals jouw voorbeeld:
    - alles in Arial
    - velden die jij invult via app.py
    """
    # witregel boven
    _p(doc, "")

    # vak (groot, vet, gecentreerd)
    _p(doc, vak, bold=True, size=20, align=WD_ALIGN_PARAGRAPH.CENTER)

    # profieldeel
    _p(doc, f"Profieldeel: {profieldeel}", size=14, align=WD_ALIGN_PARAGRAPH.CENTER)

    # leeg
    _p(doc, "")

    # opdracht
    _p(doc, f"Opdracht {opdracht_nr}:", bold=True, size=14)
    _p(doc, opdracht_titel, bold=True, size=18)

    # duur
    _p(doc, f"Duur van de opdracht:     {duur}", size=12)

    # leeg
    _p(doc, "")

    # naam / klas blokje
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Naam:"
    hdr_cells[1].text = ""
    hdr_cells = table.rows[1].cells
    hdr_cells[0].text = "Klas:"
    hdr_cells[1].text = ""

    # alles in Arial zetten
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(12)

    # extra ruimte
    _p(doc, "")
    _p(doc, "")


def build_workbook_docx_front_and_steps(meta: dict, steps: list[dict]) -> io.BytesIO:
    """
    meta = {
      "vak": "BWI",
      "profieldeel": "Wonen en interieur",
      "opdracht_nr": "1",
      "opdracht_titel": "Wallmen",
      "duur": "11 x 45 minuten",
      "docent": "Jan Jansen",
      "klas": "3B"
    }
    steps = [
      {"title": "Stap 1", "text_blocks": ["..."], "images": [b'...']},
      ...
    ]
    """
    doc = Document()

    # voorkant
    add_cover_page(
        doc,
        vak=meta.get("vak", "BWI"),
        profieldeel=meta.get("profieldeel", ""),
        opdracht_nr=meta.get("opdracht_nr", "1"),
        opdracht_titel=meta.get("opdracht_titel", "Opdracht"),
        duur=meta.get("duur", ""),
        docent=meta.get("docent", ""),
        klas=meta.get("klas", ""),
    )

    # pagina-einde zodat stappen op nieuwe pagina komen
    doc.add_page_break()

    # stappen
    for i, step in enumerate(steps, start=1):
        doc.add_heading(f"Stap {i}", level=1)
        title = step.get("title") or ""
        if title:
            _p(doc, title, bold=True, size=12)
        for txt in step.get("text_blocks", []):
            _p(doc, txt, size=11)
        # afbeeldingen zou je hier kunnen plaatsen, maar die haal jij al op in app.py

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out
