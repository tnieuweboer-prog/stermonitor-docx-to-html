import os
import io
import json
from docx import Document

# probeer pas OpenAI te importeren als we 'm echt nodig hebben
try:
    from openai import OpenAI
    HAS_OPENAI = True
except Exception:
    HAS_OPENAI = False


def read_docx_to_text(file_like) -> str:
    """Leest alle tekst uit een .docx-bestand."""
    doc = Document(file_like)
    parts = []
    for p in doc.paragraphs:
        txt = (p.text or "").strip()
        if txt:
            parts.append(txt)
    return "\n".join(parts)


def call_openai_for_lesson(text: str) -> dict:
    """Eén AI-aanroep: zet tekst om in een VMBO-lesstructuur."""
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key or not HAS_OPENAI:
        raise RuntimeError("Geen OpenAI beschikbaar.")

    client = OpenAI(api_key=api_key)

    prompt = f"""
Je krijgt hieronder lesstof uit een Word-document.
Maak hier een les van voor een VMBO-klas (basis/kader/GL).

Voor elk onderdeel:
- 1 korte, begrijpelijke titel (max 8 woorden)
- 2 of 3 korte zinnen in de je-vorm (vertellend)
- 1 controlevraag
- gebruik eenvoudige woorden

Geef ALLEEN geldig JSON in dit formaat:

{{
  "slides": [
    {{
      "title": "…",
      "text": ["…", "…", "…"],
      "check": "…"
    }}
  ]
}}

Hier is de lesstof:
{text}
"""

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
    )

    data = json.loads(resp.choices[0].message.content)
    slides = data.get("slides")
    if not slides:
        raise RuntimeError("AI gaf geen slides terug.")
    return data


def build_docx_from_lesson(lesson: dict) -> io.BytesIO:
    """Maakt een nieuw Word-bestand in les-format (kop → tekst → vraag)."""
    doc = Document()
    for slide in lesson.get("slides", []):
        title = slide.get("title") or "Lesonderdeel"
        text_lines = slide.get("text") or []
        check = slide.get("check") or ""

        # kop
        doc.add_heading(title, level=1)

        # tekstregels
        for line in text_lines:
            doc.add_paragraph(line)

        # controlevraag
        if check:
            p = doc.add_paragraph()
            p.add_run(check).bold = True

        # lege regel
        doc.add_paragraph("")

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


def local_fallback_docx(original_text: str) -> io.BytesIO:
    """Fallback als AI niet werkt of quota op is."""
    doc = Document()
    doc.add_heading("Les (lokaal gegenereerd)", level=1)
    doc.add_paragraph("Je leert vandaag iets over installatietechniek.")
    doc.add_paragraph("Lees het originele bestand erbij voor meer uitleg.")
    doc.add_paragraph("")
    doc.add_paragraph(original_text[:800])  # kort stukje van originele tekst
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


def docx_to_vmbo_lesson_json(file_like) -> io.BytesIO:
    """
    Belangrijk:
    - gebruikt nog steeds dezelfde functienaam (dus app.py werkt)
    - maakt één AI-aanroep (dus geen rate-limit)
    - geeft direct een nieuw .docx terug in les-stijl
    """
    raw_text = read_docx_to_text(file_like)
    try:
        lesson = call_openai_for_lesson(raw_text)
        return build_docx_from_lesson(lesson)
    except Exception as e:
        print(f"AI-fout: {e}")
        return local_fallback_docx(raw_text)


