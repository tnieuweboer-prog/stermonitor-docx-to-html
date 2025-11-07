import os
import json
from docx import Document


def docx_to_plain_text(file_like) -> str:
    """Leest alle tekst uit een .docx en plakt het aan elkaar."""
    doc = Document(file_like)
    parts = []
    for p in doc.paragraphs:
        txt = (p.text or "").strip()
        if txt:
            parts.append(txt)
    return "\n".join(parts)


def ai_make_lesson_from_text(full_text: str) -> dict:
    """
    Stuurt de HELE tekst in één keer naar AI en vraagt:
    'maak er een vmbo-lessonup-les van'.
    Geeft JSON terug met slides.
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Geen OPENAI_API_KEY ingesteld.")

    from openai import OpenAI
    client = OpenAI(api_key=api_key)

    prompt = f"""
Je krijgt hieronder lesstof uit een Word-document.
Maak hier een les van voor een VMBO-klas (basis/kader/GL).

Maak meerdere dia's.
Voor elke dia:
- 1 korte, begrijpelijke titel (max 8 woorden)
- 2 of 3 korte zinnen in je-vorm (vertellend)
- 1 controlevraag
- gebruik eenvoudige woorden
- herhaal de titel NIET in de tekst

Geef ALLEEN JSON in dit formaat:

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
{full_text}
"""

    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
    )

    data = json.loads(resp.choices[0].message.content)

    # klein beetje valideren
    slides = data.get("slides")
    if not slides:
        raise RuntimeError("AI gaf geen slides terug.")
    return data


def docx_to_vmbo_lesson_json(file_like) -> dict:
    """
    Hoofdfunctie voor stap 1:
    DOCX → tekst → AI → JSON (slides)
    """
    text = docx_to_plain_text(file_like)
    return ai_make_lesson_from_text(text)
