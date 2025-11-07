import os
import json
import time
from docx import Document
from openai import OpenAI, RateLimitError, APIError


def docx_to_plain_text(file_like) -> str:
    """
    Leest alle tekst uit een .docx-bestand en zet het achter elkaar.
    Dit is stap 1: Word → ruwe tekst.
    """
    doc = Document(file_like)
    parts = []
    for p in doc.paragraphs:
        txt = (p.text or "").strip()
        if txt:
            parts.append(txt)
    return "\n".join(parts)


def call_openai_json(prompt: str, model: str = "gpt-4o-mini", max_retries: int = 5) -> dict:
    """
    Roept OpenAI aan en geeft altijd JSON terug.
    Probeert een paar keer bij rate-limit.
    Gooit een nette fout bij te weinig tegoed.
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("Geen OPENAI_API_KEY ingesteld.")

    client = OpenAI(api_key=api_key)
    delay = 1.0

    for attempt in range(1, max_retries + 1):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
            )
            return json.loads(resp.choices[0].message.content)

        except RateLimitError:
            if attempt == max_retries:
                raise RuntimeError("AI-limiet bereikt. Probeer later opnieuw.")
            time.sleep(delay)
            delay *= 2

        except APIError as e:
            # dit is de “insufficient_quota”-case
            if "insufficient_quota" in str(e):
                raise RuntimeError("AI-tegoed is op bij OpenAI. Voeg tegoed toe of gebruik een andere key.")
            if attempt == max_retries:
                raise RuntimeError(f"AI-fout bleef terugkomen: {e}")
            time.sleep(delay)
            delay *= 2


def ai_make_lesson_from_text(full_text: str) -> dict:
    """
    Stuurt ALLE tekst in één keer naar AI en vraagt:
    'maak er vmbo-dia's van'.
    """
    prompt = f"""
Je krijgt hieronder lesstof uit een Word-document.
Maak hier een les van voor een VMBO-klas (basis/kader/GL).

Maak meerdere dia's.
Voor elke dia:
- 1 korte, begrijpelijke titel (max 8 woorden)
- 2 of 3 korte zinnen in de je-vorm (vertellend)
- 1 controlevraag
- eenvoudige woorden
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
    return call_openai_json(prompt)


def docx_to_vmbo_lesson_json(file_like) -> dict:
    """
    DIT is de functie die je in app.py importeert.
    DOCX → tekst → AI → JSON
    """
    full_text = docx_to_plain_text(file_like)
    return ai_make_lesson_from_text(full_text)

