import os
import io
import json
import time
from docx import Document
from openai import OpenAI, RateLimitError, APIError


def docx_to_blocks(file_like):
    """
    Leest het .docx bestand en maakt blokken: [{"title": ..., "body": ...}, ...]
    Een blok = kop + bijbehorende tekst.
    """
    doc = Document(file_like)
    blocks = []
    current_title = None
    current_body = []

    for para in doc.paragraphs:
        txt = (para.text or "").strip()
        if not txt:
            continue

        # Koppen herkennen
        is_heading = (
            (para.style and para.style.name and para.style.name.lower().startswith("heading"))
            or any(r.bold for r in para.runs)
            or (len(txt) <= 50 and txt.upper() == txt)
        )

        if is_heading:
            # vorig blok afsluiten
            if current_title or current_body:
                blocks.append({
                    "title": current_title or "Lesonderdeel",
                    "body": "\n".join(current_body).strip()
                })
            current_title = txt
            current_body = []
        else:
            current_body.append(txt)

    # laatste blok toevoegen
    if current_title or current_body:
        blocks.append({
            "title": current_title or "Lesonderdeel",
            "body": "\n".join(current_body).strip()
        })

    if not blocks:
        raise RuntimeError("Geen tekst of koppen gevonden in het document.")
    return blocks


def ai_generate_slide(client, title, body):
    """
    Eén AI-call per onderdeel.
    """
    prompt = f"""
Maak een korte dia voor een VMBO-les (basis/kader/GL).

Onderwerp: {title}
Lesstof:
{body}

Schrijf:
- "title": een pakkende titel (max 8 woorden)
- "text": 2 of 3 korte, vertellende zinnen in de je-vorm
- "check": 1 controlevraag
Gebruik eenvoudige woorden en leg kort uit hoe iets werkt.

Geef ALLEEN geldig JSON:
{{"title": "...", "text": ["...", "..."], "check": "..." }}
"""
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
    )
    slide = json.loads(resp.choices[0].message.content)
    return slide


def build_word_from_slides(slides):
    """
    Bouwt het uiteindelijke les-Word-bestand in jouw format:
    - kop
    - uitleg
    - vraag
    """
    doc = Document()
    doc.add_heading("LessonUp-les (gegenereerd met AI)", level=0)

    for idx, slide in enumerate(slides, start=1):
        title = slide.get("title") or f"Onderdeel {idx}"
        text_lines = slide.get("text") or []
        check = slide.get("check") or ""

        # onderdeel-kop
        doc.add_heading(f"{idx}️⃣ {title}", level=1)

        # uitleg
        p = doc.add_paragraph()
        p.add_run("Uitleg").bold = True

        for line in text_lines:
            doc.add_paragraph(line)

        if check:
            q = doc.add_paragraph()
            q.add_run("Vraag").bold = True
            doc.add_paragraph(check)

        doc.add_paragraph("")

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


def docx_to_vmbo_lesson_json(file_like) -> io.BytesIO:
    """
    Hoofdfunctie voor de app.
    - Splits document in blokken
    - Voor elk blok 1 AI-call met kleine pauze
    - Combineert alle resultaten tot één Word-bestand
    - Geen fallback: faalt netjes bij fouten
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY ontbreekt. Voeg je sleutel toe in de omgeving.")
    client = OpenAI(api_key=api_key)

    blocks = docx_to_blocks(file_like)
    slides = []

    for i, b in enumerate(blocks, start=1):
        title = b.get("title") or f"Onderdeel {i}"
        body = b.get("body") or ""
        try:
            slide = ai_generate_slide(client, title, body)
            slides.append(slide)
        except (RateLimitError, APIError) as e:
            raise RuntimeError(f"AI-call mislukt bij onderdeel {i}: {e}")
        # kleine pauze om limiet te vermijden
        time.sleep(2)

    return build_word_from_slides(slides)

