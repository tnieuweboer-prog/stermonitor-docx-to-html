import os
import base64
from html import escape
from typing import Optional, List, Dict

from docx import Document

# Pillow voor beeldmaten (optioneel, met fallback)
try:
    from PIL import Image
    PIL_OK = True
except Exception:
    PIL_OK = False

# --- Cloudinary (optioneel) ---
try:
    import cloudinary
    import cloudinary.uploader
except Exception:
    cloudinary = None


def _cloudinary_ready() -> bool:
    if cloudinary is None:
        return False
    url = os.getenv("CLOUDINARY_URL")
    if url:
        try:
            cloudinary.config(cloudinary_url=url, secure=True)
            return True
        except Exception:
            return False
    name = os.getenv("CLOUDINARY_CLOUD_NAME")
    key = os.getenv("CLOUDINARY_API_KEY")
    secret = os.getenv("CLOUDINARY_API_SECRET")
    if name and key and secret:
        try:
            cloudinary.config(cloud_name=name, api_key=key, api_secret=secret, secure=True)
            return True
        except Exception:
            return False
    return False


def _upload_bytes(img_bytes: bytes, folder="triade-html") -> Optional[str]:
    """Upload naar Cloudinary, retourneer secure_url; None bij geen config/fout."""
    if not _cloudinary_ready():
        return None
    try:
        res = cloudinary.uploader.upload(
            img_bytes,
            folder=folder,
            overwrite=True,
            use_filename=True,
            unique_filename=True,
            resource_type="image",
        )
        return res.get("secure_url") or res.get("url")
    except Exception:
        return None


def _image_size(img_bytes: bytes) -> Optional[tuple]:
    """Geef (w,h) in pixels terug, of None als Pillow niet beschikbaar/faalt."""
    if not PIL_OK:
        return None
    try:
        from io import BytesIO
        with Image.open(BytesIO(img_bytes)) as im:
            return im.width, im.height
    except Exception:
        return None


def _img_infos_for_paragraph(para, doc: Document) -> List[Dict]:
    """
    Vind ALLE afbeeldingen in deze paragraaf door runs te scannen op a:blip/@r:embed.
    Retourneert lijst met dicts: {"url": str, "w": int|None, "h": int|None, "small": bool}
    """
    infos: List[Dict] = []
    for run in para.runs:
        blips = run._r.xpath(".//a:blip")
        if not blips:
            continue
        for blip in blips:
            rId = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if not rId:
                continue
            try:
                part = doc.part.related_parts[rId]
                blob = part.blob
            except Exception:
                continue

            # bepaal grootte
            size = _image_size(blob)
            w = size[0] if size else None
            h = size[1] if size else None
            small = (w is not None and h is not None and w < 100 and h < 100)

            # upload naar Cloudinary of data-uri
            url = _upload_bytes(blob)
            if not url:
                b64 = base64.b64encode(blob).decode("ascii")
                url = f"data:image/png;base64,{b64}"

            infos.append({"url": url, "w": w, "h": h, "small": small})
    return infos


def _is_heading(para) -> int:
    name = (para.style.name or "").lower()
    if name.startswith("heading") or name.startswith("kop"):
        for n in ("1", "2", "3"):
            if n in name:
                return int(n)
        return 1
    return 0


def docx_to_html(file_like) -> str:
    """
    DOCX → HTML:
      - <h1..h3> voor koppen
      - <p> voor alinea's
      - Afbeeldingen:
          * standaard max 300×300 (via CSS inline style)
          * als >=2 kleine (<100×100) in dezelfde paragraaf, dan naast elkaar in een flex-rij
    """
    doc = Document(file_like)
    out = ['<div class="lesson">']

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        level = _is_heading(para)

        if level and text:
            level = min(3, max(1, level))
            out.append(f"<h{level}>{escape(text)}</h{level}>")
        elif text:
            out.append(f"<p>{escape(text)}</p>")

        # afbeeldingen in deze paragraaf
        imgs = _img_infos_for_paragraph(para, doc)
        if not imgs:
            continue

        small_count = sum(1 for i in imgs if i["small"])
        # CASE A: meerdere kleine → flex-rij
        if small_count >= 2 and small_count == len(imgs):
            out.append(
                '<div style="display:flex;gap:8px;flex-wrap:wrap;align-items:flex-start;margin:4px 0;">'
                + "".join(
                    f'<img src="{i["url"]}" alt="" loading="lazy" '
                    f'style="height:100px;max-width:300px;max-height:300px;object-fit:contain;" />'
                    for i in imgs
                )
                + "</div>"
            )
        else:
            # CASE B: normaal per stuk (max 300×300)
            for i in imgs:
                out.append(
                    f'<p><img src="{i["url"]}" alt="" loading="lazy" '
                    f'style="max-width:300px;max-height:300px;object-fit:contain;" /></p>'
                )

    out.append("</div>")
    return "\n".join(out)


