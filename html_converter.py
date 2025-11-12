import os
import base64
from html import escape
from typing import Optional, List

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT

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


def _img_urls_for_paragraph(para, doc: Document) -> List[str]:
    """
    Vind ALLE afbeeldingen in deze paragraaf door runs te scannen op a:blip/@r:embed.
    """
    urls: List[str] = []
    for run in para.runs:
        # zoek a:blip nodes
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

            # eerst proberen naar Cloudinary
            url = _upload_bytes(blob)
            if url:
                urls.append(url)
            else:
                # nette fallback: data-uri
                b64 = base64.b64encode(blob).decode("ascii")
                urls.append(f"data:image/png;base64,{b64}")
    return urls


def _is_heading(para) -> int:
    name = (para.style.name or "").lower()
    if name.startswith("heading") or name.startswith("kop"):
        # probeer level uit naam te halen (1/2/3)
        for n in ("1", "2", "3"):
            if n in name:
                return int(n)
        return 1
    return 0


def docx_to_html(file_like) -> str:
    """
    DOCX → HTML met:
      - <h1..h3> voor koppen
      - <p> voor alinea's
      - <img> voor inline-afbeeldingen (Cloudinary of data-uri)
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

        # afbeeldingen in deze paragraaf (precies gekoppeld aan de runs)
        img_urls = _img_urls_for_paragraph(para, doc)
        for url in img_urls:
            out.append(f'<p><img src="{url}" alt="" loading="lazy"></p>')

    out.append("</div>")
    return "\n".join(out)



