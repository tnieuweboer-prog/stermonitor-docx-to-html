import os
import io
import base64
from html import escape
from typing import Optional, List

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# --- Cloudinary (optioneel, auto als env aanwezig) ---
try:
    import cloudinary
    import cloudinary.uploader
except Exception:
    cloudinary = None  # module niet geïnstalleerd of omgeving zonder internet


def _cloudinary_configured() -> bool:
    """Check of Cloudinary credenties aanwezig zijn en configureer."""
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
            cloudinary.config(
                cloud_name=name, api_key=key, api_secret=secret, secure=True
            )
            return True
        except Exception:
            return False
    return False


def upload_image_bytes_to_cloudinary(
    img_bytes: bytes,
    public_id: Optional[str] = None,
    folder: str = "triade-html",
) -> Optional[str]:
    """Upload bytes → Cloudinary en retourneer secure_url. None bij fout of geen config."""
    if not _cloudinary_configured():
        return None
    try:
        res = cloudinary.uploader.upload(
            img_bytes,
            folder=folder,
            public_id=public_id,
            overwrite=True,
            resource_type="image",
            use_filename=True,
            unique_filename=True,
        )
        return res.get("secure_url") or res.get("url")
    except Exception:
        return None


# --- Extract helpers ---
def _extract_all_image_blobs(doc: Document) -> List[bytes]:
    """Pak alle image blobs uit het document in documentvolgorde van rels."""
    blobs = []
    for rel in doc.part.rels.values():
        if rel.reltype == RT.IMAGE:
            blobs.append(rel.target_part.blob)
    return blobs


def _paragraph_has_image(para) -> bool:
    """Detecteer of paragraaf een afbeelding (drawing) bevat."""
    try:
        for run in para.runs:
            if run._r.xpath(".//w:drawing"):
                return True
    except Exception:
        pass
    return False


# --- HTML converter ---
def docx_to_html(file_like) -> str:
    """
    Simpele DOCX→HTML:
      - Headings → <h1>/<h2>/<h3>
      - Paragrafen → <p>
      - Inline-afbeeldingen → upload naar Cloudinary (indien geconfigureerd) en zet <img src="">
      - Fallback: data URI als Cloudinary niet beschikbaar is
    """
    doc = Document(file_like)

    # verzamel image bytes één keer (we matchen 'next image' voor elke paragraaf met drawing)
    image_blobs = _extract_all_image_blobs(doc)
    img_idx = 0

    parts = ['<div class="lesson">']

    for para in doc.paragraphs:
        text = (para.text or "").strip()
        is_heading = (para.style.name or "").lower().startswith("heading")

        # 1) headings
        if is_heading and text:
            # bepaal level
            level = 1
            try:
                # "Heading 1", "Kop 1", etc.
                name = para.style.name
                for n in ("1", "2", "3"):
                    if n in name:
                        level = int(n)
                        break
            except Exception:
                pass
            level = min(3, max(1, level))
            parts.append(f"<h{level}>{escape(text)}</h{level}>")
            continue

        # 2) normale tekst
        if text:
            parts.append(f"<p>{escape(text)}</p>")

        # 3) inline afbeeldingen in deze paragraaf
        if _paragraph_has_image(para) and img_idx < len(image_blobs):
            # sommige paragrafen kunnen meerdere drawings hebben; we plaatsen er net zoveel als runs met drawing
            drawings_in_para = 0
            for run in para.runs:
                if run._r.xpath(".//w:drawing") and img_idx < len(image_blobs):
                    blob = image_blobs[img_idx]
                    img_idx += 1
                    drawings_in_para += 1

                    url = upload_image_bytes_to_cloudinary(blob)
                    if url:
                        parts.append(f'<p><img src="{url}" alt="" loading="lazy"></p>')
                    else:
                        # fallback: inline base64
                        b64 = base64.b64encode(blob).decode("ascii")
                        parts.append(f'<p><img src="data:image/png;base64,{b64}" alt="" loading="lazy"></p>')

    parts.append("</div>")
    return "\n".join(parts)


