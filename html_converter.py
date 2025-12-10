import os
import base64
from html import escape
from typing import Optional, List, Dict
from docx import Document

# Pillow voor beeldmaten
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


# ---------- Cloudinary Config ----------
def _cloudinary_ready() -> bool:
    """Check of Cloudinary correct is geconfigureerd via env of URL."""
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
                cloud_name=name,
                api_key=key,
                api_secret=secret,
                secure=True
            )
            return True
        except Exception:
            return False

    return False


def _upload_bytes(img_bytes: bytes, folder="triade-html") -> Optional[str]:
    """Upload naar Cloudinary, retourneer secure_url of None bij fout."""
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


# ---------- Hulpfuncties ----------
def _image_size(img_bytes: bytes) -> Optional[tuple]:
    """Bepaal (breedte, hoogte) van afbeelding met Pillow."""
    if not PIL_OK:
        return None

    try:
        from io import BytesIO
        with Image.open(BytesIO(img_bytes)) as im:
            return im.width, im.height
    except Exception:
        return None


def _img_infos_for_paragraph(para, doc: Document) -> List[Dict]:
    """Zoek alle afbeeldingen in paragraaf en retourneer info."""
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

            size = _image_size(blob)
            w = size[0] if size else None
            h = size[1] if size else None
            small = (w is not None and h is not None and w < 100 and h < 100)

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


# ---------- Hoofdconverter ----------
def docx_to_html(file_like) -> str:
    """
    DOCX → HTML met:
      • Koppen als <h1..h3>
      • Paragrafen als <p>
      • Afbeeldingen:
          - Kleine (<100×100) → naast elkaar
          - Grotere ≥100×100 → max 300×300
    """

    doc = Document(file_like)

    # HTML + CSS inclusief achtergrond via class="green"
    out = [
        "<html>",
        "<head>",
        "<style>",

        # Body reset
        "body { margin: 0; padding: 0; }",

        # Achtergrondklasse voor Stermonitor
        ".green {",
        "    background-image: url('YOUR_ASSET_URL_HERE');",
        "    background-size: cover;",
        "    background-repeat: no-repeat;",
        "    background-position: center;",
        "}",

        # Les-content container
        ".lesson {",
        "    max-width: 900px;",
        "    margin: 0;",
        "    padding: 1rem;",
        "    font-family: Arial, sans-serif;",
        "    text-align: left;",
        "    background: rgba(255,255,255,0.6);",
        "    backdrop-filter: blur(2px);",
        "}",

        "</style>",
        "</head>",

        # 👇 IMPORTANT: body krijgt de klasse "green"
        "<body class='green'>",

        "<div class='lesson'>"
    ]

    # ——— Tekst en afbeeldingen verwerken ———
    for para in doc.paragraphs:
        text = (para.text or "").strip()
        level = _is_heading(para)

        if level and text:
            out.append(f"<h{min(level,3)}>{escape(text)}</h{min(level,3)}>")
        elif text:
            out.append(f"<p>{escape(text)}</p>")

        imgs = _img_infos_for_paragraph(para, doc)
        if not imgs:
            continue

        small_imgs = [i for i in imgs if i["small"]]
        big_imgs = [i for i in imgs if not i["small"]]

        # Kleine afbeeldingen naast elkaar
        if small_imgs:
            out.append(
                '<div style="display:flex;gap:8px;flex-wrap:wrap;align-items:flex-start;margin:4px 0;">'
            )
            for i in small_imgs:
                out.append(
                    f'<img src="{i["url"]}" alt="" loading="lazy" '
                    f'style="max-width:{i["w"] or 100}px;max-height:{i["h"] or 100}px;object-fit:contain;" />'
                )
            out.append("</div>")

        # Grote afbeeldingen onder elkaar
        for i in big_imgs:
            out.append(
                f'<p><img src="{i["url"]}" alt="" loading="lazy" '
                f'style="max-width:300px;max-height:300px;object-fit:contain;" /></p>'
            )

    # HTML afsluiten
    out.append("</div>")
    out.append("</body>")
    out.append("</html>")

    return "\n".join(out)


