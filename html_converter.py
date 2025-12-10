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
        except:
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
        except:
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
    except:
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
    except:
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
            except:
                continue

            size = _image_size(blob)
            w = size[0] if size else None
            h = size[1] if size else None
            small = (w and h and w < 100 and h < 100)

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
    """ DOCX → HTML met 1 overkoepelende groene div. """

    doc = Document(file_like)

    out = [
        "<html>",
        "<head>",
        "<style>",

        "body { margin: 0; padding: 0; }",

        ".green {",
        "    background-image: url('YOUR_ASSET_URL_HERE');",
        "    background-size: cover;",
        "    background-repeat: no-repeat;",
        "    background-position: center;",
        "}",

        ".lesson {",
        "    max-width: 900px;",
        "    margin: 0;",
        "    padding: 1rem;",
        "    font-family: Arial, sans-serif;",
        "    text-align: left;",
        "    background: rgba(198,217,170,0.6);",  # hele groene achtergrond
        "    backdrop-filter: blur(2px);",
        "    border-radius: 6px;",
        "}",

        "</style>",
        "</head>",

        "<body class='green'>",

        # 👇 DIT is nu jouw volledige groene container
        "<div class='lesson light-green'>"
    ]

    # Verwerking tekst + afbeeldingen
    for para in doc.paragraphs:
        text = (para.text or "").strip()
        level = _is_heading(para)

        # Koppen blijven gewoon koppen
        if level and text:
            out.append(f"<h{min(level,3)}>{escape(text)}</h{min(level,3)}>")

        # Paragrafen worden normale <p>
        elif text:
            out.append(f"<p>{escape(text)}</p>")

        # Afbeeldingen
        imgs = _img_infos_for_paragraph(para, doc)
        if not imgs:
            continue

        small = [i for i in imgs if i["small"]]
        big = [i for i in imgs if not i["small"]]

        if small:
            out.append(
                '<div style="display:flex;gap:8px;flex-wrap:wrap;margin:4px 0;">'
            )
            for i in small:
                out.append(
                    f'<img src="{i["url"]}" alt="" '
                    f'style="max-width:{i["w"] or 100}px;max-height:{i["h"] or 100}px;object-fit:contain;" />'
                )
            out.append("</div>")

        for i in big:
            out.append(
                f'<p><img src="{i["url"]}" alt="" '
                f'style="max-width:300px;max-height:300px;object-fit:contain;" /></p>'
            )

    out.append("</div>")
    out.append("</body>")
    out.append("</html>")

    return "\n".join(out)



