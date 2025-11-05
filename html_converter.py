# html_converter.py
import re
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT

def extract_images(doc):
    images=[]
    idx=1
    for rel in doc.part.rels.values():
        if rel.reltype==RT.IMAGE:
            images.append((f"image_{idx}.{rel.target_part.partname.ext}", rel.target_part.blob))
            idx+=1
    return images

def is_word_list_paragraph(para):
    name=(para.style.name or "").lower()
    if "list" in name or "lijst" in name or "opsom" in name:
        return True
    ppr = para._p.pPr
    return ppr is not None and ppr.numPr is not None

def runs_to_html(para):
    parts=[]
    for run in para.runs:
        t = run.text.strip()
        if not t: continue
        parts.append(f"<strong>{t}</strong>" if run.bold else t)
    return " ".join(parts)

def docx_to_html(file_like, image_uploader=None, platform="Stermonitor"):
    """
    file_like: file object or path acceptable for python-docx Document()
    image_uploader: optional callable(filename, bytes) -> url (if None images ignored)
    platform: "Stermonitor" or "LessonUp"
    """
    doc = Document(file_like)
    images = extract_images(doc)
    image_urls = []
    if image_uploader:
        for fn, blob in images:
            image_urls.append(image_uploader(fn, blob))
    else:
        image_urls = [None]*len(images)

    html=[]
    buffer=""
    in_list=False
    img_i=0
    first_bold=False

    for para in doc.paragraphs:
        text=(para.text or "").strip()
        if para.style.name.startswith("Heading"):
            if buffer:
                html.append(f"<p>{buffer.strip()}</p>"); buffer=""
            if in_list: html.append("</ul>"); in_list=False
            try: lvl=int(para.style.name.split()[-1])
            except: lvl=2
            html.append(f"<h{lvl}>{text}</h{lvl}>")
            continue

        has_image = any("graphic" in r._element.xml for r in para.runs)
        if has_image:
            if buffer:
                html.append(f"<p>{buffer.strip()}</p>"); buffer=""
            if in_list:
                html.append("</ul>"); in_list=False
            url = image_urls[img_i] if img_i < len(image_urls) else None
            img_i += 1
            if url:
                if platform=="Stermonitor":
                    style='style="width:300px;height:300px;object-fit:cover;border:1px solid #ccc;border-radius:8px;padding:4px;"'
                    html.append(f'<p><img src="{url}" alt="afbeelding" {style}></p>')
                else:
                    html.append(f'<p><img src="{url}" alt="afbeelding"></p>')
            else:
                html.append("<p>[afbeelding niet geüpload]</p>")
            continue

        if is_word_list_paragraph(para):
            if buffer: html.append(f"<p>{buffer.strip()}</p>"); buffer=""
            if not in_list:
                html.append('<ul class="browser-default">' if platform=="Stermonitor" else "<ul>")
                in_list=True
            html.append(f"<li>{runs_to_html(para)}</li>")
            continue

        if any(r.bold for r in para.runs) and not is_word_list_paragraph(para):
            if buffer: html.append(f"<p>{buffer.strip()}</p>"); buffer=""
            if in_list: html.append("</ul>"); in_list=False
            if platform=="Stermonitor" and first_bold: html.append("<br>")
            html.append(f"<p>{runs_to_html(para)}</p>")
            first_bold=True
            continue

        if text:
            if in_list: html.append("</ul>"); in_list=False
            buffer += " " + runs_to_html(para)
            if re.search(r"[.!?]$", text):
                html.append(f"<p>{buffer.strip()}</p>"); buffer=""

    if buffer: html.append(f"<p>{buffer.strip()}</p>")
    if in_list: html.append("</ul>")
    return "\n".join(html)

