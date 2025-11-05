import streamlit as st
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import cloudinary
import cloudinary.uploader
import re

st.set_page_config(page_title="DOCX → Stermonitor HTML Converter")
st.title("DOCX → Stermonitor / LessonUp HTML Converter")

platform = st.selectbox("Kies platform", ["Stermonitor", "LessonUp"])
uploaded = st.file_uploader("Upload Word-bestand", type=["docx"])

# --- Cloudinary-config ---
required = ["CLOUDINARY_CLOUD_NAME","CLOUDINARY_API_KEY","CLOUDINARY_API_SECRET"]
missing = [k for k in required if k not in st.secrets]
if missing:
    st.warning("Vul Cloudinary API-gegevens in via Streamlit → Edit secrets")
else:
    cloudinary.config(
        cloud_name=st.secrets["CLOUDINARY_CLOUD_NAME"],
        api_key=st.secrets["CLOUDINARY_API_KEY"],
        api_secret=st.secrets["CLOUDINARY_API_SECRET"],
        secure=True,
    )

# --- Helpers ---
def extract_images(doc):
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    imgs=[]
    for rel in doc.part.rels.values():
        if rel.reltype==RT.IMAGE:
            imgs.append(rel.target_part.blob)
    return imgs

def upload_to_cloudinary(data):
    try:
        r=cloudinary.uploader.upload(data,folder="ster_monitor",resource_type="image")
        return r["secure_url"]
    except: return None

def is_list(p): 
    s=(p.style.name or "").lower()
    return "list" in s or "lijst" in s or "opsom" in s

def runs_to_html(p):
    out=[]
    for r in p.runs:
        t=r.text.strip()
        if not t: continue
        if r.bold: out.append(f"<strong>{t}</strong>")
        else: out.append(t)
    return " ".join(out)

# --- Converter ---
def docx_to_html(file,platform="Stermonitor"):
    doc=Document(file)
    imgs=extract_images(doc)
    urls=[upload_to_cloudinary(b) for b in imgs]
    i=0; html=[]; buf=""; in_list=False
    for p in doc.paragraphs:
        txt=p.text.strip()
        if p.style.name.startswith("Heading"):
            if buf: html.append(f"<p>{buf}</p>"); buf=""
            if in_list: html.append("</ul>"); in_list=False
            level=int(p.style.name.split()[-1]) if p.style.name.split()[-1].isdigit() else 2
            html.append(f"<h{level}>{txt}</h{level}>"); continue
        has_img=any("graphic" in r._element.xml for r in p.runs)
        if has_img:
            if buf: html.append(f"<p>{buf}</p>"); buf=""
            if in_list: html.append("</ul>"); in_list=False
            url=urls[i] if i<len(urls) else None; i+=1
            if url:
                if platform=="Stermonitor":
                    html.append(f'<p><img src="{url}" style="width:300px;height:300px;object-fit:cover;border:1px solid #ccc;border-radius:8px;padding:4px;"></p>')
                else:
                    html.append(f'<p><img src="{url}"></p>')
            continue
        if is_list(p):
            if buf: html.append(f"<p>{buf}</p>"); buf=""
            if not in_list: html.append('<ul class="browser-default">' if platform=="Stermonitor" else "<ul>"); in_list=True
            html.append(f"<li>{runs_to_html(p)}</li>"); continue
        if txt:
            if in_list: html.append("</ul>"); in_list=False
            buf+=" "+runs_to_html(p)
            if re.search(r"[.!?]$",txt): html.append(f"<p>{buf}</p>"); buf=""
    if buf: html.append(f"<p>{buf}</p>")
    if in_list: html.append("</ul>")
    return "\n".join(html)

# --- UI ---
if uploaded:
    html=docx_to_html(uploaded,platform)
    st.subheader(f"HTML voor {platform}")
    st.code(html,language="html")
    st.download_button("Download HTML",data=html,file_name="ster_monitor.html",mime="text/html")


