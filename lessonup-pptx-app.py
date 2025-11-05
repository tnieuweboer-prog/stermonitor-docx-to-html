import streamlit as st, io, re
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(page_title="DOCX → PowerPoint Converter")
st.title("DOCX → PowerPoint voor LessonUp")

uploaded = st.file_uploader("Upload Word-bestand", type=["docx"])

def extract_images(doc):
    imgs=[]; idx=1
    for rel in doc.part.rels.values():
        if rel.reltype==RT.IMAGE:
            imgs.append((f"img{idx}",rel.target_part.blob)); idx+=1
    return imgs

def is_list(p):
    s=(p.style.name or "").lower()
    return "list" in s or "lijst" in s or "opsom" in s

def apply_text_style(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.name="Arial"
            r.font.size=Pt(16)
            r.font.color.rgb=RGBColor(0,0,0)

def docx_to_pptx(file):
    prs=Presentation()
    doc=Document(file)
    imgs=extract_images(doc); im_i=0
    slide=prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text="Inhoud uit Word"
    if len(slide.placeholders)>1: slide.placeholders[1].text="Geconverteerd voor LessonUp"
    cur=slide
    for p in doc.paragraphs:
        t=p.text.strip(); has_img=any("graphic" in r._element.xml for r in p.runs)
        if not t and not has_img: continue
        if p.style.name.startswith("Heading"):
            cur=prs.slides.add_slide(prs.slide_layouts[1])
            cur.shapes.title.text=t; continue
        if has_img:
            s=prs.slides.add_slide(prs.slide_layouts[5])
            s.shapes.title.text="Afbeelding"
            if im_i<len(imgs):
                _,b=imgs[im_i]; im_i+=1
                s.shapes.add_picture(io.BytesIO(b),Inches(1),Inches(1.2),width=Inches(6))
            continue
        body=[sh for sh in cur.shapes if hasattr(sh,"text_frame")]
        if not body: body=prs.slides.add_slide(prs.slide_layouts[1]).shapes.placeholders[1].text_frame
        else: body=body[0].text_frame
        if is_list(p):
            para=body.add_paragraph(); para.text=t; para.level=0
        else:
            if body.text=="": body.text=t
            else:
                para=body.add_paragraph(); para.text=t; para.level=0
        apply_text_style(body)
    out=io.BytesIO(); prs.save(out); out.seek(0); return out

if uploaded:
    pptx_bytes=docx_to_pptx(uploaded)
    st.download_button("Download PowerPoint (.pptx)",data=pptx_bytes,file_name="lessonup_import.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
