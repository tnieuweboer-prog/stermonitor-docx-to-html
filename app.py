import streamlit as st
from docx import Document
import re

st.set_page_config(page_title="Stermonitor HTML Converter")

st.title("Stermonitor HTML Converter")
st.write(
    "Upload een Word (.docx) bestand. De converter maakt op basis van zinnen en koppen nette HTML voor Stermonitor."
)

uploaded = st.file_uploader("Kies een Word-bestand", type=["docx"])


def docx_to_html(file):
    doc = Document(file)
    html_parts = []
    buffer = ""  # tijdelijke tekstbuffer

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # koppen behouden als aparte blokken
        if para.style.name.startswith("Heading"):
            # eerst buffer wegschrijven
            if buffer:
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""
            try:
                level = int(para.style.name.split()[-1])
            except ValueError:
                level = 2
            html_parts.append(f"<h{level}>{text}</h{level}>")
        else:
            # gewone tekst toevoegen aan buffer
            buffer += " " + text
            # check of de regel eindigt met een punt, vraagteken of uitroepteken
            if re.search(r"[.!?]$", text):
                html_parts.append(f"<p>{buffer.strip()}</p>")
                buffer = ""

    # restbuffer wegschrijven
    if buffer:
        html_parts.append(f"<p>{buffer.strip()}</p>")

    return "\n".join(html_parts)


if uploaded:
    html_output = docx_to_html(uploaded)

    st.subheader("Gegenereerde HTML-code")
    st.code(html_output, language="html")

    st.download_button(
        label="Download HTML-bestand",
        data=html_output,
        file_name="ster_monitor.html",
        mime="text/html",
    )
else:
    st.info("Upload hierboven een .docx-bestand om te beginnen.")

