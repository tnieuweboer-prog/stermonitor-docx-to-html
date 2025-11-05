import streamlit as st
from docx import Document

st.set_page_config(page_title="Stermonitor HTML Converter")

st.title("Stermonitor HTML Converter")
st.write(
    "Upload een Word (.docx) bestand en krijg schone HTML terug die je kunt plakken in Stermonitor → broncode."
)

uploaded = st.file_uploader("Kies een Word-bestand", type=["docx"])


def docx_to_html(file):
    """Leest het Word-bestand en zet de inhoud om naar eenvoudige HTML."""
    doc = Document(file)
    html_parts = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Detecteer koppen (Kop 1, Kop 2, ...)
        if para.style.name.startswith("Heading"):
            try:
                level = int(para.style.name.split()[-1])
            except ValueError:
                level = 2
            html_parts.append(f"<h{level}>{text}</h{level}>")

        # Detecteer opsommingstekens
        elif text.startswith("- "):
            html_parts.append(f"<li>{text[2:]}</li>")

        else:
            html_parts.append(f"<p>{text}</p>")

    # Combineer losse <li> regels in <ul> lijsten
    final_html = []
    in_list = False
    for line in html_parts:
        if line.startswith("<li>") and not in_list:
            final_html.append("<ul>")
            in_list = True
        elif not line.startswith("<li>") and in_list:
            final_html.append("</ul>")
            in_list = False
        final_html.append(line)
    if in_list:
        final_html.append("</ul>")

    return "\n".join(final_html)


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
