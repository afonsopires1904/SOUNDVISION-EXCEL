import streamlit as st
import tempfile
from pathlib import Path
from src.extract import extract_text, parse_document, write_excel, write_pdf

st.set_page_config(
    page_title="Soundvision Extractor",
    page_icon="🔊",
    layout="centered"
)

col1, col2 = st.columns([1, 6])
with col1:
    st.image("assets/logo.png", width=60)
with col2:
    st.title("Soundvision PDF Extractor")

st.markdown("Upload a Soundvision report PDF and download the extracted data.")

uploaded_file = st.file_uploader("Upload your Soundvision PDF", type="pdf")

if uploaded_file:
    with st.spinner("Parsing PDF..."):
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_file.read())
                tmp_pdf_path = Path(tmp_pdf.name)

            text   = extract_text(tmp_pdf_path)
            groups = parse_document(text)

            total_sources = sum(len(s) for s in groups.values())
            st.success(f"✅ Found **{len(groups)} group(s)** and **{total_sources} source(s)**")

            # Preview
            with st.expander("Preview extracted data"):
                for group_name, sources in groups.items():
                    st.markdown(f"**Group: {group_name}**")
                    for source in sources:
                        st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;• {source['name']} — {len(source['enclosures'])} enclosures")

            # Generate Excel
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
                tmp_xlsx_path = Path(tmp_xlsx.name)
            write_excel(groups, tmp_xlsx_path)

            # Generate PDF
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_report:
                tmp_report_path = Path(tmp_report.name)
            write_pdf(groups, tmp_report_path)

            # Download buttons
            col1, col2 = st.columns(2)

            with open(tmp_xlsx_path, "rb") as f:
                xlsx_bytes = f.read()
            with col1:
                st.download_button(
                    label="📥 Download Excel",
                    data=xlsx_bytes,
                    file_name=uploaded_file.name.replace(".pdf", ".xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with open(tmp_report_path, "rb") as f:
                pdf_bytes = f.read()
            with col2:
                st.download_button(
                    label="📄 Download PDF Report",
                    data=pdf_bytes,
                    file_name=uploaded_file.name.replace(".pdf", "_report.pdf"),
                    mime="application/pdf",
                    use_container_width=True
                )

        except ValueError as e:
            st.error(f"❌ {e}")
        except Exception as e:
            st.error(f"❌ Unexpected error: {e}")

st.markdown("---")
st.caption("Built for internal purposes by Afonso Pires .")