import streamlit as st
import tempfile
from pathlib import Path
from src.extract import extract_text, parse_kara_section, parse_physical_config, parse_enclosure_table, write_excel

st.set_page_config(
    page_title="Soundvision Extractor",
    page_icon="🔊",
    layout="centered"
)

col1, col2 = st.columns([1, 4])
with col1:
    st.image("assets/lasmall.png", width=80)
with col2:
    st.title("Soundvision PDF Extractor")

st.markdown("Upload a Soundvision report PDF and download the data as Excel.")

uploaded_file = st.file_uploader("Upload your Soundvision PDF", type="pdf")

if uploaded_file:
    with st.spinner("Parsing PDF..."):
        try:
            # Write upload to a temp file so pdfplumber can open it
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_file.read())
                tmp_pdf_path = Path(tmp_pdf.name)

            text = extract_text(tmp_pdf_path)
            source_name, kara_block = parse_kara_section(text)
            physical   = parse_physical_config(kara_block)
            enclosures = parse_enclosure_table(kara_block)

            # Write Excel to a temp file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
                tmp_xlsx_path = Path(tmp_xlsx.name)
            write_excel(source_name, physical, enclosures, tmp_xlsx_path)

            st.success(f"✅ Extracted **{len(enclosures)} enclosures** from `{uploaded_file.name}`")

            # Show a preview of physical config
            with st.expander("Physical Configuration Preview"):
                for k, v in physical.items():
                    st.markdown(f"**{k}:** {v}")

            # Download button
            with open(tmp_xlsx_path, "rb") as f:
                xlsx_bytes = f.read()

            output_name = uploaded_file.name.replace(".pdf", ".xlsx")
            st.download_button(
                label="📥 Download Excel",
                data=xlsx_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except ValueError as e:
            st.error(f"❌ {e}")
        except Exception as e:
            st.error(f"❌ Unexpected error: {e}")

st.markdown("---")
st.caption("Built for internal use by Afonso Pires·")
