import streamlit as st
import tempfile
import pandas as pd
from pathlib import Path
from src.extract import extract_text, extract_metadata, parse_document, write_excel, write_pdf

st.set_page_config(
    page_title="Soundvision Report Data Extractor",
    page_icon="🔊",
    layout="centered"
)

st.markdown("""
<style>
.step-bar {
    display: flex; align-items: center; justify-content: center;
    gap: 0; margin: 1.5rem 0 2rem 0;
}
.step { display: flex; flex-direction: column; align-items: center; gap: 6px; }
.step-circle {
    width: 36px; height: 36px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-weight: 600; font-size: 14px;
    border: 2px solid #3A3F4B; color: #3A3F4B; background: transparent;
}
.step-circle.active { background: #F5A623; border-color: #F5A623; color: #1C1F26; }
.step-circle.done   { background: #3A3F4B; border-color: #3A3F4B; color: #F5A623; }
.step-label { font-size: 11px; color: #888; text-align: center; }
.step-label.active  { color: #F5A623; font-weight: 600; }
.step-label.done    { color: #aaa; }
.step-line { height: 2px; width: 60px; background: #3A3F4B; margin-bottom: 22px; }
.step-line.done { background: #F5A623; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 5], vertical_alignment="center")
with col1:
    st.image("assets/lasmall.png", use_container_width=True)
with col2:
    st.markdown("""
    <div style="padding-left: 8px;">
        <p style="font-size: 36px; font-weight: 700; margin: 0; line-height: 1.2; color: inherit;">Soundvision Report Data Extractor</p>
        <p style="font-size: 14px; color: #888; margin: 6px 0 0 0;">L-Acoustics system data, ready to share.</p>
    </div>
    """, unsafe_allow_html=True)

# ── How it works expander ────────────────────────────────────────────────────
with st.expander("ℹ️ How does it work?"):
    st.markdown("""
    **Soundvision Report Data Extractor** parses L-Acoustics Soundvision PDF exports and generates formatted Excel and PDF reports — ready to share with your team.

    **What it does:**
    - Extracts all flown arrays from the report — physical configuration, enclosure geometry and motor loads
    - Deduplicates L/R mirror arrays automatically — only one side is shown
    - Generates a structured Excel with one sheet per group, plus a **Report Info** page for you to fill in circuit assignments, amp IDs and channels
    - Generates a PDF report with the same data, formatted for printing or sharing

    **What you need to do in the Excel:**
    - Fill in System Engineer, Company, Venue and Date on the Report Info sheet
    - Assign circuits (A–J) to each enclosure — colours follow the resistor colour code
    - Enter Amp ID L/R and Amp Channel — these update automatically across all group sheets

    > 📌 Supported speakers: K1, K2, K3, KARA II, KIVA II, KS28, SB28, X8 and mixed arrays.
    """)

# ── Session state ─────────────────────────────────────────────────────────────
for key, default in [("groups", None), ("file_name", ""), ("report_name", ""), ("report_date", "")]:
    if key not in st.session_state:
        st.session_state[key] = default

current_step = 1 if st.session_state.groups is None else 3

def sc(n):
    if n < current_step: return "done"
    if n == current_step: return "active"
    return ""

st.markdown(f"""
<div class="step-bar">
  <div class="step">
    <div class="step-circle {sc(1)}">{'✓' if current_step > 1 else '1'}</div>
    <div class="step-label {sc(1)}">Upload</div>
  </div>
  <div class="step-line {'done' if current_step > 1 else ''}"></div>
  <div class="step">
    <div class="step-circle {sc(2)}">{'✓' if current_step > 2 else '2'}</div>
    <div class="step-label {sc(2)}">Preview</div>
  </div>
  <div class="step-line {'done' if current_step > 2 else ''}"></div>
  <div class="step">
    <div class="step-circle {sc(3)}">3</div>
    <div class="step-label {sc(3)}">Download</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Upload ────────────────────────────────────────────────────────────────────
if st.session_state.groups is None:
    uploaded_file = st.file_uploader("Upload your Soundvision PDF", type="pdf", label_visibility="collapsed")
    st.caption("📄 Only Soundvision PDF exports are supported.")
else:
    uploaded_file = None

if uploaded_file and st.session_state.groups is None:
    with st.status("Parsing PDF...", expanded=True) as status:
        st.write("Extracting text...")
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.read())
            tmp_pdf_path = Path(tmp_pdf.name)
        try:
            text = extract_text(tmp_pdf_path)
            st.write("Identifying groups and sources...")
            groups = parse_document(text)
            total = sum(len(s) for s in groups.values())
            if total == 0:
                status.update(label="No flown sources found.", state="error")
                st.warning("⚠️ No flown sources found.")
                st.stop()
            st.write(f"Found **{len(groups)} group(s)** and **{total} source(s)**.")
            name, date = extract_metadata(text)
            st.session_state.groups = groups
            st.session_state.file_name = uploaded_file.name
            st.session_state.report_name = name
            st.session_state.report_date = date
            status.update(label="Done!", state="complete")
            st.rerun()
        except Exception as e:
            status.update(label="Error", state="error")
            st.error(f"❌ {e}")
            st.stop()

# ── Download + Preview ────────────────────────────────────────────────────────
if st.session_state.groups:
    groups = st.session_state.groups
    file_name = st.session_state.file_name

    total = sum(len(s) for s in groups.values())
    st.success(f"✅ **{total} source(s)** extracted from **{len(groups)} group(s)**")
    st.caption(f"📁 File: **{st.session_state.report_name}** · {st.session_state.report_date}")

    # Generate files
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
        tmp_xlsx_path = Path(tmp_xlsx.name)
    write_excel(groups, tmp_xlsx_path,
                report_name=st.session_state.report_name,
                report_date=st.session_state.report_date)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_report:
        tmp_report_path = Path(tmp_report.name)
    write_pdf(groups, tmp_report_path,
              report_name=st.session_state.report_name,
              report_date=st.session_state.report_date)

    # Download buttons — Excel green, PDF amber
    st.markdown("""
    <style>
    div[data-testid="stDownloadButton"]:nth-of-type(1) button {
        background-color: #1a3a2a; border-color: #2d6a4f; color: #52b788;
    }
    div[data-testid="stDownloadButton"]:nth-of-type(1) button:hover {
        background-color: #2d6a4f; color: #d8f3dc;
    }
    div[data-testid="stDownloadButton"]:nth-of-type(2) button {
        background-color: #2a1a0a; border-color: #B7791F; color: #F5A623;
    }
    div[data-testid="stDownloadButton"]:nth-of-type(2) button:hover {
        background-color: #744210; color: #fefcbf;
    }
    </style>
    """, unsafe_allow_html=True)

    dl1, dl2 = st.columns(2)
    with open(tmp_xlsx_path, "rb") as f:
        xlsx_bytes = f.read()
    with dl1:
        st.download_button(
            label="📥 Download Excel",
            data=xlsx_bytes,
            file_name=file_name.replace(".pdf", ".xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with open(tmp_report_path, "rb") as f:
        pdf_bytes = f.read()
    with dl2:
        st.download_button(
            label="📄 Download PDF Report",
            data=pdf_bytes,
            file_name=file_name.replace(".pdf", "_report.pdf"),
            mime="application/pdf",
            use_container_width=True
        )

    st.markdown("---")

    # Preview below
    for group_name, sources in groups.items():
        st.markdown(f"#### {group_name}")
        for source in sources:
            with st.expander(f"**{source['name']}** — {len(source['enclosures'])} enclosures"):
                if source["physical"]:
                    st.markdown("**Physical Configuration**")
                    phys_df = pd.DataFrame(list(source["physical"].items()), columns=["Field", "Value"])
                    st.dataframe(phys_df, use_container_width=True, hide_index=True)
                if source["enclosures"]:
                    st.markdown("**Per-Enclosure Geometry**")
                    enc_df = pd.DataFrame(source["enclosures"])
                    st.dataframe(enc_df, use_container_width=True, hide_index=True)

    if st.button("🔄 Upload a new file", use_container_width=True):
        st.session_state.groups = None
        st.session_state.file_name = ""
        st.rerun()

st.markdown("---")
st.caption("ℹ️ Arrays L/R are mirrored — only one side is extracted per group.")
st.markdown("<p style='font-size:14px; color:#888; text-align:center; margin-top: 48px; margin-bottom: 16px;'>Built for internal purposes · Afonso Pires</p>", unsafe_allow_html=True)
