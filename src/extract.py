"""
Soundvision PDF Extractor
Reads all PDFs from ../data/ and outputs Excel files to ../output/

Usage:
    python extract.py                  # processes all PDFs in ../data/
    python extract.py report.pdf       # processes a specific file in ../data/
"""

import sys
import re
import pdfplumber
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Paths ─────────────────────────────────────────────────────────────────────

SRC_DIR    = Path(__file__).parent
DATA_DIR   = SRC_DIR.parent / "data"
OUTPUT_DIR = SRC_DIR.parent / "output"

# ── Parsing ───────────────────────────────────────────────────────────────────

def extract_text(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def parse_kara_section(text):
    match = re.search(
        r"(1\. Source: KARA [LR].*?)(?=2\. Source: KARA|2\. Group:|3\. Group:|$)",
        text, re.DOTALL
    )
    if not match:
        raise ValueError("Could not locate KARA source block in PDF.")
    block = match.group(1)
    source_name = re.search(r"1\. Source: (KARA [LR])", block).group(1)
    return source_name, block


def parse_physical_config(block):
    fields = {
        "Configuration":     r"Configuration:\s*(.+)",
        "Bumper":            r"Bumper:\s*(.+)",
        "# Motors":          r"# motors:\s*(\d+)",
        "Position X (m)":    r"Position \(X; Y; Z, m\):\s*([\-\d.]+);",
        "Position Y (m)":    r"Position \(X; Y; Z, m\):\s*[\-\d.]+;\s*([\-\d.]+);",
        "Position Z (m)":    r"Position \(X; Y; Z, m\):\s*[\-\d.]+;\s*[\-\d.]+;\s*([\-\d.]+)",
        "Site (°)":          r"Site:\s*([\-\d.]+)\s*°",
        "Azimuth (°)":       r"Azimuth:\s*([\-\d.]+)\s*°",
        "Bottom Elev. (m)":  r"Bottom elevation:\s*([\-\d.]+)",
        "Top Site (°)":      r"Top site:\s*([\-\d.]+)\s*°",
        "Bottom Site (°)":   r"Bottom site:\s*([\-\d.]+)\s*°",
        "Total Weight (kg)": r"Total weight \(Enclosures \+ Frames\):\s*([\d.]+)",
        "Enclosure Wt (kg)": r"Total enclosure weight:\s*([\d.]+)",
        "Front Motor (kg)":  r"Front motor load:\s*([\d.]+)",
        "Rear Motor (kg)":   r"Rear motor load:\s*([\d.]+)",
    }
    config = {}
    for key, pattern in fields.items():
        m = re.search(pattern, block)
        config[key] = m.group(1).strip() if m else "N/A"
    return config


def parse_enclosure_table(block):
    pattern = re.compile(
        r"#(\d+)\s+(KARA\s+II)\s+([\-\d.]+)\s+([\-\d.]+)\s+([\-\d.]+)\s+([\-\d.]+)\s+([\d/]+)"
    )
    return [
        {
            "Enclosure #":  int(m.group(1)),
            "Type":         m.group(2),
            "Angle (°)":    float(m.group(3)),
            "Site (°)":     float(m.group(4)),
            "Top Z (m)":    float(m.group(5)),
            "Bottom Z (m)": float(m.group(6)),
            "Panflex":      m.group(7),
        }
        for m in pattern.finditer(block)
    ]


# ── Styling helpers ───────────────────────────────────────────────────────────

def thin_border():
    s = Side(style="thin", color="B0B0B0")
    return Border(left=s, right=s, top=s, bottom=s)

HEADER_FILL  = PatternFill("solid", start_color="1F3864")
SUBHEAD_FILL = PatternFill("solid", start_color="2E75B6")
ALT_FILL     = PatternFill("solid", start_color="DCE6F1")
WHITE_FILL   = PatternFill("solid", start_color="FFFFFF")
ROW_FILL     = PatternFill("solid", start_color="4472C4")
HEADER_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=13)
SUBHEAD_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
LABEL_FONT   = Font(name="Arial", bold=True, color="1F3864", size=10)
BODY_FONT    = Font(name="Arial", size=10)
CENTER       = Alignment(horizontal="center", vertical="center")
LEFT         = Alignment(horizontal="left",   vertical="center")


# ── Excel writing ─────────────────────────────────────────────────────────────

def write_excel(source_name, physical, enclosures, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "KARA MAINS"
    ws.sheet_view.showGridLines = False

    row = 1

    # Title
    ws.merge_cells(f"A{row}:F{row}")
    c = ws[f"A{row}"]
    c.value = f"Soundvision Report — KARA MAINS"
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = CENTER
    ws.row_dimensions[row].height = 28
    row += 2

    # Physical config section header
    ws.merge_cells(f"A{row}:F{row}")
    c = ws[f"A{row}"]
    c.value = "Physical Configuration"
    c.font = SUBHEAD_FONT
    c.fill = SUBHEAD_FILL
    c.alignment = LEFT
    ws.row_dimensions[row].height = 20
    row += 1

    # Key-value pairs in two columns
    items = list(physical.items())
    for i in range(0, len(items), 2):
        fill = ALT_FILL if (row % 2 == 0) else WHITE_FILL
        for offset, idx in enumerate([i, i + 1]):
            if idx >= len(items):
                break
            key, val = items[idx]
            col_label = offset * 3 + 1

            lc = ws.cell(row=row, column=col_label, value=key)
            lc.font = LABEL_FONT
            lc.alignment = LEFT
            lc.fill = fill
            lc.border = thin_border()

            vc = ws.cell(row=row, column=col_label + 1, value=val)
            vc.font = BODY_FONT
            vc.alignment = LEFT
            vc.fill = fill
            vc.border = thin_border()

            ws.merge_cells(start_row=row, start_column=col_label + 1,
                           end_row=row, end_column=col_label + 2)
        row += 1

    row += 1  # spacer

    # Enclosure geometry section header
    ws.merge_cells(f"A{row}:G{row}")
    c = ws[f"A{row}"]
    c.value = "Per-Enclosure Geometry"
    c.font = SUBHEAD_FONT
    c.fill = SUBHEAD_FILL
    c.alignment = LEFT
    ws.row_dimensions[row].height = 20
    row += 1

    # Table headers
    headers = ["Enclosure #", "Type", "Angle (°)", "Site (°)",
               "Top Z (m)", "Bottom Z (m)", "Panflex"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = ROW_FILL
        c.alignment = CENTER
        c.border = thin_border()
    ws.row_dimensions[row].height = 18
    row += 1

    # Enclosure rows
    for enc in enclosures:
        fill = ALT_FILL if enc["Enclosure #"] % 2 == 0 else WHITE_FILL
        for col, key in enumerate(headers, 1):
            c = ws.cell(row=row, column=col, value=enc[key])
            c.font = BODY_FONT
            c.alignment = CENTER
            c.fill = fill
            c.border = thin_border()
        row += 1

    # Column widths
    for col, width in zip("ABCDEFG", [20, 16, 4, 20, 16, 4, 10]):
        ws.column_dimensions[col].width = width

    wb.save(output_path)


# ── Runner ────────────────────────────────────────────────────────────────────

def process_pdf(pdf_path):
    print(f"  Processing: {pdf_path.name}")
    text = extract_text(pdf_path)
    source_name, kara_block = parse_kara_section(text)
    physical   = parse_physical_config(kara_block)
    enclosures = parse_enclosure_table(kara_block)
    output_path = OUTPUT_DIR / (pdf_path.stem + ".xlsx")
    write_excel(source_name, physical, enclosures, output_path)
    print(f"  Saved:      {output_path.name}  ({len(enclosures)} enclosures)")


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if len(sys.argv) > 1:
        # Specific file passed as argument
        pdf_path = DATA_DIR / sys.argv[1]
        if not pdf_path.exists():
            print(f"Error: {pdf_path} not found.")
            sys.exit(1)
        pdfs = [pdf_path]
    else:
        # Process all PDFs in data/
        pdfs = sorted(DATA_DIR.glob("*.pdf"))
        if not pdfs:
            print(f"No PDF files found in {DATA_DIR}")
            sys.exit(0)

    print(f"Found {len(pdfs)} PDF(s) to process.\n")
    for pdf in pdfs:
        process_pdf(pdf)
    print("\nDone.")


if __name__ == "__main__":
    main()
