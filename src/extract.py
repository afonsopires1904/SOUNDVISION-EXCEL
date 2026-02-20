"""
Soundvision PDF Extractor — Generic
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

# ── Text extraction ───────────────────────────────────────────────────────────

def extract_text(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


# ── Generic document parsing ──────────────────────────────────────────────────

def get_canonical_name(source_name):
    name = re.sub(r'\s+[LR]\s*(\d*)$', lambda m: (' ' + m.group(1)) if m.group(1) else '', source_name).strip()
    return name


def is_mirror(name):
    return bool(re.search(r'\s+R(\s+\d+)?$', name))


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
        if m:
            config[key] = m.group(1).strip()
    return config


def parse_enclosure_table(block):
    has_panflex = bool(re.search(r"Panflex", block))
    has_angles  = bool(re.search(r"Angles \(°\)", block))

    if has_angles and has_panflex:
        # Line array with Panflex (KARA, K1, K2, K3...)
        pattern = re.compile(
            r"#(\d+)\s+([\w\s]+?)\s+([\-\d.]+)\s+([\-\d.]+)\s+([\-\d.]+)\s+([\-\d.]+)\s+([\d/]+)\s*$",
            re.MULTILINE
        )
        rows = []
        for m in pattern.finditer(block):
            rows.append({
                "Enc #":        int(m.group(1)),
                "Type":         m.group(2).strip(),
                "Angle (°)":    float(m.group(3)),
                "Site (°)":     float(m.group(4)),
                "Top Z (m)":    float(m.group(5)),
                "Bottom Z (m)": float(m.group(6)),
                "Panflex":      m.group(7),
            })
        return rows, ["Enc #", "Type", "Angle (°)", "Site (°)", "Top Z (m)", "Bottom Z (m)", "Panflex"]

    elif has_angles and not has_panflex:
        # Line array without Panflex (KIVA, X-Series arrays...)
        pattern = re.compile(
            r"#(\d+)\s+([\w\s]+?)\s+([\-\d.]+)\s+([\-\d.]+)\s+([\-\d.]+)\s+([\-\d.]+)\s*$",
            re.MULTILINE
        )
        rows = []
        for m in pattern.finditer(block):
            rows.append({
                "Enc #":        int(m.group(1)),
                "Type":         m.group(2).strip(),
                "Angle (°)":    float(m.group(3)),
                "Site (°)":     float(m.group(4)),
                "Top Z (m)":    float(m.group(5)),
                "Bottom Z (m)": float(m.group(6)),
            })
        return rows, ["Enc #", "Type", "Angle (°)", "Site (°)", "Top Z (m)", "Bottom Z (m)"]

    else:
        # Subs / point source (SB28, X8, KS28...)
        pattern = re.compile(
            r"#(\d+)\s+([\w]+(?:_C)?)\s+([\-\d.]+)\s+([\-\d.]+)\s+([\-\d.]+)\s*$",
            re.MULTILINE
        )
        rows = []
        for m in pattern.finditer(block):
            rows.append({
                "Enc #":        int(m.group(1)),
                "Type":         m.group(2).strip(),
                "Site (°)":     float(m.group(3)),
                "Top Z (m)":    float(m.group(4)),
                "Bottom Z (m)": float(m.group(5)),
            })
        return rows, ["Enc #", "Type", "Site (°)", "Top Z (m)", "Bottom Z (m)"]


def split_source_blocks(text):
    pattern = re.compile(r"^\d+\.\s+Source:\s+(.+)$", re.MULTILINE)
    matches = list(pattern.finditer(text))
    blocks = []
    for i, m in enumerate(matches):
        start = m.start()
        end   = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        blocks.append((m.group(1).strip(), text[start:end]))
    return blocks


def get_group_for_source(text, source_name):
    pattern = re.compile(rf"\d+\.\s+Source:\s+{re.escape(source_name)}")
    m = pattern.search(text)
    if not m:
        return "Unknown"
    group_pattern = re.compile(r"^\d+\.\s+Group:\s+(.+)$", re.MULTILINE)
    groups_before = list(group_pattern.finditer(text, 0, m.start()))
    if not groups_before:
        return "Unknown"
    for g in reversed(groups_before):
        name = g.group(1).strip()
        if name.upper() != "ALL":
            return name
    return "Unknown"


def source_fingerprint(physical, enclosures):
    """
    Hashable fingerprint for mirror deduplication.
    Normalises Position X and Azimuth to their absolute values so that
    L/R mirrors (which have negated X and azimuth) hash identically.
    """
    normalised = {}
    for k, v in physical.items():
        if k in ("Position X (m)", "Azimuth (°)"):
            try:
                v = str(abs(float(v)))
            except ValueError:
                pass
        normalised[k] = v
    phys_tuple = tuple(sorted(normalised.items()))
    # Sort enclosures by type so mirrored arrays (e.g. SB28_C at #1 vs #4) still match
    from collections import Counter
    enc_tuple = tuple(sorted(Counter(e["Type"] for e in enclosures).items()))
    return (phys_tuple, enc_tuple)


def parse_document(text):
    source_blocks = split_source_blocks(text)
    groups = {}

    for source_name, block in source_blocks:
        group            = get_group_for_source(text, source_name)
        physical         = parse_physical_config(block)
        enclosures, cols = parse_enclosure_table(block)

        # Skip if an identical source already exists in this group (mirror dedup)
        fp = source_fingerprint(physical, enclosures)
        if any(source_fingerprint(s["physical"], s["enclosures"]) == fp
               for s in groups.get(group, [])):
            continue

        canonical = get_canonical_name(source_name)
        entry = {
            "name":       canonical,
            "physical":   physical,
            "enclosures": enclosures,
            "columns":    cols,
        }
        if group not in groups:
            groups[group] = []
        groups[group].append(entry)

    return groups


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

def write_excel(groups, output_path):
    wb = Workbook()
    wb.remove(wb.active)

    for group_name, sources in groups.items():
        sheet_name = re.sub(r'[\\/*?:\[\]]', '', group_name)[:31]
        ws = wb.create_sheet(title=sheet_name)
        ws.sheet_view.showGridLines = False
        row = 1

        ws.merge_cells(f"A{row}:G{row}")
        c = ws[f"A{row}"]
        c.value = f"Group: {group_name}"
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=14)
        c.fill = HEADER_FILL; c.alignment = CENTER
        ws.row_dimensions[row].height = 30
        row += 2

        for source in sources:
            ws.merge_cells(f"A{row}:G{row}")
            c = ws[f"A{row}"]
            c.value = source["name"]
            c.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
            c.fill = SUBHEAD_FILL; c.alignment = LEFT
            ws.row_dimensions[row].height = 22
            row += 1

            physical = source["physical"]
            if physical:
                items = list(physical.items())
                for i in range(0, len(items), 2):
                    fill = ALT_FILL if (row % 2 == 0) else WHITE_FILL
                    for offset, idx in enumerate([i, i + 1]):
                        if idx >= len(items):
                            break
                        key, val = items[idx]
                        col_label = offset * 3 + 1
                        lc = ws.cell(row=row, column=col_label, value=key)
                        lc.font = LABEL_FONT; lc.alignment = LEFT
                        lc.fill = fill; lc.border = thin_border()
                        vc = ws.cell(row=row, column=col_label + 1, value=val)
                        vc.font = BODY_FONT; vc.alignment = LEFT
                        vc.fill = fill; vc.border = thin_border()
                        ws.merge_cells(start_row=row, start_column=col_label + 1,
                                       end_row=row, end_column=col_label + 2)
                    row += 1
                row += 1

            enclosures = source["enclosures"]
            columns    = source["columns"]
            if enclosures:
                ws.merge_cells(f"A{row}:G{row}")
                c = ws[f"A{row}"]
                c.value = "Per-Enclosure Geometry"
                c.font = Font(name="Arial", bold=True, color="1F3864", size=9)
                c.fill = ALT_FILL; c.alignment = LEFT
                ws.row_dimensions[row].height = 15
                row += 1

                for col, h in enumerate(columns, 1):
                    c = ws.cell(row=row, column=col, value=h)
                    c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
                    c.fill = ROW_FILL; c.alignment = CENTER; c.border = thin_border()
                ws.row_dimensions[row].height = 15
                row += 1

                for enc in enclosures:
                    fill = ALT_FILL if enc["Enc #"] % 2 == 0 else WHITE_FILL
                    for col, key in enumerate(columns, 1):
                        c = ws.cell(row=row, column=col, value=enc.get(key, ""))
                        c.font = BODY_FONT; c.alignment = CENTER
                        c.fill = fill; c.border = thin_border()
                    row += 1

            row += 2

        for col, width in zip("ABCDEFG", [22, 18, 4, 22, 18, 4, 12]):
            ws.column_dimensions[col].width = width

    wb.save(output_path)


# ── PDF writing ───────────────────────────────────────────────────────────────

def write_pdf(groups, output_path):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT

    doc = SimpleDocTemplate(str(output_path), pagesize=A4,
                            leftMargin=20*mm, rightMargin=20*mm,
                            topMargin=20*mm, bottomMargin=20*mm)

    NAVY  = colors.HexColor("#1F3864")
    BLUE  = colors.HexColor("#2E75B6")
    MBLUE = colors.HexColor("#4472C4")
    LBLUE = colors.HexColor("#DCE6F1")
    WHITE = colors.white

    title_style   = ParagraphStyle("t",  fontSize=15, textColor=WHITE, fontName="Helvetica-Bold", leading=18)
    source_style  = ParagraphStyle("s",  fontSize=11, textColor=WHITE, fontName="Helvetica-Bold", leading=14)
    section_style = ParagraphStyle("ss", fontSize=9,  textColor=NAVY,  fontName="Helvetica-Bold", leading=11)
    label_style   = ParagraphStyle("l",  fontSize=9,  textColor=NAVY,  fontName="Helvetica-Bold")
    value_style   = ParagraphStyle("v",  fontSize=9,  fontName="Helvetica")

    def banner(text, style, bg, width=170*mm):
        t = Table([[Paragraph(text, style)]], colWidths=[width])
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), bg),
            ("TOPPADDING",    (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("LEFTPADDING",   (0,0), (-1,-1), 10),
        ]))
        return t

    story = []
    first_group = True

    for group_name, sources in groups.items():
        if not first_group:
            story.append(PageBreak())
        first_group = False

        story.append(banner(f"Group: {group_name}", title_style, NAVY))
        story.append(Spacer(1, 5*mm))

        for source in sources:
            story.append(banner(source["name"], source_style, BLUE))
            story.append(Spacer(1, 2*mm))

            physical = source["physical"]
            if physical:
                story.append(banner("Physical Configuration", section_style, LBLUE))
                story.append(Spacer(1, 1*mm))
                items = list(physical.items())
                rows = []
                for i in range(0, len(items), 2):
                    lk, lv = items[i]
                    rk, rv = items[i+1] if i+1 < len(items) else ("", "")
                    rows.append([
                        Paragraph(lk, label_style), Paragraph(str(lv), value_style),
                        Paragraph(rk, label_style), Paragraph(str(rv), value_style),
                    ])
                pt = Table(rows, colWidths=[42*mm, 40*mm, 42*mm, 46*mm])
                ps = [("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#B0B0B0")),
                      ("TOPPADDING", (0,0), (-1,-1), 3),
                      ("BOTTOMPADDING", (0,0), (-1,-1), 3),
                      ("LEFTPADDING", (0,0), (-1,-1), 5)]
                for i in range(len(rows)):
                    ps.append(("BACKGROUND", (0,i), (-1,i), LBLUE if i%2==0 else WHITE))
                pt.setStyle(TableStyle(ps))
                story.append(pt)
                story.append(Spacer(1, 3*mm))

            enclosures = source["enclosures"]
            columns    = source["columns"]
            if enclosures:
                story.append(banner("Per-Enclosure Geometry", section_style, LBLUE))
                story.append(Spacer(1, 1*mm))
                col_w = 170*mm / len(columns)
                enc_rows = [columns] + [[str(e.get(k, "")) for k in columns] for e in enclosures]
                et = Table(enc_rows, colWidths=[col_w] * len(columns))
                es = [("BACKGROUND",    (0,0), (-1,0),  MBLUE),
                      ("TEXTCOLOR",     (0,0), (-1,0),  WHITE),
                      ("FONTNAME",      (0,0), (-1,0),  "Helvetica-Bold"),
                      ("FONTSIZE",      (0,0), (-1,-1), 8),
                      ("ALIGN",         (0,0), (-1,-1), "CENTER"),
                      ("GRID",          (0,0), (-1,-1), 0.5, colors.HexColor("#B0B0B0")),
                      ("TOPPADDING",    (0,0), (-1,-1), 3),
                      ("BOTTOMPADDING", (0,0), (-1,-1), 3)]
                for i in range(1, len(enc_rows)):
                    es.append(("BACKGROUND", (0,i), (-1,i), LBLUE if i%2==0 else WHITE))
                et.setStyle(TableStyle(es))
                story.append(et)

            story.append(Spacer(1, 6*mm))

    doc.build(story)


# ── Runner ────────────────────────────────────────────────────────────────────

def process_pdf(pdf_path):
    print(f"  Processing: {pdf_path.name}")
    text   = extract_text(pdf_path)
    groups = parse_document(text)
    total  = sum(len(s) for s in groups.values())
    print(f"  Found {len(groups)} group(s), {total} source(s)")
    xlsx_path  = OUTPUT_DIR / (pdf_path.stem + ".xlsx")
    pdf_path2  = OUTPUT_DIR / (pdf_path.stem + "_report.pdf")
    write_excel(groups, xlsx_path)
    write_pdf(groups, pdf_path2)
    print(f"  Saved: {xlsx_path.name}")
    print(f"  Saved: {pdf_path2.name}")


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if len(sys.argv) > 1:
        pdf_path = DATA_DIR / sys.argv[1]
        if not pdf_path.exists():
            print(f"Error: {pdf_path} not found.")
            sys.exit(1)
        pdfs = [pdf_path]
    else:
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