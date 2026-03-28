# Soundvision Report Data Extractor

A Python tool that parses L-Acoustics Soundvision PDF exports and generates formatted Excel and PDF reports for internal use.

## Overview

Soundvision Report Data Extractor reads Soundvision project reports and produces structured output files with physical configuration, enclosure geometry, and user-fillable fields for circuit assignments and amplifier data.

Available as a **Streamlit web app** and as a **FastAPI backend** for the native macOS/iOS companion app.

## Features

- Parses all flown arrays from Soundvision PDF exports
- Supports K1, K2, K3, KARA II, KIVA II, KS28, SB28, X8 and mixed arrays
- Deduplicates L/R mirror arrays automatically
- Generates a structured **Excel** with:
  - Report Info sheet with user-fillable fields (System Engineer, Company, Venue, Date)
  - System Summary with circuit assignment dropdowns (A–J, resistor colour code)
  - One sheet per group with physical configuration and per-enclosure geometry
  - Formulas linking circuit and amp data from Report Info to group sheets
- Generates a **PDF rigging reference** with physical configuration and enclosure geometry

## Project Structure

```
SOUNDVISION-EXCEL/
├── app.py                  # Streamlit web interface
├── backend/
│   └── server.py           # FastAPI server for native app
├── src/
│   └── extract.py          # Core parsing and output logic
├── assets/
│   └── lasmall.png         # L-Acoustics logo
├── requirements.txt
└── data/                   # Place input PDFs here (local only)
```

## Setup

```bash
git clone https://github.com/afonsopires1904/SOUNDVISION-EXCEL.git
cd SOUNDVISION-EXCEL
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

### Streamlit web app

```bash
streamlit run app.py
```

### FastAPI backend (for native app)

```bash
uvicorn backend.server:app --reload
```

### Command line

```bash
# Process all PDFs in data/
python src/extract.py

# Process a specific file
python src/extract.py my_report.pdf
```

## Live App

The Streamlit web app is deployed at: [soundvision-report-extract-afonsopires.streamlit.app](https://soundvision-report-extract-afonsopires.streamlit.app)

## Companion App

The native macOS/iOS app lives at: [SoundvisionExtractor](https://github.com/afonsopires1904/SoundvisionExtractor)

## Author

Built for internal purposes by Afonso Pires
