from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import tempfile
import sys
from pathlib import Path

# Garante que o Python encontra o src/extract.py
sys.path.insert(0, str(Path(__file__).parent.parent))
from src.extract import extract_text, extract_metadata, parse_document, write_excel, write_pdf

app = FastAPI()

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/extract")
async def extract(file: UploadFile = File(...)):
    # Guarda o PDF temporariamente
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(await file.read())
        tmp_path = Path(tmp.name)

    # Processa
    text   = extract_text(tmp_path)
    name, date = extract_metadata(text)
    groups = parse_document(text)

    # Gera Excel
    xlsx_path = tmp_path.with_suffix(".xlsx")
    write_excel(groups, xlsx_path, report_name=name, report_date=date)

    return FileResponse(
        path=xlsx_path,
        filename=f"{name}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
