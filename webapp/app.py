from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from engine import analyze_text, extract_file, load_rules

DIR = Path(__file__).resolve().parent
SAMPLES = DIR / "samples"
MAX_BYTES = 15 * 1024 * 1024
ALLOWED = {".pdf", ".docx", ".pptx", ".txt", ".md", ".html", ".htm", ".csv"}

app = FastAPI(title="Skaner greenwashingu", version="2026.09.15")
app.mount("/static", StaticFiles(directory=DIR / "static"), name="static")


@app.get("/")
def index():
    return FileResponse(DIR / "static" / "index.html")


@app.get("/api/meta")
def meta():
    rules = load_rules()
    return {
        "as_of": rules["meta"]["as_of"],
        "ecgt_apply_from": rules["meta"]["ecgt_apply_from"],
        "green_claims_status": rules["meta"]["green_claims_status"],
        "uokik_cases": rules["enforcement_pl"]["uokik_live_cases"],
        "max_mb": 15,
        "formats": sorted(ALLOWED),
    }


@app.get("/api/sample/{name}")
def sample(name: str):
    safe = {"ryzyko": "reklama_ryzyko.txt", "ok": "reklama_ok.txt"}.get(name)
    if not safe:
        raise HTTPException(404, "Nieznany przykład")
    path = SAMPLES / safe
    text = path.read_text(encoding="utf-8")
    result = analyze_text(text, filename=safe)
    result["sample"] = name
    return result


@app.post("/api/analyze")
async def analyze(
    file: UploadFile | None = File(default=None),
    text: str | None = Form(default=None),
):
    if file is not None and file.filename:
        raw = await file.read()
        if len(raw) > MAX_BYTES:
            raise HTTPException(413, "Plik jest większy niż 15 MB.")
        suffix = Path(file.filename).suffix.lower()
        if suffix and suffix not in ALLOWED:
            raise HTTPException(415, f"Nieobsługiwany format {suffix}. Użyj PDF, DOCX, PPTX, TXT lub HTML.")
        try:
            extracted = extract_file(file.filename, raw)
        except Exception as exc:
            raise HTTPException(422, f"Nie udało się odczytać pliku: {exc}") from exc
        if not extracted.strip():
            raise HTTPException(422, "Plik nie zawiera tekstu (skan bez OCR, puste slajdy albo sam obraz).")
        return analyze_text(extracted, filename=file.filename)

    pasted = (text or "").strip()
    if pasted:
        return analyze_text(pasted, filename="wklejony-tekst.txt")

    raise HTTPException(400, "Wczytaj plik albo wklej treść.")
