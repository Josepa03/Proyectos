"""
main.py — Servidor FastAPI para comparación PDF vs Excel
"""
import logging
import shutil
import uuid
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

import comparador

# ── Setup ─────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("main")

UPLOADS_DIR = Path("uploads")
OUTPUTS_DIR = Path("outputs")
UPLOADS_DIR.mkdir(exist_ok=True)
OUTPUTS_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Comparador de Facturas")
app.mount("/static", StaticFiles(directory="static"), name="static")

# ── Sesiones en memoria ───────────────────────────────────────────────────────
sessions: dict[str, dict] = {}

# ── Endpoints ─────────────────────────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def root():
    html = Path("static/index.html").read_text(encoding="utf-8")
    return HTMLResponse(content=html)


@app.post("/upload")
async def upload_files(
    pdfs: list[UploadFile] = File(...),
    excel: UploadFile = File(...),
):
    # Validaciones
    if not pdfs or all(f.filename == "" for f in pdfs):
        raise HTTPException(400, "Se requiere al menos un archivo PDF.")
    if not excel.filename:
        raise HTTPException(400, "Se requiere un archivo Excel.")
    if not excel.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "El archivo Excel debe ser .xlsx o .xls")
    for f in pdfs:
        if not f.filename.lower().endswith(".pdf"):
            raise HTTPException(400, f"'{f.filename}' no es un PDF válido.")

    session_id = str(uuid.uuid4())
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    session_dir = UPLOADS_DIR / session_id
    session_dir.mkdir()

    # Guardar PDFs
    pdf_paths = []
    for pdf_file in pdfs:
        dest = session_dir / f"{ts}_{pdf_file.filename}"
        with dest.open("wb") as f:
            shutil.copyfileobj(pdf_file.file, f)
        pdf_paths.append(str(dest))
        log.info("PDF guardado: %s", dest)

    # Guardar Excel
    excel_dest = session_dir / f"{ts}_{excel.filename}"
    with excel_dest.open("wb") as f:
        shutil.copyfileobj(excel.file, f)
    log.info("Excel guardado: %s", excel_dest)

    sessions[session_id] = {
        "pdf_paths": pdf_paths,
        "excel_path": str(excel_dest),
        "output_file": None,
    }

    return {
        "session_id": session_id,
        "pdfs": [Path(p).name for p in pdf_paths],
        "excel": excel_dest.name,
    }


@app.post("/procesar/{session_id}")
async def procesar(session_id: str):
    if session_id not in sessions:
        raise HTTPException(404, "Sesión no encontrada. Sube los archivos primero.")

    sess = sessions[session_id]
    output_name = f"resultado_{session_id[:8]}.xlsx"
    output_path = OUTPUTS_DIR / output_name

    try:
        resumen = comparador.generar_reporte(
            pdf_paths=sess["pdf_paths"],
            excel_path=sess["excel_path"],
            output_path=str(output_path),
        )
    except Exception as e:
        log.exception("Error al procesar sesión %s", session_id)
        raise HTTPException(500, f"Error durante el procesamiento: {e}")

    sess["output_file"] = output_name

    return {
        "status": "ok",
        "archivo": output_name,
        "resumen": resumen,
    }


@app.get("/descargar/{archivo}")
async def descargar(archivo: str):
    path = OUTPUTS_DIR / archivo
    if not path.exists():
        raise HTTPException(404, "Archivo no encontrado.")
    return FileResponse(
        path=str(path),
        filename=archivo,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.delete("/limpiar/{session_id}")
async def limpiar(session_id: str):
    if session_id not in sessions:
        raise HTTPException(404, "Sesión no encontrada.")
    sess = sessions.pop(session_id)

    # Borrar uploads de la sesión
    session_dir = UPLOADS_DIR / session_id
    if session_dir.exists():
        shutil.rmtree(session_dir)

    # Borrar output
    if sess.get("output_file"):
        out = OUTPUTS_DIR / sess["output_file"]
        if out.exists():
            out.unlink()

    return {"status": "limpiado", "session_id": session_id}


# ── Inicio ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    print("\n✅ App corriendo en http://localhost:8000 — Abre tu navegador\n")
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
