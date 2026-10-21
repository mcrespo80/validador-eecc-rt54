import os
import uuid
import shutil
import json
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from core.extractor import DocumentExtractor
from core.validator import EECCValidator
from core.report_pdf import generate_audit_report_pdf
from core.remediator import generate_improved_eecc_package

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
KB_DIR = os.path.join(BASE_DIR, "knowledge_base")
MODELOS_DIR = os.path.join(KB_DIR, "modelos_consejo")
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

app = FastAPI(
    title="MIC DIGITAL CONSULTING - Validador de EECC (RT 54 / CPCECABA)",
    version="2026.1"
)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

validator_engine = EECCValidator(os.path.join(KB_DIR, "reglas_rt54_cpcecaba.json"))


@app.get("/", response_class=HTMLResponse)
async def serve_dashboard():
    index_path = os.path.join(STATIC_DIR, "index.html")
    with open(index_path, "r", encoding="utf-8") as f:
        return f.read()


@app.get("/api/knowledge-base")
async def get_knowledge_base():
    """Returns the loaded RT 54 / RT 37 / CPCECABA rules and the list of official template files."""
    modelos = []
    if os.path.exists(MODELOS_DIR):
        for fname in sorted(os.listdir(MODELOS_DIR)):
            if fname.startswith("~$"):
                continue
            fpath = os.path.join(MODELOS_DIR, fname)
            if os.path.isfile(fpath):
                modelos.append({
                    "name": fname,
                    "ext": os.path.splitext(fname)[1].lower(),
                    "size_kb": round(os.path.getsize(fpath) / 1024, 1)
                })
    return {
        "rules": validator_engine.kb,
        "modelos_consejo": modelos
    }


@app.post("/api/validate")
async def validate_uploaded_file(
    file: UploadFile = File(...),
    jurisdiction: str = Form("AUTO")
):
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in (".pdf", ".docx", ".xlsx"):
        raise HTTPException(status_code=400, detail="Formato no soportado. Suba un archivo .pdf, .docx o .xlsx")

    job_id = str(uuid.uuid4())[:8]
    safe_name = f"{job_id}_{file.filename}"
    saved_path = os.path.join(UPLOADS_DIR, safe_name)

    with open(saved_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        extracted = DocumentExtractor.extract(saved_path)
        result = validator_engine.validate(extracted, jurisdiction_pref=jurisdiction)

        # Generate official branded PDF Audit Report
        report_filename = f"Reporte_Auditoria_EECC_{job_id}.pdf"
        report_path = os.path.join(REPORTS_DIR, report_filename)
        logo_path = os.path.join(BASE_DIR, "data", "brand", "emblem_clean.png")
        generate_audit_report_pdf(result, report_path, logo_path=logo_path)

        # Generate Improved / Corrected Version ready for CPCECABA presentation (PDF + DOCX)
        improved_pdf_name = f"EECC_Version_Mejorada_CPCECABA_{job_id}.pdf"
        improved_docx_name = f"EECC_Version_Mejorada_CPCECABA_{job_id}.docx"
        improved_pdf_path = os.path.join(REPORTS_DIR, improved_pdf_name)
        improved_docx_path = os.path.join(REPORTS_DIR, improved_docx_name)
        remediation_info = generate_improved_eecc_package(
            extracted, result, improved_pdf_path, improved_docx_path
        )

        # Also copy the report and improved version to G:\Otros ordenadores\Mi Portátil (1)\MC\RT54 if accessible
        g_drive_rt54 = r"G:\Otros ordenadores\Mi Portátil (1)\MC\RT54"
        if os.path.exists(g_drive_rt54):
            try:
                shutil.copy(report_path, os.path.join(g_drive_rt54, "Ultimo_Reporte_Validacion_EECC.pdf"))
                shutil.copy(improved_pdf_path, os.path.join(g_drive_rt54, "Ultima_Version_Mejorada_EECC_CPCECABA.pdf"))
                shutil.copy(improved_docx_path, os.path.join(g_drive_rt54, "Ultima_Version_Mejorada_EECC_CPCECABA.docx"))
            except Exception:
                pass

        result["report_pdf_url"] = f"/api/reports/{report_filename}"
        result["improved_pdf_url"] = f"/api/reports/{improved_pdf_name}"
        result["improved_docx_url"] = f"/api/reports/{improved_docx_name}"
        result["remediation_info"] = remediation_info
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al procesar el documento: {str(e)}")


@app.post("/api/validate-sample")
async def validate_consejo_sample(
    filename: str = Form(...),
    jurisdiction: str = Form("AUTO")
):
    """Allows testing the validator directly against any file in knowledge_base/modelos_consejo or local EECC."""
    sample_path = os.path.join(MODELOS_DIR, filename)
    if not os.path.exists(sample_path):
        # Check if it's OSETYA in PROYECTO02
        alt_path = os.path.join(r"C:\Users\mcres\PROYECTO02", filename)
        if os.path.exists(alt_path):
            sample_path = alt_path
        else:
            raise HTTPException(status_code=404, detail=f"Archivo no encontrado: {filename}")

    job_id = str(uuid.uuid4())[:8]
    extracted = DocumentExtractor.extract(sample_path)
    result = validator_engine.validate(extracted, jurisdiction_pref=jurisdiction)

    report_filename = f"Reporte_Auditoria_EECC_{job_id}.pdf"
    report_path = os.path.join(REPORTS_DIR, report_filename)
    logo_path = os.path.join(BASE_DIR, "data", "brand", "emblem_clean.png")
    generate_audit_report_pdf(result, report_path, logo_path=logo_path)

    improved_pdf_name = f"EECC_Version_Mejorada_CPCECABA_{job_id}.pdf"
    improved_docx_name = f"EECC_Version_Mejorada_CPCECABA_{job_id}.docx"
    improved_pdf_path = os.path.join(REPORTS_DIR, improved_pdf_name)
    improved_docx_path = os.path.join(REPORTS_DIR, improved_docx_name)
    remediation_info = generate_improved_eecc_package(
        extracted, result, improved_pdf_path, improved_docx_path
    )

    result["report_pdf_url"] = f"/api/reports/{report_filename}"
    result["improved_pdf_url"] = f"/api/reports/{improved_pdf_name}"
    result["improved_docx_url"] = f"/api/reports/{improved_docx_name}"
    result["remediation_info"] = remediation_info
    return JSONResponse(content=result)


@app.get("/api/reports/{filename}")
async def download_report(filename: str):
    fpath = os.path.join(REPORTS_DIR, filename)
    if not os.path.exists(fpath):
        raise HTTPException(status_code=404, detail="Reporte no encontrado")
    media = "application/pdf" if filename.lower().endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return FileResponse(fpath, media_type=media, filename=filename)


@app.get("/api/modelos/{filename}")
async def download_modelo_consejo(filename: str):
    fpath = os.path.join(MODELOS_DIR, filename)
    if not os.path.exists(fpath):
        raise HTTPException(status_code=404, detail="Modelo no encontrado")
    return FileResponse(fpath, filename=filename)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="127.0.0.1", port=8054, reload=False)
