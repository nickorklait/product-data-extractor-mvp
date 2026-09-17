import os
from pathlib import Path
from tempfile import NamedTemporaryFile, gettempdir

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

from .excel import create_batch_excel, create_excel, create_ipd_consolidated_excel, populate_ipd_rows
from .models import ExtractionResult
from .parser import extract_structured_data, extract_text_from_file


BASE_DIR = Path(__file__).resolve().parent.parent
PROJECT_ROOT = BASE_DIR.parent
IS_DATABRICKS = bool(os.getenv("DATABRICKS_APP_NAME") or os.getenv("PRODUCT_DATA_RUNTIME") == "databricks")
RUNTIME_DIR = Path(
    os.getenv(
        "PRODUCT_DATA_TEMP_DIR",
        str(Path(gettempdir()) / "product-data-extractor" if IS_DATABRICKS else BASE_DIR),
    )
)
UPLOAD_DIR = RUNTIME_DIR / "uploads"
OUTPUT_DIR = RUNTIME_DIR / "output"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
FRONTEND_DIST = PROJECT_ROOT / "frontend" / "dist"

app = FastAPI(title="Product Data Extractor MVP")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/extract", response_model=ExtractionResult)
async def extract(file: UploadFile = File(...)) -> ExtractionResult:
    suffix = Path(file.filename or "").suffix.lower()
    if suffix not in {".docx", ".pdf"}:
        raise HTTPException(status_code=400, detail="Unsupported file type. Please upload a .pdf or .docx file.")

    try:
        with NamedTemporaryFile(delete=False, suffix=suffix, dir=UPLOAD_DIR) as temporary:
            temporary.write(await file.read())
            temp_path = Path(temporary.name)

        text, parser_warnings = extract_text_from_file(temp_path)
        result = extract_structured_data(text, file.filename or temp_path.name, parser_warnings)
        return populate_ipd_rows(result)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Extraction failed: {exc}") from exc
    finally:
        if "temp_path" in locals() and temp_path.exists():
            temp_path.unlink(missing_ok=True)


@app.post("/api/export-excel")
def export_excel(result: ExtractionResult) -> StreamingResponse:
    try:
        excel_file = create_excel(result)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Excel export failed: {exc}") from exc

    filename = (result.metadata.product_name or "product-data-extraction").replace(" ", "-")
    headers = {"Content-Disposition": f'attachment; filename="{filename}.xlsx"'}
    return StreamingResponse(
        excel_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.post("/api/export-excel-batch")
def export_excel_batch(results: list[ExtractionResult]) -> StreamingResponse:
    if not results:
        raise HTTPException(status_code=400, detail="Add at least one extracted document to the batch.")
    try:
        excel_file = create_batch_excel(results)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Batch Excel export failed: {exc}") from exc

    headers = {"Content-Disposition": 'attachment; filename="product-data-batch.xlsx"'}
    return StreamingResponse(
        excel_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.post("/api/export-ipd-consolidated")
def export_ipd_consolidated(results: list[ExtractionResult]) -> StreamingResponse:
    if not results:
        raise HTTPException(status_code=400, detail="Add at least one extracted document to the IPD export.")
    try:
        excel_file = create_ipd_consolidated_excel(results)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Consolidated IPD export failed: {exc}") from exc

    headers = {"Content-Disposition": 'attachment; filename="nidar-ipd-consolidated.xlsx"'}
    return StreamingResponse(
        excel_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


if FRONTEND_DIST.exists():
    app.mount("/", StaticFiles(directory=FRONTEND_DIST, html=True), name="frontend")
