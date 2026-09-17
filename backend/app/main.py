import logging
import os
from pathlib import Path
from tempfile import NamedTemporaryFile, gettempdir
from uuid import uuid4

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

from .excel import create_batch_excel, create_excel, create_ipd_consolidated_excel, populate_ipd_rows
from .models import ExtractionResult
from .parser import extract_structured_data, extract_text_from_file
from .storage import VolumeArchive


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
VOLUME_ARCHIVE = VolumeArchive.from_environment()
LOGGER = logging.getLogger(__name__)

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


@app.get("/api/storage-status")
def storage_status() -> dict[str, bool]:
    return {"enabled": VOLUME_ARCHIVE.enabled}


@app.post("/api/extract", response_model=ExtractionResult)
async def extract(file: UploadFile = File(...)) -> ExtractionResult:
    suffix = Path(file.filename or "").suffix.lower()
    if suffix not in {".docx", ".pdf"}:
        raise HTTPException(status_code=400, detail="Unsupported file type. Please upload a .pdf or .docx file.")

    processing_id = uuid4().hex
    upload_content = b""
    try:
        upload_content = await file.read()
        with NamedTemporaryFile(delete=False, suffix=suffix, dir=UPLOAD_DIR) as temporary:
            temporary.write(upload_content)
            temp_path = Path(temporary.name)

        text, parser_warnings = extract_text_from_file(temp_path)
        result = extract_structured_data(text, file.filename or temp_path.name, parser_warnings)
        result.processing_id = processing_id
        result = populate_ipd_rows(result)
        try:
            VOLUME_ARCHIVE.write(
                "source",
                file.filename or temp_path.name,
                upload_content,
                processing_id,
            )
        except OSError:
            LOGGER.exception("Could not archive source document for processing ID %s", processing_id)
            result.extraction_warnings.append(
                "The document was extracted, but the source file could not be archived in persistent storage."
            )
        return result
    except HTTPException:
        raise
    except Exception as exc:
        if upload_content:
            try:
                VOLUME_ARCHIVE.write(
                    "failed",
                    file.filename or f"failed{suffix}",
                    upload_content,
                    processing_id,
                )
            except OSError:
                LOGGER.exception("Could not archive failed document for processing ID %s", processing_id)
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
    _archive_excel(excel_file.getvalue(), f"{filename}.xlsx", result.processing_id)
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

    _archive_excel(excel_file.getvalue(), "product-data-batch.xlsx")
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

    _archive_excel(excel_file.getvalue(), "nidar-ipd-consolidated.xlsx")
    headers = {"Content-Disposition": 'attachment; filename="nidar-ipd-consolidated.xlsx"'}
    return StreamingResponse(
        excel_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


def _archive_excel(content: bytes, filename: str, processing_id: str | None = None) -> None:
    archive_id = processing_id or uuid4().hex
    try:
        VOLUME_ARCHIVE.write("excel", filename, content, archive_id)
    except OSError:
        LOGGER.exception("Could not archive Excel output for processing ID %s", archive_id)


if FRONTEND_DIST.exists():
    app.mount("/", StaticFiles(directory=FRONTEND_DIST, html=True), name="frontend")
