# Product Data Extractor MVP

A full-stack customer-facing MVP for uploading PDF/DOCX product specification documents, extracting structured product data with deterministic rules, reviewing the result in the browser, and exporting to Excel.

This MVP uses deterministic extraction. It is designed for transparent review. AI extraction can be added later after security approval.

## Stack

- Frontend: React + Vite
- Backend: Python FastAPI
- DOCX parsing: python-docx
- PDF parsing: pdfplumber
- Excel export: pandas + openpyxl

## Architecture comparison

### Current local MVP

- React + Vite customer review interface.
- Python FastAPI extraction and export API.
- `python-docx`, `pdfplumber`, and local Tesseract OCR for text and embedded images.
- Deterministic regex, keyword, table-OCR, and Pydantic-based mapping. No external AI API.
- Browser-memory review state with editable IPD rows, field approval, source references, and a local change log.
- pandas + openpyxl exports for review and migration preparation.
- Temporary local uploads only; no database, enterprise identity, or direct IPD/PLM connection.

### Production direction described in the governance presentation

- Runs inside Orkla's Databricks environment and reads from a restricted project folder.
- Uses an approved vision-language model to understand supplier pages and pasted screenshots.
- Maps extracted values to the 165 IPD attributes with unit and source evidence per field.
- Stores editable review results in Lakebase.
- Records uploader, model output, human edits, approval, and load activity as an audit trail.
- Produces approved IPD and PLM load files; no data is loaded before superuser approval.

The presentation does not specify the production frontend framework, API framework, model name, or exact IPD/PLM integration mechanism. Those remain architecture decisions rather than confirmed technology choices.

## Windows Setup

### 1. Create Python venv in backend

```powershell
cd backend
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 2. Install backend requirements

```powershell
pip install -r requirements.txt
```

Optional OCR support for image-only Word files:

If users paste screenshots/images into a DOCX, install the local Tesseract OCR engine for Windows and restart the backend. The Python package `pytesseract` is already included in `requirements.txt`, but it needs the Windows OCR app too.

Recommended install path:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

After installing, verify in a new PowerShell window:

```powershell
tesseract --version
```

If the command is not found, add `C:\Program Files\Tesseract-OCR` to your Windows PATH and restart PowerShell.

### 3. Start FastAPI backend

```powershell
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

Backend runs at http://localhost:8000.

Health check:

```powershell
Invoke-RestMethod http://localhost:8000/api/health
```

### 4. Install frontend dependencies

Open a new PowerShell window:

```powershell
cd frontend
npm install
```

### 5. Start React frontend

```powershell
npm run dev
```

Frontend runs at http://localhost:5173.

### 6. Upload a DOCX/PDF

Open http://localhost:5173, drag and drop one or more `.docx`/`.pdf` product specifications, use **Choose files**, or use **Choose folder** for a complete source folder. Folder uploads ignore unsupported files and process supported documents one at a time so local OCR remains stable. Use the document tabs to review and edit each result.

### 7. Export Excel

After extraction completes, review the dashboard and click **Export to Excel**. The workbook includes:

1. IPD_Template: customer attribute rows from `samples/ipd-template.xlsx`, with `Data` and `UoM` filled where the deterministic extractor has a confident match.
2. Summary
3. Nutrition
4. Allergens
5. Shelf_Life_Storage
6. Physical_Chemical
7. Raw_Review

The export also includes review-oriented sheets for messy supplier formats:

- Key_Value_Candidates: generic label/value pairs found in text, tables, and OCR.
- Detected_Sections: full detected text blocks for nutrition, allergens, shelf life, storage, and physical/chemical data.
- Raw_Lines: every extracted/OCR line with line number and inferred section.
- Review_Flags: values that may need manual review, especially OCR-heavy nutrition values.

The `IPD_Template` sheet is intentionally conservative: if the app cannot confidently map a document value to a customer attribute, the cell is left blank and the source text remains available in the review sheets.

The app also supports output profiles. Roller-fabric forms are detected automatically and exported to the 65-column `Fabric_Specs_Mittet` schema. Tai Hing filament specifications are exported to the 123-field `Filament_Specs_Taihing_Nylon` schema. Both retain the supplied `Schema Dictionary` and review sheets. Documents that do not match a known profile use a flexible extracted-data workbook so captured content is not discarded.

For local port conflicts, set `VITE_BACKEND_TARGET` in `frontend/.env.local`. The normal default remains `http://127.0.0.1:8000`.

Select **IPD consolidated** to export every successfully extracted document to one organized workbook. `Overview` contains one row per product, `Allergens` shows the 14 regulated allergen groups explicitly, and `IPD_Mapped` contains only populated migration values. The complete approved attribute list remains in `IPD_All_Attributes`; generic captures, warnings, and raw OCR are separated into review sheets so unfamiliar supplier content is retained without cluttering the main overview.

Image-based allergen declarations receive a second table-oriented OCR pass. Yes/No values and declared quantities are mapped to dedicated fields and color-coded in Excel. A declaration that is detected without readable statuses is flagged for manual review rather than treated as allergen-free.

Physical, chemical, typical-value, and microbiological specification tables are preserved row by row in the browser and in a dedicated `Specifications` sheet. Limits, ranges, units, qualifiers such as `max`/`min`/`approx`, and source evidence remain separate. Confident matches are also mapped into IPD fields; unmatched rows stay visible instead of being forced into the wrong attribute.

GNT Exberry specifications are recognized as a reusable supplier-family template. Their manufacturing, colouring, physical/chemical, microbiological, general appearance, nutrition, and shelf-life tables are captured consistently across product colours. Common OCR unit substitutions are normalized but remain marked for review when the source image is ambiguous.

The browser also exposes the complete 165-attribute IPD worklist. It shows found, missing, and approved counts; sorts unresolved fields for review; allows direct value/unit edits and row approval; and records those actions in the exported `Change_Log` sheet. Source evidence identifies the PDF page or embedded Word image and includes the extracted source line.

Select **Document sheets** to retain the previous batch layout: one workbook with a `Batch_Index` sheet and one customer-shaped data sheet per successfully extracted document. Mixed output profiles can be included in the same batch.

Unknown formats use deterministic key-value capture. Their detected fields remain editable and exportable. A reviewed example and desired workbook can then be added as a reusable output profile, without changing previously supported profiles.

To update the customer migration attribute list later, replace `samples/ipd-template.xlsx` with the latest approved template using the same header columns.

## API

### GET /api/health

Returns:

```json
{ "status": "ok" }
```

### GET /api/storage-status

Returns whether persistent Unity Catalog Volume archival is configured:

```json
{ "enabled": true }
```

### POST /api/extract

Accepts a multipart file upload named `file`. Supports `.pdf` and `.docx`.

### POST /api/export-excel

Accepts the extracted JSON result and returns an `.xlsx` file download.

### POST /api/export-excel-batch

Accepts a JSON array of extracted results and returns one workbook containing a batch index and one data sheet per document.

### POST /api/export-ipd-consolidated

Accepts a JSON array of extracted results and returns one organized IPD migration workbook for the complete batch.

## Security Notes

- Uploaded files are written to temporary processing storage and deleted after extraction. When a Unity Catalog Volume is configured, a governed source copy is also archived there.
- `backend/uploads/` and `backend/output/` are ignored by git.
- No external AI APIs are used.
- Document content is not sent to an external AI service. In local mode it stays on the machine; in Databricks Apps it stays inside the configured Databricks workspace boundary.

## Databricks Apps deployment

The repository is deployable as a single Databricks App without enabling AI extraction. During deployment, Databricks installs the root Python and Node dependencies, runs the Vite production build, and starts FastAPI with `app.yaml`. FastAPI serves both the `/api` endpoints and the compiled React frontend from one origin.

Temporary processing files are deleted after extraction. The configured Unity Catalog Volume archives source documents, failed documents, and generated Excel workbooks. Excel workbooks are still streamed directly to the browser. Databricks model integration is intentionally deferred to the next milestone.

### Configure persistent file storage

Before deploying this revision, create or select a Unity Catalog Volume and attach it to the app:

1. In Catalog Explorer, create or select a catalog and schema.
2. Create a managed volume, for example `product_data_files`.
3. Open the Databricks App, then open **Settings** and **App resources**.
4. Add a **UC volume** resource with **Can read and write**.
5. Keep the resource key as `volume`; `app.yaml` maps that resource to `PRODUCT_DATA_VOLUME_PATH`.
6. Redeploy the app from the `main` branch.

The app creates this structure inside the selected volume:

```text
source/YYYY/MM/DD/<processing-id>_<source-file>
failed/YYYY/MM/DD/<processing-id>_<source-file>
excel/YYYY/MM/DD/<processing-id>_<workbook>.xlsx
```

Check `/api/storage-status` after deployment. It must return `{ "enabled": true }`. Archive failures do not block extraction or browser downloads: source archive failures appear as extraction warnings, and all failures are written to the Databricks App logs.

### Prerequisites

- A workspace region with Databricks Apps support.
- Databricks CLI 0.239.0 or newer.
- OAuth access to the target workspace.
- Workspace network access to the configured Python and npm package registries.

### Validate and deploy

Authenticate the CLI using the workspace URL:

```powershell
databricks auth login --host https://<workspace-host>
```

From the repository root, validate and deploy the development target:

```powershell
databricks bundle validate -t dev
databricks bundle deploy -t dev
databricks bundle run product_data_extractor -t dev
```

The bundle creates the `product-data-extractor-mvp` Databricks App. Databricks injects `DATABRICKS_APP_PORT`; `run_app.py` binds FastAPI to that port on `0.0.0.0`.

The same production build can be tested locally:

```powershell
npm install
npm run build
$env:PORT = "8000"
$env:PRODUCT_DATA_VOLUME_PATH = "C:\temp\product-data-volume"
python run_app.py
```

Open http://localhost:8000 and verify upload, review, and Excel download through the single FastAPI origin.

### OCR limitation in milestone 1

Native PDF text and DOCX text/tables work in Databricks. Image-only PDFs and DOCX files containing screenshots currently depend on the Tesseract system executable, which is not installed by the Databricks App package build. Those documents return an OCR warning and remain the main target for the approved Databricks vision model in milestone 2.
