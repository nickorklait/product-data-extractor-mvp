import React, { useState } from 'react';
import { createRoot } from 'react-dom/client';
import { AlertTriangle, CheckCircle2, Download, FileText, FolderOpen, Loader2, UploadCloud, XCircle } from 'lucide-react';
import './styles.css';

const API_BASE = '';

const sections = [
  { title: 'Product metadata', key: 'metadata' },
  { title: 'Nutrition', key: 'nutrition' },
  { title: 'Allergens', key: 'allergens' },
  { title: 'Shelf life', key: 'shelf_life' },
  { title: 'Storage', key: 'storage' },
  { title: 'Physical and chemical data', key: 'physical_chemical' },
];

const labels = {
  source_file: 'Source file',
  product_name: 'Product name',
  supplier: 'Supplier',
  plant: 'Plant',
  document_id: 'Document ID',
  version: 'Version',
  approval_status: 'Approval status',
  approved_date: 'Approved date',
  energy_kj: 'Energy kJ',
  energy_kcal: 'Energy kcal',
  fat_g: 'Fat g',
  saturated_fat_g: 'Saturated fat g',
  carbohydrate_g: 'Carbohydrate g',
  sugars_g: 'Sugars g',
  fibre_g: 'Fibre g',
  protein_g: 'Protein g',
  salt_g: 'Salt g',
  allergens_contains: 'Contains',
  allergens_may_contain: 'May contain',
  allergen_statement: 'Allergen statement',
  gluten: 'Gluten',
  crustaceans: 'Crustaceans',
  eggs: 'Eggs',
  fish: 'Fish',
  peanuts: 'Peanuts',
  soybeans: 'Soybeans',
  milk: 'Milk',
  nuts: 'Nuts',
  celery: 'Celery',
  mustard: 'Mustard',
  sesame: 'Sesame',
  sulphites: 'Sulphur dioxide and sulphites',
  lupin: 'Lupin',
  molluscs: 'Molluscs',
  shelf_life_value: 'Shelf life value',
  shelf_life_unit: 'Shelf life unit',
  shelf_life_text: 'Shelf life text',
  storage_temperature: 'Storage temperature',
  storage_text: 'Storage text',
  moisture: 'Moisture',
  ph: 'pH',
  density: 'Density',
  particle_size: 'Particle size',
  physical_chemical_text: 'Physical and chemical text',
};

function App() {
  const [isDragging, setIsDragging] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [error, setError] = useState('');
  const [results, setResults] = useState([]);
  const [activeIndex, setActiveIndex] = useState(0);
  const [fileName, setFileName] = useState('');
  const [uploadProgress, setUploadProgress] = useState({ completed: 0, total: 0 });
  const [candidateSelections, setCandidateSelections] = useState({});
  const [rawLineSelections, setRawLineSelections] = useState({});
  const [excelLayout, setExcelLayout] = useState('ipd-consolidated');
  const result = results[activeIndex] || null;
  const selectedCandidateIds = candidateSelections[activeIndex] || new Set();
  const selectedRawLineIds = rawLineSelections[activeIndex] || new Set();

  function setSelectedCandidateIds(update) {
    setCandidateSelections((previous) => ({
      ...previous,
      [activeIndex]: typeof update === 'function' ? update(previous[activeIndex] || new Set()) : update,
    }));
  }

  function setSelectedRawLineIds(update) {
    setRawLineSelections((previous) => ({
      ...previous,
      [activeIndex]: typeof update === 'function' ? update(previous[activeIndex] || new Set()) : update,
    }));
  }

  async function uploadFiles(fileList) {
    const selectedFiles = Array.from(fileList || []);
    if (!selectedFiles.length) return;
    const files = selectedFiles.filter((file) => /\.(pdf|docx)$/i.test(file.name));
    const skippedCount = selectedFiles.length - files.length;
    if (!files.length) {
      setError('No supported documents were found. Use PDF or DOCX.');
      return;
    }

    setIsLoading(true);
    setError('');
    setResults([]);
    setActiveIndex(0);
    setCandidateSelections({});
    setRawLineSelections({});
    setUploadProgress({ completed: 0, total: files.length });
    setFileName(files.length === 1 ? files[0].name : `${files.length} files selected`);

    try {
      const extractedResults = [];
      const failures = [];
      for (const file of files) {
        const formData = new FormData();
        formData.append('file', file);
        try {
          const response = await fetchWithRetry(`${API_BASE}/api/extract`, {
            method: 'POST',
            body: formData,
          });
          const payload = await response.json();
          if (!response.ok) throw new Error(payload.detail || 'Extraction failed.');
          extractedResults.push(payload);
        } catch (fileError) {
          failures.push(`${file.name}: ${fileError.message}`);
        }
        setUploadProgress((previous) => ({ ...previous, completed: previous.completed + 1 }));
      }
      const nextCandidateSelections = {};
      const nextRawLineSelections = {};
      extractedResults.forEach((payload, index) => {
        nextCandidateSelections[index] = new Set((payload.key_value_candidates || []).map((_, itemIndex) => itemIndex));
        nextRawLineSelections[index] = new Set((payload.raw_lines || []).map((_, itemIndex) => itemIndex));
      });
      setCandidateSelections(nextCandidateSelections);
      setRawLineSelections(nextRawLineSelections);
      setResults(extractedResults);
      const notices = [];
      if (skippedCount) notices.push(`${skippedCount} unsupported folder item${skippedCount === 1 ? ' was' : 's were'} skipped.`);
      if (failures.length) notices.push(`Some files could not be processed. ${failures.join(' ')}`);
      if (notices.length) setError(notices.join(' '));
    } catch (err) {
      setError(err.message);
    } finally {
      setIsLoading(false);
    }
  }

  async function exportExcel() {
    if (!results.length) return;
    setIsExporting(true);
    setError('');

    try {
      const exportPayloads = results.map((item, index) => buildExportPayload(
        item,
        candidateSelections[index] || new Set(),
        rawLineSelections[index] || new Set(),
      ));
      const isConsolidated = excelLayout === 'ipd-consolidated';
      const isBatch = exportPayloads.length > 1;
      const endpoint = isConsolidated
        ? '/api/export-ipd-consolidated'
        : (isBatch ? '/api/export-excel-batch' : '/api/export-excel');
      const response = await fetchWithRetry(`${API_BASE}${endpoint}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(isConsolidated || isBatch ? exportPayloads : exportPayloads[0]),
      });
      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.detail || 'Excel export failed.');
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = isConsolidated
        ? 'nidar-ipd-consolidated.xlsx'
        : (isBatch
          ? 'product-data-batch.xlsx'
          : `${exportPayloads[0].metadata?.product_name || 'product-data-extraction'}.xlsx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError(err.message);
    } finally {
      setIsExporting(false);
    }
  }

  return (
    <main className="app-shell">
      <section className="intro">
        <p className="eyebrow">Specification review</p>
        <h1>Product Data Extractor</h1>
        <p className="subtitle">Upload a product specification and extract structured data.</p>
      </section>

      <section className="excel-layout" aria-labelledby="excel-layout-title">
        <div>
          <p className="eyebrow">Excel layout</p>
          <h2 id="excel-layout-title">Choose output structure</h2>
        </div>
        <div className="layout-options" role="group" aria-label="Excel output structure">
          <button
            type="button"
            className={excelLayout === 'ipd-consolidated' ? 'active' : ''}
            onClick={() => setExcelLayout('ipd-consolidated')}
          >
            IPD consolidated
          </button>
          <button
            type="button"
            className={excelLayout === 'document-sheets' ? 'active' : ''}
            onClick={() => setExcelLayout('document-sheets')}
          >
            Document sheets
          </button>
        </div>
      </section>

      <section
        className={`upload-zone ${isDragging ? 'dragging' : ''}`}
        onDragOver={(event) => {
          event.preventDefault();
          setIsDragging(true);
        }}
        onDragLeave={() => setIsDragging(false)}
        onDrop={(event) => {
          event.preventDefault();
          setIsDragging(false);
          uploadFiles(event.dataTransfer.files);
        }}
      >
        <UploadCloud size={42} aria-hidden="true" />
        <div>
          <h2>Drop PDF or DOCX files here</h2>
          <p>Upload one file or a batch. Documents stay inside this local app.</p>
        </div>
        <div className="upload-actions">
          <label className="upload-button secondary">
            <FileText size={18} aria-hidden="true" />
            Choose files
            <input
              type="file"
              accept=".pdf,.docx"
              multiple
              onChange={(event) => uploadFiles(event.target.files)}
            />
          </label>
          <label className="upload-button">
            <FolderOpen size={18} aria-hidden="true" />
            Choose folder
            <input
              type="file"
              accept=".pdf,.docx"
              multiple
              webkitdirectory=""
              directory=""
              onChange={(event) => uploadFiles(event.target.files)}
            />
          </label>
        </div>
        {fileName && <p className="selected-file">{fileName}</p>}
      </section>

      {isLoading && (
        <section className="status-card">
          <Loader2 className="spin" size={22} aria-hidden="true" />
          <span>Extracting file {Math.min(uploadProgress.completed + 1, uploadProgress.total)} of {uploadProgress.total}...</span>
        </section>
      )}

      {error && (
        <section className="error-card">
          <XCircle size={22} aria-hidden="true" />
          <span>{error}</span>
        </section>
      )}

      {results.length > 1 && (
        <BatchNavigation results={results} activeIndex={activeIndex} onSelect={setActiveIndex} />
      )}

      {result && (
        <ResultsDashboard
          result={result}
          onExport={exportExcel}
          batchCount={results.length}
          excelLayout={excelLayout}
          isExporting={isExporting}
          selectedCandidateIds={selectedCandidateIds}
          selectedRawLineIds={selectedRawLineIds}
          setSelectedCandidateIds={setSelectedCandidateIds}
          setSelectedRawLineIds={setSelectedRawLineIds}
          onFieldChange={(section, field, value) => {
            setResults((previous) => previous.map((item, index) => (
              index === activeIndex ? updateNestedField(item, section, field, value) : item
            )));
          }}
          onDynamicFieldChange={(index, value) => {
            setResults((previous) => previous.map((item, resultIndex) => (
              resultIndex === activeIndex ? updateDynamicField(item, index, value) : item
            )));
          }}
          onIpdRowChange={(rowIndex, field, value) => {
            setResults((previous) => previous.map((item, resultIndex) => (
              resultIndex === activeIndex ? updateIpdRow(item, rowIndex, field, value) : item
            )));
          }}
          onApproveFound={() => {
            setResults((previous) => previous.map((item, resultIndex) => (
              resultIndex === activeIndex ? approveFoundIpdRows(item) : item
            )));
          }}
        />
      )}
    </main>
  );
}

function BatchNavigation({ results, activeIndex, onSelect }) {
  return (
    <nav className="batch-navigation" aria-label="Uploaded documents">
      <div>
        <p className="eyebrow">Batch review</p>
        <h2>{results.length} documents extracted</h2>
      </div>
      <div className="batch-tabs">
        {results.map((item, index) => (
          <button
            type="button"
            className={index === activeIndex ? 'active' : ''}
            onClick={() => onSelect(index)}
            key={`${item.metadata?.source_file || 'document'}-${index}`}
          >
            <span>{index + 1}</span>
            {item.metadata?.document_id || item.metadata?.product_name || `Document ${index + 1}`}
          </button>
        ))}
      </div>
    </nav>
  );
}

function ResultsDashboard({
  result,
  onExport,
  isExporting,
  selectedCandidateIds,
  selectedRawLineIds,
  setSelectedCandidateIds,
  setSelectedRawLineIds,
  onFieldChange,
  onDynamicFieldChange,
  batchCount,
  excelLayout,
  onIpdRowChange,
  onApproveFound,
}) {
  const visibleSections = ['food_ipd', 'gnt_exberry'].includes(result.output_profile) ? sections : sections.slice(0, 1);
  const profileName = profileLabel(result.output_profile);
  return (
    <section className="results">
      <div className="results-header">
        <div>
          <p className="eyebrow">Extraction result</p>
          <h2>{valueOrMissing(result.metadata?.product_name)}</h2>
          <p className="capture-summary">
            {result.raw_lines?.length || 0} raw lines captured - {result.key_value_candidates?.length || 0} key-value candidates
            {' '}({selectedCandidateIds.size + selectedRawLineIds.size} selected for Excel)
          </p>
        </div>
        <button className="export-button" onClick={onExport} disabled={isExporting}>
          {isExporting ? <Loader2 className="spin" size={18} /> : <Download size={18} />}
          {excelLayout === 'ipd-consolidated'
            ? `Export IPD overview (${batchCount})`
            : (batchCount > 1 ? `Export ${batchCount} files` : 'Export to Excel')}
        </button>
      </div>

      <section className="ipd-panel">
        <div>
          <p className="eyebrow">Excel output</p>
          <h3>{profileName}</h3>
        </div>
        <p>
          {excelLayout === 'ipd-consolidated'
            ? 'All documents will be exported into one reviewable IPD workbook with field status, source references, approvals, and change history.'
            : profileDescription(result.output_profile)}
        </p>
      </section>

      {result.ipd_rows?.length > 0 && (
        <IPDReviewTable
          rows={result.ipd_rows}
          onRowChange={onIpdRowChange}
          onApproveFound={onApproveFound}
        />
      )}

      {result.specification_rows?.length > 0 && (
        <SpecificationTable rows={result.specification_rows} />
      )}

      {result.dynamic_fields?.length > 0 && (
        <DynamicFields fields={result.dynamic_fields} onFieldChange={onDynamicFieldChange} />
      )}

      <div className="dashboard-grid">
        {visibleSections.map((section) => (
          <DataSection
            key={section.key}
            title={section.title}
            sectionKey={section.key}
            data={result[section.key]}
            onFieldChange={onFieldChange}
          />
        ))}
      </div>

      <section className="warning-panel">
        <div className="section-heading">
          <AlertTriangle size={19} aria-hidden="true" />
          <h3>Extraction warnings</h3>
        </div>
        {result.extraction_warnings?.length ? (
          <ul>
            {result.extraction_warnings.map((warning) => (
              <li key={warning}>{warning}</li>
            ))}
          </ul>
        ) : (
          <p>No warnings.</p>
        )}
      </section>

      {result.review_flags?.length > 0 && (
        <section className="warning-panel">
          <div className="section-heading">
            <AlertTriangle size={19} aria-hidden="true" />
            <h3>Review flags</h3>
          </div>
          <ul>
            {result.review_flags.map((flag) => (
              <li key={flag}>{flag}</li>
            ))}
          </ul>
        </section>
      )}

      <OtherExtractedData
        result={result}
        selectedCandidateIds={selectedCandidateIds}
        selectedRawLineIds={selectedRawLineIds}
        setSelectedCandidateIds={setSelectedCandidateIds}
        setSelectedRawLineIds={setSelectedRawLineIds}
      />

      <section className="raw-panel">
        <h3>Raw review excerpt</h3>
        <pre>{valueOrMissing(result.raw_text_excerpt)}</pre>
      </section>
    </section>
  );
}

function IPDReviewTable({ rows, onRowChange, onApproveFound }) {
  const [view, setView] = useState('attention');
  const [query, setQuery] = useState('');
  const found = rows.filter((row) => row.data).length;
  const approved = rows.filter((row) => row.data && row.approved).length;
  const missing = rows.length - found;
  const normalizedQuery = query.trim().toLowerCase();
  const entries = rows
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => {
      if (view === 'found' && !row.data) return false;
      if (view === 'missing' && row.data) return false;
      if (view === 'approved' && !row.approved) return false;
      if (view === 'attention' && row.approved) return false;
      if (!normalizedQuery) return true;
      return `${row.node} ${row.attribute} ${row.attribute_description || ''}`.toLowerCase().includes(normalizedQuery);
    })
    .sort((left, right) => statusPriority(left.row) - statusPriority(right.row));

  return (
    <section className="ipd-review">
      <div className="review-header">
        <div>
          <p className="eyebrow">Migration control</p>
          <h3>IPD attribute review</h3>
        </div>
        <button className="approve-button" type="button" onClick={onApproveFound} disabled={!found || approved === found}>
          <CheckCircle2 size={18} aria-hidden="true" />
          Approve all found
        </button>
      </div>

      <div className="review-metrics" aria-label="Extraction coverage">
        <div><strong>{rows.length}</strong><span>Total attributes</span></div>
        <div><strong>{found}</strong><span>Found</span></div>
        <div><strong>{missing}</strong><span>Not found</span></div>
        <div><strong>{approved}</strong><span>Approved</span></div>
        <div><strong>{rows.length ? Math.round((found / rows.length) * 100) : 0}%</strong><span>Coverage</span></div>
      </div>

      <div className="review-toolbar">
        <div className="review-tabs" role="group" aria-label="Filter IPD attributes">
          {[
            ['attention', 'Needs attention'],
            ['found', 'Found'],
            ['missing', 'Not found'],
            ['approved', 'Approved'],
            ['all', 'All'],
          ].map(([value, label]) => (
            <button
              type="button"
              className={view === value ? 'active' : ''}
              onClick={() => setView(value)}
              key={value}
            >
              {label}
            </button>
          ))}
        </div>
        <input
          className="review-search"
          type="search"
          value={query}
          placeholder="Search attributes"
          aria-label="Search IPD attributes"
          onChange={(event) => setQuery(event.target.value)}
        />
      </div>

      <div className="table-wrap review-table-wrap">
        <table className="review-table">
          <thead>
            <tr>
              <th>Status</th>
              <th>Node</th>
              <th>Attribute</th>
              <th>Data</th>
              <th>Unit</th>
              <th>Source</th>
              <th className="select-cell">Approve</th>
            </tr>
          </thead>
          <tbody>
            {entries.map(({ row, index }) => (
              <tr key={`${row.node}-${row.attribute}-${row.attribute_description}-${index}`}>
                <td><span className={`status-badge ${statusClass(row)}`}>{row.approved ? 'Approved' : row.status}</span></td>
                <td>{row.node}</td>
                <td>
                  <strong>{row.attribute}</strong>
                  {row.attribute_description && <small>{row.attribute_description}</small>}
                </td>
                <td>
                  <input
                    value={row.data ?? ''}
                    placeholder="Not found"
                    aria-label={`${row.attribute} data`}
                    onChange={(event) => onRowChange(index, 'data', event.target.value)}
                  />
                </td>
                <td>
                  <input
                    value={row.uom ?? ''}
                    aria-label={`${row.attribute} unit`}
                    onChange={(event) => onRowChange(index, 'uom', event.target.value)}
                  />
                </td>
                <td className="source-cell">
                  <strong>{row.source_reference || 'Not provided in document'}</strong>
                  {row.source_excerpt && <span>{row.source_excerpt}</span>}
                </td>
                <td className="select-cell">
                  <input
                    type="checkbox"
                    checked={Boolean(row.approved)}
                    disabled={!row.data}
                    aria-label={`Approve ${row.attribute}`}
                    onChange={(event) => onRowChange(index, 'approved', event.target.checked)}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="panel-note">Showing {entries.length} of {rows.length} attributes. Missing values remain blank and are never inferred.</p>
    </section>
  );
}

function statusPriority(row) {
  if (!row.data) return 0;
  if (!row.approved) return 1;
  return 2;
}

function statusClass(row) {
  if (row.approved) return 'approved';
  return row.data ? 'found' : 'missing';
}

function SpecificationTable({ rows }) {
  return (
    <section className="spec-panel">
      <div className="other-header">
        <div>
          <p className="eyebrow">Source specification</p>
          <h3>Specification table</h3>
        </div>
        <span>{rows.length} rows captured</span>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Section</th>
              <th>Parameter</th>
              <th>Value</th>
              <th>Min</th>
              <th>Max</th>
              <th>Unit</th>
              <th>Qualifier</th>
              <th>Source</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, index) => (
              <tr className={row.needs_review ? 'spec-review-row' : ''} key={`${row.section}-${row.parameter}-${index}`}>
                <td>{row.section}</td>
                <td><strong>{row.parameter}</strong></td>
                <td>{valueOrMissing(row.value)}</td>
                <td>{valueOrMissing(row.min_value)}</td>
                <td>{valueOrMissing(row.max_value)}</td>
                <td>{valueOrMissing(row.unit)}</td>
                <td>{valueOrMissing(row.qualifier)}</td>
                <td className="source-cell">
                  <strong>{row.source_reference || 'Source document'}</strong>
                  {row.source_excerpt && <span>{row.source_excerpt}</span>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="panel-note">Values remain in the supplier's original unit and limit form. Highlighted rows need OCR review.</p>
    </section>
  );
}

function DynamicFields({ fields, onFieldChange }) {
  const groups = fields.reduce((result, field, index) => {
    const section = field.section || 'Other extracted data';
    if (!result[section]) result[section] = [];
    result[section].push({ field, index });
    return result;
  }, {});

  return (
    <section className="dynamic-panel">
      <div className="other-header">
        <div>
          <p className="eyebrow">Schema fields</p>
          <h3>Extracted document fields</h3>
        </div>
        <span>{fields.length} fields</span>
      </div>
      <p className="panel-note">Review highlighted OCR values before export. Every edit is written to the matching Excel column.</p>
      <div className="dynamic-groups">
        {Object.entries(groups).map(([section, entries]) => (
          <section className="dynamic-group" key={section}>
            <h4>{section}</h4>
            <div className="dynamic-grid">
              {entries.map(({ field, index }) => (
                <label className={`editable-field ${field.needs_review ? 'needs-review' : ''}`} key={`${field.field_name}-${index}`}>
                  <span>{field.label}{field.unit ? ` (${field.unit})` : ''}</span>
                  <input
                    value={field.value ?? ''}
                    placeholder="Not found"
                    onChange={(event) => onFieldChange(index, event.target.value)}
                  />
                  {field.needs_review && <small>Check OCR value</small>}
                </label>
              ))}
            </div>
          </section>
        ))}
      </div>
    </section>
  );
}

function OtherExtractedData({
  result,
  selectedCandidateIds,
  selectedRawLineIds,
  setSelectedCandidateIds,
  setSelectedRawLineIds,
}) {
  const candidates = result.key_value_candidates || [];
  const candidateEntries = candidates.map((candidate, index) => ({ candidate, index }));
  const allergenCandidates = candidateEntries.filter(({ candidate }) => isAllergenCandidate(candidate));
  const otherCandidates = candidateEntries.filter(({ candidate }) => !isAllergenCandidate(candidate)).slice(0, 40);
  const rawHighlights = (result.raw_lines || [])
    .map((line, index) => ({ line, index }))
    .filter(({ line }) => ['allergens', 'nutrition', 'shelf_life', 'storage', 'physical_chemical'].includes(line.section))
    .slice(0, 80);

  if (!candidates.length && !rawHighlights.length) {
    return null;
  }

  return (
    <section className="other-panel">
      <div className="other-header">
        <div>
          <p className="eyebrow">Captured content</p>
          <h3>Other extracted data</h3>
        </div>
        <span>{candidates.length} candidates</span>
      </div>

      <div className="selection-tools">
        <button type="button" onClick={() => setSelectedCandidateIds(new Set(candidates.map((_, index) => index)))}>
          Select candidates
        </button>
        <button type="button" onClick={() => setSelectedCandidateIds(new Set())}>
          Clear candidates
        </button>
        <button type="button" onClick={() => setSelectedRawLineIds(new Set((result.raw_lines || []).map((_, index) => index)))}>
          Select raw lines
        </button>
        <button type="button" onClick={() => setSelectedRawLineIds(new Set())}>
          Clear raw lines
        </button>
      </div>

      {allergenCandidates.length > 0 && (
        <div className="table-block">
          <h4>Allergen checklist candidates</h4>
          <CandidateTable
            entries={allergenCandidates.slice(0, 30)}
            selectedCandidateIds={selectedCandidateIds}
            setSelectedCandidateIds={setSelectedCandidateIds}
            columns={['Key', 'Value', 'Source', 'Line']}
            rowFor={({ candidate }) => [candidate.key, candidate.value, candidate.source_reference, candidate.source_line]}
          />
        </div>
      )}

      {otherCandidates.length > 0 && (
        <div className="table-block">
          <h4>Other key-value candidates</h4>
          <CandidateTable
            entries={otherCandidates}
            selectedCandidateIds={selectedCandidateIds}
            setSelectedCandidateIds={setSelectedCandidateIds}
            columns={['Section', 'Key', 'Value', 'Source', 'Line']}
            rowFor={({ candidate }) => [
              candidate.section || 'Unclassified',
              candidate.key,
              candidate.value,
              candidate.source_reference,
              candidate.source_line,
            ]}
          />
        </div>
      )}

      {rawHighlights.length > 0 && (
        <div className="table-block">
          <h4>Raw captured lines</h4>
          <RawLineTable
            entries={rawHighlights}
            selectedRawLineIds={selectedRawLineIds}
            setSelectedRawLineIds={setSelectedRawLineIds}
          />
        </div>
      )}
    </section>
  );
}

function CandidateTable({ entries, selectedCandidateIds, setSelectedCandidateIds, columns, rowFor }) {
  return (
    <SelectableTable
      columns={columns}
      entries={entries}
      isSelected={({ index }) => selectedCandidateIds.has(index)}
      onToggle={({ index }) => toggleSetValue(setSelectedCandidateIds, index)}
      rowFor={rowFor}
    />
  );
}

function RawLineTable({ entries, selectedRawLineIds, setSelectedRawLineIds }) {
  return (
    <SelectableTable
      columns={['Line', 'Section', 'Source', 'Text']}
      entries={entries}
      isSelected={({ index }) => selectedRawLineIds.has(index)}
      onToggle={({ index }) => toggleSetValue(setSelectedRawLineIds, index)}
      rowFor={({ line }) => [line.line_number, line.section || 'Unclassified', line.source_reference, line.text]}
    />
  );
}

function SelectableTable({ columns, entries, isSelected, onToggle, rowFor }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th className="select-cell">Export</th>
            {columns.map((column) => (
              <th key={column}>{column}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {entries.map((entry) => {
            const row = rowFor(entry);
            return (
              <tr key={`${entry.index}-${row.join('|')}`}>
                <td className="select-cell">
                  <input
                    type="checkbox"
                    checked={isSelected(entry)}
                    onChange={() => onToggle(entry)}
                    aria-label="Include row in Excel export"
                  />
                </td>
                {row.map((cell, cellIndex) => (
                  <td key={`${entry.index}-${cellIndex}`}>{valueOrMissing(cell)}</td>
                ))}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

async function fetchWithRetry(url, options) {
  try {
    return await fetch(url, options);
  } catch (error) {
    await new Promise((resolve) => setTimeout(resolve, 800));
    return fetch(url, options);
  }
}

function buildExportPayload(result, selectedCandidateIds, selectedRawLineIds) {
  return {
    ...result,
    key_value_candidates: (result.key_value_candidates || []).filter((_, index) => selectedCandidateIds.has(index)),
    raw_lines: (result.raw_lines || []).filter((_, index) => selectedRawLineIds.has(index)),
  };
}

function toggleSetValue(setter, value) {
  setter((previous) => {
    const next = new Set(previous);
    if (next.has(value)) {
      next.delete(value);
    } else {
      next.add(value);
    }
    return next;
  });
}

function isAllergenCandidate(candidate) {
  const knownAllergens = [
    'gluten',
    'skalldyr',
    'egg',
    'fisk',
    'peanotter',
    'soya',
    'melk',
    'milk',
    'notter',
    'selleri',
    'seller',
    'sennep',
    'sesam',
    'sulfitt',
    'suftt',
    'lupin',
    'blotdyr',
  ];
  const key = searchableText(String(candidate.key || ''));
  const value = searchableText(String(candidate.value || ''));
  const looksLikeYesNo = /\b(ja|yes|nei|no|ye|nel|net)\b/.test(value);
  return knownAllergens.some((allergen) => key.includes(allergen)) && (candidate.section === 'allergens' || looksLikeYesNo);
}

function searchableText(value) {
  return value
    .replaceAll('\u00e6', 'ae')
    .replaceAll('\u00c6', 'ae')
    .replaceAll('\u00f8', 'o')
    .replaceAll('\u00d8', 'o')
    .replaceAll('\u00e5', 'a')
    .replaceAll('\u00c5', 'a')
    .toLowerCase();
}

function DataSection({ title, sectionKey, data = {}, onFieldChange }) {
  return (
    <section className="data-card">
      <h3>{title}</h3>
      <div className="editable-grid">
        {Object.entries(data).map(([key, value]) => (
          <label className="editable-field" key={key}>
            <span>{labels[key] || key}</span>
            {isLongField(key, value) ? (
              <textarea
                value={value ?? ''}
                placeholder="Not found"
                rows={4}
                onChange={(event) => onFieldChange(sectionKey, key, event.target.value)}
              />
            ) : (
              <input
                value={value ?? ''}
                placeholder="Not found"
                onChange={(event) => onFieldChange(sectionKey, key, event.target.value)}
              />
            )}
          </label>
        ))}
      </div>
    </section>
  );
}

function isLongField(key, value) {
  return key.endsWith('_text') || key.includes('statement') || String(value || '').length > 80;
}

function updateIpdRow(result, rowIndex, field, value) {
  const current = result.ipd_rows?.[rowIndex];
  if (!current) return result;
  const normalizedValue = field === 'approved' ? Boolean(value) : (value === '' ? null : value);
  const nextRow = { ...current, [field]: normalizedValue };
  if (field !== 'approved') nextRow.approved = false;
  nextRow.status = nextRow.approved ? 'Approved' : (nextRow.data ? 'Found' : 'Not found');

  return {
    ...result,
    ipd_rows: result.ipd_rows.map((row, index) => (index === rowIndex ? nextRow : row)),
    change_log: [
      ...(result.change_log || []),
      {
        changed_at: new Date().toISOString(),
        actor: 'Local reviewer',
        node: current.node,
        attribute: current.attribute,
        field,
        old_value: String(current[field] ?? ''),
        new_value: String(normalizedValue ?? ''),
      },
    ],
  };
}

function approveFoundIpdRows(result) {
  const changedAt = new Date().toISOString();
  const entries = [];
  const rows = (result.ipd_rows || []).map((row) => {
    if (!row.data || row.approved) return row;
    entries.push({
      changed_at: changedAt,
      actor: 'Local reviewer',
      node: row.node,
      attribute: row.attribute,
      field: 'approved',
      old_value: 'false',
      new_value: 'true',
    });
    return { ...row, approved: true, status: 'Approved' };
  });
  return { ...result, ipd_rows: rows, change_log: [...(result.change_log || []), ...entries] };
}

function updateNestedField(result, section, field, value) {
  return {
    ...result,
    [section]: {
      ...(result?.[section] || {}),
      [field]: value.trim() === '' ? null : value,
    },
  };
}

function updateDynamicField(result, index, value) {
  return {
    ...result,
    dynamic_fields: (result.dynamic_fields || []).map((field, fieldIndex) => (
      fieldIndex === index
        ? { ...field, value: value.trim() === '' ? null : value, needs_review: false }
        : field
    )),
  };
}

function profileLabel(profile) {
  if (profile === 'roller_fabric') return 'Roller fabric specification';
  if (profile === 'filament_spec') return 'Filament specification';
  if (profile === 'food_ipd') return 'IPD food migration template';
  if (profile === 'gnt_exberry') return 'GNT Exberry specification template';
  return 'Flexible extracted data workbook';
}

function profileDescription(profile) {
  if (profile === 'roller_fabric') {
    return 'Export uses the customer\'s 65-column fabric schema and includes its schema dictionary plus OCR review sheets.';
  }
  if (profile === 'filament_spec') {
    return 'Export uses the customer\'s 123-field filament schema and includes its schema dictionary plus source review sheets.';
  }
  if (profile === 'food_ipd') {
    return 'Export uses the customer IPD attribute rows and includes raw OCR lines and selected candidates for review.';
  }
  if (profile === 'gnt_exberry') {
    return 'Recognized GNT Exberry tables are captured as structured specification rows and mapped into the IPD workbook where fields match.';
  }
  return 'No known schema was detected. Export preserves dynamic fields, selected candidates, and raw source lines.';
}

function valueOrMissing(value) {
  if (value === null || value === undefined || value === '') {
    return 'Not found';
  }
  return Array.isArray(value) ? value.join(', ') : value;
}

createRoot(document.getElementById('root')).render(<App />);
