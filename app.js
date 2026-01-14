import { unpickleDataFrameToRecords } from './pyodide-loader.js?v=30';
import { buildPivot, renderPivotGrid } from './pivot.js?v=30';

const els = {
  fileInput: document.getElementById('fileInput'),
  statusText: document.getElementById('statusText'),
  columnsPreview: document.getElementById('columnsPreview'),
  gridContainer: document.getElementById('gridContainer'),
  gridSummary: document.getElementById('gridSummary'),
  zoomSelect: document.getElementById('zoomSelect'),
  exportExcel: document.getElementById('exportExcel'),
  debugLog: document.getElementById('debugLog'),
  filtersContainer: document.getElementById('filtersContainer'),
  filtersHint: document.getElementById('filtersHint'),
  clearFilters: document.getElementById('clearFilters'),

  buildingSelect: document.getElementById('buildingSelect'),
  selectAllBuildings: document.getElementById('selectAllBuildings'),
  clearBuildings: document.getElementById('clearBuildings'),
  buildingText: document.getElementById('buildingText'),
  applyBuildingText: document.getElementById('applyBuildingText'),
  clearBuildingText: document.getElementById('clearBuildingText'),
};

const state = {
  columns: [],
  records: [],
  filteredRecords: [],

  lastPivot: null,
  lastRowHeaderCols: null,

  dimCols: {
    stage: null,
    building: null,
    participant: null,
    os: null,
    row_type: null,
    id: null,
  },

  // Standard filters (empty = All). Building is required if present.
  filters: {
    participant: new Set(),
    stage: new Set(),
    building: new Set(),
    os: new Set(),
    row_type: new Set(),
  },

  // Per-section ID filters (Section value -> Set(ID values)).
  // Only active when one or more Sections are explicitly selected.
  idBySection: new Map(),

  // Manual building override input support
  knownBuildings: [],
  knownBuildingsLowerMap: new Map(),
};

const METRICS = ['h80', 'v80'];
const METRIC_LABELS = {
  h80: 'Hor 80%',
  v80: 'Ver 80%',
};

const SECTION_ORDER = ['Results', 'Location Technology', 'Handset', 'Path', 'Point'];

function sortByPreferredOrder(values, preferredOrder) {
  const pref = new Map(preferredOrder.map((v, i) => [String(v).toLowerCase(), i]));
  return Array.from(values).sort((a, b) => {
    const ak = String(a ?? '').toLowerCase();
    const bk = String(b ?? '').toLowerCase();
    const ar = pref.has(ak) ? pref.get(ak) : Number.POSITIVE_INFINITY;
    const br = pref.has(bk) ? pref.get(bk) : Number.POSITIVE_INFINITY;
    if (ar !== br) return ar - br;
    return String(a ?? '').localeCompare(String(b ?? ''));
  });
}

function sortPivotRowsByRowKeys(pivot, { preferredOrderByKey = {} } = {}) {
  if (!pivot || !Array.isArray(pivot.rows) || !pivot.rowMeta || !Array.isArray(pivot.rowKeys)) return;
  if (pivot.rows.length <= 1) return;

  const orderMaps = new Map();
  for (const [key, order] of Object.entries(preferredOrderByKey)) {
    if (!key || !Array.isArray(order)) continue;
    orderMaps.set(
      key,
      new Map(order.map((v, i) => [String(v ?? '').toLowerCase(), i]))
    );
  }

  const compareValues = (key, aVal, bVal) => {
    const map = orderMaps.get(key);
    if (map) {
      const ak = String(aVal ?? '').toLowerCase();
      const bk = String(bVal ?? '').toLowerCase();
      const ar = map.has(ak) ? map.get(ak) : Number.POSITIVE_INFINITY;
      const br = map.has(bk) ? map.get(bk) : Number.POSITIVE_INFINITY;
      if (ar !== br) return ar - br;
    }
    return String(aVal ?? '').localeCompare(String(bVal ?? ''));
  };

  pivot.rows.sort((aRowId, bRowId) => {
    const aMeta = pivot.rowMeta.get(aRowId) ?? {};
    const bMeta = pivot.rowMeta.get(bRowId) ?? {};

    for (const key of pivot.rowKeys) {
      const cmp = compareValues(key, aMeta[key], bMeta[key]);
      if (cmp) return cmp;
    }

    // Final fallback: stable, deterministic ordering.
    return String(aRowId ?? '').localeCompare(String(bRowId ?? ''));
  });
}

const PARTICIPANT_ID_SEP = '\u0001';

function makeParticipantIdKey(participant, identifier) {
  return `${toKey(participant)}${PARTICIPANT_ID_SEP}${toKey(identifier)}`;
}

function setStatus(text, { error = false } = {}) {
  if (!els.statusText) return;
  els.statusText.textContent = text;
  els.statusText.style.color = error ? 'var(--danger)' : 'var(--text)';
}

function logDebug(message) {
  if (!els.debugLog) return;
  const line = `[${new Date().toLocaleTimeString()}] ${message}`;
  els.debugLog.textContent += (els.debugLog.textContent ? '\n' : '') + line;
}

// Visible startup confirmation (helps diagnose caching / module-load issues).
setStatus('Ready. Choose a .pkl file to begin.');
logDebug('app.js initialized.');

function setGridZoom(value) {
  const v = Number(value);
  if (!Number.isFinite(v)) return;
  const clamped = Math.max(0.5, Math.min(1, v));
  document.documentElement.style.setProperty('--grid-zoom', String(clamped));
}

// Restore grid zoom preference early.
try {
  const saved = localStorage.getItem('resultsArchive.gridZoom');
  if (saved) setGridZoom(saved);
} catch {
  // ignore
}

function detectColumn(columns, candidates) {
  const lower = columns.map((c) => String(c).toLowerCase());
  for (const cand of candidates) {
    const idx = lower.indexOf(cand.toLowerCase());
    if (idx >= 0) return columns[idx];
  }
  // Fallback: contains-match
  for (const cand of candidates) {
    const idx = lower.findIndex((c) => c.includes(cand.toLowerCase()));
    if (idx >= 0) return columns[idx];
  }
  return null;
}

function setExportEnabled(enabled) {
  if (!els.exportExcel) return;
  els.exportExcel.disabled = !enabled;
}

function safeFilePart(s) {
  return String(s ?? '')
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, ' ')
    .slice(0, 80);
}

function formatDateForFilename(d) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
}

function formatDateYmdHm(d) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function buildFiltersSummaryAoA() {
  const lines = [];
  const setToText = (set) => {
    if (!set) return '';
    if (set.size === 0) return 'All';
    return Array.from(set).join(', ');
  };

  lines.push(['Results Archive Export']);
  lines.push(['Exported at', new Date().toLocaleString()]);
  lines.push(['Build', document.getElementById('buildStamp')?.textContent ?? '']);
  lines.push([]);

  lines.push(['Buildings', setToText(state.filters.building)]);
  lines.push(['Participant', setToText(state.filters.participant)]);
  lines.push(['OS', setToText(state.filters.os)]);
  lines.push(['Stage', setToText(state.filters.stage)]);
  lines.push(['Section', setToText(state.filters.row_type)]);

  if (state.idBySection && state.idBySection.size) {
    lines.push([]);
    lines.push(['Identifier filters']);
    const sections = sortByPreferredOrder(Array.from(state.idBySection.keys()), SECTION_ORDER);
    for (const sec of sections) {
      const set = state.idBySection.get(sec);
      if (!set || set.size === 0) continue;
      lines.push([`Identifier — ${sec}`, `${set.size} selected`]);
    }
  }

  return lines;
}

function exportCurrentPivotToExcel() {
  const XLSX = window.XLSX;
  if (!XLSX) {
    setStatus('Excel export library not loaded yet. Please refresh and try again.', { error: true });
    return;
  }

  const pivot = state.lastPivot;
  const rowHeaderCols = state.lastRowHeaderCols;
  if (!pivot || !Array.isArray(pivot.cols) || !Array.isArray(pivot.rows) || !rowHeaderCols) {
    setStatus('Nothing to export yet (load data + select building(s)).', { error: true });
    return;
  }

  const leftCols = rowHeaderCols.map((c) => (typeof c === 'string' ? ({ key: c, label: c }) : c));
  const leftCount = leftCols.length;
  const stages = pivot.cols;
  const metricKeys = METRICS;
  const metricCount = metricKeys.length;

  const lastCol = leftCount + stages.length * metricCount - 1;

  // Title + generated rows (like the screenshot).
  const now = new Date();
  const titleRow = ['Stage Comparison', ...Array(lastCol).fill('')];
  const generatedRow = [`Generated ${formatDateYmdHm(now)}`, ...Array(lastCol).fill('')];
  const spacerRow = Array(lastCol + 1).fill('');

  // Build a two-row header like the screenshot:
  // Row 1: left column labels + merged stage group headers.
  // Row 2: (blank under left headers) + metric subheaders (Hor/Ver 80%).
  const headerTop = leftCols.map((c) => String(c?.label ?? c?.key ?? ''));
  for (const s of stages) {
    const stageLabel = String(s).toLowerCase().startsWith('stage ') ? String(s) : `Stage ${s}`;
    headerTop.push(stageLabel);
    for (let i = 1; i < metricCount; i++) headerTop.push('');
  }

  const headerSub = Array(leftCount).fill('');
  for (const _s of stages) {
    for (const m of metricKeys) {
      headerSub.push(String(METRIC_LABELS[m] ?? m));
    }
  }

  const aoa = [titleRow, generatedRow, spacerRow, headerTop, headerSub];

  for (const rowId of pivot.rows) {
    const meta = pivot.rowMeta?.get(rowId) ?? {};
    const row = [];
    for (const c of leftCols) {
      const v = meta?.[c.key];
      row.push(v === null || v === undefined ? '' : String(v));
    }
    const rowMap = pivot.matrix?.get(rowId);
    for (const s of stages) {
      const raw = rowMap ? rowMap.get(String(s)) : undefined;
      for (const m of metricKeys) {
        const v = raw && typeof raw === 'object' ? raw[m] : undefined;
        const num = typeof v === 'number' ? v : Number(v);
        row.push(Number.isFinite(num) ? num : '');
      }
    }
    aoa.push(row);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  const lastRow = aoa.length - 1;

  // Merges: title + generated lines, vertical merges for left headers, and stage group merges.
  ws['!merges'] = ws['!merges'] || [];
  ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: lastCol } });
  ws['!merges'].push({ s: { r: 1, c: 0 }, e: { r: 1, c: lastCol } });

  const HEADER_TOP_ROW = 3;
  const HEADER_SUB_ROW = 4;
  const DATA_START_ROW = 5;

  // Merge each left header cell down across the two header rows.
  for (let c = 0; c < leftCount; c++) {
    ws['!merges'].push({ s: { r: HEADER_TOP_ROW, c }, e: { r: HEADER_SUB_ROW, c } });
  }

  // Merge stage group headers across their metric columns in the top header row.
  for (let i = 0; i < stages.length; i++) {
    const start = leftCount + i * metricCount;
    const end = start + metricCount - 1;
    if (end > start) ws['!merges'].push({ s: { r: HEADER_TOP_ROW, c: start }, e: { r: HEADER_TOP_ROW, c: end } });
  }

  // Freeze panes so the row with the 80%s stays frozen.
  // Freezing at HEADER_SUB_ROW + 1 freezes title + generated + spacer + both header rows.
  const ySplit = HEADER_SUB_ROW + 1;
  const topLeftCell = XLSX.utils.encode_cell({ r: ySplit, c: leftCount });
  ws['!sheetViews'] = [{ pane: { state: 'frozen', xSplit: leftCount, ySplit, topLeftCell, activePane: 'bottomRight' } }];

  // Column widths (roughly based on text length).
  const maxChars = new Array(lastCol + 1).fill(6);
  const sampleRows = Math.min(aoa.length, 400); // cap work
  for (let r = 0; r < sampleRows; r++) {
    const row = aoa[r];
    for (let c = 0; c <= lastCol; c++) {
      const v = row?.[c];
      const s = v === null || v === undefined ? '' : String(v);
      maxChars[c] = Math.min(60, Math.max(maxChars[c], s.length));
    }
  }
  ws['!cols'] = maxChars.map((wch, i) => {
    // Give left columns a bit more room.
    const bonus = i < leftCount ? 4 : 0;
    return { wch: Math.min(64, Math.max(8, wch + bonus)) };
  });

  // Styling (xlsx-js-style) — match the screenshot style (bold title, dark blue headers, crisp borders).
  const BORDER_THIN_BLACK = {
    top: { style: 'thin', color: { rgb: 'FF000000' } },
    bottom: { style: 'thin', color: { rgb: 'FF000000' } },
    left: { style: 'thin', color: { rgb: 'FF000000' } },
    right: { style: 'thin', color: { rgb: 'FF000000' } },
  };

  const HEADER_FILL = 'FF0F3B5E';

  const STYLE_TITLE = {
    font: { bold: true, sz: 18, color: { rgb: 'FF000000' } },
    alignment: { horizontal: 'left', vertical: 'center' },
  };

  const STYLE_GENERATED = {
    font: { italic: true, sz: 10, color: { rgb: 'FF333333' } },
    alignment: { horizontal: 'left', vertical: 'center' },
  };

  const STYLE_HDR = {
    font: { bold: true, color: { rgb: 'FFFFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: HEADER_FILL } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: BORDER_THIN_BLACK,
  };

  const STYLE_TEXT = {
    alignment: { horizontal: 'left', vertical: 'top' },
    border: BORDER_THIN_BLACK,
  };

  const STYLE_NUM = {
    alignment: { horizontal: 'right', vertical: 'top' },
    border: BORDER_THIN_BLACK,
    numFmt: '0.0',
  };

  const applyStyle = (r, c, style) => {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = ws[addr];
    if (!cell) return;
    cell.s = { ...(cell.s || {}), ...(style || {}) };
  };

  // Row heights
  ws['!rows'] = ws['!rows'] || [];
  ws['!rows'][0] = { hpt: 28 };
  ws['!rows'][1] = { hpt: 16 };
  ws['!rows'][2] = { hpt: 8 };
  ws['!rows'][HEADER_TOP_ROW] = { hpt: 22 };
  ws['!rows'][HEADER_SUB_ROW] = { hpt: 20 };

  // Title / generated styles (merged across)
  for (let c = 0; c <= lastCol; c++) {
    applyStyle(0, c, STYLE_TITLE);
    applyStyle(1, c, STYLE_GENERATED);
  }

  // Header styles (both header rows)
  for (let c = 0; c <= lastCol; c++) {
    applyStyle(HEADER_TOP_ROW, c, STYLE_HDR);
    applyStyle(HEADER_SUB_ROW, c, STYLE_HDR);
  }

  // Data styles
  for (let r = DATA_START_ROW; r <= lastRow; r++) {
    for (let c = 0; c <= lastCol; c++) {
      const isMetric = c >= leftCount;
      applyStyle(r, c, isMetric ? STYLE_NUM : STYLE_TEXT);
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Stage Comparison');

  // Filters/metadata sheet
  const ws2 = XLSX.utils.aoa_to_sheet(buildFiltersSummaryAoA());
  ws2['!cols'] = [{ wch: 24 }, { wch: 120 }];
  XLSX.utils.book_append_sheet(wb, ws2, 'Filters');

  const dt = new Date();
  const buildingPart = state.filters.building && state.filters.building.size
    ? `_${safeFilePart(Array.from(state.filters.building).slice(0, 3).join('-'))}${state.filters.building.size > 3 ? '_and_more' : ''}`
    : '';
  const filename = `Stage_Comparison_${formatDateForFilename(dt)}${buildingPart}.xlsx`;

  try {
    XLSX.writeFile(wb, filename, { compression: true });
    setStatus(`Exported Excel: ${filename}`);
  } catch (err) {
    console.error(err);
    setStatus(`Excel export failed: ${err?.message ?? String(err)}`, { error: true });
  }
}

function guessDimensionColumns(columns) {
  return {
    stage: detectColumn(columns, ['stage', 'stg', 'phase']),
    building: detectColumn(columns, ['building_id', 'building', 'bldg', 'site', 'location']),
    participant: detectColumn(columns, ['participant', 'carrier', 'name', 'user', 'person']),
    os: detectColumn(columns, ['os', 'operating system', 'platform']),
    row_type: detectColumn(columns, ['row_type', 'row type', 'type', 'section']),
    id: detectColumn(columns, ['id', 'label', 'name']),
  };
}

function toKey(v) {
  if (v === null || v === undefined) return '';
  return String(v);
}

function uniqSortedValues(records, colName, limit = 5000) {
  const s = new Set();
  for (const r of records) {
    const k = toKey(r?.[colName]);
    if (!k) continue;
    s.add(k);
    if (s.size >= limit) break;
  }
  return Array.from(s).sort((a, b) => a.localeCompare(b));
}

function getActiveFilters(exceptKeys = []) {
  const except = new Set(Array.isArray(exceptKeys) ? exceptKeys : [exceptKeys]);
  return Object.entries(state.filters).filter(([k, set]) => !except.has(k) && set && set.size > 0);
}

function filterRecordsWithActive(records, activeEntries) {
  if (!activeEntries.length) return records;
  return records.filter((r) => {
    for (const [logicalKey, selected] of activeEntries) {
      const col = state.dimCols[logicalKey];
      if (!col) continue;
      const value = toKey(r?.[col]);
      if (!selected.has(value)) return false;
    }
    return true;
  });
}

function buildingScopedRecords(records) {
  const buildingCol = state.dimCols.building;
  if (!buildingCol) return records;

  const selectedBuildings = state.filters.building;
  if (!selectedBuildings || selectedBuildings.size === 0) return records;

  return records.filter((r) => selectedBuildings.has(toKey(r?.[buildingCol])));
}

function measureTextPx(text, font) {
  const canvas = measureTextPx._canvas ?? (measureTextPx._canvas = document.createElement('canvas'));
  const ctx = canvas.getContext('2d');
  ctx.font = font;
  return ctx.measureText(text).width;
}

function autosizeSelectToWidestOption(selectEl) {
  if (!selectEl) return;
  const style = window.getComputedStyle(selectEl);
  const font = `${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;

  let max = 0;
  for (const opt of selectEl.options) {
    max = Math.max(max, measureTextPx(opt.textContent ?? '', font));
  }

  const extra = 56;
  selectEl.style.width = Math.ceil(max + extra) + 'px';
}

function enableControls(enabled) {
  if (els.clearFilters) els.clearFilters.disabled = !enabled;

  const hasBuilding = Boolean(state.dimCols.building);
  const buildingEnabled = enabled && hasBuilding;

  if (els.buildingSelect) els.buildingSelect.disabled = !buildingEnabled;
  if (els.selectAllBuildings) els.selectAllBuildings.disabled = !buildingEnabled;
  if (els.clearBuildings) els.clearBuildings.disabled = !buildingEnabled;
  if (els.buildingText) els.buildingText.disabled = !buildingEnabled;
  if (els.applyBuildingText) els.applyBuildingText.disabled = !buildingEnabled;
  if (els.clearBuildingText) els.clearBuildingText.disabled = !buildingEnabled;
}

function syncBuildingSelectFromState() {
  if (!els.buildingSelect) return;
  const selected = state.filters.building;
  for (const opt of els.buildingSelect.options) {
    opt.selected = selected.has(opt.value);
  }
}

function populateBuildingSelectOptions() {
  if (!els.buildingSelect) return;
  els.buildingSelect.innerHTML = '';

  if (!state.dimCols.building) return;

  for (const v of state.knownBuildings) {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    if (state.filters.building.has(v)) opt.selected = true;
    els.buildingSelect.appendChild(opt);
  }
}

function parseCommaList(text) {
  if (!text) return [];
  return text
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

function applyBuildingTextFilter() {
  const buildingCol = state.dimCols.building;
  if (!buildingCol) return;
  if (!els.buildingText) return;

  const raw = els.buildingText.value;
  const parts = parseCommaList(raw);

  const selected = new Set();
  const unknown = [];

  for (const p of parts) {
    const keyLower = p.toLowerCase();
    const mapped = state.knownBuildingsLowerMap.get(keyLower);
    if (mapped) {
      selected.add(mapped);
    } else {
      selected.add(p);
      unknown.push(p);
    }
  }

  state.filters.building.clear();
  for (const v of selected) state.filters.building.add(v);
  syncBuildingSelectFromState();

  if (unknown.length) logDebug(`Building override included unknown value(s): ${unknown.join(', ')}`);

  applyFilters();
  buildFiltersUI();
  render();
}

function clearBuildingTextFilter() {
  if (els.buildingText) els.buildingText.value = '';
  state.filters.building.clear();
  syncBuildingSelectFromState();
  applyFilters();
  buildFiltersUI();
  render();
}

function getRowHeaderCols() {
  const buildingCol = state.dimCols.building;
  const participantCol = state.dimCols.participant;
  const osCol = state.dimCols.os;
  const sectionCol = state.dimCols.row_type;
  const idCol = state.dimCols.id;

  const cols = [];
  if (buildingCol) cols.push({ key: buildingCol, label: 'Building' });
  if (participantCol) cols.push({ key: participantCol, label: 'Participant' });
  if (osCol) cols.push({ key: osCol, label: 'OS' });
  if (sectionCol) cols.push({ key: sectionCol, label: 'Section' });
  if (idCol) cols.push({ key: idCol, label: 'Identifier' });

  if (!cols.length && participantCol) cols.push({ key: participantCol, label: 'Participant' });
  return cols;
}

function pruneIdBySectionToSelectedSections() {
  const selectedSections = state.filters.row_type;
  if (!selectedSections || selectedSections.size === 0) {
    state.idBySection.clear();
    return;
  }
  for (const sec of Array.from(state.idBySection.keys())) {
    if (!selectedSections.has(sec)) state.idBySection.delete(sec);
  }
}

function applyFilters() {
  const buildingCol = state.dimCols.building;
  const requireBuildingSelection = Boolean(buildingCol);
  const selectedBuildings = state.filters.building;

  if (requireBuildingSelection && (!selectedBuildings || selectedBuildings.size === 0)) {
    state.filteredRecords = [];
    return;
  }

  const active = getActiveFilters([]);
  let out = active.length ? filterRecordsWithActive(state.records, active) : state.records;

  // Section-specific ID filters: only apply when Section filter is explicitly set.
  const sectionCol = state.dimCols.row_type;
  const idCol = state.dimCols.id;
  const participantCol = state.dimCols.participant;
  const selectedSections = state.filters.row_type;

  if (sectionCol && idCol && selectedSections && selectedSections.size > 0 && state.idBySection.size > 0) {
    out = out.filter((r) => {
      const sec = toKey(r?.[sectionCol]);
      // Section filter already applied above, but keep this safe.
      if (!selectedSections.has(sec)) return false;

      const set = state.idBySection.get(sec);
      if (!set || set.size === 0) return true;

      // If the set contains participant+identifier keys, filter on the pair.
      // Otherwise, fall back to identifier-only behavior.
      const sample = set.values().next().value;
      const usesPairKeys = Boolean(participantCol) && typeof sample === 'string' && sample.includes(PARTICIPANT_ID_SEP);

      if (usesPairKeys) {
        const pairKey = makeParticipantIdKey(r?.[participantCol], r?.[idCol]);
        return set.has(pairKey);
      }

      const idVal = toKey(r?.[idCol]);
      return set.has(idVal);
    });
  }

  state.filteredRecords = out;
}

function clearAllFilters() {
  // Keep Building selection intact; it's the primary scope.
  for (const k of Object.keys(state.filters)) {
    if (k === 'building') continue;
    state.filters[k].clear();
  }
  state.idBySection.clear();

  applyFilters();
  buildFiltersUI();
  render();
}

function selectionSummary(selectedCount, totalCount) {
  if (!selectedCount) return 'All';
  if (!totalCount) return `${selectedCount}`;
  return `${selectedCount}/${totalCount}`;
}

function normalizePickerValues(values) {
  const out = [];
  for (const v of values ?? []) {
    if (v && typeof v === 'object' && 'value' in v) {
      const value = String(v.value ?? '');
      const label = String(v.label ?? v.value ?? '');
      out.push({ value, label, search: (label + ' ' + value).toLowerCase() });
    } else {
      const value = String(v ?? '');
      out.push({ value, label: value, search: value.toLowerCase() });
    }
  }
  return out;
}

function openMultiSelectPicker({ title, values, selectedSet, onApply }) {
  const existing = document.querySelector('.picker-overlay');
  if (existing) existing.remove();

  const items = normalizePickerValues(values);
  const totalCount = items.length;
  const workingSelected = new Set(Array.from(selectedSet ?? []).map((v) => String(v)));

  const overlay = document.createElement('div');
  overlay.className = 'picker-overlay';
  overlay.tabIndex = -1;

  const panel = document.createElement('div');
  panel.className = 'picker';
  panel.setAttribute('role', 'dialog');
  panel.setAttribute('aria-modal', 'true');

  const head = document.createElement('div');
  head.className = 'picker-head';

  const titleEl = document.createElement('div');
  titleEl.className = 'picker-title';
  titleEl.textContent = title;

  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'btn btn-small';
  closeBtn.textContent = 'Close';

  head.appendChild(titleEl);
  head.appendChild(closeBtn);

  const body = document.createElement('div');
  body.className = 'picker-body';

  const hint = document.createElement('div');
  hint.className = 'muted';
  hint.style.fontSize = '12px';
  hint.textContent = 'Empty selection = All';

  const count = document.createElement('div');
  count.className = 'picker-count muted';
  count.style.fontSize = '12px';

  const actionRow = document.createElement('div');
  actionRow.className = 'picker-toolbar';

  const selectAllBtn = document.createElement('button');
  selectAllBtn.type = 'button';
  selectAllBtn.className = 'btn btn-small';
  selectAllBtn.textContent = 'Select all';

  const clearBtn = document.createElement('button');
  clearBtn.type = 'button';
  clearBtn.className = 'btn btn-small';
  clearBtn.textContent = 'Clear';

  actionRow.appendChild(selectAllBtn);
  actionRow.appendChild(clearBtn);

  const sel = document.createElement('select');
  sel.className = 'picker-select';
  sel.multiple = true;
  sel.size = 14;

  const footer = document.createElement('div');
  footer.className = 'picker-actions';

  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'btn';
  cancelBtn.textContent = 'Cancel';

  const applyBtn = document.createElement('button');
  applyBtn.type = 'button';
  applyBtn.className = 'btn active';
  applyBtn.textContent = 'Apply';

  footer.appendChild(cancelBtn);
  footer.appendChild(applyBtn);

  body.appendChild(hint);
  body.appendChild(count);
  body.appendChild(actionRow);
  body.appendChild(sel);

  panel.appendChild(head);
  panel.appendChild(body);
  panel.appendChild(footer);

  overlay.appendChild(panel);
  document.body.appendChild(overlay);

  const close = () => {
    overlay.remove();
  };

  const updateCount = () => {
    count.textContent = `Selected: ${selectionSummary(workingSelected.size, totalCount)}`;
  };

  const renderOptions = () => {
    sel.innerHTML = '';
    for (const it of items) {
      const opt = document.createElement('option');
      opt.value = it.value;
      opt.textContent = it.label;
      opt.selected = workingSelected.has(it.value);
      sel.appendChild(opt);
    }
  };

  const syncWorkingFromSelect = () => {
    for (const opt of sel.options) {
      if (opt.selected) workingSelected.add(opt.value);
      else workingSelected.delete(opt.value);
    }
  };

  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) close();
  });

  overlay.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  });

  closeBtn.addEventListener('click', () => close());
  cancelBtn.addEventListener('click', () => close());

  sel.addEventListener('change', () => {
    syncWorkingFromSelect();
    updateCount();
  });

  selectAllBtn.addEventListener('click', () => {
    for (const it of items) workingSelected.add(it.value);
    renderOptions();
    updateCount();
  });

  clearBtn.addEventListener('click', () => {
    workingSelected.clear();
    renderOptions();
    updateCount();
  });

  applyBtn.addEventListener('click', () => {
    const next = new Set(workingSelected);
    close();
    onApply(next);
  });

  renderOptions();
  updateCount();

  // Focus
  setTimeout(() => {
    sel.focus();
  }, 0);
}

function buildFiltersUI() {
  if (!els.filtersContainer) return;
  els.filtersContainer.innerHTML = '';

  const buildingCol = state.dimCols.building;
  const requireBuildingSelection = Boolean(buildingCol);
  const hasSelectedBuildings = state.filters.building && state.filters.building.size > 0;

  if (requireBuildingSelection && !hasSelectedBuildings) {
    if (els.filtersHint) {
      els.filtersHint.textContent = 'Select one or more Building values above to show results and enable filters.';
    }
    return;
  }

  const buildingScoped = buildingScopedRecords(state.records);

  const renderStandardFilter = ({ logicalKey, label }) => {
    const col = state.dimCols[logicalKey];
    if (!col) return false;

    let values = uniqSortedValues(buildingScoped, col);
    if (!values.length) return false;

    if (logicalKey === 'row_type') {
      values = sortByPreferredOrder(values, SECTION_ORDER);
    }

    // Prune selections no longer available under the current scope.
    if (state.filters[logicalKey] && state.filters[logicalKey].size) {
      const allowed = new Set(values);
      for (const v of Array.from(state.filters[logicalKey])) {
        if (!allowed.has(v)) state.filters[logicalKey].delete(v);
      }
    }

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'filter-btn';

    const name = document.createElement('span');
    name.className = 'filter-name';
    name.textContent = `${label}`;

    const badge = document.createElement('span');
    badge.className = 'filter-badge';
    badge.textContent = selectionSummary(state.filters[logicalKey]?.size ?? 0, values.length);

    btn.appendChild(name);
    btn.appendChild(badge);

    btn.addEventListener('click', () => {
      openMultiSelectPicker({
        title: `${label}`,
        values,
        selectedSet: state.filters[logicalKey],
        onApply: (nextSet) => {
          const set = state.filters[logicalKey];
          set.clear();
          for (const v of nextSet) set.add(v);

          if (logicalKey === 'row_type') {
            pruneIdBySectionToSelectedSections();
          }

          applyFilters();
          buildFiltersUI();
          render();
        },
      });
    });

    els.filtersContainer.appendChild(btn);
    return true;
  };

  let addedAny = false;

  // Render Participant + Section first.
  if (renderStandardFilter({ logicalKey: 'participant', label: 'Participant' })) addedAny = true;
  if (renderStandardFilter({ logicalKey: 'row_type', label: 'Section' })) addedAny = true;

  // Render section-specific ID filters.
  const sectionCol = state.dimCols.row_type;
  const idCol = state.dimCols.id;
  const participantCol = state.dimCols.participant;
  const selectedSections = state.filters.row_type;

  if (sectionCol && idCol) {
    if (!selectedSections || selectedSections.size === 0) {
      // Clear any prior section-id filters to avoid confusing invisible state.
      state.idBySection.clear();
    } else {
      pruneIdBySectionToSelectedSections();

      const otherActive = getActiveFilters(['building', 'row_type']);
      const sections = sortByPreferredOrder(Array.from(selectedSections), SECTION_ORDER);

      for (const sec of sections) {
        const sectionOnly = buildingScoped.filter((r) => toKey(r?.[sectionCol]) === sec);
        const scoped = otherActive.length ? filterRecordsWithActive(sectionOnly, otherActive) : sectionOnly;
        let values;
        let allowedKeys;

        if (participantCol) {
          const keyToLabel = new Map();
          for (const r of scoped) {
            const idVal = toKey(r?.[idCol]);
            if (!idVal) continue;
            const pVal = toKey(r?.[participantCol]);
            const key = makeParticipantIdKey(pVal, idVal);
            const label = pVal ? `${pVal} — ${idVal}` : idVal;
            if (!keyToLabel.has(key)) keyToLabel.set(key, label);
          }

          const entries = Array.from(keyToLabel.entries()).sort((a, b) => a[1].localeCompare(b[1]));
          values = entries.map(([value, label]) => ({ value, label }));
          allowedKeys = new Set(entries.map(([value]) => value));
        } else {
          const ids = uniqSortedValues(scoped, idCol);
          if (!ids.length) continue;
          values = ids;
          allowedKeys = new Set(ids);
        }

        if (!values || (Array.isArray(values) && values.length === 0)) continue;

        addedAny = true;

        const set = state.idBySection.get(sec) ?? new Set();
        state.idBySection.set(sec, set);

        // Prune selections no longer available.
        if (set.size) {
          for (const v of Array.from(set)) {
            if (!allowedKeys.has(v)) set.delete(v);
          }
        }

        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'filter-btn';

        const name = document.createElement('span');
        name.className = 'filter-name';
        name.textContent = `Identifier — ${sec}`;

        const badge = document.createElement('span');
        badge.className = 'filter-badge';
        badge.textContent = selectionSummary(set.size, values.length);

        btn.appendChild(name);
        btn.appendChild(badge);

        btn.addEventListener('click', () => {
          openMultiSelectPicker({
            title: `Identifier — ${sec}`,
            values,
            selectedSet: set,
            onApply: (nextSet) => {
              set.clear();
              for (const v of nextSet) set.add(v);
              applyFilters();
              buildFiltersUI();
              render();
            },
          });
        });

        els.filtersContainer.appendChild(btn);
      }
    }
  }

  // Render remaining filters.
  if (renderStandardFilter({ logicalKey: 'stage', label: 'Stage' })) addedAny = true;
  if (renderStandardFilter({ logicalKey: 'os', label: 'OS' })) addedAny = true;

  if (els.filtersHint) {
    if (addedAny) {
      els.filtersHint.textContent = '';
      els.filtersHint.style.display = 'none';
    } else {
      els.filtersHint.style.display = '';
      els.filtersHint.textContent = 'No filterable columns detected. Update column detection heuristics in app.js.';
    }
  }
}

function render() {
  if (!state.records.length) {
    if (els.gridContainer) els.gridContainer.innerHTML = '<div class="placeholder">Waiting for dataset…</div>';
    if (els.gridSummary) els.gridSummary.textContent = 'No data loaded.';
    state.lastPivot = null;
    state.lastRowHeaderCols = null;
    setExportEnabled(false);
    return;
  }

  applyFilters();

  if (state.dimCols.building && state.filters.building.size === 0) {
    if (els.gridContainer) els.gridContainer.innerHTML = '<div class="placeholder">Select one or more buildings to begin.</div>';
    if (els.gridSummary) els.gridSummary.textContent = 'No buildings selected.';
    state.lastPivot = null;
    state.lastRowHeaderCols = null;
    setExportEnabled(false);
    return;
  }

  if (!state.filteredRecords.length) {
    if (els.gridContainer) els.gridContainer.innerHTML = '<div class="placeholder">No rows match the selected filters.</div>';
    if (els.gridSummary) els.gridSummary.textContent = 'No matching rows.';
    state.lastPivot = null;
    state.lastRowHeaderCols = null;
    setExportEnabled(false);
    return;
  }

  const rowHeaderCols = getRowHeaderCols();
  const rowKeys = rowHeaderCols.map((c) => c.key);
  const colKey = state.dimCols.stage;

  if (!rowKeys.length) {
    setStatus('Could not detect row header columns. Update heuristics in app.js (guessDimensionColumns).', { error: true });
    return;
  }
  if (!colKey) {
    setStatus('Could not detect a Stage column. Update heuristics in app.js (guessDimensionColumns).', { error: true });
    return;
  }

  const pivot = buildPivot({
    records: state.filteredRecords,
    rowKey: rowKeys,
    colKey,
    valueKey: METRICS,
  });

  // Ensure results rows follow the same preferred Section ordering used in the filters UI.
  // (Within each other left-column grouping, e.g., Building/Participant/OS.)
  if (state.dimCols.row_type) {
    sortPivotRowsByRowKeys(pivot, {
      preferredOrderByKey: {
        [state.dimCols.row_type]: SECTION_ORDER,
      },
    });
  }

  state.lastPivot = pivot;
  state.lastRowHeaderCols = rowHeaderCols;
  setExportEnabled(true);

  if (els.gridSummary) {
    const metricText = METRICS.map((m) => METRIC_LABELS[m] ?? m).join(', ');
    els.gridSummary.textContent = `Rows: ${pivot.rows.length} • Columns: ${pivot.cols.length} • Metrics: ${metricText} • Filtered: ${state.filteredRecords.length.toLocaleString()}/${state.records.length.toLocaleString()}`;
  }

  renderPivotGrid({
    container: els.gridContainer,
    pivot,
    rowHeaderKeys: rowHeaderCols,
    metricKeys: METRICS,
    metricLabels: METRIC_LABELS,
    valueFormatter: (v) => {
      if (v === null || v === undefined || v === '') return '';
      const num = typeof v === 'number' ? v : Number(v);
      if (Number.isFinite(num) && String(v).trim() !== '') return num.toFixed(2);
      return String(v);
    },
  });
}

async function onFileSelected(file) {
  enableControls(false);
  setStatus('Reading file…');
  logDebug(`File selected: ${file?.name ?? '(unknown)'} (${file?.size ?? 0} bytes)`);

  try {
    const buf = await file.arrayBuffer();
    const bytes = new Uint8Array(buf);

    setStatus('Loading Pyodide + pandas (first load can take a bit)…');
    logDebug('Starting Pyodide unpickle…');

    const { columns, records } = await unpickleDataFrameToRecords(bytes);
    logDebug('Unpickle succeeded. Converting to pivot view…');

    state.columns = columns;
    state.dimCols = guessDimensionColumns(columns);
    state.records = records;

    // Cache known building values for case-insensitive mapping (manual input).
    state.knownBuildings = [];
    state.knownBuildingsLowerMap = new Map();
    if (state.dimCols.building) {
      const vals = uniqSortedValues(state.records, state.dimCols.building, 20000);
      state.knownBuildings = vals;
      for (const v of vals) state.knownBuildingsLowerMap.set(String(v).toLowerCase(), v);
    }

    // Reset filters on new load.
    for (const k of Object.keys(state.filters)) state.filters[k].clear();
    state.idBySection.clear();
    if (els.buildingText) els.buildingText.value = '';

    populateBuildingSelectOptions();
    syncBuildingSelectFromState();

    applyFilters();
    buildFiltersUI();

    if (els.clearFilters) {
      els.clearFilters.onclick = clearAllFilters;
    }

    if (els.columnsPreview) {
      els.columnsPreview.textContent = JSON.stringify({
        detectedDimensions: state.dimCols,
        columns: state.columns,
      }, null, 2);
    }

    enableControls(true);
    if (els.zoomSelect) els.zoomSelect.disabled = false;
    setExportEnabled(false);
    setStatus(`Loaded ${records.length.toLocaleString()} rows. Select building(s) above to begin.`);
    logDebug(`Loaded ${records.length} rows, ${columns.length} columns.`);

    render();
  } catch (err) {
    console.error(err);
    setStatus(`Load failed: ${err?.message ?? String(err)}`, { error: true });
    logDebug(`Load failed: ${err?.stack ?? (err?.message ?? String(err))}`);
    if (els.gridContainer) els.gridContainer.innerHTML = '<div class="placeholder">Failed to load dataset.</div>';
    if (els.gridSummary) els.gridSummary.textContent = 'Load failed.';
  }
}

if (els.exportExcel) {
  els.exportExcel.addEventListener('click', () => exportCurrentPivotToExcel());
}

// Grid scale control
if (els.zoomSelect) {
  // Apply saved value to the select UI if present.
  try {
    const saved = localStorage.getItem('resultsArchive.gridZoom');
    if (saved) els.zoomSelect.value = String(saved);
  } catch {
    // ignore
  }

  els.zoomSelect.addEventListener('change', () => {
    const v = els.zoomSelect.value;
    setGridZoom(v);
    try {
      localStorage.setItem('resultsArchive.gridZoom', String(v));
    } catch {
      // ignore
    }
  });
}

// Manual building override events
if (els.applyBuildingText) {
  els.applyBuildingText.addEventListener('click', () => applyBuildingTextFilter());
}
if (els.clearBuildingText) {
  els.clearBuildingText.addEventListener('click', () => clearBuildingTextFilter());
}
if (els.buildingText) {
  els.buildingText.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      applyBuildingTextFilter();
    }
  });
  els.buildingText.addEventListener('change', () => applyBuildingTextFilter());
}

// Header building selection events
if (els.buildingSelect) {
  els.buildingSelect.addEventListener('change', () => {
    state.filters.building.clear();
    for (const opt of els.buildingSelect.selectedOptions) state.filters.building.add(opt.value);
    applyFilters();
    buildFiltersUI();
    render();
  });
}
if (els.selectAllBuildings && els.buildingSelect) {
  els.selectAllBuildings.addEventListener('click', () => {
    for (const opt of els.buildingSelect.options) opt.selected = true;
    state.filters.building.clear();
    for (const opt of els.buildingSelect.options) state.filters.building.add(opt.value);
    applyFilters();
    buildFiltersUI();
    render();
  });
}
if (els.clearBuildings && els.buildingSelect) {
  els.clearBuildings.addEventListener('click', () => {
    for (const opt of els.buildingSelect.options) opt.selected = false;
    state.filters.building.clear();
    applyFilters();
    buildFiltersUI();
    render();
  });
}

// File input events
if (!els.fileInput) {
  setStatus('App initialization failed: file input not found in DOM.', { error: true });
  logDebug('Fatal: #fileInput not found.');
} else {
  els.fileInput.addEventListener('click', () => {
    els.fileInput.value = '';
  });

  els.fileInput.addEventListener('change', () => {
    const file = els.fileInput.files?.[0];
    if (!file) return;
    logDebug('fileInput change event fired.');
    onFileSelected(file);
  });
}

if (els.clearFilters) {
  els.clearFilters.addEventListener('click', () => clearAllFilters());
}
