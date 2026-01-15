import { unpickleDataFrameToRecords } from './pyodide-loader.js?v=32';
import { buildPivot, renderPivotGrid } from './pivot.js?v=31';

const els = {
  fileInput: document.getElementById('fileInput'),
  callFileInput: document.getElementById('callFileInput'),
  fileInputStatus: document.getElementById('fileInputStatus'),
  callFileInputStatus: document.getElementById('callFileInputStatus'),
  statusText: document.getElementById('statusText'),
  columnsPreview: document.getElementById('columnsPreview'),
  gridContainer: document.getElementById('gridContainer'),
  gridSummary: document.getElementById('gridSummary'),
  zoomSelect: document.getElementById('zoomSelect'),
  exportExcel: document.getElementById('exportExcel'),
  lastDatasetInfo: document.getElementById('lastDatasetInfo'),
  lastCallDatasetInfo: document.getElementById('lastCallDatasetInfo'),
  filtersDetails: document.getElementById('filtersDetails'),
  gridCard: document.getElementById('gridCard'),
  callCard: document.getElementById('callCard'),
  callSummary: document.getElementById('callSummary'),
  callTableContainer: document.getElementById('callTableContainer'),
  callLocationSourceBtn: document.getElementById('callLocationSourceBtn'),
  callViewToggleBtn: document.getElementById('callViewToggleBtn'),
  exportCallsExcel: document.getElementById('exportCallsExcel'),
  exportCallsKml: document.getElementById('exportCallsKml'),
  debugSection: document.getElementById('debugSection'),
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

function setFileStatus(kind, text) {
  const el = kind === 'call' ? els.callFileInputStatus : els.fileInputStatus;
  if (!el) return;
  el.textContent = text || 'No file chosen';
}

const state = {
  columns: [],
  records: [],
  filteredRecords: [],

  lastPivot: null,
  lastRowHeaderCols: null,

  metricCols: {
    h80: null,
    v80: null,
  },

  lastFileInfo: null,

  dimCols: {
    stage: null,
    building: null,
    participant: null,
    path_id: null,
    point_id: null,
    os: null,
    row_type: null,
    id: null,
  },

  // Standard filters (empty = All). Building is required if present.
  filters: {
    participant: new Set(),
    stage: new Set(),
    building: new Set(),
    path_id: new Set(),
    point_id: new Set(),
    os: new Set(),
    row_type: new Set(),
    location_source: new Set(),
  },

  // Per-section ID filters (Section value -> Set(ID values)).
  // Only active when one or more Sections are explicitly selected.
  idBySection: new Map(),

  // Manual building override input support
  knownBuildings: [],
  knownBuildingsLowerMap: new Map(),
};

const callState = {
  columns: [],
  records: [],
  filteredRecords: [],
  lastFileInfo: null,
  dimCols: {
    participant: null,
    stage: null,
    building: null,
    path_id: null,
    point_id: null,
    location_source: null,

    // KML/vector export columns
    actual_lat: null,
    actual_lon: null,
    actual_geoid_alt: null,
    actual_hae_alt: null,

    location_lat: null,
    location_lon: null,
    location_geoid_alt: null,
    location_hae_alt: null,
  },
};

const callUi = {
  showPreview: false,
};

const IDB_DB_NAME = 'resultsArchive';
const IDB_DB_VERSION = 1;
const IDB_STORE_FILES = 'files';
const IDB_KEY_ARCHIVE_PKL = 'archivePkl';
const IDB_KEY_CALL_PKL = 'callPkl';

const PYODIDE_CDN_URL = 'https://cdn.jsdelivr.net/pyodide/v0.25.1/full/pyodide.js';

function ensureExternalScript({ url, globalName, timeoutMs = 30000 }) {
  // If already present, no-op.
  if (globalName && globalName in window) return Promise.resolve();

  // If a script with this URL already exists, wait for it.
  const existing = Array.from(document.scripts).find((s) => s.src === url);
  if (existing) {
    if (globalName && globalName in window) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const t = setTimeout(() => reject(new Error(`Timed out loading ${url}`)), timeoutMs);
      existing.addEventListener('load', () => {
        clearTimeout(t);
        resolve();
      });
      existing.addEventListener('error', () => {
        clearTimeout(t);
        reject(new Error(`Failed to load ${url}`));
      });
    });
  }

  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = url;
    script.async = true;
    const t = setTimeout(() => {
      script.remove();
      reject(new Error(`Timed out loading ${url}`));
    }, timeoutMs);
    script.onload = () => {
      clearTimeout(t);
      resolve();
    };
    script.onerror = () => {
      clearTimeout(t);
      reject(new Error(`Failed to load ${url}`));
    };
    document.head.appendChild(script);
  });
}

async function ensurePyodideAvailable() {
  // loadPyodide is provided by pyodide.js
  if ('loadPyodide' in window) return;
  logDebug('Pyodide global not found; injecting pyodide.js…');
  await ensureExternalScript({ url: PYODIDE_CDN_URL, globalName: 'loadPyodide', timeoutMs: 60000 });
}

function formatBytes(n) {
  const num = Number(n);
  if (!Number.isFinite(num) || num <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  let i = 0;
  let v = num;
  while (v >= 1024 && i < units.length - 1) {
    v /= 1024;
    i++;
  }
  const digits = i === 0 ? 0 : (i === 1 ? 1 : 2);
  return `${v.toFixed(digits)} ${units[i]}`;
}

function openIdb() {
  return new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) {
      reject(new Error('IndexedDB not available'));
      return;
    }
    const req = indexedDB.open(IDB_DB_NAME, IDB_DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE_FILES)) {
        db.createObjectStore(IDB_STORE_FILES, { keyPath: 'key' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error('IndexedDB open failed'));
  });
}

async function idbGetPkl(key) {
  const db = await openIdb();
  try {
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE_FILES, 'readonly');
      const store = tx.objectStore(IDB_STORE_FILES);
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result ?? null);
      req.onerror = () => reject(req.error ?? new Error('IndexedDB read failed'));
    });
  } finally {
    db.close();
  }
}

async function idbDeletePkl(key) {
  const db = await openIdb();
  try {
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE_FILES, 'readwrite');
      const store = tx.objectStore(IDB_STORE_FILES);
      store.delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error ?? new Error('IndexedDB delete failed'));
      tx.onabort = () => reject(tx.error ?? new Error('IndexedDB delete aborted'));
    });
  } finally {
    db.close();
  }
}

async function idbSetPkl(key, { file, bytes }) {
  const db = await openIdb();
  try {
    const meta = {
      name: file?.name ?? '',
      size: file?.size ?? (bytes?.byteLength ?? 0),
      lastModified: file?.lastModified ?? 0,
      type: file?.type ?? 'application/octet-stream',
      savedAt: Date.now(),
    };

    const blob = file instanceof Blob ? file : new Blob([bytes ?? new Uint8Array()], { type: meta.type });

    await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE_FILES, 'readwrite');
      const store = tx.objectStore(IDB_STORE_FILES);
      store.put({ key, meta, blob });
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error ?? new Error('IndexedDB write failed'));
      tx.onabort = () => reject(tx.error ?? new Error('IndexedDB write aborted'));
    });

    return meta;
  } finally {
    db.close();
  }
}

async function refreshLastDatasetInfo() {
  if (els.lastDatasetInfo) {
    try {
      const rec = await idbGetPkl(IDB_KEY_ARCHIVE_PKL);
      const meta = rec?.meta;
      if (!meta?.name) {
        els.lastDatasetInfo.textContent = '';
      } else {
        const parts = [];
        parts.push(`Saved archive: ${meta.name}`);
        if (meta.size) parts.push(formatBytes(meta.size));
        if (meta.lastModified) parts.push(new Date(meta.lastModified).toLocaleString());
        els.lastDatasetInfo.textContent = parts.join(' • ');
      }
    } catch {
      els.lastDatasetInfo.textContent = '';
    }
  }

  if (els.lastCallDatasetInfo) {
    try {
      const rec = await idbGetPkl(IDB_KEY_CALL_PKL);
      const meta = rec?.meta;
      if (!meta?.name) {
        els.lastCallDatasetInfo.textContent = '';
      } else {
        const parts = [];
        parts.push(`Saved call data: ${meta.name}`);
        if (meta.size) parts.push(formatBytes(meta.size));
        if (meta.lastModified) parts.push(new Date(meta.lastModified).toLocaleString());
        els.lastCallDatasetInfo.textContent = parts.join(' • ');
      }
    } catch {
      els.lastCallDatasetInfo.textContent = '';
    }
  }
}

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

function sortRowIdsByRowKeys(rowIds, pivot, { preferredOrderByKey = {} } = {}) {
  if (!Array.isArray(rowIds) || !pivot || !pivot.rowMeta || !Array.isArray(pivot.rowKeys)) return rowIds;
  if (rowIds.length <= 1) return rowIds;

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

  return rowIds.slice().sort((aRowId, bRowId) => {
    const aMeta = pivot.rowMeta.get(aRowId) ?? {};
    const bMeta = pivot.rowMeta.get(bRowId) ?? {};

    for (const key of pivot.rowKeys) {
      const cmp = compareValues(key, aMeta[key], bMeta[key]);
      if (cmp) return cmp;
    }
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
setStatus('Ready. Choose .pkl file(s) to begin.');
logDebug('app.js initialized.');

// No previous-session restore: clear any legacy saved blobs and hide the info rows.
(async () => {
  try {
    if (els.lastDatasetInfo) els.lastDatasetInfo.textContent = '';
    if (els.lastCallDatasetInfo) els.lastCallDatasetInfo.textContent = '';
    await idbDeletePkl(IDB_KEY_ARCHIVE_PKL);
    await idbDeletePkl(IDB_KEY_CALL_PKL);
    logDebug('Cleared any previously saved PKLs (previous-session restore disabled).');
  } catch {
    // ignore
  }
})();

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

function guessMetricColumns(columns) {
  const h = detectColumn(columns, [
    'h80',
    'hor 80%',
    'horizontal 80%',
    'horizontal_80%',
    'horizontal80',
    'hor80',
    'horizontal 80',
  ]);
  const v = detectColumn(columns, [
    'v80',
    'ver 80%',
    'vertical 80%',
    'vertical_80%',
    'vertical80',
    'ver80',
    'vertical 80',
  ]);
  return { h80: h, v80: v };
}

function getMetricConfig() {
  const keys = [];
  const labels = {};

  const h = state.metricCols?.h80;
  const v = state.metricCols?.v80;

  if (h) {
    keys.push(h);
    labels[h] = METRIC_LABELS.h80;
  }
  if (v) {
    keys.push(v);
    labels[v] = METRIC_LABELS.v80;
  }

  return { metricKeys: keys, metricLabels: labels };
}

function setExportEnabled(enabled) {
  if (!els.exportExcel) return;
  els.exportExcel.disabled = !enabled;
}

function setCallsExportEnabled(enabled) {
  if (!els.exportCallsExcel) return;
  els.exportCallsExcel.disabled = !enabled;
}

function setCallsKmlExportEnabled(enabled) {
  if (!els.exportCallsKml) return;
  els.exportCallsKml.disabled = !enabled;
}

function canExportCallsKml() {
  if (!callState.records.length) return false;
  if (!callState.filteredRecords.length) return false;
  const c = callState.dimCols;
  const hasActualLatLon = Boolean(c.actual_lat && c.actual_lon);
  const hasActualAlt = Boolean(c.actual_geoid_alt || c.actual_hae_alt);
  const hasLocationAlt = Boolean(c.location_geoid_alt || c.location_hae_alt);
  return hasActualLatLon && hasActualAlt && hasLocationAlt;
}

function updateCallViewToggleButton() {
  const btn = els.callViewToggleBtn;
  if (!btn) return;

  const canShow = Boolean(callState.records.length);
  btn.disabled = !canShow;

  if (!canShow) {
    btn.textContent = 'View Call Data';
    return;
  }

  btn.textContent = callUi.showPreview ? 'Hide Call Data' : 'View Call Data';
}

function updateCallLocationSourceButton() {
  const btn = els.callLocationSourceBtn;
  if (!btn) return;

  const canShow = Boolean(callState.records.length) && Boolean(callState.dimCols?.location_source);
  btn.disabled = !canShow;

  if (!canShow) {
    btn.textContent = 'Location Source: N/A';
    return;
  }

  const selected = state.filters.location_source;
  btn.textContent = `Location Source: ${selectionSummary(selected?.size ?? 0, 0)}`;
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
  lines.push(['Path ID', setToText(state.filters.path_id)]);
  lines.push(['Point ID', setToText(state.filters.point_id)]);
  lines.push(['Section', setToText(state.filters.row_type)]);
  lines.push(['Location Source', setToText(state.filters.location_source)]);

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

function buildCallFiltersSummaryAoA() {
  const lines = [];
  const setToText = (set) => {
    if (!set) return '';
    if (set.size === 0) return 'All';
    return Array.from(set).join(', ');
  };

  lines.push(['Call Data Export']);
  lines.push(['Exported at', new Date().toLocaleString()]);
  lines.push(['Build', document.getElementById('buildStamp')?.textContent ?? '']);
  lines.push([]);

  lines.push(['Buildings', setToText(state.filters.building)]);
  lines.push(['Participant', setToText(state.filters.participant)]);
  lines.push(['Stage', setToText(state.filters.stage)]);
  lines.push(['Path ID', setToText(state.filters.path_id)]);
  lines.push(['Point ID', setToText(state.filters.point_id)]);
  lines.push(['Location Source', setToText(state.filters.location_source)]);

  return lines;
}

function exportCallsToExcel() {
  const XLSX = window.XLSX;
  if (!XLSX) {
    setStatus('Excel export library not loaded yet. Please refresh and try again.', { error: true });
    return;
  }

  if (!callState.records.length) {
    setStatus('No call data loaded yet.', { error: true });
    return;
  }

  const rows = callState.filteredRecords?.length ? callState.filteredRecords : [];
  if (!rows.length) {
    setStatus('No call rows match the selected filters.', { error: true });
    return;
  }

  const cols = (callState.columns ?? []).slice();
  if (!cols.length) {
    setStatus('Call data has no detected columns to export.', { error: true });
    return;
  }

  const aoa = [cols];
  for (const r of rows) {
    aoa.push(cols.map((c) => {
      const v = r?.[c];
      if (v === null || v === undefined) return '';
      if (typeof v === 'number') return v;
      return String(v);
    }));
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Freeze header row.
  const ySplit = 1;
  const topLeftCell = XLSX.utils.encode_cell({ r: ySplit, c: 0 });
  ws['!sheetViews'] = [{ pane: { state: 'frozen', xSplit: 0, ySplit, topLeftCell, activePane: 'bottomLeft' } }];

  // Autofilter the full range.
  const lastRow = aoa.length - 1;
  const lastCol = cols.length - 1;
  ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow, c: lastCol } }) };

  // Light column sizing based on first N rows.
  const maxChars = new Array(cols.length).fill(6);
  const sampleRows = Math.min(aoa.length, 300);
  for (let r = 0; r < sampleRows; r++) {
    for (let c = 0; c < cols.length; c++) {
      const v = aoa[r]?.[c];
      const s = v === null || v === undefined ? '' : String(v);
      maxChars[c] = Math.min(70, Math.max(maxChars[c], s.length));
    }
  }
  ws['!cols'] = maxChars.map((wch) => ({ wch: Math.min(72, Math.max(8, wch)) }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Call Data');

  const ws2 = XLSX.utils.aoa_to_sheet(buildCallFiltersSummaryAoA());
  ws2['!cols'] = [{ wch: 24 }, { wch: 120 }];
  XLSX.utils.book_append_sheet(wb, ws2, 'Filters');

  const dt = new Date();
  const buildingPart = state.filters.building && state.filters.building.size
    ? `_${safeFilePart(Array.from(state.filters.building).slice(0, 3).join('-'))}${state.filters.building.size > 3 ? '_and_more' : ''}`
    : '';
  const filename = `Call_Data_${formatDateForFilename(dt)}${buildingPart}.xlsx`;

  try {
    XLSX.writeFile(wb, filename, { compression: true });
    setStatus(`Exported Excel: ${filename}`);
  } catch (err) {
    console.error(err);
    setStatus(`Excel export failed: ${err?.message ?? String(err)}`, { error: true });
  }
}

function parseNumber(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  const s = String(v).trim();
  if (!s) return null;
  const n = Number(s.replace(/,/g, ''));
  return Number.isFinite(n) ? n : null;
}

function cdataSafe(s) {
  return String(s ?? '').replaceAll(']]>', ']]&gt;');
}

function xmlEscape(s) {
  return String(s ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function downloadTextFile({ filename, text, mime }) {
  const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function buildCallKmlFromRows({ rows, docName }) {
  const c = callState.dimCols;

  // Big KMLs can be slow in Google Earth.
  const hardLimit = 50000;
  if (rows.length > hardLimit) {
    const ok = confirm(`You are exporting ${rows.length.toLocaleString()} call vectors. This may be very large/slow in Google Earth. Continue?`);
    if (!ok) return null;
  }

  const STYLE_UP = 'lineUp';
  const STYLE_DOWN = 'lineDown';

  const pieces = [];
  pieces.push('<?xml version="1.0" encoding="UTF-8"?>');
  pieces.push('<kml xmlns="http://www.opengis.net/kml/2.2">');
  pieces.push('<Document>');
  pieces.push(`<name>${xmlEscape(docName || `Call Vectors (${rows.length.toLocaleString()})`)}</name>`);

  // ABGR colors: ff0000ff = red, ffff0000 = blue
  pieces.push('<Style id="lineUp"><LineStyle><color>ff0000ff</color><width>2</width></LineStyle></Style>');
  pieces.push('<Style id="lineDown"><LineStyle><color>ffff0000</color><width>2</width></LineStyle></Style>');

  const buildingCol = c.building;
  const folderMap = new Map();
  const getFolderKey = (r) => (buildingCol ? toKey(r?.[buildingCol]) : '');
  const getFolder = (key) => {
    if (!folderMap.has(key)) folderMap.set(key, []);
    return folderMap.get(key);
  };

  for (const r of rows) {
    const actualLat = parseNumber(r?.[c.actual_lat]);
    const actualLon = parseNumber(r?.[c.actual_lon]);
    if (actualLat === null || actualLon === null) continue;

    const locLat = c.location_lat ? (parseNumber(r?.[c.location_lat]) ?? actualLat) : actualLat;
    const locLon = c.location_lon ? (parseNumber(r?.[c.location_lon]) ?? actualLon) : actualLon;

    const actualGeoid = c.actual_geoid_alt ? parseNumber(r?.[c.actual_geoid_alt]) : null;
    const actualHae = c.actual_hae_alt ? parseNumber(r?.[c.actual_hae_alt]) : null;
    const geoidSep = (actualGeoid !== null && actualHae !== null) ? (actualHae - actualGeoid) : null;

    // Google Earth's "absolute" altitude behaves like meters above sea level (geoid-ish).
    // Use surveyed Geoid Alt if present; otherwise convert from HAE using the surveyed separation.
    const actualAlt = actualGeoid !== null ? actualGeoid : (actualHae !== null && geoidSep !== null ? (actualHae - geoidSep) : actualHae);

    const locGeoid = c.location_geoid_alt ? parseNumber(r?.[c.location_geoid_alt]) : null;
    const locHae = c.location_hae_alt ? parseNumber(r?.[c.location_hae_alt]) : null;
    const locAlt = locGeoid !== null
      ? locGeoid
      : (locHae !== null && geoidSep !== null ? (locHae - geoidSep) : locHae);

    if (actualAlt === null || locAlt === null) continue;

    const delta = locAlt - actualAlt;
    const styleUrl = delta >= 0 ? `#${STYLE_UP}` : `#${STYLE_DOWN}`;

    const buildingVal = buildingCol ? toKey(r?.[buildingCol]) : '';
    const stageVal = c.stage ? toKey(r?.[c.stage]) : '';
    const participantVal = c.participant ? toKey(r?.[c.participant]) : '';
    const pathVal = c.path_id ? toKey(r?.[c.path_id]) : '';
    const pointVal = c.point_id ? toKey(r?.[c.point_id]) : '';

    const nameBits = [];
    if (buildingVal) nameBits.push(buildingVal);
    if (stageVal) nameBits.push(stageVal);
    if (pointVal) nameBits.push(pointVal);
    const placemarkName = nameBits.length ? nameBits.join(' • ') : 'Call Vector';

    const descLines = [];
    if (buildingVal) descLines.push(`Building: ${buildingVal}`);
    if (stageVal) descLines.push(`Stage: ${stageVal}`);
    if (participantVal) descLines.push(`Participant: ${participantVal}`);
    if (pathVal) descLines.push(`Path ID: ${pathVal}`);
    if (pointVal) descLines.push(`Point ID: ${pointVal}`);
    if (c.location_source) descLines.push(`Location Source: ${toKey(r?.[c.location_source])}`);
    descLines.push(`Actual Alt (MSL): ${actualAlt}`);
    descLines.push(`Location Alt (MSL): ${locAlt}`);
    descLines.push(`Delta (Location-Actual): ${delta}`);
    if (geoidSep !== null) descLines.push(`Geoid Sep (HAE-Geoid): ${geoidSep}`);

    const coords = `${actualLon},${actualLat},${actualAlt} ${locLon},${locLat},${locAlt}`;

    const pm = [];
    pm.push('<Placemark>');
    pm.push(`<name>${xmlEscape(placemarkName)}</name>`);
    pm.push(`<styleUrl>${styleUrl}</styleUrl>`);
    pm.push(`<description><![CDATA[${cdataSafe(descLines.join('<br/>'))}]]></description>`);
    pm.push('<LineString>');
    pm.push('<tessellate>0</tessellate>');
    pm.push('<altitudeMode>absolute</altitudeMode>');
    pm.push(`<coordinates>${coords}</coordinates>`);
    pm.push('</LineString>');
    pm.push('</Placemark>');

    getFolder(getFolderKey(r)).push(pm.join(''));
  }

  for (const [folderKey, placemarks] of folderMap.entries()) {
    pieces.push('<Folder>');
    if (folderKey) pieces.push(`<name>${xmlEscape(folderKey)}</name>`);
    pieces.push(...placemarks);
    pieces.push('</Folder>');
  }

  pieces.push('</Document>');
  pieces.push('</kml>');

  return pieces.join('\n');
}

async function exportCallsToKml() {
  if (!canExportCallsKml()) {
    setStatus('KML export unavailable: missing required columns (Actual Lat/Lon + altitudes).', { error: true });
    return;
  }

  // If we don't have surveyed geoid separation, exporting absolute altitudes from HAE may be offset.
  const c = callState.dimCols;
  const hasGeoidSeparation = Boolean(c.actual_geoid_alt && c.actual_hae_alt);
  const willUseHaeForActual = Boolean(c.actual_hae_alt && !c.actual_geoid_alt);
  const willUseHaeForLocation = Boolean(c.location_hae_alt && !c.location_geoid_alt);
  if (!hasGeoidSeparation && (willUseHaeForActual || willUseHaeForLocation)) {
    const ok = confirm('This call dataset does not appear to include both Actual Geoid Alt and Actual HAE Alt.\n\nThe KML will still export, but absolute altitudes may be offset because Google Earth expects MSL/Geoid-like heights.\n\nContinue?');
    if (!ok) return;
  }

  const rows = callState.filteredRecords;
  const participantCol = callState.dimCols.participant;

  // If we can, split into one KML per participant.
  if (participantCol) {
    const byParticipant = new Map();
    for (const r of rows) {
      const p = toKey(r?.[participantCol]) || '(blank)';
      if (!byParticipant.has(p)) byParticipant.set(p, []);
      byParticipant.get(p).push(r);
    }

    const participants = Array.from(byParticipant.keys()).sort((a, b) => a.localeCompare(b));
    if (participants.length > 30) {
      const ok = confirm(`This will download ${participants.length.toLocaleString()} separate KML files (one per Participant). Continue?`);
      if (!ok) return;
    }

    const dt = new Date();
    const buildingPart = state.filters.building && state.filters.building.size
      ? `_${safeFilePart(Array.from(state.filters.building).slice(0, 3).join('-'))}${state.filters.building.size > 3 ? '_and_more' : ''}`
      : '';

    let exported = 0;
    for (const p of participants) {
      const subset = byParticipant.get(p) || [];
      const docName = `Call Vectors - ${p} (${subset.length.toLocaleString()})`;
      const kml = buildCallKmlFromRows({ rows: subset, docName });
      if (!kml) continue;

      const filename = `Call_Vectors_${safeFilePart(p)}_${formatDateForFilename(dt)}${buildingPart}.kml`;
      downloadTextFile({ filename, text: kml, mime: 'application/vnd.google-earth.kml+xml;charset=utf-8' });
      exported++;

      // Avoid overwhelming the browser download UI.
      // eslint-disable-next-line no-await-in-loop
      await new Promise((resolve) => setTimeout(resolve, 120));
    }

    setStatus(`Exported ${exported.toLocaleString()} KML file(s) by Participant.`);
    return;
  }

  // Fallback: single KML
  const kml = buildCallKmlFromRows({ rows, docName: `Call Vectors (${rows.length.toLocaleString()})` });
  if (!kml) return;
  const dt = new Date();
  const buildingPart = state.filters.building && state.filters.building.size
    ? `_${safeFilePart(Array.from(state.filters.building).slice(0, 3).join('-'))}${state.filters.building.size > 3 ? '_and_more' : ''}`
    : '';
  const filename = `Call_Vectors_${formatDateForFilename(dt)}${buildingPart}.kml`;
  downloadTextFile({ filename, text: kml, mime: 'application/vnd.google-earth.kml+xml;charset=utf-8' });
  setStatus(`Exported KML: ${filename}`);
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
  const { metricKeys, metricLabels } = getMetricConfig();
  const metricCount = metricKeys.length;

  if (!metricCount) {
    setStatus('Could not find metric columns (Horizontal/Vertical 80%).', { error: true });
    return;
  }

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
      headerSub.push(String(metricLabels?.[m] ?? m));
    }
  }

  const aoa = [titleRow, generatedRow, spacerRow, headerTop, headerSub];

  const exportRowIds = state.dimCols.row_type
    ? sortRowIdsByRowKeys(pivot.rows, pivot, {
      preferredOrderByKey: {
        [state.dimCols.row_type]: SECTION_ORDER,
      },
    })
    : pivot.rows;

  for (const rowId of exportRowIds) {
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
    path_id: detectColumn(columns, ['path_id', 'path id', 'path', 'pathid']),
    point_id: detectColumn(columns, ['point_id', 'point id', 'point', 'pointid']),
    os: detectColumn(columns, ['os', 'operating system', 'platform']),
    row_type: detectColumn(columns, ['row_type', 'row type', 'type', 'section']),
    id: detectColumn(columns, ['id', 'label', 'name']),
  };
}

function guessCallDimensionColumns(columns) {
  return {
    participant: detectColumn(columns, ['participant', 'carrier', 'name', 'user', 'person']),
    stage: detectColumn(columns, ['stage', 'stg', 'phase']),
    building: detectColumn(columns, ['building_id', 'building', 'bldg', 'site', 'location']),
    path_id: detectColumn(columns, ['path_id', 'path id', 'path', 'pathid']),
    point_id: detectColumn(columns, ['point_id', 'point id', 'point', 'pointid']),
    location_source: detectColumn(columns, ['location_source', 'location source', 'loc source', 'source']),

    actual_lat: detectColumn(columns, ['actual lat', 'actual_lat', 'actual latitude', 'survey lat']),
    actual_lon: detectColumn(columns, ['actual lon', 'actual_lon', 'actual longitude', 'survey lon']),
    actual_geoid_alt: detectColumn(columns, ['actual geoid alt', 'actual geoid altitude', 'geoid alt', 'geoid_alt', 'geoid altitude']),
    actual_hae_alt: detectColumn(columns, ['actual hae alt', 'actual hae altitude', 'hae alt', 'hae_alt', 'hae altitude']),

    location_lat: detectColumn(columns, ['location lat', 'location latitude', 'loc lat', 'reported lat', 'estimated lat']),
    location_lon: detectColumn(columns, ['location lon', 'location longitude', 'loc lon', 'reported lon', 'estimated lon']),
    location_geoid_alt: detectColumn(columns, ['location geoid alt', 'location geoid altitude', 'loc geoid alt']),
    // Your measured altitude column is typically just "Location Altitude" (assumed HAE).
    location_hae_alt: detectColumn(columns, [
      'location altitude (hae)',
      'location alt (hae)',
      'location hae alt',
      'location altitude hae',
      'location altitude',
      'location_altitude',
      'loc altitude',
      'loc alt',
      'loc hae alt',
      'loc alt hae',
    ]),
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

function filterRecordsWithActiveByDim(records, activeEntries, dimCols) {
  if (!activeEntries.length) return records;
  return records.filter((r) => {
    for (const [logicalKey, selected] of activeEntries) {
      const col = dimCols?.[logicalKey];
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

function buildingScopedRecordsByDim(records, dimCols) {
  const buildingCol = dimCols?.building;
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

function updateSectionsVisibility() {
  const hasData = state.records.length > 0;
  const needsBuilding = Boolean(state.dimCols.building);
  const hasSelectedBuilding = state.filters.building && state.filters.building.size > 0;
  const show = hasData && (!needsBuilding || hasSelectedBuilding);

  const callHasData = callState.records.length > 0;
  const callNeedsBuilding = Boolean(callState.dimCols.building);
  const showCalls = callHasData && (!callNeedsBuilding || hasSelectedBuilding);

  const toggle = (el, on) => {
    if (!el) return;
    el.classList.toggle('hidden', !on);
  };

  toggle(els.filtersDetails, show);
  toggle(els.gridCard, show);
  toggle(els.callCard, showCalls);
  toggle(els.debugSection, show);
}

// Initial visibility (no data yet).
updateSectionsVisibility();

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
    callState.filteredRecords = [];
    renderCallSummary();
    renderCallTable();
    setCallsExportEnabled(false);
    setCallsKmlExportEnabled(false);
    updateCallLocationSourceButton();
    updateCallViewToggleButton();
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

  // Apply the same filter sets to call data (only for columns present in callState.dimCols).
  if (callState.records.length) {
    const scopedCalls = buildingScopedRecordsByDim(callState.records, callState.dimCols);
    const callOut = active.length ? filterRecordsWithActiveByDim(scopedCalls, active, callState.dimCols) : scopedCalls;
    callState.filteredRecords = callOut;
  } else {
    callState.filteredRecords = [];
  }

  renderCallSummary();
  renderCallTable();
  setCallsExportEnabled(callState.filteredRecords.length > 0);
  updateCallLocationSourceButton();
  setCallsKmlExportEnabled(canExportCallsKml());
  updateCallViewToggleButton();
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

  const hasAnyData = state.records.length > 0 || callState.records.length > 0;
  if (!hasAnyData) {
    if (els.filtersHint) {
      els.filtersHint.style.display = '';
      els.filtersHint.textContent = 'Load a file to enable filters.';
    }
    return;
  }

  const buildingCol = state.dimCols.building;
  const requireBuildingSelection = Boolean(buildingCol);
  const hasSelectedBuildings = state.filters.building && state.filters.building.size > 0;

  if (requireBuildingSelection && !hasSelectedBuildings) {
    if (els.filtersHint) {
      els.filtersHint.textContent = 'Select one or more Building values above to show results and enable filters.';
    }
    return;
  }

  const buildingScopedArchive = buildingScopedRecords(state.records);
  const buildingScopedCalls = buildingScopedRecordsByDim(callState.records, callState.dimCols);

  const unionValues = (a, b) => {
    const s = new Set();
    for (const v of a ?? []) s.add(v);
    for (const v of b ?? []) s.add(v);
    return Array.from(s).sort((x, y) => String(x).localeCompare(String(y)));
  };

  const renderStandardFilter = ({ logicalKey, label }) => {
    const archiveCol = state.dimCols[logicalKey];
    const callCol = callState.dimCols?.[logicalKey];
    if (!archiveCol && !callCol) return false;

    const archiveValues = archiveCol ? uniqSortedValues(buildingScopedArchive, archiveCol) : [];
    const callValues = callCol ? uniqSortedValues(buildingScopedCalls, callCol) : [];
    let values = unionValues(archiveValues, callValues);
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
        const sectionOnly = buildingScopedArchive.filter((r) => toKey(r?.[sectionCol]) === sec);
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
  if (renderStandardFilter({ logicalKey: 'path_id', label: 'Path ID' })) addedAny = true;
  if (renderStandardFilter({ logicalKey: 'point_id', label: 'Point ID' })) addedAny = true;
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
    setCallsExportEnabled(callState.filteredRecords.length > 0);
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

  const { metricKeys, metricLabels } = getMetricConfig();
  if (!metricKeys.length) {
    setStatus('Could not detect metric columns (Horizontal/Vertical 80%).', { error: true });
    return;
  }

  const pivot = buildPivot({
    records: state.filteredRecords,
    rowKey: rowKeys,
    colKey,
    valueKey: metricKeys,
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
    const metricText = metricKeys.map((m) => metricLabels?.[m] ?? m).join(', ');
    els.gridSummary.textContent = `Rows: ${pivot.rows.length} • Columns: ${pivot.cols.length} • Metrics: ${metricText} • Filtered: ${state.filteredRecords.length.toLocaleString()}/${state.records.length.toLocaleString()}`;
  }

  renderPivotGrid({
    container: els.gridContainer,
    pivot,
    rowHeaderKeys: rowHeaderCols,
    metricKeys,
    metricLabels,
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

  if (file?.name) setFileStatus('archive', file.name);

  try {
    const buf = await file.arrayBuffer();
    const bytes = new Uint8Array(buf);

    setStatus('Loading Pyodide + pandas (first load can take a bit)…');
    await ensurePyodideAvailable();
    logDebug('Starting Pyodide unpickle…');

    const { columns, records } = await unpickleDataFrameToRecords(bytes);
    logDebug('Unpickle succeeded. Converting to pivot view…');

    state.columns = columns;
    state.dimCols = guessDimensionColumns(columns);
    state.metricCols = guessMetricColumns(columns);
    state.records = records;

    state.lastFileInfo = {
      name: file?.name ?? '',
      size: file?.size ?? 0,
      lastModified: file?.lastModified ?? 0,
      type: file?.type ?? '',
    };

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

    updateSectionsVisibility();

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

    updateSectionsVisibility();
  } catch (err) {
    console.error(err);
    setStatus(`Load failed: ${err?.message ?? String(err)}`, { error: true });
    logDebug(`Load failed: ${err?.stack ?? (err?.message ?? String(err))}`);
    if (els.gridContainer) els.gridContainer.innerHTML = '<div class="placeholder">Failed to load dataset.</div>';
    if (els.gridSummary) els.gridSummary.textContent = 'Load failed.';
  }
}

function renderCallTable({ maxRows = 200 } = {}) {
  if (!els.callTableContainer) return;

  const rows = callState.filteredRecords?.length ? callState.filteredRecords : callState.records;
  if (!rows || rows.length === 0) {
    els.callTableContainer.innerHTML = '<div class="placeholder">No call rows to display.</div>';
    return;
  }

  if (!callUi.showPreview) {
    els.callTableContainer.innerHTML = `
      <div class="placeholder">
        Call preview is hidden. Click “View Call Data” to show a preview (first ${maxRows.toLocaleString()} rows).
      </div>
    `;
    return;
  }

  const cols = callState.columns ?? [];
  const slice = rows.slice(0, Math.max(0, maxRows));

  const escapeHtml = (s) => String(s)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');

  const ths = cols.map((c) => `<th>${escapeHtml(c)}</th>`).join('');
  const trs = slice.map((r) => {
    const tds = cols.map((c) => `<td>${escapeHtml(r?.[c] ?? '')}</td>`).join('');
    return `<tr>${tds}</tr>`;
  }).join('');

  els.callTableContainer.innerHTML = `
    <div class="grid-scroll">
      <table class="simple-table">
        <thead><tr>${ths}</tr></thead>
        <tbody>${trs}</tbody>
      </table>
    </div>
    <div class="muted" style="margin-top:6px;">
      Showing ${slice.length.toLocaleString()} of ${rows.length.toLocaleString()} call rows.
    </div>
  `;
}

function renderCallSummary() {
  if (!els.callSummary) return;
  if (!callState.records.length) {
    els.callSummary.textContent = 'No call data loaded.';
    return;
  }

  const parts = [];
  if (callState.lastFileInfo?.name) parts.push(callState.lastFileInfo.name);
  parts.push(`Rows: ${callState.filteredRecords.length.toLocaleString()}/${callState.records.length.toLocaleString()}`);

  els.callSummary.textContent = parts.join(' • ');
}

async function loadCallDatasetFromBytes({ bytes, fileInfo, saveToIdb = false, idbFile = null }) {
  enableControls(false);
  setStatus('Loading call data…');

  try {
    setStatus('Loading Pyodide + pandas (first load can take a bit)…');
    await ensurePyodideAvailable();
    logDebug('Starting Pyodide unpickle (call data)…');

    const { columns, records } = await unpickleDataFrameToRecords(bytes);
    logDebug('Unpickle succeeded for call data.');

    callState.columns = columns;
    callState.dimCols = guessCallDimensionColumns(columns);
    callState.records = records;
    callState.filteredRecords = records;
    callState.lastFileInfo = fileInfo;

    renderCallSummary();
    renderCallTable();
    setCallsExportEnabled(callState.filteredRecords.length > 0);
    setCallsKmlExportEnabled(canExportCallsKml());
    updateCallLocationSourceButton();
    updateCallViewToggleButton();
    updateSectionsVisibility();

    if (saveToIdb) {
      try {
        await idbSetPkl(IDB_KEY_CALL_PKL, { file: idbFile, bytes });
        logDebug('Saved call dataset to IndexedDB for next launch.');
      } catch (e) {
        logDebug(`Could not save call dataset for next launch: ${e?.message ?? String(e)}`);
      }
    }

    setStatus(`Loaded call data: ${records.length.toLocaleString()} rows.`);
  } catch (err) {
    console.error(err);
    logDebug(`Call data load failed: ${err?.stack ?? (err?.message ?? String(err))}`);
    setStatus(`Call data load failed: ${err?.message ?? String(err)}`, { error: true });
    if (els.callTableContainer) els.callTableContainer.innerHTML = '<div class="placeholder">Failed to load call data.</div>';
    if (els.callSummary) els.callSummary.textContent = 'Call data load failed.';
  } finally {
    enableControls(true);
    if (els.zoomSelect) els.zoomSelect.disabled = false;
    setCallsExportEnabled(callState.filteredRecords.length > 0);
    setCallsKmlExportEnabled(canExportCallsKml());
    updateCallLocationSourceButton();
    updateCallViewToggleButton();
    updateSectionsVisibility();
  }
}

async function onCallFileSelected(file) {
  if (!file) return;
  setStatus('Reading call data file…');
  logDebug(`Call file selected: ${file?.name ?? '(unknown)'} (${file?.size ?? 0} bytes)`);

  if (file?.name) setFileStatus('call', file.name);

  const buf = await file.arrayBuffer();
  const bytes = new Uint8Array(buf);
  const fileInfo = {
    name: file?.name ?? '',
    size: file?.size ?? 0,
    lastModified: file?.lastModified ?? 0,
    type: file?.type ?? '',
  };

  await loadCallDatasetFromBytes({ bytes, fileInfo, saveToIdb: false, idbFile: null });
}

if (els.exportExcel) {
  els.exportExcel.addEventListener('click', () => exportCurrentPivotToExcel());
}

if (els.exportCallsExcel) {
  els.exportCallsExcel.addEventListener('click', () => exportCallsToExcel());
}

if (els.exportCallsKml) {
  els.exportCallsKml.addEventListener('click', () => exportCallsToKml());
}

if (els.callViewToggleBtn) {
  // Default to hidden preview; allow user to toggle and remember preference.
  try {
    const saved = localStorage.getItem('resultsArchive.callPreview');
    if (saved === '1') callUi.showPreview = true;
  } catch {
    // ignore
  }

  updateCallViewToggleButton();

  els.callViewToggleBtn.addEventListener('click', () => {
    if (!callState.records.length) return;
    callUi.showPreview = !callUi.showPreview;
    try {
      localStorage.setItem('resultsArchive.callPreview', callUi.showPreview ? '1' : '0');
    } catch {
      // ignore
    }
    updateCallViewToggleButton();
    renderCallTable();
  });
}

if (els.callLocationSourceBtn) {
  els.callLocationSourceBtn.addEventListener('click', () => {
    if (!callState.records.length) return;
    const col = callState.dimCols?.location_source;
    if (!col) return;

    // Scope options by selected Buildings + other active filters (excluding itself).
    const scoped = buildingScopedRecordsByDim(callState.records, callState.dimCols);
    const otherActive = getActiveFilters(['location_source']);
    const preFiltered = otherActive.length ? filterRecordsWithActiveByDim(scoped, otherActive, callState.dimCols) : scoped;
    const values = uniqSortedValues(preFiltered, col);

    openMultiSelectPicker({
      title: 'Location Source',
      values,
      selectedSet: state.filters.location_source,
      onApply: (nextSet) => {
        const set = state.filters.location_source;
        set.clear();
        for (const v of nextSet) set.add(v);
        applyFilters();
        buildFiltersUI();
        render();
      },
    });
  });
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
    updateSectionsVisibility();
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
    updateSectionsVisibility();
  });
}
if (els.clearBuildings && els.buildingSelect) {
  els.clearBuildings.addEventListener('click', () => {
    for (const opt of els.buildingSelect.options) opt.selected = false;
    state.filters.building.clear();
    applyFilters();
    buildFiltersUI();
    render();
    updateSectionsVisibility();
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

// Call-data file input events
if (els.callFileInput) {
  els.callFileInput.addEventListener('click', () => {
    els.callFileInput.value = '';
  });

  els.callFileInput.addEventListener('change', () => {
    const file = els.callFileInput.files?.[0];
    if (!file) return;
    logDebug('callFileInput change event fired.');
    onCallFileSelected(file);
  });
}

if (els.clearFilters) {
  els.clearFilters.addEventListener('click', () => clearAllFilters());
}
