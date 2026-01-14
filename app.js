import { unpickleDataFrameToRecords } from './pyodide-loader.js?v=30';
import { buildPivot, renderPivotGrid } from './pivot.js?v=30';

const els = {
  fileInput: document.getElementById('fileInput'),
  statusText: document.getElementById('statusText'),
  columnsPreview: document.getElementById('columnsPreview'),
  gridContainer: document.getElementById('gridContainer'),
  gridSummary: document.getElementById('gridSummary'),
  zoomSelect: document.getElementById('zoomSelect'),
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
    els.filtersHint.textContent = addedAny
      ? 'Filters: click a button to choose values (empty = All). Identifier filters appear per selected Section.'
      : 'No filterable columns detected. Update column detection heuristics in app.js.';
  }
}

function render() {
  if (!state.records.length) {
    if (els.gridContainer) els.gridContainer.innerHTML = '<div class="placeholder">Waiting for dataset…</div>';
    if (els.gridSummary) els.gridSummary.textContent = 'No data loaded.';
    return;
  }

  applyFilters();

  if (state.dimCols.building && state.filters.building.size === 0) {
    if (els.gridContainer) els.gridContainer.innerHTML = '<div class="placeholder">Select one or more buildings to begin.</div>';
    if (els.gridSummary) els.gridSummary.textContent = 'No buildings selected.';
    return;
  }

  if (!state.filteredRecords.length) {
    if (els.gridContainer) els.gridContainer.innerHTML = '<div class="placeholder">No rows match the selected filters.</div>';
    if (els.gridSummary) els.gridSummary.textContent = 'No matching rows.';
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
