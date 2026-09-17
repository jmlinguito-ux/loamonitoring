/**
 * LOA Monitoring Tool — Frontend Controller
 * Ground Data Processing Tool inspired design, Excel-like multi-cell selection & copy/paste,
 * progressive chunked rendering for 60fps performance, instant column resizing,
 * sticky headers & frozen columns, live formula calculations, and Dataverse synchronization.
 */

(function () {
  'use strict';

  const CHUNK_SIZE = 100;

  // Application State
  const state = {
    connected: false,
    activeView: 'records',
    records: [],
    filteredRecords: [],
    renderedRecordsCount: 0,
    financeLoadedCount: 0,
    financeSource: 'None',
    recentFinanceFiles: [],
    theme: 'light',
    solutions: [],
    tables: [],
    activeSolutionId: '',
    activeLoaTable: 'cr63f_loamonitoring',
    activeFinanceTable: 'cr63f_financereference',
    tableSchema: null,
    stagedEdits: {},
    columnFilters: {},
    columnFiltersOpen: {},
    colWidths: {},
    colOrder: [],
    visibleCols: new Set(),
    frozenCols: new Set(),
    sortCol: null,
    sortDir: null, // 'asc' | 'desc' | null
  };

  const financeState = {
    headers: [],
    rows: [],
    filteredRows: [],
    renderedRowsCount: 0,
    colWidths: {},
    colOrder: [],
    visibleCols: new Set(),
    frozenCols: new Set(),
    sortCol: null,
    sortDir: null,
    columnFilters: {},
    columnFiltersOpen: {}
  };

  let debounceTimer = null;
  let searchDebounce = null;
  let searchFinDebounce = null;
  let saveLayoutDebounceTimer = null;

  // Selection Options
  const CONTRACT_TYPES = [
    'AR Only',
    'CDOS ONLY',
    'CTS',
    'CTS AND DOAS',
    'CTS ONLY',
    'CTS WITH DOAS',
    'CTS WITH ERA',
    'CTS WITH MOA',
    'DOAS AND LA',
    'DOAS ONLY',
    'DOAS WITH ERA',
    'DOAS WITH ERA AND CTS',
    'EM ONLY',
    'ERA ONLY',
    'EROWA',
    'EROWA ONLY',
    'LA ONLY',
    'MOA ONLY',
    'MOA WITH DOAS',
    'NO CONTRACT',
    'NO CONTRACTS & AR',
    'NO DOCUMENTS'
  ];

  const MODES = [
    'FOR VERIFICATION',
    'NOT AVAILABLE',
    'OPEN TO LEASE',
    'OPEN TO SALE',
    'SECURED SALE',
    'SECURED LEASE',
    'SOLD TO PARTNER',
    'SS-STP',
    'TO BE IDENTIFIED',
    'RESCINDED',
    'NOT USABLE',
    'SECURED SALE - NON SP',
    'SECURED LEASE - NON SP'
  ];

  const TYPE_OPTIONS = [
    'LAND',
    'ROW',
    'LAND-SPECIAL'
  ];

  const ENTITY_OPTIONS = [
    'CIHC',
    'JJPNM',
    'LHHC',
    'LLL',
    'MBHC',
    'MGEN',
    'MNEC',
    'PIC',
    'SPBULC',
    'SPCC',
    'SPNEC',
    'SPPPHI',
    'SPTANC',
    'SPNorC',
    'TNI',
    'TSPI',
    'SPTARC',
    'SPTAGC',
    'SPPC'
  ];

  // Default Column Definitions — All Custom Columns
  const DEFAULT_COLUMNS = [
    { key: '__idx', label: '#', width: 45, fixed: true, readonly: true },
    { key: 'cr63f_control1', alt: 'Control', label: 'Control Number', width: 150, readonly: false },
    { key: 'cr63f_type', alt: 'Type', label: 'Type', width: 110, readonly: false, isOptions: true, options: TYPE_OPTIONS },
    { key: 'cr63f_indexno', alt: 'Index No', label: 'Index No', width: 120, readonly: false },
    { key: 'cr63f_areaindex', alt: 'Area-Index', label: 'Area-Index', width: 130, readonly: false },
    { key: 'cr63f_titleno', alt: 'Title No', label: 'Title No', width: 130, readonly: false },
    { key: 'cr63f_registeredowner', alt: 'Registered Owner', label: 'Registered Owner', width: 170, readonly: false },
    { key: 'cr63f_authorizedrepresentative', alt: 'Authorized Representative', label: 'Authorized Rep', width: 170, readonly: false },
    { key: 'cr63f_entity', alt: 'Entity', label: 'Entity', width: 120, readonly: false, isOptions: true, options: ENTITY_OPTIONS },
    { key: 'cr63f_province', alt: 'Province', label: 'Province', width: 120, readonly: false },
    { key: 'cr63f_municipality', alt: 'Municipality', label: 'Municipality', width: 130, readonly: false },
    { key: 'cr63f_dbadjustedbarangay', alt: 'DB Adjusted Barangay', label: 'Adjusted Barangay', width: 145, readonly: false },
    { key: 'cr63f_lotareaha', alt: 'Lot Area (Ha)', label: 'Lot Area (Ha)', width: 110, readonly: false, isNum: true },
    { key: 'new_contracttype', alt: 'Contract Type', label: 'Contract Type', width: 145, readonly: false, isOptions: true, options: CONTRACT_TYPES },
    { key: 'new_mode', alt: 'Mode', label: 'Mode', width: 165, readonly: false, isOptions: true, options: MODES },
    { key: 'cr63f_conversionstatus', alt: 'Conversion Status', label: 'Status', width: 130, readonly: false },
    { key: 'cr63f_rccprice', alt: 'RCC Price', label: 'RCC Price', width: 130, readonly: false, isNum: true, isPrice: true },
    { key: 'cr63f_rccterms', alt: 'RCC Terms', label: 'RCC Terms', width: 130, readonly: false },
    { key: 'cr63f_endorseddate', alt: 'Endorsed Date', label: 'Endorsed Date', width: 125, readonly: false, isDate: true },
    { key: 'cr63f_secureddate', alt: 'Secured Date', label: 'Secured Date', width: 125, readonly: false, isDate: true },
    { key: 'cr63f_securedendorsedate', alt: 'Secured Endorse Date', label: 'Secured Endorse Date', width: 145, readonly: false, isDate: true },
    { key: 'cr63f_totalcontractprice', alt: 'Total Contract Price', label: 'Total Contract Price', width: 145, readonly: false, isNum: true, isPrice: true },
    { key: 'cr63f_totalpaymentdeductibletotcp', alt: 'Payment Deductible', label: 'Payment Deductible', width: 145, readonly: false, isNum: true, isPrice: true },
    { key: 'cr63f_balance', alt: 'Balance', label: 'Balance', width: 140, readonly: true, isNum: true, isPrice: true, isBalance: true },
    { key: 'cr63f_paid', alt: '% Paid', label: '% Paid', width: 95, readonly: true, isNum: true, isPercent: true },
    { key: '__action', label: 'Actions', width: 55, fixed: true, readonly: true }
  ];

  // DOM Elements
  const dom = {
    // Top Bar
    dotTopConn: document.getElementById('dotTopConn'),
    lblTopConn: document.getElementById('lblTopConn'),
    btnTopConnPill: document.getElementById('btnTopConnPill'),
    dotFinanceRef: document.getElementById('dotFinanceRef'),
    lblTopFinance: document.getElementById('lblTopFinance'),
    btnTopFinancePill: document.getElementById('btnTopFinancePill'),
    btnToggleTheme: document.getElementById('btnToggleTheme'),
    tbTargetTable: document.getElementById('tbTargetTable'),

    // Nav
    tabButtons: document.querySelectorAll('.tab-btn'),
    viewPanels: document.querySelectorAll('.view-panel'),
    btnOpenNewEntry: document.getElementById('btnOpenNewEntry'),
    btnTopSyncDv: document.getElementById('btnTopSyncDv'),

    // Records View
    kpiRecordCount: document.getElementById('kpiRecordCount'),
    kpiTotalTcp: document.getElementById('kpiTotalTcp'),
    kpiTotalDeductible: document.getElementById('kpiTotalDeductible'),
    kpiTotalBalance: document.getElementById('kpiTotalBalance'),
    searchInput: document.getElementById('searchInput'),
    typeFilter: document.getElementById('typeFilter'),
    btnClearFilters: document.getElementById('btnClearFilters'),
    stagedEditsWrap: document.getElementById('stagedEditsWrap'),
    unsavedEditsChip: document.getElementById('unsavedEditsChip'),
    btnSaveStagedEdits: document.getElementById('btnSaveStagedEdits'),
    btnDiscardStagedEdits: document.getElementById('btnDiscardStagedEdits'),
    btnColumnsPicker: document.getElementById('btnColumnsPicker'),
    colPickerMenu: document.getElementById('colPickerMenu'),
    lblGridCount: document.getElementById('lblGridCount'),
    btnExportExcel: document.getElementById('btnExportExcel'),
    tableEmptyState: document.getElementById('tableEmptyState'),
    recordsTable: document.getElementById('recordsTable'),
    recordsTableHead: document.getElementById('recordsTableHead'),
    recordsTableBody: document.getElementById('recordsTableBody'),

    // Drawer / Form
    entryDrawer: document.getElementById('entryDrawer'),
    drawerTitle: document.getElementById('drawerTitle'),
    btnCloseDrawer: document.getElementById('btnCloseDrawer'),
    btnCancelDrawer: document.getElementById('btnCancelDrawer'),
    btnSaveDrawerEntry: document.getElementById('btnSaveDrawerEntry'),
    badgeMatchStatus: document.getElementById('badgeMatchStatus'),
    fRecordId: document.getElementById('fRecordId'),

    // Trigger & Identifier Fields
    fType: document.getElementById('fType'),
    fIndexNo: document.getElementById('fIndexNo'),
    fAreaIndex: document.getElementById('fAreaIndex'),
    areaIndexSuggestions: document.getElementById('areaIndexSuggestions'),
    fTitleNo: document.getElementById('fTitleNo'),
    fControl: document.getElementById('fControl'),

    // Payment Status Fields (Manual & Calculated)
    fTcp: document.getElementById('fTcp'),
    fPaymentDeductible: document.getElementById('fPaymentDeductible'),
    fBalance: document.getElementById('fBalance'),
    fPercentPaid: document.getElementById('fPercentPaid'),

    // Details Fields
    fRegisteredOwner: document.getElementById('fRegisteredOwner'),
    fAuthorizedRepresentative: document.getElementById('fAuthorizedRepresentative'),
    fEntity: document.getElementById('fEntity'),
    fProvince: document.getElementById('fProvince'),
    fMunicipality: document.getElementById('fMunicipality'),
    fDbAdjustedBarangay: document.getElementById('fDbAdjustedBarangay'),
    fLotAreaHa: document.getElementById('fLotAreaHa'),
    fContractType: document.getElementById('fContractType'),
    fMode: document.getElementById('fMode'),
    fConversionStatus: document.getElementById('fConversionStatus'),
    fRccTerms: document.getElementById('fRccTerms'),
    fRccPrice: document.getElementById('fRccPrice'),
    fSecuredDate: document.getElementById('fSecuredDate'),
    fSecuredEndorseDate: document.getElementById('fSecuredEndorseDate'),
    fEndorsedDate: document.getElementById('fEndorsedDate'),
    dynamicAttributesContainer: document.getElementById('dynamicAttributesContainer'),

    // Finance View
    lblActiveFinanceSource: document.getElementById('lblActiveFinanceSource'),
    lblActiveFinanceCount: document.getElementById('lblActiveFinanceCount'),
    lblActiveFinanceTableDisplay: document.getElementById('lblActiveFinanceTableDisplay'),
    btnFinanceGoToSettings: document.getElementById('btnFinanceGoToSettings'),
    btnUploadFinance: document.getElementById('btnUploadFinance'),
    btnClearFinanceData: document.getElementById('btnClearFinanceData'),
    searchInputFinance: document.getElementById('searchInputFinance'),
    typeFilterFinance: document.getElementById('typeFilterFinance'),
    btnClearFiltersFinance: document.getElementById('btnClearFiltersFinance'),
    btnColumnsPickerFinance: document.getElementById('btnColumnsPickerFinance'),
    colPickerMenuFinance: document.getElementById('colPickerMenuFinance'),
    lblFinanceGridCount: document.getElementById('lblFinanceGridCount'),
    btnExportExcelFinance: document.getElementById('btnExportExcelFinance'),
    financeEmptyState: document.getElementById('financeEmptyState'),
    financePreviewTable: document.getElementById('financePreviewTable'),
    financePreviewThead: document.getElementById('financePreviewThead'),
    financePreviewTbody: document.getElementById('financePreviewTbody'),

    // Settings View (Image 1 & 2)
    setConnDot: document.getElementById('setConnDot'),
    setConnText: document.getElementById('setConnText'),
    cfgEnvUrl: document.getElementById('cfgEnvUrl'),
    cfgTenantId: document.getElementById('cfgTenantId'),
    cfgClientId: document.getElementById('cfgClientId'),
    selSolution: document.getElementById('selSolution'),
    btnRefreshSolutions: document.getElementById('btnRefreshSolutions'),
    selLoaTable: document.getElementById('selLoaTable'),
    selFinanceTable: document.getElementById('selFinanceTable'),
    btnSignIn: document.getElementById('btnSignIn'),
    btnSignOut: document.getElementById('btnSignOut'),
    dvStatus: document.getElementById('dvStatus'),
    lblSignInHint: document.getElementById('lblSignInHint'),
    boxDeviceCode: document.getElementById('boxDeviceCode'),
    lblDeviceCode: document.getElementById('lblDeviceCode'),
    lblDeviceWaiting: document.getElementById('lblDeviceWaiting'),
    lnkDeviceUrl: document.getElementById('lnkDeviceUrl'),
    btnCopyCode: document.getElementById('btnCopyCode'),
    btnSaveConfig: document.getElementById('btnSaveConfig'),

    // Appearance (Image 1 Style)
    themeSeg: document.getElementById('themeSeg'),
    densitySeg: document.getElementById('densitySeg'),
    setRememberTable: document.getElementById('setRememberTable'),

    // Toast
    toastWrap: document.getElementById('toastWrap'),
  };

  // ==================== EXCEL-LIKE MULTI-CELL SELECTION & CLIPBOARD ====================
  const GRID_SELECTIONS = {}; // tableId -> { anchor:{r,c}, focus:{r,c}, dragging }
  const GRID_CFG = {};        // tableId -> { setMem, updateCell, refresh }
  let GRID_CELL_CACHE = {};   // tableId -> 2D array of data <td> elements
  let GRID_PASTE_PROXY = null;
  let GRID_GLOBAL_WIRED = false;

  function invalidateGridCache(tableId) {
    if (tableId) delete GRID_CELL_CACHE[tableId];
    else GRID_CELL_CACHE = {};
  }

  function gridCellsMatrix(tableEl) {
    if (!tableEl) return [];
    const id = tableEl.id;
    if (GRID_CELL_CACHE[id]) return GRID_CELL_CACHE[id];
    const m = Array.from(tableEl.querySelectorAll('tbody > tr')).map(tr =>
      Array.from(tr.querySelectorAll('td')).filter(td => td.dataset.colKey != null && td.dataset.colKey !== '__idx' && td.dataset.colKey !== '__action')
    );
    GRID_CELL_CACHE[id] = m;
    return m;
  }

  function gridPosOf(tbody, td) {
    const tr = td.closest('tr');
    if (!tr) return null;
    const r = Array.prototype.indexOf.call(tbody.children, tr);
    if (r < 0) return null;
    const cells = Array.from(tr.querySelectorAll('td')).filter(x => x.dataset.colKey != null && x.dataset.colKey !== '__idx' && x.dataset.colKey !== '__action');
    const c = cells.indexOf(td);
    if (c < 0) return null;
    return { r, c };
  }

  function gridNormRect(sel) {
    return {
      r1: Math.min(sel.anchor.r, sel.focus.r),
      r2: Math.max(sel.anchor.r, sel.focus.r),
      c1: Math.min(sel.anchor.c, sel.focus.c),
      c2: Math.max(sel.anchor.c, sel.focus.c),
    };
  }

  function gridClearHighlight(tableEl) {
    if (!tableEl) return;
    tableEl.querySelectorAll('td.cell-selected, td.cell-active')
      .forEach(td => td.classList.remove('cell-selected', 'cell-active'));
  }

  function gridApplyHighlight(tableEl) {
    if (!tableEl) return;
    const sel = GRID_SELECTIONS[tableEl.id];
    gridClearHighlight(tableEl);
    if (!sel) return;
    const m = gridCellsMatrix(tableEl);
    const { r1, r2, c1, c2 } = gridNormRect(sel);
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const td = m[r] && m[r][c];
        if (td) td.classList.add('cell-selected');
      }
    }
    const at = m[sel.anchor.r] && m[sel.anchor.r][sel.anchor.c];
    if (at) at.classList.add('cell-active');
  }

  function gridSelectionCount(tableEl) {
    const sel = GRID_SELECTIONS[tableEl.id];
    if (!sel) return 0;
    const { r1, r2, c1, c2 } = gridNormRect(sel);
    return (r2 - r1 + 1) * (c2 - c1 + 1);
  }

  function gridSelectionHasSelect(tableEl) {
    const sel = GRID_SELECTIONS[tableEl.id];
    if (!sel) return false;
    const m = gridCellsMatrix(tableEl);
    const { r1, r2, c1, c2 } = gridNormRect(sel);
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const td = m[r] && m[r][c];
        if (td && td.querySelector('select')) return true;
      }
    }
    return false;
  }

  function gridTargetInfo(tableEl) {
    const sel = GRID_SELECTIONS[tableEl.id];
    if (!sel) return null;
    const m = gridCellsMatrix(tableEl);
    const { r1, c1 } = gridNormRect(sel);
    const td = m[r1] && m[r1][c1];
    return { r1, c1, count: gridSelectionCount(tableEl), targetIsSelect: !!(td && td.querySelector('select')) };
  }

  function gridPasteProxy() {
    if (GRID_PASTE_PROXY && document.body.contains(GRID_PASTE_PROXY)) return GRID_PASTE_PROXY;
    const ta = document.createElement('textarea');
    ta.setAttribute('aria-hidden', 'true');
    ta.tabIndex = -1;
    ta.style.cssText = 'position:fixed;left:-9999px;top:0;width:1px;height:1px;opacity:0;';
    document.body.appendChild(ta);
    GRID_PASTE_PROXY = ta;
    return ta;
  }

  function gridCopyViaProxy(text) {
    const ta = gridPasteProxy();
    const prev = document.activeElement;
    ta.value = text;
    let ok = false;
    try {
      ta.focus();
      ta.select();
      ok = document.execCommand('copy');
    } catch (_) {
      ok = false;
    }
    ta.value = '';
    try {
      if (prev && prev.focus) prev.focus();
    } catch (_) {}
    return ok;
  }

  function gridCopyToClipboard(tableEl) {
    const count = gridSelectionCount(tableEl);
    const text = gridSelectionText(tableEl);
    const done = () => showToast(`Copied ${count} cell${count === 1 ? '' : 's'}.`, 'ok');
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done, () => {
        if (gridCopyViaProxy(text)) done();
        else showToast('Copy failed.', 'err');
      });
    } else if (gridCopyViaProxy(text)) {
      done();
    } else {
      showToast('Copy failed.', 'err');
    }
  }

  function gridFocusPasteProxy() {
    const ta = gridPasteProxy();
    ta.value = '';
    try {
      ta.focus();
    } catch (_) {}
  }

  function gridCellValue(td) {
    const ctrl = td.querySelector('select, input');
    if (ctrl) return String(ctrl.value != null ? ctrl.value : '');
    return String(td.textContent || '').trim();
  }

  function gridSelectionText(tableEl) {
    const sel = GRID_SELECTIONS[tableEl.id];
    if (!sel) return '';
    const m = gridCellsMatrix(tableEl);
    const { r1, r2, c1, c2 } = gridNormRect(sel);
    const lines = [];
    for (let r = r1; r <= r2; r++) {
      const cells = [];
      for (let c = c1; c <= c2; c++) {
        const td = m[r] && m[r][c];
        cells.push(td ? gridCellValue(td) : '');
      }
      lines.push(cells.join('\t'));
    }
    return lines.join('\n');
  }

  function gridCellEditor(td) {
    const ctrl = td.querySelector('select, input');
    if (!ctrl) return null;
    if (ctrl.tagName === 'INPUT' && ctrl.readOnly) return null;
    return ctrl;
  }

  async function gridPasteText(tableEl, text, startR, startC, verb) {
    const cfg = GRID_CFG[tableEl.id];
    if (!cfg) return;
    const m = gridCellsMatrix(tableEl);
    const block = String(text == null ? '' : text)
      .replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map(l => l.split('\t'));
    if (!block.length) return;

    const sel = GRID_SELECTIONS[tableEl.id];
    let r1 = startR, c1 = startC, selH = 1, selW = 1;
    if (sel) {
      const rect = gridNormRect(sel);
      r1 = rect.r1;
      c1 = rect.c1;
      selH = rect.r2 - rect.r1 + 1;
      selW = rect.c2 - rect.c1 + 1;
    }
    const bh = block.length;
    const bw = Math.max(1, ...block.map(rw => rw.length));
    const rows = (bh === 1) ? selH : ((selH > bh && selH % bh === 0) ? selH : bh);
    const cols = (bw === 1) ? selW : ((selW > bw && selW % bw === 0) ? selW : bw);

    let okCount = 0;
    for (let dr = 0; dr < rows; dr++) {
      const rr = r1 + dr;
      if (rr >= m.length) break;
      const srcRow = block[dr % bh];
      for (let dc = 0; dc < cols; dc++) {
        const cc = c1 + dc;
        const td = m[rr] && m[rr][cc];
        if (!td) continue;
        const ctrl = gridCellEditor(td);
        if (!ctrl) continue;

        let rawVal = String(srcRow[dc % bw] == null ? '' : srcRow[dc % bw]).trim();
        const colKey = td.dataset.colKey;
        const colDef = DEFAULT_COLUMNS.find(c => c.key === colKey) || { key: colKey };

        const isPrice = colDef.isPrice || colKey.toLowerCase().includes('price') || colKey.toLowerCase().includes('tcp');
        const isPercent = colDef.isPercent || colKey.toLowerCase().includes('paid') || colKey.includes('%');
        const isDate = colDef.isDate || colKey.toLowerCase().includes('date');

        if (isPrice) rawVal = rawVal.replace(/[₱$]/g, '').trim();
        if (isPercent) rawVal = formatPercent(rawVal);
        if (isDate) rawVal = maskDateString(rawVal);

        if (ctrl.tagName === 'SELECT') {
          const matchOpt = Array.from(ctrl.options).find(o => o.value.trim().toUpperCase() === rawVal.toUpperCase());
          if (matchOpt) {
            ctrl.value = matchOpt.value;
          } else if (rawVal) {
            ctrl.appendChild(new Option(rawVal, rawVal));
            ctrl.value = rawVal;
          } else {
            ctrl.value = '';
          }
        } else {
          ctrl.value = isPrice ? formatPrice(rawVal) : rawVal;
        }

        ctrl.classList.add('modified');
        td.classList.add('is-dirty');

        const rowIdx = parseInt(td.dataset.rowIdx, 10);
        if (cfg.setMem) cfg.setMem(rowIdx, colKey, rawVal);
        if (cfg.updateCell) cfg.updateCell(rowIdx, colKey, rawVal, ctrl, td);
        okCount++;
      }
    }

    if (!okCount) {
      showToast('Nothing to paste into the selected cells.', 'err');
      return;
    }

    GRID_SELECTIONS[tableEl.id] = null;
    gridClearHighlight(tableEl);
    showToast(`${verb || 'Pasted'} ${okCount} cell${okCount === 1 ? '' : 's'}.`, 'ok');
  }

  async function gridClearSelection(tableEl) {
    const cfg = GRID_CFG[tableEl.id];
    const sel = GRID_SELECTIONS[tableEl.id];
    if (!cfg || !sel) return;
    const m = gridCellsMatrix(tableEl);
    const rect = gridNormRect(sel);

    let okCount = 0;
    for (let r = rect.r1; r <= rect.r2; r++) {
      const tdRow = m[r];
      if (!tdRow) continue;
      for (let c = rect.c1; c <= rect.c2; c++) {
        const td = tdRow[c];
        if (!td) continue;
        const ctrl = gridCellEditor(td);
        if (!ctrl) continue;

        ctrl.value = '';
        ctrl.classList.add('modified');
        td.classList.add('is-dirty');

        const rowIdx = parseInt(td.dataset.rowIdx, 10);
        const colKey = td.dataset.colKey;
        if (cfg.setMem) cfg.setMem(rowIdx, colKey, '');
        if (cfg.updateCell) cfg.updateCell(rowIdx, colKey, '', ctrl, td);
        okCount++;
      }
    }

    if (!okCount) return;
    GRID_SELECTIONS[tableEl.id] = null;
    gridClearHighlight(tableEl);
    showToast(`Cleared ${okCount} cell${okCount === 1 ? '' : 's'}.`, 'ok');
  }

  function gridResolveActiveTable(target) {
    const ids = Object.keys(GRID_SELECTIONS).filter(id => GRID_SELECTIONS[id] && document.getElementById(id));
    if (!ids.length) return null;
    for (const id of ids) {
      const el = document.getElementById(id);
      if (target && el.contains(target)) return el;
    }
    const loose = !target || target === document.body || target === document.documentElement || target === GRID_PASTE_PROXY;
    if (loose) {
      for (const id of ids) {
        const el = document.getElementById(id);
        if (el && el.offsetParent !== null) return el;
      }
    }
    return null;
  }

  function gridInstallGlobalHandlers() {
    if (GRID_GLOBAL_WIRED) return;
    GRID_GLOBAL_WIRED = true;

    document.addEventListener('keydown', (e) => {
      const tableEl = gridResolveActiveTable(e.target);
      if (!tableEl) return;
      const sel = GRID_SELECTIONS[tableEl.id];
      if (!sel) return;

      if (e.key === 'Escape') {
        GRID_SELECTIONS[tableEl.id] = null;
        gridClearHighlight(tableEl);
        return;
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        const selCount = gridSelectionCount(tableEl);
        const ae = document.activeElement;
        const editing = !!ae && (ae.tagName === 'INPUT' || ae.tagName === 'SELECT' || ae.tagName === 'TEXTAREA');
        if (selCount > 1 || !editing) {
          e.preventDefault();
          gridClearSelection(tableEl);
        }
        return;
      }
      if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
      const key = (e.key || '').toLowerCase();
      const count = gridSelectionCount(tableEl);
      if (key === 'c') {
        if (count <= 1 && !gridSelectionHasSelect(tableEl)) return;
        e.preventDefault();
        gridCopyToClipboard(tableEl);
      } else if (key === 'v') {
        const info = gridTargetInfo(tableEl);
        if (!info || (count <= 1 && !info.targetIsSelect)) return;
        gridFocusPasteProxy();
      }
    });

    document.addEventListener('copy', (e) => {
      const tableEl = gridResolveActiveTable(e.target);
      if (!tableEl) return;
      if (gridSelectionCount(tableEl) <= 1) return;
      const text = gridSelectionText(tableEl);
      if (e.clipboardData) {
        e.clipboardData.setData('text/plain', text);
        e.preventDefault();
      }
    });

    document.addEventListener('paste', (e) => {
      const tableEl = gridResolveActiveTable(e.target);
      if (!tableEl) return;
      const sel = GRID_SELECTIONS[tableEl.id];
      if (!sel) return;
      const rect = gridNormRect(sel);
      const info = gridTargetInfo(tableEl) || { r1: rect.r1, c1: rect.c1, targetIsSelect: false };
      const text = e.clipboardData ? e.clipboardData.getData('text/plain') : '';
      const multiCell = /[\t\n]/.test(text);
      const count = gridSelectionCount(tableEl);
      if (!multiCell && count <= 1 && !info.targetIsSelect) return;
      e.preventDefault();
      gridPasteText(tableEl, text, info.r1, info.c1);
    });
  }

  function initGridSelection(tableEl, cfg) {
    if (!tableEl) return;
    GRID_CFG[tableEl.id] = cfg;
    if (tableEl.dataset.gridSelInit === '1') return;
    tableEl.dataset.gridSelInit = '1';
    const tbody = tableEl.querySelector('tbody');
    if (!tbody) return;
    let down = null;

    tbody.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      const td = e.target.closest('td');
      if (!td || td.dataset.colKey == null || td.dataset.colKey === '__idx' || td.dataset.colKey === '__action') return;
      const pos = gridPosOf(tbody, td);
      if (!pos) return;
      const sel = GRID_SELECTIONS[tableEl.id];
      if (e.shiftKey && sel) {
        sel.focus = pos;
        gridApplyHighlight(tableEl);
        e.preventDefault();
        return;
      }
      GRID_SELECTIONS[tableEl.id] = { anchor: { r: pos.r, c: pos.c }, focus: { r: pos.r, c: pos.c }, dragging: false };
      down = pos;
      gridApplyHighlight(tableEl);
    });

    tbody.addEventListener('mousemove', (e) => {
      if (!down || e.buttons !== 1) return;
      const td = e.target.closest('td');
      if (!td || td.dataset.colKey == null || td.dataset.colKey === '__idx' || td.dataset.colKey === '__action') return;
      const pos = gridPosOf(tbody, td);
      if (!pos || (pos.r === down.r && pos.c === down.c)) return;
      const sel = GRID_SELECTIONS[tableEl.id];
      if (!sel) return;
      sel.dragging = true;
      sel.focus = pos;
      gridApplyHighlight(tableEl);
      if (window.getSelection) window.getSelection().removeAllRanges();
      e.preventDefault();
    });

    document.addEventListener('mouseup', () => {
      down = null;
    });

    gridInstallGlobalHandlers();
  }

  // ==================== INITIALIZATION ====================
  function init() {
    loadTableConfig();
    setupNavigation();
    setupAppearance();
    setupRecordsView();
    setupDrawerForm();
    setupFinanceView();
    setupSettingsView();

    // Excel-like cell selection for LOA Records table
    initGridSelection(dom.recordsTable, {
      setMem: (rowIdx, colKey, val) => {
        const rec = state.filteredRecords[rowIdx];
        if (rec) {
          const id = rec.__id;
          if (id) {
            state.stagedEdits[id] = state.stagedEdits[id] || {};
            state.stagedEdits[id][colKey] = val;
          }
        }
      },
      updateCell: (rowIdx, colKey, val, inputEl, tdEl) => {
        const rec = state.filteredRecords[rowIdx];
        if (rec) {
          handleCellEdit(rec.__id, colKey, val, inputEl, tdEl);
        }
      },
      refresh: () => filterAndRenderRecords()
    });

    // Excel-like cell selection for Finance Data table
    initGridSelection(dom.financePreviewTable, {
      setMem: (rowIdx, colKey, val) => {
        const rows = (financeState.filteredRows && financeState.filteredRows.length ? financeState.filteredRows : financeState.rows);
        const row = rows[rowIdx];
        if (row) {
          row[colKey] = val;
          const origIdx = financeState.rows.indexOf(row);
          if (origIdx !== -1) financeState.rows[origIdx][colKey] = val;
        }
      },
      updateCell: (rowIdx, colKey, val, inputEl, tdEl) => {
        const rows = (financeState.filteredRows && financeState.filteredRows.length ? financeState.filteredRows : financeState.rows);
        const row = rows[rowIdx];
        if (row) {
          handleFinanceCellEdit(row.__id || rowIdx, colKey, val, inputEl, tdEl);
        }
      },
      refresh: () => filterAndRenderFinance()
    });

    // Wire scroll listeners for chunked rendering
    wireTableScrollListeners();

    // Check pywebview readiness
    if (window.pywebview && window.pywebview.api) {
      loadInitialState();
    } else {
      window.addEventListener('pywebviewready', loadInitialState);
    }
  }

  function wireTableScrollListeners() {
    const recordsWrap = dom.recordsTable.closest('.table-wrap');
    if (recordsWrap && !recordsWrap.dataset.scrollWired) {
      recordsWrap.dataset.scrollWired = '1';
      recordsWrap.addEventListener('scroll', () => {
        if (recordsWrap.scrollTop + recordsWrap.clientHeight >= recordsWrap.scrollHeight - 300) {
          if (state.renderedRecordsCount < state.filteredRecords.length) {
            renderRecordsChunk(state.renderedRecordsCount, CHUNK_SIZE);
          }
        }
      });
    }

    const financeWrap = dom.financePreviewTable.closest('.table-wrap');
    if (financeWrap && !financeWrap.dataset.scrollWired) {
      financeWrap.dataset.scrollWired = '1';
      financeWrap.addEventListener('scroll', () => {
        if (financeWrap.scrollTop + financeWrap.clientHeight >= financeWrap.scrollHeight - 300) {
          const displayRows = financeState.filteredRows && financeState.filteredRows.length ? financeState.filteredRows : (dom.searchInputFinance && dom.searchInputFinance.value ? [] : financeState.rows);
          if (financeState.renderedRowsCount < displayRows.length) {
            renderFinanceChunk(financeState.renderedRowsCount, CHUNK_SIZE);
          }
        }
      });
    }
  }

  async function loadInitialState() {
    try {
      const s = await window.pywebview.api.get_initial_state();
      applyState(s);

      if (s.table_layout) {
        applyTableLayout(s.table_layout);
      }

      const cfg = await window.pywebview.api.get_settings();
      populateSettings(cfg);

      if (s.connected) {
        loadSolutions();
        fetchRecords();
      }
    } catch (err) {
      console.warn('Init state error:', err);
    }
  }

  function applyState(s) {
    state.connected = s.connected;
    state.financeLoadedCount = s.finance_loaded_count || 0;
    state.financeSource = s.finance_source || 'None';
    state.recentFinanceFiles = s.recent_finance_files || [];
    state.theme = s.theme || 'light';
    state.activeLoaTable = s.loa_table || 'cr63f_loamonitoring';
    state.activeFinanceTable = s.finance_table || 'cr63f_financereference';
    state.activeSolutionId = s.solution_id || '';

    applyTheme(state.theme);

    // Update connection status
    if (state.connected) {
      if (dom.dotTopConn) dom.dotTopConn.classList.add('on');
      if (dom.lblTopConn) dom.lblTopConn.textContent = 'Connected';
      if (dom.setConnDot) dom.setConnDot.classList.add('on');
      if (dom.setConnText) dom.setConnText.textContent = 'DATAVERSE CONNECTED';
      if (dom.btnSignIn) dom.btnSignIn.classList.add('hidden');
      if (dom.btnSignOut) dom.btnSignOut.classList.remove('hidden');
      if (dom.dvStatus) {
        dom.dvStatus.classList.remove('hidden');
        dom.dvStatus.textContent = 'CONNECTED';
        dom.dvStatus.className = 'dv-badge connected';
      }
      if (dom.boxDeviceCode) dom.boxDeviceCode.classList.add('hidden');
    } else {
      if (dom.dotTopConn) dom.dotTopConn.classList.remove('on');
      if (dom.lblTopConn) dom.lblTopConn.textContent = 'Not connected';
      if (dom.setConnDot) dom.setConnDot.classList.remove('on');
      if (dom.setConnText) dom.setConnText.textContent = 'DATAVERSE NOT CONNECTED';
      if (dom.btnSignIn) {
        dom.btnSignIn.classList.remove('hidden');
        dom.btnSignIn.textContent = 'Connect';
      }
      if (dom.btnSignOut) dom.btnSignOut.classList.add('hidden');
      if (dom.dvStatus) dom.dvStatus.classList.add('hidden');
    }

    // Update finance reference status
    if (state.financeLoadedCount > 0) {
      dom.dotFinanceRef.classList.add('on');
      dom.lblTopFinance.textContent = `Finance Data: ${state.financeLoadedCount.toLocaleString()} rows`;
      if (dom.lblActiveFinanceCount) dom.lblActiveFinanceCount.textContent = `${state.financeLoadedCount.toLocaleString()} rows`;
      if (dom.lblActiveFinanceSource) dom.lblActiveFinanceSource.textContent = state.financeSource;
    } else {
      dom.dotFinanceRef.classList.remove('on');
      dom.lblTopFinance.textContent = 'Finance Data: None';
      if (dom.lblActiveFinanceCount) dom.lblActiveFinanceCount.textContent = '0 rows';
      if (dom.lblActiveFinanceSource) dom.lblActiveFinanceSource.textContent = 'None';
    }

    if (dom.tbTargetTable) dom.tbTargetTable.textContent = state.activeLoaTable;
    if (dom.lblActiveFinanceTableDisplay) {
      dom.lblActiveFinanceTableDisplay.textContent = state.activeFinanceTable;
    }
  }

  function applyTheme(theme) {
    state.theme = theme;
    document.documentElement.setAttribute('data-theme', theme);
    if (dom.themeSeg) updateSegActive(dom.themeSeg, 'theme', theme);
  }

  function setupAppearance() {
    if (dom.themeSeg) {
      dom.themeSeg.querySelectorAll('button').forEach(btn => {
        btn.addEventListener('click', () => {
          const t = btn.dataset.theme;
          applyTheme(t);
          window.pywebview.api.save_settings({ theme: t });
        });
      });
    }

    if (dom.densitySeg) {
      dom.densitySeg.querySelectorAll('button').forEach(btn => {
        btn.addEventListener('click', () => {
          const d = btn.dataset.density;
          applyDensity(d);
        });
      });
    }

    if (dom.setRememberTable) {
      const savedRemember = localStorage.getItem('loa_remember_publish') !== 'false';
      dom.setRememberTable.checked = savedRemember;
      dom.setRememberTable.addEventListener('change', () => {
        localStorage.setItem('loa_remember_publish', dom.setRememberTable.checked ? 'true' : 'false');
      });
    }

    const savedDensity = localStorage.getItem('loa_density') || 'comfortable';
    applyDensity(savedDensity, false);
  }

  function applyDensity(density, save = true) {
    document.documentElement.setAttribute('data-density', density);
    if (save) localStorage.setItem('loa_density', density);
    if (dom.densitySeg) updateSegActive(dom.densitySeg, 'density', density);

    let rh = 34, hh = 38;
    if (density === 'compact') { rh = 28; hh = 30; }
    applyRowHeight(rh, save, false);
    applyHeaderHeight(hh, save, false);
    if (save) saveTableLayoutBackend();
  }

  function applyRowHeight(px, save = true, notifyBackend = true) {
    document.documentElement.style.setProperty('--row-h', px + 'px');
    if (save) localStorage.setItem('loa_row_h', px);
    if (save && notifyBackend) saveTableLayoutBackend();
  }

  function applyHeaderHeight(px, save = true, notifyBackend = true) {
    document.documentElement.style.setProperty('--header-h', px + 'px');
    if (save) localStorage.setItem('loa_header_h', px);
    if (save && notifyBackend) saveTableLayoutBackend();
  }

  function updateSegActive(container, dataAttr, val) {
    if (!container) return;
    container.querySelectorAll('button').forEach(b => {
      if (b.dataset[dataAttr] === val) b.classList.add('active');
      else b.classList.remove('active');
    });
  }

  // ==================== NAVIGATION ====================
  function setupNavigation() {
    dom.tabButtons.forEach(btn => {
      btn.addEventListener('click', () => {
        const v = btn.getAttribute('data-view');
        switchView(v);
      });
    });

    dom.btnTopConnPill.addEventListener('click', () => switchView('settings'));
    dom.btnTopFinancePill.addEventListener('click', () => switchView('finance'));

    if (dom.btnToggleTheme) {
      dom.btnToggleTheme.addEventListener('click', () => {
        const nextTheme = state.theme === 'dark' ? 'light' : 'dark';
        applyTheme(nextTheme);
        window.pywebview.api.save_settings({ theme: nextTheme });
      });
    }

    // Top Bar "Sync Dataverse" button: syncs active tab
    dom.btnTopSyncDv.addEventListener('click', async () => {
      if (state.activeView === 'finance') {
        dom.btnTopSyncDv.disabled = true;
        dom.btnTopSyncDv.textContent = 'Syncing Finance...';
        try {
          const finRes = await window.pywebview.api.sync_finance_from_dataverse();
          if (finRes.success) {
            showToast(`Synced ${finRes.row_count} rows from Dataverse.`, 'ok');
            const s = await window.pywebview.api.get_initial_state();
            applyState(s);
            const sample = finRes.sample_rows || [];
            renderFinancePreview(finRes.headers || (sample.length ? Object.keys(sample[0]) : []), sample);
          } else {
            showToast(finRes.error || 'Failed to sync Finance data.', 'err');
          }
        } catch (e) {
          showToast('Error syncing finance: ' + e, 'err');
        } finally {
          dom.btnTopSyncDv.disabled = false;
          dom.btnTopSyncDv.textContent = 'Sync Dataverse';
        }
      } else {
        fetchRecords();
      }
    });
  }

  function switchView(viewName) {
    state.activeView = viewName;

    dom.tabButtons.forEach(btn => {
      if (btn.getAttribute('data-view') === viewName) {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
    });

    dom.viewPanels.forEach(p => {
      if (p.id === `view-${viewName}`) {
        p.classList.add('active');
      } else {
        p.classList.remove('active');
      }
    });

    if (viewName === 'settings' && state.connected && (!state.solutions || !state.solutions.length)) {
      loadSolutions();
    }
  }

  // ==================== NUMBER & PERCENT FORMATTING HELPERS ====================
  function parseAmount(val) {
    if (val === null || val === undefined || val === '') return 0;
    const clean = String(val).replace(/[^0-9.-]/g, '');
    const num = parseFloat(clean);
    return isNaN(num) ? 0 : num;
  }

  function formatAmount(val) {
    if (val === null || val === undefined || val === '') return '';
    const clean = String(val).replace(/[^0-9.-]/g, '');
    if (!clean || clean === '-' || clean === '.') return '';
    const num = parseFloat(clean);
    if (isNaN(num)) return String(val).replace(/[₱$]/g, '').trim();
    return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function formatPrice(val) {
    if (val === null || val === undefined || val === '') return '';
    const clean = String(val).replace(/[^0-9.-]/g, '');
    if (!clean || clean === '-' || clean === '.') return '';
    const num = parseFloat(clean);
    if (isNaN(num)) return String(val).trim();
    return '₱' + num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function maskDateString(val) {
    if (!val) return '';
    let digits = String(val).replace(/\D/g, '').slice(0, 8);
    if (digits.length >= 5) {
      return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
    } else if (digits.length >= 3) {
      return `${digits.slice(0, 2)}/${digits.slice(2)}`;
    }
    return digits;
  }

  function formatPercent(val) {
    if (val === null || val === undefined || val === '') return '0.00%';
    const s = String(val).trim();
    if (s.endsWith('%')) {
      const num = parseFloat(s.replace('%', '').trim());
      return isNaN(num) ? s : `${num.toFixed(2)}%`;
    }
    const num = parseFloat(s.replace(/,/g, ''));
    if (isNaN(num)) return s;
    if (num >= -1.0 && num <= 1.0 && num !== 0) {
      return `${(num * 100).toFixed(2)}%`;
    }
    return `${num.toFixed(2)}%`;
  }

  // ==================== TABLE CONFIG PERSISTENCE ====================
  function getSerializedLayout() {
    return {
      col_widths: state.colWidths || {},
      col_order: state.colOrder || [],
      visible_cols: Array.from(state.visibleCols || []),
      frozen_cols: Array.from(state.frozenCols || []),
      sort_col: state.sortCol || '',
      sort_dir: state.sortDir || '',
      density: localStorage.getItem('loa_density') || 'comfortable',
      row_h: parseInt(localStorage.getItem('loa_row_h'), 10) || 34,
      header_h: parseInt(localStorage.getItem('loa_header_h'), 10) || 38,
      fin_col_widths: (financeState && financeState.colWidths) || {},
      fin_col_order: (financeState && financeState.colOrder) || [],
      fin_visible_cols: (financeState && financeState.visibleCols) ? Array.from(financeState.visibleCols) : [],
      fin_frozen_cols: (financeState && financeState.frozenCols) ? Array.from(financeState.frozenCols) : [],
    };
  }

  function saveTableLayoutBackend() {
    if (!window.pywebview || !window.pywebview.api || !window.pywebview.api.save_table_layout) return;
    clearTimeout(saveLayoutDebounceTimer);
    saveLayoutDebounceTimer = setTimeout(() => {
      try {
        const layout = getSerializedLayout();
        window.pywebview.api.save_table_layout(layout);
      } catch (_) {}
    }, 150);
  }

  function saveFinanceTableConfig() {
    try {
      localStorage.setItem('fin_col_widths', JSON.stringify(financeState.colWidths));
      localStorage.setItem('fin_col_order', JSON.stringify(financeState.colOrder));
      localStorage.setItem('fin_visible_cols', JSON.stringify(Array.from(financeState.visibleCols)));
      localStorage.setItem('fin_frozen_cols', JSON.stringify(Array.from(financeState.frozenCols)));
    } catch (_) {}
    saveTableLayoutBackend();
  }

  function applyTableLayout(layout) {
    if (!layout || typeof layout !== 'object') return;

    if (layout.col_widths && Object.keys(layout.col_widths).length > 0) {
      state.colWidths = Object.assign({}, state.colWidths, layout.col_widths);
      try { localStorage.setItem('loa_col_widths', JSON.stringify(state.colWidths)); } catch (_) {}
    }

    if (Array.isArray(layout.col_order) && layout.col_order.length > 0) {
      const validKeys = new Set(DEFAULT_COLUMNS.map(c => c.key));
      const filtered = layout.col_order.filter(k => validKeys.has(k));
      DEFAULT_COLUMNS.forEach(c => {
        if (!filtered.includes(c.key)) filtered.push(c.key);
      });
      state.colOrder = filtered;
      try { localStorage.setItem('loa_col_order', JSON.stringify(state.colOrder)); } catch (_) {}
    }

    if (Array.isArray(layout.visible_cols) && layout.visible_cols.length > 0) {
      state.visibleCols = new Set(layout.visible_cols);
      DEFAULT_COLUMNS.forEach(c => {
        if (!state.visibleCols.has(c.key)) state.visibleCols.add(c.key);
      });
      try { localStorage.setItem('loa_visible_cols', JSON.stringify(Array.from(state.visibleCols))); } catch (_) {}
    }

    if (Array.isArray(layout.frozen_cols) && layout.frozen_cols.length > 0) {
      state.frozenCols = new Set(layout.frozen_cols);
      try { localStorage.setItem('loa_frozen_cols', JSON.stringify(Array.from(state.frozenCols))); } catch (_) {}
    }

    if (layout.sort_col !== undefined) {
      state.sortCol = layout.sort_col || null;
      state.sortDir = layout.sort_dir || null;
      try {
        localStorage.setItem('loa_sort_col', state.sortCol || '');
        localStorage.setItem('loa_sort_dir', state.sortDir || '');
      } catch (_) {}
    }

    if (layout.density) {
      applyDensity(layout.density, false);
    }
    if (layout.row_h) {
      applyRowHeight(layout.row_h, false, false);
    }
    if (layout.header_h) {
      applyHeaderHeight(layout.header_h, false, false);
    }

    // Finance layout
    if (layout.fin_col_widths && Object.keys(layout.fin_col_widths).length > 0) {
      financeState.colWidths = Object.assign({}, financeState.colWidths, layout.fin_col_widths);
      try { localStorage.setItem('fin_col_widths', JSON.stringify(financeState.colWidths)); } catch (_) {}
    }
    if (Array.isArray(layout.fin_col_order) && layout.fin_col_order.length > 0) {
      financeState.colOrder = layout.fin_col_order;
      try { localStorage.setItem('fin_col_order', JSON.stringify(financeState.colOrder)); } catch (_) {}
    }
    if (Array.isArray(layout.fin_visible_cols) && layout.fin_visible_cols.length > 0) {
      financeState.visibleCols = new Set(layout.fin_visible_cols);
      try { localStorage.setItem('fin_visible_cols', JSON.stringify(Array.from(financeState.visibleCols))); } catch (_) {}
    }
    if (Array.isArray(layout.fin_frozen_cols)) {
      financeState.frozenCols = new Set(layout.fin_frozen_cols);
      try { localStorage.setItem('fin_frozen_cols', JSON.stringify(Array.from(financeState.frozenCols))); } catch (_) {}
    }

    renderTableHead();
    renderColumnPickerMenu();
  }

  function loadTableConfig() {
    try {
      const sw = localStorage.getItem('loa_col_widths');
      state.colWidths = sw ? JSON.parse(sw) : {};
    } catch (_) { state.colWidths = {}; }

    try {
      const so = localStorage.getItem('loa_col_order');
      if (so) {
        const parsed = JSON.parse(so);
        const validKeys = new Set(DEFAULT_COLUMNS.map(c => c.key));
        const filtered = parsed.filter(k => validKeys.has(k));
        DEFAULT_COLUMNS.forEach(c => {
          if (!filtered.includes(c.key)) filtered.push(c.key);
        });
        state.colOrder = filtered;
      } else {
        state.colOrder = DEFAULT_COLUMNS.map(c => c.key);
      }
    } catch (_) {
      state.colOrder = DEFAULT_COLUMNS.map(c => c.key);
    }

    try {
      const sv = localStorage.getItem('loa_visible_cols');
      if (sv) {
        state.visibleCols = new Set(JSON.parse(sv));
        DEFAULT_COLUMNS.forEach(c => {
          if (!state.visibleCols.has(c.key)) state.visibleCols.add(c.key);
        });
      } else {
        state.visibleCols = new Set(DEFAULT_COLUMNS.map(c => c.key));
      }
    } catch (_) {
      state.visibleCols = new Set(DEFAULT_COLUMNS.map(c => c.key));
    }

    try {
      const sf = localStorage.getItem('loa_frozen_cols');
      state.frozenCols = sf ? new Set(JSON.parse(sf)) : new Set(['__idx', 'cr63f_control1']);
    } catch (_) {
      state.frozenCols = new Set(['__idx', 'cr63f_control1']);
    }

    try {
      const sc = localStorage.getItem('loa_sort_col');
      const sd = localStorage.getItem('loa_sort_dir');
      state.sortCol = sc || null;
      state.sortDir = sd || null;
    } catch (_) {
      state.sortCol = null;
      state.sortDir = null;
    }
  }

  function saveTableConfig() {
    try {
      localStorage.setItem('loa_col_widths', JSON.stringify(state.colWidths));
      localStorage.setItem('loa_col_order', JSON.stringify(state.colOrder));
      localStorage.setItem('loa_visible_cols', JSON.stringify(Array.from(state.visibleCols)));
      localStorage.setItem('loa_frozen_cols', JSON.stringify(Array.from(state.frozenCols)));
      localStorage.setItem('loa_sort_col', state.sortCol || '');
      localStorage.setItem('loa_sort_dir', state.sortDir || '');
    } catch (_) {}
    saveTableLayoutBackend();
  }

  function resetTableConfig() {
    localStorage.removeItem('loa_col_widths');
    localStorage.removeItem('loa_col_order');
    localStorage.removeItem('loa_visible_cols');
    localStorage.removeItem('loa_frozen_cols');
    localStorage.removeItem('loa_sort_col');
    localStorage.removeItem('loa_sort_dir');
    loadTableConfig();
    saveTableLayoutBackend();
    renderTableHead();
    renderColumnPickerMenu();
    filterAndRenderRecords();
    showToast('Table layout reset to default.', 'ok');
  }

  function getOrderedColumns() {
    const map = new Map(DEFAULT_COLUMNS.map(c => [c.key, c]));
    const ordered = [];
    state.colOrder.forEach(k => {
      if (map.has(k)) ordered.push(map.get(k));
    });
    DEFAULT_COLUMNS.forEach(c => {
      if (!state.colOrder.includes(c.key)) ordered.push(c);
    });
    return ordered;
  }

  function getVisibleColumns() {
    return getOrderedColumns().filter(c => c.fixed || state.visibleCols.has(c.key));
  }

  function getFrozenOffset(colIdx, colsList, frozenSet, colWidthsMap) {
    let offset = 0;
    for (let i = 0; i < colIdx && i < colsList.length; i++) {
      const key = colsList[i].key || colsList[i];
      if (frozenSet.has(key)) {
        const w = (colWidthsMap && colWidthsMap[key]) || colsList[i].width || 140;
        offset += w;
      }
    }
    return offset;
  }

  function isTrailingFrozen(colIdx, colsList, frozenSet) {
    if (!frozenSet) return false;
    const curKey = colsList[colIdx].key || colsList[colIdx];
    if (!frozenSet.has(curKey)) return false;

    // Only show trailing frozen divider line if at least one custom/data column is pinned
    const hasCustomPinned = Array.from(frozenSet).some(k => k !== '__idx' && k !== '__action');
    if (!hasCustomPinned) return false;

    for (let i = colIdx + 1; i < colsList.length; i++) {
      const k = colsList[i].key || colsList[i];
      if (frozenSet.has(k)) return false;
    }
    return true;
  }

  function updateTableColumnWidths(tableEl, colKey, width) {
    if (!tableEl) return;
    const th = tableEl.querySelector(`thead th[data-col-key="${colKey}"], thead th[data-fin-col="${colKey}"]`);
    if (th) {
      th.style.width = `${width}px`;
      th.style.minWidth = `${width}px`;
      th.style.maxWidth = `${width}px`;
    }
    const tds = tableEl.querySelectorAll(`tbody td[data-col-key="${colKey}"]`);
    tds.forEach(td => {
      td.style.width = `${width}px`;
      td.style.minWidth = `${width}px`;
      td.style.maxWidth = `${width}px`;
    });
  }

  function updateStickyOffsets(tableEl, colsList, frozenSet, colWidthsMap, baseOffset = 0) {
    if (!tableEl) return;
    let offset = baseOffset;
    colsList.forEach((col, colIdx) => {
      const key = col.key || col;
      if (frozenSet.has(key)) {
        const w = (colWidthsMap && colWidthsMap[key]) || col.width || 140;
        const th = tableEl.querySelector(`thead th[data-col-key="${key}"], thead th[data-fin-col="${key}"]`);
        if (th) th.style.left = `${offset}px`;
        const tds = tableEl.querySelectorAll(`tbody td[data-col-key="${key}"]`);
        tds.forEach(td => td.style.left = `${offset}px`);
        offset += w;
      }
    });
  }

  // ==================== RECORDS HUB ====================
  function setupRecordsView() {
    setupColumnPicker();

    dom.searchInput.addEventListener('input', () => {
      clearTimeout(searchDebounce);
      searchDebounce = setTimeout(filterAndRenderRecords, 120);
    });

    dom.typeFilter.addEventListener('change', filterAndRenderRecords);
    dom.btnExportExcel.addEventListener('click', exportRecordsExcel);

    if (dom.btnClearFilters) {
      dom.btnClearFilters.addEventListener('click', () => {
        dom.searchInput.value = '';
        dom.typeFilter.value = 'ALL';
        state.columnFilters = {};
        state.columnFiltersOpen = {};
        renderTableHead();
        filterAndRenderRecords();
        showToast('All filters and search reset.', 'ok');
      });
    }

    if (dom.btnSaveStagedEdits) {
      dom.btnSaveStagedEdits.addEventListener('click', saveStagedEdits);
    }

    if (dom.btnDiscardStagedEdits) {
      dom.btnDiscardStagedEdits.addEventListener('click', discardStagedEdits);
    }

    if (dom.btnFinanceGoToSettings) {
      dom.btnFinanceGoToSettings.addEventListener('click', () => switchView('settings'));
    }
  }

  function setupColumnPicker() {
    if (!dom.btnColumnsPicker || !dom.colPickerMenu) return;

    dom.btnColumnsPicker.addEventListener('click', (e) => {
      e.stopPropagation();
      dom.colPickerMenu.classList.toggle('hidden');
    });

    document.addEventListener('click', (e) => {
      if (!e.target.closest('.col-picker-container')) {
        dom.colPickerMenu.classList.add('hidden');
      }
    });

    renderColumnPickerMenu();
  }

  function renderColumnPickerMenu() {
    if (!dom.colPickerMenu) return;

    const orderedNonFixed = state.colOrder
      .map(k => DEFAULT_COLUMNS.find(c => c.key === k))
      .filter(c => c && !c.fixed);

    let html = '';
    orderedNonFixed.forEach((col, idx) => {
      const isChecked = state.visibleCols.has(col.key);
      const isFirst = idx === 0;
      const isLast = idx === orderedNonFixed.length - 1;

      html += `
        <div class="col-picker-item">
          <label class="col-picker-label">
            <input type="checkbox" data-col-toggle="${col.key}" ${isChecked ? 'checked' : ''} />
            <span>${escapeHtml(col.label)}</span>
          </label>
          <div class="col-reorder-btns">
            <button type="button" class="col-reorder-btn" data-move-up="${col.key}" ${isFirst ? 'disabled' : ''} title="Move left / earlier">◄</button>
            <button type="button" class="col-reorder-btn" data-move-down="${col.key}" ${isLast ? 'disabled' : ''} title="Move right / later">►</button>
          </div>
        </div>
      `;
    });

    html += `
      <div class="col-picker-actions">
        <button type="button" class="btn xs ghost" id="btnResetLayout" style="width: 100%; font-size: 11px;">Reset Table Layout</button>
      </div>
    `;

    dom.colPickerMenu.innerHTML = html;

    dom.colPickerMenu.querySelectorAll('[data-col-toggle]').forEach(chk => {
      chk.addEventListener('change', () => {
        const key = chk.dataset.colToggle;
        if (chk.checked) {
          state.visibleCols.add(key);
        } else {
          if (state.visibleCols.size > 1) {
            state.visibleCols.delete(key);
          } else {
            chk.checked = true;
            showToast('At least one column must remain visible.', 'err');
            return;
          }
        }
        saveTableConfig();
        renderTableHead();
        filterAndRenderRecords();
      });
    });

    dom.colPickerMenu.querySelectorAll('[data-move-up]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.moveUp;
        moveColumnOrder(key, -1);
      });
    });

    dom.colPickerMenu.querySelectorAll('[data-move-down]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.moveDown;
        moveColumnOrder(key, 1);
      });
    });

    const btnReset = dom.colPickerMenu.querySelector('#btnResetLayout');
    if (btnReset) {
      btnReset.addEventListener('click', (e) => {
        e.stopPropagation();
        resetTableConfig();
      });
    }
  }

  function moveColumnOrder(colKey, direction) {
    const idx = state.colOrder.indexOf(colKey);
    if (idx === -1) return;
    const targetIdx = idx + direction;

    if (targetIdx <= 0 || targetIdx >= state.colOrder.length - 1) return;

    const temp = state.colOrder[idx];
    state.colOrder[idx] = state.colOrder[targetIdx];
    state.colOrder[targetIdx] = temp;

    saveTableConfig();
    renderTableHead();
    renderColumnPickerMenu();
    filterAndRenderRecords();
  }

  async function fetchRecords() {
    try {
      dom.btnTopSyncDv.disabled = true;
      dom.btnTopSyncDv.textContent = 'Syncing...';

      const [res, finRes] = await Promise.all([
        window.pywebview.api.fetch_loa_records(),
        window.pywebview.api.sync_finance_from_dataverse().catch(e => ({ success: false, error: String(e) }))
      ]);

      if (res.success) {
        state.records = res.records || [];
        dom.kpiRecordCount.textContent = (res.count || 0).toLocaleString();
        dom.kpiTotalTcp.textContent = `₱${res.total_tcp_formatted || '0.00'}`;
        dom.kpiTotalDeductible.textContent = `₱${res.total_deductible_formatted || '0.00'}`;
        dom.kpiTotalBalance.textContent = `₱${res.total_balance_formatted || '0.00'}`;

        populateTypeFilter(state.records);
        renderTableHead();
        filterAndRenderRecords();
        showToast(`Loaded ${res.count.toLocaleString()} records from Dataverse.`, 'ok');
      } else {
        showToast(res.error || 'Failed to fetch Dataverse records.', 'err');
      }

      if (finRes && finRes.success) {
        const s = await window.pywebview.api.get_initial_state();
        applyState(s);
        const sample = finRes.sample_rows || [];
        renderFinancePreview(finRes.headers || (sample.length ? Object.keys(sample[0]) : []), sample);
      }
    } catch (err) {
      showToast('Error: ' + err, 'err');
    } finally {
      dom.btnTopSyncDv.disabled = false;
      dom.btnTopSyncDv.textContent = 'Sync Dataverse';
    }
  }

  function populateTypeFilter(records) {
    const typeSet = new Set(['LAND', 'ROW', 'LAND-SPECIAL']);
    records.forEach(r => {
      const t = String(r['Type'] || r['TYPE'] || r['cr63f_type'] || '').trim().toUpperCase();
      if (t) typeSet.add(t);
    });

    const currentVal = dom.typeFilter.value;
    let html = `<option value="ALL">All Types</option>`;
    Array.from(typeSet).sort().forEach(t => {
      html += `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`;
    });
    dom.typeFilter.innerHTML = html;
    if (typeSet.has(currentVal)) {
      dom.typeFilter.value = currentVal;
    } else {
      dom.typeFilter.value = 'ALL';
    }

    if (dom.fType && dom.fType.tagName === 'SELECT') {
      const curFType = dom.fType.value || 'LAND';
      dom.fType.innerHTML = `
        <option value="LAND">LAND</option>
        <option value="ROW">ROW</option>
        <option value="LAND-SPECIAL">LAND-SPECIAL</option>
      `;
      if (['LAND', 'ROW', 'LAND-SPECIAL'].includes(curFType)) {
        dom.fType.value = curFType;
      }
    }
  }

  function renderTableHead() {
    if (!dom.recordsTableHead) return;
    const visibleCols = getVisibleColumns();
    let html = '<tr>';

    visibleCols.forEach((col, colIdx) => {
      const width = state.colWidths[col.key] || col.width || 120;
      const isFilterActive = !!(state.columnFilters[col.key]);
      const isFilterOpen = !!(state.columnFiltersOpen[col.key]);
      const isSorted = state.sortCol === col.key;
      const sortIcon = isSorted ? (state.sortDir === 'asc' ? '▲' : '▼') : '';
      const isFrozen = state.frozenCols.has(col.key);
      const isLastFrz = isTrailingFrozen(colIdx, visibleCols, state.frozenCols);
      const leftOffset = isFrozen ? getFrozenOffset(colIdx, visibleCols, state.frozenCols, state.colWidths) : 0;

      const frozenClass = `${isFrozen ? 'col-frozen' : ''} ${isLastFrz ? 'col-frozen-last' : ''}`.trim();
      const leftStyle = isFrozen ? `left: ${leftOffset}px;` : '';

      if (col.key === '__idx' || col.key === '__action') {
        html += `
          <th class="${col.key === '__action' ? 'col-action' : ''} ${frozenClass}" style="width: ${width}px; min-width: ${width}px; max-width: ${width}px; ${leftStyle}">
            <div class="th-header-box" style="justify-content: center; align-items: center;">
              <span class="th-col-name" style="text-align: center;">${escapeHtml(col.label)}</span>
            </div>
            <div class="th-bottom-resizer" title="Drag to adjust header height"></div>
          </th>`;
        return;
      }

      html += `
        <th draggable="true" data-col-key="${col.key}" class="${frozenClass}" style="width: ${width}px; min-width: ${width}px; max-width: ${width}px; ${leftStyle}">
          <div class="th-header-box">
            <div class="th-title-wrap">
              <button type="button" class="th-pin-btn ${isFrozen ? 'pinned' : ''}" data-pin-col="${col.key}" title="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(col.label)}">
                <svg viewBox="0 0 24 24" width="11" height="11" fill="currentColor" stroke="none">
                  <path d="M16 12V4h1V2H7v2h1v8l-2 2v2h5.2v6h1.6v-6H18v-2l-2-2z"/>
                </svg>
              </button>
              <div class="th-sort-wrap" data-sort-col="${col.key}" title="Click to sort by ${escapeHtml(col.label)} (Drag to rearrange)">
                <span class="th-col-name">${escapeHtml(col.label)}</span>
                <span class="th-sort-icon ${isSorted ? 'active' : ''}">${sortIcon}</span>
              </div>
              <button type="button" class="th-chevron-btn ${isFilterActive ? 'active' : ''}" data-filter-toggle="${col.key}" title="Filter ${escapeHtml(col.label)}">
                <svg viewBox="0 0 24 24" width="10" height="10" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"></polyline></svg>
              </button>
            </div>
            <div class="th-filter-popup ${isFilterOpen ? '' : 'hidden'}" id="colFilterBox_${col.key}">
              <input class="th-filter-input" data-filter-col="${col.key}" value="${escapeHtml(state.columnFilters[col.key] || '')}" placeholder="Filter ${escapeHtml(col.label)}..." />
            </div>
          </div>
          <div class="th-resizer" data-col-key="${col.key}" title="Drag column width (Double-click to reset)"></div>
          <div class="th-bottom-resizer" title="Drag to adjust header height"></div>
        </th>
      `;
    });

    html += '</tr>';
    dom.recordsTableHead.innerHTML = html;
    attachHeaderResizers();
    attachHeaderFilterEvents();
    attachHeaderSortEvents();
    attachHeaderDragDropEvents();
    attachHeaderPinEvents();
  }

  function attachHeaderPinEvents() {
    dom.recordsTableHead.querySelectorAll('[data-pin-col]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.pinCol;
        if (state.frozenCols.has(key)) {
          state.frozenCols.delete(key);
        } else {
          state.frozenCols.add(key);
        }
        saveTableConfig();
        renderTableHead();
        filterAndRenderRecords();
      });
    });
  }

  function attachHeaderSortEvents() {
    dom.recordsTableHead.querySelectorAll('[data-sort-col]').forEach(wrap => {
      wrap.addEventListener('click', (e) => {
        if (e.target.closest('.th-chevron-btn') || e.target.closest('.th-resizer') || e.target.closest('.th-filter-popup') || e.target.closest('.th-pin-btn')) return;
        const colKey = wrap.dataset.sortCol;
        handleHeaderSort(colKey);
      });
    });
  }

  function handleHeaderSort(colKey) {
    if (state.sortCol === colKey) {
      if (state.sortDir === 'asc') {
        state.sortDir = 'desc';
      } else if (state.sortDir === 'desc') {
        state.sortCol = null;
        state.sortDir = null;
      } else {
        state.sortDir = 'asc';
      }
    } else {
      state.sortCol = colKey;
      state.sortDir = 'asc';
    }
    saveTableConfig();
    renderTableHead();
    filterAndRenderRecords();
  }

  function attachHeaderDragDropEvents() {
    let draggedColKey = null;

    const ths = dom.recordsTableHead.querySelectorAll('th[draggable="true"]');
    ths.forEach(th => {
      th.addEventListener('dragstart', (e) => {
        draggedColKey = th.dataset.colKey;
        th.classList.add('col-dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', draggedColKey);
      });

      th.addEventListener('dragend', () => {
        th.classList.remove('col-dragging');
        ths.forEach(t => t.classList.remove('drag-over-left', 'drag-over-right'));
      });

      th.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        const targetColKey = th.dataset.colKey;
        if (!targetColKey || targetColKey === draggedColKey) return;

        const rect = th.getBoundingClientRect();
        const midX = rect.left + rect.width / 2;
        if (e.clientX < midX) {
          th.classList.add('drag-over-left');
          th.classList.remove('drag-over-right');
        } else {
          th.classList.add('drag-over-right');
          th.classList.remove('drag-over-left');
        }
      });

      th.addEventListener('dragleave', () => {
        th.classList.remove('drag-over-left', 'drag-over-right');
      });

      th.addEventListener('drop', (e) => {
        e.preventDefault();
        th.classList.remove('drag-over-left', 'drag-over-right');
        const targetColKey = th.dataset.colKey;
        if (!draggedColKey || !targetColKey || draggedColKey === targetColKey) return;

        const rect = th.getBoundingClientRect();
        const midX = rect.left + rect.width / 2;
        const insertBefore = e.clientX < midX;

        const fromIdx = state.colOrder.indexOf(draggedColKey);
        if (fromIdx === -1) return;
        state.colOrder.splice(fromIdx, 1);

        let toIdx = state.colOrder.indexOf(targetColKey);
        if (!insertBefore) toIdx++;
        state.colOrder.splice(toIdx, 0, draggedColKey);

        saveTableConfig();
        renderTableHead();
        renderColumnPickerMenu();
        filterAndRenderRecords();
      });
    });
  }

  function attachHeaderResizers() {
    const resizers = dom.recordsTableHead.querySelectorAll('.th-resizer');
    resizers.forEach(resizer => {
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const colKey = resizer.dataset.colKey;
        const th = resizer.closest('th');
        if (!th) return;

        const startX = e.pageX;
        const startWidth = th.offsetWidth;
        resizer.classList.add('is-resizing');
        document.body.classList.add('col-resizing');

        let rAF = null;
        let currentWidth = startWidth;

        const onMouseMove = (moveEvent) => {
          const delta = moveEvent.pageX - startX;
          currentWidth = Math.max(45, Math.min(800, startWidth + delta));
          if (rAF) cancelAnimationFrame(rAF);
          rAF = requestAnimationFrame(() => {
            updateTableColumnWidths(dom.recordsTable, colKey, currentWidth);
          });
        };

        const onMouseUp = () => {
          if (rAF) cancelAnimationFrame(rAF);
          window.removeEventListener('mousemove', onMouseMove);
          window.removeEventListener('mouseup', onMouseUp);
          resizer.classList.remove('is-resizing');
          document.body.classList.remove('col-resizing');
          state.colWidths[colKey] = currentWidth;
          saveTableConfig();
          if (state.frozenCols.has(colKey)) {
            updateStickyOffsets(dom.recordsTable, getVisibleColumns(), state.frozenCols, state.colWidths, 0);
          }
        };

        window.addEventListener('mousemove', onMouseMove);
        window.addEventListener('mouseup', onMouseUp);
      });

      resizer.addEventListener('dblclick', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const colKey = resizer.dataset.colKey;
        const colDef = DEFAULT_COLUMNS.find(c => c.key === colKey);
        if (colDef) {
          state.colWidths[colKey] = colDef.width;
          saveTableConfig();
          updateTableColumnWidths(dom.recordsTable, colKey, colDef.width);
          if (state.frozenCols.has(colKey)) {
            updateStickyOffsets(dom.recordsTable, getVisibleColumns(), state.frozenCols, state.colWidths, 0);
          }
        }
      });
    });

    dom.recordsTableHead.querySelectorAll('.th-bottom-resizer').forEach(resizer => {
      attachBottomHeaderResizer(resizer.closest('th'));
    });
  }

  function attachBottomHeaderResizer(th) {
    if (!th) return;
    const resizer = th.querySelector('.th-bottom-resizer');
    if (!resizer) return;

    resizer.addEventListener('mousedown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const startY = e.clientY;
      const currentH = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--header-h')) || 38;
      document.body.classList.add('row-resizing');

      let rAF = null;
      let newH = currentH;

      const onMouseMove = (ev) => {
        const delta = ev.clientY - startY;
        newH = Math.max(26, Math.min(80, currentH + delta));
        if (rAF) cancelAnimationFrame(rAF);
        rAF = requestAnimationFrame(() => {
          document.documentElement.style.setProperty('--header-h', newH + 'px');
        });
      };

      const onMouseUp = () => {
        if (rAF) cancelAnimationFrame(rAF);
        document.body.classList.remove('row-resizing');
        window.removeEventListener('mousemove', onMouseMove);
        window.removeEventListener('mouseup', onMouseUp);
        localStorage.setItem('loa_header_h', newH);
      };

      window.addEventListener('mousemove', onMouseMove);
      window.addEventListener('mouseup', onMouseUp);
    });
  }

  function attachHeaderFilterEvents() {
    dom.recordsTableHead.querySelectorAll('[data-filter-toggle]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const colKey = btn.dataset.filterToggle;
        state.columnFiltersOpen[colKey] = !state.columnFiltersOpen[colKey];
        const popup = document.getElementById(`colFilterBox_${colKey}`);
        if (popup) {
          if (state.columnFiltersOpen[colKey]) {
            popup.classList.remove('hidden');
            const input = popup.querySelector('input');
            if (input) input.focus();
          } else {
            popup.classList.add('hidden');
          }
        }
      });
    });

    dom.recordsTableHead.querySelectorAll('.th-filter-input').forEach(input => {
      input.addEventListener('input', () => {
        const colKey = input.dataset.filterCol;
        state.columnFilters[colKey] = input.value.trim().toLowerCase();
        filterAndRenderRecords();
      });
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
          const colKey = input.dataset.filterCol;
          state.columnFiltersOpen[colKey] = false;
          const popup = document.getElementById(`colFilterBox_${colKey}`);
          if (popup) popup.classList.add('hidden');
        }
      });
    });
  }

  function getRowCellValue(row, col) {
    if (!row) return '';
    const id = row.__id;
    if (id && state.stagedEdits[id] && state.stagedEdits[id][col.key] !== undefined) {
      const val = state.stagedEdits[id][col.key];
      return col.isPercent ? formatPercent(val) : val;
    }
    let val = '';
    if (row[col.key] !== undefined && row[col.key] !== null) val = row[col.key];
    else if (col.alt && row[col.alt] !== undefined && row[col.alt] !== null) val = row[col.alt];
    else {
      const upper = (col.alt || col.key).toUpperCase();
      if (row[upper] !== undefined && row[upper] !== null) val = row[upper];
    }
    if (col.isPercent) return formatPercent(val);
    return val;
  }

  function filterAndRenderRecords() {
    const q = dom.searchInput.value.toLowerCase().trim();
    const typeFilt = dom.typeFilter.value;
    const colFilterKeys = Object.keys(state.columnFilters).filter(k => !!state.columnFilters[k]);

    state.filteredRecords = state.records.filter(r => {
      const rType = String(r['TYPE'] || r['Type'] || r['cr63f_type'] || '').trim().toUpperCase();
      if (typeFilt !== 'ALL' && rType !== typeFilt) return false;

      for (const colKey of colFilterKeys) {
        const query = state.columnFilters[colKey];
        const colDef = DEFAULT_COLUMNS.find(c => c.key === colKey) || { key: colKey };
        const val = String(getRowCellValue(r, colDef)).toLowerCase();
        if (!val.includes(query)) return false;
      }

      if (!q) return true;
      return Object.values(r).some(v => String(v).toLowerCase().includes(q));
    });

    if (state.sortCol && state.sortDir) {
      const colDef = DEFAULT_COLUMNS.find(c => c.key === state.sortCol) || { key: state.sortCol };
      const mult = state.sortDir === 'asc' ? 1 : -1;
      state.filteredRecords.sort((a, b) => {
        const valA = getRowCellValue(a, colDef);
        const valB = getRowCellValue(b, colDef);

        if (colDef.isNum || colDef.isPercent) {
          const numA = parseFloat(String(valA).replace(/[^0-9.-]/g, '')) || 0;
          const numB = parseFloat(String(valB).replace(/[^0-9.-]/g, '')) || 0;
          return (numA - numB) * mult;
        }

        return String(valA || '').localeCompare(String(valB || ''), undefined, { numeric: true, sensitivity: 'base' }) * mult;
      });
    }

    dom.lblGridCount.textContent = `${state.filteredRecords.length.toLocaleString()} records`;

    invalidateGridCache('recordsTable');

    if (!state.filteredRecords.length) {
      dom.recordsTableBody.innerHTML = '';
      dom.recordsTable.classList.add('hidden');
      dom.tableEmptyState.classList.remove('hidden');
      state.renderedRecordsCount = 0;
      return;
    }

    dom.tableEmptyState.classList.add('hidden');
    dom.recordsTable.classList.remove('hidden');
    dom.recordsTableBody.innerHTML = '';
    state.renderedRecordsCount = 0;

    renderRecordsChunk(0, CHUNK_SIZE);
  }

  function renderRecordsChunk(startIdx, count) {
    if (!dom.recordsTableBody) return;
    const visibleCols = getVisibleColumns();
    const frag = document.createDocumentFragment();
    const end = Math.min(startIdx + count, state.filteredRecords.length);

    for (let idx = startIdx; idx < end; idx++) {
      const r = state.filteredRecords[idx];
      const tr = document.createElement('tr');
      const id = r.__id || '';
      tr.dataset.rowId = id;
      tr.dataset.rowIdx = idx;

      visibleCols.forEach((col, colIdx) => {
        const td = document.createElement('td');
        td.dataset.colKey = col.key;
        td.dataset.rowIdx = idx;
        td.dataset.colIdx = colIdx;

        const width = state.colWidths[col.key] || col.width || 120;
        td.style.width = `${width}px`;
        td.style.minWidth = `${width}px`;
        td.style.maxWidth = `${width}px`;

        const isFrozen = state.frozenCols.has(col.key);
        const isLastFrz = isTrailingFrozen(colIdx, visibleCols, state.frozenCols);
        const leftOffset = isFrozen ? getFrozenOffset(colIdx, visibleCols, state.frozenCols, state.colWidths) : 0;

        if (isFrozen) {
          td.className = `${isFrozen ? 'col-frozen' : ''} ${isLastFrz ? 'col-frozen-last' : ''}`.trim();
          td.style.left = `${leftOffset}px`;
        }

        const isDirty = id && state.stagedEdits[id] && state.stagedEdits[id][col.key] !== undefined;
        if (isDirty) td.classList.add('is-dirty');

        if (col.key === '__idx') {
          td.style.fontFamily = 'var(--mono)';
          td.style.color = 'var(--ink3)';
          td.style.textAlign = 'center';
          td.textContent = String(idx + 1);
        } else if (col.key === '__action') {
          td.className = `${td.className} col-action`.trim();
          td.innerHTML = `
            <button type="button" class="btn xs ghost danger rc-del-btn" onclick="window.deleteRecord('${id}')" title="Delete record from Dataverse">
              <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M4 7h16"></path>
                <path d="M9 7V4h6v3"></path>
                <path d="M6 7l1 13a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-13"></path>
                <path d="M10 11v6"></path>
                <path d="M14 11v6"></path>
              </svg>
            </button>
          `;
        } else {
          let val = String(getRowCellValue(r, col) ?? '');
          let input;
          if (col.isOptions && col.options && !col.readonly) {
            input = document.createElement('select');
            input.className = 'cell-input select';
            if (isDirty) input.classList.add('modified');
            let optHtml = `<option value="">-- Select --</option>`;
            col.options.forEach(opt => {
              const sel = opt.trim().toUpperCase() === val.trim().toUpperCase() ? 'selected' : '';
              optHtml += `<option value="${escapeHtml(opt)}" ${sel}>${escapeHtml(opt)}</option>`;
            });
            input.innerHTML = optHtml;
            input.value = val;
            input.addEventListener('change', () => {
              handleCellEdit(id, col.key, input.value, input, td);
            });
          } else {
            input = document.createElement('input');
            input.className = 'cell-input';
            if (col.readonly) input.classList.add('cell-readonly');
            if (col.isNum) input.classList.add('num-cell');
            if (col.isBalance) input.style.fontWeight = '700';
            if (isDirty) input.classList.add('modified');

            if (col.isPrice) {
              input.value = formatPrice(val);
            } else if (col.isPercent) {
              input.value = formatPercent(val);
            } else {
              input.value = val;
            }

            input.spellcheck = false;

            if (col.isDate) {
              input.placeholder = 'MM/DD/YYYY';
              input.maxLength = 10;
              input.addEventListener('input', () => {
                input.value = maskDateString(input.value);
              });
            }

            if (col.readonly) {
              input.readOnly = true;
              input.tabIndex = -1;
            } else {
              let before = input.value;
              input.addEventListener('focus', () => {
                before = input.value;
                if (col.isPrice) {
                  const pNum = parseAmount(input.value);
                  input.value = pNum ? String(pNum) : '';
                }
              });
              input.addEventListener('blur', () => {
                if (col.isPrice) {
                  const num = parseAmount(input.value);
                  input.value = num ? formatPrice(num) : '';
                  const cleanStr = num ? String(num) : '';
                  const prevNum = parseAmount(before);
                  const prevStr = prevNum ? String(prevNum) : '';
                  if (cleanStr !== prevStr) {
                    handleCellEdit(id, col.key, cleanStr, input, td);
                  }
                } else if (col.isDate) {
                  const masked = maskDateString(input.value);
                  input.value = masked;
                  if (masked !== before) {
                    handleCellEdit(id, col.key, masked, input, td);
                  }
                } else {
                  if (input.value !== before) {
                    handleCellEdit(id, col.key, input.value, input, td);
                  }
                }
              });
              input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  input.blur();
                } else if (e.key === 'Escape') {
                  input.value = before;
                  input.blur();
                }
              });
            }
          }
          td.appendChild(input);
        }
        tr.appendChild(td);
      });

      frag.appendChild(tr);
    }

    dom.recordsTableBody.appendChild(frag);
    state.renderedRecordsCount = end;
    invalidateGridCache('recordsTable');
  }

  async function handleCellEdit(rowId, colKey, newVal, inputEl, tdEl) {
    const rec = state.records.find(r => r.__id === rowId);
    if (!rec) return;

    const colDef = DEFAULT_COLUMNS.find(c => c.key === colKey) || { key: colKey };
    const origVal = String(rec[colDef.key] || (colDef.alt ? rec[colDef.alt] : '') || '').trim();
    let trimmed = String(newVal || '').trim();
    if (colDef.isPercent) trimmed = formatPercent(trimmed);

    if (origVal === trimmed && (!state.stagedEdits[rowId] || state.stagedEdits[rowId][colKey] === undefined)) {
      if (tdEl) tdEl.classList.remove('is-dirty');
      if (inputEl) inputEl.classList.remove('modified');
      return;
    }

    state.stagedEdits[rowId] = state.stagedEdits[rowId] || {};
    state.stagedEdits[rowId][colKey] = trimmed;
    if (colDef.alt) state.stagedEdits[rowId][colDef.alt] = trimmed;

    if (tdEl) tdEl.classList.add('is-dirty');
    if (inputEl) inputEl.classList.add('modified');
    updateStagedEditsUI();

    // Live formula calculation trigger
    const formulaKeys = ['cr63f_type', 'cr63f_indexno', 'cr63f_areaindex', 'cr63f_titleno', 'Type', 'Index No', 'Area-Index', 'Title No'];
    if (formulaKeys.includes(colKey)) {
      const curType = state.stagedEdits[rowId]['cr63f_type'] || rec['cr63f_type'] || rec['Type'] || '';
      const curIdx = state.stagedEdits[rowId]['cr63f_indexno'] || rec['cr63f_indexno'] || rec['Index No'] || '';
      const curArea = state.stagedEdits[rowId]['cr63f_areaindex'] || rec['cr63f_areaindex'] || rec['Area-Index'] || '';
      const curTitle = state.stagedEdits[rowId]['cr63f_titleno'] || rec['cr63f_titleno'] || rec['Title No'] || '';

      try {
        const lk = await window.pywebview.api.lookup_finance(curType, curIdx, curArea, curTitle);
        if (lk.matched) {
          const tcpFmt = lk.total_contract_price_formatted || '0.00';
          const dedFmt = lk.total_payment_deductible_formatted || '0.00';
          const balFmt = lk.balance_formatted || '0.00';
          const formattedPct = formatPercent(lk.percent_paid_formatted || lk.percent_paid || '0.00%');

          state.stagedEdits[rowId]['cr63f_totalcontractprice'] = tcpFmt;
          state.stagedEdits[rowId]['Total Contract Price'] = tcpFmt;

          state.stagedEdits[rowId]['cr63f_totalpaymentdeductibletotcp'] = dedFmt;
          state.stagedEdits[rowId]['Total Payment Deductible to TCP'] = dedFmt;

          state.stagedEdits[rowId]['cr63f_balance'] = balFmt;
          state.stagedEdits[rowId]['Balance'] = balFmt;

          state.stagedEdits[rowId]['cr63f_paid'] = formattedPct;
          state.stagedEdits[rowId]['% Paid'] = formattedPct;

          if (tdEl) {
            const tr = tdEl.closest('tr');
            if (tr) {
              const tcpInput = tr.querySelector('[data-col-key="cr63f_totalcontractprice"] input');
              if (tcpInput) { tcpInput.value = formatPrice(tcpFmt); tcpInput.classList.add('modified'); }
              const dedInput = tr.querySelector('[data-col-key="cr63f_totalpaymentdeductibletotcp"] input');
              if (dedInput) { dedInput.value = formatPrice(dedFmt); dedInput.classList.add('modified'); }
              const balInput = tr.querySelector('[data-col-key="cr63f_balance"] input');
              if (balInput) { balInput.value = formatPrice(balFmt); balInput.classList.add('modified'); }
              const pctInput = tr.querySelector('[data-col-key="cr63f_paid"] input');
              if (pctInput) { pctInput.value = formattedPct; pctInput.classList.add('modified'); }
            }
          }
        }
      } catch (e) {
        console.warn('Inline lookup calculation error:', e);
      }
    }

    // Direct manual edits to payment fields -> live calculation of balance & % paid (never negative balance, overpayment in %)
    const paymentKeys = ['cr63f_totalcontractprice', 'Total Contract Price', 'cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP'];
    if (paymentKeys.includes(colKey)) {
      const tcpRaw = state.stagedEdits[rowId]['cr63f_totalcontractprice'] ?? rec['cr63f_totalcontractprice'] ?? rec['Total Contract Price'] ?? '';
      const dedRaw = state.stagedEdits[rowId]['cr63f_totalpaymentdeductibletotcp'] ?? rec['cr63f_totalpaymentdeductibletotcp'] ?? rec['Total Payment Deductible to TCP'] ?? '';
      const tcpNum = parseAmount(tcpRaw);
      const dedNum = parseAmount(dedRaw);
      const balNum = Math.max(0, tcpNum - dedNum);
      const balFmt = balNum.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      let pctFmt = '0.00%';
      if (tcpNum > 0) {
        pctFmt = ((dedNum / tcpNum) * 100).toFixed(2) + '%';
      } else if (dedNum > 0) {
        pctFmt = '100.00%';
      }

      state.stagedEdits[rowId]['cr63f_balance'] = balFmt;
      state.stagedEdits[rowId]['Balance'] = balFmt;
      state.stagedEdits[rowId]['cr63f_paid'] = pctFmt;
      state.stagedEdits[rowId]['% Paid'] = pctFmt;

      if (tdEl) {
        const tr = tdEl.closest('tr');
        if (tr) {
          const balInput = tr.querySelector('[data-col-key="cr63f_balance"] input');
          if (balInput) { balInput.value = formatPrice(balFmt); balInput.classList.add('modified'); }
          const pctInput = tr.querySelector('[data-col-key="cr63f_paid"] input');
          if (pctInput) { pctInput.value = pctFmt; pctInput.classList.add('modified'); }
        }
      }
    }
  }

  function updateStagedEditsUI() {
    const rowCount = Object.keys(state.stagedEdits).length;
    if (rowCount > 0) {
      dom.stagedEditsWrap.classList.remove('hidden');
      dom.unsavedEditsChip.textContent = `${rowCount} unsaved row${rowCount > 1 ? 's' : ''}`;
    } else {
      dom.stagedEditsWrap.classList.add('hidden');
    }
  }

  async function saveStagedEdits() {
    const rowIds = Object.keys(state.stagedEdits);
    if (!rowIds.length) return;

    const stagedList = rowIds.map(id => ({
      id: id,
      fields: state.stagedEdits[id]
    }));

    try {
      dom.btnSaveStagedEdits.disabled = true;
      dom.btnSaveStagedEdits.textContent = 'Saving...';

      const res = await window.pywebview.api.save_staged_records(stagedList);
      if (res.success) {
        showToast(res.message || `Saved ${res.saved_count} records to Dataverse.`, 'ok');
        stagedList.forEach(item => {
          const rec = state.records.find(r => r.__id === item.id);
          if (rec) {
            Object.assign(rec, item.fields);
          }
        });
        state.stagedEdits = {};
        updateStagedEditsUI();
        filterAndRenderRecords();
      } else {
        showToast(res.error || 'Failed to save changes.', 'err');
      }
    } catch (err) {
      showToast('Save error: ' + err, 'err');
    } finally {
      dom.btnSaveStagedEdits.disabled = false;
      dom.btnSaveStagedEdits.textContent = 'Save to Dataverse';
    }
  }

  function discardStagedEdits() {
    state.stagedEdits = {};
    updateStagedEditsUI();
    filterAndRenderRecords();
    showToast('Unsaved edits discarded.', 'ok');
  }

  window.deleteRecord = async function (id) {
    if (!confirm('Are you sure you want to delete this record from Dataverse?')) return;
    try {
      const res = await window.pywebview.api.delete_loa_record_entry(id);
      if (res.success) {
        showToast('Record deleted.', 'ok');
        state.records = state.records.filter(r => r.__id !== id);
        delete state.stagedEdits[id];
        updateStagedEditsUI();
        filterAndRenderRecords();
      } else {
        showToast(res.error || 'Failed to delete record.', 'err');
      }
    } catch (err) {
      showToast('Error: ' + err, 'err');
    }
  };

  async function exportRecordsExcel() {
    if (!state.records.length) {
      showToast('No records to export.', 'err');
      return;
    }
    const toExport = state.filteredRecords.length ? state.filteredRecords : state.records;
    const cols = (state.colOrder.length ? state.colOrder : DEFAULT_COLUMNS.map(c => c.key))
      .filter(k => state.visibleCols.has(k) && !k.startsWith('__') && !k.startsWith('_'))
      .map(k => {
        const c = DEFAULT_COLUMNS.find(dc => dc.key === k);
        return (c && c.label) ? c.label : k;
      });
    const res = await window.pywebview.api.export_records_excel(toExport, 'LOA_Monitoring_Export.xlsx', cols);
    if (res.success) {
      showToast(`Exported to ${res.file_path}`, 'ok');
    }
  }

  function formatDateForInput(dateVal) {
    if (!dateVal) return '';
    const s = String(dateVal).split('T')[0].split(' ')[0].trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const parts = s.split(/[-/]/);
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
      } else {
        return `${parts[2]}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`;
      }
    }
    return '';
  }

  // ==================== DRAWER & LOOKUP ENGINE ====================
  function setupDrawerForm() {
    if (dom.btnOpenNewEntry) {
      dom.btnOpenNewEntry.addEventListener('click', openNewEntry);
    }
    if (dom.btnCloseDrawer) {
      dom.btnCloseDrawer.addEventListener('click', closeDrawer);
    }
    if (dom.btnCancelDrawer) {
      dom.btnCancelDrawer.addEventListener('click', closeDrawer);
    }
    if (dom.btnSaveDrawerEntry) {
      dom.btnSaveDrawerEntry.addEventListener('click', handleSaveEntry);
    }

    if (dom.entryDrawer) {
      dom.entryDrawer.addEventListener('click', (e) => {
        if (e.target === dom.entryDrawer) closeDrawer();
      });
    }

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && dom.entryDrawer && dom.entryDrawer.classList.contains('open')) {
        closeDrawer();
      }
    });

    setupAreaIndexSuggestions();

    const lookupTriggers = [dom.fType, dom.fIndexNo, dom.fAreaIndex, dom.fTitleNo];
    lookupTriggers.forEach(input => {
      if (input) {
        input.addEventListener('input', triggerLiveLookup);
        input.addEventListener('change', triggerLiveLookup);
      }
    });

    [dom.fSecuredDate, dom.fSecuredEndorseDate, dom.fEndorsedDate].forEach(input => {
      if (input) {
        input.addEventListener('input', () => {
          input.value = maskDateString(input.value);
        });
      }
    });

    if (dom.fTcp) dom.fTcp.addEventListener('input', updatePaymentMath);
    if (dom.fPaymentDeductible) dom.fPaymentDeductible.addEventListener('input', updatePaymentMath);
  }

  function updatePaymentMath() {
    const tcpStr = dom.fTcp ? dom.fTcp.value.trim() : '';
    const dedStr = dom.fPaymentDeductible ? dom.fPaymentDeductible.value.trim() : '';
    if (tcpStr !== '' || dedStr !== '') {
      const tcp = parseAmount(tcpStr);
      const ded = parseAmount(dedStr);
      const bal = Math.max(0, tcp - ded);
      if (dom.fBalance && document.activeElement !== dom.fBalance) {
        dom.fBalance.value = bal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      }
      if (dom.fPercentPaid && document.activeElement !== dom.fPercentPaid) {
        if (tcp > 0) {
          const pct = (ded / tcp) * 100;
          dom.fPercentPaid.value = pct.toFixed(2) + '%';
        } else if (ded > 0) {
          dom.fPercentPaid.value = '100.00%';
        } else {
          dom.fPercentPaid.value = '0.00%';
        }
      }
    }
  }

  function setupAreaIndexSuggestions() {
    if (!dom.fAreaIndex || !dom.areaIndexSuggestions) return;

    function renderAreaSuggestions(query) {
      const q = (query || '').toLowerCase().trim();
      const selectedType = (dom.fType ? dom.fType.value : '').trim().toUpperCase();
      const areaMap = new Map();

      state.records.forEach(r => {
        const rType = String(r['cr63f_type'] || r['Type'] || '').trim().toUpperCase();
        if (selectedType && rType && rType !== selectedType) return;

        const areaVal = String(r['cr63f_areaindex'] || r['Area-Index'] || r['AREA-INDEX'] || '').trim();
        if (areaVal) {
          const upper = areaVal.toUpperCase();
          if (!areaMap.has(upper)) {
            areaMap.set(upper, { key: areaVal, rec: r });
          }
        }
      });

      let items = Array.from(areaMap.values());
      if (q) {
        items = items.filter(item => item.key.toLowerCase().includes(q));
      }

      if (!items.length) {
        dom.areaIndexSuggestions.innerHTML = '<div style="padding: 10px; font-size: 12px; color: var(--ink3); font-style: italic;">No matching Area-Index records found for ' + escapeHtml(selectedType || 'selected type') + '</div>';
        dom.areaIndexSuggestions.classList.remove('hidden');
        return;
      }

      items = items.slice(0, 40);

      let html = '';
      items.forEach(item => {
        const r = item.rec;
        const owner = r['cr63f_registeredowner'] || r['Registered Owner'] || '';
        const mun = r['cr63f_municipality'] || r['Municipality'] || '';
        const prov = r['cr63f_province'] || r['Province'] || '';
        const loc = [mun, prov].filter(Boolean).join(', ');
        const sub = [owner, loc].filter(Boolean).join(' • ');

        html += `
          <div class="area-opt" data-area-key="${escapeHtml(item.key)}">
            <div class="area-opt-main">
              <span>${escapeHtml(item.key)}</span>
              <span style="font-size: 10px; font-weight: normal; color: var(--ink3);">${escapeHtml(r['cr63f_type'] || r['Type'] || '')}</span>
            </div>
            ${sub ? `<div class="area-opt-sub">${escapeHtml(sub)}</div>` : ''}
          </div>
        `;
      });

      dom.areaIndexSuggestions.innerHTML = html;
      dom.areaIndexSuggestions.classList.remove('hidden');

      dom.areaIndexSuggestions.querySelectorAll('.area-opt').forEach(opt => {
        opt.addEventListener('click', (e) => {
          e.stopPropagation();
          const k = opt.dataset.areaKey;
          const match = areaMap.get(k.toUpperCase());
          if (match) {
            populateFormFromRecord(match.rec, k);
          }
          dom.areaIndexSuggestions.classList.add('hidden');
        });
      });
    }

    dom.fAreaIndex.addEventListener('input', () => {
      renderAreaSuggestions(dom.fAreaIndex.value);
    });

    dom.fAreaIndex.addEventListener('focus', () => {
      renderAreaSuggestions(dom.fAreaIndex.value);
    });

    if (dom.fType) {
      dom.fType.addEventListener('change', () => {
        if (!dom.areaIndexSuggestions.classList.contains('hidden')) {
          renderAreaSuggestions(dom.fAreaIndex.value);
        }
      });
    }

    document.addEventListener('click', (e) => {
      if (!e.target.closest('#fAreaIndex') && !e.target.closest('#areaIndexSuggestions')) {
        dom.areaIndexSuggestions.classList.add('hidden');
      }
    });
  }

  function populateFormFromRecord(r, areaVal) {
    if (!r) return;
    if (dom.fAreaIndex) dom.fAreaIndex.value = areaVal || r['cr63f_areaindex'] || r['Area-Index'] || '';
    if (dom.fType) dom.fType.value = r['cr63f_type'] || r['Type'] || 'LAND';
    if (dom.fIndexNo) dom.fIndexNo.value = r['cr63f_indexno'] || r['Index No'] || '';
    if (dom.fTitleNo) dom.fTitleNo.value = r['cr63f_titleno'] || r['Title No'] || '';
    if (dom.fControl) dom.fControl.value = r['cr63f_control1'] || r['Control'] || '';

    const tcpVal = r['cr63f_totalcontractprice'] ?? r['Total Contract Price'] ?? '';
    const dedVal = r['cr63f_totalpaymentdeductibletotcp'] ?? r['Total Payment Deductible to TCP'] ?? '';
    const balVal = r['cr63f_balance'] ?? r['Balance'] ?? '';
    const pctVal = r['cr63f_paid'] ?? r['% Paid'] ?? '';

    if (dom.fTcp) dom.fTcp.value = tcpVal !== '' ? formatAmount(tcpVal) : '';
    if (dom.fPaymentDeductible) dom.fPaymentDeductible.value = dedVal !== '' ? formatAmount(dedVal) : '';
    if (dom.fBalance) dom.fBalance.value = balVal !== '' ? formatAmount(balVal) : '';
    if (dom.fPercentPaid) dom.fPercentPaid.value = pctVal !== '' ? formatPercent(pctVal) : '';

    if (tcpVal !== '' || dedVal !== '') {
      updatePaymentMath();
    }

    if (dom.fRegisteredOwner) dom.fRegisteredOwner.value = r['cr63f_registeredowner'] || r['Registered Owner'] || '';
    if (dom.fAuthorizedRepresentative) dom.fAuthorizedRepresentative.value = r['cr63f_authorizedrepresentative'] || r['Authorized Representative'] || '';
    if (dom.fEntity) dom.fEntity.value = r['cr63f_entity'] || r['Entity'] || '';

    if (dom.fProvince) dom.fProvince.value = r['cr63f_province'] || r['Province'] || '';
    if (dom.fMunicipality) dom.fMunicipality.value = r['cr63f_municipality'] || r['Municipality'] || '';
    if (dom.fDbAdjustedBarangay) dom.fDbAdjustedBarangay.value = r['cr63f_dbadjustedbarangay'] || r['DB Adjusted Barangay'] || '';
    if (dom.fLotAreaHa) dom.fLotAreaHa.value = r['cr63f_lotareaha'] || r['Lot Area (Ha)'] || '';

    if (dom.fContractType) {
      dom.fContractType.value = r['new_contracttype'] || r['Contract Type'] || '';
    }
    if (dom.fMode) {
      dom.fMode.value = r['new_mode'] || r['Mode'] || '';
    }
    if (dom.fConversionStatus) dom.fConversionStatus.value = r['cr63f_conversionstatus'] || r['Conversion Status'] || '';
    if (dom.fRccTerms) dom.fRccTerms.value = r['cr63f_rccterms'] || r['RCC Terms'] || '';
    
    const rccVal = r['cr63f_rccprice'] ?? r['RCC Price'] ?? '';
    if (dom.fRccPrice) dom.fRccPrice.value = rccVal !== '' ? formatAmount(rccVal) : '';

    if (dom.fSecuredDate) dom.fSecuredDate.value = formatDateForInput(r['cr63f_secureddate'] || r['Secured Date']);
    if (dom.fSecuredEndorseDate) dom.fSecuredEndorseDate.value = formatDateForInput(r['cr63f_securedendorsedate'] || r['Secured Endorse Date']);
    if (dom.fEndorsedDate) dom.fEndorsedDate.value = formatDateForInput(r['cr63f_endorseddate'] || r['Endorsed Date']);

    triggerLiveLookup(false);
  }

  function openDrawer() {
    if (!dom.entryDrawer) return;
    dom.entryDrawer.classList.add('open');
    dom.entryDrawer.style.display = 'flex';
    setTimeout(() => {
      if (dom.fAreaIndex) dom.fAreaIndex.focus();
    }, 50);
  }

  function closeDrawer() {
    if (!dom.entryDrawer) return;
    dom.entryDrawer.classList.remove('open');
    dom.entryDrawer.style.display = 'none';
  }

  function openNewEntry() {
    if (!dom.drawerTitle) return;
    dom.drawerTitle.textContent = 'New LOA Entry / Log';
    dom.fRecordId.value = '';

    if (dom.areaIndexSuggestions) dom.areaIndexSuggestions.classList.add('hidden');

    if (dom.fType) dom.fType.value = 'LAND';
    if (dom.fIndexNo) dom.fIndexNo.value = '';
    if (dom.fAreaIndex) dom.fAreaIndex.value = '';
    if (dom.fTitleNo) dom.fTitleNo.value = '';
    if (dom.fControl) dom.fControl.value = '';

    clearAutofillFields();

    if (dom.fRegisteredOwner) dom.fRegisteredOwner.value = '';
    if (dom.fAuthorizedRepresentative) dom.fAuthorizedRepresentative.value = '';
    if (dom.fEntity) dom.fEntity.value = '';
    if (dom.fProvince) dom.fProvince.value = '';
    if (dom.fMunicipality) dom.fMunicipality.value = '';
    if (dom.fDbAdjustedBarangay) dom.fDbAdjustedBarangay.value = '';
    if (dom.fLotAreaHa) dom.fLotAreaHa.value = '';
    if (dom.fContractType) dom.fContractType.value = '';
    if (dom.fMode) dom.fMode.value = '';
    if (dom.fConversionStatus) dom.fConversionStatus.value = '';
    if (dom.fRccTerms) dom.fRccTerms.value = '';
    if (dom.fRccPrice) dom.fRccPrice.value = '';
    if (dom.fSecuredDate) dom.fSecuredDate.value = '';
    if (dom.fSecuredEndorseDate) dom.fSecuredEndorseDate.value = '';
    if (dom.fEndorsedDate) dom.fEndorsedDate.value = '';

    openDrawer();
  }

  window.openNewEntry = openNewEntry;
  window.closeDrawer = closeDrawer;

  function clearAutofillFields() {
    if (dom.fTcp) dom.fTcp.value = '';
    if (dom.fPaymentDeductible) dom.fPaymentDeductible.value = '';
    if (dom.fBalance) dom.fBalance.value = '';
    if (dom.fPercentPaid) dom.fPercentPaid.value = '';
  }

  function triggerLiveLookup(clearOnNoMatch = true) {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(async () => {
      const typeVal = dom.fType ? dom.fType.value.trim() : '';
      const indexNo = dom.fIndexNo ? dom.fIndexNo.value.trim() : '';
      const areaIndex = dom.fAreaIndex ? dom.fAreaIndex.value.trim() : '';
      const titleNo = dom.fTitleNo ? dom.fTitleNo.value.trim() : '';

      if (!typeVal && !indexNo && !areaIndex && !titleNo) {
        if (clearOnNoMatch) clearAutofillFields();
        if (dom.badgeMatchStatus) {
          dom.badgeMatchStatus.className = 'match-badge waiting';
          dom.badgeMatchStatus.textContent = 'Waiting for input...';
        }
        return;
      }

      try {
        const res = await window.pywebview.api.lookup_finance(typeVal, indexNo, areaIndex, titleNo);
        if (res.matched) {
          if (dom.fTcp) dom.fTcp.value = res.total_contract_price_formatted || '0.00';
          if (dom.fPaymentDeductible) dom.fPaymentDeductible.value = res.total_payment_deductible_formatted || '0.00';
          if (dom.fBalance) dom.fBalance.value = res.balance_formatted || '0.00';
          if (dom.fPercentPaid) dom.fPercentPaid.value = formatPercent(res.percent_paid_formatted || res.percent_paid || '0.00%');
          updatePaymentMath();

          if (dom.badgeMatchStatus) {
            dom.badgeMatchStatus.className = 'match-badge matched';
            dom.badgeMatchStatus.textContent = `✓ ${res.match_type} [${res.match_key}]`;
          }
        } else {
          if (clearOnNoMatch) clearAutofillFields();
          if (dom.badgeMatchStatus) {
            dom.badgeMatchStatus.className = 'match-badge nomatch';
            dom.badgeMatchStatus.textContent = '⚠ No match in Appended Finance';
          }
        }
      } catch (err) {
        console.error('Lookup error:', err);
      }
    }, 150);
  }

  async function handleSaveEntry() {
    const recordId = dom.fRecordId.value || null;

    const recordData = {
      'Type': dom.fType.value,
      'cr63f_type': dom.fType.value,
      'Index No': dom.fIndexNo.value.trim(),
      'cr63f_indexno': dom.fIndexNo.value.trim(),
      'Area-Index': dom.fAreaIndex.value.trim(),
      'cr63f_areaindex': dom.fAreaIndex.value.trim(),
      'Title No': dom.fTitleNo.value.trim(),
      'cr63f_titleno': dom.fTitleNo.value.trim(),
      'Control': dom.fControl.value.trim(),
      'cr63f_control1': dom.fControl.value.trim(),

      'Total Contract Price': dom.fTcp.value.trim().replace(/[₱$]/g, ''),
      'cr63f_totalcontractprice': dom.fTcp.value.trim().replace(/[₱$]/g, ''),
      'Total Payment Deductible to TCP': dom.fPaymentDeductible.value.trim().replace(/[₱$]/g, ''),
      'cr63f_totalpaymentdeductibletotcp': dom.fPaymentDeductible.value.trim().replace(/[₱$]/g, ''),
      'Balance': dom.fBalance.value.trim().replace(/[₱$]/g, ''),
      'cr63f_balance': dom.fBalance.value.trim().replace(/[₱$]/g, ''),
      '% Paid': formatPercent(dom.fPercentPaid.value.trim()),
      'cr63f_paid': formatPercent(dom.fPercentPaid.value.trim()),

      'Registered Owner': dom.fRegisteredOwner.value.trim(),
      'cr63f_registeredowner': dom.fRegisteredOwner.value.trim(),
      'Authorized Representative': dom.fAuthorizedRepresentative.value.trim(),
      'cr63f_authorizedrepresentative': dom.fAuthorizedRepresentative.value.trim(),
      'Entity': dom.fEntity.value.trim(),
      'cr63f_entity': dom.fEntity.value.trim(),

      'Province': dom.fProvince.value.trim(),
      'cr63f_province': dom.fProvince.value.trim(),
      'Municipality': dom.fMunicipality.value.trim(),
      'cr63f_municipality': dom.fMunicipality.value.trim(),
      'DB Adjusted Barangay': dom.fDbAdjustedBarangay.value.trim(),
      'cr63f_dbadjustedbarangay': dom.fDbAdjustedBarangay.value.trim(),
      'Lot Area (Ha)': dom.fLotAreaHa.value.trim(),
      'cr63f_lotareaha': dom.fLotAreaHa.value.trim(),

      'Contract Type': dom.fContractType.value.trim(),
      'new_contracttype': dom.fContractType.value.trim(),
      'Mode': dom.fMode.value.trim(),
      'new_mode': dom.fMode.value.trim(),
      'Conversion Status': dom.fConversionStatus.value.trim(),
      'cr63f_conversionstatus': dom.fConversionStatus.value.trim(),
      'RCC Terms': dom.fRccTerms.value.trim(),
      'cr63f_rccterms': dom.fRccTerms.value.trim(),
      'RCC Price': dom.fRccPrice.value.trim().replace(/[₱$]/g, ''),
      'cr63f_rccprice': dom.fRccPrice.value.trim().replace(/[₱$]/g, ''),

      'Secured Date': dom.fSecuredDate.value,
      'cr63f_secureddate': dom.fSecuredDate.value,
      'Secured Endorse Date': dom.fSecuredEndorseDate.value,
      'cr63f_securedendorsedate': dom.fSecuredEndorseDate.value,
      'Endorsed Date': dom.fEndorsedDate.value,
      'cr63f_endorseddate': dom.fEndorsedDate.value,
    };

    if (!recordData['Type']) {
      showToast('TYPE is required.', 'err');
      return;
    }

    try {
      dom.btnSaveDrawerEntry.disabled = true;
      dom.btnSaveDrawerEntry.textContent = 'Saving to Dataverse...';

      const res = await window.pywebview.api.save_loa_record_entry(recordData, recordId);
      if (res.success) {
        showToast(res.message || 'Record successfully saved.', 'ok');
        closeDrawer();
        fetchRecords();
      } else {
        showToast(res.error || 'Failed to save record.', 'err');
      }
    } catch (err) {
      showToast('Error saving: ' + err, 'err');
    } finally {
      dom.btnSaveDrawerEntry.disabled = false;
      dom.btnSaveDrawerEntry.textContent = 'Save Entry to Dataverse';
    }
  }

  // ==================== FINANCE DATA VIEW ====================
  function setupFinanceView() {
    dom.btnUploadFinance.addEventListener('click', async () => {
      const res = await window.pywebview.api.load_finance_file_dialog();
      if (res.success) {
        showToast(`Loaded ${res.row_count} rows from ${res.source}.`, 'ok');
        const s = await window.pywebview.api.get_initial_state();
        applyState(s);
        renderFinancePreview(res.headers, res.sample_rows);
      } else if (!res.cancelled) {
        showToast(res.error || 'Failed to load file.', 'err');
      }
    });

    dom.btnClearFinanceData.addEventListener('click', async () => {
      await window.pywebview.api.clear_finance_reference();
      const s = await window.pywebview.api.get_initial_state();
      applyState(s);
      financeState.headers = [];
      financeState.rows = [];
      financeState.filteredRows = [];
      financeState.renderedRowsCount = 0;
      dom.financePreviewThead.innerHTML = '';
      dom.financePreviewTbody.innerHTML = '';
      if (dom.financeEmptyState) dom.financeEmptyState.classList.remove('hidden');
      if (dom.lblFinanceGridCount) dom.lblFinanceGridCount.textContent = '0 rows';
      showToast('Finance data cleared.', 'ok');
    });

    if (dom.searchInputFinance) {
      dom.searchInputFinance.addEventListener('input', () => {
        clearTimeout(searchFinDebounce);
        searchFinDebounce = setTimeout(filterAndRenderFinance, 120);
      });
    }
    if (dom.typeFilterFinance) {
      dom.typeFilterFinance.addEventListener('change', filterAndRenderFinance);
    }
    if (dom.btnClearFiltersFinance) {
      dom.btnClearFiltersFinance.addEventListener('click', () => {
        if (dom.searchInputFinance) dom.searchInputFinance.value = '';
        if (dom.typeFilterFinance) dom.typeFilterFinance.value = 'ALL';
        financeState.columnFilters = {};
        financeState.columnFiltersOpen = {};
        filterAndRenderFinance();
        showToast('Finance filters cleared.', 'ok');
      });
    }
    if (dom.btnExportExcelFinance) {
      dom.btnExportExcelFinance.addEventListener('click', exportFinanceExcel);
    }

    setupFinanceColumnPicker();
  }

  function setupFinanceColumnPicker() {
    if (!dom.btnColumnsPickerFinance || !dom.colPickerMenuFinance) return;

    dom.btnColumnsPickerFinance.addEventListener('click', (e) => {
      e.stopPropagation();
      dom.colPickerMenuFinance.classList.toggle('hidden');
    });

    document.addEventListener('click', (e) => {
      if (!e.target.closest('#view-finance .col-picker-container')) {
        dom.colPickerMenuFinance.classList.add('hidden');
      }
    });
  }

  function renderFinanceColumnPickerMenu() {
    if (!dom.colPickerMenuFinance) return;

    const cols = financeState.colOrder.length ? financeState.colOrder : financeState.headers;
    let html = '';
    cols.forEach((colKey, idx) => {
      const isChecked = financeState.visibleCols.has(colKey);
      const isFirst = idx === 0;
      const isLast = idx === cols.length - 1;

      html += `
        <div class="col-picker-item">
          <label class="col-picker-label">
            <input type="checkbox" data-fin-col-toggle="${escapeHtml(colKey)}" ${isChecked ? 'checked' : ''} />
            <span>${escapeHtml(colKey)}</span>
          </label>
          <div class="col-reorder-btns">
            <button type="button" class="col-reorder-btn" data-fin-move-up="${escapeHtml(colKey)}" ${isFirst ? 'disabled' : ''} title="Move left">◄</button>
            <button type="button" class="col-reorder-btn" data-fin-move-down="${escapeHtml(colKey)}" ${isLast ? 'disabled' : ''} title="Move right">►</button>
          </div>
        </div>
      `;
    });

    html += `
      <div class="col-picker-actions">
        <button type="button" class="btn xs ghost" id="btnResetFinanceLayout" style="width: 100%; font-size: 11px;">Reset Columns Layout</button>
      </div>
    `;

    dom.colPickerMenuFinance.innerHTML = html;

    dom.colPickerMenuFinance.querySelectorAll('[data-fin-col-toggle]').forEach(chk => {
      chk.addEventListener('change', () => {
        const key = chk.dataset.finColToggle;
        if (chk.checked) {
          financeState.visibleCols.add(key);
        } else {
          if (financeState.visibleCols.size > 1) {
            financeState.visibleCols.delete(key);
          } else {
            chk.checked = true;
            showToast('At least one column must remain visible.', 'err');
            return;
          }
        }
        saveFinanceTableConfig();
        filterAndRenderFinance();
      });
    });

    dom.colPickerMenuFinance.querySelectorAll('[data-fin-move-up]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.finMoveUp;
        moveFinanceColumnOrder(key, -1);
      });
    });

    dom.colPickerMenuFinance.querySelectorAll('[data-fin-move-down]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.finMoveDown;
        moveFinanceColumnOrder(key, 1);
      });
    });

    const btnReset = dom.colPickerMenuFinance.querySelector('#btnResetFinanceLayout');
    if (btnReset) {
      btnReset.addEventListener('click', (e) => {
        e.stopPropagation();
        localStorage.removeItem('fin_col_widths');
        localStorage.removeItem('fin_col_order');
        localStorage.removeItem('fin_visible_cols');
        localStorage.removeItem('fin_frozen_cols');
        financeState.visibleCols = new Set(financeState.headers);
        financeState.colOrder = [...financeState.headers];
        financeState.colWidths = {};
        financeState.frozenCols = new Set();
        saveFinanceTableConfig();
        renderFinanceColumnPickerMenu();
        filterAndRenderFinance();
        showToast('Finance layout reset.', 'ok');
      });
    }
  }

  function moveFinanceColumnOrder(colKey, direction) {
    const idx = financeState.colOrder.indexOf(colKey);
    if (idx === -1) return;
    const targetIdx = idx + direction;
    if (targetIdx < 0 || targetIdx >= financeState.colOrder.length) return;

    const temp = financeState.colOrder[idx];
    financeState.colOrder[idx] = financeState.colOrder[targetIdx];
    financeState.colOrder[targetIdx] = temp;

    saveFinanceTableConfig();
    renderFinanceColumnPickerMenu();
    filterAndRenderFinance();
  }

  const SYSTEM_EXCLUDE_FIELDS = new Set([
    'createdby', 'createdon', 'createdonbehalfby', 'modifiedby', 'modifiedon',
    'modifiedonbehalfby', 'overriddencreatedon', 'versionnumber', 'importsequencenumber',
    'utcconversiontimezonecode', 'timezoneruleversionnumber', 'ownerid',
    'owningbusinessunit', 'owninguser', 'owningteam', 'statecode', 'statuscode',
    'exchangerate', 'transactioncurrencyid', 'statuscodename', 'statecodename',
    'createdbyname', 'modifiedbyname', 'owneridname', 'owningbusinessunitname',
    'owningusername', 'owningteamname', 'transactioncurrencyidname', 'createdbyyominame',
    'createdonbehalfbyname', 'createdonbehalfbyyominame', 'modifiedbyyominame',
    'modifiedonbehalfbyname', 'modifiedonbehalfbyyominame', 'owneridtype',
    'owneridyominame'
  ]);

  function isSystemOrGeneratedCol(col) {
    if (!col) return true;
    const s = String(col).trim();
    const low = s.toLowerCase();
    if (s === '#' || s === 'Actions' || s.startsWith('__') || s.startsWith('_') || s.startsWith('@')) return true;
    if (low.endsWith('_value') || low.endsWith('(base)') || low.endsWith('_base')) return true;
    if (SYSTEM_EXCLUDE_FIELDS.has(low)) return true;
    if (Array.from(SYSTEM_EXCLUDE_FIELDS).some(k => low === `_${k}_value` || low === `${k}_value`)) return true;
    if (low.endsWith('id') && low.length > 10 && (low.startsWith('cr63f_') || low.startsWith('new_') || SYSTEM_EXCLUDE_FIELDS.has(low))) return true;
    return false;
  }

  function renderFinancePreview(headers, rows) {
    if (!headers || !rows) return;
    const cleanHeaders = headers.filter(h => !isSystemOrGeneratedCol(h));

    financeState.headers = cleanHeaders;
    financeState.rows = rows;

    try {
      const fw = localStorage.getItem('fin_col_widths');
      if (fw) financeState.colWidths = JSON.parse(fw);
    } catch (_) {}

    try {
      const fo = localStorage.getItem('fin_col_order');
      if (fo) {
        const parsed = JSON.parse(fo);
        const valid = parsed.filter(k => cleanHeaders.includes(k) && !isSystemOrGeneratedCol(k));
        cleanHeaders.forEach(k => { if (!valid.includes(k)) valid.push(k); });
        financeState.colOrder = valid;
      } else {
        financeState.colOrder = [...cleanHeaders];
      }
    } catch (_) {
      financeState.colOrder = [...cleanHeaders];
    }

    try {
      const fv = localStorage.getItem('fin_visible_cols');
      if (fv) {
        const parsed = JSON.parse(fv).filter(k => cleanHeaders.includes(k) && !isSystemOrGeneratedCol(k));
        financeState.visibleCols = new Set(parsed);
        cleanHeaders.forEach(k => {
          if (!financeState.visibleCols.has(k)) financeState.visibleCols.add(k);
        });
      } else {
        financeState.visibleCols = new Set(cleanHeaders);
      }
    } catch (_) {
      financeState.visibleCols = new Set(cleanHeaders);
    }

    try {
      const ff = localStorage.getItem('fin_frozen_cols');
      financeState.frozenCols = ff ? new Set(JSON.parse(ff)) : new Set();
    } catch (_) {
      financeState.frozenCols = new Set();
    }

    // Populate type filter for finance if TYPE column exists
    if (dom.typeFilterFinance) {
      const typeCols = cleanHeaders.filter(h => h.toUpperCase().includes('TYPE'));
      const typeSet = new Set();
      rows.forEach(r => {
        typeCols.forEach(tc => {
          const val = String(r[tc] || '').trim().toUpperCase();
          if (val) typeSet.add(val);
        });
      });
      if (typeSet.size > 0) {
        let optHtml = '<option value="ALL">All Types</option>';
        Array.from(typeSet).sort().forEach(t => {
          optHtml += `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`;
        });
        dom.typeFilterFinance.innerHTML = optHtml;
      } else {
        dom.typeFilterFinance.innerHTML = '<option value="ALL">All Types</option>';
      }
    }

    renderFinanceColumnPickerMenu();
    filterAndRenderFinance();
  }

  function filterAndRenderFinance() {
    const q = dom.searchInputFinance ? dom.searchInputFinance.value.toLowerCase().trim() : '';
    const typeFilt = dom.typeFilterFinance ? dom.typeFilterFinance.value : 'ALL';
    const colFilterKeys = Object.keys(financeState.columnFilters).filter(k => !!financeState.columnFilters[k]);

    financeState.filteredRows = financeState.rows.filter(r => {
      if (typeFilt !== 'ALL') {
        const hasTypeMatch = Object.entries(r).some(([k, v]) => k.toUpperCase().includes('TYPE') && String(v).trim().toUpperCase() === typeFilt);
        if (!hasTypeMatch) return false;
      }

      for (const colKey of colFilterKeys) {
        const query = financeState.columnFilters[colKey];
        const val = String(r[colKey] ?? '').toLowerCase();
        if (!val.includes(query)) return false;
      }

      if (!q) return true;
      return Object.values(r).some(v => String(v).toLowerCase().includes(q));
    });

    if (financeState.sortCol && financeState.sortDir) {
      const col = financeState.sortCol;
      const mult = financeState.sortDir === 'asc' ? 1 : -1;
      financeState.filteredRows.sort((a, b) => {
        let valA = a[col] ?? '';
        let valB = b[col] ?? '';
        const numA = parseFloat(String(valA).replace(/[^0-9.-]+/g, ''));
        const numB = parseFloat(String(valB).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(numA) && !isNaN(numB)) return (numA - numB) * mult;
        return String(valA).localeCompare(String(valB)) * mult;
      });
    }

    if (dom.lblFinanceGridCount) {
      dom.lblFinanceGridCount.textContent = `${financeState.filteredRows.length.toLocaleString()} rows`;
    }

    renderFinanceTable();
  }

  function getFinanceVisibleColumns() {
    const order = financeState.colOrder.length ? financeState.colOrder : financeState.headers;
    return order.filter(h => financeState.visibleCols.has(h) && !isSystemOrGeneratedCol(h));
  }

  function renderFinanceTable() {
    if (!dom.financePreviewThead || !dom.financePreviewTbody) return;
    const cols = getFinanceVisibleColumns();

    if (!financeState.rows.length) {
      dom.financePreviewThead.innerHTML = '';
      dom.financePreviewTbody.innerHTML = '';
      if (dom.financeEmptyState) dom.financeEmptyState.classList.remove('hidden');
      return;
    }

    if (dom.financeEmptyState) dom.financeEmptyState.classList.add('hidden');

    let thead = '<tr>';
    thead += `<th style="width: 45px; min-width: 45px; max-width: 45px; text-align: center; position: sticky !important; top: 0 !important; z-index: 25 !important; background: var(--subtle) !important;">#</th>`;

    cols.forEach((h, colIdx) => {
      const w = financeState.colWidths[h] || 140;
      const isSorted = financeState.sortCol === h;
      const sortIcon = isSorted ? (financeState.sortDir === 'asc' ? '▲' : '▼') : '';
      const isFrozen = financeState.frozenCols.has(h);
      const isLastFrz = isTrailingFrozen(colIdx, cols, financeState.frozenCols);
      const leftOffset = isFrozen ? (45 + getFrozenOffset(colIdx, cols, financeState.frozenCols, financeState.colWidths)) : 0;
      const frozenClass = `${isFrozen ? 'col-frozen' : ''} ${isLastFrz ? 'col-frozen-last' : ''}`.trim();
      const leftStyle = isFrozen ? `left: ${leftOffset}px;` : '';
      const isFilterActive = !!(financeState.columnFilters[h]);
      const isFilterOpen = !!(financeState.columnFiltersOpen[h]);

      thead += `
        <th draggable="true" data-fin-col="${escapeHtml(h)}" class="${frozenClass}" style="width:${w}px; min-width:${w}px; max-width:${w}px; ${leftStyle}">
          <div class="th-header-box">
            <div class="th-title-wrap">
              <button type="button" class="th-pin-btn ${isFrozen ? 'pinned' : ''}" data-fin-pin="${escapeHtml(h)}" title="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(h)}">
                <svg viewBox="0 0 24 24" width="11" height="11" fill="currentColor" stroke="none">
                  <path d="M16 12V4h1V2H7v2h1v8l-2 2v2h5.2v6h1.6v-6H18v-2l-2-2z"/>
                </svg>
              </button>
              <div class="th-sort-wrap" data-fin-sort="${escapeHtml(h)}" title="Click to sort, drag to reorder">
                <span class="th-col-name">${escapeHtml(h)}</span>
                <span class="th-sort-icon ${isSorted ? 'active' : ''}">${sortIcon}</span>
              </div>
              <button type="button" class="th-chevron-btn ${isFilterActive ? 'active' : ''}" data-fin-filter-toggle="${escapeHtml(h)}" title="Filter ${escapeHtml(h)}">
                <svg viewBox="0 0 24 24" width="10" height="10" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"></polyline></svg>
              </button>
            </div>
            <div class="th-filter-popup ${isFilterOpen ? '' : 'hidden'}" id="finColFilterBox_${escapeHtml(h)}">
              <input class="th-filter-input" data-fin-filter-col="${escapeHtml(h)}" value="${escapeHtml(financeState.columnFilters[h] || '')}" placeholder="Filter ${escapeHtml(h)}..." />
            </div>
          </div>
          <div class="th-resizer" data-fin-resize="${escapeHtml(h)}" title="Drag column width"></div>
          <div class="th-bottom-resizer" title="Drag to adjust header height"></div>
        </th>
      `;
    });
    thead += `<th class="col-action" style="width: 55px; min-width: 55px; max-width: 55px; text-align: center; position: sticky !important; top: 0 !important; z-index: 20 !important; background: var(--subtle) !important;">Actions</th>`;
    thead += '</tr>';
    dom.financePreviewThead.innerHTML = thead;

    attachFinanceHeaderEvents();

    dom.financePreviewTbody.innerHTML = '';
    financeState.renderedRowsCount = 0;
    renderFinanceChunk(0, CHUNK_SIZE);
  }

  function renderFinanceChunk(startIdx, count) {
    if (!dom.financePreviewTbody) return;
    const cols = getFinanceVisibleColumns();
    const displayRows = financeState.filteredRows && financeState.filteredRows.length ? financeState.filteredRows : (dom.searchInputFinance && dom.searchInputFinance.value ? [] : financeState.rows);
    const end = Math.min(startIdx + count, displayRows.length);
    const frag = document.createDocumentFragment();

    for (let idx = startIdx; idx < end; idx++) {
      const r = displayRows[idx];
      const idOrIdx = r.__id || idx;
      const tr = document.createElement('tr');
      tr.dataset.rowIdx = idx;

      // Index column
      const idxTd = document.createElement('td');
      idxTd.style.cssText = 'width: 45px; min-width: 45px; max-width: 45px; text-align: center; font-family: var(--mono); color: var(--ink3);';
      idxTd.textContent = String(idx + 1);
      tr.appendChild(idxTd);

      cols.forEach((h, colIdx) => {
        const td = document.createElement('td');
        td.dataset.colKey = h;
        td.dataset.rowIdx = idx;
        td.dataset.colIdx = colIdx;

        let val = String(r[h] ?? '');
        const hu = h.toUpperCase();
        const isPaid = hu.includes('PAID') || hu.includes('%');
        const isMoney = hu.includes('PRICE') || hu.includes('TCP') || hu.includes('DEDUCTIBLE') || hu.includes('BALANCE');
        const isDate = hu.includes('DATE');

        const w = financeState.colWidths[h] || 140;
        td.style.width = `${w}px`;
        td.style.minWidth = `${w}px`;
        td.style.maxWidth = `${w}px`;

        const isFrozen = financeState.frozenCols.has(h);
        const isLastFrz = isTrailingFrozen(colIdx, cols, financeState.frozenCols);
        const leftOffset = isFrozen ? (45 + getFrozenOffset(colIdx, cols, financeState.frozenCols, financeState.colWidths)) : 0;

        if (isFrozen) {
          td.className = `${isFrozen ? 'col-frozen' : ''} ${isLastFrz ? 'col-frozen-last' : ''}`.trim();
          td.style.left = `${leftOffset}px`;
        }

        const input = document.createElement('input');
        input.className = 'cell-input';
        if (isMoney || isPaid) input.classList.add('num-cell');

        if (isPaid) {
          input.value = formatPercent(val);
        } else if (isMoney) {
          input.value = formatPrice(val);
        } else {
          input.value = val;
        }

        input.spellcheck = false;

        if (isDate) {
          input.placeholder = 'MM/DD/YYYY';
          input.maxLength = 10;
          input.addEventListener('input', () => {
            input.value = maskDateString(input.value);
          });
        }

        let before = input.value;
        input.addEventListener('focus', () => {
          before = input.value;
          if (isMoney) {
            const pNum = parseAmount(input.value);
            input.value = pNum ? String(pNum) : '';
          }
        });

        input.addEventListener('blur', () => {
          if (isMoney) {
            const num = parseAmount(input.value);
            input.value = num ? formatPrice(num) : '';
            const cleanStr = num ? String(num) : '';
            const prevNum = parseAmount(before);
            const prevStr = prevNum ? String(prevNum) : '';
            if (cleanStr !== prevStr) {
              handleFinanceCellEdit(idOrIdx, h, cleanStr, input, td);
            }
          } else if (isDate) {
            const masked = maskDateString(input.value);
            input.value = masked;
            if (masked !== before) {
              handleFinanceCellEdit(idOrIdx, h, masked, input, td);
            }
          } else if (isPaid) {
            const pct = formatPercent(input.value);
            input.value = pct;
            if (pct !== before) {
              handleFinanceCellEdit(idOrIdx, h, pct, input, td);
            }
          } else {
            if (input.value !== before) {
              handleFinanceCellEdit(idOrIdx, h, input.value, input, td);
            }
          }
        });

        input.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            input.blur();
          } else if (e.key === 'Escape') {
            input.value = before;
            input.blur();
          }
        });

        td.appendChild(input);
        tr.appendChild(td);
      });

      // Actions column (Red trash delete button)
      const actTd = document.createElement('td');
      actTd.className = 'col-action';
      actTd.style.cssText = 'width: 55px; min-width: 55px; max-width: 55px; text-align: center;';
      actTd.innerHTML = `
        <button type="button" class="btn xs ghost danger rc-del-btn" onclick="window.deleteFinanceRecord('${escapeHtml(String(idOrIdx))}')" title="Delete record from Finance Data">
          <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M4 7h16"></path>
            <path d="M9 7V4h6v3"></path>
            <path d="M6 7l1 13a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-13"></path>
            <path d="M10 11v6"></path>
            <path d="M14 11v6"></path>
          </svg>
        </button>
      `;
      tr.appendChild(actTd);

      frag.appendChild(tr);
    }

    dom.financePreviewTbody.appendChild(frag);
    financeState.renderedRowsCount = end;
    invalidateGridCache('financePreviewTable');
  }

  async function handleFinanceCellEdit(rowIdOrIdx, colKey, newVal, inputEl, tdEl) {
    const rows = (financeState.filteredRows && financeState.filteredRows.length ? financeState.filteredRows : financeState.rows);
    let targetRow = null;
    let targetIdx = -1;
    for (let i = 0; i < financeState.rows.length; i++) {
      if (financeState.rows[i].__id === rowIdOrIdx || i === rowIdOrIdx) {
        targetRow = financeState.rows[i];
        targetIdx = i;
        break;
      }
    }
    if (!targetRow && typeof rowIdOrIdx === 'number' && financeState.rows[rowIdOrIdx]) {
      targetRow = financeState.rows[rowIdOrIdx];
      targetIdx = rowIdOrIdx;
    }

    if (!targetRow) return;

    let trimmed = String(newVal || '').trim();
    const hu = colKey.toUpperCase();
    if (hu.includes('PAID') || hu.includes('%')) {
      trimmed = formatPercent(trimmed);
    }

    targetRow[colKey] = trimmed;
    if (tdEl) tdEl.classList.add('is-dirty');
    if (inputEl) inputEl.classList.add('modified');

    try {
      await window.pywebview.api.update_finance_cell(targetRow.__id || targetIdx, colKey, trimmed);
    } catch (e) {
      console.warn('Could not sync finance edit to engine:', e);
    }
  }

  async function exportFinanceExcel() {
    const toExport = financeState.filteredRows.length ? financeState.filteredRows : financeState.rows;
    if (!toExport.length) {
      showToast('No finance data to export.', 'err');
      return;
    }
    const cols = getFinanceVisibleColumns();
    const res = await window.pywebview.api.export_records_excel(toExport, 'Finance_Data_Export.xlsx', cols);
    if (res.success) {
      showToast(`Exported to ${res.file_path}`, 'ok');
    }
  }

  window.deleteFinanceRecord = async function (idOrIdx) {
    if (!confirm('Are you sure you want to delete this record from Finance Data?')) return;
    try {
      const res = await window.pywebview.api.delete_finance_record(idOrIdx);
      if (res.success) {
        showToast('Finance record deleted.', 'ok');
        const targetId = String(idOrIdx);
        let foundIdx = financeState.rows.findIndex(r => r.__id === targetId);
        if (foundIdx === -1) {
          const num = parseInt(idOrIdx, 10);
          if (!isNaN(num) && num >= 0 && num < financeState.rows.length) {
            foundIdx = num;
          }
        }
        if (foundIdx !== -1) {
          financeState.rows.splice(foundIdx, 1);
        }
        state.financeLoadedCount = financeState.rows.length;
        if (dom.lblActiveFinanceCount) dom.lblActiveFinanceCount.textContent = `${state.financeLoadedCount.toLocaleString()} rows`;
        if (dom.lblTopFinance) dom.lblTopFinance.textContent = state.financeLoadedCount > 0 ? `Finance Data: ${state.financeLoadedCount.toLocaleString()} rows` : 'Finance Data: None';
        filterAndRenderFinance();
      } else {
        showToast(res.error || 'Failed to delete finance record.', 'err');
      }
    } catch (err) {
      showToast('Error: ' + err, 'err');
    }
  };

  function attachFinanceHeaderEvents() {
    dom.financePreviewThead.querySelectorAll('[data-fin-pin]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const col = btn.dataset.finPin;
        if (financeState.frozenCols.has(col)) {
          financeState.frozenCols.delete(col);
        } else {
          financeState.frozenCols.add(col);
        }
        saveFinanceTableConfig();
        renderFinanceTable();
      });
    });

    dom.financePreviewThead.querySelectorAll('[data-fin-filter-toggle]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const colKey = btn.dataset.finFilterToggle;
        financeState.columnFiltersOpen[colKey] = !financeState.columnFiltersOpen[colKey];
        const popup = document.getElementById(`finColFilterBox_${colKey}`);
        if (popup) {
          if (financeState.columnFiltersOpen[colKey]) {
            popup.classList.remove('hidden');
            const input = popup.querySelector('input');
            if (input) input.focus();
          } else {
            popup.classList.add('hidden');
          }
        }
      });
    });

    dom.financePreviewThead.querySelectorAll('.th-filter-input').forEach(input => {
      input.addEventListener('input', () => {
        const colKey = input.dataset.finFilterCol;
        financeState.columnFilters[colKey] = input.value.trim().toLowerCase();
        filterAndRenderFinance();
      });
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
          const colKey = input.dataset.finFilterCol;
          financeState.columnFiltersOpen[colKey] = false;
          const popup = document.getElementById(`finColFilterBox_${colKey}`);
          if (popup) popup.classList.add('hidden');
        }
      });
    });

    dom.financePreviewThead.querySelectorAll('[data-fin-resize]').forEach(resizer => {
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const colKey = resizer.dataset.finResize;
        const th = resizer.closest('th');
        const startX = e.clientX;
        const startW = th.offsetWidth;
        resizer.classList.add('is-resizing');
        document.body.classList.add('col-resizing');

        let rAF = null;
        let newW = startW;

        const onMove = (ev) => {
          newW = Math.max(60, Math.min(600, startW + (ev.clientX - startX)));
          if (rAF) cancelAnimationFrame(rAF);
          rAF = requestAnimationFrame(() => {
            updateTableColumnWidths(dom.financePreviewTable, colKey, newW);
          });
        };

        const onUp = () => {
          if (rAF) cancelAnimationFrame(rAF);
          resizer.classList.remove('is-resizing');
          document.body.classList.remove('col-resizing');
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
          financeState.colWidths[colKey] = newW;
          saveFinanceTableConfig();
          if (financeState.frozenCols.has(colKey)) {
            updateStickyOffsets(dom.financePreviewTable, getFinanceVisibleColumns(), financeState.frozenCols, financeState.colWidths, 45);
          }
        };

        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      });
    });

    dom.financePreviewThead.querySelectorAll('.th-bottom-resizer').forEach(resizer => {
      attachBottomHeaderResizer(resizer.closest('th'));
    });

    dom.financePreviewThead.querySelectorAll('[data-fin-sort]').forEach(wrap => {
      wrap.addEventListener('click', (e) => {
        if (e.target.closest('.th-resizer') || e.target.closest('.th-bottom-resizer') || e.target.closest('.th-pin-btn') || e.target.closest('.th-chevron-btn') || e.target.closest('.th-filter-popup')) return;
        const col = wrap.dataset.finSort;
        if (financeState.sortCol === col) {
          financeState.sortDir = financeState.sortDir === 'asc' ? 'desc' : (financeState.sortDir === 'desc' ? null : 'asc');
          if (!financeState.sortDir) financeState.sortCol = null;
        } else {
          financeState.sortCol = col;
          financeState.sortDir = 'asc';
        }
        filterAndRenderFinance();
      });
    });

    let draggedCol = null;
    const ths = dom.financePreviewThead.querySelectorAll('th[draggable="true"]');
    ths.forEach(th => {
      th.addEventListener('dragstart', (e) => {
        draggedCol = th.dataset.finCol;
        th.classList.add('col-dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', draggedCol);
      });

      th.addEventListener('dragend', () => {
        th.classList.remove('col-dragging');
        ths.forEach(t => t.classList.remove('drag-over-left', 'drag-over-right'));
      });

      th.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        const targetCol = th.dataset.finCol;
        if (!targetCol || targetCol === draggedCol) return;
        const rect = th.getBoundingClientRect();
        if (e.clientX < rect.left + rect.width / 2) {
          th.classList.add('drag-over-left');
          th.classList.remove('drag-over-right');
        } else {
          th.classList.add('drag-over-right');
          th.classList.remove('drag-over-left');
        }
      });

      th.addEventListener('dragleave', () => {
        th.classList.remove('drag-over-left', 'drag-over-right');
      });

      th.addEventListener('drop', (e) => {
        e.preventDefault();
        ths.forEach(t => t.classList.remove('drag-over-left', 'drag-over-right'));
        const targetCol = th.dataset.finCol;
        if (!draggedCol || !targetCol || draggedCol === targetCol) return;
        const fromIdx = financeState.colOrder.indexOf(draggedCol);
        if (fromIdx === -1) return;
        financeState.colOrder.splice(fromIdx, 1);
        const rect = th.getBoundingClientRect();
        const insertBefore = e.clientX < rect.left + rect.width / 2;
        let toIdx = financeState.colOrder.indexOf(targetCol);
        if (!insertBefore) toIdx++;
        financeState.colOrder.splice(toIdx, 0, draggedCol);
        saveFinanceTableConfig();
        renderFinanceColumnPickerMenu();
        filterAndRenderFinance();
      });
    });
  }

  // ==================== SETTINGS & SOLUTIONS EXPLORER ====================
  function setupSettingsView() {
    dom.btnSaveConfig.addEventListener('click', async () => {
      const selectedSolOption = dom.selSolution.options[dom.selSolution.selectedIndex];
      const cfg = {
        environment_url: dom.cfgEnvUrl.value.trim(),
        tenant_id: dom.cfgTenantId.value.trim(),
        client_id: dom.cfgClientId.value.trim(),
        solution_name: selectedSolOption ? selectedSolOption.text : state.activeSolutionId,
        solution_id: dom.selSolution.value,
        loa_table: dom.selLoaTable.value,
        finance_table: dom.selFinanceTable.value,
        theme: state.theme
      };
      const res = await window.pywebview.api.save_settings(cfg);
      if (res.success) {
        showToast('Settings saved successfully.', 'ok');
        const s = await window.pywebview.api.get_initial_state();
        applyState(s);
      } else {
        showToast(res.error || 'Failed to save settings.', 'err');
      }
    });

    dom.btnSignIn.addEventListener('click', async () => {
      try {
        dom.btnSignIn.disabled = true;
        if (dom.dvStatus) {
          dom.dvStatus.classList.remove('hidden');
          dom.dvStatus.className = 'dv-badge working';
          dom.dvStatus.textContent = 'CONNECTING...';
        }

        const res = await window.pywebview.api.start_sign_in();
        dom.btnSignIn.disabled = false;

        if (res.already_connected) {
          showToast('Connected to Dataverse (cached session).', 'ok');
          const s = await window.pywebview.api.get_initial_state();
          applyState(s);
          loadSolutions();
          fetchRecords();
          return;
        }

        if (res.success) {
          dom.lblDeviceCode.textContent = res.user_code;
          dom.lnkDeviceUrl.href = res.verification_uri || 'https://login.microsoft.com/device';
          dom.boxDeviceCode.classList.remove('hidden');
          if (dom.dvStatus) {
            dom.dvStatus.classList.remove('hidden');
            dom.dvStatus.className = 'dv-badge working';
            dom.dvStatus.textContent = 'WAITING FOR SIGN-IN...';
          }
          showToast('Enter code in your browser. Connection will finish automatically.', 'ok');

          if (window.authPollTimer) clearInterval(window.authPollTimer);
          window.authPollTimer = setInterval(async () => {
            try {
              const stat = await window.pywebview.api.check_sign_in_status();
              if (stat.connected) {
                clearInterval(window.authPollTimer);
                window.authPollTimer = null;
                window.onDeviceFlowComplete(true, null);
              } else if (stat.error) {
                clearInterval(window.authPollTimer);
                window.authPollTimer = null;
                window.onDeviceFlowComplete(false, stat.error);
              }
            } catch (_) {}
          }, 2000);
        } else {
          if (dom.dvStatus) dom.dvStatus.classList.add('hidden');
          showToast(res.error || 'Failed to initiate sign-in.', 'err');
        }
      } catch (err) {
        dom.btnSignIn.disabled = false;
        if (dom.dvStatus) dom.dvStatus.classList.add('hidden');
        showToast('Sign-in error: ' + err, 'err');
      }
    });

    dom.btnCopyCode.addEventListener('click', () => {
      navigator.clipboard.writeText(dom.lblDeviceCode.textContent.replace(/\s+/g, ''));
      showToast('Code copied to clipboard.', 'ok');
    });

    dom.btnSignOut.addEventListener('click', async () => {
      await window.pywebview.api.sign_out();
      const s = await window.pywebview.api.get_initial_state();
      applyState(s);
      showToast('Signed out from Dataverse.', 'ok');
    });

    dom.btnRefreshSolutions.addEventListener('click', () => {
      loadSolutions();
    });

    dom.selSolution.addEventListener('change', () => {
      loadSolutionTables(dom.selSolution.value);
    });
  }

  window.onDeviceFlowComplete = async function (success, errorMsg) {
    if (window.authPollTimer) {
      clearInterval(window.authPollTimer);
      window.authPollTimer = null;
    }
    dom.boxDeviceCode.classList.add('hidden');
    dom.btnSignIn.disabled = false;

    if (success) {
      showToast('Successfully connected to Dataverse!', 'ok');
      const s = await window.pywebview.api.get_initial_state();
      applyState(s);
      loadSolutions();
      fetchRecords();
    } else {
      if (dom.dvStatus) dom.dvStatus.classList.add('hidden');
      showToast('Authentication failed: ' + (errorMsg || 'Timed out'), 'err');
      const s = await window.pywebview.api.get_initial_state();
      applyState(s);
    }
  };

  async function loadSolutions() {
    try {
      dom.selSolution.innerHTML = '<option value="">Loading solutions...</option>';
      const res = await window.pywebview.api.get_solutions();
      if (res.success && res.solutions && res.solutions.length) {
        state.solutions = res.solutions;
        let html = '<option value="">-- All Org Custom Tables --</option>';
        res.solutions.forEach(s => {
          const isSelected = (state.activeSolutionId && (state.activeSolutionId === s.id || state.activeSolutionId === s.uniqueName)) ||
                             s.friendlyName === 'LOA MONITORING' || s.uniqueName === 'LOAMONITORING';
          html += `<option value="${escapeHtml(s.id)}" ${isSelected ? 'selected' : ''}>${escapeHtml(s.friendlyName)} (${escapeHtml(s.uniqueName)})</option>`;
        });
        dom.selSolution.innerHTML = html;

        const currentSol = dom.selSolution.value;
        loadSolutionTables(currentSol);
      } else {
        dom.selSolution.innerHTML = '<option value="">No solutions returned</option>';
        loadSolutionTables(null);
      }
    } catch (err) {
      console.warn('Could not load solutions:', err);
      dom.selSolution.innerHTML = '<option value="">(Error loading solutions)</option>';
    }
  }

  async function loadSolutionTables(solutionId) {
    try {
      const res = await window.pywebview.api.get_solution_tables(solutionId || null);
      if (res.success && res.tables) {
        state.tables = res.tables;

        let loaHtml = '';
        res.tables.forEach(t => {
          const isSel = t.logicalName === state.activeLoaTable || (state.activeLoaTable === 'cr63f_loamonitoring' && t.logicalName === 'cr63f_loamonitoring');
          loaHtml += `<option value="${escapeHtml(t.logicalName)}" ${isSel ? 'selected' : ''}>${escapeHtml(t.displayName)} (${escapeHtml(t.logicalName)})</option>`;
        });
        dom.selLoaTable.innerHTML = loaHtml || '<option value="cr63f_loamonitoring">LOA MONITORING (cr63f_loamonitoring)</option>';

        let finHtml = '';
        res.tables.forEach(t => {
          const isSel = t.logicalName === state.activeFinanceTable || (state.activeFinanceTable === 'cr63f_financereference' && t.logicalName === 'cr63f_financereference');
          finHtml += `<option value="${escapeHtml(t.logicalName)}" ${isSel ? 'selected' : ''}>${escapeHtml(t.displayName)} (${escapeHtml(t.logicalName)})</option>`;
        });
        dom.selFinanceTable.innerHTML = finHtml || '<option value="cr63f_financereference">FINANCE REFERENCE (cr63f_financereference)</option>';
      }
    } catch (err) {
      console.warn('Could not load tables:', err);
    }
  }

  function populateSettings(cfg) {
    dom.cfgEnvUrl.value = cfg.environment_url || '';
    dom.cfgTenantId.value = cfg.tenant_id || '';
    dom.cfgClientId.value = cfg.client_id || '';
    state.activeLoaTable = cfg.loa_table || 'cr63f_loamonitoring';
    state.activeFinanceTable = cfg.finance_table || 'cr63f_financereference';
    state.activeSolutionId = cfg.solution_id || '';
  }

  // ==================== TOASTS & UTILS ====================
  function showToast(message, type = 'ok') {
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.textContent = message;
    dom.toastWrap.appendChild(t);
    setTimeout(() => {
      t.style.opacity = '0';
      setTimeout(() => t.remove(), 200);
    }, 3500);
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  init();
})();
