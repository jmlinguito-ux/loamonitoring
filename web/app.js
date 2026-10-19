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
    activeView: 'dashboard',
    records: [],
    baselineRecords: [],
    filteredRecords: [],
    renderedRecordsCount: 0,
    financeLoadedCount: 0,
    financeSource: 'None',
    recentFinanceFiles: [],
    conversionLoadedCount: 0,
    conversionSource: 'None',
    recentConversionFiles: [],
    theme: 'light',
    solutions: [],
    tables: [],
    activeSolutionId: '',
    activeLoaTable: 'cr63f_loamonitoring',
    activeFinanceTable: 'cr63f_financereference',
    activeConversionTable: 'cr63f_landconversionreference',
    activeLogsTable: 'cr63f_loamonitoringlogs',
    tableSchema: null,
    stagedEdits: {},
    financeAutopopulatedRows: new Set(),
    referenceUpdatedFieldsByRowId: {},
    pendingRefDiffSummary: null,
    columnFilters: {},
    columnFiltersOpen: {},
    colWidths: {},
    colOrder: [],
    visibleCols: new Set(),
    frozenCols: new Set(),
    sortCol: null,
    sortDir: null, // 'asc' | 'desc' | null
    showOnlyChanged: false,
  };

  const financeState = {
    headers: [],
    rows: [],
    baselineRows: [],
    stagedEdits: {},
    showOnlyChanged: false,
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

  const conversionState = {
    headers: [],
    rows: [],
    baselineRows: [],
    stagedEdits: {},
    showOnlyChanged: false,
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

  const logsState = {
    logs: [],
    filteredLogs: [],
    renderedLogsCount: 0,
    colWidths: {},
    colOrder: [],
    visibleCols: new Set(),
    frozenCols: new Set(['__idx', 'log_time']),
    sortCol: null,
    sortDir: null,
    columnFilters: {},
    columnFiltersOpen: {},
    actionFilter: 'ALL',
    wwFilter: 'ALL'
  };

  let debounceTimer = null;
  let searchDebounce = null;
  let searchLogsDebounce = null;
  let searchFinDebounce = null;
  let searchConvDebounce = null;
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
    'NO DOCUMENTS',
    'N/A'
  ];

  const MODES = [
    'FOR VERIFICATION',
    'NOT AVAILABLE',
    'OPEN TO LEASE',
    'OPEN TO SALE',
    'SECURED SALE',
    'SECURED LEASE',
    'SOLD TO PARTNER',
    'SS - STP',
    'SL - STP',
    'SS - FROM LAND STP',
    'SL - FROM LAND STP',
    'SS - FROM LAND',
    'SL - FROM LAND',
    'SS - FROM ROW',
    'SS - RESCINDED',
    'SL - RESCINDED',
    'TO BE IDENTIFIED',
    'RESCINDED',
    'NOT USABLE',
    'SECURED SALE NON SP',
    'SECURED LEASE NON SP'
  ];

  const TYPE_OPTIONS = [
    'LAND',
    'ROW',
    'LAND - SPECIAL'
  ];

  // Acquisition Sub-Level codes and their display-only Account Status descriptions.
  const ACQUISITION_SUBLEVEL_STATUS = {
    'L0.0': 'FOR SOURCING', 'L1.1': 'SOURCED', 'L1.2': 'NEGO-READY ACCOUNTS',
    'L2.1': 'ONGOING NEGO', 'L2.2': 'RETURNED TO COMDEV',
    'L2.3': 'FOR RCC ENDORSEMENT', 'L2.3-EMS': 'FOR RCC ENDORSEMENT (EM - SECURED)', 'L2.3-EM': 'FOR RCC ENDORSEMENT (EM)',
    'L3.1': 'FOR ACQUISITION RISK ANALYSIS', 'L3.1-EMS': 'FOR ACQUISITION RISK ANALYSIS (EM - SECURED)', 'L3.1-EM': 'FOR ACQUISITION RISK ANALYSIS (EM)',
    'L3.2': 'FOR LAND PAYMENT SIGNATURE', 'L3.2-EMS': 'FOR LAND PAYMENT SIGNATURE (EM - SECURED)', 'L3.2-EM': 'FOR LAND PAYMENT SIGNATURE (EM)',
    'L4.1': 'FOR PET SIGNATURE', 'L4.1-EMS': 'FOR PET SIGNATURE (EM - SECURED)', 'L4.1-EM': 'FOR PET SIGNATURE (EM)',
    'L4.2A': 'WAITING FOR CHEQUE RELEASE', 'L4.2A-EMS': 'WAITING FOR CHEQUE RELEASE (EM - SECURED)', 'L4.2A-EM': 'WAITING FOR CHEQUE RELEASE (EM)',
    'L4.2B': 'FOR CHEQUE DISBURSEMENT', 'L4.2B-EMS': 'FOR CHEQUE DISBURSEMENT (EM - SECURED)', 'L4.2B-EM': 'FOR CHEQUE DISBURSEMENT (EM)',
    'L4.2C': 'ON-HOLD WITH LLD', 'L4.2C-EMS': 'ON-HOLD WITH LLD (EM - SECURED)', 'L4.2C-EM': 'ON-HOLD WITH LLD (EM)',
    'L4.2D': 'RETURNED TO CORP FIN', 'L4.2D-EMS': 'RETURNED TO CORP FIN (EM - SECURED)', 'L4.2D-EM': 'RETURNED TO CORP FIN (EM)',
    'L4.2E': 'ON-HOLD WITH CORP FIN', 'L4.2E-EMS': 'ON-HOLD WITH CORP FIN (EM - SECURED)', 'L4.2E-EM': 'ON-HOLD WITH CORP FIN (EM)',
    'L5.1A': 'SECURED', 'L5.1-EMS': 'SECURED (EM - SECURED)', 'L5.1B': 'SECURED NON SP', 'L5.1C': 'RESCINDED',
    'L6.1A': 'RECLASS ENDORSED ONLY', 'L6.1B': 'CONVERSION ENDORSED ONLY', 'L6.1C': 'RECLASS AND CONVERSION ENDORSED',
    'L6.2A': 'RECLASS APPLIED ONLY', 'L6.2B': 'CONVERSION APPLIED ONLY',
    'L6.3A': 'RECLASS APPLIED AND CONVERSION ENDORSED', 'L6.3B': 'RECLASS ENDORSED AND CONVERSION APPLIED', 'L6.3C': 'RECLASS AND CONVERSION APPLIED',
    'L6.4A': 'RECLASS DONE ONLY', 'L6.4B': 'CONVERSION DONE ONLY',
    'L6.5A': 'RECLASS DONE AND CONVERSION ENDORSED', 'L6.5B': 'RECLASS DONE AND CONVERSION APPLIED',
    'L6.5C': 'RECLASS ENDORSED AND CONVERSION DONE', 'L6.5D': 'RECLASS APPLIED AND CONVERSION DONE', 'L6.5E': 'RECLASS AND CONVERSION DONE',
    'L6.6A': 'CONVERSION DONE WITH FINALITY ONLY', 'L6.6B': 'RECLASS ENDORSED AND CONVERSION DONE WITH FINALITY',
    'L6.6C': 'RECLASS APPLIED AND CONVERSION DONE WITH FINALITY', 'L6.7': 'RECLASS DONE AND CONVERSION DONE WITH FINALITY',
    'L7.1A': 'PRE-TITLE TRANSFER WITH PRELIMENARY TRANSACTION - EJS, DOAS AND/OR DONATION',
    'L7.1B': 'PRE-TITLE TRANSFER WITH ONGOING LO COMPLIANCES', 'L7.1C': 'PRE-TITLE TRANSFER WITHIN PROHIBITORY PERIOD',
    'L7.1D': 'WITH ONGOING GOVERNMENT AGENCIES ISSUES', 'L7.1E': 'FOR TITLE ISSUANCE - PROVISIONAL RECONSTITUTION',
    'L7.1F': 'FOR ORIGINAL REGISTRATION (RD COPY TITLE ISSUANCE)', 'L7.1G': 'FOR TITLE ISSUANCE - PARTIAL CANCELLATION',
    'L7.2': 'PRE-TITLE TRANSFER WITH PRELIMENARY TRANSACTION - ON PROCESS', 'L7.3': 'TRANSFERRABLE',
    'L7.4': 'TITLE TRANSFER ON PROCESS', 'L7.5': 'TITLE TRANSFER DONE', 'L8.1': 'SOLD TO PARTNER'
  };
  const ACQUISITION_SUBLEVEL_OPTIONS = Object.keys(ACQUISITION_SUBLEVEL_STATUS);

  // Fill a select from a list of values. A value that is stored on a record but is
  // not in the list is kept as an extra option, so loading and saving it never blanks it.
  function ensureSelectValue(select, value) {
    if (!select) return;
    const v = value === null || value === undefined ? '' : String(value).trim();
    if (v && !Array.from(select.options).some(o => o.value === v)) {
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = `${v} (not in list)`;
      select.appendChild(opt);
    }
    select.value = v;
  }

  function setSelectOptions(select, values, selected, placeholder) {
    if (!select) return;
    const items = [];
    if (placeholder !== undefined) items.push({ value: '', label: placeholder });
    values.forEach(v => items.push({ value: v, label: v }));
    select.innerHTML = items.map(o => `<option value="${escapeHtml(o.value)}">${escapeHtml(o.label)}</option>`).join('');
    ensureSelectValue(select, selected);
  }

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
    'SPPC',
    'N/A'
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
    { key: 'cr63f_possrightsstatus', alt: 'Poss Rights Status', label: 'Poss Rights Status', width: 150, readonly: false },
    { key: 'cr63f_possrightsarea', alt: 'Poss Rights Area (Ha)', label: 'Poss Rights Area (Ha)', width: 155, readonly: false, isNum: true },
    { key: 'cr63f_possrightscat', alt: 'Possrights Cat', label: 'Possrights Cat', width: 130, readonly: false },
    { key: 'cr63f_catarea', alt: 'Cat Area (Ha)', label: 'Cat Area (Ha)', width: 120, readonly: false, isNum: true },
    { key: 'new_contracttype', alt: 'Contract Type', label: 'Contract Type', width: 145, readonly: false, isOptions: true, options: CONTRACT_TYPES },
    { key: 'new_mode', alt: 'Mode', label: 'Mode', width: 165, readonly: false, isOptions: true, options: MODES },
    { key: 'cr63f_conversionstatus', alt: 'Conversion Status', label: 'Conversion Status', width: 155, readonly: false },
    { key: 'cr63f_acquisitionsublevel', alt: 'Acquisition Sub-Level', label: 'Acquisition Sub-Level', width: 155, readonly: false, isOptions: true, options: ACQUISITION_SUBLEVEL_OPTIONS },
    { key: 'cr63f_rccprice', alt: 'RCC Price', label: 'RCC Price', width: 130, readonly: false, isNum: true, isPrice: true },
    { key: 'cr63f_rccterms', alt: 'RCC Terms', label: 'RCC Terms', width: 130, readonly: false },
    { key: 'cr63f_endorseddate', alt: 'Endorsed Date', label: 'Endorsed Date', width: 125, readonly: false, isDate: true },
    { key: 'work_week', alt: 'Work Week', label: 'Work Week', width: 110, readonly: true, isPseudo: true, isWorkWeek: true },
    { key: 'cr63f_secureddate', alt: 'Secured Date', label: 'Secured Date', width: 125, readonly: false, isDate: true },
    { key: 'cr63f_securedendorsedate', alt: 'Secured Endorse Date', label: 'Secured Endorse Date', width: 145, readonly: false, isDate: true },
    { key: 'cr63f_totalcontractprice', alt: 'Total Contract Price', label: 'Total Contract Price', width: 145, readonly: false, isNum: true, isPrice: true },
    { key: 'cr63f_totalpaymentdeductibletotcp', alt: 'Payment Deductible', label: 'Payment Deductible', width: 145, readonly: false, isNum: true, isPrice: true },
    { key: 'cr63f_balance', alt: 'Balance', label: 'Balance', width: 140, readonly: true, isNum: true, isPrice: true, isBalance: true },
    { key: 'cr63f_paid', alt: '% Paid', label: '% Paid', width: 95, readonly: true, isNum: true, isPercent: true },
    { key: '__action', label: 'Actions', width: 80, fixed: true, readonly: true }
  ];

  const LOGS_COLUMNS = [
    { key: '__idx', label: '#', width: 45, fixed: true, readonly: true },
    { key: 'log_time', alt: 'Log Time', label: 'Log Time', width: 155, readonly: true },
    { key: 'work_week', alt: 'Work Week', label: 'Work Week', width: 100, readonly: true },
    { key: 'action_type', alt: 'Action Type', label: 'Action Type', width: 135, readonly: true },
    { key: 'cr63f_acquisitionsublevel', alt: 'Acquisition Sub-Level', label: 'Acquisition Sub-Level', width: 155, readonly: true },
    { key: 'cr63f_possrightsstatus', alt: 'Poss Rights Status', label: 'Poss Rights Status', width: 150, readonly: true },
    { key: 'cr63f_possrightsarea', alt: 'Poss Rights Area (Ha)', label: 'Poss Rights Area (Ha)', width: 155, readonly: true },
    { key: 'cr63f_possrightscat', alt: 'Possrights Cat', label: 'Possrights Cat', width: 130, readonly: true },
    { key: 'cr63f_catarea', alt: 'Cat Area (Ha)', label: 'Cat Area (Ha)', width: 120, readonly: true },
    { key: 'cr63f_control1', alt: 'Control', label: 'Control Number', width: 140, readonly: true },
    { key: 'area_index', alt: 'cr63f_areaindex', label: 'AREA-INDEX', width: 135, readonly: true },
    { key: 'field_name', alt: 'Field Changed', label: 'Field Changed', width: 165, readonly: true },
    { key: 'old_value', alt: 'Previous Value', label: 'Previous Value', width: 145, readonly: true },
    { key: 'new_value', alt: 'New Value', label: 'New Value', width: 145, readonly: true },
    { key: 'title_no', alt: 'cr63f_titleno', label: 'Title No', width: 140, readonly: true },
    { key: 'details', alt: 'Details', label: 'Details', width: 220, readonly: true }
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
    dotConversionRef: document.getElementById('dotConversionRef'),
    lblTopConversion: document.getElementById('lblTopConversion'),
    btnTopConversionPill: document.getElementById('btnTopConversionPill'),
    btnToggleTheme: document.getElementById('btnToggleTheme'),
    tbTargetTable: document.getElementById('tbTargetTable'),

    // Nav
    tabButtons: document.querySelectorAll('.tab-btn'),
    viewPanels: document.querySelectorAll('.view-panel'),
    viewDashboard: document.getElementById('view-dashboard'),
    btnOpenNewEntry: document.getElementById('btnOpenNewEntry'),
    btnTopSyncDv: document.getElementById('btnTopSyncDv'),

    // Records View (Land Master DB)
    kpiRecordCount: document.getElementById('kpiRecordCount'),
    kpiTotalTcp: document.getElementById('kpiTotalTcp'),
    kpiTotalDeductible: document.getElementById('kpiTotalDeductible'),
    kpiTotalBalance: document.getElementById('kpiTotalBalance'),
    kpiTotalHectares: document.getElementById('kpiTotalHectares'),
    kpiHectaresNote: document.getElementById('kpiHectaresNote'),
    kpiModeActiveBadge: document.getElementById('kpiModeActiveBadge'),
    modeFilter: document.getElementById('modeFilter'),
    modeFilterPanel: document.getElementById('modeFilterPanel'),
    searchInput: document.getElementById('searchInput'),
    typeFilter: document.getElementById('typeFilter'),
    btnClearFilters: document.getElementById('btnClearFilters'),
    stagedEditsWrap: document.getElementById('stagedEditsWrap'),
    unsavedEditsChip: document.getElementById('unsavedEditsChip'),
    btnToggleShowOnlyChanged: document.getElementById('btnToggleShowOnlyChanged'),
    btnSaveStagedEdits: document.getElementById('btnSaveStagedEdits'),
    btnDiscardStagedEdits: document.getElementById('btnDiscardStagedEdits'),
    btnColumnsPicker: document.getElementById('btnColumnsPicker'),
    colPickerMenu: document.getElementById('colPickerMenu'),
    lblGridCount: document.getElementById('lblGridCount'),
    btnUpdateFromFinance: document.getElementById('btnUpdateFromFinance'),
    btnExportExcel: document.getElementById('btnExportExcel'),
    tableEmptyState: document.getElementById('tableEmptyState'),
    recordsTable: document.getElementById('recordsTable'),
    recordsTableHead: document.getElementById('recordsTableHead'),
    recordsTableBody: document.getElementById('recordsTableBody'),

    // Entry Logs View
    kpiLogsCount: document.getElementById('kpiLogsCount'),
    kpiLogsNewEntry: document.getElementById('kpiLogsNewEntry'),
    kpiLogsManual: document.getElementById('kpiLogsManual'),
    kpiLogsFinance: document.getElementById('kpiLogsFinance'),
    searchInputLogs: document.getElementById('searchInputLogs'),
    actionFilterLogs: document.getElementById('actionFilterLogs'),
    wwFilterLogs: document.getElementById('wwFilterLogs'),
    btnClearFiltersLogs: document.getElementById('btnClearFiltersLogs'),
    btnColumnsPickerLogs: document.getElementById('btnColumnsPickerLogs'),
    colPickerMenuLogs: document.getElementById('colPickerMenuLogs'),
    lblLogsGridCount: document.getElementById('lblLogsGridCount'),
    btnRefreshLogs: document.getElementById('btnRefreshLogs'),
    btnExportExcelLogs: document.getElementById('btnExportExcelLogs'),
    logsEmptyState: document.getElementById('logsEmptyState'),
    logsTable: document.getElementById('logsTable'),
    logsTableHead: document.getElementById('logsTableHead'),
    logsTableBody: document.getElementById('logsTableBody'),

    // Dashboard View
    dashWorkWeekSelect: document.getElementById('dashWorkWeekSelect'),
    btnRefreshDashboard: document.getElementById('btnRefreshDashboard'),
    btnDashClearMode: document.getElementById('btnDashClearMode'),
    dashModeFilterBanner: document.getElementById('dashModeFilterBanner'),
    dashKpiTotalRecords: document.getElementById('dashKpiTotalRecords'),
    dashKpiActiveModeCount: document.getElementById('dashKpiActiveModeCount'),
    dashKpiWeeklyChanges: document.getElementById('dashKpiWeeklyChanges'),
    dashKpiWeeklyChangesSubtitle: document.getElementById('dashKpiWeeklyChangesSubtitle'),
    dashKpiTotalTcp: document.getElementById('dashKpiTotalTcp'),
    dashKpiTcpDelta: document.getElementById('dashKpiTcpDelta'),
    dashKpiTotalDeductible: document.getElementById('dashKpiTotalDeductible'),
    dashKpiDeductibleDelta: document.getElementById('dashKpiDeductibleDelta'),
    dashKpiTotalBalance: document.getElementById('dashKpiTotalBalance'),
    dashKpiAvgPaid: document.getElementById('dashKpiAvgPaid'),
    dashModeTotalLots: document.getElementById('dashModeTotalLots'),
    dashModeBarsContainer: document.getElementById('dashModeBarsContainer'),
    dashFinancePaidRatio: document.getElementById('dashFinancePaidRatio'),
    dashFinanceProgressBar: document.getElementById('dashFinanceProgressBar'),
    dashFinBoxTcp: document.getElementById('dashFinBoxTcp'),
    dashFinBoxDeductible: document.getElementById('dashFinBoxDeductible'),
    dashFinBoxBalance: document.getElementById('dashFinBoxBalance'),
    dashFinanceBreakdownContainer: document.getElementById('dashFinanceBreakdownContainer'),
    dashFinanceFooterStats: document.getElementById('dashFinanceFooterStats'),
    dashConvSecuredRatio: document.getElementById('dashConvSecuredRatio'),
    dashConvProgressBar: document.getElementById('dashConvProgressBar'),
    dashConvSecuredCount: document.getElementById('dashConvSecuredCount'),
    dashConvInProgressCount: document.getElementById('dashConvInProgressCount'),
    dashConvTotalArea: document.getElementById('dashConvTotalArea'),
    dashConvBarsContainer: document.getElementById('dashConvBarsContainer'),
    dashActivityFeedCount: document.getElementById('dashActivityFeedCount'),
    dashActivityFeedList: document.getElementById('dashActivityFeedList'),
    dashActivitySearchInput: document.getElementById('dashActivitySearchInput'),
    dashActivityActionFilter: document.getElementById('dashActivityActionFilter'),
    dashWwComparisonPill: document.getElementById('dashWwComparisonPill'),
    dashAuditTable: document.getElementById('dashAuditTable'),
    dashAuditTableBody: document.getElementById('dashAuditTableBody'),
    dashAuditCountLabel: document.getElementById('dashAuditCountLabel'),
    dashAuditPagination: document.getElementById('dashAuditPagination'),

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
    fPossRightsStatus: document.getElementById('fPossRightsStatus'),
    fPossRightsArea: document.getElementById('fPossRightsArea'),
    fPossRightsCat: document.getElementById('fPossRightsCat'),
    fCatArea: document.getElementById('fCatArea'),
    fContractType: document.getElementById('fContractType'),
    fMode: document.getElementById('fMode'),
    fConversionStatus: document.getElementById('fConversionStatus'),
    fAcquisitionSubLevel: document.getElementById('fAcquisitionSubLevel'),
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
    btnDownloadFinanceTemplate: document.getElementById('btnDownloadFinanceTemplate'),
    btnUploadFinance: document.getElementById('btnUploadFinance'),
    btnSyncFinanceToDv: document.getElementById('btnSyncFinanceToDv'),
    btnClearFinanceData: document.getElementById('btnClearFinanceData'),
    searchInputFinance: document.getElementById('searchInputFinance'),
    typeFilterFinance: document.getElementById('typeFilterFinance'),
    btnClearFiltersFinance: document.getElementById('btnClearFiltersFinance'),
    stagedEditsWrapFinance: document.getElementById('stagedEditsWrapFinance'),
    unsavedEditsChipFinance: document.getElementById('unsavedEditsChipFinance'),
    btnToggleShowOnlyChangedFinance: document.getElementById('btnToggleShowOnlyChangedFinance'),
    btnSaveStagedEditsFinance: document.getElementById('btnSaveStagedEditsFinance'),
    btnDiscardStagedEditsFinance: document.getElementById('btnDiscardStagedEditsFinance'),
    btnColumnsPickerFinance: document.getElementById('btnColumnsPickerFinance'),
    colPickerMenuFinance: document.getElementById('colPickerMenuFinance'),
    lblFinanceGridCount: document.getElementById('lblFinanceGridCount'),
    btnExportExcelFinance: document.getElementById('btnExportExcelFinance'),
    financeEmptyState: document.getElementById('financeEmptyState'),
    financePreviewTable: document.getElementById('financePreviewTable'),
    financePreviewThead: document.getElementById('financePreviewThead'),
    financePreviewTbody: document.getElementById('financePreviewTbody'),

    // Conversion View
    lblActiveConversionSource: document.getElementById('lblActiveConversionSource'),
    lblActiveConversionCount: document.getElementById('lblActiveConversionCount'),
    lblActiveConversionTableDisplay: document.getElementById('lblActiveConversionTableDisplay'),
    btnDownloadConversionTemplate: document.getElementById('btnDownloadConversionTemplate'),
    btnUploadConversion: document.getElementById('btnUploadConversion'),
    btnSyncConversionToDv: document.getElementById('btnSyncConversionToDv'),
    btnClearConversionData: document.getElementById('btnClearConversionData'),
    searchInputConversion: document.getElementById('searchInputConversion'),
    statusFilterConversion: document.getElementById('statusFilterConversion'),
    btnClearFiltersConversion: document.getElementById('btnClearFiltersConversion'),
    stagedEditsWrapConversion: document.getElementById('stagedEditsWrapConversion'),
    unsavedEditsChipConversion: document.getElementById('unsavedEditsChipConversion'),
    btnToggleShowOnlyChangedConversion: document.getElementById('btnToggleShowOnlyChangedConversion'),
    btnSaveStagedEditsConversion: document.getElementById('btnSaveStagedEditsConversion'),
    btnDiscardStagedEditsConversion: document.getElementById('btnDiscardStagedEditsConversion'),
    btnColumnsPickerConversion: document.getElementById('btnColumnsPickerConversion'),
    colPickerMenuConversion: document.getElementById('colPickerMenuConversion'),
    lblConversionGridCount: document.getElementById('lblConversionGridCount'),
    btnExportExcelConversion: document.getElementById('btnExportExcelConversion'),
    conversionEmptyState: document.getElementById('conversionEmptyState'),
    conversionPreviewTable: document.getElementById('conversionPreviewTable'),
    conversionPreviewThead: document.getElementById('conversionPreviewThead'),
    conversionPreviewTbody: document.getElementById('conversionPreviewTbody'),

    // Settings View (Image 1 & 2)
    setConnDot: document.getElementById('setConnDot'),
    setConnText: document.getElementById('setConnText'),
    cfgEnvUrl: document.getElementById('cfgEnvUrl'),
    cfgTenantId: document.getElementById('cfgTenantId'),
    cfgClientId: document.getElementById('cfgClientId'),
    selSolution: document.getElementById('selSolution'),
    btnRefreshSolutions: document.getElementById('btnRefreshSolutions'),
    selLoaTable: document.getElementById('selLoaTable'),
    selLogsTable: document.getElementById('selLogsTable'),
    selFinanceTable: document.getElementById('selFinanceTable'),
    selConversionTable: document.getElementById('selConversionTable'),
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
    lblAppVersion: document.getElementById('lblAppVersion'),

    // Appearance (Image 1 Style)
    themeSeg: document.getElementById('themeSeg'),
    densitySeg: document.getElementById('densitySeg'),
    setRememberTable: document.getElementById('setRememberTable'),

    // Toast
    toastWrap: document.getElementById('toastWrap'),

    // Confirmation & Progress Modal
    confirmModal: document.getElementById('confirmModal'),
    confirmModalIcon: document.getElementById('confirmModalIcon'),
    confirmModalTitle: document.getElementById('confirmModalTitle'),
    confirmModalSubtitle: document.getElementById('confirmModalSubtitle'),
    confirmModalMessage: document.getElementById('confirmModalMessage'),
    btnConfirmModalClose: document.getElementById('btnConfirmModalClose'),
    confirmModalProgressWrap: document.getElementById('confirmModalProgressWrap'),
    confirmModalProgressBar: document.getElementById('confirmModalProgressBar'),
    confirmModalProgressStatus: document.getElementById('confirmModalProgressStatus'),
    confirmModalProgressPercent: document.getElementById('confirmModalProgressPercent'),
    confirmModalFooter: document.getElementById('confirmModalFooter'),
    btnConfirmModalCancel: document.getElementById('btnConfirmModalCancel'),
    btnConfirmModalOk: document.getElementById('btnConfirmModalOk'),
    // Log Delete / Revert Modal
    // Cell hover diff tooltip
    cellHoverTooltip: document.getElementById('cellHoverTooltip'),
  };

  // ==================== CONFIRMATION & PROGRESS MODAL ENGINE ====================
  let confirmModalResolver = null;

  function showConfirmModal({
    title = 'Confirmation',
    subtitle = 'Please review before proceeding',
    message = '',
    confirmText = 'Confirm',
    cancelText = 'Cancel',
    isDanger = false,
    keepOpenOnConfirm = false
  } = {}) {
    return new Promise((resolve) => {
      confirmModalResolver = (val) => {
        if (val && keepOpenOnConfirm) {
          resolve(true);
        } else {
          if (dom.confirmModal) dom.confirmModal.classList.add('hidden');
          resolve(val);
        }
      };

      if (dom.confirmModalTitle) dom.confirmModalTitle.textContent = title;
      if (dom.confirmModalSubtitle) dom.confirmModalSubtitle.textContent = subtitle;
      if (dom.confirmModalMessage) dom.confirmModalMessage.textContent = message;
      if (dom.btnConfirmModalOk) {
        dom.btnConfirmModalOk.textContent = confirmText;
        dom.btnConfirmModalOk.className = isDanger ? 'btn danger' : 'btn primary';
        dom.btnConfirmModalOk.disabled = false;
        dom.btnConfirmModalOk.classList.remove('hidden');
      }
      if (dom.btnConfirmModalCancel) {
        dom.btnConfirmModalCancel.textContent = cancelText;
        dom.btnConfirmModalCancel.disabled = false;
        dom.btnConfirmModalCancel.classList.remove('hidden');
      }
      if (dom.btnConfirmModalClose) {
        dom.btnConfirmModalClose.classList.remove('hidden');
      }
      if (dom.confirmModalIcon) {
        dom.confirmModalIcon.className = `modal-header-icon ${isDanger ? 'danger' : ''}`;
        dom.confirmModalIcon.innerHTML = isDanger
          ? '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="7.86 2 16.14 2 22 7.86 22 16.14 16.14 22 7.86 22 2 16.14 2 7.86 7.86 2"></polygon><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>'
          : '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>';
      }

      if (dom.confirmModalProgressWrap) {
        dom.confirmModalProgressWrap.classList.add('hidden');
      }
      if (dom.confirmModalProgressBar) {
        dom.confirmModalProgressBar.style.width = '0%';
      }
      if (dom.confirmModalProgressPercent) {
        dom.confirmModalProgressPercent.textContent = '0%';
      }
      if (dom.confirmModal) {
        dom.confirmModal.classList.remove('hidden');
      }
    });
  }

  function closeConfirmModal(result = false) {
    if (dom.confirmModal) {
      dom.confirmModal.classList.add('hidden');
    }
    if (confirmModalResolver) {
      const r = confirmModalResolver;
      confirmModalResolver = null;
      r(result);
    }
  }

  function setConfirmModalProgress(current, total, statusText = '') {
    if (dom.confirmModal) {
      dom.confirmModal.classList.remove('hidden');
    }
    if (dom.confirmModalProgressWrap) {
      dom.confirmModalProgressWrap.classList.remove('hidden');
    }
    const pct = total > 0 ? Math.min(100, Math.round((current / total) * 100)) : 0;
    if (dom.confirmModalProgressBar) {
      dom.confirmModalProgressBar.style.width = `${pct}%`;
    }
    if (dom.confirmModalProgressPercent) {
      dom.confirmModalProgressPercent.textContent = `${pct}%`;
    }
    if (dom.confirmModalProgressStatus) {
      dom.confirmModalProgressStatus.textContent = statusText || `Syncing ${current.toLocaleString()} / ${total.toLocaleString()} record(s)...`;
    }
  }

  window.onFinanceSyncProgress = function (current, total) {
    setConfirmModalProgress(current, total);
  };

  window.onConversionSyncProgress = function (current, total) {
    setConfirmModalProgress(current, total);
  };

  window.onCloseBlockedBySave = function () {
    showToast('Saving to Dataverse. The app will close as soon as the save and its log are finished.', 'ok');
  };

  window.onDataverseSaveProgress = function (current, total, customMsg) {
    setConfirmModalProgress(current, total, customMsg || `Saving changes: ${current.toLocaleString()} of ${total.toLocaleString()}...`);
  };

  function setupConfirmModal() {
    if (dom.btnConfirmModalClose) {
      dom.btnConfirmModalClose.addEventListener('click', () => closeConfirmModal(false));
    }
    if (dom.btnConfirmModalCancel) {
      dom.btnConfirmModalCancel.addEventListener('click', () => closeConfirmModal(false));
    }
    if (dom.btnConfirmModalOk) {
      dom.btnConfirmModalOk.addEventListener('click', () => {
        if (confirmModalResolver) {
          const r = confirmModalResolver;
          confirmModalResolver = null;
          r(true);
        } else {
          closeConfirmModal(true);
        }
      });
    }
    if (dom.confirmModal) {
      dom.confirmModal.addEventListener('click', (e) => {
        if (e.target === dom.confirmModal) {
          if (dom.confirmModalProgressWrap && !dom.confirmModalProgressWrap.classList.contains('hidden')) {
            return;
          }
          closeConfirmModal(false);
        }
      });
    }
  }

  // ==================== EXCEL-LIKE MULTI-CELL SELECTION & CLIPBOARD ====================
  const GRID_SELECTIONS = {}; // tableId -> { anchor:{r,c}, focus:{r,c}, dragging }
  const GRID_CFG = {};        // tableId -> { setMem, updateCell, refresh }
  let GRID_CELL_CACHE = {};   // tableId -> 2D array of data <td> elements
  let GRID_PASTE_PROXY = null;
  let GRID_GLOBAL_WIRED = false;
  let GRID_ACTIVE_TABLE_ID = null;

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
    GRID_ACTIVE_TABLE_ID = tableEl.id;
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
    const invalidDates = [];
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
        if (isDate) {
          const normalizedDate = normalizeDateInput(rawVal);
          if (normalizedDate === null) {
            // Never write a scrambled date: skip the cell and report it.
            invalidDates.push(rawVal);
            continue;
          }
          rawVal = normalizedDate;
        }

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

    if (invalidDates.length) {
      showToast(`Skipped ${invalidDates.length} cell(s) with an invalid date (e.g. "${invalidDates[0]}"). Use MM/DD/YYYY.`, 'err');
    }
    if (!okCount) {
      if (!invalidDates.length) showToast('Nothing to paste into the selected cells.', 'err');
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
    if (GRID_ACTIVE_TABLE_ID && (!target || !target.matches || !target.matches('input, textarea, select, [contenteditable="true"]'))) {
      const active = document.getElementById(GRID_ACTIVE_TABLE_ID);
      if (active && active.offsetParent !== null && GRID_SELECTIONS[active.id]) return active;
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
        if (count <= 0) return;
        e.preventDefault();
        gridCopyToClipboard(tableEl);
      } else if (key === 'v') {
        if (tableEl.id === 'logsTable') {
          e.preventDefault();
          return;
        }
        const info = gridTargetInfo(tableEl);
        if (!info || (count <= 1 && !info.targetIsSelect)) return;
        gridFocusPasteProxy();
      }
    });

    document.addEventListener('copy', (e) => {
      const tableEl = gridResolveActiveTable(e.target);
      if (!tableEl) return;
      if (gridSelectionCount(tableEl) <= 0) return;
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
      if (tableEl.id === 'logsTable') {
        e.preventDefault();
        return;
      }
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
      GRID_ACTIVE_TABLE_ID = tableEl.id;
      if (tableEl.id === 'logsTable') tableEl.focus({ preventScroll: true });
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
    setupLogsView();
    setupDashboardView();
    setupFinanceView();
    setupConversionView();
    setupSettingsView();
    setupConfirmModal();

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
        const rows = financeState.filteredRows;
        const row = rows[rowIdx];
        if (row) {
          row[colKey] = val;
          const origIdx = financeState.rows.indexOf(row);
          if (origIdx !== -1) financeState.rows[origIdx][colKey] = val;
        }
      },
      updateCell: (rowIdx, colKey, val, inputEl, tdEl) => {
        const rows = financeState.filteredRows;
        const row = rows[rowIdx];
        if (row) {
          handleFinanceCellEdit(row.__id || rowIdx, colKey, val, inputEl, tdEl);
        }
      },
      refresh: () => filterAndRenderFinance()
    });

    // Excel-like cell selection for Conversion Data table
    initGridSelection(dom.conversionPreviewTable, {
      setMem: (rowIdx, colKey, val) => {
        const rows = conversionState.filteredRows;
        const row = rows[rowIdx];
        if (row) {
          row[colKey] = val;
          const origIdx = conversionState.rows.indexOf(row);
          if (origIdx !== -1) conversionState.rows[origIdx][colKey] = val;
        }
      },
      updateCell: (rowIdx, colKey, val, inputEl, tdEl) => {
        const rows = conversionState.filteredRows;
        const row = rows[rowIdx];
        if (row) {
          handleConversionCellEdit(row.__id || rowIdx, colKey, val, inputEl, tdEl);
        }
      },
      refresh: () => filterAndRenderConversion()
    });

    // Excel-like cell selection for Entry Logs table
    initGridSelection(dom.logsTable, {
      setMem: () => {},
      updateCell: () => {},
      refresh: () => filterAndRenderLogs()
    });

    // Initial render of headers & empty state across all views
    renderTableHead();
    filterAndRenderRecords();
    renderLogsHead();
    renderFinanceTable();
    renderConversionTable();

    // Wire scroll listeners for chunked rendering
    wireTableScrollListeners();
    // Wire cell hover diff tooltip
    wireCellHoverTooltip();

    // In-app updates (Settings > About & updates)
    wireUpdateControls();

    // Check pywebview readiness
    if (window.pywebview && window.pywebview.api) {
      loadInitialState();
    } else {
      window.addEventListener('pywebviewready', loadInitialState);
    }
  }

  function wireTableScrollListeners() {
    const recordsWrap = dom.recordsTable ? dom.recordsTable.closest('.table-wrap') : null;
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

    const logsWrap = dom.logsTable ? dom.logsTable.closest('.table-wrap') : null;
    if (logsWrap && !logsWrap.dataset.scrollWired) {
      logsWrap.dataset.scrollWired = '1';
      logsWrap.addEventListener('scroll', () => {
        if (logsWrap.scrollTop + logsWrap.clientHeight >= logsWrap.scrollHeight - 300) {
          if (logsState.renderedLogsCount < logsState.filteredLogs.length) {
            renderLogsChunk(logsState.renderedLogsCount, CHUNK_SIZE);
          }
        }
      });
    }

    const financeWrap = dom.financePreviewTable ? dom.financePreviewTable.closest('.table-wrap') : null;
    if (financeWrap && !financeWrap.dataset.scrollWired) {
      financeWrap.dataset.scrollWired = '1';
      financeWrap.addEventListener('scroll', () => {
        if (financeWrap.scrollTop + financeWrap.clientHeight >= financeWrap.scrollHeight - 300) {
          const displayRows = financeState.filteredRows;
          if (financeState.renderedRowsCount < displayRows.length) {
            renderFinanceChunk(financeState.renderedRowsCount, CHUNK_SIZE);
          }
        }
      });
    }

    const convWrap = dom.conversionPreviewTable ? dom.conversionPreviewTable.closest('.table-wrap') : null;
    if (convWrap && !convWrap.dataset.scrollWired) {
      convWrap.dataset.scrollWired = '1';
      convWrap.addEventListener('scroll', () => {
        if (convWrap.scrollTop + convWrap.clientHeight >= convWrap.scrollHeight - 300) {
          const displayRows = conversionState.filteredRows;
          if (conversionState.renderedRowsCount < displayRows.length) {
            renderConversionChunk(conversionState.renderedRowsCount, CHUNK_SIZE);
          }
        }
      });
    }
  }

  // ==================== CELL HOVER DIFF TOOLTIP ====================
  let tooltipHideTimer = null;

  function wireTableHoverTooltip(table) {
    const tip = dom.cellHoverTooltip;
    if (!table || !tip || table.id !== 'recordsTable') return;
    if (table.dataset.tooltipWired) return;
    table.dataset.tooltipWired = '1';

    function isLogForRecord(l, rowRecord, tLotId) {
      if (!l) return false;
      const lCtrl = String(l.cr63f_control1 || l['Control Number'] || l.Control || l['Control No'] || '').trim().toUpperCase();
      const lArea = String(l.area_index || l['AREA-INDEX'] || l['Area Index'] || l['Area-Index'] || l.cr63f_areaindex || '').trim().toUpperCase();
      const lTitle = String(l.title_no || l['Title No'] || l['TITLE NO'] || l.cr63f_titleno || '').trim().toUpperCase();
      const lId = getAuditMasterId(l);

      if (rowRecord) {
        const rId = String(rowRecord.__id || rowRecord.id || rowRecord.cr63f_loamonitoringid || '').trim();
        if (lId && rId) return lId === rId;
        const rCtrl = String(rowRecord.cr63f_control1 || rowRecord['Control Number'] || rowRecord.Control || rowRecord['Control No'] || '').trim().toUpperCase();
        if (lCtrl && rCtrl) return lCtrl === rCtrl;
        const rArea = String(rowRecord.cr63f_areaindex || rowRecord['Area Index'] || rowRecord['Area-Index'] || '').trim().toUpperCase();
        if (lArea && rArea && lArea === rArea) return true;
        const rTitle = String(rowRecord.cr63f_titleno || rowRecord['Title No'] || rowRecord['TITLE NO'] || '').trim().toUpperCase();
        if (lTitle && rTitle && lTitle === rTitle) return true;
      }

      if (tLotId) {
        const lotUpper = tLotId.toUpperCase();
        if (lCtrl && lCtrl === lotUpper) return true;
        if (lArea && lArea === lotUpper) return true;
        if (lTitle && lTitle === lotUpper) return true;
      }
      return false;
    }

    function isLogForField(l, colKey, fieldName) {
      if (!l) return false;
      const lField = String(l.field_name || l['Field Changed'] || l.cr63f_fieldchanged || '').trim();
      if (!lField || lField === 'Initial Entry') return false;

      const targetAttr = getFieldNameAttributeKey(colKey || fieldName);
      const logAttr = getFieldNameAttributeKey(lField);

      if (targetAttr && logAttr && targetAttr.toLowerCase() === logAttr.toLowerCase()) return true;

      const clean = s => String(s || '').toLowerCase().replace(/^cr63f_|^new_/, '').replace(/[^a-z0-9]/g, '');
      const normTarget = clean(fieldName || colKey);
      const normLog = clean(lField);
      return normTarget && normLog && normTarget === normLog;
    }

    function getLatestLogForCell(td) {
      if (!logsState.logs || !logsState.logs.length) return null;
      const colKey = td.dataset.colKey || '';
      const fieldName = td.dataset.fieldName || colKey || '';
      if (!colKey && !fieldName) return null;

      const tr = td.closest('tr');
      const rowId = td.dataset.rowId || (tr ? tr.dataset.rowId : null);
      const rowRecord = (rowId && state.records) ? state.records.find(r => r.__id === rowId) : null;
      const tLotId = String(td.dataset.lotId || '').trim();

      const cellLogs = [];
      for (const l of logsState.logs) {
        if (['FINANCE_UPDATE', 'CONVERSION_UPDATE'].includes(getLogActionType(l))) continue;
        if (!isLogForField(l, colKey, fieldName)) continue;
        if (!isLogForRecord(l, rowRecord, tLotId)) continue;
        cellLogs.push(l);
      }
      return cellLogs.length ? cellLogs : null;
    }

    function getPreviousValueFromLogs(cellLogs, currentLiveVal, colDef) {
      if (!cellLogs || !cellLogs.length) return '';
      const isPrice = colDef && (colDef.isPrice || /price|tcp|deductible|balance|amount|cost/i.test(`${colDef.key} ${colDef.label || ''}`));
      const isPct = colDef && colDef.isPercent;

      const valsMatch = (v1, v2) => {
        if (v1 === undefined || v1 === null || v2 === undefined || v2 === null) return false;
        const s1 = String(v1).trim();
        const s2 = String(v2).trim();
        if (!s1 || !s2 || s1 === '-' || s2 === '-' || s1 === '—' || s2 === '—') return false;
        if (isPrice) {
          const n1 = parseAmount(s1);
          const n2 = parseAmount(s2);
          if (n1 !== null && n2 !== null) return Math.abs(n1 - n2) < 0.005;
        }
        if (isPct) {
          return formatPercent(s1) === formatPercent(s2);
        }
        return s1.toLowerCase() === s2.toLowerCase();
      };

      const isValid = v => {
        if (v === undefined || v === null) return false;
        const s = String(v).trim();
        return s !== '' && s !== '-' && s !== '—' && s !== 'None' && s !== 'null' && s !== 'undefined';
      };

      // 1. First, check latestLog.old_value (Previous Value recorded when cell was edited)
      const l0 = cellLogs[0];
      const old0 = l0.old_value ?? l0.cr63f_previousvalue ?? l0['Previous Value'];
      if (isValid(old0) && !valsMatch(old0, currentLiveVal)) {
        return String(old0).trim();
      }

      // 2. Check subsequent logs in Entry Logs (e.g. 2nd log's new_value or old_value)
      for (let i = 1; i < cellLogs.length; i++) {
        const li = cellLogs[i];
        const newI = li.new_value ?? li.cr63f_newvalue ?? li['New Value'];
        if (isValid(newI) && !valsMatch(newI, currentLiveVal)) {
          return String(newI).trim();
        }
        const oldI = li.old_value ?? li.cr63f_previousvalue ?? li['Previous Value'];
        if (isValid(oldI) && !valsMatch(oldI, currentLiveVal)) {
          return String(oldI).trim();
        }
      }

      // 3. Fallback: if old0 was valid even if somehow matched currentLiveVal
      if (isValid(old0)) return String(old0).trim();

      // 4. Fallback to second log's new_value
      if (cellLogs.length > 1) {
        const l1 = cellLogs[1];
        const n1 = l1.new_value ?? l1.cr63f_newvalue ?? l1['New Value'] ?? l1.old_value;
        if (isValid(n1)) return String(n1).trim();
      }

      return '';
    }

    function showTip(td) {
      const colKey = td.dataset.colKey || '';
      const fieldName = td.dataset.fieldName || colKey || '';
      const colDef = DEFAULT_COLUMNS.find(c => c.key === colKey || c.alt === colKey || c.label === fieldName) || { key: colKey, label: fieldName };
      const tr = td.closest('tr');
      const rowId = td.dataset.rowId || (tr ? tr.dataset.rowId : null);
      const rowRecord = (rowId && state.records) ? state.records.find(r => r.__id === rowId) : null;
      const lotId = td.dataset.lotId || (rowRecord ? (rowRecord.cr63f_areaindex || rowRecord['Area-Index'] || rowRecord.cr63f_control1 || '') : '');
      const isDirty = td.classList.contains('is-dirty');

      const cellLogs = getLatestLogForCell(td);
      if (!isDirty && !cellLogs) {
        hideTip();
        return;
      }

      const latestLog = cellLogs ? cellLogs[0] : null;

      // Current live value in the cell (what the user actually sees in Land Master table)
      const cellInput = td.querySelector('input, select');
      let liveVal = cellInput ? cellInput.value : (rowRecord ? (rowRecord[colDef.key] ?? (colDef.alt ? rowRecord[colDef.alt] : '') ?? '') : (td.textContent || '').trim());

      let currVal = liveVal;
      let prevVal = '';
      let actionType = td.dataset.diffType || 'manual';

      if (isDirty) {
        actionType = 'manual';
        const baseRec = (state.baselineRecords || []).find(b => b.__id === rowId) || {};
        const baseRaw = baseRec[colDef.key] ?? (colDef.alt ? baseRec[colDef.alt] : '') ?? '';
        prevVal = String(baseRaw);
      } else if (cellLogs && cellLogs.length) {
        actionType = String(latestLog.action_type || latestLog['Action Type'] || 'EDIT').toUpperCase();
        prevVal = latestLog.old_value ?? latestLog.cr63f_previousvalue ?? latestLog['Previous Value'] ?? '';
      } else {
        prevVal = td.dataset.prevValue || '';
      }

      let badgeClass = 'manual';
      let badgeLabel = '✎ Manual Edit';

      if (isDirty) {
        badgeClass = 'manual';
        badgeLabel = '✎ Unsaved Edit';
      } else if (actionType.includes('REVERT')) {
        badgeClass = 'ref';
        badgeLabel = '⟳ Reverted Field';
      } else if (actionType.includes('REF') || actionType.includes('FINANCE') || actionType.includes('CONVERSION') || td.dataset.diffType === 'ref') {
        badgeClass = 'ref';
        badgeLabel = '⟳ Reference Update';
      }

      const prevDisplay = formatCellChangeValue(fieldName, prevVal);
      const currDisplay = formatCellChangeValue(fieldName, currVal);

      const canRevert = isDirty || isLogRevertible(latestLog);
      const revertBtnHtml = canRevert
        ? `<button type="button" class="btn xs primary cell-tooltip-revert-btn" data-row-id="${escapeHtml(rowId || '')}" data-col-key="${escapeHtml(colDef.key || colKey)}" data-field-name="${escapeHtml(fieldName)}" data-revert-val="${escapeHtml(String(prevVal))}" onclick="window.revertCellFromTooltip(event, this)">
            <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"></polyline><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"></path></svg>
            Revert Field
          </button>`
        : `<button type="button" class="btn xs ghost cell-tooltip-revert-btn" style="opacity: 0.55; cursor: not-allowed;" title="No previous entry log found for this cell" disabled>
            Revert Field
          </button>`;

      tip.innerHTML = `
        <div class="cell-tooltip-header">
          <span class="cell-tooltip-field-name">${escapeHtml(fieldName)}</span>
          <span class="cell-tooltip-badge ${badgeClass}">${escapeHtml(badgeLabel)}</span>
        </div>
        <div class="cell-tooltip-body">
          <div class="cell-tooltip-diff-row">
            <span class="cell-tooltip-label">Was</span>
            <span class="cell-tooltip-value old">${escapeHtml(prevDisplay)}</span>
          </div>
          <div class="cell-tooltip-diff-row">
            <span class="cell-tooltip-label">Now</span>
            <span class="cell-tooltip-value new">${escapeHtml(currDisplay)}</span>
          </div>
        </div>
        <div class="cell-tooltip-footer" style="display: flex; align-items: center; justify-content: space-between; gap: 8px;">
          <div>
            <span class="cell-tooltip-footer-label">Record</span>
            <span class="cell-tooltip-footer-lot">${escapeHtml(lotId)}</span>
          </div>
          ${revertBtnHtml}
        </div>
      `;

      tip.removeAttribute('aria-hidden');
      tip._sourceCell = td;
      tip._rowId = rowId;
      tip._latestLog = latestLog;
      tip._prevVal = prevVal;
      tip._targetRevertVal = prevVal;
      tip._colKey = colDef.key || colKey;
      tip._fieldName = fieldName;
      positionTip(tip, td);
      tip.style.pointerEvents = 'auto';
      tip.classList.add('visible');

      const revBtn = tip.querySelector('.cell-tooltip-revert-btn:not([disabled])');
      if (revBtn) {
        revBtn.onmousedown = (e) => { e.stopPropagation(); };
        revBtn.onclick = (e) => {
          e.stopPropagation();
          e.preventDefault();
          window.revertCellFromTooltip(e, revBtn);
        };
      }
    }

    function positionTip(tip, td) {
      const rect = td.getBoundingClientRect();
      const margin = 6;
      const vw = window.innerWidth;
      const vh = window.innerHeight;

      tip.style.left = '-9999px';
      tip.style.top = '-9999px';

      const tipW = tip.offsetWidth || 260;
      const tipH = tip.offsetHeight || 130;

      // Position fixed right above the cell
      let x = rect.left;
      let y = rect.top - tipH - margin;

      // Flip below cell if not enough space above
      if (y < 8) y = rect.bottom + margin;

      // Clamp horizontally to viewport
      if (x + tipW > vw - 8) x = rect.right - tipW;
      x = Math.max(8, Math.min(x, vw - tipW - 8));
      y = Math.max(8, Math.min(y, vh - tipH - 8));

      tip.style.left = Math.round(x) + 'px';
      tip.style.top = Math.round(y) + 'px';
    }

    function hideTip() {
      tip.classList.remove('visible');
      tip.style.pointerEvents = 'none';
      tip.setAttribute('aria-hidden', 'true');
      tip._sourceCell = null;
      tip._isHovered = false;
      if (tooltipHideTimer) clearTimeout(tooltipHideTimer);
      tooltipHideTimer = setTimeout(() => {
        if (!tip.classList.contains('visible')) {
          tip.innerHTML = '';
        }
      }, 250);
    }

    function handleCellHoverOrClick(td) {
      if (!td) return;
      if (tip && tip._isHovered) return;
      const colKey = td.dataset.colKey || '';
      if (colKey === '__idx' || colKey === '__action' || colKey === 'cr63f_balance' || colKey === 'cr63f_paid') return;

      const isDirty = td.classList.contains('is-dirty');
      const cellLogs = getLatestLogForCell(td);
      if (isDirty || cellLogs) {
        if (tooltipHideTimer) { clearTimeout(tooltipHideTimer); tooltipHideTimer = null; }
        showTip(td);
      }
    }

    table.addEventListener('mouseover', (e) => {
      const td = e.target.closest('td');
      handleCellHoverOrClick(td);
    });

    table.addEventListener('mousemove', (e) => {
      const td = e.target.closest('td');
      if (!td) return;
      if (tip.classList.contains('visible') && (tip._sourceCell === td || tip._isHovered)) return;
      handleCellHoverOrClick(td);
    });

    table.addEventListener('click', (e) => {
      const td = e.target.closest('td');
      handleCellHoverOrClick(td);
    });

    table.addEventListener('focusin', (e) => {
      const td = e.target.closest('td');
      handleCellHoverOrClick(td);
    });

    table.addEventListener('mouseout', (e) => {
      const td = e.target.closest('td');
      if (!td) return;
      const related = e.relatedTarget;
      if (td.contains(related)) return;
      if (tip && (tip === related || tip.contains(related))) return;
      if (tooltipHideTimer) clearTimeout(tooltipHideTimer);
      tooltipHideTimer = setTimeout(() => {
        if (!tip._isHovered) {
          hideTip();
        }
      }, 1200);
    });
  }

  window.revertCellFromTooltip = async function (e, btnEl) {
    if (e) {
      if (e.stopPropagation) e.stopPropagation();
      if (e.preventDefault) e.preventDefault();
    }
    const tip = dom.cellHoverTooltip;
    if (!tip || !tip._sourceCell || !tip._sourceCell.closest('#recordsTable')) return;
    const cell = tip._sourceCell;
    const btn = btnEl || (e && e.target ? e.target.closest('.cell-tooltip-revert-btn') : null);
    const rowId = (btn && btn.dataset.rowId) || tip._rowId || cell.dataset.rowId;
    const colKey = (btn && btn.dataset.colKey) || tip._colKey || cell.dataset.colKey;
    const latestLog = tip._latestLog;
    const isDirty = cell.classList.contains('is-dirty');
    tip.classList.remove('visible');
    tip.setAttribute('aria-hidden', 'true');
    tip._sourceCell = null;

    if (!isDirty && latestLog) {
      const logId = latestLog.__id || latestLog.id || latestLog.cr63f_loamonitoringlogid;
      if (logId) await window.revertLogRecord(String(logId));
      return;
    }
    // Undo a staged edit locally. A saved change uses the version-checked
    // Dataverse revert above, rather than a blind staged overwrite.
    const staged = state.stagedEdits[rowId];
    if (!isDirty || !staged) return;
    const column = DEFAULT_COLUMNS.find(c => c.key === colKey) || { key: colKey };
    delete staged[column.key];
    if (column.alt) delete staged[column.alt];
    if (STAGED_PAYMENT_KEYS.includes(column.key)) {
      STAGED_DERIVED_KEYS.forEach(key => delete staged[key]);
      if (STAGED_PAYMENT_KEYS.some(key => key in staged)) {
        const record = state.records.find(row => String(row.__id) === String(rowId));
        if (record) recalcStagedPayment(rowId, record, cell);
      }
    }
    if (!Object.keys(staged).some(key => !key.startsWith('__'))) delete state.stagedEdits[rowId];
    updateStagedEditsUI();
    recalculateKPIs();
    filterAndRenderRecords();
    showToast('Unsaved field edit discarded.', 'ok');
  };

  function wireCellHoverTooltip() {
    if (dom.cellHoverTooltip && !dom.cellHoverTooltip.dataset.hoverWired) {
      dom.cellHoverTooltip.dataset.hoverWired = '1';
      dom.cellHoverTooltip.addEventListener('mouseenter', () => {
        dom.cellHoverTooltip._isHovered = true;
        if (tooltipHideTimer) { clearTimeout(tooltipHideTimer); tooltipHideTimer = null; }
      });
      dom.cellHoverTooltip.addEventListener('mouseleave', (e) => {
        dom.cellHoverTooltip._isHovered = false;
        const rel = e.relatedTarget;
        if (dom.cellHoverTooltip._sourceCell && dom.cellHoverTooltip._sourceCell.contains(rel)) {
          return;
        }
        if (tooltipHideTimer) clearTimeout(tooltipHideTimer);
        tooltipHideTimer = setTimeout(() => {
          if (!dom.cellHoverTooltip._isHovered) {
            hideTip();
          }
        }, 800);
      });
    }
    if (dom.recordsTable) wireTableHoverTooltip(dom.recordsTable);
    if (dom.financePreviewTable) wireTableHoverTooltip(dom.financePreviewTable);
    if (dom.conversionPreviewTable) wireTableHoverTooltip(dom.conversionPreviewTable);
  }

  function updateCellChangePreview(rowId, colDef, value, inputEl, tdEl) {
    if (!tdEl || !colDef) return;
    const record = state.records.find(r => r.__id === rowId) || {};
    const baseline = (state.baselineRecords || []).find(r => r.__id === rowId) || record;
    const previousRaw = baseline[colDef.key] ?? (colDef.alt ? baseline[colDef.alt] : '') ?? '';
    const previous = String(previousRaw).trim();
    const current = String(value ?? '').trim();
    let same = previous === current;

    if (colDef.isPrice || /PRICE|TCP|DEDUCTIBLE|BALANCE|COST|AMOUNT/i.test(`${colDef.key} ${colDef.label || ''}`)) {
      const previousNumber = parseAmount(previous);
      const currentNumber = parseAmount(current);
      if (previousNumber !== null && currentNumber !== null) same = previousNumber === currentNumber;
    } else if (colDef.isPercent) {
      same = formatPercent(previous) === formatPercent(current);
    }

    if (same) {
      if (state.stagedEdits[rowId]) {
        delete state.stagedEdits[rowId][colDef.key];
        if (colDef.alt) delete state.stagedEdits[rowId][colDef.alt];
        if (Object.keys(state.stagedEdits[rowId]).length === 0) delete state.stagedEdits[rowId];
        updateStagedEditsUI();
      }
      if (!tdEl.classList.contains('cell-reference-updated')) {
        tdEl.classList.remove('is-dirty');
        delete tdEl.dataset.prevValue;
        delete tdEl.dataset.currValue;
        delete tdEl.dataset.fieldName;
        delete tdEl.dataset.diffType;
        delete tdEl.dataset.lotId;
        if (inputEl) inputEl.classList.remove('modified');
      }
    } else {
      tdEl.classList.add('is-dirty');
      if (inputEl) inputEl.classList.add('modified');
      tdEl.dataset.prevValue = previous;
      tdEl.dataset.currValue = current;
      tdEl.dataset.fieldName = colDef.label || colDef.key;
      tdEl.dataset.diffType = 'manual';
      tdEl.dataset.lotId = String(record.cr63f_areaindex || record['Area-Index'] || rowId);
        if (tdEl._refreshDiffTooltip) {
          tdEl._refreshDiffTooltip();
        } else if (tdEl._lastPointer) {
          tdEl.dispatchEvent(new MouseEvent('mouseover', {
            bubbles: true,
            clientX: tdEl._lastPointer.clientX,
            clientY: tdEl._lastPointer.clientY
          }));
        }
        return;
    }

    if (tdEl._hideDiffTooltip) tdEl._hideDiffTooltip();
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

      if (s.finance_loaded_count > 0) {
        try {
          const finPrev = await window.pywebview.api.get_finance_preview();
          if (finPrev && finPrev.rows && finPrev.rows.length) {
            const sample = finPrev.rows;
            renderFinancePreview(Object.keys(sample[0] || {}), sample);
          }
        } catch (_) {}
      }

      if (s.conversion_loaded_count > 0) {
        try {
          const convPrev = await window.pywebview.api.get_conversion_preview();
          if (convPrev && convPrev.rows && convPrev.rows.length) {
            const sample = convPrev.rows;
            renderConversionPreview(Object.keys(sample[0] || {}), sample);
          }
        } catch (_) {}
      }

      if (s.connected) {
        loadSolutions();
        await Promise.all([
          fetchRecords(),
          fetchEntryLogs()
        ]);
      } else {
        // Load entry logs from local persistent cache if offline / not yet connected
        fetchEntryLogs();
      }
      if (s.connected && state.activeView === 'dashboard') {
        await loadDashboardAnalytics(currentDashWeek(), selectedDashboardMode);
      }
    } catch (err) {
      console.warn('Init state error:', err);
    }
  }

  // Settings > About label and the sidebar "v1.0.x by JMLI" line. Prefer the
  // installed version reported by the updater over the build constant.
  function setAppVersionLabels(version) {
    const v = String(version || '').trim();
    if (!v) return;
    if (dom.lblAppVersion) dom.lblAppVersion.textContent = `VERSION ${v}`;
    const brand = document.getElementById('lblBrandVersion');
    if (brand) brand.textContent = `v${v} by JMLI`;
  }

  function applyState(s) {
    state.connected = s.connected;
    state.financeLoadedCount = s.finance_loaded_count || 0;
    state.financeSource = s.finance_source || 'None';
    setAppVersionLabels((s.update && s.update.currentVersion) || s.app_version);
    if (s.update) renderUpdate(s.update);
    state.recentFinanceFiles = s.recent_finance_files || [];
    state.conversionLoadedCount = s.conversion_loaded_count || 0;
    state.conversionSource = s.conversion_source || 'None';
    state.recentConversionFiles = s.recent_conversion_files || [];
    state.theme = s.theme || 'light';
    state.activeLoaTable = s.loa_table || 'cr63f_loamonitoring';
    state.activeFinanceTable = s.finance_table || 'cr63f_financereference';
    state.activeConversionTable = s.conversion_table || 'cr63f_landconversionreference';
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
      dom.lblTopFinance.textContent = `${state.financeLoadedCount.toLocaleString()} rows`;
      if (dom.lblActiveFinanceCount) dom.lblActiveFinanceCount.textContent = `${state.financeLoadedCount.toLocaleString()} rows`;
      if (dom.lblActiveFinanceSource) dom.lblActiveFinanceSource.textContent = state.financeSource;
    } else {
      dom.dotFinanceRef.classList.remove('on');
      dom.lblTopFinance.textContent = 'Not loaded';
      if (dom.lblActiveFinanceCount) dom.lblActiveFinanceCount.textContent = '0 rows';
      if (dom.lblActiveFinanceSource) dom.lblActiveFinanceSource.textContent = 'None';
    }

    // Update conversion reference status
    if (state.conversionLoadedCount > 0) {
      if (dom.dotConversionRef) dom.dotConversionRef.classList.add('on');
      if (dom.lblTopConversion) dom.lblTopConversion.textContent = `${state.conversionLoadedCount.toLocaleString()} rows`;
      if (dom.lblActiveConversionCount) dom.lblActiveConversionCount.textContent = `${state.conversionLoadedCount.toLocaleString()} rows`;
      if (dom.lblActiveConversionSource) dom.lblActiveConversionSource.textContent = state.conversionSource;
    } else {
      if (dom.dotConversionRef) dom.dotConversionRef.classList.remove('on');
      if (dom.lblTopConversion) dom.lblTopConversion.textContent = 'Not loaded';
      if (dom.lblActiveConversionCount) dom.lblActiveConversionCount.textContent = '0 rows';
      if (dom.lblActiveConversionSource) dom.lblActiveConversionSource.textContent = 'None';
    }

    if (dom.tbTargetTable) dom.tbTargetTable.textContent = state.activeLoaTable;
    if (dom.lblActiveFinanceTableDisplay) {
      dom.lblActiveFinanceTableDisplay.textContent = state.activeFinanceTable;
    }
    if (dom.lblActiveConversionTableDisplay) {
      dom.lblActiveConversionTableDisplay.textContent = state.activeConversionTable;
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
  function updateTopRefreshAction() {
    if (!dom.btnTopSyncDv) return;
    const actions = {
      records: { label: 'Sync Land Master', description: 'Reload Land Master from Dataverse' },
      finance: { label: 'Sync Finance', description: 'Load finance data from Dataverse' },
      conversion: { label: 'Sync Conversion', description: 'Load conversion data from Dataverse' }
    };
    const action = actions[state.activeView];
    dom.btnTopSyncDv.classList.toggle('hidden', !action);
    dom.btnTopSyncDv.setAttribute('aria-label', action?.description || 'Refresh data');
    dom.btnTopSyncDv.title = action?.description || '';
    if (!dom.btnTopSyncDv.disabled) {
      dom.btnTopSyncDv.textContent = action?.label || '';
    }
  }

  function setupNavigation() {
    dom.tabButtons.forEach(btn => {
      btn.addEventListener('click', () => {
        const v = btn.getAttribute('data-view');
        switchView(v);
      });
    });

    dom.btnTopConnPill.addEventListener('click', () => switchView('settings'));
    dom.btnTopFinancePill.addEventListener('click', () => switchView('finance'));
    if (dom.btnTopConversionPill) {
      dom.btnTopConversionPill.addEventListener('click', () => switchView('conversion'));
    }

    if (dom.btnToggleTheme) {
      dom.btnToggleTheme.addEventListener('click', () => {
        const nextTheme = state.theme === 'dark' ? 'light' : 'dark';
        applyTheme(nextTheme);
        window.pywebview.api.save_settings({ theme: nextTheme });
      });
    }

    // Top Bar "Sync Dataverse" button: syncs active tab
    dom.btnTopSyncDv.addEventListener('click', async () => {
      if (state.activeView === 'conversion') {
        dom.btnTopSyncDv.disabled = true;
        dom.btnTopSyncDv.textContent = 'Syncing Conversion...';
        try {
          const convRes = await window.pywebview.api.sync_conversion_from_dataverse();
          if (convRes.success) {
            showToast(`Synced ${convRes.row_count} rows from Dataverse.`, 'ok');
            const s = await window.pywebview.api.get_initial_state();
            applyState(s);
            const sample = convRes.sample_rows || [];
            renderConversionPreview(convRes.headers || (sample.length ? Object.keys(sample[0]) : []), sample);
          } else {
            showToast(convRes.error || 'Failed to sync Conversion data.', 'err');
          }
        } catch (e) {
          showToast('Error syncing conversion: ' + e, 'err');
        } finally {
          dom.btnTopSyncDv.disabled = false;
          updateTopRefreshAction();
        }
      } else if (state.activeView === 'finance') {
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
          updateTopRefreshAction();
        }
      } else {
        fetchRecords();
      }
    });
    updateTopRefreshAction();
  }

  function switchView(viewName) {
    state.activeView = viewName;

    dom.tabButtons.forEach(btn => {
      if (btn.getAttribute('data-view') === viewName) {
        btn.classList.add('active');
        btn.setAttribute('aria-current', 'page');
      } else {
        btn.classList.remove('active');
        btn.removeAttribute('aria-current');
      }
    });
    updateTopRefreshAction();

    dom.viewPanels.forEach(p => {
      if (p.id === `view-${viewName}`) {
        p.classList.add('active');
      } else {
        p.classList.remove('active');
      }
    });

    if (viewName === 'records') {
      filterAndRenderRecords();
    } else if (viewName === 'settings' && state.connected && (!state.solutions || !state.solutions.length)) {
      loadSolutions();
    } else if (viewName === 'logs') {
      fetchEntryLogs();
    } else if (viewName === 'dashboard') {
      loadDashboardAnalytics(currentDashWeek(), selectedDashboardMode);
    }
  }

  // ==================== NUMBER & PERCENT FORMATTING HELPERS ====================
  // True when an amount field holds nothing (so 0 stays 0 and blank stays blank).
  function isBlankAmount(val) {
    return val === null || val === undefined || String(val).replace(/[₱$,\s]/g, '') === '' ||
      ['-', '—'].includes(String(val).trim());
  }

  function parseAmount(val) {
    if (val === null || val === undefined || val === '') return 0;
    const raw = String(val).trim();
    const clean = raw.replace(/[^0-9.-]/g, '');
    let num = parseFloat(clean);
    if (isNaN(num)) return 0;
    // Accounting negative, e.g. (1,234): same sign rule as parse_float in Python.
    if (/^\(.*\)$/.test(raw)) num = -Math.abs(num);
    return num;
  }

  function formatAmount(val) {
    if (val === null || val === undefined || val === '') return '';
    const raw = String(val).trim();
    const clean = raw.replace(/[^0-9.-]/g, '');
    if (!clean || clean === '-' || clean === '.') return '';
    let num = parseFloat(clean);
    if (isNaN(num)) return raw.replace(/[₱$]/g, '').trim();
    if (/^\(.*\)$/.test(raw) || (raw.startsWith('-') && !clean.startsWith('-'))) {
      num = -Math.abs(num);
    }
    return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function formatAmountWithCommas(val) {
    if (val === null || val === undefined) return '';
    let raw = String(val).trim();
    if (!raw) return '';
    const isNeg = /^\(.*\)$/.test(raw) || raw.startsWith('-');
    let s = raw.replace(/[^0-9.-]/g, '');
    if (!s || s === '-' || s === '.') return s;

    const parts = s.split('.');
    const integerPart = parts[0];
    const decimalPart = parts.length > 1 ? parts.slice(1).join('') : null;

    const sign = (isNeg || integerPart.startsWith('-')) ? '-' : '';
    const cleanInt = integerPart.replace(/^-/, '');
    const formattedInt = cleanInt ? parseInt(cleanInt, 10).toLocaleString('en-US') : '';

    if (decimalPart !== null) {
      return `${sign}${formattedInt || '0'}.${decimalPart.slice(0, 2)}`;
    }
    return `${sign}${formattedInt}`;
  }

  function attachLiveAmountFormatting(input, onUpdate) {
    if (!input) return;
    input.addEventListener('input', () => {
      const raw = input.value;
      const cursorPos = input.selectionStart;
      const oldLen = raw.length;

      if (/[0-9]/.test(raw)) {
        const formatted = formatAmountWithCommas(raw);
        input.value = formatted;
        const newLen = formatted.length;
        const newPos = Math.max(0, cursorPos + (newLen - oldLen));
        try {
          input.setSelectionRange(newPos, newPos);
        } catch (_) {}
      }

      if (onUpdate) onUpdate();
    });

    input.addEventListener('blur', () => {
      if (input.value.trim() !== '') {
        const num = parseAmount(input.value);
        if (num !== null && !isNaN(num)) {
          input.value = num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
      }
      if (onUpdate) onUpdate();
    });
  }

  function formatPrice(val) {
    if (val === null || val === undefined || val === '') return '';
    const clean = String(val).replace(/[^0-9.-]/g, '');
    if (!clean || clean === '-' || clean === '.') return '';
    const num = parseFloat(clean);
    if (isNaN(num)) return String(val).trim();
    return '₱' + num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function formatCellChangeValue(fieldName, value) {
    if (value === null || value === undefined || String(value).trim() === '') return '—';
    const raw = String(value).trim();
    if (['-', '—', 'null', 'undefined'].includes(raw.toLowerCase())) return '—';
    if (/%|PAID/i.test(String(fieldName || ''))) return formatPercent(raw);
    if (!/(PRICE|TCP|DEDUCTIBLE|BALANCE|COST|AMOUNT)/i.test(String(fieldName || ''))) return raw;

    const clean = raw.replace(/[₱$,\s]/g, '');
    const match = clean.match(/^(-?)(\d+)(?:\.(\d*))?$/);
    if (!match) return raw;
    const groupedInteger = match[2].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    const fraction = match[3] || '';
    const decimals = `.${fraction.padEnd(2, '0')}`;
    return `₱${match[1]}${groupedInteger}${decimals}`;
  }

  // Live mask while typing. Digits only (01052026) are shaped as MM/DD/YYYY.
  // Once the user types a separator (1/5/2026, 2026-10-12) the value is left as
  // typed, apart from removing other characters, so it is never reshuffled.
  function maskDateString(val) {
    if (!val) return '';
    const s = String(val);
    if (/[/\-.]/.test(s)) return s.replace(/[^\d/\-.]/g, '').slice(0, 10);
    const digits = s.replace(/\D/g, '').slice(0, 8);
    if (digits.length >= 5) {
      return `${digits.slice(0, 2)}/${digits.slice(2, 4)}/${digits.slice(4)}`;
    } else if (digits.length >= 3) {
      return `${digits.slice(0, 2)}/${digits.slice(2)}`;
    }
    return digits;
  }

  // Final form of a typed or pasted date: 'MM/DD/YYYY', '' for blank, or null
  // when it is not a real date. Accepts M/D/YYYY, M/D/YY, YYYY-MM-DD, MMDDYYYY
  // and Excel serial numbers. Month-first; D/M only when the first number is
  // above 12 (same rule as the Python engine).
  function normalizeDateInput(val) {
    const s = String(val === null || val === undefined ? '' : val).trim();
    if (!s || ['-', '—', 'none', 'null'].includes(s.toLowerCase())) return '';
    let y, m, d;
    let match;
    if ((match = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T].*)?$/))) {
      [y, m, d] = [+match[1], +match[2], +match[3]];
    } else if ((match = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})(?:[ T].*)?$/))) {
      let a = +match[1], b = +match[2];
      y = +match[3];
      if (match[3].length === 2) y += y >= 50 ? 1900 : 2000;
      if (a > 12 && b <= 12) [a, b] = [b, a];
      [m, d] = [a, b];
    } else if ((match = s.match(/^(\d{2})(\d{2})(\d{4})$/))) {
      [m, d, y] = [+match[1], +match[2], +match[3]];
    } else if (/^\d{5}(\.\d+)?$/.test(s)) {
      // Excel serial day number (days since 1899-12-30).
      const dt = new Date(Date.UTC(1899, 11, 30) + Math.floor(+s) * 86400000);
      [y, m, d] = [dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()];
    } else {
      return null;
    }
    const check = new Date(Date.UTC(y, m - 1, d));
    if (m < 1 || m > 12 || check.getUTCMonth() !== m - 1 || check.getUTCDate() !== d || y < 1753 || y > 9999) {
      return null;
    }
    return `${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}/${y}`;
  }

  // Shared blur rule for date inputs: normalise, or warn and put the old value back.
  function commitDateInput(input, before) {
    const normalized = normalizeDateInput(input.value);
    if (normalized === null) {
      showToast(`"${input.value}" is not a valid date. Use MM/DD/YYYY.`, 'err');
      input.value = before;
      return null;
    }
    input.value = normalized;
    return normalized;
  }

  // Today's local date as MM/DD/YYYY (the Endorsed Date default).
  function todayMMDDYYYY() {
    const now = new Date();
    return `${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')}/${now.getFullYear()}`;
  }

  function formatDateMMDDYYYY(dateVal) {
    if (!dateVal) return '';
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
      const y = dateVal.getFullYear();
      if (y <= 1900) return '';
      const m = String(dateVal.getMonth() + 1).padStart(2, '0');
      const d = String(dateVal.getDate()).padStart(2, '0');
      return `${m}/${d}/${y}`;
    }
    const s = String(dateVal).split('T')[0].split(' ')[0].trim();
    if (!s || s === '-' || s.toLowerCase() === 'none' || s.toLowerCase() === 'null') return '';
    const isoMatch = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (isoMatch) {
      const [, y, m, d] = isoMatch;
      if (parseInt(y, 10) <= 1900) return '';
      return `${m.padStart(2, '0')}/${d.padStart(2, '0')}/${y}`;
    }
    const usMatch = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
    if (usMatch) {
      let [, m, d, y] = usMatch;
      if (y.length === 2) {
        y = parseInt(y, 10) >= 50 ? `19${y}` : `20${y}`;
      }
      if (parseInt(y, 10) <= 1900) return '';
      return `${m.padStart(2, '0')}/${d.padStart(2, '0')}/${y}`;
    }
    try {
      const dt = new Date(s);
      if (!isNaN(dt.getTime())) {
        const y = dt.getUTCFullYear();
        if (y <= 1900) return '';
        const m = String(dt.getUTCMonth() + 1).padStart(2, '0');
        const d = String(dt.getUTCDate()).padStart(2, '0');
        return `${m}/${d}/${y}`;
      }
    } catch (_) {}
    return '';
  }

  function getWorkWeek(dateVal) {
    if (!dateVal) return '';
    const s = String(dateVal).split('T')[0].split(' ')[0].trim();
    if (!s || s.includes('1900')) return '';
    let y, m, d;
    const isoMatch = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    const usMatch = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (isoMatch) {
      y = parseInt(isoMatch[1], 10);
      m = parseInt(isoMatch[2], 10) - 1;
      d = parseInt(isoMatch[3], 10);
    } else if (usMatch) {
      y = parseInt(usMatch[3], 10);
      const first = parseInt(usMatch[1], 10);
      const second = parseInt(usMatch[2], 10);
      if (first > 12 && second <= 12) {
        // Same fallback as the Python engine: not a valid M/D date, so read it as D/M.
        m = second - 1;
        d = first;
      } else {
        m = first - 1;
        d = second;
      }
    } else {
      const dt = new Date(s);
      if (isNaN(dt.getTime())) return '';
      y = dt.getFullYear();
      m = dt.getMonth();
      d = dt.getDate();
    }
    if (y <= 1900) return '';
    const dateObj = new Date(Date.UTC(y, m, d));
    // An impossible date (13/13/2026, 2/30/2026) has no work week: Date.UTC would
    // roll it into another month, while the Python engine returns ''.
    if (dateObj.getUTCFullYear() !== y || dateObj.getUTCMonth() !== m || dateObj.getUTCDate() !== d) return '';
    const pyWeekday = (dateObj.getUTCDay() + 6) % 7;
    // Tuesday-to-Monday weeks (decision 2026-10-14), same rule as loa_engine.get_work_week_info:
    // a week belongs to the year its Saturday falls in, and WW-01 is the first week
    // whose Saturday is in that year, so every year starts at WW-01.
    const daysSinceTue = (pyWeekday + 6) % 7;
    const startUtc = new Date(dateObj.getTime() - daysSinceTue * 86400000);
    const satUtc = new Date(startUtc.getTime() + 4 * 86400000);
    const primaryYear = satUtc.getUTCFullYear();
    const jan1 = new Date(Date.UTC(primaryYear, 0, 1));
    const daysToFirstSaturday = (6 - jan1.getUTCDay() + 7) % 7;
    const ww1Start = new Date(jan1.getTime() + (daysToFirstSaturday - 4) * 86400000);
    const weekNum = Math.floor((startUtc.getTime() - ww1Start.getTime()) / (7 * 86400000)) + 1;
    return `WW-${String(weekNum).padStart(2, '0')}`;
  }

  function formatPercent(val) {
    // Blank stays blank (as in Python): an account without a TCP has no % Paid, not 0%.
    if (val === null || val === undefined || String(val).trim() === '') return '';
    const s = String(val).trim();
    const display = num => `${num.toLocaleString('en-US', { useGrouping: false, minimumFractionDigits: 2, maximumFractionDigits: 10 })}%`;
    if (s.endsWith('%')) {
      const num = parseFloat(s.replace('%', '').trim());
      return isNaN(num) ? s : display(num);
    }
    const num = parseFloat(s.replace(/,/g, ''));
    if (isNaN(num)) return s;
    // Stored % Paid is on a 0-100 scale: 0.5 means 0.50%, not 50%.
    return display(num);
  }

  // A % Paid as Dataverse will store it: one decimal place (cr63f_paid precision 1).
  function storedPercentText(val) {
    const num = parseFloat(String(val === null || val === undefined ? '' : val).replace(/[%,\s]/g, ''));
    if (isNaN(num)) return formatPercent(val);
    return `${(Math.round(num * 10) / 10).toFixed(1)}%`;
  }

  function formatLogValue(fieldName, value) {
    const raw = value === null || value === undefined ? '' : String(value).trim();
    if (!raw || ['-', 'Previous', 'Updated', 'Created'].includes(raw)) return raw || '-';
    const field = String(fieldName || '').trim().toUpperCase();
    if (field === '% PAID' || field.includes('PERCENT PAID')) return formatPercent(raw);
    if (['TOTAL CONTRACT PRICE', 'TOTAL PAYMENT DEDUCTIBLE TO TCP', 'BALANCE', 'RCC PRICE'].includes(field)) {
      return formatPrice(raw);
    }
    return raw;
  }

  // ==================== UNIFIED TABLE SORT COMPARATOR ====================
  function isEmptyTableValue(v) {
    if (v === null || v === undefined) return true;
    const s = String(v).trim();
    return s === '' || s === '-' || s === '—' || s.toLowerCase() === 'n/a' || s.toLowerCase() === 'null' || s.toLowerCase() === 'undefined';
  }

  function compareTableValues(valA, valB, colDef = {}, mult = 1) {
    const emptyA = isEmptyTableValue(valA);
    const emptyB = isEmptyTableValue(valB);

    if (emptyA && emptyB) return 0;
    if (emptyA) return 1;  // empty always at bottom
    if (emptyB) return -1; // empty always at bottom

    const strA = String(valA).trim();
    const strB = String(valB).trim();

    // 1. Work Week comparison (e.g. WW-05 vs WW-38)
    const isWwA = /^WW-\d+$/i.test(strA);
    const isWwB = /^WW-\d+$/i.test(strB);
    if ((colDef.isWorkWeek || isWwA) && isWwA && isWwB) {
      const numA = parseInt(strA.replace(/[^0-9]/g, ''), 10) || 0;
      const numB = parseInt(strB.replace(/[^0-9]/g, ''), 10) || 0;
      return (numA - numB) * mult;
    }

    // 2. Numeric / Currency / Percentage comparison
    const isExplicitNumeric = colDef.isNum || colDef.isPrice || colDef.isPercent || colDef.isBalance;
    const cleanNumStrA = strA.replace(/^[₱$\s]+/, '').replace(/,/g, '').replace(/%$/, '').trim();
    const cleanNumStrB = strB.replace(/^[₱$\s]+/, '').replace(/,/g, '').replace(/%$/, '').trim();
    const isNumA = /^-?\d+(\.\d+)?$/.test(cleanNumStrA);
    const isNumB = /^-?\d+(\.\d+)?$/.test(cleanNumStrB);

    if (isExplicitNumeric || (isNumA && isNumB)) {
      const numA = parseFloat(cleanNumStrA);
      const numB = parseFloat(cleanNumStrB);
      if (!isNaN(numA) && !isNaN(numB)) {
        return (numA - numB) * mult;
      }
    }

    // 3. Date / Timestamp comparison
    const isDatePattern = (s) => /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(s) || /^\d{1,2}[-/]\d{1,2}[-/]\d{4}/.test(s);
    if (colDef.isDate || (isDatePattern(strA) && isDatePattern(strB))) {
      const parsedA = Date.parse(strA.replace(/-/g, '/'));
      const parsedB = Date.parse(strB.replace(/-/g, '/'));
      if (!isNaN(parsedA) && !isNaN(parsedB)) {
        return (parsedA - parsedB) * mult;
      }
    }

    // 4. Natural Alphanumeric String comparison
    return strA.localeCompare(strB, undefined, { numeric: true, sensitivity: 'base' }) * mult;
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
      fin_sort_col: (financeState && financeState.sortCol) || '',
      fin_sort_dir: (financeState && financeState.sortDir) || '',
      conv_col_widths: (conversionState && conversionState.colWidths) || {},
      conv_col_order: (conversionState && conversionState.colOrder) || [],
      conv_visible_cols: (conversionState && conversionState.visibleCols) ? Array.from(conversionState.visibleCols) : [],
      conv_frozen_cols: (conversionState && conversionState.frozenCols) ? Array.from(conversionState.frozenCols) : [],
      conv_sort_col: (conversionState && conversionState.sortCol) || '',
      conv_sort_dir: (conversionState && conversionState.sortDir) || '',
      logs_col_widths: (logsState && logsState.colWidths) || {},
      logs_col_order: (logsState && logsState.colOrder) || [],
      logs_visible_cols: (logsState && logsState.visibleCols) ? Array.from(logsState.visibleCols) : [],
      logs_frozen_cols: (logsState && logsState.frozenCols) ? Array.from(logsState.frozenCols) : [],
      logs_sort_col: (logsState && logsState.sortCol) || '',
      logs_sort_dir: (logsState && logsState.sortDir) || '',
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
      localStorage.setItem('fin_sort_col', financeState.sortCol || '');
      localStorage.setItem('fin_sort_dir', financeState.sortDir || '');
    } catch (_) {}
    saveTableLayoutBackend();
  }

  function saveConversionTableConfig() {
    try {
      localStorage.setItem('conv_col_widths', JSON.stringify(conversionState.colWidths));
      localStorage.setItem('conv_col_order', JSON.stringify(conversionState.colOrder));
      localStorage.setItem('conv_visible_cols', JSON.stringify(Array.from(conversionState.visibleCols)));
      localStorage.setItem('conv_frozen_cols', JSON.stringify(Array.from(conversionState.frozenCols)));
      localStorage.setItem('conv_sort_col', conversionState.sortCol || '');
      localStorage.setItem('conv_sort_dir', conversionState.sortDir || '');
    } catch (_) {}
    saveTableLayoutBackend();
  }

  function saveLogsTableConfig() {
    try {
      localStorage.setItem('logs_col_widths', JSON.stringify(logsState.colWidths));
      localStorage.setItem('logs_col_order', JSON.stringify(logsState.colOrder));
      localStorage.setItem('logs_visible_cols', JSON.stringify(Array.from(logsState.visibleCols)));
      localStorage.setItem('logs_frozen_cols', JSON.stringify(Array.from(logsState.frozenCols)));
      localStorage.setItem('logs_sort_col', logsState.sortCol || '');
      localStorage.setItem('logs_sort_dir', logsState.sortDir || '');
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
    if (layout.fin_sort_col !== undefined) {
      financeState.sortCol = layout.fin_sort_col || null;
      financeState.sortDir = layout.fin_sort_dir || null;
      try {
        localStorage.setItem('fin_sort_col', financeState.sortCol || '');
        localStorage.setItem('fin_sort_dir', financeState.sortDir || '');
      } catch (_) {}
    }

    // Conversion layout
    if (layout.conv_col_widths && Object.keys(layout.conv_col_widths).length > 0) {
      conversionState.colWidths = Object.assign({}, conversionState.colWidths, layout.conv_col_widths);
      try { localStorage.setItem('conv_col_widths', JSON.stringify(conversionState.colWidths)); } catch (_) {}
    }
    if (Array.isArray(layout.conv_col_order) && layout.conv_col_order.length > 0) {
      conversionState.colOrder = layout.conv_col_order;
      try { localStorage.setItem('conv_col_order', JSON.stringify(conversionState.colOrder)); } catch (_) {}
    }
    if (Array.isArray(layout.conv_visible_cols) && layout.conv_visible_cols.length > 0) {
      conversionState.visibleCols = new Set(layout.conv_visible_cols);
      try { localStorage.setItem('conv_visible_cols', JSON.stringify(Array.from(conversionState.visibleCols))); } catch (_) {}
    }
    if (Array.isArray(layout.conv_frozen_cols)) {
      conversionState.frozenCols = new Set(layout.conv_frozen_cols);
      try { localStorage.setItem('conv_frozen_cols', JSON.stringify(Array.from(conversionState.frozenCols))); } catch (_) {}
    }
    if (layout.conv_sort_col !== undefined) {
      conversionState.sortCol = layout.conv_sort_col || null;
      conversionState.sortDir = layout.conv_sort_dir || null;
      try {
        localStorage.setItem('conv_sort_col', conversionState.sortCol || '');
        localStorage.setItem('conv_sort_dir', conversionState.sortDir || '');
      } catch (_) {}
    }

    // Logs layout
    if (layout.logs_col_widths && Object.keys(layout.logs_col_widths).length > 0) {
      logsState.colWidths = Object.assign({}, logsState.colWidths, layout.logs_col_widths);
      try { localStorage.setItem('logs_col_widths', JSON.stringify(logsState.colWidths)); } catch (_) {}
    }
    if (Array.isArray(layout.logs_col_order) && layout.logs_col_order.length > 0) {
      const validKeys = new Set(LOGS_COLUMNS.map(c => c.key));
      const filtered = layout.logs_col_order.filter(k => validKeys.has(k));
      LOGS_COLUMNS.forEach(c => {
        if (!filtered.includes(c.key)) {
          if (c.key === 'cr63f_control1') {
            const actIdx = filtered.indexOf('action_type');
            if (actIdx !== -1) filtered.splice(actIdx + 1, 0, c.key);
            else filtered.push(c.key);
          } else if (c.key === 'log_time') {
            const idxIdx = filtered.indexOf('__idx');
            if (idxIdx !== -1) filtered.splice(idxIdx + 1, 0, c.key);
            else filtered.unshift(c.key);
          } else {
            filtered.push(c.key);
          }
        }
      });
      const ctrlIdx = filtered.indexOf('cr63f_control1');
      const areaIdx = filtered.indexOf('area_index');
      if (ctrlIdx !== -1 && areaIdx !== -1 && ctrlIdx > areaIdx) {
        filtered.splice(ctrlIdx, 1);
        const newAreaIdx = filtered.indexOf('area_index');
        filtered.splice(newAreaIdx, 0, 'cr63f_control1');
      }
      const timeIdx = filtered.indexOf('log_time');
      const wwIdx = filtered.indexOf('work_week');
      if (timeIdx !== -1 && wwIdx !== -1 && timeIdx > wwIdx) {
        filtered.splice(timeIdx, 1);
        const newWwIdx = filtered.indexOf('work_week');
        filtered.splice(newWwIdx, 0, 'log_time');
      }
      logsState.colOrder = filtered;
      try { localStorage.setItem('logs_col_order', JSON.stringify(logsState.colOrder)); } catch (_) {}
    }
    if (Array.isArray(layout.logs_visible_cols) && layout.logs_visible_cols.length > 0) {
      logsState.visibleCols = new Set(layout.logs_visible_cols);
      LOGS_COLUMNS.forEach(c => {
        if (!logsState.visibleCols.has(c.key)) logsState.visibleCols.add(c.key);
      });
      try { localStorage.setItem('logs_visible_cols', JSON.stringify(Array.from(logsState.visibleCols))); } catch (_) {}
    }
    if (Array.isArray(layout.logs_frozen_cols)) {
      logsState.frozenCols = new Set(layout.logs_frozen_cols);
      try { localStorage.setItem('logs_frozen_cols', JSON.stringify(Array.from(logsState.frozenCols))); } catch (_) {}
    }
    if (layout.logs_sort_col !== undefined) {
      logsState.sortCol = layout.logs_sort_col || null;
      logsState.sortDir = layout.logs_sort_dir || null;
      try {
        localStorage.setItem('logs_sort_col', logsState.sortCol || '');
        localStorage.setItem('logs_sort_dir', logsState.sortDir || '');
      } catch (_) {}
    }

    renderTableHead();
    renderColumnPickerMenu();
    renderLogsHead();
    renderLogsColumnPickerMenu();
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

    loadLogsTableConfig();
  }

  function loadLogsTableConfig() {
    const LOGS_LAYOUT_VERSION = 'v3_ctrl_ww_order';
    const savedVer = localStorage.getItem('logs_layout_ver');
    if (savedVer !== LOGS_LAYOUT_VERSION) {
      localStorage.setItem('logs_layout_ver', LOGS_LAYOUT_VERSION);
      logsState.colOrder = LOGS_COLUMNS.map(c => c.key);
      logsState.visibleCols = new Set(LOGS_COLUMNS.map(c => c.key));
      try {
        localStorage.setItem('logs_col_order', JSON.stringify(logsState.colOrder));
        localStorage.setItem('logs_visible_cols', JSON.stringify(Array.from(logsState.visibleCols)));
      } catch (_) {}
    }

    try {
      const sw = localStorage.getItem('logs_col_widths');
      logsState.colWidths = sw ? JSON.parse(sw) : {};
    } catch (_) { logsState.colWidths = {}; }

    try {
      const so = localStorage.getItem('logs_col_order');
      if (so) {
        const parsed = JSON.parse(so);
        const validKeys = new Set(LOGS_COLUMNS.map(c => c.key));
        const filtered = parsed.filter(k => validKeys.has(k));
        LOGS_COLUMNS.forEach(c => {
          if (!filtered.includes(c.key)) {
            if (c.key === 'cr63f_control1') {
              const actIdx = filtered.indexOf('action_type');
              if (actIdx !== -1) filtered.splice(actIdx + 1, 0, c.key);
              else filtered.push(c.key);
            } else if (c.key === 'log_time') {
              const idxIdx = filtered.indexOf('__idx');
              if (idxIdx !== -1) filtered.splice(idxIdx + 1, 0, c.key);
              else filtered.unshift(c.key);
            } else {
              filtered.push(c.key);
            }
          }
        });
        // If cr63f_control1 was positioned after area_index in older saved layouts, migrate it before area_index
        const ctrlIdx = filtered.indexOf('cr63f_control1');
        const areaIdx = filtered.indexOf('area_index');
        if (ctrlIdx !== -1 && areaIdx !== -1 && ctrlIdx > areaIdx) {
          filtered.splice(ctrlIdx, 1);
          const newAreaIdx = filtered.indexOf('area_index');
          filtered.splice(newAreaIdx, 0, 'cr63f_control1');
        }
        const timeIdx = filtered.indexOf('log_time');
        const wwIdx = filtered.indexOf('work_week');
        if (timeIdx !== -1 && wwIdx !== -1 && timeIdx > wwIdx) {
          filtered.splice(timeIdx, 1);
          const newWwIdx = filtered.indexOf('work_week');
          filtered.splice(newWwIdx, 0, 'log_time');
        }
        logsState.colOrder = filtered;
      } else {
        logsState.colOrder = LOGS_COLUMNS.map(c => c.key);
      }
    } catch (_) {
      logsState.colOrder = LOGS_COLUMNS.map(c => c.key);
    }

    try {
      const sv = localStorage.getItem('logs_visible_cols');
      if (sv) {
        logsState.visibleCols = new Set(JSON.parse(sv));
        LOGS_COLUMNS.forEach(c => {
          if (!logsState.visibleCols.has(c.key)) logsState.visibleCols.add(c.key);
        });
      } else {
        logsState.visibleCols = new Set(LOGS_COLUMNS.map(c => c.key));
      }
    } catch (_) {
      logsState.visibleCols = new Set(LOGS_COLUMNS.map(c => c.key));
    }

    try {
      const sf = localStorage.getItem('logs_frozen_cols');
      logsState.frozenCols = sf ? new Set(JSON.parse(sf)) : new Set(['__idx', 'log_time']);
    } catch (_) {
      logsState.frozenCols = new Set(['__idx', 'log_time']);
    }

    try {
      const sc = localStorage.getItem('logs_sort_col');
      const sd = localStorage.getItem('logs_sort_dir');
      logsState.sortCol = sc || null;
      logsState.sortDir = sd || null;
    } catch (_) {
      logsState.sortCol = null;
      logsState.sortDir = null;
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

  function getLogsOrderedColumns() {
    const map = new Map(LOGS_COLUMNS.map(c => [c.key, c]));
    const ordered = [];
    logsState.colOrder.forEach(k => {
      if (map.has(k)) ordered.push(map.get(k));
    });
    LOGS_COLUMNS.forEach(c => {
      if (!logsState.colOrder.includes(c.key)) ordered.push(c);
    });
    return ordered;
  }

  function getLogsVisibleColumns() {
    return getLogsOrderedColumns().filter(c => c.fixed || logsState.visibleCols.has(c.key));
  }

  function getLogActionType(log, fallback = '') {
    const nameAction = String(log.cr63f_name || log.name || '').match(/\(([^()]+)\)\s*$/);
    const details = String(log.cr63f_details || log.details || log.Details || '');
    const detailsAction = details.match(/\[SOURCE:([^\]]+)\]/);
    const legacyReference = /Updated Finance Reference\b/i.test(details) ? 'FINANCE_UPDATE'
      : (/Updated Conversion Reference\b/i.test(details) ? 'CONVERSION_UPDATE' : '');
    return String(log.action_type || log['Action Type'] || log.cr63f_actiontype || (nameAction && nameAction[1]) || (detailsAction && detailsAction[1]) || legacyReference || fallback).toUpperCase().replace(/ /g, '_');
  }

  function getAuditMasterId(log) {
    const marker = ['cr63f_name', 'name', 'cr63f_details', 'details', 'Details']
      .map(key => String(log[key] || '')).join(' ').match(/\[MASTER:([^\]]+)\]/);
    return String(log.__master_record_id || (marker && marker[1]) || log.record_id || log.cr63f_loamonitoringid || '').trim();
  }

  function applyReferenceSaveResult(referenceState, result) {
    const saved = new Set((result && result.saved_ids || []).map(String));
    referenceState.baselineRows = referenceState.baselineRows || [];
    referenceState.rows.forEach((row, index) => {
      const id = String(row.__id || index);
      if (!saved.has(id)) return;
      referenceState.baselineRows[index] = JSON.parse(JSON.stringify(row));
      delete referenceState.stagedEdits[id];
    });
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
    const th = tableEl.querySelector(`thead th[data-col-key="${colKey}"], thead th[data-fin-col="${colKey}"], thead th[data-conv-col="${colKey}"], thead th[data-log-col="${colKey}"]`);
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
        const th = tableEl.querySelector(`thead th[data-col-key="${key}"], thead th[data-fin-col="${key}"], thead th[data-conv-col="${key}"], thead th[data-log-col="${key}"]`);
        if (th) th.style.left = `${offset}px`;
        const tds = tableEl.querySelectorAll(`tbody td[data-col-key="${key}"]`);
        tds.forEach(td => td.style.left = `${offset}px`);
        offset += w;
      }
    });
  }

  // ==================== RECORDS HUB ====================
  // ==================== MULTI-SELECT FILTER ====================
  // Land Master Type and Mode filters (2026-10-14). Nothing ticked means all. Each
  // option shows how many accounts it matches in the current view.
  class MultiSelectFilter {
    constructor(host, opts) {
      this.host = host;
      this.allLabel = opts.allLabel;
      this.noun = opts.noun;
      this.label = opts.label;
      this.onChange = opts.onChange;
      this.selected = new Set();
      this.options = [];
      this.allCount = 0;
      this.isOpen = false;

      host.innerHTML = '';
      this.button = document.createElement('button');
      this.button.type = 'button';
      this.button.className = 'select multi-filter-btn';
      this.button.setAttribute('aria-haspopup', 'true');
      this.button.setAttribute('aria-expanded', 'false');
      this.button.setAttribute('aria-label', this.label);
      this.panel = document.createElement('div');
      this.panel.className = 'multi-filter-panel';
      this.panel.hidden = true;
      host.append(this.button, this.panel);

      this.button.addEventListener('click', () => (this.isOpen ? this.close() : this.open()));
      this.panel.addEventListener('change', (e) => {
        const box = e.target.closest('input[type="checkbox"]');
        if (!box) return;
        if (box.checked) this.selected.add(box.value); else this.selected.delete(box.value);
        this.emit();
      });
      this.panel.addEventListener('click', (e) => {
        if (e.target.closest('.multi-filter-clear')) {
          this.selected.clear();
          this.emit();
        }
      });
      document.addEventListener('mousedown', (e) => {
        if (this.isOpen && !this.host.contains(e.target)) this.close();
      });
      document.addEventListener('keydown', (e) => {
        if (!this.isOpen) return;
        if (e.key === 'Escape') {
          this.close();
          this.button.focus();
        } else if (e.key === 'Tab') {
          // Close once keyboard focus has actually left the filter.
          setTimeout(() => { if (!this.host.contains(document.activeElement)) this.close(); }, 0);
        }
      });
      window.addEventListener('resize', () => this.close());
      window.addEventListener('scroll', (e) => {
        if (this.isOpen && !this.panel.contains(e.target)) this.close();
      }, true);
      this.renderButton();
      this.renderPanel();
    }

    values() { return Array.from(this.selected); }

    selectedLabels() { return this.options.filter(o => this.selected.has(o.value)).map(o => o.label); }

    setOptions(options, allCount) {
      this.options = options;
      this.allCount = allCount;
      const keys = new Set(options.map(o => o.value));
      Array.from(this.selected).forEach(v => { if (!keys.has(v)) this.selected.delete(v); });
      this.renderButton();
      this.renderPanel();
    }

    setSelected(values) {
      this.selected = new Set(values);
      this.renderButton();
      this.renderPanel();
    }

    emit() {
      this.renderButton();
      this.renderPanel();
      this.onChange();
    }

    renderButton() {
      const n = this.selected.size;
      let text;
      if (n === 0) {
        text = `${this.allLabel} (${this.allCount.toLocaleString()})`;
      } else if (n === 1) {
        const opt = this.options.find(o => this.selected.has(o.value));
        text = opt ? `${opt.label} (${opt.count.toLocaleString()})` : this.values()[0];
      } else {
        const total = this.options.filter(o => this.selected.has(o.value)).reduce((sum, o) => sum + o.count, 0);
        text = `${n} selected (${total.toLocaleString()})`;
      }
      this.button.innerHTML = `<span class="multi-filter-text">${escapeHtml(text)}</span><span class="multi-filter-caret" aria-hidden="true">▾</span>`;
      this.button.classList.toggle('active', n > 0);
      if (n > 1) this.button.dataset.tip = this.selectedLabels().join('\n');
      else this.button.removeAttribute('data-tip');
    }

    renderPanel() {
      // Keep keyboard focus on the same option when the list is rebuilt.
      const active = document.activeElement;
      const focusVal = active && this.panel.contains(active) && active.type === 'checkbox' ? active.value : null;
      const n = this.selected.size;
      const rows = this.options.map(o =>
        `<label class="multi-filter-row"><input type="checkbox" value="${escapeHtml(o.value)}"${this.selected.has(o.value) ? ' checked' : ''}>`
        + `<span class="multi-filter-name">${escapeHtml(o.label)}</span>`
        + `<span class="multi-filter-count">${o.count.toLocaleString()}</span></label>`).join('');
      this.panel.innerHTML =
        `<div class="multi-filter-head"><span>${n ? `${n} selected` : `Showing all ${this.noun}`}</span>`
        + `${n ? '<button type="button" class="multi-filter-clear">Show all</button>' : ''}</div>`
        + `<div class="multi-filter-list">${rows || `<div class="multi-filter-empty">No ${this.noun} in this view</div>`}</div>`;
      if (focusVal !== null) {
        const box = Array.from(this.panel.querySelectorAll('input[type="checkbox"]')).find(i => i.value === focusVal);
        if (box) box.focus({ preventScroll: true });
      }
    }

    open() {
      this.isOpen = true;
      this.panel.hidden = false;
      this.button.setAttribute('aria-expanded', 'true');
      this.host.classList.add('open');
      this.position();
    }

    close() {
      if (!this.isOpen) return;
      this.isOpen = false;
      this.panel.hidden = true;
      this.button.setAttribute('aria-expanded', 'false');
      this.host.classList.remove('open');
    }

    position() {
      // Fixed, so the panel is not clipped by the toolbar or its card.
      const r = this.button.getBoundingClientRect();
      const vw = document.documentElement.clientWidth;
      const vh = document.documentElement.clientHeight;
      this.panel.style.minWidth = `${Math.round(Math.max(r.width, 230))}px`;
      const pw = this.panel.offsetWidth;
      const ph = this.panel.offsetHeight;
      const left = Math.max(8, Math.min(r.left, vw - pw - 8));
      let top = r.bottom + 4;
      if (top + ph > vh - 8 && r.top - 4 - ph >= 8) top = r.top - 4 - ph;
      this.panel.style.left = `${Math.round(left)}px`;
      this.panel.style.top = `${Math.round(top)}px`;
    }
  }

  let typeFilterUI = null;
  let modeFilterUI = null;
  function getTypeFilter() {
    if (!typeFilterUI) {
      typeFilterUI = new MultiSelectFilter(dom.typeFilter, {
        allLabel: 'All Types', noun: 'types', label: 'Filter Land Master by type',
        onChange: () => { populateModeFilter(state.records); filterAndRenderRecords(); },
      });
    }
    return typeFilterUI;
  }
  function getModeFilter() {
    if (!modeFilterUI) {
      modeFilterUI = new MultiSelectFilter(dom.modeFilterPanel, {
        allLabel: 'All Modes', noun: 'modes', label: 'Filter Land Master by mode',
        onChange: () => { populateTypeFilter(state.records); updateModeBadge(); filterAndRenderRecords(); },
      });
    }
    return modeFilterUI;
  }

  function setupRecordsView() {
    setupColumnPicker();

    dom.searchInput.addEventListener('input', () => {
      clearTimeout(searchDebounce);
      searchDebounce = setTimeout(filterAndRenderRecords, 120);
    });

    getTypeFilter();
    getModeFilter();

    if (dom.btnUpdateFromFinance) {
      dom.btnUpdateFromFinance.addEventListener('click', updateRecordsFromFinance);
    }
    dom.btnExportExcel.addEventListener('click', exportRecordsExcel);

    if (dom.btnClearFilters) {
      dom.btnClearFilters.addEventListener('click', () => {
        dom.searchInput.value = '';
        getTypeFilter().setSelected([]);
        getModeFilter().setSelected([]);
        updateModeBadge();
        state.columnFilters = {};
        state.columnFiltersOpen = {};
        populateTypeFilter(state.records);
        populateModeFilter(state.records);
        populateTypeFilter(state.records);
        populateModeFilter(state.records);
        renderTableHead();
        filterAndRenderRecords();
        showToast('All filters and search reset.', 'ok');
      });
    }

    if (dom.btnToggleShowOnlyChanged) {
      dom.btnToggleShowOnlyChanged.addEventListener('click', () => {
        state.showOnlyChanged = !state.showOnlyChanged;
        dom.btnToggleShowOnlyChanged.textContent = state.showOnlyChanged ? 'Show All' : 'Show Changed Only';
        dom.btnToggleShowOnlyChanged.className = state.showOnlyChanged ? 'btn xs primary' : 'btn xs ghost';
        filterAndRenderRecords();
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
            <button type="button" class="col-reorder-btn" data-move-up="${col.key}" ${isFirst ? 'disabled' : ''} title="Move left">◄</button>
            <button type="button" class="col-reorder-btn" data-move-down="${col.key}" ${isLast ? 'disabled' : ''} title="Move right">►</button>
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

      // Reloading Land Master also refreshes the reference tables from Dataverse,
      // but never replaces a file the user loaded or unsaved reference edits
      // (the explicit Sync Finance / Sync Conversion buttons still do that).
      const keepRows = { success: true, skipped: true };
      const finEdits = Object.keys(financeState.stagedEdits || {}).length > 0;
      const convEdits = Object.keys(conversionState.stagedEdits || {}).length > 0;
      const [res, finRes, convRes] = await Promise.all([
        window.pywebview.api.fetch_loa_records(),
        finEdits ? Promise.resolve(keepRows)
          : window.pywebview.api.sync_finance_from_dataverse(null, true).catch(e => ({ success: false, error: String(e) })),
        convEdits ? Promise.resolve(keepRows)
          : window.pywebview.api.sync_conversion_from_dataverse(null, true).catch(e => ({ success: false, error: String(e) }))
      ]);

      if (res.success) {
        state.records = res.records || [];
        state.baselineRecords = JSON.parse(JSON.stringify(state.records));
        state.stagedEdits = {};
        state.financeAutopopulatedRows = new Set();
        recalculateKPIs();

        populateTypeFilter(state.records);
        populateModeFilter(state.records);
        renderTableHead();
        filterAndRenderRecords();
        showToast(`Loaded ${res.count.toLocaleString()} records from Dataverse.`, 'ok');
        if (window.__markDataSynced) window.__markDataSynced();
      } else {
        showToast(res.error || 'Failed to fetch Dataverse records.', 'err');
      }

      if (finRes && finRes.success && !finRes.skipped) {
        const s = await window.pywebview.api.get_initial_state();
        applyState(s);
        const sample = finRes.sample_rows || [];
        renderFinancePreview(finRes.headers || (sample.length ? Object.keys(sample[0]) : []), sample);
      }

      if (convRes && convRes.success && !convRes.skipped) {
        const s = await window.pywebview.api.get_initial_state();
        applyState(s);
        const sample = convRes.sample_rows || [];
        renderConversionPreview(convRes.headers || (sample.length ? Object.keys(sample[0]) : []), sample);
      }
    } catch (err) {
      showToast('Error: ' + err, 'err');
    } finally {
      dom.btnTopSyncDv.disabled = false;
      updateTopRefreshAction();
    }
  }

  function normalizeModeFilterValue(value) {
    return String(value || '').trim().toUpperCase().replace(/\s*-\s*/g, '-').replace(/\s+/g, ' ');
  }

  function populateTypeFilter(records) {
    const modeSel = getModeFilter().values();
    const typeCounts = new Map();
    (records || []).forEach(r => {
      const rowMode = normalizeModeFilterValue(r['new_mode'] || r['Mode'] || r['MODE'] || r['cr63f_mode'] || '');
      if (modeSel.length && !modeSel.includes(rowMode)) return;
      const type = String(r['Type'] || r['TYPE'] || r['cr63f_type'] || '').trim().toUpperCase();
      if (type) typeCounts.set(type, (typeCounts.get(type) || 0) + 1);
    });
    const options = Array.from(typeCounts.keys()).sort().map(t => ({ value: t, label: t, count: typeCounts.get(t) }));
    const total = options.reduce((sum, o) => sum + o.count, 0);
    getTypeFilter().setOptions(options, total);

    if (dom.fType && dom.fType.tagName === 'SELECT') {
      setSelectOptions(dom.fType, TYPE_OPTIONS, dom.fType.value || 'LAND');
    }
  }

  function populateModeFilter(records) {
    const typeSel = getTypeFilter().values();
    const modeLabels = new Map();
    const modeCounts = new Map();
    const modeScopeRows = (records || []).filter(r => {
      const rowType = String(r['Type'] || r['TYPE'] || r['cr63f_type'] || '').trim().toUpperCase();
      return !typeSel.length || typeSel.includes(rowType);
    });
    MODES.forEach(mode => {
      const key = normalizeModeFilterValue(mode);
      modeLabels.set(key, mode);
      modeCounts.set(key, 0);
    });

    modeScopeRows.forEach(r => {
      const rawMode = String(r['new_mode'] || r['Mode'] || r['MODE'] || r['cr63f_mode'] || '').trim();
      const key = normalizeModeFilterValue(rawMode);
      if (!key) return;
      if (!modeLabels.has(key)) modeLabels.set(key, rawMode);
      modeCounts.set(key, (modeCounts.get(key) || 0) + 1);
    });

    const options = Array.from(modeLabels.keys())
      .filter(key => (modeCounts.get(key) || 0) > 0)
      .map(key => ({ value: key, label: modeLabels.get(key), count: modeCounts.get(key) }));
    getModeFilter().setOptions(options, modeScopeRows.length);
    updateModeBadge();
  }

  function updateModeBadge() {
    if (!dom.kpiModeActiveBadge) return;
    const labels = getModeFilter().selectedLabels();
    dom.kpiModeActiveBadge.textContent = labels.length === 0
      ? 'ALL MODES'
      : labels.length === 1 ? String(labels[0]).toUpperCase() : `${labels.length} MODES`;
  }

  function renderTableHead() {
    if (!dom.recordsTableHead) return;
    const visibleCols = getVisibleColumns();
    let html = '<tr>';

    visibleCols.forEach((col, colIdx) => {
      const width = (col.key === '__action' ? Math.max(80, state.colWidths[col.key] || 0) : (state.colWidths[col.key] || col.width || 120));
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
            <div class="th-bottom-resizer" title="Drag to resize the header"></div>
          </th>`;
        return;
      }

      html += `
        <th draggable="true" data-col-key="${col.key}" class="${frozenClass}" style="width: ${width}px; min-width: ${width}px; max-width: ${width}px; ${leftStyle}">
          <div class="th-header-box">
            <div class="th-title-wrap">
              <button type="button" class="th-pin-btn ${isFrozen ? 'pinned' : ''}" data-pin-col="${col.key}" aria-label="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(col.label)}" title="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(col.label)}">
                <svg viewBox="0 0 24 24" width="11" height="11" fill="currentColor" stroke="none">
                  <path d="M16 12V4h1V2H7v2h1v8l-2 2v2h5.2v6h1.6v-6H18v-2l-2-2z"/>
                </svg>
              </button>
              <button type="button" class="th-sort-wrap" data-sort-col="${col.key}" aria-label="Sort by ${escapeHtml(col.label)}" title="Sort by ${escapeHtml(col.label)} · drag to reorder">
                <span class="th-col-name">${escapeHtml(col.label)}</span>
                <span class="th-sort-icon ${isSorted ? 'active' : ''}" aria-hidden="true">${sortIcon}</span>
              </button>
              <button type="button" class="th-chevron-btn ${isFilterActive ? 'active' : ''}" data-filter-toggle="${col.key}" aria-label="Filter ${escapeHtml(col.label)}" title="Filter ${escapeHtml(col.label)}">
                <svg viewBox="0 0 24 24" width="10" height="10" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"></polyline></svg>
              </button>
            </div>
            <div class="th-filter-popup ${isFilterOpen ? '' : 'hidden'}" id="colFilterBox_${col.key}">
              <input class="th-filter-input" data-filter-col="${col.key}" aria-label="Filter ${escapeHtml(col.label)} column" autocomplete="off" value="${escapeHtml(state.columnFilters[col.key] || '')}" placeholder="Filter ${escapeHtml(col.label)}…" />
            </div>
          </div>
          <div class="th-resizer" data-col-key="${col.key}" title="Drag to resize · double-click to reset"></div>
          <div class="th-bottom-resizer" title="Drag to resize the header"></div>
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

    let recordsFilterDebounceTimer = null;
    dom.recordsTableHead.querySelectorAll('.th-filter-input').forEach(input => {
      input.addEventListener('input', () => {
        const colKey = input.dataset.filterCol;
        state.columnFilters[colKey] = input.value.trim().toLowerCase();
        if (recordsFilterDebounceTimer) clearTimeout(recordsFilterDebounceTimer);
        recordsFilterDebounceTimer = setTimeout(() => {
          filterAndRenderRecords();
        }, 120);
      });
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
          const colKey = input.dataset.filterCol;
          state.columnFiltersOpen[colKey] = false;
          const popup = document.getElementById(`colFilterBox_${colKey}`);
          if (popup) popup.classList.add('hidden');
        }
        if (e.key === 'Enter') e.preventDefault();
        e.stopPropagation();
      });
    });
  }

  function getRowCellValue(row, col) {
    if (!row) return '';
    if (col.key === 'work_week' || col.isWorkWeek) {
      const id = row.__id;
      const staged = (id && state.stagedEdits[id]) ? state.stagedEdits[id] : {};
      const endorsedDateVal = staged['cr63f_endorseddate'] ?? staged['Endorsed Date'] ?? row['cr63f_endorseddate'] ?? row['Endorsed Date'] ?? row['endorsed_date'] ?? '';
      return getWorkWeek(endorsedDateVal);
    }
    const id = row.__id;
    if (id && state.stagedEdits[id] && state.stagedEdits[id][col.key] !== undefined) {
      const val = state.stagedEdits[id][col.key];
      return col.isPercent ? formatPercent(val) : val;
    }
    if (col.isPercent && row.__display_percent_paid) {
      return formatPercent(row.__display_percent_paid);
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
    // Comma-separated search (decision 2026-10-14): "TST-ABC-000101, TST-248" shows the
    // records that match ANY of the terms. A single term works as before.
    const searchTerms = q.split(',').map(t => t.trim()).filter(Boolean);
    const typeSel = getTypeFilter().values();
    const modeSel = getModeFilter().values();
    const colFilterKeys = Object.keys(state.columnFilters).filter(k => !!state.columnFilters[k]);

    state.filteredRecords = state.records.filter(r => {
      if (state.showOnlyChanged) {
        const id = r.__id || '';
        const isChanged = Boolean(r.__finance_changed) ||
          Boolean(r.__conversion_changed) ||
          (id && state.financeAutopopulatedRows && state.financeAutopopulatedRows.has(id)) ||
          (id && state.stagedEdits && state.stagedEdits[id] && Object.keys(state.stagedEdits[id]).length > 0) ||
          (r.__changed_fields && r.__changed_fields.length > 0);
        if (!isChanged) return false;
      }

      const rType = String(r['TYPE'] || r['Type'] || r['cr63f_type'] || '').trim().toUpperCase();
      if (typeSel.length && !typeSel.includes(rType)) return false;

      const rMode = normalizeModeFilterValue(r['new_mode'] || r['Mode'] || r['MODE'] || r['cr63f_mode'] || '');
      if (modeSel.length && !modeSel.includes(rMode)) return false;

      for (const colKey of colFilterKeys) {
        const query = state.columnFilters[colKey];
        const colDef = DEFAULT_COLUMNS.find(c => c.key === colKey) || { key: colKey };
        const val = String(getRowCellValue(r, colDef)).toLowerCase();
        if (!val.includes(query)) return false;
      }

      if (!searchTerms.length) return true;
      const haystack = Object.values(r).map(v => String(v).toLowerCase()).join('\u0001');
      return searchTerms.some(t => haystack.includes(t));
    });

    if (state.sortCol && state.sortDir) {
      const colDef = DEFAULT_COLUMNS.find(c => c.key === state.sortCol) || { key: state.sortCol };
      const mult = state.sortDir === 'asc' ? 1 : -1;
      state.filteredRecords.sort((a, b) => {
        const valA = getRowCellValue(a, colDef);
        const valB = getRowCellValue(b, colDef);
        return compareTableValues(valA, valB, colDef, mult);
      });
    }

    dom.lblGridCount.textContent = `${state.filteredRecords.length.toLocaleString()} records`;

    if (dom.recordsTableHead && !dom.recordsTableHead.children.length) {
      renderTableHead();
    }

    recalculateKPIs();
    invalidateGridCache('recordsTable');

    if (!state.filteredRecords.length) {
      dom.recordsTableBody.innerHTML = '';
      if (dom.tableEmptyState) dom.tableEmptyState.classList.add('hidden');
      dom.recordsTable.classList.remove('hidden');
      state.renderedRecordsCount = 0;

      const visibleCols = getVisibleColumns();
      const colSpan = Math.max(1, visibleCols.length);
      const isInitialEmpty = !state.records.length;
      const title = isInitialEmpty ? 'No LOA records loaded' : 'No matching records found';
      const desc = isInitialEmpty
        ? 'Click <strong>Reload Land Master</strong> above or add a <strong>New Entry</strong>.'
        : 'Try adjusting your search terms, Type, Mode, or column filters.';

      const emptyTr = document.createElement('tr');
      emptyTr.className = 'empty-table-row';
      emptyTr.innerHTML = `
        <td colspan="${colSpan}" style="text-align: center; padding: 48px 16px; background: transparent; border: none;">
          <div style="display: inline-flex; flex-direction: column; align-items: center; justify-content: center; gap: 6px;">
            <svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" style="color: var(--ink3); opacity: 0.7;">
              <circle cx="11" cy="11" r="8"></circle>
              <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
            </svg>
            <p style="font-size: 13.5px; font-weight: 600; color: var(--ink); margin: 4px 0 2px;">${title}</p>
            <p style="font-size: 12px; color: var(--ink3); margin: 0;">${desc}</p>
          </div>
        </td>
      `;
      dom.recordsTableBody.appendChild(emptyTr);
      return;
    }

    if (dom.tableEmptyState) dom.tableEmptyState.classList.add('hidden');
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

      const isFinancePopulated = Boolean(r.__finance_changed) ||
        (id && state.financeAutopopulatedRows && state.financeAutopopulatedRows.has(id));

      if (isFinancePopulated) {
        tr.classList.add('row-finance-populated');
      }

      visibleCols.forEach((col, colIdx) => {
        const td = document.createElement('td');
        td.dataset.colKey = col.key;
        td.dataset.rowIdx = idx;
        td.dataset.colIdx = colIdx;
        td.dataset.colKey = col.key;
        td.dataset.rowId = id;
        td.dataset.fieldName = col.label || col.key;
        td.dataset.lotId = String(r.cr63f_areaindex || r['Area-Index'] || r.cr63f_control1 || r['Control Number'] || id);

        const width = (col.key === '__action' ? Math.max(80, state.colWidths[col.key] || 0) : (state.colWidths[col.key] || col.width || 120));
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

        const checkKeys = [col.key, col.alt, col.label].filter(Boolean);
        const rowRefSet = (id && state.referenceUpdatedFieldsByRowId && state.referenceUpdatedFieldsByRowId[id]) ? state.referenceUpdatedFieldsByRowId[id] : null;
        const rowChangedArr = Array.isArray(r.__changed_fields) ? r.__changed_fields : [];
        const isRefUpdated = checkKeys.some(k => (rowRefSet && rowRefSet.has(k)) || rowChangedArr.includes(k));
        if (isRefUpdated) {
          td.classList.add('cell-reference-updated');
        }

        // Attach diff data attributes for hover tooltip
        if (isRefUpdated || isDirty) {
          const baseRec = (state.baselineRecords && state.baselineRecords.find(b => b.__id === id)) || {};
          const prevRaw = baseRec[col.key] ?? (col.alt ? baseRec[col.alt] : '') ?? '';
          const currRaw = (state.stagedEdits[id] && state.stagedEdits[id][col.key] !== undefined)
            ? state.stagedEdits[id][col.key]
            : (r[col.key] ?? (col.alt ? r[col.alt] : '') ?? '');
          const areaIdx = r.cr63f_areaindex || r['Area-Index'] || id;
          td.dataset.prevValue = String(prevRaw);
          td.dataset.currValue = String(currRaw);
          td.dataset.fieldName = col.label || col.key;
          td.dataset.diffType = isRefUpdated ? 'ref' : 'manual';
          td.dataset.lotId = String(areaIdx);
        }

        if (col.key === '__idx') {
          td.style.fontFamily = 'var(--mono)';
          td.style.color = 'var(--ink3)';
          td.style.textAlign = 'center';
          td.textContent = String(idx + 1);
          td.title = 'Double-click to edit this account';
          td.addEventListener('dblclick', () => openEditEntry(id));
        } else if (col.key === '__action') {
          td.className = `${td.className} col-action`.trim();
          td.innerHTML = `
            <button type="button" class="btn xs ghost rc-edit-btn" onclick="window.openEditEntry('${id}')" aria-label="Edit account" title="Edit account">
              <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M12 20h9"></path>
                <path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"></path>
              </svg>
            </button>
            <button type="button" class="btn xs ghost danger rc-del-btn" onclick="window.deleteRecord('${id}')" aria-label="Delete Land Master record" title="Delete record">
              <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M4 7h16"></path>
                <path d="M9 7V4h6v3"></path>
                <path d="M6 7l1 13a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-13"></path>
                <path d="M10 11v6"></path>
                <path d="M14 11v6"></path>
              </svg>
            </button>
          `;
        } else if (col.key === 'work_week' || col.isWorkWeek) {
          const wwVal = String(getRowCellValue(r, col) ?? '');
          td.style.textAlign = 'center';
          if (wwVal) {
            td.innerHTML = `<span class="ww-pill">${escapeHtml(wwVal)}</span>`;
          } else {
            td.innerHTML = `<span style="color: var(--ink3); font-size: 11px;">-</span>`;
          }
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
              const label = col.key === 'cr63f_acquisitionsublevel'
                ? `${opt} — ${ACQUISITION_SUBLEVEL_STATUS[opt] || ''}`
                : opt;
              optHtml += `<option value="${escapeHtml(opt)}" ${sel}>${escapeHtml(label)}</option>`;
            });
            // A stored value that is not in the list stays visible (same rule as
            // the drawer), so the grid never shows it as blank.
            if (val.trim() && !col.options.some(opt => opt.trim().toUpperCase() === val.trim().toUpperCase())) {
              optHtml += `<option value="${escapeHtml(val)}" selected>${escapeHtml(val)} (not in list)</option>`;
            }
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
                  input.value = isBlankAmount(input.value) ? '' : formatAmountWithCommas(pNum);
                }
              });
              if (col.isPrice) {
                input.addEventListener('input', () => {
                  const raw = input.value;
                  const cursorPos = input.selectionStart;
                  const oldLen = raw.length;
                  if (/[0-9]/.test(raw)) {
                    const formatted = formatAmountWithCommas(raw);
                    input.value = formatted;
                    const newLen = formatted.length;
                    const newPos = Math.max(0, cursorPos + (newLen - oldLen));
                    try { input.setSelectionRange(newPos, newPos); } catch (_) {}
                  }
                });
              }
              input.addEventListener('input', () => {
                updateCellChangePreview(id, col, input.value, input, td);
              });
              input.addEventListener('blur', () => {
                if (col.isPrice) {
                  const num = parseAmount(input.value);
                  const typedBlank = isBlankAmount(input.value);
                  input.value = typedBlank ? '' : formatPrice(num);
                  const cleanStr = typedBlank ? '' : String(num);
                  const prevNum = parseAmount(before);
                  const prevStr = isBlankAmount(before) ? '' : String(prevNum);
                  if (cleanStr !== prevStr) {
                    handleCellEdit(id, col.key, cleanStr, input, td);
                  }
                } else if (col.isDate) {
                  const normalized = commitDateInput(input, before);
                  if (normalized !== null && normalized !== (normalizeDateInput(before) ?? before)) {
                    handleCellEdit(id, col.key, normalized, input, td);
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

  const STAGED_PAYMENT_KEYS = ['cr63f_totalcontractprice', 'Total Contract Price', 'cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP'];
  const STAGED_DERIVED_KEYS = ['cr63f_balance', 'Balance', 'cr63f_paid', '% Paid'];

  // Balance and % Paid for one grid row, from its staged or stored TCP and Deductible.
  // A blank TCP cannot be calculated, so both are staged blank and cleared on save.
  // cr63f_paid keeps one decimal, as Dataverse stores it.
  function recalcStagedPayment(rowId, rec, tdEl) {
    if (!state.stagedEdits[rowId]) state.stagedEdits[rowId] = {};
    const staged = state.stagedEdits[rowId];
    const tcpRaw = staged['cr63f_totalcontractprice'] ?? rec['cr63f_totalcontractprice'] ?? rec['Total Contract Price'] ?? '';
    const dedRaw = staged['cr63f_totalpaymentdeductibletotcp'] ?? rec['cr63f_totalpaymentdeductibletotcp'] ?? rec['Total Payment Deductible to TCP'] ?? '';
    const tcpBlank = String(tcpRaw).trim() === '';
    const tcpNum = parseAmount(tcpRaw);
    const dedNum = parseAmount(dedRaw);
    let balFmt = '';
    let pctFmt = '';
    if (!tcpBlank) {
      balFmt = Math.max(0, tcpNum - dedNum).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      if (tcpNum > 0) pctFmt = ((dedNum / tcpNum) * 100).toFixed(1) + '%';
      else pctFmt = dedNum > 0 ? '100.0%' : '0.0%';
    }
    staged['cr63f_balance'] = balFmt;
    staged['Balance'] = balFmt;
    staged['cr63f_paid'] = pctFmt;
    staged['% Paid'] = pctFmt;

    if (tdEl) {
      const tr = tdEl.closest('tr');
      if (tr) {
        const balInput = tr.querySelector('[data-col-key="cr63f_balance"] input');
        if (balInput) { balInput.value = balFmt ? formatPrice(balFmt) : ''; balInput.classList.add('modified'); }
        const pctInput = tr.querySelector('[data-col-key="cr63f_paid"] input');
        if (pctInput) { pctInput.value = pctFmt; pctInput.classList.add('modified'); }
      }
    }
  }

  const REQUIRED_EXISTING_KEYS = ['cr63f_type', 'new_mode'];

  async function handleCellEdit(rowId, colKey, newVal, inputEl, tdEl) {

    const rec = state.records.find(r => r.__id === rowId);
    if (!rec) return;

    const colDef = DEFAULT_COLUMNS.find(c => c.key === colKey) || { key: colKey };
    const origVal = String(rec[colDef.key] || (colDef.alt ? rec[colDef.alt] : '') || '').trim();
    let trimmed = String(newVal || '').trim();
    if (colDef.isPercent) trimmed = formatPercent(trimmed);
    const sameValue = colDef.isDate
      ? (normalizeDateInput(origVal) ?? origVal) === (normalizeDateInput(trimmed) ?? trimmed)
      : origVal === trimmed;

    // Type and Mode cannot be cleared on an existing account (the server refuses it too).
    if (REQUIRED_EXISTING_KEYS.includes(colDef.key) && !trimmed && origVal) {
      showToast(`${colDef.label || colKey} cannot be cleared on an existing account (stored value: ${origVal}).`, 'err');
      if (inputEl) inputEl.value = origVal;
      return;
    }

    if (sameValue) {
      if (state.stagedEdits[rowId]) {
        delete state.stagedEdits[rowId][colKey];
        if (colDef.alt) delete state.stagedEdits[rowId][colDef.alt];
        if (STAGED_PAYMENT_KEYS.includes(colKey)) {
          // Balance and % Paid were staged from the old amount: drop them, then
          // recalculate from what is still staged (if the other amount is staged).
          STAGED_DERIVED_KEYS.forEach(k => delete state.stagedEdits[rowId][k]);
          const otherAmountStaged = STAGED_PAYMENT_KEYS.some(k => state.stagedEdits[rowId] && k in state.stagedEdits[rowId]);
          if (otherAmountStaged) recalcStagedPayment(rowId, rec, tdEl);
          if (tdEl && !otherAmountStaged) {
            const tr = tdEl.closest('tr');
            if (tr) {
              const balInput = tr.querySelector('[data-col-key="cr63f_balance"] input');
              if (balInput) { balInput.value = formatPrice(rec['cr63f_balance'] ?? rec['Balance'] ?? ''); balInput.classList.remove('modified'); }
              const pctInput = tr.querySelector('[data-col-key="cr63f_paid"] input');
              if (pctInput) { pctInput.value = rec['cr63f_paid'] ?? rec['% Paid'] ?? ''; pctInput.classList.remove('modified'); }
            }
          }
        }
        if (Object.keys(state.stagedEdits[rowId]).length === 0) delete state.stagedEdits[rowId];
      }
      if (tdEl && !tdEl.classList.contains('cell-reference-updated')) tdEl.classList.remove('is-dirty');
      if (inputEl) inputEl.classList.remove('modified');
      updateStagedEditsUI();
      return;
    }

    state.stagedEdits[rowId] = state.stagedEdits[rowId] || {};
    state.stagedEdits[rowId][colKey] = trimmed;
    if (colDef.alt) state.stagedEdits[rowId][colDef.alt] = trimmed;

    if (tdEl) tdEl.classList.add('is-dirty');
    if (inputEl) inputEl.classList.add('modified');
    if (tdEl) updateCellChangePreview(rowId, colDef, trimmed, inputEl, tdEl);
    updateStagedEditsUI();

    // Live formula calculation trigger
    const formulaKeys = ['cr63f_type', 'cr63f_indexno', 'cr63f_areaindex', 'cr63f_titleno', 'new_mode', 'Type', 'Index No', 'Area-Index', 'Title No', 'Mode'];
    if (formulaKeys.includes(colKey)) {
      const curType = state.stagedEdits[rowId]['cr63f_type'] || rec['cr63f_type'] || rec['Type'] || '';
      const curIdx = state.stagedEdits[rowId]['cr63f_indexno'] || rec['cr63f_indexno'] || rec['Index No'] || '';
      const curArea = state.stagedEdits[rowId]['cr63f_areaindex'] || rec['cr63f_areaindex'] || rec['Area-Index'] || '';
      const curTitle = state.stagedEdits[rowId]['cr63f_titleno'] || rec['cr63f_titleno'] || rec['Title No'] || '';

      try {
        // Mode picks the SALE or LEASE Finance row (TYPE-aware matching).
        const curMode = state.stagedEdits[rowId]['new_mode'] || rec['new_mode'] || rec['Mode'] || '';
        const lk = await window.pywebview.api.lookup_finance(curType, curIdx, curArea, curTitle, curMode);
        if (lk.matched) {
          const tcpFmt = lk.total_contract_price_formatted || '0.00';
          const dedFmt = lk.total_payment_deductible_formatted || '0.00';
          const balFmt = lk.balance_formatted || '0.00';
          const formattedPct = storedPercentText(lk.percent_paid_formatted || lk.percent_paid || '0');

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
    if (STAGED_PAYMENT_KEYS.includes(colKey)) {
      recalcStagedPayment(rowId, rec, tdEl);
    }

    // Direct manual edits to Endorsed Date -> live recalculation of Work Week (WW-NN)
    const endorsedKeys = ['cr63f_endorseddate', 'Endorsed Date', 'endorsed_date'];
    if (endorsedKeys.includes(colKey)) {
      const ww = getWorkWeek(trimmed);
      state.stagedEdits[rowId]['work_week'] = ww;
      state.stagedEdits[rowId]['Work Week'] = ww;

      if (tdEl) {
        const tr = tdEl.closest('tr');
        if (tr) {
          const wwCell = tr.querySelector('[data-col-key="work_week"]');
          if (wwCell) {
            if (ww) {
              wwCell.innerHTML = `<span class="ww-pill">${escapeHtml(ww)}</span>`;
            } else {
              wwCell.innerHTML = `<span style="color: var(--ink3); font-size: 11px;">-</span>`;
            }
            wwCell.classList.add('is-dirty');
          }
        }
      }
    }

    recalculateKPIs();
  }

  // Exact Mode lists per payment group, same table as app.py PAYMENT_GROUP_OF_MODE
  // (tests/test_mode_group_parity.py). Any other Mode, e.g. OPEN TO SALE or OPEN TO
  // LEASE, is OTHER: not a secured sale or lease, so not in % Paid or the balance.
  const PAYMENT_GROUP_OF_MODE = {
    'SECURED SALE': 'DIRECT_SALE', 'SS - FROM LAND': 'DIRECT_SALE', 'SS - FROM ROW': 'DIRECT_SALE',
    'SECURED LEASE': 'LEASE_DIRECT', 'SL - FROM LAND': 'LEASE_DIRECT',
    'SS - STP': 'STP_SALE', 'SS - FROM LAND STP': 'STP_SALE', 'SOLD TO PARTNER': 'STP_SALE',
    'SL - STP': 'LEASE_STP', 'SL - FROM LAND STP': 'LEASE_STP',
    'SECURED SALE NON SP': 'NON_SP', 'SECURED LEASE NON SP': 'NON_SP',
    'RESCINDED': 'RESCINDED', 'SS - RESCINDED': 'RESCINDED', 'SL - RESCINDED': 'RESCINDED',
  };

  function normalizePaymentMode(mode) {
    return String(mode ?? '').toUpperCase().replace(/[\u2013\u2014]/g, '-')
      .replace(/\bNON\s*-?\s*SP\b/g, 'NON SP')
      .replace(/\s*-\s*/g, ' - ')
      .replace(/\s+/g, ' ').trim()
      .replace(' - NON SP', ' NON SP');
  }

  function paymentGroupOf(mode) {
    return PAYMENT_GROUP_OF_MODE[normalizePaymentMode(mode)] || 'OTHER';
  }

  function recalculateKPIs() {
    let sumTcp = 0;
    let sumDed = 0;
    let sumBal = 0;
    let sumHa = 0;
    let noArea = 0;
    const shownRecords = state.filteredRecords !== undefined ? state.filteredRecords : state.records;
    const stagedOf = (r) => (state.stagedEdits && r.__id && state.stagedEdits[r.__id]) ? state.stagedEdits[r.__id] : {};
    // Same groups as the Overview (app.py _payment_group). Decision 2026-10-14:
    // records, contract price and payments count every account shown (payments made
    // before turnover to partner are still money paid); Outstanding Balance counts only
    // what the company still owes: direct sale + secured lease.
    const groupOf = (r) => {
      const staged = stagedOf(r);
      const group = paymentGroupOf(staged['new_mode'] ?? staged['Mode'] ?? r['new_mode'] ?? r['Mode'] ?? r['MODE'] ?? r['cr63f_mode'] ?? '');
      return group === 'LEASE_DIRECT' ? 'LEASE' : group;
    };
    const BALANCE_GROUPS = ['DIRECT_SALE', 'LEASE'];
    const groups = {};

    shownRecords.forEach(r => {
      const g = groupOf(r);
      const staged = stagedOf(r);
      const tcp = parseAmount(staged['Total Contract Price'] ?? staged['cr63f_totalcontractprice'] ?? r['Total Contract Price'] ?? r['cr63f_totalcontractprice']);
      const ded = parseAmount(staged['Total Payment Deductible to TCP'] ?? staged['cr63f_totalpaymentdeductibletotcp'] ?? r['Total Payment Deductible to TCP'] ?? r['cr63f_totalpaymentdeductibletotcp']);
      // Same rule as the Overview: each account's own balance, clamped at 0,
      // calculated from TCP and Deductible (a stale stored Balance is ignored).
      const bal = Math.max(0, tcp - ded);
      const areaRaw = String(staged['cr63f_lotareaha'] ?? staged['Lot Area (Ha)'] ?? r['cr63f_lotareaha'] ?? r['Lot Area (Ha)'] ?? r['Lot Area (ha)'] ?? r['Lot Area'] ?? '').trim();
      const ha = areaRaw ? parseAmount(areaRaw) : 0;
      if (!areaRaw || !ha) noArea++;
      sumHa += ha;
      sumTcp += tcp;
      sumDed += ded;
      if (BALANCE_GROUPS.includes(g)) sumBal += bal;
      const t = groups[g] || (groups[g] = { count: 0, tcp: 0, ded: 0, bal: 0, ha: 0 });
      t.count++; t.tcp += tcp; t.ded += ded; t.bal += bal; t.ha += ha;
    });

    if (dom.kpiRecordCount) dom.kpiRecordCount.textContent = shownRecords.length.toLocaleString();
    if (dom.kpiTotalTcp) dom.kpiTotalTcp.textContent = formatPrice(sumTcp) || '₱0.00';
    if (dom.kpiTotalDeductible) dom.kpiTotalDeductible.textContent = formatPrice(sumDed) || '₱0.00';
    if (dom.kpiTotalBalance) dom.kpiTotalBalance.textContent = formatPrice(sumBal) || '₱0.00';
    const hectares = (v) => `${v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 4 })} ha`;
    if (dom.kpiTotalHectares) dom.kpiTotalHectares.textContent = hectares(sumHa);
    if (dom.kpiHectaresNote) {
      dom.kpiHectaresNote.textContent = noArea
        ? `${noArea.toLocaleString()} account${noArea === 1 ? '' : 's'} with no lot area`
        : 'Every account shown';
    }

    const GROUP_LABELS = [
      ['DIRECT_SALE', 'Secured sale'],
      ['LEASE', 'Secured lease'],
      ['STP_SALE', 'Turned over – sale'],
      ['LEASE_STP', 'Turned over – lease'],
      ['NON_SP', 'Non-SP'],
      ['RESCINDED', 'Rescinded'],
      ['OTHER', 'No sale or lease mode']
    ];
    const present = GROUP_LABELS.filter(([g]) => groups[g] && groups[g].count);
    const money = (v) => formatPrice(v) || '₱0.00';
    const breakdown = (key, total, counted, footer) => present.map(([g, label]) =>
      `${label}: ${money(groups[g][key])}${counted.includes(g) ? '' : ' (not counted)'}`)
      .concat([`Total: ${money(total)}`]).concat(footer ? ['', footer] : []).join('\n');
    const setTip = (el, text) => { const card = el && el.closest('.kpi-card'); if (card) card.title = text; };
    const all = present.map(([g]) => g);
    setTip(dom.kpiRecordCount, present.map(([g, label]) => `${label}: ${groups[g].count.toLocaleString()}`)
      .concat([`Total: ${shownRecords.length.toLocaleString()}`]).join('\n'));
    setTip(dom.kpiTotalHectares, present.map(([g, label]) => `${label}: ${hectares(groups[g].ha)}`)
      .concat([`Total: ${hectares(sumHa)}`])
      .concat(noArea ? ['', `${noArea.toLocaleString()} account${noArea === 1 ? ' has' : 's have'} no lot area.`] : []).join('\n'));
    setTip(dom.kpiTotalTcp, breakdown('tcp', sumTcp, all, ''));
    setTip(dom.kpiTotalDeductible, breakdown('ded', sumDed, all, 'Includes payments made before turnover to partner.'));
    setTip(dom.kpiTotalBalance, breakdown('bal', sumBal, BALANCE_GROUPS,
      'Grey rows are not counted: only secured sale and secured lease balances are yours to pay.'));

    const exclNote = document.getElementById('kpiExcludedNote');
    if (exclNote) {
      const turned = (groups.STP_SALE?.count || 0) + (groups.LEASE_STP?.count || 0);
      const parts = [];
      if (turned) parts.push(`${turned.toLocaleString()} turned over to partner`);
      if (groups.NON_SP?.count) parts.push(`${groups.NON_SP.count.toLocaleString()} non-SP`);
      if (groups.RESCINDED?.count) parts.push(`${groups.RESCINDED.count.toLocaleString()} rescinded`);
      if (groups.OTHER?.count) parts.push(`${groups.OTHER.count.toLocaleString()} with no sale/lease Mode`);
      exclNote.textContent = parts.length ? `Balance excludes ${parts.join(' and ')}` : 'Balance: secured sale and secured lease';
    }

    updateModeBadge();
  }

  function updateStagedEditsUI() {
    const stagedCount = Object.keys(state.stagedEdits).length;
    const refCount = state.financeAutopopulatedRows ? state.financeAutopopulatedRows.size : 0;
    const totalCount = Math.max(stagedCount, refCount);
    if (totalCount > 0) {
      dom.stagedEditsWrap.classList.remove('hidden');
      dom.unsavedEditsChip.textContent = `${totalCount.toLocaleString()} modified row${totalCount > 1 ? 's' : ''}`;
      if (dom.btnToggleShowOnlyChanged) {
        dom.btnToggleShowOnlyChanged.textContent = state.showOnlyChanged ? 'Show All' : 'Show Changed Only';
        dom.btnToggleShowOnlyChanged.className = state.showOnlyChanged ? 'btn xs primary' : 'btn xs ghost';
      }
    } else {
      dom.stagedEditsWrap.classList.add('hidden');
      state.showOnlyChanged = false;
    }
  }

  async function saveStagedEdits() {
    const rowIds = Object.keys(state.stagedEdits);
    const hasRefUpdates = state.financeAutopopulatedRows && state.financeAutopopulatedRows.size > 0;
    const refCount = hasRefUpdates ? state.financeAutopopulatedRows.size : 0;
    const totalCount = Math.max(rowIds.length, refCount);

    if (totalCount === 0) {
      showToast('No unsaved changes to save.', 'info');
      return;
    }

    const stagedList = rowIds.map(id => {
      const fields = { ...state.stagedEdits[id] };
      const actionType = fields.__action_type || 'MANUAL_EDIT';
      delete fields.__action_type;
      delete fields.__action_types;
      const referenceFields = state.referenceUpdatedFieldsByRowId?.[id];
      const referenceRecord = state.records.find(record => String(record.__id) === String(id));
      if (referenceFields && referenceRecord) {
        for (const key of referenceFields) {
          const stagedValue = fields[key];
          const referenceValue = referenceRecord[key];
          const numericReferenceField = /price|deductible|balance|paid/i.test(key);
          const sameValue = referenceValue !== undefined && (numericReferenceField
            ? Number.isFinite(parseAmount(stagedValue))
              && Number.isFinite(parseAmount(referenceValue))
              && Math.abs(parseAmount(stagedValue) - parseAmount(referenceValue)) < 0.005
            : String(stagedValue ?? '').trim() === String(referenceValue ?? '').trim());
          if (Object.prototype.hasOwnProperty.call(fields, key) && sameValue) {
            delete fields[key];
          }
        }
      }
      return { id, fields, action_type: actionType };
    }).filter(item => Object.keys(item.fields).length > 0);

    const confirmed = await showConfirmModal({
      title: 'Save Changes to Dataverse',
      subtitle: 'Dataverse Sync & Audit Logging',
      message: `You are about to save changes for ${totalCount.toLocaleString()} modified record(s) to Dataverse. All updated values will be recorded in the Land Master DB and logged in the Entry Logs audit trail. Previous and updated values can be inspected at any time by hovering over any highlighted cell in the table.`,
      confirmText: 'Save to Dataverse & Log',
      cancelText: 'Cancel',
      isDanger: false,
      keepOpenOnConfirm: true
    });

    if (!confirmed) return;

    try {
      if (dom.btnConfirmModalOk) {
        dom.btnConfirmModalOk.disabled = true;
        dom.btnConfirmModalOk.textContent = 'Saving...';
      }
      if (dom.btnConfirmModalCancel) {
        dom.btnConfirmModalCancel.classList.add('hidden');
      }
      if (dom.btnConfirmModalClose) {
        dom.btnConfirmModalClose.classList.add('hidden');
      }

      setConfirmModalProgress(0, totalCount, `Preparing to save ${totalCount.toLocaleString()} record(s)...`);

      dom.btnSaveStagedEdits.disabled = true;
      dom.btnSaveStagedEdits.textContent = 'Saving...';

      let totalSaved = 0;

      // 1. If reference updates are staged, save them to Dataverse
      if (hasRefUpdates) {
        const baseList = (state.baselineRecords && state.baselineRecords.length) ? state.baselineRecords : state.records;
        const refRes = await window.pywebview.api.populate_loa_records_from_references(baseList, true);
        if (refRes && refRes.success) {
          totalSaved += (refRes.saved_count ?? 0);
          state.records = refRes.records || state.records;
        } else if (refRes && !refRes.success) {
          // Accounts that did save are kept; saving again only retries the others.
          const savedPart = refRes.saved_count ? `${refRes.saved_count} account(s) saved; ` : '';
          const failedPart = (refRes.errors && refRes.errors.length) ? `${refRes.errors.length} failed: ` : '';
          throw new Error(savedPart + failedPart + (refRes.error || 'Failed to save reference updates.'));
        }
      }

      // 2. If manual cell edits are staged, save them to Dataverse
      if (stagedList.length > 0) {
        const res = await window.pywebview.api.save_staged_records(stagedList);
        if (res && res.success) {
          totalSaved += (res.saved_count ?? 0);
          stagedList.forEach(item => {
            const rec = state.records.find(r => r.__id === item.id);
            if (rec) {
              Object.assign(rec, item.fields);
            }
          });
        } else if (res && !res.success) {
          throw new Error(res.error || 'Failed to save manual edits.');
        }
      }

      closeConfirmModal(false);

      showToast(`Successfully saved ${totalSaved.toLocaleString()} record(s) to Dataverse and logged changes.`, 'ok');

      // Clear all highlight-causing flags from records so the table renders clean
      state.records.forEach(rec => {
        delete rec.__finance_changed;
        delete rec.__changed_fields;
        delete rec.__ref_updated;
      });

      state.baselineRecords = JSON.parse(JSON.stringify(state.records));
      state.stagedEdits = {};
      state.financeAutopopulatedRows = new Set();
      state.referenceUpdatedFieldsByRowId = {};
      state.pendingRefDiffSummary = null;
      state.showOnlyChanged = false;
      if (dom.btnToggleShowOnlyChanged) {
        dom.btnToggleShowOnlyChanged.textContent = 'Show Changed Only';
        dom.btnToggleShowOnlyChanged.className = 'btn xs ghost';
      }
      recalculateKPIs();
      updateStagedEditsUI();
      filterAndRenderRecords();

      if (typeof fetchEntryLogs === 'function') {
        fetchEntryLogs();
      }
    } catch (err) {
      closeConfirmModal(false);
      showToast('Save error: ' + (err.message || err), 'err');
    } finally {
      dom.btnSaveStagedEdits.disabled = false;
      dom.btnSaveStagedEdits.textContent = 'Save to Dataverse';
    }
  }

  async function discardStagedEdits() {
    const stagedCount = Object.keys(state.stagedEdits).length;
    const refCount = state.financeAutopopulatedRows ? state.financeAutopopulatedRows.size : 0;
    const rowCount = Math.max(stagedCount, refCount);
    if (rowCount > 0) {
      const ok = await showConfirmModal({
        title: 'Discard Unsaved Edits',
        subtitle: 'Revert Staged Changes',
        message: `Are you sure you want to discard all staged changes for ${rowCount.toLocaleString()} record(s)? This will revert all modified fields back to their previous values.`,
        confirmText: 'Discard Changes',
        cancelText: 'Cancel',
        isDanger: true
      });
      if (!ok) return;
    }
    if (state.baselineRecords && state.baselineRecords.length) {
      state.records = JSON.parse(JSON.stringify(state.baselineRecords));
    }
    // Clear all highlight-causing flags from records
    state.records.forEach(rec => {
      delete rec.__finance_changed;
      delete rec.__changed_fields;
      delete rec.__ref_updated;
    });
    state.stagedEdits = {};
    state.financeAutopopulatedRows = new Set();
    state.referenceUpdatedFieldsByRowId = {};
    state.pendingRefDiffSummary = null;
    state.showOnlyChanged = false;
    if (dom.btnToggleShowOnlyChanged) {
      dom.btnToggleShowOnlyChanged.textContent = 'Show Changed Only';
      dom.btnToggleShowOnlyChanged.className = 'btn xs ghost';
    }
    recalculateKPIs();
    updateStagedEditsUI();
    filterAndRenderRecords();
    showToast('Unsaved edits discarded.', 'ok');
  }

  window.deleteRecord = async function (id) {
    const ok = await showConfirmModal({
      title: 'Delete LOA Record',
      subtitle: 'Dataverse Permanent Deletion',
      message: 'Delete this account from Dataverse? A DELETE row with its last values is written to Entry Logs first. This cannot be undone with Revert.',
      confirmText: 'Delete',
      cancelText: 'Cancel',
      isDanger: true
    });
    if (!ok) return;

    try {
      const res = await window.pywebview.api.delete_loa_record_entry(id);
      if (res.success) {
        showToast('Record deleted.', 'ok');
        state.records = state.records.filter(r => r.__id !== id);
        if (state.baselineRecords) {
          state.baselineRecords = state.baselineRecords.filter(r => r.__id !== id);
        }
        delete state.stagedEdits[id];
        recalculateKPIs();
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
    const toExport = state.filteredRecords;
    if (!toExport.length) {
      showToast('No matching records to export.', 'err');
      return;
    }
    const cols = (state.colOrder.length ? state.colOrder : DEFAULT_COLUMNS.map(c => c.key))
      .filter(k => state.visibleCols.has(k) && !k.startsWith('__') && !k.startsWith('_'))
      .map(k => {
        // Values are read by the Dataverse name (key); the label is only the heading.
        // Reading by label left columns blank where the grid label differs from the
        // Dataverse display name, e.g. Payment Deductible (fixed 2026-10-15).
        const c = DEFAULT_COLUMNS.find(dc => dc.key === k);
        return { key: k, label: (c && c.label) ? c.label : k };
      });
    await runExport(() => window.pywebview.api.export_records_excel(toExport, 'LOA_Monitoring_Export.xlsx', cols));
  }

  // One result handler for every Excel export: a failure (file open in Excel, folder
  // not writable) used to show nothing; a cancelled save dialog stays silent.
  async function runExport(call) {
    try {
      const res = await call();
      if (res && res.success) showToast(`Exported to ${res.file_path}`, 'ok');
      else if (!(res && res.cancelled)) showToast((res && res.error) || 'The export could not be saved.', 'err');
    } catch (err) {
      showToast('The export could not be saved: ' + (err && err.message ? err.message : err), 'err');
    }
  }

  async function updateRecordsFromFinance() {
    if (!state.records || !state.records.length) {
      showToast('No LOA records loaded in the hub to update.', 'err');
      return;
    }
    const hasFinance = (financeState.rows && financeState.rows.length) || state.financeLoadedCount > 0;
    const hasConversion = (conversionState.rows && conversionState.rows.length) || state.conversionLoadedCount > 0;
    if (!hasFinance && !hasConversion) {
      showToast('No Reference Data loaded. Please upload or sync Finance or Conversion Data first.', 'err');
      return;
    }

    try {
      if (dom.btnUpdateFromFinance) {
        dom.btnUpdateFromFinance.disabled = true;
        dom.btnUpdateFromFinance.textContent = 'Checking...';
      }

      // Step 1: Preview matching and changed counts using baseline data
      const baseList = (state.baselineRecords && state.baselineRecords.length) ? state.baselineRecords : state.records;
      const previewRes = await window.pywebview.api.populate_loa_records_from_references(baseList, false);
      if (!previewRes.success) {
        showToast(previewRes.error || 'Failed to analyze records from Reference Data.', 'err');
        return;
      }

      const fCount = previewRes.matched_count ?? 0;
      const uCount = previewRes.updated_count ?? 0;

      // Accounts whose reference key matched conflicting duplicate rows are skipped, not guessed.
      if (previewRes.warning) {
        showToast(previewRes.warning, 'err');
      }

      if (uCount === 0) {
        showToast(`All ${fCount} matched records are already up to date with Reference Data (0 changes found).`, 'info');
        return;
      }

      // Step 2: Immediately stage and highlight changes on the table
      state.records = previewRes.records || [];
      state.financeAutopopulatedRows = new Set();
      state.referenceUpdatedFieldsByRowId = {};
      state.pendingRefDiffSummary = previewRes.diff_summary || [];

      state.records.forEach((rec, idx) => {
        const changedFields = rec.__changed_fields || [];
        if (changedFields.length > 0) {
          const rowId = rec.__id || String(idx);
          state.financeAutopopulatedRows.add(rowId);
          state.referenceUpdatedFieldsByRowId[rowId] = new Set(changedFields);
        }
      });

      // Automatically isolate and display only the changed rows with glowing cells
      state.showOnlyChanged = true;
      recalculateKPIs();
      updateStagedEditsUI();
      filterAndRenderRecords();

      showToast(`Staged updates for ${uCount.toLocaleString()} record(s) from Reference Data. Changed rows and cells are highlighted in green. Review and click 'Save to Dataverse' on the toolbar when ready to commit.`, 'ok');
    } catch (err) {
      showToast('Error: ' + err, 'err');
    } finally {
      if (dom.btnUpdateFromFinance) {
        dom.btnUpdateFromFinance.disabled = false;
        dom.btnUpdateFromFinance.textContent = 'Update from References';
      }
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
    if (dom.fAcquisitionSubLevel) {
      ACQUISITION_SUBLEVEL_OPTIONS.forEach(code => {
        const option = document.createElement('option');
        option.value = code;
        option.textContent = `${code} — ${ACQUISITION_SUBLEVEL_STATUS[code]}`;
        dom.fAcquisitionSubLevel.appendChild(option);
      });
    }
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

    // An account loaded by picking its Area-Index keeps its stored values: while the
    // Area-Index still names that account, a lookup only compares with the reference.
    const onIdentityInput = () => {
      if (drawerPicked && dom.fRecordId.value && dom.fAreaIndex
          && dom.fAreaIndex.value.trim().toUpperCase() === drawerPicked.area.toUpperCase()) {
        triggerLiveLookup(false, false);
        return;
      }
      triggerLiveLookup();
    };
    const lookupTriggers = [dom.fType, dom.fIndexNo, dom.fAreaIndex, dom.fTitleNo];
    lookupTriggers.forEach(input => {
      if (input) {
        input.addEventListener('input', onIdentityInput);
        input.addEventListener('change', onIdentityInput);
      }
    });

    [dom.fSecuredDate, dom.fSecuredEndorseDate, dom.fEndorsedDate].forEach(input => {
      if (input) {
        let beforeDate = input.value;
        input.addEventListener('focus', () => { beforeDate = input.value; });
        input.addEventListener('input', () => {
          input.value = maskDateString(input.value);
        });
        input.addEventListener('blur', () => { commitDateInput(input, beforeDate); });
      }
    });

    attachLiveAmountFormatting(dom.fTcp, updatePaymentMath);
    attachLiveAmountFormatting(dom.fPaymentDeductible, updatePaymentMath);
    attachLiveAmountFormatting(dom.fBalance, updatePaymentMath);
    attachLiveAmountFormatting(dom.fRccPrice, null);
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
          dom.fPercentPaid.value = pct.toFixed(1) + '%';
        } else if (ded > 0) {
          dom.fPercentPaid.value = '100.0%';
        } else {
          dom.fPercentPaid.value = '0.0%';
        }
      }
    }
  }

  function setupAreaIndexSuggestions() {
    if (!dom.fAreaIndex || !dom.areaIndexSuggestions) return;

    function renderAreaSuggestions(query) {
      // Editing an existing account: the suggestion list (copy another account
      // as a template) does not apply.
      if (dom.fRecordId && dom.fRecordId.value && !drawerPicked) {
        dom.areaIndexSuggestions.classList.add('hidden');
        return;
      }
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
          if (match) pickExistingArea(k, match.rec);
          dom.areaIndexSuggestions.classList.add('hidden');
        });
      });
    }

    dom.fAreaIndex.addEventListener('input', () => {
      // Typing another Area-Index after picking an account: this is a new account,
      // so stop updating the picked one.
      if (drawerPicked && dom.fRecordId.value
          && dom.fAreaIndex.value.trim().toUpperCase() !== drawerPicked.area.toUpperCase()) {
        setPickedDrawerMode(false, true);
      }
      renderAreaSuggestions(dom.fAreaIndex.value);
      clearTimeout(autoPickTimer);
      autoPickTimer = setTimeout(autoPickExactArea, 350);
    });
    // Tab / Enter / leaving the field with a complete Area-Index loads it too, so
    // typing gives the same result as clicking the suggestion.
    dom.fAreaIndex.addEventListener('change', autoPickExactArea);
    dom.fAreaIndex.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); autoPickExactArea(); }
    });

    if (dom.fControl) {
      // Picked account: its own Control Number updates it; any other number creates
      // a new account that uses it as a template.
      dom.fControl.addEventListener('input', () => {
        if (!drawerPicked) return;
        const same = dom.fControl.value.trim().toUpperCase() === drawerPicked.control.toUpperCase();
        if (same && !dom.fRecordId.value) setPickedDrawerMode(true);
        else if (!same && dom.fRecordId.value) setPickedDrawerMode(false);
      });
    }

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

    // Mode decides whether the SALE or LEASE Finance row applies, so re-run the lookup when it changes.
    if (dom.fMode) {
      dom.fMode.addEventListener('change', () => triggerLiveLookup(false));
    }

    document.addEventListener('click', (e) => {
      if (!e.target.closest('#fAreaIndex') && !e.target.closest('#areaIndexSuggestions')) {
        dom.areaIndexSuggestions.classList.add('hidden');
      }
    });
  }

  // New LOA Entry / Log (decision 2026-10-14): picking an existing Area-Index loads
  // that account, so its Control Number shows like the other fields and Save updates
  // it and logs the changes. Typing a different Control Number (or Area-Index) turns
  // the drawer back into a new entry that uses the picked account as a template.
  let drawerPicked = null; // { id, control, area, baseline, label }
  // True while the drawer was opened with New LOA Entry / Log. In that drawer the
  // Endorsed Date is always today: on open, after picking or loading an account, when
  // switching between update and new, and a Finance lookup never replaces it.
  let drawerIsEntryLog = false;
  let autoPickTimer = null;
  let lastAutoPickedArea = '';

  function applyEntryLogEndorsedDate() {
    if (!drawerIsEntryLog || !dom.fEndorsedDate) return;
    dom.fEndorsedDate.value = todayMMDDYYYY();
    dom.fEndorsedDate.title = "New LOA Entry / Log always uses today's date.";
  }

  const recordAreaIndex = (r) => String(r['cr63f_areaindex'] || r['Area-Index'] || r['AREA-INDEX'] || '').trim();

  // A typed Area-Index that exactly matches an existing account loads it, the same as
  // clicking its suggestion.
  function autoPickExactArea() {
    clearTimeout(autoPickTimer);
    if (!drawerIsEntryLog || !dom.fAreaIndex) return;
    const typed = dom.fAreaIndex.value.trim().toUpperCase();
    if (!typed || typed === lastAutoPickedArea) return;
    const rec = (state.records || []).find(r => recordAreaIndex(r).toUpperCase() === typed);
    if (!rec) return;
    pickExistingArea(recordAreaIndex(rec), rec);
    if (dom.areaIndexSuggestions) dom.areaIndexSuggestions.classList.add('hidden');
  }

  function setPickedDrawerMode(update, clearControl = false) {
    if (!drawerPicked) return;
    if (update) {
      dom.fRecordId.value = drawerPicked.id;
      drawerBaseline = drawerPicked.baseline;
      if (dom.drawerTitle) dom.drawerTitle.textContent = `Update account ${drawerPicked.label}`;
      if (dom.btnSaveDrawerEntry) dom.btnSaveDrawerEntry.textContent = 'Save Changes';
      applyEntryLogEndorsedDate();
    } else {
      dom.fRecordId.value = '';
      drawerBaseline = null;
      if (dom.drawerTitle) dom.drawerTitle.textContent = 'New LOA Entry / Log';
      if (dom.btnSaveDrawerEntry) dom.btnSaveDrawerEntry.textContent = 'Save Entry to Dataverse';
      if (clearControl && dom.fControl) dom.fControl.value = '';
      applyEntryLogEndorsedDate();
    }
  }

  function pickExistingArea(areaKey, templateRec) {
    const wanted = String(areaKey || '').trim().toUpperCase();
    lastAutoPickedArea = wanted;
    const accounts = (state.records || []).filter(r => recordAreaIndex(r).toUpperCase() === wanted);
    const rec = accounts.length === 1 ? accounts[0] : null;
    const hasGridEdits = rec && state.stagedEdits && state.stagedEdits[rec.__id]
      && Object.keys(state.stagedEdits[rec.__id]).length;

    if (rec && rec.__id && !hasGridEdits) {
      dom.fRecordId.value = String(rec.__id);
      populateFormFromRecord(rec, areaKey);
      const control = dom.fControl ? dom.fControl.value.trim() : '';
      drawerPicked = {
        id: String(rec.__id), control, area: dom.fAreaIndex ? dom.fAreaIndex.value.trim() : areaKey,
        baseline: drawerBaseline, label: control || areaKey,
      };
      setPickedDrawerMode(true);
      showToast(`Loaded ${control || areaKey}. Save updates this account and logs the changes. Type a different Control Number to create a new account from it instead.`, 'info');
      return;
    }

    // Several accounts share this Area-Index, or the account has unsaved grid edits:
    // use the picked one as a template for a new account.
    drawerPicked = null;
    dom.fRecordId.value = '';
    populateFormFromRecord(templateRec, areaKey);
    drawerBaseline = null;
    if (dom.fControl) { dom.fControl.value = ''; dom.fControl.focus(); }
    applyEntryLogEndorsedDate();
    showToast(accounts.length > 1
      ? `${accounts.length} accounts share ${areaKey}. Filled from one as a template; enter the Control Number for this account.`
      : hasGridEdits
        ? `${areaKey} has unsaved grid edits, so it was used as a template. Enter a new Control Number, or save the grid edits first.`
        : `Filled from ${areaKey}. Enter a new Control Number for this account.`, 'info');
  }

  function populateFormFromRecord(r, areaVal) {
    if (!r) return;
    if (dom.fAreaIndex) dom.fAreaIndex.value = areaVal || r['cr63f_areaindex'] || r['Area-Index'] || '';
    if (dom.fType) ensureSelectValue(dom.fType, r['cr63f_type'] || r['Type'] || '');
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
    drawerStored = { tcp: tcpVal, ded: dedVal, bal: balVal, pct: pctVal };
    drawerAutofilled = false;

    if (dom.fRegisteredOwner) dom.fRegisteredOwner.value = r['cr63f_registeredowner'] || r['Registered Owner'] || '';
    if (dom.fAuthorizedRepresentative) dom.fAuthorizedRepresentative.value = r['cr63f_authorizedrepresentative'] || r['Authorized Representative'] || '';
    if (dom.fEntity) ensureSelectValue(dom.fEntity, r['cr63f_entity'] || r['Entity'] || '');

    if (dom.fProvince) dom.fProvince.value = r['cr63f_province'] || r['Province'] || '';
    if (dom.fMunicipality) dom.fMunicipality.value = r['cr63f_municipality'] || r['Municipality'] || '';
    if (dom.fDbAdjustedBarangay) dom.fDbAdjustedBarangay.value = r['cr63f_dbadjustedbarangay'] || r['DB Adjusted Barangay'] || '';
    if (dom.fLotAreaHa) dom.fLotAreaHa.value = r['cr63f_lotareaha'] || r['Lot Area (Ha)'] || '';
    if (dom.fPossRightsStatus) dom.fPossRightsStatus.value = r['cr63f_possrightsstatus'] || r['Poss Rights Status'] || '';
    if (dom.fPossRightsArea) dom.fPossRightsArea.value = r['cr63f_possrightsarea'] || r['Poss Rights Area (Ha)'] || '';
    if (dom.fPossRightsCat) dom.fPossRightsCat.value = r['cr63f_possrightscat'] || r['Possrights Cat'] || '';
    if (dom.fCatArea) dom.fCatArea.value = r['cr63f_catarea'] || r['Cat Area (Ha)'] || '';

    // ensureSelectValue keeps a stored value that is not in the list, so it is never blanked.
    if (dom.fContractType) ensureSelectValue(dom.fContractType, r['new_contracttype'] || r['Contract Type'] || '');
    if (dom.fMode) ensureSelectValue(dom.fMode, r['new_mode'] || r['Mode'] || '');
    if (dom.fConversionStatus) dom.fConversionStatus.value = r['cr63f_conversionstatus'] || r['Conversion Status'] || '';
    if (dom.fAcquisitionSubLevel) {
      dom.fAcquisitionSubLevel.value = r['cr63f_acquisitionsublevel'] || r['Acquisition Sub-Level'] || '';
    }
    if (dom.fRccTerms) dom.fRccTerms.value = r['cr63f_rccterms'] || r['RCC Terms'] || '';
    
    const rccVal = r['cr63f_rccprice'] ?? r['RCC Price'] ?? '';
    if (dom.fRccPrice) dom.fRccPrice.value = rccVal !== '' ? formatAmount(rccVal) : '';

    if (dom.fSecuredDate) dom.fSecuredDate.value = formatDateMMDDYYYY(r['cr63f_secureddate'] || r['Secured Date']);
    if (dom.fSecuredEndorseDate) dom.fSecuredEndorseDate.value = formatDateMMDDYYYY(r['cr63f_securedendorsedate'] || r['Secured Endorse Date']);
    if (dom.fEndorsedDate) {
      dom.fEndorsedDate.value = formatDateMMDDYYYY(r['cr63f_endorseddate'] || r['Endorsed Date']);
      dom.fEndorsedDate.title = '';
    }

    drawerBaseline = collectDrawerData();
    // Endorsed Date defaults to today (decision 2026-10-14). Filled after the baseline,
    // so it is saved with the account's next save and dropped if the drawer is closed
    // without saving. Only a truly blank date is filled; a stored 01/01/1900 placeholder
    // is kept, the same rule as Update from References (decided 2026-10-08).
    const storedEndorsed = String(r['cr63f_endorseddate'] ?? r['Endorsed Date'] ?? r['endorsed_date'] ?? '').trim();
    if (dom.fEndorsedDate && ['', '-', 'None', 'null'].includes(storedEndorsed)) {
      dom.fEndorsedDate.value = todayMMDDYYYY();
      dom.fEndorsedDate.title = "Defaulted to today's date; saved when you save this account.";
    }
    // Opening an account shows its stored values. The reference is compared, not applied.
    triggerLiveLookup(false, false);
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
    drawerPicked = null;
    drawerIsEntryLog = true;
    lastAutoPickedArea = '';

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
    if (dom.fPossRightsStatus) dom.fPossRightsStatus.value = '';
    if (dom.fPossRightsArea) dom.fPossRightsArea.value = '';
    if (dom.fPossRightsCat) dom.fPossRightsCat.value = '';
    if (dom.fCatArea) dom.fCatArea.value = '';
    if (dom.fContractType) dom.fContractType.value = '';
    if (dom.fMode) dom.fMode.value = '';
    if (dom.fConversionStatus) dom.fConversionStatus.value = '';
    if (dom.fAcquisitionSubLevel) dom.fAcquisitionSubLevel.value = '';
    if (dom.fRccTerms) dom.fRccTerms.value = '';
    if (dom.fRccPrice) dom.fRccPrice.value = '';
    if (dom.fSecuredDate) dom.fSecuredDate.value = '';
    if (dom.fSecuredEndorseDate) dom.fSecuredEndorseDate.value = '';
    
    // Endorsed Date: today (MM/DD/YYYY).
    applyEntryLogEndorsedDate();

    drawerBaseline = null;
    drawerStored = null;
    drawerAutofilled = false;
    if (dom.btnSaveDrawerEntry) dom.btnSaveDrawerEntry.textContent = 'Save Entry to Dataverse';
    openDrawer();
  }

  // Open an existing account in the drawer. The drawer shows its stored values;
  // saving sends only the fields changed here (plus identity keys) through the
  // same version-checked, audited save as before.
  function openEditEntry(recordId) {
    const rec = (state.records || []).find(r => String(r.__id) === String(recordId));
    if (!rec) {
      showToast('This account is not loaded. Reload Land Master and try again.', 'err');
      return;
    }
    if (state.stagedEdits && state.stagedEdits[recordId] && Object.keys(state.stagedEdits[recordId]).length) {
      showToast('This account has unsaved grid edits. Save or discard them first.', 'err');
      return;
    }
    const label = rec.cr63f_control1 || rec['Control Number'] || rec.cr63f_areaindex || rec['Area-Index'] || '';
    drawerPicked = null;
    drawerIsEntryLog = false;
    if (dom.drawerTitle) dom.drawerTitle.textContent = `Edit account ${label}`;
    if (dom.areaIndexSuggestions) dom.areaIndexSuggestions.classList.add('hidden');
    clearAutofillFields();
    dom.fRecordId.value = String(recordId);
    populateFormFromRecord(rec);
    if (dom.btnSaveDrawerEntry) dom.btnSaveDrawerEntry.textContent = 'Save Changes';
    openDrawer();
  }

  window.openEditEntry = openEditEntry;
  window.openNewEntry = openNewEntry;
  window.closeDrawer = closeDrawer;

  function clearAutofillFields() {
    if (dom.fTcp) dom.fTcp.value = '';
    if (dom.fPaymentDeductible) dom.fPaymentDeductible.value = '';
    if (dom.fBalance) dom.fBalance.value = '';
    if (dom.fPercentPaid) dom.fPercentPaid.value = '';
  }

  function setAmountFields(tcp, ded, bal, pct) {
    const has = v => v !== '' && v !== null && v !== undefined;
    if (dom.fTcp) dom.fTcp.value = has(tcp) ? formatAmount(tcp) : '';
    if (dom.fPaymentDeductible) dom.fPaymentDeductible.value = has(ded) ? formatAmount(ded) : '';
    if (dom.fBalance) dom.fBalance.value = has(bal) ? formatAmount(bal) : '';
    if (dom.fPercentPaid) dom.fPercentPaid.value = has(pct) ? formatPercent(pct) : '';
  }

  // A lookup that finds nothing must not destroy the open account's amounts: restore the
  // stored ones, or clear them when the current amounts came from an earlier lookup.
  function restoreAmountsAfterNoMatch() {
    if (!drawerAutofilled) return;
    drawerAutofilled = false;
    if (drawerStored) {
      setAmountFields(drawerStored.tcp, drawerStored.ded, drawerStored.bal, drawerStored.pct);
    } else {
      clearAutofillFields();
    }
  }

  // applyToForm = true: typing in the drawer. A match fills the amounts from the reference.
  // applyToForm = false: opening an account. Stored values stay; the badge reports whether the
  // reference differs from them.
  // Each lookup gets a number; a reply that comes back after a newer lookup started
  // (e.g. after switching to another Area-Index) is ignored, so it can never write the
  // previous account's reference values over the current one.
  let lookupSeq = 0;
  function triggerLiveLookup(clearOnNoMatch = true, applyToForm = true) {
    clearTimeout(debounceTimer);
    const seq = ++lookupSeq;
    debounceTimer = setTimeout(async () => {
      if (seq !== lookupSeq) return;
      const typeVal = dom.fType ? dom.fType.value.trim() : '';
      const indexNo = dom.fIndexNo ? dom.fIndexNo.value.trim() : '';
      const areaIndex = dom.fAreaIndex ? dom.fAreaIndex.value.trim() : '';
      const titleNo = dom.fTitleNo ? dom.fTitleNo.value.trim() : '';

      if (!typeVal && !indexNo && !areaIndex && !titleNo) {
        if (clearOnNoMatch && applyToForm) restoreAmountsAfterNoMatch();
        if (dom.badgeMatchStatus) {
          dom.badgeMatchStatus.className = 'match-badge waiting';
          dom.badgeMatchStatus.textContent = 'Waiting for input...';
        }
        return;
      }

      try {
        const modeVal = dom.fMode ? dom.fMode.value.trim() : '';
        const res = await window.pywebview.api.lookup_finance(typeVal, indexNo, areaIndex, titleNo, modeVal);
        if (seq !== lookupSeq) return;
        const storedTcp = dom.fTcp ? parseAmount(dom.fTcp.value) : 0;
        const storedDed = dom.fPaymentDeductible ? parseAmount(dom.fPaymentDeductible.value) : 0;
        let refDiffers = false;
        if (res.matched) {
          if (applyToForm) {
            if (dom.fTcp) dom.fTcp.value = res.total_contract_price_formatted || '0.00';
            if (dom.fPaymentDeductible) dom.fPaymentDeductible.value = res.total_payment_deductible_formatted || '0.00';
            if (dom.fBalance) dom.fBalance.value = res.balance_formatted || '0.00';
            if (dom.fPercentPaid) dom.fPercentPaid.value = storedPercentText(res.percent_paid_formatted || res.percent_paid || '0');
            if (res.secured_date && dom.fSecuredDate) dom.fSecuredDate.value = formatDateMMDDYYYY(res.secured_date);
            if (res.secured_endorse_date && dom.fSecuredEndorseDate) dom.fSecuredEndorseDate.value = formatDateMMDDYYYY(res.secured_endorse_date);
            if (res.endorsed_date && dom.fEndorsedDate && !drawerIsEntryLog) dom.fEndorsedDate.value = formatDateMMDDYYYY(res.endorsed_date);
            drawerAutofilled = true;
            updatePaymentMath();
          } else {
            refDiffers = Math.abs(storedTcp - Number(res.total_contract_price || 0)) > 0.005
              || Math.abs(storedDed - Number(res.total_payment_deductible || 0)) > 0.005;
          }
        } else if (applyToForm && clearOnNoMatch) {
          restoreAmountsAfterNoMatch();
        }

        // Live Conversion Status lookup
        let convMatched = false;
        let convType = '';
        let convStatus = '';
        if (areaIndex) {
          try {
            const convRes = await window.pywebview.api.lookup_conversion(areaIndex);
            if (seq !== lookupSeq) return;
            if (convRes && convRes.matched && convRes.status) {
              convMatched = true;
              convType = convRes.match_type || 'Conversion';
              convStatus = convRes.status;
              if (applyToForm) {
                if (dom.fConversionStatus) dom.fConversionStatus.value = convStatus;
              } else if (dom.fConversionStatus
                  && String(dom.fConversionStatus.value || '').trim().toUpperCase() !== convStatus.trim().toUpperCase()) {
                refDiffers = true;
              }
            }
          } catch (eConv) {
            console.error('Conversion lookup error:', eConv);
          }
        }

        if (dom.badgeMatchStatus) {
          if (!applyToForm && (res.matched || convMatched) && refDiffers) {
            dom.badgeMatchStatus.className = 'match-badge nomatch';
            dom.badgeMatchStatus.textContent = '⚠ Reference differs from the stored values. Stored values kept; use Update from References to apply.';
          } else if (res.matched && convMatched) {
            dom.badgeMatchStatus.className = 'match-badge matched';
            dom.badgeMatchStatus.textContent = `✓ Ref: ${res.match_type} | Conv: ${convType}`;
          } else if (res.matched) {
            dom.badgeMatchStatus.className = 'match-badge matched';
            dom.badgeMatchStatus.textContent = `✓ ${res.match_type} [${res.match_key}]`;
          } else if (convMatched) {
            dom.badgeMatchStatus.className = 'match-badge matched';
            dom.badgeMatchStatus.textContent = `✓ Conv: ${convType} (${convStatus})`;
          } else {
            dom.badgeMatchStatus.className = 'match-badge nomatch';
            dom.badgeMatchStatus.textContent = '⚠ No match in Appended Reference';
          }
        }
      } catch (err) {
        console.error('Lookup error:', err);
      }
    }, 150);
  }

  // Drawer state for the open account. drawerBaseline: the fields as the drawer showed them
  // on open (a save sends only what changed). drawerStored: the stored amounts, restored when
  // a lookup finds nothing. drawerAutofilled: the amount fields currently hold a lookup result.
  let drawerBaseline = null;
  let drawerStored = null;
  let drawerAutofilled = false;
  const DRAWER_IDENTITY_KEYS = [
    'Type', 'cr63f_type', 'Index No', 'cr63f_indexno', 'Area-Index', 'cr63f_areaindex',
    'Title No', 'cr63f_titleno', 'Control', 'cr63f_control1',
  ];

  // Every drawer field as a save sends it, keyed by the display and Dataverse names.
  function collectDrawerData() {
    return {
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
      'Poss Rights Status': dom.fPossRightsStatus ? dom.fPossRightsStatus.value.trim() : '',
      'cr63f_possrightsstatus': dom.fPossRightsStatus ? dom.fPossRightsStatus.value.trim() : '',
      'Poss Rights Area (Ha)': dom.fPossRightsArea ? dom.fPossRightsArea.value.trim() : '',
      'cr63f_possrightsarea': dom.fPossRightsArea ? dom.fPossRightsArea.value.trim() : '',
      'Possrights Cat': dom.fPossRightsCat ? dom.fPossRightsCat.value.trim() : '',
      'cr63f_possrightscat': dom.fPossRightsCat ? dom.fPossRightsCat.value.trim() : '',
      'Cat Area (Ha)': dom.fCatArea ? dom.fCatArea.value.trim() : '',
      'cr63f_catarea': dom.fCatArea ? dom.fCatArea.value.trim() : '',

      'Contract Type': dom.fContractType.value.trim(),
      'new_contracttype': dom.fContractType.value.trim(),
      'Mode': dom.fMode.value.trim(),
      'new_mode': dom.fMode.value.trim(),
      'Conversion Status': dom.fConversionStatus.value.trim(),
      'cr63f_conversionstatus': dom.fConversionStatus.value.trim(),
      'Acquisition Sub-Level': dom.fAcquisitionSubLevel.value.trim(),
      'cr63f_acquisitionsublevel': dom.fAcquisitionSubLevel.value.trim(),
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
  }

  async function handleSaveEntry() {
    const recordId = dom.fRecordId.value || null;
    const recordData = collectDrawerData();

    // Dates must be real dates: an unreadable one is refused here rather than
    // saved as blank.
    for (const [label, keys] of [
      ['Endorsed Date', ['Endorsed Date', 'cr63f_endorseddate']],
      ['Secured Date', ['Secured Date', 'cr63f_secureddate']],
      ['Secured Endorse Date', ['Secured Endorse Date', 'cr63f_securedendorsedate']],
    ]) {
      const normalized = normalizeDateInput(recordData[keys[0]]);
      if (normalized === null) {
        showToast(`${label} "${recordData[keys[0]]}" is not a valid date. Use MM/DD/YYYY.`, 'err');
        return;
      }
      keys.forEach(k => { recordData[k] = normalized; });
    }

    if (!recordData['Type']) {
      showToast('TYPE is required.', 'err');
      return;
    }

    const isNew = !recordId;
    if (isNew && !recordData['cr63f_acquisitionsublevel']) {
      showToast('Acquisition Sub-Level is required.', 'err');
      return;
    }
    // Control Number identifies the account in Entry Logs and Revert, so a new entry needs one.
    if (isNew && !String(recordData['cr63f_control1'] || '').trim()) {
      showToast('Control Number is required for a new entry.', 'err');
      if (dom.fControl) dom.fControl.focus();
      return;
    }
    // An existing account sends only the fields changed in the drawer (plus its identity
    // keys). A field the drawer never touched stays exactly as stored, so opening an account
    // and saving another field cannot rewrite its amounts, status or Mode.
    // Identity fields (Type, Index No, AREA-INDEX, Title No, Control Number) count as changes
    // too: they used to be skipped here, so editing only the Title No said "No changes to
    // save" and nothing was written (fixed 2026-10-14).
    let payload = recordData;
    if (!isNew && drawerBaseline) {
      const changedKeys = Object.keys(recordData).filter(k =>
        String(recordData[k] ?? '').trim() !== String(drawerBaseline[k] ?? '').trim());
      if (changedKeys.length === 0) {
        showToast('No changes to save.', 'info');
        return;
      }
      payload = {};
      Object.keys(recordData).forEach(k => {
        if (DRAWER_IDENTITY_KEYS.includes(k) || changedKeys.includes(k)) payload[k] = recordData[k];
      });
    }

    const confirmed = await showConfirmModal({
      title: isNew ? 'Create New Master Record' : 'Save Record Changes',
      subtitle: 'Dataverse & Audit Log Confirmation',
      message: isNew
        ? `Are you sure you want to create this new Land Master record in Dataverse? A NEW_ENTRY audit log will be recorded in the Entry Logs repository.`
        : `Are you sure you want to save modifications to this Land Master record in Dataverse? All field changes will be recorded in the Entry Logs repository.`,
      confirmText: isNew ? 'Create & Log Record' : 'Save & Log Changes',
      cancelText: 'Cancel',
      isDanger: false
    });
    if (!confirmed) return;

    try {
      dom.btnSaveDrawerEntry.disabled = true;
      dom.btnSaveDrawerEntry.textContent = 'Saving to Dataverse...';

      let res = await window.pywebview.api.save_loa_record_entry(payload, recordId);
      if (!res.success && res.needs_confirmation) {
        // AREA-INDEX already used by another account: allowed for parcels in the same
        // area, but only after the user confirms it is not the same account typed twice.
        const shareArea = await showConfirmModal({
          title: 'AREA-INDEX Already Used',
          subtitle: 'Possible duplicate account',
          message: `${res.message} Choose Cancel if this is the same account entered again.`,
          confirmText: 'Save as Another Account',
          cancelText: 'Cancel',
          isDanger: false
        });
        if (!shareArea) {
          showToast('New entry was not saved.', 'info');
          return;
        }
        res = await window.pywebview.api.save_loa_record_entry(recordData, recordId, true);
      }
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
      dom.btnSaveDrawerEntry.textContent = dom.fRecordId.value ? 'Save Changes' : 'Save Entry to Dataverse';
    }
  }

  // ==================== FINANCE DATA VIEW ====================
  function setupFinanceView() {
    dom.btnUploadFinance.addEventListener('click', async () => {
      const stagedCount = Object.keys(financeState.stagedEdits || {}).length;
      if (stagedCount > 0) {
        const ok = await showConfirmModal({
          title: 'Unsaved Finance Changes',
          subtitle: 'Upload Conflict',
          message: `You have ${stagedCount} unsaved modified row(s) in Finance Data. Uploading a new file will overwrite these changes. Do you want to continue and discard unsaved edits?`,
          confirmText: 'Discard & Upload',
          cancelText: 'Cancel',
          isDanger: true
        });
        if (!ok) return;
      }
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
      const stagedCount = Object.keys(financeState.stagedEdits || {}).length;
      if (stagedCount > 0) {
        const ok = await showConfirmModal({
          title: 'Clear Finance Data',
          subtitle: 'Unsaved Edits',
          message: `You have ${stagedCount} unsaved modified row(s). Clearing will discard them. Are you sure?`,
          confirmText: 'Clear Data',
          cancelText: 'Cancel',
          isDanger: true
        });
        if (!ok) return;
      }
      await window.pywebview.api.clear_finance_reference();
      const s = await window.pywebview.api.get_initial_state();
      applyState(s);
      financeState.headers = [];
      financeState.rows = [];
      financeState.baselineRows = [];
      financeState.stagedEdits = {};
      financeState.showOnlyChanged = false;
      financeState.filteredRows = [];
      financeState.renderedRowsCount = 0;
      updateStagedEditsUIFinance();
      if (dom.lblFinanceGridCount) dom.lblFinanceGridCount.textContent = '0 rows';
      renderFinanceTable();
      showToast('Finance data cleared.', 'ok');
    });

    if (dom.btnToggleShowOnlyChangedFinance) {
      dom.btnToggleShowOnlyChangedFinance.addEventListener('click', toggleShowOnlyChangedFinance);
    }
    if (dom.btnSaveStagedEditsFinance) {
      dom.btnSaveStagedEditsFinance.addEventListener('click', saveStagedEditsFinance);
    }
    if (dom.btnDiscardStagedEditsFinance) {
      dom.btnDiscardStagedEditsFinance.addEventListener('click', discardStagedEditsFinance);
    }

    if (dom.searchInputFinance) {
      dom.searchInputFinance.addEventListener('input', () => {
        clearTimeout(searchFinDebounce);
        searchFinDebounce = setTimeout(filterAndRenderFinance, 120);
      });
    }
    if (dom.typeFilterFinance) {
      dom.typeFilterFinance.addEventListener('change', () => filterAndRenderFinance());
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
    if (dom.btnDownloadFinanceTemplate) {
      dom.btnDownloadFinanceTemplate.addEventListener('click', downloadFinanceTemplate);
    }
    if (dom.btnSyncFinanceToDv) {
      dom.btnSyncFinanceToDv.addEventListener('click', syncFinanceToDataverse);
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
        localStorage.removeItem('fin_sort_col');
        localStorage.removeItem('fin_sort_dir');
        financeState.visibleCols = new Set(financeState.headers);
        financeState.colOrder = [...financeState.headers];
        financeState.colWidths = {};
        financeState.frozenCols = new Set();
        financeState.sortCol = null;
        financeState.sortDir = null;
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
    financeState.baselineRows = JSON.parse(JSON.stringify(rows));
    financeState.stagedEdits = {};
    financeState.showOnlyChanged = false;
    updateStagedEditsUIFinance();

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

    try {
      const sc = localStorage.getItem('fin_sort_col');
      const sd = localStorage.getItem('fin_sort_dir');
      financeState.sortCol = sc || null;
      financeState.sortDir = sd || null;
    } catch (_) {
      financeState.sortCol = null;
      financeState.sortDir = null;
    }

    // Populate type filter for finance if TYPE column exists
    if (dom.typeFilterFinance) {
      const currentType = String(dom.typeFilterFinance.value || 'ALL').trim().toUpperCase();
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
        dom.typeFilterFinance.value = typeSet.has(currentType) ? currentType : 'ALL';
      } else {
        dom.typeFilterFinance.innerHTML = '<option value="ALL">All Types</option>';
      }
    }

    renderFinanceColumnPickerMenu();
    filterAndRenderFinance();
  }

  function filterAndRenderFinance(skipHeadRender = false) {
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

      if (!q) {
        if (financeState.showOnlyChanged) {
          const id = r.__id;
          const rowIdx = String(financeState.rows.indexOf(r));
          return Boolean((id && financeState.stagedEdits[id]) || financeState.stagedEdits[rowIdx]);
        }
        return true;
      }
      const matchesQ = Object.values(r).some(v => String(v).toLowerCase().includes(q));
      if (!matchesQ) return false;
      if (financeState.showOnlyChanged) {
        const id = r.__id;
        const rowIdx = String(financeState.rows.indexOf(r));
        return Boolean((id && financeState.stagedEdits[id]) || financeState.stagedEdits[rowIdx]);
      }
      return true;
    });

    if (financeState.sortCol && financeState.sortDir) {
      const col = financeState.sortCol;
      const mult = financeState.sortDir === 'asc' ? 1 : -1;
      const colDef = {
        key: col,
        isPrice: /PRICE|TCP|DEDUCTIBLE|BALANCE|COST|AMOUNT/i.test(col),
        isPercent: /PAID|PERCENT|%/i.test(col),
        isDate: /DATE/i.test(col),
        isNum: /AREA|HA/i.test(col)
      };
      financeState.filteredRows.sort((a, b) => {
        const valA = a[col] ?? '';
        const valB = b[col] ?? '';
        return compareTableValues(valA, valB, colDef, mult);
      });
    }

    if (dom.lblFinanceGridCount) {
      dom.lblFinanceGridCount.textContent = `${financeState.filteredRows.length.toLocaleString()} rows`;
    }

    if (skipHeadRender) {
      // Only re-render the body rows — keep the header (and the focused filter input) intact
      const displayRows = financeState.filteredRows;
      renderFinanceBody(displayRows);
    } else {
      renderFinanceTable();
    }
  }

  const DEFAULT_FINANCE_HEADERS = [
    'Control No.',
    'Type',
    'Area-Index',
    'Title No.',
    'Total Contract Price',
    'Total Payment Deductible',
    'Outstanding Balance',
    'DAR Status',
    'Secured Date'
  ];

  function getFinanceVisibleColumns() {
    let order = financeState.colOrder.length ? financeState.colOrder : financeState.headers;
    if (!order.length) order = DEFAULT_FINANCE_HEADERS;
    return order.filter(h => (!financeState.visibleCols.size || financeState.visibleCols.has(h)) && !isSystemOrGeneratedCol(h));
  }

  function renderFinanceTable() {
    if (!dom.financePreviewThead || !dom.financePreviewTbody) return;
    const cols = getFinanceVisibleColumns();

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
              <button type="button" class="th-pin-btn ${isFrozen ? 'pinned' : ''}" data-fin-pin="${escapeHtml(h)}" aria-label="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(h)}" title="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(h)}">
                <svg viewBox="0 0 24 24" width="11" height="11" fill="currentColor" stroke="none">
                  <path d="M16 12V4h1V2H7v2h1v8l-2 2v2h5.2v6h1.6v-6H18v-2l-2-2z"/>
                </svg>
              </button>
              <button type="button" class="th-sort-wrap" data-fin-sort="${escapeHtml(h)}" aria-label="Sort by ${escapeHtml(h)}" title="Sort · drag to reorder">
                <span class="th-col-name">${escapeHtml(h)}</span>
                <span class="th-sort-icon ${isSorted ? 'active' : ''}" aria-hidden="true">${sortIcon}</span>
              </button>
              <button type="button" class="th-chevron-btn ${isFilterActive ? 'active' : ''}" data-fin-filter-toggle="${escapeHtml(h)}" aria-label="Filter ${escapeHtml(h)}" title="Filter ${escapeHtml(h)}">
                <svg viewBox="0 0 24 24" width="10" height="10" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"></polyline></svg>
              </button>
            </div>
            <div class="th-filter-popup ${isFilterOpen ? '' : 'hidden'}" id="finColFilterBox_${escapeHtml(h)}">
              <input class="th-filter-input" data-fin-filter-col="${escapeHtml(h)}" aria-label="Filter ${escapeHtml(h)} column" autocomplete="off" value="${escapeHtml(financeState.columnFilters[h] || '')}" placeholder="Filter ${escapeHtml(h)}…" />
            </div>
          </div>
          <div class="th-resizer" data-fin-resize="${escapeHtml(h)}" title="Drag to resize · double-click to reset"></div>
          <div class="th-bottom-resizer" title="Drag to resize the header"></div>
        </th>
      `;
    });
    thead += '</tr>';
    dom.financePreviewThead.innerHTML = thead;

    attachFinanceHeaderEvents();

    const displayRows = financeState.filteredRows;

    if (!displayRows.length) {
      dom.financePreviewTbody.innerHTML = '';
      const colSpan = Math.max(1, cols.length + 1);
      const isInitialEmpty = !financeState.rows.length;
      const title = isInitialEmpty ? 'No Finance Reference data loaded' : 'No matching finance records';
      const desc = isInitialEmpty
        ? 'Click <strong>Upload Excel / CSV</strong> or <strong>Sync to Dataverse</strong> to load reference data.'
        : 'Try adjusting your search terms, Type filter, or column filters.';

      const emptyTr = document.createElement('tr');
      emptyTr.className = 'empty-table-row';
      emptyTr.innerHTML = `
        <td colspan="${colSpan}" style="text-align: center; padding: 48px 16px; background: transparent; border: none;">
          <div style="display: inline-flex; flex-direction: column; align-items: center; justify-content: center; gap: 6px;">
            <svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" style="color: var(--ink3); opacity: 0.7;">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
              <line x1="16" y1="13" x2="8" y2="13"></line>
              <line x1="16" y1="17" x2="8" y2="17"></line>
              <polyline points="10 9 9 9 8 9"></polyline>
            </svg>
            <p style="font-size: 13.5px; font-weight: 600; color: var(--ink); margin: 4px 0 2px;">${title}</p>
            <p style="font-size: 12px; color: var(--ink3); margin: 0;">${desc}</p>
          </div>
        </td>
      `;
      dom.financePreviewTbody.appendChild(emptyTr);
      if (dom.financeEmptyState) dom.financeEmptyState.classList.add('hidden');
      return;
    }

    if (dom.financeEmptyState) dom.financeEmptyState.classList.add('hidden');
    dom.financePreviewTbody.innerHTML = '';
    financeState.renderedRowsCount = 0;
    renderFinanceChunk(0, CHUNK_SIZE);
  }

  // Re-renders only the tbody (keeps the header/focused input intact)
  function renderFinanceBody(displayRows) {
    if (!dom.financePreviewTbody) return;
    const cols = getFinanceVisibleColumns();
    if (!displayRows || !displayRows.length) {
      dom.financePreviewTbody.innerHTML = '';
      const colSpan = Math.max(1, cols.length + 1);
      const emptyTr = document.createElement('tr');
      emptyTr.className = 'empty-table-row';
      emptyTr.innerHTML = `<td colspan="${colSpan}" style="text-align: center; padding: 48px 16px;"><p style="font-size: 13px; color: var(--ink3);">No matching finance records</p></td>`;
      dom.financePreviewTbody.appendChild(emptyTr);
      return;
    }
    dom.financePreviewTbody.innerHTML = '';
    financeState.renderedRowsCount = 0;
    renderFinanceChunk(0, CHUNK_SIZE);
  }

  function renderFinanceChunk(startIdx, count) {
    if (!dom.financePreviewTbody) return;
    const cols = getFinanceVisibleColumns();
    const displayRows = financeState.filteredRows;
    const end = Math.min(startIdx + count, displayRows.length);
    const frag = document.createDocumentFragment();

    for (let idx = startIdx; idx < end; idx++) {
      const r = displayRows[idx];
      // A row without a Dataverse ID is addressed by its position in ALL rows
      // (what the Python side indexes), never its position in the filtered or
      // sorted view, which pointed edits and deletes at a different row.
      const fullIdx = financeState.rows.indexOf(r);
      const idOrIdx = r.__id || (fullIdx >= 0 ? fullIdx : idx);
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

        const rowKey = r.__id || String(idx);
        const isDirty = Boolean(
          (financeState.stagedEdits && financeState.stagedEdits[rowKey]?.[h] !== undefined) ||
          (r.__id && financeState.stagedEdits && financeState.stagedEdits[r.__id]?.[h] !== undefined)
        );
        if (isDirty) {
          td.classList.add('is-dirty');
          const baseRow = financeState.baselineRows && (financeState.baselineRows.find(b => b.__id && b.__id === r.__id) || financeState.baselineRows[idx]);
          const prevVal = baseRow ? (baseRow[h] ?? '') : '';
          const currVal = (financeState.stagedEdits[rowKey]?.[h] !== undefined) ? financeState.stagedEdits[rowKey][h] : (r[h] ?? '');
          const lotId = r['Area Index'] || r['Corridor Index'] || r['cr63f_areaindex'] || rowKey;
          td.dataset.prevValue = String(prevVal);
          td.dataset.currValue = String(currVal);
          td.dataset.fieldName = h;
          td.dataset.diffType = 'manual';
          td.dataset.lotId = String(lotId);
        }

        const input = document.createElement('input');
        input.className = 'cell-input';
        if (isDirty) input.classList.add('modified');
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
            input.value = isBlankAmount(input.value) ? '' : String(pNum);
          }
        });

        input.addEventListener('blur', () => {
          if (isMoney) {
            const num = parseAmount(input.value);
            const typedBlank = isBlankAmount(input.value);
                  input.value = typedBlank ? '' : formatPrice(num);
            const cleanStr = typedBlank ? '' : String(num);
            const prevNum = parseAmount(before);
            const prevStr = isBlankAmount(before) ? '' : String(prevNum);
            if (cleanStr !== prevStr) {
              handleFinanceCellEdit(idOrIdx, h, cleanStr, input, td);
            }
          } else if (isDate) {
            const normalized = commitDateInput(input, before);
            if (normalized !== null && normalized !== (normalizeDateInput(before) ?? before)) {
              handleFinanceCellEdit(idOrIdx, h, normalized, input, td);
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



      frag.appendChild(tr);
    }

    dom.financePreviewTbody.appendChild(frag);
    financeState.renderedRowsCount = end;
    invalidateGridCache('financePreviewTable');
  }

  function updateStagedEditsUIFinance() {
    if (!dom.stagedEditsWrapFinance) return;
    const count = Object.keys(financeState.stagedEdits || {}).length;
    if (count > 0) {
      dom.stagedEditsWrapFinance.classList.remove('hidden');
      if (dom.unsavedEditsChipFinance) {
        dom.unsavedEditsChipFinance.textContent = `${count.toLocaleString()} modified row${count > 1 ? 's' : ''}`;
      }
      if (dom.btnToggleShowOnlyChangedFinance) {
        dom.btnToggleShowOnlyChangedFinance.textContent = financeState.showOnlyChanged ? 'Show All' : 'Show Changed Only';
        dom.btnToggleShowOnlyChangedFinance.className = financeState.showOnlyChanged ? 'btn xs primary' : 'btn xs ghost';
      }
    } else {
      dom.stagedEditsWrapFinance.classList.add('hidden');
      financeState.showOnlyChanged = false;
      if (dom.btnToggleShowOnlyChangedFinance) {
        dom.btnToggleShowOnlyChangedFinance.textContent = 'Show Changed Only';
        dom.btnToggleShowOnlyChangedFinance.className = 'btn xs ghost';
      }
    }
  }

  function toggleShowOnlyChangedFinance() {
    financeState.showOnlyChanged = !financeState.showOnlyChanged;
    updateStagedEditsUIFinance();
    filterAndRenderFinance();
  }

  async function saveStagedEditsFinance() {
    const rowIds = Object.keys(financeState.stagedEdits || {});
    if (rowIds.length === 0) {
      showToast('No unsaved finance changes to save.', 'info');
      return;
    }

    const confirmed = await showConfirmModal({
      title: 'Save Finance Changes to Dataverse',
      subtitle: 'Finance Reference Sync & Audit Logging',
      message: `You are about to save changes for ${rowIds.length.toLocaleString()} modified finance record(s) to Dataverse. All updated values will be recorded in cr63f_financereference and logged in the Entry Logs audit trail.`,
      confirmText: 'Save to Dataverse & Log',
      cancelText: 'Cancel',
      isDanger: false,
      keepOpenOnConfirm: true
    });

    if (!confirmed) return;

    try {
      if (dom.btnConfirmModalOk) {
        dom.btnConfirmModalOk.disabled = true;
        dom.btnConfirmModalOk.textContent = 'Saving...';
      }
      if (dom.btnConfirmModalCancel) dom.btnConfirmModalCancel.classList.add('hidden');
      if (dom.btnConfirmModalClose) dom.btnConfirmModalClose.classList.add('hidden');

      setConfirmModalProgress(0, rowIds.length, `Preparing to save ${rowIds.length.toLocaleString()} finance record(s)...`);

      if (dom.btnSaveStagedEditsFinance) {
        dom.btnSaveStagedEditsFinance.disabled = true;
        dom.btnSaveStagedEditsFinance.textContent = 'Saving...';
      }

      const stagedList = rowIds.map(id => ({
        id: id,
        fields: financeState.stagedEdits[id]
      }));

      const res = await window.pywebview.api.save_staged_finance_records(stagedList);
      applyReferenceSaveResult(financeState, res);
      updateStagedEditsUIFinance();
      filterAndRenderFinance();
      if (res && res.logged_count) fetchEntryLogs();
      if (res && res.success) {
        closeConfirmModal(false);
        showToast(res.message || `Successfully saved ${res.saved_count || rowIds.length} finance record(s) to Dataverse.`, 'ok');

        financeState.showOnlyChanged = false;
        updateStagedEditsUIFinance();
        filterAndRenderFinance();

        if (typeof fetchEntryLogs === 'function') {
          fetchEntryLogs();
        }
      } else {
        throw new Error((res && res.error) || 'Failed to save finance edits.');
      }
    } catch (err) {
      closeConfirmModal(false);
      showToast('Save error: ' + (err.message || err), 'err');
    } finally {
      if (dom.btnSaveStagedEditsFinance) {
        dom.btnSaveStagedEditsFinance.disabled = false;
        dom.btnSaveStagedEditsFinance.textContent = 'Save to Dataverse';
      }
    }
  }

  async function discardStagedEditsFinance() {
    const rowCount = Object.keys(financeState.stagedEdits || {}).length;
    if (rowCount > 0) {
      const ok = await showConfirmModal({
        title: 'Discard Unsaved Finance Edits',
        subtitle: 'Revert Staged Changes',
        message: `Are you sure you want to discard all staged changes for ${rowCount.toLocaleString()} finance record(s)? This will revert all modified fields back to their loaded values.`,
        confirmText: 'Discard Changes',
        cancelText: 'Cancel',
        isDanger: true
      });
      if (!ok) return;
    }

    if (financeState.baselineRows && financeState.baselineRows.length) {
      financeState.rows = JSON.parse(JSON.stringify(financeState.baselineRows));
    }
    financeState.stagedEdits = {};
    financeState.showOnlyChanged = false;
    updateStagedEditsUIFinance();
    filterAndRenderFinance();

    try {
      await window.pywebview.api.revert_staged_finance_records(financeState.rows);
    } catch (e) {
      console.warn('Could not revert engine finance rows:', e);
    }
    showToast('Discarded all unsaved finance changes.', 'info');
  }

  async function handleFinanceCellEdit(rowIdOrIdx, colKey, newVal, inputEl, tdEl) {
    let targetRow = null;
    let targetIdx = -1;
    for (let i = 0; i < financeState.rows.length; i++) {
      if (financeState.rows[i].__id === rowIdOrIdx || i === rowIdOrIdx || String(i) === String(rowIdOrIdx)) {
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

    const rowKey = targetRow.__id || String(targetIdx);
    const baseRow = financeState.baselineRows && (financeState.baselineRows.find(r => r.__id && r.__id === targetRow.__id) || financeState.baselineRows[targetIdx]);
    const baseVal = baseRow ? String(baseRow[colKey] ?? '').trim() : '';

    if (trimmed !== baseVal) {
      if (!financeState.stagedEdits[rowKey]) financeState.stagedEdits[rowKey] = {};
      financeState.stagedEdits[rowKey][colKey] = trimmed;
      if (tdEl) tdEl.classList.add('is-dirty');
      if (inputEl) inputEl.classList.add('modified');
      if (tdEl) {
        tdEl.dataset.prevValue = String(baseVal);
        tdEl.dataset.currValue = String(trimmed);
        tdEl.dataset.fieldName = colKey;
        tdEl.dataset.diffType = 'manual';
        const lotId = targetRow['Area Index'] || targetRow['Corridor Index'] || targetRow['cr63f_areaindex'] || rowKey;
        tdEl.dataset.lotId = String(lotId);
      }
    } else {
      if (financeState.stagedEdits[rowKey]) {
        delete financeState.stagedEdits[rowKey][colKey];
        if (Object.keys(financeState.stagedEdits[rowKey]).length === 0) {
          delete financeState.stagedEdits[rowKey];
        }
      }
      if (tdEl) {
        tdEl.classList.remove('is-dirty');
        delete tdEl.dataset.prevValue;
        delete tdEl.dataset.currValue;
        delete tdEl.dataset.fieldName;
        delete tdEl.dataset.diffType;
        delete tdEl.dataset.lotId;
      }
      if (inputEl) inputEl.classList.remove('modified');
    }

    updateStagedEditsUIFinance();

    try {
      await window.pywebview.api.update_finance_cell(targetRow.__id || targetIdx, colKey, trimmed);
    } catch (e) {
      console.warn('Could not sync finance edit to engine:', e);
    }
  }

  async function exportFinanceExcel() {
    const toExport = financeState.filteredRows;
    if (!toExport.length) {
      showToast('No finance data to export.', 'err');
      return;
    }
    const cols = getFinanceVisibleColumns();
    await runExport(() => window.pywebview.api.export_records_excel(toExport, 'Finance_Data_Export.xlsx', cols));
  }

  async function downloadFinanceTemplate() {
    try {
      const res = await window.pywebview.api.download_finance_template();
      if (res.success) {
        showToast(`Template downloaded to ${res.file_path}`, 'ok');
      } else if (!res.cancelled) {
        showToast(res.error || 'Failed to download finance template.', 'err');
      }
    } catch (err) {
      showToast('Error: ' + err, 'err');
    }
  }

  async function syncFinanceToDataverse() {
    if (!financeState.rows || !financeState.rows.length) {
      showToast('No finance data loaded to sync.', 'err');
      return;
    }
    // Edited cells are saved through "Save" (one audit row per changed field);
    // a sync would upload them with only a summary audit row.
    if (Object.keys(financeState.stagedEdits || {}).length) {
      showToast('Save or discard your edited finance rows first, then sync.', 'err');
      return;
    }

    const rowCount = financeState.rows.length.toLocaleString();
    const confirmed = await showConfirmModal({
      title: 'Sync Finance Data to Dataverse',
      subtitle: 'Table: cr63f_financereference',
      message: `Sync ${rowCount} finance row(s) to Dataverse? Rows that already exist (same TYPE, Corridor Index, Area Index and Title) are updated, new rows are created, and rows whose key matches several existing rows are skipped and listed.`,
      confirmText: 'Upload & Sync',
      cancelText: 'Cancel',
      isDanger: false,
      keepOpenOnConfirm: true
    });

    if (!confirmed) return;

    try {
      // Switch modal to live progress bar mode
      if (dom.confirmModal) {
        dom.confirmModal.classList.remove('hidden');
      }
      if (dom.confirmModalTitle) dom.confirmModalTitle.textContent = 'Syncing Finance Data to Dataverse';
      if (dom.confirmModalSubtitle) dom.confirmModalSubtitle.textContent = 'Dataverse Reference Sync (cr63f_financereference)';
      if (dom.confirmModalMessage) dom.confirmModalMessage.textContent = `Uploading ${rowCount} finance record(s)... Please do not close the window.`;

      if (dom.confirmModalProgressWrap) {
        dom.confirmModalProgressWrap.classList.remove('hidden');
      }
      if (dom.confirmModalProgressBar) {
        dom.confirmModalProgressBar.style.width = '0%';
      }
      if (dom.confirmModalProgressStatus) {
        dom.confirmModalProgressStatus.textContent = `Preparing ${rowCount} records...`;
      }
      if (dom.confirmModalProgressPercent) {
        dom.confirmModalProgressPercent.textContent = '0%';
      }
      if (dom.btnConfirmModalOk) {
        dom.btnConfirmModalOk.disabled = true;
        dom.btnConfirmModalOk.textContent = 'Syncing...';
      }
      if (dom.btnConfirmModalCancel) {
        dom.btnConfirmModalCancel.classList.add('hidden');
      }
      if (dom.btnConfirmModalClose) {
        dom.btnConfirmModalClose.classList.add('hidden');
      }
      if (dom.btnSyncFinanceToDv) {
        dom.btnSyncFinanceToDv.disabled = true;
        dom.btnSyncFinanceToDv.textContent = 'Syncing...';
      }

      showToast('Syncing Finance Data to Dataverse...', 'info');

      const res = await window.pywebview.api.sync_uploaded_finance_to_dataverse();
      if (res.success || res.synced_count) {
        if (dom.confirmModalProgressBar) dom.confirmModalProgressBar.style.width = '100%';
        if (dom.confirmModalProgressPercent) dom.confirmModalProgressPercent.textContent = '100%';
        if (dom.confirmModalProgressStatus) {
          dom.confirmModalProgressStatus.textContent = res.message || 'Sync finished.';
        }
        showToast(res.success ? res.message : `${res.message} ${res.error || ''}`.trim(), res.success ? (res.skipped_count ? 'info' : 'ok') : 'err');
        if (dom.lblActiveFinanceSource) {
          dom.lblActiveFinanceSource.textContent = `Dataverse: cr63f_financereference (${res.synced_count} records)`;
        }
        // Reload the rows so they carry the Dataverse IDs the sync just assigned
        // (later edits are then saved to the right row).
        try {
          const prev = await window.pywebview.api.get_finance_preview();
          if (prev && Array.isArray(prev.rows)) renderFinancePreview(financeState.headers, prev.rows);
        } catch (_) {}
        financeState.baselineRows = JSON.parse(JSON.stringify(financeState.rows));
        financeState.stagedEdits = {};
        financeState.showOnlyChanged = false;
        updateStagedEditsUIFinance();
        filterAndRenderFinance();
        setTimeout(() => {
          if (dom.confirmModal) dom.confirmModal.classList.add('hidden');
          if (dom.btnConfirmModalClose) dom.btnConfirmModalClose.classList.remove('hidden');
        }, 1200);
      } else {
        if (dom.confirmModal) dom.confirmModal.classList.add('hidden');
        if (dom.btnConfirmModalClose) dom.btnConfirmModalClose.classList.remove('hidden');
        showToast(res.error || 'Failed to sync finance data to Dataverse.', 'err');
      }
    } catch (err) {
      if (dom.confirmModal) dom.confirmModal.classList.add('hidden');
      if (dom.btnConfirmModalClose) dom.btnConfirmModalClose.classList.remove('hidden');
      showToast('Error: ' + err, 'err');
    } finally {
      if (dom.btnSyncFinanceToDv) {
        dom.btnSyncFinanceToDv.disabled = false;
        dom.btnSyncFinanceToDv.textContent = 'Sync to Dataverse';
      }
      if (dom.btnConfirmModalOk) {
        dom.btnConfirmModalOk.disabled = false;
        dom.btnConfirmModalOk.textContent = 'Confirm';
      }
      if (dom.btnConfirmModalCancel) {
        dom.btnConfirmModalCancel.classList.remove('hidden');
      }
      if (dom.btnConfirmModalClose) {
        dom.btnConfirmModalClose.classList.remove('hidden');
      }
    }
  }

  window.deleteFinanceRecord = async function (idOrIdx) {
    const ok = await showConfirmModal({
      title: 'Delete Finance Record',
      subtitle: 'Finance Reference Deletion',
      message: 'Are you sure you want to delete this record from Finance Data?',
      confirmText: 'Delete',
      cancelText: 'Cancel',
      isDanger: true
    });
    if (!ok) return;

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
        if (dom.lblTopFinance) dom.lblTopFinance.textContent = state.financeLoadedCount > 0 ? `${state.financeLoadedCount.toLocaleString()} rows` : 'Not loaded';
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

    let financeFilterDebounceTimer = null;
    dom.financePreviewThead.querySelectorAll('.th-filter-input').forEach(input => {
      input.addEventListener('input', () => {
        const colKey = input.dataset.finFilterCol;
        financeState.columnFilters[colKey] = input.value.trim().toLowerCase();
        if (financeFilterDebounceTimer) clearTimeout(financeFilterDebounceTimer);
        financeFilterDebounceTimer = setTimeout(() => {
          filterAndRenderFinance(true); // skipHeadRender keeps the focused input alive
        }, 120);
      });
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
          const colKey = input.dataset.finFilterCol;
          financeState.columnFiltersOpen[colKey] = false;
          const popup = document.getElementById(`finColFilterBox_${colKey}`);
          if (popup) popup.classList.add('hidden');
        }
        if (e.key === 'Enter') e.preventDefault();
        e.stopPropagation();
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
        saveFinanceTableConfig();
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

  // ==================== CONVERSION DATA VIEW ====================
  function setupConversionView() {
    if (dom.btnUploadConversion) {
      dom.btnUploadConversion.addEventListener('click', async () => {
        const stagedCount = Object.keys(conversionState.stagedEdits || {}).length;
        if (stagedCount > 0) {
          const ok = await showConfirmModal({
            title: 'Unsaved Conversion Changes',
            subtitle: 'Upload Conflict',
            message: `You have ${stagedCount} unsaved modified row(s) in Conversion Data. Uploading a new file will overwrite these changes. Do you want to continue and discard unsaved edits?`,
            confirmText: 'Discard & Upload',
            cancelText: 'Cancel',
            isDanger: true
          });
          if (!ok) return;
        }
        const res = await window.pywebview.api.load_conversion_file_dialog();
        if (res.success) {
          showToast(`Loaded ${res.row_count} rows from ${res.source}.`, 'ok');
          const s = await window.pywebview.api.get_initial_state();
          applyState(s);
          renderConversionPreview(res.headers, res.sample_rows);
        } else if (!res.cancelled) {
          showToast(res.error || 'Failed to load file.', 'err');
        }
      });
    }

    if (dom.btnClearConversionData) {
      dom.btnClearConversionData.addEventListener('click', async () => {
        const stagedCount = Object.keys(conversionState.stagedEdits || {}).length;
        if (stagedCount > 0) {
          const ok = await showConfirmModal({
            title: 'Clear Conversion Data',
            subtitle: 'Unsaved Edits',
            message: `You have ${stagedCount} unsaved modified row(s). Clearing will discard them. Are you sure?`,
            confirmText: 'Clear Data',
            cancelText: 'Cancel',
            isDanger: true
          });
          if (!ok) return;
        }
        await window.pywebview.api.clear_conversion_reference();
        const s = await window.pywebview.api.get_initial_state();
        applyState(s);
        conversionState.headers = [];
        conversionState.rows = [];
        conversionState.baselineRows = [];
        conversionState.stagedEdits = {};
        conversionState.showOnlyChanged = false;
        conversionState.filteredRows = [];
        conversionState.renderedRowsCount = 0;
        updateStagedEditsUIConversion();
        if (dom.lblConversionGridCount) dom.lblConversionGridCount.textContent = '0 rows';
        renderConversionTable();
        showToast('Conversion data cleared.', 'ok');
      });
    }

    if (dom.btnToggleShowOnlyChangedConversion) {
      dom.btnToggleShowOnlyChangedConversion.addEventListener('click', toggleShowOnlyChangedConversion);
    }
    if (dom.btnSaveStagedEditsConversion) {
      dom.btnSaveStagedEditsConversion.addEventListener('click', saveStagedEditsConversion);
    }
    if (dom.btnDiscardStagedEditsConversion) {
      dom.btnDiscardStagedEditsConversion.addEventListener('click', discardStagedEditsConversion);
    }

    if (dom.searchInputConversion) {
      dom.searchInputConversion.addEventListener('input', () => {
        clearTimeout(searchConvDebounce);
        searchConvDebounce = setTimeout(filterAndRenderConversion, 120);
      });
    }
    if (dom.statusFilterConversion) {
      dom.statusFilterConversion.addEventListener('change', () => filterAndRenderConversion());
    }
    if (dom.btnClearFiltersConversion) {
      dom.btnClearFiltersConversion.addEventListener('click', () => {
        if (dom.searchInputConversion) dom.searchInputConversion.value = '';
        if (dom.statusFilterConversion) dom.statusFilterConversion.value = 'ALL';
        conversionState.columnFilters = {};
        conversionState.columnFiltersOpen = {};
        filterAndRenderConversion();
        showToast('Conversion filters cleared.', 'ok');
      });
    }
    if (dom.btnDownloadConversionTemplate) {
      dom.btnDownloadConversionTemplate.addEventListener('click', downloadConversionTemplate);
    }
    if (dom.btnSyncConversionToDv) {
      dom.btnSyncConversionToDv.addEventListener('click', syncConversionToDataverse);
    }
    if (dom.btnExportExcelConversion) {
      dom.btnExportExcelConversion.addEventListener('click', exportConversionExcel);
    }

    setupConversionColumnPicker();
  }

  function setupConversionColumnPicker() {
    if (!dom.btnColumnsPickerConversion || !dom.colPickerMenuConversion) return;

    dom.btnColumnsPickerConversion.addEventListener('click', (e) => {
      e.stopPropagation();
      dom.colPickerMenuConversion.classList.toggle('hidden');
    });

    document.addEventListener('click', (e) => {
      if (!e.target.closest('#view-conversion .col-picker-container')) {
        dom.colPickerMenuConversion.classList.add('hidden');
      }
    });
  }

  function renderConversionColumnPickerMenu() {
    if (!dom.colPickerMenuConversion) return;

    const cols = conversionState.colOrder.length ? conversionState.colOrder : conversionState.headers;
    let html = '';
    cols.forEach((colKey, idx) => {
      const isChecked = conversionState.visibleCols.has(colKey);
      const isFirst = idx === 0;
      const isLast = idx === cols.length - 1;

      html += `
        <div class="col-picker-item">
          <label class="col-picker-label">
            <input type="checkbox" data-conv-col-toggle="${escapeHtml(colKey)}" ${isChecked ? 'checked' : ''} />
            <span>${escapeHtml(colKey)}</span>
          </label>
          <div class="col-reorder-btns">
            <button type="button" class="col-reorder-btn" data-conv-move-up="${escapeHtml(colKey)}" ${isFirst ? 'disabled' : ''} title="Move left">◄</button>
            <button type="button" class="col-reorder-btn" data-conv-move-down="${escapeHtml(colKey)}" ${isLast ? 'disabled' : ''} title="Move right">►</button>
          </div>
        </div>
      `;
    });

    html += `
      <div class="col-picker-actions">
        <button type="button" class="btn xs ghost" id="btnResetConversionLayout" style="width: 100%; font-size: 11px;">Reset Columns Layout</button>
      </div>
    `;

    dom.colPickerMenuConversion.innerHTML = html;

    dom.colPickerMenuConversion.querySelectorAll('[data-conv-col-toggle]').forEach(chk => {
      chk.addEventListener('change', () => {
        const key = chk.dataset.convColToggle;
        if (chk.checked) {
          conversionState.visibleCols.add(key);
        } else {
          if (conversionState.visibleCols.size > 1) {
            conversionState.visibleCols.delete(key);
          } else {
            chk.checked = true;
            showToast('At least one column must remain visible.', 'err');
            return;
          }
        }
        saveConversionTableConfig();
        filterAndRenderConversion();
      });
    });

    dom.colPickerMenuConversion.querySelectorAll('[data-conv-move-up]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.convMoveUp;
        moveConversionColumnOrder(key, -1);
      });
    });

    dom.colPickerMenuConversion.querySelectorAll('[data-conv-move-down]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.convMoveDown;
        moveConversionColumnOrder(key, 1);
      });
    });

    const btnReset = dom.colPickerMenuConversion.querySelector('#btnResetConversionLayout');
    if (btnReset) {
      btnReset.addEventListener('click', (e) => {
        e.stopPropagation();
        localStorage.removeItem('conv_col_widths');
        localStorage.removeItem('conv_col_order');
        localStorage.removeItem('conv_visible_cols');
        localStorage.removeItem('conv_frozen_cols');
        localStorage.removeItem('conv_sort_col');
        localStorage.removeItem('conv_sort_dir');
        conversionState.visibleCols = new Set(conversionState.headers);
        conversionState.colOrder = [...conversionState.headers];
        conversionState.colWidths = {};
        conversionState.frozenCols = new Set();
        conversionState.sortCol = null;
        conversionState.sortDir = null;
        saveConversionTableConfig();
        renderConversionColumnPickerMenu();
        filterAndRenderConversion();
        showToast('Conversion layout reset.', 'ok');
      });
    }
  }

  function moveConversionColumnOrder(colKey, direction) {
    const idx = conversionState.colOrder.indexOf(colKey);
    if (idx === -1) return;
    const targetIdx = idx + direction;
    if (targetIdx < 0 || targetIdx >= conversionState.colOrder.length) return;

    const temp = conversionState.colOrder[idx];
    conversionState.colOrder[idx] = conversionState.colOrder[targetIdx];
    conversionState.colOrder[targetIdx] = temp;

    saveConversionTableConfig();
    renderConversionColumnPickerMenu();
    filterAndRenderConversion();
  }

  function getConversionDarStatus(row) {
    const entries = Object.entries(row || {});
    for (const acceptedKeys of [new Set(['darstatus', 'cr63fdarstatus']), new Set(['conversionstatus', 'cr63fconversionstatus'])]) {
      for (const [key, value] of entries) {
        const normalizedKey = String(key).toLowerCase().replace(/[^a-z0-9]/g, '');
        if (acceptedKeys.has(normalizedKey) && value != null && String(value).trim()) return String(value).trim();
      }
    }
    return '';
  }

  function renderConversionPreview(headers, rows) {
    if (!headers || !rows) return;
    const cleanHeaders = headers.filter(h => !isSystemOrGeneratedCol(h));

    conversionState.headers = cleanHeaders;
    conversionState.rows = rows;
    conversionState.baselineRows = JSON.parse(JSON.stringify(rows));
    conversionState.stagedEdits = {};
    conversionState.showOnlyChanged = false;
    updateStagedEditsUIConversion();

    try {
      const cw = localStorage.getItem('conv_col_widths');
      if (cw) conversionState.colWidths = JSON.parse(cw);
    } catch (_) {}

    try {
      const co = localStorage.getItem('conv_col_order');
      if (co) {
        const parsed = JSON.parse(co);
        const valid = parsed.filter(k => cleanHeaders.includes(k) && !isSystemOrGeneratedCol(k));
        cleanHeaders.forEach(k => { if (!valid.includes(k)) valid.push(k); });
        conversionState.colOrder = valid;
      } else {
        conversionState.colOrder = [...cleanHeaders];
      }
    } catch (_) {
      conversionState.colOrder = [...cleanHeaders];
    }

    try {
      const cv = localStorage.getItem('conv_visible_cols');
      if (cv) {
        const parsed = JSON.parse(cv).filter(k => cleanHeaders.includes(k) && !isSystemOrGeneratedCol(k));
        conversionState.visibleCols = new Set(parsed);
        cleanHeaders.forEach(k => {
          if (!conversionState.visibleCols.has(k)) conversionState.visibleCols.add(k);
        });
      } else {
        conversionState.visibleCols = new Set(cleanHeaders);
      }
    } catch (_) {
      conversionState.visibleCols = new Set(cleanHeaders);
    }

    try {
      const cf = localStorage.getItem('conv_frozen_cols');
      conversionState.frozenCols = cf ? new Set(JSON.parse(cf)) : new Set();
    } catch (_) {
      conversionState.frozenCols = new Set();
    }

    try {
      const sc = localStorage.getItem('conv_sort_col');
      const sd = localStorage.getItem('conv_sort_dir');
      conversionState.sortCol = sc || null;
      conversionState.sortDir = sd || null;
    } catch (_) {
      conversionState.sortCol = null;
      conversionState.sortDir = null;
    }

    // Populate the Conversion filter from the DAR Status field only.
    if (dom.statusFilterConversion) {
      const currentStatus = String(dom.statusFilterConversion.value || 'ALL').trim().toLowerCase();
      const statusSet = new Set();
      rows.forEach(r => {
        const val = getConversionDarStatus(r);
        if (val) statusSet.add(val);
      });
      if (statusSet.size > 0) {
        let optHtml = '<option value="ALL">All DAR Status</option>';
        Array.from(statusSet).sort().forEach(st => {
          optHtml += `<option value="${escapeHtml(st)}">${escapeHtml(st)}</option>`;
        });
        dom.statusFilterConversion.innerHTML = optHtml;
        const matchingStatus = Array.from(statusSet).find(st => st.toLowerCase() === currentStatus);
        dom.statusFilterConversion.value = matchingStatus || 'ALL';
      } else {
        dom.statusFilterConversion.innerHTML = '<option value="ALL">All DAR Status</option>';
      }
    }

    renderConversionColumnPickerMenu();
    filterAndRenderConversion();
  }

  function filterAndRenderConversion(skipHeadRender = false) {
    const q = dom.searchInputConversion ? dom.searchInputConversion.value.toLowerCase().trim() : '';
    const statusFilt = dom.statusFilterConversion ? dom.statusFilterConversion.value : 'ALL';
    const colFilterKeys = Object.keys(conversionState.columnFilters).filter(k => !!conversionState.columnFilters[k]);

    conversionState.filteredRows = conversionState.rows.filter(r => {
      if (statusFilt !== 'ALL') {
        if (getConversionDarStatus(r).toLowerCase() !== statusFilt.trim().toLowerCase()) return false;
      }

      for (const colKey of colFilterKeys) {
        const query = conversionState.columnFilters[colKey];
        const val = String(r[colKey] ?? '').toLowerCase();
        if (!val.includes(query)) return false;
      }

      if (!q) {
        if (conversionState.showOnlyChanged) {
          const id = r.__id;
          const rowIdx = String(conversionState.rows.indexOf(r));
          return Boolean((id && conversionState.stagedEdits[id]) || conversionState.stagedEdits[rowIdx]);
        }
        return true;
      }
      const matchesQ = Object.values(r).some(v => String(v).toLowerCase().includes(q));
      if (!matchesQ) return false;
      if (conversionState.showOnlyChanged) {
        const id = r.__id;
        const rowIdx = String(conversionState.rows.indexOf(r));
        return Boolean((id && conversionState.stagedEdits[id]) || conversionState.stagedEdits[rowIdx]);
      }
      return true;
    });

    if (conversionState.sortCol && conversionState.sortDir) {
      const col = conversionState.sortCol;
      const mult = conversionState.sortDir === 'asc' ? 1 : -1;
      const colDef = {
        key: col,
        isDate: /DATE/i.test(col),
        isNum: /AREA|HA/i.test(col)
      };
      conversionState.filteredRows.sort((a, b) => {
        const valA = a[col] ?? '';
        const valB = b[col] ?? '';
        return compareTableValues(valA, valB, colDef, mult);
      });
    }

    if (dom.lblConversionGridCount) {
      dom.lblConversionGridCount.textContent = `${conversionState.filteredRows.length.toLocaleString()} rows`;
    }

    if (skipHeadRender) {
      const displayRows = conversionState.filteredRows;
      renderConversionBody(displayRows);
    } else {
      renderConversionTable();
    }
  }

  const DEFAULT_CONVERSION_HEADERS = [
    'AREA-INDEX',
    'TITLE NUMBER',
    'LOT NUMBER/SURVEY NUMBER',
    'DAR STATUS',
    'NIA STATUS',
    'CONVERSION CLASS',
    'DAR BATCH',
    'DAR UPDATE',
    'ENTITY',
    'MUNICIPALITY/CITY',
    'BARANGAY',
    'AREA APPLIED (HAS.)'
  ];

  function getConversionVisibleColumns() {
    let order = conversionState.colOrder.length ? conversionState.colOrder : conversionState.headers;
    if (!order.length) order = DEFAULT_CONVERSION_HEADERS;
    return order.filter(h => (!conversionState.visibleCols.size || conversionState.visibleCols.has(h)) && !isSystemOrGeneratedCol(h));
  }

  function renderConversionTable() {
    if (!dom.conversionPreviewThead || !dom.conversionPreviewTbody) return;
    const cols = getConversionVisibleColumns();

    let thead = '<tr>';
    thead += `<th style="width: 45px; min-width: 45px; max-width: 45px; text-align: center; position: sticky !important; top: 0 !important; z-index: 25 !important; background: var(--subtle) !important;">#</th>`;

    cols.forEach((h, colIdx) => {
      const w = conversionState.colWidths[h] || 140;
      const isSorted = conversionState.sortCol === h;
      const sortIcon = isSorted ? (conversionState.sortDir === 'asc' ? '▲' : '▼') : '';
      const isFrozen = conversionState.frozenCols.has(h);
      const isLastFrz = isTrailingFrozen(colIdx, cols, conversionState.frozenCols);
      const leftOffset = isFrozen ? (45 + getFrozenOffset(colIdx, cols, conversionState.frozenCols, conversionState.colWidths)) : 0;
      const frozenClass = `${isFrozen ? 'col-frozen' : ''} ${isLastFrz ? 'col-frozen-last' : ''}`.trim();
      const leftStyle = isFrozen ? `left: ${leftOffset}px;` : '';
      const isFilterActive = !!(conversionState.columnFilters[h]);
      const isFilterOpen = !!(conversionState.columnFiltersOpen[h]);

      thead += `
        <th draggable="true" data-conv-col="${escapeHtml(h)}" class="${frozenClass}" style="width:${w}px; min-width:${w}px; max-width:${w}px; ${leftStyle}">
          <div class="th-header-box">
            <div class="th-title-wrap">
              <button type="button" class="th-pin-btn ${isFrozen ? 'pinned' : ''}" data-conv-pin="${escapeHtml(h)}" aria-label="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(h)}" title="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(h)}">
                <svg viewBox="0 0 24 24" width="11" height="11" fill="currentColor" stroke="none">
                  <path d="M16 12V4h1V2H7v2h1v8l-2 2v2h5.2v6h1.6v-6H18v-2l-2-2z"/>
                </svg>
              </button>
              <button type="button" class="th-sort-wrap" data-conv-sort="${escapeHtml(h)}" aria-label="Sort by ${escapeHtml(h)}" title="Sort · drag to reorder">
                <span class="th-col-name">${escapeHtml(h)}</span>
                <span class="th-sort-icon ${isSorted ? 'active' : ''}" aria-hidden="true">${sortIcon}</span>
              </button>
              <button type="button" class="th-chevron-btn ${isFilterActive ? 'active' : ''}" data-conv-filter-toggle="${escapeHtml(h)}" aria-label="Filter ${escapeHtml(h)}" title="Filter ${escapeHtml(h)}">
                <svg viewBox="0 0 24 24" width="10" height="10" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"></polyline></svg>
              </button>
            </div>
            <div class="th-filter-popup ${isFilterOpen ? '' : 'hidden'}" id="convColFilterBox_${escapeHtml(h)}">
              <input class="th-filter-input" data-conv-filter-col="${escapeHtml(h)}" aria-label="Filter ${escapeHtml(h)} column" autocomplete="off" value="${escapeHtml(conversionState.columnFilters[h] || '')}" placeholder="Filter ${escapeHtml(h)}…" />
            </div>
          </div>
          <div class="th-resizer" data-conv-resize="${escapeHtml(h)}" title="Drag to resize · double-click to reset"></div>
          <div class="th-bottom-resizer" title="Drag to resize the header"></div>
        </th>
      `;
    });
    thead += '</tr>';
    dom.conversionPreviewThead.innerHTML = thead;

    attachConversionHeaderEvents();

    const displayRows = conversionState.filteredRows;

    if (!displayRows.length) {
      dom.conversionPreviewTbody.innerHTML = '';
      const colSpan = Math.max(1, cols.length + 1);
      const isInitialEmpty = !conversionState.rows.length;
      const title = isInitialEmpty ? 'No Land Conversion Reference data loaded' : 'No matching conversion records';
      const desc = isInitialEmpty
        ? 'Click <strong>Upload Excel / CSV</strong> or <strong>Sync to Dataverse</strong> to load reference data.'
        : 'Try adjusting your search terms, DAR Status filter, or column filters.';

      const emptyTr = document.createElement('tr');
      emptyTr.className = 'empty-table-row';
      emptyTr.innerHTML = `
        <td colspan="${colSpan}" style="text-align: center; padding: 48px 16px; background: transparent; border: none;">
          <div style="display: inline-flex; flex-direction: column; align-items: center; justify-content: center; gap: 6px;">
            <svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" style="color: var(--ink3); opacity: 0.7;">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
              <line x1="16" y1="13" x2="8" y2="13"></line>
              <line x1="16" y1="17" x2="8" y2="17"></line>
              <polyline points="10 9 9 9 8 9"></polyline>
            </svg>
            <p style="font-size: 13.5px; font-weight: 600; color: var(--ink); margin: 4px 0 2px;">${title}</p>
            <p style="font-size: 12px; color: var(--ink3); margin: 0;">${desc}</p>
          </div>
        </td>
      `;
      dom.conversionPreviewTbody.appendChild(emptyTr);
      if (dom.conversionEmptyState) dom.conversionEmptyState.classList.add('hidden');
      return;
    }

    if (dom.conversionEmptyState) dom.conversionEmptyState.classList.add('hidden');
    dom.conversionPreviewTbody.innerHTML = '';
    conversionState.renderedRowsCount = 0;
    renderConversionChunk(0, CHUNK_SIZE);
  }

  // Re-renders only the tbody (keeps the header/focused input intact)
  function renderConversionBody(displayRows) {
    if (!dom.conversionPreviewTbody) return;
    const cols = getConversionVisibleColumns();
    if (!displayRows || !displayRows.length) {
      dom.conversionPreviewTbody.innerHTML = '';
      const colSpan = Math.max(1, cols.length + 1);
      const emptyTr = document.createElement('tr');
      emptyTr.className = 'empty-table-row';
      emptyTr.innerHTML = `<td colspan="${colSpan}" style="text-align: center; padding: 48px 16px;"><p style="font-size: 13px; color: var(--ink3);">No matching conversion records</p></td>`;
      dom.conversionPreviewTbody.appendChild(emptyTr);
      return;
    }
    dom.conversionPreviewTbody.innerHTML = '';
    conversionState.renderedRowsCount = 0;
    renderConversionChunk(0, CHUNK_SIZE);
  }

  function renderConversionChunk(startIdx, count) {
    if (!dom.conversionPreviewTbody) return;
    const cols = getConversionVisibleColumns();
    const displayRows = conversionState.filteredRows;
    const end = Math.min(startIdx + count, displayRows.length);
    const frag = document.createDocumentFragment();

    for (let idx = startIdx; idx < end; idx++) {
      const r = displayRows[idx];
      // A row without a Dataverse ID is addressed by its position in ALL rows
      // (what the Python side indexes), never its position in the filtered or
      // sorted view, which pointed edits and deletes at a different row.
      const fullIdx = conversionState.rows.indexOf(r);
      const idOrIdx = r.__id || (fullIdx >= 0 ? fullIdx : idx);
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
        const isMoney = hu.includes('PRICE') || hu.includes('TCP') || hu.includes('COST');
        const isDate = hu.includes('DATE');

        const w = conversionState.colWidths[h] || 140;
        td.style.width = `${w}px`;
        td.style.minWidth = `${w}px`;
        td.style.maxWidth = `${w}px`;

        const isFrozen = conversionState.frozenCols.has(h);
        const isLastFrz = isTrailingFrozen(colIdx, cols, conversionState.frozenCols);
        const leftOffset = isFrozen ? (45 + getFrozenOffset(colIdx, cols, conversionState.frozenCols, conversionState.colWidths)) : 0;

        if (isFrozen) {
          td.className = `${isFrozen ? 'col-frozen' : ''} ${isLastFrz ? 'col-frozen-last' : ''}`.trim();
          td.style.left = `${leftOffset}px`;
        }

        const rowKey = r.__id || String(idx);
        const isDirty = Boolean(
          (conversionState.stagedEdits && conversionState.stagedEdits[rowKey]?.[h] !== undefined) ||
          (r.__id && conversionState.stagedEdits && conversionState.stagedEdits[r.__id]?.[h] !== undefined)
        );
        if (isDirty) {
          td.classList.add('is-dirty');
          const baseRow = conversionState.baselineRows && (conversionState.baselineRows.find(b => b.__id && b.__id === r.__id) || conversionState.baselineRows[idx]);
          const prevVal = baseRow ? (baseRow[h] ?? '') : '';
          const currVal = (conversionState.stagedEdits[rowKey]?.[h] !== undefined) ? conversionState.stagedEdits[rowKey][h] : (r[h] ?? '');
          const lotId = r['AREA-INDEX'] || r['cr63f_areaindex'] || r['TITLE NUMBER'] || rowKey;
          td.dataset.prevValue = String(prevVal);
          td.dataset.currValue = String(currVal);
          td.dataset.fieldName = h;
          td.dataset.diffType = 'manual';
          td.dataset.lotId = String(lotId);
        }

        const input = document.createElement('input');
        input.className = 'cell-input';
        if (isDirty) input.classList.add('modified');
        if (isMoney) input.classList.add('num-cell');

        if (isMoney) {
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
            input.value = isBlankAmount(input.value) ? '' : String(pNum);
          }
        });

        input.addEventListener('blur', () => {
          if (isMoney) {
            const num = parseAmount(input.value);
            const typedBlank = isBlankAmount(input.value);
                  input.value = typedBlank ? '' : formatPrice(num);
            const cleanStr = typedBlank ? '' : String(num);
            const prevNum = parseAmount(before);
            const prevStr = isBlankAmount(before) ? '' : String(prevNum);
            if (cleanStr !== prevStr) {
              handleConversionCellEdit(idOrIdx, h, cleanStr, input, td);
            }
          } else if (isDate) {
            const normalized = commitDateInput(input, before);
            if (normalized !== null && normalized !== (normalizeDateInput(before) ?? before)) {
              handleConversionCellEdit(idOrIdx, h, normalized, input, td);
            }
          } else {
            if (input.value !== before) {
              handleConversionCellEdit(idOrIdx, h, input.value, input, td);
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


      frag.appendChild(tr);
    }

    dom.conversionPreviewTbody.appendChild(frag);
    conversionState.renderedRowsCount = end;
    invalidateGridCache('conversionPreviewTable');
  }

  function updateStagedEditsUIConversion() {
    if (!dom.stagedEditsWrapConversion) return;
    const count = Object.keys(conversionState.stagedEdits || {}).length;
    if (count > 0) {
      dom.stagedEditsWrapConversion.classList.remove('hidden');
      if (dom.unsavedEditsChipConversion) {
        dom.unsavedEditsChipConversion.textContent = `${count.toLocaleString()} modified row${count > 1 ? 's' : ''}`;
      }
      if (dom.btnToggleShowOnlyChangedConversion) {
        dom.btnToggleShowOnlyChangedConversion.textContent = conversionState.showOnlyChanged ? 'Show All' : 'Show Changed Only';
        dom.btnToggleShowOnlyChangedConversion.className = conversionState.showOnlyChanged ? 'btn xs primary' : 'btn xs ghost';
      }
    } else {
      dom.stagedEditsWrapConversion.classList.add('hidden');
      conversionState.showOnlyChanged = false;
      if (dom.btnToggleShowOnlyChangedConversion) {
        dom.btnToggleShowOnlyChangedConversion.textContent = 'Show Changed Only';
        dom.btnToggleShowOnlyChangedConversion.className = 'btn xs ghost';
      }
    }
  }

  function toggleShowOnlyChangedConversion() {
    conversionState.showOnlyChanged = !conversionState.showOnlyChanged;
    updateStagedEditsUIConversion();
    filterAndRenderConversion();
  }

  async function saveStagedEditsConversion() {
    const rowIds = Object.keys(conversionState.stagedEdits || {});
    if (rowIds.length === 0) {
      showToast('No unsaved conversion changes to save.', 'info');
      return;
    }

    const confirmed = await showConfirmModal({
      title: 'Save Conversion Changes to Dataverse',
      subtitle: 'Conversion Reference Sync & Audit Logging',
      message: `You are about to save changes for ${rowIds.length.toLocaleString()} modified conversion record(s) to Dataverse. All updated values will be recorded in cr63f_landconversionreference and logged in the Entry Logs audit trail.`,
      confirmText: 'Save to Dataverse & Log',
      cancelText: 'Cancel',
      isDanger: false,
      keepOpenOnConfirm: true
    });

    if (!confirmed) return;

    try {
      if (dom.btnConfirmModalOk) {
        dom.btnConfirmModalOk.disabled = true;
        dom.btnConfirmModalOk.textContent = 'Saving...';
      }
      if (dom.btnConfirmModalCancel) dom.btnConfirmModalCancel.classList.add('hidden');
      if (dom.btnConfirmModalClose) dom.btnConfirmModalClose.classList.add('hidden');

      setConfirmModalProgress(0, rowIds.length, `Preparing to save ${rowIds.length.toLocaleString()} conversion record(s)...`);

      if (dom.btnSaveStagedEditsConversion) {
        dom.btnSaveStagedEditsConversion.disabled = true;
        dom.btnSaveStagedEditsConversion.textContent = 'Saving...';
      }

      const stagedList = rowIds.map(id => ({
        id: id,
        fields: conversionState.stagedEdits[id]
      }));

      const res = await window.pywebview.api.save_staged_conversion_records(stagedList);
      applyReferenceSaveResult(conversionState, res);
      updateStagedEditsUIConversion();
      filterAndRenderConversion();
      if (res && res.logged_count) fetchEntryLogs();
      if (res && res.success) {
        closeConfirmModal(false);
        showToast(res.message || `Successfully saved ${res.saved_count || rowIds.length} conversion record(s) to Dataverse.`, 'ok');

        conversionState.showOnlyChanged = false;
        updateStagedEditsUIConversion();
        filterAndRenderConversion();

        if (typeof fetchEntryLogs === 'function') {
          fetchEntryLogs();
        }
      } else {
        throw new Error((res && res.error) || 'Failed to save conversion edits.');
      }
    } catch (err) {
      closeConfirmModal(false);
      showToast('Save error: ' + (err.message || err), 'err');
    } finally {
      if (dom.btnSaveStagedEditsConversion) {
        dom.btnSaveStagedEditsConversion.disabled = false;
        dom.btnSaveStagedEditsConversion.textContent = 'Save to Dataverse';
      }
    }
  }

  async function discardStagedEditsConversion() {
    const rowCount = Object.keys(conversionState.stagedEdits || {}).length;
    if (rowCount > 0) {
      const ok = await showConfirmModal({
        title: 'Discard Unsaved Conversion Edits',
        subtitle: 'Revert Staged Changes',
        message: `Are you sure you want to discard all staged changes for ${rowCount.toLocaleString()} conversion record(s)? This will revert all modified fields back to their loaded values.`,
        confirmText: 'Discard Changes',
        cancelText: 'Cancel',
        isDanger: true
      });
      if (!ok) return;
    }

    if (conversionState.baselineRows && conversionState.baselineRows.length) {
      conversionState.rows = JSON.parse(JSON.stringify(conversionState.baselineRows));
    }
    conversionState.stagedEdits = {};
    conversionState.showOnlyChanged = false;
    updateStagedEditsUIConversion();
    filterAndRenderConversion();

    try {
      await window.pywebview.api.revert_staged_conversion_records(conversionState.rows);
    } catch (e) {
      console.warn('Could not revert engine conversion rows:', e);
    }
    showToast('Discarded all unsaved conversion changes.', 'info');
  }

  async function handleConversionCellEdit(rowIdOrIdx, colKey, newVal, inputEl, tdEl) {
    let targetRow = null;
    let targetIdx = -1;
    for (let i = 0; i < conversionState.rows.length; i++) {
      if (conversionState.rows[i].__id === rowIdOrIdx || i === rowIdOrIdx || String(i) === String(rowIdOrIdx)) {
        targetRow = conversionState.rows[i];
        targetIdx = i;
        break;
      }
    }
    if (!targetRow && typeof rowIdOrIdx === 'number' && conversionState.rows[rowIdOrIdx]) {
      targetRow = conversionState.rows[rowIdOrIdx];
      targetIdx = rowIdOrIdx;
    }

    if (!targetRow) return;

    let trimmed = String(newVal || '').trim();
    targetRow[colKey] = trimmed;

    const rowKey = targetRow.__id || String(targetIdx);
    const baseRow = conversionState.baselineRows && (conversionState.baselineRows.find(r => r.__id && r.__id === targetRow.__id) || conversionState.baselineRows[targetIdx]);
    const baseVal = baseRow ? String(baseRow[colKey] ?? '').trim() : '';

    if (trimmed !== baseVal) {
      if (!conversionState.stagedEdits[rowKey]) conversionState.stagedEdits[rowKey] = {};
      conversionState.stagedEdits[rowKey][colKey] = trimmed;
      if (tdEl) tdEl.classList.add('is-dirty');
      if (inputEl) inputEl.classList.add('modified');
      if (tdEl) {
        tdEl.dataset.prevValue = String(baseVal);
        tdEl.dataset.currValue = String(trimmed);
        tdEl.dataset.fieldName = colKey;
        tdEl.dataset.diffType = 'manual';
        const lotId = targetRow['AREA-INDEX'] || targetRow['cr63f_areaindex'] || targetRow['TITLE NUMBER'] || rowKey;
        tdEl.dataset.lotId = String(lotId);
      }
    } else {
      if (conversionState.stagedEdits[rowKey]) {
        delete conversionState.stagedEdits[rowKey][colKey];
        if (Object.keys(conversionState.stagedEdits[rowKey]).length === 0) {
          delete conversionState.stagedEdits[rowKey];
        }
      }
      if (tdEl) {
        tdEl.classList.remove('is-dirty');
        delete tdEl.dataset.prevValue;
        delete tdEl.dataset.currValue;
        delete tdEl.dataset.fieldName;
        delete tdEl.dataset.diffType;
        delete tdEl.dataset.lotId;
      }
      if (inputEl) inputEl.classList.remove('modified');
    }

    updateStagedEditsUIConversion();

    try {
      await window.pywebview.api.update_conversion_cell(targetRow.__id || targetIdx, colKey, trimmed);
    } catch (e) {
      console.warn('Could not sync conversion edit to engine:', e);
    }
  }

  async function exportConversionExcel() {
    const toExport = conversionState.filteredRows;
    if (!toExport.length) {
      showToast('No conversion data to export.', 'err');
      return;
    }
    const cols = getConversionVisibleColumns();
    await runExport(() => window.pywebview.api.export_records_excel(toExport, 'Conversion_Data_Export.xlsx', cols));
  }

  async function downloadConversionTemplate() {
    try {
      const res = await window.pywebview.api.download_conversion_template();
      if (res.success) {
        showToast(`Template downloaded to ${res.file_path}`, 'ok');
      } else if (!res.cancelled) {
        showToast(res.error || 'Failed to download conversion template.', 'err');
      }
    } catch (err) {
      showToast('Error: ' + err, 'err');
    }
  }

  async function syncConversionToDataverse() {
    if (!conversionState.rows || !conversionState.rows.length) {
      showToast('No conversion data loaded to sync.', 'err');
      return;
    }
    // Edited cells are saved through "Save" (one audit row per changed field);
    // a sync would upload them with only a summary audit row.
    if (Object.keys(conversionState.stagedEdits || {}).length) {
      showToast('Save or discard your edited conversion rows first, then sync.', 'err');
      return;
    }

    const rowCount = conversionState.rows.length.toLocaleString();
    const tableName = state.activeConversionTable || 'cr63f_landconversionreference';
    const confirmed = await showConfirmModal({
      title: 'Sync Conversion Data to Dataverse',
      subtitle: `Table: ${tableName}`,
      message: `Sync ${rowCount} conversion row(s) to Dataverse? Rows that already exist (same AREA-INDEX, or Title Number when there is none) are updated, new rows are created, and rows whose key matches several existing rows are skipped and listed.`,
      confirmText: 'Upload & Sync',
      cancelText: 'Cancel',
      isDanger: false,
      keepOpenOnConfirm: true
    });

    if (!confirmed) return;

    try {
      if (dom.confirmModal) {
        dom.confirmModal.classList.remove('hidden');
      }
      if (dom.confirmModalTitle) dom.confirmModalTitle.textContent = 'Syncing Conversion Data to Dataverse';
      if (dom.confirmModalSubtitle) dom.confirmModalSubtitle.textContent = `Dataverse Reference Sync (${tableName})`;
      if (dom.confirmModalMessage) dom.confirmModalMessage.textContent = `Uploading ${rowCount} conversion record(s)... Please do not close the window.`;

      if (dom.confirmModalProgressWrap) {
        dom.confirmModalProgressWrap.classList.remove('hidden');
      }
      if (dom.confirmModalProgressBar) {
        dom.confirmModalProgressBar.style.width = '0%';
      }
      if (dom.confirmModalProgressStatus) {
        dom.confirmModalProgressStatus.textContent = `Preparing ${rowCount} records...`;
      }
      if (dom.confirmModalProgressPercent) {
        dom.confirmModalProgressPercent.textContent = '0%';
      }
      if (dom.btnConfirmModalOk) {
        dom.btnConfirmModalOk.disabled = true;
        dom.btnConfirmModalOk.textContent = 'Syncing...';
      }
      if (dom.btnConfirmModalCancel) {
        dom.btnConfirmModalCancel.classList.add('hidden');
      }
      if (dom.btnConfirmModalClose) {
        dom.btnConfirmModalClose.classList.add('hidden');
      }
      if (dom.btnSyncConversionToDv) {
        dom.btnSyncConversionToDv.disabled = true;
        dom.btnSyncConversionToDv.textContent = 'Syncing...';
      }

      showToast('Syncing Conversion Data to Dataverse...', 'info');

      const res = await window.pywebview.api.sync_uploaded_conversion_to_dataverse();
      if (res.success || res.synced_count) {
        if (dom.confirmModalProgressBar) dom.confirmModalProgressBar.style.width = '100%';
        if (dom.confirmModalProgressPercent) dom.confirmModalProgressPercent.textContent = '100%';
        if (dom.confirmModalProgressStatus) {
          dom.confirmModalProgressStatus.textContent = res.message || 'Sync finished.';
        }
        showToast(res.success ? res.message : `${res.message} ${res.error || ''}`.trim(), res.success ? (res.skipped_count ? 'info' : 'ok') : 'err');
        if (dom.lblActiveConversionSource) {
          dom.lblActiveConversionSource.textContent = `Dataverse: ${tableName} (${res.synced_count} records)`;
        }
        // Reload the rows so they carry the Dataverse IDs the sync just assigned
        // (later edits are then saved to the right row).
        try {
          const prev = await window.pywebview.api.get_conversion_preview();
          if (prev && Array.isArray(prev.rows)) renderConversionPreview(conversionState.headers, prev.rows);
        } catch (_) {}
        conversionState.baselineRows = JSON.parse(JSON.stringify(conversionState.rows));
        conversionState.stagedEdits = {};
        conversionState.showOnlyChanged = false;
        updateStagedEditsUIConversion();
        filterAndRenderConversion();
        setTimeout(() => {
          if (dom.confirmModal) dom.confirmModal.classList.add('hidden');
          if (dom.btnConfirmModalClose) dom.btnConfirmModalClose.classList.remove('hidden');
        }, 1200);
      } else {
        if (dom.confirmModal) dom.confirmModal.classList.add('hidden');
        if (dom.btnConfirmModalClose) dom.btnConfirmModalClose.classList.remove('hidden');
        showToast(res.error || 'Failed to sync conversion data to Dataverse.', 'err');
      }
    } catch (err) {
      if (dom.confirmModal) dom.confirmModal.classList.add('hidden');
      if (dom.btnConfirmModalClose) dom.btnConfirmModalClose.classList.remove('hidden');
      showToast('Error: ' + err, 'err');
    } finally {
      if (dom.btnSyncConversionToDv) {
        dom.btnSyncConversionToDv.disabled = false;
        dom.btnSyncConversionToDv.textContent = 'Sync to Dataverse';
      }
      if (dom.btnConfirmModalOk) {
        dom.btnConfirmModalOk.disabled = false;
        dom.btnConfirmModalOk.textContent = 'Confirm';
      }
      if (dom.btnConfirmModalCancel) {
        dom.btnConfirmModalCancel.classList.remove('hidden');
      }
      if (dom.btnConfirmModalClose) {
        dom.btnConfirmModalClose.classList.remove('hidden');
      }
    }
  }

  window.deleteConversionRecord = async function (idOrIdx) {
    const ok = await showConfirmModal({
      title: 'Delete Conversion Record',
      subtitle: 'Conversion Reference Deletion',
      message: 'Are you sure you want to delete this record from Conversion Data?',
      confirmText: 'Delete',
      cancelText: 'Cancel',
      isDanger: true
    });
    if (!ok) return;

    try {
      const res = await window.pywebview.api.delete_conversion_record(idOrIdx);
      if (res.success) {
        showToast('Conversion record deleted.', 'ok');
        const targetId = String(idOrIdx);
        let foundIdx = conversionState.rows.findIndex(r => r.__id === targetId);
        if (foundIdx === -1) {
          const num = parseInt(idOrIdx, 10);
          if (!isNaN(num) && num >= 0 && num < conversionState.rows.length) {
            foundIdx = num;
          }
        }
        if (foundIdx !== -1) {
          conversionState.rows.splice(foundIdx, 1);
        }
        state.conversionLoadedCount = conversionState.rows.length;
        if (dom.lblActiveConversionCount) dom.lblActiveConversionCount.textContent = `${state.conversionLoadedCount.toLocaleString()} rows`;
        if (dom.lblTopConversion) dom.lblTopConversion.textContent = state.conversionLoadedCount > 0 ? `${state.conversionLoadedCount.toLocaleString()} rows` : 'Not loaded';
        filterAndRenderConversion();
      } else {
        showToast(res.error || 'Failed to delete conversion record.', 'err');
      }
    } catch (err) {
      showToast('Error: ' + err, 'err');
    }
  };

  function attachConversionHeaderEvents() {
    dom.conversionPreviewThead.querySelectorAll('[data-conv-pin]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const col = btn.dataset.convPin;
        if (conversionState.frozenCols.has(col)) {
          conversionState.frozenCols.delete(col);
        } else {
          conversionState.frozenCols.add(col);
        }
        saveConversionTableConfig();
        renderConversionTable();
      });
    });

    dom.conversionPreviewThead.querySelectorAll('[data-conv-filter-toggle]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const colKey = btn.dataset.convFilterToggle;
        conversionState.columnFiltersOpen[colKey] = !conversionState.columnFiltersOpen[colKey];
        const popup = document.getElementById(`convColFilterBox_${colKey}`);
        if (popup) {
          if (conversionState.columnFiltersOpen[colKey]) {
            popup.classList.remove('hidden');
            const input = popup.querySelector('input');
            if (input) input.focus();
          } else {
            popup.classList.add('hidden');
          }
        }
      });
    });

    let convFilterDebounceTimer = null;
    dom.conversionPreviewThead.querySelectorAll('.th-filter-input').forEach(input => {
      input.addEventListener('input', () => {
        const colKey = input.dataset.convFilterCol;
        conversionState.columnFilters[colKey] = input.value.trim().toLowerCase();
        if (convFilterDebounceTimer) clearTimeout(convFilterDebounceTimer);
        convFilterDebounceTimer = setTimeout(() => {
          filterAndRenderConversion(true); // skipHeadRender keeps the focused input alive
        }, 120);
      });
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
          const colKey = input.dataset.convFilterCol;
          conversionState.columnFiltersOpen[colKey] = false;
          const popup = document.getElementById(`convColFilterBox_${colKey}`);
          if (popup) popup.classList.add('hidden');
        }
        if (e.key === 'Enter') e.preventDefault();
        e.stopPropagation();
      });
    });

    dom.conversionPreviewThead.querySelectorAll('[data-conv-resize]').forEach(resizer => {
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const colKey = resizer.dataset.convResize;
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
            updateTableColumnWidths(dom.conversionPreviewTable, colKey, newW);
          });
        };

        const onUp = () => {
          if (rAF) cancelAnimationFrame(rAF);
          resizer.classList.remove('is-resizing');
          document.body.classList.remove('col-resizing');
          window.removeEventListener('mousemove', onMove);
          window.removeEventListener('mouseup', onUp);
          conversionState.colWidths[colKey] = newW;
          saveConversionTableConfig();
          if (conversionState.frozenCols.has(colKey)) {
            updateStickyOffsets(dom.conversionPreviewTable, getConversionVisibleColumns(), conversionState.frozenCols, conversionState.colWidths, 45);
          }
        };

        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp);
      });
    });

    dom.conversionPreviewThead.querySelectorAll('.th-bottom-resizer').forEach(resizer => {
      attachBottomHeaderResizer(resizer.closest('th'));
    });

    dom.conversionPreviewThead.querySelectorAll('[data-conv-sort]').forEach(wrap => {
      wrap.addEventListener('click', (e) => {
        if (e.target.closest('.th-resizer') || e.target.closest('.th-bottom-resizer') || e.target.closest('.th-pin-btn') || e.target.closest('.th-chevron-btn') || e.target.closest('.th-filter-popup')) return;
        const col = wrap.dataset.convSort;
        if (conversionState.sortCol === col) {
          conversionState.sortDir = conversionState.sortDir === 'asc' ? 'desc' : (conversionState.sortDir === 'desc' ? null : 'asc');
          if (!conversionState.sortDir) conversionState.sortCol = null;
        } else {
          conversionState.sortCol = col;
          conversionState.sortDir = 'asc';
        }
        saveConversionTableConfig();
        filterAndRenderConversion();
      });
    });

    let draggedCol = null;
    const ths = dom.conversionPreviewThead.querySelectorAll('th[draggable="true"]');
    ths.forEach(th => {
      th.addEventListener('dragstart', (e) => {
        draggedCol = th.dataset.convCol;
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
        const targetCol = th.dataset.convCol;
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
        const targetCol = th.dataset.convCol;
        if (!draggedCol || !targetCol || draggedCol === targetCol) return;
        const fromIdx = conversionState.colOrder.indexOf(draggedCol);
        if (fromIdx === -1) return;
        conversionState.colOrder.splice(fromIdx, 1);
        const rect = th.getBoundingClientRect();
        const insertBefore = e.clientX < rect.left + rect.width / 2;
        let toIdx = conversionState.colOrder.indexOf(targetCol);
        if (!insertBefore) toIdx++;
        conversionState.colOrder.splice(toIdx, 0, draggedCol);
        saveConversionTableConfig();
        renderConversionColumnPickerMenu();
        filterAndRenderConversion();
      });
    });
  }

  // ==================== ENTRY LOGS VIEW ====================
  function setupLogsView() {
    if (dom.searchInputLogs) {
      dom.searchInputLogs.addEventListener('input', () => {
        clearTimeout(searchLogsDebounce);
        searchLogsDebounce = setTimeout(filterAndRenderLogs, 120);
      });
    }
    // Do not pass the DOM Event as filterAndRenderLogs' skipHeadRender argument.
    // Doing so updates the count but leaves the previous table rows rendered.
    if (dom.actionFilterLogs) dom.actionFilterLogs.addEventListener('change', () => {
      populateLogsWwFilter(logsState.logs);
      filterAndRenderLogs();
    });
    if (dom.wwFilterLogs) dom.wwFilterLogs.addEventListener('change', () => {
      populateLogsActionFilter(logsState.logs);
      filterAndRenderLogs();
    });
    if (dom.btnClearFiltersLogs) {
      dom.btnClearFiltersLogs.addEventListener('click', () => {
        if (dom.searchInputLogs) dom.searchInputLogs.value = '';
        if (dom.actionFilterLogs) dom.actionFilterLogs.value = 'ALL';
        if (dom.wwFilterLogs) dom.wwFilterLogs.value = 'ALL';
        populateLogsActionFilter(logsState.logs);
        populateLogsWwFilter(logsState.logs);
        logsState.columnFilters = {};
        logsState.columnFiltersOpen = {};
        filterAndRenderLogs();
        showToast('Logs filters reset.', 'ok');
      });
    }
    if (dom.btnRefreshLogs) dom.btnRefreshLogs.addEventListener('click', fetchEntryLogs);
    if (dom.btnExportExcelLogs) dom.btnExportExcelLogs.addEventListener('click', exportLogsExcel);

    setupLogsColumnPicker();
  }

  function setupLogsColumnPicker() {
    if (!dom.btnColumnsPickerLogs || !dom.colPickerMenuLogs) return;

    dom.btnColumnsPickerLogs.addEventListener('click', (e) => {
      e.stopPropagation();
      dom.colPickerMenuLogs.classList.toggle('hidden');
    });

    document.addEventListener('click', (e) => {
      if (!e.target.closest('#view-logs .col-picker-container')) {
        dom.colPickerMenuLogs.classList.add('hidden');
      }
    });
  }

  function renderLogsColumnPickerMenu() {
    if (!dom.colPickerMenuLogs) return;

    const orderedNonFixed = logsState.colOrder
      .map(k => LOGS_COLUMNS.find(c => c.key === k))
      .filter(c => c && !c.fixed);

    let html = '';
    orderedNonFixed.forEach((col, idx) => {
      const isChecked = logsState.visibleCols.has(col.key);
      const isFirst = idx === 0;
      const isLast = idx === orderedNonFixed.length - 1;

      html += `
        <div class="col-picker-item">
          <label class="col-picker-label">
            <input type="checkbox" data-log-col-toggle="${escapeHtml(col.key)}" ${isChecked ? 'checked' : ''} />
            <span>${escapeHtml(col.label)}</span>
          </label>
          <div class="col-reorder-btns">
            <button type="button" class="col-reorder-btn" data-log-move-up="${escapeHtml(col.key)}" ${isFirst ? 'disabled' : ''} title="Move left">◄</button>
            <button type="button" class="col-reorder-btn" data-log-move-down="${escapeHtml(col.key)}" ${isLast ? 'disabled' : ''} title="Move right">►</button>
          </div>
        </div>
      `;
    });

    html += `
      <div class="col-picker-actions">
        <button type="button" class="btn xs ghost" id="btnResetLogsLayout" style="width: 100%; font-size: 11px;">Reset Columns Layout</button>
      </div>
    `;

    dom.colPickerMenuLogs.innerHTML = html;

    dom.colPickerMenuLogs.querySelectorAll('[data-log-col-toggle]').forEach(chk => {
      chk.addEventListener('change', () => {
        const key = chk.dataset.logColToggle;
        if (chk.checked) {
          logsState.visibleCols.add(key);
        } else {
          if (logsState.visibleCols.size > 1) {
            logsState.visibleCols.delete(key);
          } else {
            chk.checked = true;
            showToast('At least one column must remain visible.', 'err');
            return;
          }
        }
        saveLogsTableConfig();
        filterAndRenderLogs();
      });
    });

    dom.colPickerMenuLogs.querySelectorAll('[data-log-move-up]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.logMoveUp;
        moveLogsColumnOrder(key, -1);
      });
    });

    dom.colPickerMenuLogs.querySelectorAll('[data-log-move-down]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const key = btn.dataset.logMoveDown;
        moveLogsColumnOrder(key, 1);
      });
    });

    const btnReset = dom.colPickerMenuLogs.querySelector('#btnResetLogsLayout');
    if (btnReset) {
      btnReset.addEventListener('click', (e) => {
        e.stopPropagation();
        localStorage.removeItem('logs_col_widths');
        localStorage.removeItem('logs_col_order');
        localStorage.removeItem('logs_visible_cols');
        localStorage.removeItem('logs_frozen_cols');
        localStorage.removeItem('logs_sort_col');
        localStorage.removeItem('logs_sort_dir');
        logsState.visibleCols = new Set(LOGS_COLUMNS.map(c => c.key));
        logsState.colOrder = LOGS_COLUMNS.map(c => c.key);
        logsState.colWidths = {};
        logsState.frozenCols = new Set(['__idx', 'log_time']);
        logsState.sortCol = null;
        logsState.sortDir = null;
        saveLogsTableConfig();
        renderLogsColumnPickerMenu();
        filterAndRenderLogs();
        showToast('Logs layout reset to default.', 'ok');
      });
    }
  }

  function moveLogsColumnOrder(colKey, direction) {
    const idx = logsState.colOrder.indexOf(colKey);
    if (idx === -1) return;
    const targetIdx = idx + direction;
    if (targetIdx < 0 || targetIdx >= logsState.colOrder.length) return;

    const temp = logsState.colOrder[idx];
    logsState.colOrder[idx] = logsState.colOrder[targetIdx];
    logsState.colOrder[targetIdx] = temp;

    saveLogsTableConfig();
    renderLogsColumnPickerMenu();
    filterAndRenderLogs();
  }

  async function fetchEntryLogs() {
    try {
      if (dom.btnRefreshLogs) {
        dom.btnRefreshLogs.disabled = true;
        dom.btnRefreshLogs.textContent = 'Refreshing...';
      }
      const res = await window.pywebview.api.fetch_entry_logs();
      if (res.success) {
        logsState.logs = (res.logs || []).map((l, i) => {
          if (!l.__id) {
            l.__id = l.id || l.cr63f_loamonitoringlogid || `log_cached_${i}_${Date.now()}`;
          }
          return l;
        });
        populateLogsActionFilter(logsState.logs);
        populateLogsWwFilter(logsState.logs);
        // Rebuild once more after either facet may have cleared a stale selection.
        populateLogsActionFilter(logsState.logs);
        populateLogsWwFilter(logsState.logs);
        filterAndRenderLogs();
        recalculateLogsKPIs();
      } else {
        showToast(res.error || 'Failed to fetch entry logs.', 'err');
      }
    } catch (err) {
      showToast('Error fetching logs: ' + err, 'err');
    } finally {
      if (dom.btnRefreshLogs) {
        dom.btnRefreshLogs.disabled = false;
        dom.btnRefreshLogs.textContent = 'Refresh Logs';
      }
    }
  }

  function getLogActionFilterKey(log) {
    const action = getLogActionType(log);
    if (action.includes('REVERT')) return 'REVERT';
    if (action.includes('DELETE')) return 'DELETE';
    if (action.includes('REFERENCE')) return 'REFERENCE_UPDATE';
    if (action.includes('FINANCE')) return 'FINANCE_UPDATE';
    if (action.includes('CONVERSION')) return 'CONVERSION_UPDATE';
    if (action.includes('BASELINE')) return 'BASELINE';
    if (action.includes('NEW')) return 'NEW_ENTRY';
    return 'MANUAL_EDIT';
  }

  function populateLogsActionFilter(logs) {
    if (!dom.actionFilterLogs) return;
    const selectedWeek = String(dom.wwFilterLogs ? dom.wwFilterLogs.value : 'ALL').trim().toUpperCase();
    const currentVal = String(dom.actionFilterLogs.value || 'ALL').trim().toUpperCase();
    const labels = {
      NEW_ENTRY: 'New Entry',
      MANUAL_EDIT: 'Manual Edit',
      REFERENCE_UPDATE: 'Reference Update',
      FINANCE_UPDATE: 'Finance Update',
      CONVERSION_UPDATE: 'Conversion Update',
      REVERT: 'Revert',
      DELETE: 'Deleted',
      BASELINE: 'Baseline'
    };
    const counts = new Map();
    (logs || []).forEach(log => {
      const week = String(log.work_week || log['Work Week'] || '').trim().toUpperCase();
      if (selectedWeek !== 'ALL' && week !== selectedWeek) return;
      const key = getLogActionFilterKey(log);
      counts.set(key, (counts.get(key) || 0) + 1);
    });

    let html = '<option value="ALL">All Actions</option>';
    Object.keys(labels).forEach(key => {
      const count = counts.get(key) || 0;
      if (count > 0) html += `<option value="${key}">${labels[key]} (${count.toLocaleString()})</option>`;
    });
    dom.actionFilterLogs.innerHTML = html;
    dom.actionFilterLogs.value = counts.has(currentVal) ? currentVal : 'ALL';
  }

  function populateLogsWwFilter(logs) {
    if (!dom.wwFilterLogs) return;
    const selectedAction = String(dom.actionFilterLogs ? dom.actionFilterLogs.value : 'ALL').trim().toUpperCase();
    const currentVal = dom.wwFilterLogs.value;
    const wwSet = new Set();
    logs.forEach(l => {
      if (selectedAction !== 'ALL' && getLogActionFilterKey(l) !== selectedAction) return;
      const ww = String(l.work_week || l['Work Week'] || '').trim();
      if (ww) wwSet.add(ww);
    });
    const sortedWw = Array.from(wwSet).sort().reverse();
    let optHtml = '<option value="ALL">All Work Weeks</option>';
    sortedWw.forEach(w => {
      optHtml += `<option value="${escapeHtml(w)}" ${w === currentVal ? 'selected' : ''}>${escapeHtml(w)}</option>`;
    });
    dom.wwFilterLogs.innerHTML = optHtml;
    if (currentVal !== 'ALL' && !wwSet.has(currentVal)) dom.wwFilterLogs.value = 'ALL';
  }

  function recalculateLogsKPIs() {
    const all = logsState.logs || [];
    let newCount = 0, manualCount = 0, finCount = 0;
    all.forEach(l => {
      const act = getLogActionType(l);
      if (act.includes('NEW')) newCount++;
      else if (act.includes('REFERENCE')) finCount++;
      else manualCount++;
    });
    if (dom.kpiLogsCount) dom.kpiLogsCount.textContent = all.length.toLocaleString();
    if (dom.kpiLogsNewEntry) dom.kpiLogsNewEntry.textContent = newCount.toLocaleString();
    if (dom.kpiLogsManual) dom.kpiLogsManual.textContent = manualCount.toLocaleString();
    if (dom.kpiLogsFinance) dom.kpiLogsFinance.textContent = finCount.toLocaleString();
  }

  function filterAndRenderLogs(skipHeadRender = false) {
    const q = dom.searchInputLogs ? dom.searchInputLogs.value.toLowerCase().trim() : '';
    // An empty value (an option that no longer exists) means All, never "match nothing".
    const actFilter = (dom.actionFilterLogs ? dom.actionFilterLogs.value.toUpperCase() : 'ALL') || 'ALL';
    const wwFilter = (dom.wwFilterLogs ? dom.wwFilterLogs.value.toUpperCase() : 'ALL') || 'ALL';
    const colFilterKeys = Object.keys(logsState.columnFilters).filter(k => !!logsState.columnFilters[k]);

    logsState.filteredLogs = logsState.logs.filter(l => {
      const act = getLogActionType(l);
      if (actFilter !== 'ALL' && getLogActionFilterKey(l) !== actFilter) return false;
      const ww = String(l.work_week || l['Work Week'] || '').toUpperCase();
      if (wwFilter !== 'ALL' && ww !== wwFilter) return false;

      for (const colKey of colFilterKeys) {
        const query = logsState.columnFilters[colKey];
        const colDef = LOGS_COLUMNS.find(c => c.key === colKey);
        const rawValue = colKey === 'action_type'
          ? getLogActionType(l)
          : l[colKey] ?? (colDef ? l[colDef.alt] : undefined) ?? (colDef ? l[colDef.label] : '') ?? '';
        const val = String(rawValue).toLowerCase();
        if (!val.includes(query)) return false;
      }

      if (!q) return true;
      return Object.values(l).some(v => String(v).toLowerCase().includes(q));
    });

    if (logsState.sortCol && logsState.sortDir) {
      const colKey = logsState.sortCol;
      const mult = logsState.sortDir === 'asc' ? 1 : -1;
      const colDef = LOGS_COLUMNS.find(c => c.key === colKey) || { key: colKey };
      logsState.filteredLogs.sort((a, b) => {
        const valA = colKey === 'action_type'
          ? getLogActionType(a)
          : (a[colKey] ?? (colDef ? a[colDef.alt] : undefined) ?? (colDef ? a[colDef.label] : '') ?? '');
        const valB = colKey === 'action_type'
          ? getLogActionType(b)
          : (b[colKey] ?? (colDef ? b[colDef.alt] : undefined) ?? (colDef ? b[colDef.label] : '') ?? '');
        return compareTableValues(valA, valB, colDef, mult);
      });
    }

    if (dom.lblLogsGridCount) dom.lblLogsGridCount.textContent = `${logsState.filteredLogs.length.toLocaleString()} logs`;

    invalidateGridCache('logsTable');

    if (!skipHeadRender) {
      // Save focused filter input state before re-rendering header
      const activeEl = document.activeElement;
      let focusedFilterCol = null;
      let focusedSelStart = 0;
      let focusedSelEnd = 0;
      if (activeEl && activeEl.classList.contains('th-filter-input') && activeEl.dataset.logFilterCol) {
        focusedFilterCol = activeEl.dataset.logFilterCol;
        focusedSelStart = activeEl.selectionStart || 0;
        focusedSelEnd = activeEl.selectionEnd || 0;
      }

      renderLogsHead();

      // Restore focus and cursor to the filter input after re-render
      if (focusedFilterCol) {
        const restored = dom.logsTableHead && dom.logsTableHead.querySelector(`[data-log-filter-col="${focusedFilterCol}"]`);
        if (restored) {
          restored.focus();
          try { restored.setSelectionRange(focusedSelStart, focusedSelEnd); } catch (_) {}
        }
      }
    }
  }

  function renderLogsHead() {
    if (!dom.logsTableHead) return;
    const visibleCols = getLogsVisibleColumns();
    let thead = '<tr>';

    visibleCols.forEach((col, colIdx) => {
      const width = logsState.colWidths[col.key] || col.width || 130;
      const isSorted = logsState.sortCol === col.key;
      const sortIcon = isSorted ? (logsState.sortDir === 'asc' ? '▲' : '▼') : '';
      const isFrozen = logsState.frozenCols.has(col.key);
      const isLastFrz = isTrailingFrozen(colIdx, visibleCols, logsState.frozenCols);
      const leftOffset = isFrozen ? getFrozenOffset(colIdx, visibleCols, logsState.frozenCols, logsState.colWidths) : 0;
      const frozenClass = `${isFrozen ? 'col-frozen' : ''} ${isLastFrz ? 'col-frozen-last' : ''}`.trim();
      const leftStyle = isFrozen ? `left: ${leftOffset}px;` : '';
      const isFilterActive = !!(logsState.columnFilters[col.key]);
      const isFilterOpen = !!(logsState.columnFiltersOpen[col.key]);

      if (col.key === '__idx' || col.key === '__action') {
        thead += `
          <th class="${col.key === '__action' ? 'col-action' : ''} ${frozenClass}" style="width: ${width}px; min-width: ${width}px; max-width: ${width}px; ${leftStyle}">
            <div class="th-header-box" style="justify-content: center; align-items: center;">
              <span class="th-col-name" style="text-align: center;">${escapeHtml(col.label)}</span>
            </div>
            <div class="th-bottom-resizer" title="Drag to resize the header"></div>
          </th>`;
        return;
      }

      thead += `
        <th draggable="true" data-log-col="${col.key}" class="${frozenClass}" style="width: ${width}px; min-width: ${width}px; max-width: ${width}px; ${leftStyle}">
          <div class="th-header-box">
            <div class="th-title-wrap">
              <button type="button" class="th-pin-btn ${isFrozen ? 'pinned' : ''}" data-log-pin="${col.key}" aria-label="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(col.label)}" title="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(col.label)}">
                <svg viewBox="0 0 24 24" width="11" height="11" fill="currentColor" stroke="none">
                  <path d="M16 12V4h1V2H7v2h1v8l-2 2v2h5.2v6h1.6v-6H18v-2l-2-2z"/>
                </svg>
              </button>
              <button type="button" class="th-sort-wrap" data-log-sort="${col.key}" aria-label="Sort by ${escapeHtml(col.label)}" title="Sort by ${escapeHtml(col.label)} · drag to reorder">
                <span class="th-col-name">${escapeHtml(col.label)}</span>
                <span class="th-sort-icon ${isSorted ? 'active' : ''}" aria-hidden="true">${sortIcon}</span>
              </button>
              <button type="button" class="th-chevron-btn ${isFilterActive ? 'active' : ''}" data-log-filter-toggle="${col.key}" aria-label="Filter ${escapeHtml(col.label)}" title="Filter ${escapeHtml(col.label)}">
                <svg viewBox="0 0 24 24" width="10" height="10" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"></polyline></svg>
              </button>
            </div>
            <div class="th-filter-popup ${isFilterOpen ? '' : 'hidden'}" id="logColFilterBox_${col.key}">
              <input class="th-filter-input" data-log-filter-col="${col.key}" aria-label="Filter ${escapeHtml(col.label)} column" autocomplete="off" value="${escapeHtml(logsState.columnFilters[col.key] || '')}" placeholder="Filter ${escapeHtml(col.label)}…" />
            </div>
          </div>
          <div class="th-resizer" data-log-resize="${col.key}" title="Drag to resize · double-click to reset"></div>
          <div class="th-bottom-resizer" title="Drag to resize the header"></div>
        </th>
      `;
    });

    thead += '</tr>';
    dom.logsTableHead.innerHTML = thead;

    attachLogsHeaderEvents();

    if (!logsState.filteredLogs.length) {
      if (dom.logsTableBody) {
        dom.logsTableBody.innerHTML = '';
        const visibleCols = getLogsVisibleColumns();
        const colSpan = Math.max(1, visibleCols.length);
        const isInitialEmpty = !logsState.logs.length;
        const title = isInitialEmpty ? 'No Entry Logs found' : 'No matching logs found';
        const desc = isInitialEmpty
          ? 'Changes made to Land Master DB or New Entries will appear here.'
          : 'Try adjusting your search terms, Action Type, or Work Week filters.';

        const emptyTr = document.createElement('tr');
        emptyTr.className = 'empty-table-row';
        emptyTr.innerHTML = `
          <td colspan="${colSpan}" style="text-align: center; padding: 48px 16px; background: transparent; border: none;">
            <div style="display: inline-flex; flex-direction: column; align-items: center; justify-content: center; gap: 6px;">
              <svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" style="color: var(--ink3); opacity: 0.7;">
                <circle cx="12" cy="12" r="10"></circle>
                <polyline points="12 6 12 12 16 14"></polyline>
              </svg>
              <p style="font-size: 13.5px; font-weight: 600; color: var(--ink); margin: 4px 0 2px;">${title}</p>
              <p style="font-size: 12px; color: var(--ink3); margin: 0;">${desc}</p>
            </div>
          </td>
        `;
        dom.logsTableBody.appendChild(emptyTr);
      }
      if (dom.logsEmptyState) dom.logsEmptyState.classList.add('hidden');
      return;
    }
    if (dom.logsEmptyState) dom.logsEmptyState.classList.add('hidden');
    if (dom.logsTableBody) dom.logsTableBody.innerHTML = '';
    logsState.renderedLogsCount = 0;
    renderLogsChunk(0, CHUNK_SIZE);
  }

  const CANONICAL_FIELD_ALIASES = {
    'cr63f_totalcontractprice': ['cr63f_totalcontractprice', 'total contract price', 'totalcontractprice', 'tcp', 'total contract price (finance)'],
    'cr63f_totalpaymentdeductibletotcp': ['cr63f_totalpaymentdeductibletotcp', 'payment deductible', 'paymentdeductible', 'total payment deductible to tcp', 'total payment deductible', 'payment deductible to tcp', 'totalpaymentdeductibletotcp'],
    'cr63f_balance': ['cr63f_balance', 'balance', 'outstanding balance'],
    'cr63f_paid': ['cr63f_paid', '% paid', 'paid', 'percent paid', '%paid'],
    'cr63f_conversionstatus': ['cr63f_conversionstatus', 'conversion status', 'conversionstatus', 'dar status', 'status'],
    'cr63f_acquisitionsublevel': ['cr63f_acquisitionsublevel', 'acquisition sub-level', 'acquisition sublevel', 'acquisitionsublevel'],
    'cr63f_rccprice': ['cr63f_rccprice', 'rcc price', 'rccprice'],
    'cr63f_rccterms': ['cr63f_rccterms', 'rcc terms', 'rccterms'],
    'cr63f_secureddate': ['cr63f_secureddate', 'secured date', 'secureddate'],
    'cr63f_securedendorsedate': ['cr63f_securedendorsedate', 'secured endorse date', 'securedendorsedate'],
    'cr63f_endorseddate': ['cr63f_endorseddate', 'endorsed date', 'endorseddate'],
    'work_week': ['work_week', 'work week', 'workweek', 'cr63f_workweek'],
    'cr63f_control1': ['cr63f_control1', 'cr63f_control', 'control', 'control number', 'control no', 'control no.'],
    'cr63f_titleno': ['cr63f_titleno', 'title no', 'title no.', 'title number', 'titleno'],
    'cr63f_areaindex': ['cr63f_areaindex', 'area-index', 'area index', 'areaindex'],
    'cr63f_indexno': ['cr63f_indexno', 'index no', 'index no.', 'indexno'],
    'cr63f_type': ['cr63f_type', 'type'],
    'cr63f_lotareaha': ['cr63f_lotareaha', 'lot area (ha)', 'lot area', 'lotareaha'],
    'cr63f_possrightsstatus': ['cr63f_possrightsstatus', 'poss rights status', 'possrights status', 'possrightsstatus'],
    'cr63f_possrightsarea': ['cr63f_possrightsarea', 'poss rights area (ha)', 'poss rights area', 'possrightsarea'],
    'cr63f_possrightscat': ['cr63f_possrightscat', 'possrights cat', 'poss rights cat', 'possrightscat'],
    'cr63f_catarea': ['cr63f_catarea', 'cat area (ha)', 'cat area', 'catarea'],
    'new_contracttype': ['new_contracttype', 'cr63f_contracttype', 'contract type', 'contracttype'],
    'new_mode': ['new_mode', 'cr63f_mode', 'mode'],
    'cr63f_registeredowner': ['cr63f_registeredowner', 'registered owner', 'registeredowner', 'owner'],
    'cr63f_authorizedrepresentative': ['cr63f_authorizedrepresentative', 'authorized representative', 'authorized rep', 'authorizedrepresentative'],
    'cr63f_entity': ['cr63f_entity', 'entity'],
    'cr63f_province': ['cr63f_province', 'province'],
    'cr63f_municipality': ['cr63f_municipality', 'municipality'],
    'cr63f_dbadjustedbarangay': ['cr63f_dbadjustedbarangay', 'db adjusted barangay', 'adjusted barangay', 'dbadjustedbarangay'],
  };

  function getFieldNameAttributeKey(fieldName) {
    if (!fieldName) return '';
    const clean = s => String(s || '').toLowerCase().replace(/^cr63f_|^new_/, '').replace(/[^a-z0-9]/g, '');
    const norm = clean(fieldName);
    if (!norm) return fieldName;
    for (const [key, aliases] of Object.entries(CANONICAL_FIELD_ALIASES)) {
      if (clean(key) === norm) return key;
      for (const a of aliases) {
        if (clean(a) === norm) return key;
      }
    }
    for (let i = 0; i < DEFAULT_COLUMNS.length; i++) {
      const col = DEFAULT_COLUMNS[i];
      const kNorm = clean(col.key);
      const lNorm = clean(col.label);
      const aNorm = clean(col.alt);
      if (norm === kNorm || norm === lNorm || norm === aNorm) {
        return col.key;
      }
    }
    return fieldName;
  }

  function revertValuesEqual(left, right, fieldName) {
    const fLow = String(fieldName || '').toLowerCase();
    const isNumeric = ['price', 'deductible', 'balance', 'tcp', 'amount', 'paid', '%'].some(k => fLow.includes(k));
    if (isNumeric) {
      const leftNum = parseAmount(left);
      const rightNum = parseAmount(right);
      if (leftNum !== null && rightNum !== null) {
        return Math.abs(leftNum - rightNum) <= 0.005;
      }
    }
    const norm = val => String(val || '').replace(/[\s,₱$%]/g, '').trim().toLowerCase();
    return norm(left) === norm(right);
  }

  function isLogRevertible(log) {
    if (!log) return false;
    if (['FINANCE_UPDATE', 'CONVERSION_UPDATE'].includes(getLogActionType(log))) return false;
    if (getLogActionType(log).includes('DELETE')) return false;
    const targetField = String(log.field_name || log['Field Changed'] || log.cr63f_fieldchanged || '').trim();
    if (!targetField || targetField === 'Initial Entry' || targetField === 'createdon' || targetField === 'cr63f_createdon') {
      return false;
    }

    const fLow = targetField.toLowerCase();
    const nonRevertibleCalculatedFields = new Set(['balance', '% paid', 'paid', 'percent paid']);
    if (nonRevertibleCalculatedFields.has(fLow)) {
      return false;
    }

    const targetCtrl = String(log.cr63f_control1 || log['Control Number'] || log.Control || log['Control No'] || log.control_no || '').trim().toUpperCase();
    const targetLot = String(log.area_index || log['AREA-INDEX'] || log['Area-Index'] || log['Area Index'] || log.cr63f_areaindex || '').trim().toUpperCase();
    const targetTitle = String(log.title_no || log['Title No'] || log['TITLE NO'] || log['Title No.'] || log.cr63f_titleno || '').trim().toUpperCase();
    const targetIdx = String(log.index_no || log['Index No'] || log['INDEX NO'] || log['Corridor Index'] || log.cr63f_indexno || '').trim().toUpperCase();
    const targetId = getAuditMasterId(log);

    if (!targetCtrl && !targetLot && !targetTitle && !targetIdx && !targetId) {
      return false;
    }

    const logId = String(log.__id || log.id || log.cr63f_loamonitoringlogid || '');

    // 1. Must be the FIRST (most recent) audit log in logsState.logs for this same record and field
    const allLogs = logsState.logs || [];
    for (let i = 0; i < allLogs.length; i++) {
      const other = allLogs[i];
      const otherField = String(other.field_name || other['Field Changed'] || other.cr63f_fieldchanged || '').trim();
      const oFieldLow = otherField.toLowerCase();
      if (oFieldLow !== fLow && getFieldNameAttributeKey(oFieldLow) !== getFieldNameAttributeKey(fLow)) continue;

      const oCtrl = String(other.cr63f_control1 || other['Control Number'] || other.Control || other['Control No'] || other.control_no || '').trim().toUpperCase();
      const oLot = String(other.area_index || other['AREA-INDEX'] || other['Area-Index'] || other['Area Index'] || other.cr63f_areaindex || '').trim().toUpperCase();
      const oTitle = String(other.title_no || other['Title No'] || other['TITLE NO'] || other['Title No.'] || other.cr63f_titleno || '').trim().toUpperCase();
      const oIdx = String(other.index_no || other['Index No'] || other['INDEX NO'] || other['Corridor Index'] || other.cr63f_indexno || '').trim().toUpperCase();
      if (['FINANCE_UPDATE', 'CONVERSION_UPDATE'].includes(getLogActionType(other))) continue;
      const oId = getAuditMasterId(other);

      let isSameRecord = false;
      if (targetId && oId) isSameRecord = targetId === oId;
      else if (targetCtrl && oCtrl) isSameRecord = targetCtrl === oCtrl;
      else if (targetLot && oLot && targetLot === oLot) isSameRecord = true;
      else if (targetTitle && oTitle && targetTitle === oTitle) isSameRecord = true;
      else if (targetIdx && oIdx && targetIdx === oIdx) isSameRecord = true;

      if (isSameRecord) {
        const oIdStr = String(other.__id || other.id || other.cr63f_loamonitoringlogid || '');
        if (oIdStr === logId) {
          // This IS the most recent log for this record + field!
          break;
        } else {
          // A newer log exists for this record + field!
          return false;
        }
      }
    }

    // 2. Check if current master record value matches expected new_value
    if (state.records && state.records.length) {
      const master = state.records.find(r => {
        if (!r) return false;
        const rCtrl = String(r.cr63f_control1 || r['Control Number'] || r.Control || r['Control No'] || r.control_no || '').trim().toUpperCase();
        const rArea = String(r.cr63f_areaindex || r['Area Index'] || r['Area-Index'] || r['AREA-INDEX'] || r.area_index || '').trim().toUpperCase();
        const rTitle = String(r.cr63f_titleno || r['Title No'] || r['TITLE NO'] || r['Title No.'] || r.title_no || '').trim().toUpperCase();
        const rIdx = String(r.cr63f_indexno || r['Index No'] || r['INDEX NO'] || r['Corridor Index'] || r.index_no || '').trim().toUpperCase();
        const rId = String(r.__id || r.id || r.cr63f_loamonitoringid || '').trim();

        if (targetId && rId) return targetId === rId;
        if (targetCtrl && rCtrl) return targetCtrl === rCtrl;
        if (targetLot && rArea && targetLot === rArea) return true;
        if (targetTitle && rTitle && targetTitle === rTitle) return true;
        if (targetIdx && rIdx && targetIdx === rIdx) return true;
        return false;
      });

      if (master) {
        const expectedNew = log.new_value ?? log.cr63f_newvalue ?? log['New Value'] ?? '';
        const currentVal = master[targetField] ?? master[getFieldNameAttributeKey(targetField)] ?? '';
        if (!revertValuesEqual(currentVal, expectedNew, targetField)) {
          return false;
        }
      }
    }

    return true;
  }

  function formatLogDateTime(raw) {
    if (!raw) return '-';
    const s = String(raw).trim();
    if (!s || s === '-') return '-';
    if (s.includes('T') || s.endsWith('Z')) {
      const d = new Date(s);
      if (!isNaN(d.getTime())) {
        const pad = n => String(n).padStart(2, '0');
        const mm = pad(d.getMonth() + 1);
        const dd = pad(d.getDate());
        const yyyy = d.getFullYear();
        let h = d.getHours();
        const ampm = h >= 12 ? 'PM' : 'AM';
        h = h % 12 || 12;
        return `${mm}/${dd}/${yyyy} ${pad(h)}:${pad(d.getMinutes())} ${ampm}`;
      }
    }
    return s;
  }

function renderLogsChunk(startIdx, count) {
    if (!dom.logsTableBody) return;
    const end = Math.min(startIdx + count, logsState.filteredLogs.length);
    const frag = document.createDocumentFragment();
    const visibleCols = getLogsVisibleColumns();

    for (let idx = startIdx; idx < end; idx++) {
      const r = logsState.filteredLogs[idx];
      const tr = document.createElement('tr');
      tr.dataset.rowIdx = idx;

      visibleCols.forEach((col, colIdx) => {
        const td = document.createElement('td');
        const width = logsState.colWidths[col.key] || col.width || 130;
        const isFrozen = logsState.frozenCols.has(col.key);
        const isLastFrz = isTrailingFrozen(colIdx, visibleCols, logsState.frozenCols);
        const leftOffset = isFrozen ? getFrozenOffset(colIdx, visibleCols, logsState.frozenCols, logsState.colWidths) : 0;

        td.style.width = `${width}px`;
        td.style.minWidth = `${width}px`;
        td.style.maxWidth = `${width}px`;
        td.dataset.colKey = col.key;

        if (isFrozen) {
          td.classList.add('col-frozen');
          td.style.left = `${leftOffset}px`;
        }
        if (isLastFrz) {
          td.classList.add('col-frozen-last');
        }

        if (col.key === '__idx') {
          td.style.fontFamily = 'var(--mono)';
          td.style.color = 'var(--ink3)';
          td.style.textAlign = 'center';
          td.textContent = String(idx + 1);
        } else if (col.key === 'work_week') {
          const ww = r.work_week || r['Work Week'] || '';
          td.style.textAlign = 'center';
          td.innerHTML = ww ? `<span class="ww-pill">${escapeHtml(ww)}</span>` : `<span style="color: var(--ink3); font-size: 11px;">-</span>`;
        } else if (col.key === 'action_type') {
          const act = getLogActionType(r, 'LOG');
          let cls = 'action-edit';
          let label = 'Manual Edit';
          if (act.includes('NEW')) {
            cls = 'action-new';
            label = 'New Entry';
          } else if (act.includes('REFERENCE') || act.includes('FINANCE') || act.includes('CONVERSION')) {
            cls = 'action-finance';
            label = 'Reference Update';
          } else if (act.includes('REVERT')) {
            cls = 'action-revert';
            label = 'Revert';
          } else if (act.includes('DELETE')) {
            cls = 'action-delete';
            label = 'Deleted';
          } else if (act.includes('MANUAL')) {
            cls = 'action-edit';
            label = 'Manual Edit';
          } else {
            label = act.replace(/_/g, ' ');
          }
          td.style.textAlign = 'center';
          td.innerHTML = `<span class="action-badge ${cls}">${escapeHtml(label)}</span>`;
        } else if (col.key === 'field_name') {
          const fn = r.field_name || r['Field Changed'] || r.cr63f_fieldchanged || '-';
          td.innerHTML = `<span class="log-diff-field" title="${escapeHtml(fn)}">${escapeHtml(fn)}</span>`;
        } else if (col.key === 'old_value') {
          const fieldName = r.field_name || r['Field Changed'] || r.cr63f_fieldchanged || '';
          const oldV = formatLogValue(fieldName, r.old_value ?? r['Previous Value'] ?? r.cr63f_previousvalue ?? '-');
          td.innerHTML = `<span class="log-old-val" title="${escapeHtml(String(oldV))}">${escapeHtml(String(oldV || '-'))}</span>`;
        } else if (col.key === 'new_value') {
          const fieldName = r.field_name || r['Field Changed'] || r.cr63f_fieldchanged || '';
          const newV = formatLogValue(fieldName, r.new_value ?? r['New Value'] ?? r.cr63f_newvalue ?? '-');
          td.innerHTML = `<span class="log-new-val" title="${escapeHtml(String(newV))}">${escapeHtml(String(newV || '-'))}</span>`;
        } else if (col.key === 'cr63f_control1') {
          const ctrlVal = r.cr63f_control1 ?? r.Control ?? r['Control Number'] ?? r.cr63f_control ?? r['Control No'] ?? r['Control No.'] ?? r.control_no ?? '';
          td.textContent = String(ctrlVal || '-');
          td.style.fontFamily = 'var(--mono)';
          td.style.fontSize = '12px';
        } else if (col.key === 'log_time') {
          const lt = r.log_time ?? r['Log Time'] ?? r.cr63f_logtime ?? r.createdon ?? r.timestamp ?? '';
          td.textContent = formatLogDateTime(lt);
          td.style.fontFamily = 'var(--mono)';
          td.style.fontSize = '12px';
        } else if (col.key === 'area_index') {
          const ai = r.area_index ?? r['Area-Index'] ?? r['Area Index'] ?? r.cr63f_areaindex ?? r['AREA-INDEX'] ?? '';
          td.textContent = String(ai || '-');
          td.style.fontFamily = 'var(--mono)';
          td.style.fontSize = '12px';
        } else if (col.key === 'title_no') {
          const tn = r.title_no ?? r['Title No'] ?? r['Title Number'] ?? r.cr63f_titleno ?? r['TITLE NO'] ?? '';
          td.textContent = String(tn || '-');
        } else {
          const val = r[col.key] ?? (col.alt ? r[col.alt] : undefined) ?? r[col.label] ?? '';
          if (col.isPrice) {
            td.textContent = (val !== '' && val !== null && val !== undefined) ? formatPrice(val) : '';
            td.style.textAlign = 'right';
          } else if (col.isPercent) {
            td.textContent = (val !== '' && val !== null && val !== undefined) ? formatPercent(val) : '';
            td.style.textAlign = 'right';
          } else {
            td.textContent = col.key === 'details'
              ? String(val).replace(/\s*\[(?:MASTER|SOURCE):[^\]]+\]/g, '').trim()
              : String(val);
          }
        }
        tr.appendChild(td);
      });
      frag.appendChild(tr);
    }
    dom.logsTableBody.appendChild(frag);
    logsState.renderedLogsCount = end;
  }

  function attachLogsHeaderEvents() {
    if (!dom.logsTableHead) return;

    // 1. Column Resizers
    dom.logsTableHead.querySelectorAll('[data-log-resize]').forEach(resizer => {
      resizer.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const colKey = resizer.dataset.logResize;
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
            updateTableColumnWidths(dom.logsTable, colKey, currentWidth);
          });
        };

        const onMouseUp = () => {
          if (rAF) cancelAnimationFrame(rAF);
          window.removeEventListener('mousemove', onMouseMove);
          window.removeEventListener('mouseup', onMouseUp);
          resizer.classList.remove('is-resizing');
          document.body.classList.remove('col-resizing');
          logsState.colWidths[colKey] = currentWidth;
          saveLogsTableConfig();
          if (logsState.frozenCols.has(colKey)) {
            updateStickyOffsets(dom.logsTable, getLogsVisibleColumns(), logsState.frozenCols, logsState.colWidths, 0);
          }
        };

        window.addEventListener('mousemove', onMouseMove);
        window.addEventListener('mouseup', onMouseUp);
      });

      resizer.addEventListener('dblclick', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const colKey = resizer.dataset.logResize;
        const colDef = LOGS_COLUMNS.find(c => c.key === colKey);
        if (colDef) {
          logsState.colWidths[colKey] = colDef.width;
          saveLogsTableConfig();
          updateTableColumnWidths(dom.logsTable, colKey, colDef.width);
          if (logsState.frozenCols.has(colKey)) {
            updateStickyOffsets(dom.logsTable, getLogsVisibleColumns(), logsState.frozenCols, logsState.colWidths, 0);
          }
        }
      });
    });

    // 2. Bottom Resizer for Header Height
    dom.logsTableHead.querySelectorAll('.th-bottom-resizer').forEach(resizer => {
      attachBottomHeaderResizer(resizer.closest('th'));
    });

    // 3. Freeze / Pin Buttons
    dom.logsTableHead.querySelectorAll('[data-log-pin]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const colKey = btn.dataset.logPin;
        if (logsState.frozenCols.has(colKey)) {
          logsState.frozenCols.delete(colKey);
        } else {
          logsState.frozenCols.add(colKey);
        }
        saveLogsTableConfig();
        renderLogsHead();
        filterAndRenderLogs();
      });
    });

    // 4. Sort Wrap
    dom.logsTableHead.querySelectorAll('[data-log-sort]').forEach(wrap => {
      wrap.addEventListener('click', (e) => {
        if (e.target.closest('.th-chevron-btn') || e.target.closest('.th-resizer') || e.target.closest('.th-filter-popup') || e.target.closest('.th-pin-btn')) return;
        const colKey = wrap.dataset.logSort;
        if (logsState.sortCol === colKey) {
          if (logsState.sortDir === 'asc') logsState.sortDir = 'desc';
          else if (logsState.sortDir === 'desc') { logsState.sortCol = null; logsState.sortDir = null; }
          else logsState.sortDir = 'asc';
        } else {
          logsState.sortCol = colKey;
          logsState.sortDir = 'asc';
        }
        saveLogsTableConfig();
        renderLogsHead();
        filterAndRenderLogs();
      });
    });

    // 5. Filter Popups
    dom.logsTableHead.querySelectorAll('[data-log-filter-toggle]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const colKey = btn.dataset.logFilterToggle;
        logsState.columnFiltersOpen[colKey] = !logsState.columnFiltersOpen[colKey];
        const popup = document.getElementById(`logColFilterBox_${colKey}`);
        if (popup) {
          if (logsState.columnFiltersOpen[colKey]) {
            popup.classList.remove('hidden');
            const input = popup.querySelector('input');
            if (input) input.focus();
          } else {
            popup.classList.add('hidden');
          }
        }
      });
    });

    let logsFilterDebounceTimer = null;
    dom.logsTableHead.querySelectorAll('.th-filter-input').forEach(input => {
      input.addEventListener('input', () => {
        const colKey = input.dataset.logFilterCol;
        logsState.columnFilters[colKey] = input.value.trim().toLowerCase();
        // filterAndRenderLogs restores this input's focus and cursor after rebuilding the table.
        if (logsFilterDebounceTimer) clearTimeout(logsFilterDebounceTimer);
        logsFilterDebounceTimer = setTimeout(() => {
          filterAndRenderLogs();
        }, 120);
      });
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
          const colKey = input.dataset.logFilterCol;
          logsState.columnFiltersOpen[colKey] = false;
          const popup = document.getElementById(`logColFilterBox_${colKey}`);
          if (popup) popup.classList.add('hidden');
        }
        // Prevent Escape or Enter from bubbling and triggering other handlers
        if (e.key === 'Enter') e.preventDefault();
        e.stopPropagation();
      });
    });

    // 6. Drag & Drop Reordering
    let draggedLogCol = null;
    const ths = dom.logsTableHead.querySelectorAll('th[draggable="true"]');
    ths.forEach(th => {
      th.addEventListener('dragstart', (e) => {
        draggedLogCol = th.dataset.logCol;
        th.classList.add('col-dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', draggedLogCol);
      });

      th.addEventListener('dragend', () => {
        th.classList.remove('col-dragging');
        ths.forEach(t => t.classList.remove('drag-over-left', 'drag-over-right'));
      });

      th.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        const targetCol = th.dataset.logCol;
        if (!targetCol || targetCol === draggedLogCol) return;
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
        const targetCol = th.dataset.logCol;
        if (!draggedLogCol || !targetCol || draggedLogCol === targetCol) return;
        const fromIdx = logsState.colOrder.indexOf(draggedLogCol);
        if (fromIdx === -1) return;
        logsState.colOrder.splice(fromIdx, 1);
        const rect = th.getBoundingClientRect();
        const insertBefore = e.clientX < rect.left + rect.width / 2;
        let toIdx = logsState.colOrder.indexOf(targetCol);
        if (!insertBefore) toIdx++;
        logsState.colOrder.splice(toIdx, 0, draggedLogCol);
        saveLogsTableConfig();
        renderLogsColumnPickerMenu();
        filterAndRenderLogs();
      });
    });
  }

  async function exportLogsExcel() {
    const toExport = logsState.filteredLogs;
    if (!toExport.length) {
      showToast('No logs to export.', 'err');
      return;
    }
    const cols = LOGS_COLUMNS.filter(c => !c.key.startsWith('__')).map(c => c.label);
    await runExport(() => window.pywebview.api.export_logs_excel(toExport, 'Entry_Logs_Export.xlsx', cols));
  }

  function revertMasterRecordInState(logItem, upRec) {
    if (!logItem) return;
    const targetCtrl = String(logItem.cr63f_control1 || logItem['Control Number'] || logItem.Control || logItem['Control No'] || logItem.control_no || '').trim().toUpperCase();
    const targetLot = String(logItem.area_index || logItem['AREA-INDEX'] || logItem['Area-Index'] || logItem['Area Index'] || logItem.cr63f_areaindex || '').trim().toUpperCase();
    const targetTitle = String(logItem.title_no || logItem['Title No'] || logItem['TITLE NO'] || logItem['Title No.'] || logItem.cr63f_titleno || '').trim().toUpperCase();
    const targetIdx = String(logItem.index_no || logItem['Index No'] || logItem['INDEX NO'] || logItem['Corridor Index'] || logItem.cr63f_indexno || '').trim().toUpperCase();
    const targetField = String(logItem.field_name || logItem['Field Changed'] || logItem.cr63f_fieldchanged || '').trim();
    const oldRaw = logItem.old_value !== undefined ? logItem.old_value : (logItem['Previous Value'] !== undefined ? logItem['Previous Value'] : logItem.cr63f_previousvalue);
    const oldVal = (oldRaw === '-' || oldRaw === 'None' || oldRaw === 'null' || oldRaw === 'Previous' || oldRaw === 'Created') ? '' : String(oldRaw ?? '').trim();

    const cleanNum = oldVal.replace(/[₱$,%]/g, '').trim();

    const updateSingleRecord = (r) => {
      if (!r) return false;
      const rCtrl = String(r.cr63f_control1 || r['Control Number'] || r.Control || r['Control No'] || r.control_no || '').trim().toUpperCase();
      const rArea = String(r.cr63f_areaindex || r['Area Index'] || r['Area-Index'] || r['AREA-INDEX'] || r.area_index || '').trim().toUpperCase();
      const rTitle = String(r.cr63f_titleno || r['Title No'] || r['TITLE NO'] || r['Title No.'] || r.title_no || '').trim().toUpperCase();
      const rIdx = String(r.cr63f_indexno || r['Index No'] || r['INDEX NO'] || r['Corridor Index'] || r.index_no || '').trim().toUpperCase();

      let isMatch = false;
      if (upRec && upRec.__id && r.__id && upRec.__id === r.__id) isMatch = true;
      else if (targetCtrl && rCtrl && targetCtrl === rCtrl) isMatch = true;
      else if (targetLot && rArea && targetLot === rArea) isMatch = true;
      else if (targetTitle && rTitle && (targetTitle === rTitle || rTitle.includes(targetTitle) || targetTitle.includes(rTitle))) isMatch = true;
      else if (targetIdx && rIdx && targetIdx === rIdx) isMatch = true;

      if (!isMatch) return false;

      if (upRec && typeof upRec === 'object') {
        Object.assign(r, upRec);
        if (r.__id && state.stagedEdits?.[r.__id]) delete state.stagedEdits[r.__id];
        return true;
      }

      const fLow = targetField.toLowerCase();
      if (fLow.includes('price') || fLow.includes('contract price') || fLow === 'tcp') {
        const num = parseFloat(cleanNum);
        const formatted = !isNaN(num) ? '₱' + num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : oldVal;
        r['cr63f_totalcontractprice'] = cleanNum;
        r['Total Contract Price'] = formatted;
        r['TOTAL CONTRACT PRICE'] = formatted;
      } else if (fLow.includes('deductible')) {
        const num = parseFloat(cleanNum);
        const formatted = !isNaN(num) ? '₱' + num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : oldVal;
        r['cr63f_totalpaymentdeductibletotcp'] = cleanNum;
        r['Payment Deductible'] = formatted;
        r['Total Payment Deductible to TCP'] = formatted;
        r['TOTAL PAYMENT DEDUCTIBLE TO TCP'] = formatted;
      } else if (fLow.includes('dar') || fLow.includes('conversion') || fLow === 'status') {
        r['cr63f_conversionstatus'] = oldVal;
        r['Conversion Status'] = oldVal;
        r['DAR STATUS'] = oldVal;
        r['DAR Status'] = oldVal;
        r['Status'] = oldVal;
      } else if (fLow === 'mode') {
        r['new_mode'] = oldVal;
        r['Mode'] = oldVal;
        r['MODE'] = oldVal;
        r['cr63f_mode'] = oldVal;
      } else if (fLow.includes('contract type')) {
        r['new_contracttype'] = oldVal;
        r['Contract Type'] = oldVal;
        r['cr63f_contracttype'] = oldVal;
      } else if (fLow.includes('endorsed date') || fLow === 'endorsed date') {
        r['cr63f_endorseddate'] = oldVal;
        r['Endorsed Date'] = oldVal;
        const ww = getWorkWeek(oldVal);
        r['work_week'] = ww;
        r['Work Week'] = ww;
      } else if (fLow.includes('lot area')) {
        r['cr63f_lotareaha'] = cleanNum;
        r['Lot Area (Ha)'] = cleanNum;
        r['Lot Area (ha)'] = cleanNum;
      } else {
        r[targetField] = oldVal;
        const colDef = DEFAULT_COLUMNS.find(c => c.label === targetField || c.alt === targetField || c.key === targetField);
        if (colDef) {
          r[colDef.key] = oldVal;
          if (colDef.alt) r[colDef.alt] = oldVal;
        }
      }

      // Financial balance & % paid recalculation if TCP or Deductible was changed
      const shouldRecalc = fLow.includes('price') || fLow.includes('contract price') || fLow === 'tcp' || fLow.includes('deductible');
      if (shouldRecalc) {
        // Same rule as the server: a blank TCP cannot be calculated (blank result); % Paid keeps one decimal.
        const tcpRaw = r['cr63f_totalcontractprice'] ?? r['Total Contract Price'] ?? '';
        const tcpBlank = String(tcpRaw).trim() === '';
        const tcp = parseAmount(tcpRaw);
        const ded = parseAmount(r['cr63f_totalpaymentdeductibletotcp'] ?? r['Payment Deductible'] ?? r['Total Payment Deductible to TCP']);
        const bal = tcpBlank ? '' : Math.max(0, tcp - ded);
        let pct = '';
        if (!tcpBlank) pct = tcp > 0 ? (ded / tcp * 100).toFixed(1) + '%' : (ded > 0 ? '100.0%' : '0.0%');
        r['cr63f_balance'] = bal;
        r['Balance'] = tcpBlank ? '' : '₱' + bal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        r['BALANCE'] = r['Balance'];
        r['cr63f_paid'] = pct;
        r['% Paid'] = pct;
        r['Paid'] = pct;
      }

      // Clear any staged edit for this record
      if (r.__id && state.stagedEdits && state.stagedEdits[r.__id]) {
        delete state.stagedEdits[r.__id];
      }

      return true;
    };

    if (Array.isArray(state.records)) {
      state.records.forEach(updateSingleRecord);
    }
    if (Array.isArray(state.baselineRecords)) {
      state.baselineRecords.forEach(updateSingleRecord);
    }
    if (Array.isArray(state.filteredRecords)) {
      state.filteredRecords.forEach(updateSingleRecord);
    }

    recalculateKPIs();
    filterAndRenderRecords();
  }

  window.revertLogRecord = async function (idOrIdx) {
    if (Object.keys(state.stagedEdits || {}).length) {
      showToast('Save or discard staged Land Master edits before reverting a saved change.', 'err');
      return;
    }
    const logItem = logsState.logs.find(l => String(l.__id || l.id || l.cr63f_loamonitoringlogid || '') === String(idOrIdx));
    if (!logItem) { showToast('Log record not found. Refresh and try again.', 'err'); return; }
    const field = logItem.field_name || logItem.cr63f_fieldchanged || logItem['Field Changed'] || 'field';
    const rawPrev = logItem.old_value ?? logItem.cr63f_previousvalue ?? logItem['Previous Value'] ?? '';
    const previous = (rawPrev === '' || rawPrev === '-' || rawPrev === null || rawPrev === 'None' || rawPrev === 'null')
      ? '(blank)'
      : formatLogValue(field, rawPrev);
    const confirmed = await showConfirmModal({
      title: 'Revert Land Master Change',
      subtitle: 'Dataverse update operation',
      message: `Restore ${field} to ${previous}? The field will be updated in Land Master DB and Dataverse, and a new REVERT audit log will record the change.`,
      confirmText: 'Revert',
      cancelText: 'Cancel',
      isDanger: true
    });
    if (!confirmed) return;
    try {
      const res = await window.pywebview.api.revert_log_record(String(idOrIdx));
      if (!res || !res.success) throw new Error((res && res.error) || 'Revert failed.');
      showToast(res.message || 'Field reverted successfully in Dataverse and Land Master DB.', 'ok');
      if (res.updated_record) revertMasterRecordInState(logItem, res.updated_record);
      await Promise.all([fetchRecords(), fetchEntryLogs()]);
    } catch (err) {
      showToast('Revert failed: ' + (err.message || err), 'err');
    }
  };

  // ==================== OVERVIEW (DASHBOARD) VIEW ====================
  // Analyst-first layout: "This week" (deltas, changed accounts, needs attention) above "Portfolio".
  let selectedDashboardMode = null;
  let dashAuditSearchQuery = '';
  let dashSourceFilter = 'ALL';
  let dashAuditCurrentPage = 1;
  let dashAuditPageSize = 25;
  let dashChangedAccounts = [];
  let dashExpandedAccounts = new Set();
  let dashLastSyncedAt = null;

  const DASH_SOURCE_LABELS = { reference: 'Reference update', manual: 'Manual edit', new: 'New entry', revert: 'Revert' };
  const DASH_SOURCE_SHORT = { ALL: 'All', reference: 'Reference', manual: 'Manual', new: 'New', revert: 'Reverted' };

  function dashEl(id) { return document.getElementById(id); }

  // The newest week is "follow the latest": when new activity starts a newer
  // week, the Overview moves to it. An older week the user picked stays put.
  function currentDashWeek() {
    const sel = dom.dashWorkWeekSelect;
    if (!sel || !sel.value || sel.selectedIndex <= 0) return 'ALL';
    return sel.value;
  }

  function reloadDashboard() {
    return loadDashboardAnalytics(currentDashWeek(), selectedDashboardMode);
  }

  // ₱347.1M on cards, the exact figure goes in the title attribute.
  function formatPesoCompact(n) {
    const v = Number(n) || 0;
    const a = Math.abs(v);
    const sign = v < 0 ? '−' : '';
    if (a >= 1e9) return `${sign}₱${(a / 1e9).toFixed(a >= 1e10 ? 1 : 2)}B`;
    if (a >= 1e6) return `${sign}₱${(a / 1e6).toFixed(1)}M`;
    if (a >= 1e3) return `${sign}₱${(a / 1e3).toFixed(0)}K`;
    return `${sign}₱${a.toFixed(0)}`;
  }
  function formatPesoExact(n) {
    const v = Number(n) || 0;
    return `${v < 0 ? '−' : ''}₱${Math.abs(v).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  }
  function titleCaseStatus(s) {
    const t = String(s || '').trim();
    if (!t) return '(blank)';
    return t.toLowerCase().replace(/\b(dar|nia|ww|row|carp|tcp)\b/g, m => m.toUpperCase()).replace(/^./, c => c.toUpperCase());
  }
  function formatShortDateTime(raw) {
    if (!raw) return '—';
    const d = new Date(raw);
    if (isNaN(d.getTime())) return String(raw);
    return d.toLocaleString('en-US', { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
  }
  function formatWeekRange(start, end) {
    if (!start || !end) return '';
    const o = { month: 'short', day: 'numeric' };
    const a = new Date(start + 'T00:00:00'), b = new Date(end + 'T00:00:00');
    if (isNaN(a) || isNaN(b)) return '';
    return `${a.toLocaleDateString('en-US', o)} – ${b.toLocaleDateString('en-US', o)}`;
  }
  function shortWeek(label) {
    const m = String(label || '').match(/WW-\d+/i);
    return m ? m[0].toUpperCase() : String(label || '');
  }

  // goodWhen: 'up' (more is good), 'down' (less is good)
  function deltaChip(value, { money = false, goodWhen = 'up', suffix = '' } = {}) {
    const v = Number(value) || 0;
    if (Math.abs(v) < (money ? 0.5 : 0.0001)) return '<span class="ov-chip mut">No change</span>';
    const up = v > 0;
    const good = goodWhen === 'up' ? up : !up;
    const txt = money ? formatPesoCompact(Math.abs(v)) : Math.abs(v).toLocaleString();
    const exact = money ? ` title="${escapeHtml((up ? '+' : '−') + formatPesoExact(Math.abs(v)).replace('−', ''))}"` : '';
    return `<span class="ov-chip ${good ? 'pos' : 'neg'}"${exact}>${up ? '▲' : '▼'} ${escapeHtml(txt)}${escapeHtml(suffix)}</span>`;
  }

  function updateDashSyncedLabel() {
    const el = dashEl('dashSyncedLabel');
    if (!el) return;
    if (!dashLastSyncedAt) { el.textContent = ''; return; }
    el.textContent = 'Synced ' + dashLastSyncedAt.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
    el.title = dashLastSyncedAt.toLocaleString();
  }
  window.__markDataSynced = function () { dashLastSyncedAt = new Date(); updateDashSyncedLabel(); };

  function setupDashboardView() {
    if (dom.dashActivitySearchInput) {
      dom.dashActivitySearchInput.addEventListener('input', () => {
        dashAuditSearchQuery = dom.dashActivitySearchInput.value.toLowerCase().trim();
        dashAuditCurrentPage = 1;
        renderDashboardAuditTable();
      });
    }
    if (dom.dashWorkWeekSelect) {
      dom.dashWorkWeekSelect.addEventListener('change', () => { dashExpandedAccounts.clear(); reloadDashboard(); });
    }
    const stepWeek = (dir) => {
      const sel = dom.dashWorkWeekSelect;
      if (!sel || !sel.options.length) return;
      // Options are newest first, so "previous week" moves down the list.
      const next = sel.selectedIndex + (dir < 0 ? 1 : -1);
      if (next < 0 || next >= sel.options.length) return;
      sel.selectedIndex = next;
      dashExpandedAccounts.clear();
      reloadDashboard();
    };
    const prev = dashEl('btnDashPrevWeek');
    const nextBtn = dashEl('btnDashNextWeek');
    if (prev) prev.addEventListener('click', () => stepWeek(-1));
    if (nextBtn) nextBtn.addEventListener('click', () => stepWeek(1));

    if (dom.btnRefreshDashboard) {
      dom.btnRefreshDashboard.addEventListener('click', async () => {
        const btn = dom.btnRefreshDashboard;
        const label = btn.querySelector('span');
        btn.disabled = true;
        if (label) label.textContent = 'Syncing…';
        try {
          if (state.connected) await fetchRecords();
          await fetchEntryLogs();
          await reloadDashboard();
        } finally {
          btn.disabled = false;
          if (label) label.textContent = 'Sync';
        }
      });
    }
    const btnReport = dashEl('btnDashWeeklyReport');
    if (btnReport) {
      btnReport.addEventListener('click', async () => {
        try {
          btnReport.disabled = true;
          const res = await window.pywebview.api.export_weekly_report(currentDashWeek(), selectedDashboardMode);
          if (res && res.success) showToast('Weekly report saved to ' + res.path, 'ok');
          else if (!(res && res.cancelled)) showToast((res && res.error) || 'Could not export the weekly report.', 'err');
        } catch (err) {
          showToast('Could not export the weekly report: ' + err, 'err');
        } finally {
          btnReport.disabled = false;
        }
      });
    }
    if (dom.btnDashClearMode) {
      dom.btnDashClearMode.addEventListener('click', () => {
        selectedDashboardMode = null;
        reloadDashboard();
      });
    }
    updateDashSyncedLabel();
  }

  async function loadDashboardAnalytics(workWeek = 'ALL', mode = null) {
    try {
      if (dom.viewDashboard) dom.viewDashboard.setAttribute('aria-busy', 'true');
      selectedDashboardMode = mode || null;
      const res = await window.pywebview.api.get_dashboard_analytics(workWeek, selectedDashboardMode);
      if (res.success) renderDashboard(res);
      else showToast(res.error || 'Failed to load the overview.', 'err');
    } catch (err) {
      showToast('Error loading the overview: ' + err, 'err');
    } finally {
      if (dom.viewDashboard) dom.viewDashboard.removeAttribute('aria-busy');
    }
  }

  function renderDashboard(data) {
    if (!data) return;
    window.__lastDashboardData = data;
    const ws = data.week_summary || {};
    const conv = data.conversion_metrics || {};
    const fb = data.financial_breakdown || {};
    const weeks = data.available_work_weeks || [];
    const curWeek = data.selected_work_week || '';
    const hasWeek = curWeek && curWeek.toUpperCase() !== 'ALL';

    // ---- Week picker ----
    if (dom.dashWorkWeekSelect) {
      if (weeks.length) {
        dom.dashWorkWeekSelect.innerHTML = weeks.map(w =>
          `<option value="${escapeHtml(w)}" ${w === curWeek ? 'selected' : ''}>${escapeHtml(shortWeek(w))}</option>`).join('');
        dom.dashWorkWeekSelect.disabled = false;
      } else {
        dom.dashWorkWeekSelect.innerHTML = '<option value="ALL">No weeks yet</option>';
        dom.dashWorkWeekSelect.disabled = true;
      }
      const idx = dom.dashWorkWeekSelect.selectedIndex;
      const prev = dashEl('btnDashPrevWeek'), next = dashEl('btnDashNextWeek');
      if (prev) prev.disabled = !weeks.length || idx >= weeks.length - 1;
      if (next) next.disabled = !weeks.length || idx <= 0;
    }
    const range = formatWeekRange(ws.week_start, ws.week_end);
    const prevW = shortWeek(data.comparison_previous_week || ws.previous_week || '');
    const rangeEl = dashEl('dashWeekRange');
    if (rangeEl) rangeEl.textContent = [range, prevW ? `vs ${prevW}` : ''].filter(Boolean).join(' · ');
    const cmpEl = dashEl('dashCompareLabel');
    if (cmpEl) cmpEl.textContent = hasWeek ? (prevW ? `${shortWeek(curWeek)} compared with ${prevW}` : `${shortWeek(curWeek)} · first recorded week`) : 'No logged changes yet';

    // ---- Mode filter banner ----
    const hasModeFilter = Boolean(data.selected_mode && String(data.selected_mode).toUpperCase() !== 'ALL');
    if (dom.dashModeFilterBanner) {
      dom.dashModeFilterBanner.classList.toggle('hidden', !hasModeFilter);
      const t = dashEl('dashModeFilterText');
      if (t && hasModeFilter) t.innerHTML = `Showing <b>${escapeHtml(String(data.selected_mode).toUpperCase())}</b> accounts only — ${(data.total_records || 0).toLocaleString()} of ${(data.all_records_count || 0).toLocaleString()}`;
    }

    // ---- This week KPIs ----
    if (dom.dashKpiWeeklyChanges) dom.dashKpiWeeklyChanges.textContent = (ws.accounts_changed || 0).toLocaleString();
    if (dom.dashKpiWeeklyChangesSubtitle) {
      const edits = ws.field_edits || 0, adds = ws.new_entries || 0;
      dom.dashKpiWeeklyChangesSubtitle.innerHTML = `${edits.toLocaleString()} field edit${edits === 1 ? '' : 's'}` + (adds ? ` · <span class="ov-chip acc">${adds} new entr${adds === 1 ? 'y' : 'ies'}</span>` : '');
    }
    // Money tiles (decision 2026-10-14): Payments made count every account, including
    // money paid before turnover to partner. Balance to pay counts only what the company
    // still owes: direct sale + secured lease. Turned-over balances are the partner's and
    // rescinded balances are not owed. The hover text lists every group.
    const gt = fb.group_totals || {};
    const gAmt = (g, k) => Number((gt[g] || {})[k]) || 0;
    const MONEY_GROUPS = [
      ['DIRECT_SALE', 'Secured sale'],
      ['LEASE_DIRECT', 'Secured lease'],
      ['STP_SALE', 'Turned over – sale'],
      ['LEASE_STP', 'Turned over – lease'],
      ['NON_SP', 'Non-SP'],
      ['RESCINDED', 'Rescinded'],
      ['OTHER', 'No sale or lease mode']
    ];
    const moneyTip = (k, counted, footer) => MONEY_GROUPS
      .filter(([g]) => gt[g] && (gt[g].count || gAmt(g, k)))
      .map(([g, label]) => `${label}: ${formatPesoExact(gAmt(g, k))}${counted.includes(g) ? '' : ' (not counted)'}`)
      .concat([`Total: ${formatPesoExact(counted.reduce((sum, g) => sum + gAmt(g, k), 0))}`])
      .concat(footer ? ['', footer] : []).join('\n');
    const paymentsMade = Number(fb.payments_made_total) || 0;
    const balanceToPay = Number(fb.balance_to_pay) || 0;
    const turnedPaid = Number(fb.turned_over_paid) || 0;
    const turnedBal = Number(fb.turned_over_balance) || 0;
    if (dom.dashKpiTotalDeductible) {
      dom.dashKpiTotalDeductible.textContent = formatPesoCompact(paymentsMade);
      dom.dashKpiTotalDeductible.title = moneyTip('paid', MONEY_GROUPS.map(([g]) => g),
        'Includes payments made before turnover to partner.');
    }
    if (dom.dashKpiDeductibleDelta) dom.dashKpiDeductibleDelta.innerHTML =
      `${deltaChip(ws.deductible_delta, { money: true, goodWhen: 'up' })}<span>this week${turnedPaid ? ` · incl. ${escapeHtml(formatPesoCompact(turnedPaid))} on accounts turned over to partner` : ''}</span>`;
    if (dom.dashKpiTotalBalance) {
      dom.dashKpiTotalBalance.textContent = formatPesoCompact(balanceToPay);
      dom.dashKpiTotalBalance.title = moneyTip('balance', ['DIRECT_SALE', 'LEASE_DIRECT'],
        'Grey rows are not counted: only secured sale and secured lease balances are yours to pay.');
    }
    const balDelta = dashEl('dashKpiBalanceDelta');
    if (balDelta) balDelta.innerHTML =
      `${deltaChip(ws.own_balance_delta, { money: true, goodWhen: 'down' })}<span>secured sale + secured lease${turnedBal ? ` · excl. ${escapeHtml(formatPesoCompact(turnedBal))} now with partner` : ''}</span>`;
    const sec = dashEl('dashKpiSecured');
    if (sec) sec.textContent = (conv.secured_count || 0).toLocaleString();
    const secD = dashEl('dashKpiSecuredDelta');
    if (secD) secD.innerHTML = `${deltaChip(ws.secured_delta, { goodWhen: 'up' })}<span>${escapeHtml((Number(conv.secured_percent_num) || 0).toFixed(1))}% of ${(conv.total_lots || 0).toLocaleString()}</span>`;

    // ---- Changed accounts ----
    dashChangedAccounts = Array.isArray(data.changed_accounts) ? data.changed_accounts : [];
    renderDashSourceFilter(ws.source_counts || {});
    renderDashboardAuditTable();

    // ---- Needs attention ----
    renderDashAttention(data.needs_attention || []);

    // ---- Portfolio ----
    const pl = dashEl('dashPortfolioLabel');
    if (pl) pl.textContent = `${hasModeFilter ? String(data.selected_mode).toUpperCase() + ' · ' : 'All '}${(data.total_records || 0).toLocaleString()} accounts${hasWeek ? ' · as of ' + shortWeek(curWeek) : ''}`;

    // Payment progress (decisions 2026-10-13/14), kept to one line per group with the
    // amounts in hover text. Headline %, bar and status rows: DIRECT SALE only
    // (SECURED SALE, SS - FROM ROW). Secured lease, turned over to partner and
    // rescinded are listed below as counts.
    const peso = (v) => formatPesoExact(v);
    const directCount = fb.direct_sale_count || 0;
    const directPaidPct = Number(fb.direct_sale_paid_percent_num) || 0;
    const directTip = [
      '# Secured sale',
      `Lots: ${directCount.toLocaleString()}`,
      `Contract price: ${peso(gAmt('DIRECT_SALE', 'tcp'))}`,
      `Paid: ${peso(gAmt('DIRECT_SALE', 'paid'))}`,
      `% paid: ${directPaidPct.toFixed(1)}%`,
      `Balance: ${peso(gAmt('DIRECT_SALE', 'balance'))}`,
      '',
      'Modes: SECURED SALE, SS - FROM LAND, SS - FROM ROW.'
    ].join('\n');
    if (dom.dashFinancePaidRatio) {
      dom.dashFinancePaidRatio.textContent = directCount ? `${directPaidPct.toFixed(1)}%` : '—';
      dom.dashFinancePaidRatio.title = directTip;
    }
    const finSub = dashEl('dashFinanceSub');
    if (finSub) {
      finSub.textContent = directCount
        ? `secured sale paid · ${formatPesoCompact(fb.direct_sale_tcp || 0)} contract · ${formatPesoCompact(gAmt('DIRECT_SALE', 'balance'))} balance`
        : 'no secured sale accounts in this view';
      finSub.title = directTip;
    }
    const finSegs = [
      { name: 'FULLY PAID', count: fb.fully_paid_count || 0, pct: fb.fully_paid_percent_num || 0, cls: 'ov-c-pos',
        tip: 'Payments cover the full contract price.' },
      { name: 'PARTIALLY PAID', count: fb.partially_paid_count || 0, pct: fb.partially_paid_percent_num || 0, cls: 'ov-c-mid',
        tip: 'Some payment made, with a balance still to pay.' },
      { name: 'NO PAYMENT YET', count: fb.zero_paid_count || 0, pct: fb.zero_paid_percent_num || 0, cls: 'ov-c-none',
        tip: 'Nothing paid against the contract price yet.' }
    ];
    renderStack(dashEl('dashFinanceStack'), directCount ? finSegs : [], 'lots');

    const leaseCount = fb.lease_direct_count || 0;
    const leasePct = Number(fb.lease_direct_paid_percent_num) || 0;
    const turnedSale = fb.stp_sale_count || 0;
    const turnedLease = fb.lease_stp_count || 0;
    // Secured lease gets its own % paid (paid ÷ its contract price), shown as a
    // second small progress block; it stays out of the direct sale % above because a
    // lease contract price is total rent over the whole term (decision 2026-10-14).
    const leaseTcp = gAmt('LEASE_DIRECT', 'tcp');
    const leasePaid = gAmt('LEASE_DIRECT', 'paid');
    const leaseBal = gAmt('LEASE_DIRECT', 'balance');
    const leaseTip = [
      '# Secured lease',
      `Lots: ${leaseCount.toLocaleString()}`,
      `Contract price (total rent): ${peso(leaseTcp)}`,
      `Paid: ${peso(leasePaid)}`,
      `% paid: ${leasePct.toFixed(1)}%`,
      `Balance (rent remaining): ${peso(leaseBal)}`,
      '',
      'Reaches 100% only at the end of the lease term.',
      'Modes: SECURED LEASE, SL - FROM LAND.'
    ].join('\n');
    const leasePaidW = leaseTcp > 0 ? Math.min(100, (leasePaid / leaseTcp) * 100) : 0;
    const leaseBlock = leaseCount
      ? `<div class="ov-sub-prog ov-tip" title="${escapeHtml(leaseTip)}">`
        + `<div class="ov-sub-hd"><b>${leasePct.toFixed(1)}%</b><span>secured lease paid · ${escapeHtml(formatPesoCompact(leaseTcp))} contract · ${escapeHtml(formatPesoCompact(leaseBal))} balance · ${leaseCount.toLocaleString()} lots</span></div>`
        + `<div class="ov-stack ov-stack-sm" role="img" aria-label="Secured lease ${leasePct.toFixed(1)}% paid"><i class="ov-c-pos" style="width:${leasePaidW}%"></i></div>`
        + `</div>`
      : '';
    const others = [
      { name: 'TURNED OVER TO PARTNER', count: turnedSale + turnedLease, right: `${(turnedSale + turnedLease).toLocaleString()} lots`,
        tip: [
          '# Turned over to partner',
          `Sale: ${turnedSale.toLocaleString()} lots`,
          `Lease: ${turnedLease.toLocaleString()} lots`,
          `Paid before turnover: ${peso(gAmt('STP_SALE', 'paid') + gAmt('LEASE_STP', 'paid'))}`,
          `Partner's balance: ${peso(gAmt('STP_SALE', 'balance') + gAmt('LEASE_STP', 'balance'))} (not counted)`,
          '',
          'The partner\'s balance is not counted in Balance to pay.'
        ].join('\n') },
      { name: 'NON-SP', count: fb.non_sp_count || 0, right: `${(fb.non_sp_count || 0).toLocaleString()} lots`,
        tip: ['# Non-SP', `Paid: ${peso(gAmt('NON_SP', 'paid'))}`, `Balance: ${peso(gAmt('NON_SP', 'balance'))} (not counted)`,
          '', 'Balance is not counted in Balance to pay.', 'Modes: SECURED SALE NON SP, SECURED LEASE NON SP.'].join('\n') },
      { name: 'RESCINDED', count: fb.rescinded_count || 0, right: `${(fb.rescinded_count || 0).toLocaleString()} lots`,
        tip: ['# Rescinded', `Paid: ${peso(gAmt('RESCINDED', 'paid'))}`, `Balance: ${peso(gAmt('RESCINDED', 'balance'))} (not counted)`,
          '', 'The remaining balance is not owed.', 'Modes: RESCINDED, SS - RESCINDED, SL - RESCINDED.'].join('\n') },
      { name: 'MODE NOT SALE OR LEASE', count: fb.other_count || 0, right: `${(fb.other_count || 0).toLocaleString()} lots`,
        tip: 'Mode is blank, or not a sale or lease.' }
    ].filter(o => o.count > 0);
    if (dom.dashFinanceBreakdownContainer) {
      const statusRows = directCount ? finSegs.map(s =>
        `<div class="ov-li ov-tip" title="${escapeHtml(`# ${s.name.charAt(0)}${s.name.slice(1).toLowerCase()}\nLots: ${s.count.toLocaleString()} of ${directCount.toLocaleString()}\nShare: ${s.pct.toFixed(1)}%\n\n${s.tip}`)}"><span class="ov-sw ${s.cls}"></span><span class="ov-nm">${escapeHtml(s.name)}</span><span class="ov-ct">${s.count.toLocaleString()} lots · ${s.pct.toFixed(1)}%</span></div>`).join('') : '';
      const otherRows = others.length
        ? `<div class="ov-legend-sep">Not in % paid · hover for details</div>` + others.map(o =>
            `<div class="ov-li ov-li-muted ov-tip" title="${escapeHtml(o.tip)}"><span class="ov-sw ov-c-hatch"></span><span class="ov-nm">${escapeHtml(o.name)}</span><span class="ov-ct">${escapeHtml(o.right)}</span></div>`).join('')
        : '';
      dom.dashFinanceBreakdownContainer.innerHTML = (statusRows + leaseBlock + otherRows) || '<p class="ov-empty">No accounts in this view.</p>';
    }
    if (dom.dashFinanceFooterStats) {
      const paidTip = '# Payments made\n' + moneyTip('paid', MONEY_GROUPS.map(([g]) => g), 'Includes payments made before turnover to partner.');
      const balTip = '# Balance to pay\n' + moneyTip('balance', ['DIRECT_SALE', 'LEASE_DIRECT'], 'Grey rows are not counted: only secured sale and secured lease balances are yours to pay.');
      dom.dashFinanceFooterStats.innerHTML =
        `<span class="ov-tip" title="${escapeHtml(paidTip)}">Payments made <b>${escapeHtml(formatPesoCompact(paymentsMade))}</b></span> · `
        + `<span class="ov-tip" title="${escapeHtml(balTip)}">Balance to pay <b>${escapeHtml(formatPesoCompact(balanceToPay))}</b></span>`;
    }

    if (dom.dashConvSecuredRatio) dom.dashConvSecuredRatio.textContent = `${(Number(conv.secured_percent_num) || 0).toFixed(1)}%`;
    const convSub = dashEl('dashConvSub');
    if (convSub) convSub.textContent = `secured · ${(Number(conv.total_applied_area) || 0).toLocaleString('en-US', { maximumFractionDigits: 1 })} ha applied`;
    const totalLots = conv.total_lots || 0;
    const pctOf = (n) => totalLots ? (n / totalLots) * 100 : 0;
    // Each status gets its own colour (decision 2026-10-14): shades of blue for secured,
    // amber/orange for in progress, red for pulled-out/withdrawn, grey for not yet endorsed.
    // The bar is split by status in the same order and colours as the list.
    const breakdown = (conv.breakdown || []).slice().sort((x, y) => (convRank(y.status) - convRank(x.status)) || ((y.count || 0) - (x.count || 0)));
    const CONV_COLORS = {
      'DAR FINALITY SECURED': '#1e3a8a',
      'DAR CO SECURED': '#3b82f6',
      'DAR APPLIED': '#d97706',
      'DAR FOR APPLICATION': '#fbbf24',
      'DAR PULLED-OUT/ WITHDRAWN': '#dc2626',
      'DAR PULLED-OUT/WITHDRAWN': '#dc2626',
      'NOT YET ENDORSED': '#cbd5e1'
    };
    const RANK_SHADES = { 3: ['#1d4ed8', '#60a5fa', '#93c5fd'], 2: ['#ea580c', '#f59e0b', '#fcd34d', '#fdba74'], 1: ['#94a3b8', '#cbd5e1', '#e2e8f0'] };
    const usedPerRank = { 1: 0, 2: 0, 3: 0 };
    const convColor = (status) => {
      const key = String(status || '').trim().toUpperCase().replace(/\s+/g, ' ');
      if (CONV_COLORS[key]) return CONV_COLORS[key];
      if (/PULLED|WITHDRAWN|CANCEL|REVOKED|DENIED|REJECTED/.test(key)) return '#dc2626';
      const r = convRank(status);
      const shades = RANK_SHADES[r] || RANK_SHADES[1];
      return shades[(usedPerRank[r]++) % shades.length];
    };
    const convSegs = breakdown.map(b => ({
      name: String(b.status || '').toUpperCase(), count: b.count || 0, pct: pctOf(b.count || 0), color: convColor(b.status)
    }));
    renderStack(dashEl('dashConvStack'), convSegs, 'lots');
    if (dom.dashConvBarsContainer) {
      dom.dashConvBarsContainer.innerHTML = convSegs.length
        ? convSegs.map(c => `<div class="ov-li" title="${escapeHtml(`${c.name}: ${c.count.toLocaleString()} lots (${c.pct.toFixed(1)}%)`)}"><span class="ov-sw" style="background:${c.color}"></span><span class="ov-nm">${escapeHtml(c.name)}</span><span class="ov-ct">${c.count.toLocaleString()}</span></div>`).join('')
        : '<p class="ov-empty">No conversion statuses yet.</p>';
    }

    // ---- Accounts by mode (click to filter the whole page) ----
    if (dom.dashModeBarsContainer) {
      const modes = (data.modes_distribution || []).slice();
      // SECURED LEASE sits directly after SECURED SALE (requested 2026-10-14).
      const modeName = (m) => String(m.mode).toUpperCase();
      const leaseIdx = modes.findIndex(m => modeName(m) === 'SECURED LEASE');
      if (leaseIdx > -1 && modes.some(m => modeName(m) === 'SECURED SALE')) {
        const [lease] = modes.splice(leaseIdx, 1);
        modes.splice(modes.findIndex(m => modeName(m) === 'SECURED SALE') + 1, 0, lease);
      }
      if (dom.dashModeTotalLots) dom.dashModeTotalLots.textContent = hasModeFilter ? 'Click again to clear' : 'Click to filter';
      // Every mode is listed (the card scrolls); it used to stop at the 8 largest (2026-10-14).
      dom.dashModeBarsContainer.innerHTML = modes.length ? modes.map(m => {
        const active = Boolean(data.selected_mode && data.selected_mode.toUpperCase() === String(m.mode).toUpperCase());
        const muted = /TO BE IDENTIFIED|UNSET|NONE/.test(String(m.mode).toUpperCase());
        return `<button type="button" class="ov-mode ${active ? 'active' : ''} ${hasModeFilter && !active ? 'dim' : ''}" data-mode-name="${escapeHtml(m.mode)}" aria-pressed="${active}">
            <span class="ov-mode-t"><b>${escapeHtml(String(m.mode || '').toUpperCase())}</b><span>${(m.count || 0).toLocaleString()} · ${escapeHtml(m.percent || '')}</span></span>
            <span class="ov-bar"><i class="${muted ? 'muted' : ''}" style="width:${Math.min(100, Number(m.percent_val) || 0)}%"></i></span>
          </button>`;
      }).join('') : '<p class="ov-empty">No mode data yet.</p>';
      dom.dashModeBarsContainer.querySelectorAll('.ov-mode').forEach(btn => {
        btn.addEventListener('click', () => {
          const m = btn.dataset.modeName;
          selectedDashboardMode = (data.selected_mode && data.selected_mode.toUpperCase() === m.toUpperCase()) ? null : m;
          dashExpandedAccounts.clear();
          reloadDashboard();
        });
      });
    }
    updateDashSyncedLabel();
  }

  function convRank(status) {
    // Mirrors loa_engine.conversion_rank: whole words plus a negation check, so
    // NOT SECURED / UNSECURED / DISAPPROVED / NOT YET CONVERTED are not "secured".
    const s = String(status || '').trim().toUpperCase().replace(/\s+/g, ' ');
    // Not started, or ended before conversion: rank 1 (never "in progress").
    const notStarted = ['', '-', 'UNKNOWN', 'NONE', 'N/A', 'UNSET', 'NOT YET ENDORSED'];
    const ended = ['CANCELED TITLE', 'CANCELLED TITLE', 'WITHDRAWN', 'PULLED-OUT', 'PULLED OUT'];
    if (notStarted.includes(s) || s.startsWith('NOT YET ENDORSED') || ended.includes(s)) return 1;
    const secured = /\b(SECURED|EXEMPT|EXEMPTED|APPROVED|COMPLETED|CONVERTED)\b/.test(s);
    const negated = /\b(NOT|NO|YET|PENDING|DENIED|REJECTED|REVOKED|CANCELLED|CANCELED|INSPECTION)\b/.test(s);
    return secured && !negated ? 3 : 2;
  }

  function renderStack(el, segs, unit) {
    if (!el) return;
    el.innerHTML = segs.filter(s => s.pct > 0).map(s =>
      `<i class="${s.cls || ''}" style="width:${Math.min(100, s.pct)}%${s.color ? `;background:${s.color}` : ''}" title="${escapeHtml(s.name)}: ${s.count.toLocaleString()} ${unit} (${s.pct.toFixed(1)}%)"></i>`).join('');
    el.setAttribute('aria-label', segs.map(s => `${s.name} ${s.pct.toFixed(1)}%`).join(', '));
  }
  function renderLegend(segs, unit) {
    return segs.map(s => `<div class="ov-li"><span class="ov-sw ${s.cls}"></span><span class="ov-nm">${escapeHtml(s.name)}</span><span class="ov-ct">${s.count.toLocaleString()} ${unit} · ${s.pct.toFixed(1)}%</span></div>`).join('');
  }

  function renderDashSourceFilter(counts) {
    const el = dashEl('dashSourceFilter');
    if (!el) return;
    const total = dashChangedAccounts.length;
    const keys = ['ALL', 'reference', 'manual', 'new'];
    if (counts.revert) keys.push('revert');
    if (dashSourceFilter !== 'ALL' && !keys.includes(dashSourceFilter)) dashSourceFilter = 'ALL';
    el.innerHTML = keys.map(k => {
      const n = k === 'ALL' ? total : (counts[k] || 0);
      return `<button type="button" class="${dashSourceFilter === k ? 'on' : ''}" data-src="${k}" aria-pressed="${dashSourceFilter === k}">${DASH_SOURCE_SHORT[k]} <span>${n}</span></button>`;
    }).join('');
    el.querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
      dashSourceFilter = b.dataset.src;
      dashAuditCurrentPage = 1;
      renderDashSourceFilter(counts);
      renderDashboardAuditTable();
    }));
  }

  function renderDashAttention(items) {
    const list = dashEl('dashAttentionList');
    const count = dashEl('dashAttentionCount');
    if (count) {
      count.textContent = items.length.toLocaleString();
      count.className = 'ov-chip ' + (items.length ? 'neg' : 'pos');
    }
    if (!list) return;
    if (!items.length) {
      list.innerHTML = '<div class="ov-att-empty"><b>All clear</b><span>No payment drops, overpayments or status regressions this week.</span></div>';
      return;
    }
    const icon = {
      payment_down: '<polyline points="23 18 13.5 8.5 8.5 13.5 1 6"></polyline><polyline points="17 18 23 18 23 12"></polyline>',
      over_paid: '<circle cx="12" cy="12" r="9"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line>',
      status_back: '<polyline points="1 4 1 10 7 10"></polyline><path d="M3.5 15a9 9 0 1 0 2.1-9.4L1 10"></path>'
    };
    const shown = items;
    list.innerHTML = shown.map(n => `
      <div class="ov-att-item">
        <span class="ov-att-ico ${n.severity === 'high' ? 'neg' : 'warn'}"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icon[n.kind] || icon.over_paid}</svg></span>
        <div class="ov-att-body">
          <b>${escapeHtml(n.title)}</b>
          <span>${escapeHtml(n.account || '—')} · ${escapeHtml(n.kind === 'status_back' ? n.detail.split(' → ').map(titleCaseStatus).join(' → ') : n.detail)}</span>
          ${n.log_time ? `<span class="ov-muted">${escapeHtml(DASH_SOURCE_LABELS[n.source] || 'Change')}, ${escapeHtml(formatShortDateTime(n.log_time))}</span>` : ''}
          <button type="button" class="ov-link" data-acct="${escapeHtml(n.account || '')}">Review in Land Master</button>
        </div>
      </div>`).join('') + (items.length > shown.length ? `<div class="ov-att-more">+ ${items.length - shown.length} more — see the weekly report</div>` : '');
    list.querySelectorAll('.ov-link[data-acct]').forEach(b => b.addEventListener('click', () => window.jumpToDashboardRecord(b.dataset.acct)));
  }

  function renderDashboardAuditTable() {
    if (!dom.dashAuditTableBody) return;
    const q = dashAuditSearchQuery;
    const filtered = dashChangedAccounts.filter(a => {
      if (dashSourceFilter !== 'ALL' && !(a.sources || []).includes(dashSourceFilter)) return false;
      if (q) {
        const hay = [a.account, a.area_index, a.title_no, a.control_number, a.mode, a.last_by,
          ...(a.changes || []).flatMap(c => [c.field_name, c.old_value, c.new_value])].join(' ').toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });

    const total = filtered.length;
    const totalPages = Math.max(1, Math.ceil(total / dashAuditPageSize));
    dashAuditCurrentPage = Math.min(Math.max(1, dashAuditCurrentPage), totalPages);
    const startIdx = (dashAuditCurrentPage - 1) * dashAuditPageSize;
    const pageItems = filtered.slice(startIdx, startIdx + dashAuditPageSize);

    if (dom.dashAuditCountLabel) {
      dom.dashAuditCountLabel.textContent = total
        ? `Showing ${(startIdx + 1).toLocaleString()}–${(startIdx + pageItems.length).toLocaleString()} of ${total.toLocaleString()} account${total === 1 ? '' : 's'}`
        : 'No changes';
    }

    if (!pageItems.length) {
      const msg = (q || dashSourceFilter !== 'ALL') ? 'No accounts match. Try clearing the search or source filter.' : 'No accounts changed in this work week.';
      dom.dashAuditTableBody.innerHTML = `<tr><td colspan="5" class="ov-empty-row">${escapeHtml(msg)}</td></tr>`;
    } else {
      dom.dashAuditTableBody.innerHTML = pageItems.map(a => {
        const open = dashExpandedAccounts.has(a.account);
        const changes = a.changes || [];
        const fieldChips = a.is_new && !changes.length
          ? '<span class="ov-chip mut">New account</span>'
          : changes.slice(0, 2).map(c => `<span class="ov-chip mut">${escapeHtml(prettyFieldName(c.field_name))}</span>`).join('')
            + (changes.length > 2 ? `<span class="ov-chip mut" title="${escapeHtml(changes.slice(2).map(c => prettyFieldName(c.field_name)).join(', '))}">+${changes.length - 2} more</span>` : '');
        const primarySrc = a.is_new ? 'new' : ((a.sources || [])[0] || 'manual');
        const sub = [a.area_index && a.area_index !== a.account ? a.area_index : '', a.mode ? titleCaseStatus(a.mode) : '', a.contract_price ? formatPesoCompact(a.contract_price) : ''].filter(Boolean).join(' · ');
        let html = `
          <tr class="ov-row ${open ? 'open' : ''}" data-acct="${escapeHtml(a.account)}">
            <td class="ov-col-caret"><button type="button" class="ov-caret" aria-expanded="${open}" aria-label="${open ? 'Collapse' : 'Expand'} ${escapeHtml(a.account)}">${open ? '▾' : '▸'}</button></td>
            <td class="ov-acct"><b>${escapeHtml(a.account)}</b><span>${escapeHtml(sub)}</span></td>
            <td><div class="ov-fields">${fieldChips}</div></td>
            <td><span class="ov-chip src-${primarySrc}">${escapeHtml(DASH_SOURCE_LABELS[primarySrc] || primarySrc)}</span>${(a.sources || []).length > 1 ? `<span class="ov-muted"> +${a.sources.length - 1}</span>` : ''}</td>
            <td class="ov-when"><span>${escapeHtml(formatShortDateTime(a.last_time))}</span><span class="ov-by">${escapeHtml(a.last_by || 'Unknown')}</span></td>
          </tr>`;
        if (open) {
          // One line per change: field name, then old → new (wraps inside its own line).
          const lines = changes.map(c => {
            const shown = (v) => { const t = String(formatLogValue(c.field_name, v) || '').trim(); return t === '-' ? '' : t; };
            const oldV = shown(c.old_value);
            const newV = shown(c.new_value);
            const meta = [DASH_SOURCE_LABELS[c.source] || '', c.log_time ? formatShortDateTime(c.log_time) : '', c.updated_by || ''].filter(Boolean).join(' · ');
            return `<div class="ov-dline"${meta ? ` title="${escapeHtml(meta)}"` : ''}>
              <span class="ov-df">${escapeHtml(prettyFieldName(c.field_name))}</span>
              <span class="ov-dvals">
                <span class="ov-do${oldV ? '' : ' blank'}">${escapeHtml(oldV || '(blank)')}</span>
                <span class="ov-darrow" aria-hidden="true">→</span>
                <span class="ov-dn${newV ? '' : ' blank'}">${escapeHtml(newV || '(blank)')}</span>
              </span>
            </div>`;
          }).join('');
          html += `
          <tr class="ov-expand"><td></td><td colspan="4">
            ${lines ? `<div class="ov-diff">${lines}</div>` : '<p class="ov-muted">New account added this week.</p>'}
            <div class="ov-rowacts"><button type="button" class="btn sm" data-open="${escapeHtml(a.account)}">Open in Land Master</button></div>
          </td></tr>`;
        }
        return html;
      }).join('');

      dom.dashAuditTableBody.querySelectorAll('tr.ov-row').forEach(tr => {
        tr.addEventListener('click', (e) => {
          if (e.target.closest('a, input, select')) return;
          const acct = tr.dataset.acct;
          if (dashExpandedAccounts.has(acct)) dashExpandedAccounts.delete(acct); else dashExpandedAccounts.add(acct);
          renderDashboardAuditTable();
        });
      });
      dom.dashAuditTableBody.querySelectorAll('[data-open]').forEach(b => b.addEventListener('click', () => window.jumpToDashboardRecord(b.dataset.open)));
    }
    renderDashboardPaginationControls(totalPages);
  }

  function prettyFieldName(name) {
    const n = String(name || '').trim();
    const map = { 'total payment deductible to tcp': 'Payment', 'total contract price': 'Contract price', '% paid': '% paid', 'cr63f_conversionstatus': 'Conversion status' };
    const k = n.toLowerCase();
    if (map[k]) return map[k];
    if (/^cr63f_|^new_/.test(n)) return n.replace(/^(cr63f_|new_)/, '');
    return n.charAt(0).toUpperCase() + n.slice(1).toLowerCase().replace(/\b(dar|nia|tcp|rcc)\b/g, m => m.toUpperCase());
  }

  function renderDashboardPaginationControls(totalPages) {
    const el = dom.dashAuditPagination;
    if (!el) return;
    if (totalPages <= 1) { el.innerHTML = ''; return; }
    el.innerHTML = `
      <button type="button" class="dash-page-btn" id="btnDashAuditPrev" aria-label="Previous page" ${dashAuditCurrentPage <= 1 ? 'disabled' : ''}>‹</button>
      <span class="ov-pageinfo">Page ${dashAuditCurrentPage} of ${totalPages}</span>
      <button type="button" class="dash-page-btn" id="btnDashAuditNext" aria-label="Next page" ${dashAuditCurrentPage >= totalPages ? 'disabled' : ''}>›</button>
      <select class="dash-page-size-select" id="selDashAuditPageSize" aria-label="Accounts per page">
        ${[25, 50, 100].map(n => `<option value="${n}" ${dashAuditPageSize === n ? 'selected' : ''}>${n} / page</option>`).join('')}
      </select>`;
    const p = el.querySelector('#btnDashAuditPrev'), n = el.querySelector('#btnDashAuditNext'), s = el.querySelector('#selDashAuditPageSize');
    if (p) p.addEventListener('click', () => { dashAuditCurrentPage--; renderDashboardAuditTable(); });
    if (n) n.addEventListener('click', () => { dashAuditCurrentPage++; renderDashboardAuditTable(); });
    if (s) s.addEventListener('change', () => { dashAuditPageSize = parseInt(s.value, 10) || 25; dashAuditCurrentPage = 1; renderDashboardAuditTable(); });
  }

  window.jumpToDashboardRecord = function (query) {
    if (!query) return;
    const cleanQ = String(query).trim();
    if (dom.searchInput) dom.searchInput.value = cleanQ;
    // Clear both filters: a Type left selected could hide the account (0 records).
    getTypeFilter().setSelected([]);
    getModeFilter().setSelected([]);
    populateTypeFilter(state.records);
    populateModeFilter(state.records);
    switchView('records');
    filterAndRenderRecords();
    showToast(`Showing Land Master records matching "${cleanQ}".`, 'ok');
  };

  // One-time (repeatable) fix: Balance and % Paid recalculated for every account
  // from its own TCP and Deductible. Preview first, then the audited save.
  async function recalculatePaymentFields() {
    const btn = document.getElementById('btnRecalcPayments');
    if (Object.keys(state.stagedEdits || {}).length || (state.financeAutopopulatedRows && state.financeAutopopulatedRows.size)) {
      showToast('Save or discard your Land Master edits first.', 'err');
      return;
    }
    try {
      if (btn) { btn.disabled = true; btn.textContent = 'Checking…'; }
      if (!state.records || !state.records.length) await fetchRecords();
      const preview = await window.pywebview.api.recalculate_payment_fields(false);
      if (!preview || !preview.success) {
        showToast((preview && preview.error) || 'Could not check the accounts.', 'err');
        return;
      }
      if (!preview.to_change) {
        showToast(`All accounts already have the right Balance and % Paid (${preview.skipped_no_tcp} without a TCP left as they are).`, 'info');
        return;
      }
      const sample = (preview.samples || []).slice(0, 3)
        .map(s => `${s.account}: % Paid ${s.paid}`).join('; ');
      const ok = await showConfirmModal({
        title: 'Recalculate Balance & % Paid',
        subtitle: 'Land Master maintenance',
        message: `${preview.to_change.toLocaleString()} account(s) will get a recalculated Balance and/or % Paid `
          + `(e.g. ${sample}). ${preview.unchanged.toLocaleString()} are already right; `
          + `${preview.skipped_no_tcp.toLocaleString()} without a TCP are left as they are. `
          + `This sends about ${preview.estimated_requests.toLocaleString()} requests to Dataverse and can take several minutes; `
          + 'every account gets an audit snapshot. Export Land Master to Excel first if you have not made a backup.',
        confirmText: 'Recalculate & Log',
        cancelText: 'Cancel',
        isDanger: false,
        keepOpenOnConfirm: true
      });
      if (!ok) return;
      if (dom.btnConfirmModalOk) { dom.btnConfirmModalOk.disabled = true; dom.btnConfirmModalOk.textContent = 'Saving...'; }
      if (dom.btnConfirmModalCancel) dom.btnConfirmModalCancel.classList.add('hidden');
      if (dom.btnConfirmModalClose) dom.btnConfirmModalClose.classList.add('hidden');
      setConfirmModalProgress(0, preview.to_change, 'Checking audit history...');
      const res = await window.pywebview.api.recalculate_payment_fields(true);
      closeConfirmModal(false);
      if (res && res.success) {
        showToast(res.message, 'ok');
      } else {
        showToast(`${(res && res.message) || ''} ${(res && res.error) || 'Recalculation failed.'}`.trim(), 'err');
      }
      await fetchRecords();
    } catch (err) {
      closeConfirmModal(false);
      showToast('Error: ' + err, 'err');
    } finally {
      if (dom.btnConfirmModalOk) { dom.btnConfirmModalOk.disabled = false; dom.btnConfirmModalOk.textContent = 'Confirm'; }
      if (dom.btnConfirmModalCancel) dom.btnConfirmModalCancel.classList.remove('hidden');
      if (dom.btnConfirmModalClose) dom.btnConfirmModalClose.classList.remove('hidden');
      if (btn) { btn.disabled = false; btn.textContent = 'Recalculate…'; }
    }
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
        logs_table: dom.selLogsTable ? dom.selLogsTable.value : state.activeLogsTable,
        finance_table: dom.selFinanceTable.value,
        conversion_table: dom.selConversionTable ? dom.selConversionTable.value : state.activeConversionTable,
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

    const btnRecalc = document.getElementById('btnRecalcPayments');
    if (btnRecalc) btnRecalc.addEventListener('click', recalculatePaymentFields);

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
        // Exactly one option is selected: the saved solution, else LOA MONITORING.
        // (Marking both made the browser pick whichever came last.)
        const saved = res.solutions.find(s => state.activeSolutionId &&
          (state.activeSolutionId === s.id || state.activeSolutionId === s.uniqueName));
        const fallback = res.solutions.find(s => s.uniqueName === 'LOAMONITORING' || s.friendlyName === 'LOA MONITORING');
        const chosen = saved || fallback;
        res.solutions.forEach(s => {
          const isSelected = chosen && s.id === chosen.id;
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

        if (dom.selLogsTable) {
          let logsHtml = '';
          res.tables.forEach(t => {
            const isSel = t.logicalName === state.activeLogsTable ||
                          (state.activeLogsTable === 'cr63f_loamonitoringlogs' && t.logicalName === 'cr63f_loamonitoringlogs') ||
                          (!state.activeLogsTable && t.logicalName.toLowerCase().includes('log'));
            logsHtml += `<option value="${escapeHtml(t.logicalName)}" ${isSel ? 'selected' : ''}>${escapeHtml(t.displayName)} (${escapeHtml(t.logicalName)})</option>`;
          });
          dom.selLogsTable.innerHTML = logsHtml || '<option value="cr63f_loamonitoringlogs">LOA MONITORING LOGS (cr63f_loamonitoringlogs)</option>';
        }

        let finHtml = '';
        res.tables.forEach(t => {
          const isSel = t.logicalName === state.activeFinanceTable || (state.activeFinanceTable === 'cr63f_financereference' && t.logicalName === 'cr63f_financereference');
          finHtml += `<option value="${escapeHtml(t.logicalName)}" ${isSel ? 'selected' : ''}>${escapeHtml(t.displayName)} (${escapeHtml(t.logicalName)})</option>`;
        });
        dom.selFinanceTable.innerHTML = finHtml || '<option value="cr63f_financereference">FINANCE REFERENCE (cr63f_financereference)</option>';

        if (dom.selConversionTable) {
          let convHtml = '';
          res.tables.forEach(t => {
            const isSel = t.logicalName === state.activeConversionTable || (state.activeConversionTable === 'cr63f_landconversionreference' && t.logicalName === 'cr63f_landconversionreference');
            convHtml += `<option value="${escapeHtml(t.logicalName)}" ${isSel ? 'selected' : ''}>${escapeHtml(t.displayName)} (${escapeHtml(t.logicalName)})</option>`;
          });
          dom.selConversionTable.innerHTML = convHtml || '<option value="cr63f_landconversionreference">LAND CONVERSION REFERENCE (cr63f_landconversionreference)</option>';
        }
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
    state.activeLogsTable = cfg.logs_table || 'cr63f_loamonitoringlogs';
    state.activeFinanceTable = cfg.finance_table || 'cr63f_financereference';
    state.activeConversionTable = cfg.conversion_table || 'cr63f_landconversionreference';
    state.activeSolutionId = cfg.solution_id || '';
    if (dom.selLogsTable) dom.selLogsTable.value = state.activeLogsTable;
    if (dom.selConversionTable) dom.selConversionTable.value = state.activeConversionTable;
  }

  // ==================== APPLICATION UPDATES (GitHub Releases) ====================
  const UPDATE_BUSY = ['checking', 'downloading', 'installing'];
  let promptedUpdateVersion = '';
  let promptedRestartVersion = '';

  function updateApi() {
    return (window.pywebview && window.pywebview.api) ? window.pywebview.api : null;
  }

  function renderUpdate(u) {
    if (!u) return;
    const $ = (id) => document.getElementById(id);
    const status = u.status || 'idle';
    const busy = UPDATE_BUSY.includes(status);
    const managed = u.configured !== false && u.active !== false;

    if (u.currentVersion) setAppVersionLabels(u.currentVersion);

    const titles = {
      idle: 'Updates are checked when the app starts',
      checking: 'Checking for updates…',
      up_to_date: 'Up to date',
      available: `Version ${u.latestVersion} is available`,
      downloading: `Downloading update… ${Number(u.progress) || 0}%`,
      ready: `Version ${u.latestVersion} is ready to install`,
      installing: 'Installing update…',
      error: 'Update problem',
      disabled: 'In-app updates are off for this copy',
    };
    const statusEl = $('updateStatus');
    if (statusEl) statusEl.textContent = titles[status] || u.message || '';

    const hint = $('updateHint');
    if (hint) {
      if (status === 'error') hint.textContent = u.message || 'Try again later.';
      else if (!managed) hint.textContent = (u.message ? u.message + ' ' : '') + 'Install the latest Setup.exe from the GitHub releases page to receive updates in the app.';
      else if (status === 'ready') hint.textContent = 'The app will close, install the update and reopen. Save your work first.';
      else if (status === 'available') hint.textContent = 'Download it now; you choose when to restart.';
      else hint.textContent = 'New versions are published as GitHub releases and checked each time the app starts.';
    }

    const dot = $('updateDot');
    if (dot) {
      dot.className = 'update-dot';
      if (busy) dot.classList.add('busy');
      else if (status === 'available' || status === 'ready') dot.classList.add('new');
      else if (status === 'up_to_date') dot.classList.add('ready');
      else if (status === 'error') dot.classList.add('error');
    }

    const pct = Math.max(0, Math.min(100, Number(u.progress) || 0));
    const progress = $('updateProgress');
    if (progress) {
      progress.classList.toggle('hidden', !['downloading', 'ready'].includes(status));
      progress.setAttribute('aria-valuenow', String(pct));
    }
    const bar = $('updateProgressBar');
    if (bar) bar.style.width = pct + '%';

    const notes = $('updateNotes');
    if (notes) {
      notes.textContent = u.releaseNotes || '';
      notes.classList.toggle('hidden', !u.releaseNotes || !['available', 'downloading', 'ready'].includes(status));
    }

    const latest = $('updateLatest');
    if (latest) latest.textContent = u.latestVersion ? `LATEST ${u.latestVersion}` : '';

    const check = $('btnCheckUpdate');
    if (check) {
      check.disabled = busy || !managed || status === 'ready';
      check.textContent = status === 'checking' ? 'Checking…' : 'Check for updates';
    }
    const download = $('btnDownloadUpdate');
    if (download) download.classList.toggle('hidden', status !== 'available');
    const install = $('btnInstallUpdate');
    if (install) install.classList.toggle('hidden', status !== 'ready');
  }

  // The shared confirm modal holds one pending answer; never open an update
  // prompt over a save/delete confirmation the user is answering.
  function confirmModalBusy() {
    return !!(dom.confirmModal && !dom.confirmModal.classList.contains('hidden'));
  }

  async function startUpdateDownload() {
    const api = updateApi();
    if (!api || !api.download_update) return;
    const res = await api.download_update();
    if (res && res.update) renderUpdate(res.update);
  }

  async function installUpdateNow() {
    const api = updateApi();
    if (!api || !api.install_update) return;
    const res = await api.install_update();
    if (res && res.update) renderUpdate(res.update);
    if (res && res.ok === false) showToast(res.error || 'Could not install the update.', 'err');
  }

  async function confirmInstallUpdate(version) {
    const ok = await showConfirmModal({
      title: 'Install update',
      subtitle: version ? `LOA Monitoring Tool ${version}` : 'Downloaded update',
      message: 'The app will close, install the update and reopen automatically. Unsaved grid edits will be lost, so save them first.',
      confirmText: 'Update and restart',
      cancelText: 'Later',
    });
    if (ok) await installUpdateNow();
  }

  window.onUpdateStatus = async function (u) {
    renderUpdate(u);
    if (!u) return;
    if (u.status === 'available' && u.latestVersion && promptedUpdateVersion !== u.latestVersion) {
      promptedUpdateVersion = u.latestVersion;
      if (confirmModalBusy()) {
        showToast(`Version ${u.latestVersion} is available. Download it from Settings.`, 'info');
        return;
      }
      const ok = await showConfirmModal({
        title: 'Update available',
        subtitle: `LOA Monitoring Tool ${u.latestVersion} (you have ${u.currentVersion || 'an older version'})`,
        message: u.releaseNotes ? `What's new:\n\n${u.releaseNotes}` : 'A new version has been published.',
        confirmText: 'Download update',
        cancelText: 'Later',
      });
      if (ok) await startUpdateDownload();
    } else if (u.status === 'ready' && u.latestVersion && promptedRestartVersion !== u.latestVersion) {
      promptedRestartVersion = u.latestVersion;
      if (confirmModalBusy()) {
        showToast(`Version ${u.latestVersion} is downloaded. Restart from Settings to install it.`, 'info');
        return;
      }
      await confirmInstallUpdate(u.latestVersion);
    } else if (u.status === 'error' && !u.quiet) {
      showToast(u.message || 'Update failed.', 'err');
    }
  };

  function wireUpdateControls() {
    const check = document.getElementById('btnCheckUpdate');
    if (check) check.addEventListener('click', async () => {
      const api = updateApi();
      if (!api || !api.check_for_updates) return;
      const res = await api.check_for_updates(true);
      if (res && res.update) renderUpdate(res.update);
    });
    const download = document.getElementById('btnDownloadUpdate');
    if (download) download.addEventListener('click', startUpdateDownload);
    const install = document.getElementById('btnInstallUpdate');
    if (install) install.addEventListener('click', () => {
      const latest = document.getElementById('updateLatest');
      confirmInstallUpdate(latest ? latest.textContent.replace(/^LATEST\s*/, '') : '');
    });
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

/**
 * Styled hover cards for every `title` in the app (added 2026-10-14).
 * The native browser tooltip is plain and slow, so on first hover each `title`
 * moves to data-tip and is shown in a themed card instead. Text written as
 * "Label: value" lines becomes aligned rows; a "Total:" row is emphasised; a
 * paragraph after a blank line becomes a muted footnote. Code that later sets
 * el.title again still works: the new text is picked up on the next hover.
 */
(function () {
  'use strict';
  const SHOW_DELAY = 220;
  let tipEl = null;
  let current = null;
  let showTimer = null;

  function esc(text) {
    return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function ensureTip() {
    if (tipEl) return tipEl;
    tipEl = document.createElement('div');
    tipEl.className = 'app-tip';
    tipEl.setAttribute('role', 'tooltip');
    tipEl.id = 'appHoverTip';
    document.body.appendChild(tipEl);
    return tipEl;
  }

  const ROW = /^([^:\n]{1,64}):\s+(.+)$/;

  function renderBody(text) {
    const blocks = String(text).trim().split(/\n\s*\n/);
    return blocks.map((block, bi) => {
      const lines = block.split('\n').map(l => l.trim()).filter(Boolean);
      const rowLines = lines.filter(l => ROW.test(l) && !/[.!?]$/.test(l.split(':').slice(1).join(':').trim()));
      const isTable = rowLines.length >= Math.max(1, Math.ceil(lines.length * 0.6)) && !(bi > 0 && lines.length === 1 && /[.!?]$/.test(lines[0]));
      if (!isTable) {
        const cls = bi > 0 ? 'app-tip-note' : 'app-tip-text';
        return `<div class="${cls}">${lines.map(esc).join('<br>')}</div>`;
      }
      const rows = lines.map(line => {
        const m = line.match(ROW);
        if (!m) return `<div class="app-tip-text app-tip-span">${esc(line)}</div>`;
        let [, label, value] = m;
        const total = /^total\b/i.test(label.trim());
        let flag = '';
        const fm = value.match(/\s*\((not counted|excluded|not in % paid)\)\s*$/i);
        if (fm) { flag = fm[1]; value = value.slice(0, fm.index); }
        return `<div class="app-tip-row${total ? ' total' : ''}${flag ? ' muted' : ''}">`
          + `<span class="app-tip-label">${esc(label.trim())}</span>`
          + `<span class="app-tip-value">${esc(value.trim())}</span></div>`;
      }).join('');
      return `<div class="app-tip-rows">${rows}</div>`;
    }).join('');
  }

  function headingFor(el) {
    if (el.dataset.tipTitle) return el.dataset.tipTitle;
    const card = el.closest('.kpi-card, .ov-kpi');
    if (card && (el === card || el.classList.contains('kpi-value') || el.classList.contains('ov-kpi-v'))) {
      const label = card.querySelector('.kpi-label, .ov-kpi-l');
      if (label) return label.textContent.trim();
    }
    return '';
  }

  function claim(el) {
    const t = el.getAttribute('title');
    if (t !== null) {
      el.dataset.tip = t;
      el.removeAttribute('title');
      if (!el.getAttribute('aria-label') && !el.textContent.trim() && t.trim()) el.setAttribute('aria-label', t.trim());
    }
    return el.dataset.tip || '';
  }

  function place(el) {
    const tip = ensureTip();
    const r = el.getBoundingClientRect();
    const vw = document.documentElement.clientWidth;
    const vh = document.documentElement.clientHeight;
    const tw = tip.offsetWidth;
    const th = tip.offsetHeight;
    const gap = 8;
    // Large targets (cards, panels): open under the top area near the pointer side.
    let left = r.width > tw ? r.left + 12 : r.left + r.width / 2 - tw / 2;
    left = Math.max(8, Math.min(left, vw - tw - 8));
    let top = r.bottom + gap;
    let below = true;
    if (top + th > vh - 8 && r.top - gap - th >= 8) { top = r.top - gap - th; below = false; }
    top = Math.max(8, Math.min(top, vh - th - 8));
    tip.style.left = `${Math.round(left)}px`;
    tip.style.top = `${Math.round(top)}px`;
    tip.dataset.side = below ? 'below' : 'above';
  }

  function show(el) {
    const text = claim(el);
    if (!text.trim()) return;
    const tip = ensureTip();
    let body = text;
    let heading = headingFor(el);
    const hm = body.match(/^#\s+(.+)\n?/);
    if (hm) { heading = hm[1].trim(); body = body.slice(hm[0].length); }
    tip.innerHTML = (heading ? `<div class="app-tip-head">${esc(heading)}</div>` : '') + (body.trim() ? renderBody(body) : '');
    tip.classList.toggle('compact', !heading && !/\n/.test(body) && body.length <= 60);
    tip.classList.add('measuring');
    place(el);
    tip.classList.remove('measuring');
    tip.classList.add('visible');
    el.setAttribute('aria-describedby', 'appHoverTip');
  }

  function hide() {
    clearTimeout(showTimer);
    showTimer = null;
    if (current) current.removeAttribute('aria-describedby');
    current = null;
    if (tipEl) tipEl.classList.remove('visible');
  }

  function targetOf(node) {
    if (!node || !node.closest) return null;
    if (node.closest('.cell-hover-tooltip, .app-tip')) return null;
    return node.closest('[title], [data-tip]');
  }

  function enter(el) {
    if (el === current) return;
    hide();
    claim(el);
    if (!el.dataset.tip || !el.dataset.tip.trim()) return;
    current = el;
    showTimer = setTimeout(() => { if (current === el && el.isConnected) show(el); }, SHOW_DELAY);
  }

  document.addEventListener('mouseover', (e) => {
    const el = targetOf(e.target);
    if (el) enter(el);
    else if (current) hide();
  }, true);
  document.addEventListener('mouseout', (e) => {
    if (!current) return;
    const to = e.relatedTarget;
    if (!to || !current.contains(to)) {
      const next = targetOf(to);
      if (next !== current) hide();
    }
  }, true);
  document.addEventListener('focusin', (e) => {
    const el = targetOf(e.target);
    if (el && e.target.matches(':focus-visible')) { enter(el); }
  });
  document.addEventListener('focusout', hide);
  ['mousedown', 'keydown', 'wheel', 'scroll'].forEach(type =>
    window.addEventListener(type, (e) => { if (type === 'keydown' && e.key !== 'Escape' && !current) return; hide(); }, { capture: true, passive: true }));
  window.addEventListener('blur', hide);
  window.addEventListener('resize', hide);
})();
