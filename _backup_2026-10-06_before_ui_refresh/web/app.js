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
    { key: '__action', label: 'Actions', width: 55, fixed: true, readonly: true }
  ];

  const LOGS_COLUMNS = [
    { key: '__idx', label: '#', width: 45, fixed: true, readonly: true },
    { key: 'log_time', alt: 'Log Time', label: 'Log Time', width: 155, readonly: true },
    { key: 'work_week', alt: 'Work Week', label: 'Work Week', width: 100, readonly: true },
    { key: 'action_type', alt: 'Action Type', label: 'Action Type', width: 135, readonly: true },
    { key: 'cr63f_acquisitionsublevel', alt: 'Acquisition Sub-Level', label: 'Acquisition Sub-Level', width: 155, readonly: true },
    { key: 'cr63f_control1', alt: 'Control', label: 'Control Number', width: 140, readonly: true },
    { key: 'area_index', alt: 'cr63f_areaindex', label: 'AREA-INDEX', width: 135, readonly: true },
    { key: 'field_name', alt: 'Field Changed', label: 'Field Changed', width: 165, readonly: true },
    { key: 'old_value', alt: 'Previous Value', label: 'Previous Value', width: 145, readonly: true },
    { key: 'new_value', alt: 'New Value', label: 'New Value', width: 145, readonly: true },
    { key: 'title_no', alt: 'cr63f_titleno', label: 'Title No', width: 140, readonly: true },
    { key: 'details', alt: 'Details', label: 'Details', width: 220, readonly: true },
    { key: '__action', label: 'Actions', width: 75, fixed: true, readonly: true }
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
    btnCheckUpdates: document.getElementById('btnCheckUpdates'),

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
    if (!table || !tip) return;
    if (table.dataset.tooltipWired) return;
    table.dataset.tooltipWired = '1';

    function isLogForRecord(l, rowRecord, tLotId) {
      if (!l) return false;
      const lCtrl = String(l.cr63f_control1 || l['Control Number'] || l.Control || l['Control No'] || '').trim().toUpperCase();
      const lArea = String(l.area_index || l['AREA-INDEX'] || l['Area Index'] || l['Area-Index'] || l.cr63f_areaindex || '').trim().toUpperCase();
      const lTitle = String(l.title_no || l['Title No'] || l['TITLE NO'] || l.cr63f_titleno || '').trim().toUpperCase();
      const lId = String(l.record_id || l.cr63f_loamonitoringid || '').trim();

      if (rowRecord) {
        const rId = String(rowRecord.__id || rowRecord.id || rowRecord.cr63f_loamonitoringid || '').trim();
        if (lId && rId && lId === rId) return true;
        const rCtrl = String(rowRecord.cr63f_control1 || rowRecord['Control Number'] || rowRecord.Control || rowRecord['Control No'] || '').trim().toUpperCase();
        if (lCtrl && rCtrl && lCtrl === rCtrl) return true;
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
        prevVal = getPreviousValueFromLogs(cellLogs, liveVal, colDef);
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

      const canRevert = Boolean(prevVal !== undefined && prevVal !== null && String(prevVal).trim() !== '' && String(prevVal).trim() !== '—' && String(prevVal).trim() !== '-');
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
    const btn = btnEl || (e && e.target ? e.target.closest('.cell-tooltip-revert-btn') : null) || (tip ? tip.querySelector('.cell-tooltip-revert-btn') : null);

    const rowId = (btn && btn.dataset.rowId) || (tip && tip._rowId) || (tip && tip._sourceCell && tip._sourceCell.dataset.rowId);
    const colKey = (btn && btn.dataset.colKey) || (tip && tip._colKey) || (tip && tip._sourceCell && tip._sourceCell.dataset.colKey) || '';
    const fieldName = (btn && btn.dataset.fieldName) || (tip && tip._fieldName) || colKey;
    let targetRevertVal = (btn && btn.dataset.revertVal !== undefined && btn.dataset.revertVal !== '')
      ? btn.dataset.revertVal
      : ((tip && (tip._targetRevertVal ?? tip._prevVal)) || '');

    // Hide tooltip immediately
    if (tip) {
      tip.classList.remove('visible');
      tip.style.pointerEvents = 'none';
      tip.setAttribute('aria-hidden', 'true');
      tip._sourceCell = null;
      tip._isHovered = false;
    }

    if (!rowId) return;

    const tr = document.querySelector(`#recordsTable tbody tr[data-row-id="${rowId}"]`) || (tip && tip._sourceCell ? tip._sourceCell.closest('tr') : null);
    const colDef = DEFAULT_COLUMNS.find(c => c.key === colKey || c.alt === colKey || c.label === fieldName) || { key: colKey, label: fieldName };
    const td = tr ? tr.querySelector(`td[data-col-key="${colDef.key}"]`) : (tip ? tip._sourceCell : null);
    const cellInput = td ? td.querySelector('input, select') : null;
    const rowRecord = state.records.find(r => String(r.__id) === String(rowId));

    if (!rowRecord && !tr) return;

    // Clean targetRevertVal for storage in stagedEdits
    let cleanVal = String(targetRevertVal ?? '').trim();
    if (colDef.isPrice || /price|tcp|deductible|balance|amount|cost/i.test(`${colDef.key} ${fieldName}`)) {
      cleanVal = cleanVal.replace(/[₱$,\s]/g, '');
    } else if (colDef.isPercent) {
      cleanVal = formatPercent(cleanVal);
    }

    // Formatted value for putting directly into the cell input
    let formattedVal = cleanVal;
    if (colDef.isPrice || /price|tcp|deductible|balance|amount|cost/i.test(`${colDef.key} ${fieldName}`)) {
      formattedVal = formatPrice(cleanVal);
    } else if (colDef.isPercent) {
      formattedVal = formatPercent(cleanVal);
    }

    // 1. Put the WAS value directly into the cell DOM input!
    if (cellInput) {
      cellInput.value = formattedVal;
      cellInput.classList.add('modified');
    }
    if (td) {
      td.classList.add('is-dirty');
      td.classList.remove('cell-reference-updated');
      td.dataset.currValue = formattedVal;
      td.dataset.diffType = 'manual';
    }

    // 2. Stage the edit in state.stagedEdits with action_type = 'REVERT'
    state.stagedEdits[rowId] = state.stagedEdits[rowId] || {};
    state.stagedEdits[rowId][colDef.key] = cleanVal;
    if (colDef.alt) state.stagedEdits[rowId][colDef.alt] = cleanVal;
    state.stagedEdits[rowId].__action_type = 'REVERT';

    // 3. Clear reference-updated status on this field so it's treated as a staged manual revert
    if (state.referenceUpdatedFieldsByRowId && state.referenceUpdatedFieldsByRowId[rowId]) {
      state.referenceUpdatedFieldsByRowId[rowId].delete(colDef.key);
      if (colDef.alt) state.referenceUpdatedFieldsByRowId[rowId].delete(colDef.alt);
    }
    if (rowRecord && Array.isArray(rowRecord.__changed_fields)) {
      rowRecord.__changed_fields = rowRecord.__changed_fields.filter(f => f !== colDef.key && f !== colDef.alt && f !== colDef.label);
    }

    // 4. Live calculation of financial fields (balance & % paid)
    const paymentKeys = ['cr63f_totalcontractprice', 'Total Contract Price', 'cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP'];
    if (paymentKeys.includes(colDef.key) || /price|tcp|deductible/i.test(`${colDef.key} ${fieldName}`)) {
      const tcpRaw = state.stagedEdits[rowId]['cr63f_totalcontractprice'] ?? (rowRecord ? (rowRecord['cr63f_totalcontractprice'] ?? rowRecord['Total Contract Price'] ?? '') : '');
      const dedRaw = state.stagedEdits[rowId]['cr63f_totalpaymentdeductibletotcp'] ?? (rowRecord ? (rowRecord['cr63f_totalpaymentdeductibletotcp'] ?? rowRecord['Total Payment Deductible to TCP'] ?? '') : '');
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

      if (tr) {
        const balInput = tr.querySelector('[data-col-key="cr63f_balance"] input');
        if (balInput) {
          balInput.value = formatPrice(balFmt);
          balInput.classList.add('modified');
        }
        const pctInput = tr.querySelector('[data-col-key="cr63f_paid"] input');
        if (pctInput) {
          pctInput.value = pctFmt;
          pctInput.classList.add('modified');
        }
      }
    }

    // 5. Flash pulse animation
    if (td) {
      td.classList.add('cell-reverted-highlight');
      setTimeout(() => { td.classList.remove('cell-reverted-highlight'); }, 1500);
    }

    // 6. Show the staged edits bar ("1 modified row" with [Save to Dataverse] and Discard)
    updateStagedEditsUI();
    recalculateKPIs();

    showToast(`Reverted ${fieldName} to ${formattedVal || '(blank)'}. Click "Save to Dataverse" to commit.`, 'ok');
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
        await loadDashboardAnalytics(dom.dashWorkWeekSelect ? dom.dashWorkWeekSelect.value : 'ALL');
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
      dom.lblTopFinance.textContent = `Finance Data: ${state.financeLoadedCount.toLocaleString()} rows`;
      if (dom.lblActiveFinanceCount) dom.lblActiveFinanceCount.textContent = `${state.financeLoadedCount.toLocaleString()} rows`;
      if (dom.lblActiveFinanceSource) dom.lblActiveFinanceSource.textContent = state.financeSource;
    } else {
      dom.dotFinanceRef.classList.remove('on');
      dom.lblTopFinance.textContent = 'Finance Data: None';
      if (dom.lblActiveFinanceCount) dom.lblActiveFinanceCount.textContent = '0 rows';
      if (dom.lblActiveFinanceSource) dom.lblActiveFinanceSource.textContent = 'None';
    }

    // Update conversion reference status
    if (state.conversionLoadedCount > 0) {
      if (dom.dotConversionRef) dom.dotConversionRef.classList.add('on');
      if (dom.lblTopConversion) dom.lblTopConversion.textContent = `Conversion Data: ${state.conversionLoadedCount.toLocaleString()} rows`;
      if (dom.lblActiveConversionCount) dom.lblActiveConversionCount.textContent = `${state.conversionLoadedCount.toLocaleString()} rows`;
      if (dom.lblActiveConversionSource) dom.lblActiveConversionSource.textContent = state.conversionSource;
    } else {
      if (dom.dotConversionRef) dom.dotConversionRef.classList.remove('on');
      if (dom.lblTopConversion) dom.lblTopConversion.textContent = 'Conversion Data: None';
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
      records: { label: 'Reload Land Master', description: 'Reload Land Master from Dataverse' },
      finance: { label: 'Load from Dataverse', description: 'Load finance data from Dataverse' },
      conversion: { label: 'Load from Dataverse', description: 'Load conversion data from Dataverse' }
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
      loadDashboardAnalytics(dom.dashWorkWeekSelect ? dom.dashWorkWeekSelect.value : 'ALL');
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

  function formatAmountWithCommas(val) {
    if (val === null || val === undefined) return '';
    let s = String(val).trim();
    if (!s) return '';
    s = s.replace(/[^0-9.-]/g, '');
    if (!s || s === '-' || s === '.') return s;

    const parts = s.split('.');
    const integerPart = parts[0];
    const decimalPart = parts.length > 1 ? parts.slice(1).join('') : null;

    const sign = integerPart.startsWith('-') ? '-' : '';
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
      m = parseInt(usMatch[1], 10) - 1;
      d = parseInt(usMatch[2], 10);
    } else {
      const dt = new Date(s);
      if (isNaN(dt.getTime())) return '';
      y = dt.getFullYear();
      m = dt.getMonth();
      d = dt.getDate();
    }
    if (y <= 1900) return '';
    const dateObj = new Date(Date.UTC(y, m, d));
    const pyWeekday = (dateObj.getUTCDay() + 6) % 7;
    const daysSinceWed = (pyWeekday >= 2) ? (pyWeekday - 2) : (pyWeekday + 5);
    const wedUtc = new Date(dateObj.getTime() - daysSinceWed * 86400000);
    const midUtc = new Date(wedUtc.getTime() + 3 * 86400000);
    const primaryYear = midUtc.getUTCFullYear();
    const jan1 = new Date(Date.UTC(primaryYear, 0, 1));
    const jan1PyWeekday = (jan1.getUTCDay() + 6) % 7;
    const jan1DaysToWed = (jan1PyWeekday >= 2) ? (jan1PyWeekday - 2) : (jan1PyWeekday + 5);
    const ww1Wed = new Date(jan1.getTime() - jan1DaysToWed * 86400000);
    let weekNum = Math.floor((wedUtc.getTime() - ww1Wed.getTime()) / (7 * 86400000)) + 1;
    if (weekNum < 1) {
      const primaryYearPrev = primaryYear - 1;
      const jan1Prev = new Date(Date.UTC(primaryYearPrev, 0, 1));
      const jan1PrevPyWeekday = (jan1Prev.getUTCDay() + 6) % 7;
      const jan1PrevDaysToWed = (jan1PrevPyWeekday >= 2) ? (jan1PrevPyWeekday - 2) : (jan1PrevPyWeekday + 5);
      const ww1WedPrev = new Date(jan1Prev.getTime() - jan1PrevDaysToWed * 86400000);
      weekNum = Math.floor((wedUtc.getTime() - ww1WedPrev.getTime()) / (7 * 86400000)) + 1;
    }
    return `WW-${String(weekNum).padStart(2, '0')}`;
  }

  function formatPercent(val) {
    if (val === null || val === undefined || val === '') return '0.00%';
    const s = String(val).trim();
    const display = num => `${num.toLocaleString('en-US', { useGrouping: false, minimumFractionDigits: 2, maximumFractionDigits: 10 })}%`;
    if (s.endsWith('%')) {
      const num = parseFloat(s.replace('%', '').trim());
      return isNaN(num) ? s : display(num);
    }
    const num = parseFloat(s.replace(/,/g, ''));
    if (isNaN(num)) return s;
    if (num > -1.0 && num < 1.0 && num !== 0) {
      return display(num * 100);
    }
    return display(num);
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
    return String(log.action_type ?? log['Action Type'] ?? log.cr63f_actiontype ?? fallback).toUpperCase();
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
  function setupRecordsView() {
    setupColumnPicker();

    dom.searchInput.addEventListener('input', () => {
      clearTimeout(searchDebounce);
      searchDebounce = setTimeout(filterAndRenderRecords, 120);
    });

    dom.typeFilter.addEventListener('change', () => {
      populateModeFilter(state.records);
      filterAndRenderRecords();
    });

    if (dom.modeFilter) {
      dom.modeFilter.addEventListener('change', () => {
        populateTypeFilter(state.records);
        if (dom.modeFilterPanel) dom.modeFilterPanel.value = dom.modeFilter.value;
        if (dom.kpiModeActiveBadge) {
          dom.kpiModeActiveBadge.textContent = dom.modeFilter.value === 'ALL' ? 'ALL MODES' : dom.modeFilter.value;
        }
        filterAndRenderRecords();
      });
    }

    if (dom.modeFilterPanel) {
      dom.modeFilterPanel.addEventListener('change', () => {
        populateTypeFilter(state.records);
        if (dom.kpiModeActiveBadge) {
          dom.kpiModeActiveBadge.textContent = dom.modeFilterPanel.value === 'ALL' ? 'ALL MODES' : dom.modeFilterPanel.value;
        }
        filterAndRenderRecords();
      });
    }

    if (dom.btnUpdateFromFinance) {
      dom.btnUpdateFromFinance.addEventListener('click', updateRecordsFromFinance);
    }
    dom.btnExportExcel.addEventListener('click', exportRecordsExcel);

    if (dom.btnClearFilters) {
      dom.btnClearFilters.addEventListener('click', () => {
        dom.searchInput.value = '';
        dom.typeFilter.value = 'ALL';
        if (dom.modeFilter) dom.modeFilter.value = 'ALL';
        if (dom.modeFilterPanel) dom.modeFilterPanel.value = 'ALL';
        if (dom.kpiModeActiveBadge) dom.kpiModeActiveBadge.textContent = 'ALL MODES';
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

      const [res, finRes, convRes] = await Promise.all([
        window.pywebview.api.fetch_loa_records(),
        window.pywebview.api.sync_finance_from_dataverse().catch(e => ({ success: false, error: String(e) })),
        window.pywebview.api.sync_conversion_from_dataverse().catch(e => ({ success: false, error: String(e) }))
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
      } else {
        showToast(res.error || 'Failed to fetch Dataverse records.', 'err');
      }

      if (finRes && finRes.success) {
        const s = await window.pywebview.api.get_initial_state();
        applyState(s);
        const sample = finRes.sample_rows || [];
        renderFinancePreview(finRes.headers || (sample.length ? Object.keys(sample[0]) : []), sample);
      }

      if (convRes && convRes.success) {
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
    const selectedMode = normalizeModeFilterValue((dom.modeFilterPanel || dom.modeFilter)?.value || 'ALL');
    const typeCounts = new Map();
    (records || []).forEach(r => {
      const rowMode = normalizeModeFilterValue(r['new_mode'] || r['Mode'] || r['MODE'] || r['cr63f_mode'] || '');
      if (selectedMode !== 'ALL' && rowMode !== selectedMode) return;
      const type = String(r['Type'] || r['TYPE'] || r['cr63f_type'] || '').trim().toUpperCase();
      if (type) typeCounts.set(type, (typeCounts.get(type) || 0) + 1);
    });

    const currentVal = String(dom.typeFilter.value || 'ALL').trim().toUpperCase();
    let html = `<option value="ALL">All Types (${Array.from(typeCounts.values()).reduce((sum, n) => sum + n, 0).toLocaleString()})</option>`;
    Array.from(typeCounts.keys()).sort().forEach(t => {
      html += `<option value="${escapeHtml(t)}">${escapeHtml(t)} (${typeCounts.get(t).toLocaleString()})</option>`;
    });
    dom.typeFilter.innerHTML = html;
    if (typeCounts.has(currentVal)) {
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

  function populateModeFilter(records) {
    const selectedType = String(dom.typeFilter ? dom.typeFilter.value : 'ALL').trim().toUpperCase();
    const modeLabels = new Map();
    const modeCounts = new Map();
    const modeScopeRows = (records || []).filter(r => {
      const rowType = String(r['Type'] || r['TYPE'] || r['cr63f_type'] || '').trim().toUpperCase();
      return selectedType === 'ALL' || rowType === selectedType;
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

    const currentVal = normalizeModeFilterValue((dom.modeFilterPanel || dom.modeFilter)?.value || 'ALL');
    let html = `<option value="ALL">All Modes (${modeScopeRows.length.toLocaleString()})</option>`;
    const availableModes = Array.from(modeLabels.keys()).filter(key => (modeCounts.get(key) || 0) > 0);
    availableModes.forEach(key => {
      const label = modeLabels.get(key);
      html += `<option value="${escapeHtml(key)}">${escapeHtml(label)} (${modeCounts.get(key).toLocaleString()})</option>`;
    });
    const selectedMode = availableModes.includes(currentVal) ? currentVal : 'ALL';

    if (dom.modeFilter) {
      dom.modeFilter.innerHTML = html;
      dom.modeFilter.value = selectedMode;
    }
    if (dom.modeFilterPanel) {
      dom.modeFilterPanel.innerHTML = html;
      dom.modeFilterPanel.value = selectedMode;
    }

    const activeVal = (dom.modeFilterPanel || dom.modeFilter)?.value || 'ALL';
    if (dom.kpiModeActiveBadge) {
      dom.kpiModeActiveBadge.textContent = activeVal === 'ALL' ? 'ALL MODES' : activeVal;
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
              <button type="button" class="th-pin-btn ${isFrozen ? 'pinned' : ''}" data-pin-col="${col.key}" aria-label="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(col.label)}" title="${isFrozen ? 'Unfreeze' : 'Freeze'} ${escapeHtml(col.label)}">
                <svg viewBox="0 0 24 24" width="11" height="11" fill="currentColor" stroke="none">
                  <path d="M16 12V4h1V2H7v2h1v8l-2 2v2h5.2v6h1.6v-6H18v-2l-2-2z"/>
                </svg>
              </button>
              <button type="button" class="th-sort-wrap" data-sort-col="${col.key}" aria-label="Sort by ${escapeHtml(col.label)}" title="Click to sort by ${escapeHtml(col.label)} (Drag to rearrange)">
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
    const typeFilt = dom.typeFilter.value;
    const modeFilt = (dom.modeFilterPanel || dom.modeFilter)?.value || 'ALL';
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
      if (typeFilt !== 'ALL' && rType !== typeFilt) return false;

      const rMode = normalizeModeFilterValue(r['new_mode'] || r['Mode'] || r['MODE'] || r['cr63f_mode'] || '');
      if (modeFilt !== 'ALL' && rMode !== normalizeModeFilterValue(modeFilt)) return false;

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
        } else if (col.key === '__action') {
          td.className = `${td.className} col-action`.trim();
          td.innerHTML = `
            <button type="button" class="btn xs ghost danger rc-del-btn" onclick="window.deleteRecord('${id}')" aria-label="Delete Land Master record" title="Delete record from Dataverse">
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
                  input.value = pNum ? formatAmountWithCommas(pNum) : '';
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

    if (origVal === trimmed) {
      if (state.stagedEdits[rowId]) {
        delete state.stagedEdits[rowId][colKey];
        if (colDef.alt) delete state.stagedEdits[rowId][colDef.alt];
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

  function recalculateKPIs() {
    let sumTcp = 0;
    let sumDed = 0;
    let sumBal = 0;
    const targetRecords = state.filteredRecords !== undefined ? state.filteredRecords : state.records;

    targetRecords.forEach(r => {
      const id = r.__id;
      const staged = (state.stagedEdits && id && state.stagedEdits[id]) ? state.stagedEdits[id] : {};
      const tcp = parseAmount(staged['Total Contract Price'] ?? staged['cr63f_totalcontractprice'] ?? r['Total Contract Price'] ?? r['cr63f_totalcontractprice']);
      const ded = parseAmount(staged['Total Payment Deductible to TCP'] ?? staged['cr63f_totalpaymentdeductibletotcp'] ?? r['Total Payment Deductible to TCP'] ?? r['cr63f_totalpaymentdeductibletotcp']);
      const bal = parseAmount(staged['Balance'] ?? staged['cr63f_balance'] ?? r['Balance'] ?? r['cr63f_balance']);
      sumTcp += tcp;
      sumDed += ded;
      sumBal += bal;
    });

    if (dom.kpiRecordCount) dom.kpiRecordCount.textContent = targetRecords.length.toLocaleString();
    if (dom.kpiTotalTcp) dom.kpiTotalTcp.textContent = formatPrice(sumTcp) || '₱0.00';
    if (dom.kpiTotalDeductible) dom.kpiTotalDeductible.textContent = formatPrice(sumDed) || '₱0.00';
    if (dom.kpiTotalBalance) dom.kpiTotalBalance.textContent = formatPrice(sumBal) || '₱0.00';

    const activeMode = (dom.modeFilterPanel || dom.modeFilter)?.value || 'ALL';
    if (dom.kpiModeActiveBadge) {
      dom.kpiModeActiveBadge.textContent = activeMode === 'ALL' ? 'ALL MODES' : activeMode;
    }
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
          throw new Error(refRes.error || 'Failed to save reference updates.');
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
      message: 'Are you sure you want to delete this record from Dataverse?',
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
        const c = DEFAULT_COLUMNS.find(dc => dc.key === k);
        return (c && c.label) ? c.label : k;
      });
    const res = await window.pywebview.api.export_records_excel(toExport, 'LOA_Monitoring_Export.xlsx', cols);
    if (res.success) {
      showToast(`Exported to ${res.file_path}`, 'ok');
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
    if (dom.fAcquisitionSubLevel) {
      dom.fAcquisitionSubLevel.value = r['cr63f_acquisitionsublevel'] || r['Acquisition Sub-Level'] || '';
    }
    if (dom.fRccTerms) dom.fRccTerms.value = r['cr63f_rccterms'] || r['RCC Terms'] || '';
    
    const rccVal = r['cr63f_rccprice'] ?? r['RCC Price'] ?? '';
    if (dom.fRccPrice) dom.fRccPrice.value = rccVal !== '' ? formatAmount(rccVal) : '';

    if (dom.fSecuredDate) dom.fSecuredDate.value = formatDateMMDDYYYY(r['cr63f_secureddate'] || r['Secured Date']);
    if (dom.fSecuredEndorseDate) dom.fSecuredEndorseDate.value = formatDateMMDDYYYY(r['cr63f_securedendorsedate'] || r['Secured Endorse Date']);
    if (dom.fEndorsedDate) dom.fEndorsedDate.value = formatDateMMDDYYYY(r['cr63f_endorseddate'] || r['Endorsed Date']);

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
    if (dom.fAcquisitionSubLevel) dom.fAcquisitionSubLevel.value = '';
    if (dom.fRccTerms) dom.fRccTerms.value = '';
    if (dom.fRccPrice) dom.fRccPrice.value = '';
    if (dom.fSecuredDate) dom.fSecuredDate.value = '';
    if (dom.fSecuredEndorseDate) dom.fSecuredEndorseDate.value = '';
    
    // Autofill Endorsed Date with current local date (MM/DD/YYYY)
    const now = new Date();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const yyyy = now.getFullYear();
    if (dom.fEndorsedDate) dom.fEndorsedDate.value = `${mm}/${dd}/${yyyy}`;

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
          if (res.secured_date && dom.fSecuredDate) dom.fSecuredDate.value = formatDateMMDDYYYY(res.secured_date);
          if (res.secured_endorse_date && dom.fSecuredEndorseDate) dom.fSecuredEndorseDate.value = formatDateMMDDYYYY(res.secured_endorse_date);
          if (res.endorsed_date && dom.fEndorsedDate) dom.fEndorsedDate.value = formatDateMMDDYYYY(res.endorsed_date);
          updatePaymentMath();
        } else {
          if (clearOnNoMatch) clearAutofillFields();
        }

        // Live Conversion Status lookup
        let convMatched = false;
        let convType = '';
        let convStatus = '';
        if (areaIndex) {
          try {
            const convRes = await window.pywebview.api.lookup_conversion(areaIndex);
            if (convRes && convRes.matched && convRes.status) {
              convMatched = true;
              convType = convRes.match_type || 'Conversion';
              convStatus = convRes.status;
              if (dom.fConversionStatus) {
                dom.fConversionStatus.value = convStatus;
              }
            }
          } catch (eConv) {
            console.error('Conversion lookup error:', eConv);
          }
        }

        if (dom.badgeMatchStatus) {
          if (res.matched && convMatched) {
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

    if (!recordData['Type']) {
      showToast('TYPE is required.', 'err');
      return;
    }

    const isNew = !recordId;
    if (isNew && !recordData['cr63f_acquisitionsublevel']) {
      showToast('Acquisition Sub-Level is required.', 'err');
      return;
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
              <button type="button" class="th-sort-wrap" data-fin-sort="${escapeHtml(h)}" aria-label="Sort by ${escapeHtml(h)}" title="Click to sort, drag to reorder">
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
          <div class="th-resizer" data-fin-resize="${escapeHtml(h)}" title="Drag column width (Double-click to reset)"></div>
          <div class="th-bottom-resizer" title="Drag to adjust header height"></div>
        </th>
      `;
    });
    thead += `<th class="col-action" style="width: 55px; min-width: 55px; max-width: 55px; text-align: center; position: sticky !important; top: 0 !important; z-index: 20 !important; background: var(--subtle) !important;">Actions</th>`;
    thead += '</tr>';
    dom.financePreviewThead.innerHTML = thead;

    attachFinanceHeaderEvents();

    const displayRows = financeState.filteredRows;

    if (!displayRows.length) {
      dom.financePreviewTbody.innerHTML = '';
      const colSpan = Math.max(1, cols.length + 2);
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
      const colSpan = Math.max(1, cols.length + 2);
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
        <button type="button" class="btn xs ghost danger rc-del-btn" onclick="window.deleteFinanceRecord('${escapeHtml(String(idOrIdx))}')" aria-label="Delete Finance record" title="Delete record from Finance Data">
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
      if (res && res.success) {
        closeConfirmModal(false);
        showToast(res.message || `Successfully saved ${res.saved_count || rowIds.length} finance record(s) to Dataverse.`, 'ok');

        financeState.baselineRows = JSON.parse(JSON.stringify(financeState.rows));
        financeState.stagedEdits = {};
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
    const res = await window.pywebview.api.export_records_excel(toExport, 'Finance_Data_Export.xlsx', cols);
    if (res.success) {
      showToast(`Exported to ${res.file_path}`, 'ok');
    }
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

    const rowCount = financeState.rows.length.toLocaleString();
    const confirmed = await showConfirmModal({
      title: 'Sync Finance Data to Dataverse',
      subtitle: 'Table: cr63f_financereference',
      message: `Upload and sync ${rowCount} finance record(s) to Dataverse?`,
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
      if (res.success) {
        if (dom.confirmModalProgressBar) dom.confirmModalProgressBar.style.width = '100%';
        if (dom.confirmModalProgressPercent) dom.confirmModalProgressPercent.textContent = '100%';
        if (dom.confirmModalProgressStatus) {
          dom.confirmModalProgressStatus.textContent = `Completed! ${res.synced_count.toLocaleString()} records synced successfully.`;
        }
        showToast(res.message || `Successfully synced ${res.synced_count} finance records.`, 'ok');
        if (dom.lblActiveFinanceSource) {
          dom.lblActiveFinanceSource.textContent = `Dataverse: cr63f_financereference (${res.synced_count} records)`;
        }
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
              <button type="button" class="th-sort-wrap" data-conv-sort="${escapeHtml(h)}" aria-label="Sort by ${escapeHtml(h)}" title="Click to sort, drag to reorder">
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
          <div class="th-resizer" data-conv-resize="${escapeHtml(h)}" title="Drag column width (Double-click to reset)"></div>
          <div class="th-bottom-resizer" title="Drag to adjust header height"></div>
        </th>
      `;
    });
    thead += `<th class="col-action" style="width: 55px; min-width: 55px; max-width: 55px; text-align: center; position: sticky !important; top: 0 !important; z-index: 20 !important; background: var(--subtle) !important;">Actions</th>`;
    thead += '</tr>';
    dom.conversionPreviewThead.innerHTML = thead;

    attachConversionHeaderEvents();

    const displayRows = conversionState.filteredRows;

    if (!displayRows.length) {
      dom.conversionPreviewTbody.innerHTML = '';
      const colSpan = Math.max(1, cols.length + 2);
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
      const colSpan = Math.max(1, cols.length + 2);
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
              handleConversionCellEdit(idOrIdx, h, cleanStr, input, td);
            }
          } else if (isDate) {
            const masked = maskDateString(input.value);
            input.value = masked;
            if (masked !== before) {
              handleConversionCellEdit(idOrIdx, h, masked, input, td);
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

      // Actions column (Red trash delete button)
      const actTd = document.createElement('td');
      actTd.className = 'col-action';
      actTd.style.cssText = 'width: 55px; min-width: 55px; max-width: 55px; text-align: center;';
      actTd.innerHTML = `
        <button type="button" class="btn xs ghost danger rc-del-btn" onclick="window.deleteConversionRecord('${escapeHtml(String(idOrIdx))}')" aria-label="Delete Conversion record" title="Delete record from Conversion Data">
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
      if (res && res.success) {
        closeConfirmModal(false);
        showToast(res.message || `Successfully saved ${res.saved_count || rowIds.length} conversion record(s) to Dataverse.`, 'ok');

        conversionState.baselineRows = JSON.parse(JSON.stringify(conversionState.rows));
        conversionState.stagedEdits = {};
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
    const res = await window.pywebview.api.export_records_excel(toExport, 'Conversion_Data_Export.xlsx', cols);
    if (res.success) {
      showToast(`Exported to ${res.file_path}`, 'ok');
    }
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

    const rowCount = conversionState.rows.length.toLocaleString();
    const tableName = state.activeConversionTable || 'cr63f_landconversionreference';
    const confirmed = await showConfirmModal({
      title: 'Sync Conversion Data to Dataverse',
      subtitle: `Table: ${tableName}`,
      message: `Upload and sync ${rowCount} conversion record(s) to Dataverse?`,
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
      if (res.success) {
        if (dom.confirmModalProgressBar) dom.confirmModalProgressBar.style.width = '100%';
        if (dom.confirmModalProgressPercent) dom.confirmModalProgressPercent.textContent = '100%';
        if (dom.confirmModalProgressStatus) {
          dom.confirmModalProgressStatus.textContent = `Completed! ${res.synced_count.toLocaleString()} records synced successfully.`;
        }
        showToast(res.message || `Successfully synced ${res.synced_count} conversion records.`, 'ok');
        if (dom.lblActiveConversionSource) {
          dom.lblActiveConversionSource.textContent = `Dataverse: ${tableName} (${res.synced_count} records)`;
        }
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
        if (dom.lblTopConversion) dom.lblTopConversion.textContent = state.conversionLoadedCount > 0 ? `Conversion Data: ${state.conversionLoadedCount.toLocaleString()} rows` : 'Conversion Data: None';
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
    if (action.includes('REFERENCE')) return 'REFERENCE_UPDATE';
    if (action.includes('FINANCE')) return 'FINANCE_UPDATE';
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
      REVERT: 'Revert',
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
    const actFilter = dom.actionFilterLogs ? dom.actionFilterLogs.value.toUpperCase() : 'ALL';
    const wwFilter = dom.wwFilterLogs ? dom.wwFilterLogs.value.toUpperCase() : 'ALL';
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
            <div class="th-bottom-resizer" title="Drag to adjust header height"></div>
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
              <button type="button" class="th-sort-wrap" data-log-sort="${col.key}" aria-label="Sort by ${escapeHtml(col.label)}" title="Click to sort by ${escapeHtml(col.label)} (Drag to rearrange)">
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
          <div class="th-resizer" data-log-resize="${col.key}" title="Drag column width (Double-click to reset)"></div>
          <div class="th-bottom-resizer" title="Drag to adjust header height"></div>
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
    const targetId = String(log.record_id || log.cr63f_loamonitoringid || '').trim();

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
      const oIdx = String(other.index_no || other['Index No'] || other['INDEX NO'] || other['Corridor Index'] || other.index_no || '').trim().toUpperCase();
      const oId = String(other.record_id || other.cr63f_loamonitoringid || '').trim();

      let isSameRecord = false;
      if (targetId && oId && targetId === oId) isSameRecord = true;
      else if (targetCtrl && oCtrl && targetCtrl === oCtrl) isSameRecord = true;
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

        if (targetId && rId && targetId === rId) return true;
        if (targetCtrl && rCtrl && targetCtrl === rCtrl) return true;
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
        } else if (col.key === '__action') {
          td.className = `col-action ${isFrozen ? 'col-frozen' : ''} ${isLastFrz ? 'col-frozen-last' : ''}`.trim();
          const recId = r.__id || r.id || r.cr63f_loamonitoringlogid || idx;
          const canRevert = isLogRevertible(r);
          const revertBtn = canRevert
            ? `<button type="button" class="rc-revert-btn" aria-label="Revert field in Land Master" title="Revert field in Land Master DB" onclick="window.revertLogRecord('${escapeHtml(recId)}')"><svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"></polyline><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"></path></svg></button>`
            : '<span class="action-btn-placeholder">—</span>';
          td.innerHTML = revertBtn;
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
          td.innerHTML = `<span class="log-diff-field">${escapeHtml(fn)}</span>`;
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
          td.textContent = String(lt || '-');
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
            td.textContent = String(val);
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
    const res = await window.pywebview.api.export_logs_excel(toExport, 'Entry_Logs_Export.xlsx', cols);
    if (res.success) showToast(`Exported to ${res.file_path}`, 'ok');
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
        const tcp = parseAmount(r['cr63f_totalcontractprice'] ?? r['Total Contract Price']);
        const ded = parseAmount(r['cr63f_totalpaymentdeductibletotcp'] ?? r['Payment Deductible'] ?? r['Total Payment Deductible to TCP']);
        const bal = Math.max(0, tcp - ded);
        const pct = tcp > 0 ? (ded / tcp * 100).toFixed(2) + '%' : '0.00%';
        r['cr63f_balance'] = bal;
        r['Balance'] = '₱' + bal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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

  // ==================== DASHBOARD VIEW ====================
  let selectedDashboardMode = null;
  let dashboardActivitySearch = '';
  let dashAuditSearchQuery = '';
  let dashAuditFilterAction = 'ALL';
  let dashAuditCurrentPage = 1;
  let dashAuditPageSize = 10;
  let dashAuditFlattenedData = [];

  function setupDashboardView() {
    if (dom.dashActivitySearchInput) {
      dom.dashActivitySearchInput.addEventListener('input', () => {
        dashboardActivitySearch = dom.dashActivitySearchInput.value.toLowerCase().trim();
        dashAuditSearchQuery = dashboardActivitySearch;
        dashAuditCurrentPage = 1;
        renderDashboardAuditTable();
      });
    }
    if (dom.dashActivityActionFilter) {
      dom.dashActivityActionFilter.addEventListener('change', () => {
        dashAuditFilterAction = dom.dashActivityActionFilter.value;
        dashAuditCurrentPage = 1;
        renderDashboardAuditTable();
      });
    }
    if (dom.dashWorkWeekSelect) {
      dom.dashWorkWeekSelect.addEventListener('change', () => loadDashboardAnalytics(dom.dashWorkWeekSelect.value, selectedDashboardMode));
    }
    if (dom.btnRefreshDashboard) {
      dom.btnRefreshDashboard.addEventListener('click', () => loadDashboardAnalytics(dom.dashWorkWeekSelect ? dom.dashWorkWeekSelect.value : 'ALL', selectedDashboardMode));
    }
    if (dom.btnDashClearMode) {
      dom.btnDashClearMode.addEventListener('click', () => {
        selectedDashboardMode = null;
        loadDashboardAnalytics(dom.dashWorkWeekSelect ? dom.dashWorkWeekSelect.value : 'ALL', null);
      });
    }
  }

  async function loadDashboardAnalytics(workWeek = 'ALL', mode = null) {
    try {
      if (dom.btnRefreshDashboard) {
        dom.btnRefreshDashboard.disabled = true;
        dom.btnRefreshDashboard.textContent = 'Loading…';
      }
      if (dom.viewDashboard) dom.viewDashboard.setAttribute('aria-busy', 'true');
      selectedDashboardMode = mode || null;
      const res = await window.pywebview.api.get_dashboard_analytics(workWeek, selectedDashboardMode);
      if (res.success) renderDashboard(res);
      else showToast(res.error || 'Failed to load dashboard analytics.', 'err');
    } catch (err) {
      showToast('Error loading dashboard: ' + err, 'err');
    } finally {
      if (dom.btnRefreshDashboard) {
        dom.btnRefreshDashboard.disabled = false;
        dom.btnRefreshDashboard.textContent = 'Refresh overview';
      }
      if (dom.viewDashboard) dom.viewDashboard.removeAttribute('aria-busy');
    }
  }

  function renderDashboard(data) {
    if (!data) return;
    window.__lastDashboardData = data;
    if (dom.dashWorkWeekSelect && data.available_work_weeks) {
      const cur = data.selected_work_week || dom.dashWorkWeekSelect.value || 'ALL';
      let optHtml = '<option value="ALL">All Time / All Weeks</option>';
      data.available_work_weeks.forEach(w => {
        optHtml += `<option value="${escapeHtml(w)}" ${w === cur ? 'selected' : ''}>${escapeHtml(w)}</option>`;
      });
      dom.dashWorkWeekSelect.innerHTML = optHtml;
    }

    if (dom.dashKpiTotalRecords) dom.dashKpiTotalRecords.textContent = (data.total_records || 0).toLocaleString();
    if (dom.dashKpiActiveModeCount) {
      if (data.selected_mode) {
        dom.dashKpiActiveModeCount.innerHTML = `Filtered: ${escapeHtml(data.selected_mode)} <span id="dashKpiClearModeLink" title="Clear mode filter" style="cursor: pointer; text-decoration: underline; margin-left: 4px; opacity: 0.85;">✕</span>`;
        const clearLink = dom.dashKpiActiveModeCount.querySelector('#dashKpiClearModeLink');
        if (clearLink) {
          clearLink.addEventListener('click', (e) => {
            e.stopPropagation();
            selectedDashboardMode = null;
            loadDashboardAnalytics(dom.dashWorkWeekSelect ? dom.dashWorkWeekSelect.value : 'ALL', null);
          });
        }
      } else {
        dom.dashKpiActiveModeCount.textContent = 'All modes active';
      }
    }
    if (dom.dashKpiWeeklyChanges) dom.dashKpiWeeklyChanges.textContent = (data.weekly_changes_count || 0).toLocaleString();
    if (dom.dashKpiTotalTcp) dom.dashKpiTotalTcp.textContent = data.total_tcp_formatted || '₱0.00';
    if (dom.dashKpiTotalDeductible) dom.dashKpiTotalDeductible.textContent = data.total_deductible_formatted || '₱0.00';
    if (dom.dashKpiTotalBalance) dom.dashKpiTotalBalance.textContent = data.total_balance_formatted || '₱0.00';
    if (dom.dashKpiAvgPaid) dom.dashKpiAvgPaid.textContent = `${data.overall_paid_percent || '0.00%'} overall settled`;

    // Mode Filter Banner & Clear Button (banner hidden to match clean reference layout)
    const hasModeFilter = Boolean(data.selected_mode && String(data.selected_mode).toUpperCase() !== 'ALL');
    if (dom.btnDashClearMode) {
      if (hasModeFilter) dom.btnDashClearMode.classList.remove('hidden');
      else dom.btnDashClearMode.classList.add('hidden');
    }
    if (dom.dashModeFilterBanner) {
      dom.dashModeFilterBanner.classList.add('hidden');
    }

    // Widget 1: Mode Distribution Breakdown (Card 1: 1-column vertical clickable mode rows)
    if (dom.dashModeBarsContainer) {
      const modes = data.modes_distribution || [];
      if (dom.dashModeTotalLots) {
        dom.dashModeTotalLots.textContent = `${(data.total_records || 0).toLocaleString()} total lots`;
        if (hasModeFilter) {
          dom.dashModeTotalLots.style.cursor = 'pointer';
          dom.dashModeTotalLots.title = 'Click to clear mode filter';
          dom.dashModeTotalLots.onclick = () => {
            selectedDashboardMode = null;
            loadDashboardAnalytics(dom.dashWorkWeekSelect ? dom.dashWorkWeekSelect.value : 'ALL', null);
          };
        } else {
          dom.dashModeTotalLots.style.cursor = 'default';
          dom.dashModeTotalLots.title = '';
          dom.dashModeTotalLots.onclick = null;
        }
      }

      if (!modes.length) {
        dom.dashModeBarsContainer.innerHTML = '<p style="color: var(--ink3); font-size: 12px; font-style: italic; padding: 12px 0;">No mode data available.</p>';
      } else {
        const renderModeRow = (m) => {
          if (!m) return '';
          const isSelected = Boolean(data.selected_mode && data.selected_mode.toUpperCase() === m.mode.toUpperCase());
          return `
            <button type="button" class="mode-bar-row clickable ${isSelected ? 'active' : ''}" data-mode-name="${escapeHtml(m.mode)}" aria-pressed="${isSelected ? 'true' : 'false'}" title="${isSelected ? 'Click to clear filter (' + escapeHtml(m.mode) + ')' : 'Filter overview by ' + escapeHtml(m.mode)}">
              <div class="mode-bar-meta">
                <span class="mode-bar-label">${escapeHtml(m.mode)}</span>
                <span class="mode-bar-count" title="${m.count.toLocaleString()} lots; ${escapeHtml(m.percent)} of accounts">${m.count.toLocaleString()} · ${escapeHtml(m.percent)}</span>
              </div>
              <div class="mode-bar-track">
                <div class="mode-bar-fill" style="width: ${Math.min(100, m.percent_val)}%;"></div>
              </div>
            </button>
          `;
        };

        const rowsHtml = modes.slice(0, 8).map(renderModeRow).join('');
        dom.dashModeBarsContainer.innerHTML = rowsHtml;

        // Attach click events on mode bars to filter or toggle off
        dom.dashModeBarsContainer.querySelectorAll('.mode-bar-row.clickable').forEach(row => {
          row.addEventListener('click', () => {
            const clickedMode = row.dataset.modeName;
            if (data.selected_mode && data.selected_mode.toUpperCase() === clickedMode.toUpperCase()) {
              selectedDashboardMode = null;
            } else {
              selectedDashboardMode = clickedMode;
            }
            loadDashboardAnalytics(dom.dashWorkWeekSelect ? dom.dashWorkWeekSelect.value : 'ALL', selectedDashboardMode);
          });
        });
      }
    }

    // Widget 2: Financial Progress & Settlement (Image 2)
    if (dom.dashFinancePaidRatio) dom.dashFinancePaidRatio.textContent = `${data.overall_paid_percent || '0.00%'} Paid`;
    if (dom.dashFinanceProgressBar) {
      const pNum = Math.min(100, Math.max(0, data.overall_paid_percent_num || 0));
      dom.dashFinanceProgressBar.style.width = `${pNum}%`;
    }
    if (dom.dashFinBoxTcp) dom.dashFinBoxTcp.textContent = data.total_tcp_formatted || '₱0.00';
    if (dom.dashFinBoxDeductible) dom.dashFinBoxDeductible.textContent = data.total_deductible_formatted || '₱0.00';
    if (dom.dashFinBoxBalance) dom.dashFinBoxBalance.textContent = data.total_balance_formatted || '₱0.00';

    if (dom.dashFinanceBreakdownContainer) {
      const fb = data.financial_breakdown || {};
      dom.dashFinanceBreakdownContainer.innerHTML = `
        <div class="fin-settle-grid">
          <div class="fin-settle-item">
            <div class="fin-settle-header">
              <span class="fin-settle-label"><span class="status-dot green"></span>Settled · 100%</span>
              <span class="fin-settle-val">${(fb.fully_paid_count || 0).toLocaleString()} lots <span class="fin-settle-subpct">${escapeHtml(fb.fully_paid_percent || '0.0%')}</span></span>
            </div>
            <div class="fin-settle-track">
              <div class="fin-settle-fill green" style="width: ${Math.min(100, fb.fully_paid_percent_num || 0)}%;"></div>
            </div>
          </div>
          <div class="fin-settle-item">
            <div class="fin-settle-header">
              <span class="fin-settle-label"><span class="status-dot amber"></span>Partial · &gt;0%</span>
              <span class="fin-settle-val">${(fb.partially_paid_count || 0).toLocaleString()} lots <span class="fin-settle-subpct">${escapeHtml(fb.partially_paid_percent || '0.0%')}</span></span>
            </div>
            <div class="fin-settle-track">
              <div class="fin-settle-fill amber" style="width: ${Math.min(100, fb.partially_paid_percent_num || 0)}%;"></div>
            </div>
          </div>
          <div class="fin-settle-item">
            <div class="fin-settle-header">
              <span class="fin-settle-label"><span class="status-dot slate"></span>Unpaid · 0%</span>
              <span class="fin-settle-val">${(fb.zero_paid_count || 0).toLocaleString()} lots <span class="fin-settle-subpct">${escapeHtml(fb.zero_paid_percent || '0.0%')}</span></span>
            </div>
            <div class="fin-settle-track">
              <div class="fin-settle-fill slate" style="width: ${Math.min(100, fb.zero_paid_percent_num || 0)}%;"></div>
            </div>
          </div>
        </div>
      `;
    }

    if (dom.dashFinanceFooterStats) {
      const fb = data.financial_breakdown || {};
      dom.dashFinanceFooterStats.innerHTML = `
        <div class="dash-tile">
          <div class="dash-tile-icon-box slate">
            <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="20" x2="18" y2="10"></line><line x1="12" y1="20" x2="12" y2="4"></line><line x1="6" y1="20" x2="6" y2="14"></line></svg>
          </div>
          <div class="dash-tile-info">
            <span class="dash-tile-label">AVG TCP / LOT</span>
            <span class="dash-tile-val">${escapeHtml(fb.avg_tcp_formatted || '₱0.00')}</span>
          </div>
        </div>
        <div class="dash-tile">
          <div class="dash-tile-icon-box green">
            <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="6" width="20" height="12" rx="2"></rect><circle cx="12" cy="12" r="2"></circle><path d="M6 12h.01M18 12h.01"></path></svg>
          </div>
          <div class="dash-tile-info">
            <span class="dash-tile-label">AVG PAID / LOT</span>
            <span class="dash-tile-val text-green">${escapeHtml(fb.avg_paid_formatted || '₱0.00')}</span>
          </div>
        </div>
      `;
    }

    // Widget 3: Conversion Status & Progress (Image 2)
    const conv = data.conversion_metrics || {};
    if (dom.dashConvSecuredRatio) {
      dom.dashConvSecuredRatio.textContent = `${conv.secured_percent || '0.0%'} Secured`;
    }
    if (dom.dashConvProgressBar) {
      const convNum = Math.min(100, Math.max(0, conv.secured_percent_num || 0));
      dom.dashConvProgressBar.style.width = `${convNum}%`;
    }
    if (dom.dashConvSecuredCount) {
      dom.dashConvSecuredCount.textContent = `${(conv.secured_count || 0).toLocaleString()} lots`;
    }
    if (dom.dashConvInProgressCount) {
      dom.dashConvInProgressCount.textContent = `${(conv.in_progress_count || 0).toLocaleString()} lots`;
    }
    if (dom.dashConvTotalArea) {
      dom.dashConvTotalArea.textContent = conv.total_applied_area_formatted || '0.00 Has.';
    }
    if (dom.dashConvBarsContainer) {
      const breakdown = conv.breakdown || [];
      if (!breakdown.length) {
        dom.dashConvBarsContainer.innerHTML = '<p style="color: var(--ink3); font-size: 12px; font-style: italic; padding: 12px 0;">No conversion status breakdown available.</p>';
      } else {
        const renderConvItem = (item) => {
          if (!item) return '';
          const stUpper = item.status.toUpperCase();
          let fillClass = 'blue';
          if (stUpper.includes('SECURED') || stUpper.includes('FINALITY') || stUpper.includes('APPROVED') || stUpper.includes('EXEMPT')) {
            fillClass = 'green';
          } else if (stUpper.includes('DONE') || stUpper.includes('INSPECTION')) {
            fillClass = 'cyan';
          } else if (stUpper.includes('APPLICATION') || stUpper.includes('PROGRESS')) {
            fillClass = 'amber';
          } else if (stUpper.includes('NOT YET') || stUpper.includes('PENDING') || stUpper.includes('UNSET')) {
            fillClass = 'slate';
          }
          return `
            <div class="conv-status-item" title="${escapeHtml(item.status)}: ${item.count.toLocaleString()} lots (${escapeHtml(item.percent)})">
              <div class="conv-status-meta">
                <span class="conv-status-name">${escapeHtml(item.status)}</span>
                <span class="conv-status-stat">${item.count.toLocaleString()} · ${escapeHtml(item.percent)}</span>
              </div>
              <div class="conv-status-track">
                <div class="conv-status-fill ${fillClass}" style="width: ${Math.min(100, item.percent_val)}%;"></div>
              </div>
            </div>
          `;
        };

        const items = breakdown.slice(0, 6);
        const rowsHtml = items.map(renderConvItem).join('');
        dom.dashConvBarsContainer.innerHTML = rowsHtml;
      }
    }

    // Work Week comparison pill
    if (dom.dashWwComparisonPill) {
      if (data.comparison_previous_week && data.selected_work_week && data.selected_work_week !== 'ALL') {
        dom.dashWwComparisonPill.textContent = `${data.comparison_previous_week} → ${data.selected_work_week}`;
        dom.dashWwComparisonPill.style.display = 'inline-flex';
      } else if (data.selected_work_week && data.selected_work_week !== 'ALL') {
        dom.dashWwComparisonPill.textContent = data.selected_work_week;
        dom.dashWwComparisonPill.style.display = 'inline-flex';
      } else {
        dom.dashWwComparisonPill.textContent = 'All weeks';
        dom.dashWwComparisonPill.style.display = 'inline-flex';
      }
    }

    // Flatten recent activity changes into audit table dataset
    dashAuditFlattenedData = [];
    const recentActivity = data.recent_activity || [];
    recentActivity.forEach(account => {
      const title = account.area_index || account.title_no || account.control_number || 'Account';
      const subTitle = account.control_number ? `Control: ${account.control_number}` : (account.title_no || '');
      const changes = account.changes || [];
      changes.forEach(c => {
        let dateDisp = c.log_time ? formatAuditDate(c.log_time) : (account.current_week || '—');
        dashAuditFlattenedData.push({
          date: dateDisp,
          rawDate: c.log_time || '',
          account: title,
          subAccount: subTitle,
          fieldName: c.field_name || 'Field',
          actionType: c.action_type || 'UPDATED',
          oldValue: c.old_value,
          newValue: c.new_value,
          contractPrice: c.contract_price || '—',
          acquisitionSublevel: c.acquisition_sublevel || '—',
          updatedBy: c.updated_by || 'j.inguito',
          rawAccount: account
        });
      });
    });

    // Fallback to logsState.logs if recent_activity has 0 items
    if (dashAuditFlattenedData.length === 0 && logsState.logs && logsState.logs.length > 0) {
      const selectedWw = data.selected_work_week || 'ALL';
      const matchingLogs = logsState.logs.filter(l => {
        if (!selectedWw || selectedWw === 'ALL') return true;
        const lWw = String(l.work_week || l['Work Week'] || '').trim().toUpperCase();
        return lWw === selectedWw.toUpperCase();
      });
      matchingLogs.slice(0, 100).forEach(log => {
        const title = log.area_index || log['Area-Index'] || log['AREA-INDEX'] || log.title_no || log['Title No'] || log.cr63f_control1 || 'Account';
        const subTitle = log.cr63f_control1 ? `Control: ${log.cr63f_control1}` : '';
        const fieldName = log.field_name || log['Field Changed'] || log.cr63f_fieldchanged || 'Field';
        const dateDisp = log.log_time ? formatAuditDate(log.log_time) : (log.work_week || '—');
        dashAuditFlattenedData.push({
          date: dateDisp,
          rawDate: log.log_time || '',
          account: title,
          subAccount: subTitle,
          fieldName: fieldName,
          actionType: log.action_type || log['Action Type'] || 'UPDATED',
          oldValue: log.old_value !== undefined ? log.old_value : log['Previous Value'],
          newValue: log.new_value !== undefined ? log.new_value : log['New Value'],
          contractPrice: log.cr63f_totalcontractprice || log['Total Contract Price'] ? formatPrice(log.cr63f_totalcontractprice || log['Total Contract Price']) : '—',
          acquisitionSublevel: log.cr63f_acquisitionsublevel || log['Acquisition Sub-Level'] || '—',
          updatedBy: log.author || log.createdby || log.User || 'j.inguito',
          rawAccount: log
        });
      });
    }

    renderDashboardAuditTable();
  }

  function formatAuditDate(dateStr) {
    if (!dateStr) return '—';
    try {
      const d = new Date(dateStr);
      if (!isNaN(d.getTime())) {
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        const y = d.getFullYear();
        let hours = d.getHours();
        const mins = String(d.getMinutes()).padStart(2, '0');
        const ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12 || 12;
        return `${m}/${day}/${y} ${hours}:${mins} ${ampm}`;
      }
    } catch (_) {}
    return String(dateStr);
  }

  function getAuditBadge(actionType) {
    const act = String(actionType || '').toUpperCase();
    let badgeClass = 'blue';
    let badgeText = 'Updated';
    if (act.includes('REVERT')) {
      badgeClass = 'amber';
      badgeText = 'Reverted';
    } else if (act.includes('NEW') || act.includes('ADD')) {
      badgeClass = 'green';
      badgeText = 'Added';
    } else if (act.includes('REFERENCE') || act.includes('FINANCE') || act.includes('SYNC')) {
      badgeClass = 'purple';
      badgeText = 'Reference';
    } else if (act.includes('CHANGE')) {
      badgeClass = 'blue';
      badgeText = 'Changed';
    } else {
      badgeClass = 'blue';
      badgeText = act ? act.charAt(0) + act.slice(1).toLowerCase() : 'Updated';
    }
    return `<span class="badge-audit ${badgeClass}">${escapeHtml(badgeText)}</span>`;
  }

  function renderDashboardAuditTable() {
    if (!dom.dashAuditTableBody) return;

    const q = dashAuditSearchQuery;
    const actFilter = dashAuditFilterAction;

    const filtered = dashAuditFlattenedData.filter(item => {
      if (actFilter !== 'ALL') {
        const act = item.actionType.toUpperCase();
        if (actFilter === 'UPDATED' && !act.includes('UPDATE')) return false;
        if (actFilter === 'CHANGED' && !act.includes('CHANGE') && !act.includes('UPDATE')) return false;
        if (actFilter === 'ADDED' && !act.includes('ADD') && !act.includes('NEW')) return false;
        if (actFilter === 'REVERTED' && !act.includes('REVERT')) return false;
      }

      if (q) {
        const searchTarget = [
          item.date,
          item.account,
          item.subAccount,
          item.fieldName,
          item.actionType,
          item.oldValue,
          item.newValue,
          item.contractPrice,
          item.acquisitionSublevel,
          item.updatedBy
        ].join(' ').toLowerCase();
        if (!searchTarget.includes(q)) return false;
      }

      return true;
    });

    const totalCount = filtered.length;
    const totalPages = Math.max(1, Math.ceil(totalCount / dashAuditPageSize));
    if (dashAuditCurrentPage > totalPages) dashAuditCurrentPage = totalPages;
    if (dashAuditCurrentPage < 1) dashAuditCurrentPage = 1;

    const startIdx = totalCount > 0 ? (dashAuditCurrentPage - 1) * dashAuditPageSize : 0;
    const endIdx = Math.min(startIdx + dashAuditPageSize, totalCount);
    const pageItems = filtered.slice(startIdx, endIdx);

    if (dom.dashAuditCountLabel) {
      if (totalCount === 0) {
        dom.dashAuditCountLabel.textContent = 'Showing 0 changes';
      } else {
        dom.dashAuditCountLabel.textContent = `Showing ${(startIdx + 1).toLocaleString()}–${endIdx.toLocaleString()} of ${totalCount.toLocaleString()} changes`;
      }
    }

    if (pageItems.length === 0) {
      dom.dashAuditTableBody.innerHTML = `
        <tr>
          <td colspan="9" style="text-align: center; padding: 40px 16px; color: var(--ink3); font-style: italic;">
            ${q || actFilter !== 'ALL' ? 'No matching weekly changes. Try adjusting your filters.' : 'No weekly changes between these work weeks.'}
          </td>
        </tr>
      `;
    } else {
      let tbodyHtml = '';
      pageItems.forEach(item => {
        const oldDisp = formatLogValue(item.fieldName, item.oldValue);
        const newDisp = formatLogValue(item.fieldName, item.newValue);
        const badgeHtml = getAuditBadge(item.actionType);

        tbodyHtml += `
          <tr>
            <td>
              <span style="font-weight: 500; font-family: var(--mono); font-size: 11.5px; color: var(--ink2);">${escapeHtml(item.date)}</span>
            </td>
            <td>
              <div style="display: flex; flex-direction: column; gap: 2px;">
                <span style="font-weight: 600; color: var(--ink); cursor: pointer;" title="Filter Land Master records by ${escapeHtml(item.account)}" onclick="window.jumpToDashboardRecord('${escapeHtml(item.account)}')">
                  ${escapeHtml(item.account)}
                </span>
                <span style="font-size: 11px; color: var(--ink3);">${escapeHtml(item.fieldName)}</span>
              </div>
            </td>
            <td>${badgeHtml}</td>
            <td>
              <span style="color: var(--ink3); font-family: var(--mono); font-size: 11.5px; max-width: 140px; display: inline-block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${escapeHtml(String(oldDisp))}">
                ${escapeHtml(String(oldDisp || '—'))}
              </span>
            </td>
            <td>
              <span style="color: var(--ok, #10b981); font-weight: 600; font-family: var(--mono); font-size: 11.5px; max-width: 140px; display: inline-block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${escapeHtml(String(newDisp))}">
                ${escapeHtml(String(newDisp || '—'))}
              </span>
            </td>
            <td>
              <span style="font-family: var(--mono); font-size: 11.5px;">${escapeHtml(item.contractPrice)}</span>
            </td>
            <td>
              <span style="font-size: 11.5px; color: var(--ink2); max-width: 180px; display: inline-block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${escapeHtml(item.acquisitionSublevel)}">
                ${escapeHtml(item.acquisitionSublevel)}
              </span>
            </td>
            <td>
              <span style="font-size: 11.5px; color: var(--ink3);">${escapeHtml(item.updatedBy)}</span>
            </td>
            <td style="text-align: right;">
              <button type="button" class="btn-row-action" title="View ${escapeHtml(item.account)} in Land Master DB" onclick="window.jumpToDashboardRecord('${escapeHtml(item.account)}')">···</button>
            </td>
          </tr>
        `;
      });
      dom.dashAuditTableBody.innerHTML = tbodyHtml;
    }

    if (dom.dashAuditPagination) {
      renderDashboardPaginationControls(totalPages);
    }
  }

  function renderDashboardPaginationControls(totalPages) {
    if (!dom.dashAuditPagination) return;
    let pagHtml = '';

    pagHtml += `
      <button type="button" class="dash-page-btn" ${dashAuditCurrentPage <= 1 ? 'disabled' : ''} id="btnDashAuditPrev" aria-label="Previous page">
        <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"></polyline></svg>
      </button>
    `;

    let startPage = 1;
    let endPage = totalPages;
    if (totalPages > 5) {
      if (dashAuditCurrentPage <= 3) {
        startPage = 1;
        endPage = 5;
      } else if (dashAuditCurrentPage + 2 >= totalPages) {
        startPage = totalPages - 4;
        endPage = totalPages;
      } else {
        startPage = dashAuditCurrentPage - 2;
        endPage = dashAuditCurrentPage + 2;
      }
    }

    for (let p = startPage; p <= endPage; p++) {
      pagHtml += `
        <button type="button" class="dash-page-btn ${p === dashAuditCurrentPage ? 'active' : ''}" data-page="${p}">${p}</button>
      `;
    }

    pagHtml += `
      <button type="button" class="dash-page-btn" ${dashAuditCurrentPage >= totalPages ? 'disabled' : ''} id="btnDashAuditNext" aria-label="Next page">
        <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"></polyline></svg>
      </button>
    `;

    pagHtml += `
      <select class="dash-page-size-select" id="selDashAuditPageSize" aria-label="Rows per page">
        <option value="10" ${dashAuditPageSize === 10 ? 'selected' : ''}>10 / page</option>
        <option value="25" ${dashAuditPageSize === 25 ? 'selected' : ''}>25 / page</option>
        <option value="50" ${dashAuditPageSize === 50 ? 'selected' : ''}>50 / page</option>
      </select>
    `;

    dom.dashAuditPagination.innerHTML = pagHtml;

    const btnPrev = dom.dashAuditPagination.querySelector('#btnDashAuditPrev');
    if (btnPrev && dashAuditCurrentPage > 1) {
      btnPrev.addEventListener('click', () => {
        dashAuditCurrentPage--;
        renderDashboardAuditTable();
      });
    }

    const btnNext = dom.dashAuditPagination.querySelector('#btnDashAuditNext');
    if (btnNext && dashAuditCurrentPage < totalPages) {
      btnNext.addEventListener('click', () => {
        dashAuditCurrentPage++;
        renderDashboardAuditTable();
      });
    }

    dom.dashAuditPagination.querySelectorAll('[data-page]').forEach(btn => {
      btn.addEventListener('click', () => {
        const p = parseInt(btn.dataset.page, 10);
        if (!isNaN(p) && p !== dashAuditCurrentPage) {
          dashAuditCurrentPage = p;
          renderDashboardAuditTable();
        }
      });
    });

    const selSize = dom.dashAuditPagination.querySelector('#selDashAuditPageSize');
    if (selSize) {
      selSize.addEventListener('change', () => {
        dashAuditPageSize = parseInt(selSize.value, 10) || 10;
        dashAuditCurrentPage = 1;
        renderDashboardAuditTable();
      });
    }
  }

  window.jumpToDashboardRecord = function (query) {
    if (!query) return;
    const cleanQ = String(query).trim();
    if (dom.searchInput) {
      dom.searchInput.value = cleanQ;
    }
    if (dom.modeFilter) dom.modeFilter.value = 'ALL';
    if (dom.modeFilterPanel) dom.modeFilterPanel.value = 'ALL';
    if (dom.kpiModeActiveBadge) dom.kpiModeActiveBadge.textContent = 'ALL';
    switchView('records');
    filterAndRenderRecords();
    showToast(`Showing Land Master records matching "${cleanQ}".`, 'ok');
  };

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

    if (dom.btnCheckUpdates) {
      dom.btnCheckUpdates.addEventListener('click', () => {
        showToast('You are using the latest version (v1.0.6).', 'ok');
      });
    }

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
