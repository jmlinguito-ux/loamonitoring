// Run with: node tests/test_revert_reference_ui.js
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('web/app.js', 'utf8');

function section(start, end) {
  const startIndex = source.indexOf(start);
  const endIndex = source.indexOf(end, startIndex);
  assert.notEqual(startIndex, -1, `Missing start: ${start}`);
  assert.notEqual(endIndex, -1, `Missing end: ${end}`);
  return source.slice(startIndex, endIndex);
}

const context = { logsState: { logs: [] }, state: { records: [] },
  getFieldNameAttributeKey: field => field };
vm.createContext(context);
vm.runInContext(section('  function getLogActionType(', '  function getFrozenOffset('), context);
vm.runInContext(section('  function revertValuesEqual(', 'function renderLogsChunk('), context);

const finance = { rows: [{ __id: 'ok', value: 20 }, { __id: 'failed', value: 40 }],
  baselineRows: [{ __id: 'ok', value: 10 }, { __id: 'failed', value: 30 }],
  stagedEdits: { ok: { value: 20 }, failed: { value: 40 } } };
context.applyReferenceSaveResult(finance, { success: false, saved_ids: ['ok'], failed_ids: ['failed'] });
assert.deepEqual(Object.keys(finance.stagedEdits), ['failed']);
assert.equal(finance.baselineRows[0].value, 20);
assert.equal(finance.baselineRows[1].value, 30);

for (const action of ['FINANCE_UPDATE', 'CONVERSION_UPDATE']) {
  assert.equal(context.isLogRevertible({ action_type: action, field_name: 'Mode', area_index: 'A1' }), false);
  assert.equal(context.isLogRevertible({ cr63f_name: `A1 - Mode (${action})`, field_name: 'Mode', area_index: 'A1' }), false);
  assert.equal(context.isLogRevertible({ details: `Updated [SOURCE:${action}]`, field_name: 'Mode', area_index: 'A1' }), false);
}

assert.equal(context.getAuditMasterId({ details: 'Changed Mode [MASTER:m1]' }), 'm1');
const logColumns = section('  const LOGS_COLUMNS = [', '  // DOM Elements');
assert.equal(logColumns.includes('__action'), false);
const financeRendering = section('  function renderFinanceTable(', '  function updateStagedEditsUIFinance(');
assert.equal(financeRendering.includes('col-action'), false);
assert.equal(financeRendering.includes('deleteFinanceRecord'), false);
const logsRendering = section('function renderLogsChunk(', '  function attachLogsHeader');
// Decision 2026-10-15: Entry Logs table is read-only; row-level action column and revert were removed.
assert.equal(logsRendering.includes('col-action'), false);
assert.equal(logsRendering.includes('revertLogRecord'), false);
// Neither reference grid has a row delete button (decision 2026-10-12).
const conversionRendering = section('  function renderConversionChunk(', '  function updateStagedEditsUIConversion(');
assert.equal(conversionRendering.includes('deleteConversionRecord'), false);
const dashboardRendering = source.slice(source.indexOf('  function renderDashboardAuditTable('));
assert.equal(dashboardRendering.includes('data-revert='), false);
const tooltipStart = source.indexOf('  window.revertCellFromTooltip = async function');
const tooltipEnd = source.indexOf('\n  };', tooltipStart) + '\n  };'.length;
let revertedId = null;
context.window = { revertLogRecord: async id => { revertedId = id; } };
context.dom = { cellHoverTooltip: {
  _sourceCell: { closest: () => ({}), dataset: {}, classList: { contains: () => false } },
  _rowId: 'm1', _colKey: 'new_mode', _latestLog: { __id: 'l1' },
  classList: { remove() {} }, setAttribute() {},
} };
vm.runInContext(source.slice(tooltipStart, tooltipEnd), context);
context.window.revertCellFromTooltip(null, null).then(() => {
  assert.equal(revertedId, 'l1', 'Saved Land Master changes use the version-checked backend revert.');
  console.log('Reference partial-save and Land Master-only Revert UI checks passed.');
}).catch(error => { console.error(error); process.exitCode = 1; });
