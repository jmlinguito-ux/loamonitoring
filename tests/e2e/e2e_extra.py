"""Second E2E batch: the remaining UI interactions (grid tools, tooltips, filters, Overview tools, drawer lookup)."""
import json
import os
import sys

from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
src = open(os.path.join(HERE, 'e2e_test.py')).read()
# Reuse helpers from the main script (everything before the first test).
exec(src.split('# =====================================================================================\n# 1. Startup')[0])


def fill_drawer(values):
    for sel, val in values.items():
        page.evaluate("""([sel, val]) => { const e = document.querySelector(sel); e.value = val;
            e.dispatchEvent(new Event('input', {bubbles: true})); e.dispatchEvent(new Event('change', {bubbles: true}));
            e.dispatchEvent(new Event('blur')); }""", [sel, val])
    page.wait_for_timeout(500)


def set_header_filter(sel, value):
    page.evaluate("""([sel, value]) => { const i = document.querySelector(sel); i.value = value;
        i.dispatchEvent(new Event('input', {bubbles: true})); }""", [sel, value])
    page.wait_for_timeout(700)


@step('Grid tools', 'header filter, freeze, resize, copy')
def t_grid_tools():
    nav('records')
    set_header_filter('#recordsTableHead [data-filter-col="cr63f_municipality"]', 'nasug')
    page.wait_for_timeout(700)
    n = page.locator('#recordsTableBody tr').count()
    check('Grid tools', 'Column header filter narrows rows (Municipality contains "nasug")', n == 14, n)
    set_header_filter('#recordsTableHead [data-filter-col="cr63f_control1"]', 'C-01')
    page.wait_for_timeout(700)
    n = page.locator('#recordsTableBody tr').count()
    check('Grid tools', 'Two header filters combine (Control contains C-01)', n == 5, n)
    page.click('#btnClearFilters')
    page.wait_for_timeout(600)
    check('Grid tools', 'Clear filters also clears header filters', page.locator('#recordsTableBody tr').count() == 14,
          page.locator('#recordsTableBody tr').count())
    page.evaluate("document.querySelector('#recordsTableHead [data-pin-col=\"cr63f_areaindex\"]').click()")
    page.wait_for_timeout(500)
    frozen = page.locator('#recordsTableBody tr:first-child td[data-col-key="cr63f_areaindex"].col-frozen').count()
    check('Grid tools', 'Freezing a column pins it', frozen == 1, frozen)
    page.evaluate("document.querySelector('#recordsTableHead [data-pin-col=\"cr63f_areaindex\"]').click()")
    th = page.locator('#recordsTableHead th[data-col-key="cr63f_registeredowner"]')
    w0 = th.bounding_box()['width']
    handle = th.locator('.th-resizer').first
    box = handle.bounding_box()
    page.mouse.move(box['x'] + box['width'] / 2, box['y'] + box['height'] / 2)
    page.mouse.down()
    page.mouse.move(box['x'] + 80, box['y'] + box['height'] / 2, steps=5)
    page.mouse.up()
    page.wait_for_timeout(400)
    w1 = th.bounding_box()['width']
    check('Grid tools', 'Dragging a header edge resizes the column', w1 > w0 + 40, (w0, w1))
    text = page.evaluate("""() => {
        const a = document.querySelector('#recordsTableBody tr:nth-child(1) td[data-col-key="cr63f_control1"]');
        const b = document.querySelector('#recordsTableBody tr:nth-child(2) td[data-col-key="cr63f_type"]');
        a.dispatchEvent(new MouseEvent('mousedown', {bubbles: true, button: 0}));
        b.dispatchEvent(new MouseEvent('mousedown', {bubbles: true, button: 0, shiftKey: true}));
        document.dispatchEvent(new MouseEvent('mouseup', {bubbles: true}));
        const dt = new DataTransfer();
        a.dispatchEvent(new ClipboardEvent('copy', {clipboardData: dt, bubbles: true, cancelable: true}));
        return dt.getData('text/plain'); }""")
    rows = [r.split('\t') for r in text.strip().split('\n')] if text else []
    check('Grid tools', 'Copying a 2x2 block gives tab-separated text', len(rows) == 2 and all(len(r) == 2 for r in rows), text)


@step('Grid tools', 'changed-only, tooltip, keyboard')
def t_grid_more():
    type_cell('C-012', 'cr63f_province', 'RIZAL')
    page.click('#btnToggleShowOnlyChanged')
    page.wait_for_timeout(500)
    check('Grid tools', '"Show Changed Only" shows just the edited row', page.locator('#recordsTableBody tr').count() == 1,
          page.locator('#recordsTableBody tr').count())
    page.hover(f'{row_sel("C-012")} td[data-col-key="cr63f_province"]')
    page.wait_for_timeout(700)
    tip = page.text_content('#cellHoverTooltip') or ''
    check('Grid tools', 'Hover tooltip shows old → new for an unsaved edit', 'BATANGAS' in tip and 'RIZAL' in tip, tip[:200])
    page.evaluate("window.revertCellFromTooltip(null, document.querySelector('.cell-tooltip-revert-btn'))")
    page.wait_for_timeout(600)
    check('Grid tools', 'Tooltip "Revert" undoes the unsaved edit', not page.is_visible('#stagedEditsWrap'),
          page.text_content('#unsavedEditsChip'))
    page.click('#btnClearFilters')
    page.wait_for_timeout(400)
    if page.is_visible('#btnToggleShowOnlyChanged'):
        page.click('#btnToggleShowOnlyChanged')
    sel = cell_input('C-012', 'cr63f_province')
    page.focus(sel)
    page.fill(sel, 'XX')
    page.keyboard.press('Escape')
    page.wait_for_timeout(300)
    check('Grid tools', 'Escape cancels typing in a cell', cell_value('C-012', 'cr63f_province') == 'BATANGAS', cell_value('C-012', 'cr63f_province'))
    page.focus(sel)
    page.fill(sel, 'PAMPANGA')
    page.keyboard.press('Enter')
    page.wait_for_timeout(400)
    check('Grid tools', 'Enter commits the cell (staged)', page.is_visible('#stagedEditsWrap'), page.text_content('#unsavedEditsChip'))
    page.click('#btnDiscardStagedEdits')
    try:
        modal_ok(500, timeout=1500)
    except Exception:
        pass


@step('References', 'no reference data, then cancel')
def t_refs_cancel():
    m = mark()
    page.click('#btnUpdateFromFinance')
    page.wait_for_timeout(800)
    check('References', 'Update from References without reference data explains what to load',
          any('Reference Data' in t for _, t in toasts(m)), toasts(m))
    nav('finance')
    answer_dialog(env.finance_file)
    page.click('#btnUploadFinance')
    page.wait_for_timeout(1500)
    set_select('#typeFilterFinance', next((o for o in page.eval_on_selector_all('#typeFilterFinance option', 'os => os.map(o => o.value)') if o.upper() == 'LEASE'), 'ALL'))
    check('Finance', 'Finance TYPE filter shows only LEASE rows', page.locator('#financePreviewTbody tr').count() == 1,
          page.locator('#financePreviewTbody tr').count())
    page.click('#btnClearFiltersFinance')
    page.wait_for_timeout(300)
    page.click('#btnColumnsPickerFinance')
    page.evaluate("""() => { const c = document.querySelector('#colPickerMenuFinance [data-fin-col-toggle="% Paid"]');
        c.checked = false; c.dispatchEvent(new Event('change', {bubbles: true})); }""")
    page.wait_for_timeout(400)
    check('Finance', 'Finance column picker hides a column', page.locator('#financePreviewThead [data-fin-sort="% Paid"]').count() == 0)
    page.mouse.click(5, 5)
    nav('records')
    before = {r['cr63f_control1']: r['cr63f_totalcontractprice'] for r in S.rows[LOA].values()}
    page.click('#btnUpdateFromFinance')
    page.wait_for_timeout(2000)
    check('References', 'Changes are staged (not saved yet)', page.is_visible('#stagedEditsWrap'))
    page.click('#btnDiscardStagedEdits')
    try:
        modal_ok(800, timeout=2000)
    except Exception:
        pass
    after = {r['cr63f_control1']: r['cr63f_totalcontractprice'] for r in S.rows[LOA].values()}
    check('References', 'Discarding staged reference changes writes nothing', before == after)
    check('References', 'Grid shows stored values again after discard', cell_value('C-001', 'cr63f_totalcontractprice') == '₱1,000,000.00',
          cell_value('C-001', 'cr63f_totalcontractprice'))


@step('New entry', 'live finance lookup')
def t_drawer_lookup():
    page.click('#btnOpenNewEntry')
    page.wait_for_timeout(400)
    fill_drawer({'#fType': 'LAND', '#fMode': 'SECURED SALE', '#fAreaIndex': 'AI-007'})
    page.evaluate("document.getElementById('areaIndexSuggestions').classList.add('hidden')")
    page.wait_for_timeout(1200)
    tcp = page.input_value('#fTcp')
    check('New entry', 'Typing an AREA-INDEX fills TCP/Deductible from Finance', tcp.replace(',', '').startswith('900000'), tcp)
    page.click('#btnCancelDrawer')
    page.wait_for_timeout(300)
    check('New entry', 'Cancel closes the drawer', not page.is_visible('#entryDrawer'))


@step('Entry Logs', 'columns, sort, header filter')
def t_logs_tools():
    type_cell('C-012', 'cr63f_province', 'RIZAL')
    page.click('#btnSaveStagedEdits')
    modal_ok(2000)
    nav('logs', 2000)
    page.click('#btnColumnsPickerLogs')
    page.evaluate("""() => { const c = document.querySelector('#colPickerMenuLogs [data-log-col-toggle="details"]');
        c.checked = false; c.dispatchEvent(new Event('change', {bubbles: true})); }""")
    page.wait_for_timeout(400)
    check('Entry Logs', 'Logs column picker hides Details', page.locator('#logsTableHead [data-log-sort="details"]').count() == 0)
    page.mouse.click(5, 5)
    set_header_filter('#logsTableHead [data-log-filter-col="field_name"]', 'Province')
    page.wait_for_timeout(600)
    vals = page.eval_on_selector_all('#logsTableBody td[data-col-key="field_name"]', 'es => es.map(e => e.textContent.trim())')
    check('Entry Logs', 'Header filter on Field Changed', vals and all(v == 'Province' for v in vals), vals)
    page.click('#btnClearFiltersLogs')
    page.wait_for_timeout(500)
    page.click('#logsTableHead [data-log-sort="field_name"]')
    page.wait_for_timeout(500)
    vals = page.eval_on_selector_all('#logsTableBody td[data-col-key="field_name"]', 'es => es.map(e => e.textContent.trim())')
    check('Entry Logs', 'Sorting logs by Field Changed', vals == sorted(vals) or vals == sorted(vals, reverse=True), vals[:6])
    kpi = page.text_content('#kpiLogsCount')
    check('Entry Logs', 'Logs KPI count shows', kpi and kpi.strip() != '0', kpi)
    page.click('#btnRefreshLogs')
    page.wait_for_timeout(1500)
    check('Entry Logs', 'Refresh Logs reloads', page.locator('#logsTableBody tr').count() >= 1)


@step('Overview', 'search, source filter, open account')
def t_overview_tools():
    nav('dashboard', 2500)
    total = page.locator('#dashAuditTableBody tr.ov-row').count()
    page.fill('#dashActivitySearchInput', 'C-012')
    page.dispatch_event('#dashActivitySearchInput', 'input')
    page.wait_for_timeout(500)
    check('Overview', 'Search in Changed accounts', page.locator('#dashAuditTableBody tr.ov-row').count() == 1, (total, page.locator('#dashAuditTableBody tr.ov-row').count()))
    page.fill('#dashActivitySearchInput', '')
    page.dispatch_event('#dashActivitySearchInput', 'input')
    page.click('#dashSourceFilter [data-src="manual"]')
    page.wait_for_timeout(500)
    srcs = page.eval_on_selector_all('#dashAuditTableBody tr.ov-row td:nth-child(4)', 'es => es.map(e => e.textContent.trim())')
    check('Overview', 'Source filter "Manual" lists only manual changes', srcs and all('Manual' in s for s in srcs), srcs)
    page.click('#dashSourceFilter [data-src="ALL"]')
    page.wait_for_timeout(300)
    page.locator('#dashAuditTableBody tr.ov-row', has_text='C-012').first.click()
    page.wait_for_timeout(400)
    page.click('#dashAuditTableBody [data-open]')
    page.wait_for_timeout(1500)
    check('Overview', '"Open in Land Master" jumps to the account', page.is_visible('#view-records')
          and page.locator('#recordsTableBody tr').count() == 1, page.locator('#recordsTableBody tr').count())


@step('Settings', 'density and solutions refresh')
def t_settings_more():
    nav('settings', 1500)
    page.click('#densitySeg [data-density="compact"]')
    page.wait_for_timeout(300)
    check('Settings', 'Density switch applies', page.evaluate("document.documentElement.getAttribute('data-density')") == 'compact')
    page.click('#densitySeg [data-density="comfortable"]')
    page.click('#btnRefreshSolutions')
    page.wait_for_timeout(1500)
    check('Settings', 'Refresh solutions keeps the list', page.locator('#selSolution option').count() >= 1)
    nav('conversion')
    answer_dialog(env.conversion_file)
    page.click('#btnUploadConversion')
    page.wait_for_timeout(1200)
    opts = page.eval_on_selector_all('#statusFilterConversion option', 'os => os.map(o => o.value)')
    target = next((o for o in opts if 'APPLICATION' in o.upper()), None)
    if target:
        set_select('#statusFilterConversion', target)
        check('Conversion', 'Status filter narrows rows', page.locator('#conversionPreviewTbody tr').count() == 1,
              page.locator('#conversionPreviewTbody tr').count())
    else:
        check('Conversion', 'Status filter lists DAR statuses', False, opts)
    page.click('#btnClearFiltersConversion')
    nav('finance')
    page.click('#btnClearFinanceData')
    try:
        modal_ok(800, timeout=2000)
    except Exception:
        pass
    page.wait_for_timeout(500)
    check('Finance', 'Clear Finance Data empties the table', len(env.api.finance_rows) == 0 and env.api.settings.get('remembered_finance_file') == '',
          len(env.api.finance_rows))


for test in (t_grid_tools, t_grid_more, t_refs_cancel, t_drawer_lookup, t_logs_tools, t_overview_tools, t_settings_more):
    nav('records', 500) if test is not t_grid_tools else None
    test()

check('General', 'No JavaScript errors on the page', not PAGE_ERRORS, PAGE_ERRORS[:3])
check('General', 'No Python exceptions through the bridge', not page.evaluate('window.__bridgeErrors'), page.evaluate('window.__bridgeErrors'))
check('General', 'Dataverse never rejected a payload as invalid (400)', not [r for r in S.rejections if r[2] == 400], S.rejections[:3])
json.dump({'results': RESULTS}, open(os.path.join(harness.WORK, 'results_extra.json'), 'w'), indent=1, default=str)
browser.close()
pw.stop()
fails = [r for r in RESULTS if not r['ok']]
print(f"\n{len(RESULTS) - len(fails)} passed, {len(fails)} failed")
