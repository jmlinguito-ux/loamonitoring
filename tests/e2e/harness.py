"""Run the real LOA app (Python Api + web UI) against fake_dataverse in a browser."""
import json
import os
import shutil
import sys
import threading
import types
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import openpyxl

import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.abspath(os.path.join(HERE, '..', '..'))   # the LOA project folder
WORK = os.path.join(tempfile.gettempdir(), 'loa_e2e_work')  # a throw-away copy of the app runs here
APP_DIR = os.path.join(WORK, 'app')
FILES = os.path.join(WORK, 'files')

sys.path.insert(0, HERE)
import fake_dataverse as fdv  # noqa: E402


class FakeWindow:
    """Stands in for the pywebview window: records JS callbacks, answers file dialogs."""

    def __init__(self):
        self.js = []
        self.dialog_answers = []
        self.dialogs = []

    def evaluate_js(self, script):
        self.js.append(script)

    def create_file_dialog(self, dialog_type, **kwargs):
        self.dialogs.append((dialog_type, kwargs))
        answer = self.dialog_answers.pop(0) if self.dialog_answers else None
        return (answer,) if answer else None


def _install_webview_stub():
    stub = types.ModuleType('webview')
    stub.OPEN_DIALOG, stub.SAVE_DIALOG, stub.FOLDER_DIALOG = 10, 30, 20
    stub.Window = object
    stub.create_window = lambda *a, **k: None
    stub.start = lambda *a, **k: None
    sys.modules['webview'] = stub


def seed(store):
    """Land Master accounts, earlier audit history and nothing else (references come from files)."""
    loa = 'cr63f_loamonitoring'

    def acct(ctrl, typ, area, mode, tcp=None, ded=None, bal=None, paid=None, **extra):
        values = {
            'cr63f_name': area, 'cr63f_control1': ctrl, 'cr63f_type': typ, 'cr63f_areaindex': area, 'new_mode': mode,
            'cr63f_titleno': extra.pop('title', f'TCT-T-{ctrl[-3:]}'), 'cr63f_registeredowner': extra.pop('owner', f'Owner {ctrl}'),
            'cr63f_province': 'BATANGAS', 'cr63f_municipality': extra.pop('mun', 'NASUGBU'), 'cr63f_entity': 'MBHC',
            'new_contracttype': 'DOAS', 'cr63f_acquisitionsublevel': 'L5.1A',
            'cr63f_totalcontractprice': tcp, 'cr63f_totalpaymentdeductibletotcp': ded,
            'cr63f_balance': bal, 'cr63f_paid': paid, 'cr63f_lotareaha': extra.pop('ha', 1.25),
        }
        values.update(extra)
        return store.insert(loa, values, created='2026-09-01T01:00:00Z', user='Seeder')

    ids = {}
    ids['C-001'] = acct('C-001', 'LAND', 'AI-001', 'SECURED SALE', 1000000, 250000, 750000, 0.3,  # legacy fraction-scale %
                        cr63f_endorseddate='2026-01-05', cr63f_conversionstatus='DAR FOR APPLICATION')
    ids['C-002'] = acct('C-002', 'LAND', 'AI-002', 'SECURED LEASE', 500000, 0, 500000, 0)
    ids['C-003'] = acct('C-003', 'ROW', 'AI-003', 'SECURED SALE', None, None, None, None, cr63f_indexno='ROW-IDX-1',
                        cr63f_endorseddate='2026-03-10')
    ids['C-004'] = acct('C-004', 'LAND - SPECIAL', 'AI-004', 'OLD MODE X', 800000, 100000, 700000, 12.5)
    ids['C-005'] = acct('C-005', 'LAND', 'AI-005', 'SECURED SALE', 300000, 300000, 0, 100,
                        cr63f_endorseddate='1900-01-01')
    for n in range(6, 15):
        ctrl = f'C-{n:03d}'
        ids[ctrl] = acct(ctrl, 'LAND', 'AI-010' if n in (10, 11) else f'AI-{n:03d}', 'SECURED SALE',
                         100000 * n, 10000 * n, 90000 * n, 10.0, cr63f_endorseddate='2026-09-02',
                         cr63f_conversionstatus='DAR SECURED' if n % 2 else 'NOT YET ENDORSED')
    # C-001 already has audit history (a NEW_ENTRY snapshot); the others do not.
    row = store.rows[loa][ids['C-001']]
    snapshot = {k: v for k, v in row.items() if k in fdv.STORE.cols('cr63f_loamonitoringlogs') and v is not None}
    snapshot.update({'cr63f_name': f"AI-001 - Full Record [MASTER:{ids['C-001']}] (NEW_ENTRY)",
                     'cr63f_actiontype': 'NEW_ENTRY', 'cr63f_fieldchanged': 'Full Record', 'cr63f_previousvalue': '-',
                     'cr63f_newvalue': 'Created',
                     'cr63f_details': f"Created new entry AI-001 [MASTER:{ids['C-001']}] [SOURCE:NEW_ENTRY]"})
    store.insert('cr63f_loamonitoringlogs', snapshot, created='2026-09-01T01:00:00Z', user='Seeder')
    return ids


def write_reference_files():
    os.makedirs(FILES, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    # Column C has no header (the A1 bug case); its junk must not shift TCP.
    ws.append(['TYPE', 'Corridor Index', None, 'Area Index', 'Title Mod', 'Total Contract Price',
               'Total Payment Deductible to TCP', '% Paid'])
    ws.append(['SALE', '', 'junk', 'AI-001', 'TCT-T-001', 1200000, 300000, '25%'])
    ws.append(['LEASE', '', 'junk', 'AI-002', 'TCT-T-002', 500000, 125000, '25%'])
    ws.append(['ROW', 'ROW-IDX-1', 'junk', '', '', 650000, 65000, '10%'])
    ws.append(['SALE', '', 'junk', 'AI-006', '', 600000, 60000, '10%'])          # unchanged
    ws.append(['SALE', '', 'junk', 'AI-007', '', 900000, 450000, '50%'])         # changes TCP + ded
    ws.append(['SALE', '', 'junk', 'AI-008', '', 1, 2, ''])                      # conflicting duplicate...
    ws.append(['SALE', '', 'junk', 'AI-008', '', 3, 4, ''])                      # ...same key, different amounts
    ws.append(['SA', '', 'junk', 'AI-004', '', 999, 9, ''])                      # SA never supplies LAND
    finance = os.path.join(FILES, 'finance.xlsx')
    wb.save(finance)
    conversion = os.path.join(FILES, 'conversion.csv')
    with open(conversion, 'w', encoding='utf-8', newline='') as f:
        f.write('AREA-INDEX,TITLE NUMBER,DAR STATUS,NIA STATUS,CONVERSION CLASS,DAR UPDATE,AREA APPLIED (HAS.)\n')
        f.write('AI-001,TCT-T-001,DAR SECURED,NIA DONE,Class A,11/23/2025,1.25\n')
        f.write('AI-002,TCT-T-002,DAR FOR APPLICATION,NIA FOR APPLICATION,Class A,11/23/2025,0.5\n')
        f.write('AI-009,TCT-T-009,DAR FINALITY SECURED,,Class B,1/5/2026,2\n')
    return finance, conversion


def start():
    shutil.rmtree(WORK, ignore_errors=True)
    shutil.copytree(SRC, APP_DIR, ignore=shutil.ignore_patterns('.git', 'web_profile', 'build', 'dist', '__pycache__', 'tests',
                                                                '.build-venv', 'outputs', 'scratch',
                                                                'app_settings.json', '.dv_token_cache.bin'))
    store = fdv.STORE
    server = fdv.serve()
    env = f'http://127.0.0.1:{server.server_address[1]}'
    with open(os.path.join(APP_DIR, 'dataverse_config.json'), 'w') as f:
        json.dump({'environment_url': env, 'tenant_id': 't', 'client_id': 'c', 'solution_name': 'LOA MONITORING',
                   'solution_id': 'sol-1', 'loa_table': 'cr63f_loamonitoring', 'finance_table': 'cr63f_financereference',
                   'conversion_table': 'cr63f_landconversionreference', 'logs_table': 'cr63f_loamonitoringlogs'}, f)
    ids = seed(store)
    finance, conversion = write_reference_files()

    _install_webview_stub()
    sys.path.insert(0, APP_DIR)
    os.chdir(APP_DIR)
    import dataverse as dv
    dv.DataverseClient.try_silent_token = lambda self: True
    dv.DataverseClient.token = lambda self: 'fake-token'
    dv.DataverseClient.current_user_name = lambda self: 'tester@company.com'
    dv.DataverseClient._throttle_wait = lambda self, resp, attempt: 0.05  # keep retries fast in tests
    import app as app_module
    api = app_module.Api()
    window = FakeWindow()
    api.set_window(window)

    bridge = _serve_bridge(api)
    return types.SimpleNamespace(api=api, app=app_module, dv=dv, store=store, ids=ids, window=window,
                                 url=f'http://127.0.0.1:{bridge.server_address[1]}/web/index.html',
                                 finance_file=finance, conversion_file=conversion, files=FILES)


def _serve_bridge(api):
    web_dir = os.path.join(APP_DIR, 'web')

    class Bridge(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def do_GET(self):
            path = self.path.split('?')[0]
            if path == '/stub.js':
                data = STUB_JS.encode()
                ctype = 'application/javascript'
            else:
                local = os.path.join(web_dir, path[len('/web/'):]) if path.startswith('/web/') else ''
                if not local or not os.path.isfile(local):
                    self.send_response(404)
                    self.send_header('Content-Length', '0')
                    self.end_headers()
                    return
                data = open(local, 'rb').read()
                if local.endswith('index.html'):
                    data = data.replace(b'<head>', b'<head><script src="/stub.js"></script>', 1)
                ctype = {'html': 'text/html', 'js': 'application/javascript', 'css': 'text/css',
                         'png': 'image/png', 'ico': 'image/x-icon'}.get(local.rsplit('.', 1)[-1], 'application/octet-stream')
            self.send_response(200)
            self.send_header('Content-Type', ctype)
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def do_POST(self):
            name = self.path.rsplit('/', 1)[-1]
            args = json.loads(self.rfile.read(int(self.headers.get('Content-Length') or 0)) or b'[]')
            try:
                result = getattr(api, name)(*args)
            except Exception as exc:  # surfaced to the page like pywebview does
                result = {'__bridge_error__': f'{type(exc).__name__}: {exc}'}
            data = json.dumps(result, default=str).encode()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)

    server = ThreadingHTTPServer(('127.0.0.1', 0), Bridge)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server


STUB_JS = r"""
window.__bridgeErrors = [];
window.pywebview = { api: new Proxy({}, { get: (_, name) => async (...args) => {
  const r = await fetch('/api/' + name, { method: 'POST', body: JSON.stringify(args) });
  const out = await r.json();
  if (out && out.__bridge_error__) { window.__bridgeErrors.push(name + ': ' + out.__bridge_error__); throw new Error(out.__bridge_error__); }
  return out;
}})};
window.addEventListener('load', () => setTimeout(() => window.dispatchEvent(new Event('pywebviewready')), 50));
"""
