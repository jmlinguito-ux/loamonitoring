"""In-memory stand-in for the Dataverse Web API used by the LOA Monitoring Tool.

Behaves like the real service where the app depends on it: entity metadata,
FormattedValue annotations (en-US, dates without leading zeros), paging with
@odata.nextLink, $filter "a eq 'v' or ...", ETags with If-Match (412),
Prefer: return=representation, 400 for unknown columns / wrong types / text
longer than the column's MaxLength, and injectable 429s and failures.
"""
import json
import re
import threading
import uuid
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import unquote, urlparse, parse_qs

PAGE_SIZE = 7  # small on purpose: every table read goes through paging

S, M, D, DT, MEMO = 'String', 'Money', 'Decimal', 'DateTime', 'Memo'

LOA_COLUMNS = [
    ('cr63f_name', 'Name', S), ('cr63f_control1', 'Control Number', S), ('cr63f_type', 'Type', S),
    ('cr63f_indexno', 'Index No', S), ('cr63f_areaindex', 'Area-Index', S), ('cr63f_titleno', 'Title No', S),
    ('cr63f_registeredowner', 'Registered Owner', S), ('cr63f_authorizedrepresentative', 'Authorized Representative', S),
    ('cr63f_entity', 'Entity', S), ('cr63f_province', 'Province', S), ('cr63f_municipality', 'Municipality', S),
    ('cr63f_dbadjustedbarangay', 'DB Adjusted Barangay', S), ('cr63f_lotareaha', 'Lot Area (Ha)', D),
    ('new_contracttype', 'Contract Type', S), ('new_mode', 'Mode', S), ('cr63f_conversionstatus', 'Conversion Status', S),
    ('cr63f_rccterms', 'RCC Terms', S), ('cr63f_rccprice', 'RCC Price', M),
    ('cr63f_totalcontractprice', 'Total Contract Price', M),
    ('cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP', M),
    ('cr63f_balance', 'Balance', M), ('cr63f_paid', '% Paid', D),
    ('cr63f_secureddate', 'Secured Date', DT), ('cr63f_securedendorsedate', 'Secured Endorse Date', DT),
    ('cr63f_endorseddate', 'Endorsed Date', DT), ('cr63f_acquisitionsublevel', 'Acquisition Sub-Level', S),
]
LOG_COLUMNS = [c for c in LOA_COLUMNS] + [
    ('cr63f_control', 'Control', S), ('cr63f_actiontype', 'Action Type', S), ('cr63f_workweek', 'Work Week', S),
    ('cr63f_logtime', 'Log Time', S), ('cr63f_fieldchanged', 'Field Changed', S),
    ('cr63f_previousvalue', 'Previous Value', S), ('cr63f_newvalue', 'New Value', S), ('cr63f_details', 'Details', MEMO),
]
FIN_COLUMNS = [
    ('cr63f_name', 'Name', S), ('cr63f_type', 'TYPE', S), ('cr63f_corridorindex', 'Corridor Index', S),
    ('cr63f_areaindex', 'Area Index', S), ('cr63f_titlemod', 'Title Mod', S),
    ('cr63f_totalcontractprice', 'Total Contract Price', M),
    ('cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP', M),
    ('cr63f_paid', '% Paid', D), ('cr63f_balance', 'Balance', M),
]
CONV_COLUMNS = [
    ('cr63f_name', 'Name', S), ('cr63f_areaindex', 'AREA-INDEX', S), ('cr63f_titlenumber', 'TITLE NUMBER', S),
    ('cr63f_lotorsurveynumber', 'LOT NUMBER/SURVEY NUMBER', S), ('cr63f_darstatus', 'DAR STATUS', S),
    ('cr63f_niastatus', 'NIA STATUS', S), ('cr63f_conversionclass', 'CONVERSION CLASS', S),
    ('cr63f_darbatch', 'DAR BATCH', S), ('cr63f_darupdate', 'DAR UPDATE', DT), ('cr63f_entity', 'ENTITY', S),
    ('cr63f_municipality', 'MUNICIPALITY/CITY', S), ('cr63f_barangay', 'BARANGAY', S),
    ('cr63f_areaappliedhas', 'AREA APPLIED (HAS.)', D),
]
TABLES = {
    'cr63f_loamonitoring': ('cr63f_loamonitorings', LOA_COLUMNS),
    'cr63f_loamonitoringlogs': ('cr63f_loamonitoringlogses', LOG_COLUMNS),
    'cr63f_financereference': ('cr63f_financereferences', FIN_COLUMNS),
    'cr63f_landconversionreference': ('cr63f_landconversionreferences', CONV_COLUMNS),
}
PRECISION = {'cr63f_paid': 1, 'cr63f_lotareaha': 4, 'cr63f_areaappliedhas': 4}
MAX_LEN = {S: 100, MEMO: 4000}
SET_TO_TABLE = {v[0]: k for k, v in TABLES.items()}


class Store:
    def __init__(self):
        self.lock = threading.RLock()
        self.rows = {t: {} for t in TABLES}
        self.faults = {}          # name -> remaining count
        self.requests = []        # (method, path, status)
        self.rejections = []      # 400/412 bodies, for the report
        self.user = 'Tester One'

    # ---- helpers
    def cols(self, table):
        return {c[0]: c for c in TABLES[table][1]}

    def pk(self, table):
        return f'{table}id'

    def fault(self, name):
        with self.lock:
            n = self.faults.get(name, 0)
            if n > 0:
                self.faults[name] = n - 1
                return True
        return False

    def insert(self, table, values, created=None, user=None):
        rid = str(uuid.uuid4())
        row = dict(values)
        row[self.pk(table)] = rid
        row['createdon'] = created or datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        row['_createdby'] = user or self.user
        row['_version'] = 1
        self.rows[table][rid] = row
        return rid

    def find(self, table, **match):
        return [r for r in self.rows[table].values() if all(r.get(k) == v for k, v in match.items())]


STORE = Store()


def _fmt(table, col, value):
    """FormattedValue text, as Dataverse returns it for an en-US user."""
    if value is None or value == '':
        return None
    kind = STORE.cols(table)[col][2]
    if kind == M:
        return f"₱{float(value):,.2f}"
    if kind == D:
        p = PRECISION.get(col, 2)
        return f"{float(value):,.{p}f}"
    if kind == DT:
        d = datetime.fromisoformat(str(value).replace('Z', '+00:00'))
        return f"{d.month}/{d.day}/{d.year}"
    return None


def _created_fmt(value):
    d = datetime.fromisoformat(value.replace('Z', '+00:00')).astimezone()
    return f"{d.month}/{d.day}/{d.year} {d.strftime('%I:%M %p').lstrip('0')}"


def _render(table, row):
    out = {'@odata.etag': f'W/"{row["_version"]}"', STORE.pk(table): row[STORE.pk(table)]}
    for col, _, kind in TABLES[table][1]:
        value = row.get(col)
        out[col] = value
        f = _fmt(table, col, value)
        if f is not None:
            out[f'{col}@OData.Community.Display.V1.FormattedValue'] = f
    out['createdon'] = row['createdon']
    out['createdon@OData.Community.Display.V1.FormattedValue'] = _created_fmt(row['createdon'])
    out['_createdby_value'] = 'user-guid'
    out['_createdby_value@OData.Community.Display.V1.FormattedValue'] = row['_createdby']
    return out


def _validate(table, payload, creating):
    cols = STORE.cols(table)
    clean = {}
    for key, value in payload.items():
        if key not in cols:
            return None, f"Invalid property '{key}' was found in entity 'Microsoft.Dynamics.CRM.{table}'."
        kind = cols[key][2]
        if value is None:
            clean[key] = None
            continue
        if kind in (M, D):
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                return None, f"Cannot convert the literal '{value}' to the expected type 'Edm.Decimal' ({key})."
            clean[key] = round(float(value), PRECISION.get(key, 2))
        elif kind == DT:
            if not isinstance(value, str) or not re.fullmatch(r'\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}Z?)?', value):
                return None, f"Cannot convert the literal '{value}' to the expected type 'Edm.DateTimeOffset' ({key})."
            clean[key] = value
        else:
            if not isinstance(value, str):
                return None, f"A value of type '{type(value).__name__}' was given for String column {key}."
            if len(value) > MAX_LEN[kind]:
                return None, f"A validation error occurred. The length of the '{key}' attribute of the '{table}' entity exceeded the maximum allowed length of '{MAX_LEN[kind]}'."
            clean[key] = value
    return clean, ''


def _metadata_attrs(table):
    attrs = [{
        'LogicalName': col, 'SchemaName': col, 'AttributeType': kind if kind != MEMO else 'Memo',
        'DisplayName': {'UserLocalizedLabel': {'Label': disp}}, 'IsValidForCreate': True,
        'IsValidForUpdate': True, 'IsValidForRead': True,
    } for col, disp, kind in TABLES[table][1]]
    attrs.append({'LogicalName': STORE.pk(table), 'SchemaName': STORE.pk(table), 'AttributeType': 'Uniqueidentifier',
                  'DisplayName': {'UserLocalizedLabel': {'Label': 'Id'}}, 'IsValidForCreate': True,
                  'IsValidForUpdate': False, 'IsValidForRead': True})
    attrs.append({'LogicalName': 'createdon', 'SchemaName': 'CreatedOn', 'AttributeType': 'DateTime',
                  'DisplayName': {'UserLocalizedLabel': {'Label': 'Created On'}}, 'IsValidForCreate': False,
                  'IsValidForUpdate': False, 'IsValidForRead': True})
    attrs.append({'LogicalName': 'createdby', 'SchemaName': 'CreatedBy', 'AttributeType': 'Lookup',
                  'DisplayName': {'UserLocalizedLabel': {'Label': 'Created By'}}, 'IsValidForCreate': False,
                  'IsValidForUpdate': False, 'IsValidForRead': True})
    return attrs


def _parse_filter(text):
    clauses = []
    for part in re.split(r'\s+or\s+', text or ''):
        m = re.fullmatch(r"\s*(\w+)\s+eq\s+'((?:[^']|'')*)'\s*", part)
        if not m:
            return None
        clauses.append((m.group(1), m.group(2).replace("''", "'")))
    return clauses


class Handler(BaseHTTPRequestHandler):
    protocol_version = 'HTTP/1.1'

    def log_message(self, *args):
        pass

    def _send(self, status, body=None, headers=None):
        data = b'' if body is None else json.dumps(body).encode()
        self.send_response(status)
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        if body is not None:
            self.send_header('Content-Type', 'application/json; odata.metadata=minimal')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        if data:
            self.wfile.write(data)
        STORE.requests.append((self.command, self.path, status))

    def _error(self, status, message):
        STORE.rejections.append((self.command, self.path, status, message))
        self._send(status, {'error': {'code': '0x80040265', 'message': message}})

    def _body(self):
        n = int(self.headers.get('Content-Length') or 0)
        return json.loads(self.rfile.read(n) or b'{}')

    def _route(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        assert path.startswith('/api/data/v9.2/'), path
        return path[len('/api/data/v9.2/'):], parse_qs(parsed.query)

    def _throttled(self):
        if STORE.fault('throttle'):
            STORE.requests.append((self.command, self.path, 429))
            body = json.dumps({'error': {'message': 'Number of requests exceeded the limit of 6000 over time window of 300 seconds.'}}).encode()
            self.send_response(429)
            self.send_header('Retry-After', '1')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return True
        return False

    # ---------------------------------------------------------------- GET
    def do_GET(self):
        if self._throttled():
            return
        rest, query = self._route()
        with STORE.lock:
            m = re.fullmatch(r"EntityDefinitions\(LogicalName='(\w+)'\)", rest)
            if m:
                t = m.group(1)
                if t not in TABLES:
                    return self._error(404, f'Could not find entity {t}')
                return self._send(200, {'LogicalName': t, 'EntitySetName': TABLES[t][0], 'PrimaryIdAttribute': STORE.pk(t),
                                        'PrimaryNameAttribute': 'cr63f_name', 'DisplayName': {'UserLocalizedLabel': {'Label': t}}})
            m = re.fullmatch(r"EntityDefinitions\(LogicalName='(\w+)'\)/Attributes(/Microsoft\.Dynamics\.CRM\.(\w+))?", rest)
            if m:
                t, kind = m.group(1), m.group(3)
                if kind == 'DateTimeAttributeMetadata':
                    return self._send(200, {'value': [{'LogicalName': c, 'DateTimeBehavior': {'Value': 'DateOnly'}}
                                                      for c, _, k in TABLES[t][1] if k == DT]})
                if kind == 'PicklistAttributeMetadata':
                    return self._send(200, {'value': []})
                return self._send(200, {'value': _metadata_attrs(t)})
            if rest == 'EntityDefinitions':
                return self._send(200, {'value': [{'MetadataId': f'meta-{t}', 'LogicalName': t, 'EntitySetName': s,
                                                   'DisplayName': {'UserLocalizedLabel': {'Label': t}}, 'IsCustomEntity': True,
                                                   'PrimaryNameAttribute': 'cr63f_name'} for t, (s, _) in TABLES.items()]})
            if rest == 'solutions':
                return self._send(200, {'value': [{'solutionid': 'sol-1', 'uniquename': 'loa', 'friendlyname': 'LOA MONITORING',
                                                   'version': '1.0', 'ismanaged': False}]})
            m = re.fullmatch(r"solutions\(([^)]+)\)/solution_solutioncomponent", rest)
            if m:
                return self._send(200, {'value': [{'objectid': f'meta-{t}'} for t in TABLES]})
            m = re.fullmatch(r"(\w+)\(([^)]+)\)", rest)
            if m and m.group(1) in SET_TO_TABLE:
                t = SET_TO_TABLE[m.group(1)]
                row = STORE.rows[t].get(m.group(2))
                if not row:
                    return self._error(404, 'Not found')
                return self._send(200, _render(t, row))
            if rest in SET_TO_TABLE:
                t = SET_TO_TABLE[rest]
                rows = list(STORE.rows[t].values())
                if '$filter' in query:
                    clauses = _parse_filter(query['$filter'][0])
                    if clauses is None:
                        return self._error(400, f"Unsupported filter {query['$filter'][0]}")
                    for col, _ in clauses:
                        if col not in STORE.cols(t):
                            return self._error(400, f"Could not find a property named '{col}' on type 'Microsoft.Dynamics.CRM.{t}'.")
                    rows = [r for r in rows if any(str(r.get(c) or '') == v for c, v in clauses)]
                rows.sort(key=lambda r: r['createdon'])
                skip = int((query.get('$skiptoken') or ['0'])[0])
                page = rows[skip:skip + PAGE_SIZE]
                body = {'value': [_render(t, r) for r in page]}
                if skip + PAGE_SIZE < len(rows):
                    q = parse_qs(urlparse(self.path).query)
                    q['$skiptoken'] = [str(skip + PAGE_SIZE)]
                    from urllib.parse import urlencode
                    body['@odata.nextLink'] = f"http://{self.headers['Host']}/api/data/v9.2/{rest}?{urlencode({k: v[0] for k, v in q.items()})}"
                return self._send(200, body)
        return self._error(404, f'No route {rest}')

    # ---------------------------------------------------------------- POST
    def do_POST(self):
        if self._throttled():
            return
        rest, _ = self._route()
        payload = self._body()
        if rest == 'PublishXml' or rest.startswith('EntityDefinitions'):
            return self._error(400, 'Schema changes are not expected in tests')
        if rest not in SET_TO_TABLE:
            return self._error(404, rest)
        t = SET_TO_TABLE[rest]
        if t == 'cr63f_loamonitoringlogs' and STORE.fault('log_post_fail'):
            return self._error(500, 'Injected audit write failure')
        clean, err = _validate(t, payload, True)
        if err:
            return self._error(400, err)
        with STORE.lock:
            rid = STORE.insert(t, clean)
        self._send(204, None, {'OData-EntityId': f"http://{self.headers['Host']}/api/data/v9.2/{rest}({rid})"})

    # ---------------------------------------------------------------- PATCH
    def do_PATCH(self):
        if self._throttled():
            return
        rest, query = self._route()
        m = re.fullmatch(r"(\w+)\(([^)]+)\)", rest)
        t = SET_TO_TABLE.get(m.group(1)) if m else None
        if not t:
            return self._error(404, rest)
        payload = self._body()
        if STORE.fault(f'patch_fail_{t}'):
            return self._error(500, 'Injected update failure')
        with STORE.lock:
            row = STORE.rows[t].get(m.group(2))
            if not row:
                return self._error(404, 'Record not found')
            if_match = self.headers.get('If-Match')
            if if_match and if_match != f'W/"{row["_version"]}"':
                return self._error(412, 'The version of the existing record doesn\'t match the RowVersion property provided.')
            clean, err = _validate(t, payload, False)
            if err:
                return self._error(400, err)
            row.update(clean)
            row['_version'] += 1
            if 'return=representation' in (self.headers.get('Prefer') or ''):
                return self._send(200, {'@odata.etag': f'W/"{row["_version"]}"', STORE.pk(t): row[STORE.pk(t)]})
        self._send(204)

    # ---------------------------------------------------------------- DELETE
    def do_DELETE(self):
        if self._throttled():
            return
        rest, _ = self._route()
        m = re.fullmatch(r"(\w+)\(([^)]+)\)", rest)
        t = SET_TO_TABLE.get(m.group(1)) if m else None
        if not t:
            return self._error(404, rest)
        if STORE.fault(f'delete_fail_{t}'):
            return self._error(403, 'Injected: user lacks delete privilege')
        with STORE.lock:
            if STORE.rows[t].pop(m.group(2), None) is None:
                return self._error(404, 'Record not found')
        self._send(204)


def serve(port=0):
    server = ThreadingHTTPServer(('127.0.0.1', port), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server
