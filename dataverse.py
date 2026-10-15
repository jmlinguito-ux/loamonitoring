"""Dataverse connector for LOA Monitoring Tool.

Provides authenticated MSAL device-code connection, metadata discovery, and OData
v9.2 CRUD operations for `cr63f_loamonitoring` and `APPENDED_FINANCE` tables.
"""

from __future__ import annotations

import json
import math
import os
import re
from urllib.parse import quote
import time
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple

import requests
from app_paths import data_file
from loa_engine import parse_log_datetime

try:
    import msal
except ImportError:
    msal = None

API_VERSION = 'v9.2'
_TIMEOUT = 60


def config_path() -> str:
    return data_file('dataverse_config.json')


def _token_cache_path() -> str:
    return data_file('.dv_token_cache.bin')


# Connection used whenever a value is missing (decided 2026-10-13).
DEFAULT_CONFIG = {
    "environment_url": "https://orgfa075889.crm.dynamics.com",
    "tenant_id": "70234e6c-22dd-42ec-b3ca-fee4c6a42328",
    "client_id": "c16cfa67-0d61-46b4-8813-94197d53ce6b",
    "solution_name": "LOA MONITORING (LOAMONITORING)",
    "solution_id": "34c6d4d3-b4b0-f111-aaac-7ced8d3b4de7",
    "loa_table": "cr63f_loamonitoring",
    "logs_table": "cr63f_loamonitoringlogs",
    "finance_table": "cr63f_financereference",
    "conversion_table": "cr63f_landconversionreference",
}
# Old default solution: replaced by LOA MONITORING when found in a saved config.
_OLD_DEFAULT_SOLUTION_ID = "a7f41cec-5786-f111-ab0f-7ced8d3b4de7"


def load_config() -> dict:
    path = config_path()
    if not os.path.exists(path):
        # Fallback to local config if present
        local_cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dataverse_config.json')
        if os.path.exists(local_cfg):
            path = local_cfg
        else:
            return dict(DEFAULT_CONFIG)
    try:
        with open(path, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
    except (OSError, ValueError):
        cfg = {}
    if not isinstance(cfg, dict):
        cfg = {}
    changed = False
    if str(cfg.get('solution_id') or '') == _OLD_DEFAULT_SOLUTION_ID:
        cfg['solution_id'] = DEFAULT_CONFIG['solution_id']
        cfg['solution_name'] = DEFAULT_CONFIG['solution_name']
        changed = True
    for key, value in DEFAULT_CONFIG.items():
        if not str(cfg.get(key) or '').strip():
            cfg[key] = value
            changed = True
    if changed:
        try:
            save_config(cfg)
        except Exception:
            pass
    for key in ('environment_url', 'tenant_id', 'client_id'):
        if not cfg.get(key):
            raise ValueError(f'dataverse_config.json is missing "{key}".')
    cfg['environment_url'] = cfg['environment_url'].rstrip('/')
    if cfg.get('finance_table') in ('cr63f_appendedfinance', '', None):
        cfg['finance_table'] = 'cr63f_financereference'
        try:
            save_config(cfg)
        except Exception:
            pass
    if cfg.get('conversion_table') in ('', None):
        cfg['conversion_table'] = 'cr63f_landconversionreference'
        try:
            save_config(cfg)
        except Exception:
            pass
    if cfg.get('loa_table') in ('', None):
        cfg['loa_table'] = 'cr63f_loamonitoring'
    if cfg.get('logs_table') in ('', None):
        cfg['logs_table'] = 'cr63f_loamonitoringlogs'
    return cfg


def save_config(cfg: dict) -> None:
    path = config_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, indent=2)


def _s(val: Any) -> str:
    return '' if val is None else str(val).strip()


def _norm(s: Any) -> str:
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def _clean_key(s: Any) -> str:
    cleaned = re.sub(r'[^a-z0-9]', '', str(s or '').lower())
    for prefix in ('cr63f', 'new', 'crc5e'):
        if cleaned.startswith(prefix):
            cleaned = cleaned[len(prefix):]
    return cleaned


def _extract_label(disp: Any, default: str = '') -> str:
    if not isinstance(disp, dict):
        return default or ''
    user_label = disp.get('UserLocalizedLabel')
    if isinstance(user_label, dict) and user_label.get('Label'):
        return str(user_label.get('Label')).strip()
    localized = disp.get('LocalizedLabels')
    if isinstance(localized, list) and len(localized) > 0 and isinstance(localized[0], dict) and localized[0].get('Label'):
        return str(localized[0].get('Label')).strip()
    return default or ''


def _clean_dataverse_datetime(val: Any, behavior: str = 'UserLocal') -> Optional[str]:
    """Sanitize datetime string for Dataverse CrmDateTime (range: 1753-01-01 to 9999-12-31).
    Returns ISO date string, or None if invalid or out of range.

    ``behavior`` is the column's DateTimeBehavior. App times are local (PH)
    wall-clock times, so for UserLocal columns they are converted to UTC before
    the 'Z' suffix is added; labelling local time as 'Z' shifted them by +8h.
    TimeZoneIndependent columns get the wall-clock time without a zone, and
    DateOnly columns get the date only.
    """
    if val is None:
        return None
    s = str(val).strip()
    if not s or s.lower() in ('none', 'null', '-', 'undefined'):
        return None

    # Handle 0001-01-01 or 01/01/0001 fast check
    if '0001' in s or '0000' in s:
        return None

    explicit_utc = s.endswith('Z') or bool(re.search(r'[+-]\d{2}:?\d{2}$', s))
    if explicit_utc and 'T' in s:
        try:
            aware = datetime.fromisoformat(s.replace('Z', '+00:00'))
            if aware.tzinfo is not None:
                if not (1753 <= aware.year <= 9999):
                    return None
                if behavior == 'DateOnly':
                    return aware.astimezone().strftime('%Y-%m-%d')
                utc_value = aware.astimezone(timezone.utc)
                if behavior == 'TimeZoneIndependent':
                    return aware.astimezone().strftime('%Y-%m-%dT%H:%M:%S')
                return utc_value.strftime('%Y-%m-%dT%H:%M:%SZ')
        except ValueError:
            pass

    formats = [
        ('%m/%d/%Y %I:%M:%S %p', True),
        ('%m/%d/%Y %I:%M %p', True),
        ('%m/%d/%Y %H:%M:%S', True),
        ('%m/%d/%Y %H:%M', True),
        ('%Y-%m-%d %H:%M:%S', True),
        ('%Y-%m-%d %H:%M', True),
        ('%Y-%m-%dT%H:%M:%SZ', True),
        ('%Y-%m-%dT%H:%M:%S', True),
        ('%m/%d/%Y', False),
        ('%Y-%m-%d', False),
        ('%m/%d/%y', False),
        ('%Y/%m/%d', False),
    ]

    dt = None
    has_time = False
    for fmt, with_t in formats:
        try:
            dt = datetime.strptime(s, fmt)
            has_time = with_t
            break
        except (ValueError, TypeError):
            continue

    if not dt:
        parts = re.split(r'[-/.]', s.split('T')[0].split(' ')[0])
        if len(parts) == 3:
            try:
                p0, p1, p2 = int(parts[0]), int(parts[1]), int(parts[2])
                if p0 > 100:
                    year, month, day = p0, p1, p2
                elif p2 > 100:
                    year = p2
                    if p0 > 12 and 1 <= p1 <= 12:
                        month, day = p1, p0
                    else:
                        month, day = p0, p1
                elif p2 < 100:
                    year = 2000 + p2 if p2 < 50 else 1900 + p2
                    if p0 > 12 and 1 <= p1 <= 12:
                        month, day = p1, p0
                    else:
                        month, day = p0, p1
                else:
                    return None
                dt = datetime(year, month, day)
            except Exception:
                return None
        else:
            return None

    if dt.year < 1753 or dt.year > 9999:
        return None

    if behavior == 'DateOnly' or not (has_time and (dt.hour != 0 or dt.minute != 0 or dt.second != 0)):
        return dt.strftime('%Y-%m-%d')
    if behavior == 'TimeZoneIndependent':
        return dt.strftime('%Y-%m-%dT%H:%M:%S')
    # UserLocal (Dataverse's default): local wall-clock time -> UTC.
    return dt.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')


_BLANK_TEXT = ('', 'none', 'null', '-', '—', 'undefined', 'n/a', 'na', '#n/a')


def _parse_number_strict(val: Any) -> Optional[float]:
    """A number written as money/percent text, or None if it is not one.

    Accepts currency symbols, thousands separators, spaces, a trailing % and
    accounting negatives '(1,234.00)'. Anything else (e.g. '1.2M', 'N/A') is
    not a number; the caller decides whether that is an error.
    """
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        if math.isnan(val) or math.isinf(val):
            return None
        return float(val)
    s = str(val).strip()
    negative = False
    if s.startswith('(') and s.endswith(')'):
        negative, s = True, s[1:-1].strip()
    s = re.sub(r'[₱$\s,]', '', s)
    if s.endswith('%'):
        s = s[:-1]
    if s.startswith('(') and s.endswith(')'):
        negative, s = True, s[1:-1]
    if not re.fullmatch(r'[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?', s):
        return None
    number = float(s)
    return -abs(number) if negative else number


class ConcurrencyError(RuntimeError):
    """The record changed in Dataverse after it was read (HTTP 412)."""


class DataverseClient:
    def __init__(self, config: Optional[dict] = None):
        self.config = config or load_config()
        self.env = self.config['environment_url'].rstrip('/')
        self.api_root = f'{self.env}/api/data/{API_VERSION}/'
        self.scope = f'{self.env}/.default'
        self._token: Optional[str] = None
        self._token_expiry = 0.0
        self._app: Optional[msal.PublicClientApplication] = None
        self._cache: Optional[msal.SerializableTokenCache] = None
        self._pending_flow: Optional[dict] = None
        self._meta_cache: Dict[str, dict] = {}
        self._datetime_behavior_cache: Dict[str, Dict[str, str]] = {}
        self._entity_info_cache: Dict[str, dict] = {}
        self._entity_defs_cache: Optional[List[dict]] = None
        self._solution_tables_cache: Dict[str, List[dict]] = {}
        self._picklist_cache: Dict[str, Dict[str, List[dict]]] = {}

        self._session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(pool_connections=8, pool_maxsize=16, max_retries=2)
        self._session.mount('https://', adapter)
        self._session.mount('http://', adapter)

    # -- Authentication --------------------------------------------------------
    def _msal_app(self) -> msal.PublicClientApplication:
        if msal is None:
            raise RuntimeError('msal library is not installed.')
        if self._app is None:
            self._cache = msal.SerializableTokenCache()
            cache_file = _token_cache_path()
            if os.path.exists(cache_file):
                try:
                    with open(cache_file, 'r', encoding='utf-8') as f:
                        self._cache.deserialize(f.read())
                except Exception:
                    pass
            authority = f"https://login.microsoftonline.com/{self.config['tenant_id']}"
            self._app = msal.PublicClientApplication(
                self.config['client_id'],
                authority=authority,
                token_cache=self._cache
            )
        return self._app

    def _save_cache(self) -> None:
        try:
            if self._cache is not None and self._cache.has_state_changed:
                with open(_token_cache_path(), 'w', encoding='utf-8') as f:
                    f.write(self._cache.serialize())
        except Exception:
            pass

    def try_silent_token(self) -> bool:
        """Attempt to acquire a token silently from local cache."""
        try:
            app = self._msal_app()
            accounts = app.get_accounts()
            if not accounts:
                return False
            result = app.acquire_token_silent([self.scope], account=accounts[0])
            self._save_cache()
            if result and 'access_token' in result:
                self._store_token(result)
                return True
        except Exception:
            pass
        return False

    def signed_in(self) -> bool:
        """Check if currently signed in with a valid or refreshable token."""
        if self._token and time.time() < self._token_expiry:
            return True
        return self.try_silent_token()

    def begin_device_flow(self) -> dict:
        """Start interactive device code flow. Returns {user_code, verification_uri, message}."""
        app = self._msal_app()
        flow = app.initiate_device_flow(scopes=[self.scope])
        if not flow or not isinstance(flow, dict):
            raise RuntimeError('Failed to initiate device flow with identity provider.')
        if 'user_code' not in flow:
            err = flow.get('error_description') or flow.get('error') or 'Failed to start device flow.'
            raise RuntimeError(err)
        self._pending_flow = flow
        return {
            'user_code': flow.get('user_code', ''),
            'verification_uri': flow.get('verification_uri', ''),
            'message': flow.get('message', '')
        }

    def complete_device_flow(self, flow: Optional[dict] = None) -> dict:
        """Poll and wait for user to authenticate via device code flow."""
        target_flow = getattr(self, '_pending_flow', None) or flow
        if not target_flow:
            raise RuntimeError('No active device code flow to complete.')
        app = self._msal_app()
        result = app.acquire_token_by_device_flow(target_flow)
        self._save_cache()
        if result and isinstance(result, dict) and 'access_token' in result:
            self._store_token(result)
            return {'success': True, 'account': result.get('id_token_claims', {})}
        err = (result.get('error_description') or result.get('error')) if isinstance(result, dict) else 'Authentication failed'
        raise RuntimeError(err)

    def sign_out(self) -> None:
        """Clear cached tokens."""
        self._token = None
        self._token_expiry = 0
        try:
            cache_file = _token_cache_path()
            if os.path.exists(cache_file):
                os.remove(cache_file)
            self._app = None
            self._cache = None
        except Exception:
            pass

    def current_user_name(self) -> str:
        """Signed-in account's user name (e-mail), or '' when unknown."""
        try:
            accounts = self._msal_app().get_accounts()
            return str(accounts[0].get('username') or '') if accounts else ''
        except Exception:
            return ''

    def _store_token(self, result: dict) -> None:
        self._token = result['access_token']
        self._token_expiry = time.time() + result.get('expires_in', 3600) - 60

    def token(self) -> str:
        """Get valid access token, refreshing silently if expired."""
        if self._token and time.time() < self._token_expiry:
            return self._token
        if self.try_silent_token():
            return self._token  # type: ignore
        raise PermissionError('Not signed in to Dataverse. Please authenticate first.')

    # -- HTTP helpers ----------------------------------------------------------
    def _headers(self, with_formatting: bool = True) -> dict:
        h = {
            'Authorization': f'Bearer {self.token()}',
            'Accept': 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'Content-Type': 'application/json; charset=utf-8',
        }
        if with_formatting:
            h['Prefer'] = 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
        return h

    # Dataverse service protection: 429 (too many requests) or 503 means the
    # request was NOT processed, so it is safe to wait and send it again, even
    # for a POST. Retry-After says how long to wait.
    _THROTTLE_STATUSES = (429, 503)
    _THROTTLE_ATTEMPTS = 6

    def _throttle_wait(self, resp, attempt: int) -> float:
        try:
            retry_after = float(resp.headers.get('Retry-After'))
        except (TypeError, ValueError, AttributeError):
            retry_after = 0.0
        return min(60.0, max(retry_after, 2.0 * (2 ** attempt)))

    def _send(self, method: str, url: str, **kwargs):
        """One HTTP request, retried while Dataverse answers 429/503."""
        sender = getattr(self._session, method)
        for attempt in range(self._THROTTLE_ATTEMPTS):
            resp = sender(url, **kwargs)
            status = getattr(resp, 'status_code', None)
            if isinstance(status, int) and status in self._THROTTLE_STATUSES and attempt < self._THROTTLE_ATTEMPTS - 1:
                time.sleep(self._throttle_wait(resp, attempt))
                continue
            return resp
        return resp

    def get(self, url_or_path: str, params: Optional[dict] = None, with_formatting: bool = True) -> dict:
        url = url_or_path if url_or_path.startswith('http') else f'{self.api_root}{url_or_path.lstrip("/")}'
        resp = self._send('get', url, headers=self._headers(with_formatting), params=params, timeout=_TIMEOUT)
        if not resp.ok:
            raise RuntimeError(f"Dataverse GET {url} failed ({resp.status_code}): {resp.text}")
        return resp.json()

    def post(self, entity_set: str, payload: dict) -> Tuple[int, str, dict]:
        url = f'{self.api_root}{entity_set.lstrip("/")}'
        resp = self._send('post', url, headers=self._headers(False), json=payload, timeout=_TIMEOUT)
        if not (200 <= resp.status_code < 300):
            raise RuntimeError(f"Dataverse POST {entity_set} failed ({resp.status_code}): {resp.text}")
        entity_id = resp.headers.get('OData-EntityId', '')
        created_guid = ""
        if entity_id:
            m = re.search(r'\(([^)]+)\)', entity_id)
            if m:
                created_guid = m.group(1).strip("'\"")
        body = resp.json() if resp.content and resp.status_code != 204 else {}
        return resp.status_code, created_guid, body

    def patch(self, entity_set: str, record_id: str, payload: dict, expected_etag: Optional[str] = None) -> int:
        return self.patch_with_version(entity_set, record_id, payload, expected_etag=expected_etag)[0]

    def patch_with_version(self, entity_set: str, record_id: str, payload: dict,
                           expected_etag: Optional[str] = None,
                           select_attr: Optional[str] = None) -> Tuple[int, str]:
        """PATCH a record; returns (status, new etag or '').

        expected_etag: sent as If-Match, so the save fails (412) when someone
        else changed the record since it was read. select_attr: ask Dataverse to
        return the updated record's version (Prefer: return=representation),
        so the app can make its next conditional save without another read.
        """
        clean_id = record_id.strip("{}'\"")
        url = f'{self.api_root}{entity_set.lstrip("/")}({clean_id})'
        if select_attr:
            url += f'?$select={select_attr}'
        # Dataverse sometimes closes a connection before returning the PATCH
        # response (RemoteDisconnected / Connection aborted). PATCH replaces
        # the same field values, so retrying this request is idempotent. Do not
        # apply this retry to POST audit writes, where it could duplicate logs.
        retryable_errors = (requests.exceptions.ConnectionError, requests.exceptions.Timeout)
        resp = None
        headers = self._headers(False)
        if expected_etag:
            headers['If-Match'] = expected_etag
        if select_attr:
            headers['Prefer'] = 'return=representation'
        attempts = 1 if expected_etag else 3
        for attempt in range(attempts):
            try:
                resp = self._send('patch', url, headers=headers, json=payload, timeout=_TIMEOUT)
                break
            except retryable_errors:
                if attempt == attempts - 1:
                    raise
                time.sleep(0.5 * (2 ** attempt))
        if resp.status_code == 412:
            raise ConcurrencyError('This record was changed in Dataverse by someone else since it was loaded. Reload and try again.')
        if not (200 <= resp.status_code < 300):
            raise RuntimeError(f"Dataverse PATCH {entity_set}({record_id}) failed ({resp.status_code}): {resp.text}")
        new_etag = ''
        if select_attr and resp.status_code == 200:
            try:
                body = resp.json()
                candidate = body.get('@odata.etag') if isinstance(body, dict) else ''
                new_etag = candidate if isinstance(candidate, str) else ''
            except Exception:
                new_etag = ''
        if not new_etag and resp is not None:
            header_etag = resp.headers.get('ETag') or resp.headers.get('etag') or ''
            if header_etag and isinstance(header_etag, str):
                new_etag = header_etag
        return resp.status_code, new_etag

    def delete(self, entity_set: str, record_id: str, expected_etag: Optional[str] = None) -> int:
        """Delete one row. With expected_etag (If-Match) the delete fails with
        ConcurrencyError when the row changed after it was read."""
        clean_id = record_id.strip("{}'\"")
        url = f'{self.api_root}{entity_set.lstrip("/")}({clean_id})'
        headers = self._headers(False)
        if expected_etag:
            headers['If-Match'] = expected_etag
        resp = self._send('delete', url, headers=headers, timeout=_TIMEOUT)
        if resp.status_code == 412:
            raise ConcurrencyError('This record was changed in Dataverse by someone else since it was loaded. Reload and try again.')
        if not (200 <= resp.status_code < 300):
            raise RuntimeError(f"Dataverse DELETE {entity_set}({record_id}) failed ({resp.status_code}): {resp.text}")
        return resp.status_code

    # -- Metadata Discovery ----------------------------------------------------
    def get_entity_info(self, logical_name: str) -> dict:
        """Resolve entity set name and primary key attribute for an entity."""
        clean = logical_name.strip().lower()
        if clean in self._entity_info_cache:
            return self._entity_info_cache[clean]

        url = f"EntityDefinitions(LogicalName='{clean}')"
        data = self.get(url, params={'$select': 'EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,DisplayName,LogicalName'})
        info = {
            'logical_name': clean,
            'entity_set': data.get('EntitySetName', f"{clean}s"),
            'primary_id': data.get('PrimaryIdAttribute', f"{clean}id"),
            'primary_name': data.get('PrimaryNameAttribute', f"{clean}name"),
            'display_name': _extract_label(data.get('DisplayName'), clean)
        }
        self._entity_info_cache[clean] = info
        return info

    def get_datetime_behaviors(self, logical_name: str) -> Dict[str, str]:
        """Map each DateTime column to its behavior: UserLocal, DateOnly or TimeZoneIndependent."""
        clean = logical_name.strip().lower()
        if clean in self._datetime_behavior_cache:
            return self._datetime_behavior_cache[clean]
        behaviors: Dict[str, str] = {}
        try:
            url = (f"EntityDefinitions(LogicalName='{clean}')/Attributes/"
                   f"Microsoft.Dynamics.CRM.DateTimeAttributeMetadata")
            data = self.get(url, params={'$select': 'LogicalName,DateTimeBehavior'}, with_formatting=False)
            for attr in data.get('value', []):
                value = (attr.get('DateTimeBehavior') or {}).get('Value')
                if attr.get('LogicalName') and value:
                    behaviors[attr['LogicalName']] = value
        except Exception:
            # Unknown behavior falls back to UserLocal, Dataverse's default.
            behaviors = {}
        self._datetime_behavior_cache[clean] = behaviors
        return behaviors

    def get_entity_attributes(self, logical_name: str) -> Dict[str, dict]:
        """Fetch all attributes and their display names for an entity."""
        clean = logical_name.strip().lower()
        if clean in self._meta_cache:
            return self._meta_cache[clean]

        url = f"EntityDefinitions(LogicalName='{clean}')/Attributes"
        data = self.get(url, params={
            '$select': 'LogicalName,DisplayName,AttributeType,SchemaName,IsValidForCreate,IsValidForUpdate,IsValidForRead'
        })
        attr_map = {}
        for a in data.get('value', []):
            log_name = a.get('LogicalName')
            disp = _extract_label(a.get('DisplayName'), log_name)
            attr_map[log_name] = {
                'logical': log_name,
                'schema': a.get('SchemaName', log_name),
                'display': disp or log_name,
                'type': a.get('AttributeType'),
                'can_create': a.get('IsValidForCreate', True),
                'can_update': a.get('IsValidForUpdate', True),
                'can_read': a.get('IsValidForRead', True),
            }
        self._meta_cache[clean] = attr_map
        return attr_map

    def create_action_type_column(self, logical_name: str = 'cr63f_loamonitoringlogs') -> dict:
        """Add the approved text Action Type column if the audit table lacks it."""
        table = logical_name.strip().lower()
        try:
            attrs = self.get_entity_attributes(table)
            if 'cr63f_actiontype' in attrs:
                return {'created': False, 'logical_name': 'cr63f_actiontype'}
        except Exception:
            pass

        schema_timeout = 180
        label = {'@odata.type': 'Microsoft.Dynamics.CRM.Label',
                 'LocalizedLabels': [{'@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel', 'Label': 'Action Type', 'LanguageCode': 1033}],
                 'UserLocalizedLabel': {'Label': 'Action Type', 'LanguageCode': 1033}}
        payload = {
            '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata',
            'AttributeType': 'String',
            'AttributeTypeName': {'Value': 'StringType'},
            'SchemaName': 'cr63f_ActionType',
            'DisplayName': label,
            'RequiredLevel': {
                'Value': 'None', 'CanBeChanged': True,
                'ManagedPropertyLogicalName': 'canmodifyrequirementlevelsettings',
            },
            'MaxLength': 100,
            'FormatName': {'Value': 'Text'},
        }
        payload['DisplayName']['LocalizedLabels'][0]['@odata.type'] = 'Microsoft.Dynamics.CRM.LocalizedLabel'
        payload['DisplayName']['UserLocalizedLabel']['@odata.type'] = 'Microsoft.Dynamics.CRM.LocalizedLabel'
        path = f"EntityDefinitions(LogicalName='{table}')/Attributes"
        url = f'{self.api_root}{path}'

        try:
            response = self._session.post(
                url, headers=self._headers(False), json=payload, timeout=schema_timeout
            )
            if not (200 <= response.status_code < 300):
                if 'already exists' in response.text.lower() or 'duplicate' in response.text.lower() or '0x80040203' in response.text.lower():
                    self._meta_cache.pop(table, None)
                    return {'created': False, 'logical_name': 'cr63f_actiontype'}
                raise RuntimeError(
                    f"Dataverse could not add Action Type to {table} "
                    f"({response.status_code}): {response.text}"
                )
        except requests.exceptions.ReadTimeout:
            print(f"Notice: Action Type attribute creation on {table} read timed out (in progress on server).")

        publish_xml = (
            '<importexportxml><entities><entity>'
            f'{table}'
            '</entity></entities></importexportxml>'
        )
        try:
            publish_response = self._session.post(
                f'{self.api_root}PublishXml', headers=self._headers(False),
                json={'ParameterXml': publish_xml}, timeout=schema_timeout
            )
            if not (200 <= publish_response.status_code < 300):
                print(f"Notice: PublishXml for {table} returned status {publish_response.status_code}")
        except requests.exceptions.ReadTimeout:
            print(f"Notice: PublishXml for {table} read timed out (publishing continues in background).")

        self._meta_cache.pop(table, None)
        try:
            self.get_entity_attributes(table)
        except Exception:
            pass
        return {'created': True, 'logical_name': 'cr63f_actiontype'}

    def ensure_acquisition_sublevel_column(self, logical_name: str) -> dict:
        """Ensure the optional text Acquisition Sub-Level column exists on a table."""
        table = logical_name.strip().lower()
        try:
            attrs = self.get_entity_attributes(table)
            if 'cr63f_acquisitionsublevel' in attrs:
                return {'created': False, 'logical_name': 'cr63f_acquisitionsublevel', 'table': table}
        except Exception:
            pass

        schema_timeout = 180
        display = 'Acquisition Sub-Level'
        label = {
            '@odata.type': 'Microsoft.Dynamics.CRM.Label',
            'LocalizedLabels': [{
                '@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel',
                'Label': display, 'LanguageCode': 1033,
            }],
            'UserLocalizedLabel': {
                '@odata.type': 'Microsoft.Dynamics.CRM.LocalizedLabel',
                'Label': display, 'LanguageCode': 1033,
            },
        }
        payload = {
            '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata',
            'AttributeType': 'String',
            'AttributeTypeName': {'Value': 'StringType'},
            'SchemaName': 'cr63f_AcquisitionSubLevel',
            'DisplayName': label,
            'RequiredLevel': {
                'Value': 'None', 'CanBeChanged': True,
                'ManagedPropertyLogicalName': 'canmodifyrequirementlevelsettings',
            },
            'MaxLength': 100,
            'FormatName': {'Value': 'Text'},
        }
        path = f"EntityDefinitions(LogicalName='{table}')/Attributes"

        try:
            response = self._session.post(
                f'{self.api_root}{path}', headers=self._headers(False),
                json=payload, timeout=schema_timeout,
            )
            if not (200 <= response.status_code < 300):
                if 'already exists' in response.text.lower() or 'duplicate' in response.text.lower() or '0x80040203' in response.text.lower():
                    self._meta_cache.pop(table, None)
                    return {'created': False, 'logical_name': 'cr63f_acquisitionsublevel', 'table': table}
                raise RuntimeError(
                    f"Dataverse could not add Acquisition Sub-Level to {table} "
                    f"({response.status_code}): {response.text}"
                )
        except requests.exceptions.ReadTimeout:
            print(f"Notice: Acquisition Sub-Level attribute creation on {table} read timed out (in progress on server).")

        publish_xml = (
            '<importexportxml><entities><entity>'
            f'{table}'
            '</entity></entities></importexportxml>'
        )
        try:
            publish_response = self._session.post(
                f'{self.api_root}PublishXml', headers=self._headers(False),
                json={'ParameterXml': publish_xml}, timeout=schema_timeout,
            )
            if not (200 <= publish_response.status_code < 300):
                print(f"Notice: PublishXml for {table} returned status {publish_response.status_code}")
        except requests.exceptions.ReadTimeout:
            print(f"Notice: PublishXml for {table} read timed out (publishing continues in background).")

        self._meta_cache.pop(table, None)
        try:
            self.get_entity_attributes(table)
        except Exception:
            pass
        return {'created': True, 'logical_name': 'cr63f_acquisitionsublevel', 'table': table}

    def get_solutions(self) -> List[Dict[str, Any]]:
        """Fetch all visible solutions from the environment."""
        url = ("solutions?$select=solutionid,uniquename,friendlyname,version,ismanaged,description"
               "&$filter=isvisible eq true&$orderby=friendlyname asc")
        try:
            data = self.get(url)
            return [{
                'id': s.get('solutionid'),
                'uniqueName': s.get('uniquename'),
                'friendlyName': s.get('friendlyname') or s.get('uniquename'),
                'version': s.get('version', ''),
                'isManaged': bool(s.get('ismanaged')),
                'description': s.get('description') or '',
            } for s in data.get('value', [])]
        except Exception as exc:
            raise RuntimeError(f'Could not fetch solutions: {exc}') from exc

    def _entity_definitions(self) -> List[dict]:
        """Org-wide table metadata, cached per session."""
        if self._entity_defs_cache is not None:
            return self._entity_defs_cache
        url = "EntityDefinitions?$select=MetadataId,LogicalName,DisplayName,EntitySetName,PrimaryNameAttribute,IsCustomEntity&$filter=IsPrivate eq false"
        self._entity_defs_cache = self.get(url).get('value', [])
        return self._entity_defs_cache

    def get_solution_tables(self, solution_id: Optional[str] = None) -> List[Dict[str, Any]]:
        """Fetch tables belonging to a solution or all custom tables."""
        cache_key = str(solution_id or '')
        if cache_key in self._solution_tables_cache:
            return self._solution_tables_cache[cache_key]

        table_meta_ids = set()
        resolved_sol_id = solution_id
        if solution_id:
            if len(solution_id) != 36 or '-' not in solution_id:
                try:
                    sol_literal = "'" + str(solution_id).replace("'", "''") + "'"
                    sol_url = f"solutions?$select=solutionid,uniquename&$filter=uniquename eq {sol_literal} or friendlyname eq {sol_literal}"
                    sols = self.get(sol_url).get('value', [])
                    if sols:
                        resolved_sol_id = sols[0].get('solutionid')
                except Exception:
                    pass

            try:
                comp_url = f"solutions({resolved_sol_id})/solution_solutioncomponent?$filter=componenttype eq 1&$select=objectid"
                comp_data = self.get(comp_url)
                for comp in comp_data.get('value', []):
                    oid = comp.get('objectid')
                    if oid:
                        table_meta_ids.add(str(oid).lower())
            except Exception:
                pass

        entity_defs = self._entity_definitions()
        tables = []
        for ent in entity_defs:
            meta_id = str(ent.get('MetadataId', '')).lower()
            if table_meta_ids and meta_id not in table_meta_ids:
                continue
            logical = ent.get('LogicalName')
            disp = _extract_label(ent.get('DisplayName'), logical)
            tables.append({
                'metadataId': meta_id,
                'logicalName': logical,
                'displayName': disp or logical,
                'entitySetName': ent.get('EntitySetName') or f"{logical}s",
                'isCustom': bool(ent.get('IsCustomEntity')),
                'primaryName': ent.get('PrimaryNameAttribute') or '',
            })

        if table_meta_ids and not tables:
            # Fallback to custom / cr63f_ tables
            for ent in entity_defs:
                logical = ent.get('LogicalName', '')
                if ent.get('IsCustomEntity') or logical.startswith('cr63f_'):
                    disp = _extract_label(ent.get('DisplayName'), logical)
                    tables.append({
                        'metadataId': str(ent.get('MetadataId', '')).lower(),
                        'logicalName': logical,
                        'displayName': disp or logical,
                        'entitySetName': ent.get('EntitySetName') or f"{logical}s",
                        'isCustom': bool(ent.get('IsCustomEntity')),
                        'primaryName': ent.get('PrimaryNameAttribute') or '',
                    })

        tables.sort(key=lambda t: (not t['isCustom'], t['displayName'].lower()))
        self._solution_tables_cache[cache_key] = tables
        return tables

    def get_picklist_options(self, logical_name: str) -> Dict[str, List[Dict[str, Any]]]:
        """Fetch option labels and values for picklist/choice attributes."""
        clean = logical_name.strip().lower()
        if clean in self._picklist_cache:
            return self._picklist_cache[clean]

        out: Dict[str, List[Dict[str, Any]]] = {}
        try:
            url = f"EntityDefinitions(LogicalName='{clean}')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options,IsGlobal),GlobalOptionSet($select=Options)"
            data = self.get(url)
            for item in data.get('value', []):
                attr_log = item.get('LogicalName')
                options_raw = (item.get('OptionSet') or {}).get('Options') or (item.get('GlobalOptionSet') or {}).get('Options') or []
                opt_list = []
                for o in options_raw:
                    val = o.get('Value')
                    lbl = _extract_label(o.get('Label'), str(val))
                    opt_list.append({'value': val, 'label': lbl})
                if attr_log and opt_list:
                    out[attr_log] = opt_list
        except Exception:
            pass

        self._picklist_cache[clean] = out
        return out

    # -- High-level Operations -------------------------------------------------
    def fetch_records(self, logical_name: str, max_records: Optional[int] = None, custom_only: bool = False, filter_query: Optional[str] = None) -> List[Dict[str, Any]]:
        """Fetch all records from a Dataverse table and format into clean dictionaries."""
        info = self.get_entity_info(logical_name)
        attrs = self.get_entity_attributes(logical_name)
        entity_set = info['entity_set']
        primary_id = info['primary_id']

        clean_tbl = logical_name.strip().lower()
        system_cols = {
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
        }

        if custom_only:
            filtered_attrs = {}
            for log_name, meta in attrs.items():
                low = log_name.lower()
                if low in system_cols:
                    continue
                if low == f"{clean_tbl}id":
                    continue
                if low.endswith(('_base', 'yominame', 'name')) and not meta.get('is_custom'):
                    continue
                if meta.get('type') in ('Virtual', 'PartyList'):
                    continue
                if meta.get('is_custom') or any(low.startswith(p) for p in ('cr63f_', 'new_')):
                    filtered_attrs[log_name] = meta
            attrs = filtered_attrs if filtered_attrs else attrs

        url = f"{self.api_root}{entity_set}"
        if filter_query:
            sep = '&' if '?' in url else '?'
            # Encode the filter so values containing &, #, + or % (e.g. titles
            # like 'T-1 & T-2') cannot end or alter the query string.
            encoded_filter = quote(filter_query, safe="'(),=")
            url = f"{url}{sep}$filter={encoded_filter}"

        all_records: List[Dict[str, Any]] = []
        headers = self._headers(with_formatting=True)
        headers['Prefer'] = 'odata.maxpagesize=5000, odata.include-annotations="OData.Community.Display.V1.FormattedValue"'

        while url:
            if not url.startswith('http'):
                url = f"{self.api_root}{url.lstrip('/')}"
            resp = self._send('get', url, headers=headers, timeout=_TIMEOUT)
            if not resp.ok:
                raise RuntimeError(f"Dataverse GET {url} failed ({resp.status_code}): {resp.text}")
            data = resp.json()
            values = data.get('value', [])
            for row in values:
                rec_id = row.get(primary_id, '') or row.get(f"{clean_tbl}id", '') or row.get('id', '')
                rec: Dict[str, Any] = {
                    '__id': rec_id,
                    '__etag': row.get('@odata.etag', ''),
                }
                # createdon is replaced below by its display text (minute
                # precision, month-first). Keep the raw ISO UTC value so audit
                # history can be ordered exactly.
                if row.get('createdon'):
                    rec['__createdon_raw'] = row['createdon']
                # Lookup columns (Created By, Modified By, Owner) come back as
                # _createdby_value plus a FormattedValue annotation holding the
                # name. Keep the name under the column's own name ('createdby')
                # so the Overview and the weekly report can show who saved.
                if not custom_only:
                    for k, v in row.items():
                        m = re.fullmatch(r'_(\w+)_value@OData\.Community\.Display\.V1\.FormattedValue', k)
                        if m and v:
                            rec[m.group(1)] = v
                # Store all row attributes
                for k, v in row.items():
                    if not k.startswith('@') and '@' not in k:
                        if custom_only:
                            low_k = k.lower().strip()
                            if k.startswith('_') or low_k.endswith('_value') or low_k in system_cols:
                                continue
                        rec[k] = v if v is not None else ''
                # Store formatted and display name attributes
                for log_name, meta in attrs.items():
                    disp = meta['display']
                    fmt_key = f"{log_name}@OData.Community.Display.V1.FormattedValue"
                    val = row.get(fmt_key, row.get(log_name, ''))
                    if val is not None and val != '':
                        rec[disp] = val
                        rec[log_name] = val
                    elif disp not in rec:
                        rec[disp] = ''
                if rec.get('__createdon_raw'):
                    raw_dt = parse_log_datetime(rec['__createdon_raw'])
                    if raw_dt:
                        local_disp = raw_dt.strftime('%m/%d/%Y %I:%M %p')
                        rec['createdon'] = local_disp
                        if 'Created On' in rec:
                            rec['Created On'] = local_disp
                all_records.append(rec)

            if max_records and len(all_records) >= max_records:
                return all_records[:max_records]

            url = data.get('@odata.nextLink', '')

        return all_records

    def get_record(self, table_name: str, record_id: str) -> Dict[str, Any]:
        """Read a single current record and retain the version for conditional writes."""
        info = self.get_entity_info(table_name)
        clean_id = str(record_id).strip("{}'\"")
        row = self.get(f"{info['entity_set']}({clean_id})")
        rec = {k: v for k, v in row.items() if '@' not in k}
        rec['__id'] = row.get(info['primary_id']) or clean_id
        rec['__etag'] = row.get('@odata.etag', '')
        # Choice labels are required by the app's comparisons and save mapper.
        # Keep numeric/date logical values raw so rollback retains full precision.
        for key, meta in self.get_entity_attributes(table_name).items():
            formatted = row.get(f'{key}@OData.Community.Display.V1.FormattedValue')
            if formatted is not None:
                rec[meta['display']] = formatted
                if meta['type'] in ('Picklist', 'State', 'Status'):
                    rec[key] = formatted
        return rec

    def save_record(self, table_name: str, record_data: Dict[str, Any], record_id: Optional[str] = None,
                    expected_etag: Optional[str] = None, strict: bool = True) -> Dict[str, Any]:
        """Create or update a record in any Dataverse table with automatic field mapping & casting.

        strict=True: a number, whole number or date that cannot be read raises
        ValueError instead of being saved as blank. The audit table uses
        strict=False, because a snapshot may carry display text in a typed column.
        """
        info = self.get_entity_info(table_name)
        attrs = self.get_entity_attributes(table_name)

        # Build payload mapping display names / logical names to valid Dataverse attributes
        payload: Dict[str, Any] = {}
        for k, v in record_data.items():
            if k.startswith('__') or k.startswith('@'):
                continue
            # find matching logical attribute
            target_attr = None
            if k in attrs and (attrs[k]['can_create'] or attrs[k]['can_update']):
                target_attr = attrs[k]
            else:
                k_norm = _norm(k)
                for log_name, meta in attrs.items():
                    if _norm(meta['display']) == k_norm or _norm(log_name) == k_norm:
                        if meta['can_create'] or meta['can_update']:
                            target_attr = meta
                            break

                # Prefix-stripped clean match (e.g. new_contracttype vs cr63f_contracttype vs Contract Type)
                if not target_attr:
                    k_clean = _clean_key(k)
                    if k_clean:
                        # Pass 1: exact clean match across all attributes
                        for log_name, meta in attrs.items():
                            log_clean = _clean_key(log_name)
                            disp_clean = _clean_key(meta['display'])
                            if log_clean == k_clean or disp_clean == k_clean:
                                if meta['can_create'] or meta['can_update']:
                                    target_attr = meta
                                    break
                        # Pass 2: fallback to number-stripped match only if no exact clean match
                        if not target_attr:
                            for log_name, meta in attrs.items():
                                log_clean = _clean_key(log_name)
                                disp_clean = _clean_key(meta['display'])
                                if (log_clean.rstrip('1234567890') == k_clean.rstrip('1234567890') or
                                    disp_clean.rstrip('1234567890') == k_clean.rstrip('1234567890')):
                                    if meta['can_create'] or meta['can_update']:
                                        target_attr = meta
                                        break

            if target_attr:
                log_name = target_attr['logical']
                attr_type = target_attr['type']
                # Cast based on type
                if v == '' or v is None:
                    if log_name not in payload:
                        payload[log_name] = None
                elif str(v).strip().lower() in _BLANK_TEXT and attr_type in ('Integer', 'BigInt', 'Decimal', 'Double', 'Money', 'DateTime'):
                    # An explicit placeholder ('-', 'None') means blank.
                    if log_name not in payload:
                        payload[log_name] = None
                elif attr_type in ('Integer', 'BigInt', 'Decimal', 'Double', 'Money'):
                    number = _parse_number_strict(v)
                    if number is None:
                        if strict:
                            raise ValueError(f"'{v}' is not a valid number for {target_attr['display']}.")
                        if log_name not in payload:
                            payload[log_name] = None
                    else:
                        payload[log_name] = int(round(number)) if attr_type in ('Integer', 'BigInt') else number
                elif attr_type in ('Picklist', 'State', 'Status'):
                    try:
                        payload[log_name] = int(str(v).strip())
                    except Exception:
                        # If passed a label, look up choice integer value
                        opts = self.get_picklist_options(table_name).get(log_name, [])
                        v_norm = _norm(v)
                        v_clean = _clean_key(v)
                        # Exact label first; a partial match is used only when it is
                        # unique, so 'LEASE' can never resolve to 'SUBLEASE'.
                        matched_opt = next((o['value'] for o in opts if _norm(o['label']) == v_norm), None)
                        if matched_opt is None:
                            matched_opt = next((o['value'] for o in opts if _clean_key(o['label']) == v_clean), None)
                        if matched_opt is None and v_norm:
                            partial = [o['value'] for o in opts if v_norm in _norm(o['label'])]
                            matched_opt = partial[0] if len(partial) == 1 else None
                        if matched_opt is not None:
                            payload[log_name] = matched_opt
                        else:
                            # An unmatched label must not be dropped silently: the
                            # caller would be told the save succeeded without it.
                            raise ValueError(f"'{v}' is not a valid choice for {log_name}")
                elif attr_type == 'DateTime':
                    clean_dt = _clean_dataverse_datetime(v, self.get_datetime_behaviors(table_name).get(log_name, 'UserLocal'))
                    if clean_dt is not None:
                        payload[log_name] = clean_dt
                    else:
                        text = str(v).strip()
                        # 0001-01-01 style placeholders are blanks on purpose.
                        if strict and not re.search(r'(^|\D)0{3}[01](\D|$)', text):
                            raise ValueError(f"'{v}' is not a valid date for {target_attr['display']} (use MM/DD/YYYY).")
                        if log_name not in payload:
                            payload[log_name] = None
                elif attr_type == 'Boolean':
                    payload[log_name] = bool(v) if str(v).lower() not in ('false', '0', 'no', '') else False
                else:
                    payload[log_name] = str(v).strip()

        # Primary name attribute safeguard on create
        primary_name = info.get('primary_name')
        if not record_id and primary_name and (primary_name not in payload or not payload[primary_name]):
            def_name = (
                payload.get('cr63f_areaindex') or record_data.get('area_index') or record_data.get('AREA-INDEX') or
                record_data.get('cr63f_control1') or record_data.get('Control') or
                record_data.get('field_name') or record_data.get('details') or 'LOA Entry'
            )
            payload[primary_name] = str(def_name).strip()

        if record_id:
            # Update. The new version tag lets the next save of this record be
            # conditional too (see DataverseClient.patch_with_version).
            _, new_etag = self.patch_with_version(
                info['entity_set'], record_id, payload, expected_etag=expected_etag,
                select_attr=info.get('primary_id'))
            return {'success': True, 'action': 'update', 'id': record_id, 'etag': new_etag}
        else:
            # Create
            status, created_id, _ = self.post(info['entity_set'], payload)
            return {'success': True, 'action': 'create', 'id': created_id}

    def save_loa_record(self, record_data: Dict[str, Any], record_id: Optional[str] = None, expected_etag: Optional[str] = None) -> Dict[str, Any]:
        """Create or update a record in cr63f_loamonitoring."""
        table_name = self.config.get('loa_table', 'cr63f_loamonitoring')
        if expected_etag:
            return self.save_record(table_name, record_data, record_id, expected_etag=expected_etag)
        return self.save_record(table_name, record_data, record_id)

    def save_log_record(self, record_data: Dict[str, Any], record_id: Optional[str] = None) -> Dict[str, Any]:
        """Create or update a record in cr63f_loamonitoringlogs (entry logs repository)."""
        table_name = self.config.get('logs_table', 'cr63f_loamonitoringlogs')
        return self.save_record(table_name, record_data, record_id, strict=False)

    def save_finance_record(self, record_data: Dict[str, Any], record_id: Optional[str] = None) -> Dict[str, Any]:
        """Create or update a record in cr63f_financereference."""
        table_name = self.config.get('finance_table', 'cr63f_financereference')
        return self.save_record(table_name, record_data, record_id)

    def save_conversion_record(self, record_data: Dict[str, Any], record_id: Optional[str] = None) -> Dict[str, Any]:
        """Create or update a record in cr63f_landconversionreference."""
        table_name = self.config.get('conversion_table', 'cr63f_landconversionreference')
        return self.save_record(table_name, record_data, record_id)

    def delete_log_record(self, record_id: str) -> Dict[str, Any]:
        """Delete a record from cr63f_loamonitoringlogs."""
        target_str = str(record_id or '').strip()
        guid_match = re.search(r'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}', target_str)
        if not guid_match:
            return {'success': True, 'action': 'skipped', 'id': target_str, 'message': 'Non-GUID log ID skipped in Dataverse'}
        clean_guid = guid_match.group(0)
        table_name = self.config.get('logs_table', 'cr63f_loamonitoringlogs')
        info = self.get_entity_info(table_name)
        self.delete(info['entity_set'], clean_guid)
        return {'success': True, 'action': 'delete', 'id': clean_guid}

    def delete_loa_record(self, record_id: str) -> Dict[str, Any]:
        """Delete a record from cr63f_loamonitoring (used to undo a create whose audit row failed)."""
        table_name = self.config.get('loa_table', 'cr63f_loamonitoring')
        info = self.get_entity_info(table_name)
        self.delete(info['entity_set'], record_id)
        return {'success': True, 'action': 'delete', 'id': record_id}

    def delete_conversion_record(self, record_id: str) -> Dict[str, Any]:
        """Delete a record from cr63f_landconversionreference."""
        table_name = self.config.get('conversion_table', 'cr63f_landconversionreference')
        info = self.get_entity_info(table_name)
        self.delete(info['entity_set'], record_id)
        return {'success': True, 'action': 'delete', 'id': record_id}

    def delete_finance_record(self, record_id: str) -> Dict[str, Any]:
        """Delete a record from cr63f_financereference."""
        table_name = self.config.get('finance_table', 'cr63f_financereference')
        info = self.get_entity_info(table_name)
        self.delete(info['entity_set'], record_id)
        return {'success': True, 'action': 'delete', 'id': record_id}


