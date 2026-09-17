"""Dataverse connector for LOA Monitoring Tool.

Provides authenticated MSAL device-code connection, metadata discovery, and OData
v9.2 CRUD operations for `cr63f_loamonitoring` and `APPENDED_FINANCE` tables.
"""

from __future__ import annotations

import concurrent.futures as cf
import json
import os
import re
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

import requests
from app_paths import data_file

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


def load_config() -> dict:
    path = config_path()
    if not os.path.exists(path):
        # Fallback to local config if present
        local_cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dataverse_config.json')
        if os.path.exists(local_cfg):
            path = local_cfg
        else:
            return {
                "tenant_id": "70234e6c-22dd-42ec-b3ca-fee4c6a42328",
                "client_id": "c16cfa67-0d61-46b4-8813-94197d53ce6b",
                "environment_url": "https://orgfa075889.crm.dynamics.com",
                "solution_name": "SOURCING AND NEGO (BATANGASNEGO)",
                "solution_id": "a7f41cec-5786-f111-ab0f-7ced8d3b4de7",
                "loa_table": "cr63f_loamonitoring",
                "finance_table": "cr63f_appendedfinance"
            }
    with open(path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
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
    if cfg.get('loa_table') in ('', None):
        cfg['loa_table'] = 'cr63f_loamonitoring'
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

    def get(self, url_or_path: str, params: Optional[dict] = None, with_formatting: bool = True) -> dict:
        url = url_or_path if url_or_path.startswith('http') else f'{self.api_root}{url_or_path.lstrip("/")}'
        resp = self._session.get(url, headers=self._headers(with_formatting), params=params, timeout=_TIMEOUT)
        if not resp.ok:
            raise RuntimeError(f"Dataverse GET {url} failed ({resp.status_code}): {resp.text}")
        return resp.json()

    def post(self, entity_set: str, payload: dict) -> Tuple[int, str, dict]:
        url = f'{self.api_root}{entity_set.lstrip("/")}'
        resp = self._session.post(url, headers=self._headers(False), json=payload, timeout=_TIMEOUT)
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

    def patch(self, entity_set: str, record_id: str, payload: dict) -> int:
        clean_id = record_id.strip("{}'\"")
        url = f'{self.api_root}{entity_set.lstrip("/")}({clean_id})'
        resp = self._session.patch(url, headers=self._headers(False), json=payload, timeout=_TIMEOUT)
        if not (200 <= resp.status_code < 300):
            raise RuntimeError(f"Dataverse PATCH {entity_set}({record_id}) failed ({resp.status_code}): {resp.text}")
        return resp.status_code

    def delete(self, entity_set: str, record_id: str) -> int:
        clean_id = record_id.strip("{}'\"")
        url = f'{self.api_root}{entity_set.lstrip("/")}({clean_id})'
        resp = self._session.delete(url, headers=self._headers(False), timeout=_TIMEOUT)
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
        data = self.get(url, params={'$select': 'EntitySetName,PrimaryIdAttribute,DisplayName,LogicalName'})
        info = {
            'logical_name': clean,
            'entity_set': data.get('EntitySetName', f"{clean}s"),
            'primary_id': data.get('PrimaryIdAttribute', f"{clean}id"),
            'display_name': _extract_label(data.get('DisplayName'), clean)
        }
        self._entity_info_cache[clean] = info
        return info

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

    def get_solutions(self) -> List[Dict[str, Any]]:
        """Fetch all visible solutions from the environment."""
        url = (f"solutions?$select=solutionid,uniquename,friendlyname,version,ismanaged,description"
               f"&$filter=isvisible eq true&$orderby=friendlyname asc")
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
                    sol_url = f"solutions?$select=solutionid,uniquename&$filter=uniquename eq '{solution_id}' or friendlyname eq '{solution_id}'"
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
            url = f"EntityDefinitions(LogicalName='{clean}')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)"
            data = self.get(url)
            for item in data.get('value', []):
                attr_log = item.get('LogicalName')
                options_raw = (item.get('OptionSet') or {}).get('Options') or []
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
    def fetch_records(self, logical_name: str, max_records: Optional[int] = None, custom_only: bool = False) -> List[Dict[str, Any]]:
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

        all_records: List[Dict[str, Any]] = []
        headers = self._headers(with_formatting=True)
        headers['Prefer'] = 'odata.maxpagesize=5000, odata.include-annotations="OData.Community.Display.V1.FormattedValue"'

        while url:
            if not url.startswith('http'):
                url = f"{self.api_root}{url.lstrip('/')}"
            resp = self._session.get(url, headers=headers, timeout=_TIMEOUT)
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
                all_records.append(rec)

            if max_records and len(all_records) >= max_records:
                return all_records[:max_records]

            url = data.get('@odata.nextLink', '')

        return all_records

    def save_loa_record(self, record_data: Dict[str, Any], record_id: Optional[str] = None) -> Dict[str, Any]:
        """Create or update a record in cr63f_loamonitoring."""
        table_name = self.config.get('loa_table', 'cr63f_loamonitoring')
        info = self.get_entity_info(table_name)
        attrs = self.get_entity_attributes(table_name)

        # Build payload mapping display names / logical names to valid Dataverse attributes
        payload: Dict[str, Any] = {}
        for k, v in record_data.items():
            if k.startswith('__'):
                continue
            # find matching logical attribute
            target_attr = None
            if k in attrs and (attrs[k]['can_create'] or attrs[k]['can_update']):
                target_attr = attrs[k]
            else:
                for log_name, meta in attrs.items():
                    if _norm(meta['display']) == _norm(k) or _norm(log_name) == _norm(k):
                        if meta['can_create'] or meta['can_update']:
                            target_attr = meta
                            break

            if target_attr:
                log_name = target_attr['logical']
                attr_type = target_attr['type']
                # Cast based on type
                if v == '' or v is None:
                    payload[log_name] = None
                elif attr_type in ('Integer', 'BigInt'):
                    try:
                        payload[log_name] = int(float(str(v).replace(',', '').strip()))
                    except Exception:
                        payload[log_name] = None
                elif attr_type in ('Decimal', 'Double', 'Money'):
                    try:
                        payload[log_name] = float(str(v).replace(',', '').replace('₱', '').replace('$', '').strip())
                    except Exception:
                        payload[log_name] = None
                elif attr_type in ('Picklist', 'State', 'Status'):
                    try:
                        payload[log_name] = int(str(v).strip())
                    except Exception:
                        # If passed a label, look up choice integer value
                        opts = self.get_picklist_options(table_name).get(log_name, [])
                        matched_opt = next((o['value'] for o in opts if _norm(o['label']) == _norm(v)), None)
                        payload[log_name] = matched_opt if matched_opt is not None else None
                elif attr_type == 'DateTime':
                    v_str = str(v).strip()
                    if not v_str:
                        payload[log_name] = None
                    else:
                        parts = re.split(r'[-/]', v_str.split('T')[0].split(' ')[0])
                        if len(parts) == 3:
                            if len(parts[0]) == 4:  # YYYY-MM-DD
                                payload[log_name] = f"{parts[0]}-{int(parts[1]):02d}-{int(parts[2]):02d}"
                            else:  # MM/DD/YYYY
                                payload[log_name] = f"{parts[2]}-{int(parts[0]):02d}-{int(parts[1]):02d}"
                        else:
                            payload[log_name] = v_str
                elif attr_type == 'Boolean':
                    payload[log_name] = bool(v) if str(v).lower() not in ('false', '0', 'no', '') else False
                else:
                    payload[log_name] = str(v).strip()

        if record_id:
            # Update
            self.patch(info['entity_set'], record_id, payload)
            return {'success': True, 'action': 'update', 'id': record_id}
        else:
            # Create
            status, created_id, _ = self.post(info['entity_set'], payload)
            return {'success': True, 'action': 'create', 'id': created_id}

