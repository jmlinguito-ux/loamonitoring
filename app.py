"""LOA Monitoring Tool - Desktop Application Backend.

Pywebview + Python controller providing Dataverse integration, live Excel lookup
autofill engine, explicit record logging/saving, reference data sync, and batch processing.
"""

from __future__ import annotations

import json
import os
import sys
import threading
import time
from typing import Any, Dict, List, Optional

# Ensure directory is on sys.path
_base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
if _base_dir and _base_dir not in sys.path:
    sys.path.insert(0, _base_dir)

import webview

from app_paths import data_file, ensure_user_data
import dataverse as dv
from loa_engine import (
    FinanceLookupEngine,
    format_currency,
    format_percent,
    norm_str,
    parse_float,
    read_tabular_file,
    write_excel_file,
)

APP_TITLE = "LOA Monitoring Tool"


def _settings_path() -> str:
    return data_file('app_settings.json')


def load_app_settings() -> dict:
    path = _settings_path()
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "remembered_finance_file": "",
        "recent_finance_files": [],
        "theme": "light",
        "page_size": 25,
        "table_layout": {}
    }


def save_app_settings(settings: dict) -> None:
    path = _settings_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass


class Api:
    def __init__(self):
        self.dv_client: Optional[dv.DataverseClient] = None
        self.finance_engine: FinanceLookupEngine = FinanceLookupEngine()
        self.loa_records: List[Dict[str, Any]] = []
        self.finance_rows: List[Dict[str, Any]] = []
        self.finance_source_name: str = "None"
        self.batch_rows: List[Dict[str, Any]] = []
        self.active_flow: Optional[dict] = None
        self._window: Optional[webview.Window] = None

        # Load settings
        self.settings = load_app_settings()

        # Try initializing Dataverse client silently
        self._init_dataverse()

        # Try auto-loading remembered finance file if exists
        self._auto_load_remembered_file()

    def set_window(self, window: webview.Window) -> None:
        self._window = window

    def _init_dataverse(self) -> None:
        try:
            cfg = dv.load_config()
            self.dv_client = dv.DataverseClient(cfg)
            self.dv_client.try_silent_token()
        except Exception as e:
            print(f"Dataverse init notice: {e}")

    def _auto_load_remembered_file(self) -> None:
        rem = self.settings.get('remembered_finance_file')
        if rem and os.path.isfile(rem):
            try:
                headers, rows = read_tabular_file(rem)
                self.finance_rows = rows
                self.finance_engine.load_rows(rows)
                self.finance_source_name = f"File: {os.path.basename(rem)}"
            except Exception as e:
                print(f"Could not auto-load remembered finance file: {e}")

    # -- State & Connection ----------------------------------------------------
    def get_initial_state(self) -> dict:
        is_auth = False
        if self.dv_client:
            try:
                is_auth = bool(self.dv_client.token())
            except Exception:
                is_auth = False

        cfg = self.dv_client.config if self.dv_client else dv.load_config()

        return {
            'connected': is_auth,
            'environment_url': cfg.get('environment_url', ''),
            'loa_table': cfg.get('loa_table', 'cr63f_loamonitoring'),
            'finance_table': cfg.get('finance_table', 'cr63f_financereference'),
            'solution_name': cfg.get('solution_name', 'LOA MONITORING'),
            'solution_id': cfg.get('solution_id', '34c6d4d3-b4b0-f111-aaac-7ced8d3b4de7'),
            'finance_loaded_count': len(self.finance_rows),
            'finance_source': self.finance_source_name,
            'remembered_finance_file': self.settings.get('remembered_finance_file', ''),
            'recent_finance_files': self.settings.get('recent_finance_files', []),
            'theme': self.settings.get('theme', 'light'),
            'table_layout': self.settings.get('table_layout', {}),
            'loa_record_count': len(self.loa_records),
        }

    # -- MSAL Auth -------------------------------------------------------------
    def start_sign_in(self) -> dict:
        if not self.dv_client:
            self._init_dataverse()
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse client could not be initialized.'}
        try:
            if self.dv_client.try_silent_token():
                return {'success': True, 'already_connected': True}

            flow = self.dv_client.begin_device_flow()
            self.active_flow = flow
            self._auth_status = {'in_progress': True, 'connected': False, 'error': None}

            # Run polling in background daemon thread
            def _wait_for_auth():
                try:
                    self.dv_client.complete_device_flow()
                    self.active_flow = None
                    self._auth_status = {'in_progress': False, 'connected': True, 'error': None}
                    if self._window:
                        try:
                            self._window.evaluate_js('window.onDeviceFlowComplete && window.onDeviceFlowComplete(true, null)')
                        except Exception:
                            pass
                except Exception as exc:
                    self.active_flow = None
                    self._auth_status = {'in_progress': False, 'connected': False, 'error': str(exc)}
                    if self._window:
                        try:
                            msg = json.dumps(str(exc))
                            self._window.evaluate_js(f'window.onDeviceFlowComplete && window.onDeviceFlowComplete(false, {msg})')
                        except Exception:
                            pass

            threading.Thread(target=_wait_for_auth, daemon=True).start()

            return {
                'success': True,
                'user_code': flow.get('user_code'),
                'verification_uri': flow.get('verification_uri'),
                'message': flow.get('message'),
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def check_sign_in_status(self) -> dict:
        """Called by frontend to poll status while waiting for device login."""
        if not self.dv_client:
            return {'connected': False, 'in_progress': False, 'error': None}
        is_conn = False
        try:
            is_conn = bool(self.dv_client.token())
        except Exception:
            is_conn = False
        status = getattr(self, '_auth_status', {})
        return {
            'connected': is_conn,
            'in_progress': status.get('in_progress', False),
            'error': status.get('error', None)
        }

    def complete_sign_in(self) -> dict:
        """Fallback manual completion if needed."""
        if not self.dv_client:
            return {'success': False, 'error': 'No active authentication flow.'}
        try:
            res = self.dv_client.complete_device_flow()
            self.active_flow = None
            return res
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def sign_out(self) -> dict:
        if self.dv_client:
            self.dv_client.sign_out()
        return {'success': True}

    # -- Live Finance Lookup ---------------------------------------------------
    def lookup_finance(
        self,
        type_val: Any,
        index_no: Any,
        area_index: Any,
        title_mod: Any
    ) -> dict:
        """Called live from the UI form as the user modifies any key field."""
        res = self.finance_engine.lookup(
            type_val=type_val,
            index_no=index_no,
            area_index_val=area_index,
            title_mod_val=title_mod
        )
        return res

    # -- Reference Data (APPENDED_FINANCE) -------------------------------------
    def load_finance_file_dialog(self) -> dict:
        if not self._window:
            return {'success': False, 'error': 'Window not attached.'}
        file_types = ('Excel and CSV Files (*.xlsx;*.xlsm;*.csv)', 'All files (*.*)')
        result = self._window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=file_types
        )
        if not result or not result[0]:
            return {'success': False, 'cancelled': True}

        file_path = result[0]
        return self._load_finance_file_path(file_path)

    def load_finance_file_recent(self, file_path: str) -> dict:
        return self._load_finance_file_path(file_path)

    def _load_finance_file_path(self, file_path: str) -> dict:
        try:
            headers, rows = read_tabular_file(file_path)
            self.finance_rows = rows
            self.finance_engine.load_rows(rows)
            filename = os.path.basename(file_path)
            self.finance_source_name = f"File: {filename}"

            # Save remembered path and recent list
            self.settings['remembered_finance_file'] = file_path
            recent = self.settings.get('recent_finance_files', [])
            if file_path in recent:
                recent.remove(file_path)
            recent.insert(0, file_path)
            self.settings['recent_finance_files'] = recent[:6]
            save_app_settings(self.settings)

            return {
                'success': True,
                'row_count': len(rows),
                'headers': headers,
                'source': self.finance_source_name,
                'file_path': file_path,
                'sample_rows': rows
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def sync_finance_from_dataverse(self, table_name: Optional[str] = None) -> dict:
        if not self.dv_client:
            self._init_dataverse()
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse client not configured.'}
        try:
            target_table = table_name or self.dv_client.config.get('finance_table') or 'cr63f_financereference'
            if target_table in ('cr63f_appendedfinance', '', None):
                target_table = 'cr63f_financereference'
                self.dv_client.config['finance_table'] = target_table
                try:
                    dv.save_config(self.dv_client.config)
                except Exception:
                    pass

            try:
                rows = self.dv_client.fetch_records(target_table, custom_only=True)
            except Exception as fe:
                if target_table != 'cr63f_financereference':
                    target_table = 'cr63f_financereference'
                    self.dv_client.config['finance_table'] = target_table
                    try:
                        dv.save_config(self.dv_client.config)
                    except Exception:
                        pass
                    rows = self.dv_client.fetch_records(target_table, custom_only=True)
                else:
                    raise fe

            # Normalize % Paid and format currency for custom finance rows
            for r in rows:
                for k in list(r.keys()):
                    ku = k.upper()
                    if '% PAID' in ku or 'PAID' in ku:
                        r[k] = format_percent(r[k])

            self.finance_rows = rows
            self.finance_engine.load_rows(rows)
            self.finance_source_name = f"Dataverse: {target_table}"

            # Extract clean, non-duplicate custom column headers
            fin_attrs = self.dv_client.get_entity_attributes(target_table)
            clean_tbl = target_table.strip().lower()
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
            clean_headers = []
            clean_cols = []
            for log_name, meta in fin_attrs.items():
                low = log_name.lower()
                if low in system_cols or low == f"{clean_tbl}id" or low.endswith('_base') or meta.get('type') in ('Virtual', 'PartyList'):
                    continue
                if meta.get('is_custom') or any(low.startswith(p) for p in ('cr63f_', 'new_')):
                    disp = meta.get('display') or log_name
                    if disp not in clean_headers:
                        clean_headers.append(disp)
                        clean_cols.append({
                            'key': log_name,
                            'label': disp,
                            'isMoney': any(m in low for m in ('price', 'deductible', 'tcp')),
                            'isPercent': 'paid' in low or '%' in disp
                        })

            if not clean_headers and rows:
                clean_headers = [k for k in rows[0].keys() if not k.startswith('__') and not k.startswith('cr63f_')]

            return {
                'success': True,
                'row_count': len(rows),
                'source': self.finance_source_name,
                'table_name': target_table,
                'headers': clean_headers,
                'columns': clean_cols,
                'sample_rows': rows
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def update_record_cell(self, record_id: str, field_name: str, new_value: Any) -> dict:
        """Update a single cell directly in Dataverse upon blur/enter or batch."""
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        try:
            res = self.dv_client.save_loa_record({field_name: new_value}, record_id=record_id)
            # Update memory
            for r in self.loa_records:
                if r.get('__id') == record_id:
                    r[field_name] = new_value
                    break
            return {'success': True, 'result': res}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def save_staged_records(self, staged_records: list) -> dict:
        """Save a list of modified records to Dataverse."""
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        saved_count = 0
        errors = []
        try:
            for item in staged_records:
                rec_id = item.get('id')
                fields = item.get('fields', {})
                if not fields:
                    continue
                try:
                    self.dv_client.save_loa_record(fields, record_id=rec_id)
                    saved_count += 1
                    for r in self.loa_records:
                        if r.get('__id') == rec_id:
                            r.update(fields)
                            break
                except Exception as ex:
                    errors.append(f"Record {rec_id}: {ex}")

            if errors:
                return {'success': False, 'error': f"Saved {saved_count} records. Errors: {'; '.join(errors[:3])}"}
            return {'success': True, 'saved_count': saved_count, 'message': f"Successfully saved {saved_count} records to Dataverse."}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def clear_finance_reference(self) -> dict:
        self.finance_rows = []
        self.finance_engine = FinanceLookupEngine()
        self.finance_source_name = "None"
        self.settings['remembered_finance_file'] = ""
        save_app_settings(self.settings)
        return {'success': True}

    def delete_finance_record(self, record_id_or_idx: Any) -> dict:
        try:
            target_id = str(record_id_or_idx)
            found_idx = -1
            found_row = None

            # Check if record_id_or_idx matches __id
            for idx, r in enumerate(self.finance_rows):
                if r.get('__id') == target_id:
                    found_idx = idx
                    found_row = r
                    break

            # If not found by __id, check if it's integer index
            if found_idx == -1:
                try:
                    num_idx = int(record_id_or_idx)
                    if 0 <= num_idx < len(self.finance_rows):
                        found_idx = num_idx
                        found_row = self.finance_rows[num_idx]
                except (ValueError, TypeError):
                    pass

            if found_row is None or found_idx == -1:
                return {'success': False, 'error': 'Finance record not found.'}

            # If from Dataverse and has __id, delete in Dataverse
            if self.dv_client and found_row.get('__id'):
                try:
                    target_table = self.dv_client.config.get('finance_table', 'cr63f_financereference')
                    info = self.dv_client.get_entity_info(target_table)
                    self.dv_client.delete(info['entity_set'], found_row['__id'])
                except Exception as de:
                    print(f"Notice: Could not delete from Dataverse: {de}")

            # Remove from local memory & reload lookup engine
            self.finance_rows.pop(found_idx)
            self.finance_engine.load_rows(self.finance_rows)

            return {
                'success': True,
                'remaining_count': len(self.finance_rows),
                'message': 'Finance record deleted.'
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def update_finance_cell(self, record_id_or_idx: Any, field_name: str, new_value: Any) -> dict:
        """Direct inline edit of a finance cell in table."""
        try:
            target_id = str(record_id_or_idx)
            found_idx = -1
            found_row = None

            for idx, r in enumerate(self.finance_rows):
                if r.get('__id') == target_id:
                    found_idx = idx
                    found_row = r
                    break

            if found_idx == -1:
                try:
                    num_idx = int(record_id_or_idx)
                    if 0 <= num_idx < len(self.finance_rows):
                        found_idx = num_idx
                        found_row = self.finance_rows[num_idx]
                except (ValueError, TypeError):
                    pass

            if found_row is None or found_idx == -1:
                return {'success': False, 'error': 'Finance row not found.'}

            self.finance_rows[found_idx][field_name] = new_value
            # Reload engine so live lookup calculations use updated values
            self.finance_engine.load_rows(self.finance_rows)

            # If row has __id and Dataverse is connected, update in Dataverse
            if self.dv_client and found_row.get('__id'):
                try:
                    target_table = self.dv_client.config.get('finance_table', 'cr63f_financereference')
                    info = self.dv_client.get_entity_info(target_table)
                    self.dv_client.patch(info['entity_set'], found_row['__id'], {field_name: new_value})
                except Exception as de:
                    print(f"Notice: Could not patch Dataverse finance record: {de}")

            return {'success': True, 'field': field_name, 'value': new_value}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def get_finance_preview(self, limit: Optional[int] = None) -> dict:
        rows = self.finance_rows if limit is None else self.finance_rows[:limit]
        return {
            'source': self.finance_source_name,
            'total_count': len(self.finance_rows),
            'rows': rows
        }

    # -- LOA Records (cr63f_loamonitoring) -------------------------------------
    def fetch_loa_records(self) -> dict:
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse client not configured.'}
        try:
            table_name = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
            records = self.dv_client.fetch_records(table_name, custom_only=True)

            # Ensure % Paid is properly formatted as a percentage string
            for r in records:
                for pct_key in ('% Paid', 'cr63f_paid', '% PAID', 'Percent Paid'):
                    if pct_key in r and r[pct_key] != '':
                        r[pct_key] = format_percent(r[pct_key])

            self.loa_records = records

            # Extract ALL custom columns from Dataverse entity attributes
            attrs = self.dv_client.get_entity_attributes(table_name)
            clean_tbl = table_name.strip().lower()
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
            computed_keys = {
                'cr63f_totalcontractprice', 'cr63f_totalpaymentdeductibletotcp',
                'cr63f_balance', 'cr63f_paid', 'Total Contract Price',
                'Total Payment Deductible to TCP', 'Balance', '% Paid'
            }

            custom_cols = []
            # Priority order for preferred display
            priority_order = [
                'cr63f_type', 'cr63f_indexno', 'cr63f_areaindex', 'cr63f_titleno',
                'cr63f_control1', 'cr63f_registeredowner', 'cr63f_authorizedrepresentative',
                'cr63f_entity', 'cr63f_province', 'cr63f_municipality', 'cr63f_dbadjustedbarangay',
                'cr63f_lotareaha', 'new_contracttype', 'new_mode', 'cr63f_conversionstatus',
                'cr63f_rccprice', 'cr63f_rccterms', 'cr63f_endorseddate', 'cr63f_secureddate',
                'cr63f_securedendorsedate', 'cr63f_totalcontractprice',
                'cr63f_totalpaymentdeductibletotcp', 'cr63f_balance', 'cr63f_paid'
            ]

            processed_keys = set()
            # Add prioritized columns first
            for key in priority_order:
                if key in attrs:
                    meta = attrs[key]
                    low = key.lower()
                    disp = meta.get('display') or key
                    is_money = any(m in low for m in ('price', 'deductible', 'balance', 'tcp'))
                    is_percent = 'paid' in low or '%' in disp
                    custom_cols.append({
                        'key': key,
                        'label': disp,
                        'type': meta.get('type', 'String'),
                        'editable': key not in computed_keys,
                        'isMoney': is_money,
                        'isPercent': is_percent,
                        'width': 150 if is_money or len(disp) > 15 else 130
                    })
                    processed_keys.add(key)

            # Add any remaining custom attributes
            for log_name, meta in attrs.items():
                if log_name in processed_keys:
                    continue
                low = log_name.lower()
                if low in system_cols or low == f"{clean_tbl}id" or low.endswith('_base') or meta.get('type') in ('Virtual', 'PartyList'):
                    continue
                if meta.get('is_custom') or any(low.startswith(p) for p in ('cr63f_', 'new_')):
                    disp = meta.get('display') or log_name
                    is_money = any(m in low for m in ('price', 'deductible', 'balance', 'tcp'))
                    is_percent = 'paid' in low or '%' in disp
                    custom_cols.append({
                        'key': log_name,
                        'label': disp,
                        'type': meta.get('type', 'String'),
                        'editable': log_name not in computed_keys,
                        'isMoney': is_money,
                        'isPercent': is_percent,
                        'width': 140
                    })
                    processed_keys.add(log_name)

            # Calculate KPI summaries (clamping balance to 0 on overpayment)
            total_tcp = sum(parse_float(r.get('Total Contract Price', r.get('cr63f_totalcontractprice', 0))) or 0 for r in records)
            total_deductible = sum(parse_float(r.get('Total Payment Deductible to TCP', r.get('cr63f_totalpaymentdeductibletotcp', 0))) or 0 for r in records)
            total_balance = max(0.0, total_tcp - total_deductible)

            return {
                'success': True,
                'records': records,
                'columns': custom_cols,
                'count': len(records),
                'total_tcp': total_tcp,
                'total_tcp_formatted': format_currency(total_tcp),
                'total_deductible': total_deductible,
                'total_deductible_formatted': format_currency(total_deductible),
                'total_balance': total_balance,
                'total_balance_formatted': format_currency(total_balance),
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def save_loa_record_entry(self, record_data: dict, record_id: Optional[str] = None) -> dict:
        """Save a new log entry or update existing record in Dataverse upon user clicking Save."""
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        try:
            res = self.dv_client.save_loa_record(record_data, record_id=record_id)

            # Update local memory
            if record_id:
                for idx, r in enumerate(self.loa_records):
                    if r.get('__id') == record_id:
                        self.loa_records[idx].update(record_data)
                        break
            else:
                new_rec = dict(record_data)
                new_rec['__id'] = res.get('id', '')
                self.loa_records.insert(0, new_rec)

            return {
                'success': True,
                'action': res.get('action'),
                'id': res.get('id'),
                'message': f"Record successfully {'updated' if record_id else 'created'} in Dataverse."
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def delete_loa_record_entry(self, record_id: str) -> dict:
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        try:
            table_name = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
            info = self.dv_client.get_entity_info(table_name)
            self.dv_client.delete(info['entity_set'], record_id)
            self.loa_records = [r for r in self.loa_records if r.get('__id') != record_id]
            return {'success': True, 'message': 'Record successfully deleted from Dataverse.'}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    # -- Batch Processing & Export ---------------------------------------------
    def upload_batch_file_dialog(self) -> dict:
        if not self._window:
            return {'success': False, 'error': 'Window not attached.'}
        file_types = ('Excel and CSV Files (*.xlsx;*.xlsm;*.csv)', 'All files (*.*)')
        result = self._window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=file_types
        )
        if not result or not result[0]:
            return {'success': False, 'cancelled': True}

        file_path = result[0]
        try:
            headers, raw_rows = read_tabular_file(file_path)
            # Run lookup calculation across all rows
            processed_rows: List[Dict[str, Any]] = []
            for r in raw_rows:
                type_val = r.get("TYPE", r.get("Type", ""))
                index_no = r.get("INDEX NO", r.get("Index No", r.get("Corridor Index", "")))
                area_index = r.get("AREA-INDEX", r.get("Area Index", r.get("AREA INDEX", "")))
                title_mod = r.get("TITLE MOD", r.get("Title Mod", ""))

                lookup_res = self.finance_engine.lookup(type_val, index_no, area_index, title_mod)

                row_copy = dict(r)
                if lookup_res['matched']:
                    row_copy['Total Contract Price'] = lookup_res['total_contract_price_formatted']
                    row_copy['Total Payment Deductible to TCP'] = lookup_res['total_payment_deductible_formatted']
                    row_copy['Balance'] = lookup_res['balance_formatted']
                    row_copy['% Paid'] = lookup_res['percent_paid_formatted']
                    row_copy['__match_status'] = f"Matched via {lookup_res['match_type']}"
                else:
                    row_copy['__match_status'] = "No Finance Match"

                processed_rows.append(row_copy)

            self.batch_rows = processed_rows

            # Ensure all calculation columns are in headers
            for col in ['Total Contract Price', 'Total Payment Deductible to TCP', 'Balance', '% Paid']:
                if col not in headers:
                    headers.append(col)

            return {
                'success': True,
                'file_name': os.path.basename(file_path),
                'row_count': len(processed_rows),
                'headers': headers,
                'rows': processed_rows[:200]
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def export_records_excel(
        self,
        records: List[Dict[str, Any]],
        default_name: str = "LOA_Monitoring_Export.xlsx",
        columns: Optional[List[str]] = None,
        destination_path: Optional[str] = None
    ) -> dict:
        if not records:
            return {'success': False, 'error': 'No records to export.'}

        if destination_path:
            save_path = destination_path
        else:
            if not self._window:
                return {'success': False, 'error': 'Window not attached.'}
            result = self._window.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename=default_name,
                file_types=('Excel Files (*.xlsx)',)
            )
            if not result or not result[0]:
                return {'success': False, 'cancelled': True}
            save_path = result[0]

        if not save_path.lower().endswith('.xlsx'):
            save_path += '.xlsx'

        try:
            # System and internal Dataverse attributes to exclude from export
            system_exclude = {
                '__id', '__etag', '__idx', '__action', '__match_status', 'createdby', 'createdon',
                'createdonbehalfby', 'modifiedby', 'modifiedon', 'modifiedonbehalfby',
                'overriddencreatedon', 'versionnumber', 'importsequencenumber',
                'utcconversiontimezonecode', 'timezoneruleversionnumber', 'ownerid',
                'owningbusinessunit', 'owninguser', 'owningteam', 'statecode', 'statuscode',
                'exchangerate', 'transactioncurrencyid', 'statuscodename', 'statecodename',
                'createdbyname', 'modifiedbyname', 'owneridname', 'owningbusinessunitname',
                'owningusername', 'owningteamname', 'transactioncurrencyidname',
                'createdbyyominame', 'createdonbehalfbyname', 'createdonbehalfbyyominame',
                'modifiedbyyominame', 'modifiedonbehalfbyname', 'modifiedonbehalfbyyominame',
                'owneridtype', 'owneridyominame'
            }

            def is_excluded_col(col_name: Any) -> bool:
                if not col_name:
                    return True
                s = str(col_name).strip()
                low = s.lower()
                if s in ('#', 'Actions', '__idx', '__action', '__id', '__etag', '__match_status'):
                    return True
                if s.startswith('_') or s.startswith('@'):
                    return True
                if low.endswith(('_value', '(base)', '_base')):
                    return True
                if low in system_exclude:
                    return True
                if any(low == f"_{k}_value" or low == f"{k}_value" for k in system_exclude):
                    return True
                if (low.startswith('cr63f_') or low.startswith('new_')) and low.endswith('id') and len(low) > 8:
                    return True
                return False

            # Determine export columns
            if columns and len(columns) > 0:
                export_cols = [c for c in columns if not is_excluded_col(c)]
            else:
                # Filter keys from records: exclude system cols and logical cr63f_ prefixes if display name exists
                collected: List[str] = []
                for r in records:
                    for k in r.keys():
                        if is_excluded_col(k):
                            continue
                        low = k.lower().strip()
                        # If logical column like cr63f_... and display label is also in record, omit raw cr63f_
                        if (k.startswith('cr63f_') or k.startswith('new_')) and any(k != other and not other.startswith(('cr63f_', 'new_')) for other in r.keys()):
                            continue
                        if k not in collected:
                            collected.append(k)
                export_cols = collected

            # Build clean row dictionaries using only export_cols
            cleaned_rows = []
            for r in records:
                row_dict = {}
                for col in export_cols:
                    val = r.get(col, '')
                    if (val is None or val == '') and not col.startswith(('cr63f_', 'new_')):
                        clean_c = col.lower().replace(' ', '').replace('-', '').replace('_', '')
                        for rk, rv in r.items():
                            if rv is not None and rv != '' and not is_excluded_col(rk):
                                clean_k = rk.lower().replace(' ', '').replace('-', '').replace('_', '')
                                if clean_k == clean_c:
                                    val = rv
                                    break
                    row_dict[col] = val
                cleaned_rows.append(row_dict)

            write_excel_file(save_path, export_cols, cleaned_rows)
            return {'success': True, 'file_path': save_path}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    # -- Settings & Schema Discovery -------------------------------------------
    def get_settings(self) -> dict:
        cfg = dv.load_config()
        return {
            'environment_url': cfg.get('environment_url', ''),
            'tenant_id': cfg.get('tenant_id', ''),
            'client_id': cfg.get('client_id', ''),
            'solution_name': cfg.get('solution_name', 'LOA MONITORING'),
            'solution_id': cfg.get('solution_id', '34c6d4d3-b4b0-f111-aaac-7ced8d3b4de7'),
            'loa_table': cfg.get('loa_table', 'cr63f_loamonitoring'),
            'finance_table': cfg.get('finance_table', 'cr63f_financereference'),
            'theme': self.settings.get('theme', 'light'),
        }

    def save_settings(self, new_settings: dict) -> dict:
        try:
            cfg = {
                'environment_url': new_settings.get('environment_url', '').rstrip('/'),
                'tenant_id': new_settings.get('tenant_id', ''),
                'client_id': new_settings.get('client_id', ''),
                'solution_name': new_settings.get('solution_name', ''),
                'solution_id': new_settings.get('solution_id', ''),
                'loa_table': new_settings.get('loa_table', 'cr63f_loamonitoring'),
                'finance_table': new_settings.get('finance_table', 'cr63f_financereference'),
            }
            dv.save_config(cfg)
            self._init_dataverse()

            self.settings['theme'] = new_settings.get('theme', 'light')
            save_app_settings(self.settings)

            return {'success': True, 'message': 'Settings saved successfully.'}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def get_solutions(self) -> dict:
        if not self.dv_client:
            self._init_dataverse()
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse client is not initialized.'}
        try:
            solutions = self.dv_client.get_solutions()
            return {'success': True, 'solutions': solutions}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def get_solution_tables(self, solution_id: Optional[str] = None) -> dict:
        if not self.dv_client:
            self._init_dataverse()
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse client is not initialized.'}
        try:
            tables = self.dv_client.get_solution_tables(solution_id)
            return {'success': True, 'tables': tables}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def get_table_schema(self, table_name: Optional[str] = None) -> dict:
        if not self.dv_client:
            self._init_dataverse()
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse client is not initialized.'}
        try:
            tbl = table_name or self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
            attrs = self.dv_client.get_entity_attributes(tbl)
            picklists = self.dv_client.get_picklist_options(tbl)
            return {
                'success': True,
                'table_name': tbl,
                'attributes': attrs,
                'picklists': picklists,
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def save_table_layout(self, layout: dict) -> dict:
        """Persist user column widths, ordering, frozen columns, and density to disk."""
        try:
            current_layout = self.settings.get('table_layout', {})
            current_layout.update(layout)
            self.settings['table_layout'] = current_layout
            save_app_settings(self.settings)
            return {'success': True}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def get_table_layout(self) -> dict:
        return self.settings.get('table_layout', {})

    def discover_tables(self) -> dict:
        return self.get_solution_tables()


def _resource_dir() -> str:
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, 'web')


def _index_url() -> str:
    index_html = os.path.join(_resource_dir(), 'index.html')
    return index_html


def main() -> None:
    api = Api()
    window = webview.create_window(
        title=APP_TITLE,
        url=_index_url(),
        js_api=api,
        width=1380,
        height=860,
        min_size=(960, 600),
        background_color='#0F172A'
    )
    api.set_window(window)
    web_storage = os.path.join(ensure_user_data(), 'web_profile')
    os.makedirs(web_storage, exist_ok=True)
    webview.start(storage_path=web_storage, debug=False)


if __name__ == '__main__':
    main()

