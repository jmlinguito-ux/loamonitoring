"""LOA Monitoring Tool - Desktop Application Backend.

Pywebview + Python controller providing Dataverse integration, live Excel lookup
autofill engine, explicit record logging/saving, reference data sync, and batch processing.
"""

from __future__ import annotations

import json
import os
import re
import sys
import threading
import time
import uuid
from decimal import Decimal, InvalidOperation
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

# Ensure directory is on sys.path
_base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
if _base_dir and _base_dir not in sys.path:
    sys.path.insert(0, _base_dir)

from updater import UpdateService, run_startup_hooks

# Velopack starts the exe with install/update/uninstall arguments and expects it
# to handle them and exit before any window is created.
if __name__ == '__main__':
    run_startup_hooks()

import webview

from app_paths import data_file, ensure_user_data
import dataverse as dv
from loa_engine import (
    _strict_title_key as _strict_title,
    ConversionLookupEngine,
    FinanceLookupEngine,
    calculate_paid_percent,
    clean_title_variants,
    conversion_rank,
    format_calculated_percent,
    format_currency,
    finance_mode_class,
    format_percent,
    get_work_week,
    get_work_week_info,
    get_work_week_label,
    norm_key,
    norm_str,
    paid_percent_for_storage,
    parse_float,
    parse_log_datetime,
    read_tabular_file,
    set_text_safe,
    write_excel_file,
)

APP_TITLE = "LOA Monitoring Tool"
# Single source for the app version: shown in Settings > About, and read by
# build-release.ps1 / the GitHub release workflow as the release version.
APP_VERSION = "1.0.11"  # stamped by the release workflow from its version input


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
        "remembered_conversion_file": "",
        "recent_conversion_files": [],
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


# Compatibility hooks retained for older tests/extensions. They deliberately
# perform no filesystem I/O; audit data is session-only and sourced from Dataverse.
def load_cached_logs() -> List[Dict[str, Any]]:
    return []


def save_cached_logs(logs: List[Dict[str, Any]]) -> None:
    return None


def load_tombstoned_signatures() -> set:
    return set()


def save_tombstoned_signature(log_entry: Dict[str, Any]) -> None:
    return None



def _extract_num(row: dict, keys: list) -> Optional[float]:
    for k in keys:
        if k in row and row[k] is not None and str(row[k]).strip() != '':
            f = parse_float(row[k])
            if f is not None:
                return f
    return None


def _norm_pct(val: Any) -> Optional[float]:
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None
    f = parse_float(s)
    if f is None:
        return None
    # Stored % Paid is already on a 0-100 scale; values below 1 are small
    # percentages (0.5 = 0.50%), not fractions.
    return f


def _extract_pct(row: dict, keys: list) -> Optional[float]:
    for k in keys:
        if k in row and row[k] is not None and str(row[k]).strip() != '':
            f = _norm_pct(row[k])
            if f is not None:
                return f
    return None


def _decimal_amount(value: Any) -> Optional[Decimal]:
    if value is None or str(value).strip() in ('', '-', 'None', 'null'):
        return None
    try:
        return Decimal(str(value).replace('₱', '').replace('$', '').replace(',', '').strip())
    except (InvalidOperation, ValueError):
        parsed = parse_float(value)
        return Decimal(str(parsed)) if parsed is not None else None


def _record_amount(record: Dict[str, Any], keys: tuple) -> float:
    for key in keys:
        value = record.get(key)
        if value not in (None, ''):
            parsed = parse_float(value)
            if parsed is not None:
                return parsed
    return 0.0


def _record_tcp(record: Dict[str, Any]) -> float:
    return _record_amount(record, ('cr63f_totalcontractprice', 'Total Contract Price'))


def _record_deductible(record: Dict[str, Any]) -> float:
    return _record_amount(record, ('cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP', 'Payment Deductible'))


def _record_balance(record: Dict[str, Any]) -> float:
    """One account's outstanding balance, never negative (overpayment shows as 0)."""
    return max(0.0, _record_tcp(record) - _record_deductible(record))


_TCP_KEYS = ('cr63f_totalcontractprice', 'Total Contract Price')
_DEDUCTIBLE_KEYS = ('cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP', 'Payment Deductible')


def _none_if_blank(value: Any) -> Any:
    """None for an empty value; any other value (including a numeric 0) is kept."""
    return None if value is None or value == '' else value


def _derived_payment_fields(record: Dict[str, Any]) -> Dict[str, Any]:
    """Balance and % Paid calculated from TCP and Deductible, as Dataverse will store them.

    Used by every manual save path, so the stored values never depend on what
    the client calculated. Returns {} when the record carries neither amount
    key (nothing to derive). A blank Deductible counts as 0. A blank TCP cannot
    be calculated, so Balance and % Paid are cleared (None) rather than left
    from the previous TCP. Both the Dataverse and display keys are set.
    """
    if not any(key in record for key in _TCP_KEYS + _DEDUCTIBLE_KEYS):
        return {}
    tcp = _decimal_amount(_get_field_val(record, list(_TCP_KEYS)))
    if tcp is None:
        return {'cr63f_balance': None, 'Balance': None, 'cr63f_paid': None, '% Paid': None}
    ded = _decimal_amount(_get_field_val(record, list(_DEDUCTIBLE_KEYS))) or Decimal(0)
    balance = f"{max(Decimal(0), tcp - ded):.2f}"
    pct = f"{paid_percent_for_storage(tcp, ded):.1f}"
    return {'cr63f_balance': balance, 'Balance': balance, 'cr63f_paid': pct, '% Paid': pct}


def _refresh_display_percent(record: Dict[str, Any]) -> None:
    """Keep the grid's calculated % Paid in step with the TCP and Deductible just saved."""
    display = format_calculated_percent(
        _get_field_val(record, ['cr63f_totalcontractprice', 'Total Contract Price']),
        _get_field_val(record, ['cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP', 'Payment Deductible']),
    )
    if display:
        record['__display_percent_paid'] = display
    else:
        record.pop('__display_percent_paid', None)


def _required_field_cleared(stored_record: Dict[str, Any], incoming: Dict[str, Any]) -> str:
    """Error text when a save would blank Type or Mode on an existing account, else ''.

    Used by the drawer and the grid, so neither path can erase a stored value
    (for example a Mode missing from the drop-down list).
    """
    for label, keys in (('TYPE', ['cr63f_type', 'Type', 'TYPE']), ('Mode', ['new_mode', 'Mode', 'MODE'])):
        sent = next((str(incoming[k] if incoming[k] is not None else '').strip() for k in keys if k in incoming), None)
        stored = str(_get_field_val(stored_record, keys) or '').strip()
        if sent == '' and stored:
            return f'{label} cannot be cleared on an existing account (stored value: {stored}).'
    return ''


def _dialog_path(result: Any) -> str:
    """The chosen path from a pywebview file dialog, or '' when cancelled.

    pywebview 4.x/5.x return a plain string from a save dialog and a tuple from
    an open dialog; 6.x returns a tuple from both. Taking result[0] of a string
    gave its first letter ('C'), so the file was saved as 'C.xlsx'.
    """
    if not result:
        return ''
    if isinstance(result, str):
        return result
    try:
        return str(result[0] or '')
    except (TypeError, IndexError, KeyError):
        return ''


def _set_version(record: Dict[str, Any], etag: str) -> None:
    """Remember the version Dataverse returned after a save (for the next If-Match).

    Without a returned version the old one is dropped: the next save of that
    account is then unconditional rather than refused for a change we made.
    """
    if etag:
        record['__etag'] = etag
    else:
        record.pop('__etag', None)


def _reference_value_differs(key: str, stored: Any, incoming: Any) -> bool:
    """Does a file value differ from the stored reference value? (% columns compared as numbers.)"""
    if 'PAID' in str(key).upper() or '%' in str(key):
        return _is_different_pct(stored, incoming)
    if isinstance(incoming, datetime):
        return _norm_date_str(stored) != incoming.strftime('%m/%d/%Y')
    if isinstance(incoming, (int, float, Decimal)) and not isinstance(incoming, bool):
        stored_num = _decimal_amount(stored)
        return stored_num is None or stored_num != _decimal_amount(incoming)
    return not _audit_values_equal(CANONICAL_FIELD_MAP.get(key, key), stored, incoming)


def _audit_week_label(timestamp: Any) -> str:
    """Entry Logs work week with its year, e.g. '2026 · WW-40' (same label as the Overview)."""
    moment = parse_log_datetime(timestamp)
    return get_work_week_label(moment) if moment else ''


def _endorsed_date_value(record: Dict[str, Any]) -> str:
    """The record's Endorsed Date, or '' when blank or a 1900 placeholder."""
    value = str(record.get('cr63f_endorseddate') or record.get('Endorsed Date') or record.get('ENDORSED DATE') or '').strip()
    if not value or value in ('-', 'None', 'null') or '1900' in value or value.startswith(('0001', '01/01/0001')):
        return ''
    return value


def _record_mode(record: Dict[str, Any]) -> str:
    """Land Master Mode in upper case; an account with no Mode is grouped as 'BLANK'."""
    return str(record.get('new_mode') or record.get('Mode') or record.get('MODE') or '').strip().upper() or 'BLANK'


def _payment_group(mode: Any) -> str:
    """Payment group of a Mode (decision 2026-10-13/14).

    DIRECT_SALE  SECURED SALE, SS - FROM ROW ...   tracked as paid/unpaid
    STP_SALE     SS - STP, SS - FROM LAND STP      turned over to partner; buyer pays balance
    LEASE_DIRECT SECURED LEASE                     rent basis
    LEASE_STP    SL - STP, SL - FROM LAND STP      turned over to partner
    RESCINDED    RESCINDED, SS - RESCINDED
    OTHER        blank or not a sale/lease Mode
    """
    m = str(mode or '').strip().upper()
    if 'RESCIND' in m:
        return 'RESCINDED'
    is_stp = bool(re.search(r'\bSTP\b', m)) or 'SOLD TO PARTNER' in m
    mode_class = finance_mode_class(m)
    if mode_class == 'LEASE':
        return 'LEASE_STP' if is_stp else 'LEASE_DIRECT'
    if mode_class == 'SALE':
        return 'STP_SALE' if is_stp else 'DIRECT_SALE'
    return 'OTHER'


def _endorsed_is_blank(record: Dict[str, Any]) -> bool:
    """True only for a truly empty Endorsed Date. A 1900 placeholder is a stored value
    and is never overwritten by a reference update (decided 2026-10-08)."""
    value = str(record.get('cr63f_endorseddate') or record.get('Endorsed Date') or record.get('ENDORSED DATE') or '').strip()
    return value in ('', '-', 'None', 'null')


def _account_keys_of(record: Dict[str, Any]) -> set:
    """Upper-cased Control Number, AREA-INDEX and Title No that identify one account in the logs."""
    values = (
        record.get('cr63f_control1') or record.get('cr63f_control') or record.get('Control Number') or record.get('Control') or '',
        record.get('cr63f_areaindex') or record.get('Area-Index') or record.get('AREA-INDEX') or record.get('Area Index') or '',
        record.get('cr63f_titleno') or record.get('Title No') or record.get('TITLE NO') or '',
    )
    return {str(value).strip().upper() for value in values if str(value).strip()}


def _endorsed_work_week(record: Dict[str, Any]) -> str:
    """Land Master Work Week: derived from Endorsed Date, never from the save time."""
    endorsed = _endorsed_date_value(record)
    return get_work_week(endorsed) if endorsed else ''


def _odata_literal(value: str) -> str:
    """Quote a value for an OData $filter: wrap in quotes and double any quote inside."""
    return "'" + str(value).replace("'", "''") + "'"


def _json_safe(value: Any) -> Any:
    """Convert values returned through pywebview's JSON bridge to JSON-safe types.

    Decimal values are emitted as strings to avoid losing precision in money and
    calculated percentages while still allowing the UI to parse and format them.
    """
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, dict):
        return {key: _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    return value


def _is_different_num(a_val: Any, b_val: Any) -> bool:
    fa = _decimal_amount(a_val)
    fb = _decimal_amount(b_val)
    if fa is None and fb is None:
        return False
    if (fa is None and (fb is None or fb == 0)) or (fb is None and (fa is None or fa == 0)):
        return False
    if fa is None or fb is None:
        return True
    return fa != fb


def _is_different_pct(a_val: Any, b_val: Any) -> bool:
    if a_val is None and b_val is None:
        return False
    # If both format to identical string (e.g. '1.00%' == '1.00%'), there is NO change
    sa = format_percent(a_val).strip()
    sb = format_percent(b_val).strip()
    if sa and sb and sa == sb:
        return False

    pa = _norm_pct(a_val)
    pb = _norm_pct(b_val)
    if pa is None and pb is None:
        return False
    if (pa is None and (pb is None or abs(pb) < 0.005)) or (pb is None and (pa is None or abs(pa) < 0.005)):
        return False
    if pa is None or pb is None:
        return True
    # Same precision as the stored Dataverse value (1 decimal).
    return abs(pa - pb) >= 0.05


CANONICAL_FIELD_MAP = {
    'cr63f_totalcontractprice': 'Total Contract Price',
    'Total Contract Price': 'Total Contract Price',
    'TOTAL CONTRACT PRICE': 'Total Contract Price',
    'cr63f_totalpaymentdeductibletotcp': 'Total Payment Deductible to TCP',
    'Total Payment Deductible to TCP': 'Total Payment Deductible to TCP',
    'Payment Deductible': 'Total Payment Deductible to TCP',
    'TOTAL PAYMENT DEDUCTIBLE TO TCP': 'Total Payment Deductible to TCP',
    'cr63f_balance': 'Balance',
    'Balance': 'Balance',
    'BALANCE': 'Balance',
    'cr63f_paid': '% Paid',
    '% Paid': '% Paid',
    'Paid': '% Paid',
    '% PAID': '% Paid',
    'cr63f_conversionstatus': 'Conversion Status',
    'Conversion Status': 'Conversion Status',
    'cr63f_acquisitionsublevel': 'Acquisition Sub-Level',
    'Acquisition Sub-Level': 'Acquisition Sub-Level',
    'DAR STATUS': 'Conversion Status',
    'DAR Status': 'Conversion Status',
    'Status': 'Conversion Status',
    'cr63f_rccprice': 'RCC Price',
    'RCC Price': 'RCC Price',
    'cr63f_lotareaha': 'Lot Area (ha)',
    'Lot Area (ha)': 'Lot Area (ha)',
    'cr63f_endorsement': 'Endorsement',
    'Endorsement': 'Endorsement',
    'cr63f_buyername': 'Buyer Name',
    'Buyer Name': 'Buyer Name',
    'cr63f_titleno': 'Title No',
    'Title No': 'Title No',
    'cr63f_areaindex': 'AREA-INDEX',
    'Area-Index': 'AREA-INDEX',
    'AREA-INDEX': 'AREA-INDEX',
    'cr63f_indexno': 'Index No',
    'Index No': 'Index No',
    'cr63f_type': 'Type',
    'Type': 'Type',
    'TYPE': 'Type',
}

DATAVERSE_ATTR_TO_DISPLAY = CANONICAL_FIELD_MAP


FIELD_NAME_TO_DATAVERSE_ATTR = {
    'Total Contract Price': 'cr63f_totalcontractprice',
    'Total Payment Deductible to TCP': 'cr63f_totalpaymentdeductibletotcp',
    'Payment Deductible': 'cr63f_totalpaymentdeductibletotcp',
    'Balance': 'cr63f_balance',
    '% Paid': 'cr63f_paid',
    'Conversion Status': 'cr63f_conversionstatus',
    'Acquisition Sub-Level': 'cr63f_acquisitionsublevel',
    'cr63f_acquisitionsublevel': 'cr63f_acquisitionsublevel',
    'DAR STATUS': 'cr63f_conversionstatus',
    'DAR Status': 'cr63f_conversionstatus',
    'Status': 'cr63f_conversionstatus',
    'RCC Price': 'cr63f_rccprice',
    'RCC Terms': 'cr63f_rccterms',
    'Mode': 'new_mode',
    'Contract Type': 'new_contracttype',
    'Type': 'cr63f_type',
    'TYPE': 'cr63f_type',
    'Entity': 'cr63f_entity',
    'Registered Owner': 'cr63f_registeredowner',
    'Authorized Representative': 'cr63f_authorizedrepresentative',
    'Authorized Rep': 'cr63f_authorizedrepresentative',
    'Province': 'cr63f_province',
    'Municipality': 'cr63f_municipality',
    'DB Adjusted Barangay': 'cr63f_dbadjustedbarangay',
    'Adjusted Barangay': 'cr63f_dbadjustedbarangay',
    'Lot Area (Ha)': 'cr63f_lotareaha',
    'Lot Area (ha)': 'cr63f_lotareaha',
    'Lot Area': 'cr63f_lotareaha',
    'Endorsed Date': 'cr63f_endorseddate',
    'Secured Date': 'cr63f_secureddate',
    'Secured Endorse Date': 'cr63f_securedendorsedate',
    'Title No': 'cr63f_titleno',
    'Title No.': 'cr63f_titleno',
    'TITLE NO': 'cr63f_titleno',
    'Index No': 'cr63f_indexno',
    'Index No.': 'cr63f_indexno',
    'INDEX NO': 'cr63f_indexno',
    'AREA-INDEX': 'cr63f_areaindex',
    'Area-Index': 'cr63f_areaindex',
    'Area Index': 'cr63f_areaindex',
    'Control Number': 'cr63f_control1',
    'Control': 'cr63f_control1',
}


def _extract_val(rec: dict, keys: list) -> str:
    """Safely extract a non-empty string value for any alias in keys from record dictionary."""
    if not rec or not isinstance(rec, dict):
        return ''
    for k in keys:
        if k in rec and rec[k] is not None:
            s = str(rec[k]).strip()
            if s != '':
                return s
        uk = k.upper()
        if uk in rec and rec[uk] is not None:
            s = str(rec[uk]).strip()
            if s != '':
                return s
        lk = k.lower()
        if lk in rec and rec[lk] is not None:
            s = str(rec[lk]).strip()
            if s != '':
                return s
    return ''


REFERENCE_ONLY_ACTIONS = {'FINANCE_UPDATE', 'CONVERSION_UPDATE'}


def _audit_source_action(record: dict) -> str:
    action = str(record.get('action_type') or record.get('Action Type') or
                 record.get('cr63f_actiontype') or '').strip().upper().replace(' ', '_')
    if not action:
        match = re.search(r'\(([^()]+)\)\s*$', str(record.get('cr63f_name') or record.get('name') or ''))
        action = match.group(1).strip().upper().replace(' ', '_') if match else ''
    if not action:
        details = str(record.get('cr63f_details') or record.get('details') or record.get('Details') or '')
        match = re.search(r'\[SOURCE:([^\]]+)\]', details)
        if match:
            action = match.group(1).strip().upper()
        elif re.search(r'Updated Finance Reference\b', details, re.I):
            action = 'FINANCE_UPDATE'
        elif re.search(r'Updated Conversion Reference\b', details, re.I):
            action = 'CONVERSION_UPDATE'
    for token, canonical in (('BASELINE', 'BASELINE'), ('REVERT', 'REVERT'), ('DELETE', 'DELETE'),
                             ('RECALC', 'RECALCULATE'),
                             ('REFERENCE', 'REFERENCE_UPDATE'), ('FINANCE', 'FINANCE_UPDATE'),
                             ('CONVERSION', 'CONVERSION_UPDATE'), ('NEW', 'NEW_ENTRY')):
        if token in action:
            return canonical
    return 'MANUAL_EDIT'


def _audit_master_id(record: dict) -> str:
    marker_text = ' '.join(str(record.get(key) or '') for key in
                           ('cr63f_name', 'name', 'cr63f_details', 'details', 'Details'))
    match = re.search(r'\[MASTER:([^\]]+)\]', marker_text)
    return str(record.get('__master_record_id') or (match.group(1) if match else '') or
               record.get('cr63f_loamonitoringid') or '').strip()


def _audit_identity_keys(record: dict, master: bool = False) -> set:
    """Only strong account keys can suppress another account's baseline."""
    record_id = str(record.get('__id') or '') if master else _audit_master_id(record)
    control = str(record.get('cr63f_control1') or record.get('cr63f_control') or
                  record.get('Control Number') or record.get('Control') or '').strip().upper()
    keys = {f'MASTER:{record_id}'} if record_id else set()
    if control:
        keys.add(f'CTRL:{control}')
    return keys


def _audit_history_keys(record: dict) -> set:
    keys = _audit_identity_keys(record)
    master_keys = {key for key in keys if key.startswith('MASTER:')}
    return master_keys or keys


def build_log_entry(
    action_type: str,
    work_week: str,
    now_str: str,
    area_index: str,
    index_no: str,
    title_no: str,
    field_name: str,
    old_value: Any,
    new_value: Any,
    details: Optional[str] = None,
    full_record: Optional[Dict[str, Any]] = None
) -> dict:
    old_str = _format_log_value(field_name, old_value)
    new_str = _format_log_value(field_name, new_value)
    det = details or f"Modified {field_name} from {old_str} to {new_str}"

    entry: Dict[str, Any] = {}
    if full_record and isinstance(full_record, dict):
        for k, v in full_record.items():
            if not k.startswith('__') and not k.startswith('@'):
                entry[k] = v

    # Explicitly map and ensure all master lot attributes are populated with both logical and display aliases
    if full_record and isinstance(full_record, dict):
        rec = full_record
        ctrl = _extract_val(rec, ['cr63f_control1', 'cr63f_control', 'Control', 'Control Number', 'Control No.', 'Control No'])
        r_type = _extract_val(rec, ['cr63f_type', 'Type', 'TYPE'])
        idx_no = index_no or _extract_val(rec, ['cr63f_indexno', 'Index No', 'INDEX NO', 'Index No.', 'index_no'])
        a_idx = area_index or _extract_val(rec, ['cr63f_areaindex', 'Area-Index', 'AREA-INDEX', 'Area Index', 'area_index'])
        t_no = title_no or _extract_val(rec, ['cr63f_titleno', 'Title No', 'TITLE NO', 'Title No.', 'Title Number', 'title_no', 'Title Mod'])
        owner = _extract_val(rec, ['cr63f_registeredowner', 'Registered Owner', 'registered_owner', 'Owner'])
        rep = _extract_val(rec, ['cr63f_authorizedrepresentative', 'Authorized Representative', 'Authorized Rep', 'authorized_representative', 'Rep'])
        ent = _extract_val(rec, ['cr63f_entity', 'Entity', 'entity'])
        prov = _extract_val(rec, ['cr63f_province', 'Province', 'province'])
        muni = _extract_val(rec, ['cr63f_municipality', 'Municipality', 'municipality', 'MUNICIPALITY/CITY'])
        brgy = _extract_val(rec, ['cr63f_dbadjustedbarangay', 'DB Adjusted Barangay', 'Adjusted Barangay', 'Barangay', 'BARANGAY', 'db_adjusted_barangay'])
        area_ha = _extract_val(rec, ['cr63f_lotareaha', 'Lot Area (Ha)', 'Lot Area (ha)', 'Lot Area', 'AREA APPLIED (HAS.)', 'lot_area_ha'])
        ctype = _extract_val(rec, ['new_contracttype', 'cr63f_contracttype', 'Contract Type', 'contract_type'])
        mode_val = _extract_val(rec, ['new_mode', 'cr63f_mode', 'Mode', 'MODE', 'mode'])
        cstatus = _extract_val(rec, ['cr63f_conversionstatus', 'Conversion Status', 'DAR STATUS', 'DAR Status', 'Status', 'conversion_status'])
        acquisition_sublevel = _extract_val(rec, ['cr63f_acquisitionsublevel', 'Acquisition Sub-Level'])
        rcc_p = _extract_val(rec, ['cr63f_rccprice', 'RCC Price', 'rcc_price'])
        rcc_t = _extract_val(rec, ['cr63f_rccterms', 'RCC Terms', 'rcc_terms'])
        e_date = _extract_val(rec, ['cr63f_endorseddate', 'Endorsed Date', 'endorsed_date'])
        s_date = _extract_val(rec, ['cr63f_secureddate', 'Secured Date', 'secured_date'])
        se_date = _extract_val(rec, ['cr63f_securedendorsedate', 'Secured Endorse Date', 'secured_endorse_date'])
        tcp_val_r = _extract_val(rec, ['cr63f_totalcontractprice', 'Total Contract Price', 'TCP', 'total_contract_price'])
        ded_val_r = _extract_val(rec, ['cr63f_totalpaymentdeductibletotcp', 'Payment Deductible', 'Total Payment Deductible to TCP', 'Total Payment Deductible', 'payment_deductible'])
        bal_val_r = _extract_val(rec, ['cr63f_balance', 'Balance', 'Outstanding Balance', 'balance'])
        paid_val_r = _extract_val(rec, ['cr63f_paid', 'cr63f_percentpaid', '% Paid', 'Percent Paid', 'Paid', 'paid'])

        entry.update({
            'cr63f_control1': ctrl, 'cr63f_control': ctrl, 'Control': ctrl, 'Control Number': ctrl,
            'cr63f_type': r_type, 'Type': r_type,
            'cr63f_registeredowner': owner, 'Registered Owner': owner,
            'cr63f_authorizedrepresentative': rep, 'Authorized Representative': rep, 'Authorized Rep': rep,
            'cr63f_entity': ent, 'Entity': ent,
            'cr63f_province': prov, 'Province': prov,
            'cr63f_municipality': muni, 'Municipality': muni,
            'cr63f_dbadjustedbarangay': brgy, 'DB Adjusted Barangay': brgy, 'Adjusted Barangay': brgy,
            'cr63f_lotareaha': area_ha, 'Lot Area (Ha)': area_ha, 'Lot Area (ha)': area_ha,
            'new_contracttype': ctype, 'cr63f_contracttype': ctype, 'Contract Type': ctype,
            'new_mode': mode_val, 'cr63f_mode': mode_val, 'Mode': mode_val,
            'cr63f_conversionstatus': cstatus, 'Conversion Status': cstatus,
            'cr63f_acquisitionsublevel': acquisition_sublevel, 'Acquisition Sub-Level': acquisition_sublevel,
            'cr63f_rccprice': rcc_p, 'RCC Price': rcc_p,
            'cr63f_rccterms': rcc_t, 'RCC Terms': rcc_t,
            'cr63f_endorseddate': e_date, 'Endorsed Date': e_date,
            'cr63f_secureddate': s_date, 'Secured Date': s_date,
            'cr63f_securedendorsedate': se_date, 'Secured Endorse Date': se_date,
            'cr63f_totalcontractprice': tcp_val_r, 'Total Contract Price': tcp_val_r,
            'cr63f_totalpaymentdeductibletotcp': ded_val_r, 'Payment Deductible': ded_val_r, 'Total Payment Deductible to TCP': ded_val_r,
            'cr63f_balance': bal_val_r, 'Balance': bal_val_r,
            'cr63f_paid': paid_val_r, '% Paid': paid_val_r,
        })

    # Overlay specific audit values and aliases
    master_id = _audit_master_id(full_record or {})
    if action_type not in REFERENCE_ONLY_ACTIONS and not master_id and full_record:
        # A master snapshot has no audit field; a log's __id identifies the log.
        if not any(key in full_record for key in ('field_name', 'Field Changed', 'cr63f_fieldchanged')):
            master_id = str(full_record.get('__id') or '').strip()
    if action_type in REFERENCE_ONLY_ACTIONS:
        master_id = ''
        entry.pop('cr63f_loamonitoringid', None)
    marker = f' [MASTER:{master_id}]' if master_id else ''
    if marker and marker not in det:
        det += marker
    source_marker = f' [SOURCE:{action_type}]'
    if source_marker not in det:
        det += source_marker
    if master_id:
        entry['__master_record_id'] = master_id
    entry.update({
        # Primary Name
        'cr63f_name': f"{area_index or 'Lot'} - {field_name}{marker} ({action_type})",
        'name': f"{area_index or 'Lot'} - {field_name}{marker} ({action_type})",

        # Action Type
        'action_type': action_type,
        'Action Type': action_type,
        'cr63f_actiontype': action_type,

        # Work Week
        'work_week': work_week,
        'Work Week': work_week,
        'cr63f_workweek': work_week,

        # Log Time / Timestamp
        'log_time': now_str,
        'Log Time': now_str,
        'timestamp': now_str,
        'cr63f_logtime': now_str,

        # AREA-INDEX
        'area_index': area_index,
        'AREA-INDEX': area_index,
        'Area-Index': area_index,
        'Area Index': area_index,
        'cr63f_areaindex': area_index,

        # Index No
        'index_no': index_no,
        'Index No': index_no,
        'INDEX NO': index_no,
        'cr63f_indexno': index_no,

        # Title No
        'title_no': title_no,
        'Title No': title_no,
        'TITLE NO': title_no,
        'cr63f_titleno': title_no,

        # Field Changed
        'field_name': field_name,
        'Field Changed': field_name,
        'cr63f_fieldchanged': field_name,

        # Previous Value
        'old_value': old_str,
        'Previous Value': old_str,
        'cr63f_previousvalue': old_str,

        # New Value
        'new_value': new_str,
        'New Value': new_str,
        'cr63f_newvalue': new_str,

        # Details
        'details': det,
        'Details': det,
        'cr63f_details': det,
    })
    return entry


def _format_log_value(field_name: str, value: Any) -> str:
    """Return a consistent, human-readable audit value for a changed field."""
    raw = str(value if value is not None else '').strip()
    if raw in ('', '-', 'Previous', 'Updated', 'Created'):
        return raw

    canonical_name = CANONICAL_FIELD_MAP.get(field_name, field_name)
    if canonical_name in {
        'Total Contract Price', 'Total Payment Deductible to TCP',
        'Balance', 'RCC Price'
    }:
        numeric_value = _norm_currency_str(value)
        return f"₱{numeric_value}" if numeric_value else raw
    if canonical_name == '% Paid':
        percent_value = _extract_pct({'value': value}, ['value'])
        return f"{percent_value:.2f}%" if percent_value is not None else raw
    return raw


LOG_COMPARE_SPECS = [
    # (canonical_name, [keys_to_check], type: 'currency' | 'percent' | 'date' | 'text')
    ('Total Contract Price', ['cr63f_totalcontractprice', 'Total Contract Price'], 'currency'),
    ('Total Payment Deductible to TCP', ['cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP', 'Payment Deductible'], 'currency'),
    ('Conversion Status', ['cr63f_conversionstatus', 'Conversion Status', 'Status'], 'text'),
    ('Acquisition Sub-Level', ['cr63f_acquisitionsublevel', 'Acquisition Sub-Level'], 'text'),
    ('Contract Type', ['new_contracttype', 'cr63f_contracttype', 'Contract Type'], 'text'),
    ('Mode', ['new_mode', 'cr63f_mode', 'Mode'], 'text'),
    ('Registered Owner', ['cr63f_registeredowner', 'Registered Owner'], 'text'),
    ('Authorized Representative', ['cr63f_authorizedrepresentative', 'Authorized Representative', 'Authorized Rep'], 'text'),
    ('Entity', ['cr63f_entity', 'Entity'], 'text'),
    ('Province', ['cr63f_province', 'Province'], 'text'),
    ('Municipality', ['cr63f_municipality', 'Municipality'], 'text'),
    ('DB Adjusted Barangay', ['cr63f_dbadjustedbarangay', 'DB Adjusted Barangay', 'Adjusted Barangay'], 'text'),
    ('Lot Area (Ha)', ['cr63f_lotareaha', 'Lot Area (Ha)', 'Lot Area (ha)', 'Lot Area'], 'text'),
    ('RCC Price', ['cr63f_rccprice', 'RCC Price'], 'currency'),
    ('RCC Terms', ['cr63f_rccterms', 'RCC Terms'], 'text'),
    ('Secured Date', ['cr63f_secureddate', 'Secured Date'], 'date'),
    ('Secured Endorse Date', ['cr63f_securedendorsedate', 'Secured Endorse Date'], 'date'),
    ('Endorsed Date', ['cr63f_endorseddate', 'Endorsed Date'], 'date'),
    ('Title No', ['cr63f_titleno', 'Title No', 'TITLE NO'], 'text'),
    ('AREA-INDEX', ['cr63f_areaindex', 'Area-Index', 'AREA-INDEX', 'Area Index'], 'text'),
    ('Control Number', ['cr63f_control1', 'cr63f_control', 'Control', 'Control Number'], 'text'),
    ('Type', ['cr63f_type', 'Type', 'TYPE'], 'text'),
    ('Index No', ['cr63f_indexno', 'Index No', 'INDEX NO'], 'text'),
]


def _norm_currency_str(val: Any) -> str:
    if val is None or str(val).strip() in ('', '-', 'None', 'null'):
        return ""
    s = str(val).replace('₱', '').replace(',', '').strip()
    try:
        number = Decimal(s)
        if not number.is_finite():
            return str(val).strip()
        rendered = format(number, ',f')
        if '.' not in rendered:
            rendered += '.00'
        else:
            whole, fraction = rendered.split('.', 1)
            rendered = f"{whole}.{fraction.ljust(2, '0')}"
        return rendered
    except (InvalidOperation, ValueError, TypeError):
        return str(val).strip()


def _norm_pct_str(val: Any) -> str:
    if val is None or str(val).strip() in ('', '-', 'None', 'null'):
        return ""
    s = str(val).replace('%', '').strip()
    try:
        f = float(s)
        return f"{f:.2f}%"
    except (ValueError, TypeError):
        return str(val).strip()


def _norm_date_str(val: Any) -> str:
    if val is None or str(val).strip() in ('', '-', 'None', 'null'):
        return ""
    s = str(val).split('T')[0].split(' ')[0].strip()
    parts = re.split(r'[-/]', s)
    if len(parts) == 3:
        if len(parts[0]) == 4:
            return f"{int(parts[1]):02d}/{int(parts[2]):02d}/{parts[0]}"
        elif len(parts[2]) == 4:
            return f"{int(parts[0]):02d}/{int(parts[1]):02d}/{parts[2]}"
    return s


def _norm_text_str(val: Any) -> str:
    if val is None or str(val).strip() in ('None', 'null'):
        return ""
    return str(val).strip()


def _get_field_val(record: dict, keys: list) -> Any:
    for k in keys:
        if k in record and record[k] is not None:
            v_str = str(record[k]).strip()
            if v_str not in ('', 'None', 'null'):
                return record[k]
    return ""


REFERENCE_AUDIT_FIELDS = {
    'Total Contract Price',
    'Total Payment Deductible to TCP',
    'Conversion Status',
    # Filled with the update date when a reference update changes an account
    # that has no Endorsed Date yet.
    'Endorsed Date',
}
CALCULATED_AUDIT_FIELDS = {'Balance', '% Paid', 'Paid', 'Percent Paid'}


def _audit_values_equal(field_name: str, left: Any, right: Any) -> bool:
    """Compare stored values without display-format rounding."""
    canonical = CANONICAL_FIELD_MAP.get(field_name, field_name)
    spec = next((entry for entry in LOG_COMPARE_SPECS if entry[0] == canonical), None)
    if canonical == 'Balance' or (spec and spec[2] == 'currency'):
        left_norm = _norm_currency_str(left)
        right_norm = _norm_currency_str(right)
        if not left_norm or not right_norm:
            return left_norm == right_norm
        try:
            return Decimal(left_norm.replace(',', '')) == Decimal(right_norm.replace(',', ''))
        except InvalidOperation:
            return left_norm == right_norm
    if spec and spec[2] == 'percent':
        return _norm_pct_str(left) == _norm_pct_str(right)
    if spec and spec[2] == 'date':
        return _norm_date_str(left) == _norm_date_str(right)
    return _norm_text_str(left) == _norm_text_str(right)


def derive_field_diffs_from_snapshots(
    raw_records: List[Dict[str, Any]],
    active_session_logs: Optional[List[Dict[str, Any]]] = None
) -> List[Dict[str, Any]]:
    """
    Derives individual field diff records by chronologically sorting lot snapshots
    from cr63f_loamonitoringlogs and comparing consecutive snapshots by immutable
    master ID, with Control Number as the fallback for legacy history.
    Snapshots with zero changes are excluded.
    """
    if not raw_records and not active_session_logs:
        return []

    explicit_diff_logs: List[Dict[str, Any]] = []
    snapshots: List[Dict[str, Any]] = []

    def get_log_timestamp(rec: dict) -> Any:
        """Use the immutable Dataverse creation time as the audit timestamp."""
        return (
            rec.get('createdon') or rec.get('cr63f_logtime') or
            rec.get('timestamp') or rec.get('log_time') or
            rec.get('Log Time') or ''
        )

    def get_source_action(rec: dict) -> str:
        """Preserve how the save was initiated instead of inferring it from the field."""
        return _audit_source_action(rec)

    def normalize_audit_week(rec: Dict[str, Any]) -> Dict[str, Any]:
        """Audit work week is always derived from log time (Tuesday-Monday)."""
        normalized = dict(rec)
        calculated = _audit_week_label(get_log_timestamp(normalized))
        ww = calculated or normalized.get('work_week') or normalized.get('Work Week') or normalized.get('cr63f_workweek')
        if ww:
            normalized['work_week'] = ww
            normalized['Work Week'] = ww
            normalized['cr63f_workweek'] = ww
        return normalized

    for r in (raw_records or []):
        f_name = r.get('field_name') or r.get('Field Changed') or r.get('cr63f_fieldchanged') or ''
        f_name_str = str(f_name).strip()
        o_val = r.get('old_value') or r.get('Previous Value') or r.get('cr63f_previousvalue')
        n_val = r.get('new_value') or r.get('New Value') or r.get('cr63f_newvalue')

        if f_name_str and f_name_str not in ('All Fields', 'Full Record') and (o_val or n_val or f_name_str == 'Initial Entry'):
            canonical = CANONICAL_FIELD_MAP.get(f_name_str, f_name_str)
            action = get_source_action(r)
            if action == 'BASELINE':
                snapshots.append(dict(r))
                continue
            if action == 'REVERT':
                # A revert restores a known state; its embedded record is the
                # anchor the next change is compared with.
                snapshots.append(dict(r))
            if canonical in CALCULATED_AUDIT_FIELDS:
                continue
            if action == 'REFERENCE_UPDATE' and canonical not in REFERENCE_AUDIT_FIELDS:
                continue
            explicit_diff_logs.append(normalize_audit_week(r))
        else:
            snapshots.append(dict(r))

    if active_session_logs:
        for r in active_session_logs:
            f_name = r.get('field_name') or r.get('Field Changed') or r.get('cr63f_fieldchanged') or ''
            f_name_str = str(f_name).strip()
            if f_name_str and f_name_str not in ('All Fields', 'Full Record'):
                canonical = CANONICAL_FIELD_MAP.get(f_name_str, f_name_str)
                action = get_source_action(r)
                if action == 'BASELINE':
                    snapshots.append(dict(r))
                    continue
                if action == 'REVERT':
                    snapshots.append(dict(r))
                if canonical in CALCULATED_AUDIT_FIELDS:
                    continue
                if action == 'REFERENCE_UPDATE' and canonical not in REFERENCE_AUDIT_FIELDS:
                    continue
                explicit_diff_logs.append(normalize_audit_week(r))
            else:
                snapshots.append(dict(r))

    control_master_ids: Dict[str, set] = {}
    for snapshot in snapshots:
        master_id = _audit_master_id(snapshot)
        if master_id:
            for key in _audit_identity_keys(snapshot):
                if key.startswith('CTRL:'):
                    control_master_ids.setdefault(key, set()).add(master_id)

    def get_lot_key(rec: dict) -> str:
        master_id = _audit_master_id(rec)
        if master_id:
            return f'MASTER:{master_id}'
        ctrl = str(rec.get('cr63f_control1') or rec.get('cr63f_control') or rec.get('Control') or rec.get('Control Number') or rec.get('Control No.') or rec.get('Control No') or rec.get('control_no') or '').strip().upper()
        if ctrl:
            master_ids = control_master_ids.get(f'CTRL:{ctrl}', set())
            if len(master_ids) == 1:
                return f'MASTER:{next(iter(master_ids))}'
            return f"CTRL:{ctrl}"
        ai = str(rec.get('cr63f_areaindex') or rec.get('AREA-INDEX') or rec.get('Area-Index') or rec.get('area_index') or rec.get('Area Index') or '').strip().upper()
        if ai:
            return f"AI:{ai}"
        t_no = str(rec.get('cr63f_titleno') or rec.get('Title No') or rec.get('title_no') or rec.get('TITLE NO') or rec.get('Title No.') or '').strip().upper()
        if t_no:
            return f"TITLE:{t_no}"
        rec_id = str(rec.get('__id') or rec.get('id') or rec.get('cr63f_loamonitoringlogid') or '').strip()
        return f"ID:{rec_id}"

    def get_time_key(rec: dict) -> Tuple[datetime, int]:
        # Order by the real time, never by text: Dataverse returns createdon as
        # display text ('10/2/2026 9:00 AM'), which sorts before '9/30/2026'.
        # The raw ISO value (seconds precision) wins when the fetch kept it.
        # A baseline is the pre-change state, so it goes first on a tie; other
        # ties keep their original order (the sort is stable).
        moment = (
            parse_log_datetime(rec.get('__createdon_raw')) or
            parse_log_datetime(get_log_timestamp(rec)) or
            datetime.min
        )
        return (moment, 0 if get_source_action(rec) == 'BASELINE' else 1)

    lot_groups: Dict[str, List[Dict[str, Any]]] = {}
    for s in snapshots:
        k = get_lot_key(s)
        if k.startswith('MASTER:'):
            s['__master_record_id'] = k[len('MASTER:'):]
        lot_groups.setdefault(k, []).append(s)

    derived_diffs: List[Dict[str, Any]] = []

    for lot_key, group in lot_groups.items():
        group.sort(key=get_time_key)

        for i, curr_s in enumerate(group):
            ts_str = get_log_timestamp(curr_s) or datetime.now().strftime('%m/%d/%Y %I:%M %p')
            # Never trust a stale work-week value stored on the source row. The
            # audit week belongs to the log timestamp, not to Endorsed Date.
            ww = _audit_week_label(ts_str) or get_work_week_label(datetime.now())
            ctrl_val = curr_s.get('cr63f_control1') or curr_s.get('cr63f_control') or curr_s.get('Control') or curr_s.get('Control Number') or ''
            area_idx = curr_s.get('cr63f_areaindex') or curr_s.get('Area-Index') or curr_s.get('area_index') or ''
            index_no = curr_s.get('cr63f_indexno') or curr_s.get('Index No') or curr_s.get('index_no') or ''
            title_no = curr_s.get('cr63f_titleno') or curr_s.get('Title No') or curr_s.get('title_no') or ''
            rec_log_id = curr_s.get('__id') or curr_s.get('id') or curr_s.get('cr63f_loamonitoringlogid') or ''

            action_type = get_source_action(curr_s)
            # Baselines and reverts exist only to anchor comparisons. Their own
            # field rows are already in explicit_diff_logs, so they are not
            # diffed again here. Baselines are never exposed as Entry Logs.
            if action_type in ('BASELINE', 'REVERT'):
                continue

            if action_type == 'NEW_ENTRY':
                init_entry = build_log_entry(
                    action_type='NEW_ENTRY', work_week=ww, now_str=ts_str,
                    area_index=area_idx, index_no=index_no, title_no=title_no,
                    field_name='Initial Entry', old_value='-', new_value='Created',
                    details=f"Initial record entry for {ctrl_val or area_idx or lot_key}",
                    full_record=curr_s
                )
                if rec_log_id:
                    init_entry['__id'] = f"{rec_log_id}_init"
                derived_diffs.append(init_entry)
                continue

            if i == 0:
                # Legacy history may begin with an old snapshot that has no
                # Action Type metadata. Preserve its historical Initial Entry
                # meaning; explicit BASELINE rows are handled above.
                explicit_action = str(
                    curr_s.get('action_type') or curr_s.get('Action Type') or
                    curr_s.get('cr63f_actiontype') or ''
                ).strip()
                source_name = str(curr_s.get('cr63f_name') or curr_s.get('name') or '').upper()
                source_details = str(curr_s.get('cr63f_details') or curr_s.get('details') or curr_s.get('Details') or '')
                has_source_marker = bool(re.search(r'\[SOURCE:[^\]]+\]', source_details) or
                                         re.search(r'\(([^()]+)\)\s*$', source_name))
                if not explicit_action and not has_source_marker:
                    init_entry = build_log_entry(
                        action_type='NEW_ENTRY', work_week=ww, now_str=ts_str,
                        area_index=area_idx, index_no=index_no, title_no=title_no,
                        field_name='Initial Entry', old_value='-', new_value='Created',
                        details=f"Initial record entry for {ctrl_val or area_idx or lot_key}",
                        full_record=curr_s
                    )
                    if rec_log_id:
                        init_entry['__id'] = f"{rec_log_id}_init"
                    derived_diffs.append(init_entry)
                continue

            prev_s = group[i - 1]
            snapshot_diffs: List[Dict[str, Any]] = []

            for canon_name, keys, f_type in LOG_COMPARE_SPECS:
                if canon_name in CALCULATED_AUDIT_FIELDS:
                    continue
                if action_type == 'REFERENCE_UPDATE' and canon_name not in REFERENCE_AUDIT_FIELDS:
                    continue

                p_val = _get_field_val(prev_s, keys)
                c_val = _get_field_val(curr_s, keys)

                if f_type == 'currency':
                    p_norm = _norm_currency_str(p_val)
                    c_norm = _norm_currency_str(c_val)
                    old_disp = f"₱{p_norm}" if p_norm else "-"
                    new_disp = f"₱{c_norm}" if c_norm else "-"
                elif f_type == 'percent':
                    p_norm = _norm_pct_str(p_val)
                    c_norm = _norm_pct_str(c_val)
                    old_disp = p_norm if p_norm else "-"
                    new_disp = c_norm if c_norm else "-"
                elif f_type == 'date':
                    p_norm = _norm_date_str(p_val)
                    c_norm = _norm_date_str(c_val)
                    old_disp = p_norm if p_norm else "-"
                    new_disp = c_norm if c_norm else "-"
                else:
                    p_norm = _norm_text_str(p_val)
                    c_norm = _norm_text_str(c_val)
                    old_disp = p_norm if p_norm else "-"
                    new_disp = c_norm if c_norm else "-"

                if p_norm != c_norm:
                    diff_log = build_log_entry(
                        action_type=action_type,
                        work_week=ww,
                        now_str=ts_str,
                        area_index=area_idx,
                        index_no=index_no,
                        title_no=title_no,
                        field_name=canon_name,
                        old_value=old_disp,
                        new_value=new_disp,
                        details=f"Updated {canon_name} from {old_disp} to {new_disp}",
                        full_record=curr_s
                    )
                    if rec_log_id:
                        diff_log['__id'] = f"{rec_log_id}_{canon_name.replace(' ', '_')}"
                    snapshot_diffs.append(diff_log)

            if snapshot_diffs:
                derived_diffs.extend(snapshot_diffs)

    seen_events = set()
    combined_logs: List[Dict[str, Any]] = []

    for item in (explicit_diff_logs + derived_diffs):
        item_id = str(item.get('__id') or item.get('id') or item.get('cr63f_loamonitoringlogid') or '').strip()
        ctrl = str(item.get('cr63f_control1') or item.get('cr63f_control') or item.get('Control') or item.get('Control Number') or item.get('Control No') or '').strip().upper()
        a_idx = str(item.get('area_index') or item.get('cr63f_areaindex') or item.get('Area-Index') or '').strip().upper()
        f_name = str(item.get('field_name') or item.get('cr63f_fieldchanged') or '').strip()
        n_val = str(item.get('new_value') or item.get('cr63f_newvalue') or '').strip()
        t_str = str(item.get('log_time') or item.get('createdon') or item.get('timestamp') or '').strip()

        # Separate saves can happen within one minute and may write the same
        # value again. Only collapse the same persisted audit row.
        event_sig = ('ID', item_id) if item_id else (ctrl or a_idx, f_name, n_val, t_str)
        if event_sig not in seen_events:
            seen_events.add(event_sig)
            combined_logs.append(item)

    combined_logs.sort(key=get_time_key, reverse=True)
    return combined_logs


class Api:
    def __init__(self):
        self.dv_client: Optional[dv.DataverseClient] = None
        self.finance_engine: FinanceLookupEngine = FinanceLookupEngine()
        self.conversion_engine: ConversionLookupEngine = ConversionLookupEngine()
        self.loa_records: List[Dict[str, Any]] = []
        # Audit logs are session-only. Dataverse is the sole source of truth.
        self.log_records: List[Dict[str, Any]] = []
        # Newly written audit rows remain in memory until Dataverse includes
        # them in a subsequent fetch (Dataverse reads can be briefly delayed).
        self.pending_session_logs: List[Dict[str, Any]] = []
        self.finance_rows: List[Dict[str, Any]] = []
        self.finance_source_name: str = "None"
        self.conversion_rows: List[Dict[str, Any]] = []
        self.conversion_source_name: str = "None"
        self.batch_rows: List[Dict[str, Any]] = []
        self.active_flow: Optional[dict] = None
        self._window: Optional[webview.Window] = None

        # Load settings
        self.settings = load_app_settings()

        # Try initializing Dataverse client silently
        self._init_dataverse()

        # Try auto-loading remembered reference files if exist
        self._auto_load_remembered_file()

    def _add_log_records(self, new_logs: List[Dict[str, Any]]) -> None:
        """Add new log entries to the active session only; never persist a local log cache."""
        if not new_logs:
            return
        for le in reversed(new_logs):
            label = _audit_week_label(le.get('log_time') or le.get('createdon') or le.get('timestamp'))
            if label:
                le['work_week'] = le['Work Week'] = label
            self.log_records.insert(0, le)
        seen_sig = set()
        deduped = []
        for l in self.log_records:
            log_id = str(l.get('__id') or l.get('id') or '').strip()
            a_idx = str(l.get('area_index') or l.get('cr63f_areaindex') or '').strip().upper()
            f_name = str(l.get('field_name') or l.get('cr63f_fieldchanged') or '').strip()
            n_val = str(l.get('new_value') or l.get('cr63f_newvalue') or '').strip()
            t_str = str(l.get('log_time') or l.get('createdon') or l.get('timestamp') or '').strip()
            sig = (log_id, a_idx, f_name, n_val, t_str) if log_id else ('', a_idx, f_name, n_val, t_str)
            if sig not in seen_sig:
                seen_sig.add(sig)
                deduped.append(l)
        self.log_records = deduped

    def _logged_account_keys(self) -> set:
        """Strong account identities with master audit history in Dataverse.

        One query for the whole logs table, so a bulk reference update does not
        query the logs once per account.
        """
        logs_table = self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs')
        keys: set = set()
        for row in self.dv_client.fetch_records(logs_table, custom_only=False):
            if _audit_source_action(row) not in REFERENCE_ONLY_ACTIONS:
                keys.update(_audit_history_keys(row))
        return keys

    def _logs_control_columns(self) -> List[str]:
        """Control Number columns present on the logs table (cr63f_control1 and/or cr63f_control)."""
        logs_table = self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs')
        try:
            attrs = self.dv_client.get_entity_attributes(logs_table) or {}
        except Exception:
            attrs = {}
        found = [name for name in ('cr63f_control1', 'cr63f_control') if name in attrs]
        return found or ['cr63f_control']

    def _write_baseline_if_missing(
        self, record: Dict[str, Any], now_str: str, known_keys: Optional[set] = None,
        require_master_identity: bool = False, keys_lock: Optional[threading.Lock] = None
    ) -> Optional[Dict[str, Any]]:
        """Persist one hidden pre-change snapshot when this account has no audit history.

        known_keys: account keys already known to have history (from _logged_account_keys).
        When given, the per-account Dataverse query is skipped; keys of accounts that get a
        baseline are added to it so one bulk run never writes two baselines for an account.
        """
        if not self.dv_client:
            raise RuntimeError('Dataverse is not connected.')

        control = str(
            record.get('cr63f_control1') or record.get('cr63f_control') or
            record.get('Control Number') or record.get('Control') or ''
        ).strip()
        area_index = str(
            record.get('cr63f_areaindex') or record.get('Area-Index') or
            record.get('AREA-INDEX') or record.get('Area Index') or ''
        ).strip()
        title_no = str(
            record.get('cr63f_titleno') or record.get('Title No') or
            record.get('TITLE NO') or ''
        ).strip()
        account_keys = _audit_identity_keys(record, master=True)
        history_keys = {key for key in account_keys if key.startswith('MASTER:')} if require_master_identity else account_keys
        for row in self.pending_session_logs:
            if _audit_source_action(row) in REFERENCE_ONLY_ACTIONS:
                continue
            row_keys = _audit_history_keys(row)
            if history_keys.intersection(row_keys):
                return None
        if known_keys is not None:
            if not account_keys:
                raise RuntimeError('Cannot establish an audit baseline without a master ID or Control Number.')
            if keys_lock:
                with keys_lock:
                    if history_keys & known_keys:
                        return None
            elif history_keys & known_keys:
                return None
        else:
            logs_table = self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs')
            clauses = []
            if control:
                # Filter only on Control columns that exist on the logs table;
                # a missing column would make Dataverse reject the query.
                for logical in self._logs_control_columns():
                    clauses.append(f"{logical} eq {_odata_literal(control)}")
            if not control and not record.get('__id'):
                raise RuntimeError('Cannot establish an audit baseline without a master ID or Control Number.')

            history = self.dv_client.fetch_records(
                logs_table, custom_only=False, filter_query=' or '.join(clauses) or None
            )
            history = [
                row for row in history
                if _audit_source_action(row) not in REFERENCE_ONLY_ACTIONS
                and history_keys.intersection(_audit_history_keys(row))
            ]
            if history:
                return None

        baseline = build_log_entry(
            action_type='BASELINE', work_week=get_work_week(datetime.now()),
            now_str=now_str, area_index=area_index,
            index_no=str(record.get('cr63f_indexno') or record.get('Index No') or ''),
            title_no=title_no, field_name='Full Record', old_value='-',
            new_value='Baseline Snapshot', full_record=record,
        )
        saved = self.dv_client.save_log_record(baseline)
        if saved and saved.get('id'):
            baseline['__id'] = saved['id']
        if known_keys is not None:
            if keys_lock:
                with keys_lock:
                    known_keys.update(_audit_history_keys(baseline))
            else:
                known_keys.update(_audit_history_keys(baseline))
        return baseline

    def _save_master(self, fields: Dict[str, Any], record_id: str, etag: Any = None) -> Dict[str, Any]:
        """Update one Land Master account, conditional on ``etag`` when known (If-Match)."""
        if etag:
            return self.dv_client.save_loa_record(fields, record_id=record_id, expected_etag=etag)
        return self.dv_client.save_loa_record(fields, record_id=record_id)

    def _persist_account_updates(self, items_to_save: List[Dict[str, Any]], now_str: str,
                                 progress_label: str = 'Saving accounts and audit rows'):
        """Save changed Land Master accounts one by one, each with its audit snapshot.

        Each item: {'id', 'fields', 'etag', 'previous_fields', 'previous_record', 'snapshot'}.
        Per account: hidden BASELINE if the account has no history, master PATCH
        conditional on 'etag' (If-Match), then the audit snapshot; if the snapshot
        fails, the master change is rolled back. 6 workers; throttled requests are
        retried by the Dataverse client. Returns (saved_items, audit_rows, errors).
        """
        saved_items: List[Dict[str, Any]] = []
        saved_audit_snapshots: List[Dict[str, Any]] = []
        save_errors: List[str] = []
        total = len(items_to_save)
        self._report_progress(0, total, 'Checking audit history...')
        # One logs query for the whole run; if it fails, each account is
        # checked on its own.
        known_keys = None
        if items_to_save:
            try:
                known_keys = self._logged_account_keys()
            except Exception:
                known_keys = None
        progress_lock = threading.Lock()
        completed = [0]

        def save_account(item):
            """Baseline, master update and audit row for ONE account.

            Each account is finished (or rolled back) before the next one starts
            in this worker, so an interruption leaves at most a few accounts in
            flight, never thousands of updated accounts without their audit rows.
            """
            result = {'item': item, 'baseline': None, 'saved': False, 'error': ''}
            try:
                try:
                    result['baseline'] = self._write_baseline_if_missing(
                        item['previous_record'], now_str, known_keys=known_keys, keys_lock=progress_lock)
                except Exception as ex:
                    result['error'] = f"Audit baseline failed for {item.get('id')}: {ex}; Land Master was not updated."
                    return result
                try:
                    res = self._save_master(item['fields'], item['id'], item.get('etag'))
                except Exception as ex:
                    result['error'] = f"Master save error for {item.get('id')}: {ex}"
                    return result
                new_etag = (res or {}).get('etag') or ''
                try:
                    saved_log = self.dv_client.save_log_record(item['snapshot'])
                    if saved_log and saved_log.get('id'):
                        item['snapshot']['__id'] = saved_log['id']
                    item['new_etag'] = new_etag
                    result['saved'] = True
                except Exception as ex:
                    try:
                        self._save_master(item['previous_fields'], item['id'], new_etag)
                        result['error'] = f"Audit save failed for {item['id']} ({ex}); Land Master update was rolled back."
                    except Exception as rollback_error:
                        result['error'] = (
                            f"Audit save failed for {item['id']} ({ex}) and Land Master rollback failed "
                            f"({rollback_error}). Check Dataverse.")
                return result
            finally:
                with progress_lock:
                    completed[0] += 1
                    done = completed[0]
                self._report_progress(done, total, f'{progress_label}: {done} of {total}...')

        if items_to_save:
            with ThreadPoolExecutor(max_workers=min(6, len(items_to_save))) as executor:
                for outcome in executor.map(save_account, items_to_save):
                    if outcome['baseline']:
                        saved_audit_snapshots.append(outcome['baseline'])
                    if outcome['saved']:
                        saved_items.append(outcome['item'])
                        saved_audit_snapshots.append(outcome['item']['snapshot'])
                    if outcome['error']:
                        save_errors.append(outcome['error'])
        self.pending_session_logs.extend(saved_audit_snapshots)
        self.log_records = derive_field_diffs_from_snapshots(saved_audit_snapshots) + self.log_records
        return saved_items, saved_audit_snapshots, save_errors

    def recalculate_payment_fields(self, persist_to_dataverse: bool = False) -> dict:
        """Recalculate Balance and % Paid for every loaded account from its own TCP and Deductible.

        The one-time fix for accounts still holding % Paid on the old fraction
        scale (1 meaning 100%) or a stale Balance. No reference data is used.
        - Accounts without a TCP are skipped (nothing to calculate; their values
          are left as they are) and counted.
        - Only accounts whose stored Balance or % Paid differs from the
          calculated value are changed (Balance exact, % Paid at 0.05).
        persist_to_dataverse=False returns the preview; True saves through the
        same audited, version-checked path as Update from References.
        """
        if not self.loa_records:
            return {'success': False, 'error': 'Load Land Master first.'}
        now_str = datetime.now().strftime('%m/%d/%Y %I:%M %p')
        ww = get_work_week(datetime.now())
        items: List[Dict[str, Any]] = []
        skipped_no_tcp = 0
        samples: List[Dict[str, str]] = []
        for record in self.loa_records:
            tcp = _decimal_amount(_get_field_val(record, list(_TCP_KEYS)))
            if tcp is None:
                skipped_no_tcp += 1
                continue
            ded = _decimal_amount(_get_field_val(record, list(_DEDUCTIBLE_KEYS))) or Decimal(0)
            balance = max(Decimal(0), tcp - ded)
            pct = paid_percent_for_storage(tcp, ded)
            stored_balance = _decimal_amount(_get_field_val(record, ['cr63f_balance', 'Balance']))
            stored_pct = _norm_pct(_get_field_val(record, ['cr63f_paid', '% Paid']))
            balance_differs = stored_balance is None or stored_balance != balance
            pct_differs = stored_pct is None or abs(Decimal(str(stored_pct)) - pct) >= Decimal('0.05')
            if not (balance_differs or pct_differs):
                continue
            if not record.get('__id'):
                continue
            new_record = dict(record)
            new_record.update({'cr63f_balance': balance, 'Balance': format_currency(float(balance)),
                               'cr63f_paid': pct, '% Paid': f"{pct:.1f}%"})
            area_index = str(_get_field_val(record, ['cr63f_areaindex', 'Area-Index']) or '')
            snapshot = build_log_entry(
                action_type='RECALCULATE', work_week=ww, now_str=now_str, area_index=area_index,
                index_no=str(_get_field_val(record, ['cr63f_indexno', 'Index No']) or ''),
                title_no=str(_get_field_val(record, ['cr63f_titleno', 'Title No']) or ''),
                field_name='Full Record', old_value='Previous Snapshot', new_value='Updated Snapshot',
                details='Recalculated Balance and % Paid from TCP and Deductible', full_record=new_record)
            items.append({
                'id': record['__id'], 'etag': record.get('__etag') or '',
                'fields': {'cr63f_balance': balance, 'cr63f_paid': pct},
                'previous_fields': {'cr63f_balance': _none_if_blank(_get_field_val(record, ['cr63f_balance', 'Balance'])),
                                    'cr63f_paid': _none_if_blank(_get_field_val(record, ['cr63f_paid', '% Paid']))},
                'previous_record': dict(record), 'snapshot': snapshot, 'new_record': new_record,
            })
            if len(samples) < 20:
                samples.append({
                    'account': str(_get_field_val(record, ['cr63f_control1', 'Control Number']) or area_index),
                    'balance': f"{_get_field_val(record, ['cr63f_balance', 'Balance']) or '-'} → {format_currency(float(balance))}",
                    'paid': f"{_get_field_val(record, ['cr63f_paid', '% Paid']) or '-'} → {pct:.1f}%",
                })
        result = {
            'success': True, 'to_change': len(items), 'skipped_no_tcp': skipped_no_tcp,
            'unchanged': len(self.loa_records) - len(items) - skipped_no_tcp, 'samples': samples,
            'estimated_requests': len(items) * 3,
        }
        if not persist_to_dataverse:
            return _json_safe(result)
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        saved_items, _, errors = self._persist_account_updates(items, now_str, 'Recalculating accounts')
        saved_ids = set()
        for item in saved_items:
            saved_ids.add(str(item['id']))
            item['new_record']['__display_percent_paid'] = format_calculated_percent(
                _get_field_val(item['new_record'], list(_TCP_KEYS)), _get_field_val(item['new_record'], list(_DEDUCTIBLE_KEYS)))
            _set_version(item['new_record'], item.get('new_etag') or '')
        by_id = {str(item['id']): item['new_record'] for item in saved_items}
        self.loa_records = [by_id.get(str(r.get('__id')), r) if str(r.get('__id')) in saved_ids else r
                            for r in self.loa_records]
        result.update({
            'success': not errors, 'saved_count': len(saved_items), 'errors': errors[:20],
            'error': '; '.join(errors[:3]),
            'message': f"Recalculated Balance and % Paid for {len(saved_items)} account(s)"
                       + (f"; {len(errors)} failed" if errors else '') + '.',
        })
        return _json_safe(result)

    def _report_progress(self, done: int, total: int, message: str) -> None:
        """Update the save progress bar (no-op without a window)."""
        if not self._window:
            return
        try:
            self._window.evaluate_js(
                f"window.onDataverseSaveProgress && window.onDataverseSaveProgress({int(done)}, {int(total)}, {json.dumps(message)})"
            )
        except Exception:
            pass

    def set_window(self, window: webview.Window) -> None:
        self._window = window

    # -- Application updates (GitHub Releases via Velopack) -------------------
    @property
    def _updates(self) -> UpdateService:
        # Created on first use so tests that build Api without __init__ still work.
        svc = self.__dict__.get('_updater')
        if svc is None:
            svc = UpdateService(APP_VERSION, self._emit_update_status)
            self.__dict__['_updater'] = svc
        return svc

    def _emit_update_status(self, status: dict) -> None:
        window = getattr(self, '_window', None)
        if not window:
            return
        try:
            window.evaluate_js(f"window.onUpdateStatus && window.onUpdateStatus({json.dumps(status)})")
        except Exception:
            pass  # The page may not be ready during early startup.

    def get_update_state(self) -> dict:
        return {'ok': True, 'update': self._updates.state()}

    def check_for_updates(self, manual: bool = True) -> dict:
        return {'ok': True, 'update': self._updates.check(bool(manual))}

    def download_update(self) -> dict:
        return {'ok': True, 'update': self._updates.download()}

    def install_update(self) -> dict:
        # Leave the install folder so it is not held open as the working
        # directory while Velopack replaces it.
        try:
            import tempfile
            os.chdir(tempfile.gettempdir())
        except Exception:
            pass
        status = self._updates.apply_and_restart()
        failed = status.get('status') == 'error'
        return {'ok': not failed, 'update': status, 'error': status.get('message') if failed else None}

    def _init_dataverse(self) -> None:
        try:
            cfg = dv.load_config()
            self.dv_client = dv.DataverseClient(cfg)
            if self.dv_client.try_silent_token():
                self._ensure_acquisition_sublevel_schema()
        except Exception as e:
            print(f"Dataverse init notice: {e}")

    def _ensure_acquisition_sublevel_schema(self) -> None:
        """Create the approved optional text field in both source tables if absent."""
        if not self.dv_client:
            return
        for table in (
            self.dv_client.config.get('loa_table', 'cr63f_loamonitoring'),
            self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs'),
        ):
            try:
                self.dv_client.ensure_acquisition_sublevel_column(table)
            except Exception as exc:
                # Keep Dataverse sign-in usable; surface migration failure in app logs.
                print(f"Acquisition Sub-Level schema notice for {table}: {exc}")

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

        rem_conv = self.settings.get('remembered_conversion_file')
        if rem_conv and os.path.isfile(rem_conv):
            try:
                headers, rows = read_tabular_file(rem_conv)
                self.conversion_rows = rows
                self.conversion_engine.load_rows(rows)
                self.conversion_source_name = f"File: {os.path.basename(rem_conv)}"
            except Exception as e:
                print(f"Could not auto-load remembered conversion file: {e}")

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
            'conversion_table': cfg.get('conversion_table', 'cr63f_landconversionreference'),
            'logs_table': cfg.get('logs_table', 'cr63f_loamonitoringlogs'),
            'solution_name': cfg.get('solution_name', 'LOA MONITORING'),
            'solution_id': cfg.get('solution_id', '34c6d4d3-b4b0-f111-aaac-7ced8d3b4de7'),
            'finance_loaded_count': len(self.finance_rows),
            'finance_source': self.finance_source_name,
            'remembered_finance_file': self.settings.get('remembered_finance_file', ''),
            'recent_finance_files': self.settings.get('recent_finance_files', []),
            'conversion_loaded_count': len(self.conversion_rows),
            'conversion_source': self.conversion_source_name,
            'remembered_conversion_file': self.settings.get('remembered_conversion_file', ''),
            'recent_conversion_files': self.settings.get('recent_conversion_files', []),
            'theme': self.settings.get('theme', 'light'),
            'table_layout': self.settings.get('table_layout', {}),
            'loa_record_count': len(self.loa_records),
            'app_version': APP_VERSION,
            'update': self._updates.state(),
        }

    # -- MSAL Auth -------------------------------------------------------------
    def start_sign_in(self) -> dict:
        if not self.dv_client:
            self._init_dataverse()
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse client could not be initialized.'}
        try:
            if self.dv_client.try_silent_token():
                self._ensure_acquisition_sublevel_schema()
                return {'success': True, 'already_connected': True}

            # A sign-in is already waiting for its code: show the same code
            # instead of starting a second flow (whose expiry would later
            # report a failure over the one the user completed).
            if self.active_flow and getattr(self, '_auth_status', {}).get('in_progress'):
                flow = self.active_flow
                return {
                    'success': True,
                    'user_code': flow.get('user_code'),
                    'verification_uri': flow.get('verification_uri'),
                    'message': flow.get('message'),
                }

            flow = self.dv_client.begin_device_flow()
            self.active_flow = flow
            self._auth_status = {'in_progress': True, 'connected': False, 'error': None}

            # Run polling in background daemon thread
            def _wait_for_auth():
                try:
                    self.dv_client.complete_device_flow()
                    self._ensure_acquisition_sublevel_schema()
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
            self._ensure_acquisition_sublevel_schema()
            self.active_flow = None
            return res
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def sign_out(self) -> dict:
        if self.dv_client:
            self.dv_client.sign_out()
        return {'success': True}

    # -- Live Finance & Conversion Lookups -------------------------------------
    def lookup_finance(
        self,
        type_val: Any,
        index_no: Any,
        area_index: Any,
        title_mod: Any,
        mode: Any = None
    ) -> dict:
        """Called live from the UI form as the user modifies any key field."""
        res = self.finance_engine.lookup(
            type_val=type_val,
            index_no=index_no,
            area_index_val=area_index,
            title_mod_val=title_mod,
            mode_val=mode,
        )
        return res

    def lookup_conversion(
        self,
        area_index: Any
    ) -> dict:
        """Called live from the UI form to populate Conversion Status."""
        res = self.conversion_engine.lookup(
            area_index_val=area_index
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
        file_path = _dialog_path(result)
        if not file_path:
            return {'success': False, 'cancelled': True}
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
            # Loaded by the user in this session: automatic reloads keep it.
            self._finance_file_this_session = True

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

    def sync_finance_from_dataverse(self, table_name: Optional[str] = None, keep_loaded_file: bool = False) -> dict:
        """Load the finance reference from Dataverse.

        keep_loaded_file=True (used by the automatic reload with Land Master):
        a finance file the user loaded is kept instead of being silently replaced
        by the Dataverse table, which may be older.
        """
        if keep_loaded_file and getattr(self, '_finance_file_this_session', False) and str(self.finance_source_name).startswith('File:'):
            return {'success': True, 'skipped': True, 'source': self.finance_source_name,
                    'message': 'Kept the loaded finance file.'}
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
            self._finance_file_this_session = False

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

    def clear_finance_reference(self) -> dict:
        self.finance_rows = []
        self.finance_engine = FinanceLookupEngine()
        self.finance_source_name = "None"
        self.settings['remembered_finance_file'] = ""
        save_app_settings(self.settings)
        return {'success': True}

    def delete_finance_record(self, record_id_or_idx: Any) -> dict:
        return self._delete_reference_row('finance', record_id_or_idx)

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

            baselines = getattr(self, '_reference_edit_baselines', {})
            self._reference_edit_baselines = baselines
            # Keyed by the id the screen uses (row ID or position), the same key
            # _save_staged_reference_records looks up. Keying by the Dataverse ID
            # while the screen still sent a position lost the "before" values, so
            # the save saw no change and silently wrote nothing.
            key = ('finance', target_id)
            if key not in baselines or baselines[key][0] is not found_row:
                baselines[key] = (found_row, dict(found_row))
            self.finance_rows[found_idx][field_name] = new_value
            # Reload engine so live lookup calculations use updated values
            self.finance_engine.load_rows(self.finance_rows)
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

    def download_finance_template(self, destination_path: Optional[str] = None) -> dict:
        """Generate and save standard Excel template for Finance Reference."""
        default_name = "Finance_Reference_Template.xlsx"
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
            save_path = _dialog_path(result)
            if not save_path:
                return {'success': False, 'cancelled': True}

        if not save_path.lower().endswith('.xlsx'):
            save_path += '.xlsx'

        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "APPENDED_FINANCE"

            headers = [
                "TYPE",
                "Corridor Index",
                "Area Index",
                "Title Mod",
                "Total Contract Price",
                "Total Payment Deductible to TCP",
                "% Paid",
                "Balance"
            ]

            sample_rows = [
                ["ROW", "C-101", "", "T-12345", 1500000.00, 500000.00, "33.33%", 1000000.00],
                ["NON-ROW", "", "AI-202", "T-67890", 2500000.00, 1000000.00, "40.00%", 1500000.00],
                ["NON-ROW", "", "", "T-30303", 800000.00, 200000.00, "25.00%", 600000.00]
            ]

            header_fill = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
            header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            data_font = Font(name="Segoe UI", size=10)
            thin_border = Border(
                left=Side(style="thin", color="E2E8F0"),
                right=Side(style="thin", color="E2E8F0"),
                top=Side(style="thin", color="E2E8F0"),
                bottom=Side(style="thin", color="E2E8F0")
            )

            # Write header
            for col_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=col_idx, value=h)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            # Write sample data
            for row_idx, r in enumerate(sample_rows, start=2):
                for col_idx, val in enumerate(r, start=1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    cell.font = data_font
                    cell.border = thin_border
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0.00'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif str(val).endswith('%'):
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

            ws.row_dimensions[1].height = 26
            for col_idx in range(1, len(headers) + 1):
                col_letter = get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = 24

            wb.save(save_path)
            return {'success': True, 'file_path': save_path}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def sync_uploaded_finance_to_dataverse(self) -> dict:
        """Upload the loaded finance rows to Dataverse without creating duplicates."""
        return self._sync_reference_rows('finance')

    @staticmethod
    def _reference_row_key(kind: str, row: Dict[str, Any]) -> tuple:
        """Natural key of a reference row, used to find the Dataverse row a file row updates.

        Finance: TYPE + Corridor Index + Area Index + Title (title prefix removed).
        Conversion: AREA-INDEX, or the title number when the row has no AREA-INDEX.
        An all-blank key cannot be matched.
        """
        get = FinanceLookupEngine._get_field
        if kind == 'finance':
            key = (
                norm_key(get(row, ['TYPE', 'Type', 'cr63f_type'])),
                norm_key(get(row, ['Corridor Index', 'CORRIDOR INDEX', 'INDEX NO', 'cr63f_corridorindex', 'cr63f_indexno'])),
                norm_key(get(row, ['Area Index', 'AREA INDEX', 'AREA-INDEX', 'cr63f_areaindex'])),
                norm_key(_strict_title(get(row, ['Title Mod', 'TITLE MOD', 'TITLE NO', 'Title No', 'cr63f_titlemod', 'cr63f_titleno']))),
            )
            return key if any(key[1:]) else ()
        area = norm_key(get(row, ['AREA-INDEX', 'AREA INDEX', 'Area Index', 'cr63f_areaindex']))
        if area:
            return ('AREA', area)
        title = norm_key(get(row, ['TITLE NUMBER', 'Title Number', 'TITLE NO', 'cr63f_titlenumber']))
        return ('TITLE', title) if title else ()

    def _sync_reference_rows(self, kind: str) -> dict:
        """Create or update reference rows in Dataverse from the loaded rows.

        - A row loaded from Dataverse (it has an ID) is updated, and only the
          values that differ are sent.
        - A row from a file is matched to the existing Dataverse row with the
          same key (_reference_row_key) and updated; it is created only when no
          row has that key. Re-syncing the same file therefore never duplicates.
        - A key shared by several Dataverse rows, or by several file rows, is
          skipped and reported (never guessed).
        - One summary audit row is written for the sync.
        """
        if not self.dv_client:
            self._init_dataverse()
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected. Please connect in Settings.'}
        rows = self.finance_rows if kind == 'finance' else self.conversion_rows
        if not rows:
            return {'success': False, 'error': f'No {kind} rows loaded to sync.'}
        default_table = 'cr63f_financereference' if kind == 'finance' else 'cr63f_landconversionreference'
        table_name = self.dv_client.config.get(f'{kind}_table', default_table)
        save = self.dv_client.save_finance_record if kind == 'finance' else self.dv_client.save_conversion_record
        progress_js = 'onFinanceSyncProgress' if kind == 'finance' else 'onConversionSyncProgress'
        try:
            existing = self.dv_client.fetch_records(table_name, custom_only=True)
        except Exception as exc:
            return {'success': False, 'error': f'Could not read {table_name} before syncing: {exc}'}
        existing_by_id = {str(r.get('__id')): r for r in existing if r.get('__id')}
        existing_by_key: Dict[tuple, List[Dict[str, Any]]] = {}
        for r in existing:
            key = self._reference_row_key(kind, r)
            if key:
                existing_by_key.setdefault(key, []).append(r)
        file_key_counts: Dict[tuple, int] = {}
        for row in rows:
            if not row.get('__id'):
                key = self._reference_row_key(kind, row)
                if key:
                    file_key_counts[key] = file_key_counts.get(key, 0) + 1

        counts = {'created': 0, 'updated': 0, 'unchanged': 0, 'skipped': 0, 'failed': 0}
        errors: List[str] = []
        skipped: List[str] = []
        total_rows = len(rows)
        for i, row in enumerate(rows):
            label = ' / '.join(k for k in self._reference_row_key(kind, row) if k) or f'row {i + 1}'
            try:
                payload = {k: v for k, v in row.items() if not k.startswith('_')}
                target = existing_by_id.get(str(row.get('__id'))) if row.get('__id') else None
                if target is None:
                    key = self._reference_row_key(kind, row)
                    matches = existing_by_key.get(key, []) if key else []
                    if key and file_key_counts.get(key, 0) > 1:
                        counts['skipped'] += 1
                        skipped.append(f'{label} (appears {file_key_counts[key]} times in the file)')
                        continue
                    if len(matches) > 1:
                        counts['skipped'] += 1
                        skipped.append(f'{label} (matches {len(matches)} existing Dataverse rows)')
                        continue
                    target = matches[0] if matches else None
                if target is not None:
                    changed = {
                        k: v for k, v in payload.items()
                        if _reference_value_differs(k, target.get(k), v)
                    }
                    if not changed:
                        counts['unchanged'] += 1
                    else:
                        save(changed, record_id=target['__id'])
                        counts['updated'] += 1
                    row['__id'] = target['__id']
                else:
                    res = save(payload)
                    if res.get('id'):
                        row['__id'] = res['id']
                    counts['created'] += 1
            except Exception as e:
                counts['failed'] += 1
                errors.append(f'{label}: {e}')
            finally:
                if self._window and (i % 5 == 0 or i == total_rows - 1):
                    try:
                        self._window.evaluate_js(f"window.{progress_js} && window.{progress_js}({i + 1}, {total_rows})")
                    except Exception:
                        pass

        summary = (f"{counts['created']} created, {counts['updated']} updated, {counts['unchanged']} unchanged, "
                   f"{counts['skipped']} skipped, {counts['failed']} failed")
        audit_error = ''
        if counts['created'] or counts['updated']:
            source = self.finance_source_name if kind == 'finance' else self.conversion_source_name
            log = build_log_entry(
                action_type=f'{kind.upper()}_UPDATE', work_week=get_work_week(datetime.now()),
                now_str=datetime.now().strftime('%m/%d/%Y %I:%M %p'),
                area_index='', index_no='', title_no='', field_name='Bulk Sync',
                old_value='-', new_value=summary,
                details=f'Synced {source} to {table_name}: {summary}')
            try:
                saved_log = self.dv_client.save_log_record(log)
                log['__id'] = (saved_log or {}).get('id', '')
                self.pending_session_logs.append(dict(log))
                self._add_log_records([log])
            except Exception as exc:
                audit_error = f'The sync summary could not be written to Entry Logs: {exc}'
        if counts['created'] or counts['updated']:
            self._set_reference_source(kind, f"Dataverse: {table_name}")
        problems = errors + ([audit_error] if audit_error else [])
        return {
            'success': not errors and not audit_error,
            'total_rows': total_rows,
            'synced_count': counts['created'] + counts['updated'] + counts['unchanged'],
            'error_count': counts['failed'],
            'skipped_count': counts['skipped'],
            'counts': counts,
            'errors': problems[:5],
            'skipped': skipped[:5],
            'error': '; '.join(problems[:3]),
            'message': f"Sync to {table_name}: {summary}." + (
                f" Skipped: {'; '.join(skipped[:3])}{' …' if len(skipped) > 3 else ''}" if skipped else ''),
        }

    def _delete_reference_row(self, kind: str, record_id_or_idx: Any) -> dict:
        """Delete one Finance or Conversion row.

        A row stored in Dataverse is removed locally only after Dataverse
        confirms the delete, and the delete is audited like a reference edit
        (audit row first; removed again if the delete fails). A row that exists
        only in a loaded file is just removed from the loaded rows.
        """
        rows = self.finance_rows if kind == 'finance' else self.conversion_rows
        engine = self.finance_engine if kind == 'finance' else self.conversion_engine
        target_id = str(record_id_or_idx)
        found_idx = next((i for i, r in enumerate(rows) if str(r.get('__id') or '') == target_id), -1)
        if found_idx == -1:
            try:
                num_idx = int(record_id_or_idx)
                if 0 <= num_idx < len(rows):
                    found_idx = num_idx
            except (ValueError, TypeError):
                pass
        if found_idx == -1:
            return {'success': False, 'error': f'{kind.title()} record not found.'}
        found_row = rows[found_idx]
        if found_row.get('__id'):
            if not self.dv_client:
                return {'success': False, 'error': 'Dataverse is not connected; nothing was deleted.'}
            default_table = 'cr63f_financereference' if kind == 'finance' else 'cr63f_landconversionreference'
            target_table = self.dv_client.config.get(f'{kind}_table', default_table)
            key = ' / '.join(k for k in self._reference_row_key(kind, found_row) if k)
            log = build_log_entry(
                action_type=f'{kind.upper()}_UPDATE', work_week=get_work_week(datetime.now()),
                now_str=datetime.now().strftime('%m/%d/%Y %I:%M %p'),
                area_index=str(_get_field_val(found_row, ['cr63f_areaindex', 'AREA-INDEX', 'Area Index']) or ''),
                index_no='', title_no='', field_name='Deleted Row', old_value=key or 'Row', new_value='Deleted',
                details=f'Deleted {kind.title()} Reference row {key}',
                full_record={k: v for k, v in found_row.items() if not k.startswith('_')})
            try:
                saved_log = self.dv_client.save_log_record(log)
                log['__id'] = (saved_log or {}).get('id', '')
            except Exception as exc:
                return {'success': False, 'error': f'Nothing was deleted: the audit row could not be written ({exc}).'}
            try:
                info = self.dv_client.get_entity_info(target_table)
                self.dv_client.delete(info['entity_set'], found_row['__id'])
            except Exception as exc:
                try:
                    if log.get('__id'):
                        self.dv_client.delete_log_record(log['__id'])
                except Exception:
                    pass
                return {'success': False, 'error': f'Could not delete the row in Dataverse: {exc}'}
            self.pending_session_logs.append(dict(log))
            self._add_log_records([log])
        rows.pop(found_idx)
        engine.load_rows(rows)
        return {
            'success': True,
            'remaining_count': len(rows),
            'message': f'{kind.title()} record deleted.',
        }

    def _set_reference_source(self, kind: str, name: str) -> None:
        if kind == 'finance':
            self.finance_source_name = name
        else:
            self.conversion_source_name = name

    # -- Reference Data (LAND_CONVERSION_REFERENCE) ---------------------------
    def load_conversion_file_dialog(self) -> dict:
        if not self._window:
            return {'success': False, 'error': 'Window not attached.'}
        file_types = ('Excel and CSV Files (*.xlsx;*.xlsm;*.csv)', 'All files (*.*)')
        result = self._window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=file_types
        )
        file_path = _dialog_path(result)
        if not file_path:
            return {'success': False, 'cancelled': True}
        return self._load_conversion_file_path(file_path)

    def load_conversion_file_recent(self, file_path: str) -> dict:
        return self._load_conversion_file_path(file_path)

    def _load_conversion_file_path(self, file_path: str) -> dict:
        try:
            headers, rows = read_tabular_file(file_path)
            self.conversion_rows = rows
            self.conversion_engine.load_rows(rows)
            filename = os.path.basename(file_path)
            self.conversion_source_name = f"File: {filename}"
            # Loaded by the user in this session: automatic reloads keep it.
            self._conversion_file_this_session = True

            # Save remembered path and recent list
            self.settings['remembered_conversion_file'] = file_path
            recent = self.settings.get('recent_conversion_files', [])
            if file_path in recent:
                recent.remove(file_path)
            recent.insert(0, file_path)
            self.settings['recent_conversion_files'] = recent[:6]
            save_app_settings(self.settings)

            return {
                'success': True,
                'row_count': len(rows),
                'headers': headers,
                'source': self.conversion_source_name,
                'file_path': file_path,
                'sample_rows': rows
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def sync_conversion_from_dataverse(self, table_name: Optional[str] = None, keep_loaded_file: bool = False) -> dict:
        """Load the conversion reference from Dataverse.

        keep_loaded_file=True (used by the automatic reload with Land Master):
        a conversion file the user loaded is kept instead of being silently replaced
        by the Dataverse table, which may be older.
        """
        if keep_loaded_file and getattr(self, '_conversion_file_this_session', False) and str(self.conversion_source_name).startswith('File:'):
            return {'success': True, 'skipped': True, 'source': self.conversion_source_name,
                    'message': 'Kept the loaded conversion file.'}
        if not self.dv_client:
            self._init_dataverse()
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse client not configured.'}
        try:
            target_table = table_name or self.dv_client.config.get('conversion_table') or 'cr63f_landconversionreference'
            try:
                rows = self.dv_client.fetch_records(target_table, custom_only=True)
            except Exception as fe:
                if target_table != 'cr63f_landconversionreference':
                    target_table = 'cr63f_landconversionreference'
                    self.dv_client.config['conversion_table'] = target_table
                    try:
                        dv.save_config(self.dv_client.config)
                    except Exception:
                        pass
                    rows = self.dv_client.fetch_records(target_table, custom_only=True)
                else:
                    raise fe

            self.conversion_rows = rows
            self.conversion_engine.load_rows(rows)
            self.conversion_source_name = f"Dataverse: {target_table}"
            self._conversion_file_this_session = False

            conv_attrs = self.dv_client.get_entity_attributes(target_table)
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
            for log_name, meta in conv_attrs.items():
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
                            'isStatus': 'status' in low or 'class' in low,
                            'isDate': meta.get('type') == 'DateTime'
                        })

            if not clean_headers and rows:
                clean_headers = [k for k in rows[0].keys() if not k.startswith('__') and not k.startswith('cr63f_')]

            return {
                'success': True,
                'row_count': len(rows),
                'source': self.conversion_source_name,
                'table_name': target_table,
                'headers': clean_headers,
                'columns': clean_cols,
                'sample_rows': rows
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def clear_conversion_reference(self) -> dict:
        self.conversion_rows = []
        self.conversion_engine = ConversionLookupEngine()
        self.conversion_source_name = "None"
        self.settings['remembered_conversion_file'] = ""
        save_app_settings(self.settings)
        return {'success': True}

    def delete_conversion_record(self, record_id_or_idx: Any) -> dict:
        return self._delete_reference_row('conversion', record_id_or_idx)

    def update_conversion_cell(self, record_id_or_idx: Any, field_name: str, new_value: Any) -> dict:
        """Direct inline edit of a conversion cell in table."""
        try:
            target_id = str(record_id_or_idx)
            found_idx = -1
            found_row = None

            for idx, r in enumerate(self.conversion_rows):
                if r.get('__id') == target_id:
                    found_idx = idx
                    found_row = r
                    break

            if found_idx == -1:
                try:
                    num_idx = int(record_id_or_idx)
                    if 0 <= num_idx < len(self.conversion_rows):
                        found_idx = num_idx
                        found_row = self.conversion_rows[num_idx]
                except (ValueError, TypeError):
                    pass

            if found_row is None or found_idx == -1:
                return {'success': False, 'error': 'Conversion row not found.'}

            baselines = getattr(self, '_reference_edit_baselines', {})
            self._reference_edit_baselines = baselines
            # Keyed by the id the screen uses (row ID or position), the same key
            # _save_staged_reference_records looks up. Keying by the Dataverse ID
            # while the screen still sent a position lost the "before" values, so
            # the save saw no change and silently wrote nothing.
            key = ('conversion', target_id)
            if key not in baselines or baselines[key][0] is not found_row:
                baselines[key] = (found_row, dict(found_row))
            self.conversion_rows[found_idx][field_name] = new_value
            self.conversion_engine.load_rows(self.conversion_rows)
            return {'success': True, 'field': field_name, 'value': new_value}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def get_conversion_preview(self, limit: Optional[int] = None) -> dict:
        rows = self.conversion_rows if limit is None else self.conversion_rows[:limit]
        return {
            'source': self.conversion_source_name,
            'total_count': len(self.conversion_rows),
            'rows': rows
        }

    def download_conversion_template(self, destination_path: Optional[str] = None) -> dict:
        """Generate and save standard Excel template for Land Conversion Reference."""
        default_name = "Land_Conversion_Reference_Template.xlsx"
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
            save_path = _dialog_path(result)
            if not save_path:
                return {'success': False, 'cancelled': True}

        if not save_path.lower().endswith('.xlsx'):
            save_path += '.xlsx'

        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "CONVERSION_REFERENCE"

            headers = [
                "AREA-INDEX",
                "TITLE NUMBER",
                "LOT NUMBER/SURVEY NUMBER",
                "DAR STATUS",
                "NIA STATUS",
                "CONVERSION CLASS",
                "DAR BATCH",
                "DAR UPDATE",
                "ENTITY",
                "MUNICIPALITY/CITY",
                "BARANGAY",
                "AREA APPLIED (HAS.)"
            ]

            sample_rows = [
                ["TST-1383", "TCT-T-10004", "101, PSD-10001", "DAR SECURED", "NIA INSPECTION DONE", "Class A", "DAR-BAT3-19", "11/23/2025", "MBHC", "NASUGBU", "LATAG", 2.25],
                ["TST-7742", "TCT-T-100005", "1-A-1-A", "DAR FOR APPLICATION", "NIA FOR APPLICATION", "Class A", "DAR-BAT3-36", "11/23/2025", "PIC", "TUY", "MAGAHIS", 0.381],
            ]

            header_fill = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")
            header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            data_font = Font(name="Segoe UI", size=10)
            thin_border = Border(
                left=Side(style="thin", color="E2E8F0"),
                right=Side(style="thin", color="E2E8F0"),
                top=Side(style="thin", color="E2E8F0"),
                bottom=Side(style="thin", color="E2E8F0")
            )

            for col_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=col_idx, value=h)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            for row_idx, r in enumerate(sample_rows, start=2):
                for col_idx, val in enumerate(r, start=1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    cell.font = data_font
                    cell.border = thin_border
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0.0000'
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

            ws.row_dimensions[1].height = 26
            for col_idx in range(1, len(headers) + 1):
                col_letter = get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = 24

            wb.save(save_path)
            return {'success': True, 'file_path': save_path}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def sync_uploaded_conversion_to_dataverse(self) -> dict:
        """Upload the loaded conversion rows to Dataverse without creating duplicates."""
        return self._sync_reference_rows('conversion')

    def populate_loa_records_from_references(
        self,
        records: Optional[List[Dict[str, Any]]] = None,
        persist_to_dataverse: bool = False
    ) -> dict:
        """Auto-populate financial fields and Conversion Status for LOA records using FinanceLookupEngine & ConversionLookupEngine.
        If persist_to_dataverse=True, immediately saves modified master records and writes audit diff logs to Dataverse."""
        target_records = records if records is not None else self.loa_records
        if persist_to_dataverse and records is not None and self.loa_records:
            # Save from the app's current copy of each account (with the
            # version it was loaded at), not the screen's older baseline, so a
            # retry after a partial failure does not redo saved accounts and a
            # change made by someone else is detected (If-Match).
            current_by_id = {str(row.get('__id')): row for row in self.loa_records if row.get('__id')}
            target_records = [current_by_id.get(str(row.get('__id')), row) if row.get('__id') else row
                              for row in records]
        if not target_records:
            return {'success': False, 'error': 'No LOA records available to update.'}
        has_fin = bool(self.finance_rows or getattr(self.finance_engine, 'raw_rows', []))
        has_conv = bool(self.conversion_rows or getattr(self.conversion_engine, 'raw_rows', []))
        if not has_fin and not has_conv:
            return {'success': False, 'error': 'No Reference Data available for lookup. Please load Finance or Conversion Data first.'}

        matched_count = 0
        updated_count = 0
        updated_records = []
        now_str = datetime.now().strftime('%m/%d/%Y %I:%M %p')

        items_to_save = []
        all_diff_logs = []
        missing_record_ids = []
        # Accounts whose reference key matches several rows with different
        # values. They are skipped (never guessed) and reported back.
        ambiguous_accounts: List[str] = []

        for r in target_records:
            type_val = r.get('cr63f_type') or r.get('Type') or r.get('TYPE') or ''
            index_no = r.get('cr63f_indexno') or r.get('Index No') or r.get('INDEX NO') or ''
            area_index = r.get('cr63f_areaindex') or r.get('Area-Index') or r.get('AREA-INDEX') or r.get('Area Index') or ''
            title_no = r.get('cr63f_titleno') or r.get('Title No') or r.get('TITLE NO') or ''
            title_mod = r.get('Title Mod') or r.get('cr63f_titlemod') or title_no
            ww = get_work_week(datetime.now())

            new_r = dict(r)
            record_matched = False
            record_changed = False
            patch_fields = {}
            row_logs = []

            # 1. Finance Lookup
            fin_changed_fields = []
            if has_fin:
                res_fin = self.finance_engine.lookup(
                    type_val=type_val,
                    index_no=index_no,
                    area_index_val=area_index,
                    title_mod_val=title_mod,
                    mode_val=r.get('new_mode') or r.get('Mode') or r.get('MODE') or '',
                )
                if res_fin.get('ambiguous'):
                    ambiguous_accounts.append(
                        f"{area_index or index_no or title_no} (Finance: {res_fin.get('ambiguous_count')} conflicting rows of the same TYPE)"
                    )
                if res_fin.get('matched'):
                    record_matched = True
                    finance_row = res_fin.get('matched_row') or {}
                    tcp_raw = _get_field_val(finance_row, [
                        'Total Contract Price', 'Total Contract Price (Finance)',
                        'TotalContractPrice', 'TOTAL CONTRACT PRICE', 'TCP', 'tcp',
                        'cr63f_totalcontractprice'
                    ])
                    ded_raw = _get_field_val(finance_row, [
                        'Total Payment Deductible to TCP', 'Total Payment Deductible',
                        'Payment Deductible', 'Payment Deductible to TCP',
                        'TOTAL PAYMENT DEDUCTIBLE TO TCP',
                        'cr63f_totalpaymentdeductibletotcp'
                    ])
                    tcp_val = _decimal_amount(tcp_raw)
                    if tcp_val is None:
                        tcp_val = _decimal_amount(res_fin.get('total_contract_price'))
                    ded_val = _decimal_amount(ded_raw)
                    if ded_val is None:
                        ded_val = _decimal_amount(res_fin.get('total_payment_deductible'))
                    tcp_fmt = format_currency(float(tcp_val)) if tcp_val is not None else ''
                    ded_fmt = format_currency(float(ded_val)) if ded_val is not None else ''

                    curr_tcp = _extract_num(new_r, [
                        'Total Contract Price', 'cr63f_totalcontractprice', 'TOTAL CONTRACT PRICE',
                        'TotalContractPrice', 'TCP', 'tcp', 'Total Contract Price (Finance)', 'Total Contract Price (TCP)'
                    ])
                    curr_ded = _extract_num(new_r, [
                        'Total Payment Deductible to TCP', 'Payment Deductible', 'cr63f_totalpaymentdeductibletotcp',
                        'Total Payment Deductible', 'Payment Deductible to TCP', 'Total Payment deductible to TCP',
                        'TOTAL PAYMENT DEDUCTIBLE TO TCP', 'Payment Deductibles', 'cr63f_paymentdeductible',
                        'Payment Deductible (Finance)', 'Total Payment Deductible (Finance)'
                    ])
                    curr_bal = _extract_num(new_r, [
                        'Balance', 'cr63f_balance', 'BALANCE', 'Remaining Balance', 'Balance (Finance)', 'cr63f_remainingbalance'
                    ])

                    curr_tcp_fmt = str(r.get('Total Contract Price') or r.get('cr63f_totalcontractprice') or (f"₱{curr_tcp:,.2f}" if curr_tcp is not None else ''))
                    curr_ded_fmt = str(r.get('Total Payment Deductible to TCP') or r.get('Payment Deductible') or r.get('cr63f_totalpaymentdeductibletotcp') or (f"₱{curr_ded:,.2f}" if curr_ded is not None else ''))
                    curr_bal_fmt = str(r.get('Balance') or r.get('cr63f_balance') or (f"₱{curr_bal:,.2f}" if curr_bal is not None else ''))

                    curr_tcp_exact = _decimal_amount(_get_field_val(new_r, [
                        'cr63f_totalcontractprice', 'Total Contract Price', 'TCP'
                    ]))
                    curr_ded_exact = _decimal_amount(_get_field_val(new_r, [
                        'cr63f_totalpaymentdeductibletotcp',
                        'Total Payment Deductible to TCP', 'Payment Deductible'
                    ]))
                    if tcp_val is not None and _is_different_num(curr_tcp_exact, tcp_val):
                        fin_changed_fields.extend(['cr63f_totalcontractprice', 'Total Contract Price'])
                        patch_fields['cr63f_totalcontractprice'] = tcp_val
                        new_r['cr63f_totalcontractprice'] = tcp_val
                        new_r['Total Contract Price'] = tcp_fmt
                        row_logs.append(build_log_entry('REFERENCE_UPDATE', ww, now_str, area_index, index_no, title_no, 'Total Contract Price', curr_tcp_fmt, tcp_fmt, full_record=new_r))

                    if ded_val is not None and _is_different_num(curr_ded_exact, ded_val):
                        fin_changed_fields.extend(['cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP', 'Payment Deductible'])
                        patch_fields['cr63f_totalpaymentdeductibletotcp'] = ded_val
                        new_r['cr63f_totalpaymentdeductibletotcp'] = ded_val
                        new_r['Total Payment Deductible to TCP'] = ded_fmt
                        new_r['Payment Deductible'] = ded_fmt
                        row_logs.append(build_log_entry('REFERENCE_UPDATE', ww, now_str, area_index, index_no, title_no, 'Total Payment Deductible to TCP', curr_ded_fmt, ded_fmt, full_record=new_r))

                    # Balance and % Paid are always derived from the final TCP and
                    # Deductible and saved whenever the stored value differs, so
                    # Dataverse never keeps a stale Balance or % Paid. The
                    # reference table's own % Paid is never copied.
                    final_tcp = tcp_val if tcp_val is not None else curr_tcp_exact
                    final_ded = ded_val if ded_val is not None else curr_ded_exact
                    if final_tcp is not None and final_ded is not None:
                        calculated_balance = max(Decimal(0), final_tcp - final_ded)
                        curr_bal_exact = _decimal_amount(_get_field_val(new_r, ['cr63f_balance', 'Balance']))
                        if curr_bal_exact is None or curr_bal_exact != calculated_balance:
                            fin_changed_fields.extend(['cr63f_balance', 'Balance'])
                            patch_fields['cr63f_balance'] = calculated_balance
                            new_r['cr63f_balance'] = calculated_balance
                            new_r['Balance'] = format_currency(float(calculated_balance))
                            row_logs.append(build_log_entry('REFERENCE_UPDATE', ww, now_str, area_index, index_no, title_no, 'Balance', curr_bal_fmt, format_currency(float(calculated_balance)), full_record=new_r))

                        stored_pct = paid_percent_for_storage(final_tcp, final_ded)
                        curr_pct_raw = _get_field_val(new_r, ['cr63f_paid', '% Paid'])
                        curr_pct = _norm_pct(curr_pct_raw)
                        if stored_pct is not None and (
                            curr_pct is None or abs(Decimal(str(curr_pct)) - stored_pct) >= Decimal('0.05')
                        ):
                            fin_changed_fields.extend(['cr63f_paid', '% Paid'])
                            patch_fields['cr63f_paid'] = stored_pct
                            new_r['cr63f_paid'] = f"{stored_pct:.1f}%"
                            new_r['% Paid'] = f"{stored_pct:.1f}%"
                            row_logs.append(build_log_entry('REFERENCE_UPDATE', ww, now_str, area_index, index_no, title_no, '% Paid', curr_pct_raw, stored_pct, full_record=new_r))

                    display_pct = format_calculated_percent(final_tcp, final_ded)
                    if display_pct:
                        new_r['__display_percent_paid'] = display_pct

                    fin_changed = bool(fin_changed_fields)
                    new_r['__finance_matched'] = True
                    new_r['__finance_changed'] = fin_changed
                    new_r['__finance_changed_fields'] = fin_changed_fields
                    if fin_changed:
                        record_changed = True

            # 2. Conversion Lookup
            conv_changed_fields = []
            if has_conv:
                res_conv = self.conversion_engine.lookup(
                    area_index_val=area_index
                )
                if res_conv.get('ambiguous'):
                    ambiguous_accounts.append(
                        f"{area_index} (Conversion: {res_conv.get('ambiguous_count')} conflicting rows)"
                    )
                if res_conv.get('matched'):
                    record_matched = True
                    conv_status = res_conv.get('conversion_status') or ''
                    curr_status = str(new_r.get('cr63f_conversionstatus') or new_r.get('Conversion Status') or '').strip()
                    if conv_status and norm_key(conv_status) != norm_key(curr_status):
                        conv_changed_fields.extend(['cr63f_conversionstatus', 'Conversion Status'])
                        patch_fields['cr63f_conversionstatus'] = conv_status
                        patch_fields['Conversion Status'] = conv_status
                        new_r['cr63f_conversionstatus'] = conv_status
                        new_r['Conversion Status'] = conv_status
                        row_logs.append(build_log_entry('REFERENCE_UPDATE', ww, now_str, area_index, index_no, title_no, 'Conversion Status', curr_status, conv_status, full_record=new_r))

                    conv_changed = bool(conv_changed_fields)
                    new_r['__conversion_matched'] = True
                    new_r['__conversion_changed'] = conv_changed
                    new_r['__conversion_changed_fields'] = conv_changed_fields
                    if conv_changed:
                        record_changed = True

            # A reference value actually changed (not just a Balance/% Paid
            # correction) and the account has no Endorsed Date: pre-fill it
            # with today's date. Only a truly blank date is filled; a stored
            # value, including a 1900 placeholder, is never overwritten.
            reference_value_changed = any(
                key in patch_fields for key in
                ('cr63f_totalcontractprice', 'cr63f_totalpaymentdeductibletotcp', 'cr63f_conversionstatus')
            )
            endorsed_changed_fields = []
            if reference_value_changed and _endorsed_is_blank(new_r):
                today = datetime.now().strftime('%m/%d/%Y')
                patch_fields['cr63f_endorseddate'] = today
                new_r['cr63f_endorseddate'] = today
                new_r['Endorsed Date'] = today
                new_r['work_week'] = new_r['Work Week'] = get_work_week(today)
                endorsed_changed_fields = ['cr63f_endorseddate', 'Endorsed Date']
                row_logs.append(build_log_entry('REFERENCE_UPDATE', ww, now_str, area_index, index_no, title_no, 'Endorsed Date', '', today, full_record=new_r))

            all_changed = fin_changed_fields + conv_changed_fields + endorsed_changed_fields
            new_r['__changed_fields'] = all_changed

            if record_matched:
                matched_count += 1
            if record_changed:
                updated_count += 1
                rec_id = new_r.get('__id')
                if rec_id and patch_fields:
                    snapshot_record = dict(new_r)
                    snapshot = build_log_entry(
                        action_type='REFERENCE_UPDATE', work_week=ww,
                        now_str=now_str, area_index=area_index,
                        index_no=index_no, title_no=title_no,
                        field_name='Full Record',
                        old_value='Previous Snapshot',
                        new_value='Updated Snapshot',
                        full_record=snapshot_record
                    )
                    previous_fields = {}
                    for key in patch_fields:
                        canonical = CANONICAL_FIELD_MAP.get(key, key)
                        spec = next((entry for entry in LOG_COMPARE_SPECS if entry[0] == canonical), None)
                        previous_fields[key] = _none_if_blank(_get_field_val(r, spec[1] if spec else [key, canonical]))
                    items_to_save.append({
                        'id': rec_id, 'fields': patch_fields, 'etag': r.get('__etag') or '',
                        'previous_fields': previous_fields,
                        'previous_record': dict(r), 'snapshot': snapshot
                    })
                elif patch_fields:
                    missing_record_ids.append(area_index or title_no or 'unknown account')
                if row_logs:
                    all_diff_logs.extend(row_logs)

            updated_records.append(new_r)

        # Direct Persistence to Dataverse if requested
        saved_to_dv = False
        save_errors = [
            f"Missing Dataverse record ID for {account}"
            for account in missing_record_ids
        ] if persist_to_dataverse else []
        saved_items = []
        saved_audit_snapshots = []
        if persist_to_dataverse:
            if self.dv_client:
                saved_items, snapshots, errors = self._persist_account_updates(items_to_save, now_str)
                saved_audit_snapshots.extend(snapshots)
                save_errors.extend(errors)
                saved_to_dv = True
                # Saved accounts carry their new version for the next conditional save.
                for item in saved_items:
                    for row in updated_records:
                        if str(row.get('__id')) == str(item['id']):
                            if item.get('new_etag'):
                                row['__etag'] = item['new_etag']
                            else:
                                row.pop('__etag', None)
                            break
            else:
                save_errors.append('Dataverse is not connected.')

        if records is None and not persist_to_dataverse:
            self.loa_records = updated_records
        elif persist_to_dataverse and saved_items:
            saved_ids = {str(item['id']) for item in saved_items}
            updated_by_id = {str(row.get('__id')): row for row in updated_records if row.get('__id')}
            self.loa_records = [
                updated_by_id.get(str(row.get('__id')), row)
                if str(row.get('__id')) in saved_ids else row
                for row in self.loa_records
            ]

        return _json_safe({
            'success': len(save_errors) == 0 if persist_to_dataverse else True,
            'total_records': len(target_records),
            'matched_count': matched_count,
            'updated_count': updated_count,
            'saved_count': len(saved_items) if persist_to_dataverse else 0,
            'logged_count': len(saved_items) if persist_to_dataverse else 0,
            'saved_to_dataverse': saved_to_dv,
            'errors': save_errors,
            'error': '; '.join(save_errors[:3]) if save_errors else '',
            'records': updated_records,
            'diff_summary': all_diff_logs,
            'ambiguous_accounts': ambiguous_accounts,
            'warning': (
                f"{len(ambiguous_accounts)} account(s) skipped because the reference data has conflicting duplicate rows: "
                + '; '.join(ambiguous_accounts[:5]) + (' …' if len(ambiguous_accounts) > 5 else '')
            ) if ambiguous_accounts else '',
            'message': f"Updated {len(saved_items)} record(s) in Land Master DB and recorded audit snapshots in Dataverse." if persist_to_dataverse else f"Matched {matched_count} record(s) with {updated_count} change(s)."
        })

    def populate_loa_records_from_finance(self, records: Optional[List[Dict[str, Any]]] = None) -> dict:
        """Alias for populate_loa_records_from_references for backward compatibility."""
        return self.populate_loa_records_from_references(records)

    # -- LOA Records (cr63f_loamonitoring) -------------------------------------
    def fetch_loa_records(self) -> dict:
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse client not configured.'}
        try:
            table_name = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
            records = self.dv_client.fetch_records(table_name, custom_only=True)

            # Ensure % Paid and Work Week are properly formatted / derived
            for r in records:
                for pct_key in ('% Paid', 'cr63f_paid', '% PAID', 'Percent Paid'):
                    if pct_key in r and r[pct_key] != '':
                        r[pct_key] = format_percent(r[pct_key])
                display_pct = format_calculated_percent(
                    _get_field_val(r, ['cr63f_totalcontractprice', 'Total Contract Price']),
                    _get_field_val(r, ['cr63f_totalpaymentdeductibletotcp', 'Total Payment Deductible to TCP', 'Payment Deductible']),
                )
                if display_pct:
                    r['__display_percent_paid'] = display_pct
                ed_val = r.get('cr63f_endorseddate') or r.get('Endorsed Date') or r.get('ENDORSED DATE') or ''
                ww_val = get_work_week(ed_val) if ed_val else ''
                r['Work Week'] = ww_val
                r['work_week'] = ww_val

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
                'Total Payment Deductible to TCP', 'Balance', '% Paid',
                'work_week', 'Work Week'
            }

            custom_cols = []
            # Priority order for preferred display
            priority_order = [
                'cr63f_type', 'cr63f_indexno', 'cr63f_areaindex', 'cr63f_titleno',
                'cr63f_control1', 'cr63f_registeredowner', 'cr63f_authorizedrepresentative',
                'cr63f_entity', 'cr63f_province', 'cr63f_municipality', 'cr63f_dbadjustedbarangay',
                'cr63f_lotareaha', 'new_contracttype', 'new_mode', 'cr63f_conversionstatus',
                'cr63f_rccprice', 'cr63f_rccterms', 'cr63f_endorseddate', 'work_week', 'cr63f_secureddate',
                'cr63f_securedendorsedate', 'cr63f_totalcontractprice',
                'cr63f_totalpaymentdeductibletotcp', 'cr63f_balance', 'cr63f_paid'
            ]

            processed_keys = set()
            # Add prioritized columns first
            for key in priority_order:
                if key == 'work_week':
                    custom_cols.append({
                        'key': 'work_week',
                        'label': 'Work Week',
                        'type': 'String',
                        'editable': False,
                        'isMoney': False,
                        'isPercent': False,
                        'width': 110
                    })
                    processed_keys.add('work_week')
                    continue
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
            total_tcp = sum(_record_tcp(r) for r in records)
            total_deductible = sum(_record_deductible(r) for r in records)
            # Sum of each account's own balance (each clamped at 0). Netting the
            # totals let overpaid accounts cancel other accounts' balances.
            total_balance = sum(_record_balance(r) for r in records)

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

    def fetch_entry_logs(self) -> dict:
        """Fetch audit rows directly from Dataverse; do not merge, derive, or persist local copies."""
        if not self.dv_client:
            return {
                'success': False,
                'logs': self.log_records,
                'error': 'Dataverse is not connected. Current session logs were not refreshed.'
            }
        try:
            table_name = self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs')
            raw_logs = self.dv_client.fetch_records(table_name, custom_only=False)
            raw_ids = {
                str(row.get('__id') or row.get('id') or row.get('cr63f_loamonitoringlogid') or '').strip()
                for row in raw_logs
            }
            raw_ids.discard('')
            self.pending_session_logs = [
                row for row in self.pending_session_logs
                if str(row.get('__id') or row.get('id') or row.get('cr63f_loamonitoringlogid') or '').strip() not in raw_ids
            ]
            # Source remains Dataverse-only. Snapshot rows are expanded into
            # field-level audit comparisons in memory for display.
            self.log_records = derive_field_diffs_from_snapshots(
                [dict(row) for row in raw_logs],
                active_session_logs=self.pending_session_logs
            )
            for idx, row in enumerate(self.log_records):
                row.setdefault('__id', str(row.get('id') or row.get('cr63f_loamonitoringlogid') or f'log_session_{idx}'))
            self.log_records.sort(
                key=lambda r: parse_log_datetime(
                    r.get('createdon') or r.get('cr63f_logtime') or r.get('timestamp') or r.get('log_time')
                ) or datetime.min,
                reverse=True,
            )
            return {
                'success': True,
                'logs': self.log_records,
                'count': len(self.log_records),
                'source': f"Dataverse: {table_name}"
            }
        except Exception as e:
            return {
                'success': False,
                'logs': self.log_records,
                'count': len(self.log_records),
                'error': f"Could not refresh Dataverse logs: {e}"
            }

    def delete_log_record(self, record_id: str, revert_master: bool = False) -> dict:
        """Revert an audited Land Master field and preserve its history."""
        if not revert_master:
            return {'success': False, 'error': 'Audit logs cannot be deleted. Use Revert for an eligible field.'}
        target_str = str(record_id).strip()
        target_log = None
        target_idx = -1

        for idx, r in enumerate(self.log_records):
            r_id = str(r.get('__id') or r.get('id') or r.get('cr63f_loamonitoringlogid') or '').strip()
            if r_id and r_id == target_str:
                target_log = r
                target_idx = idx
                break

        if target_log is None and target_str.isdigit():
            idx = int(target_str)
            if 0 <= idx < len(self.log_records):
                target_log = self.log_records[idx]
                target_idx = idx

        if target_log is None and target_str.startswith('log_idx_'):
            try:
                idx = int(target_str.replace('log_idx_', ''))
                if 0 <= idx < len(self.log_records):
                    target_log = self.log_records[idx]
                    target_idx = idx
            except Exception:
                pass

        if revert_master and target_log is None:
            return {'success': False, 'error': 'The selected audit record is not available in this session. Refresh logs and try again.'}


        if revert_master and target_log:
            if _audit_source_action(target_log) in REFERENCE_ONLY_ACTIONS:
                return {'success': False, 'error': 'Reference-table audit logs cannot be reverted. Revert is available only for Land Master changes.'}
            if _audit_source_action(target_log) == 'DELETE':
                return {'success': False, 'error': 'A deleted account cannot be restored with Revert. Create it again with New Entry; its last values are in this log row.'}
            master_id = _audit_master_id(target_log)
            ctrl_no = str(target_log.get('cr63f_control1') or target_log.get('cr63f_control') or target_log.get('Control') or target_log.get('Control Number') or target_log.get('Control No') or '').strip()
            area_idx = str(target_log.get('area_index') or target_log.get('cr63f_areaindex') or target_log.get('Area-Index') or target_log.get('Area Index') or target_log.get('AREA-INDEX') or target_log.get('AREA INDEX') or '').strip()
            title_no = str(target_log.get('title_no') or target_log.get('cr63f_titleno') or target_log.get('Title No') or target_log.get('TITLE NO') or target_log.get('Title No.') or target_log.get('TITLE NO.') or target_log.get('TITLE NUMBER') or '').strip()
            index_no = str(target_log.get('index_no') or target_log.get('cr63f_indexno') or target_log.get('Index No') or target_log.get('INDEX NO') or target_log.get('Corridor Index') or '').strip()
            field_name = str(target_log.get('field_name') or target_log.get('cr63f_fieldchanged') or target_log.get('Field Changed') or '').strip()
            old_val = str(target_log.get('old_value') if target_log.get('old_value') is not None else (target_log.get('cr63f_previousvalue') if target_log.get('cr63f_previousvalue') is not None else '')).strip()

            if field_name.lower() in {'balance', '% paid', 'paid', 'percent paid'}:
                return {
                    'success': False,
                    'error': f'{field_name} is calculated from Total Contract Price and Total Payment Deductible to TCP. Revert the deductible or contract price change instead.'
                }

            if not ctrl_no and not area_idx and not title_no:
                log_name = str(target_log.get('cr63f_name') or target_log.get('name') or '').strip()
                if ' - ' in log_name:
                    area_idx = log_name.split(' - ')[0].strip()

            if field_name and field_name != 'Initial Entry' and (master_id or ctrl_no or area_idx or title_no or index_no):
                # Ensure self.loa_records is loaded
                if not self.loa_records and self.dv_client:
                    try:
                        self.fetch_loa_records()
                    except Exception:
                        pass

                # Reverts must never guess between matching Land Master records.
                strict_matches = [mr for mr in self.loa_records if (
                    (master_id and str(mr.get('__id') or '') == master_id) or
                    (not master_id and ctrl_no and str(mr.get('cr63f_control1') or mr.get('Control') or mr.get('Control Number') or '').strip().upper() == ctrl_no.upper()) or
                    (not master_id and not ctrl_no and area_idx and str(mr.get('cr63f_areaindex') or mr.get('Area-Index') or mr.get('Area Index') or '').strip().upper() == area_idx.upper())
                )]
                if master_id and not strict_matches:
                    strict_matches = [{'__id': master_id}]
                if len(strict_matches) != 1:
                    return {'success': False, 'error': 'Cannot safely identify exactly one Land Master record for this audit log.'}
                # Work on a copy: the loaded record changes only after Dataverse
                # accepts the revert and its REVERT audit row.
                live_master = strict_matches[0]
                if not (self.dv_client and live_master.get('__id')):
                    return {'success': False, 'error': 'Dataverse is not connected or this record has no Dataverse ID; nothing was reverted.'}
                table = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
                try:
                    found_master = self.dv_client.get_record(table, live_master['__id'])
                    if not isinstance(found_master, dict) or not found_master.get('__etag'):
                        raise RuntimeError('The current Dataverse record version is unavailable; refresh and try again.')
                except Exception as exc:
                    return {'success': False, 'error': f'Cannot read the current Land Master record: {exc}'}
                original_master = dict(found_master)
                if found_master:
                    # Sanitize old_val
                    if old_val in ('-', 'None', 'null', 'Previous', 'Created', ''):
                        sanitized_val = ''
                    else:
                        f_low = field_name.lower()
                        if any(k in f_low for k in ('price', 'deductible', 'balance', 'tcp', 'amount')):
                            sanitized_val = old_val.replace('₱', '').replace('$', '').replace(',', '').strip()
                        elif 'paid' in f_low or '%' in field_name:
                            sanitized_val = old_val.replace('%', '').replace(',', '').strip()
                        else:
                            sanitized_val = old_val

                    attr_name = FIELD_NAME_TO_DATAVERSE_ATTR.get(field_name, field_name)
                    canonical_disp = DATAVERSE_ATTR_TO_DISPLAY.get(attr_name, field_name)
                    expected_new = next((target_log[key] for key in ('new_value', 'cr63f_newvalue', 'New Value')
                                         if target_log.get(key) is not None), '')
                    current_value = found_master.get(attr_name, found_master.get(canonical_disp, found_master.get(field_name, '')))
                    if not _audit_values_equal(canonical_disp, current_value, expected_new):
                        return {'success': False, 'error': 'This field has changed since this audit log. Revert the latest change first.'}

                    found_master[attr_name] = sanitized_val
                    found_master[field_name] = sanitized_val
                    found_master[canonical_disp] = sanitized_val

                    # Field-specific alias propagation
                    f_low = field_name.lower()
                    if 'total contract price' in f_low or f_low == 'tcp':
                        num_f = parse_float(sanitized_val)
                        disp_p = format_currency(num_f) if num_f is not None else sanitized_val
                        found_master['cr63f_totalcontractprice'] = sanitized_val
                        found_master['Total Contract Price'] = disp_p
                        found_master['TOTAL CONTRACT PRICE'] = disp_p
                    elif 'deductible' in f_low:
                        num_f = parse_float(sanitized_val)
                        disp_p = format_currency(num_f) if num_f is not None else sanitized_val
                        found_master['cr63f_totalpaymentdeductibletotcp'] = sanitized_val
                        found_master['Payment Deductible'] = disp_p
                        found_master['Total Payment Deductible to TCP'] = disp_p
                        found_master['TOTAL PAYMENT DEDUCTIBLE TO TCP'] = disp_p
                    elif 'dar' in f_low or 'conversion' in f_low or f_low == 'status':
                        found_master['cr63f_conversionstatus'] = sanitized_val
                        found_master['Conversion Status'] = sanitized_val
                        found_master['DAR STATUS'] = sanitized_val
                        found_master['DAR Status'] = sanitized_val
                        found_master['Status'] = sanitized_val
                    elif f_low == 'mode':
                        found_master['new_mode'] = sanitized_val
                        found_master['Mode'] = sanitized_val
                        found_master['MODE'] = sanitized_val
                        found_master['cr63f_mode'] = sanitized_val
                    elif 'contract type' in f_low:
                        found_master['new_contracttype'] = sanitized_val
                        found_master['Contract Type'] = sanitized_val
                        found_master['cr63f_contracttype'] = sanitized_val
                    elif 'lot area' in f_low:
                        found_master['cr63f_lotareaha'] = sanitized_val
                        found_master['Lot Area (Ha)'] = sanitized_val
                        found_master['Lot Area (ha)'] = sanitized_val
                    elif 'endorsed' in f_low or attr_name == 'cr63f_endorseddate':
                        ww = get_work_week(sanitized_val) if sanitized_val else ''
                        found_master['cr63f_endorseddate'] = sanitized_val
                        found_master['Endorsed Date'] = sanitized_val
                        found_master['work_week'] = ww
                        found_master['Work Week'] = ww

                    # Balance and % Paid are individually auditable fields. Do
                    # not overwrite their requested old value by immediately
                    # recalculating it. Recalculate dependants only when an
                    # authoritative input (TCP or Deductible) is reverted.
                    should_recalculate = (
                        'total contract price' in f_low or f_low == 'tcp' or
                        'deductible' in f_low
                    )
                    if should_recalculate:
                        tcp_f = _decimal_amount(found_master.get('cr63f_totalcontractprice') or found_master.get('Total Contract Price') or 0)
                        ded_f = _decimal_amount(found_master.get('cr63f_totalpaymentdeductibletotcp') or found_master.get('Payment Deductible') or found_master.get('Total Payment Deductible to TCP') or 0)
                        if tcp_f is not None and ded_f is not None:
                            calc_bal = max(Decimal(0), tcp_f - ded_f)
                            calc_pct = paid_percent_for_storage(tcp_f, ded_f) or Decimal(0)
                            found_master['cr63f_balance'] = calc_bal
                            found_master['Balance'] = format_currency(float(calc_bal))
                            found_master['BALANCE'] = format_currency(float(calc_bal))
                            found_master['cr63f_paid'] = calc_pct
                            percent_display = format_calculated_percent(tcp_f, ded_f)
                            found_master['% Paid'] = percent_display
                            found_master['% PAID'] = percent_display
                            found_master['Paid'] = percent_display

                    # Save to Dataverse first. Balance and % Paid are included only
                    # when they were recalculated from a TCP/Deductible revert.
                    patch_payload = {attr_name: sanitized_val if sanitized_val != '' else None}
                    if should_recalculate:
                        for dependant in ('cr63f_balance', 'cr63f_paid'):
                            if dependant in found_master:
                                patch_payload[dependant] = found_master[dependant]
                    try:
                        self._validate_control_change(original_master, patch_payload, found_master['__id'])
                        revert_res = self.dv_client.save_loa_record(patch_payload, record_id=found_master['__id'],
                                                                    expected_etag=original_master['__etag'])
                    except Exception as pex:
                        return {'success': False, 'error': f"Failed to revert record in Dataverse ({attr_name}): {pex}"}

                    # Revert is an auditable source operation: retain the original
                    # log, write a new REVERT row, and roll every patched field
                    # back if that audit row cannot be created.
                    revert_log = build_log_entry(
                        action_type='REVERT',
                        work_week=get_work_week(datetime.now()),
                        now_str=datetime.now().strftime('%m/%d/%Y %I:%M %p'),
                        area_index=area_idx,
                        index_no=index_no,
                        title_no=title_no,
                        field_name=field_name,
                        old_value=target_log.get('new_value') or target_log.get('cr63f_newvalue') or '',
                        new_value=old_val,
                        details=f"Reverted {field_name} from audit log {target_str}",
                        full_record=found_master
                    )
                    try:
                        saved_log = self.dv_client.save_log_record(revert_log)
                        revert_log['__id'] = saved_log.get('id', '')
                    except Exception as audit_error:
                        rollback_payload = {}
                        for key in patch_payload:
                            previous = original_master.get(key)
                            if key == attr_name and (previous is None or previous == ''):
                                previous = original_master.get(canonical_disp, original_master.get(field_name))
                            rollback_payload[key] = previous if previous not in ('',) else None
                        try:
                            current = self.dv_client.get_record(table, found_master['__id'])
                            if not current.get('__etag') or any(
                                not _audit_values_equal(DATAVERSE_ATTR_TO_DISPLAY.get(key, key), current.get(key), value)
                                for key, value in patch_payload.items()
                            ):
                                raise RuntimeError('The record changed again; automatic rollback would overwrite newer values.')
                            self.dv_client.save_loa_record(rollback_payload, record_id=found_master['__id'],
                                                           expected_etag=current['__etag'])
                        except Exception as rollback_error:
                            return {'success': False, 'error': f'Audit log failed ({audit_error}); rollback also failed ({rollback_error}). Check Dataverse.'}
                        return {'success': False, 'error': f'Revert rolled back because its audit log could not be written: {audit_error}'}

                    # Both writes succeeded: now update the loaded record.
                    _set_version(found_master, (revert_res or {}).get('etag') or '')
                    live_master.clear()
                    live_master.update(found_master)
                    if not any(row is live_master for row in self.loa_records):
                        self.loa_records.append(live_master)
                    self.pending_session_logs.append(dict(revert_log))
                    self._add_log_records([revert_log])
                    return _json_safe({'success': True, 'reverted': True, 'updated_record': dict(found_master), 'message': 'Land Master record reverted and a REVERT audit log was written.'})

        return {'success': False, 'error': 'This audit row has no eligible Land Master field to revert.'}

    def revert_log_record(self, record_id: str) -> dict:
        """Public API for the Entry Logs Revert action."""
        return self.delete_log_record(record_id, revert_master=True)

    def _find_or_fetch_master_record(
        self,
        area_index: Optional[str] = None,
        title_no: Optional[str] = None,
        index_no: Optional[str] = None,
        record_id: Optional[str] = None,
        control_no: Optional[str] = None
    ) -> Dict[str, Any]:
        """Find one Land Master record, strongest key first, never guessing.

        Order: record id, Control Number, AREA-INDEX, Title No, Index No. Each
        key is checked across all records before moving to the next, and a
        non-id key is only accepted when exactly one record carries it. (The
        old loop returned the first row matching *any* key, so an earlier
        record sharing an AREA-INDEX beat the exact record id.)
        """
        norm_id = str(record_id or '').strip()
        records = [r for r in (self.loa_records or []) if isinstance(r, dict)]

        if norm_id:
            for r in records:
                if str(r.get('__id') or '').strip() == norm_id:
                    return dict(r)

        key_specs = (
            (control_no, ('cr63f_control1', 'Control', 'Control Number', 'cr63f_control')),
            (area_index, ('cr63f_areaindex', 'Area-Index', 'Area Index', 'AREA-INDEX')),
            (title_no, ('cr63f_titleno', 'Title No', 'Title Number', 'TITLE NO')),
            (index_no, ('cr63f_indexno', 'Index No', 'INDEX NO')),
        )
        for wanted, fields in key_specs:
            wanted_norm = str(wanted or '').strip().upper()
            if not wanted_norm:
                continue
            matches = [
                r for r in records
                if any(str(r.get(f) or '').strip().upper() == wanted_norm for f in fields)
            ]
            if len(matches) == 1:
                return dict(matches[0])
            if len(matches) > 1:
                # Ambiguous on this key: a weaker key cannot be more reliable.
                return {}

        # Query Dataverse when the record is not loaded in this session.
        if self.dv_client:
            try:
                table_name = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
                filters = []
                if str(area_index or '').strip():
                    filters.append(f"cr63f_areaindex eq {_odata_literal(str(area_index).strip())}")
                if str(title_no or '').strip():
                    filters.append(f"cr63f_titleno eq {_odata_literal(str(title_no).strip())}")
                if filters:
                    dv_records = self.dv_client.fetch_records(
                        table_name, max_records=2, custom_only=True, filter_query=' or '.join(filters)
                    )
                    if dv_records and len(dv_records) == 1:
                        return dict(dv_records[0])
            except Exception:
                pass

        return {}

    def _existing_masters(self, logical: str, fields: tuple, value: str) -> List[Dict[str, Any]]:
        """Land Master records whose ``fields`` equal ``value`` (case-insensitive).

        Uses the loaded records; when none are loaded, asks Dataverse.
        """
        wanted = value.strip().upper()
        if self.loa_records:
            return [r for r in self.loa_records if any(str(r.get(f) or '').strip().upper() == wanted for f in fields)]
        if self.dv_client:
            try:
                table_name = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
                return list(self.dv_client.fetch_records(
                    table_name, max_records=5, custom_only=True,
                    filter_query=f"{logical} eq {_odata_literal(value.strip())}"
                ) or [])
            except Exception:
                return []
        return []

    def _validate_control_change(self, original: dict, fields: dict, record_id: str,
                                 reserved: Optional[dict] = None) -> None:
        keys = ('cr63f_control1', 'cr63f_control', 'Control Number', 'Control', 'Control No.', 'Control No')
        incoming = [str(fields[key] or '').strip() for key in keys if key in fields]
        if not incoming:
            return
        if len({value.upper() for value in incoming}) > 1:
            raise ValueError('Conflicting Control Number values were supplied.')
        control = incoming[0]
        old_control = str(_get_field_val(original, list(keys)) or '').strip()
        if control.upper() == old_control.upper():
            return
        if not control:
            raise ValueError('Control Number cannot be cleared on an existing account.')
        if not original.get('__id') or str(original['__id']) != str(record_id):
            raise ValueError('Cannot safely identify the account whose Control Number is being changed.')
        table = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
        server_rows = self.dv_client.fetch_records(
            table, custom_only=False, filter_query=f'cr63f_control1 eq {_odata_literal(control)}')
        for row in list(self.loa_records) + list(server_rows):
            if str(row.get('__id') or '') == str(record_id):
                continue
            if str(_get_field_val(row, list(keys)) or '').strip().upper() == control.upper():
                raise ValueError(f'Control Number {control} already belongs to another account.')
        if reserved is not None and control.upper() in reserved and reserved[control.upper()] != record_id:
            raise ValueError(f'Control Number {control} is also assigned to another staged account.')
        # A rename must retain its immutable ID in a writable audit column.
        logs_table = self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs')
        attrs = self.dv_client.get_entity_attributes(logs_table)
        identity_attrs = ('cr63f_details', 'cr63f_name')
        if not any(attrs.get(key, {}).get('can_create') for key in identity_attrs):
            raise ValueError('Control Number changes require a writable Details or Name audit column to retain the master record ID.')
        if reserved is not None:
            reserved[control.upper()] = record_id
        fields['cr63f_control1'] = control

    def _check_new_entry_keys(self, record_data: Dict[str, Any], allow_shared_area_index: bool) -> Optional[Dict[str, Any]]:
        """Duplicate rules for a new Land Master account (decision 2026-10-07).

        - Control Number is required and must be unique: it is the account's
          key for Entry Logs history and Revert.
        - AREA-INDEX may be shared (e.g. several parcels in one area), but only
          after the user confirms: the first save returns needs_confirmation
          naming the existing account(s); the UI resends with
          allow_shared_area_index=True.
        Returns an error/confirmation response, or None when the entry may be saved.
        """
        control = str(record_data.get('cr63f_control1') or record_data.get('Control') or record_data.get('Control Number') or '').strip()
        if not control:
            return {'success': False, 'error': 'Control Number is required for a new entry.'}
        if self._existing_masters('cr63f_control1', ('cr63f_control1', 'Control', 'Control Number'), control):
            return {
                'success': False,
                'error': f"Control Number {control} already exists in Land Master DB. Edit that account instead of creating a new one.",
            }

        area_index = str(record_data.get('cr63f_areaindex') or record_data.get('Area-Index') or record_data.get('AREA-INDEX') or '').strip()
        if area_index and not allow_shared_area_index:
            existing = self._existing_masters('cr63f_areaindex', ('cr63f_areaindex', 'Area-Index', 'Area Index', 'AREA-INDEX'), area_index)
            if existing:
                controls = [
                    str(r.get('cr63f_control1') or r.get('Control') or r.get('Control Number') or '').strip() or '(no Control Number)'
                    for r in existing
                ]
                return {
                    'success': False,
                    'needs_confirmation': True,
                    'existing_controls': controls,
                    'message': (
                        f"AREA-INDEX {area_index} is already used by {', '.join(controls)}. "
                        f"Save {control} as another account in the same area?"
                    ),
                }
        return None

    def save_loa_record_entry(self, record_data: dict, record_id: Optional[str] = None,
                              allow_shared_area_index: bool = False) -> dict:
        """Save new log entry to cr63f_loamonitoringlogs and upsert master record in cr63f_loamonitoring."""
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        if not record_id and not str(record_data.get('cr63f_acquisitionsublevel') or record_data.get('Acquisition Sub-Level') or '').strip():
            return {'success': False, 'error': 'Acquisition Sub-Level is required for a new entry.'}
        try:
            master_table = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
            logs_table = self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs')

            now_str = datetime.now().strftime('%m/%d/%Y %I:%M %p')
            ww = get_work_week(datetime.now())

            if record_id:
                # Update existing record in master table (cr63f_loamonitoring)
                found_rec = self._find_or_fetch_master_record(
                    record_id=record_id,
                    area_index=record_data.get('cr63f_areaindex') or record_data.get('Area-Index'),
                    title_no=record_data.get('cr63f_titleno') or record_data.get('Title No')
                )
                self._validate_control_change(found_rec, record_data, record_id)
                # A blank Type or Mode sent for an existing account would erase a
                # stored value (for example a Mode missing from the drop-down list).
                # Refuse that instead of saving it.
                blank_error = _required_field_cleared(found_rec, record_data)
                if blank_error:
                    return {'success': False, 'error': blank_error}

                merged_rec = dict(found_rec)
                merged_rec.update(record_data)
                # Balance and % Paid are always derived here, never taken from
                # the client, so a manual save cannot store a stale value.
                derived = _derived_payment_fields(merged_rec)
                if derived:
                    record_data = {**record_data, **derived}
                    merged_rec.update(derived)

                # The hidden pre-change snapshot lets this first manual update
                # be compared after the app restarts.
                baseline = self._write_baseline_if_missing(
                    found_rec, now_str, require_master_identity=(
                        'cr63f_control1' in record_data and
                        str(record_data['cr63f_control1']).strip().upper() !=
                        str(_get_field_val(found_rec, ['cr63f_control1', 'Control', 'Control Number']) or '').strip().upper()))

                # Conditional on the version this account was loaded at: if someone
                # else saved it since, Dataverse refuses (412) instead of this save
                # overwriting their change, and the snapshot below stays true.
                res = self._save_master(record_data, record_id, found_rec.get('__etag'))
                new_etag = (res or {}).get('etag') or ''

                # Persist one post-change snapshot; field-level rows are derived
                # from the previous Dataverse snapshot when logs are loaded.
                area_idx = record_data.get('cr63f_areaindex') or record_data.get('Area-Index') or merged_rec.get('cr63f_areaindex') or ''
                index_no = record_data.get('cr63f_indexno') or record_data.get('Index No') or merged_rec.get('cr63f_indexno') or ''
                title_no = record_data.get('cr63f_titleno') or record_data.get('Title No') or merged_rec.get('cr63f_titleno') or ''
                log_entry = build_log_entry(
                    action_type='MANUAL_EDIT',
                    work_week=ww,
                    now_str=now_str,
                    area_index=area_idx,
                    index_no=index_no,
                    title_no=title_no,
                    field_name='Full Record',
                    old_value='Previous Snapshot',
                    new_value='Updated Snapshot',
                    full_record=merged_rec
                )
                try:
                    saved_log = self.dv_client.save_log_record(log_entry)
                    if saved_log and saved_log.get('id'):
                        log_entry['__id'] = saved_log['id']
                except Exception as audit_error:
                    try:
                        rollback_fields = {
                            key: _none_if_blank(_get_field_val(found_rec, [key, CANONICAL_FIELD_MAP.get(key, key)]))
                            for key in record_data
                            if not key.startswith('__')
                        }
                        rollback_res = self._save_master(rollback_fields, record_id, new_etag)
                        for row in self.loa_records:
                            if row.get('__id') == record_id:
                                _set_version(row, (rollback_res or {}).get('etag') or '')
                                break
                    except Exception as rollback_error:
                        return {'success': False, 'error': f'Audit save failed ({audit_error}); Land Master rollback failed ({rollback_error}). Check Dataverse.'}
                    return {'success': False, 'error': f'Land Master update was rolled back because its audit snapshot could not be saved: {audit_error}'}

                if baseline:
                    self.pending_session_logs.append(dict(baseline))
                self.pending_session_logs.append(dict(log_entry))
                self.log_records = derive_field_diffs_from_snapshots(
                    ([baseline] if baseline else []) + [log_entry]
                ) + self.log_records
                for idx, r in enumerate(self.loa_records):
                    if r.get('__id') == record_id:
                        self.loa_records[idx].update(record_data)
                        _set_version(self.loa_records[idx], new_etag)
                        _refresh_display_percent(self.loa_records[idx])
                        endorsed_ww = _endorsed_work_week(self.loa_records[idx])
                        self.loa_records[idx]['work_week'] = endorsed_ww
                        self.loa_records[idx]['Work Week'] = endorsed_ww
                        break

                msg = f"Record successfully updated in Dataverse master table ({master_table})."
            else:
                # 0. Refuse duplicates: AREA-INDEX and Control Number identify
                # an account everywhere else (lookups, audit history, reverts).
                key_problem = self._check_new_entry_keys(record_data, allow_shared_area_index)
                if key_problem:
                    return key_problem

                # Endorsed Date defaults to today's date (the form pre-fills it;
                # this covers a cleared field or another caller).
                if not _endorsed_date_value(record_data):
                    record_data = dict(record_data)
                    today = datetime.now().strftime('%m/%d/%Y')
                    record_data['cr63f_endorseddate'] = today
                    record_data['Endorsed Date'] = today

                # Balance and % Paid come from TCP and Deductible, never from the client.
                derived = _derived_payment_fields(record_data)
                if derived:
                    record_data = {**record_data, **derived}

                # 1. Save master record in cr63f_loamonitoring
                res_master = self.dv_client.save_loa_record(record_data)
                created_id = res_master.get('id', '')
                new_rec = dict(record_data)
                new_rec['__id'] = created_id
                _refresh_display_percent(new_rec)
                endorsed_ww = _endorsed_work_week(new_rec)
                new_rec['work_week'] = endorsed_ww
                new_rec['Work Week'] = endorsed_ww
                self.loa_records.insert(0, new_rec)

                # 2. Save log record in cr63f_loamonitoringlogs
                area_idx = record_data.get('cr63f_areaindex') or record_data.get('Area-Index') or ''
                index_no = record_data.get('cr63f_indexno') or record_data.get('Index No') or ''
                title_no = record_data.get('cr63f_titleno') or record_data.get('Title No') or ''
                log_entry = build_log_entry(
                    action_type='NEW_ENTRY',
                    work_week=ww,
                    now_str=now_str,
                    area_index=area_idx,
                    index_no=index_no,
                    title_no=title_no,
                    field_name='Full Record',
                    old_value='-',
                    new_value='Created',
                    details=f"Created new entry {area_idx}",
                    full_record=new_rec
                )

                try:
                    res_log = self.dv_client.save_log_record(log_entry)
                    log_entry['__id'] = res_log.get('id', '')
                except Exception as audit_error:
                    # Same rule as updates: no Land Master change without its
                    # audit row. Remove the account that was just created.
                    try:
                        if created_id:
                            self.dv_client.delete_loa_record(created_id)
                        self.loa_records = [r for r in self.loa_records if r.get('__id') != created_id]
                    except Exception as rollback_error:
                        return {
                            'success': False,
                            'id': created_id,
                            'error': f'Initial Entry audit row could not be saved ({audit_error}) and the new account could not be removed ({rollback_error}). Check Dataverse.',
                        }
                    return {
                        'success': False,
                        'error': f'New entry was not saved because its Initial Entry audit row could not be written: {audit_error}',
                    }
                self.pending_session_logs.append(dict(log_entry))
                self.log_records = derive_field_diffs_from_snapshots([log_entry]) + self.log_records

                msg = f"New entry saved to Master DB ({master_table}) and logged ({logs_table})."
                res = res_master

            return {
                'success': True,
                'action': res.get('action'),
                'id': res.get('id'),
                'target_table': master_table,
                'message': msg
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def save_staged_records(self, staged_list: List[Dict[str, Any]]) -> dict:
        """Save staged row updates into Dataverse master table (cr63f_loamonitoring) and write change logs in parallel."""
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        master_table = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
        logs_table = self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs')

        total_items = len(staged_list)
        if total_items == 0:
            return {'success': True, 'saved_count': 0, 'message': 'No changes to save.'}

        lock = threading.Lock()
        saved_count = 0
        errors = []
        now_str = datetime.now().strftime('%m/%d/%Y %I:%M %p')
        log_entries_to_save = []
        baseline_snapshots = []
        audit_contexts = {}
        reserved_controls = {}

        def save_single_record(item):
            nonlocal saved_count
            rec_id = item.get('id')
            fields = item.get('fields', {})
            if not rec_id or not fields:
                return

            with lock:
                found_rec = next((r for r in self.loa_records if r.get('__id') == rec_id), {})
                orig_rec = dict(found_rec)
            if not orig_rec:
                orig_rec = self._find_or_fetch_master_record(record_id=rec_id, area_index=fields.get('cr63f_areaindex') or fields.get('Area-Index'), title_no=fields.get('cr63f_titleno') or fields.get('Title No'))
            try:
                with lock:
                    self._validate_control_change(orig_rec, fields, rec_id, reserved_controls)
            except Exception as exc:
                with lock:
                    errors.append(f'Row {rec_id}: {exc}')
                return
            blank_error = _required_field_cleared(orig_rec, fields)
            if blank_error:
                with lock:
                    errors.append(f'Row {rec_id}: {blank_error}')
                return
            merged_rec = dict(orig_rec)
            merged_rec.update(fields)
            # Any staged change to a payment field (TCP, Deductible, or a typed
            # Balance / % Paid) recalculates all four on the server, so the
            # stored values always follow the saved amounts.
            if any(key in fields for key in _TCP_KEYS + _DEDUCTIBLE_KEYS + ('cr63f_balance', 'Balance', 'cr63f_paid', '% Paid')):
                derived = _derived_payment_fields(merged_rec)
                if derived:
                    fields.update(derived)
                    merged_rec.update(derived)

            area_idx = orig_rec.get('cr63f_areaindex') or orig_rec.get('Area-Index') or fields.get('cr63f_areaindex') or fields.get('Area-Index') or rec_id
            index_no = orig_rec.get('cr63f_indexno') or orig_rec.get('Index No') or fields.get('cr63f_indexno') or fields.get('Index No') or ''
            title_no = orig_rec.get('cr63f_titleno') or orig_rec.get('Title No') or fields.get('cr63f_titleno') or fields.get('Title No') or ''
            # Audit work week is based on the log/save time, not Endorsed Date.
            ww = get_work_week(datetime.now())

            # Work out what really changes before anything is written. A staged
            # value equal to the stored one is a no-op: nothing is saved and no
            # audit row is written, so no Land Master change lacks its audit row.
            changed_fields = []
            seen_canon = set()
            for f_key, new_v in fields.items():
                if f_key.startswith('__') or f_key in ('work_week', 'Work Week'):
                    continue
                canon_name = CANONICAL_FIELD_MAP.get(f_key, f_key)
                if canon_name in CALCULATED_AUDIT_FIELDS:
                    continue
                if canon_name in seen_canon:
                    continue

                spec = next((entry for entry in LOG_COMPARE_SPECS if entry[0] == canon_name), None)
                keys = spec[1] if spec else [f_key, canon_name]
                old_v = _get_field_val(orig_rec, keys)
                updated_v = next(
                    (fields[key] for key in keys if key in fields),
                    _get_field_val(merged_rec, keys)
                )
                if not _audit_values_equal(canon_name, old_v, updated_v):
                    seen_canon.add(canon_name)
                    changed_fields.append(canon_name)

            if not changed_fields:
                return

            try:
                baseline = self._write_baseline_if_missing(
                    orig_rec, now_str, require_master_identity=(
                        'cr63f_control1' in fields and
                        str(fields['cr63f_control1']).strip().upper() !=
                        str(_get_field_val(orig_rec, ['cr63f_control1', 'Control', 'Control Number']) or '').strip().upper()))
                if baseline:
                    with lock:
                        baseline_snapshots.append(baseline)

                res = self._save_master(fields, rec_id, orig_rec.get('__etag'))
                new_etag = (res or {}).get('etag') or ''
                with lock:
                    saved_count += 1

                item_logs = []
                if changed_fields:
                    act_type = item.get('action_type') or 'MANUAL_EDIT'
                    item_logs.append(build_log_entry(
                        action_type=act_type,
                        work_week=ww,
                        now_str=now_str,
                        area_index=area_idx,
                        index_no=index_no,
                        title_no=title_no,
                        field_name='Full Record',
                        old_value='Previous Snapshot',
                        new_value='Updated Snapshot',
                        full_record=merged_rec
                    ))

                with lock:
                    for r in self.loa_records:
                        if r.get('__id') == rec_id:
                            r.update(fields)
                            _set_version(r, new_etag)
                            _refresh_display_percent(r)
                            # Land Master Work Week always comes from Endorsed
                            # Date; the save week belongs to the audit log only.
                            endorsed_ww = _endorsed_work_week(r)
                            r['work_week'] = endorsed_ww
                            r['Work Week'] = endorsed_ww
                            break
                    if item_logs:
                        log_entries_to_save.extend(item_logs)
                        for log_entry in item_logs:
                            audit_contexts[id(log_entry)] = (rec_id, orig_rec, fields, new_etag)

            except Exception as e:
                with lock:
                    errors.append(f"Row {rec_id}: {e}")

        # Phase 1: Parallel save of LOA master records
        max_workers = min(10, max(1, total_items))
        completed_records = 0
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(save_single_record, itm) for itm in staged_list]
            for future in as_completed(futures):
                with lock:
                    completed_records += 1
                    cur_prog = completed_records
                if self._window:
                    try:
                        self._window.evaluate_js(
                            f"window.onDataverseSaveProgress && window.onDataverseSaveProgress({cur_prog}, {total_items}, 'Saving master records: {cur_prog} of {total_items}...')"
                        )
                    except Exception:
                        pass

        # Phase 2: Parallel save of audit log records
        if log_entries_to_save:
            total_logs = len(log_entries_to_save)
            completed_logs = 0
            log_workers = min(10, max(1, total_logs))

            def save_single_log(log_entry):
                try:
                    saved_log = self.dv_client.save_log_record(log_entry)
                    if saved_log and saved_log.get('id'):
                        log_entry['__id'] = saved_log['id']
                    return log_entry
                except Exception as exc:
                    return exc

            saved_logs = []
            with ThreadPoolExecutor(max_workers=log_workers) as log_executor:
                log_futures = {
                    log_executor.submit(save_single_log, le): le
                    for le in log_entries_to_save
                }
                for lf in as_completed(log_futures):
                    result = lf.result()
                    if isinstance(result, Exception):
                        rec_id, original, changed, saved_etag = audit_contexts[id(log_futures[lf])]
                        rollback_fields = {}
                        for key in changed:
                            canonical = CANONICAL_FIELD_MAP.get(key, key)
                            spec = next((entry for entry in LOG_COMPARE_SPECS if entry[0] == canonical), None)
                            rollback_fields[key] = _none_if_blank(_get_field_val(original, spec[1] if spec else [key, canonical]))
                        try:
                            rollback_res = self._save_master(rollback_fields, rec_id, saved_etag)
                            with lock:
                                saved_count -= 1
                                for row in self.loa_records:
                                    if row.get('__id') == rec_id:
                                        row.clear()
                                        row.update(original)
                                        _set_version(row, (rollback_res or {}).get('etag') or '')
                                        break
                            errors.append(f"Audit save failed for {rec_id} ({result}); Land Master update was rolled back.")
                        except Exception as rollback_error:
                            errors.append(
                                f"Audit save failed for {rec_id} ({result}) and Land Master rollback failed ({rollback_error}). Check Dataverse."
                            )
                    else:
                        saved_logs.append(result)
                    with lock:
                        completed_logs += 1
                        cur_log_prog = completed_logs
                    if self._window:
                        try:
                            self._window.evaluate_js(
                                f"window.onDataverseSaveProgress && window.onDataverseSaveProgress({cur_log_prog}, {total_logs}, 'Writing audit logs: {cur_log_prog} of {total_logs}...')"
                            )
                        except Exception:
                            pass

            with lock:
                newly_saved_rows = baseline_snapshots + saved_logs
                self.pending_session_logs.extend(dict(log) for log in newly_saved_rows)
                self.log_records = derive_field_diffs_from_snapshots(newly_saved_rows) + self.log_records

        if errors:
            return {'success': False, 'error': '; '.join(errors[:3])}

        return {
            'success': True,
            'saved_count': saved_count,
            'message': f"Saved {saved_count} updated record(s) to Master DB ({master_table}) and wrote {len(log_entries_to_save)} audit snapshot(s) to {logs_table}."
        }

    def save_staged_finance_records(self, staged_list: List[Dict[str, Any]]) -> dict:
        return self._save_staged_reference_records(staged_list, 'finance')

    def save_staged_conversion_records(self, staged_list: List[Dict[str, Any]]) -> dict:
        return self._save_staged_reference_records(staged_list, 'conversion')

    def _save_staged_reference_records(self, staged_list: List[Dict[str, Any]], kind: str) -> dict:
        """Count a row as saved only when its reference write and audit writes succeed."""
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        rows = self.finance_rows if kind == 'finance' else self.conversion_rows
        engine = self.finance_engine if kind == 'finance' else self.conversion_engine
        default_table = 'cr63f_financereference' if kind == 'finance' else 'cr63f_landconversionreference'
        table = self.dv_client.config.get(f'{kind}_table', default_table)
        save = self.dv_client.save_finance_record if kind == 'finance' else self.dv_client.save_conversion_record
        saved_ids, failed_ids, errors, persisted_logs = [], [], [], []
        blocked = getattr(self, '_reference_save_blocked', set())
        self._reference_save_blocked = blocked
        now_str = datetime.now().strftime('%m/%d/%Y %I:%M %p')
        baselines = getattr(self, '_reference_edit_baselines', {})
        for position, item in enumerate(staged_list):
            input_id = str(item.get('id', ''))
            fields = {key: value for key, value in item.get('fields', {}).items() if not key.startswith('__')}
            original = {}
            written_id = ''
            item_logs = []
            source_written = False
            try:
                if (kind, input_id) in blocked:
                    raise RuntimeError('A previous save has an unresolved rollback/audit error. Check Dataverse before retrying this row.')
                row = next((record for record in rows if str(record.get('__id') or '') == input_id), None)
                if row is None and input_id.isdigit() and int(input_id) < len(rows):
                    row = rows[int(input_id)]
                if row is None:
                    raise ValueError('The selected reference row is unavailable. Refresh and try again.')
                original = dict(row)
                staged_baseline = baselines.get((kind, input_id))
                if not staged_baseline:
                    # The same row may have been edited under its other key
                    # (Dataverse ID vs position); never lose its "before" values.
                    for other in (str(row.get('__id') or ''), str(rows.index(row))):
                        candidate = baselines.get((kind, other))
                        if other and candidate and candidate[0] is row:
                            staged_baseline = candidate
                            break
                if staged_baseline and staged_baseline[0] is row:
                    original = dict(staged_baseline[1])
                changed = {key: value for key, value in fields.items()
                           if not _audit_values_equal(CANONICAL_FIELD_MAP.get(key, key), original.get(key), value)}
                if not changed:
                    saved_ids.append(input_id)
                    baselines.pop((kind, input_id), None)
                    continue
                merged = {**original, **changed}
                area = str(_get_field_val(merged, ['cr63f_areaindex', 'AREA-INDEX', 'Area Index', 'Corridor Index']) or input_id)
                title = str(_get_field_val(merged, ['cr63f_titlenumber', 'TITLE NUMBER', 'Title Mod', 'Title No']) or '')
                index = str(_get_field_val(merged, ['Corridor Index', 'INDEX NO', 'Index No']) or '')
                master = self._find_or_fetch_master_record(area_index=area, title_no=title, index_no=index)
                snapshot = {**master, **merged}
                index = index or str(master.get('cr63f_indexno') or master.get('Index No') or '')
                title = title or str(master.get('cr63f_titleno') or master.get('Title No') or '')
                written_id = str(original.get('__id') or '')
                result = save(changed, record_id=written_id) if written_id else save(merged)
                if not result or result.get('success') is False:
                    raise RuntimeError('Dataverse did not confirm the reference write.')
                source_written = True
                written_id = written_id or str(result.get('id') or '')
                if not written_id:
                    raise RuntimeError('Dataverse did not return the created reference ID.')
                for key, value in changed.items():
                    log = build_log_entry(
                        action_type=f'{kind.upper()}_UPDATE', work_week=get_work_week(datetime.now()),
                        now_str=now_str, area_index=area, index_no=index, title_no=title,
                        field_name=key, old_value=original.get(key, ''), new_value=value,
                        details=f'Updated {kind.title()} Reference {key}', full_record=snapshot)
                    logged = self.dv_client.save_log_record(log)
                    if not isinstance(logged, dict) or logged.get('success') is False or not logged.get('id'):
                        raise RuntimeError('Dataverse did not confirm the audit row ID.')
                    log['__id'] = logged['id']
                    item_logs.append(log)
                row.update(changed)
                row['__id'] = written_id
                saved_ids.append(input_id)
                baselines.pop((kind, input_id), None)
                persisted_logs.extend(item_logs)
            except Exception as exc:
                problem = str(exc)
                if source_written:
                    try:
                        if not written_id:
                            raise RuntimeError('The created reference ID is unknown.')
                        if original.get('__id'):
                            current = self.dv_client.get_record(table, written_id)
                            if not isinstance(current, dict) or not current.get('__etag'):
                                raise RuntimeError('The current reference version is unavailable.')
                            for key, value in changed.items():
                                current_value = _get_field_val(current, [key, CANONICAL_FIELD_MAP.get(key, key)])
                                if not _audit_values_equal(CANONICAL_FIELD_MAP.get(key, key), current_value, value):
                                    raise RuntimeError('The reference changed again; rollback would overwrite newer values.')
                            rollback = {key: _none_if_blank(original.get(key)) for key in changed}
                            self.dv_client.save_record(table, rollback, record_id=written_id, expected_etag=current['__etag'])
                        else:
                            info = self.dv_client.get_entity_info(table)
                            self.dv_client.delete(info['entity_set'], written_id)
                        for log in item_logs:
                            self.dv_client.delete_log_record(log['__id'])
                        problem += '; reference update rolled back.'
                    except Exception as rollback_error:
                        blocked.add((kind, input_id))
                        problem += f'; rollback/audit cleanup failed ({rollback_error}). Check Dataverse before retrying.'
                failed_ids.append(input_id)
                errors.append(f'{kind.title()} row {input_id}: {problem}')
            finally:
                if self._window:
                    try:
                        self._window.evaluate_js(
                            f"window.onDataverseSaveProgress && window.onDataverseSaveProgress({position + 1}, {len(staged_list)}, 'Saving {kind} records and audit logs...')")
                    except Exception:
                        pass
        if persisted_logs:
            self.pending_session_logs.extend(dict(log) for log in persisted_logs)
            self._add_log_records(persisted_logs)
        engine.load_rows(rows)
        return _json_safe({
            'success': not errors, 'saved_count': len(saved_ids), 'saved_ids': saved_ids,
            'failed_ids': failed_ids, 'errors': errors, 'error': '; '.join(errors[:3]),
            'logged_count': len(persisted_logs),
            'message': f'Saved {len(saved_ids)} {kind} record(s) and {len(persisted_logs)} audit changes. {len(failed_ids)} failed.'})

    def revert_staged_finance_records(self, baseline_rows: List[Dict[str, Any]]) -> dict:
        """Revert in-memory finance records and engine back to baseline."""
        self.finance_rows = list(baseline_rows or [])
        for key in list(getattr(self, '_reference_edit_baselines', {})):
            if key[0] == 'finance':
                self._reference_edit_baselines.pop(key)
        self.finance_engine.load_rows(self.finance_rows)
        return {'success': True}

    def revert_staged_conversion_records(self, baseline_rows: List[Dict[str, Any]]) -> dict:
        """Revert in-memory conversion records and engine back to baseline."""
        self.conversion_rows = list(baseline_rows or [])
        for key in list(getattr(self, '_reference_edit_baselines', {})):
            if key[0] == 'conversion':
                self._reference_edit_baselines.pop(key)
        self.conversion_engine.load_rows(self.conversion_rows)
        return {'success': True}

    def get_dashboard_analytics(self, work_week: Optional[str] = None, mode: Optional[str] = None) -> dict:
        """Compute weekly intelligence metrics, mode distribution, conversion analytics, and change activity."""
        all_records = self.loa_records or []
        logs = self.log_records or []
        conv_rows = self.conversion_rows or []

        selected_mode_input = (mode or '').strip().upper()
        if selected_mode_input and selected_mode_input != 'ALL':
            records = [r for r in all_records if _record_mode(r) == selected_mode_input]
            active_mode = selected_mode_input
        else:
            records = all_records
            active_mode = None

        def log_datetime(log: Dict[str, Any]) -> Optional[datetime]:
            return parse_log_datetime(
                log.get('createdon') or log.get('cr63f_logtime') or log.get('timestamp')
                or log.get('log_time') or log.get('Log Time')
            )

        def log_week_label(log: Dict[str, Any]) -> str:
            dt = log_datetime(log)
            if dt:
                return get_work_week_label(dt)
            ww = str(log.get('work_week') or log.get('Work Week') or '').strip()
            return ww

        def bare_week(label: str) -> str:
            # '2026 · WW-40' -> 'WW-40'; 'WW-40' -> 'WW-40'
            return str(label or '').split('·')[-1].strip().upper()

        week_dates = {}
        for log in logs:
            label = log_week_label(log)
            dt = log_datetime(log)
            if label and dt:
                week_dates[label] = max(week_dates.get(label, dt), dt)
        available_ww = [label for label, _ in sorted(week_dates.items(), key=lambda item: item[1], reverse=True)]
        selected_ww = (work_week or 'ALL').strip()
        if selected_ww.upper() == 'ALL':
            selected_ww = available_ww[0] if available_ww else 'ALL'
        elif selected_ww not in available_ww:
            # A bare 'WW-40' selects the most recent year that has that week.
            # Exact week comparison only: substring matching let 'WW-1'
            # match WW-10..WW-19 and 'WW-01' match every year.
            wanted = bare_week(selected_ww)
            selected_ww = next((w for w in available_ww if bare_week(w) == wanted), selected_ww)
        total_tcp = sum(_record_tcp(r) for r in records)
        total_ded = sum(_record_deductible(r) for r in records)
        total_bal = sum(_record_balance(r) for r in records)
        overall_paid_pct = ((total_ded / total_tcp) * 100) if total_tcp > 0 else (100.0 if total_ded > 0 else 0.0)

        # Mode distribution breakdown (calculated against all_records)
        mode_counts = {}
        for r in all_records:
            m = _record_mode(r)
            mode_counts[m] = mode_counts.get(m, 0) + 1

        modes_distribution = []
        total_recs = len(all_records) or 1
        for m_name, count in sorted(mode_counts.items(), key=lambda x: x[1], reverse=True):
            pct = (count / total_recs) * 100.0
            modes_distribution.append({
                'mode': m_name,
                'count': count,
                'percent': f"{pct:.1f}%",
                'percent_val': round(pct, 1),
                'is_active': (active_mode == m_name)
            })

        # Financial Settlement Breakdown (decision 2026-10-13).
        # Settled / Partially paid / Unpaid apply to DIRECT SALE accounts only:
        # - STP SALE (SS - STP, SS - FROM LAND STP): the buyer owes the balance after
        #   turnover, so it is counted but not tracked as paid/unpaid.
        # - LEASE (SECURED LEASE, SL - ...): TCP is total rent over the lease term, so
        #   % paid reflects time elapsed, not arrears; counted only ("LEASE - RENT BASIS").
        # - RESCINDED and accounts with no sale/lease Mode: counted only.
        rec_count = len(records)
        fully_paid_count = 0
        partially_paid_count = 0
        zero_paid_count = 0
        direct_tcp = 0.0
        direct_ded = 0.0
        stp_sale_count = 0
        lease_direct_count = 0
        lease_direct_tcp = 0.0
        lease_direct_ded = 0.0
        lease_stp_count = 0
        rescinded_count = 0
        other_count = 0

        group_totals: Dict[str, Dict[str, float]] = {}
        for r in records:
            mode = _record_mode(r)
            g_tcp, g_ded = _record_tcp(r), _record_deductible(r)
            gt = group_totals.setdefault(_payment_group(mode), {'count': 0, 'tcp': 0.0, 'paid': 0.0, 'balance': 0.0})
            gt['count'] += 1
            gt['tcp'] += g_tcp
            gt['paid'] += g_ded
            gt['balance'] += max(0.0, g_tcp - g_ded)
            mode_class = finance_mode_class(mode)
            is_stp = bool(re.search(r'\bSTP\b', mode)) or 'SOLD TO PARTNER' in mode
            if 'RESCIND' in mode:
                rescinded_count += 1
                continue
            if mode_class == 'LEASE':
                if is_stp:
                    lease_stp_count += 1
                else:
                    lease_direct_count += 1
                    lease_direct_tcp += _record_tcp(r)
                    lease_direct_ded += _record_deductible(r)
                continue
            if mode_class != 'SALE':
                other_count += 1
                continue
            if is_stp:
                stp_sale_count += 1
                continue
            tcp_v = _record_tcp(r)
            ded_v = _record_deductible(r)
            direct_tcp += tcp_v
            direct_ded += ded_v
            bal_v = max(0.0, tcp_v - ded_v)
            # Same rule as % Paid: TCP 0 with a deductible counts as 100% paid.
            if (tcp_v > 0 and bal_v <= 0.01) or (tcp_v <= 0 and ded_v > 0):
                fully_paid_count += 1
            elif ded_v > 0 and ded_v < tcp_v:
                partially_paid_count += 1
            else:
                zero_paid_count += 1

        direct_count = fully_paid_count + partially_paid_count + zero_paid_count
        direct_paid_pct = ((direct_ded / direct_tcp) * 100) if direct_tcp > 0 else (100.0 if direct_ded > 0 else 0.0)
        avg_tcp = (direct_tcp / direct_count) if direct_count > 0 else 0.0
        avg_ded = (direct_ded / direct_count) if direct_count > 0 else 0.0

        def _pct_of_direct(n: int) -> float:
            return round(n / (direct_count or 1) * 100, 1)

        financial_breakdown = {
            'direct_sale_count': direct_count,
            'direct_sale_tcp': direct_tcp,
            'direct_sale_paid': direct_ded,
            'direct_sale_paid_percent_num': round(direct_paid_pct, 2),
            'fully_paid_count': fully_paid_count,
            'fully_paid_percent': f"{_pct_of_direct(fully_paid_count):.1f}%",
            'fully_paid_percent_num': _pct_of_direct(fully_paid_count),
            'partially_paid_count': partially_paid_count,
            'partially_paid_percent': f"{_pct_of_direct(partially_paid_count):.1f}%",
            'partially_paid_percent_num': _pct_of_direct(partially_paid_count),
            'zero_paid_count': zero_paid_count,
            'zero_paid_percent': f"{_pct_of_direct(zero_paid_count):.1f}%",
            'zero_paid_percent_num': _pct_of_direct(zero_paid_count),
            'stp_sale_count': stp_sale_count,
            'lease_count': lease_direct_count + lease_stp_count,
            'lease_direct_count': lease_direct_count,
            'direct_sale_balance': group_totals.get('DIRECT_SALE', {}).get('balance', 0.0),
            'lease_direct_balance': group_totals.get('LEASE_DIRECT', {}).get('balance', 0.0),
            'group_totals': group_totals,
            'lease_direct_tcp': lease_direct_tcp,
            'lease_direct_paid': lease_direct_ded,
            'lease_direct_paid_percent_num': round((lease_direct_ded / lease_direct_tcp * 100) if lease_direct_tcp > 0 else 0.0, 1),
            'lease_stp_count': lease_stp_count,
            'rescinded_count': rescinded_count,
            'other_count': other_count,
            'avg_tcp_formatted': f"₱{format_currency(avg_tcp)}",
            'avg_paid_formatted': f"₱{format_currency(avg_ded)}",
        }

        # Conversion Reference Analytics (filtered by active records)
        conv_status_counts = {}
        total_applied_area = 0.0
        secured_count = 0
        in_progress_count = 0

        # Land Master only, never the Conversion reference table: an empty
        # Land Master (or a Mode filter with no accounts) shows zero, not reference data.
        for r in records:
            c_status = str(r.get('cr63f_conversionstatus') or r.get('Conversion Status') or '').strip().upper()
            area_val = parse_float(r.get('cr63f_lotareaha') or r.get('Lot Area (Ha)') or r.get('cr63f_areaappliedhas') or 0.0)

            # Counts reflect Land Master only, so they agree with the grid and
            # exports. Reference data reaches Land Master via Update from References.
            if not c_status or c_status in ('UNKNOWN', '-', ''):
                c_status = 'NOT YET ENDORSED'

            conv_status_counts[c_status] = conv_status_counts.get(c_status, 0) + 1
            rank = conversion_rank(c_status)
            if rank == 3:
                secured_count += 1
            elif rank == 2:
                in_progress_count += 1

            if area_val:
                total_applied_area += area_val

        total_conv_lots = rec_count

        conv_secured_pct = (secured_count / total_conv_lots * 100.0) if total_conv_lots > 0 else 0.0

        conversion_breakdown = []
        for s_name, count in sorted(conv_status_counts.items(), key=lambda x: x[1], reverse=True)[:8]:
            pct = (count / (total_conv_lots or 1)) * 100.0
            conversion_breakdown.append({
                'status': s_name,
                'count': count,
                'percent': f"{pct:.1f}%",
                'percent_val': round(pct, 1)
            })

        # Compare the latest recorded value for each account + field between the
        # selected work week and the nearest earlier work week, or retrieve direct audit
        # log changes for the selected period if comparisons yield zero deltas.
        # The calendar week before the selected one, even if it had no logs.
        # (Using the previous week *with logs* made "vs WW-39" really mean WW-37.)
        previous_ww = ''
        if selected_ww in week_dates:
            _, sel_start, _ = get_work_week_info(week_dates[selected_ww])
            previous_ww = get_work_week_label(sel_start - timedelta(days=7))

        def account_key(log: Dict[str, Any]) -> str:
            return str(log.get('cr63f_control1') or log.get('cr63f_control') or log.get('Control') or log.get('Control Number') or log.get('area_index') or log.get('cr63f_areaindex') or log.get('title_no') or log.get('cr63f_titleno') or '').strip().upper()

        def field_key(log: Dict[str, Any]) -> str:
            return str(log.get('field_name') or log.get('cr63f_fieldchanged') or log.get('Field Changed') or '').strip()

        def build_index(fields: tuple) -> Dict[str, List[Dict[str, Any]]]:
            index: Dict[str, List[Dict[str, Any]]] = {}
            for mr in all_records:
                values = {str(mr.get(f) or '').strip() for f in fields}
                values.discard('')
                for value in values:
                    index.setdefault(value, []).append(mr)
            return index

        ctrl_index = build_index(('Control Number', 'cr63f_control1', 'Control'))
        area_index_map = build_index(('Area-Index', 'cr63f_areaindex', 'Area Index'))
        title_index = build_index(('Title No', 'cr63f_titleno', 'Title'))

        def match_master_rec(log: Dict[str, Any]) -> Optional[Dict[str, Any]]:
            ctrl_no = str(log.get('cr63f_control1') or log.get('cr63f_control') or log.get('Control') or log.get('Control Number') or '').strip()
            area_idx = str(log.get('area_index') or log.get('cr63f_areaindex') or log.get('Area-Index') or '').strip()
            title_no = str(log.get('title_no') or log.get('cr63f_titleno') or log.get('Title No') or '').strip()
            if not ctrl_no and not area_idx and not title_no:
                return None
            # Strongest key first across all records, and only a unique match:
            # a Title match on an earlier row must not beat an exact Control match.
            # The indexes are built once per call (scanning every account for
            # every log made a busy week take many seconds).
            for wanted, index in ((ctrl_no, ctrl_index), (area_idx, area_index_map), (title_no, title_index)):
                if not wanted:
                    continue
                matches = index.get(wanted, [])
                if len(matches) == 1:
                    return matches[0]
                if matches:
                    return None
            return None

        # ---- Weekly insights (UI v2): week-over-week deltas, account-grouped changes, exceptions ----
        insights = self._build_weekly_insights(
            logs=logs,
            all_records=all_records,
            records=records,
            selected_ww=selected_ww,
            previous_ww=previous_ww,
            active_mode=active_mode,
            log_week_label=log_week_label,
            log_datetime=log_datetime,
            account_key=account_key,
            field_key=field_key,
            match_master_rec=match_master_rec,
        )

        return {
            'success': True,
            'selected_work_week': selected_ww,
            'selected_mode': active_mode,
            'available_work_weeks': available_ww,
            'total_records': len(records),
            'all_records_count': len(all_records),
            'total_tcp': total_tcp,
            'total_tcp_formatted': f"₱{format_currency(total_tcp)}",
            'total_deductible': total_ded,
            'total_deductible_formatted': f"₱{format_currency(total_ded)}",
            'total_balance': total_bal,
            'total_balance_formatted': f"₱{format_currency(total_bal)}",
            'overall_paid_percent': f"{overall_paid_pct:.2f}%",
            'overall_paid_percent_num': round(overall_paid_pct, 2),
            'modes_distribution': modes_distribution,
            'financial_breakdown': financial_breakdown,
            'conversion_metrics': {
                'total_lots': total_conv_lots,
                'secured_count': secured_count,
                'in_progress_count': in_progress_count,
                'total_applied_area': round(total_applied_area, 2),
                'total_applied_area_formatted': f"{total_applied_area:,.2f} Has.",
                'secured_percent': f"{conv_secured_pct:.1f}%",
                'secured_percent_num': round(conv_secured_pct, 1),
                'breakdown': conversion_breakdown
            },
            'comparison_previous_week': previous_ww,
            'week_summary': insights['week_summary'],
            'changed_accounts': insights['changed_accounts'],
            'needs_attention': insights['needs_attention'],
        }

    # -- Overview helpers (UI v2) ---------------------------------------------
    @staticmethod
    def _field_kind(field_name: Any) -> str:
        f = re.sub(r'[^a-z0-9]', '', str(field_name or '').lower())
        if not f:
            return ''
        if 'deductible' in f or f in ('paymentdeductibles', 'totalpayment'):
            return 'deductible'
        if 'contractprice' in f or f == 'tcp':
            return 'tcp'
        if 'conversionstatus' in f or 'darstatus' in f:
            return 'conversion'
        return ''

    @staticmethod
    def _conversion_rank(status: Any) -> int:
        return conversion_rank(status)

    @staticmethod
    def _source_of(action_type: Any) -> str:
        a = str(action_type or '').upper()
        if 'REVERT' in a:
            return 'revert'
        if 'DELETE' in a:
            return 'manual'
        if 'NEW' in a or 'ADD' in a:
            return 'new'
        if 'REFERENCE' in a or 'FINANCE' in a or 'SYNC' in a or 'CONVERSION' in a:
            return 'reference'
        return 'manual'

    def _build_weekly_insights(self, logs, all_records, records, selected_ww, previous_ww, active_mode,
                               log_week_label, log_datetime, account_key, field_key, match_master_rec) -> dict:
        """Summarise one work week for the Overview: deltas, per-account changes and exceptions."""

        def fnum(v: Any) -> float:
            return parse_float(v) or 0.0

        def rec_mode(rec: Optional[Dict[str, Any]], log: Dict[str, Any]) -> str:
            m = ''
            if rec:
                m = str(rec.get('new_mode') or rec.get('Mode') or rec.get('MODE') or '').strip().upper()
            if not m:
                m = str(log.get('new_mode') or log.get('Mode') or log.get('MODE') or '').strip().upper()
            return m

        session_user = ''
        try:
            session_user = self.dv_client.current_user_name() if self.dv_client else ''
        except Exception:
            session_user = ''
        is_single_week = bool(selected_ww) and selected_ww.upper() != 'ALL'
        week_logs = [l for l in logs if (log_week_label(l) == selected_ww)] if is_single_week else list(logs)
        week_logs.sort(key=lambda l: log_datetime(l) or datetime.min)

        accounts: Dict[str, Dict[str, Any]] = {}
        new_entries = 0
        field_edits = 0
        tcp_delta = 0.0
        ded_delta = 0.0
        balance_delta = 0.0
        secured_delta = 0
        direct_ded_delta = 0.0      # direct sale only (decision 2026-10-14)
        direct_balance_delta = 0.0

        for log in week_logs:
            acct = account_key(log)
            if not acct:
                continue
            rec = match_master_rec(log)
            mode_val = rec_mode(rec, log)
            if active_mode and mode_val and mode_val != active_mode:
                continue
            fld = field_key(log)
            action = log.get('action_type') or log.get('Action Type') or ''
            source = self._source_of(action)
            ts = log_datetime(log)
            author = str(log.get('createdby') or log.get('author') or log.get('modifiedby')
                         or log.get('cr63f_createdby') or log.get('User') or '').strip()
            if not author and not (log.get('createdon') or log.get('__createdon_raw')):
                # Saved in this session and not fetched back yet: it is ours.
                author = session_user
            author = author or 'Unknown'

            entry = accounts.get(acct)
            if entry is None:
                tcp_raw = (rec or {}).get('Total Contract Price') or (rec or {}).get('cr63f_totalcontractprice') \
                    or log.get('cr63f_totalcontractprice') or log.get('Total Contract Price')
                entry = accounts[acct] = {
                    'account': acct,
                    'control_number': str(log.get('cr63f_control1') or log.get('cr63f_control') or log.get('Control Number') or (rec or {}).get('cr63f_control1') or ''),
                    'area_index': str(log.get('area_index') or log.get('cr63f_areaindex') or log.get('Area-Index') or (rec or {}).get('cr63f_areaindex') or ''),
                    'title_no': str(log.get('title_no') or log.get('cr63f_titleno') or log.get('Title No') or ''),
                    'mode': mode_val,
                    'contract_price': fnum(tcp_raw),
                    'sources': [],
                    'fields': {},
                    'last_time': '',
                    'last_by': author,
                    'is_new': False,
                    'rec': rec,
                }
            if source not in entry['sources']:
                entry['sources'].append(source)
            entry['last_time'] = str(log.get('log_time') or log.get('createdon') or log.get('timestamp') or entry['last_time'])
            entry['last_by'] = author

            if source == 'new' or (fld and fld.lower() == 'initial entry'):
                if not entry['is_new']:
                    entry['is_new'] = True
                    new_entries += 1
                    new_tcp = fnum(log.get('cr63f_totalcontractprice') or log.get('Total Contract Price'))
                    new_ded = fnum(log.get('cr63f_totalpaymentdeductibletotcp') or log.get('Total Payment Deductible to TCP'))
                    tcp_delta += new_tcp
                    ded_delta += new_ded
                    balance_delta += max(0.0, new_tcp - new_ded)
                    if _payment_group(entry['mode']) == 'DIRECT_SALE':
                        direct_ded_delta += new_ded
                        direct_balance_delta += max(0.0, new_tcp - new_ded)
                    if self._conversion_rank(log.get('cr63f_conversionstatus') or log.get('Conversion Status')) == 3:
                        secured_delta += 1
                continue
            if not fld:
                continue

            old_v = log.get('old_value', log.get('cr63f_previousvalue', log.get('Previous Value', '')))
            new_v = log.get('new_value', log.get('cr63f_newvalue', log.get('New Value', '')))
            if _norm_text_str(old_v) == _norm_text_str(new_v):
                continue
            field_edits += 1
            f_entry = entry['fields'].get(fld)
            if f_entry is None:
                entry['fields'][fld] = {
                    'field_name': fld, 'old_value': old_v, 'new_value': new_v,
                    'source': source, 'log_id': str(log.get('__id') or log.get('id') or ''),
                    'log_time': str(log.get('log_time') or log.get('createdon') or ''), 'updated_by': author,
                }
            else:
                f_entry.update({'new_value': new_v, 'source': source,
                                'log_id': str(log.get('__id') or log.get('id') or f_entry['log_id']),
                                'log_time': str(log.get('log_time') or log.get('createdon') or f_entry['log_time']),
                                'updated_by': author})

        needs_attention: List[Dict[str, Any]] = []
        changed_accounts: List[Dict[str, Any]] = []
        for acct, entry in accounts.items():
            # Balance change per account, each balance clamped at 0 like the
            # portfolio total (tcp_delta - ded_delta would count overpayment).
            if not entry['is_new']:
                tcp_f = next((f for fld, f in entry['fields'].items() if self._field_kind(fld) == 'tcp'), None)
                ded_f = next((f for fld, f in entry['fields'].items() if self._field_kind(fld) == 'deductible'), None)
                if tcp_f or ded_f:
                    rec = entry.get('rec') or {}
                    cur_tcp = fnum(rec.get('cr63f_totalcontractprice') or rec.get('Total Contract Price'))
                    cur_ded = fnum(rec.get('cr63f_totalpaymentdeductibletotcp') or rec.get('Total Payment Deductible to TCP'))
                    old_tcp, new_tcp = (fnum(tcp_f['old_value']), fnum(tcp_f['new_value'])) if tcp_f else (cur_tcp, cur_tcp)
                    old_ded, new_ded = (fnum(ded_f['old_value']), fnum(ded_f['new_value'])) if ded_f else (cur_ded, cur_ded)
                    balance_delta += max(0.0, new_tcp - new_ded) - max(0.0, old_tcp - old_ded)
                    if _payment_group(entry['mode']) == 'DIRECT_SALE':
                        direct_balance_delta += max(0.0, new_tcp - new_ded) - max(0.0, old_tcp - old_ded)
            changes = []
            for fld, f in entry['fields'].items():
                if _norm_text_str(f['old_value']) == _norm_text_str(f['new_value']):
                    continue
                kind = self._field_kind(fld)
                if not entry['is_new']:
                    if kind == 'tcp':
                        tcp_delta += fnum(f['new_value']) - fnum(f['old_value'])
                    elif kind == 'deductible':
                        diff = fnum(f['new_value']) - fnum(f['old_value'])
                        ded_delta += diff
                        if _payment_group(entry['mode']) == 'DIRECT_SALE':
                            direct_ded_delta += diff
                        if diff < -0.005:
                            needs_attention.append({
                                'kind': 'payment_down', 'severity': 'high', 'account': acct,
                                'title': 'Payment went down',
                                'detail': f"₱{format_currency(fnum(f['old_value']))} → ₱{format_currency(fnum(f['new_value']))}",
                                'log_time': f['log_time'], 'source': f['source'],
                            })
                    elif kind == 'conversion':
                        r_old, r_new = self._conversion_rank(f['old_value']), self._conversion_rank(f['new_value'])
                        if r_new == 3 and r_old < 3:
                            secured_delta += 1
                        elif r_old == 3 and r_new < 3:
                            secured_delta -= 1
                        if r_new < r_old:
                            needs_attention.append({
                                'kind': 'status_back', 'severity': 'medium', 'account': acct,
                                'title': 'Status moved backwards',
                                'detail': f"{str(f['old_value']).strip() or '(blank)'} → {str(f['new_value']).strip() or '(blank)'}",
                                'log_time': f['log_time'], 'source': f['source'],
                            })
                changes.append(f)
            if not changes and not entry['is_new']:
                continue
            changed_accounts.append({
                'account': acct,
                'control_number': entry['control_number'],
                'area_index': entry['area_index'],
                'title_no': entry['title_no'],
                'mode': entry['mode'],
                'contract_price': entry['contract_price'],
                'contract_price_formatted': f"₱{format_currency(entry['contract_price'])}" if entry['contract_price'] else '—',
                'is_new': entry['is_new'],
                'sources': entry['sources'],
                'last_time': entry['last_time'],
                'last_by': entry['last_by'],
                'changes': changes,
            })
        changed_accounts.sort(key=lambda a: parse_log_datetime(a['last_time']) or datetime.min, reverse=True)

        # Portfolio-level data quality: deductible above contract price.
        for r in records:
            tcp_v = fnum(r.get('Total Contract Price', r.get('cr63f_totalcontractprice', 0)))
            ded_v = fnum(r.get('Total Payment Deductible to TCP', r.get('cr63f_totalpaymentdeductibletotcp', 0)))
            if tcp_v > 0 and ded_v - tcp_v > 0.01:
                acct = str(r.get('cr63f_control1') or r.get('Control Number') or r.get('cr63f_areaindex') or r.get('Area-Index') or '').strip()
                needs_attention.append({
                    'kind': 'over_paid', 'severity': 'high', 'account': acct,
                    'title': 'Deductible exceeds TCP',
                    'detail': f"{ded_v / tcp_v * 100:.1f}% paid · balance is clamped at 0",
                    'log_time': '', 'source': '',
                })
        sev_order = {'high': 0, 'medium': 1}
        needs_attention.sort(key=lambda n: (sev_order.get(n['severity'], 2), n['kind']))

        week_start = week_end = ''
        if is_single_week:
            any_dt = next((log_datetime(l) for l in week_logs if log_datetime(l)), None)
            if any_dt:
                _, wk_start, wk_end = get_work_week_info(any_dt)  # Tuesday .. Monday
                week_start, week_end = wk_start.strftime('%Y-%m-%d'), wk_end.strftime('%Y-%m-%d')

        return {
            'week_summary': {
                'work_week': selected_ww,
                'previous_week': previous_ww,
                'week_start': week_start,
                'week_end': week_end,
                'accounts_changed': len(changed_accounts),
                'field_edits': field_edits,
                'new_entries': new_entries,
                'tcp_delta': round(tcp_delta, 2),
                'deductible_delta': round(ded_delta, 2),
                'balance_delta': round(balance_delta, 2),
                'direct_sale_deductible_delta': round(direct_ded_delta, 2),
                'direct_sale_balance_delta': round(direct_balance_delta, 2),
                'secured_delta': secured_delta,
                'source_counts': {
                    k: sum(1 for a in changed_accounts if k in a['sources'])
                    for k in ('reference', 'manual', 'new', 'revert')
                },
            },
            'changed_accounts': changed_accounts,
            'needs_attention': needs_attention,
        }

    def export_weekly_report(self, work_week: Optional[str] = None, mode: Optional[str] = None) -> dict:
        """Export the Overview for one work week as an Excel workbook (Summary, Changed accounts, Needs attention)."""
        data = self.get_dashboard_analytics(work_week, mode)
        if not data.get('success'):
            return {'success': False, 'error': 'Could not build the weekly overview.'}
        wk = str(data.get('selected_work_week') or 'All weeks')
        safe_wk = re.sub(r'[^A-Za-z0-9-]+', '_', wk).strip('_') or 'All'
        default_name = f"LOA_Weekly_Report_{safe_wk}.xlsx"
        if not self._window:
            return {'success': False, 'error': 'Window not attached.'}
        result = self._window.create_file_dialog(webview.SAVE_DIALOG, save_filename=default_name,
                                                 file_types=('Excel Files (*.xlsx)',))
        save_path = _dialog_path(result)
        if not save_path:
            return {'success': False, 'cancelled': True}
        if not save_path.lower().endswith('.xlsx'):
            save_path += '.xlsx'
        try:
            self._write_weekly_report(save_path, data)
            return {'success': True, 'path': save_path}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    @staticmethod
    def _write_weekly_report(save_path: str, data: dict) -> None:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter

        ws_sum = data.get('week_summary') or {}
        wb = openpyxl.Workbook()
        head_font = Font(bold=True, color='FFFFFF')
        head_fill = PatternFill('solid', fgColor='1D5FD1')
        money = '"₱"#,##0.00'

        s = wb.active
        s.title = 'Summary'
        s['A1'] = 'LOA Monitoring — Weekly report'
        s['A1'].font = Font(bold=True, size=14)
        period = str(data.get('selected_work_week') or 'All weeks')
        if ws_sum.get('week_start'):
            period += f" ({ws_sum['week_start']} to {ws_sum['week_end']})"
        if ws_sum.get('previous_week'):
            period += f" vs {ws_sum['previous_week']}"
        s['A2'] = period
        s['A2'].font = Font(color='5B6170')
        rows = [
            ('This week', None, None),
            ('Accounts changed', ws_sum.get('accounts_changed', 0), None),
            ('Field edits', ws_sum.get('field_edits', 0), None),
            ('New entries', ws_sum.get('new_entries', 0), None),
            ('Contract value change', ws_sum.get('tcp_delta', 0.0), money),
            ('Payments deducted change', ws_sum.get('deductible_delta', 0.0), money),
            ('Balance change', ws_sum.get('balance_delta', 0.0), money),
            ('Lots newly DAR secured (net)', ws_sum.get('secured_delta', 0), None),
            ('Needs attention', len(data.get('needs_attention') or []), None),
            ('', None, None),
            ('Portfolio', None, None),
            ('Accounts', data.get('total_records', 0), None),
            ('Contract value', data.get('total_tcp', 0.0), money),
            ('Payments deducted', data.get('total_deductible', 0.0), money),
            ('Balance outstanding', data.get('total_balance', 0.0), money),
            ('% paid', (data.get('overall_paid_percent_num') or 0) / 100.0, '0.00%'),
            ('DAR secured lots', (data.get('conversion_metrics') or {}).get('secured_count', 0), None),
            ('% secured', ((data.get('conversion_metrics') or {}).get('secured_percent_num') or 0) / 100.0, '0.0%'),
        ]
        r = 4
        for label, val, fmt in rows:
            s.cell(row=r, column=1, value=label)
            if val is None and label:
                s.cell(row=r, column=1).font = Font(bold=True)
            if val is not None:
                c = s.cell(row=r, column=2, value=val)
                if fmt:
                    c.number_format = fmt
            r += 1
        s.column_dimensions['A'].width = 32
        s.column_dimensions['B'].width = 22

        def table(ws, headers, body, widths):
            for i, h in enumerate(headers, 1):
                c = ws.cell(row=1, column=i, value=h)
                c.font, c.fill = head_font, head_fill
                c.alignment = Alignment(vertical='center')
                ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
            for ri, row in enumerate(body, 2):
                for ci, v in enumerate(row, 1):
                    set_text_safe(ws.cell(row=ri, column=ci), v)
            ws.freeze_panes = 'A2'

        src_label = {'reference': 'Reference update', 'manual': 'Manual edit', 'new': 'New entry', 'revert': 'Revert'}
        body = []
        for a in data.get('changed_accounts') or []:
            if not a.get('changes'):
                body.append([a['account'], a.get('area_index'), a.get('mode'), a.get('contract_price') or None,
                             'New account', '', '', 'New entry', a.get('last_time'), a.get('last_by')])
            for ch in a.get('changes') or []:
                body.append([a['account'], a.get('area_index'), a.get('mode'), a.get('contract_price') or None,
                             ch.get('field_name'), str(ch.get('old_value') or ''), str(ch.get('new_value') or ''),
                             src_label.get(ch.get('source'), ch.get('source')), ch.get('log_time'), ch.get('updated_by')])
        t = wb.create_sheet('Changed accounts')
        table(t, ['Account', 'Area index', 'Mode', 'Contract price', 'Field', 'Previous value', 'New value',
                  'Source', 'When', 'By'], body, [16, 14, 16, 18, 30, 24, 24, 18, 22, 14])
        for ri in range(2, len(body) + 2):
            t.cell(row=ri, column=4).number_format = money

        n = wb.create_sheet('Needs attention')
        table(n, ['Flag', 'Account', 'Detail', 'When'],
              [[x.get('title'), x.get('account'), x.get('detail'), x.get('log_time')] for x in data.get('needs_attention') or []],
              [24, 18, 60, 22])
        wb.save(save_path)

    def export_logs_excel(self, logs: list, filename: Optional[str] = None, columns: Optional[list] = None) -> dict:
        """Export entry logs to Excel file."""
        if not logs:
            return {'success': False, 'error': 'No logs to export.'}
        default_name = filename or f"Entry_Logs_Export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        if not self._window:
            return {'success': False, 'error': 'Window not attached.'}
        result = self._window.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename=default_name,
            file_types=('Excel Files (*.xlsx)',)
        )
        save_path = _dialog_path(result)
        if not save_path:
            return {'success': False, 'cancelled': True}
        if not save_path.lower().endswith('.xlsx'):
            save_path += '.xlsx'

        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Entry Logs"
            ws.views.sheetView[0].showGridLines = True

            cols = columns or ['Log Time', 'Work Week', 'Action Type', 'Control Number', 'AREA-INDEX', 'Field Changed', 'Previous Value', 'New Value', 'Title No', 'Details']
            hdr_font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
            hdr_fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
            thin_border = Border(
                left=Side(style='thin', color='CBD5E1'),
                right=Side(style='thin', color='CBD5E1'),
                top=Side(style='thin', color='CBD5E1'),
                bottom=Side(style='thin', color='CBD5E1')
            )

            for col_idx, col_name in enumerate(cols, 1):
                cell = ws.cell(row=1, column=col_idx, value=col_name)
                cell.font = hdr_font
                cell.fill = hdr_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')

            key_map = {
                'Log Time': ['log_time', 'createdon', 'timestamp'],
                'Work Week': ['work_week', 'Work Week'],
                'Action Type': ['action_type', 'Action Type', 'cr63f_actiontype'],
                'Control Number': ['cr63f_control1', 'cr63f_control', 'Control', 'Control Number', 'Control No'],
                'AREA-INDEX': ['area_index', 'cr63f_areaindex', 'Area-Index', 'AREA-INDEX'],
                'Title No': ['title_no', 'cr63f_titleno', 'Title No'],
                'Field Changed': ['field_name', 'Field Changed', 'cr63f_fieldchanged'],
                'Previous Value': ['old_value', 'Previous Value', 'cr63f_previousvalue'],
                'New Value': ['new_value', 'New Value', 'cr63f_newvalue'],
                'Details': ['details', 'Details']
            }

            for row_idx, log_row in enumerate(logs, 2):
                for col_idx, col_name in enumerate(cols, 1):
                    possible_keys = key_map.get(col_name, [col_name, col_name.lower().replace(' ', '_')])
                    val = ''
                    for k in possible_keys:
                        if k in log_row and log_row[k] is not None:
                            val = log_row[k]
                            break
                    cell = ws.cell(row=row_idx, column=col_idx)
                    set_text_safe(cell, str(val))
                    cell.font = Font(name='Segoe UI', size=9.5)
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical='center')

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            wb.save(save_path)
            return {'success': True, 'file_path': save_path}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def delete_loa_record_entry(self, record_id: str) -> dict:
        """Delete a Land Master account, audited like every other change.

        A DELETE audit row (with the account's full last snapshot) is written
        first; if it cannot be written, nothing is deleted. If the delete itself
        then fails, that audit row is removed again so the log never claims a
        deletion that did not happen.
        """
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        record = next((dict(r) for r in self.loa_records if str(r.get('__id') or '') == str(record_id)), None)
        try:
            if record is None:
                table_name = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
                record = self.dv_client.get_record(table_name, record_id)
        except Exception as exc:
            return {'success': False, 'error': f'Cannot read the account to delete: {exc}'}
        record['__id'] = record.get('__id') or record_id
        control = str(_get_field_val(record, ['cr63f_control1', 'Control', 'Control Number']) or '').strip()
        area_idx = str(_get_field_val(record, ['cr63f_areaindex', 'Area-Index', 'AREA-INDEX']) or '').strip()
        log_entry = build_log_entry(
            action_type='DELETE', work_week=get_work_week(datetime.now()),
            now_str=datetime.now().strftime('%m/%d/%Y %I:%M %p'),
            area_index=area_idx,
            index_no=str(_get_field_val(record, ['cr63f_indexno', 'Index No']) or ''),
            title_no=str(_get_field_val(record, ['cr63f_titleno', 'Title No']) or ''),
            field_name='Deleted Record', old_value='Active', new_value='Deleted',
            details=f"Deleted account {control or area_idx or record_id}",
            full_record=record,
        )
        try:
            saved_log = self.dv_client.save_log_record(log_entry)
            log_entry['__id'] = (saved_log or {}).get('id', '')
        except Exception as exc:
            return {'success': False, 'error': f'Nothing was deleted: the audit row could not be written ({exc}).'}
        try:
            table_name = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
            info = self.dv_client.get_entity_info(table_name)
            self.dv_client.delete(info['entity_set'], record_id)
        except Exception as exc:
            try:
                if log_entry.get('__id'):
                    self.dv_client.delete_log_record(log_entry['__id'])
            except Exception as cleanup_error:
                return {'success': False, 'error': f'Delete failed ({exc}) and its audit row could not be removed ({cleanup_error}). Check Dataverse.'}
            return {'success': False, 'error': f'Delete failed: {exc}'}
        self.loa_records = [r for r in self.loa_records if r.get('__id') != record_id]
        self.pending_session_logs.append(dict(log_entry))
        self._add_log_records([log_entry])
        return {'success': True, 'message': f"Account {control or area_idx} deleted from Dataverse and logged in Entry Logs."}

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
        file_path = _dialog_path(result)
        if not file_path:
            return {'success': False, 'cancelled': True}
        try:
            headers, raw_rows = read_tabular_file(file_path)
            # Run lookup calculation across all rows
            processed_rows: List[Dict[str, Any]] = []
            for r in raw_rows:
                type_val = r.get("TYPE", r.get("Type", ""))
                index_no = r.get("INDEX NO", r.get("Index No", r.get("Corridor Index", "")))
                area_index = r.get("AREA-INDEX", r.get("Area Index", r.get("AREA INDEX", "")))
                title_mod = r.get("TITLE MOD", r.get("Title Mod", ""))

                mode_val = r.get("MODE", r.get("Mode", ""))
                lookup_res = self.finance_engine.lookup(type_val, index_no, area_index, title_mod, mode_val)

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
            save_path = _dialog_path(result)
            if not save_path:
                return {'success': False, 'cancelled': True}

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
            'conversion_table': cfg.get('conversion_table', 'cr63f_landconversionreference'),
            'logs_table': cfg.get('logs_table', 'cr63f_loamonitoringlogs'),
            'theme': self.settings.get('theme', 'light'),
        }

    def save_settings(self, new_settings: dict) -> dict:
        try:
            # Only the keys that were sent are changed. A theme-only save used
            # to rewrite the connection with blanks, so it reset on next start.
            # A blank value keeps the saved one (load_config fills defaults).
            conn_keys = ('environment_url', 'tenant_id', 'client_id', 'solution_name', 'solution_id',
                         'loa_table', 'finance_table', 'conversion_table', 'logs_table')
            updates = {k: str(new_settings[k]).strip() for k in conn_keys
                       if k in new_settings and str(new_settings[k] or '').strip()}
            if updates:
                cfg = dv.load_config()
                cfg.update(updates)
                cfg['environment_url'] = cfg['environment_url'].rstrip('/')
                dv.save_config(cfg)
                self._init_dataverse()

            if new_settings.get('theme'):
                self.settings['theme'] = new_settings['theme']
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


def _clear_web_cache(web_storage: str) -> None:
    """Drop WebView2's HTTP/code caches so edited web files load on the next start.

    Cookies, local storage and sign-in live elsewhere in the profile and are kept.
    """
    import shutil
    default_dir = os.path.join(web_storage, 'EBWebView', 'Default')
    for name in ('Cache', 'Code Cache', 'GPUCache', 'Service Worker'):
        target = os.path.join(default_dir, name)
        if os.path.isdir(target):
            shutil.rmtree(target, ignore_errors=True)


def main() -> None:
    api = Api()
    window = webview.create_window(
        title='',
        url=_index_url(),
        js_api=api,
        width=1380,
        height=860,
        min_size=(960, 600),
        background_color='#0F172A'
    )
    api.set_window(window)

    def _auto_check_updates():
        # Quiet check shortly after the page loads; a result only interrupts the
        # user when a new version is found.
        threading.Timer(1.5, lambda: api.check_for_updates(False)).start()

    window.events.loaded += _auto_check_updates

    def _hide_title_bar_icon():
        # Blank title-bar icon (decision 2026-10-14). WinForms ShowIcon=False hides
        # only the caption icon; the taskbar keeps the app icon.
        try:
            form = getattr(window, 'native', None)
            if form is None:
                return
            from System import Action  # pythonnet, loaded by pywebview's WinForms backend

            def _apply():
                form.ShowIcon = False
            if form.InvokeRequired:
                form.Invoke(Action(_apply))
            else:
                _apply()
        except Exception:
            pass  # cosmetic only; other platforms/backends have no WinForms form

    window.events.shown += _hide_title_bar_icon

    web_storage = os.path.join(ensure_user_data(), 'web_profile')
    os.makedirs(web_storage, exist_ok=True)
    _clear_web_cache(web_storage)
    webview.start(storage_path=web_storage, debug=False)


if __name__ == '__main__':
    main()

