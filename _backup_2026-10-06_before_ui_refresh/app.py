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
from datetime import datetime
from typing import Any, Dict, List, Optional

# Ensure directory is on sys.path
_base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
if _base_dir and _base_dir not in sys.path:
    sys.path.insert(0, _base_dir)

import webview

from app_paths import data_file, ensure_user_data
import dataverse as dv
from loa_engine import (
    ConversionLookupEngine,
    FinanceLookupEngine,
    calculate_paid_percent,
    clean_title_variants,
    format_calculated_percent,
    format_currency,
    format_percent,
    get_work_week,
    norm_key,
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
    has_pct_sign = "%" in s
    f = parse_float(s)
    if f is None:
        return None
    if has_pct_sign:
        return f
    # Only true fractions between -1.0 and 1.0 exclusive (in Dataverse, 1.0 = 1.00%, 100.0 = 100.00%)
    if -1.0 < f < 1.0 and f != 0:
        return f * 100.0
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
    return abs(pa - pb) > 0.05


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
    entry.update({
        # Primary Name
        'cr63f_name': f"{area_index or 'Lot'} - {field_name} ({action_type})",
        'name': f"{area_index or 'Lot'} - {field_name} ({action_type})",

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
        if -1.0 < f < 1.0 and f != 0.0:
            f *= 100.0
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
}
CALCULATED_AUDIT_FIELDS = {'Balance', '% Paid', 'Paid', 'Percent Paid'}


def _audit_values_equal(field_name: str, left: Any, right: Any) -> bool:
    """Compare stored values without display-format rounding."""
    canonical = CANONICAL_FIELD_MAP.get(field_name, field_name)
    spec = next((entry for entry in LOG_COMPARE_SPECS if entry[0] == canonical), None)
    if spec and spec[2] == 'currency':
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
    return _norm_text_str(left) == _norm_text_str(right)


def derive_field_diffs_from_snapshots(
    raw_records: List[Dict[str, Any]],
    active_session_logs: Optional[List[Dict[str, Any]]] = None
) -> List[Dict[str, Any]]:
    """
    Derives individual field diff records by chronologically sorting lot snapshots
    from cr63f_loamonitoringlogs and comparing consecutive snapshots using Control Number as primary key.
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
        action = str(
            rec.get('action_type') or rec.get('Action Type') or
            rec.get('cr63f_actiontype') or ''
        ).strip().upper().replace(' ', '_')

        # Some Dataverse schemas only retain the primary name. build_log_entry
        # includes the source action in parentheses there, so use it as fallback.
        if not action:
            name = str(rec.get('cr63f_name') or rec.get('name') or '')
            match = re.search(r'\(([^()]+)\)\s*$', name)
            if match:
                action = match.group(1).strip().upper().replace(' ', '_')

        if 'BASELINE' in action:
            return 'BASELINE'
        if 'REVERT' in action:
            return 'REVERT'
        if 'REFERENCE' in action:
            return 'REFERENCE_UPDATE'
        if 'FINANCE' in action:
            return 'FINANCE_UPDATE'
        if 'NEW' in action:
            return 'NEW_ENTRY'
        return 'MANUAL_EDIT'

    def normalize_audit_week(rec: Dict[str, Any]) -> Dict[str, Any]:
        """Audit work week is always derived from log time (Wednesday-Tuesday)."""
        normalized = dict(rec)
        timestamp = get_log_timestamp(normalized)
        calculated = get_work_week(timestamp) if timestamp else ''
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
                if canonical in CALCULATED_AUDIT_FIELDS:
                    continue
                if action == 'REFERENCE_UPDATE' and canonical not in REFERENCE_AUDIT_FIELDS:
                    continue
                explicit_diff_logs.append(normalize_audit_week(r))
            else:
                snapshots.append(dict(r))

    def get_lot_key(rec: dict) -> str:
        ctrl = str(rec.get('cr63f_control1') or rec.get('cr63f_control') or rec.get('Control') or rec.get('Control Number') or rec.get('Control No.') or rec.get('Control No') or rec.get('control_no') or '').strip().upper()
        if ctrl:
            return f"CTRL:{ctrl}"
        ai = str(rec.get('cr63f_areaindex') or rec.get('AREA-INDEX') or rec.get('Area-Index') or rec.get('area_index') or rec.get('Area Index') or '').strip().upper()
        if ai:
            return f"AI:{ai}"
        t_no = str(rec.get('cr63f_titleno') or rec.get('Title No') or rec.get('title_no') or rec.get('TITLE NO') or rec.get('Title No.') or '').strip().upper()
        if t_no:
            return f"TITLE:{t_no}"
        rec_id = str(rec.get('__id') or rec.get('id') or rec.get('cr63f_loamonitoringlogid') or '').strip()
        return f"ID:{rec_id}"

    def get_time_key(rec: dict) -> str:
        return str(get_log_timestamp(rec))

    lot_groups: Dict[str, List[Dict[str, Any]]] = {}
    for s in snapshots:
        k = get_lot_key(s)
        lot_groups.setdefault(k, []).append(s)

    derived_diffs: List[Dict[str, Any]] = []

    for lot_key, group in lot_groups.items():
        group.sort(key=get_time_key)

        for i, curr_s in enumerate(group):
            ts_str = get_log_timestamp(curr_s) or datetime.now().strftime('%m/%d/%Y %I:%M %p')
            # Never trust a stale work-week value stored on the source row. The
            # audit week belongs to the log timestamp, not to Endorsed Date.
            ww = get_work_week(ts_str) or get_work_week(datetime.now())
            ctrl_val = curr_s.get('cr63f_control1') or curr_s.get('cr63f_control') or curr_s.get('Control') or curr_s.get('Control Number') or ''
            area_idx = curr_s.get('cr63f_areaindex') or curr_s.get('Area-Index') or curr_s.get('area_index') or ''
            index_no = curr_s.get('cr63f_indexno') or curr_s.get('Index No') or curr_s.get('index_no') or ''
            title_no = curr_s.get('cr63f_titleno') or curr_s.get('Title No') or curr_s.get('title_no') or ''
            rec_log_id = curr_s.get('__id') or curr_s.get('id') or curr_s.get('cr63f_loamonitoringlogid') or ''

            action_type = get_source_action(curr_s)
            # Baselines exist only to anchor the first post-deployment comparison.
            # They are never exposed as user-facing Entry Logs.
            if action_type == 'BASELINE':
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
                if not explicit_action and 'NEW_ENTRY' not in source_name:
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

    def _write_baseline_if_missing(
        self, record: Dict[str, Any], now_str: str
    ) -> Optional[Dict[str, Any]]:
        """Persist one hidden pre-change snapshot when this account has no audit history."""
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
        account_keys = {value.upper() for value in (control, area_index, title_no) if value}
        for row in self.pending_session_logs:
            row_keys = {
                str(row.get('cr63f_control') or row.get('cr63f_control1') or row.get('Control Number') or '').strip().upper(),
                str(row.get('cr63f_areaindex') or row.get('Area-Index') or row.get('AREA-INDEX') or '').strip().upper(),
                str(row.get('cr63f_titleno') or row.get('Title No') or row.get('TITLE NO') or '').strip().upper(),
            }
            row_keys.discard('')
            if account_keys.intersection(row_keys):
                return None
        clauses = []
        for logical, value in (
            ('cr63f_control', control),
            ('cr63f_areaindex', area_index),
            ('cr63f_titleno', title_no),
        ):
            if value:
                escaped = value.replace("'", "''")
                clauses.append(f"{logical} eq '{escaped}'")
        if not clauses:
            raise RuntimeError('Cannot establish an audit baseline without a Control Number, AREA-INDEX, or Title No.')

        logs_table = self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs')
        history = self.dv_client.fetch_records(
            logs_table, custom_only=False, filter_query=' or '.join(clauses)
        )
        identity = control or area_index or title_no
        history = [
            row for row in history
            if identity.upper() in {
                str(row.get('cr63f_control') or row.get('cr63f_control1') or row.get('Control') or row.get('Control Number') or '').strip().upper(),
                str(row.get('cr63f_areaindex') or row.get('Area-Index') or row.get('AREA-INDEX') or row.get('Area Index') or '').strip().upper(),
                str(row.get('cr63f_titleno') or row.get('Title No') or row.get('TITLE NO') or '').strip().upper(),
            }
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
        return baseline


    def set_window(self, window: webview.Window) -> None:
        self._window = window

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
            if not result or not result[0]:
                return {'success': False, 'cancelled': True}
            save_path = result[0]

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
        """Batch upload all currently loaded finance rows into Dataverse (cr63f_financereference)."""
        if not self.dv_client:
            self._init_dataverse()
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected. Please connect in Settings.'}
        if not self.finance_rows:
            return {'success': False, 'error': 'No finance rows loaded to sync.'}

        table_name = self.dv_client.config.get('finance_table', 'cr63f_financereference')
        success_count = 0
        error_count = 0
        errors = []
        total_rows = len(self.finance_rows)

        for i, row in enumerate(self.finance_rows):
            try:
                rec_id = row.get('__id') or None
                clean_payload = {k: v for k, v in row.items() if not k.startswith('__') and not k.startswith('_')}
                res = self.dv_client.save_finance_record(clean_payload, record_id=rec_id)
                if res.get('id') and not rec_id:
                    row['__id'] = res['id']
                success_count += 1
            except Exception as e:
                error_count += 1
                errors.append(str(e))

            if self._window and (i % 2 == 0 or i == total_rows - 1):
                try:
                    self._window.evaluate_js(f"window.onFinanceSyncProgress && window.onFinanceSyncProgress({i + 1}, {total_rows})")
                except Exception:
                    pass

        self.finance_source_name = f"Dataverse: {table_name}"

        return {
            'success': True,
            'total_rows': total_rows,
            'synced_count': success_count,
            'error_count': error_count,
            'errors': errors[:5],
            'message': f"Successfully synced {success_count} finance record(s) to Dataverse {table_name}."
        }

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
        if not result or not result[0]:
            return {'success': False, 'cancelled': True}

        file_path = result[0]
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

    def sync_conversion_from_dataverse(self, table_name: Optional[str] = None) -> dict:
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
                return {'success': False, 'error': 'Conversion record not found.'}

            if self.dv_client and found_row.get('__id'):
                try:
                    target_table = self.dv_client.config.get('conversion_table', 'cr63f_landconversionreference')
                    info = self.dv_client.get_entity_info(target_table)
                    self.dv_client.delete(info['entity_set'], found_row['__id'])
                except Exception as de:
                    print(f"Notice: Could not delete from Dataverse: {de}")

            self.conversion_rows.pop(found_idx)
            self.conversion_engine.load_rows(self.conversion_rows)

            return {
                'success': True,
                'remaining_count': len(self.conversion_rows),
                'message': 'Conversion record deleted.'
            }
        except Exception as e:
            return {'success': False, 'error': str(e)}

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
            if not result or not result[0]:
                return {'success': False, 'cancelled': True}
            save_path = result[0]

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
                ["NAS-1383", "TCT-T-56366", "748, PSD-34336", "DAR SECURED", "NIA INSPECTION DONE", "Class A", "DAR-BAT3-19", "11/23/2025", "MBHC", "NASUGBU", "LATAG", 2.25],
                ["TUY-7742", "TCT-T-121628", "2-B-4-E", "DAR FOR APPLICATION", "NIA FOR APPLICATION", "Class A", "DAR-BAT3-36", "11/23/2025", "PIC", "TUY", "MAGAHIS", 0.381],
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
        """Batch upload all currently loaded conversion rows into Dataverse (cr63f_landconversionreference)."""
        if not self.dv_client:
            self._init_dataverse()
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected. Please connect in Settings.'}
        if not self.conversion_rows:
            return {'success': False, 'error': 'No conversion rows loaded to sync.'}

        table_name = self.dv_client.config.get('conversion_table', 'cr63f_landconversionreference')
        success_count = 0
        error_count = 0
        errors = []
        total_rows = len(self.conversion_rows)

        for i, row in enumerate(self.conversion_rows):
            try:
                rec_id = row.get('__id') or None
                clean_payload = {k: v for k, v in row.items() if not k.startswith('__') and not k.startswith('_')}
                res = self.dv_client.save_conversion_record(clean_payload, record_id=rec_id)
                if res.get('id') and not rec_id:
                    row['__id'] = res['id']
                success_count += 1
            except Exception as e:
                error_count += 1
                errors.append(str(e))

            if self._window and (i % 2 == 0 or i == total_rows - 1):
                try:
                    self._window.evaluate_js(f"window.onConversionSyncProgress && window.onConversionSyncProgress({i + 1}, {total_rows})")
                except Exception:
                    pass

        self.conversion_source_name = f"Dataverse: {table_name}"

        return {
            'success': True,
            'total_rows': total_rows,
            'synced_count': success_count,
            'error_count': error_count,
            'errors': errors[:5],
            'message': f"Successfully synced {success_count} conversion record(s) to Dataverse {table_name}."
        }

    def populate_loa_records_from_references(
        self,
        records: Optional[List[Dict[str, Any]]] = None,
        persist_to_dataverse: bool = False
    ) -> dict:
        """Auto-populate financial fields and Conversion Status for LOA records using FinanceLookupEngine & ConversionLookupEngine.
        If persist_to_dataverse=True, immediately saves modified master records and writes audit diff logs to Dataverse."""
        target_records = records if records is not None else self.loa_records
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
                    title_mod_val=title_mod
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

                    reference_finance_changed = any(
                        name in fin_changed_fields for name in (
                            'cr63f_totalcontractprice',
                            'cr63f_totalpaymentdeductibletotcp',
                        )
                    )
                    if reference_finance_changed:
                        final_tcp = _decimal_amount(
                            new_r.get('cr63f_totalcontractprice')
                            if new_r.get('cr63f_totalcontractprice') is not None
                            else new_r.get('Total Contract Price') or tcp_val or curr_tcp
                        )
                        final_ded = _decimal_amount(
                            new_r.get('cr63f_totalpaymentdeductibletotcp')
                            if new_r.get('cr63f_totalpaymentdeductibletotcp') is not None
                            else new_r.get('Total Payment Deductible to TCP') or ded_val or curr_ded
                        )
                        if final_tcp is not None and final_ded is not None:
                            calculated_balance = max(Decimal(0), final_tcp - final_ded)
                            patch_fields['cr63f_balance'] = calculated_balance
                            new_r['cr63f_balance'] = calculated_balance
                            new_r['Balance'] = format_currency(float(calculated_balance))

                    # % Paid is display-only for Update from References. Never
                    # copy the reference value or write the calculated value to
                    # the Dataverse master/audit change list.
                    final_tcp_for_display = tcp_val if tcp_val is not None else curr_tcp_exact
                    final_ded_for_display = ded_val if ded_val is not None else curr_ded_exact
                    display_pct = format_calculated_percent(final_tcp_for_display, final_ded_for_display)
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

            all_changed = fin_changed_fields + conv_changed_fields
            new_r['__changed_fields'] = all_changed

            if record_matched:
                matched_count += 1
            if record_changed:
                updated_count += 1
                rec_id = new_r.get('__id')
                if rec_id and patch_fields:
                    snapshot_record = dict(new_r)
                    for percent_key in ('cr63f_paid', '% Paid', 'Paid'):
                        if percent_key in r:
                            snapshot_record[percent_key] = r[percent_key]
                        else:
                            snapshot_record.pop(percent_key, None)
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
                        previous_fields[key] = _get_field_val(r, spec[1] if spec else [key, canonical]) or None
                    items_to_save.append({
                        'id': rec_id, 'fields': patch_fields,
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
                eligible_items = []
                for item in items_to_save:
                    try:
                        baseline = self._write_baseline_if_missing(item['previous_record'], now_str)
                        if baseline:
                            saved_audit_snapshots.append(baseline)
                        eligible_items.append(item)
                    except Exception as ex:
                        save_errors.append(f"Audit baseline failed for {item.get('id')}: {ex}; Land Master was not updated.")
                items_to_save = eligible_items

                # Phase 1: Save changed master records in parallel
                if items_to_save:
                    master_workers = min(10, max(1, len(items_to_save)))
                    completed_m = 0
                    with ThreadPoolExecutor(max_workers=master_workers) as executor:
                        future_to_item = {executor.submit(self.dv_client.save_loa_record, itm['fields'], record_id=itm['id']): itm for itm in items_to_save}
                        for f in as_completed(future_to_item):
                            completed_m += 1
                            try:
                                f.result()
                                saved_items.append(future_to_item[f])
                            except Exception as ex:
                                itm = future_to_item[f]
                                save_errors.append(f"Master save error for {itm.get('id')}: {ex}")
                            if self._window:
                                try:
                                    self._window.evaluate_js(
                                        f"window.onDataverseSaveProgress && window.onDataverseSaveProgress({completed_m}, {len(items_to_save)}, 'Updating master records: {completed_m} of {len(items_to_save)}...')"
                                    )
                                except Exception:
                                    pass

                # One full source snapshot per changed account. Field-level
                # deltas are derived in memory from adjacent Dataverse rows.
                if saved_items:
                    log_workers = min(10, len(saved_items))
                    completed_l = 0
                    with ThreadPoolExecutor(max_workers=log_workers) as log_executor:
                        log_futures = {
                            log_executor.submit(self.dv_client.save_log_record, item['snapshot']): item
                            for item in saved_items
                        }
                        for lf in as_completed(log_futures):
                            completed_l += 1
                            try:
                                saved_log = lf.result()
                                snapshot = log_futures[lf]['snapshot']
                                if saved_log and saved_log.get('id'):
                                    snapshot['__id'] = saved_log['id']
                                saved_audit_snapshots.append(snapshot)
                            except Exception as ex:
                                failed_item = log_futures[lf]
                                try:
                                    self.dv_client.save_loa_record(
                                        failed_item['previous_fields'], record_id=failed_item['id']
                                    )
                                    saved_items.remove(failed_item)
                                    save_errors.append(
                                        f"Audit save failed for {failed_item['id']} ({ex}); Land Master update was rolled back."
                                    )
                                except Exception as rollback_error:
                                    save_errors.append(
                                        f"Audit save failed for {failed_item['id']} ({ex}) and Land Master rollback failed ({rollback_error}). Check Dataverse."
                                    )
                            if self._window:
                                try:
                                    self._window.evaluate_js(
                                        f"window.onDataverseSaveProgress && window.onDataverseSaveProgress({completed_l}, {len(saved_items)}, 'Writing audit snapshots: {completed_l} of {len(saved_items)}...')"
                                    )
                                except Exception:
                                    pass
                self.pending_session_logs.extend(saved_audit_snapshots)
                self.log_records = derive_field_diffs_from_snapshots(saved_audit_snapshots) + self.log_records
                saved_to_dv = True
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
            self.log_records.sort(key=lambda r: str(r.get('createdon') or r.get('cr63f_logtime') or r.get('timestamp') or r.get('log_time') or ''), reverse=True)
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

        updated_master_rec = None
        reverted = False

        if revert_master and target_log:
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

            if field_name and field_name != 'Initial Entry' and (ctrl_no or area_idx or title_no or index_no):
                # Ensure self.loa_records is loaded
                if not self.loa_records and self.dv_client:
                    try:
                        self.fetch_loa_records()
                    except Exception:
                        pass

                found_master = None
                for mr in self.loa_records:
                    m_ctrl = str(mr.get('cr63f_control1') or mr.get('cr63f_control') or mr.get('Control') or mr.get('Control Number') or mr.get('Control No') or '').strip()
                    m_aidx = str(mr.get('cr63f_areaindex') or mr.get('Area Index') or mr.get('Area-Index') or mr.get('AREA-INDEX') or mr.get('AREA INDEX') or mr.get('area_index') or '').strip()
                    m_tno = str(mr.get('cr63f_titleno') or mr.get('Title No') or mr.get('TITLE NO') or mr.get('Title No.') or mr.get('TITLE NO.') or mr.get('title_no') or mr.get('TITLE NUMBER') or '').strip()
                    m_ino = str(mr.get('cr63f_indexno') or mr.get('Index No') or mr.get('INDEX NO') or mr.get('Corridor Index') or mr.get('index_no') or '').strip()

                    is_match = False
                    if ctrl_no and m_ctrl and ctrl_no.upper() == m_ctrl.upper():
                        is_match = True
                    elif area_idx and m_aidx and area_idx.upper() == m_aidx.upper():
                        is_match = True
                    elif title_no and m_tno and (title_no.upper() == m_tno.upper() or title_no.upper() in m_tno.upper() or m_tno.upper() in title_no.upper()):
                        is_match = True
                    elif index_no and m_ino and index_no.upper() == m_ino.upper():
                        is_match = True

                    if is_match:
                        found_master = mr
                        break

                # Reverts must never guess between matching Land Master records.
                strict_matches = [mr for mr in self.loa_records if (
                    (ctrl_no and str(mr.get('cr63f_control1') or mr.get('Control') or mr.get('Control Number') or '').strip().upper() == ctrl_no.upper()) or
                    (not ctrl_no and area_idx and str(mr.get('cr63f_areaindex') or mr.get('Area-Index') or mr.get('Area Index') or '').strip().upper() == area_idx.upper())
                )]
                if len(strict_matches) != 1:
                    return {'success': False, 'error': 'Cannot safely identify exactly one Land Master record for this audit log.'}
                found_master = strict_matches[0]
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
                    expected_new = target_log.get('new_value') or target_log.get('cr63f_newvalue') or target_log.get('New Value') or ''
                    current_value = found_master.get(attr_name, found_master.get(canonical_disp, found_master.get(field_name, '')))
                    def revert_values_equal(left: Any, right: Any) -> bool:
                        # Currency/percentage values must be compared numerically:
                        # 102000, 102000.0 and ₱102,000.00 are the same value.
                        f_low_cmp = field_name.lower()
                        is_numeric = any(k in f_low_cmp for k in (
                            'price', 'deductible', 'balance', 'tcp', 'amount', 'paid', '%'
                        ))
                        if is_numeric:
                            left_num = parse_float(left)
                            right_num = parse_float(right)
                            if left_num is not None and right_num is not None:
                                return abs(left_num - right_num) <= 0.005
                        normalize = lambda value: re.sub(r'[\s,₱$%]', '', str(value or '')).lower()
                        return normalize(left) == normalize(right)

                    if not revert_values_equal(current_value, expected_new):
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
                            calc_pct = calculate_paid_percent(tcp_f, ded_f) or Decimal(0)
                            found_master['cr63f_balance'] = calc_bal
                            found_master['Balance'] = format_currency(float(calc_bal))
                            found_master['BALANCE'] = format_currency(float(calc_bal))
                            found_master['cr63f_paid'] = calc_pct
                            percent_display = format_calculated_percent(tcp_f, ded_f)
                            found_master['% Paid'] = percent_display
                            found_master['% PAID'] = percent_display
                            found_master['Paid'] = percent_display

                    # If connected to Dataverse and record has __id, update Dataverse master table
                    if self.dv_client and found_master.get('__id'):
                        try:
                            patch_payload = {attr_name: sanitized_val if sanitized_val != '' else None}
                            if 'cr63f_balance' in found_master:
                                patch_payload['cr63f_balance'] = found_master['cr63f_balance']
                            if 'cr63f_paid' in found_master:
                                patch_payload['cr63f_paid'] = found_master['cr63f_paid']
                            self.dv_client.save_loa_record(patch_payload, record_id=found_master['__id'])
                        except Exception as pex:
                            return {'success': False, 'error': f"Failed to revert record in Dataverse ({attr_name}): {pex}"}

                    updated_master_rec = dict(found_master)
                    reverted = True

                    # Revert is an auditable source operation: retain the original
                    # log, write a new REVERT row, and roll the field back if that
                    # audit row cannot be created.
                    if self.dv_client and found_master.get('__id'):
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
                            # Best-effort rollback keeps an unaudited revert from
                            # silently remaining in Dataverse.
                            try:
                                rollback_value = target_log.get('new_value') or target_log.get('cr63f_newvalue') or None
                                self.dv_client.save_loa_record({attr_name: rollback_value}, record_id=found_master['__id'])
                            except Exception as rollback_error:
                                return {'success': False, 'error': f'Audit log failed ({audit_error}); rollback also failed ({rollback_error}). Check Dataverse.'}
                            return {'success': False, 'error': f'Revert rolled back because its audit log could not be written: {audit_error}'}
                        self.pending_session_logs.append(dict(revert_log))
                        self._add_log_records([revert_log])
                        return _json_safe({'success': True, 'reverted': True, 'updated_record': updated_master_rec, 'message': 'Land Master record reverted and a REVERT audit log was written.'})

        return {'success': False, 'error': 'This audit row has no eligible Land Master field to revert.'}

    def revert_log_record(self, record_id: str) -> dict:
        """Public API for the Entry Logs Revert action."""
        return self.delete_log_record(record_id, revert_master=True)

    def revert_cell_value(self, record_id: str, col_key: str, target_value: str) -> dict:
        """Revert a single field on a Land Master record to target_value and write a REVERT log."""
        if not record_id or not col_key:
            return {'success': False, 'error': 'Missing record ID or column key.'}

        if not self.loa_records and self.dv_client:
            try:
                self.fetch_loa_records()
            except Exception:
                pass

        target_id_str = str(record_id).strip()
        found_master = None
        for mr in (self.loa_records or []):
            if str(mr.get('__id') or mr.get('id') or mr.get('cr63f_loamonitoringid') or '').strip() == target_id_str:
                found_master = mr
                break

        if not found_master:
            for mr in (self.loa_records or []):
                m_ctrl = str(mr.get('cr63f_control1') or mr.get('Control Number') or mr.get('Control') or '').strip().upper()
                m_area = str(mr.get('cr63f_areaindex') or mr.get('Area-Index') or mr.get('Area Index') or '').strip().upper()
                if target_id_str.upper() in (m_ctrl, m_area):
                    found_master = mr
                    break

        if not found_master:
            found_master = self._find_or_fetch_master_record(record_id=target_id_str)

        if not found_master:
            return {'success': False, 'error': f'Land Master record {record_id} not found.'}

        attr_name = FIELD_NAME_TO_DATAVERSE_ATTR.get(col_key, col_key)
        canonical_disp = DATAVERSE_ATTR_TO_DISPLAY.get(attr_name, col_key)
        current_val = str(found_master.get(attr_name, found_master.get(canonical_disp, found_master.get(col_key, ''))) or '')

        raw_target = str(target_value or '').strip()
        if raw_target in ('-', 'None', 'null', 'Previous', 'Created', ''):
            sanitized_val = ''
        else:
            f_low = col_key.lower()
            if any(k in f_low for k in ('price', 'deductible', 'balance', 'tcp', 'amount')):
                sanitized_val = raw_target.replace('₱', '').replace('$', '').replace(',', '').strip()
            elif 'paid' in f_low or '%' in col_key:
                sanitized_val = raw_target.replace('%', '').replace(',', '').strip()
            else:
                sanitized_val = raw_target

        found_master[attr_name] = sanitized_val
        found_master[col_key] = sanitized_val
        found_master[canonical_disp] = sanitized_val

        f_low = col_key.lower()
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

        should_recalculate = (
            'total contract price' in f_low or f_low == 'tcp' or
            'deductible' in f_low
        )
        if should_recalculate:
            tcp_f = _decimal_amount(found_master.get('cr63f_totalcontractprice') or found_master.get('Total Contract Price') or 0)
            ded_f = _decimal_amount(found_master.get('cr63f_totalpaymentdeductibletotcp') or found_master.get('Payment Deductible') or found_master.get('Total Payment Deductible to TCP') or 0)
            if tcp_f is not None and ded_f is not None:
                calc_bal = max(Decimal(0), tcp_f - ded_f)
                calc_pct = calculate_paid_percent(tcp_f, ded_f) or Decimal(0)
                found_master['cr63f_balance'] = calc_bal
                found_master['Balance'] = format_currency(float(calc_bal))
                found_master['BALANCE'] = format_currency(float(calc_bal))
                found_master['cr63f_paid'] = calc_pct
                percent_display = format_calculated_percent(tcp_f, ded_f)
                found_master['% Paid'] = percent_display
                found_master['% PAID'] = percent_display
                found_master['Paid'] = percent_display

        if self.dv_client and found_master.get('__id'):
            try:
                patch_payload = {attr_name: sanitized_val if sanitized_val != '' else None}
                if 'cr63f_balance' in found_master:
                    patch_payload['cr63f_balance'] = found_master['cr63f_balance']
                if 'cr63f_paid' in found_master:
                    patch_payload['cr63f_paid'] = found_master['cr63f_paid']
                self.dv_client.save_loa_record(patch_payload, record_id=found_master['__id'])
            except Exception as pex:
                return {'success': False, 'error': f"Failed to revert record in Dataverse ({attr_name}): {pex}"}

        area_idx = str(found_master.get('cr63f_areaindex') or found_master.get('Area-Index') or '').strip()
        index_no = str(found_master.get('cr63f_indexno') or found_master.get('Index No') or '').strip()
        title_no = str(found_master.get('cr63f_titleno') or found_master.get('Title No') or '').strip()

        revert_log = build_log_entry(
            action_type='REVERT',
            work_week=get_work_week(datetime.now()),
            now_str=datetime.now().strftime('%m/%d/%Y %I:%M %p'),
            area_index=area_idx,
            index_no=index_no,
            title_no=title_no,
            field_name=canonical_disp,
            old_value=current_val,
            new_value=sanitized_val,
            details=f"Reverted {canonical_disp} to {sanitized_val}",
            full_record=found_master
        )

        if self.dv_client:
            try:
                saved_log = self.dv_client.save_log_record(revert_log)
                revert_log['__id'] = saved_log.get('id', '')
            except Exception:
                pass

        self.pending_session_logs.append(dict(revert_log))
        self._add_log_records([revert_log])

        return _json_safe({
            'success': True,
            'reverted': True,
            'updated_record': dict(found_master),
            'message': f'{canonical_disp} reverted successfully to {sanitized_val or "(blank)"}.'
        })

    def _find_or_fetch_master_record(
        self,
        area_index: Optional[str] = None,
        title_no: Optional[str] = None,
        index_no: Optional[str] = None,
        record_id: Optional[str] = None,
        control_no: Optional[str] = None
    ) -> Dict[str, Any]:
        """Find a master Land Master DB record by id, area_index, title_no, index_no, or control_no from memory or Dataverse."""
        norm_area = str(area_index or '').strip().upper()
        norm_title = str(title_no or '').strip().upper()
        norm_idx = str(index_no or '').strip().upper()
        norm_id = str(record_id or '').strip()
        norm_ctrl = str(control_no or '').strip().upper()

        # 1. Search in-memory loa_records
        for r in (self.loa_records or []):
            if not isinstance(r, dict):
                continue
            if norm_id and str(r.get('__id') or '').strip() == norm_id:
                return dict(r)

            r_area = str(r.get('cr63f_areaindex') or r.get('Area-Index') or r.get('Area Index') or r.get('AREA-INDEX') or '').strip().upper()
            if norm_area and r_area and r_area == norm_area:
                return dict(r)

            r_title = str(r.get('cr63f_titleno') or r.get('Title No') or r.get('Title Number') or r.get('TITLE NO') or '').strip().upper()
            if norm_title and r_title and r_title == norm_title:
                return dict(r)

            r_idx = str(r.get('cr63f_indexno') or r.get('Index No') or r.get('INDEX NO') or '').strip().upper()
            if norm_idx and r_idx and r_idx == norm_idx:
                return dict(r)

            r_ctrl = str(r.get('cr63f_control1') or r.get('Control') or r.get('Control Number') or r.get('cr63f_control') or '').strip().upper()
            if norm_ctrl and r_ctrl and r_ctrl == norm_ctrl:
                return dict(r)

        # 2. Query Dataverse cr63f_loamonitoring master table if dv_client is available
        if self.dv_client:
            try:
                table_name = self.dv_client.config.get('loa_table', 'cr63f_loamonitoring')
                filters = []
                if norm_area:
                    filters.append(f"cr63f_areaindex eq '{norm_area}'")
                if norm_title:
                    filters.append(f"cr63f_titleno eq '{norm_title}'")
                if filters:
                    filter_str = ' or '.join(filters)
                    dv_records = self.dv_client.fetch_records(table_name, max_records=1, custom_only=True, filter_query=filter_str)
                    if dv_records and len(dv_records) > 0:
                        return dict(dv_records[0])
            except Exception:
                pass

        return {}

    def save_loa_record_entry(self, record_data: dict, record_id: Optional[str] = None) -> dict:
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
                merged_rec = dict(found_rec)
                merged_rec.update(record_data)

                # The hidden pre-change snapshot lets this first manual update
                # be compared after the app restarts.
                baseline = self._write_baseline_if_missing(found_rec, now_str)

                res = self.dv_client.save_loa_record(record_data, record_id=record_id)

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
                            key: _get_field_val(found_rec, [key, CANONICAL_FIELD_MAP.get(key, key)]) or None
                            for key in record_data
                            if not key.startswith('__')
                        }
                        self.dv_client.save_loa_record(rollback_fields, record_id=record_id)
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
                        break

                msg = f"Record successfully updated in Dataverse master table ({master_table})."
            else:
                # 1. Save master record in cr63f_loamonitoring
                res_master = self.dv_client.save_loa_record(record_data)
                created_id = res_master.get('id', '')
                new_rec = dict(record_data)
                new_rec['__id'] = created_id
                new_rec['work_week'] = ww
                new_rec['Work Week'] = ww
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
                    return {
                        'success': False,
                        'id': created_id,
                        'error': f'Land Master account was created, but its Initial Entry audit row could not be saved: {audit_error}',
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
            merged_rec = dict(orig_rec)
            merged_rec.update(fields)

            area_idx = orig_rec.get('cr63f_areaindex') or orig_rec.get('Area-Index') or fields.get('cr63f_areaindex') or fields.get('Area-Index') or rec_id
            index_no = orig_rec.get('cr63f_indexno') or orig_rec.get('Index No') or fields.get('cr63f_indexno') or fields.get('Index No') or ''
            title_no = orig_rec.get('cr63f_titleno') or orig_rec.get('Title No') or fields.get('cr63f_titleno') or fields.get('Title No') or ''
            # Audit work week is based on the log/save time, not Endorsed Date.
            ww = get_work_week(datetime.now())

            try:
                baseline = self._write_baseline_if_missing(orig_rec, now_str)
                if baseline:
                    with lock:
                        baseline_snapshots.append(baseline)

                self.dv_client.save_loa_record(fields, record_id=rec_id)
                with lock:
                    saved_count += 1

                # Dataverse stores one full snapshot per saved account. Entry Logs
                # expands that snapshot into field-level comparisons in memory.
                changed_fields = []
                change_logs = []
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
                        act_type = item.get('action_type') or 'MANUAL_EDIT'
                        change_logs.append(build_log_entry(
                            action_type=act_type, work_week=ww,
                            now_str=now_str, area_index=area_idx,
                            index_no=index_no, title_no=title_no,
                            field_name=canon_name, old_value=old_v,
                            new_value=updated_v
                        ))

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
                            r['work_week'] = ww
                            r['Work Week'] = ww
                            break
                    if item_logs:
                        log_entries_to_save.extend(item_logs)
                        for log_entry in item_logs:
                            audit_contexts[id(log_entry)] = (rec_id, orig_rec, fields)

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
                        rec_id, original, changed = audit_contexts[id(log_futures[lf])]
                        rollback_fields = {}
                        for key in changed:
                            canonical = CANONICAL_FIELD_MAP.get(key, key)
                            spec = next((entry for entry in LOG_COMPARE_SPECS if entry[0] == canonical), None)
                            rollback_fields[key] = _get_field_val(original, spec[1] if spec else [key, canonical]) or None
                        try:
                            self.dv_client.save_loa_record(rollback_fields, record_id=rec_id)
                            with lock:
                                saved_count -= 1
                                for row in self.loa_records:
                                    if row.get('__id') == rec_id:
                                        row.clear()
                                        row.update(original)
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
        """Save staged finance row updates into Dataverse finance table and write change logs."""
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        finance_table = self.dv_client.config.get('finance_table', 'cr63f_financereference')
        logs_table = self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs')

        total_items = len(staged_list)
        if total_items == 0:
            return {'success': True, 'saved_count': 0, 'message': 'No changes to save.'}

        lock = threading.Lock()
        saved_count = 0
        errors = []
        now_str = datetime.now().strftime('%m/%d/%Y %I:%M %p')
        ww = get_work_week(datetime.now())
        log_entries_to_save = []

        def save_single_finance(item):
            nonlocal saved_count
            rec_id = item.get('id')
            fields = item.get('fields', {})
            if rec_id is None or not fields:
                return

            with lock:
                found_rec = next((r for r in self.finance_rows if str(r.get('__id') or '') == str(rec_id)), {})
                if not found_rec:
                    try:
                        idx = int(rec_id)
                        if 0 <= idx < len(self.finance_rows):
                            found_rec = self.finance_rows[idx]
                    except (ValueError, TypeError):
                        pass
                orig_rec = dict(found_rec)

            merged_rec = dict(orig_rec)
            merged_rec.update(fields)

            area_idx = orig_rec.get('Area Index') or orig_rec.get('Corridor Index') or orig_rec.get('cr63f_areaindex') or fields.get('Area Index') or fields.get('Corridor Index') or str(rec_id)
            index_no = orig_rec.get('Corridor Index') or orig_rec.get('INDEX NO') or orig_rec.get('Index No') or fields.get('Corridor Index') or ''
            title_no = orig_rec.get('Title Mod') or orig_rec.get('Title No') or fields.get('Title Mod') or ''

            # Look up master Land Master DB record to create full-row snapshot
            master_rec = self._find_or_fetch_master_record(area_index=area_idx, title_no=title_no, index_no=index_no)
            log_snapshot = dict(master_rec)
            log_snapshot.update(merged_rec)
            if not index_no and master_rec:
                index_no = master_rec.get('cr63f_indexno') or master_rec.get('Index No') or ''
            if not title_no and master_rec:
                title_no = master_rec.get('cr63f_titleno') or master_rec.get('Title No') or ''

            try:
                dv_id = orig_rec.get('__id')
                if dv_id:
                    self.dv_client.save_finance_record(fields, record_id=str(dv_id))
                else:
                    res = self.dv_client.save_finance_record(merged_rec)
                    if res and res.get('id'):
                        merged_rec['__id'] = res['id']

                with lock:
                    saved_count += 1

                item_logs = []
                for f_key, new_v in fields.items():
                    if f_key.startswith('__'):
                        continue
                    old_v = orig_rec.get(f_key, '')
                    if str(old_v).strip() != str(new_v).strip():
                        log_entry = build_log_entry(
                            action_type='FINANCE_UPDATE',
                            work_week=ww,
                            now_str=now_str,
                            area_index=area_idx,
                            index_no=index_no,
                            title_no=title_no,
                            field_name=f_key,
                            old_value=old_v,
                            new_value=new_v,
                            details=f"Updated Finance Reference {f_key} from {old_v} to {new_v}",
                            full_record=log_snapshot
                        )
                        item_logs.append(log_entry)

                with lock:
                    for r in self.finance_rows:
                        if (r.get('__id') and str(r.get('__id')) == str(rec_id)) or r == orig_rec:
                            r.update(fields)
                            if merged_rec.get('__id'):
                                r['__id'] = merged_rec['__id']
                            break
                    if item_logs:
                        log_entries_to_save.extend(item_logs)

            except Exception as e:
                with lock:
                    errors.append(f"Finance Row {rec_id}: {e}")

        max_workers = min(10, max(1, total_items))
        completed_records = 0
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(save_single_finance, itm) for itm in staged_list]
            for future in as_completed(futures):
                with lock:
                    completed_records += 1
                    cur_prog = completed_records
                if self._window:
                    try:
                        self._window.evaluate_js(
                            f"window.onDataverseSaveProgress && window.onDataverseSaveProgress({cur_prog}, {total_items}, 'Saving finance reference records: {cur_prog} of {total_items}...')"
                        )
                    except Exception:
                        pass

        if log_entries_to_save:
            total_logs = len(log_entries_to_save)
            completed_logs = 0
            log_workers = min(10, max(1, total_logs))

            def save_single_log(log_entry):
                try:
                    self.dv_client.save_log_record(log_entry)
                except Exception:
                    pass

            with ThreadPoolExecutor(max_workers=log_workers) as log_executor:
                log_futures = [log_executor.submit(save_single_log, le) for le in log_entries_to_save]
                for lf in as_completed(log_futures):
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
                self._add_log_records(log_entries_to_save)

        self.finance_engine.load_rows(self.finance_rows)

        if errors and saved_count == 0:
            return {'success': False, 'error': '; '.join(errors[:3])}

        return {
            'success': True,
            'saved_count': saved_count,
            'message': f"Saved {saved_count} finance record(s) to Dataverse ({finance_table}) and logged {len(log_entries_to_save)} changes to {logs_table}."
        }

    def save_staged_conversion_records(self, staged_list: List[Dict[str, Any]]) -> dict:
        """Save staged conversion row updates into Dataverse conversion table and write change logs."""
        if not self.dv_client:
            return {'success': False, 'error': 'Dataverse is not connected.'}
        conv_table = self.dv_client.config.get('conversion_table', 'cr63f_landconversionreference')
        logs_table = self.dv_client.config.get('logs_table', 'cr63f_loamonitoringlogs')

        total_items = len(staged_list)
        if total_items == 0:
            return {'success': True, 'saved_count': 0, 'message': 'No changes to save.'}

        lock = threading.Lock()
        saved_count = 0
        errors = []
        now_str = datetime.now().strftime('%m/%d/%Y %I:%M %p')
        ww = get_work_week(datetime.now())
        log_entries_to_save = []

        def save_single_conversion(item):
            nonlocal saved_count
            rec_id = item.get('id')
            fields = item.get('fields', {})
            if rec_id is None or not fields:
                return

            with lock:
                found_rec = next((r for r in self.conversion_rows if str(r.get('__id') or '') == str(rec_id)), {})
                if not found_rec:
                    try:
                        idx = int(rec_id)
                        if 0 <= idx < len(self.conversion_rows):
                            found_rec = self.conversion_rows[idx]
                    except (ValueError, TypeError):
                        pass
                orig_rec = dict(found_rec)

            merged_rec = dict(orig_rec)
            merged_rec.update(fields)

            area_idx = orig_rec.get('cr63f_areaindex') or orig_rec.get('AREA-INDEX') or orig_rec.get('Area Index') or fields.get('cr63f_areaindex') or fields.get('AREA-INDEX') or str(rec_id)
            title_no = orig_rec.get('cr63f_titlenumber') or orig_rec.get('TITLE NUMBER') or orig_rec.get('Title No') or fields.get('cr63f_titlenumber') or fields.get('TITLE NUMBER') or ''

            # Look up master Land Master DB record to create full-row snapshot
            master_rec = self._find_or_fetch_master_record(area_index=area_idx, title_no=title_no)
            log_snapshot = dict(master_rec)
            log_snapshot.update(merged_rec)
            if not title_no and master_rec:
                title_no = master_rec.get('cr63f_titleno') or master_rec.get('Title No') or ''
            log_index_no = master_rec.get('cr63f_indexno') or master_rec.get('Index No') or ''

            try:
                dv_id = orig_rec.get('__id')
                if dv_id:
                    self.dv_client.save_conversion_record(fields, record_id=str(dv_id))
                else:
                    res = self.dv_client.save_conversion_record(merged_rec)
                    if res and res.get('id'):
                        merged_rec['__id'] = res['id']

                with lock:
                    saved_count += 1

                item_logs = []
                for f_key, new_v in fields.items():
                    if f_key.startswith('__'):
                        continue
                    old_v = orig_rec.get(f_key, '')
                    if str(old_v).strip() != str(new_v).strip():
                        log_entry = build_log_entry(
                            action_type='CONVERSION_UPDATE',
                            work_week=ww,
                            now_str=now_str,
                            area_index=area_idx,
                            index_no=log_index_no,
                            title_no=title_no,
                            field_name=f_key,
                            old_value=old_v,
                            new_value=new_v,
                            details=f"Updated Conversion Reference {f_key} from {old_v} to {new_v}",
                            full_record=log_snapshot
                        )
                        item_logs.append(log_entry)

                with lock:
                    for r in self.conversion_rows:
                        if (r.get('__id') and str(r.get('__id')) == str(rec_id)) or r == orig_rec:
                            r.update(fields)
                            if merged_rec.get('__id'):
                                r['__id'] = merged_rec['__id']
                            break
                    if item_logs:
                        log_entries_to_save.extend(item_logs)

            except Exception as e:
                with lock:
                    errors.append(f"Conversion Row {rec_id}: {e}")

        max_workers = min(10, max(1, total_items))
        completed_records = 0
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(save_single_conversion, itm) for itm in staged_list]
            for future in as_completed(futures):
                with lock:
                    completed_records += 1
                    cur_prog = completed_records
                if self._window:
                    try:
                        self._window.evaluate_js(
                            f"window.onDataverseSaveProgress && window.onDataverseSaveProgress({cur_prog}, {total_items}, 'Saving conversion reference records: {cur_prog} of {total_items}...')"
                        )
                    except Exception:
                        pass

        if log_entries_to_save:
            total_logs = len(log_entries_to_save)
            completed_logs = 0
            log_workers = min(10, max(1, total_logs))

            def save_single_log(log_entry):
                try:
                    self.dv_client.save_log_record(log_entry)
                except Exception:
                    pass

            with ThreadPoolExecutor(max_workers=log_workers) as log_executor:
                log_futures = [log_executor.submit(save_single_log, le) for le in log_entries_to_save]
                for lf in as_completed(log_futures):
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
                self._add_log_records(log_entries_to_save)

        self.conversion_engine.load_rows(self.conversion_rows)

        if errors and saved_count == 0:
            return {'success': False, 'error': '; '.join(errors[:3])}

        return {
            'success': True,
            'saved_count': saved_count,
            'message': f"Saved {saved_count} conversion record(s) to Dataverse ({conv_table}) and logged {len(log_entries_to_save)} changes to {logs_table}."
        }

    def revert_staged_finance_records(self, baseline_rows: List[Dict[str, Any]]) -> dict:
        """Revert in-memory finance records and engine back to baseline."""
        self.finance_rows = list(baseline_rows or [])
        self.finance_engine.load_rows(self.finance_rows)
        return {'success': True}

    def revert_staged_conversion_records(self, baseline_rows: List[Dict[str, Any]]) -> dict:
        """Revert in-memory conversion records and engine back to baseline."""
        self.conversion_rows = list(baseline_rows or [])
        self.conversion_engine.load_rows(self.conversion_rows)
        return {'success': True}

    def get_dashboard_analytics(self, work_week: Optional[str] = None, mode: Optional[str] = None) -> dict:
        """Compute weekly intelligence metrics, mode distribution, conversion analytics, and change activity."""
        all_records = self.loa_records or []
        logs = self.log_records or []
        conv_rows = self.conversion_rows or []

        selected_mode_input = (mode or '').strip().upper()
        if selected_mode_input and selected_mode_input != 'ALL':
            records = [r for r in all_records if str(r.get('new_mode') or r.get('Mode') or r.get('MODE') or '').strip().upper() == selected_mode_input]
            active_mode = selected_mode_input
        else:
            records = all_records
            active_mode = None

        def log_datetime(log: Dict[str, Any]) -> Optional[datetime]:
            raw = str(log.get('createdon') or log.get('cr63f_logtime') or log.get('timestamp') or log.get('log_time') or log.get('Log Time') or '').strip()
            if not raw:
                return None
            try:
                return datetime.fromisoformat(raw.replace('Z', '+00:00')).replace(tzinfo=None)
            except ValueError:
                pass
            for fmt in ('%m/%d/%Y %I:%M %p', '%m/%d/%Y %H:%M', '%Y-%m-%d %H:%M:%S'):
                try:
                    return datetime.strptime(raw, fmt)
                except ValueError:
                    continue
            return None

        def log_week_label(log: Dict[str, Any]) -> str:
            dt = log_datetime(log)
            if dt:
                return f"{dt.year} · {get_work_week(dt)}"
            ww = str(log.get('work_week') or log.get('Work Week') or '').strip()
            return ww

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
            for w in available_ww:
                if selected_ww.upper() in w.upper() or w.upper() in selected_ww.upper():
                    selected_ww = w
                    break
        filtered_logs = logs
        if selected_ww and selected_ww.upper() != 'ALL':
            sel_u = selected_ww.upper()
            filtered_logs = [
                l for l in filtered_logs
                if log_week_label(l) == selected_ww
                or sel_u in log_week_label(l).upper()
                or log_week_label(l).upper() in sel_u
                or str(l.get('work_week') or l.get('Work Week') or '').strip().upper() == sel_u
            ]
        if active_mode:
            filtered_logs = [l for l in filtered_logs if str(l.get('new_mode') or l.get('Mode') or l.get('MODE') or '').strip().upper() == active_mode or active_mode in str(l.get('details', '')).upper()]

        total_tcp = sum(parse_float(r.get('Total Contract Price', r.get('cr63f_totalcontractprice', 0))) or 0 for r in records)
        total_ded = sum(parse_float(r.get('Total Payment Deductible to TCP', r.get('cr63f_totalpaymentdeductibletotcp', 0))) or 0 for r in records)
        total_bal = max(0.0, total_tcp - total_ded)
        overall_paid_pct = ((total_ded / total_tcp) * 100) if total_tcp > 0 else (100.0 if total_ded > 0 else 0.0)

        # Mode distribution breakdown (calculated against all_records)
        mode_counts = {}
        for r in all_records:
            m = str(r.get('new_mode') or r.get('Mode') or r.get('MODE') or 'TO BE IDENTIFIED').strip().upper()
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

        # Financial Settlement Breakdown (for records)
        rec_count = len(records)
        fully_paid_count = 0
        partially_paid_count = 0
        zero_paid_count = 0

        for r in records:
            tcp_v = parse_float(r.get('Total Contract Price', r.get('cr63f_totalcontractprice', 0))) or 0.0
            ded_v = parse_float(r.get('Total Payment Deductible to TCP', r.get('cr63f_totalpaymentdeductibletotcp', 0))) or 0.0
            bal_v = max(0.0, tcp_v - ded_v)
            if tcp_v > 0 and bal_v <= 0.01:
                fully_paid_count += 1
            elif ded_v > 0 and ded_v < tcp_v:
                partially_paid_count += 1
            else:
                zero_paid_count += 1

        avg_tcp = (total_tcp / rec_count) if rec_count > 0 else 0.0
        avg_ded = (total_ded / rec_count) if rec_count > 0 else 0.0

        financial_breakdown = {
            'fully_paid_count': fully_paid_count,
            'fully_paid_percent': f"{(fully_paid_count / (rec_count or 1) * 100):.1f}%",
            'fully_paid_percent_num': round(fully_paid_count / (rec_count or 1) * 100, 1),
            'partially_paid_count': partially_paid_count,
            'partially_paid_percent': f"{(partially_paid_count / (rec_count or 1) * 100):.1f}%",
            'partially_paid_percent_num': round(partially_paid_count / (rec_count or 1) * 100, 1),
            'zero_paid_count': zero_paid_count,
            'zero_paid_percent': f"{(zero_paid_count / (rec_count or 1) * 100):.1f}%",
            'zero_paid_percent_num': round(zero_paid_count / (rec_count or 1) * 100, 1),
            'avg_tcp_formatted': f"₱{format_currency(avg_tcp)}",
            'avg_paid_formatted': f"₱{format_currency(avg_ded)}",
        }

        # Conversion Reference Analytics (filtered by active records)
        conv_status_counts = {}
        total_applied_area = 0.0
        secured_count = 0
        in_progress_count = 0

        if records:
            for r in records:
                c_status = str(r.get('cr63f_conversionstatus') or r.get('Conversion Status') or '').strip().upper()
                area_val = parse_float(r.get('cr63f_lotareaha') or r.get('Lot Area (Ha)') or r.get('cr63f_areaappliedhas') or 0.0)

                if (not c_status or c_status in ('UNKNOWN', '-', '')) and hasattr(self, 'conversion_engine') and self.conversion_engine:
                    area_idx = r.get('cr63f_areaindex') or r.get('Area-Index') or r.get('Area Index')
                    lk = self.conversion_engine.lookup(area_idx)
                    if lk.get('matched'):
                        c_status = str(lk.get('conversion_status') or lk.get('dar_status') or '').strip().upper()
                        if not area_val:
                            lk_row = lk.get('matched_row') or {}
                            area_val = parse_float(lk_row.get('AREA APPLIED (HAS.)') or lk_row.get('cr63f_areaappliedhas') or 0.0)

                if not c_status or c_status in ('UNKNOWN', '-', ''):
                    c_status = 'NOT YET ENDORSED'

                conv_status_counts[c_status] = conv_status_counts.get(c_status, 0) + 1
                is_secured = any(k in c_status for k in ('SECURED', 'EXEMPT', 'APPROVED', 'COMPLETED', 'CONVERTED')) and 'INSPECTION' not in c_status
                if is_secured:
                    secured_count += 1
                elif c_status not in ('NOT YET ENDORSED', 'UNSET', 'NONE', 'N/A'):
                    in_progress_count += 1

                if area_val:
                    total_applied_area += area_val

            total_conv_lots = rec_count
        else:
            source_conv = conv_rows
            for cr in source_conv:
                c_status = str(cr.get('DAR STATUS') or cr.get('DAR Status') or cr.get('cr63f_conversionstatus') or cr.get('cr63f_darstatus') or 'UNKNOWN').strip().upper()
                if c_status and c_status not in ('UNKNOWN', '-', ''):
                    conv_status_counts[c_status] = conv_status_counts.get(c_status, 0) + 1
                    is_secured = any(k in c_status for k in ('SECURED', 'EXEMPT', 'APPROVED', 'COMPLETED', 'CONVERTED')) and 'INSPECTION' not in c_status
                    if is_secured:
                        secured_count += 1
                    elif c_status not in ('NOT YET ENDORSED', 'UNSET', 'NONE', 'N/A'):
                        in_progress_count += 1

                area_val = parse_float(cr.get('AREA APPLIED (HAS.)') or cr.get('Area Applied (Has.)') or cr.get('cr63f_areaappliedhas') or cr.get('cr63f_area') or 0)
                if area_val:
                    total_applied_area += area_val
            total_conv_lots = len(source_conv)

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
        selected_pos = available_ww.index(selected_ww) if selected_ww in available_ww else -1
        previous_ww = available_ww[selected_pos + 1] if selected_pos >= 0 and selected_pos + 1 < len(available_ww) else ''

        def audit_value(log: Dict[str, Any]) -> Any:
            return log.get('new_value', log.get('cr63f_newvalue', log.get('New Value', '')))

        def account_key(log: Dict[str, Any]) -> str:
            return str(log.get('cr63f_control1') or log.get('cr63f_control') or log.get('Control') or log.get('Control Number') or log.get('area_index') or log.get('cr63f_areaindex') or log.get('title_no') or log.get('cr63f_titleno') or '').strip().upper()

        def field_key(log: Dict[str, Any]) -> str:
            return str(log.get('field_name') or log.get('cr63f_fieldchanged') or log.get('Field Changed') or '').strip()

        def match_master_rec(log: Dict[str, Any]) -> Optional[Dict[str, Any]]:
            ctrl_no = str(log.get('cr63f_control1') or log.get('cr63f_control') or log.get('Control') or log.get('Control Number') or '').strip()
            area_idx = str(log.get('area_index') or log.get('cr63f_areaindex') or log.get('Area-Index') or '').strip()
            title_no = str(log.get('title_no') or log.get('cr63f_titleno') or log.get('Title No') or '').strip()
            if not ctrl_no and not area_idx and not title_no:
                return None
            for mr in all_records:
                if ctrl_no and str(mr.get('Control Number') or mr.get('cr63f_control1') or mr.get('Control') or '').strip() == ctrl_no:
                    return mr
                if area_idx and str(mr.get('Area-Index') or mr.get('cr63f_areaindex') or mr.get('Area Index') or '').strip() == area_idx:
                    return mr
                if title_no and str(mr.get('Title No') or mr.get('cr63f_titleno') or mr.get('Title') or '').strip() == title_no:
                    return mr
            return None

        def latest_values(week: str) -> Dict[Tuple[str, str], Dict[str, Any]]:
            found: Dict[Tuple[str, str], Dict[str, Any]] = {}
            for log in logs:
                if log_week_label(log) != week:
                    continue
                key = (account_key(log), field_key(log))
                if not all(key) or key[1].lower() == 'initial entry':
                    continue
                if key not in found or (log_datetime(log) or datetime.min) >= (log_datetime(found[key]) or datetime.min):
                    found[key] = log
            return found

        current_values = latest_values(selected_ww) if selected_ww and previous_ww else {}
        previous_values = latest_values(previous_ww) if previous_ww else {}
        grouped_changes: Dict[str, Dict[str, Any]] = {}

        # 1. First attempt: Week-over-week delta comparison
        if current_values and previous_values:
            for key, current_log in current_values.items():
                previous_log = previous_values.get(key)
                if not previous_log:
                    continue
                old_value, new_value = audit_value(previous_log), audit_value(current_log)
                if _norm_text_str(old_value) == _norm_text_str(new_value):
                    continue

                matched_rec = match_master_rec(current_log)
                if active_mode:
                    acct_m = ''
                    if matched_rec:
                        acct_m = str(matched_rec.get('new_mode') or matched_rec.get('Mode') or matched_rec.get('MODE') or '').strip().upper()
                    if not acct_m:
                        acct_m = str(current_log.get('new_mode') or current_log.get('Mode') or current_log.get('MODE') or '').strip().upper()
                    if acct_m and acct_m != active_mode:
                        continue

                acct = key[0]
                group = grouped_changes.setdefault(acct, {
                    'control_number': str(current_log.get('cr63f_control1') or current_log.get('cr63f_control') or current_log.get('Control') or current_log.get('Control Number') or ''),
                    'area_index': str(current_log.get('area_index') or current_log.get('cr63f_areaindex') or current_log.get('Area-Index') or ''),
                    'title_no': str(current_log.get('title_no') or current_log.get('cr63f_titleno') or current_log.get('Title No') or ''),
                    'current_week': selected_ww, 'previous_week': previous_ww, 'changes': []
                })

                tcp_raw = current_log.get('cr63f_totalcontractprice') or current_log.get('Total Contract Price')
                if matched_rec and not tcp_raw:
                    tcp_raw = matched_rec.get('Total Contract Price') or matched_rec.get('cr63f_totalcontractprice')

                sublevel_raw = current_log.get('cr63f_acquisitionsublevel') or current_log.get('Acquisition Sub-Level')
                if matched_rec and not sublevel_raw:
                    sublevel_raw = matched_rec.get('Acquisition Sub-Level') or matched_rec.get('cr63f_acquisitionsublevel') or matched_rec.get('sub_level')

                updated_by_raw = current_log.get('createdby') or current_log.get('author') or current_log.get('modifiedby') or current_log.get('cr63f_createdby') or current_log.get('User') or 'j.inguito'
                log_time_val = current_log.get('log_time') or current_log.get('createdon') or current_log.get('timestamp') or ''

                group['changes'].append({
                    'field_name': key[1],
                    'old_value': old_value,
                    'new_value': new_value,
                    'action_type': current_log.get('action_type') or current_log.get('Action Type') or 'UPDATE',
                    'log_time': log_time_val,
                    'contract_price': f"₱{format_currency(parse_float(tcp_raw))}" if tcp_raw else '—',
                    'acquisition_sublevel': str(sublevel_raw or '—'),
                    'updated_by': str(updated_by_raw or '—')
                })

        # 2. If week-over-week deltas yielded no changes (or single week / ALL selected),
        # populate grouped_changes directly from candidate log records
        if not grouped_changes and filtered_logs:
            sorted_candidate_logs = sorted(
                filtered_logs,
                key=lambda l: log_datetime(l) or datetime.min,
                reverse=True
            )
            for current_log in sorted_candidate_logs[:200]:
                fld = field_key(current_log)
                if not fld or fld.lower() == 'initial entry':
                    continue
                acct = account_key(current_log) or 'ACCOUNT'
                matched_rec = match_master_rec(current_log)
                if active_mode:
                    acct_m = ''
                    if matched_rec:
                        acct_m = str(matched_rec.get('new_mode') or matched_rec.get('Mode') or matched_rec.get('MODE') or '').strip().upper()
                    if not acct_m:
                        acct_m = str(current_log.get('new_mode') or current_log.get('Mode') or current_log.get('MODE') or '').strip().upper()
                    if acct_m and acct_m != active_mode:
                        continue

                old_val = current_log.get('old_value', current_log.get('cr63f_previousvalue', current_log.get('Previous Value', '—')))
                new_val = current_log.get('new_value', current_log.get('cr63f_newvalue', current_log.get('New Value', '')))
                if _norm_text_str(old_val) == _norm_text_str(new_val) and old_val not in (None, '', '—'):
                    continue

                group = grouped_changes.setdefault(acct, {
                    'control_number': str(current_log.get('cr63f_control1') or current_log.get('cr63f_control') or current_log.get('Control') or current_log.get('Control Number') or ''),
                    'area_index': str(current_log.get('area_index') or current_log.get('cr63f_areaindex') or current_log.get('Area-Index') or ''),
                    'title_no': str(current_log.get('title_no') or current_log.get('cr63f_titleno') or current_log.get('Title No') or ''),
                    'current_week': log_week_label(current_log) or selected_ww,
                    'previous_week': previous_ww,
                    'changes': []
                })

                tcp_raw = current_log.get('cr63f_totalcontractprice') or current_log.get('Total Contract Price')
                if matched_rec and not tcp_raw:
                    tcp_raw = matched_rec.get('Total Contract Price') or matched_rec.get('cr63f_totalcontractprice')

                sublevel_raw = current_log.get('cr63f_acquisitionsublevel') or current_log.get('Acquisition Sub-Level')
                if matched_rec and not sublevel_raw:
                    sublevel_raw = matched_rec.get('Acquisition Sub-Level') or matched_rec.get('cr63f_acquisitionsublevel') or matched_rec.get('sub_level')

                updated_by_raw = current_log.get('createdby') or current_log.get('author') or current_log.get('modifiedby') or current_log.get('cr63f_createdby') or current_log.get('User') or 'j.inguito'
                log_time_val = current_log.get('log_time') or current_log.get('createdon') or current_log.get('timestamp') or ''
                act_type = current_log.get('action_type') or current_log.get('Action Type') or ('ADDED' if not old_val or old_val == '—' else 'UPDATED')

                group['changes'].append({
                    'field_name': fld,
                    'old_value': old_val,
                    'new_value': new_val,
                    'action_type': act_type,
                    'log_time': log_time_val,
                    'contract_price': f"₱{format_currency(parse_float(tcp_raw))}" if tcp_raw else '—',
                    'acquisition_sublevel': str(sublevel_raw or '—'),
                    'updated_by': str(updated_by_raw or '—')
                })

        recent_activity = list(grouped_changes.values())
        total_period_changes = sum(len(g.get('changes', [])) for g in recent_activity)

        return {
            'success': True,
            'selected_work_week': selected_ww,
            'selected_mode': active_mode,
            'available_work_weeks': available_ww,
            'total_records': len(records),
            'all_records_count': len(all_records),
            'weekly_changes_count': total_period_changes,
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
            'recent_activity': recent_activity
        }

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
        if not result or not result[0]:
            return {'success': False, 'cancelled': True}
        save_path = result[0]
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
                    cell = ws.cell(row=row_idx, column=col_idx, value=str(val))
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
            'conversion_table': cfg.get('conversion_table', 'cr63f_landconversionreference'),
            'logs_table': cfg.get('logs_table', 'cr63f_loamonitoringlogs'),
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
                'conversion_table': new_settings.get('conversion_table', 'cr63f_landconversionreference'),
                'logs_table': new_settings.get('logs_table', 'cr63f_loamonitoringlogs'),
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
        title='',
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

