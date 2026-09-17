"""LOA Monitoring Engine.

Pure-Python computation and lookup engine implementing the Excel lookup formula:
  =LET(
      Result,
          IF(
              [@TYPE]="ROW",
              XLOOKUP(
                  [@[INDEX NO]],
                  FILTER(APPENDED_FINANCE[Corridor Index], APPENDED_FINANCE[TYPE]="ROW"),
                  FILTER(APPENDED_FINANCE[% Paid], APPENDED_FINANCE[TYPE]="ROW"),
                  ""
              ),
              XLOOKUP(
                  [@[AREA-INDEX]],
                  APPENDED_FINANCE[Area Index],
                  APPENDED_FINANCE[% Paid],
                  ""
              )
          ),
      IF(
          Result <> "",
          Result,
          IF(
              [@[TITLE MOD]] <> "",
              XLOOKUP(
                  [@[TITLE MOD]],
                  APPENDED_FINANCE[Title Mod],
                  APPENDED_FINANCE[% Paid],
                  ""
              ),
              ""
          )
      )
  )

And derived calculation:
  Balance = Total Contract Price - Total Payment Deductible to TCP
"""

from __future__ import annotations

import csv
import io
import math
import os
import re
from typing import Any, Dict, List, Optional, Tuple


def norm_str(val: Any) -> str:
    """Normalize string by stripping and standardizing whitespace."""
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0") and (s[:-2].isdigit() or (s.startswith("-") and s[1:-2].isdigit())):
        s = s[:-2]
    return s


def norm_key(val: Any) -> str:
    """Normalize lookup keys (case-insensitive, collapsed whitespace, punctuation trimmed)."""
    s = norm_str(val).upper()
    # collapse multiple spaces / dashes
    s = re.sub(r'[\s\-_]+', ' ', s).strip()
    return s


def parse_float(val: Any) -> Optional[float]:
    """Parse numeric float from string/number, stripping currency symbols and commas."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if math.isnan(val) or math.isinf(val):
            return None
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    # Remove currency symbols (₱, $, etc.), commas, spaces
    clean = re.sub(r'[^\d.\-+]', '', s)
    if not clean or clean in ('-', '+', '.'):
        return None
    try:
        return float(clean)
    except ValueError:
        return None


def format_currency(val: Any) -> str:
    """Format float as currency with commas and 2 decimals."""
    f = parse_float(val)
    if f is None:
        return ""
    return f"{f:,.2f}"


def format_percent(val: Any) -> str:
    """Format percent value (e.g. 0.75 or 75% -> 75.00%)."""
    if val is None or val == "":
        return ""
    if isinstance(val, str) and "%" in val:
        f = parse_float(val)
        if f is not None:
            return f"{f:.2f}%"
        return val.strip()
    f = parse_float(val)
    if f is None:
        return str(val)
    # If standard 0.0 - 1.0 fraction
    if -1.0 <= f <= 1.0 and f != 0:
        pct = f * 100.0
    else:
        pct = f
    return f"{pct:.2f}%"


def clean_title_mod(val: Any) -> str:
    """Extract stripped numerical/alphanumeric core from Title Number or Title Mod for robust matching."""
    if val is None:
        return ""
    s = str(val).strip()
    # Strip common title prefixes: T-, TCT-, OCT-, CCT-, CLOA-, EP-, NO., etc.
    s_clean = re.sub(r'^(?:TCT|OCT|CCT|CLOA|EP|KATIBAYAN|TITLE|NO|NUM|NUMBER)\s*[-.:#\s]*', '', s, flags=re.IGNORECASE).strip()
    # If pure digits exist, extract digits and strip leading zeros
    digits = re.sub(r'[^\d]', '', s_clean)
    if digits:
        return digits.lstrip('0') or digits
    # Fallback to alphanumeric without spaces/punctuation
    return re.sub(r'[^A-Za-z0-9]', '', s_clean).upper()


class FinanceLookupEngine:
    """Fast indexed lookup engine for APPENDED_FINANCE reference data aligned with Excel LET/XLOOKUP formulas."""

    def __init__(self, finance_rows: Optional[List[Dict[str, Any]]] = None):
        self.raw_rows: List[Dict[str, Any]] = []
        # Index 1: TYPE == "ROW", key: Corridor Index
        self.row_corridor_index: Dict[str, Dict[str, Any]] = {}
        # Index 2: Area Index
        self.area_index: Dict[str, Dict[str, Any]] = {}
        # Index 3: Title Mod
        self.title_mod_index: Dict[str, Dict[str, Any]] = {}

        if finance_rows:
            self.load_rows(finance_rows)

    def load_rows(self, rows: List[Dict[str, Any]]) -> None:
        self.raw_rows = list(rows)
        self.row_corridor_index.clear()
        self.area_index.clear()
        self.title_mod_index.clear()

        for r in rows:
            type_val = norm_key(self._get_field(r, ["TYPE", "Type", "type", "ROW_TYPE"]))
            corridor_idx = norm_key(self._get_field(r, [
                "Corridor Index", "CorridorIndex", "CORRIDOR INDEX", "INDEX NO", "Index No", "index_no", "cr63f_indexno"
            ]))
            area_idx = norm_key(self._get_field(r, [
                "Area Index", "AreaIndex", "AREA INDEX", "AREA-INDEX", "area_index", "cr63f_areaindex"
            ]))
            title_mod_raw = self._get_field(r, [
                "Title Mod", "TitleMod", "TITLE MOD", "title_mod", "TITLE", "TITLE NO", "Title No", "titleno", "cr63f_titleno"
            ])
            title_mod = norm_key(title_mod_raw)
            title_mod_clean = clean_title_mod(title_mod_raw)

            # Index for TYPE="ROW" (Corridor Index)
            if type_val == "ROW" and corridor_idx:
                if corridor_idx not in self.row_corridor_index:
                    self.row_corridor_index[corridor_idx] = r

            # Index for Area Index
            if area_idx and area_idx not in self.area_index:
                self.area_index[area_idx] = r

            # Index for Title Mod (both normalized string and stripped core digits)
            if title_mod and title_mod not in self.title_mod_index:
                self.title_mod_index[title_mod] = r
            if title_mod_clean and title_mod_clean not in self.title_mod_index:
                self.title_mod_index[title_mod_clean] = r

    @staticmethod
    def _get_field(row: Dict[str, Any], candidate_names: List[str]) -> Any:
        for name in candidate_names:
            if name in row and row[name] not in (None, ""):
                return row[name]
        # case-insensitive search
        lower_map = {k.lower().replace(" ", "").replace("_", "").replace("-", "").replace("(", "").replace(")", ""): v for k, v in row.items()}
        for name in candidate_names:
            k = name.lower().replace(" ", "").replace("_", "").replace("-", "").replace("(", "").replace(")", "")
            if k in lower_map and lower_map[k] not in (None, ""):
                return lower_map[k]
        return ""

    def lookup(
        self,
        type_val: Any,
        index_no: Any,
        area_index_val: Any,
        title_mod_val: Any
    ) -> Dict[str, Any]:
        """Execute the exact Excel formula logic and derive finance metrics."""
        norm_t = norm_key(type_val)
        norm_idx = norm_key(index_no)
        norm_area = norm_key(area_index_val)
        norm_title = norm_key(title_mod_val)
        clean_title = clean_title_mod(title_mod_val)

        matched_row: Optional[Dict[str, Any]] = None
        match_type = ""
        match_key = ""

        # Step 1: Check primary logic
        # IF [@TYPE]="ROW", XLOOKUP([@INDEX NO], Corridor Index [TYPE="ROW"])
        if norm_t == "ROW":
            if norm_idx and norm_idx in self.row_corridor_index:
                matched_row = self.row_corridor_index[norm_idx]
                match_type = "ROW (Corridor Index)"
                match_key = norm_str(index_no)
        else:
            # ELSE XLOOKUP([@AREA-INDEX], Area Index)
            if norm_area and norm_area in self.area_index:
                matched_row = self.area_index[norm_area]
                match_type = "Area Index"
                match_key = norm_str(area_index_val)

        # Step 2: Fallback to Title Mod if no result
        # IF Result <> "", Result, IF([@TITLE MOD] <> "", XLOOKUP([@TITLE MOD], Title Mod), "")
        if not matched_row:
            if norm_title and norm_title in self.title_mod_index:
                matched_row = self.title_mod_index[norm_title]
                match_type = "Title Mod (Fallback)"
                match_key = norm_str(title_mod_val)
            elif clean_title and clean_title in self.title_mod_index:
                matched_row = self.title_mod_index[clean_title]
                match_type = "Title Mod (Clean Fallback)"
                match_key = clean_title

        if not matched_row:
            return {
                "matched": False,
                "match_type": "",
                "match_key": "",
                "total_contract_price": None,
                "total_payment_deductible": None,
                "balance": None,
                "percent_paid": None,
                "total_contract_price_formatted": "",
                "total_payment_deductible_formatted": "",
                "balance_formatted": "",
                "percent_paid_formatted": "",
            }

        # Extract values supporting all column naming variations
        tcp_raw = self._get_field(matched_row, [
            "Total Contract Price", "Total Contract Price (Finance)", "TotalContractPrice",
            "TOTAL CONTRACT PRICE", "TOTAL CONTRACT PRICE (FINANCE)", "TCP", "tcp", "cr63f_totalcontractprice"
        ])
        deductible_raw = self._get_field(matched_row, [
            "Total Payment deductible to TCP", "Total Payment Deductible to TCP", "Total Payment Deductible",
            "Total Payment deductible to TCP (Finance)", "Total Payment Deductible to TCP (Finance)",
            "Total Payment Deductible (Finance)", "Payment Deductible (Finance)",
            "Payment Deductible", "Payment Deductible to TCP", "TotalPaymentDeductibleToTCP",
            "TOTAL PAYMENT DEDUCTIBLE TO TCP", "TOTAL PAYMENT DEDUCTIBLE TO TCP (FINANCE)",
            "cr63f_totalpaymentdeductibletotcp"
        ])
        pct_paid_raw = self._get_field(matched_row, [
            "% Paid", "% Paid (Finance)", "% PAID (FINANCE)", "% PAID", "%Paid",
            "Percent Paid", "Percentage Paid", "Percent Paid (Finance)", "cr63f_paid", "cr63f_percentpaid", "PercentPaid"
        ])
        balance_raw = self._get_field(matched_row, [
            "Balance", "Balance (Finance)", "BALANCE", "BALANCE (FINANCE)", "balance", "cr63f_balance"
        ])

        tcp_num = parse_float(tcp_raw)
        deductible_num = parse_float(deductible_raw)
        pct_paid_num = parse_float(pct_paid_raw)
        balance_num = parse_float(balance_raw)

        # Compute balance (balance is never negative on overpayment)
        if tcp_num is not None and deductible_num is not None:
            computed_balance = max(0.0, tcp_num - deductible_num)
        elif balance_num is not None:
            computed_balance = max(0.0, balance_num)
        else:
            computed_balance = None

        # Compute % Paid (reflect actual overpayment in percent paid)
        if tcp_num is not None and tcp_num > 0 and deductible_num is not None:
            pct_val = (deductible_num / tcp_num) * 100.0
        elif pct_paid_num is not None:
            pct_val = pct_paid_num
        elif deductible_num is not None and deductible_num > 0 and (tcp_num is None or tcp_num == 0):
            pct_val = 100.0
        else:
            pct_val = None

        return {
            "matched": True,
            "match_type": match_type,
            "match_key": match_key,
            "total_contract_price": tcp_num,
            "total_payment_deductible": deductible_num,
            "balance": computed_balance,
            "percent_paid": pct_val,
            "total_contract_price_formatted": format_currency(tcp_num),
            "total_payment_deductible_formatted": format_currency(deductible_num),
            "balance_formatted": format_currency(computed_balance),
            "percent_paid_formatted": format_percent(pct_val),
            "matched_row": matched_row,
        }


def read_tabular_file(file_path: str) -> Tuple[List[str], List[Dict[str, Any]]]:
    """Read CSV or Excel (.xlsx/.xls) file into headers and row dicts."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.xlsx', '.xlsm', '.xltx', '.xltm'):
        import openpyxl
        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        sheet = wb.active
        rows_iter = sheet.iter_rows(values_only=True)
        try:
            raw_headers = next(rows_iter)
        except StopIteration:
            return [], []
        headers = [str(h or '').strip() for h in raw_headers if h is not None and str(h).strip()]
        header_len = len(headers)
        rows: List[Dict[str, Any]] = []
        for raw_row in rows_iter:
            if not any(raw_row):
                continue
            row_dict = {}
            for i, h in enumerate(headers):
                val = raw_row[i] if i < len(raw_row) else ""
                row_dict[h] = "" if val is None else val
            rows.append(row_dict)
        return headers, rows

    elif ext == '.csv':
        for enc in ('utf-8-sig', 'utf-8', 'cp1252', 'latin-1'):
            try:
                with open(file_path, 'r', encoding=enc, newline='') as f:
                    reader = csv.reader(f)
                    try:
                        raw_headers = next(reader)
                    except StopIteration:
                        return [], []
                    headers = [str(h or '').strip() for h in raw_headers if h is not None and str(h).strip()]
                    rows: List[Dict[str, Any]] = []
                    for raw_row in reader:
                        if not any(raw_row):
                            continue
                        row_dict = {}
                        for i, h in enumerate(headers):
                            val = raw_row[i] if i < len(raw_row) else ""
                            row_dict[h] = "" if val is None else val
                        rows.append(row_dict)
                    return headers, rows
            except UnicodeDecodeError:
                continue
        raise ValueError(f"Could not decode CSV file {file_path}")
    else:
        raise ValueError(f"Unsupported file format: {ext}")


def write_excel_file(file_path: str, headers: List[str], rows: List[Dict[str, Any]], sheet_name: str = "LOA Records") -> None:
    """Write rows to Excel file with formatted headers and widths."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Header styling
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1A56DB", end_color="1A56DB", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB'),
    )

    ws.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_data in rows:
        row_vals = [row_data.get(h, "") for h in headers]
        ws.append(row_vals)

    # Auto-adjust column widths
    for col_num, h in enumerate(headers, 1):
        max_len = len(str(h))
        for row in range(2, len(rows) + 2):
            cell_val = ws.cell(row=row, column=col_num).value
            if cell_val is not None:
                max_len = max(max_len, len(str(cell_val)))
        col_letter = get_column_letter(col_num)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    wb.save(file_path)

