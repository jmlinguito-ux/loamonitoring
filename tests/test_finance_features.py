"""Unit tests for Finance Template and LOA Update from Finance features."""

import os
import unittest
import tempfile
from unittest.mock import patch
from app import Api
import openpyxl


class TestFinanceFeatures(unittest.TestCase):
    def setUp(self):
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            self.app = Api()
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_download_finance_template(self):
        template_path = os.path.join(self.temp_dir.name, "Test_Finance_Template.xlsx")
        res = self.app.download_finance_template(destination_path=template_path)
        self.assertTrue(res.get("success"), f"Failed to download template: {res}")
        self.assertTrue(os.path.exists(template_path))

        wb = openpyxl.load_workbook(template_path)
        sheet = wb.active
        self.assertEqual(sheet.title, "APPENDED_FINANCE")

        headers = [cell.value for cell in sheet[1]]
        expected_headers = [
            "TYPE",
            "Corridor Index",
            "Area Index",
            "Title Mod",
            "Total Contract Price",
            "Total Payment Deductible to TCP",
            "% Paid",
            "Balance"
        ]
        self.assertEqual(headers, expected_headers)

        # Verify sample rows exist
        self.assertGreaterEqual(sheet.max_row, 3)

    def test_populate_loa_records_from_finance(self):
        # 1. Setup finance reference data in engine
        finance_rows = [
            {
                "TYPE": "ROW",
                "Corridor Index": "C-100",
                "Total Contract Price": "1,500,000.00",
                "Total Payment Deductible to TCP": "500,000.00",
                "% Paid": "33.33%",
                "Balance": "1,000,000.00"
            },
            {
                "TYPE": "NON-ROW",
                "Area Index": "AI-200",
                "Total Contract Price": "2,000,000.00",
                "Total Payment Deductible to TCP": "1,000,000.00",
                "% Paid": "50.00%",
                "Balance": "1,000,000.00"
            },
            {
                "TYPE": "NON-ROW",
                "Title Mod": "T-55555",
                "Total Contract Price": "3,000,000.00",
                "Total Payment Deductible to TCP": "750,000.00",
                "% Paid": "25.00%",
                "Balance": "2,250,000.00"
            }
        ]
        self.app.finance_rows = finance_rows
        self.app.finance_engine.load_rows(finance_rows)

        # 2. Setup LOA records requiring updates
        loa_records = [
            {
                "__id": "rec-1",
                "cr63f_type": "ROW",
                "cr63f_indexno": "C-100",
                "cr63f_areaindex": "",
                "cr63f_titleno": "T-111",
                "Total Contract Price": "",
                "Total Payment Deductible to TCP": ""
            },
            {
                "__id": "rec-2",
                "cr63f_type": "NON-ROW",
                "cr63f_indexno": "",
                "cr63f_areaindex": "AI-200",
                "cr63f_titleno": "T-222",
                "Total Contract Price": "",
                "Total Payment Deductible to TCP": ""
            },
            {
                "__id": "rec-3",
                "cr63f_type": "NON-ROW",
                "cr63f_indexno": "",
                "cr63f_areaindex": "NON-MATCHING",
                "cr63f_titleno": "TCT-55555",
                "Total Contract Price": "",
                "Total Payment Deductible to TCP": ""
            },
            {
                "__id": "rec-4",
                "cr63f_type": "NON-ROW",
                "cr63f_indexno": "",
                "cr63f_areaindex": "NON-MATCHING-2",
                "cr63f_titleno": "T-99999",
                "Total Contract Price": "",
                "Total Payment Deductible to TCP": ""
            }
        ]

        res = self.app.populate_loa_records_from_finance(records=loa_records)
        self.assertTrue(res.get("success"))
        self.assertEqual(res.get("matched_count"), 3)
        self.assertEqual(res.get("updated_count"), 3)

        updated = res.get("records")

        # Rec 1: Matched via Corridor Index
        self.assertEqual(updated[0]["Total Contract Price"], "1,500,000.00")
        self.assertEqual(updated[0]["Total Payment Deductible to TCP"], "500,000.00")
        self.assertEqual(updated[0]["Balance"], "1,000,000.00")
        self.assertEqual(updated[0]["__display_percent_paid"], "33.3333333333%")

        # Rec 2: Matched via Area Index
        self.assertEqual(updated[1]["Total Contract Price"], "2,000,000.00")
        self.assertEqual(updated[1]["Total Payment Deductible to TCP"], "1,000,000.00")
        self.assertEqual(updated[1]["Balance"], "1,000,000.00")
        self.assertEqual(updated[1]["__display_percent_paid"], "50.00%")

        # Rec 3: Matched via Title Mod fallback (stripped core 55555)
        self.assertEqual(updated[2]["Total Contract Price"], "3,000,000.00")
        self.assertEqual(updated[2]["Total Payment Deductible to TCP"], "750,000.00")
        self.assertEqual(updated[2]["Balance"], "2,250,000.00")
        self.assertEqual(updated[2]["__display_percent_paid"], "25.00%")

        # Rec 4: Unmatched
        self.assertEqual(updated[3]["Total Contract Price"], "")

    def test_populate_loa_records_changed_detection(self):
        # Setup finance data
        finance_rows = [
            {
                "TYPE": "NON-ROW",
                "Area Index": "LIA-4527",
                "Total Contract Price": "1,000,000.00",
                "Total Payment Deductible to TCP": "200,000.00",
                "% Paid": "20.00%",
                "Balance": "800,000.00"
            },
            {
                "TYPE": "NON-ROW",
                "Area Index": "LIA-9999",
                "Total Contract Price": "5,000,000.00",
                "Total Payment Deductible to TCP": "1,000,000.00",
                "% Paid": "20.00%",
                "Balance": "4,000,000.00"
            }
        ]
        self.app.finance_rows = finance_rows
        self.app.finance_engine.load_rows(finance_rows)

        loa_records = [
            # LIA-4527 has identical finance numbers already present (Dataverse attributes with currency formatting and decimal % Paid)
            {
                "__id": "rec-unchanged",
                "cr63f_type": "NON-ROW",
                "cr63f_areaindex": "LIA-4527",
                "cr63f_totalcontractprice": "₱1,000,000.00",
                "cr63f_totalpaymentdeductibletotcp": "₱200,000.00",
                "cr63f_balance": "₱800,000.00",
                "cr63f_paid": "20.00",
                "Total Contract Price": "₱1,000,000.00",
                "Payment Deductible": "₱200,000.00",
                "% Paid": "20.00%",
                "Balance": "₱800,000.00"
            },
            # LIA-9999 has different finance numbers (empty or different)
            {
                "__id": "rec-changed",
                "cr63f_type": "NON-ROW",
                "cr63f_areaindex": "LIA-9999",
                "Total Contract Price": "",
                "Payment Deductible": "",
                "% Paid": "",
                "Balance": ""
            }
        ]

        res = self.app.populate_loa_records_from_finance(records=loa_records)
        self.assertTrue(res.get("success"))
        self.assertEqual(res.get("matched_count"), 2)
        # Only 1 record actually changed
        self.assertEqual(res.get("updated_count"), 1)

        updated = res.get("records")
        # Unchanged record
        self.assertTrue(updated[0].get("__finance_matched"))
        self.assertFalse(updated[0].get("__finance_changed"))

        # Changed record
        self.assertTrue(updated[1].get("__finance_matched"))
        self.assertTrue(updated[1].get("__finance_changed"))

    def test_clean_dataverse_datetime_handles_crm_min_and_invalid_dates(self):
        from dataverse import _clean_dataverse_datetime

        # Dataverse CrmDateTime requires 1753-01-01 <= date <= 9999-12-31.
        # Dates like 01/01/0001 (uninitialized .NET DateTime.MinValue), 0001-01-01, '-',
        # or out-of-range dates must be converted to None to avoid error 0x80040239.
        self.assertIsNone(_clean_dataverse_datetime('01/01/0001 00:00:00'))
        self.assertIsNone(_clean_dataverse_datetime('01/01/0001'))
        self.assertIsNone(_clean_dataverse_datetime('0001-01-01'))
        self.assertIsNone(_clean_dataverse_datetime('0001-01-01T00:00:00Z'))
        self.assertIsNone(_clean_dataverse_datetime('-'))
        self.assertIsNone(_clean_dataverse_datetime(''))
        self.assertIsNone(_clean_dataverse_datetime(None))
        self.assertIsNone(_clean_dataverse_datetime('None'))
        self.assertIsNone(_clean_dataverse_datetime('00/00/0000'))
        self.assertIsNone(_clean_dataverse_datetime('invalid-date'))

        # Valid dates should format cleanly
        self.assertEqual(_clean_dataverse_datetime('05/20/2024'), '2024-05-20')
        self.assertEqual(_clean_dataverse_datetime('2024-05-20'), '2024-05-20')
        self.assertEqual(_clean_dataverse_datetime('05/20/24'), '2024-05-20')
        # App times are local wall-clock times. User-local columns get them in UTC
        # (labelling local time as 'Z' shifted it by the UTC offset).
        from datetime import datetime, timezone
        expected_utc = datetime(2026, 9, 29, 13, 46).astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        self.assertEqual(_clean_dataverse_datetime('09/29/2026 01:46 PM'), expected_utc)
        self.assertEqual(_clean_dataverse_datetime('09/29/2026 01:46 PM', 'TimeZoneIndependent'), '2026-09-29T13:46:00')
        self.assertEqual(_clean_dataverse_datetime('09/29/2026 01:46 PM', 'DateOnly'), '2026-09-29')
        # Already-UTC ISO values are kept as UTC.
        self.assertEqual(_clean_dataverse_datetime('2026-09-29T05:46:00Z'), '2026-09-29T05:46:00Z')


if __name__ == "__main__":
    unittest.main()


