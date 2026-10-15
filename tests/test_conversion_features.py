"""Unit tests for Conversion Reference Engine and LOA Update from References features."""

import os
import tempfile
import unittest
import openpyxl

from unittest.mock import patch
from app import Api
from loa_engine import ConversionLookupEngine, clean_title_variants, clean_title_mod


class TestConversionFeatures(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            self.app = Api()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_title_normalization_and_variants(self):
        # 1. Standard prefix variants
        variants_1 = clean_title_variants("TCT-T-10003")
        variants_2 = clean_title_variants("T-10003")
        variants_3 = clean_title_variants("10003")
        self.assertIn("10003", variants_1)
        self.assertIn("10003", variants_2)
        self.assertIn("10003", variants_3)

        # 2. Parenthesized variants e.g. TCT (T-121628) 055-2025002708
        complex_variants = clean_title_variants("TCT (T-121628) 055-2025002708")
        self.assertIn("121628", complex_variants)
        self.assertIn("2025002708", complex_variants)

        # 3. Clean title mod
        self.assertEqual(clean_title_mod("TCT-CLOA-1001"), "1001")
        self.assertEqual(clean_title_mod("OCT-P-00123"), "123")

    def test_conversion_lookup_uses_area_index_only(self):
        conversion_data = [
            {
                "AREA-INDEX": "NAS-1397",
                "TITLE NUMBER": "TCT-T-10003",
                "DAR STATUS": "DAR SECURED",
                "NIA STATUS": "NIA INSPECTION DONE",
                "CONVERSION CLASS": "Class A"
            },
            {
                "AREA-INDEX": "TST-7786",
                "TITLE NUMBER": "TCT (T-121628) 055-2025002708",
                "DAR STATUS": "DAR FOR APPLICATION",
                "NIA STATUS": "NIA FOR APPLICATION",
                "CONVERSION CLASS": "Class A"
            },
            {
                "AREA-INDEX": "CAL-9999",
                "TITLE NUMBER": "TCT-CLOA-8888",
                "DAR STATUS": "DAR FINALITY SECURED",
                "NIA STATUS": "NIA SECURED",
                "CONVERSION CLASS": "Class B"
            }
        ]

        engine = ConversionLookupEngine(conversion_data)

        # Area Index is the sole key; conflicting/missing titles are irrelevant.
        res1 = engine.lookup(area_index_val="NAS-1397")
        self.assertTrue(res1["matched"])
        self.assertEqual(res1["match_type"], "Area Index")
        self.assertEqual(res1["conversion_status"], "DAR SECURED")

        # Matching Area Index wins even when the supplied title belongs elsewhere.
        res2 = engine.lookup(area_index_val="NAS-1397")
        self.assertTrue(res2["matched"])
        self.assertEqual(res2["match_type"], "Area Index")
        self.assertEqual(res2["conversion_status"], "DAR SECURED")

        # Titles never create a match when Area Index is absent or different.
        res3 = engine.lookup(area_index_val="")
        self.assertFalse(res3["matched"])
        res4 = engine.lookup(area_index_val="UNKNOWN-AREA")
        self.assertFalse(res4["matched"])

        # Non-match
        res_none = engine.lookup(area_index_val="UNKNOWN-000")
        self.assertFalse(res_none["matched"])
        self.assertEqual(res_none["conversion_status"], "")

    def test_download_conversion_template(self):
        template_path = os.path.join(self.temp_dir.name, "Test_Conversion_Template.xlsx")
        res = self.app.download_conversion_template(destination_path=template_path)
        self.assertTrue(res["success"])
        self.assertTrue(os.path.exists(template_path))

        wb = openpyxl.load_workbook(template_path)
        sheet = wb.active
        self.assertEqual(sheet.title, "CONVERSION_REFERENCE")
        headers = [cell.value for cell in sheet[1]]
        self.assertIn("AREA-INDEX", headers)
        self.assertIn("TITLE NUMBER", headers)
        self.assertIn("DAR STATUS", headers)
        self.assertIn("NIA STATUS", headers)
        self.assertIn("CONVERSION CLASS", headers)

    def test_populate_loa_records_from_references(self):
        # Setup Conversion and Finance reference data
        conversion_rows = [
            {
                "AREA-INDEX": "NAS-1397",
                "TITLE NUMBER": "TCT-T-10003",
                "DAR STATUS": "DAR SECURED",
            },
            {
                "AREA-INDEX": "TST-7786",
                "TITLE NUMBER": "TCT (T-121628) 055-2025002708",
                "DAR STATUS": "DAR FOR APPLICATION",
            }
        ]
        finance_rows = [
            {
                "TYPE": "LAND",
                "Area Index": "NAS-1397",
                "Title Mod": "T-10003",
                "Total Contract Price": "5,000,000.00",
                "Total Payment Deductible to TCP": "2,000,000.00",
                "% Paid": "40.00%",
                "Balance": "3,000,000.00"
            }
        ]

        self.app.conversion_rows = conversion_rows
        self.app.conversion_engine.load_rows(conversion_rows)
        self.app.finance_rows = finance_rows
        self.app.finance_engine.load_rows(finance_rows)

        loa_records = [
            {
                "__id": "rec-1",
                "cr63f_type": "LAND",
                "cr63f_areaindex": "NAS-1397",
                "cr63f_titleno": "TCT-T-10003",
                "cr63f_conversionstatus": "NOT YET ENDORSED",
                "cr63f_totalcontractprice": "",
                "cr63f_totalpaymentdeductibletotcp": "",
                "cr63f_balance": "",
                "cr63f_paid": ""
            },
            {
                "__id": "rec-2",
                "cr63f_type": "LAND",
                "cr63f_areaindex": "TST-7786",
                "cr63f_titleno": "T-121628",
                "cr63f_conversionstatus": "DAR FOR APPLICATION", # Already same
                "cr63f_totalcontractprice": "",
                "cr63f_totalpaymentdeductibletotcp": "",
                "cr63f_balance": "",
                "cr63f_paid": ""
            }
        ]

        res = self.app.populate_loa_records_from_references(records=loa_records)
        self.assertTrue(res["success"])
        updated = res["records"]

        # Rec-1: Both Finance and Conversion matched and changed
        self.assertTrue(updated[0]["__finance_matched"])
        self.assertTrue(updated[0]["__finance_changed"])
        self.assertTrue(updated[0]["__conversion_matched"])
        self.assertTrue(updated[0]["__conversion_changed"])
        self.assertEqual(updated[0]["cr63f_conversionstatus"], "DAR SECURED")
        self.assertIn("cr63f_conversionstatus", updated[0]["__changed_fields"])
        self.assertIn("cr63f_totalcontractprice", updated[0]["__changed_fields"])

        # Rec-2: Conversion matched by exact Area Index; status was already DAR FOR APPLICATION (no diff)
        self.assertTrue(updated[1]["__conversion_matched"])
        self.assertFalse(updated[1]["__conversion_changed"])
        self.assertEqual(updated[1]["__conversion_changed_fields"], [])


if __name__ == "__main__":
    unittest.main()
