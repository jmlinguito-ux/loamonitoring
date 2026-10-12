"""Unit tests for LOA engine lookup and calculations."""

import unittest
from loa_engine import FinanceLookupEngine, norm_key, parse_float, format_currency, format_percent


class TestFinanceLookupEngine(unittest.TestCase):
    def setUp(self):
        # Sample APPENDED_FINANCE dataset
        self.finance_data = [
            {
                "TYPE": "ROW",
                "Corridor Index": "C-101",
                "Area Index": "AI-999",
                "Title Mod": "T-101-MOD",
                "Total Contract Price": "1,500,000.00",
                "Total Payment Deductible to TCP": "500,000.00",
                "% Paid": "33.33%",
            },
            {
                "TYPE": "NON-ROW",
                "Corridor Index": "C-202",
                "Area Index": "AI-202",
                "Title Mod": "T-202-MOD",
                "Total Contract Price": "2,000,000.00",
                "Total Payment Deductible to TCP": "1,000,000.00",
                "% Paid": "50.00%",
            },
            {
                "TYPE": "ROW",
                "Corridor Index": "C-303",
                "Area Index": "AI-303",
                "Title Mod": "T-303-SPECIAL",
                "Total Contract Price": "5,000,000.00",
                "Total Payment Deductible to TCP": "1,250,000.00",
                "% Paid": "25.00%",
            },
            {
                "TYPE": "NON-ROW",
                "Corridor Index": "",
                "Area Index": "",
                "Title Mod": "T-FALLBACK-01",
                "Total Contract Price": "800,000.00",
                "Total Payment Deductible to TCP": "200,000.00",
                "% Paid": "25.00%",
            }
        ]
        self.engine = FinanceLookupEngine(self.finance_data)

    def test_row_type_lookup(self):
        # When TYPE == "ROW", should look up INDEX NO in Corridor Index (where TYPE="ROW")
        res = self.engine.lookup(
            type_val="ROW",
            index_no="C-101",
            area_index_val="NON_MATCHING",
            title_mod_val=""
        )
        self.assertTrue(res["matched"])
        self.assertEqual(res["match_type"], "ROW (Corridor Index)")
        self.assertEqual(res["total_contract_price"], 1500000.0)
        self.assertEqual(res["total_payment_deductible"], 500000.0)
        self.assertEqual(res["balance"], 1000000.0)
        self.assertEqual(res["balance_formatted"], "1,000,000.00")
        self.assertEqual(res["percent_paid_formatted"], "33.3333333333%")

    def test_non_row_type_lookup(self):
        # When TYPE != "ROW", should look up AREA-INDEX in Area Index
        res = self.engine.lookup(
            type_val="LOT",
            index_no="NON_MATCHING",
            area_index_val="AI-202",
            title_mod_val=""
        )
        self.assertTrue(res["matched"])
        self.assertEqual(res["match_type"], "Area Index")
        self.assertEqual(res["total_contract_price"], 2000000.0)
        self.assertEqual(res["total_payment_deductible"], 1000000.0)
        self.assertEqual(res["balance"], 1000000.0)
        self.assertEqual(res["percent_paid_formatted"], "50.00%")

    def test_fallback_title_mod_lookup(self):
        # If primary lookup fails, should fall back to TITLE MOD lookup
        res = self.engine.lookup(
            type_val="LOT",
            index_no="C-999",
            area_index_val="UNKNOWN_AREA",
            title_mod_val="T-FALLBACK-01"
        )
        self.assertTrue(res["matched"])
        self.assertEqual(res["match_type"], "Title Mod (Normalized Exact)")
        self.assertEqual(res["total_contract_price"], 800000.0)
        self.assertEqual(res["total_payment_deductible"], 200000.0)
        self.assertEqual(res["balance"], 600000.0)

    def test_clean_title_mod_fallback_lookup(self):
        # Testing title mod cleanup matching (e.g., 'T-55555' vs '55555')
        engine = FinanceLookupEngine([{
            "TYPE": "LAND",
            "Area Index": "OTHER",
            "Title Mod": "55555",
            "Total Contract Price (Finance)": "3,000,000.00",
            "Total Payment Deductible to TCP (Finance)": "1,000,000.00",
            "% Paid (Finance)": "33.33%",
            "Balance (Finance)": "2,000,000.00",
        }])
        res = engine.lookup(
            type_val="LAND",
            index_no="",
            area_index_val="NONE",
            title_mod_val="T-55555"
        )
        self.assertTrue(res["matched"])
        self.assertEqual(res["match_type"], "Title Mod (Normalized Exact)")
        self.assertEqual(res["total_contract_price"], 3000000.0)
        self.assertEqual(res["total_payment_deductible"], 1000000.0)
        self.assertEqual(res["balance"], 2000000.0)
        self.assertEqual(res["percent_paid_formatted"], "33.3333333333%")

    def test_title_fallback_preserves_title_prefix_number_segments(self):
        # Similar titles share a suffix, but the leading title-number segment
        # distinguishes the intended ROW reference from a sale record.
        sale = {
            "TYPE": "SALE",
            "Area Index": "PAG-122",
            "Title Mod": "TCT-066-2000001234",
            "Total Contract Price": "3,000,030",
            "Total Payment Deductible to TCP": "3,000,030",
        }
        row_reference = {
            "TYPE": "ROW",
            "Corridor Index": "OTHER-ROW",
            "Area Index": "CON-1601",
            "Title Mod": "043-2000001234",
            "Total Contract Price": "300,000",
            "Total Payment Deductible to TCP": "150,000",
        }
        engine = FinanceLookupEngine([sale, row_reference])

        result = engine.lookup(
            type_val="ROW",
            index_no="TARROW-TST-00901",
            area_index_val="1601",
            title_mod_val="TCT-043-2000001234",
        )

        self.assertTrue(result["matched"])
        self.assertIs(result["matched_row"], row_reference)
        self.assertEqual(result["total_contract_price"], 300000.0)
        self.assertEqual(result["total_payment_deductible"], 150000.0)

    def test_ambiguous_shared_title_suffix_does_not_pick_first_reference(self):
        engine = FinanceLookupEngine([
            {"TYPE": "SALE", "Title Mod": "TCT-066-2000001234", "Total Contract Price": "3000030"},
            {"TYPE": "ROW", "Title Mod": "043-2000001234", "Total Contract Price": "300000"},
        ])

        result = engine.lookup(
            type_val="LOT",
            index_no="",
            area_index_val="NO-MATCH",
            title_mod_val="TCT-999-2000001234",
        )

        self.assertFalse(result["matched"])

    def test_row_title_fallback_rejects_non_row_finance_type(self):
        lease_reference = {
            "TYPE": "LEASE",
            "Area Index": "CON-33",
            "Title Mod": "10006",
            "Registered Owner": "ELIZABETH S. MALLARI",
            "Total Contract Price": "900,000",
            "Total Payment Deductible to TCP": "300,000",
        }
        engine = FinanceLookupEngine([lease_reference])

        result = engine.lookup(
            type_val="ROW",
            index_no="TARROW-TST-00902",
            area_index_val="4745",
            title_mod_val="TCT-10006",
        )

        self.assertFalse(result["matched"])

class TestWorkWeekCalculation(unittest.TestCase):
    def test_work_week_derivation(self):
        from loa_engine import get_work_week
        from datetime import datetime

        # Test Wednesday-to-Tuesday cycle
        self.assertEqual(get_work_week("09/17/2026"), "WW-38")
        self.assertEqual(get_work_week("01/01/2026"), "WW-01")
        self.assertEqual(get_work_week("12/31/2025"), "WW-01")
        self.assertEqual(get_work_week("01/05/2025"), "WW-01")
        # 2024 starts on a Monday, so its WW-01 is Wed Jan 3 - Tue Jan 9, 2024
        # (the week holding Jan 1 belongs to 2023). Sep 1, 2024 is WW-35.
        # The old count skipped WW-01 in such years and returned WW-36.
        self.assertEqual(get_work_week("09/01/2024"), "WW-35")
        self.assertEqual(get_work_week("01/03/2024"), "WW-01")
        self.assertEqual(get_work_week("01/03/2029"), "WW-01")
        self.assertEqual(get_work_week("01/09/2029"), "WW-01")
        self.assertEqual(get_work_week("01/02/2029"), "WW-53")  # last week of 2028
        self.assertEqual(get_work_week("2026-09-17"), "WW-38")
        self.assertEqual(get_work_week(datetime(2026, 9, 17)), "WW-38")
        self.assertEqual(get_work_week(""), "")
        self.assertEqual(get_work_week("01/01/1900"), "")


if __name__ == "__main__":
    unittest.main()


