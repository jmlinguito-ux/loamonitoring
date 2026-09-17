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
        self.assertEqual(res["percent_paid_formatted"], "33.33%")

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
        self.assertEqual(res["match_type"], "Title Mod (Fallback)")
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
        self.assertEqual(res["match_type"], "Title Mod (Clean Fallback)")
        self.assertEqual(res["total_contract_price"], 3000000.0)
        self.assertEqual(res["total_payment_deductible"], 1000000.0)
        self.assertEqual(res["balance"], 2000000.0)
        self.assertEqual(res["percent_paid_formatted"], "33.33%")

    def test_overpayment_balance_and_percent(self):
        engine = FinanceLookupEngine([{
            "TYPE": "LAND",
            "Area Index": "AI-OVERPAY",
            "Total Contract Price": "1,000,000.00",
            "Total Payment Deductible to TCP": "1,500,000.00",
        }])
        res = engine.lookup(type_val="LAND", index_no="", area_index_val="AI-OVERPAY", title_mod_val="")
        self.assertTrue(res["matched"])
        self.assertEqual(res["balance"], 0.0)
        self.assertEqual(res["balance_formatted"], "0.00")
        self.assertEqual(res["percent_paid"], 150.0)
        self.assertEqual(res["percent_paid_formatted"], "150.00%")


if __name__ == "__main__":
    unittest.main()

