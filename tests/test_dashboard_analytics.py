"""Unit tests for Dashboard Analytics, Mode Filtering on Conversion Status, and Financial Settlement Breakdown."""

import unittest
from unittest.mock import patch
from app import Api


class TestDashboardAnalytics(unittest.TestCase):
    def setUp(self):
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            self.app = Api()

        self.app.loa_records = [
            {
                "__id": "1",
                "cr63f_areaindex": "NAS-001",
                "cr63f_titleno": "T-100",
                "new_mode": "SECURED SALE",
                "cr63f_conversionstatus": "DAR SECURED",
                "cr63f_lotareaha": "2.5",
                "cr63f_totalcontractprice": "1000000",
                "cr63f_totalpaymentdeductibletotcp": "1000000",
                "cr63f_balance": "0",
                "cr63f_paid": "100.00%",
                "work_week": "WW-38"
            },
            {
                "__id": "2",
                "cr63f_areaindex": "NAS-002",
                "cr63f_titleno": "T-101",
                "new_mode": "SECURED SALE",
                "cr63f_conversionstatus": "DAR FOR APPLICATION",
                "cr63f_lotareaha": "1.5",
                "cr63f_totalcontractprice": "2000000",
                "cr63f_totalpaymentdeductibletotcp": "500000",
                "cr63f_balance": "1500000",
                "cr63f_paid": "25.00%",
                "work_week": "WW-38"
            },
            {
                "__id": "3",
                "cr63f_areaindex": "NAS-003",
                "cr63f_titleno": "T-102",
                "new_mode": "SS - STP",
                "cr63f_conversionstatus": "DAR INSPECTION DONE",
                "cr63f_lotareaha": "3.0",
                "cr63f_totalcontractprice": "1500000",
                "cr63f_totalpaymentdeductibletotcp": "0",
                "cr63f_balance": "1500000",
                "cr63f_paid": "0.00%",
                "work_week": "WW-37"
            }
        ]

    def test_dashboard_analytics_all_modes(self):
        res = self.app.get_dashboard_analytics(work_week="ALL", mode="ALL")
        self.assertTrue(res["success"])
        self.assertEqual(res["total_records"], 3)
        self.assertEqual(res["total_tcp"], 4500000.0)
        self.assertEqual(res["total_deductible"], 1500000.0)
        
        # Financial Breakdown
        fb = res["financial_breakdown"]
        self.assertEqual(fb["fully_paid_count"], 1)
        self.assertEqual(fb["partially_paid_count"], 1)
        self.assertEqual(fb["zero_paid_count"], 1)

        # Conversion Metrics
        cm = res["conversion_metrics"]
        self.assertEqual(cm["total_lots"], 3)
        self.assertEqual(cm["secured_count"], 1)
        self.assertEqual(cm["in_progress_count"], 2)
        self.assertEqual(cm["total_applied_area"], 7.0)

    def test_dashboard_analytics_mode_filtered(self):
        # Filter by SECURED SALE (2 records)
        res = self.app.get_dashboard_analytics(work_week="ALL", mode="SECURED SALE")
        self.assertTrue(res["success"])
        self.assertEqual(res["total_records"], 2)
        self.assertEqual(res["selected_mode"], "SECURED SALE")
        self.assertEqual(res["total_tcp"], 3000000.0)
        self.assertEqual(res["total_deductible"], 1500000.0)

        # Financial Breakdown reflects only SECURED SALE
        fb = res["financial_breakdown"]
        self.assertEqual(fb["fully_paid_count"], 1)
        self.assertEqual(fb["partially_paid_count"], 1)
        self.assertEqual(fb["zero_paid_count"], 0)

        # Conversion metrics are now properly filtered!
        cm = res["conversion_metrics"]
        self.assertEqual(cm["total_lots"], 2)
        self.assertEqual(cm["secured_count"], 1)
        self.assertEqual(cm["in_progress_count"], 1)
        self.assertEqual(cm["total_applied_area"], 4.0)

        statuses = [b["status"] for b in cm["breakdown"]]
        self.assertIn("DAR SECURED", statuses)
        self.assertIn("DAR FOR APPLICATION", statuses)
        self.assertNotIn("DAR INSPECTION DONE", statuses)

    def test_dashboard_analytics_recent_activity_audit_columns(self):
        self.app.log_records = [
            {
                "action_type": "REFERENCE_UPDATE",
                "area_index": "NAS-001",
                "work_week": "WW-38",
                "log_time": "2026-09-17T10:00:00Z",
                "field_name": "Total Contract Price",
                "old_value": "0",
                "new_value": "900,000.00",
                "cr63f_totalcontractprice": "1000000",
                "cr63f_acquisitionsublevel": "L7.5 - TITLE TRANSFER DONE",
                "author": "j.inguito"
            },
            {
                "action_type": "REFERENCE_UPDATE",
                "area_index": "NAS-001",
                "work_week": "WW-39",
                "log_time": "2026-09-24T10:00:00Z",
                "field_name": "Total Contract Price",
                "old_value": "900,000.00",
                "new_value": "1,000,000.00",
                "cr63f_totalcontractprice": "1000000",
                "cr63f_acquisitionsublevel": "L7.5 - TITLE TRANSFER DONE",
                "author": "j.inguito"
            }
        ]
        res = self.app.get_dashboard_analytics(work_week="ALL")
        self.assertTrue(res["success"])
        self.assertGreaterEqual(len(res["changed_accounts"]), 1)
        account_act = res["changed_accounts"][0]
        self.assertEqual(account_act["area_index"], "NAS-001")
        self.assertEqual(account_act["contract_price_formatted"], "₱1,000,000.00")
        self.assertGreaterEqual(len(account_act["changes"]), 1)
        change = account_act["changes"][0]
        self.assertEqual(change["field_name"], "Total Contract Price")
        self.assertEqual(change["old_value"], "900,000.00")
        self.assertEqual(change["new_value"], "1,000,000.00")
        self.assertEqual(change["updated_by"], "j.inguito")
        self.assertEqual(change["log_time"], "2026-09-24T10:00:00Z")

    def test_dashboard_analytics_mode_filter_and_distribution(self):
        # Without mode filter
        res_all = self.app.get_dashboard_analytics(work_week="ALL", mode=None)
        self.assertTrue(res_all["success"])
        self.assertEqual(res_all["total_records"], 3)
        self.assertIsNone(res_all["selected_mode"])
        modes = res_all["modes_distribution"]
        self.assertEqual(len(modes), 2)
        # Check active flag when filtering by SS - STP
        res_filtered = self.app.get_dashboard_analytics(work_week="ALL", mode="SS - STP")
        self.assertTrue(res_filtered["success"])
        self.assertEqual(res_filtered["total_records"], 1)
        self.assertEqual(res_filtered["selected_mode"], "SS - STP")
        modes_filtered = res_filtered["modes_distribution"]
        ss_mode = next((m for m in modes_filtered if m["mode"] == "SS - STP"), None)
        self.assertIsNotNone(ss_mode)
        self.assertTrue(ss_mode["is_active"])
        other_mode = next((m for m in modes_filtered if m["mode"] != "SS - STP"), None)
        self.assertIsNotNone(other_mode)
        self.assertFalse(other_mode["is_active"])


if __name__ == "__main__":
    unittest.main()
