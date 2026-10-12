"""Unit tests for Excel export column filtering."""

import os
import unittest
import tempfile
from unittest.mock import patch
from app import Api
import openpyxl


class TestExportColumnFiltering(unittest.TestCase):
    def setUp(self):
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            self.app = Api()
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_export_excludes_system_and_value_columns(self):
        records = [
            {
                '__id': 'guid-1234',
                '__etag': 'etag-abc',
                '_owninguser_value': 'user-guid-001',
                '_transactioncurrencyid_value': 'curr-guid-002',
                '_modifiedonbehalfby_value': 'mod-guid-003',
                '_createdonbehalfby_value': 'cre-guid-004',
                '_modifiedby_value': 'modby-guid-005',
                '_owningbusinessunit_value': 'bu-guid-006',
                '_createdby_value': 'creby-guid-007',
                '_owningteam_value': 'team-guid-008',
                '_ownerid_value': 'owner-guid-009',
                'createdon': '2026-01-01',
                'modifiedon': '2026-01-02',
                'statuscode': 1,
                'statecode': 0,
                'Corridor Index': 'C-100',
                'Total Contract Price': '1,500,000.00',
                'Total Payment Deductible': '500,000.00',
                '% Paid': '33.33%'
            }
        ]
        out_path = os.path.join(self.temp_dir.name, "test_export.xlsx")
        res = self.app.export_records_excel(records, destination_path=out_path)
        self.assertTrue(res.get("success"))

        wb = openpyxl.load_workbook(out_path)
        sheet = wb.active
        headers = [cell.value for cell in sheet[1]]

        # Verify excluded columns are not present
        forbidden = [
            '_owninguser_value', '_transactioncurrencyid_value',
            '_modifiedonbehalfby_value', '_createdonbehalfby_value',
            '_modifiedby_value', '_owningbusinessunit_value',
            '_createdby_value', '_owningteam_value', '_ownerid_value',
            '__id', '__etag', 'createdon', 'modifiedon', 'statuscode', 'statecode'
        ]
        for f in forbidden:
            self.assertNotIn(f, headers)

        # Verify expected custom columns are present
        self.assertIn('Corridor Index', headers)
        self.assertIn('Total Contract Price', headers)
        self.assertIn('Total Payment Deductible', headers)
        self.assertIn('% Paid', headers)

    def test_export_with_explicit_columns_filters_system(self):
        records = [
            {
                'Corridor Index': 'C-100',
                'Total Contract Price': '1,500,000.00',
                '_ownerid_value': 'owner-123',
                '#': 1
            }
        ]
        cols = ['#', 'Corridor Index', '_ownerid_value', 'Total Contract Price']
        out_path = os.path.join(self.temp_dir.name, "test_export_explicit.xlsx")
        res = self.app.export_records_excel(records, columns=cols, destination_path=out_path)
        self.assertTrue(res.get("success"))

        wb = openpyxl.load_workbook(out_path)
        sheet = wb.active
        headers = [cell.value for cell in sheet[1]]

        self.assertNotIn('#', headers)
        self.assertNotIn('_ownerid_value', headers)
        self.assertEqual(headers, ['Corridor Index', 'Total Contract Price'])


if __name__ == "__main__":
    unittest.main()
