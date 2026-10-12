"""Unit tests for Reference Tables (Finance & Conversion) Staged Edits & Dataverse Sync."""

import unittest
from unittest.mock import MagicMock, patch
from app import Api


class TestReferenceStaging(unittest.TestCase):
    def setUp(self):
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            self.app = Api()
        # Mock DataverseClient
        self.mock_dv = MagicMock()
        self.mock_dv.config = {
            'loa_table': 'cr63f_loamonitoring',
            'finance_table': 'cr63f_financereference',
            'conversion_table': 'cr63f_landconversionreference',
            'logs_table': 'cr63f_loamonitoringlogs'
        }
        self.mock_dv.save_finance_record.return_value = {'success': True, 'id': 'fin-101'}
        self.mock_dv.save_conversion_record.return_value = {'success': True, 'id': 'conv-202'}
        self.mock_dv.save_log_record.return_value = {'success': True, 'id': 'log-303'}
        self.app.dv_client = self.mock_dv

    def test_save_staged_finance_records(self):
        self.app.finance_rows = [
            {
                '__id': 'fin-101',
                'TYPE': 'ROW',
                'Corridor Index': 'COR-1',
                'Area Index': 'AI-100',
                'Total Contract Price': '1,000,000.00',
                'Balance': '500,000.00'
            }
        ]
        self.app.finance_engine.load_rows(self.app.finance_rows)

        staged_list = [
            {
                'id': 'fin-101',
                'fields': {
                    'Total Contract Price': '1,200,000.00',
                    'Balance': '700,000.00'
                }
            }
        ]

        res = self.app.save_staged_finance_records(staged_list)
        self.assertTrue(res['success'])
        self.assertEqual(res['saved_count'], 1)

        # Check Dataverse client save call
        self.mock_dv.save_finance_record.assert_called()
        self.mock_dv.save_log_record.assert_called()

        # Verify updated values in finance_rows
        self.assertEqual(self.app.finance_rows[0]['Total Contract Price'], '1,200,000.00')
        self.assertEqual(self.app.finance_rows[0]['Balance'], '700,000.00')

        # Verify audit logs in log_records
        self.assertGreaterEqual(len(self.app.log_records), 1)
        fin_log = next((l for l in self.app.log_records if l.get('action_type') == 'FINANCE_UPDATE'), None)
        self.assertIsNotNone(fin_log)
        self.assertEqual(fin_log['action_type'], 'FINANCE_UPDATE')
        self.assertIn('AI-100', fin_log.get('area_index', '') + fin_log.get('cr63f_areaindex', ''))

    def test_save_staged_conversion_records(self):
        self.app.conversion_rows = [
            {
                '__id': 'conv-202',
                'AREA-INDEX': 'AI-200',
                'TITLE NUMBER': 'T-999',
                'DAR STATUS': 'FOR APPLICATION'
            }
        ]
        self.app.conversion_engine.load_rows(self.app.conversion_rows)

        staged_list = [
            {
                'id': 'conv-202',
                'fields': {
                    'DAR STATUS': 'DAR FINALITY SECURED'
                }
            }
        ]

        res = self.app.save_staged_conversion_records(staged_list)
        self.assertTrue(res['success'])
        self.assertEqual(res['saved_count'], 1)

        # Check Dataverse client save call
        self.mock_dv.save_conversion_record.assert_called()
        self.mock_dv.save_log_record.assert_called()

        # Verify updated value in conversion_rows
        self.assertEqual(self.app.conversion_rows[0]['DAR STATUS'], 'DAR FINALITY SECURED')

        # Verify audit logs in log_records
        self.assertGreaterEqual(len(self.app.log_records), 1)
        conv_log = next((l for l in self.app.log_records if l.get('action_type') == 'CONVERSION_UPDATE'), None)
        self.assertIsNotNone(conv_log)
        self.assertEqual(conv_log['action_type'], 'CONVERSION_UPDATE')
        self.assertEqual(conv_log.get('field_name'), 'DAR STATUS')
        self.assertEqual(conv_log.get('new_value'), 'DAR FINALITY SECURED')

    def test_revert_staged_records(self):
        # Baseline rows
        orig_fin = [{'__id': '1', 'Area Index': 'A', 'Total Contract Price': '100'}]
        self.app.finance_rows = [{'__id': '1', 'Area Index': 'A', 'Total Contract Price': '200'}]
        res = self.app.revert_staged_finance_records(orig_fin)
        self.assertTrue(res['success'])
        self.assertEqual(self.app.finance_rows[0]['Total Contract Price'], '100')

        orig_conv = [{'__id': '2', 'AREA-INDEX': 'B', 'DAR STATUS': 'OLD'}]
        self.app.conversion_rows = [{'__id': '2', 'AREA-INDEX': 'B', 'DAR STATUS': 'NEW'}]
        res2 = self.app.revert_staged_conversion_records(orig_conv)
        self.assertTrue(res2['success'])
        self.assertEqual(self.app.conversion_rows[0]['DAR STATUS'], 'OLD')


if __name__ == '__main__':
    unittest.main()
