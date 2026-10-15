import unittest
from unittest.mock import MagicMock, patch

from dataverse import _clean_dataverse_datetime, _parse_number_strict, DataverseClient
from loa_engine import clean_title_mod, clean_title_variants, get_work_week
from app import Api


class TestAuditReviewFixes(unittest.TestCase):
    def test_clean_dataverse_datetime_day_first(self):
        """Dates with Day > 12 should be recognized and normalized to YYYY-MM-DD."""
        self.assertEqual(_clean_dataverse_datetime('25/12/2026'), '2026-12-25')
        self.assertEqual(_clean_dataverse_datetime('31/01/2025'), '2025-01-31')
        self.assertEqual(_clean_dataverse_datetime('15/06/24'), '2024-06-15')
        # Standard MM/DD/YYYY should remain intact
        self.assertEqual(_clean_dataverse_datetime('12/25/2026'), '2026-12-25')
        self.assertEqual(_clean_dataverse_datetime('2026-12-25'), '2026-12-25')

    def test_parse_number_strict_nan_and_inf(self):
        """NaN and Infinity should return None, avoiding invalid JSON in Dataverse payloads."""
        self.assertIsNone(_parse_number_strict(float('nan')))
        self.assertIsNone(_parse_number_strict(float('inf')))
        self.assertIsNone(_parse_number_strict(float('-inf')))
        self.assertEqual(_parse_number_strict(1234.5), 1234.5)
        self.assertEqual(_parse_number_strict('₱1,234.50'), 1234.5)

    def test_clean_title_mod_deterministic(self):
        """clean_title_mod should deterministically pick the main/longest numeric core."""
        val1 = clean_title_mod('TCT-066-2019006115')
        val2 = clean_title_mod('OCT-P-00123')
        for _ in range(10):
            self.assertEqual(clean_title_mod('TCT-066-2019006115'), val1)
            self.assertEqual(clean_title_mod('OCT-P-00123'), val2)
        self.assertEqual(val2, '123')

    def test_save_loa_record_entry_derives_diff_when_baseline_none(self):
        """When editing an existing account whose baseline is None, in-memory diffs must be derived."""
        api = Api()
        api.dv_client = MagicMock()
        api.dv_client.save_loa_record.return_value = {'action': 'update', 'id': 'rec_1', 'etag': 'etag_2'}
        api.dv_client.save_log_record.return_value = {'id': 'log_saved_1'}

        orig_record = {
            '__id': 'rec_1',
            '__etag': 'etag_1',
            'cr63f_control1': 'CTRL-100',
            'cr63f_areaindex': 'AI-100',
            'cr63f_totalcontractprice': 1000.0,
            'Total Contract Price': '₱1,000.00'
        }
        api.loa_records = [dict(orig_record)]
        api.log_records = []

        # Simulate existing record with history -> baseline is None
        with patch.object(api, '_write_baseline_if_missing', return_value=None):
            res = api.save_loa_record_entry(
                {'cr63f_totalcontractprice': 2500.0},
                record_id='rec_1'
            )

        self.assertTrue(res.get('success'), f"Expected success but got error: {res.get('error')}")
        self.assertGreater(len(api.log_records), 0)
        tcp_diff = next((l for l in api.log_records if l.get('field_name') == 'Total Contract Price'), None)
        self.assertIsNotNone(tcp_diff, "Field diff for Total Contract Price was not derived in memory.")
        self.assertIn('1,000.00', str(tcp_diff.get('old_value')))
        self.assertIn('2,500.00', str(tcp_diff.get('new_value')))

    def test_save_staged_records_derives_diff_when_baseline_none(self):
        """Batch save of existing accounts without baselines must derive in-memory field diffs."""
        api = Api()
        api.dv_client = MagicMock()
        api.dv_client.save_loa_record.return_value = {'action': 'update', 'id': 'rec_2', 'etag': 'etag_2'}
        api.dv_client.save_log_record.return_value = {'id': 'log_saved_2'}

        orig_record = {
            '__id': 'rec_2',
            '__etag': 'etag_1',
            'cr63f_control1': 'CTRL-200',
            'cr63f_areaindex': 'AI-200',
            'cr63f_totalcontractprice': 3000.0,
            'Total Contract Price': '₱3,000.00'
        }
        api.loa_records = [dict(orig_record)]
        api.log_records = []

        with patch.object(api, '_write_baseline_if_missing', return_value=None):
            res = api.save_staged_records([{
                'id': 'rec_2',
                'fields': {'cr63f_totalcontractprice': 5000.0}
            }])

        self.assertTrue(res.get('success'), f"Expected success but got error: {res.get('error')}")
        self.assertGreater(len(api.log_records), 0)
        tcp_diff = next((l for l in api.log_records if l.get('field_name') == 'Total Contract Price'), None)
        self.assertIsNotNone(tcp_diff, "Field diff for Total Contract Price was not derived in staged save.")
        self.assertIn('3,000.00', str(tcp_diff.get('old_value')))
        self.assertIn('5,000.00', str(tcp_diff.get('new_value')))

    def test_revert_tcp_to_blank_clears_balance_and_paid(self):
        """Reverting TCP back to blank must clear Balance and % Paid instead of setting 0."""
        api = Api()
        api.dv_client = MagicMock()
        api.dv_client.config = {'loa_table': 'cr63f_loamonitoring', 'logs_table': 'cr63f_loamonitoringlogs'}
        api.dv_client.get_record.return_value = {
            '__id': 'rec_3',
            '__etag': 'etag_1',
            'cr63f_control1': 'CTRL-300',
            'cr63f_totalcontractprice': '1000000.00',
            'cr63f_balance': '1000000.00',
            'cr63f_paid': '0.0'
        }
        api.dv_client.save_loa_record.return_value = {'status': 200, 'etag': 'etag_2'}
        api.dv_client.save_log_record.return_value = {'id': 'log_saved_3'}

        api.loa_records = [{
            '__id': 'rec_3',
            '__etag': 'etag_1',
            'cr63f_control1': 'CTRL-300',
            'cr63f_totalcontractprice': '1000000.00',
            'cr63f_balance': '1000000.00',
            'cr63f_paid': '0.0'
        }]
        target_log = {
            '__id': 'log_rev_1',
            'cr63f_control1': 'CTRL-300',
            'field_name': 'Total Contract Price',
            'old_value': '-',
            'new_value': '₱1,000,000.00'
        }
        api.log_records = [target_log]

        res = api.revert_log_record('log_rev_1')
        self.assertTrue(res.get('success'), f"Expected success but got: {res.get('error')}")

        saved_payload = api.dv_client.save_loa_record.call_args[0][0]
        self.assertIsNone(saved_payload.get('cr63f_totalcontractprice'))
        self.assertIsNone(saved_payload.get('cr63f_balance'), "Balance should be cleared (None) when TCP is cleared.")
        self.assertIsNone(saved_payload.get('cr63f_paid'), "% Paid should be cleared (None) when TCP is cleared.")

    def test_save_record_exact_attribute_match_priority(self):
        """Numbered attributes (e.g. control2 vs control1) must match exact clean name first."""
        client = DataverseClient.__new__(DataverseClient)
        client._entity_info_cache = {}
        client._meta_cache = {}
        client.get_entity_info = MagicMock(return_value={'entity_set': 'tests', 'primary_id': 'testid'})
        client.get_entity_attributes = MagicMock(return_value={
            'cr63f_control1': {'logical': 'cr63f_control1', 'display': 'Control 1', 'type': 'String', 'can_create': True, 'can_update': True},
            'cr63f_control2': {'logical': 'cr63f_control2', 'display': 'Control 2', 'type': 'String', 'can_create': True, 'can_update': True},
        })
        client.patch_with_version = MagicMock(return_value=(200, 'etag_new'))

        # Saving 'control2' should resolve to cr63f_control2, NOT cr63f_control1
        client.save_record('test_table', {'control2': 'VAL_2'}, record_id='123')
        payload = client.patch_with_version.call_args[0][2]
        self.assertIn('cr63f_control2', payload)
        self.assertNotIn('cr63f_control1', payload)

    def test_clean_dataverse_datetime_and_work_week_dots(self):
        """Dates with dot separators (10.15.2026, 2026.10.15) should be supported."""
        self.assertEqual(_clean_dataverse_datetime('10.15.2026'), '2026-10-15')
        self.assertEqual(_clean_dataverse_datetime('2026.10.15'), '2026-10-15')
        self.assertEqual(_clean_dataverse_datetime('15.10.2026'), '2026-10-15')
        ww = get_work_week('10.15.2026')
        self.assertTrue(ww.startswith('WW-'), f"Expected valid work week but got: {ww}")

    def test_delete_loa_record_entry_not_found(self):
        """Deleting a non-existent account should return clean error without TypeError."""
        api = Api()
        api.dv_client = MagicMock()
        api.dv_client.config = {'loa_table': 'cr63f_loamonitoring'}
        api.dv_client.get_record.return_value = None
        api.loa_records = []

        res = api.delete_loa_record_entry('missing_guid')
        self.assertFalse(res.get('success'))
        self.assertIn('missing_guid', res.get('error', ''))
        self.assertIn('was not found', res.get('error', ''))

    def test_patch_with_version_etag_header_fallback(self):
        """When body omits ETag (e.g. 204 or header-only), ETag header must be captured."""
        client = DataverseClient.__new__(DataverseClient)
        client.api_root = 'https://fake.crm.dynamics.com/api/data/v9.2/'
        client._headers = MagicMock(return_value={})
        mock_resp = MagicMock()
        mock_resp.status_code = 204
        mock_resp.headers = {'ETag': 'W/"header_etag_123"'}
        client._send = MagicMock(return_value=mock_resp)

        status, new_etag = client.patch_with_version('tests', 'guid1', {'name': 'test'})
        self.assertEqual(status, 204)
        self.assertEqual(new_etag, 'W/"header_etag_123"')

    def test_clean_title_variants_sorted(self):
        """clean_title_variants should return a sorted list for deterministic indexing."""
        variants = clean_title_variants('TCT-066-2019006115')
        self.assertIsInstance(variants, list)
        self.assertEqual(variants, sorted(variants))

    def test_delete_finance_record(self):
        """delete_finance_record on DataverseClient should call delete on the finance table."""
        client = DataverseClient.__new__(DataverseClient)
        client.config = {'finance_table': 'cr63f_financereference'}
        client.get_entity_info = MagicMock(return_value={'entity_set': 'cr63f_financereferences'})
        client.delete = MagicMock(return_value=204)

        res = client.delete_finance_record('fin_1')
        self.assertTrue(res.get('success'))
        client.delete.assert_called_once_with('cr63f_financereferences', 'fin_1')

    def test_export_records_excel_retains_custom_columns(self):
        """Export without explicit column selection must retain custom columns without display alias."""
        api = Api()
        api._window = MagicMock()
        with patch('app._dialog_path', return_value='test_export.xlsx'), \
             patch('app.write_excel_file') as mock_write:
            records = [{
                '__id': '1',
                'Control Number': 'CTRL-1',
                'cr63f_totalcontractprice': 1000,
                'Total Contract Price': '₱1,000.00',
                'cr63f_customnotes': 'Special Note 1',
            }]
            res = api.export_records_excel(records)
            self.assertTrue(res.get('success'))
            cols_exported = mock_write.call_args[0][1]
            self.assertIn('cr63f_customnotes', cols_exported)
            self.assertIn('Total Contract Price', cols_exported)
            # Raw cr63f_totalcontractprice should be omitted because Total Contract Price alias exists
            self.assertNotIn('cr63f_totalcontractprice', cols_exported)


if __name__ == '__main__':
    unittest.main()

