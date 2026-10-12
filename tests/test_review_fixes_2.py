"""Regression tests for the remaining 2026-10-07 review items (6-21)."""

import unittest
from datetime import datetime
from unittest.mock import MagicMock, patch

from app import Api, _odata_literal
from loa_engine import ConversionLookupEngine, FinanceLookupEngine, parse_float


def _api():
    with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
        app = Api()
    app.loa_records = []
    app.log_records = []
    return app


def _mock_dv():
    dv = MagicMock()
    dv.config = {}
    dv.fetch_records.return_value = []
    dv.save_loa_record.return_value = {'success': True, 'action': 'update', 'id': 'rec'}
    dv.save_log_record.return_value = {'success': True, 'action': 'create', 'id': 'log-1'}
    return dv


class TestNumberParsing(unittest.TestCase):
    def test_accounting_negative_and_scientific(self):
        self.assertEqual(parse_float('(1,234.00)'), -1234.0)
        self.assertEqual(parse_float('₱(500.50)'), -500.5)
        self.assertEqual(parse_float('1.2E+05'), 120000.0)
        self.assertEqual(parse_float('₱1,234.00'), 1234.0)
        self.assertEqual(parse_float('-75'), -75.0)


class TestLookupDuplicates(unittest.TestCase):
    def test_conflicting_area_index_rows_are_not_guessed(self):
        engine = FinanceLookupEngine([
            {'TYPE': 'SALE', 'Area Index': 'AI-1', 'Total Contract Price': '100', 'Total Payment Deductible to TCP': '10'},
            {'TYPE': 'SALE', 'Area Index': 'AI-1', 'Total Contract Price': '200', 'Total Payment Deductible to TCP': '10'},
            {'TYPE': 'SALE', 'Title Mod': 'T-5', 'Total Contract Price': '999'},
        ])
        res = engine.lookup('SALE', '', 'AI-1', 'T-5')
        self.assertFalse(res['matched'])
        self.assertTrue(res['ambiguous'])
        self.assertEqual(res['ambiguous_count'], 2)

    def test_identical_duplicates_still_match(self):
        row = {'TYPE': 'SALE', 'Area Index': 'AI-1', 'Total Contract Price': '100', 'Total Payment Deductible to TCP': '10'}
        engine = FinanceLookupEngine([row, dict(row)])
        self.assertTrue(engine.lookup('SALE', '', 'AI-1', '')['matched'])

    def test_conflicting_conversion_rows(self):
        engine = ConversionLookupEngine([
            {'cr63f_areaindex': 'AI-1', 'cr63f_darstatus': 'DAR SECURED'},
            {'cr63f_areaindex': 'AI-1', 'cr63f_darstatus': 'DAR FOR APPLICATION'},
        ])
        res = engine.lookup('AI-1')
        self.assertFalse(res['matched'])
        self.assertTrue(res['ambiguous'])

    def test_non_row_account_never_takes_row_finance_by_title(self):
        engine = FinanceLookupEngine([
            {'TYPE': 'ROW', 'Corridor Index': 'C-1', 'Title Mod': 'T-12345', 'Total Contract Price': '500'},
        ])
        self.assertFalse(engine.lookup('SALE', '', 'NO-MATCH', 'T-12345')['matched'])
        self.assertTrue(engine.lookup('ROW', 'NO-MATCH', '', 'T-12345')['matched'])

    def test_title_fallback_rejects_conflicting_number_segments(self):
        engine = FinanceLookupEngine([
            {'TYPE': 'SALE', 'Title Mod': 'TCT-066-2019006115', 'Total Contract Price': '300'},
        ])
        self.assertFalse(engine.lookup('SALE', '', 'X', 'TCT-999-2019006115')['matched'])
        self.assertTrue(engine.lookup('SALE', '', 'X', '2019006115')['matched'])

    def test_update_from_references_reports_conflicts(self):
        app = _api()
        rows = [
            {'TYPE': 'SALE', 'Area Index': 'AI-1', 'Total Contract Price': '100', 'Total Payment Deductible to TCP': '10'},
            {'TYPE': 'SALE', 'Area Index': 'AI-1', 'Total Contract Price': '200', 'Total Payment Deductible to TCP': '10'},
        ]
        app.finance_rows = rows
        app.finance_engine.load_rows(rows)
        res = app.populate_loa_records_from_references(records=[{'__id': 'r1', 'cr63f_type': 'SALE', 'cr63f_areaindex': 'AI-1'}])
        self.assertEqual(res['updated_count'], 0)
        self.assertEqual(len(res['ambiguous_accounts']), 1)
        self.assertIn('AI-1', res['warning'])


class TestMasterLookup(unittest.TestCase):
    def test_record_id_beats_earlier_area_index_match(self):
        app = _api()
        app.loa_records = [
            {'__id': 'other', 'cr63f_areaindex': 'AI-1'},
            {'__id': 'target', 'cr63f_areaindex': 'AI-1'},
        ]
        self.assertEqual(app._find_or_fetch_master_record(record_id='target', area_index='AI-1')['__id'], 'target')

    def test_ambiguous_area_index_returns_nothing(self):
        app = _api()
        app.loa_records = [{'__id': 'a', 'cr63f_areaindex': 'AI-1'}, {'__id': 'b', 'cr63f_areaindex': 'AI-1'}]
        self.assertEqual(app._find_or_fetch_master_record(area_index='AI-1'), {})

    def test_odata_literal_escapes_quotes(self):
        self.assertEqual(_odata_literal("O'Brien"), "'O''Brien'")


class TestSavesKeepEndorsedWeek(unittest.TestCase):
    def test_staged_save_keeps_endorsed_date_week(self):
        app = _api()
        app.dv_client = _mock_dv()
        app.loa_records = [{'__id': 'r1', 'cr63f_areaindex': 'AI-1', 'cr63f_rccprice': '1',
                            'cr63f_endorseddate': '01/15/2026', 'work_week': 'WW-03'}]
        res = app.save_staged_records([{'id': 'r1', 'fields': {'cr63f_rccprice': '2'}}])
        self.assertTrue(res['success'], res)
        self.assertEqual(app.loa_records[0]['work_week'], 'WW-03')


class TestNewEntry(unittest.TestCase):
    def _data(self, **extra):
        data = {'cr63f_areaindex': 'AI-9', 'cr63f_control1': 'CTRL-9', 'cr63f_acquisitionsublevel': 'L1'}
        data.update(extra)
        return data

    # The three scenarios chosen on 2026-10-07 (existing account LIA-4527 / CTRL-100).
    def _app_with_existing(self):
        app = _api()
        app.dv_client = _mock_dv()
        app.dv_client.save_loa_record.return_value = {'success': True, 'action': 'create', 'id': 'new-id'}
        app.loa_records = [{'__id': 'x', 'cr63f_areaindex': 'LIA-4527', 'cr63f_control1': 'CTRL-100'}]
        return app

    def test_new_parcel_in_same_area_asks_then_saves(self):
        app = self._app_with_existing()
        data = self._data(cr63f_areaindex='lia-4527', cr63f_control1='CTRL-200')
        res = app.save_loa_record_entry(data)
        self.assertFalse(res['success'])
        self.assertTrue(res['needs_confirmation'])
        self.assertEqual(res['existing_controls'], ['CTRL-100'])
        self.assertIn('CTRL-100', res['message'])
        app.dv_client.save_loa_record.assert_not_called()
        res = app.save_loa_record_entry(data, None, True)
        self.assertTrue(res['success'], res)
        app.dv_client.save_loa_record.assert_called_once()

    def test_reused_control_number_is_blocked(self):
        app = self._app_with_existing()
        res = app.save_loa_record_entry(self._data(cr63f_areaindex='LIA-9000', cr63f_control1='ctrl-100'), None, True)
        self.assertFalse(res['success'])
        self.assertNotIn('needs_confirmation', res)
        self.assertIn('already exists', res['error'])
        app.dv_client.save_loa_record.assert_not_called()

    def test_blank_control_number_is_blocked(self):
        app = self._app_with_existing()
        res = app.save_loa_record_entry(self._data(cr63f_areaindex='LIA-4527', cr63f_control1=''))
        self.assertFalse(res['success'])
        self.assertIn('Control Number is required', res['error'])
        app.dv_client.save_loa_record.assert_not_called()

    def test_new_area_saves_without_prompt(self):
        app = self._app_with_existing()
        res = app.save_loa_record_entry(self._data(cr63f_areaindex='LIA-9001', cr63f_control1='CTRL-300'))
        self.assertTrue(res['success'], res)

    def test_failed_audit_row_removes_new_account(self):
        app = _api()
        app.dv_client = _mock_dv()
        app.loa_records = [{'__id': 'x', 'cr63f_areaindex': 'AI-1'}]
        app.dv_client.save_loa_record.return_value = {'success': True, 'action': 'create', 'id': 'new-id'}
        app.dv_client.save_log_record.side_effect = RuntimeError('logs table down')
        res = app.save_loa_record_entry(self._data())
        self.assertFalse(res['success'])
        app.dv_client.delete_loa_record.assert_called_once_with('new-id')
        self.assertNotIn('new-id', [r.get('__id') for r in app.loa_records])


class TestRevertSafety(unittest.TestCase):
    def _setup(self):
        app = _api()
        app.dv_client = _mock_dv()
        app.log_records = [{'__id': 'log-1', 'area_index': 'AI-1', 'field_name': 'Total Contract Price',
                            'old_value': '₱3,000,000.00', 'new_value': '₱5,000,000.00'}]
        app.loa_records = [{'__id': 'm1', 'cr63f_areaindex': 'AI-1', 'cr63f_totalcontractprice': '5000000',
                            'cr63f_totalpaymentdeductibletotcp': '1000000', 'cr63f_balance': '4000000',
                            'cr63f_paid': '20.00%'}]
        app.dv_client.get_record.return_value = {**app.loa_records[0], '__etag': 'W/"1"'}
        return app

    def test_failed_dataverse_save_leaves_record_unchanged(self):
        app = self._setup()
        app.dv_client.save_loa_record.side_effect = RuntimeError('offline')
        res = app.revert_log_record('log-1')
        self.assertFalse(res['success'])
        self.assertEqual(app.loa_records[0]['cr63f_totalcontractprice'], '5000000')
        self.assertEqual(app.loa_records[0]['cr63f_balance'], '4000000')

    def test_failed_audit_row_rolls_back_dependants(self):
        app = self._setup()
        app.dv_client.save_log_record.side_effect = RuntimeError('logs table down')
        app.dv_client.get_record.side_effect = [
            {**app.loa_records[0], '__etag': 'W/"1"'},
            {**app.loa_records[0], 'cr63f_totalcontractprice': '3000000.00',
             'cr63f_balance': 2000000, 'cr63f_paid': '33.3', '__etag': 'W/"2"'},
        ]
        res = app.revert_log_record('log-1')
        self.assertFalse(res['success'])
        rollback = app.dv_client.save_loa_record.call_args_list[-1][0][0]
        self.assertEqual(rollback['cr63f_totalcontractprice'], '5000000')
        self.assertEqual(rollback['cr63f_balance'], '4000000')
        self.assertEqual(rollback['cr63f_paid'], '20.00%')
        self.assertEqual(app.loa_records[0]['cr63f_totalcontractprice'], '5000000')

    def test_success_updates_loaded_record(self):
        app = self._setup()
        res = app.revert_log_record('log-1')
        self.assertTrue(res['success'], res)
        self.assertEqual(app.loa_records[0]['cr63f_totalcontractprice'], '3000000.00')

    def test_unsafe_cell_revert_is_gone(self):
        self.assertFalse(hasattr(Api, 'revert_cell_value'))


class TestOverviewNumbers(unittest.TestCase):
    def test_portfolio_balance_is_per_account(self):
        app = _api()
        app.loa_records = [
            {'cr63f_totalcontractprice': '100', 'cr63f_totalpaymentdeductibletotcp': '150'},  # overpaid: 0
            {'cr63f_totalcontractprice': '100', 'cr63f_totalpaymentdeductibletotcp': '20'},   # 80
            {'cr63f_totalcontractprice': '0', 'cr63f_totalpaymentdeductibletotcp': '5'},      # 100% paid
        ]
        res = app.get_dashboard_analytics()
        self.assertEqual(res['total_balance'], 80.0)
        self.assertEqual(res['financial_breakdown']['fully_paid_count'], 2)

    def test_previous_week_is_the_calendar_week(self):
        app = _api()
        app.log_records = [
            {'__id': f'l{i}', 'action_type': 'MANUAL_EDIT', 'cr63f_control1': 'C1', 'field_name': 'Mode',
             'old_value': 'A', 'new_value': 'B', 'log_time': t}
            for i, t in enumerate(['09/10/2026 10:00 AM', '10/01/2026 10:00 AM'])
        ]
        res = app.get_dashboard_analytics(work_week='ALL')
        self.assertEqual(res['selected_work_week'], '2026 · WW-40')
        self.assertEqual(res['comparison_previous_week'], '2026 · WW-39')

    def test_status_counts_use_land_master_only(self):
        app = _api()
        app.conversion_rows = [{'cr63f_areaindex': 'AI-1', 'cr63f_darstatus': 'DAR SECURED'}]
        app.conversion_engine.load_rows(app.conversion_rows)
        app.loa_records = [{'cr63f_areaindex': 'AI-1', 'cr63f_conversionstatus': ''}]
        res = app.get_dashboard_analytics()
        self.assertEqual(res['conversion_metrics']['secured_count'], 0)


class TestDataverseClient(unittest.TestCase):
    def _client(self):
        import dataverse as dv
        client = dv.DataverseClient.__new__(dv.DataverseClient)
        client.config = {}
        client._meta_cache = {}
        client._datetime_behavior_cache = {'t': {}}
        client.get_entity_info = lambda table: {'entity_set': 'ts', 'primary_name': 'name'}
        client.get_entity_attributes = lambda table: {
            'new_mode': {'logical': 'new_mode', 'display': 'Mode', 'type': 'Picklist', 'can_create': True, 'can_update': True},
        }
        client.get_picklist_options = lambda table: {'new_mode': [
            {'value': 1, 'label': 'SUBLEASE'}, {'value': 2, 'label': 'LEASE'}, {'value': 3, 'label': 'SALE - DIRECT'},
            {'value': 4, 'label': 'SALE - BROKER'},
        ]}
        client.patch_with_version = MagicMock(return_value=(204, ''))
        client.patch = client.patch_with_version
        return client

    def test_choice_label_matches_exactly(self):
        client = self._client()
        client.save_record('t', {'new_mode': 'LEASE'}, record_id='r1')
        self.assertEqual(client.patch.call_args[0][2]['new_mode'], 2)

    def test_ambiguous_partial_choice_is_not_guessed(self):
        # 'SALE' matches two choices, so it is refused rather than guessed or dropped silently.
        client = self._client()
        with self.assertRaises(ValueError):
            client.save_record('t', {'new_mode': 'SALE'}, record_id='r1')
        client.patch.assert_not_called()


if __name__ == '__main__':
    unittest.main()
