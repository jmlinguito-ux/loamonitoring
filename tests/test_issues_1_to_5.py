"""Regression coverage for reference audit, account identity, and revert concurrency."""
import unittest
from unittest.mock import MagicMock

from app import Api, build_log_entry, derive_field_diffs_from_snapshots
from dataverse import DataverseClient


def api():
    instance = Api.__new__(Api)
    instance.dv_client = MagicMock()
    instance.dv_client.config = {}
    instance.dv_client.fetch_records.return_value = []
    instance.dv_client.get_entity_attributes.return_value = {
        'cr63f_details': {'can_create': True},
    }
    instance.dv_client.save_log_record.return_value = {'success': True, 'id': 'log-saved'}
    instance.dv_client.save_loa_record.return_value = {'success': True, 'id': 'm1'}
    instance.dv_client.save_finance_record.return_value = {'success': True, 'id': 'f1'}
    instance.dv_client.save_conversion_record.return_value = {'success': True, 'id': 'c1'}
    instance.loa_records = []
    instance.log_records = []
    instance.pending_session_logs = []
    instance.finance_rows = []
    instance.conversion_rows = []
    instance.finance_engine = MagicMock()
    instance.conversion_engine = MagicMock()
    instance._window = None
    return instance


def snapshot(master, action='MANUAL_EDIT', timestamp='10/08/2026 10:00 AM'):
    return build_log_entry(action, 'WW-41', timestamp, 'AREA-1', '', '',
                           'Full Record', 'Previous', 'Updated', full_record=master)


class ReferenceSaveTests(unittest.TestCase):
    def test_inline_edits_retain_original_values_until_source_and_audit_are_saved(self):
        for kind in ('finance', 'conversion'):
            with self.subTest(kind=kind):
                app = api()
                rows = getattr(app, f'{kind}_rows')
                rows.append({'__id': 'r1', 'Area Index': 'A1', 'Total Contract Price': 100})
                update = getattr(app, f'update_{kind}_cell')
                self.assertTrue(update('r1', 'Total Contract Price', 150)['success'])
                self.assertTrue(update('r1', 'Total Contract Price', 200)['success'])
                result = getattr(app, f'save_staged_{kind}_records')([
                    {'id': 'r1', 'fields': {'Total Contract Price': 200}}])
                self.assertTrue(result['success'], result)
                getattr(app.dv_client, f'save_{kind}_record').assert_called_once_with(
                    {'Total Contract Price': 200}, record_id='r1')
                log = app.dv_client.save_log_record.call_args.args[0]
                self.assertEqual(log['old_value'], '₱100.00')
                self.assertEqual(log['new_value'], '₱200.00')

    def test_reference_log_revert_is_rejected_for_both_sources(self):
        for action in ('FINANCE_UPDATE', 'CONVERSION_UPDATE'):
            with self.subTest(action=action):
                app = api()
                app.log_records = [{'__id': 'l1', 'cr63f_name': f'AREA - TCP ({action})',
                                    'field_name': 'Total Contract Price', 'old_value': '100', 'new_value': '200'}]
                result = app.revert_log_record('l1')
                self.assertFalse(result['success'])
                app.dv_client.save_loa_record.assert_not_called()
                app.dv_client.get_record.assert_not_called()

    def test_legacy_reference_details_still_identify_the_source_table(self):
        app = api()
        app.log_records = [{'__id': 'l1', 'cr63f_details': 'Updated Finance Reference Total Contract Price from 100 to 200',
                            'field_name': 'Total Contract Price', 'old_value': 100, 'new_value': 200}]
        self.assertFalse(app.revert_log_record('l1')['success'])
        app.dv_client.save_loa_record.assert_not_called()

    def test_audit_failure_rolls_back_reference_and_displays_no_phantom_log(self):
        for kind in ('finance', 'conversion'):
            with self.subTest(kind=kind):
                app = api()
                rows = getattr(app, f'{kind}_rows')
                rows.append({'__id': 'r1', 'Area Index': 'AREA-1', 'Total Contract Price': '100'})
                app.dv_client.save_log_record.side_effect = RuntimeError('audit down')
                app.dv_client.get_record.return_value = {
                    '__id': 'r1', '__etag': 'W/"2"', 'Total Contract Price': '200'}
                save = getattr(app, f'save_staged_{kind}_records')
                result = save([{'id': 'r1', 'fields': {'Total Contract Price': '200'}}])
                self.assertFalse(result['success'])
                self.assertEqual(result['saved_count'], 0)
                self.assertEqual(result['failed_ids'], ['r1'])
                self.assertEqual(rows[0]['Total Contract Price'], '100')
                self.assertEqual(app.log_records, [])
                self.assertEqual(app.pending_session_logs, [])
                rollback = app.dv_client.save_record.call_args
                self.assertEqual(rollback.args[1], {'Total Contract Price': '100'})
                self.assertEqual(rollback.kwargs['expected_etag'], 'W/"2"')

    def test_partial_success_returns_only_confirmed_ids_and_persisted_logs(self):
        app = api()
        app.finance_rows = [{'__id': f'f{i}', 'Area Index': f'A{i}', 'Total Contract Price': '100'} for i in (1, 2)]
        app.dv_client.save_finance_record.side_effect = [
            {'success': True, 'id': 'f1'}, RuntimeError('reference down')]
        result = app.save_staged_finance_records([
            {'id': f'f{i}', 'fields': {'Total Contract Price': '200'}} for i in (1, 2)])
        self.assertFalse(result['success'])
        self.assertEqual(result['saved_ids'], ['f1'])
        self.assertEqual(result['failed_ids'], ['f2'])
        self.assertEqual(result['saved_count'], 1)
        self.assertEqual(result['logged_count'], 1)
        self.assertEqual(app.log_records[0]['__id'], 'log-saved')
        self.assertEqual(app.finance_rows[1]['Total Contract Price'], '100')

    def test_partial_field_audit_failure_cleans_up_confirmed_audit_row(self):
        app = api()
        app.finance_rows = [{'__id': 'f1', 'Area Index': 'A1', 'Total Contract Price': 0, 'Balance': 0}]
        app.dv_client.save_log_record.side_effect = [
            {'success': True, 'id': '11111111-2222-3333-4444-555555555555'}, RuntimeError('second audit failed')]
        app.dv_client.get_record.return_value = {'__etag': 'W/"2"', 'Total Contract Price': 200, 'Balance': 100}
        result = app.save_staged_finance_records([{'id': 'f1', 'fields': {'Total Contract Price': 200, 'Balance': 100}}])
        self.assertFalse(result['success'])
        self.assertEqual(app.dv_client.save_record.call_args.args[1], {'Total Contract Price': 0, 'Balance': 0})
        app.dv_client.delete_log_record.assert_called_once_with('11111111-2222-3333-4444-555555555555')
        self.assertEqual(app.log_records, [])

    def test_new_reference_is_deleted_if_audit_creation_fails(self):
        app = api()
        app.finance_rows = [{'Area Index': 'A1', 'Total Contract Price': 100}]
        app.dv_client.get_entity_info.return_value = {'entity_set': 'finances'}
        app.dv_client.save_log_record.side_effect = RuntimeError('audit down')
        result = app.save_staged_finance_records([{'id': '0', 'fields': {'Total Contract Price': 200}}])
        self.assertFalse(result['success'])
        app.dv_client.delete.assert_called_once_with('finances', 'f1')
        self.assertNotIn('__id', app.finance_rows[0])

    def test_failed_rollback_blocks_blind_retry(self):
        app = api()
        app.finance_rows = [{'__id': 'f1', 'Area Index': 'A1', 'Total Contract Price': 100}]
        app.dv_client.save_log_record.side_effect = RuntimeError('audit down')
        app.dv_client.get_record.return_value = {'__etag': 'W/"2"', 'Total Contract Price': 300}
        items = [{'id': 'f1', 'fields': {'Total Contract Price': 200}}]
        first = app.save_staged_finance_records(items)
        self.assertIn('Check Dataverse', first['error'])
        app.dv_client.save_record.assert_not_called()
        second = app.save_staged_finance_records(items)
        self.assertFalse(second['success'])
        self.assertEqual(app.dv_client.save_finance_record.call_count, 1)


class AccountIdentityTests(unittest.TestCase):
    def test_loaded_duplicate_control_is_rejected_before_any_write(self):
        app = api()
        app.loa_records = [
            {'__id': 'm1', 'cr63f_control1': 'C1'}, {'__id': 'm2', 'cr63f_control1': 'C2'}]
        result = app.save_staged_records([{'id': 'm2', 'fields': {'cr63f_control1': 'c1'}}])
        self.assertFalse(result['success'])
        self.assertIn('already belongs', result['error'])
        app.dv_client.save_loa_record.assert_not_called()
        app.dv_client.save_log_record.assert_not_called()

    def test_server_duplicate_is_checked_even_when_loaded_rows_are_incomplete(self):
        app = api()
        original = {'__id': 'm1', 'cr63f_control1': 'C1'}
        app.loa_records = [original]
        app.dv_client.fetch_records.return_value = [{'__id': 'm2', 'cr63f_control1': 'C2'}]
        with self.assertRaisesRegex(ValueError, 'already belongs'):
            app._validate_control_change(original, {'Control Number': 'C2'}, 'm1')

    def test_batch_reservations_prevent_duplicate_new_controls(self):
        app = api()
        app.loa_records = [{'__id': 'm1', 'cr63f_control1': 'C1'}, {'__id': 'm2', 'cr63f_control1': 'C2'}]
        reservations = {}
        app._validate_control_change(app.loa_records[0], {'cr63f_control1': 'NEW'}, 'm1', reservations)
        with self.assertRaisesRegex(ValueError, 'another staged'):
            app._validate_control_change(app.loa_records[1], {'cr63f_control1': 'new'}, 'm2', reservations)

    def test_blank_control_and_unavailable_audit_identity_storage_are_rejected(self):
        app = api()
        original = {'__id': 'm1', 'cr63f_control1': 'C1'}
        with self.assertRaisesRegex(ValueError, 'cannot be cleared'):
            app._validate_control_change(original, {'cr63f_control1': ''}, 'm1')
        app.dv_client.get_entity_attributes.return_value = {}
        with self.assertRaisesRegex(ValueError, 'audit column'):
            app._validate_control_change(original, {'cr63f_control1': 'C2'}, 'm1')

    def test_control_rename_stays_in_one_audit_history_after_reload(self):
        original = {'__id': 'm1', 'cr63f_control1': 'C1', 'new_mode': 'SALE'}
        baseline = snapshot(original, 'BASELINE', '10/08/2026 09:00 AM')
        updated = snapshot({**original, 'cr63f_control1': 'C3', 'new_mode': 'LEASE'})
        # The actual logs table may retain only Details, not Name or internal helpers.
        for row in (baseline, updated):
            for key in ('__master_record_id', 'cr63f_loamonitoringid', 'cr63f_name', 'name',
                        'action_type', 'Action Type', 'cr63f_actiontype'):
                row.pop(key, None)
        logs = derive_field_diffs_from_snapshots([baseline, updated])
        self.assertEqual({log['field_name'] for log in logs}, {'Control Number', 'Mode'})
        self.assertTrue(all(log['__master_record_id'] == 'm1' for log in logs))

    def test_rename_creates_identity_baseline_when_legacy_control_history_exists(self):
        app = api()
        record = {'__id': 'm1', 'cr63f_control1': 'C1', 'new_mode': 'SALE'}
        app.dv_client.fetch_records.return_value = [{'cr63f_control': 'C1', 'action_type': 'MANUAL_EDIT'}]
        baseline = app._write_baseline_if_missing(record, '10/08/2026 09:00 AM', require_master_identity=True)
        self.assertIsNotNone(baseline)
        self.assertIn('[MASTER:m1]', baseline['cr63f_details'])

    def test_shared_area_and_title_do_not_skip_another_accounts_baseline(self):
        app = api()
        first = {'__id': 'm1', 'cr63f_control1': 'C1', 'cr63f_areaindex': 'AREA', 'cr63f_titleno': 'TITLE'}
        app.dv_client.fetch_records.return_value = [snapshot(first)]
        known = app._logged_account_keys()
        app.pending_session_logs = [snapshot(first)]
        second = {**first, '__id': 'm2', 'cr63f_control1': 'C2'}
        self.assertIsNotNone(app._write_baseline_if_missing(second, '10/08/2026 10:00 AM', known))

    def test_reference_only_logs_are_not_master_baselines(self):
        app = api()
        reference = {'__id': 'f1', 'cr63f_control1': 'C1', 'action_type': 'FINANCE_UPDATE'}
        app.dv_client.fetch_records.return_value = [reference]
        app.pending_session_logs = [reference]
        self.assertEqual(app._logged_account_keys(), set())
        self.assertIsNotNone(app._write_baseline_if_missing(
            {'__id': 'm1', 'cr63f_control1': 'C1'}, '10/08/2026 10:00 AM'))


class RevertConcurrencyTests(unittest.TestCase):
    def make_app(self):
        app = api()
        master = {'__id': 'm1', 'cr63f_control1': 'C1', 'cr63f_totalcontractprice': 200,
                  'cr63f_totalpaymentdeductibletotcp': 0, 'cr63f_balance': 200, 'cr63f_paid': 0}
        app.loa_records = [master]
        log = build_log_entry('MANUAL_EDIT', 'WW-41', '10/08/2026 09:00 AM', 'A1', '', '',
                              'Total Contract Price', 100, 200, full_record=master)
        log['__id'] = 'l1'
        app.log_records = [log]
        app.dv_client.get_record.return_value = {**master, '__etag': 'W/"10"'}
        return app

    def test_newer_server_value_blocks_revert_despite_matching_cached_value(self):
        app = self.make_app()
        app.dv_client.get_record.return_value['cr63f_totalcontractprice'] = 300
        result = app.revert_log_record('l1')
        self.assertFalse(result['success'])
        self.assertIn('changed since', result['error'])
        app.dv_client.save_loa_record.assert_not_called()

    def test_revert_sends_fresh_etag_and_not_cached_version(self):
        app = self.make_app()
        result = app.revert_log_record('l1')
        self.assertTrue(result['success'], result)
        self.assertEqual(app.dv_client.save_loa_record.call_args.kwargs['expected_etag'], 'W/"10"')

    def test_even_small_currency_changes_block_stale_revert(self):
        app = self.make_app()
        app.dv_client.get_record.return_value['cr63f_totalcontractprice'] = '200.001'
        self.assertFalse(app.revert_log_record('l1')['success'])
        app.dv_client.save_loa_record.assert_not_called()

    def test_date_revert_compares_server_iso_date_with_audit_display_date(self):
        app = self.make_app()
        app.log_records[0].update(field_name='Endorsed Date', old_value='10/01/2026', new_value='10/08/2026')
        app.dv_client.get_record.return_value['cr63f_endorseddate'] = '2026-10-08T00:00:00Z'
        result = app.revert_log_record('l1')
        self.assertTrue(result['success'], result)
        self.assertEqual(app.dv_client.save_loa_record.call_args.args[0], {'cr63f_endorseddate': '10/01/2026'})

    def test_concurrent_patch_conflict_preserves_cache_and_writes_no_revert_log(self):
        app = self.make_app()
        app.dv_client.save_loa_record.side_effect = RuntimeError('record changed during save')
        result = app.revert_log_record('l1')
        self.assertFalse(result['success'])
        self.assertEqual(app.loa_records[0]['cr63f_totalcontractprice'], 200)
        app.dv_client.save_log_record.assert_not_called()

    def test_audit_failure_does_not_roll_back_over_newer_server_change(self):
        app = self.make_app()
        app.dv_client.get_record.side_effect = [
            {**app.loa_records[0], '__etag': 'W/"10"'},
            {**app.loa_records[0], 'cr63f_totalcontractprice': 300, '__etag': 'W/"12"'}]
        app.dv_client.save_log_record.side_effect = RuntimeError('audit down')
        result = app.revert_log_record('l1')
        self.assertFalse(result['success'])
        self.assertIn('rollback also failed', result['error'])
        self.assertEqual(app.dv_client.save_loa_record.call_count, 1)

    def test_missing_record_version_and_failed_fresh_read_block_writes(self):
        app = self.make_app()
        app.dv_client.get_record.return_value.pop('__etag')
        self.assertFalse(app.revert_log_record('l1')['success'])
        app.dv_client.get_record.side_effect = RuntimeError('offline')
        self.assertFalse(app.revert_log_record('l1')['success'])
        app.dv_client.save_loa_record.assert_not_called()


class TransportConcurrencyTests(unittest.TestCase):
    def test_conditional_patch_sends_if_match_and_reports_412(self):
        client = DataverseClient.__new__(DataverseClient)
        client.api_root = 'https://example/api/'
        client._headers = MagicMock(return_value={})
        client._session = MagicMock()
        client._session.patch.return_value = MagicMock(status_code=412)
        with self.assertRaisesRegex(RuntimeError, 'changed in Dataverse'):
            client.patch('masters', 'm1', {'x': 1}, expected_etag='W/"10"')
        self.assertEqual(client._session.patch.call_args.kwargs['headers']['If-Match'], 'W/"10"')
        self.assertEqual(client._session.patch.call_count, 1)

    def test_fresh_record_keeps_precision_version_and_choice_labels(self):
        client = DataverseClient.__new__(DataverseClient)
        client.get_entity_info = MagicMock(return_value={'entity_set': 'masters', 'primary_id': 'masterid'})
        client.get_entity_attributes = MagicMock(return_value={
            'money': {'display': 'Money', 'type': 'Money'}, 'mode': {'display': 'Mode', 'type': 'Picklist'}})
        client.get = MagicMock(return_value={
            'masterid': 'm1', '@odata.etag': 'W/"10"', 'money': 1.234,
            'money@OData.Community.Display.V1.FormattedValue': '$1.23', 'mode': 7,
            'mode@OData.Community.Display.V1.FormattedValue': 'SALE'})
        record = client.get_record('master', 'm1')
        self.assertEqual(record['money'], 1.234)
        self.assertEqual(record['mode'], 'SALE')
        self.assertEqual(record['__etag'], 'W/"10"')


if __name__ == '__main__':
    unittest.main()
