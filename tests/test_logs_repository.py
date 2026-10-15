"""Unit tests for LOA Monitoring Logs repository routing, diff tracking, and analytics."""

from datetime import datetime
from decimal import Decimal

import json
import unittest
from unittest.mock import MagicMock, patch
from app import Api, _is_different_pct
import app as app_module
import dataverse as dv


class TestLogsRepository(unittest.TestCase):
    def setUp(self):
        self._test_cache = []
        self._test_tombstones = set()
        self.patch_save = patch('app.save_cached_logs', side_effect=lambda logs: setattr(self, '_test_cache', list(logs)))
        self.patch_load = patch('app.load_cached_logs', side_effect=lambda: list(getattr(self, '_test_cache', [])))
        self.patch_save.start()
        self.patch_load.start()

        def _mock_save_tomb(entry):
            a_idx = str(entry.get('area_index') or entry.get('cr63f_areaindex') or '').strip().upper()
            f_name = str(entry.get('field_name') or entry.get('cr63f_fieldchanged') or '').strip()
            n_val = str(entry.get('new_value') or entry.get('cr63f_newvalue') or '').strip()
            t_str = str(entry.get('log_time') or entry.get('createdon') or '').strip()
            log_id = str(entry.get('__id') or entry.get('id') or '').strip()
            if log_id:
                self._test_tombstones.add(f"ID:{log_id}")
            if a_idx and f_name:
                self._test_tombstones.add(f"EVT:{a_idx}|{f_name}|{n_val}|{t_str}")
                self._test_tombstones.add(f"FIELD:{a_idx}|{f_name}|{n_val}")

        self.patch_save_tomb = patch('app.save_tombstoned_signature', side_effect=_mock_save_tomb)
        self.patch_load_tomb = patch('app.load_tombstoned_signatures', side_effect=lambda: set(self._test_tombstones))
        self.patch_save_tomb.start()
        self.patch_load_tomb.start()
        self.addCleanup(patch.stopall)

        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            self.app = Api()
        self.app.log_records = []
        # Mock Dataverse client to test routing without live network
        self.mock_dv = MagicMock()
        self.mock_dv.fetch_records.return_value = []
        self.mock_dv.config = {
            'loa_table': 'cr63f_loamonitoring',
            'finance_table': 'cr63f_financereference',
            'logs_table': 'cr63f_loamonitoringlogs',
            'environment_url': 'https://yourorg.crm.dynamics.com',
            'tenant_id': 'test-tenant',
            'client_id': 'test-client',
            'solution_name': 'LOA MONITORING',
            'solution_id': 'test-sol-id'
        }
        self.app.dv_client = self.mock_dv
        self.mock_dv.get_record.side_effect = lambda table, record_id: {
            **next(row for row in self.app.loa_records if row['__id'] == record_id),
            '__etag': 'W/"1"',
        }

    def test_default_config_has_logs_table(self):
        cfg = dv.load_config()
        self.assertEqual(cfg.get('logs_table'), 'cr63f_loamonitoringlogs')
        self.assertEqual(cfg.get('loa_table'), 'cr63f_loamonitoring')
        self.assertEqual(cfg.get('finance_table'), 'cr63f_financereference')

    def test_action_type_schema_migration_adds_only_approved_text_column(self):
        client = dv.DataverseClient({
            'environment_url': 'https://example.crm.dynamics.com',
            'tenant_id': 'test-tenant', 'client_id': 'test-client',
        })
        client.token = MagicMock(return_value='mock-token')
        client.get_entity_attributes = MagicMock(side_effect=[{}, {
            'cr63f_actiontype': {
                'logical': 'cr63f_actiontype', 'display': 'Action Type',
                'type': 'String', 'can_create': True, 'can_update': True,
            }
        }])
        response = MagicMock(status_code=204, text='')
        client._session.post = MagicMock(return_value=response)

        result = client.create_action_type_column()

        self.assertTrue(result['created'])
        create_call = client._session.post.call_args_list[0]
        payload = create_call.kwargs['json']
        self.assertEqual(payload['SchemaName'], 'cr63f_ActionType')
        self.assertEqual(payload['AttributeType'], 'String')
        self.assertEqual(payload['MaxLength'], 100)
        self.assertEqual(client._session.post.call_args_list[1].args[0].endswith('/PublishXml'), True)

    def test_acquisition_sublevel_schema_migration_adds_optional_text_column(self):
        client = dv.DataverseClient({
            'environment_url': 'https://example.crm.dynamics.com',
            'tenant_id': 'test-tenant', 'client_id': 'test-client',
        })
        client.token = MagicMock(return_value='mock-token')
        client.get_entity_attributes = MagicMock(side_effect=[{}, {
            'cr63f_acquisitionsublevel': {
                'logical': 'cr63f_acquisitionsublevel', 'display': 'Acquisition Sub-Level',
                'type': 'String', 'can_create': True, 'can_update': True,
            }
        }])
        client._session.post = MagicMock(return_value=MagicMock(status_code=204, text=''))

        result = client.ensure_acquisition_sublevel_column('cr63f_loamonitoring')

        self.assertTrue(result['created'])
        payload = client._session.post.call_args_list[0].kwargs['json']
        self.assertEqual(payload['SchemaName'], 'cr63f_AcquisitionSubLevel')
        self.assertEqual(payload['AttributeType'], 'String')
        self.assertEqual(payload['RequiredLevel']['Value'], 'None')
        self.assertEqual(client._session.post.call_args_list[1].args[0].endswith('/PublishXml'), True)

    def test_get_settings_includes_logs_table(self):
        settings = self.app.get_settings()
        self.assertIn('logs_table', settings)
        self.assertEqual(settings['logs_table'], 'cr63f_loamonitoringlogs')

    def test_new_entry_saves_to_master_and_logs(self):
        self.mock_dv.save_loa_record.return_value = {
            'success': True,
            'action': 'create',
            'id': 'master-id-12345'
        }
        self.mock_dv.save_log_record.return_value = {
            'success': True,
            'action': 'create',
            'id': 'log-id-12345'
        }

        new_entry_data = {
            'Type': 'LAND',
            'Area-Index': 'AREA-001',
            'cr63f_control1': 'CTRL-001',
            'Index No': 'IDX-001',
            'Acquisition Sub-Level': 'L0.0',
            'cr63f_acquisitionsublevel': 'L0.0',
        }

        res = self.app.save_loa_record_entry(new_entry_data, record_id=None)
        self.assertTrue(res.get('success'))
        # A blank Endorsed Date defaults to today's date.
        today = datetime.now().strftime('%m/%d/%Y')
        self.mock_dv.save_loa_record.assert_called_once_with(
            dict(new_entry_data, cr63f_endorseddate=today, **{'Endorsed Date': today})
        )
        self.mock_dv.save_log_record.assert_called_once()
        self.assertEqual(self.app.loa_records[0]['__id'], 'master-id-12345')
        self.assertEqual(len(self.app.log_records), 1)
        self.assertEqual(self.app.log_records[0]['action_type'], 'NEW_ENTRY')

    def test_new_entry_requires_acquisition_sublevel(self):
        result = self.app.save_loa_record_entry({'Type': 'LAND', 'Area-Index': 'AREA-001'})
        self.assertFalse(result['success'])
        self.assertIn('Acquisition Sub-Level is required', result['error'])
        self.mock_dv.save_loa_record.assert_not_called()

    def test_update_existing_entry_updates_master_and_logs(self):
        self.mock_dv.save_loa_record.return_value = {
            'success': True,
            'action': 'update',
            'id': 'master-rec-999'
        }
        self.mock_dv.save_log_record.return_value = {
            'success': True,
            'action': 'create',
            'id': 'log-rec-999'
        }

        existing_entry_data = {
            'Type': 'LAND',
            'Area-Index': 'AREA-001',
            'Index No': 'IDX-001',
            'Acquisition Sub-Level': 'L1.1',
            'cr63f_acquisitionsublevel': 'L1.1',
        }

        self.app.loa_records = [{
            '__id': 'master-rec-999',
            'cr63f_control1': 'CTRL-999',
            'Type': 'LAND',
            'Index No': 'IDX-001',
            'Area-Index': 'AREA-OLD'
        }]

        res = self.app.save_loa_record_entry(existing_entry_data, record_id='master-rec-999')
        self.assertTrue(res.get('success'))
        self.mock_dv.save_loa_record.assert_called_once_with(existing_entry_data, record_id='master-rec-999')
        self.assertEqual(self.mock_dv.save_log_record.call_count, 2)  # baseline + update snapshot
        self.assertEqual(self.app.loa_records[0]['Area-Index'], 'AREA-001')
        self.assertEqual(self.app.log_records[0]['action_type'], 'MANUAL_EDIT')

    def test_save_staged_records_diff_logging(self):
        self.mock_dv.save_loa_record.return_value = {
            'success': True,
            'action': 'update',
            'id': 'rec-1'
        }
        log_ids = iter(f'log-staged-{i}' for i in range(4))
        self.mock_dv.save_log_record.side_effect = lambda record: {
            'success': True, 'action': 'create', 'id': next(log_ids)
        }

        self.app.loa_records = [
            {'__id': 'rec-1', 'cr63f_rccprice': '1000', 'Area-Index': 'LIA-001'},
            {'__id': 'rec-2', 'cr63f_rccprice': '2000', 'Area-Index': 'LIA-002'}
        ]

        staged = [
            {'id': 'rec-1', 'fields': {'cr63f_rccprice': '1500'}},
            {'id': 'rec-2', 'fields': {'cr63f_rccprice': '2500'}}
        ]

        res = self.app.save_staged_records(staged)
        self.assertTrue(res.get('success'))
        self.assertEqual(res.get('saved_count'), 2)
        self.assertEqual(self.mock_dv.save_loa_record.call_count, 2)
        self.assertEqual(self.mock_dv.save_log_record.call_count, 4)  # baseline + update per account
        self.assertEqual(len(self.app.log_records), 2)
        self.assertEqual(self.app.log_records[0]['action_type'], 'MANUAL_EDIT')

    def test_dashboard_analytics_summary(self):
        self.app.loa_records = [
            {
                '__id': 'r1',
                'cr63f_totalcontractprice': '1,000,000.00',
                'cr63f_totalpaymentdeductibletotcp': '200,000.00',
                'cr63f_balance': '800,000.00',
                'cr63f_paid': '20.00%',
                'Mode': 'BUYOUT',
                'Endorsed Date': '09/17/2026'
            },
            {
                '__id': 'r2',
                'cr63f_totalcontractprice': '2,000,000.00',
                'cr63f_totalpaymentdeductibletotcp': '1,000,000.00',
                'cr63f_balance': '1,000,000.00',
                'cr63f_paid': '50.00%',
                'Mode': 'DIRECT PURCHASE',
                'Endorsed Date': '09/17/2026'
            }
        ]

        self.app.log_records = [
            {
            'action_type': 'REFERENCE_UPDATE',
                'area_index': 'LIA-101',
                'work_week': 'WW-38',
                'log_time': '2026-09-17T10:00:00Z',
                'field_name': 'Total Contract Price',
                'old_value': '0',
                'new_value': '900,000.00'
            },
            {
                'action_type': 'REFERENCE_UPDATE',
                'area_index': 'LIA-101',
                'work_week': 'WW-39',
                'log_time': '2026-09-24T10:00:00Z',
                'field_name': 'Total Contract Price',
                'old_value': '900,000.00',
                'new_value': '1,000,000.00'
            },
        ]

        analytics = self.app.get_dashboard_analytics('ALL')
        self.assertTrue(analytics.get('success'))
        self.assertEqual(analytics['total_records'], 2)
        self.assertEqual(len(analytics['changed_accounts']), 1)
        self.assertEqual(analytics['total_tcp_formatted'], '₱3,000,000.00')
        self.assertEqual(analytics['total_deductible_formatted'], '₱1,200,000.00')
        self.assertEqual(analytics['total_balance_formatted'], '₱1,800,000.00')
        self.assertEqual(analytics['overall_paid_percent'], '40.00%')
        self.assertEqual(len(analytics['modes_distribution']), 2)

    def test_delete_log_record(self):
        self.app.log_records = [
            {'__id': 'log-101', 'action_type': 'NEW_ENTRY', 'area_index': 'LIA-101'},
            {'__id': 'log-102', 'action_type': 'MANUAL_EDIT', 'area_index': 'LIA-102'}
        ]
        self.mock_dv.delete_log_record.return_value = {'success': True, 'action': 'delete', 'id': 'log-101'}

        res = self.app.delete_log_record('log-101')
        self.assertFalse(res.get('success'))
        self.mock_dv.delete_log_record.assert_not_called()
        self.assertEqual(len(self.app.log_records), 2)
    def test_save_staged_records_batch_parallel(self):
        self.mock_dv.save_loa_record.return_value = {'success': True, 'action': 'update', 'id': 'rec'}
        log_ids = iter(f'log-{i}' for i in range(20))
        self.mock_dv.save_log_record.side_effect = lambda record: {
            'success': True, 'action': 'create', 'id': next(log_ids)
        }
        self.app.loa_records = [
            {'__id': f'rec-{i}', 'cr63f_rccprice': '1000', 'Area-Index': f'LIA-{i}'}
            for i in range(10)
        ]
        staged = [
            {'id': f'rec-{i}', 'fields': {'cr63f_rccprice': f'{1500 + i}'}}
            for i in range(10)
        ]
        res = self.app.save_staged_records(staged)
        self.assertTrue(res['success'])
        self.assertEqual(res['saved_count'], 10)
        self.assertEqual(self.mock_dv.save_loa_record.call_count, 10)
        self.assertEqual(self.mock_dv.save_log_record.call_count, 20)  # baseline + update snapshot per account
        self.assertEqual(len(self.app.log_records), 10)


    def test_log_entry_full_column_aliases(self):
        self.mock_dv.save_loa_record.return_value = {'success': True, 'action': 'update', 'id': 'rec-1'}
        self.mock_dv.save_log_record.return_value = {'success': True, 'action': 'create', 'id': 'log-1'}
        self.app.loa_records = [
            {'__id': 'rec-1', 'cr63f_totalcontractprice': '1000000', 'Area-Index': 'LIA-001', 'Index No': 'C-1', 'Title No': 'T-100'}
        ]
        staged = [
            {'id': 'rec-1', 'fields': {'cr63f_totalcontractprice': '1250000'}}
        ]
        res = self.app.save_staged_records(staged)
        self.assertTrue(res['success'])
        self.assertEqual(len(self.app.log_records), 1)
        log = self.app.log_records[0]
        self.assertEqual(log['field_name'], 'Total Contract Price')
        self.assertEqual(log['Field Changed'], 'Total Contract Price')
        self.assertEqual(log['cr63f_fieldchanged'], 'Total Contract Price')
        self.assertEqual(log['old_value'], '₱1,000,000.00')
        self.assertEqual(log['Previous Value'], '₱1,000,000.00')
        self.assertEqual(log['new_value'], '₱1,250,000.00')
        self.assertEqual(log['New Value'], '₱1,250,000.00')
        self.assertEqual(log['area_index'], 'LIA-001')
        self.assertEqual(log['AREA-INDEX'], 'LIA-001')
        self.assertEqual(log['index_no'], 'C-1')
        self.assertEqual(log['Index No'], 'C-1')
        self.assertEqual(log['title_no'], 'T-100')
        self.assertEqual(log['Title No'], 'T-100')
        self.assertEqual(log['action_type'], 'MANUAL_EDIT')

    def test_populate_from_references_direct_persist(self):
        self.mock_dv.save_loa_record.return_value = {'success': True, 'action': 'update', 'id': 'rec-1'}
        self.mock_dv.save_log_record.return_value = {'success': True, 'action': 'create', 'id': 'log-1'}
        self.app.finance_rows = [
            {
                'Type': 'NON-ROW',
                'AREA-INDEX': 'LIA-101',
                'Title No': 'T-101',
                'Total Contract Price': '1500000',
                'Total Payment Deductible to TCP': '300000',
                'Balance': '1200000',
                '% Paid': '20.00%'
            }
        ]
        self.app.finance_engine.load_rows(self.app.finance_rows)

        self.app.conversion_rows = [
            {
                'AREA-INDEX': 'LIA-101',
                'Title No': 'T-101',
                'Status': 'DAR SECURED'
            }
        ]
        self.app.conversion_engine.load_rows(self.app.conversion_rows)

        self.app.loa_records = [
            {
                '__id': 'rec-1',
                'Type': 'NON-ROW',
                'Area-Index': 'LIA-101',
                'Title No': 'T-101',
                'Total Contract Price': '1,000,000.00',
                'Total Payment Deductible to TCP': '0.00',
                'Balance': '1,000,000.00',
                '% Paid': '0.00%',
                'Conversion Status': 'PENDING'
            }
        ]

        res = self.app.populate_loa_records_from_references(persist_to_dataverse=True)
        # Responses cross the pywebview JSON bridge; Decimal finance values must
        # be returned as exact strings instead of causing a serialization error.
        json.dumps(res)
        self.assertTrue(res['success'])
        self.assertEqual(res['updated_count'], 1)
        self.assertEqual(res['saved_count'], 1)
        self.assertEqual(self.mock_dv.save_loa_record.call_count, 1)
        # The first-ever update writes a hidden baseline and one changed snapshot.
        self.assertEqual(self.mock_dv.save_log_record.call_count, 2)
        snapshot = self.mock_dv.save_log_record.call_args_list[-1][0][0]
        self.assertEqual(snapshot['field_name'], 'Full Record')
        self.assertEqual(snapshot['action_type'], 'REFERENCE_UPDATE')
        self.assertEqual(snapshot['cr63f_actiontype'], 'REFERENCE_UPDATE')
        self.assertNotIn('AUDIT_DIFFS_JSON:', snapshot.get('cr63f_details', ''))
        saved_fields = self.mock_dv.save_loa_record.call_args[0][0]
        # % Paid is always saved, rounded to 2 decimals, so Dataverse never goes stale.
        self.assertEqual(saved_fields['cr63f_paid'], Decimal('20.00'))
        self.assertEqual(res['records'][0]['__display_percent_paid'], '20.00%')
        self.mock_dv.fetch_records.return_value = [dict(row) for row in self.app.pending_session_logs]
        loaded = self.app.fetch_entry_logs()
        self.assertTrue(loaded['success'])
        self.assertGreaterEqual(len(self.app.log_records), 3)

        # Check that individual log rows have all required column aliases populated
        for log in self.app.log_records:
            self.assertTrue(bool(log.get('Field Changed')))
            self.assertTrue(bool(log.get('Previous Value')))
            self.assertTrue(bool(log.get('New Value')))
            self.assertEqual(log.get('AREA-INDEX'), 'LIA-101')
            self.assertEqual(log.get('action_type'), 'REFERENCE_UPDATE')

        tcp_log = next(log for log in self.app.log_records if log.get('field_name') == 'Total Contract Price')
        self.assertEqual(tcp_log['old_value'], '₱1,000,000.00')
        self.assertEqual(tcp_log['new_value'], '₱1,500,000.00')
        self.assertFalse(any(log.get('field_name') == '% Paid' for log in self.app.log_records))

    def test_full_record_snapshot_in_logs(self):
        """Ensure full record snapshot columns are copied into log entries so logs do not have blank columns."""
        self.mock_dv.save_loa_record.return_value = {'success': True, 'action': 'update', 'id': 'rec-full'}
        self.mock_dv.save_log_record.return_value = {'success': True, 'action': 'create', 'id': 'log-full'}

        self.app.finance_rows = [{
            'AREA-INDEX': 'SNAP-01',
            'Title No': 'T-SNAP',
            'Total Contract Price': '3,500,000.00',
            'Total Payment Deductible to TCP': '1,000,000.00',
            'Balance': '2,500,000.00',
            '% Paid': '28.57%'
        }]
        self.app.finance_engine.load_rows(self.app.finance_rows)

        self.app.loa_records = [{
            '__id': 'rec-full',
            'cr63f_control1': 'CTRL-999',
            'cr63f_type': 'LAND',
            'cr63f_areaindex': 'SNAP-01',
            'cr63f_titleno': 'T-SNAP',
            'cr63f_registeredowner': 'Juan Dela Cruz',
            'cr63f_authorizedrepresentative': 'Maria Santos',
            'cr63f_entity': 'CIHC',
            'cr63f_province': 'Bulacan',
            'cr63f_municipality': 'San Jose',
            'cr63f_lotareaha': '2.50',
            'new_contracttype': 'DOAS ONLY',
            'new_mode': 'OPEN TO SALE',
            'cr63f_conversionstatus': 'DAR PENDING',
            'cr63f_totalcontractprice': '1,000,000.00',
            'cr63f_totalpaymentdeductibletotcp': '0.00',
            'cr63f_balance': '1,000,000.00',
            'cr63f_paid': '0.00%'
        }]

        res = self.app.populate_loa_records_from_references(persist_to_dataverse=True)
        self.assertTrue(res['success'])
        self.assertGreaterEqual(len(self.app.log_records), 1)

        # Verify that all master lot attributes are preserved in the log record
        sample_log = self.app.log_records[0]
        self.assertEqual(sample_log.get('cr63f_control1'), 'CTRL-999')
        self.assertEqual(sample_log.get('cr63f_registeredowner'), 'Juan Dela Cruz')
        self.assertEqual(sample_log.get('cr63f_authorizedrepresentative'), 'Maria Santos')
        self.assertEqual(sample_log.get('cr63f_entity'), 'CIHC')
        self.assertEqual(sample_log.get('cr63f_province'), 'Bulacan')
        self.assertEqual(sample_log.get('cr63f_municipality'), 'San Jose')
        self.assertEqual(sample_log.get('cr63f_lotareaha'), '2.50')
        self.assertEqual(sample_log.get('new_contracttype'), 'DOAS ONLY')
        self.assertEqual(sample_log.get('new_mode'), 'OPEN TO SALE')
        self.assertIn('cr63f_totalcontractprice', sample_log)

    def test_subsequent_populate_detects_zero_changes(self):
        """Ensure that after master records are updated from references, running the update again detects 0 changes."""
        self.mock_dv.save_loa_record.return_value = {'success': True, 'action': 'update', 'id': 'rec-1'}
        self.mock_dv.save_log_record.return_value = {'success': True, 'action': 'create', 'id': 'log-1'}

        self.app.finance_rows = [{
            'Type': 'LAND',
            'AREA-INDEX': 'TEST-001',
            'Title No': 'T-001',
            'Total Contract Price': '5,000,000.00',
            'Total Payment Deductible to TCP': '1,000,000.00',
            'Balance': '4,000,000.00',
            '% Paid': '20.00%'
        }]
        self.app.finance_engine.load_rows(self.app.finance_rows)

        self.app.conversion_rows = [{
            'AREA-INDEX': 'TEST-001',
            'Title No': 'T-001',
            'Status': 'DAR APPROVED'
        }]
        self.app.conversion_engine.load_rows(self.app.conversion_rows)

        self.app.loa_records = [{
            '__id': 'rec-1',
            'Type': 'LAND',
            'Area-Index': 'TEST-001',
            'Title No': 'T-001',
            'Total Contract Price': '2,000,000.00',
            'Total Payment Deductible to TCP': '0.00',
            'Balance': '2,000,000.00',
            '% Paid': '0.00%',
            'Conversion Status': 'PENDING'
        }]

        # Step 1: First update run with persistence
        res1 = self.app.populate_loa_records_from_references(persist_to_dataverse=True)
        self.assertTrue(res1['success'])
        self.assertEqual(res1['updated_count'], 1)
        self.assertEqual(res1['saved_count'], 1)

        # Verify internal master records were updated in app state
        self.assertIn('5,000,000.00', self.app.loa_records[0]['Total Contract Price'])
        self.assertEqual(self.app.loa_records[0]['Conversion Status'], 'DAR APPROVED')

        # Step 2: Subsequent run on self.app.loa_records must detect 0 changes
        res2 = self.app.populate_loa_records_from_references(self.app.loa_records, persist_to_dataverse=False)
        self.assertTrue(res2['success'])
        self.assertEqual(res2['updated_count'], 0)
        self.assertEqual(res2['diff_summary'], [])

    def test_save_record_percentage_casting_and_clean_key(self):
        """Ensure save_record correctly parses percentage and currency strings without setting None, and _clean_key matches correctly."""
        self.assertEqual(dv._clean_key('cr63f_contracttype'), 'contracttype')
        self.assertEqual(dv._clean_key('new_contracttype1'), 'contracttype1')
        self.assertEqual(dv._clean_key('Contract Type'), 'contracttype')

        real_client = dv.DataverseClient({'environment_url': 'https://example.crm.dynamics.com', 'tenant_id': 't', 'client_id': 'c'})
        real_client.token = MagicMock(return_value='mock-token')
        mock_attrs = {
            'cr63f_paid': {'logical': 'cr63f_paid', 'display': '% Paid', 'type': 'Decimal', 'can_create': True, 'can_update': True},
            'cr63f_totalcontractprice': {'logical': 'cr63f_totalcontractprice', 'display': 'Total Contract Price', 'type': 'Money', 'can_create': True, 'can_update': True},
            'new_contracttype': {'logical': 'new_contracttype', 'display': 'Contract Type', 'type': 'String', 'can_create': True, 'can_update': True}
        }
        real_client.get_entity_attributes = MagicMock(return_value=mock_attrs)
        real_client.get_entity_info = MagicMock(return_value={'entity_set': 'cr63f_loamonitorings', 'primary_id': 'cr63f_loamonitoringid', 'primary_name': 'cr63f_name'})
        real_client._session = MagicMock()
        mock_patch_resp = MagicMock()
        mock_patch_resp.status_code = 204
        mock_patch_resp.ok = True
        real_client._session.patch.return_value = mock_patch_resp

        res = real_client.save_record('cr63f_loamonitoring', {
            'cr63f_paid': '25.00%',
            'Total Contract Price': '₱1,500,000.00',
            'cr63f_contracttype': 'DOAS ONLY'
        }, record_id='rec-123')

        self.assertTrue(res['success'])
        patch_kwargs = real_client._session.patch.call_args[1]
        sent_payload = patch_kwargs['json']

        self.assertEqual(sent_payload['cr63f_paid'], 25.0)
        self.assertEqual(sent_payload['cr63f_totalcontractprice'], 1500000.0)
        self.assertEqual(sent_payload['new_contracttype'], 'DOAS ONLY')

    def test_percent_paid_identical_no_diff(self):
        """Ensure that 1.0, '1.0', and '1.00%' are recognized as identical and not flagged as a change."""
        self.assertFalse(app_module._is_different_pct('1.00%', 1.0))
        self.assertFalse(app_module._is_different_pct(1.0, 1.0))
        self.assertFalse(app_module._is_different_pct('1.00%', '1.00%'))

        # Set up a record with 1.00% matching reference 1.00%
        self.app.finance_rows = [{
            'Type': 'LAND',
            'AREA-INDEX': 'LIA-7726',
            'Title No': 'T-7726',
            'Total Contract Price': '1466664.00',
            'Total Payment Deductible to TCP': '14667.00',
            'Balance': '1451997.00',
            '% Paid': '1.00%'
        }]
        self.app.finance_engine.load_rows(self.app.finance_rows)

        self.app.loa_records = [{
            '__id': 'rec-7726',
            'Type': 'LAND',
            'Area-Index': 'LIA-7726',
            'Title No': 'T-7726',
            'Total Contract Price': '₱1,466,664.00',
            'Total Payment Deductible to TCP': '₱14,667.00',
            'Balance': '₱1,451,997.00',
            'cr63f_paid': '1.00%',
            '% Paid': '1.00%'
        }]

        res = self.app.populate_loa_records_from_references(persist_to_dataverse=False)
        self.assertTrue(res['success'])
        self.assertEqual(res['updated_count'], 0)
        self.assertEqual(len(res['diff_summary']), 0)

    def test_format_percent_scale_safe(self):
        from loa_engine import format_percent
        self.assertEqual(format_percent(1.0), '1.00%')
        self.assertEqual(format_percent('1.00%'), '1.00%')
        self.assertEqual(format_percent(100.0), '100.00%')
        self.assertEqual(format_percent(94.3), '94.30%')
        # Stored % Paid is on a 0-100 scale: 0.25 is a quarter of one percent.
        self.assertEqual(format_percent(0.25), '0.25%')
        self.assertEqual(format_percent(0.5), '0.50%')

    def test_derive_field_diffs_from_snapshots_multi_field_and_zero_change(self):
        from app import derive_field_diffs_from_snapshots
        raw_snapshots = [
            # Snapshot 1: initial entry at 09:00
            {
                '__id': 'log-snap-1',
                'createdon': '2026-09-18T09:00:00Z',
                'cr63f_areaindex': 'LIA-100',
                'cr63f_totalcontractprice': '1000000',
                'cr63f_totalpaymentdeductibletotcp': '100000',
                'cr63f_balance': '900000',
                'cr63f_paid': '10.00%',
                'cr63f_registeredowner': 'Juan Dela Cruz',
                'cr63f_entity': 'CIHC',
                'cr63f_province': 'Batangas',
                'cr63f_municipality': 'Calaca',
                'cr63f_conversionstatus': 'PENDING'
            },
            # Snapshot 2: update at 10:00 (% Paid changed from 10% to 50%, Deductible to 500k, Balance to 500k)
            {
                '__id': 'log-snap-2',
                'createdon': '2026-09-18T10:00:00Z',
                'cr63f_areaindex': 'LIA-100',
                'cr63f_totalcontractprice': '1000000',
                'cr63f_totalpaymentdeductibletotcp': '500000',
                'cr63f_balance': '500000',
                'cr63f_paid': '50.00%',
                'cr63f_registeredowner': 'Juan Dela Cruz',
                'cr63f_entity': 'CIHC',
                'cr63f_province': 'Batangas',
                'cr63f_municipality': 'Calaca',
                'cr63f_conversionstatus': 'PENDING'
            },
            # Snapshot 3: redundant snapshot at 11:00 with EXACT SAME values (0 changes)
            {
                '__id': 'log-snap-3',
                'createdon': '2026-09-18T11:00:00Z',
                'cr63f_areaindex': 'LIA-100',
                'cr63f_totalcontractprice': '1000000',
                'cr63f_totalpaymentdeductibletotcp': '500000',
                'cr63f_balance': '500000',
                'cr63f_paid': '50.00%',
                'cr63f_registeredowner': 'Juan Dela Cruz',
                'cr63f_entity': 'CIHC',
                'cr63f_province': 'Batangas',
                'cr63f_municipality': 'Calaca',
                'cr63f_conversionstatus': 'PENDING'
            }
        ]

        derived = derive_field_diffs_from_snapshots(raw_snapshots)
        # Snapshot 1 produces 1 initial entry
        # Snapshot 2 produces only the authoritative deductible diff
        # Snapshot 3 has ZERO changes and MUST be completely omitted!
        self.assertEqual(len(derived), 2)

        # Newest logs first
        diff_fields = [d['field_name'] for d in derived]
        self.assertIn('Total Payment Deductible to TCP', diff_fields)
        self.assertNotIn('% Paid', diff_fields)
        self.assertNotIn('Balance', diff_fields)
        self.assertIn('Initial Entry', diff_fields)

    def test_acquisition_sublevel_snapshot_change_is_audited(self):
        from app import derive_field_diffs_from_snapshots
        derived = derive_field_diffs_from_snapshots([
            {
                'createdon': '2026-09-18T08:00:00Z',
                'cr63f_control1': 'ACQ-001',
                'cr63f_acquisitionsublevel': 'L1.1',
            },
            {
                'createdon': '2026-09-18T09:00:00Z',
                'cr63f_control1': 'ACQ-001',
                'cr63f_acquisitionsublevel': 'L1.2',
            },
        ])
        change = next(row for row in derived if row['field_name'] == 'Acquisition Sub-Level')
        self.assertEqual(change['old_value'], 'L1.1')
        self.assertEqual(change['new_value'], 'L1.2')

    def test_snapshot_diff_preserves_manual_action_for_financial_field(self):
        from app import derive_field_diffs_from_snapshots

        derived = derive_field_diffs_from_snapshots([
            {
                '__id': 'manual-snap-1',
                'createdon': '2026-09-17T09:00:00Z',
                'cr63f_control1': 'CTRL-MANUAL',
                'cr63f_totalcontractprice': '1000000',
                'action_type': 'MANUAL_EDIT',
            },
            {
                '__id': 'manual-snap-2',
                'createdon': '2026-09-17T10:00:00Z',
                'cr63f_control1': 'CTRL-MANUAL',
                'cr63f_totalcontractprice': '1250000',
                'action_type': 'MANUAL_EDIT',
            },
        ])

        change = next(d for d in derived if d['field_name'] == 'Total Contract Price')
        self.assertEqual(change['action_type'], 'MANUAL_EDIT')
        self.assertEqual(change['work_week'], '2026 · WW-38')  # Entry Logs weeks carry their year

    def test_snapshot_diff_maps_reference_action_and_recalculates_work_week(self):
        from app import derive_field_diffs_from_snapshots

        derived = derive_field_diffs_from_snapshots([
            {
                '__id': 'reference-snap-1',
                'createdon': '2026-09-17T09:00:00Z',
                'cr63f_control1': 'CTRL-REFERENCE',
                'cr63f_balance': '900000',
                'cr63f_workweek': 'WW-39',
                'cr63f_name': 'LOT - Full Record (MANUAL_EDIT)',
            },
            {
                '__id': 'reference-snap-2',
                'createdon': '2026-09-24T09:00:00Z',
                'cr63f_control1': 'CTRL-REFERENCE',
                'cr63f_balance': '500000',
                'cr63f_workweek': 'WW-38',
                'cr63f_name': 'LOT - Full Record (REFERENCE_UPDATE)',
            },
        ])

        self.assertFalse(any(d['field_name'] in ('Balance', '% Paid') for d in derived))
        self.assertFalse(any(d['action_type'] == 'REFERENCE_UPDATE' for d in derived))

    def test_reference_snapshots_show_every_changed_account_without_prior_history(self):
        from app import build_log_entry, derive_field_diffs_from_snapshots

        source_rows = []
        for account in ('CTRL-A', 'CTRL-B'):
            changes = [
                build_log_entry('REFERENCE_UPDATE', 'WW-39', '09/24/2026 10:00 AM', account, '', '',
                                'Total Payment Deductible to TCP', '₱100.00', '₱200.00'),
            ]
            baseline = build_log_entry(
                'BASELINE', 'WW-39', '09/24/2026 09:00 AM', account, '', '',
                'Full Record', '-', 'Baseline Snapshot',
                full_record={'cr63f_control1': account, 'cr63f_totalpaymentdeductibletotcp': '100.00'}
            )
            baseline['__id'] = f'baseline-{account}'
            baseline['createdon'] = '2026-09-24T09:00:00Z'
            row = build_log_entry(
                'REFERENCE_UPDATE', 'WW-39', '09/24/2026 10:00 AM', account, '', '',
                'Full Record', 'Previous Snapshot', 'Updated Snapshot',
                full_record={'cr63f_control1': account, 'cr63f_totalpaymentdeductibletotcp': '200.00'}
            )
            row['__id'] = f'snapshot-{account}'
            row['createdon'] = '2026-09-24T10:00:00Z'
            source_rows.extend([baseline, row])

        derived = derive_field_diffs_from_snapshots(source_rows)
        self.assertEqual(len(derived), 2)
        self.assertEqual({row['cr63f_control1'] for row in derived}, {'CTRL-A', 'CTRL-B'})
        self.assertTrue(all(row['action_type'] == 'REFERENCE_UPDATE' for row in derived))

    def test_finance_lookup_ignores_reference_percent_and_preserves_calculation(self):
        from loa_engine import FinanceLookupEngine, format_calculated_percent

        engine = FinanceLookupEngine()
        engine.load_rows([{
            'Type': 'NON-ROW', 'AREA-INDEX': 'AI-1',
            'Total Contract Price': '300',
            'Total Payment Deductible to TCP': '100',
            '% Paid': '99.9%'
        }])
        result = engine.lookup(type_val='NON-ROW', index_no='', area_index_val='AI-1', title_mod_val='')
        self.assertTrue(result['matched'])
        self.assertEqual(result['percent_paid_formatted'], '33.3333333333%')
        self.assertEqual(format_calculated_percent('300', '100'), '33.3333333333%')

    def test_manual_audit_failure_restores_land_master_value(self):
        self.app.loa_records = [{
            '__id': 'master-manual-rollback',
            'cr63f_areaindex': 'AI-MANUAL-ROLLBACK',
            'cr63f_rccprice': '1000',
        }]
        self.mock_dv.save_loa_record.return_value = {'success': True}
        self.mock_dv.save_log_record.side_effect = [
            {'success': True, 'action': 'create', 'id': 'baseline-manual'},
            RuntimeError('audit unavailable'),
        ]

        result = self.app.save_staged_records([{
            'id': 'master-manual-rollback',
            'fields': {'cr63f_rccprice': '1500'},
        }])

        self.assertFalse(result['success'])
        self.assertIn('rolled back', result['error'])
        self.assertEqual(self.mock_dv.save_loa_record.call_count, 2)
        self.assertEqual(self.app.loa_records[0]['cr63f_rccprice'], '1000')
        self.assertEqual(self.app.log_records, [])

    def test_reference_audit_failure_restores_land_master_value(self):
        self.app.loa_records = [{
            '__id': 'master-reference-rollback',
            'cr63f_areaindex': 'AI-REFERENCE-ROLLBACK',
            'cr63f_totalcontractprice': '100',
            'cr63f_totalpaymentdeductibletotcp': '50',
            'cr63f_balance': '50',
            'cr63f_paid': '50.00%',
        }]
        self.app.finance_rows = [{'source': 'loaded'}]
        self.app.finance_engine.lookup = MagicMock(return_value={
            'matched': True,
            'total_contract_price_formatted': '₱200.00',
            'total_contract_price': 200,
            'total_payment_deductible_formatted': '₱100.00',
            'total_payment_deductible': 100,
            'balance_formatted': '₱100.00',
            'balance': 100,
            'percent_paid_formatted': '50.00%',
            'percent_paid': 50,
        })
        self.mock_dv.save_loa_record.return_value = {'success': True}
        self.mock_dv.save_log_record.side_effect = [
            {'success': True, 'action': 'create', 'id': 'baseline-reference'},
            RuntimeError('audit unavailable'),
        ]

        result = self.app.populate_loa_records_from_references(persist_to_dataverse=True)

        self.assertFalse(result['success'])
        self.assertIn('rolled back', result['error'])
        self.assertEqual(self.mock_dv.save_loa_record.call_count, 2)
        self.assertEqual(self.app.loa_records[0]['cr63f_totalcontractprice'], '100')
        self.assertEqual(self.app.log_records, [])

    def test_explicit_log_recalculates_stale_work_week_from_createdon(self):
        from app import derive_field_diffs_from_snapshots

        derived = derive_field_diffs_from_snapshots([{
            '__id': 'explicit-log-1',
            'createdon': '2026-09-17T10:33:00Z',
            'cr63f_control1': 'CTRL-EXPLICIT',
            'field_name': 'Registered Owner',
            'old_value': 'Old Owner',
            'new_value': 'New Owner',
            'work_week': 'WW-39',
            'action_type': 'MANUAL_EDIT',
        }])

        self.assertEqual(derived[0]['work_week'], '2026 · WW-38')  # Entry Logs weeks carry their year

    def test_fetch_entry_logs_derives_diffs_and_skips_zero_change(self):
        self.mock_dv.fetch_records.return_value = [
            {
                'createdon': '2026-09-18T08:00:00Z',
                'cr63f_areaindex': 'LOT-200',
                'cr63f_conversionstatus': 'PENDING',
                'cr63f_totalcontractprice': '2,000,000.00'
            },
            {
                'createdon': '2026-09-18T09:00:00Z',
                'cr63f_areaindex': 'LOT-200',
                'cr63f_conversionstatus': 'DAR SECURED',
                'cr63f_totalcontractprice': '2,000,000.00'
            }
        ]
        res = self.app.fetch_entry_logs()
        self.assertTrue(res['success'])
        # Initial entry + 1 diff (Conversion Status from PENDING to DAR SECURED)
        self.assertEqual(len(res['logs']), 2)
        status_log = next(l for l in res['logs'] if l['field_name'] == 'Conversion Status')
        self.assertEqual(status_log['old_value'], 'PENDING')
        self.assertEqual(status_log['new_value'], 'DAR SECURED')

    def test_logs_persist_across_restarts(self):
        # 1. Add log entry in app instance 1
        new_entry_data = {
            'Type': 'LAND',
            'Area-Index': 'PERSIST-AREA-001',
            'cr63f_control1': 'CTRL-PERSIST-001',
            'Index No': 'PERSIST-IDX-001',
            'Acquisition Sub-Level': 'L0.0',
            'cr63f_acquisitionsublevel': 'L0.0'
        }
        self.mock_dv.save_loa_record.return_value = {'success': True, 'action': 'create', 'id': 'p-id-1'}
        self.mock_dv.save_log_record.return_value = {'success': True, 'action': 'create', 'id': 'p-log-1'}

        res1 = self.app.save_loa_record_entry(new_entry_data)
        self.assertTrue(res1['success'])
        self.assertTrue(any(l.get('area_index') == 'PERSIST-AREA-001' for l in self.app.log_records))

        # 2. Create new app instance (simulating app exit & restart)
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            restarted_app = Api()

        # 3. Re-fetch from Dataverse after restart; no local cache is used.
        restarted_app.dv_client = self.mock_dv
        self.mock_dv.fetch_records.return_value = [dict(self.app.pending_session_logs[-1])]
        fetched = restarted_app.fetch_entry_logs()
        self.assertTrue(fetched['success'])
        found = next((l for l in restarted_app.log_records if l.get('area_index') == 'PERSIST-AREA-001'), None)
        self.assertIsNotNone(found)
        self.assertEqual(found.get('action_type'), 'NEW_ENTRY')

    def test_delete_log_record_log_only(self):
        self.app.log_records = [{
            '__id': 'log-del-1',
            'area_index': 'LIA-001',
            'field_name': 'Mode',
            'old_value': 'LEASE',
            'new_value': 'PURCHASE'
        }]
        self.app.loa_records = [{
            '__id': 'master-1',
            'cr63f_areaindex': 'LIA-001',
            'new_mode': 'PURCHASE',
            'Mode': 'PURCHASE'
        }]

        res = self.app.delete_log_record('log-del-1', revert_master=False)
        self.assertFalse(res['success'])
        self.assertEqual(len(self.app.log_records), 1)
        # Master record should remain unchanged
        self.assertEqual(self.app.loa_records[0]['new_mode'], 'PURCHASE')

    def test_delete_log_record_revert_master_value(self):
        self.app.log_records = [{
            '__id': 'log-del-2',
            'area_index': 'LIA-002',
            'field_name': 'Total Contract Price',
            'old_value': '₱1,000,000.00',
            'new_value': '₱2,000,000.00'
        }]
        self.app.loa_records = [{
            '__id': 'master-2',
            'cr63f_areaindex': 'LIA-002',
            'cr63f_totalcontractprice': '2000000',
            'Total Contract Price': '2000000',
            'cr63f_totalpaymentdeductibletotcp': '500000',
            'cr63f_balance': '1500000',
            'cr63f_paid': '25.00%'
        }]

        res = self.app.delete_log_record('log-del-2', revert_master=True)
        self.assertTrue(res['success'])
        self.assertTrue(res['reverted'])
        self.assertEqual(len(self.app.log_records), 2)
        # Master record should be reverted to 1,000,000
        master = self.app.loa_records[0]
        self.assertEqual(master['cr63f_totalcontractprice'], '1000000.00')
        self.assertEqual(master['cr63f_balance'], 500000.0)
        self.assertEqual(float(master['cr63f_paid']), 50.0)
        self.assertEqual(master['% Paid'], '50.00%')

    def test_delete_log_record_revert_dar_status(self):
        self.app.log_records = [{
            '__id': 'log-del-3',
            'area_index': 'AI-200',
            'title_no': 'T-999',
            'field_name': 'DAR STATUS',
            'old_value': 'FOR APPLICATION',
            'new_value': 'DAR FINALITY SECURED'
        }]
        self.app.loa_records = [{
            '__id': 'master-3',
            'cr63f_areaindex': 'AI-200',
            'cr63f_titleno': 'T-999',
            'cr63f_conversionstatus': 'DAR FINALITY SECURED',
            'Conversion Status': 'DAR FINALITY SECURED'
        }]

        res = self.app.delete_log_record('log-del-3', revert_master=True)
        self.assertTrue(res['success'])
        self.assertTrue(res['reverted'])
        self.assertEqual(len(self.app.log_records), 2)
        master = self.app.loa_records[0]
        self.assertEqual(master['cr63f_conversionstatus'], 'FOR APPLICATION')
        self.assertEqual(master['Conversion Status'], 'FOR APPLICATION')

    def test_delete_log_record_by_index_fallback(self):
        self.app.log_records = [{
            '__id': '',
            'area_index': 'AI-300',
            'field_name': 'Mode',
            'old_value': 'PURCHASE',
            'new_value': 'EXCHANGE'
        }]
        self.app.loa_records = [{
            '__id': 'master-4',
            'cr63f_areaindex': 'AI-300',
            'new_mode': 'EXCHANGE',
            'Mode': 'EXCHANGE'
        }]

        res = self.app.delete_log_record('0', revert_master=True)
        self.assertTrue(res['success'])
        self.assertTrue(res['reverted'])
        self.assertEqual(len(self.app.log_records), 2)
        master = self.app.loa_records[0]
        self.assertEqual(master['new_mode'], 'PURCHASE')


    def test_delete_log_revert_field_writes_to_dataverse_master(self):
        self.app.log_records = [{
            '__id': '11111111-2222-3333-4444-555555555555_Total_Contract_Price',
            'area_index': 'LIA-777',
            'field_name': 'Total Contract Price',
            'old_value': '₱3,000,000.00',
            'new_value': '₱5,000,000.00'
        }]
        self.app.loa_records = [{
            '__id': 'master-guid-777',
            'cr63f_areaindex': 'LIA-777',
            'cr63f_totalcontractprice': '5000000',
            'Total Contract Price': '5000000',
            'cr63f_totalpaymentdeductibletotcp': '1000000',
            'cr63f_balance': '4000000',
            'cr63f_paid': '20.00%'
        }]
        self.mock_dv.save_loa_record.return_value = {'success': True, 'action': 'update', 'id': 'master-guid-777'}
        self.mock_dv.delete_log_record.return_value = {'success': True, 'action': 'delete', 'id': '11111111-2222-3333-4444-555555555555'}

        res = self.app.delete_log_record('11111111-2222-3333-4444-555555555555_Total_Contract_Price', revert_master=True)
        self.assertTrue(res['success'])
        self.assertTrue(res['reverted'])
        self.mock_dv.save_loa_record.assert_called_once()
        call_args = self.mock_dv.save_loa_record.call_args
        payload = call_args[0][0]
        self.assertEqual(payload['cr63f_totalcontractprice'], '3000000.00')
        self.assertEqual(payload['cr63f_balance'], 2000000.0)
        self.assertEqual(Decimal(str(payload['cr63f_paid'])), Decimal('33.3'))  # cr63f_paid keeps 1 decimal
        self.assertEqual(call_args[1]['record_id'], 'master-guid-777')
        self.mock_dv.delete_log_record.assert_not_called()

    def test_delete_log_revert_field_surfaces_dataverse_error(self):
        self.app.log_records = [{
            '__id': 'log-del-err',
            'area_index': 'LIA-888',
            'field_name': 'Total Contract Price',
            'old_value': '₱3,000,000.00',
            'new_value': '₱5,000,000.00'
        }]
        self.app.loa_records = [{
            '__id': 'master-guid-888',
            'cr63f_areaindex': 'LIA-888',
            'cr63f_totalcontractprice': '5000000',
            'Total Contract Price': '5000000'
        }]
        self.mock_dv.save_loa_record.side_effect = RuntimeError("Dataverse server timeout")

        res = self.app.delete_log_record('log-del-err', revert_master=True)
        self.assertFalse(res['success'])
        self.assertIn("Failed to revert record in Dataverse", res['error'])

    def test_revert_balance_is_blocked_because_it_is_calculated(self):
        self.app.log_records = [{
            '__id': 'balance-log',
            'cr63f_control1': 'CTRL-REVERT-BAL',
            'area_index': 'AI-REVERT-BAL',
            'field_name': 'Balance',
            'old_value': '₱0.00',
            'new_value': '₱102,000.00',
        }]
        self.app.loa_records = [{
            '__id': 'master-revert-bal',
            'cr63f_control1': 'CTRL-REVERT-BAL',
            'cr63f_areaindex': 'AI-REVERT-BAL',
            'cr63f_totalcontractprice': '8202627',
            'cr63f_totalpaymentdeductibletotcp': '8100627',
            'cr63f_balance': 102000.0,
            'cr63f_paid': '98.80%',
        }]
        res = self.app.delete_log_record('balance-log', revert_master=True)

        self.assertFalse(res['success'])
        self.assertIn('is calculated', res['error'])
        self.mock_dv.save_loa_record.assert_not_called()
        self.mock_dv.save_log_record.assert_not_called()

    def test_revert_percent_is_blocked_because_it_is_calculated(self):
        self.app.log_records = [{
            '__id': 'paid-log',
            'cr63f_control1': 'CTRL-REVERT-PAID',
            'area_index': 'AI-REVERT-PAID',
            'field_name': '% Paid',
            'old_value': '1.00%',
            'new_value': '98.80%',
        }]
        self.app.loa_records = [{
            '__id': 'master-revert-paid',
            'cr63f_control1': 'CTRL-REVERT-PAID',
            'cr63f_areaindex': 'AI-REVERT-PAID',
            'cr63f_totalcontractprice': '8202627',
            'cr63f_totalpaymentdeductibletotcp': '8100627',
            'cr63f_balance': '102000.00',
            'cr63f_paid': '98.8%',
        }]
        res = self.app.delete_log_record('paid-log', revert_master=True)

        self.assertFalse(res['success'])
        self.assertIn('is calculated', res['error'])
        self.mock_dv.save_loa_record.assert_not_called()
        self.mock_dv.save_log_record.assert_not_called()

    def test_revert_deductible_accepts_numeric_format_and_recalculates_dependants(self):
        self.app.log_records = [{
            '__id': 'deductible-log',
            'cr63f_control1': 'CTRL-REVERT-DED',
            'area_index': 'AI-REVERT-DED',
            'field_name': 'Total Payment Deductible to TCP',
            'old_value': '₱8,202,627.00',
            'new_value': '₱8,100,627.00',
        }]
        self.app.loa_records = [{
            '__id': 'master-revert-ded',
            'cr63f_control1': 'CTRL-REVERT-DED',
            'cr63f_areaindex': 'AI-REVERT-DED',
            'cr63f_totalcontractprice': '8202627',
            'cr63f_totalpaymentdeductibletotcp': 8100627.0,
            'cr63f_balance': '102000.00',
            'cr63f_paid': '98.80%',
        }]
        self.mock_dv.save_loa_record.return_value = {'success': True, 'id': 'master-revert-ded'}
        self.mock_dv.save_log_record.return_value = {'success': True, 'id': 'revert-deductible-log'}

        res = self.app.delete_log_record('deductible-log', revert_master=True)

        self.assertTrue(res['success'])
        payload = self.mock_dv.save_loa_record.call_args[0][0]
        self.assertEqual(payload['cr63f_totalpaymentdeductibletotcp'], '8202627.00')
        self.assertEqual(payload['cr63f_balance'], 0.0)
        self.assertEqual(float(payload['cr63f_paid']), 100.0)
        self.assertEqual(len(self.app.pending_session_logs), 1)
        self.assertEqual(self.app.pending_session_logs[0]['action_type'], 'REVERT')

    def test_fresh_revert_log_remains_visible_while_dataverse_read_is_delayed(self):
        pending = {
            '__id': 'pending-revert-log',
            'createdon': '2026-09-24T11:50:00Z',
            'cr63f_control1': 'CTRL-PENDING',
            'field_name': 'Total Payment Deductible to TCP',
            'old_value': '₱8,202,627.00',
            'new_value': '₱8,100,627.00',
            'action_type': 'REVERT',
            'work_week': 'WW-39',
        }
        self.app.pending_session_logs = [pending]
        self.mock_dv.fetch_records.return_value = []

        res = self.app.fetch_entry_logs()

        self.assertTrue(res['success'])
        self.assertEqual(len(res['logs']), 1)
        self.assertEqual(res['logs'][0]['action_type'], 'REVERT')

    def test_calculated_fields_never_appear_as_field_changed(self):
        raw_snapshots = [
            {
                '__id': 'snap-1',
                'action_type': 'MANUAL_EDIT',
                'cr63f_areaindex': 'AI-999',
                'createdon': '2026-09-23T10:00:00Z',
                'Total Contract Price': '1000000',
                'Balance': '500000',
                '% Paid': '50.00%'
            },
            {
                '__id': 'snap-2',
                'action_type': 'MANUAL_EDIT',
                'cr63f_areaindex': 'AI-999',
                'createdon': '2026-09-23T10:05:00Z',
                'Total Contract Price': '1000000',
                'Balance': '400000',
                '% Paid': '60.00%'
            }
        ]

        derived = app_module.derive_field_diffs_from_snapshots(raw_snapshots)
        field_names = [d.get('field_name') for d in derived]
        self.assertNotIn('Balance', field_names)
        self.assertNotIn('% Paid', field_names)

    def test_reference_audit_compares_currency_without_cent_rounding(self):
        from app import derive_field_diffs_from_snapshots

        rows = [
            {
                '__id': 'decimal-base', 'createdon': '2026-09-23T10:00:00Z',
                'cr63f_control1': 'CTRL-DECIMAL', 'action_type': 'BASELINE',
                'cr63f_totalcontractprice': '1000.001',
            },
            {
                '__id': 'decimal-update', 'createdon': '2026-09-23T10:05:00Z',
                'cr63f_control1': 'CTRL-DECIMAL', 'action_type': 'REFERENCE_UPDATE',
                'cr63f_totalcontractprice': '1000.002',
            },
        ]
        derived = derive_field_diffs_from_snapshots(rows)
        self.assertEqual(len(derived), 1)
        self.assertEqual(derived[0]['old_value'], '₱1,000.001')
        self.assertEqual(derived[0]['new_value'], '₱1,000.002')

    def test_save_staged_finance_records_full_master_snapshot(self):
        """Ensure save_staged_finance_records looks up Land Master DB row and snapshots all master lot attributes into the log."""
        self.mock_dv.save_finance_record.return_value = {'success': True, 'action': 'update', 'id': 'fin-1'}
        self.mock_dv.save_log_record.return_value = {'success': True, 'action': 'create', 'id': 'log-fin-1'}

        self.app.loa_records = [{
            '__id': 'master-rec-501',
            'cr63f_control1': 'CTRL-501',
            'cr63f_type': 'LAND',
            'cr63f_areaindex': 'LIA-501',
            'cr63f_titleno': 'T-501',
            'cr63f_registeredowner': 'Juan Pedro',
            'cr63f_authorizedrepresentative': 'Ana Pedro',
            'cr63f_entity': 'CIHC',
            'cr63f_province': 'Batangas',
            'cr63f_municipality': 'Calaca',
            'cr63f_dbadjustedbarangay': 'Dacanlao',
            'cr63f_lotareaha': '1.75',
            'new_contracttype': 'DOAS ONLY',
            'new_mode': 'OPEN TO SALE',
            'cr63f_conversionstatus': 'DAR SECURED'
        }]

        self.app.finance_rows = [{
            '__id': 'fin-1',
            'Area Index': 'LIA-501',
            'Title Mod': 'T-501',
            'Total Contract Price': '2,000,000.00',
            'Total Payment Deductible to TCP': '500,000.00'
        }]

        staged = [{
            'id': 'fin-1',
            'fields': {'Total Contract Price': '3,000,000.00'}
        }]

        res = self.app.save_staged_finance_records(staged)
        self.assertTrue(res['success'])
        self.assertEqual(len(self.app.log_records), 1)
        log = self.app.log_records[0]
        self.assertEqual(log['field_name'], 'Total Contract Price')
        self.assertEqual(log['old_value'], '₱2,000,000.00')
        self.assertEqual(log['new_value'], '₱3,000,000.00')
        self.assertEqual(log['cr63f_registeredowner'], 'Juan Pedro')
        self.assertEqual(log['Registered Owner'], 'Juan Pedro')
        self.assertEqual(log['cr63f_authorizedrepresentative'], 'Ana Pedro')
        self.assertEqual(log['cr63f_province'], 'Batangas')
        self.assertEqual(log['cr63f_municipality'], 'Calaca')
        self.assertEqual(log['cr63f_dbadjustedbarangay'], 'Dacanlao')
        self.assertEqual(log['cr63f_lotareaha'], '1.75')
        self.assertEqual(log['new_contracttype'], 'DOAS ONLY')
        self.assertEqual(log['new_mode'], 'OPEN TO SALE')
        self.assertEqual(log['cr63f_control1'], 'CTRL-501')

    def test_save_staged_conversion_records_full_master_snapshot(self):
        """Ensure save_staged_conversion_records looks up Land Master DB row and snapshots all master lot attributes into the log."""
        self.mock_dv.save_conversion_record.return_value = {'success': True, 'action': 'update', 'id': 'conv-1'}
        self.mock_dv.save_log_record.return_value = {'success': True, 'action': 'create', 'id': 'log-conv-1'}

        self.app.loa_records = [{
            '__id': 'master-rec-601',
            'cr63f_control1': 'CTRL-601',
            'cr63f_type': 'LAND',
            'cr63f_areaindex': 'LIA-601',
            'cr63f_titleno': 'T-601',
            'cr63f_registeredowner': 'Maria Santos',
            'cr63f_authorizedrepresentative': 'Carlos Santos',
            'cr63f_entity': 'MBHC',
            'cr63f_province': 'Batangas',
            'cr63f_municipality': 'Tuy',
            'cr63f_dbadjustedbarangay': 'Magahis',
            'cr63f_lotareaha': '0.85',
            'new_contracttype': 'JOINT VENTURE',
            'new_mode': 'DIRECT PURCHASE',
            'cr63f_conversionstatus': 'DAR PENDING'
        }]

        self.app.conversion_rows = [{
            '__id': 'conv-1',
            'cr63f_areaindex': 'LIA-601',
            'cr63f_titlenumber': 'T-601',
            'DAR STATUS': 'DAR FOR APPLICATION'
        }]

        staged = [{
            'id': 'conv-1',
            'fields': {'DAR STATUS': 'DAR FINALITY SECURED'}
        }]

        res = self.app.save_staged_conversion_records(staged)
        self.assertTrue(res['success'])
        self.assertEqual(len(self.app.log_records), 1)
        log = self.app.log_records[0]
        self.assertEqual(log['field_name'], 'DAR STATUS')
        self.assertEqual(log['old_value'], 'DAR FOR APPLICATION')
        self.assertEqual(log['new_value'], 'DAR FINALITY SECURED')
        self.assertEqual(log['cr63f_registeredowner'], 'Maria Santos')
        self.assertEqual(log['Registered Owner'], 'Maria Santos')
        self.assertEqual(log['cr63f_authorizedrepresentative'], 'Carlos Santos')
        self.assertEqual(log['cr63f_province'], 'Batangas')
        self.assertEqual(log['cr63f_municipality'], 'Tuy')
        self.assertEqual(log['cr63f_lotareaha'], '0.85')
        self.assertEqual(log['new_contracttype'], 'JOINT VENTURE')
        self.assertEqual(log['new_mode'], 'DIRECT PURCHASE')
        self.assertEqual(log['cr63f_control1'], 'CTRL-601')

    def test_derive_field_diffs_by_control_number(self):
        """Ensure derive_field_diffs_from_snapshots uses Control Number as key to compare changes even across area index modifications."""
        from app import derive_field_diffs_from_snapshots
        raw_snapshots = [
            {
                '__id': 'snap-ctrl-1',
                'createdon': '2026-09-24T08:00:00Z',
                'cr63f_control1': 'CTRL-888',
                'cr63f_areaindex': 'OLD-AI-01',
                'cr63f_totalcontractprice': '1000000',
                'cr63f_registeredowner': 'Pedro Santos'
            },
            {
                '__id': 'snap-ctrl-2',
                'createdon': '2026-09-24T09:00:00Z',
                'cr63f_control1': 'CTRL-888',
                'cr63f_areaindex': 'NEW-AI-01',
                'cr63f_totalcontractprice': '1500000',
                'cr63f_registeredowner': 'Pedro Santos'
            }
        ]
        derived = derive_field_diffs_from_snapshots(raw_snapshots)
        # Initial entry + 2 diffs (AREA-INDEX from OLD-AI-01 to NEW-AI-01, Total Contract Price from 1M to 1.5M)
        self.assertEqual(len(derived), 3)
        diff_fields = [d['field_name'] for d in derived]
        self.assertIn('Total Contract Price', diff_fields)
        self.assertIn('AREA-INDEX', diff_fields)
        self.assertIn('Initial Entry', diff_fields)

        tcp_diff = next(d for d in derived if d['field_name'] == 'Total Contract Price')
        self.assertEqual(tcp_diff['old_value'], '₱1,000,000.00')
        self.assertEqual(tcp_diff['new_value'], '₱1,500,000.00')
        self.assertEqual(tcp_diff['cr63f_control1'], 'CTRL-888')

        ai_diff = next(d for d in derived if d['field_name'] == 'AREA-INDEX')
        self.assertEqual(ai_diff['old_value'], 'OLD-AI-01')
        self.assertEqual(ai_diff['new_value'], 'NEW-AI-01')
        self.assertEqual(ai_diff['cr63f_control1'], 'CTRL-888')

if __name__ == '__main__':
    unittest.main()
