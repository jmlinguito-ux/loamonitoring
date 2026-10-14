"""Regression tests for the 2026-10-08 consistency and logic fixes (review items 1-13)."""

import unittest
from unittest.mock import MagicMock, patch

from app import Api, _derived_payment_fields, derive_field_diffs_from_snapshots
from loa_engine import FinanceLookupEngine


def _api():
    with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
        app = Api()
    app.dv_client = MagicMock()
    app.dv_client.config = {}
    app.dv_client.fetch_records.return_value = []
    app.dv_client.save_loa_record.return_value = {'success': True, 'id': 'rec-1'}
    app.dv_client.save_log_record.return_value = {'success': True, 'id': 'log-1'}
    return app


class TestBlankTcpClearsDerivedFields(unittest.TestCase):
    def test_blank_tcp_clears_balance_and_percent(self):
        fields = _derived_payment_fields({'cr63f_totalcontractprice': '', 'cr63f_totalpaymentdeductibletotcp': '5'})
        self.assertEqual(fields, {'cr63f_balance': None, 'Balance': None, 'cr63f_paid': None, '% Paid': None})

    def test_absent_amount_keys_derive_nothing(self):
        self.assertEqual(_derived_payment_fields({'cr63f_areaindex': 'AI-1'}), {})


class TestRevertIsAnAnchor(unittest.TestCase):
    def test_change_after_revert_is_not_compared_with_pre_revert_state(self):
        def snap(log_id, created, tcp, mode, action='MANUAL_EDIT'):
            return {
                '__id': log_id, '__createdon_raw': created, 'createdon': created, 'action_type': action,
                'cr63f_control1': 'C-1', 'cr63f_areaindex': 'AI-1', 'field_name': 'Full Record',
                'old_value': 'Previous Snapshot', 'new_value': 'Updated Snapshot',
                'cr63f_totalcontractprice': tcp, 'new_mode': mode,
            }
        revert_row = {
            '__id': 'rev', '__createdon_raw': '2026-10-02T01:00:20Z', 'createdon': '10/2/2026 9:00 AM',
            'action_type': 'REVERT', 'cr63f_control1': 'C-1', 'cr63f_areaindex': 'AI-1',
            'field_name': 'Total Contract Price', 'old_value': '200', 'new_value': '100',
            'cr63f_totalcontractprice': '100', 'new_mode': 'SALE',
        }
        logs = derive_field_diffs_from_snapshots([
            snap('a', '2026-10-02T01:00:00Z', '100', 'SALE', action='BASELINE'),
            snap('b', '2026-10-02T01:00:10Z', '200', 'SALE'),
            revert_row,
            snap('c', '2026-10-02T01:00:30Z', '100', 'LEASE'),
        ])
        tcp_rows = [l for l in logs if l['field_name'] == 'Total Contract Price']
        # The real edit 100 -> 200 and the REVERT 200 -> 100 only. Without the
        # anchor, the post-revert save would add a spurious 200 -> 100 edit.
        self.assertEqual(sorted(l['action_type'] for l in tcp_rows), ['MANUAL_EDIT', 'REVERT'])
        mode_rows = [l for l in logs if l['field_name'] == 'Mode']
        self.assertEqual(len(mode_rows), 1)
        self.assertEqual(mode_rows[0]['action_type'], 'MANUAL_EDIT')


class TestNoOpStagedEditIsSkipped(unittest.TestCase):
    def test_staged_value_equal_to_stored_saves_nothing(self):
        app = _api()
        app.loa_records = [{'__id': 'rec-1', 'cr63f_control1': 'C-1', 'cr63f_areaindex': 'AI-1',
                            'cr63f_totalcontractprice': '1000', 'cr63f_totalpaymentdeductibletotcp': '250'}]
        with patch.object(app, '_write_baseline_if_missing', return_value=None):
            res = app.save_staged_records([{'id': 'rec-1', 'action_type': 'MANUAL_EDIT',
                                            'fields': {'cr63f_areaindex': 'AI-1'}}])
        self.assertTrue(res['success'], res)
        app.dv_client.save_loa_record.assert_not_called()
        app.dv_client.save_log_record.assert_not_called()


class TestStagedPaymentFieldRecalculates(unittest.TestCase):
    def test_typed_balance_is_replaced_by_calculated_value(self):
        app = _api()
        app.loa_records = [{'__id': 'rec-1', 'cr63f_control1': 'C-1', 'cr63f_areaindex': 'AI-1',
                            'cr63f_totalcontractprice': '1000', 'cr63f_totalpaymentdeductibletotcp': '250'}]
        with patch.object(app, '_write_baseline_if_missing', return_value=None):
            res = app.save_staged_records([{'id': 'rec-1', 'action_type': 'MANUAL_EDIT',
                                            'fields': {'cr63f_balance': '1', 'Balance': '1'}}])
        self.assertTrue(res['success'], res)
        # Balance alone is a calculated field: nothing real changed, so nothing is saved.
        app.dv_client.save_loa_record.assert_not_called()


class TestBlankTypeOrModeRefused(unittest.TestCase):
    def test_blank_type_cannot_replace_stored_type(self):
        app = _api()
        app.loa_records = [{'__id': 'rec-1', 'cr63f_control1': 'C-1', 'cr63f_areaindex': 'AI-1',
                            'cr63f_type': 'LAND - SPECIAL', 'new_mode': 'SS - STP'}]
        res = app.save_loa_record_entry({'cr63f_type': '', 'cr63f_areaindex': 'AI-1'}, record_id='rec-1')
        self.assertFalse(res['success'])
        self.assertIn('TYPE cannot be cleared', res['error'])
        app.dv_client.save_loa_record.assert_not_called()

    def test_blank_mode_cannot_replace_stored_mode(self):
        app = _api()
        app.loa_records = [{'__id': 'rec-1', 'cr63f_control1': 'C-1', 'cr63f_areaindex': 'AI-1',
                            'cr63f_type': 'LAND', 'new_mode': 'SS - STP'}]
        res = app.save_loa_record_entry({'new_mode': '', 'cr63f_areaindex': 'AI-1'}, record_id='rec-1')
        self.assertFalse(res['success'])
        self.assertIn('Mode cannot be cleared', res['error'])


class TestBlankModeGroupedAsBlank(unittest.TestCase):
    def test_accounts_without_mode_are_grouped_as_blank(self):
        app = _api()
        app.loa_records = [{'cr63f_control1': 'C-1'}, {'new_mode': 'SS - STP'}]
        app.log_records = []
        res = app.get_dashboard_analytics('ALL')
        modes = {m['mode']: m['count'] for m in res['modes_distribution']}
        self.assertEqual(modes, {'BLANK': 1, 'SS - STP': 1})


class TestBatchBaselineUsesKnownKeys(unittest.TestCase):
    def test_known_account_gets_no_query_and_no_baseline(self):
        app = _api()
        baseline = app._write_baseline_if_missing({'cr63f_control1': 'C-1', 'cr63f_areaindex': 'AI-1'},
                                                  '10/08/2026 09:00 AM', known_keys={'CTRL:C-1'})
        self.assertIsNone(baseline)
        app.dv_client.fetch_records.assert_not_called()
        app.dv_client.save_log_record.assert_not_called()

    def test_new_account_gets_one_baseline_and_is_recorded(self):
        app = _api()
        known = set()
        first = app._write_baseline_if_missing({'cr63f_control1': 'C-2', 'cr63f_areaindex': 'AI-2'},
                                               '10/08/2026 09:00 AM', known_keys=known)
        second = app._write_baseline_if_missing({'cr63f_control1': 'C-2', 'cr63f_areaindex': 'AI-2'},
                                                '10/08/2026 09:00 AM', known_keys=known)
        self.assertIsNotNone(first)
        self.assertIsNone(second)
        self.assertEqual(app.dv_client.save_log_record.call_count, 1)
        app.dv_client.fetch_records.assert_not_called()


class TestFinanceLookupFixes(unittest.TestCase):
    def test_untyped_row_with_corridor_index_matches_row_account(self):
        engine = FinanceLookupEngine()
        engine.load_rows([{'Corridor Index': 'CI-9', 'Total Contract Price': '500', 'Total Payment Deductible to TCP': '100'}])
        res = engine.lookup('ROW', 'CI-9', '', '')
        self.assertTrue(res.get('matched'), res)

    def test_agreeing_duplicate_titles_match(self):
        engine = FinanceLookupEngine()
        row = {'TYPE': 'SALE', 'Title Mod': 'TCT-066-2019006115',
               'Total Contract Price': '1,000', 'Total Payment Deductible to TCP': '100'}
        engine.load_rows([dict(row), dict(row)])
        res = engine.lookup('SALE', '', '', 'TCT-066-2019006115', 'SECURED SALE')
        self.assertTrue(res.get('matched'), res)
        self.assertFalse(res.get('ambiguous'), res)


if __name__ == '__main__':
    unittest.main()


class TestStagedTitleChangeIsLogged(unittest.TestCase):
    """2026-10-14: a grid edit of Title No + Registered Owner logged only the owner,
    because the snapshot was stamped with the row's old Title No."""

    def test_title_and_owner_both_appear_in_entry_logs(self):
        app = _api()
        before = {'__id': 'rec-1', 'cr63f_control1': 'BAT-NAS-004411', 'cr63f_areaindex': 'NAS-272',
                  'cr63f_indexno': 'I-1', 'cr63f_titleno': 'TCT-OLD', 'cr63f_registeredowner': 'JUAN'}
        app.loa_records = [dict(before)]
        baseline = {**before, 'cr63f_actiontype': 'BASELINE', 'cr63f_fieldchanged': 'Full Record',
                    'cr63f_details': '[MASTER:rec-1] [SOURCE:BASELINE]', 'cr63f_loamonitoringid': 'rec-1',
                    '__createdon_raw': '2026-10-14T08:05:00Z', '__id': 'base-1'}
        with patch.object(app, '_write_baseline_if_missing', return_value=None):
            res = app.save_staged_records([{'id': 'rec-1', 'action_type': 'MANUAL_EDIT', 'fields': {
                'cr63f_titleno': 'TCT-T-52585', 'cr63f_registeredowner': 'DOMINGO MALABANAN ET AL'}}])
        self.assertTrue(res['success'], res)
        snapshot = app.dv_client.save_log_record.call_args[0][0]
        self.assertEqual(snapshot['cr63f_titleno'], 'TCT-T-52585')
        snapshot = {**snapshot, '__createdon_raw': '2026-10-14T08:06:00Z', '__id': 'snap-1'}
        fields = {l['field_name'] for l in derive_field_diffs_from_snapshots([baseline, snapshot])}
        self.assertIn('Title No', fields)
        self.assertIn('Registered Owner', fields)


class TestEveryFieldChangeIsLogged(unittest.TestCase):
    """2026-10-14: every editable field, changed or cleared, from the grid or the
    drawer, appears in Entry Logs. A cleared field used to read back the old value
    from a display-name copy ('Registered Owner', 'Authorized Rep', 'Status' ...)."""

    BEFORE = {
        'cr63f_control1': 'C-1', 'cr63f_type': 'SALE', 'cr63f_indexno': 'I-1', 'cr63f_areaindex': 'AI-1',
        'cr63f_titleno': 'T-1', 'cr63f_registeredowner': 'OWN', 'cr63f_authorizedrepresentative': 'REP',
        'cr63f_entity': 'ENT', 'cr63f_province': 'PROV', 'cr63f_municipality': 'MUN',
        'cr63f_dbadjustedbarangay': 'BRGY', 'cr63f_lotareaha': '1.5', 'new_contracttype': 'CT',
        'new_mode': 'SECURED SALE', 'cr63f_conversionstatus': 'CS', 'cr63f_acquisitionsublevel': 'ASL',
        'cr63f_rccprice': '100', 'cr63f_rccterms': 'TERMS', 'cr63f_endorseddate': '10/01/2026',
        'cr63f_secureddate': '10/02/2026', 'cr63f_securedendorsedate': '10/03/2026',
        'cr63f_totalcontractprice': '1000', 'cr63f_totalpaymentdeductibletotcp': '200',
    }
    NOT_CLEARABLE = {'cr63f_control1', 'cr63f_type', 'new_mode'}

    def _record(self):
        from app import LOG_COMPARE_SPECS
        rec = {'__id': 'rec-1', '__etag': 'W/"1"', **self.BEFORE}
        for _canon, keys, _type in LOG_COMPARE_SPECS:
            if keys[0] in self.BEFORE:
                for alias in keys[1:]:
                    rec[alias] = self.BEFORE[keys[0]]
        return rec

    def _logged_fields(self, app, record):
        baseline = {**record, 'cr63f_actiontype': 'BASELINE', 'cr63f_fieldchanged': 'Full Record',
                    'cr63f_details': '[MASTER:rec-1] [SOURCE:BASELINE]', 'cr63f_loamonitoringid': 'rec-1',
                    '__createdon_raw': '2026-10-14T08:05:00Z', '__id': 'base'}
        snapshot = {**app.dv_client.save_log_record.call_args[0][0],
                    '__createdon_raw': '2026-10-14T08:06:00Z', '__id': 'snap'}
        return {l['field_name'] for l in derive_field_diffs_from_snapshots([baseline, snapshot])}

    @staticmethod
    def _new_value(value):
        if '/' in value:
            return '11/11/2026'
        return str(float(value) + 1) if value[0].isdigit() else value + 'X'

    def _cases(self):
        from app import LOG_COMPARE_SPECS
        for canon, keys, _type in LOG_COMPARE_SPECS:
            key = keys[0]
            if key not in self.BEFORE:
                continue
            yield canon, keys, self._new_value(self.BEFORE[key])
            if key not in self.NOT_CLEARABLE:
                yield canon, keys, ''

    def test_grid_save_logs_every_field(self):
        for canon, keys, value in self._cases():
            with self.subTest(field=canon, value=value):
                app = _api()
                record = self._record()
                app.loa_records = [dict(record)]
                with patch.object(app, '_write_baseline_if_missing', return_value=None):
                    res = app.save_staged_records([{'id': 'rec-1', 'action_type': 'MANUAL_EDIT',
                                                    'fields': {keys[0]: value}}])
                self.assertTrue(res['success'], res)
                self.assertIn(canon, self._logged_fields(app, record))

    def test_drawer_save_logs_every_field(self):
        identity = ['cr63f_type', 'cr63f_indexno', 'cr63f_areaindex', 'cr63f_titleno', 'cr63f_control1']
        for canon, keys, value in self._cases():
            with self.subTest(field=canon, value=value):
                app = _api()
                app.dv_client.get_record.return_value = None
                app.dv_client.save_loa_record.return_value = {'success': True, 'id': 'rec-1', 'etag': 'W/"2"'}
                record = self._record()
                app.loa_records = [dict(record)]
                payload = {key: self.BEFORE[key] for key in identity}
                payload[keys[0]] = value
                payload[keys[1]] = value
                with patch.object(app, '_write_baseline_if_missing', return_value=None):
                    res = app.save_loa_record_entry(payload, 'rec-1')
                self.assertTrue(res.get('success'), res.get('error'))
                self.assertIn(canon, self._logged_fields(app, record))
