"""Regression tests for the 2026-10-07 code review fixes.

1. Entry Logs ordering uses real time, not text.
2. Secured-status ranking rejects negated statuses.
3. Stored % Paid is never rescaled.
4. Update from References always saves Balance and % Paid.
5. Overview work weeks carry the work-week year and match exactly.
"""

import unittest
from datetime import datetime
from decimal import Decimal
from unittest.mock import MagicMock, patch

from app import Api, derive_field_diffs_from_snapshots
from loa_engine import (
    conversion_rank,
    format_percent,
    get_work_week,
    get_work_week_label,
    paid_percent_for_storage,
    parse_log_datetime,
)


def _snapshot(log_id, createdon, tcp, action='MANUAL_EDIT'):
    return {
        '__id': log_id, 'createdon': createdon, 'action_type': action,
        'cr63f_control1': 'C-1', 'cr63f_areaindex': 'AI-1',
        'field_name': 'Full Record', 'old_value': 'Previous Snapshot', 'new_value': 'Updated Snapshot',
        'cr63f_totalcontractprice': tcp,
    }


class TestLogOrdering(unittest.TestCase):
    def test_month_boundary_does_not_drop_a_change(self):
        # Dataverse display text for createdon: '10/2/...' sorted before '9/29/...' as text.
        logs = derive_field_diffs_from_snapshots([
            _snapshot('a', '9/30/2026 4:00 PM', '100'),
            _snapshot('b', '10/2/2026 9:00 AM', '200'),
            _snapshot('c', '9/29/2026 1:00 PM', '50', action='BASELINE'),
        ])
        changes = [(l['old_value'], l['new_value']) for l in logs if l['field_name'] == 'Total Contract Price']
        self.assertEqual(changes, [('₱100.00', '₱200.00'), ('₱50.00', '₱100.00')])

    def test_year_boundary_order(self):
        logs = derive_field_diffs_from_snapshots([
            _snapshot('a', '12/30/2026 9:00 AM', '100', action='BASELINE'),
            _snapshot('b', '1/5/2027 9:00 AM', '300'),
        ])
        changes = [(l['old_value'], l['new_value']) for l in logs if l['field_name'] == 'Total Contract Price']
        self.assertEqual(changes, [('₱100.00', '₱300.00')])

    def test_raw_createdon_breaks_same_minute_ties(self):
        first = _snapshot('x', '10/2/2026 9:00 AM', '100', action='BASELINE')
        second = _snapshot('y', '10/2/2026 9:00 AM', '200')
        third = _snapshot('z', '10/2/2026 9:00 AM', '300')
        first['__createdon_raw'] = '2026-10-02T01:00:05Z'
        second['__createdon_raw'] = '2026-10-02T01:00:40Z'
        third['__createdon_raw'] = '2026-10-02T01:00:20Z'
        logs = derive_field_diffs_from_snapshots([second, third, first])
        changes = [(l['old_value'], l['new_value']) for l in logs if l['field_name'] == 'Total Contract Price']
        self.assertCountEqual(changes, [('₱100.00', '₱300.00'), ('₱300.00', '₱200.00')])

    def test_parse_log_datetime_formats(self):
        self.assertEqual(parse_log_datetime('10/2/2026 9:05 PM'), datetime(2026, 10, 2, 21, 5))
        self.assertEqual(parse_log_datetime('10/02/2026 09:05 AM'), datetime(2026, 10, 2, 9, 5))
        self.assertIsNotNone(parse_log_datetime('2026-10-02T01:00:05Z'))
        self.assertIsNone(parse_log_datetime(''))
        self.assertIsNone(parse_log_datetime('not a date'))


class TestSecuredRank(unittest.TestCase):
    def test_secured_statuses(self):
        for status in ('DAR SECURED', 'DAR FINALITY SECURED', 'EXEMPT', 'APPROVED', 'CONVERTED', 'dar  secured'):
            self.assertEqual(conversion_rank(status), 3, status)

    def test_negated_statuses_are_not_secured(self):
        for status in ('NOT SECURED', 'UNSECURED', 'DISAPPROVED', 'NOT YET CONVERTED',
                       'PENDING APPROVED', 'APPROVED - FOR INSPECTION', 'DENIED'):
            self.assertEqual(conversion_rank(status), 2, status)

    def test_not_started(self):
        for status in ('', None, '-', 'NOT YET ENDORSED', 'N/A'):
            self.assertEqual(conversion_rank(status), 1, status)

    def test_api_wrapper_uses_shared_rank(self):
        self.assertEqual(Api._conversion_rank('NOT SECURED'), 2)


class TestPercentScale(unittest.TestCase):
    def test_small_percentages_are_not_multiplied(self):
        self.assertEqual(format_percent(0.5), '0.50%')
        self.assertEqual(format_percent('0.99'), '0.99%')
        self.assertEqual(format_percent(1.0), '1.00%')
        self.assertEqual(format_percent('45.5%'), '45.50%')

    def test_storage_value_matches_dataverse_precision(self):
        # cr63f_paid has precision 1 in Dataverse.
        self.assertEqual(paid_percent_for_storage('300', '100'), Decimal('33.3'))
        self.assertEqual(paid_percent_for_storage('3', '2'), Decimal('66.7'))
        self.assertEqual(paid_percent_for_storage('0', '5'), Decimal('100.0'))
        self.assertIsNone(paid_percent_for_storage('', '5'))


class TestAlwaysSavePaid(unittest.TestCase):
    def setUp(self):
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            self.app = Api()
        finance = [{'TYPE': 'SALE', 'Area Index': 'AI-1',
                    'Total Contract Price': '1,000,000.00', 'Total Payment Deductible to TCP': '250,000.00'}]
        self.app.finance_rows = finance
        self.app.finance_engine.load_rows(finance)

    def _record(self, **overrides):
        rec = {'__id': 'rec-1', 'cr63f_type': 'SALE', 'cr63f_areaindex': 'AI-1',
               'cr63f_totalcontractprice': '1000000', 'cr63f_totalpaymentdeductibletotcp': '250000',
               'cr63f_balance': '750000', 'cr63f_paid': '25.00%'}
        rec.update(overrides)
        return rec

    def test_up_to_date_record_is_untouched(self):
        res = self.app.populate_loa_records_from_references(records=[self._record()])
        self.assertEqual(res['updated_count'], 0)

    def test_stale_percent_is_corrected_even_when_tcp_unchanged(self):
        res = self.app.populate_loa_records_from_references(records=[self._record(cr63f_paid='10.00%')])
        self.assertEqual(res['updated_count'], 1)
        rec = res['records'][0]
        self.assertIn('cr63f_paid', rec['__changed_fields'])
        self.assertEqual(rec['cr63f_paid'], '25.0%')

    def test_value_rounded_by_dataverse_is_not_a_change(self):
        # Dataverse stores 33.3 for 33.33; the next run must not see a change.
        finance = [{'TYPE': 'SALE', 'Area Index': 'AI-3', 'Total Contract Price': '300',
                    'Total Payment Deductible to TCP': '100'}]
        self.app.finance_rows = finance
        self.app.finance_engine.load_rows(finance)
        res = self.app.populate_loa_records_from_references(records=[{
            '__id': 'r3', 'cr63f_type': 'SALE', 'cr63f_areaindex': 'AI-3',
            'cr63f_totalcontractprice': '300', 'cr63f_totalpaymentdeductibletotcp': '100',
            'cr63f_balance': '200', 'cr63f_paid': '33.3'}])
        self.assertEqual(res['updated_count'], 0)

    def test_stale_balance_is_corrected(self):
        res = self.app.populate_loa_records_from_references(records=[self._record(cr63f_balance='999')])
        self.assertEqual(res['updated_count'], 1)
        self.assertIn('cr63f_balance', res['records'][0]['__changed_fields'])

    def test_persist_writes_percent(self):
        self.app.dv_client = MagicMock()
        self.app.dv_client.config = {}
        self.app.dv_client.fetch_records.return_value = []
        self.app.dv_client.save_loa_record.return_value = {'success': True, 'id': 'rec-1'}
        self.app.dv_client.save_log_record.return_value = {'success': True, 'id': 'log-1'}
        self.app.loa_records = [self._record(cr63f_totalpaymentdeductibletotcp='100000',
                                             cr63f_balance='900000', cr63f_paid='10.00%')]
        res = self.app.populate_loa_records_from_references(persist_to_dataverse=True)
        self.assertTrue(res['success'], res.get('error'))
        saved = self.app.dv_client.save_loa_record.call_args[0][0]
        self.assertEqual(saved['cr63f_totalpaymentdeductibletotcp'], Decimal('250000.00'))
        self.assertEqual(saved['cr63f_balance'], Decimal('750000.00'))
        self.assertEqual(saved['cr63f_paid'], Decimal('25.00'))


class TestWorkWeekYear(unittest.TestCase):
    def test_label_uses_work_week_year(self):
        # Wed Dec 29 2027 - Tue Jan 4 2028: Saturday is Jan 1 2028.
        self.assertEqual(get_work_week_label(datetime(2027, 12, 29)), '2028 · WW-01')
        self.assertEqual(get_work_week_label(datetime(2027, 1, 1)), '2027 · WW-01')
        self.assertNotEqual(get_work_week_label(datetime(2027, 12, 30)), get_work_week_label(datetime(2027, 1, 1)))
        # Jan 1 2029 belongs to the last week of 2028.
        self.assertEqual(get_work_week_label(datetime(2029, 1, 1)), '2028 · WW-53')
        self.assertEqual(get_work_week(datetime(2026, 10, 7)), 'WW-41')

    def test_every_year_starts_at_ww01_without_gaps(self):
        from datetime import timedelta
        for year in range(2018, 2040):
            day, labels = datetime(year - 1, 12, 20), []
            while day <= datetime(year + 1, 1, 10):
                label = get_work_week_label(day)
                if label.startswith(str(year)) and (not labels or labels[-1] != label):
                    labels.append(label)
                day += timedelta(days=1)
            numbers = [int(label[-2:]) for label in labels]
            self.assertEqual(numbers, list(range(1, len(numbers) + 1)), year)
        self.assertEqual(get_work_week_label(datetime(2029, 1, 3)), '2029 · WW-01')

    def _app_with_logs(self, times):
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            app = Api()
        app.loa_records = []
        app.log_records = [
            {'__id': f'l{i}', 'action_type': 'MANUAL_EDIT', 'cr63f_control1': f'C{i}', 'area_index': f'A{i}',
             'field_name': 'Mode', 'old_value': 'A', 'new_value': 'B', 'log_time': t}
            for i, t in enumerate(times)
        ]
        return app

    def test_week_filter_is_exact(self):
        # WW-01 (Jan 1 2026) must not pick up WW-10..WW-19 or another year's WW-01.
        app = self._app_with_logs(['01/01/2026 10:00 AM', '03/05/2026 10:00 AM', '01/04/2027 10:00 AM'])
        res = app.get_dashboard_analytics(work_week='2026 · WW-01')
        self.assertEqual(res['selected_work_week'], '2026 · WW-01')
        self.assertEqual(res['week_summary']['accounts_changed'], 1)
        # 'WW-1' is not a real week label, so it must not match WW-10..WW-19.
        res = app.get_dashboard_analytics(work_week='WW-1')
        self.assertEqual(res['week_summary']['accounts_changed'], 0)

    def test_bare_week_picks_most_recent_year(self):
        app = self._app_with_logs(['01/01/2026 10:00 AM', '01/04/2027 10:00 AM'])
        res = app.get_dashboard_analytics(work_week='WW-01')
        self.assertEqual(res['selected_work_week'], '2027 · WW-01')
        self.assertEqual(res['week_summary']['accounts_changed'], 1)

    def test_year_end_days_are_not_merged_with_january(self):
        app = self._app_with_logs(['01/01/2027 10:00 AM', '12/30/2027 10:00 AM'])
        weeks = app.get_dashboard_analytics(work_week='ALL')['available_work_weeks']
        self.assertEqual(weeks, ['2028 · WW-01', '2027 · WW-01'])


if __name__ == '__main__':
    unittest.main()
