"""Endorsed Date defaults to today for new entries and Update from References."""

import unittest
from datetime import datetime
from unittest.mock import MagicMock, patch

from app import Api
from loa_engine import get_work_week


def _api():
    with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
        app = Api()
    app.loa_records = []
    app.log_records = []
    finance = [{'TYPE': 'SALE', 'Area Index': 'AI-1',
                'Total Contract Price': '1,000', 'Total Payment Deductible to TCP': '500'}]
    app.finance_rows = finance
    app.finance_engine.load_rows(finance)
    return app


def _record(**extra):
    rec = {'__id': 'r1', 'cr63f_type': 'SALE', 'cr63f_areaindex': 'AI-1',
           'cr63f_totalcontractprice': '1000', 'cr63f_totalpaymentdeductibletotcp': '100',
           'cr63f_balance': '900', 'cr63f_paid': '10.00%'}
    rec.update(extra)
    return rec


class TestEndorsedDateDefault(unittest.TestCase):
    def setUp(self):
        self.today = datetime.now().strftime('%m/%d/%Y')

    def test_reference_update_fills_blank_endorsed_date(self):
        res = _api().populate_loa_records_from_references(records=[_record()])
        rec = res['records'][0]
        self.assertEqual(rec['cr63f_endorseddate'], self.today)
        self.assertEqual(rec['work_week'], get_work_week(self.today))
        self.assertIn('cr63f_endorseddate', rec['__changed_fields'])
        self.assertIn('Endorsed Date', [log['field_name'] for log in res['diff_summary']])

    def test_reference_update_never_overwrites_existing_date(self):
        res = _api().populate_loa_records_from_references(records=[_record(cr63f_endorseddate='03/15/2026')])
        rec = res['records'][0]
        self.assertEqual(rec['cr63f_endorseddate'], '03/15/2026')
        self.assertNotIn('cr63f_endorseddate', rec['__changed_fields'])

    def test_placeholder_1900_date_is_never_overwritten(self):
        # Decided 2026-10-08: only a truly blank Endorsed Date is stamped.
        res = _api().populate_loa_records_from_references(records=[_record(cr63f_endorseddate='01/01/1900')])
        self.assertEqual(res['records'][0]['cr63f_endorseddate'], '01/01/1900')
        self.assertNotIn('cr63f_endorseddate', res['records'][0]['__changed_fields'])

    def test_no_reference_change_leaves_date_blank(self):
        # Already up to date with the reference: nothing changes, no date is added.
        res = _api().populate_loa_records_from_references(records=[_record(
            cr63f_totalpaymentdeductibletotcp='500', cr63f_balance='500', cr63f_paid='50.00%')])
        self.assertEqual(res['updated_count'], 0)
        self.assertFalse(res['records'][0].get('cr63f_endorseddate'))

    def test_persisted_reference_update_saves_date(self):
        app = _api()
        app.dv_client = MagicMock()
        app.dv_client.config = {}
        app.dv_client.fetch_records.return_value = []
        app.dv_client.save_loa_record.return_value = {'success': True, 'id': 'r1'}
        app.dv_client.save_log_record.return_value = {'success': True, 'id': 'log-1'}
        app.loa_records = [_record()]
        res = app.populate_loa_records_from_references(persist_to_dataverse=True)
        self.assertTrue(res['success'], res.get('error'))
        self.assertEqual(app.dv_client.save_loa_record.call_args[0][0]['cr63f_endorseddate'], self.today)

    def test_new_entry_keeps_a_typed_date(self):
        app = _api()
        app.dv_client = MagicMock()
        app.dv_client.config = {}
        app.dv_client.fetch_records.return_value = []
        app.dv_client.save_loa_record.return_value = {'success': True, 'action': 'create', 'id': 'n1'}
        app.dv_client.save_log_record.return_value = {'success': True, 'id': 'log-1'}
        app.save_loa_record_entry({'cr63f_areaindex': 'AI-7', 'cr63f_control1': 'C-7', 'cr63f_acquisitionsublevel': 'L1',
                                   'cr63f_endorseddate': '02/03/2026'})
        self.assertEqual(app.dv_client.save_loa_record.call_args[0][0]['cr63f_endorseddate'], '02/03/2026')


if __name__ == '__main__':
    unittest.main()
