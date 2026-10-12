"""Manual saves derive Balance and % Paid on the server (review item 4, 2026-10-08).

The grid and drawer also calculate these in JavaScript, but the stored values
must not depend on what the client sent. The helper below is the single rule.
"""

import unittest
from unittest.mock import MagicMock, patch

from app import Api, _derived_payment_fields


class TestDerivedPaymentFields(unittest.TestCase):
    def test_balance_and_percent_from_amounts(self):
        fields = _derived_payment_fields({
            'cr63f_totalcontractprice': '1000000', 'cr63f_totalpaymentdeductibletotcp': '250000'})
        self.assertEqual(fields, {'cr63f_balance': '750000.00', 'Balance': '750000.00',
                                  'cr63f_paid': '25.0', '% Paid': '25.0'})

    def test_percent_rounds_to_dataverse_precision(self):
        fields = _derived_payment_fields({'Total Contract Price': '300', 'Total Payment Deductible to TCP': '100'})
        self.assertEqual(fields['cr63f_paid'], '33.3')

    def test_overpayment_balance_is_zero_and_percent_exceeds_100(self):
        fields = _derived_payment_fields({'cr63f_totalcontractprice': '100', 'cr63f_totalpaymentdeductibletotcp': '150'})
        self.assertEqual(fields['cr63f_balance'], '0.00')
        self.assertEqual(fields['cr63f_paid'], '150.0')

    def test_zero_contract_price_with_deductible_is_fully_paid(self):
        fields = _derived_payment_fields({'cr63f_totalcontractprice': '0', 'cr63f_totalpaymentdeductibletotcp': '5'})
        self.assertEqual(fields['cr63f_paid'], '100.0')

    def test_blank_counterpart_counts_as_zero(self):
        fields = _derived_payment_fields({'cr63f_totalcontractprice': '1,000.00', 'cr63f_totalpaymentdeductibletotcp': ''})
        self.assertEqual(fields['cr63f_balance'], '1000.00')
        self.assertEqual(fields['cr63f_paid'], '0.0')

    def test_nothing_derived_without_amounts(self):
        self.assertEqual(_derived_payment_fields({'cr63f_areaindex': 'AI-1'}), {})


class TestManualSavesDeriveValues(unittest.TestCase):
    def setUp(self):
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            self.app = Api()
        self.app.dv_client = MagicMock()
        self.app.dv_client.config = {}
        self.app.dv_client.save_loa_record.return_value = {'success': True, 'id': 'rec-1'}
        self.app.dv_client.save_log_record.return_value = {'success': True, 'id': 'log-1'}
        self.app.dv_client.fetch_records.return_value = []
        self.app.loa_records = [{
            '__id': 'rec-1', 'cr63f_control1': 'C-1', 'cr63f_areaindex': 'AI-1', 'cr63f_type': 'SALE',
            'cr63f_totalcontractprice': '1000000', 'cr63f_totalpaymentdeductibletotcp': '250000',
            'cr63f_balance': '750000', 'cr63f_paid': '25.0',
        }]

    def _saved_master_payload(self):
        return self.app.dv_client.save_loa_record.call_args[0][0]

    def test_drawer_update_ignores_client_balance_and_percent(self):
        record_data = {
            'cr63f_areaindex': 'AI-1', 'cr63f_control1': 'C-1',
            'cr63f_totalcontractprice': '2,000,000.00', 'Total Contract Price': '2,000,000.00',
            'cr63f_totalpaymentdeductibletotcp': '500,000.00', 'Total Payment Deductible to TCP': '500,000.00',
            'cr63f_balance': '999', 'Balance': '999', 'cr63f_paid': '10.00%', '% Paid': '10.00%',
        }
        with patch.object(self.app, '_write_baseline_if_missing', return_value=None):
            res = self.app.save_loa_record_entry(record_data, record_id='rec-1')
        self.assertTrue(res['success'], res)
        payload = self._saved_master_payload()
        self.assertEqual(payload['cr63f_balance'], '1500000.00')
        self.assertEqual(payload['Balance'], '1500000.00')
        self.assertEqual(payload['cr63f_paid'], '25.0')
        self.assertEqual(self.app.loa_records[0]['__display_percent_paid'], '25.00%')

    def test_drawer_new_entry_derives_values(self):
        record_data = {
            'cr63f_control1': 'C-9', 'cr63f_areaindex': 'AI-9', 'cr63f_type': 'LAND',
            'cr63f_acquisitionsublevel': 'L1.1',
            'cr63f_totalcontractprice': '300', 'cr63f_totalpaymentdeductibletotcp': '100',
            'cr63f_balance': '999', 'cr63f_paid': '1.00%',
        }
        self.app.loa_records = []
        res = self.app.save_loa_record_entry(record_data)
        self.assertTrue(res['success'], res)
        payload = self._saved_master_payload()
        self.assertEqual(payload['cr63f_balance'], '200.00')
        self.assertEqual(payload['cr63f_paid'], '33.3')

    def test_staged_amount_change_recalculates_balance_and_percent(self):
        staged = [{'id': 'rec-1', 'action_type': 'MANUAL_EDIT', 'fields': {
            'cr63f_totalcontractprice': '2,000,000.00', 'Total Contract Price': '2,000,000.00',
            'cr63f_balance': '1', 'Balance': '1', 'cr63f_paid': '3.0%', '% Paid': '3.0%'}}]
        with patch.object(self.app, '_write_baseline_if_missing', return_value=None):
            res = self.app.save_staged_records(staged)
        self.assertTrue(res['success'], res)
        # Stored Deductible is 250,000 and the staged TCP is 2,000,000.
        payload = self._saved_master_payload()
        self.assertEqual(payload['cr63f_balance'], '1750000.00')
        self.assertEqual(payload['cr63f_paid'], '12.5')
        self.assertEqual(self.app.loa_records[0]['__display_percent_paid'], '12.50%')

    def test_staged_unrelated_edit_does_not_touch_payment_fields(self):
        staged = [{'id': 'rec-1', 'action_type': 'MANUAL_EDIT', 'fields': {'new_mode': 'SECURED SALE'}}]
        with patch.object(self.app, '_write_baseline_if_missing', return_value=None):
            res = self.app.save_staged_records(staged)
        self.assertTrue(res['success'], res)
        payload = self._saved_master_payload()
        self.assertEqual(set(payload), {'new_mode'})


if __name__ == '__main__':
    unittest.main()


class TestPaidDisplayUsesCalculation(unittest.TestCase):
    """Review item 3: a stored fraction (1 = 100%) must not show as 1.00%."""

    def test_loaded_record_shows_calculated_percent(self):
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            app = Api()
        app.dv_client = MagicMock()
        app.dv_client.config = {'loa_table': 'cr63f_loamonitoring'}
        app.dv_client.fetch_records.return_value = [{
            '__id': 'rec-1', 'cr63f_totalcontractprice': '1000', 'cr63f_totalpaymentdeductibletotcp': '1000',
            'cr63f_balance': '0', 'cr63f_paid': '1',
        }]
        app.dv_client.get_entity_attributes.return_value = {}
        res = app.fetch_loa_records()
        self.assertTrue(res['success'], res)
        self.assertEqual(res['records'][0]['__display_percent_paid'], '100.00%')

    def test_stored_value_is_shown_only_when_amounts_are_missing(self):
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            app = Api()
        app.dv_client = MagicMock()
        app.dv_client.config = {'loa_table': 'cr63f_loamonitoring'}
        app.dv_client.fetch_records.return_value = [{
            '__id': 'rec-2', 'cr63f_totalcontractprice': '', 'cr63f_totalpaymentdeductibletotcp': '',
            'cr63f_paid': '0.5',
        }]
        app.dv_client.get_entity_attributes.return_value = {}
        res = app.fetch_loa_records()
        self.assertNotIn('__display_percent_paid', res['records'][0])
