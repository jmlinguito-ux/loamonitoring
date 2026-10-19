"""Poss Rights Status / Poss Rights Area (Ha) / Possrights Cat / Cat Area (Ha)."""
import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
from app import Api, build_log_entry, _audit_values_equal, LOG_COMPARE_SPECS  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIELDS = [
    ('cr63f_possrightsstatus', 'Poss Rights Status', 'text'),
    ('cr63f_possrightsarea', 'Poss Rights Area (Ha)', 'number'),
    ('cr63f_possrightscat', 'Possrights Cat', 'text'),
    ('cr63f_catarea', 'Cat Area (Ha)', 'number'),
]


def read(*parts):
    with open(os.path.join(ROOT, *parts), encoding='utf-8') as f:
        return f.read()


class BackendTests(unittest.TestCase):
    def test_every_field_is_tracked_by_entry_logs(self):
        for logical, display, kind in FIELDS:
            spec = next(s for s in LOG_COMPARE_SPECS if s[0] == display)
            self.assertIn(logical, spec[1])
            self.assertIn(display, spec[1])
            self.assertEqual(spec[2], kind)
            self.assertEqual(app_module.CANONICAL_FIELD_MAP[logical], display)
            self.assertEqual(app_module.FIELD_NAME_TO_DATAVERSE_ATTR[display], logical)

    def test_snapshot_carries_all_names_of_each_field(self):
        rec = {'__id': 'r1', 'cr63f_control1': 'C-1', 'cr63f_possrightsstatus': 'GRANTED',
               'Cat Area (Ha)': '1.25'}
        entry = build_log_entry(action_type='MANUAL_EDIT', work_week='WW-1', now_str='10/19/2026 10:00 AM',
                                area_index='A', index_no='1', title_no='T', field_name='Full Record',
                                old_value='Previous Snapshot', new_value='Updated Snapshot', full_record=rec)
        self.assertEqual(entry['cr63f_possrightsstatus'], 'GRANTED')
        self.assertEqual(entry['Poss Rights Status'], 'GRANTED')
        self.assertEqual(entry['cr63f_catarea'], '1.25')
        self.assertEqual(entry['Cat Area (Ha)'], '1.25')
        self.assertEqual(entry['cr63f_possrightsarea'], '')

    def test_area_comparison_is_numeric(self):
        self.assertTrue(_audit_values_equal('Poss Rights Area (Ha)', '1', '1.00'))
        self.assertTrue(_audit_values_equal('cr63f_catarea', 1.5, '1.50'))
        self.assertFalse(_audit_values_equal('Cat Area (Ha)', '1.25', '1.26'))
        self.assertFalse(_audit_values_equal('Cat Area (Ha)', '', '0'))
        self.assertTrue(_audit_values_equal('Cat Area (Ha)', None, ''))

    def _api(self):
        api = Api()
        api.dv_client = MagicMock()
        api.dv_client.config = {'loa_table': 'cr63f_loamonitoring', 'logs_table': 'cr63f_loamonitoringlogs'}
        api.dv_client.save_loa_record.return_value = {'action': 'update', 'id': 'rec_1', 'etag': 'e2'}
        api.dv_client.save_log_record.return_value = {'id': 'log_1'}
        api.loa_records = [{'__id': 'rec_1', '__etag': 'e1', 'cr63f_control1': 'C-1', 'cr63f_areaindex': 'AI-1',
                            'cr63f_possrightsarea': '1.00', 'cr63f_possrightsstatus': 'OLD'}]
        api.log_records = []
        return api

    def test_saving_a_change_logs_a_field_row(self):
        api = self._api()
        with patch.object(api, '_write_baseline_if_missing', return_value=None):
            res = api.save_staged_records([{'id': 'rec_1', 'fields': {
                'cr63f_possrightsstatus': 'GRANTED', 'cr63f_catarea': '2.5'}}])
        self.assertTrue(res['success'], res)
        names = {row.get('field_name') for row in api.log_records}
        self.assertIn('Poss Rights Status', names)
        self.assertIn('Cat Area (Ha)', names)
        saved_snapshot = api.dv_client.save_log_record.call_args[0][0]
        self.assertEqual(saved_snapshot['cr63f_possrightsstatus'], 'GRANTED')
        self.assertEqual(saved_snapshot['cr63f_catarea'], '2.5')

    def test_same_area_in_another_format_is_not_a_change(self):
        api = self._api()
        with patch.object(api, '_write_baseline_if_missing', return_value=None):
            api.save_staged_records([{'id': 'rec_1', 'fields': {'cr63f_possrightsarea': '1'}}])
        api.dv_client.save_loa_record.assert_not_called()
        api.dv_client.save_log_record.assert_not_called()


class AreaDiffTests(unittest.TestCase):
    def test_number_text_normalises(self):
        from app import _norm_number_str
        self.assertEqual(_norm_number_str('1.00'), '1')
        self.assertEqual(_norm_number_str('1,234.50'), '1234.5')
        self.assertEqual(_norm_number_str(10), '10')
        self.assertEqual(_norm_number_str(''), '')
        self.assertEqual(_norm_number_str('-'), '')

    def test_reformatted_area_is_not_shown_as_a_change(self):
        from app import derive_field_diffs_from_snapshots
        base = {'__id': 'b', 'cr63f_actiontype': 'BASELINE', 'cr63f_fieldchanged': 'Full Record',
                'cr63f_details': '[MASTER:r1] [SOURCE:BASELINE]', '__createdon_raw': '2026-10-14T08:05:00Z',
                'cr63f_control1': 'C1', 'cr63f_catarea': '1.00', 'cr63f_lotareaha': '2.50', 'cr63f_possrightsarea': '3'}
        snap = dict(base, __id='s', cr63f_actiontype='MANUAL_EDIT', cr63f_details='[MASTER:r1] [SOURCE:MANUAL_EDIT]',
                    __createdon_raw='2026-10-14T08:06:00Z', cr63f_catarea='1', cr63f_lotareaha='2.5',
                    cr63f_possrightsarea='3.25')
        rows = derive_field_diffs_from_snapshots([base, snap])
        self.assertEqual([r['field_name'] for r in rows], ['Poss Rights Area (Ha)'])
        self.assertEqual((rows[0]['old_value'], rows[0]['new_value']), ('3.00', '3.25'))


class RevertTests(unittest.TestCase):
    def _api(self, current):
        api = Api()
        api.dv_client = MagicMock()
        api.dv_client.config = {'loa_table': 'cr63f_loamonitoring', 'logs_table': 'cr63f_loamonitoringlogs'}
        api.dv_client.get_record.return_value = dict(current)
        api.dv_client.save_loa_record.return_value = {'etag': 'e2'}
        api.dv_client.save_log_record.return_value = {'id': 'rev_log'}
        api.loa_records = [dict(current)]
        return api

    def test_revert_area_when_stored_value_is_formatted(self):
        current = {'__id': 'r1', '__etag': 'e1', 'cr63f_control1': 'C-1', 'cr63f_catarea': '2.50'}
        api = self._api(current)
        api.log_records = [{'__id': 'l1', 'cr63f_control1': 'C-1', 'field_name': 'Cat Area (Ha)',
                            'old_value': '1.25', 'new_value': '2.5'}]
        res = api.revert_log_record('l1')
        self.assertTrue(res.get('success'), res)
        self.assertEqual(api.dv_client.save_loa_record.call_args[0][0], {'cr63f_catarea': '1.25'})

    def test_revert_text_field_to_blank(self):
        current = {'__id': 'r1', '__etag': 'e1', 'cr63f_control1': 'C-1', 'cr63f_possrightsstatus': 'GRANTED'}
        api = self._api(current)
        api.log_records = [{'__id': 'l1', 'cr63f_control1': 'C-1', 'field_name': 'Poss Rights Status',
                            'old_value': '-', 'new_value': 'GRANTED'}]
        res = api.revert_log_record('l1')
        self.assertTrue(res.get('success'), res)
        self.assertEqual(api.dv_client.save_loa_record.call_args[0][0], {'cr63f_possrightsstatus': None})


class FrontendTests(unittest.TestCase):
    def test_columns_form_and_aliases_exist(self):
        js, html = read('web', 'app.js'), read('web', 'index.html')
        for logical, display, _ in FIELDS:
            col = f"key: '{logical}', alt: '{display}', label: '{display}'"
            self.assertEqual(js.count(col), 2, f'{display}: grid and logs column')
            self.assertIn(f"'{logical}': ['{logical}'", js)
        for dom_id in ('fPossRightsStatus', 'fPossRightsArea', 'fPossRightsCat', 'fCatArea'):
            self.assertIn(f'id="{dom_id}"', html)
            self.assertIn(f"{dom_id}: document.getElementById('{dom_id}')", js)
            self.assertRegex(js, rf"dom\.{dom_id}\.value = ''")          # cleared for a new entry
            self.assertRegex(js, rf"dom\.{dom_id}\) dom\.{dom_id}\.value = r\[")  # filled when editing
        for logical, display, _ in FIELDS:
            self.assertIn(f"'{logical}': dom.", js)                      # sent on save
            self.assertIn(f"'{display}': dom.", js)


if __name__ == '__main__':
    unittest.main()
