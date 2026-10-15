"""Regression tests for 2026-10-13: Land Master-only conversion card and persistent connection settings."""

import json
import os
import tempfile
import unittest
from unittest.mock import patch

import dataverse as dv
from app import Api


class TestConversionCardUsesLandMasterOnly(unittest.TestCase):
    def setUp(self):
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            self.app = Api()

    def test_empty_land_master_ignores_conversion_reference(self):
        self.app.loa_records = []
        self.app.conversion_rows = [
            {'DAR STATUS': 'DAR SECURED', 'AREA APPLIED (HAS.)': '5'},
            {'DAR STATUS': 'DAR APPLIED', 'AREA APPLIED (HAS.)': '3'},
        ]
        metrics = self.app.get_dashboard_analytics()['conversion_metrics']
        self.assertEqual(metrics['total_lots'], 0)
        self.assertEqual(metrics['secured_count'], 0)
        self.assertEqual(metrics['breakdown'], [])
        self.assertEqual(metrics['total_applied_area'], 0)


class TestConnectionSettingsPersist(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.path = os.path.join(self.tmp, 'dataverse_config.json')
        self.patcher = patch('dataverse.config_path', return_value=self.path)
        self.patcher.start()
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            self.app = Api()
        self.app.settings = {}

    def tearDown(self):
        self.patcher.stop()

    def _write(self, cfg):
        with open(self.path, 'w', encoding='utf-8') as f:
            json.dump(cfg, f)

    def test_theme_only_save_keeps_connection(self):
        self._write(dict(dv.DEFAULT_CONFIG))
        with patch('app.save_app_settings'), patch.object(Api, '_init_dataverse'):
            self.assertTrue(self.app.save_settings({'theme': 'dark'})['success'])
        cfg = dv.load_config()
        for key, value in dv.DEFAULT_CONFIG.items():
            self.assertEqual(cfg[key], value)

    def test_old_default_solution_is_replaced(self):
        self._write({**dv.DEFAULT_CONFIG, 'solution_id': 'a7f41cec-5786-f111-ab0f-7ced8d3b4de7',
                     'solution_name': 'SOURCING AND NEGO (BATANGASNEGO)'})
        cfg = dv.load_config()
        self.assertEqual(cfg['solution_id'], dv.DEFAULT_CONFIG['solution_id'])
        self.assertEqual(cfg['solution_name'], 'LOA MONITORING (LOAMONITORING)')

    def test_missing_values_filled_from_defaults(self):
        self._write({'environment_url': 'https://yourorg.crm.dynamics.com/'})
        cfg = dv.load_config()
        self.assertEqual(cfg['environment_url'], 'https://yourorg.crm.dynamics.com')
        self.assertEqual(cfg['tenant_id'], dv.DEFAULT_CONFIG['tenant_id'])
        self.assertEqual(cfg['loa_table'], 'cr63f_loamonitoring')


if __name__ == '__main__':
    unittest.main()
