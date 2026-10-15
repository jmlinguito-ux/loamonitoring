"""Mode → payment group: exact lists, identical in app.py and web/app.js (2026-10-15)."""

import json
import os
import shutil
import subprocess
import unittest

from app import _payment_group

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

MODES = [
    'SECURED SALE', 'SS - FROM LAND', 'SS - FROM ROW', 'SECURED LEASE', 'SL - FROM LAND',
    'SS - STP', 'SS - FROM LAND STP', 'SL - STP', 'SL - FROM LAND STP', 'SOLD TO PARTNER',
    'SECURED SALE NON SP', 'SECURED LEASE NON SP', 'RESCINDED', 'SS - RESCINDED', 'SL - RESCINDED',
    # spacing / dash / case variants of listed modes
    'ss-stp', 'SL-FROM LAND', 'SECURED SALE-NON SP', 'SECURED LEASE NON-SP', 'SECURED SALE NONSP',
    '  secured   sale ', 'SS – FROM ROW',
    # not listed: never a secured sale or lease
    'OPEN TO SALE', 'OPEN TO LEASE', 'FOR VERIFICATION', 'NOT AVAILABLE', 'TO BE IDENTIFIED',
    'NOT USABLE', 'SALE', 'LEASE', 'SUBLEASE', 'STP', '',
]

EXPECTED = {
    'SECURED SALE': 'DIRECT_SALE', 'SS - FROM LAND': 'DIRECT_SALE', 'SS - FROM ROW': 'DIRECT_SALE',
    'SECURED LEASE': 'LEASE_DIRECT', 'SL - FROM LAND': 'LEASE_DIRECT',
    'SS - STP': 'STP_SALE', 'SS - FROM LAND STP': 'STP_SALE', 'SOLD TO PARTNER': 'STP_SALE',
    'SL - STP': 'LEASE_STP', 'SL - FROM LAND STP': 'LEASE_STP',
    'SECURED SALE NON SP': 'NON_SP', 'SECURED LEASE NON SP': 'NON_SP',
    'RESCINDED': 'RESCINDED', 'SS - RESCINDED': 'RESCINDED', 'SL - RESCINDED': 'RESCINDED',
    'ss-stp': 'STP_SALE', 'SL-FROM LAND': 'LEASE_DIRECT', 'SECURED SALE-NON SP': 'NON_SP',
    'SECURED LEASE NON-SP': 'NON_SP', 'SECURED SALE NONSP': 'NON_SP', '  secured   sale ': 'DIRECT_SALE',
    'SS – FROM ROW': 'DIRECT_SALE',
}


def _js_groups(modes):
    with open(os.path.join(ROOT, 'web', 'app.js'), encoding='utf8') as fh:
        src = fh.read()
    start = src.index('  const PAYMENT_GROUP_OF_MODE = {')
    end = src.index('  function recalculateKPIs() {')
    script = src[start:end] + '\nconsole.log(JSON.stringify(%s.map(paymentGroupOf)));' % json.dumps(modes)
    out = subprocess.check_output(['node', '-e', script], cwd=ROOT)
    return json.loads(out.decode('utf8'))


class TestModeGroups(unittest.TestCase):
    def test_python_groups(self):
        for mode in MODES:
            self.assertEqual(_payment_group(mode), EXPECTED.get(mode, 'OTHER'), repr(mode))

    @unittest.skipUnless(shutil.which('node'), 'node is not installed')
    def test_javascript_matches_python(self):
        js = _js_groups(MODES)
        for mode, group in zip(MODES, js):
            self.assertEqual(group, _payment_group(mode), repr(mode))


if __name__ == '__main__':
    unittest.main()


class TestOpenModesNotSecured(unittest.TestCase):
    """OPEN TO SALE / OPEN TO LEASE are not secured: not in % Paid, status counts or balance."""

    def test_open_modes_stay_out_of_paid_figures(self):
        from unittest.mock import patch
        from app import Api
        with patch('dataverse.DataverseClient.try_silent_token', return_value=None):
            app = Api()
        app.loa_records = [
            {'new_mode': 'SECURED SALE', 'cr63f_totalcontractprice': '100', 'cr63f_totalpaymentdeductibletotcp': '40'},
            {'new_mode': 'SECURED LEASE', 'cr63f_totalcontractprice': '50', 'cr63f_totalpaymentdeductibletotcp': '10'},
            {'new_mode': 'OPEN TO SALE', 'cr63f_totalcontractprice': '1000', 'cr63f_totalpaymentdeductibletotcp': '0'},
            {'new_mode': 'OPEN TO LEASE', 'cr63f_totalcontractprice': '500', 'cr63f_totalpaymentdeductibletotcp': '0'},
        ]
        fb = app.get_dashboard_analytics()['financial_breakdown']
        self.assertEqual(fb['direct_sale_count'], 1)
        self.assertEqual(fb['lease_direct_count'], 1)
        self.assertEqual(fb['other_count'], 2)
        self.assertEqual(fb['balance_to_pay'], 100.0)   # 60 secured sale + 40 secured lease
        self.assertEqual(fb['zero_paid_count'], 0)       # OPEN TO SALE is not "no payment yet"
