"""Work week: web/app.js getWorkWeek and loa_engine.get_work_week agree (2026-10-15)."""

import json
import os
import shutil
import subprocess
import unittest

from loa_engine import get_work_week

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DATES = ['10/14/2026', '10.14.2026', '2026-10-14', '2026.10.14', '10/14/26', '14/10/2026', '1/5/2026',
         '01-05-2026', '2026/1/5', '10/14/2026 12:00 AM', '2026-10-14T08:00:00Z', '12/31/2025',
         '12/30/2025', '1/1/2027', '2/29/2028', '13/13/2026', '2/30/2026', '10/14/1900', '']


def _js_work_weeks(dates):
    with open(os.path.join(ROOT, 'web', 'app.js'), encoding='utf8') as fh:
        src = fh.read()
    start = src.index('  function getWorkWeek(dateVal) {')
    depth, i = 0, src.index('{', start)
    while True:
        depth += {'{': 1, '}': -1}.get(src[i], 0)
        if depth == 0:
            break
        i += 1
    script = src[start:i + 1] + '\nconsole.log(JSON.stringify(%s.map(getWorkWeek)));' % json.dumps(dates)
    return json.loads(subprocess.check_output(['node', '-e', script], cwd=ROOT).decode('utf8'))


class TestWorkWeekParity(unittest.TestCase):
    @unittest.skipUnless(shutil.which('node'), 'node is not installed')
    def test_javascript_matches_python(self):
        for date, js in zip(DATES, _js_work_weeks(DATES)):
            self.assertEqual(js, get_work_week(date), repr(date))

    def test_impossible_dates_have_no_week(self):
        self.assertEqual(get_work_week('13/13/2026'), '')
        self.assertEqual(get_work_week('2/30/2026'), '')

    def test_month_first_dashes(self):
        self.assertEqual(get_work_week('01-05-2026'), get_work_week('1/5/2026'))
