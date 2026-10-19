"""Every audited field (LOG_COMPARE_SPECS) resolves to the same Land Master column in web/app.js."""
import json
import os
import shutil
import subprocess
import unittest

from app import LOG_COMPARE_SPECS

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

NODE_SCRIPT = r"""
const fs = require('fs');
const src = fs.readFileSync(process.argv[1], 'utf8').replace(/\r\n/g, '\n');
const grab = (start, end) => { const i = src.indexOf(start); return src.slice(i, src.indexOf(end, i) + end.length); };
const code = 'const TYPE_OPTIONS=[],ENTITY_OPTIONS=[],CONTRACT_TYPES=[],MODES=[],ACQUISITION_SUBLEVEL_OPTIONS=[];'
  + grab('const DEFAULT_COLUMNS = [', '];')
  + grab('const CANONICAL_FIELD_ALIASES = {', '};')
  + grab('function getFieldNameAttributeKey(fieldName) {', '\n  }\n')
  + 'return {getFieldNameAttributeKey, DEFAULT_COLUMNS};';
const lib = new Function(code)();
const specs = JSON.parse(process.argv[2]);
const out = specs.map(([canon, logical]) => [canon, lib.getFieldNameAttributeKey(canon),
  lib.DEFAULT_COLUMNS.some(c => c.key === logical)]);
process.stdout.write(JSON.stringify(out));
"""


@unittest.skipUnless(shutil.which('node'), 'node is not installed')
class AuditFieldParity(unittest.TestCase):
    def test_python_audit_fields_match_js_columns(self):
        specs = [[canon, keys[0]] for canon, keys, _ in LOG_COMPARE_SPECS]
        res = subprocess.run(['node', '-e', NODE_SCRIPT, os.path.join(ROOT, 'web', 'app.js'), json.dumps(specs)],
                             capture_output=True, text=True, check=True)
        for (canon, logical), (_, js_key, has_col) in zip(specs, json.loads(res.stdout)):
            self.assertEqual(js_key, logical, f'{canon}: JS resolves to {js_key}')
            self.assertTrue(has_col, f'{canon}: no Land Master column {logical}')


if __name__ == '__main__':
    unittest.main()
