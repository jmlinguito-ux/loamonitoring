"""Each saved Land Master row gets its audit row at once; closing waits for a save."""
import os
import sys
from concurrent.futures import Future
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
from app import Api  # noqa: E402


def make_api(order):
    api = Api()
    api.dv_client = MagicMock()
    api.dv_client.config = {'loa_table': 'cr63f_loamonitoring', 'logs_table': 'cr63f_loamonitoringlogs'}

    def save_master(*a, **k):
        order.append(('master', k.get('record_id') or a[1]))
        return {'action': 'update', 'id': 'x', 'etag': 'etag_new'}

    def save_log(entry):
        order.append(('log', entry.get('field_name')))
        return {'id': 'log_' + str(len(order))}

    api.dv_client.save_loa_record.side_effect = save_master
    api.dv_client.save_log_record.side_effect = save_log
    api.loa_records = [
        {'__id': f'rec_{i}', '__etag': 'etag_old', 'cr63f_control1': f'CTRL-{i}',
         'cr63f_areaindex': f'AI-{i}', 'cr63f_conversionstatus': 'OLD'}
        for i in range(3)
    ]
    api.log_records = []
    return api


class ImmediateAuditTests(unittest.TestCase):
    def test_log_row_written_right_after_its_master_row(self):
        order = []
        api = make_api(order)
        staged = [{'id': f'rec_{i}', 'fields': {'cr63f_conversionstatus': 'DAR FINALITY SECURED'}} for i in range(3)]
        with patch.object(api, '_write_baseline_if_missing', return_value=None), \
                patch.object(app_module, 'ThreadPoolExecutor', _SerialPool):
            res = api.save_staged_records(staged)
        self.assertTrue(res['success'], res)
        kinds = [k for k, _ in order]
        # master, log, master, log, master, log: never all masters first.
        self.assertEqual(kinds, ['master', 'log'] * 3)

    def test_audit_failure_rolls_back_that_row_and_reports(self):
        order = []
        api = make_api(order)
        api.dv_client.save_log_record.side_effect = RuntimeError('boom')
        with patch.object(api, '_write_baseline_if_missing', return_value=None):
            res = api.save_staged_records(
                [{'id': 'rec_0', 'fields': {'cr63f_conversionstatus': 'NEW'}}])
        self.assertFalse(res['success'])
        self.assertIn('rolled back', res['error'])
        # Master saved once, then restored to the old value.
        calls = api.dv_client.save_loa_record.call_args_list
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[1][0][0]['cr63f_conversionstatus'], 'OLD')
        self.assertEqual(api.loa_records[0]['cr63f_conversionstatus'], 'OLD')


class _SerialPool:
    """Runs each submitted job immediately, one after another."""
    def __init__(self, *a, **k):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def submit(self, fn, *args):
        fut = Future()
        try:
            fut.set_result(fn(*args))
        except Exception as exc:  # pragma: no cover
            fut.set_exception(exc)
        return fut


class CloseGuardTests(unittest.TestCase):
    def test_counter_and_deferred_close(self):
        api = Api()
        api._window = MagicMock()
        app_module._save_begin(api)
        app_module._save_begin(api)
        api._close_after_save = True
        app_module._save_end(api)
        api._window.destroy.assert_not_called()      # one save still running
        app_module._save_end(api)
        api._window.destroy.assert_called_once()     # last save done: close now
        self.assertEqual(api._active_saves, 0)

    def test_guard_decorator_counts_and_releases_on_error(self):
        seen = []

        class Probe(Api):
            @app_module._guard_save
            def work(self, fail=False):
                seen.append(self._active_saves)
                if fail:
                    raise ValueError('x')
                return 'ok'

        p = Probe()
        self.assertEqual(p.work(), 'ok')
        with self.assertRaises(ValueError):
            p.work(fail=True)
        self.assertEqual(seen, [1, 1])
        self.assertEqual(p._active_saves, 0)


if __name__ == '__main__':
    unittest.main()
