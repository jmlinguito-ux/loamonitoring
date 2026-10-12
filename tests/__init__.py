import os
import tempfile
import atexit

_test_temp_cache = tempfile.NamedTemporaryFile(delete=False, suffix='_test_entry_logs.json')
_test_temp_cache.write(b'[]')
_test_temp_cache.close()

_test_temp_tombstone = tempfile.NamedTemporaryFile(delete=False, suffix='_test_tombstones.json')
_test_temp_tombstone.write(b'[]')
_test_temp_tombstone.close()

os.environ['LOA_LOGS_CACHE_TEST_PATH'] = _test_temp_cache.name
os.environ['LOA_TOMBSTONES_TEST_PATH'] = _test_temp_tombstone.name

def _cleanup():
    for f in (_test_temp_cache.name, _test_temp_tombstone.name):
        try:
            if os.path.exists(f):
                os.remove(f)
        except Exception:
            pass

atexit.register(_cleanup)

