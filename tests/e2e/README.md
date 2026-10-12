# End-to-end tests (no Dataverse needed)

These run the **real app** (Python `Api` + the web UI in `web/`) in a headless browser
against `fake_dataverse.py`, an in-memory stand-in for the Dataverse Web API. It checks
metadata, paging, `$filter`, version tags (412 on a stale save), and rejects unknown columns,
wrong types and over-long text the way Dataverse does. It can also simulate throttling (429) and failures.

Nothing touches your real Dataverse. The app is copied to a temporary folder and run there.

```powershell
pip install playwright openpyxl msal requests
python -m playwright install chromium
python tests\e2e\e2e_test.py     # ~135 checks: every screen and save path
python tests\e2e\e2e_extra.py    # ~35 checks: grid tools, tooltips, filters, Overview tools
```

Each script prints PASS/FAIL per check and ends with a total. Screenshots of failures and a
`results.json` are written to `%TEMP%\loa_e2e_work`.

Not covered here: Microsoft sign-in (device code) and the packaged .exe; test those by hand.
