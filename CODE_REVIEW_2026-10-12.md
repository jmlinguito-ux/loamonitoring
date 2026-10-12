# LOA Monitoring Tool: Pre-launch Review (2026-10-12)

Scope: the current code in `C:\Users\Jinguito\Desktop\CSV\LOA`: `app.py`, `loa_engine.py`, `dataverse.py`, `web/app.js`, `web/index.html`, the build spec and the last build output. The 2026-10-08 fixes were taken as the starting point.

What was run:
- Unit suite: **188 tests pass** (local time and `TZ=UTC`). `node --check web/app.js` and `tests/test_revert_reference_ui.js` pass.
- Work Week: JS `getWorkWeek` and Python `get_work_week` agree on every day from 2015 to 2034.
- Each issue below marked "Confirmed" was reproduced with a small script. The others were found by reading the code.

No Dataverse calls were made, and the code was not changed.

## Status after "apply all" (2026-10-12)

A full end-to-end test of every feature followed and found and fixed six more issues: see [`TEST_REPORT_2026-10-12.md`](TEST_REPORT_2026-10-12.md).

Everything below was applied. Tests: **227 Python tests pass** (188 earlier + 39 new in `tests/test_launch_review_2026_10_12.py`, local time and `TZ=UTC`), plus `node tests/test_launch_review_ui.js` and `node tests/test_revert_reference_ui.js`. A headless-browser run of the real page with a stubbed Python side showed no script errors, and confirmed the date, Mode, zero-amount, version and DELETE-badge behaviour. Not tested here: real Dataverse calls and a Windows build (`build.ps1`).

| # | What changed |
|---|---|
| A1 | `read_tabular_file` keeps each value under its own header; repeated headers become `Name (2)`; workbooks are closed after reading. |
| A2 | Typing with separators is no longer reshuffled. `normalizeDateInput` accepts M/D/YYYY, M/D/YY, YYYY-MM-DD, MMDDYYYY and Excel serials, and returns MM/DD/YYYY. Invalid dates are refused with a message (grid, drawer, Finance and Conversion grids), and pasted invalid dates are skipped and listed. |
| A3 | `save_record(strict=True)` raises on unreadable numbers and dates (Land Master and reference tables). `-`/`N/A` mean blank and `(1,234)` is −1234. The logs table stays lenient. |
| A4 | Grid saves refuse blank Type/Mode (same rule as the drawer, `_required_field_cleared`), and the grid also stops it before staging. Unlisted stored values show as "(not in list)". |
| A5 | Drawer, grid, Update from References and Revert save with `If-Match`. A 412 means "changed in Dataverse by someone else". The new version is returned by the PATCH and kept. Bulk saves use the app's current copy of each account. |
| A6 | Delete writes a `DELETE` audit row with the full snapshot first: no audit row means no delete, and a failed delete removes the audit row. DELETE rows show in Entry Logs (filter "Deleted"), cannot be reverted, and the confirmation says so. |
| A7 | 429/503 are retried with `Retry-After` (up to 6 attempts) for every request. Update from References now finishes each account on its own (baseline → master → audit row, with rollback) using 6 workers. Progress shows from "Checking audit history...". A partial failure reports how many accounts saved, and saving again retries only the failed ones. |
| B1 | Sync to Dataverse matches rows by key and updates them, sends only changed values, creates only new keys, skips and lists ambiguous keys, writes one `Bulk Sync` audit row, reports failures and skips on screen, and is blocked while edits are unsaved. |
| B2 | Reference row delete: an audit row is written first, and the row leaves the screen only after Dataverse confirms the delete. A file-only row is removed locally. |
| B3 | Created By names are read from the lookup's display value. Rows saved in this session show the signed-in user. |
| B4 | A typed 0 stays 0 in Land Master, Finance and Conversion amount cells. |
| B5 | The fake update check is removed. Settings > About shows `APP_VERSION` (now 1.0.7) from `app.py`. |
| B6 | Overview indexes accounts once per call: the 2,800-account test week takes 0.7 s (was 11.5 s). |
| B7 | `_dialog_path` handles both pywebview return types in all eight file dialogs, and `requirements.txt` pins `pywebview>=6.0`. |
| B8 | `build.ps1` builds from a clean `.build-venv` and runs the tests first. The spec excludes unrelated packages and the non-Windows pywebview back-ends. |
| C | Removed `update_record_cell`. README no longer claims a recent-files list. Blank % Paid shows blank. The Land Master KPI Balance uses Σ max(0, TCP − Deductible). A second Sign-in click reuses the waiting code. Text starting with `=` stays text in all Excel exports. The per-account history check filters only on Control columns that exist (closes open decision 3). |

Still yours to decide (not changed): the one-time % Paid recompute (now safe to run: throttling, per-account save and progress are in place) and the 8 LAND - SPECIAL accounts with only an SA row.

---

## A. Fix before launch (data can be lost or wrong)

| # | Issue | Where | Status |
|---|---|---|---|
| A1 | **An empty header column shifts every column to its right.** `read_tabular_file` drops blank headers, then still reads values by position. In a Finance file laid out as `AREA-INDEX, (blank), TCP`, the TCP column gets the blank column's value. Update from References would then write wrong amounts. | `loa_engine.py` 838–848 (xlsx), 860–869 (csv) | Confirmed |
| A2 | **Dates typed or pasted without leading zeros are scrambled, then erased.** `maskDateString` keeps the digits only, so `1/5/2026` becomes `15/20/26` and `2026-10-12` becomes `20/26/1012`. On save, `_clean_dataverse_datetime` rejects the value and the date is stored as **blank** without any error. Excel copies dates as `1/5/2026`, so a multi-cell paste into Endorsed, Secured or Secured Endorse Date erases those dates. | `app.js` 2153 (mask), 893 (paste), `dataverse.py` 927–932 | Confirmed |
| A3 | **Values that cannot be parsed are saved as blank.** `save_record` turns a Money, Decimal or Integer value it cannot parse (for example `(1,234.00)`, `1.2M` or `₱ 1 000`) into `null`, and it does the same for a bad date. It should raise an error, as it already does for a choice label that matches nothing. | `dataverse.py` 891–904, 927–932 | Confirmed |
| A4 | **The grid can blank Type and Mode, but the drawer cannot.** `save_staged_records` has no Type/Mode guard: choosing "-- Select --" or pasting a blank cell saves `null`. The grid also shows a stored value that is not in the list as "-- Select --". The drawer handles both cases correctly (`ensureSelectValue`). | `app.py` 3225 (missing guard; compare 3066), `app.js` 3740–3752 | Confirmed |
| A5 | **No protection against two people editing the same account.** Drawer, grid and reference saves PATCH without `If-Match` (only Revert uses one), and they build the audit snapshot from the copy loaded when the app started. If user B changed Mode after user A loaded the data, A's next edit records the *old* Mode in its snapshot. Entry Logs and the Overview then show a "Mode changed back" edit that never happened. Before writing, re-read the record (`get_record` gives an etag), build the snapshot from it, and PATCH with that etag. | `app.py` 3057–3089, 3253–3319, 2358 | Code reading. **High if more than one person will use it; low if only you.** |
| A6 | **Deleting a Land Master account writes no audit row.** Every grid row has a delete button. Deleting writes no log row, cannot be undone, and contradicts the README rule that every Land Master change is audited. Either remove the button or write a `DELETE` snapshot first and refuse the delete if that write fails. | `app.js` 3719, `app.py` 4263 | Code reading |
| A7 | **A large Update from References is not safe yet.** This matters for the pending one-time recompute (Decision 1, about 2,800 accounts). (a) The run makes no 429/`Retry-After` handling. It needs a baseline, a master PATCH and a log POST per account (≈8,000 requests), but Dataverse allows about 6,000 requests per user per 5 minutes. (b) Baselines are written one at a time before the progress bar starts, so the window looks frozen for many minutes. (c) All master records are saved before any audit row. If the network drops in phase 2, rollbacks also fail and accounts are left changed with no audit row. Retry on 429, show baseline progress, and save each account's master and log together. | `app.py` 2343–2413, `dataverse.py` 352–406 | Code reading |

## B. Should fix (wrong numbers or misleading screens)

| # | Issue | Where |
|---|---|---|
| B1 | **Sync to Dataverse (Finance and Conversion) duplicates rows.** Rows loaded from a file have no ID, so each sync creates them again. Reloading the same week's file and syncing doubles the table. Duplicates whose amounts differ then cause accounts to be skipped as "conflicting duplicates". The path writes no reference audit rows, and the UI says "Completed!" even when `error_count > 0`. | `app.py` 1696–1738, 2040–2082; `app.js` 5897–5906 |
| B2 | **Deleting a Finance or Conversion row reports success even when the Dataverse delete failed.** The row disappears locally and comes back on the next sync. | `app.py` 1543–1549, 1890–1896 |
| B3 | **"By" is always "Unknown"** in Changed accounts and the weekly report. The Web API returns Created By as `_createdby_value@…FormattedValue`, but `fetch_records` looks for `createdby`, so the name is never read. | `dataverse.py` 812–820, `app.py` 3901 |
| B4 | **A typed 0 becomes blank in grid amount cells.** The code uses `num ? … : ''`. A TCP of 0 is saved as blank, which also clears Balance and % Paid. | `app.js` 3815–3820 |
| B5 | **"Check for updates" never checks.** It always says "latest version (v1.0.6)". Remove it or make it check a real version source. | `app.js` 8919–8922 |
| B6 | **The Overview is slow for a busy week.** One reference run on 2,800 accounts took **11.5 s** to build its week in a test here (each log is matched against every account three times). Build a key → account map once per call. | `app.py` 3763–3783, 3893 |
| B7 | **Export dialogs depend on the pywebview version.** pywebview 4.x–5.x returns a plain string from a save dialog, so `result[0]` saves to a file named `C.xlsx` in the working folder. The current build uses 6.x and works, but `requirements.txt` allows `>=4.4.1`. Pin `pywebview>=6`, or use the string/tuple check that `export_weekly_report` already has. | `app.py` 1621, 1966, 4195, 4351 (compare 4083) |
| B8 | **The build bundles packages the app does not use.** The last build includes torch, transformers, scipy, sklearn, sympy, botocore and others, and its package file is 427 MB. A one-file exe unpacks all of that to `%TEMP%` on every launch: startup is slow and antivirus is more likely to flag it. Build from a clean virtual environment with only `requirements.txt` (plus `excludes=` in the spec). | `LOA-Monitoring-Tool.spec`, `build/…/PKG` |

## C. Minor and cleanup

- `update_record_cell` writes to Land Master with no audit row. The UI does not call it, but it is still exposed to the page: remove it.
- The README says the app remembers recent files, but nothing in the UI uses `load_*_file_recent`.
- The grid shows a blank % Paid as `0.00%`, while Python shows it as blank.
- The Land Master KPI adds up the *stored* Balance, while the Overview adds up the *calculated* Balance clamped at 0. The two totals can differ.
- If a reference save fails part-way, the retry recalculates from the old baseline and writes duplicate (harmless) snapshots for the accounts that had already saved.
- Clicking Sign in twice starts two device flows. When the first one expires, it can report failure after the second one succeeded.
- Excel exports write text that starts with `=` as a formula.
- Still open from 2026-10-08: the one-time % Paid recompute, the 8 LAND - SPECIAL accounts with only an SA row, and whether `cr63f_control` exists on the logs table. Every drawer and grid save filters on that column, so the save fails if it is missing.

## Suggested order

1. A1, A2, A3, A4: small, local fixes, and they stop silent data loss.
2. A6 and B1/B2: remove or audit destructive actions.
3. Decide how many people will use the tool. If more than one, A5.
4. A7 before running the one-time recompute.
5. B8 (clean build) and B7 (pin pywebview) before handing out the exe.
6. The rest as time allows.
