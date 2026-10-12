# LOA Monitoring Tool: Full Feature Test (2026-10-12)

## How it was tested

The **real app** was run end to end: the real Python `Api` and the real `web/` UI in a headless Chromium browser. It ran against an in-memory stand-in for Dataverse (`tests/e2e/fake_dataverse.py`), so your live data was never touched.

The stand-in behaves like Dataverse where the app depends on it:
- FormattedValues (en-US, dates without leading zeros);
- paging with `nextLink`;
- `$filter`;
- version tags, with a 412 on a stale save;
- `return=representation`;
- 400s for unknown columns, wrong types and text longer than the column allows.

Throttling (429) and failures (audit write, delete, update) were injected on purpose. Seed data covers:
- LAND, ROW, LAND - SPECIAL;
- sale and lease;
- a legacy fraction-scale % Paid;
- a 1900 placeholder date;
- a Mode not in the list;
- a shared AREA-INDEX;
- accounts with and without audit history.

The reference files include a blank-header column, a conflicting duplicate key, an SA-only row and a ROW corridor row.

| Suite | Result |
|---|---|
| `tests/e2e/e2e_test.py` (every screen and save path) | **133 passed, 0 failed** (172 Dataverse requests); 150 after round 3 |
| `tests/e2e/e2e_extra.py` (grid tools, tooltips, filters, Overview tools) | **34 passed, 0 failed** |
| Python unit tests (`python -m unittest discover -s tests`) | **230 passed** (local time and UTC) |
| `node tests/test_launch_review_ui.js`, `node tests/test_revert_reference_ui.js` | passed |

Across all runs: no JavaScript errors, no Python exceptions, and Dataverse never rejected a payload as invalid.

## What was covered

| Area | Checked |
|---|---|
| Start-up | Paged load of all accounts; Balance KPI; connection status; Work Week from Endorsed Date (and none for 1900); unlisted Mode shown "(not in list)"; calculated % Paid shown over a legacy stored value |
| Land Master grid | Search, Type and Mode filters, clear filters, header filters (combined), column picker, sort, freeze, resize, copy a 2×2 block, Show Changed Only, hover tooltip (old → new), tooltip undo, Enter/Escape |
| Grid edits | Text, money (0 stays 0), dates (2/3/2026 → 02/03/2026, Work Week follows), invalid date refused, Mode cannot be cleared, unsaved counter, save → Dataverse, BASELINE + MANUAL_EDIT snapshots, discard, Excel paste (valid date converted, invalid skipped and reported) |
| Two users | A colleague's newer save is detected (refused, nothing written, no audit row); after reload the edit saves and the snapshot contains the colleague's value |
| New Entry drawer | Endorsed Date pre-filled; live Balance/% Paid; live finance lookup by AREA-INDEX; create (amounts, dates, NEW_ENTRY row); duplicate Control refused; invalid date refused (on leaving the field and on Save); shared AREA-INDEX confirmation; Acquisition Sub-Level required; suggestion fills the form without the Control Number; cancel |
| Delete account | Confirmation text; DELETE audit row with last values; Dataverse failure → kept, no audit row left; audit failure → nothing deleted |
| Update from References | Missing-reference message; conflicting duplicate reported, not guessed; staging and highlighting; discard writes nothing; save with 3 throttled requests; SALE vs LEASE rows; ROW by Corridor Index; SA-only untouched; Endorsed Date filled when blank; Conversion Status; one snapshot per account; re-run finds 0 changes |
| Entry Logs | Load; Deleted and Reference filters; search; BASELINE hidden; year-qualified week filter; column picker; header filter; sort; KPI; refresh; Excel export |
| Revert | Text field; TCP (Balance and % Paid recalculated); reverting twice refused; DELETE rows refused |
| Overview | Changed accounts this week; "By" shows names; Needs attention; week stepper; expand diff; mode filter + clear; search; source filter; Open in Land Master; weekly report (3 sheets, names in "By"); Sync timing; Sync keeps a loaded finance file; follows the newest week |
| Finance | Load Excel (blank header column correct); edit while a search is active hits the right row; discard; template; TYPE filter; column picker; Sync to Dataverse (creates, skips the duplicated key, summary audit row, no duplicates on re-sync); sync blocked while edits are unsaved; edit after sync saved and audited; delete (failure kept, success audited); export; reload from Dataverse; clear |
| Conversion | Load CSV; template; sync (date and decimal types right, no duplicates on re-sync); edit saved and audited; search; status filter; delete; export; reload; clear |
| Settings | About version; solutions and table lists; save config (all four tables kept); theme; density; refresh solutions; connection status |

## Bugs found by this run (all fixed and re-tested)

| # | Bug | Impact | Fix |
|---|---|---|---|
| E1 | **Reloading Land Master replaced a loaded Finance/Conversion file** with the Dataverse table, and wiped unsaved reference edits. This ran at start-up, on Sync Land Master, after every New Entry save, and on Overview Sync. | After loading this week's finance file and saving one entry, **Update from References silently used the older Dataverse data**. | A file loaded in this session, or unsaved reference edits, is kept on reload. The explicit Sync Finance/Conversion buttons still load Dataverse. At start-up Dataverse still wins over an old remembered file. |
| E2 | **Reference edits made after "Sync to Dataverse" were not saved.** The screen said "Saved 1 record(s)" but wrote nothing. | Lost Finance/Conversion edits, with no audit row. | The "before" values are keyed by the id the screen sends (and also looked up under the row's other key). Both screens reload their rows with the new IDs after a sync. |
| E3 | **Editing or deleting a file-loaded reference row while a search/filter/sort was active changed a different row** (filtered position used as the full-list position). | Wrong Finance/Conversion row edited or deleted. | Rows without an ID are addressed by their position in the full list. |
| E4 | **The Overview stayed on the week shown at start-up**, so this week's activity was hidden after new work. | Changed accounts looked missing. | With the newest week selected, the Overview follows the newest week. An older week you picked stays. |
| E5 | **The new "Deleted" Entry Logs filter showed nothing** (a gap in the previous round), and Conversion edits were filed under "Manual Edit". | Filter unusable or misleading. | Deleted and Conversion Update options added; an empty filter value means All. |
| E6 | **Picking an existing AREA-INDEX suggestion in New Entry copied its Control Number**, so the save was always refused as a duplicate. | Suggestion feature unusable. | The suggestion fills the form as a template but leaves Control Number blank (with a prompt). |

## Decisions applied after this report (round 3)

| Decision | What changed | Tested |
|---|---|---|
| SA = Special Acquisition | A LAND - SPECIAL account with no SALE/LEASE match uses its SA row; SALE/LEASE still win; LAND/ROW never use SA; CWA never | Unit tests (incl. TAA-1 ₱14,000,000 case, SALE-wins, conflicting SA); E2E: C-004 now takes its SA amounts |
| Recalculate Balance & % Paid | Settings > Maintenance: preview count, then audited, version-checked save of every account with a TCP; repeatable | Unit tests (preview, only Balance/% Paid written, If-Match, RECALCULATE snapshot, no-TCP skipped, no Entry Log noise); E2E: legacy 0.1% → 10.0%, stale Balance fixed, second run finds nothing |
| Edit account in the drawer | Pencil button / double-click row number; only changed fields saved; suggestions off while editing | E2E: open, stored values, "No changes", save, untouched fields unchanged, audit row, new entry afterwards is blank |
| Revert in Entry Logs | Revert button on revertible rows; none on DELETE rows; none in the Overview | E2E: Revert from the log row; DELETE row and Overview have none |
| No delete on reference rows | Conversion delete button removed (Finance already had none) | JS check + E2E |

Final runs after round 3: **E2E 150 + 34 passed, 0 failed**; **235 unit tests** passed; both JS checks passed.

Still open: none of the earlier questions. Before launch, run the Windows check below, then use Settings > Maintenance > Recalculate (after an Excel backup).

## Not tested here

- **Microsoft sign-in** (device code) and sign-out, because they need Microsoft's servers.
- **The packaged .exe** (`build.ps1` runs on Windows).
- **Your real Dataverse schema.** The stand-in uses the columns the app expects, so a column whose real type or length differs (for example a 100-character limit on `cr63f_name`) would only show up live.

Before handing the app out, test once on Windows against Dataverse: sign in, then one grid edit, one new entry, one delete and one Update from References.

## Re-running

```powershell
pip install playwright openpyxl msal requests
python -m playwright install chromium
python tests\e2e\e2e_test.py
python tests\e2e\e2e_extra.py
```
