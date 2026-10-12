# LOA Monitoring Tool — Project Memory & Technical Knowledge Base

This document persists all system architecture, business rules, calculations, data models, and feature behaviors established across all conversations and iterations.

---

## 1. System Architecture & Database Partitioning

### Dataverse Tables
1. **Land Master DB (`cr63f_loamonitoring`)**:
   - Master repository of all land acquisition records and current baseline state.
   - Fields: `cr63f_control1` (Control Number), `cr63f_indexno` (Index No / Corridor Index for ROW), `cr63f_type` (Type), `cr63f_areaindex` (Area-Index), `cr63f_titleno` (Title No), `cr63f_registeredowner` (Registered Owner), `cr63f_authorizedrepresentative` (Authorized Rep), `cr63f_entity` (Entity), `cr63f_province` (Province), `cr63f_municipality` (Municipality), `cr63f_dbadjustedbarangay` (DB Adjusted Barangay), `cr63f_lotareaha` (Lot Area Ha), `new_contracttype` (Contract Type), `new_mode` (Mode), `cr63f_conversionstatus` (Conversion Status), `cr63f_rccterms` (RCC Terms), `cr63f_rccprice` (RCC Price), `cr63f_totalcontractprice` (Total Contract Price), `cr63f_totalpaymentdeductibletotcp` (Total Payment Deductible to TCP), `cr63f_balance` (Balance), `cr63f_paid` (% Paid, stored on a 0–100 scale: 25.00 = 25%), `cr63f_secureddate` (Secured Date), `cr63f_securedendorsedate` (Secured Endorse Date), `cr63f_endorseddate` (Endorsed Date).

2. **Entry Logs Repository (`cr63f_loamonitoringlogs`)**:
   - Audit trail and historical change ledger.
   - Action types: `NEW_ENTRY`, `MANUAL_EDIT`, `REFERENCE_UPDATE` (legacy `FINANCE_UPDATE`), `REVERT`, and hidden `BASELINE` (a pre-change snapshot written before an account's first audited change; never shown).
   - Each save writes one full-record snapshot row; field-level diffs (`Old Value ➔ New Value`) are derived in memory by comparing consecutive snapshots of the same account (Control Number, else Area-Index, else Title No).
   - Snapshots are ordered by real time: the raw ISO `createdon` (kept as `__createdon_raw` by `dataverse.fetch_records`) or the parsed display time. Never sort log times as text.
   - Audit `Work Week` is derived from the log time (not Endorsed Date).
   - Stores full master record snapshot columns (`cr63f_control1`, `cr63f_registeredowner`, `cr63f_authorizedrepresentative`, `cr63f_entity`, `cr63f_province`, `cr63f_municipality`, `cr63f_lotareaha`, `new_contracttype`, `new_mode`, `cr63f_conversionstatus`, `cr63f_totalcontractprice`, `cr63f_totalpaymentdeductibletotcp`, `cr63f_balance`, `cr63f_paid`, etc.) alongside the audit diff metadata so logs do not have empty/skipped columns.
   - **Storage**: Dataverse is the only source of truth. There is no local log cache (`load_cached_logs`/`save_cached_logs` are no-op stubs). Newly written rows are kept in memory (`pending_session_logs`) until a Dataverse fetch returns them.

3. **Finance Reference (`cr63f_financereference`)**:
   - Appended finance dataset synced from Dataverse or loaded via Excel files for weekly reconciliation and autofill matching.

4. **Land Conversion Reference (`cr63f_landconversionreference`)**:
   - Master reference dataset containing DAR, NIA, and land conversion classifications.
   - Fields: `cr63f_areaindex` (AREA-INDEX), `cr63f_titlenumber` (TITLE NUMBER), `cr63f_darstatus` (DAR STATUS), `cr63f_niastatus` (NIA STATUS), `cr63f_conversionclass` (CONVERSION CLASS), `cr63f_lotorsurveynumber`.
   - Used for live autofill in New Entry drawer and bulk reconciliation in Land Master DB.

---

## 2. Core Business Logic & Calculations

### Work Week Derivation (Wednesday-to-Tuesday Cycle)
- Land Master `Work Week` is derived from `Endorsed Date`; audit/Overview weeks are derived from the log time. Same algorithm (Ground Data Processing Tool):
  - Week cycle runs from **Wednesday (day 1)** to **Tuesday (day 7)**.
  - Formatted as `WW-XX` (e.g., `WW-38`).
  - **Year rule (corrected 2026-10-07)**: a week belongs to the year its Saturday falls in, and **WW-01 is the first Wednesday–Tuesday week whose Saturday is in that year**, so every year starts at WW-01 with no gaps (52 or 53 weeks). Example: 2029 WW-01 = Wed Jan 3 – Tue Jan 9, 2029; Wed Dec 27, 2028 – Tue Jan 2, 2029 is 2028 WW-53. (The earlier count started from the Wednesday on/before Jan 1, which made years starting on Sun/Mon/Tue — 2018, 2019, 2023, 2024, 2029, 2030 — begin at WW-02. Dates in those years now show one week number lower than before, e.g. Sep 1, 2024 is WW-35, not WW-36.) The Ground Data Processing Tool needs the same correction to stay in step.
  - Overview and Entry Logs weeks are year-qualified with that work-week year: `2026 · WW-40` (`get_work_week_label`). Week selection is exact; a bare `WW-40` picks the most recent year. The Overview compares with the previous calendar week (even if it had no logs). The `cr63f_workweek` value saved in Dataverse stays `WW-XX`.
  - Land Master `Work Week` always follows Endorsed Date, including right after a save or new entry (never the save week).
  - **Endorsed Date default (2026-10-07)**: a new entry gets today's date when its Endorsed Date is blank (the form pre-fills it). Update from References fills a truly blank Endorsed Date with today's date when it changes a reference value (TCP, Deductible or Conversion Status). An existing date is never overwritten, a 1900 placeholder included (decided 2026-10-08), and a Balance/% Paid-only correction doesn't set it. The fill is audited as an `Endorsed Date` change.
  - Dates are read month-first (M/D/YYYY); D/M is used only when the first number is above 12 (Python and JS behave the same).
- **1900 Date Handling Rule**:
  - Placeholder dates with year `<= 1900` or containing `'1900'` (e.g., `01/01/1900`) return empty string `""` / `-` without assigning an invalid `WW-53` badge.
  - For Work Week display only. A 1900 Endorsed Date is a stored value: no reference update overwrites it (decided 2026-10-08).

### Reference Lookup & Cascade Matching Logic
When updating records from References ("Update from References" button):
1. **Finance Metrics Cascade**:
   - ROW Types (`TYPE == "ROW"`): Matches Land `INDEX NO` against Finance `Corridor Index`.
   - Non-ROW Types (`TYPE != "ROW"`): Matches Land `AREA-INDEX` against Finance `Area Index`.
   - Fallback Matching: Matches Land `TITLE NO` against Finance `Title Mod` (exact after removing the title-type prefix), then cleaned numeric title IDs.
   - The title fallback never crosses between ROW and non-ROW finance rows (both directions) and rejects titles whose number segments conflict (`TCT-999-2019006115` ≠ `TCT-066-2019006115`).
   - **TYPE-aware matching (decision 2026-10-07; SA rule updated 2026-10-12, see below)**: Finance TYPE decides which rows an account may use. ROW accounts → Finance `ROW` rows only (Corridor Index, then Area Index, then title). LAND / LAND - SPECIAL accounts → `SALE` row when Mode is a sale (SECURED SALE, SS - …, SOLD TO PARTNER), `LEASE` row when Mode is a lease (SECURED LEASE, SL - …), either when Mode is neither. `ROW`, `CWA` and `SA` rows (CWA/SA are contract types) never supply LAND accounts. Untyped rows (Excel without TYPE) still match LAND accounts. Live data check: 3,219 accounts match exactly one row; 8 LAND - SPECIAL accounts have only an SA row and get no match.
   - **Duplicates**: a duplicate is the same key **and the same TYPE** with different amounts. When a key (Corridor Index, Area Index, conversion Area Index, or title) matches several eligible rows, they are used only if they agree (same TCP and Deductible / same DAR status). Conflicting duplicates are skipped and reported (`ambiguous_accounts`, shown as a warning toast) — never resolved by taking the first row.
2. **Conversion Status**:
   - Matched on `AREA-INDEX` only (`ConversionLookupEngine.lookup`). There is no Title No fallback.
   - Populates LOA `cr63f_conversionstatus` directly from `cr63f_darstatus` (e.g. `DAR FINALITY SECURED`, `DAR FOR APPLICATION`, `DAR SECURED`).

### Financial Formulas
- `TOTAL CONTRACT PRICE (TCP)`: Pulled from matched Finance TCP.
- `PAYMENT DEDUCTIBLES`: Pulled from matched Finance Total Payment Deductible to TCP.
- `BALANCE = max(0.0, TCP - PAYMENT DEDUCTIBLES)`.
- `% PAID = (PAYMENT DEDUCTIBLES / TCP) * 100.0` (TCP = 0 with deductible > 0 → 100%).
- **% Paid is always saved (decision 2026-10-07)**: every save path (manual edit, new entry, revert, Update from References) writes `cr63f_paid` rounded to **1 decimal**, the Dataverse column's precision (`paid_percent_for_storage`); changes are compared at 0.05. Update from References also rewrites Balance and % Paid whenever the stored value differs from the calculated one, even if TCP/Deductible did not change. The finance reference's own % Paid is never copied.
- **Manual saves derive Balance and % Paid on the server (2026-10-08)**: the drawer save, new entry and grid staged save recalculate `cr63f_balance` and `cr63f_paid` from the TCP and Deductible they are saving (`_derived_payment_fields`), so the client's own values are never stored. A grid save recalculates only when its change touches TCP, Deductible, Balance or % Paid. The loaded record's `__display_percent_paid` is refreshed after each save.
- **Displayed % Paid is calculated** wherever TCP and Deductible are present (grid via `__display_percent_paid`, drawer via `updatePaymentMath`). The stored value is shown only when an amount is missing.
- **Blank TCP (2026-10-08)**: Balance and % Paid cannot be calculated without a TCP, so a save that blanks TCP clears both (they are not left from the old TCP). A blank Deductible counts as 0.
- **Stored % Paid precision (2026-10-08)**: the drawer lookup, grid lookup and grid calculation show one decimal, the precision that is stored.
- **Type and Mode cannot be blanked (2026-10-08)**: a save that sends a blank Type or Mode for an existing account is refused, so a Mode missing from the list can never erase a stored value. The drop-downs contain every stored value (`SS - STP`, `SL - STP`, `SS - FROM LAND STP`, `SL - FROM LAND STP`, `SS - FROM ROW`, `SS - RESCINDED`, `LAND - SPECIAL`, Entity and Contract Type `N/A`); a stored value that is not listed is kept as an extra option.
- **Mode grouping (2026-10-08)**: an account with no Mode is grouped as `BLANK` in the Overview, not as `TO BE IDENTIFIED`.
- **Drawer on open (2026-10-08)**: opening an account shows its stored values. The reference is compared, not applied: if it differs, the badge says so and the stored values are kept (apply with Update from References). Saving sends only the fields changed in the drawer, plus the identity keys. A lookup that finds nothing while typing restores the account's stored amounts (or clears amounts that came from an earlier lookup); amounts typed by hand are never cleared by a lookup.
- **Revert is an anchor (2026-10-08)**: a REVERT row's embedded record is the state the next change is compared with, so a later edit does not show the reverted change again.
- **No-op staged edits (2026-10-08)**: a grid edit equal to the stored value is not saved and writes no audit row. Staged Balance or % Paid edits are recalculated from TCP and Deductible on save.
- **Audit baseline in bulk (2026-10-08)**: Update from References checks the logs once for the whole run (one query) instead of once per account.
- **Unmatched picklist label (2026-10-08)**: a choice label that matches no option, or matches more than one, raises an error instead of being dropped silently.
- Stored % values are never rescaled: 0.5 means 0.50%, not 50%. **Live data (checked 2026-10-07):** about 2,550 accounts still hold fractions (1,528 hold `1` meaning 100%; 1,020 hold 0.1–0.9) and about 410 hold 0–100 values. Update from References rewrites matched accounts to 0–100 (about 2,800 on its first run); a one-time recompute for all accounts is pending the user's decision.

### Pre-launch rules (2026-10-12)
- **Version-checked saves**: drawer, grid, Update from References and Revert PATCH with `If-Match` = the `__etag` the account was loaded at. Dataverse refuses (412, `ConcurrencyError`) if someone else saved it since; the user reloads. The new version comes back from the PATCH (`Prefer: return=representation`, `$select` = primary id) and is kept in `loa_records`; without one, the next save is unconditional. Update from References persists from the app's current copy of each account (`self.loa_records`), not the screen's baseline.
- **Bulk save order**: each account is finished on its own (baseline if needed → master PATCH → audit row → rollback on audit failure) by 6 workers. Progress starts at "Checking audit history...". HTTP 429/503 are retried with `Retry-After` (up to 6 attempts, `DataverseClient._send`).
- **Delete is audited**: `delete_loa_record_entry` writes a `DELETE` row (field `Deleted Record`, `Active ➔ Deleted`, full last snapshot) first; no audit row, no delete; failed delete removes that row. DELETE rows cannot be reverted. Reference row deletes write a `FINANCE_UPDATE`/`CONVERSION_UPDATE` `Deleted Row` audit row the same way and only leave the screen when Dataverse confirms.
- **Strict saves**: `save_record(strict=True)` (Land Master, references) raises on a number/whole number/date it cannot read; `-`, `N/A`, `None` mean blank; `(1,234)` is −1234. The logs table uses `strict=False`.
- **Type/Mode** cannot be cleared by the drawer or the grid (`_required_field_cleared`). Grid drop-downs show unlisted stored values as "(not in list)".
- **Dates in the UI**: `normalizeDateInput` accepts M/D/YYYY, M/D/YY, YYYY-MM-DD, MMDDYYYY and Excel serials, returns MM/DD/YYYY or null (refused with a message). Pasted invalid dates are skipped and reported.
- **File reading**: values stay under their own header even when a header cell is blank; repeated headers get ` (2)`.
- **Reference Sync to Dataverse** (`_sync_reference_rows`): rows matched by key (Finance: TYPE + Corridor Index + Area Index + strict Title; Conversion: AREA-INDEX else Title Number) are updated with changed values only; new keys are created; keys matching several Dataverse rows or repeated in the file are skipped and listed; one `Bulk Sync` summary audit row. Blocked while reference edits are unsaved.
- **Authors**: `fetch_records` (non-custom fetches) maps `_createdby_value@…FormattedValue` to `createdby`; rows saved in this session show the signed-in user.
- **Per-account history check** filters on the Control columns that exist on the logs table (`cr63f_control1` and/or `cr63f_control`).
- **Land Master KPI Balance** = Σ max(0, TCP − Deductible), same as the Overview. Blank % Paid shows blank.
- **Build**: `build.ps1` (clean `.build-venv`); the spec excludes unrelated packages. `APP_VERSION` in `app.py` is the only version (Settings > About); the fake update check was removed.

### End-to-end test findings (2026-10-12)
- **Reference files are kept on reload**: `fetchRecords` (Land Master reload, after a drawer save, Overview Sync, start-up) calls `sync_*_from_dataverse(None, keep_loaded_file=True)`; a file loaded in this session (`_finance_file_this_session` / `_conversion_file_this_session`) is kept, and unsaved reference edits skip the reload. The explicit Sync Finance / Sync Conversion buttons still load Dataverse.
- **Reference edit baselines** are keyed by the id the screen sends (Dataverse ID or position in ALL rows); the save also looks under the row's other key. Rows without an ID are addressed by their position in `rows`, never in the filtered/sorted view. After a Sync to Dataverse both screens reload their rows so they carry the new IDs.
- **Reference grids** have no per-row delete button (decision 2026-10-12).
- **Overview week**: when the newest week is selected the Overview follows the newest week (requests `ALL`); an older week the user picked stays.
- **Entry Logs action filter** lists Deleted and Conversion Update; an empty filter value means All.
- **New Entry suggestions** copy an account's values as a template but not its Control Number.
- **Revert** is available from the Entry Logs row button and from the hover tooltip of a changed Land Master cell (`revertCellFromTooltip` → `revert_log_record`). The Overview has no Revert button.
- `tests/e2e/`: real Api + UI against `fake_dataverse.py` (Playwright). Run before every release.

### Decisions 2026-10-12 (round 3)
- **SA = Special Acquisition.** `FinanceLookupEngine.lookup` runs the normal cascade first; a LAND - SPECIAL account (`_SPECIAL_ACCOUNT_TYPES`) with no match and no conflict runs it again on SA rows only (Mode ignored; match type ends "(Special Acquisition)"). LAND and ROW accounts never use SA; CWA is never used. Conflicting SA duplicates are reported.
- **Recalculate Balance & % Paid** (`recalculate_payment_fields`, Settings > Maintenance): preview, then saves via `_persist_account_updates` (shared with Update from References: baseline → If-Match PATCH → `RECALCULATE` snapshot → rollback). Accounts without a TCP are skipped. Only Balance/% Paid change, so Entry Logs and the Overview show no field rows for it.
- **Edit account in the drawer**: pencil button / double-click row number (`openEditEntry`). Sends only changed fields; AREA-INDEX suggestions off while editing.
- **Revert in Entry Logs**: an Actions cell per row with a Revert button when `isLogRevertible` allows (latest change of the field, current value equals the logged new value, not calculated / DELETE / reference rows). The Overview has no Revert.
- **No delete buttons on Finance or Conversion rows** (`delete_*_record` stay in the API, audited).

### Conversion Status Ranking
- `loa_engine.conversion_rank` (mirrored by `convRank` in app.js): 3 = secured (whole word SECURED/EXEMPT/APPROVED/COMPLETED/CONVERTED and no negating word NOT/NO/YET/PENDING/DENIED/REJECTED/REVOKED/CANCELLED/INSPECTION), 1 = not started or ended (blank, NOT YET ENDORSED (including NOT YET ENDORSED FOR CONVERSION), N/A, CANCELED TITLE / CANCELLED TITLE, WITHDRAWN, PULLED-OUT / PULLED OUT), else 2 = in progress. Decided 2026-10-08: ended statuses rank 1, so a move to them is not counted as progress.

### UI Highlighting Rule
- Only rows and specific cells that have actual numerical or textual value changes upon clicking "Update from References" are highlighted (`.row-finance-populated`, `.cell-reference-updated`).
- Changed cells feature an emerald green highlight with glowing border outline, top-corner dot indicator, and hover tooltip (`Updated from Reference Data`).
- Highlights remain active while edits are staged and clear automatically when staged edits are saved to Dataverse or discarded. Rows with unchanged values remain uncolored.

---

## 3. UI Views & Features

1. **Land Master DB (`#view-records`)**:
   - Grid with column search filters, type filters, mode filters, column visibility picker, sticky columns, and Excel-like multi-cell selection & copy/paste.
   - "Update from References" button triggers cascading reconciliation of Finance and Conversion Status. Immediately updates the active grid with glowing emerald cells (`.cell-reference-updated`), isolates affected records using the auto-activated `Show Changed Only` filter toggle, and presents a confirmation modal with a scrollable diff breakdown (`AREA-INDEX`, `Field Changed`, `Previous Value`, `New Value`).
   - On confirmation, batch updates Dataverse master records and writes complete snapshot audit diffs in parallel with a live progress indicator (`Updating master records: X of Y...` -> `Writing audit logs: X of Y...`), synchronizing `state.baselineRecords` so subsequent runs detect 0 repeat changes.
   - If canceled, changes remain staged on the grid for manual review with the option to click "Save to Dataverse" or "Discard".
   - **New-entry duplicate rules (decision 2026-10-07)**: Control Number is required and must be unique (refused otherwise). AREA-INDEX may be shared by several accounts (e.g. parcels in one area), but the first save returns `needs_confirmation` naming the existing Control Number(s); the UI shows "AREA-INDEX Already Used" and resends with `allow_shared_area_index=True` only if the user confirms. If the Initial Entry audit row cannot be saved, the new account is deleted again (no Land Master change without its audit row).
   - Record lookups (`_find_or_fetch_master_record`, Overview account matching) go id → Control Number → AREA-INDEX → Title No → Index No and accept a non-id key only when exactly one record has it.
   - "Save to Dataverse" triggers a Confirmation Modal informing the user of the number of manually staged records that will be updated in Dataverse and recorded in the Entry Logs audit trail.
   - Live progress bar indicator displays real-time progress while batch writing to Dataverse.

2. **Entry Logs (`#view-logs`)**:
   - Table displaying all logged changes and new entries.
   - Parity features: resizable columns, freezable columns, 3-state sorting, header filter inputs, drag-and-drop column reordering, column visibility picker (`#btnColumnsPickerLogs`), and Excel-like multi-cell selection & copy/paste.
   - Audit rows cannot be deleted. **Revert** (`revert_log_record`, from the Entry Logs row button or the hover tooltip of a changed Land Master cell) checks the field still holds the logged new value, writes the old value back, and records a `REVERT` row. The loaded record changes only after both Dataverse writes succeed; if the `REVERT` row fails, every patched field (including recalculated Balance and % Paid) is restored.
   - Action Type filter (`NEW_ENTRY`, `MANUAL_EDIT`, `REFERENCE_UPDATE`, `FINANCE_UPDATE`, `REVERT`, `ALL`).
   - Work Week filter dropdown.
   - Excel export functionality.

3. **Overview (`#view-dashboard`) — UI v2, Option A (2026-10-06)**:
   - Analyst-first. Header: title, work week stepper (‹ WW-40 · Sep 30 – Oct 6 · vs WW-39 ›), "Synced h:mm" label, **Sync** (reloads Land Master, references and Entry Logs, then recomputes) and **Weekly report** (Excel export via `export_weekly_report`).
   - **This week**: 4 KPI cards — Accounts changed (field edits + new entries), Payments deducted, Balance outstanding, DAR secured — each a portfolio total with a week-over-week delta chip (▲/▼, green = good direction).
   - **Changed accounts** table: one row per account (first old → last new value per field within the week), source filter (All / Reference / Manual / New / Reverted), search, click a row to expand old → new diffs with **Open in Land Master**.
   - **Needs attention**: payment went down, deductible exceeds TCP, conversion status moved backwards (rank: secured 3 > in progress 2 > not endorsed 1).
   - **Balance outstanding** is the sum of each account's own balance (clamped at 0), so overpaid accounts don't offset others; the weekly balance change is computed the same way. TCP 0 with a deductible counts as fully paid. Conversion counts use Land Master status only (no fallback to the reference table).
   - **Portfolio**: Payment progress (stacked settled/partial/unpaid), Conversion progress (stacked secured/in progress/not endorsed + status legend), Accounts by mode (click to filter the whole page; banner with Clear filter).
   - Backend: `get_dashboard_analytics` additionally returns `week_summary`, `changed_accounts`, `needs_attention` (built in `_build_weekly_insights`). Missing log authors show as `Unknown` (no hard-coded fallback name).
   - Money on cards is compact (₱162.6M) with the exact value in the hover title.
   - Shell: no top bar; light sidebar with brand, nav, + New entry, context Sync button, and a Data status box (Dataverse / Finance / Conversion). `--mono` now maps to the sans stack with tabular figures.

4. **Finance Data (`#view-finance`)**:
   - Reference table preview with column picker, filtering, and Excel file upload / Dataverse sync capabilities.

5. **Conversion Data (`#view-conversion`)**:
   - Full parity reference table preview with column picker, live search, status filter, template download, Excel/CSV file upload, Dataverse sync, Excel export, and inline grid editing.

6. **Settings & Connection (`#view-settings`)**:
   - Azure AD Device Flow authentication.
   - Solution selector and dropdown mapping for `loa_table`, `logs_table`, `finance_table`, and `conversion_table`.
   - About & Updates card showing `APP_VERSION`, update status dot, release notes, progress bar, manual "Check for updates", "Download update", and "Update and restart" buttons.
   - Dark / Light mode and density controls.

---

## 4. Build & Deployment
- **Portable / Single-File EXE**: Built with `build.ps1` (`pyinstaller --noconfirm LOA-Monitoring-Tool.spec` -> `dist/LOA-Monitoring-Tool.exe`).
- **In-App Updates & Releases (Velopack)**:
  - Architecture mirrors `productivity-tool` (Ground Data Processing Tool): uses `updater.py` with `UpdateService` pointing at the static feed `https://github.com/jmlinguito-ux/loamonitoring/releases/latest/download/`.
  - Packaging: `build-release.ps1` sets `LOA_BUILD_ONEDIR=1` to create PyInstaller directory distribution and packages via `vpk pack` into `Releases/` (`Setup.exe`, `releases.win.json`, full/delta `.nupkg`).
  - Publishing: `publish-github-release.ps1` or automated via `.github/workflows/release.yml`.
- All tests located under `tests/` (`python -m unittest discover -s tests -p "test_*.py"`).

---

## 5. Dataverse Client Rules
- Choice (picklist) labels are matched exactly; a partial match is accepted only if it is unique (`LEASE` never becomes `SUBLEASE`).
- DateTime values follow each column's `DateTimeBehavior`: UserLocal → local time converted to UTC with `Z`; TimeZoneIndependent → wall-clock time without zone; DateOnly → date only.
- `$filter` values are quoted with `_odata_literal` (quotes doubled) and the filter is URL-encoded.
- `fetch_records` replaces raw values with display-formatted values; the raw `createdon` is kept as `__createdon_raw` for exact ordering.
