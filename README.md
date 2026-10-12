# LOA Monitoring Tool — Desktop Application

A Windows desktop application (Python + pywebview/WebView2) for monitoring land acquisition (LOA) accounts in Microsoft Dataverse. It reconciles the Land Master table against Finance and Land Conversion reference data and keeps an audit trail of every change.

For the full business rules, field mappings and design decisions, see [`MEMORY.md`](MEMORY.md).

---

## Key Features

1. **Update from References**
   - **Finance:** matching is TYPE-aware. ROW accounts use Finance `ROW` rows (Corridor Index, then Area-Index, then title). Other accounts use the `SALE` or `LEASE` row picked by their Mode (on `AREA-INDEX`, then title). A **LAND - SPECIAL** account with no SALE/LEASE match uses its `SA` (Special Acquisition) row. `CWA` rows, and `SA` rows for any other account, never supply amounts. The title fallback rejects titles whose numbers conflict.
   - **Conversion:** `Conversion Status` comes from the Land Conversion reference, matched on `AREA-INDEX` only.
   - **Duplicates:** same key and same TYPE with different amounts. The account is skipped and reported, never guessed.
   - **Derived fields:** `Balance = max(0, TCP − Deductible)` and `% Paid = Deductible ÷ TCP × 100` (stored to 1 decimal, 0–100 scale). They are saved on every save path and corrected whenever the stored value is stale.

2. **Land Master DB**
   - Grid editing with staged changes, then **Save to Dataverse**.
   - **Edit account** (pencil button on a row, or double-click the row number) opens the account in the form; only the fields changed there are saved.
   - New entries need a unique Control Number. An `AREA-INDEX` already used by another account asks for confirmation (several parcels can share an area). `Endorsed Date` defaults to today.
   - `Work Week` (Wednesday–Tuesday) comes from `Endorsed Date`.

3. **Entry Logs (audit trail)**
   - Each save writes one full-record snapshot to `cr63f_loamonitoringlogs`. Field-level changes are derived by comparing consecutive snapshots in time order.
   - Logs cannot be deleted from the app. **Revert** (button on the Entry Logs row, or the hover tooltip of a changed Land Master cell) writes the old value back plus a `REVERT` row, and is rolled back if that row can't be saved. It is offered only on the latest change of a Land Master field (not Balance / % Paid, deletes or reference-table rows).
   - Every Land Master change is undone if its audit row fails. **Deleting an account** writes a `DELETE` row with the account's last values first; nothing is deleted if that row can't be written. A deleted account cannot be restored with Revert.
   - Saves are conditional on the version each account was loaded at. If someone else saved the account in the meantime, the save is refused ("changed in Dataverse by someone else") instead of overwriting their change; reload and try again.

4. **Overview**
   - Shows one work week at a time (labelled with its year, e.g. `2026 · WW-40`) compared with the previous calendar week.
   - KPI cards, changed accounts, needs-attention alerts and portfolio charts. Balances are summed per account, each clamped at 0.

5. **Reference data**
   - Finance and Conversion reference tables can be loaded from Dataverse or from Excel/CSV files. A file you load stays in use until you click **Sync Finance** / **Sync Conversion** (reloading Land Master, saving an entry or the Overview Sync no longer replace it, and never discard unsaved reference edits). At start-up the app uses Dataverse when connected; the last file is reopened only when it is not.
   - **Sync to Dataverse** never duplicates rows: a file row updates the Dataverse row with the same key (Finance: TYPE + Corridor Index + Area Index + Title; Conversion: AREA-INDEX, else Title Number) and only changed values are sent. A key that matches several rows is skipped and listed. Each sync writes one summary row to Entry Logs. Save or discard edited cells before syncing.
   - Reference rows have no delete button: correct the source file and sync instead (deleting a row only made matching accounts stop updating).

6. **Maintenance (Settings)**
   - **Recalculate Balance & % Paid** recalculates every account with a TCP from its own TCP and Deductible (fixes % Paid stored on the old fraction scale and stale balances). It shows how many accounts would change before saving, saves through the same audited, version-checked path, and can be run again safely. Export Land Master first as a backup.

7. **Data entry rules**
   - Dates are typed or pasted as `M/D/YYYY`, `MM/DD/YYYY`, `YYYY-MM-DD`, `MMDDYYYY` or an Excel date number, and are stored as `MM/DD/YYYY`. An impossible date is refused with a message; it is never saved as blank.
   - An amount or date that cannot be read is refused by the save (`'1.2M' is not a valid number`) instead of being stored as blank. `-` and `N/A` mean blank. A typed `0` stays `0`.
   - Type and Mode cannot be cleared on an existing account, from the drawer or the grid. A stored value that is not in a drop-down list is shown as "(not in list)".
   - Bulk saves retry automatically when Dataverse asks the app to slow down (HTTP 429/503).

---

## Running from source

```powershell
pip install -r requirements.txt
python app.py
```

Tests (no Dataverse connection needed):

```powershell
python -m unittest discover -s tests -p "test_*.py"
node tests/test_revert_reference_ui.js
node tests/test_launch_review_ui.js
```

End-to-end tests of every screen against a stand-in Dataverse (headless browser): see [`tests/e2e/README.md`](tests/e2e/README.md).

## Standalone executable & In-App Updates

### Portable / Standalone Exe (Onefile)
Build with `build.ps1`, which creates a clean virtual environment (`.build-venv`) with only `requirements.txt`, runs the tests and then PyInstaller:

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

Output: `dist/LOA-Monitoring-Tool.exe`.

### In-App Updates & GitHub Releases (Velopack)
The tool integrates with [Velopack](https://velopack.io) (matching the Ground Data Processing Tool updater architecture) to check for, download, and apply releases directly from GitHub Releases (`jmlinguito-ux/loamonitoring`).

1. **Automatic background check:** Shortly after startup, the app checks the latest release feed on GitHub (`/releases/latest/download/`). If a newer release is detected, an "Update available" modal prompts the user with release notes.
2. **Download & Restart:** The update is downloaded in the background. Once ready, clicking "Update and restart" swaps the files in place and restarts the application seamlessly.
3. **Settings panel:** The About card in Settings provides manual "Check for updates", progress bar, and release notes display.

To package a new release locally or in CI:
```powershell
dotnet tool install -g vpk --version 1.2.0
.\build-release.ps1 -Version 1.0.8 -ReleaseNotes .\release-notes.md
```

To publish to GitHub Releases:
```powershell
$env:GH_TOKEN = '<github_token_with_contents_write>'
.\publish-github-release.ps1 -Publish
```
GitHub Actions workflow is also configured in `.github/workflows/release.yml` for 1-click releases via the Actions tab.

---

## Project Structure

```
LOA/
├── .github/workflows/     # CI/CD: release.yml for automated GitHub Releases
├── app.py                 # pywebview API bridge: saves, audit trail, Overview analytics
├── app_paths.py           # User data folder and settings paths
├── build.ps1              # Clean-environment build of portable onefile exe
├── build-release.ps1      # Velopack release packager for Windows installer
├── publish-github-release.ps1 # GitHub upload helper using vpk
├── updater.py             # Velopack background update service & lifecycle manager
├── dataverse.py           # Dataverse client (MSAL device-code auth, OData CRUD, metadata)
├── dataverse_config.json  # Tenant, client ID, environment URL, table names
├── loa_engine.py          # Lookups, % Paid, work weeks, status ranking, file I/O
├── MEMORY.md              # Business rules and decisions (keep in sync with code)
├── tests/                 # unittest suite
└── web/                   # UI: index.html, app.js, style.css
```
