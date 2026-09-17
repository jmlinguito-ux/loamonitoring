# LOA Monitoring Tool — Desktop Application

A high-performance Windows desktop application built with Python and pywebview/WebView2, integrated with Microsoft Dataverse (`cr63f_loamonitoring`) and implementing real-time Excel lookup autofill from `APPENDED_FINANCE`.

---

## Key Features

1. **Live Excel Lookup & Autofill Engine**:
   Implements the exact `=LET(...)` formula logic:
   - **Primary Check**:
     - When `TYPE = "ROW"`: Looks up `INDEX NO` against `APPENDED_FINANCE[Corridor Index]` (filtered by `TYPE="ROW"`).
     - Otherwise: Looks up `AREA-INDEX` against `APPENDED_FINANCE[Area Index]`.
   - **Fallback Check**:
     - If no primary match and `TITLE MOD` is populated: Looks up `TITLE MOD` against `APPENDED_FINANCE[Title Mod]`.
   - **Autofills**:
     - `Total Contract Price`
     - `Total Payment Deductible to TCP`
     - `Balance` = `Total Contract Price - Total Payment Deductible to TCP`
     - `% Paid`

2. **Form Entry & Logging**:
   - Interactive entry drawer with real-time feedback badges.
   - **Explicit Save Button**: Entries are treated as logs/confirmed updates and written to Dataverse only when the user clicks **Save Entry to Dataverse**.

3. **Dataverse Integration**:
   - Direct connection to `cr63f_loamonitoring` using MSAL device-code authentication.
   - Metadata discovery and OData v9.2 CRUD operations.

4. **Appended Finance Reference Hub**:
   - Pull reference data directly from Dataverse or upload local Excel/CSV files.
   - Automatic file remembering and recent files selector.

5. **Batch Processing & Export**:
   - Bulk upload CSV/Excel files and calculate all derived financial fields across hundreds of records in seconds.
   - Export enriched records to Excel.

---

## Standalone Executable

A standalone single-file Windows executable has been built and is ready to run directly without installing Python:
- **Location**: [`dist/LOA-Monitoring-Tool.exe`](file:///c:/Users/Jinguito/Desktop/CSV/LOA/dist/LOA-Monitoring-Tool.exe)

To rebuild the `.exe` at any time:
```powershell
pyinstaller LOA-Monitoring-Tool.spec --noconfirm --clean
```

---

## Installation & Running from Source

---

## Project Structure

```
LOA/
├── app.py                     # Main application entrypoint & pywebview API bridge
├── app_paths.py               # Path management & user data persistence
├── dataverse.py               # Dataverse connector (MSAL auth, OData CRUD, metadata)
├── dataverse_config.json      # Connection settings (tenant, client ID, environment URL)
├── loa_engine.py              # Pure-Python Excel lookup & calculation engine
├── requirements.txt           # Python dependencies
├── tests/
│   └── test_loa_engine.py     # Unit tests verifying lookup logic and math
└── web/
    ├── app.js                 # UI logic, live input listeners & grid renderer
    ├── index.html             # Main interface layout (Hub, Form, Reference, Settings)
    └── style.css              # Modern IBM Plex theme & responsive styles
```

