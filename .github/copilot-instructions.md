# Copilot Instructions for Homeopathy Name Search App

## Project Overview
- This is a single-file Kivy desktop/mobile app for searching homeopathic remedy names.
- Main file: `homeopathy_name_search_app.py` (contains both UI and data logic).
- Data source: `remedies.xlsx` (must be in the same folder; expects columns "Latin" and "Common").

## Architecture & Data Flow
- UI is defined inline using Kivy's KV language string.
- Data is loaded from Excel using pandas and openpyxl.
- Two lookup maps are built: Latin→Common and Common→Latin (case-insensitive).
- Search direction is auto-detected or user-selected via Spinner.
- Results are displayed in a scrollable list; fuzzy matching is supported for multi-word queries.

## Developer Workflows
- **Run app (desktop):**
  ```pwsh
  python homeopathy_name_search_app.py
  ```
- **Dependencies:**
  - Install with: `pip install kivy pandas openpyxl`
- **Excel file requirements:**
  - Must have headers: "Latin" and "Common" (case-insensitive).
  - To change headers, update `latin_col` and `common_col` in the code.
- **Android build:**
  - Use Buildozer; consider converting Excel to CSV to reduce APK size.

## Project-Specific Patterns
- All logic is in one file; no external modules.
- Error messages are shown in the UI label (`status_text`).
- If pandas is missing, the app disables Excel loading and prompts for install.
- Fuzzy search splits queries into words and matches each part in both columns.
- UI widgets are accessed via `self.ids` (Kivy convention).

## Integration Points
- External dependencies: pandas, openpyxl, kivy.
- Data file: `remedies.xlsx` (or fallback to `remedies.csv`).
- No network or API calls; all data is local.

## Example Patterns
- To add new search columns, update mapping logic in `_load_df` and UI hints.
- To customize UI, edit the `KV` string in the main file.

---

If any section is unclear or missing, please provide feedback to improve these instructions.
