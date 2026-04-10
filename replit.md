# Dengel Merchant Solver

## Overview
A Streamlit web application for Awash Bank branch teams that matches a "Master Merchant" Excel list (staff assignments + contact details) against a "Corporate Inactive" Excel list. Identifies which inactive merchants are assigned to which staff for follow-up and activation.

## Branding
- App name: **Dengel Merchant Solver**
- Theme: Awash Birr Pro–inspired (orange-red gradient, clean card layout)
- Colors: Primary #C0311E → #E85B2C gradient, accent gold, light gray background
- Works on mobile browsers, desktop, and web — share the deployed URL to any branch

## Tech Stack
- **Language:** Python 3.12
- **Web Framework:** Streamlit (responsive, works on mobile)
- **Data Processing:** Pandas
- **Excel Integration:** openpyxl, xlrd, xlsxwriter

## Project Structure
- `app.py` — Main Streamlit application (UI, logic, data processing, Excel export)
- `requirements.txt` — Python dependencies
- `data/` (created at runtime):
  - `master_store.xlsx` — Persisted master merchant list
  - `inactive_merchant_tracker.json` — Cross-day "still inactive" tracking

## Running the App
Workflow "Start application":
```
streamlit run app.py --server.port 5000 --server.address 0.0.0.0 --server.headless true
```

## Key Features
- Upload Master Merchant + Corporate Inactive Excel files
- Auto-detects Merchant Code and Phone Number columns
- Matches inactive merchants to assigned staff (excludes DORMANT)
- Tracks "Still_Inactive" status across days via JSON tracker
- Exports multi-sheet Excel (Matched, Unmatched, Excluded_Dormant_Staff, Dashboard)
- No data visualization charts in the UI (removed); dashboard data available in Excel export only
- Deployable as a shared web link — works on mobile and desktop browsers
