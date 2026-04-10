# Merchant Activation Matcher

## Overview
A Streamlit-based web application that matches a "Master Merchant" Excel list (with staff assignments and contact details) against a "Corporate Inactive" Excel list (inactive merchant codes). It identifies which inactive merchants are assigned to which staff members for follow-up and activation efforts.

## Tech Stack
- **Language:** Python 3.12
- **Web Framework:** Streamlit
- **Data Processing:** Pandas
- **Excel Integration:** openpyxl, xlrd, xlsxwriter

## Project Structure
- `app.py` — Main Streamlit application with all logic and UI
- `requirements.txt` — Python dependencies
- `data/` (created at runtime) — Stores persisted master list and inactive merchant tracker JSON

## Running the App
The app runs via the "Start application" workflow:
```
streamlit run app.py --server.port 5000 --server.address 0.0.0.0 --server.headless true
```

## Key Features
- Upload Master Merchant and Corporate Inactive Excel files
- Auto-detects Merchant Code and Phone Number columns
- Matches inactive merchants to assigned staff
- Tracks "still inactive" status across multiple days using a JSON tracker
- Exports multi-sheet Excel reports with Dashboard, Matched, and Unmatched sheets
- Excludes DORMANT-assigned merchants from the main matched list
