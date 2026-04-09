# Merchant Activation Matcher

A small local web app that:

- Takes your **Master Merchant** Excel (with staff assignment, till number, phone, etc.)
- Takes **Corporate Inactive** Excel (merchant name + merchant code)
- Matches by **merchant code** and outputs a list of inactive merchants **with assigned staff details**
- Lets you **update/append** new merchants into your master list

## Setup (Windows)

Open PowerShell in this folder and run:

```bash
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

## Expected columns

The app is flexible about column names (it tries to auto-detect), but you should have:

- Master: merchant code, assigned staff, and any extra columns (till number, phone, etc.)
- Corporate inactive: merchant code (required) and merchant name (optional)

If your columns use different names (e.g. `MerchantCode`, `MERCHANT CODE`, `MID`), the app will attempt to detect them.
