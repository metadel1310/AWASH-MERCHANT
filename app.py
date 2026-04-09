from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, Optional

import pandas as pd
import streamlit as st


MASTER_STORE_PATH = "data/master_store.xlsx"
INACTIVE_TRACKER_PATH = "data/inactive_merchant_tracker.json"


@dataclass(frozen=True)
class ColumnGuess:
    merchant_code: str
    merchant_name: Optional[str] = None
    assigned_staff: Optional[str] = None
    phone: Optional[str] = None


def _norm_col(s: str) -> str:
    s = str(s or "").strip().lower()
    s = re.sub(r"[\s\-_]+", " ", s)
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    return s.strip()


def _normalize_merchant_code(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    s = s.replace("\u200b", "").replace("\ufeff", "")
    s = re.sub(r"\s+", "", s)
    # Common Excel artifact when numeric IDs are read as floats in some workflows
    # (even though we read with dtype=str, some exports may still include ".0")
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = s.upper()
    return s


def _auto_detect_6digit_code_column(df: pd.DataFrame) -> Optional[str]:
    """
    Pick the column that most consistently contains 6-digit merchant codes.

    We score each column by the fraction of non-empty values that look like a 6-digit code.
    """

    def _clean_cell(v) -> str:
        if v is None:
            return ""
        s = _normalize_merchant_code(v)
        # strip common wrappers like "MID: 123456" but keep only digits/spaces separators
        s = re.sub(r"[^\d]", "", s)
        return s

    def is_exact_6_digits(v) -> bool:
        s = _clean_cell(v)
        return bool(re.fullmatch(r"\d{6}", s))

    def is_long_numeric(v) -> bool:
        s = _clean_cell(v)
        return len(s) >= 8 and s.isdigit()

    def is_short_numeric(v) -> bool:
        s = _clean_cell(v)
        return 1 <= len(s) <= 3 and s.isdigit()

    best_col = None
    best_score = 0.0

    for col in df.columns:
        series = df[col]
        # sample up to first 2000 rows for speed on huge sheets
        sample = series.head(2000)
        non_empty = sample.dropna().astype(str).map(lambda x: x.strip())
        non_empty = non_empty[non_empty != ""]
        n = int(non_empty.shape[0])
        if n == 0:
            continue
        hits6 = int(non_empty.map(is_exact_6_digits).sum())
        long_count = int(non_empty.map(is_long_numeric).sum())
        short_count = int(non_empty.map(is_short_numeric).sum())

        # Primary: fraction of exact 6-digit values
        frac6 = hits6 / n
        # Penalize phone-number-like columns (many long numeric strings)
        frac_long = long_count / n
        # Penalize left-most index columns (many 1-3 digit sequences)
        frac_short = short_count / n

        score = frac6 - (0.75 * frac_long) - (0.35 * frac_short)

        # Require enough evidence: at least 10 rows looking like codes in the sample
        if hits6 < 10:
            continue
        if score > best_score:
            best_score = score
            best_col = col

    # Require a strong signal to avoid picking a random numeric column
    if best_col is not None and best_score >= 0.4:
        return best_col
    return None


def _auto_detect_phone_column_251(df: pd.DataFrame) -> Optional[str]:
    """Pick column where values (digits) most often start with 251 (e.g. Ethiopia mobile)."""

    def digits_only(v) -> str:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        return re.sub(r"\D+", "", str(v).strip())

    def looks_like_251_phone(v) -> bool:
        d = digits_only(v)
        return d.startswith("251") and len(d) >= 10

    best_col = None
    best_score = 0.0

    for col in df.columns:
        sample = df[col].head(2000).dropna()
        if sample.empty:
            continue
        sample = sample.astype(str).map(lambda x: x.strip())
        sample = sample[sample != ""]
        n = int(sample.shape[0])
        if n < 3:
            continue
        hits = int(sample.map(looks_like_251_phone).sum())
        score = hits / n
        if score > best_score:
            best_score = score
            best_col = col

    if best_col is not None and best_score >= 0.25:
        return best_col
    return None


def _guess_columns(df: pd.DataFrame) -> ColumnGuess:
    cols = list(df.columns)
    norm = {_norm_col(c): c for c in cols}

    def pick(candidates: Iterable[str]) -> Optional[str]:
        for cand in candidates:
            if cand in norm:
                return norm[cand]
        for cand in candidates:
            for k, orig in norm.items():
                if cand in k:
                    return orig
        return None

    merchant_code = pick(
        [
            "merchant code",
            "merchantcode",
            "mid",
            "merchant id",
            "merchantid",
            "code",
        ]
    )
    if not merchant_code:
        merchant_code = _auto_detect_6digit_code_column(df)
    if not merchant_code:
        raise ValueError(
            "Could not find a Merchant Code column by header or by detecting a 6-digit code column. "
            "Please add a header like 'Merchant Code' or ensure one column contains mostly 6-digit codes."
        )

    merchant_name = pick(["merchant name", "merchantname", "name", "business name"])
    assigned_staff = pick(
        ["assigned staff", "staff", "relationship officer", "ro", "agent", "sales rep"]
    )
    # Prefer the bank’s “Telephone” column name; fall back to common aliases, then 251-pattern detection.
    phone = pick(
        [
            "telephone",
            "telephone no",
            "telephone number",
            "phone",
            "mobile",
            "cell",
            "tel",
            "phone number",
            "msisdn",
            "contact",
            "merchant phone",
        ]
    )
    if not phone:
        phone = _auto_detect_phone_column_251(df)
    return ColumnGuess(
        merchant_code=merchant_code,
        merchant_name=merchant_name,
        assigned_staff=assigned_staff,
        phone=phone,
    )


def _read_excel(uploaded) -> pd.DataFrame:
    if uploaded is None:
        raise ValueError("No file provided.")
    return pd.read_excel(uploaded, dtype=str)


def _ensure_data_dir():
    import os

    os.makedirs("data", exist_ok=True)


def _load_inactive_tracker() -> dict:
    try:
        with open(INACTIVE_TRACKER_PATH, encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_inactive_tracker(tracker: dict) -> None:
    _ensure_data_dir()
    with open(INACTIVE_TRACKER_PATH, "w", encoding="utf-8") as f:
        json.dump(tracker, f, indent=2)


def _apply_still_inactive_and_update_tracker(
    matched: pd.DataFrame, merchant_code_col: Optional[str]
) -> pd.DataFrame:
    """
    Flag merchants who were already on a matched inactive list on a *previous* calendar day.
    Tracker stores last_seen per 6-digit code; codes that disappear from the matched list are
    dropped from the tracker (treated as reactivated). Same-day re-runs are not flagged.
    """
    if matched.empty or not merchant_code_col or merchant_code_col not in matched.columns:
        return matched

    tracker = _load_inactive_tracker()
    today = datetime.now().strftime("%Y-%m-%d")

    codes_series = matched[merchant_code_col].map(_six_digit_code_display)
    flags: list[str] = []
    for code in codes_series:
        if not code or len(code) != 6:
            flags.append("No")
            continue
        rec = tracker.get(code)
        if rec is None:
            flags.append("No")
            continue
        last = str(rec.get("last_seen", "")).strip()
        flags.append("Yes" if last and last < today else "No")

    out = matched.copy()
    out["Still_Inactive"] = flags

    cols = [c for c in out.columns if c != "Still_Inactive"]
    if merchant_code_col in cols:
        i = cols.index(merchant_code_col) + 1
        cols = cols[:i] + ["Still_Inactive"] + cols[i:]
    else:
        cols = ["Still_Inactive"] + cols
    out = out[cols]

    matched_codes = {c for c in codes_series if c and len(c) == 6}
    for k in list(tracker.keys()):
        if k not in matched_codes:
            del tracker[k]
    for code in matched_codes:
        prev = tracker.get(code, {})
        first = prev.get("first_seen") or today
        tracker[code] = {"first_seen": first, "last_seen": today}
    _save_inactive_tracker(tracker)

    return out


def _add_matched_row_numbers(matched: pd.DataFrame) -> pd.DataFrame:
    """Left-most column No. = 1, 2, 3, … for the matched inactive export."""
    if matched.empty:
        return matched
    out = matched.copy()
    if "No." in out.columns:
        out = out.drop(columns=["No."])
    out.insert(0, "No.", list(range(1, len(out) + 1)))
    return out


def _load_master_store() -> Optional[pd.DataFrame]:
    try:
        return pd.read_excel(MASTER_STORE_PATH, dtype=str)
    except Exception:
        return None


def _save_master_store(df: pd.DataFrame) -> None:
    _ensure_data_dir()
    with pd.ExcelWriter(MASTER_STORE_PATH, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Master")


def _dedupe_master(df: pd.DataFrame, merchant_code_col: str) -> pd.DataFrame:
    tmp = df.copy()
    tmp["_merchant_code_norm"] = tmp[merchant_code_col].map(_normalize_merchant_code)
    tmp = tmp.sort_values(by=["_merchant_code_norm"])
    tmp = tmp.drop_duplicates(subset=["_merchant_code_norm"], keep="last")
    tmp = tmp.drop(columns=["_merchant_code_norm"])
    return tmp


def _merge_side_col(df: pd.DataFrame, col: Optional[str], side: str) -> Optional[str]:
    """Resolve column name after merge (inactive vs master suffixes)."""
    if not col:
        return None
    suffixed = f"{col}_{side}"
    if suffixed in df.columns:
        return suffixed
    if col in df.columns:
        return col
    return None


def _six_digit_code_display(v) -> str:
    s = _normalize_merchant_code(v)
    d = re.sub(r"\D", "", s)
    if len(d) == 6:
        return d
    return d if len(d) == 6 else s


def _clean_cell(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip()
    if s.lower() == "nan":
        return ""
    return s


def _is_account_no_column(name: str) -> bool:
    n = _norm_col(str(name))
    if not n:
        return False
    if ("account" in n or "acct" in n) and ("no" in n or "number" in n):
        return True
    return False


def _format_account_for_excel(v) -> str:
    """
    Prefix a leading 0 on digit-only account values so Excel keeps them as text-like values
    (Excel often drops a single leading zero when the cell was numeric).
    """
    s = _clean_cell(v)
    if not s:
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    t = s.strip()
    if t.isdigit():
        return ("0" + t) if not t.startswith("0") else t
    return s


def _is_staff_email_column(name: str) -> bool:
    """Staff mail / email columns are for contacting staff, not for the inactive-merchant list."""
    n = _norm_col(name)
    if not n:
        return False
    has_mail = "mail" in n or "email" in n or "e mail" in n
    if "staff" in n and has_mail:
        return True
    if "relationship" in n and "officer" in n and has_mail:
        return True
    if re.search(r"(^|\s)ro[\s_]*(e[\s_-]*mail|mail)\b", n) or n.startswith("ro mail"):
        return True
    return False


def _output_master_columns(master_cols: list[str]) -> list[str]:
    """Master column order minus staff mail/email columns (not shown on inactive lists)."""
    return [c for c in master_cols if not _is_staff_email_column(c)]


def _split_dormant_staff(
    matched_raw: pd.DataFrame, master_guess: ColumnGuess
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Rows assigned to staff name 'DORMANT' are excluded from the active match list."""
    if matched_raw.empty:
        return matched_raw, pd.DataFrame(columns=matched_raw.columns)
    scol = _merge_side_col(matched_raw, master_guess.assigned_staff, "master")
    if not scol:
        return matched_raw, pd.DataFrame(columns=matched_raw.columns)
    is_dormant = (
        matched_raw[scol]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("DORMANT")
    )
    excluded = matched_raw[is_dormant].copy()
    kept = matched_raw[~is_dormant].copy()
    return kept, excluded


def _furnish_master_like(
    df: pd.DataFrame,
    master_df: pd.DataFrame,
    master_guess: ColumnGuess,
    inactive_guess: ColumnGuess,
    kind: str,
) -> pd.DataFrame:
    """
    Output matches the Master Excel layout: same columns in the same order (Telephone, Account No., etc.).
    Values come from the master side of the merge; merchant code is normalized to 6 digits.
    Unmatched rows: same headers; code/name filled from the inactive file where possible.
    Staff email columns are omitted (those are for contacting staff, not for this list).
    """
    master_cols = _output_master_columns(list(master_df.columns))
    out: dict[str, pd.Series] = {}

    for col in master_cols:
        merged_m = _merge_side_col(df, col, "master")

        if kind == "matched":
            if merged_m is not None:
                s = df[merged_m]
                if master_guess.merchant_code and col == master_guess.merchant_code:
                    out[col] = s.map(_six_digit_code_display)
                elif _is_account_no_column(col):
                    out[col] = s.map(_format_account_for_excel)
                else:
                    out[col] = s.map(_clean_cell)
            else:
                out[col] = pd.Series([""] * len(df), index=df.index, dtype=object)
            continue

        # unmatched
        filled = False
        if master_guess.merchant_code and col == master_guess.merchant_code:
            code_i = _merge_side_col(df, inactive_guess.merchant_code, "inactive")
            if code_i is not None:
                out[col] = df[code_i].map(_six_digit_code_display)
                filled = True
        if not filled and master_guess.merchant_name and col == master_guess.merchant_name:
            name_i = _merge_side_col(df, inactive_guess.merchant_name, "inactive")
            if name_i is not None:
                out[col] = df[name_i].map(_clean_cell)
                filled = True
        if not filled:
            if merged_m is not None:
                s = df[merged_m]
                if master_guess.merchant_code and col == master_guess.merchant_code:
                    out[col] = s.map(_six_digit_code_display)
                elif _is_account_no_column(col):
                    out[col] = s.map(_format_account_for_excel)
                else:
                    out[col] = s.map(_clean_cell)
            else:
                out[col] = pd.Series([""] * len(df), index=df.index, dtype=object)

    result = pd.DataFrame(out, columns=master_cols)

    if (
        kind == "matched"
        and master_guess.assigned_staff
        and master_guess.assigned_staff in result.columns
    ):
        sk = (
            result[master_guess.assigned_staff]
            .fillna("")
            .astype(str)
            .str.strip()
            .map(lambda x: x.lower() if x else "\uffff")
        )
        result = result.assign(_sk=sk).sort_values("_sk", kind="mergesort").drop(columns="_sk")

    return result


def _build_output(
    master_df: pd.DataFrame,
    inactive_df: pd.DataFrame,
    master_guess: ColumnGuess,
    inactive_guess: ColumnGuess,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    master = master_df.copy()
    inactive = inactive_df.copy()

    master["_merchant_code_norm"] = master[master_guess.merchant_code].map(
        _normalize_merchant_code
    )
    inactive["_merchant_code_norm"] = inactive[inactive_guess.merchant_code].map(
        _normalize_merchant_code
    )

    master = master[master["_merchant_code_norm"] != ""]
    inactive = inactive[inactive["_merchant_code_norm"] != ""]

    out = inactive.merge(
        master,
        how="left",
        on="_merchant_code_norm",
        suffixes=("_inactive", "_master"),
        indicator=True,
    )

    matched_raw = out[out["_merge"] == "both"].copy()
    unmatched_raw = out[out["_merge"] != "both"].copy()

    matched_active, excluded_dormant = _split_dormant_staff(matched_raw, master_guess)

    matched = _furnish_master_like(
        matched_active, master_df, master_guess, inactive_guess, "matched"
    )
    unmatched = _furnish_master_like(
        unmatched_raw, master_df, master_guess, inactive_guess, "unmatched"
    )
    excluded_dormant_furnished = _furnish_master_like(
        excluded_dormant, master_df, master_guess, inactive_guess, "matched"
    )

    total_master = len(master)
    corporate_rows = len(inactive)
    inactive_in_master = len(matched_raw)
    followup_count = len(matched_active)
    dormant_count = len(excluded_dormant)
    unmatched_corporate = len(unmatched_raw)
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M")
    active_portfolio = max(0, total_master - inactive_in_master)

    dashboard_df = _build_dashboard_summary(
        total_master=total_master,
        corporate_rows=corporate_rows,
        inactive_in_master=inactive_in_master,
        followup_count=followup_count,
        dormant_count=dormant_count,
        unmatched_corporate=unmatched_corporate,
        run_date=run_date,
    )
    chart_df = _dashboard_chart_df(active_portfolio, followup_count, dormant_count)

    return matched, unmatched, excluded_dormant_furnished, dashboard_df, chart_df


def _build_dashboard_summary(
    total_master: int,
    corporate_rows: int,
    inactive_in_master: int,
    followup_count: int,
    dormant_count: int,
    unmatched_corporate: int,
    run_date: str,
) -> pd.DataFrame:
    """Single table for Excel + app: counts and shares (decimals for Excel % format)."""
    active = max(0, total_master - inactive_in_master)
    pm = total_master if total_master else 1
    pc = corporate_rows if corporate_rows else 1

    rows: list[dict] = [
        {
            "Metric": "Report date",
            "Count": "",
            "Pct_of_master": None,
            "Pct_of_corporate_file": None,
            "Detail": run_date,
        },
        {
            "Metric": "Total merchants (master portfolio)",
            "Count": total_master,
            "Pct_of_master": None,
            "Pct_of_corporate_file": None,
            "Detail": "All rows with a valid merchant code",
        },
        {
            "Metric": "Active — not on this corporate inactive file",
            "Count": active,
            "Pct_of_master": active / pm,
            "Pct_of_corporate_file": None,
            "Detail": "Still transacting vs this extract",
        },
        {
            "Metric": "Inactive — on file, staff follow-up list",
            "Count": followup_count,
            "Pct_of_master": followup_count / pm,
            "Pct_of_corporate_file": None,
            "Detail": "Matched to master, excluding DORMANT",
        },
        {
            "Metric": "Inactive — on file, DORMANT assignment",
            "Count": dormant_count,
            "Pct_of_master": dormant_count / pm,
            "Pct_of_corporate_file": None,
            "Detail": "Excluded from main matched list",
        },
        {
            "Metric": "Inactive on file — in your master (subtotal)",
            "Count": inactive_in_master,
            "Pct_of_master": inactive_in_master / pm,
            "Pct_of_corporate_file": None,
            "Detail": "Follow-up + DORMANT",
        },
        {
            "Metric": "Corporate inactive file — total rows",
            "Count": corporate_rows,
            "Pct_of_master": None,
            "Pct_of_corporate_file": None,
            "Detail": "Rows with a detected merchant code",
        },
        {
            "Metric": "Corporate inactive — not found in master",
            "Count": unmatched_corporate,
            "Pct_of_master": None,
            "Pct_of_corporate_file": unmatched_corporate / pc,
            "Detail": "Other branch / wrong code / timing",
        },
    ]
    return pd.DataFrame(rows)


def _dashboard_chart_df(active: int, followup: int, dormant: int) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Status": ["Active", "Inactive (follow-up)", "Inactive (DORMANT)"],
            "Merchants": [active, followup, dormant],
        }
    )


def _write_dashboard_worksheet(
    workbook,
    dashboard_df: pd.DataFrame,
    chart_df: Optional[pd.DataFrame],
) -> None:
    ws = workbook.add_worksheet("Dashboard")
    hdr = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#2F5597",
            "font_color": "white",
            "border": 1,
            "valign": "vcenter",
        }
    )
    cell = workbook.add_format({"border": 1, "valign": "vcenter"})
    pct = workbook.add_format({"border": 1, "num_format": "0.0%", "align": "right"})
    num = workbook.add_format({"border": 1, "num_format": "#,##0", "align": "right"})
    meta = workbook.add_format(
        {"italic": True, "font_color": "#404040", "border": 1, "text_wrap": True}
    )
    sub = workbook.add_format(
        {"bold": True, "font_color": "#2F5597", "bottom": 1, "valign": "vcenter"}
    )

    for c, name in enumerate(dashboard_df.columns):
        ws.write(0, c, name, hdr)

    for r, row in enumerate(dashboard_df.values, start=1):
        for c, val in enumerate(row):
            if c == 0:
                ws.write(r, c, str(val) if val is not None else "", cell)
            elif c == 1:
                if val == "" or val is None or (isinstance(val, float) and pd.isna(val)):
                    ws.write(r, c, "", cell)
                elif isinstance(val, (int, float)) and not isinstance(val, bool):
                    v = float(val)
                    if v == int(v):
                        ws.write_number(r, c, int(v), num)
                    else:
                        ws.write_number(r, c, v, num)
                else:
                    try:
                        v = float(str(val).replace(",", ""))
                        ws.write_number(r, c, int(v) if v == int(v) else v, num)
                    except ValueError:
                        ws.write(r, c, str(val), cell)
            elif c in (2, 3):
                if val is not None and not (isinstance(val, float) and pd.isna(val)):
                    ws.write_number(r, c, float(val), pct)
                else:
                    ws.write(r, c, "", cell)
            else:
                ws.write(r, c, str(val) if val is not None else "", meta)

    ws.set_column(0, 0, 46)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 3, 18)
    ws.set_column(4, 4, 40)
    ws.freeze_panes(1, 0)

    if chart_df is not None and not chart_df.empty:
        base = len(dashboard_df) + 2
        ws.write(base, 0, "Portfolio split (insert Excel chart from this range)", sub)
        for c, name in enumerate(chart_df.columns):
            ws.write(base + 1, c, name, hdr)
        for i, row in enumerate(chart_df.values, start=2):
            for c, val in enumerate(row):
                if c == 1 and isinstance(val, (int, float)):
                    ws.write_number(base + i, c, int(val), num)
                else:
                    ws.write(base + i, c, str(val), cell)
        ws.set_column(0, 1, 22)


def _to_excel_bytes(
    sheets: dict[str, pd.DataFrame],
    dashboard_df: Optional[pd.DataFrame] = None,
    chart_df: Optional[pd.DataFrame] = None,
) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            safe = name[:31]
            df.to_excel(writer, index=False, sheet_name=safe)

        if dashboard_df is not None and not dashboard_df.empty:
            _write_dashboard_worksheet(writer.book, dashboard_df, chart_df)

    return bio.getvalue()


st.set_page_config(page_title="Inactive Merchant Matcher", layout="wide")
st.title("Inactive Merchant Matcher (Master vs Corporate List)")

st.write(
    "Upload your **Master Merchant** Excel and the **Corporate Inactive** Excel. "
    "The app matches by **Merchant Code**. The export mirrors your **Master** columns (minus **Staff Mail** / staff email), "
    "including **Telephone** and **Account No.** Merchants assigned to staff named **DORMANT** are excluded from the main matched list. "
    "If a merchant is still inactive on a **later day**, they are flagged **Still_Inactive**. "
    "Each download includes a **Dashboard** sheet with active/inactive counts and percentages."
)

tab_match, tab_master = st.tabs(["Match Inactive → Staff", "Master List (Store/Update)"])

with tab_master:
    st.subheader("Master list store")
    st.caption(
        "Optional: keep a saved master list inside this app folder so you don't have to upload it every time."
    )

    stored = _load_master_store()
    if stored is None:
        st.info("No saved master list yet. Upload one to save it.")
    else:
        st.success(f"Loaded saved master list: {MASTER_STORE_PATH}")
        st.dataframe(stored.head(50), use_container_width=True)

    master_upload_for_store = st.file_uploader(
        "Upload Master Excel to save/replace",
        type=["xlsx", "xls"],
        key="master_store_upload",
    )

    col_a, col_b = st.columns(2)
    with col_a:
        replace_store = st.button("Save (replace store)", type="primary")
    with col_b:
        append_store = st.button("Append into store (dedupe by Merchant Code)")

    if master_upload_for_store is not None and (replace_store or append_store):
        try:
            incoming = _read_excel(master_upload_for_store)
            incoming_guess = _guess_columns(incoming)

            if replace_store or stored is None:
                combined = incoming
            else:
                combined = pd.concat([stored, incoming], ignore_index=True)

            combined = _dedupe_master(combined, incoming_guess.merchant_code)
            _save_master_store(combined)
            st.success("Master list saved.")
            st.dataframe(combined.head(50), use_container_width=True)
        except Exception as e:
            st.error(str(e))

with tab_match:
    st.subheader("Match corporate inactive list to staff assignments")

    use_saved_master = st.checkbox(
        "Use saved master list (if available)",
        value=True,
    )

    left, right = st.columns(2)
    with left:
        master_upload = st.file_uploader(
            "Upload Master Merchant Excel",
            type=["xlsx", "xls"],
            key="master_match_upload",
            disabled=use_saved_master and _load_master_store() is not None,
        )
    with right:
        inactive_upload = st.file_uploader(
            "Upload Corporate Inactive Excel",
            type=["xlsx", "xls"],
            key="inactive_upload",
        )

    only_matched = st.checkbox("Show only matched (assigned) rows", value=True)

    st.caption(
        "**Still inactive:** merchants who were already on your matched inactive export on a **previous calendar day**, "
        "and appear again today, get **Still_Inactive = Yes**. History is stored in `data/inactive_merchant_tracker.json` "
        "(codes that drop off the inactive list are removed from tracking)."
    )
    if st.button("Clear still-inactive history", key="clear_inactive_tracker"):
        _save_inactive_tracker({})
        st.success("Carryover history cleared. Next run will not flag any merchant as still inactive.")

    if st.button("Run matching", type="primary", disabled=inactive_upload is None):
        try:
            if use_saved_master and _load_master_store() is not None:
                master_df = _load_master_store()
            else:
                if master_upload is None:
                    raise ValueError("Please upload a Master Merchant Excel (or enable saved master).")
                master_df = _read_excel(master_upload)

            inactive_df = _read_excel(inactive_upload)

            master_guess = _guess_columns(master_df)
            inactive_guess = _guess_columns(inactive_df)

            matched, unmatched, excluded_dormant, dashboard_df, chart_df = _build_output(
                master_df=master_df,
                inactive_df=inactive_df,
                master_guess=master_guess,
                inactive_guess=inactive_guess,
            )

            matched = _apply_still_inactive_and_update_tracker(
                matched, master_guess.merchant_code
            )
            matched = _add_matched_row_numbers(matched)

            st.subheader("Portfolio overview")
            dr = dashboard_df.reset_index(drop=True)

            def _dash_count(row_idx: int) -> int:
                v = dr.iloc[row_idx]["Count"]
                if v == "" or v is None or (isinstance(v, float) and pd.isna(v)):
                    return 0
                try:
                    return int(v)
                except (TypeError, ValueError):
                    return 0

            total_master = _dash_count(1)
            active_n = _dash_count(2)
            follow_n = _dash_count(3)
            dorm_n = _dash_count(4)
            corp_rows = _dash_count(6)
            not_in_master = _dash_count(7)
            pm = total_master if total_master else 1

            k1, k2, k3, k4 = st.columns(4)
            k1.metric(
                "Active",
                f"{active_n:,}",
                f"{100 * active_n / pm:.1f}% of master",
            )
            k2.metric(
                "Inactive (follow-up)",
                f"{follow_n:,}",
                f"{100 * follow_n / pm:.1f}% of master",
            )
            k3.metric(
                "Inactive (DORMANT)",
                f"{dorm_n:,}",
                f"{100 * dorm_n / pm:.1f}% of master",
            )
            k4.metric(
                "Corporate file",
                f"{corp_rows:,} rows",
                f"{not_in_master:,} not in master",
            )

            st.caption("Split of your master portfolio (active vs inactive on this file)")
            st.bar_chart(chart_df.set_index("Status"))

            with st.expander("Full metrics table (Excel **Dashboard** sheet)"):
                st.dataframe(dashboard_df, use_container_width=True, hide_index=True)

            st.write("Detected columns")
            st.json(
                {
                    "master": master_guess.__dict__,
                    "inactive": inactive_guess.__dict__,
                }
            )

            if len(excluded_dormant) > 0:
                st.info(
                    f"Excluded **{len(excluded_dormant)}** merchant(s) assigned to staff **DORMANT** "
                    "(see sheet *Excluded_Dormant_Staff* in the download)."
                )

            if only_matched:
                st.subheader(f"Matched inactive merchants ({len(matched)})")
                st.dataframe(matched, use_container_width=True)
            else:
                st.subheader(f"Matched inactive merchants ({len(matched)})")
                st.dataframe(matched, use_container_width=True)
                st.subheader(f"Unmatched inactive merchants ({len(unmatched)})")
                st.dataframe(unmatched, use_container_width=True)
                if len(excluded_dormant) > 0:
                    st.subheader(f"Excluded (DORMANT staff) — not in matched list ({len(excluded_dormant)})")
                    st.dataframe(excluded_dormant, use_container_width=True)

            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            sheets: dict[str, pd.DataFrame] = {
                "Matched_Inactive": matched,
                "Unmatched_Inactive": unmatched,
            }
            if len(excluded_dormant) > 0:
                sheets["Excluded_Dormant_Staff"] = excluded_dormant
            xbytes = _to_excel_bytes(
                sheets,
                dashboard_df=dashboard_df,
                chart_df=chart_df,
            )
            st.download_button(
                "Download Excel results",
                data=xbytes,
                file_name=f"inactive_matched_{now}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(str(e))
