#!/usr/bin/env python3
"""
att_salesforce_prep.py

Reconcile an AT&T Premier asset export against a Salesforce "AT&T Client Assets"
report and produce a Salesforce-import-ready CSV containing ONLY the assets that
are not already in Salesforce.

WORKFLOW
--------
Before each run, two files are placed in INPUT_DIR:
  1. The AT&T Premier asset export (.xlsx) — has junk rows above the real header.
  2. The Salesforce "AT&T Client Assets" report (.xlsx, sometimes .csv).

This script:
  * Reads both files as pure strings (no number coercion — long IMEI/ICCID/phone
    values keep every digit, no scientific notation, no trailing ".0").
  * Locates the IMEI column in each file by fuzzy (case/space-insensitive) match.
  * Anti-joins: keeps only Premier rows whose IMEI is NOT already in Salesforce
    (the manual "VLOOKUP + filter for Not Found" step).
  * Builds the Salesforce upload table with the required columns and formatting.
  * Quarantines rows with an invalid Universal ID into a separate REVIEW csv.
  * Permanently locks any Wireless Number ever marked Cancelled or Chargeback:
    once seen, that number can never be uploaded back to Active (or any other
    status) by a later run. Blocked attempts are written to a BLOCKED_HISTORY
    csv instead of the clean upload file. The lock is remembered across runs
    in state/cancelled_chargeback_history.csv (do not delete this file).
  * Writes ATT_Import_<date>.csv (and _REVIEW.csv / _BLOCKED_HISTORY.csv if
    needed) and prints a summary.

It deliberately STOPS at producing CSVs. A human runs the Salesforce import wizard
afterward — there is no Salesforce API / write code here.
"""

import argparse
import glob
import os
import re
import shutil
import sys
import time
from datetime import date, datetime

import pandas as pd

# ===========================================================================
# CONFIG
# ===========================================================================

# Input/output default to folders sitting NEXT TO this script, so the whole
# tool folder can be copied to any machine and "just work". The run_att_prep.bat
# runner also passes these explicitly. Drop the two downloaded files in "input".
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(SCRIPT_DIR, "input")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")

# Filename globs for auto-detection (used when --premier/--sf aren't given).
# Premier "Inventory" export downloads as detail_report.xlsx (maybe with a
# " (1)" suffix). The Salesforce report is "All Active ATT Client(s) Assets-...".
PREMIER_GLOB = "detail_report*.xlsx"

SF_GLOB = "*ATT*Client*Assets*"  # matches the .xlsx (or .csv) Salesforce export

# FALLBACK ONLY. The script now auto-detects the Premier header row by matching
# the known column names (so it survives the export gaining/losing a preamble
# row). This value is used only if auto-detection fails. A real detail_report.xlsx
# has 5 metadata rows with the header on row 6.
PREMIER_SKIP_ROWS = 5

# Salesforce report has a normal header row at the top.
SF_SKIP_ROWS = 0

# SAFETY: when files are auto-detected (no --premier/--sf given, i.e. the weekly
# scheduled run), refuse to run if the newest matching file is older than this
# many days. This stops a forgotten download from silently reprocessing last
# week's files. Legit weekly downloads are 0-3 days old; last week's are ~7.
MAX_INPUT_AGE_DAYS = 6

# SAFETY: after a successful run, move the input files that live inside the
# input folder into input/_processed/<date>/ so they can never be re-used next
# week. (Files passed explicitly with --premier/--sf are left where they are.)
ARCHIVE_PROCESSED_INPUTS = True

# Salesforce "Client" code for the batch. In many orgs this is a constant
# master/parent account code shared by every row in the report, with the
# individual sub-accounts identified by Universal ID rather than this field.
# TODO <<FILL>>: set this to YOUR Salesforce client/account identifier.
# Override per-run with --client if a different batch ever uses another code.
CLIENT_VALUE = "YOUR-CLIENT-CODE"

# Asset Owner Type remap. Comparison is case-insensitive on the trimmed value.
OWNER_TYPE_MAP = {"standard": "Client", "no contract": "ICD"}

# TODO <<FILL>>: regex a VALID Universal ID must match. Rows that fail are quarantined.
# Example below matches things like "U-12345" / "U12345". Tighten to your real format.
UNIVERSAL_ID_REGEX = r"^U-?\w+$"

# Once a Wireless Number is ever seen with a Cancelled or Chargeback status, it
# is permanently locked: no later run may upload that number back to Active
# (or any other status). This file is the persistent memory of that lock and
# must survive between runs — it is NOT part of the per-run output/ folder and
# is never archived or deleted automatically.
HISTORY_DIR = os.path.join(SCRIPT_DIR, "state")
HISTORY_PATH = os.path.join(HISTORY_DIR, "cancelled_chargeback_history.csv")
HISTORY_COLUMNS = [
    "Phone Digits", "Phone Number", "Blocked Status",
    "Identifier", "Universal ID", "First Marked", "Last Seen",
]

# Status strings that trigger the permanent lock. Matching is case/space/punct
# -insensitive (see _norm_status_key), so "Cancelled", "canceled", "CANCELED",
# "Charge Back", "chargeback", etc. all match.
CANCELLED_STATUS_KEYS = {"cancelled", "canceled"}
CHARGEBACK_STATUS_KEYS = {"chargeback"}

# ---------------------------------------------------------------------------
# Premier column aliases (case/space-insensitive). Add real header spellings
# from an actual export here. The FIRST alias that matches a header wins.
# TODO <<FILL>>: confirm/extend these against a real Premier export.
# ---------------------------------------------------------------------------
# First alias = real header confirmed in detail_report.xlsx; rest are fallbacks.
PREMIER_ALIASES = {
    "imei":           ["Device IMEI", "IMEI", "IMEI/MEID", "IMEI / MEID", "Identifier"],
    "iccid":          ["SIM number (ICCID)", "ICCID", "ICC ID", "SIM ICCID", "SIM Card Number"],
    "wireless_number":["Wireless number", "Wireless Number", "Phone Number", "MSISDN", "Mobile Number", "WTN"],
    "universal_id":   ["UDL 3", "Universal ID", "Universal Id", "UniversalID", "Univ ID"],
    "status":         ["Status", "Current Status", "Line Status", "Device Status"],
    "contract_type":  ["Contract type", "Contract Type", "Contract", "Agreement Type", "Term Type"],
}

# Salesforce IMEI column aliases (case/space-insensitive).
# First alias = real header confirmed in the "All Active ATT Clients Assets" report.
SF_IMEI_ALIASES = ["Client Asset: Identifier", "Identifier", "IMEI", "Asset Identifier", "Device IMEI"]

# Output column order + headers — matches the desired Salesforce upload layout.
OUTPUT_COLUMNS = [
    "Phone Number",
    "Current Status",
    "Universal ID",
    "Identifier",
    "ICCID",
    "Asset Owner Type",
    "Client",
]

# ===========================================================================
# HELPERS
# ===========================================================================


def _norm_header(text: str) -> str:
    """Normalize a header for fuzzy comparison: lowercase, strip non-alphanumerics."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def digits_only(value) -> str:
    """Return only the digit characters of a value. NaN/None -> ''."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\D", "", str(value))


def repair_numeric_string(value) -> str:
    """
    Repair string artifacts left by spreadsheet number parsing.

    Handles trailing '.0' (e.g. '123456789.0') and scientific notation
    (e.g. '1.23457E+14') back into a plain integer string, while leaving
    genuine text untouched.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return ""
    # Scientific notation like 1.23457E+14
    if re.fullmatch(r"[-+]?\d+(\.\d+)?[eE][-+]?\d+", s):
        try:
            return str(int(float(s)))
        except (ValueError, OverflowError):
            return s
    # Trailing .0 (whole number that got floated)
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".")[0]
    return s


def find_column(df_columns, aliases, file_label):
    """
    Find the first column in df_columns matching any alias (fuzzy).
    Raises a clear error listing actual headers if nothing matches.
    """
    norm_to_actual = {_norm_header(c): c for c in df_columns}
    for alias in aliases:
        key = _norm_header(alias)
        if key in norm_to_actual:
            return norm_to_actual[key]
    raise SystemExit(
        f"\nERROR: Could not locate a required column in {file_label}.\n"
        f"  Tried aliases: {aliases}\n"
        f"  Headers actually present:\n    "
        + "\n    ".join(repr(c) for c in df_columns)
        + "\n  --> Add the correct spelling to the alias list in the CONFIG block.\n"
    )


def _read_raw(path):
    """Read an xlsx/csv with NO header and every value as a string."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, dtype=str, header=None, engine="openpyxl")
    if ext == ".csv":
        return pd.read_csv(path, dtype=str, header=None, keep_default_na=False)
    raise SystemExit(f"ERROR: Unsupported file type '{ext}' for {path}")


def read_with_header_detect(path, alias_groups, fallback_skiprows, label,
                            min_groups, scan_rows=30):
    """
    Read a file as strings, auto-detecting which row is the real header.

    Instead of trusting a fixed row count (which the AT&T Premier export changes
    by a row now and then), this scans the first `scan_rows` rows and picks the
    one whose cells best match the expected column names in `alias_groups`
    (a list of alias-lists, one per field we need). The row matching the most
    distinct fields wins; ties go to the earliest row. If nothing matches at
    least `min_groups` fields, it falls back to `fallback_skiprows` and warns.

    Returns (dataframe, header_row_index).
    """
    raw = _read_raw(path)
    norm_groups = [{_norm_header(a) for a in aliases} for aliases in alias_groups]

    best_row, best_score = None, -1
    limit = min(scan_rows, len(raw))
    for r in range(limit):
        cells = {
            _norm_header(c)
            for c in raw.iloc[r].tolist()
            if c is not None and str(c).strip() != ""
        }
        score = sum(1 for g in norm_groups if g & cells)
        if score > best_score:
            best_score, best_row = score, r

    if best_row is None or best_score < min_groups:
        print(
            f"  NOTE: could not confidently auto-detect the {label} header "
            f"(best match {max(best_score, 0)} of {len(norm_groups)} fields). "
            f"Falling back to the configured skip of {fallback_skiprows} row(s)."
        )
        header_row = fallback_skiprows
    else:
        header_row = best_row

    header = [("" if v is None else str(v)).strip() for v in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1:].copy()
    data.columns = header
    data = data.reset_index(drop=True)
    return data, header_row


def autodetect(input_dir, pattern, label):
    """Return the most recently modified file in input_dir matching pattern."""
    matches = glob.glob(os.path.join(input_dir, pattern))
    matches = [m for m in matches if os.path.isfile(m) and not os.path.basename(m).startswith("~$")]
    if not matches:
        raise SystemExit(
            f"ERROR: No {label} file found in {input_dir} matching '{pattern}'.\n"
            f"  Pass --{'premier' if label == 'Premier' else 'sf'} explicitly, "
            f"or fix the glob in the CONFIG block."
        )
    newest = max(matches, key=os.path.getmtime)
    return newest


def check_freshness(path, label, max_age_days):
    """Stop the run if an auto-detected file is older than max_age_days."""
    age_days = (time.time() - os.path.getmtime(path)) / 86400.0
    if age_days > max_age_days:
        modified = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
        raise SystemExit(
            "\n" + "!" * 70 + "\n"
            f"STOPPING: the {label} file looks STALE.\n"
            f"  File   : {os.path.basename(path)}\n"
            f"  Last changed: {modified}  ({age_days:.1f} days ago)\n"
            f"  This is older than the {max_age_days}-day limit, which usually means\n"
            f"  this week's file was NOT downloaded into the input folder.\n\n"
            f"  Nothing was processed. Download this week's report(s), drop them in\n"
            f"  the input folder, and run again. (To override, raise --max-age-days.)\n"
            + "!" * 70 + "\n"
        )


def archive_inputs(used_paths, input_dir, stamp):
    """Move used input files (that live inside input_dir) into input/_processed/<date>/."""
    dest_dir = os.path.join(input_dir, "_processed", stamp)
    moved = []
    for p in used_paths:
        try:
            inside = os.path.commonpath([os.path.abspath(p), os.path.abspath(input_dir)]) == os.path.abspath(input_dir)
        except ValueError:
            inside = False  # different drive
        if not inside:
            continue  # explicitly-passed file outside the input folder: leave it alone
        os.makedirs(dest_dir, exist_ok=True)
        target = os.path.join(dest_dir, os.path.basename(p))
        n = 1
        while os.path.exists(target):
            stem, ext = os.path.splitext(os.path.basename(p))
            target = os.path.join(dest_dir, f"{stem} ({n}){ext}")
            n += 1
        shutil.move(p, target)
        moved.append(target)
    return moved


def format_phone(value) -> str:
    """Format a phone number as 000-000-0000. Drop a leading '1' on 11-digit numbers."""
    d = digits_only(value)
    if len(d) == 11 and d.startswith("1"):
        d = d[1:]
    if len(d) == 10:
        return f"{d[0:3]}-{d[3:6]}-{d[6:10]}"
    return d  # leave anything non-standard as plain digits for review


def remap_owner_type(value, unmapped_collector):
    """Remap Contract Type via OWNER_TYPE_MAP; collect anything unmapped."""
    raw = "" if value is None else str(value).strip()
    mapped = OWNER_TYPE_MAP.get(raw.lower())
    if mapped is not None:
        return mapped
    if raw != "":
        unmapped_collector.add(raw)
    return raw


def _norm_status_key(value) -> str:
    """Normalize a status string for matching (lowercase, letters only)."""
    return re.sub(r"[^a-z]", "", str(value).lower())


def classify_blocked_status(value):
    """Return 'Cancelled', 'Chargeback', or None if the status isn't locked."""
    key = _norm_status_key(value)
    if key in CANCELLED_STATUS_KEYS:
        return "Cancelled"
    if key in CHARGEBACK_STATUS_KEYS:
        return "Chargeback"
    return None


def load_history(path):
    """Load the persistent Cancelled/Chargeback lock list. Empty if none yet."""
    if not os.path.exists(path):
        return pd.DataFrame(columns=HISTORY_COLUMNS)
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    for col in HISTORY_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[HISTORY_COLUMNS]


def save_history(df, path):
    """Persist the updated Cancelled/Chargeback lock list."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    df.to_csv(path, index=False)


# ===========================================================================
# MAIN
# ===========================================================================


def main():
    parser = argparse.ArgumentParser(
        description="Prepare a Salesforce-import-ready CSV of AT&T assets not yet in Salesforce."
    )
    parser.add_argument("--premier", help="Path to the AT&T Premier export (.xlsx).")
    parser.add_argument("--sf", help="Path to the Salesforce 'AT&T Client Assets' report (.xlsx/.csv).")
    parser.add_argument("--input-dir", default=INPUT_DIR, help="Override input folder for auto-detection.")
    parser.add_argument("--output-dir", default=OUTPUT_DIR, help="Override output folder.")
    parser.add_argument("--client", default=None, help="Client/Account name for the batch (overrides CLIENT_VALUE; optional).")
    parser.add_argument("--max-age-days", type=float, default=MAX_INPUT_AGE_DAYS,
                        help="Reject auto-detected input files older than this (safety against stale files).")
    parser.add_argument("--no-archive", action="store_true",
                        help="Do not move processed input files into input/_processed after a successful run.")
    args = parser.parse_args()

    input_dir = args.input_dir
    output_dir = args.output_dir
    client_value = args.client if args.client is not None else CLIENT_VALUE

    # Track which files were auto-detected vs. explicitly provided.
    premier_auto = args.premier is None
    sf_auto = args.sf is None
    premier_path = args.premier or autodetect(input_dir, PREMIER_GLOB, "Premier")
    sf_path = args.sf or autodetect(input_dir, SF_GLOB, "Salesforce")

    # Stale-file guard: only for auto-detected files (the unattended weekly run).
    if premier_auto:
        check_freshness(premier_path, "Premier", args.max_age_days)
    if sf_auto:
        check_freshness(sf_path, "Salesforce", args.max_age_days)

    print("=" * 70)
    print("AT&T -> Salesforce import prep")
    print("=" * 70)
    print(f"Premier export : {premier_path}")
    print(f"Salesforce rpt : {sf_path}")
    print(f"Output folder  : {output_dir}")
    print("-" * 70)

    # --- Read both files as strings, auto-detecting the header row ----------
    # Premier: require at least 3 of the 6 known fields to land on the header.
    premier, premier_hdr = read_with_header_detect(
        premier_path, list(PREMIER_ALIASES.values()), PREMIER_SKIP_ROWS,
        "Premier export", min_groups=3,
    )
    # Salesforce: a normal report; just needs the IMEI/Identifier column.
    sf, sf_hdr = read_with_header_detect(
        sf_path, [SF_IMEI_ALIASES], SF_SKIP_ROWS, "Salesforce report", min_groups=1,
    )
    print(f"Header row detected      : Premier=row {premier_hdr + 1}, Salesforce=row {sf_hdr + 1}")
    print("-" * 70)

    premier_rows_read = len(premier)

    # --- Locate IMEI columns ------------------------------------------------
    premier_imei_col = find_column(premier.columns, PREMIER_ALIASES["imei"], "the Premier export")
    sf_imei_col = find_column(sf.columns, SF_IMEI_ALIASES, "the Salesforce report")

    # --- Locate the other Premier source columns ----------------------------
    col_iccid = find_column(premier.columns, PREMIER_ALIASES["iccid"], "the Premier export")
    col_phone = find_column(premier.columns, PREMIER_ALIASES["wireless_number"], "the Premier export")
    col_univ = find_column(premier.columns, PREMIER_ALIASES["universal_id"], "the Premier export")
    col_status = find_column(premier.columns, PREMIER_ALIASES["status"], "the Premier export")
    col_contract = find_column(premier.columns, PREMIER_ALIASES["contract_type"], "the Premier export")

    # --- Normalize IMEI on both sides ---------------------------------------
    premier["_imei_norm"] = premier[premier_imei_col].map(digits_only)
    sf_imei_set = set(sf[sf_imei_col].map(digits_only)) - {""}

    # Drop blank-IMEI Premier rows.
    blank_imei = int((premier["_imei_norm"] == "").sum())
    premier = premier[premier["_imei_norm"] != ""].copy()

    # --- Anti-join: keep rows NOT already in Salesforce ---------------------
    in_sf_mask = premier["_imei_norm"].isin(sf_imei_set)
    already_in_sf = int(in_sf_mask.sum())
    not_found = premier[~in_sf_mask].copy()

    # --- Build output table -------------------------------------------------
    unmapped_owner = set()
    out = pd.DataFrame()
    out["Identifier"] = not_found["_imei_norm"]  # already digits-only
    out["ICCID"] = not_found[col_iccid].map(repair_numeric_string).map(digits_only)
    out["Phone Number"] = not_found[col_phone].map(format_phone)
    out["Universal ID"] = not_found[col_univ].map(lambda v: "" if v is None or pd.isna(v) else str(v).strip())
    out["Current Status"] = not_found[col_status].map(lambda v: "" if v is None or pd.isna(v) else str(v).strip())
    out["Asset Owner Type"] = not_found[col_contract].map(lambda v: remap_owner_type(v, unmapped_owner))
    out["Client"] = client_value

    out = out[OUTPUT_COLUMNS]

    # --- Enforce permanent Cancelled/Chargeback lock ------------------------
    # A Wireless Number that has EVER been marked Cancelled or Chargeback may
    # never be uploaded back to Active (or anything else) by a later run.
    history = load_history(HISTORY_PATH)
    locked_status_by_digits = dict(zip(history["Phone Digits"], history["Blocked Status"]))

    out["_phone_digits"] = out["Phone Number"].map(digits_only)
    out["_new_blocked_status"] = out["Current Status"].map(classify_blocked_status)
    is_locked = out["_phone_digits"].map(lambda d: d != "" and d in locked_status_by_digits)
    attempts_change = out["_new_blocked_status"].isna()  # not itself Cancelled/Chargeback this run

    blocked_mask = is_locked & attempts_change
    blocked = out[blocked_mask].copy()
    blocked["Locked Status"] = blocked["_phone_digits"].map(locked_status_by_digits)
    blocked["Attempted Status"] = blocked["Current Status"]
    blocked = blocked.drop(columns=["_phone_digits", "_new_blocked_status"])
    blocked = blocked[["Phone Number", "Locked Status", "Attempted Status"] +
                       [c for c in OUTPUT_COLUMNS if c not in ("Phone Number", "Current Status")]]

    out = out[~blocked_mask].copy()

    # Update the persistent lock list: newly-Cancelled/Chargeback numbers get
    # added (or refreshed), and numbers we just blocked get their "Last Seen"
    # bumped so the audit trail shows they're still being attempted.
    stamp_today = date.today().isoformat()
    history_by_digits = history.set_index("Phone Digits").to_dict(orient="index") if len(history) else {}

    newly_locked = out[out["_new_blocked_status"].notna()]
    for _, row in newly_locked.iterrows():
        d = row["_phone_digits"]
        if d == "":
            continue
        existing = history_by_digits.get(d)
        history_by_digits[d] = {
            "Phone Number": row["Phone Number"],
            "Blocked Status": row["_new_blocked_status"],
            "Identifier": row["Identifier"],
            "Universal ID": row["Universal ID"],
            "First Marked": existing["First Marked"] if existing else stamp_today,
            "Last Seen": stamp_today,
        }

    for d in blocked["Phone Number"].map(digits_only):
        if d in history_by_digits:
            history_by_digits[d]["Last Seen"] = stamp_today

    if history_by_digits:
        updated_history = pd.DataFrame(
            [{"Phone Digits": d, **rec} for d, rec in history_by_digits.items()]
        )[HISTORY_COLUMNS]
        save_history(updated_history, HISTORY_PATH)

    out = out.drop(columns=["_phone_digits", "_new_blocked_status"])

    # --- Validate Universal ID; quarantine failures -------------------------
    univ_re = re.compile(UNIVERSAL_ID_REGEX)
    valid_mask = out["Universal ID"].map(lambda v: bool(univ_re.match(v)) if v else False)
    clean = out[valid_mask].copy()
    review = out[~valid_mask].copy()

    # --- Write outputs ------------------------------------------------------
    os.makedirs(output_dir, exist_ok=True)
    stamp = date.today().strftime("%Y%m%d")
    clean_path = os.path.join(output_dir, f"ATT_Import_{stamp}.csv")
    review_path = os.path.join(output_dir, f"ATT_Import_{stamp}_REVIEW.csv")
    blocked_path = os.path.join(output_dir, f"ATT_Import_{stamp}_BLOCKED_HISTORY.csv")

    clean.to_csv(clean_path, index=False)
    if len(review) > 0:
        review.to_csv(review_path, index=False)
    if len(blocked) > 0:
        blocked.to_csv(blocked_path, index=False)

    # --- Summary ------------------------------------------------------------
    print(f"Rows read (Premier)      : {premier_rows_read}")
    print(f"  blank IMEI dropped     : {blank_imei}")
    print(f"Already in Salesforce    : {already_in_sf}")
    print(f"New (not found in SF)    : {len(not_found)}")
    if len(blocked) > 0:
        print(f"Blocked (Cancelled/Chargeback lock): {len(blocked)}  -> {blocked_path}")
    else:
        print(f"Blocked (Cancelled/Chargeback lock): 0")
    print(f"Clean to upload          : {len(clean)}  -> {clean_path}")
    if len(review) > 0:
        print(f"Held for review (bad UID): {len(review)}  -> {review_path}")
    else:
        print(f"Held for review (bad UID): 0")

    if unmapped_owner:
        print("-" * 70)
        print("WARNING: unmapped Asset Owner Type value(s) kept as-is:")
        for v in sorted(unmapped_owner):
            print(f"  - {v!r}")
        print("  Add a mapping to OWNER_TYPE_MAP if these should be remapped.")

    # --- Archive the processed inputs so they can't be re-used next week ----
    if ARCHIVE_PROCESSED_INPUTS and not args.no_archive:
        moved = archive_inputs([premier_path, sf_path], input_dir, stamp)
        if moved:
            print("-" * 70)
            print("Archived processed input file(s) so next week starts clean:")
            for m in moved:
                print(f"  -> {m}")

    print("=" * 70)
    print("Done. Run the Salesforce import wizard with the clean CSV above.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(130)
