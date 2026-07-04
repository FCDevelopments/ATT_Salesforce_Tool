# ATT_Salesforce_Tool

Reconciles a weekly **AT&T Premier** device/line export against a **Salesforce "Client Assets"** report and produces a clean, Salesforce-import-ready CSV containing only the assets that are genuinely new — replacing a manual, line-by-line VLOOKUP comparison.

> Non-technical quick-start guide: see [`README - START HERE.txt`](./README%20-%20START%20HERE.txt) — written for the end user who runs this weekly.

## Highlights

- **Header auto-detection** — survives the AT&T export format shifting rows between weeks
- **String-safe parsing** — IMEI/ICCID/phone values keep every digit (no scientific notation, no trailing `.0`)
- **Bad-data quarantine** — rows with an invalid Universal ID go to a separate `_REVIEW.csv` instead of failing the import
- **Permanent status lock** — any line ever marked Cancelled/Chargeback can never be re-uploaded as Active (persisted in `state/`)
- **Stale-input guard** — refuses to silently reprocess last week's files (>6 days old)
- **Self-archiving** — processed inputs move to `input/_processed/<date>/` after each run

## Setup

1. Run `Install_Requirements.bat` once (installs Python + pandas/openpyxl if missing).
2. Open `att_salesforce_prep.py` and fill the `TODO <<FILL>>` config values (`CLIENT_VALUE`, `UNIVERSAL_ID_REGEX`, column aliases).
3. Run `Setup_Weekly_Schedule.bat` to schedule the Monday 8:30 AM run — or `Run_Now.bat` to run on demand.

## Weekly workflow

Drop the two downloaded reports into `input/`:

| Report | Source | Filename pattern |
|---|---|---|
| Device export | AT&T Premier → Inventory | `detail_report*.xlsx` |
| Asset report | Salesforce → Reports | `*ATT*Client*Assets*` |

Outputs land in `output/`: `ATT_Import_<date>.csv` (upload this), plus `_REVIEW.csv` / `_BLOCKED_HISTORY.csv` when rows need human attention.

## Stack

Python · pandas · openpyxl · Windows Task Scheduler (`.bat` launchers, no admin required)
