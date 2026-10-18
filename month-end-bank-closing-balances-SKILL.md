---
name: "month-end-bank-closing-balances"
description: "Pull month-end closing bank balances from the corporate bank statement CSVs (ANZ, BoM, NAB, CBA, BoQ) and write them straight into the \"Checks - Actuals\" tab of the DFC Cash Flow and Funding - Working Model. Use when the user asks to run month-end balances, update the cash flow model, pull closing bank balances, update the cash position, or do the monthly bank reconciliation/cash flow prep."
---

## Purpose

Each month, Dennis Family Corporation saves updated bank statement exports (CSV) into a "Cashflow Automations" folder, organised into subfolders per bank: "ANZ Bank Statements", "BoM Bank Statements", "NAB Bank Statements", "CBA Bank Statements", "BoQ Bank Statements". That same folder holds the live model file `DFC Cash Flow and Funding - Working Model.xlsx`. This skill extracts the closing balance for every known bank account and writes it directly into that model's "Checks - Actuals" tab — it does NOT produce a separate standalone summary file; the model itself is the deliverable.

Known folder location: `S:\Corporate Accounts\1 FINANCE 2027\Monthly Financial Reporting\Cashflow Automations`. If this folder isn't already connected, request access to it (ask the user to add it via the folder picker if it's on a network drive). If the file or folder has moved, ask rather than guessing.

## Step 0 — Check the file isn't open

Before writing, check for a lock file next to the workbook: `~$DFC Cash Flow and Funding - Working Model.xlsx`. If it exists, the file is currently open in Excel by someone — do NOT write to it. Tell the user who (the lock file's first line is usually the editor's name) and ask them to close it first. Writing while it's open risks the user's own unsaved changes overwriting yours, or yours being silently lost.

## Step 1 — Understand the target sheet layout

Open `DFC Cash Flow and Funding - Working Model.xlsx`, sheet "Checks - Actuals". Layout (confirmed structure — verify it still matches before writing, since the model can evolve):
- Column B = Bank, Column C = Account Name, Column D = Account Number (as text, spaced format e.g. `452 193 991`), for a fixed block of account rows (currently rows 35–67, one row per known account, in a fixed order — see the canonical table in Step 4). Row 68 is a blank buffer row. Row 69 is `TOTAL` = `=SUM(E35:E68)` per column — never edit row 69 or insert/delete rows in this block, since that range reference and everything else in this 40+ tab workbook that references these rows would break.
- Row 34 (and row 3, which row 34 mirrors via `=E3` style formulas) holds month-end dates across columns E, F, G... one column per month, running from Jul-25 onward for many years. These are already populated with dates whether or not the data rows below have values yet — find the correct column for a given month by matching row 34's date, don't assume a fixed column number.
- Only ever write to cells in the intersection of an account's row (35–67) and a month's column (E onward) — the actual closing balance value, as a plain number (not a formula, matching the sheet's existing convention). Leave every other cell, formula, and sheet in the workbook completely untouched.

If the sheet layout has changed materially from this description (e.g. columns shifted, rows renumbered), stop and confirm with the user rather than guessing.

## Step 2 — Determine which month(s) to update

Scan every statement CSV to find the latest transaction/statement date available (see Step 3 for per-bank parsing). Determine the target month(s):
- If the user asked for a specific month, use that.
- Otherwise, update every month-end column from the earliest one that's still blank across the account rows through to the latest month-end for which statement data exists (this sheet may start out fully blank and need a historical backfill, or may already have prior months filled and just need the newest month added — handle both the same way: compute and write every month-end that has statement coverage and isn't already correct).
- Always tell the user which month(s) were written.

## Step 3 — Parse each bank's CSVs (deterministic, not eyeballed)

These are real financial balances — write and run a Python script (via bash) to parse them exactly; never hand-total or eyeball figures from a CSV preview. Confirmed layouts:

**BoM** ("BoM Bank Statements"): one file per account, columns `Account Number, Value Date, Transaction Description, Cheque Serial Number, Account Currency, Amount, Amount Indicator, Balance, Balance Indicator, BSB, Trancode`. Dates `d/m/Y`. Strip quotes/spaces/commas from numbers. If `Balance Indicator` is `DR`, the balance is negative. Some accounts are split across two files named `... 1_2.csv` / `... 2_2.csv` (or similar) covering different date ranges of the same account — combine files sharing the same base name before picking a balance.

**ANZ** ("ANZ Bank Statements"): usually one combined file with ALL accounts stacked, columns include `Account Number, Account Name, Statement Date, Closing Ledger Balance`. Dates like `1-Jul-25` (`%d-%b-%y`). Group rows by `Account Number` before picking a balance. Numbers are quoted with commas. Account Number format is `BSB-accountnumber` (e.g. `013006-836879334`) — use only the digits after the hyphen.

**NAB** ("NAB Bank Statements"): one file per account with a metadata header — line 1 `Account Name:, <name>`, line 2 `Account Number:, <number>`, then a transaction table starting at the row beginning `Date,Narrative,...,Running Balance`. Dates `d/m/Y`. Don't assume row order (can be ascending or descending) — compare dates directly.

**CBA** ("CBA Bank Statements"): one file per account, no account identifiers in-file — identify from the filename (e.g. `FY26 CBA Leakes Acc 1247.csv` → last digits `1247`). Columns: `Process date, Description, Currency Code, Debit, Credit, Balance`. Dates `d/m/Y`.

**BoQ** ("BoQ Bank Statements"): one file per account, no account identifiers in-file — identify from filename like CBA. Columns: `Date, Description, Debit, Credit, Balance`, with `$` and comma-formatted numbers — strip both. Dates `d/m/Y`.

For every account and every target month-end date, select the row with the latest date ≤ that month-end (regardless of file order), and take its balance. This automatically satisfies "roll the last known balance forward" — if an account had no transactions in a given month, the closing balance is simply whatever it was at the last transaction before that month-end, carried into that month's column unchanged. If an account has literally no statement rows at all on or before a target date (no file, or file starts later), do NOT invent a number — leave that cell as-is and flag it to the user as "no statement data available."

## Step 4 — Match extracted accounts to the sheet's rows

Match by the **last 4 digits** of the account number (reliable — confirmed unique across all known accounts) against column D in rows 35–67. Canonical reference (bank, name, full account number, in row order as they currently appear in the sheet — reconfirm row numbers against column B/C/D before writing, since rows could shift if the sheet is edited):

| Last 4 | Bank | Account Name | Full Account Number |
|---|---|---|---|
| 3991 | BoM | DF Real Estate Management (QLD) | 452 193 991 |
| 0813 | BoM | Masco | 555 270 813 |
| 0739 | ANZ | Masco P/L | 473 170 739 |
| 1966 | ANZ | Masco P/L (Bpay) | 838 491 966 |
| 8090 | BoM | Masco Family Office | 417 238 090 |
| 7874 | BoM | Dennis Family Homes | 472 237 874 |
| 1101 | BoM | Masco | 553 201 101 |
| 3441 | BoM | DFC (Services) Pty Ltd | 463 313 441 |
| 7286 | BoM | DFC (Proj Management) | 498 557 286 |
| 0862 | BoM | Dennis Family AFSL | 554 430 862 |
| 5672 | BoM | Burnside Commercial | 416 925 672 |
| 7040 | BoM | Burnside Properties | 459 917 040 |
| 0526 | NAB | Millwood Farm Services | 629 270 526 |
| 2851 | NAB | Hunt Club Commercial | 730 042 851 |
| 3792 | NAB | HCV Commdev | 928 743 792 |
| 7754 | NAB | Manor Central Nominees | 756 847 754 |
| 8814 | CBA | Ranfurlie Property Management | 14 218 814 |
| 1385 | CBA | Rose Grange | 12 951 385 |
| 8180 | CBA | TC Commdev | 13 148 180 |
| 1247 | CBA | Leakes | 14 301 247 |
| 6862 | ANZ | Westbrook Project | 836 796 862 |
| 6889 | ANZ | Westbrook Proceeds | 836 796 889 |
| 9279 | ANZ | Westbrook North Project | 837 479 279 |
| 9299 | ANZ | Westbrook North Proceeds | 838 599 299 |
| 2624 | ANZ | Peppercorn Project | 838 192 624 |
| 6268 | ANZ | Peppercorn Proceeds | 838 286 268 |
| 9334 | ANZ | Ashbury Project | 836 879 334 |
| 9253 | ANZ | Ashbury Proceeds | 838 299 253 |
| 8815 | ANZ | Cora Proceeds | 838 518 815 |
| 8807 | ANZ | Cora Project | 838 518 807 |
| 0987 | ANZ | Manor Commercial | 838 270 987 |
| 9037 | BoQ | Donnybrae Transaction | 22 839 037 |
| 2317 | BoQ | Donnybrae Settlement | 23 032 317 |

If a statement file's account doesn't match any row in the sheet (a genuinely new account), do NOT insert a new row (this would break the `TOTAL` formula range and any other references into this block) — flag it to the user by name/account number so they can decide whether and how to add it manually.

## Step 5 — Write the values

Using openpyxl, open the workbook, write each computed closing balance (plain number, matching the sheet's existing number formatting) into the correct row/column intersection, and save. Do not run a full-workbook recalculation (e.g. LibreOffice `recalc.py`) on this file — it has thousands of formulas linking to other external files, which a headless recalculation cannot resolve and would permanently break. Saving with openpyxl alone preserves formula text everywhere else untouched; the user needs to open the file in Excel afterward and let it recalculate/update links normally to restore cached values throughout the rest of the workbook (this is expected and the user has already accepted this trade-off — don't re-litigate it, just remind them once in the summary).

## Step 6 — Report back

Tell the user: which month(s) were written, how many accounts were updated, any accounts with no statement data this run (cell left unchanged), and any unmatched/new accounts found in statements but not present in the sheet. Remind them to open the file in Excel and let it recalculate/refresh links before trusting other figures elsewhere in the workbook.

## Notes for future runs

- If a bank changes its export format, the parsing rule for that bank needs updating — don't force new data into old column assumptions; flag it if a file doesn't match the documented layout.
- If the user confirms a new/unmatched account should be tracked permanently, they'll need to add a row to the sheet themselves (and update the `TOTAL` sum range) before this skill can populate it — then add it to the canonical table above so future runs recognise it.
- If the sheet layout, file name, or folder location changes, update this skill accordingly rather than silently adapting on the fly.
