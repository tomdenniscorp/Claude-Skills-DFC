---
name: "rose-grange-statement-parse"
description: "Manually triggered monthly/quarterly run for the Rose Grange / Tarneit Central CBA facility. Parses new CBA Rollover Notices (monthly) and Line Fee Invoices (quarterly), checks them, and populates 'Rose Grange - Loan Schedule - FY27 Statement Verification.xlsx' (Statement Checks, Statement Processing, Statement Month-End Schedule). Use when the user says 'run the statement parse', 'process the rollover notice', 'update the loan schedule', 'add the line fee invoice', or similar. Replaces rose-grange-fy27-ledger-update."
---

# Rose Grange Statement Parse → Loan Schedule

Run on demand. The typical runs are:
- **Monthly**, after the CBA Rollover Rate Set Notice arrives (~23rd).
- **Quarterly**, after the Line Fee Invoice arrives (mid-Mar, Jun, Sep, Dec).

Everything is deterministic and done by two scripts. Your job is to run them, read the report, and stop if anything needs a decision.

## Locations
- Root folder: `S:\Corporate Accounts\1 FINANCE 2027\Monthly Financial Reporting\Interest Rate Statement Automations`. If it isn't mounted, ask the user to connect it with the folder picker (network drive).
- Workbook: `Loan Schedules/Rose Grange/Rose Grange - Loan Schedule - FY27 Statement Verification.xlsx`.
- PDFs: `Loan Schedules/Rose Grange/FYxx/`.
- Scripts: `Skills/scripts/rg_parse.py` (PDF extraction plus document-level checks) and `Skills/scripts/rg_ledger.py` (ingest, rebuild and verify).
- Backups: every write first copies the workbook to `Loan Schedules/Rose Grange/_backups/`.
- Recalc: use `scripts/recalc.py` from the **xlsx** skill folder. Find it with `find /sessions -path "*skills/xlsx/scripts/recalc.py" 2>/dev/null | head -1`.

In bash, the root appears under `/sessions/<id>/mnt/Interest Rate Statement Automations`. Set:
```bash
ROOT=$(find /sessions -maxdepth 3 -type d -name "Interest Rate Statement Automations" 2>/dev/null | head -1)
RG="$ROOT/Loan Schedules/Rose Grange"; WB="$RG/Rose Grange - Loan Schedule - FY27 Statement Verification.xlsx"
S="$ROOT/Skills/scripts"; RECALC=$(find /sessions -path "*skills/xlsx/scripts/recalc.py" 2>/dev/null | head -1)
```

## How the workbook works (read once)
- **Statement Checks** is the single source of truth: one row per bank document, with extracted figures in blue. Every run reads this tab, merges in the new PDFs, and **re-renders all three tabs**. Formula ranges are therefore always exact, and there are no buffer rows to manage.
- Preserved across runs: the "Notes (manual)" section, column P "Comment", and any custom Continuity text written as a `="text"` formula in column L.
- Regenerated each run: Summary, Automated flags, Processing, and Month-End.
- **Statement Processing** is the continuous bill-period ledger.
  - Implied rows (amber) are inserted automatically where a rollover notice is missing. Their rate is back-solved from the next notice's payment confirmation.
  - "Interest – realised" equals the payment the next notice confirms. This only differs from "stated" when the balance changed mid-period.
  - Side tranches sit in their own block.
  - The facility limit is an input in `E3`.
- **Statement Month-End Schedule** gives the calendar-month accrual and cash split for journals.
  - Interest accrual prorates realised interest by day overlap.
  - Line fee accrues daily on the facility limit at 1.40%.
  - Cash columns use bill end dates (interest) and the invoice debit date in column N (line fee).
  - Reconciliation rows at the bottom must all show PASS.
  - Formulas are plain `SUMPRODUCT`/`INDEX`/`MATCH`. **Never** use `LET`, `XLOOKUP`, or `IF` inside `SUMPRODUCT`, because they mis-evaluate outside dynamic-array Excel.
- Orange rows are "Legacy schedule – PDF pending". They were backfilled from the old Loan Schedule. When the real PDF for that period is ingested, it replaces the legacy row automatically and is compared against it.

## Process

1. **Intake.** If there are new or unnamed PDFs in the root folder or loose in `Loan Schedules/Rose Grange`, run `rose-grange-statement-intake-filing` first. It names and files them into FYxx. If the user just names the file(s), use those.
2. **Find what's new.** Compare the PDFs in the current FY folder (and the previous one near year-end) against the workbook's Source Document column. Alternatively, pass every PDF in the FY folder: anything already ingested is skipped as "already in workbook", and `(duplicate)` copies are ignored.
3. **Dry run:**
   ```bash
   python3 "$S/rg_ledger.py" ingest "$WB" "<pdf1>" ["<pdf2>" ...] --dry-run
   ```
   Read the JSON report: `status`, `added`, `replaced_legacy`, `skipped`, `checks`, `blocking`, `warnings`.
   - `NO CHANGE`: tell the user nothing new was found and stop.
   - `ESCALATE` (exit code 2): **do not write.** Show the user the `blocking` items in plain English and follow "When it escalates" below.
   - `ROUTINE`: continue.
4. **Write:** run the same command without `--dry-run`. It backs up the workbook first and then re-renders it.
5. **Recalculate.** LibreOffice needs a local writable copy, so:
   ```bash
   cp "$WB" /tmp/rg.xlsx && python3 "$RECALC" /tmp/rg.xlsx 90 && cp /tmp/rg.xlsx "$WB"
   ```
   The recalc must return `total_errors: 0`.
6. **Verify:**
   ```bash
   python3 "$S/rg_ledger.py" verify "$WB"
   ```
   It must return `"status": "PASS"`. The verify step:
   - scans every cell for formula errors;
   - checks the ledger is continuous;
   - recomputes every Month-End row in Python (interest accrual, interest cash, line fee accrual, line fee cash, days captured) and compares it to the sheet to the cent;
   - confirms the reconciliation rows and that there are no Calc or Continuity FAILs.

   If verify fails, restore the latest `_backups` copy and report the failure. Don't hand-patch formulas.
7. **Report back briefly.** For a routine run, use 3–5 lines:
   - what was added (period, rate, amount);
   - the check results (for example, "calc PASS, continuity PASS – 23-Oct notice confirms Sep–Oct interest of $248,303.84");
   - for the month(s) now complete, the Month-End **Interest accrual, Line fee accrual, Total accrual, Interest paid and Line fee paid**, ready for the journal;
   - any `partial_months` or warnings.

   Then present the workbook file.

## Checks performed

**Document level (`rg_parse.py`):**
- The facility deal ref and bank ref are correct.
- The date and document type were read.
- All fields were extracted.
- There is a single loan line.
- Balance × rate × days / 365 equals the stated amount (within $1).
- The Interest Due Date equals the period end. Where the bank's Next End Date and Interest Due Date disagree, the parser uses whichever one reconciles to the stated interest and raises a WARN.
- The facility balance is unchanged.
- The repricing end date equals the effective date.
- The margin is 0%.
- For line fees: the calculation lines are contiguous (CBA splits at 31-Dec, and the parser sums the lines), days are inclusive, the fee recalculates, total = sum of lines, the rate is 1.40%, the base is the $66.8m limit, and the debit date is 1–10 days after the period end.

**Workbook level (`rg_ledger.py ingest`):**
- Duplicate or already-ingested files are skipped.
- Continuity: the interest paid on the new notice equals the prior period's Interest Due (within $1).
- A gap between periods is a WARN, and an implied period is inserted automatically.
- Rate movement over 75bp is a WARN.
- A balance change vs the prior period is a FAIL.
- A conflict with an existing PDF row for the same period is a FAIL.
- When a PDF replaces a legacy row, it is compared against that row.
- Line fee continuity is checked (each invoice starts the day after the prior one ended).

**Post-write:** recalc returns 0 errors, and verify passes (see step 6).

## When it escalates (ask the user, don't guess)
- **Continuity FAIL (payment ≠ prior interest due):** compare the notice with the prior period. If a mid-period repayment or increase explains it, treat it as a balance change (below). Otherwise it's a bank or data error, so flag it to the user.
- **Balance change (`Outstandings Increase` / `Decrease`, repayment):**
  - First, record any repayment or break-cost PDFs with `--force`. This puts them under Other Notices only.
  - Then ingest the rollover notice with `--force`.
  - Then, in Statement Checks, overwrite the affected row's K (Prior-Period Payment Confirmed) with the total realised interest across all the documents. The Mar-2023 precedent is $147,501.98 + $11,129.42 stub = $158,631.40.
  - Set its L to `="See note – confirmed via two documents due to a mid-period principal repayment"` and add a manual Note.
  - Run `rg_ledger.py rebuild "$WB"`, then recalc and verify.
- **Drawdown of a separate amount (side tranche):** add a row under "Side Tranches" in Statement Checks (same columns, L as a `="Confirmed by …"` text), record the repayment under Other Notices, then rebuild, recalc and verify.
- **Unknown document type:** show the user the first lines of the PDF and ask how to treat it.
- **Backlog of several months:** pass all the PDFs in one ingest call. They are processed in date order, and continuity is checked pair by pair.

## Manual corrections
Edit **Statement Checks** directly. Use only the blue input cells, Notes, column P Comment, or the L text override. Then run `rg_ledger.py rebuild "$WB"`, recalc and verify. Never edit Processing or Month-End by hand, because they are overwritten on every run.

## Don'ts
- Don't write when the status is `ESCALATE` unless the user agrees to `--force`.
- Don't delete or rename PDFs.
- Don't add tabs between the three ledger tabs. Extra tabs placed after them are preserved.
- Don't change the line-fee rate or facility limit unless a new invoice or facility letter says so. They are inputs in Processing `E3` and the per-row blue rate cells.
