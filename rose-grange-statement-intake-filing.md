# Statement Intake & Filing (all facilities)

Skill name: `rose-grange-statement-intake-filing`
Description: Use whenever a new PDF (CBA rollover notice, line fee invoice, repayment, break cost, drawdown, etc.) is saved into the "Interest Rate Statement Automations" folder, for any facility (Rose Grange, Leakes Commercial, etc.). Names it using the rose-grange-statement-naming skill (now covers all facilities), then files it into Loan Schedules/[Facility]/FYxx by Australian financial year.

Loops the `rose-grange-statement-naming` skill (now general for all facilities) with filing into the financial-year folder structure.

## Folder structure
Root: "Interest Rate Statement Automations" (S:\Corporate Accounts\1 FINANCE 2027\Monthly Financial Reporting\Interest Rate Statement Automations). If not mounted, request access with the directory tool (note: UNC paths can't be mounted; use a mapped drive letter or a local copy).

```
Interest Rate Statement Automations/
└── Loan Schedules/              (holds all facilities)
    ├── Rose Grange/             (facility folder: holds the Excel workbook)
    │   ├── Rose Grange - Loan Schedule - FY27 Statement Verification.xlsx
    │   ├── FY23/  FY24/  FY25/  FY26/  FY27/ ...   (statement PDFs)
    └── Leakes/                  (same layout; create if missing)
        └── FYxx/
```
Each facility gets its own sibling folder under "Loan Schedules" named with the same short facility name used in filenames.

## Trigger
A new PDF appears (or is found unfiled) in the root folder, or loose in a facility folder under "Loan Schedules". Any PDF not already inside an FYxx folder is "new".

## Process
1. **Name** each new PDF by following the `rose-grange-statement-naming` skill exactly (load via the Skill tool): read the PDF, identify the facility short name, document "Date:" and type, rename to `CBA [Type] - YYYY-MM-DD - [Facility][ (qualifier)].pdf`. Apply its special cases. Never delete or overwrite.
2. **Determine FY** from the document date in the new filename (Australian FY, 1 July - 30 June): month 7-12 -> FY = year+1; month 1-6 -> FY = year. Label `FY` + last two digits (2022-09-26 -> FY23; 2026-09-14 -> FY27; 2026-06-23 -> FY26).
3. **File**: create `Loan Schedules/[Facility]/FYxx` if missing, and move the renamed PDF into it. If a same-named file exists, keep both using the naming skill's duplicate/collision rules.
4. **Do not** move or rename .xlsx workbooks; they live in `Loan Schedules/[Facility]/`.
5. **Report**: list old name -> new name -> destination folder, grouped by facility. Flag ambiguities (unclear facility, unclear type, unreadable date) rather than guessing; ask if the date is missing.
6. **Next step**: for Rose Grange, offer `rose-grange-fy27-ledger-update`. For other facilities (e.g. Leakes), note no ledger workbook skill exists yet and offer to set one up.

## Notes
- Folders are organised by facility, then FY only — not by document type.
- Re-sorting existing files: apply steps 2-3 to PDFs already named correctly (skip step 1).
