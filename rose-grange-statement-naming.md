# CBA Statement Naming (all facilities)

Skill name: `rose-grange-statement-naming`
Description: Use whenever a new PDF (CBA interest rate statement, bill/rollover notice, line fee invoice, repayment, break cost, or similar) is saved into the "Interest Rate Statement Automation(s)" folder, for ANY facility (Rose Grange, Leakes Commercial, etc.). Reads the document, identifies the borrowing entity/facility, and renames it to the standard convention so statements can be told apart and sorted chronologically.

## When to use
Trigger whenever a PDF is added to (or found unrenamed in) the folder "Interest Rate Statement Automation(s)" — this holds CBA facility documents for multiple borrowing entities/trusts. Known facilities include:
- **Rose Grange** — Rose Grange Nominees Pty Ltd as trustee for The Rose Grange Commercial Trust (deal ref "ROSE GRANGE COM BILAT AUD66.8M 15DEC17", bank ref AUSNA1000072; the property is sometimes called "Tarneit Central" — same facility, use "Rose Grange" in filenames).
- **Leakes** — Leakes Commercial (trust/entity). Use "Leakes" in filenames.
- Any other facility: derive the name as described below.

A file already matching the "CBA ... - YYYY-MM-DD - [Facility]..." pattern does not need touching.

## Process
1. List the folder and identify any PDF whose name does NOT already start with "CBA ".
2. Read each such PDF (short 1-2 page CBA notice).
3. Extract:
   - The document date (the "Date:" field near the top of the letter).
   - The **facility short name**: take it from the document itself (customer/borrower name, "The ___ Commercial Trust", deal reference) and, if helpful, the original filename. Reduce to the distinctive short name in Title Case, dropping boilerplate such as "Nominees Pty Ltd", "as trustee for", "The", "Commercial Trust", "Pty Ltd" (e.g. "The Leakes Commercial Trust" -> "Leakes"; "Rose Grange Commercial Trust" -> "Rose Grange"). If the entity is new/unclear or two entities could share a short name, flag it and confirm with the user rather than guessing. Keep the same short name consistently for a given facility.
   - The document type, from its own heading/description:
     - "*** ROLLOVER RATE SET NOTICE ***" -> Rollover Notice (or Drawdown Notice if it's a NEW loan being drawn down rather than an existing one repricing — check whether the amount is new money or a reprice of an existing "Current Loan").
     - "*** FEE INVOICE ***" with "Line Fee in Arrears" -> Line Fee Invoice.
     - "*** FEE INVOICE ***" with "Break Cost" -> Break Cost Notice.
     - A repayment/payment notice (principal + interest being debited, loan amount decreasing) -> Loan Repayment Notice.
     - "Rollover Intent Notice" or similar preliminary/pre-notice -> Rollover Intent Notice.
     - Anything else -> use the document's own stated title, Title Case, as the type.
4. Rename the file to:
   `CBA [Type] - YYYY-MM-DD - [Facility].pdf`
   - `[Type]` is one of: Rollover Notice, Line Fee Invoice, Loan Repayment Notice, Break Cost Notice, Rollover Intent Notice, Drawdown Notice, or another short Title-Case type taken from the document itself.
   - `YYYY-MM-DD` is the document's own "Date:" field (not an effective date, due date, or filename hint), zero-padded.
   - `[Facility]` is the short name from step 3 (e.g. Rose Grange, Leakes). Always end with " - [Facility].pdf".
5. Special cases:
   - **Side tranches / amounts that aren't the main facility** (e.g. a temporary $10m drawdown or repayment running alongside the main balance): append a short parenthetical, e.g. `CBA Drawdown Notice - 2023-06-27 - Rose Grange (10m tranche).pdf`. For a principal repayment that reduced the main facility itself, append `(10m principal repayment)` or similar describing what happened.
   - **Duplicates**: if two files have identical content and the same document date, keep both but append ` (duplicate)` before `.pdf` on the second — never delete a file.
   - If a filename collision would occur for a genuinely distinct document (e.g. two different Line Fee Invoices on the same date), disambiguate with a short suffix describing the difference (e.g. the period covered).
6. After renaming, briefly summarise what was renamed (old name -> new name), grouped by facility, and flag anything ambiguous (unclear facility, document type that doesn't fit, unreadable date) rather than guessing silently.

## Do not
- Do not rename workbooks (e.g. "Loan Schedule - FY27 Statement Verification.xlsx") — this skill only applies to source statement PDFs.
- Do not delete or overwrite files.
- Do not invent a date if the PDF's own "Date:" field is missing or unreadable — ask the user.
- Do not label a non-Rose-Grange document "Rose Grange".

## Context
Renamed files feed per-facility verification workbooks (Statement Checks / Statement Processing / Statement Month-End Schedule). Consistent naming keeps "Source Document" references meaningful. After renaming, offer to run the matching workbook update (currently `rose-grange-fy27-ledger-update` exists for Rose Grange only; other facilities may need their own workbook set up first).
