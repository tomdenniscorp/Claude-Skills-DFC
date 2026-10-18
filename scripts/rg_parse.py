#!/usr/bin/env python3
"""
rg_parse.py - Extract figures from CBA facility notices (Rose Grange / Tarneit Central).

Handles:
  rollover      *** ROLLOVER RATE SET NOTICE ***
  line_fee      *** FEE INVOICE *** + "Line Fee in Arrears"
  break_cost    *** FEE INVOICE *** + "Break Cost"
  repayment     *** PAYMENT NOTICE ***
  drawdown      *** DRAWDOWN RATE SET NOTICE ***
  intent        *** ROLLOVER INTENT NOTICE ***   (preliminary - info only)
  unknown       anything else

Usage:
  python3 rg_parse.py <pdf> [<pdf> ...]          -> JSON list to stdout
Each record carries "parse_checks": a list of {check, status, detail} produced
from the document alone (no workbook needed). status in PASS / WARN / FAIL.
"""
import json
import re
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

FACILITY_TOKEN = "ROSE GRANGE COM BILAT AUD66.8M"
BANK_REF = "AUSNA1000072"
DAY_BASIS = 365
TOL = 1.00  # $ tolerance for recalculation checks

D = r"(\d{1,2}-[A-Za-z]{3}-\d{4})"
AMT = r"([\d,]+\.\d{2})"


def _d(s):
    return datetime.strptime(s.strip(), "%d-%b-%Y").date() if s else None


def _a(s):
    return float(s.replace(",", "")) if s else None


def pdf_text(path):
    try:
        out = subprocess.run(["pdftotext", "-layout", str(path), "-"],
                             capture_output=True, text=True, check=True).stdout
        if out.strip():
            return out
    except Exception:
        pass
    import pdfplumber  # fallback
    with pdfplumber.open(str(path)) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)


def _section(text, start_pat, end_pats):
    m = re.search(start_pat, text)
    if not m:
        return ""
    rest = text[m.end():]
    cut = len(rest)
    for p in end_pats:
        e = re.search(p, rest)
        if e:
            cut = min(cut, e.start())
    return rest[:cut]


def _chk(lst, name, ok, detail, warn_only=False):
    lst.append({"check": name,
                "status": "PASS" if ok else ("WARN" if warn_only else "FAIL"),
                "detail": detail})


def classify(text):
    t = text.upper()
    title = re.search(r"\*\*\*\s*(.+?)\s*\*\*\*", text)
    title = title.group(1).strip().upper() if title else ""
    if "ROLLOVER INTENT" in title:
        return "intent"
    if "ROLLOVER RATE SET" in title:
        return "rollover"
    if "DRAWDOWN RATE SET" in title:
        return "drawdown"
    if "PAYMENT NOTICE" in title:
        return "repayment"
    if "FEE INVOICE" in title:
        if "LINE FEE" in t:
            return "line_fee"
        if "BREAK COST" in t:
            return "break_cost"
    return "unknown"


def parse_rollover(text, rec, checks):
    eff = re.search(r"Effective Date:\s*" + D, text)
    rec["effective_date"] = _d(eff.group(1)) if eff else None
    rec["description"] = (re.search(r"Description:\s*(.+)", text) or [None, None])[1]

    rep = _section(text, r"LOANS repricing:", [r"Borrower ", r"\n\s*\n\s*\n"])
    rep_lines = re.findall(r"^\s*(.+?)\s{2,}" + AMT + r"\s+" + D + r"\s*$", rep, re.M)
    rec["repricing_loans"] = [{"desc": d.strip(), "amount": _a(a), "current_end": _d(e)} for d, a, e in rep_lines]

    inc = re.search(r"Outstandings Increase\s+" + AMT, text)
    dec = re.search(r"Outstandings Decrease\s+" + AMT, text)
    rec["outstandings_increase"] = _a(inc.group(1)) if inc else 0.0
    rec["outstandings_decrease"] = _a(dec.group(1)) if dec else 0.0

    ip = _section(text, r"INTEREST payments:", [r"Borrower ", r"Calculated as follows", r"We will"])
    ip_lines = re.findall(r"^\s*(BBSY.*?|[A-Za-z].*?)\s{2,}" + AMT + r"\s*$", ip, re.M)
    ip_lines = [(d, a) for d, a in ip_lines if not d.strip().lower().startswith("description")]
    rec["interest_payment"] = round(sum(_a(a) for _, a in ip_lines), 2) if ip_lines else None

    nl = _section(text, r"new LOANS:", [r"Borrower ", r"Commonwealth Bank of Australia\s+Page"])
    nl_lines = re.findall(r"^\s*(.+?)\s{2,}" + AMT + r"\s+" + D + r"\s*$", nl, re.M)
    rec["new_loans"] = [{"desc": d.strip(), "amount": _a(a), "next_end": _d(e)} for d, a, e in nl_lines]

    rs = _section(text, r"rates have\s+been set as follows:", [r"Margin:"])
    rs_lines = re.findall(r"^\s*(.+?)\s{2,}" + AMT + r"\s+([\d.]+)%\s+" + D + r"\s*$", rs, re.M)
    rec["rate_lines"] = [{"desc": d.strip(), "amount": _a(a), "rate": float(r) / 100, "next_end": _d(e)}
                         for d, a, r, e in rs_lines]
    margin = re.search(r"Margin:\s*([\d.]+)%", text)
    rec["margin"] = float(margin.group(1)) / 100 if margin else None
    due = re.search(r"Interest Due:\s*AUD\s*" + AMT, text)
    due_d = re.search(r"Interest Due Date:\s*" + D, text)
    rec["interest_due"] = _a(due.group(1)) if due else None
    rec["interest_due_date"] = _d(due_d.group(1)) if due_d else None

    # Normalised ledger fields
    if rec["rate_lines"]:
        rl = rec["rate_lines"][0]
        rec["period_start"] = rec["effective_date"]
        rec["period_end"] = rl["next_end"]
        rec["loan_balance"] = rl["amount"]
        rec["rate"] = rl["rate"]
    rec["stated_amount"] = rec["interest_due"]
    rec["prior_payment_confirmed"] = rec["interest_payment"]

    # --- document-level checks ---
    req = ["effective_date", "period_end", "loan_balance", "rate", "interest_due", "interest_payment"]
    missing = [k for k in req if rec.get(k) is None]
    _chk(checks, "All required fields extracted", not missing, "missing: " + ", ".join(missing) if missing else "ok")
    if missing:
        return
    _chk(checks, "Single loan line (no split tranches)",
         len(rec["rate_lines"]) == 1 and len(rec["new_loans"]) == 1 and len(rec["repricing_loans"]) == 1,
         f"{len(rec['repricing_loans'])} repricing / {len(rec['new_loans'])} new / {len(rec['rate_lines'])} rate lines")
    def _recalc(end):
        dd = (end - rec["period_start"]).days
        return dd, rec["loan_balance"] * rec["rate"] * dd / DAY_BASIS
    days, recalc = _recalc(rec["period_end"])
    if (abs(recalc - rec["interest_due"]) > TOL and rec["interest_due_date"]
            and rec["interest_due_date"] != rec["period_end"]):
        d2, r2 = _recalc(rec["interest_due_date"])
        if abs(r2 - rec["interest_due"]) <= TOL:
            _chk(checks, "Bank date inconsistency resolved", False,
                 f"Next End Date {rec['period_end']} but Interest Due Date {rec['interest_due_date']}; "
                 f"stated interest only reconciles to the Interest Due Date, so that is used as Period End",
                 warn_only=True)
            rec["period_end"], days, recalc = rec["interest_due_date"], d2, r2
    rec["days"] = days
    rec["recalculated_amount"] = round(recalc, 2)
    _chk(checks, "Interest Due recalculates (bal x rate x days / 365)",
         abs(recalc - rec["interest_due"]) <= TOL,
         f"stated {rec['interest_due']:,.2f} vs recalc {recalc:,.2f} ({days} days)")
    _chk(checks, "Interest Due Date = Period End",
         rec["interest_due_date"] == rec["period_end"], f"{rec['interest_due_date']} vs {rec['period_end']}",
         warn_only=True)
    old_bal = rec["repricing_loans"][0]["amount"] if rec["repricing_loans"] else None
    rec["prior_balance"] = old_bal
    rec["balance_change"] = round((rec["loan_balance"] or 0) - (old_bal or 0), 2)
    _chk(checks, "Facility balance unchanged (repricing amount = new loan amount)",
         rec["balance_change"] == 0,
         f"repricing {old_bal:,.2f} -> new {rec['loan_balance']:,.2f}"
         + (f"; outstandings increase {rec['outstandings_increase']:,.2f}" if rec["outstandings_increase"] else "")
         + (f"; outstandings decrease {rec['outstandings_decrease']:,.2f}" if rec["outstandings_decrease"] else ""))
    if rec["repricing_loans"]:
        _chk(checks, "Repricing loan's current end = effective date",
             rec["repricing_loans"][0]["current_end"] == rec["effective_date"],
             f"{rec['repricing_loans'][0]['current_end']} vs {rec['effective_date']}")
    _chk(checks, "Margin is 0.00% (all-in rate = BBSY bid)", (rec["margin"] or 0) == 0,
         f"margin {rec['margin']}", warn_only=True)


def parse_line_fee(text, rec, checks):
    # An invoice can hold several calculation lines (CBA splits at 31-Dec) - sum them.
    lines = re.findall(D + r"\s+" + D + r"\s+(\d+)\s+AUD\s+" + AMT + r"\s+([\d.]+)%\s+" + AMT, text)
    eff = re.search(r"Effective\s+" + D, text)
    tot = re.search(r"Totalling\s+AUD\s+" + AMT, text, re.I)
    rec["effective_date"] = _d(eff.group(1)) if eff else None
    rec["total"] = _a(tot.group(1)) if tot else None
    rec["calc_lines"] = [{"start": _d(a), "end": _d(b), "days": int(c), "base": _a(d), "rate": float(r) / 100,
                          "amount": _a(f)} for a, b, c, d, r, f in lines]
    if lines:
        cl = rec["calc_lines"]
        rec["period_start"], rec["period_end"] = cl[0]["start"], cl[-1]["end"]
        rec["days_stated"] = sum(x["days"] for x in cl)
        rec["loan_balance"] = cl[0]["base"]
        rec["rate"] = cl[0]["rate"]
        rec["stated_amount"] = round(sum(x["amount"] for x in cl), 2)
    req = ["period_start", "period_end", "loan_balance", "rate", "stated_amount", "effective_date"]
    missing = [k for k in req if rec.get(k) is None]
    _chk(checks, "All required fields extracted", not missing, "missing: " + ", ".join(missing) if missing else "ok")
    if missing:
        return
    cl = rec["calc_lines"]
    _chk(checks, "Calculation lines are contiguous with a single base and rate",
         all((cl[i + 1]["start"] - cl[i]["end"]).days == 1 for i in range(len(cl) - 1))
         and len({(x["base"], x["rate"]) for x in cl}) == 1, f"{len(cl)} line(s)")
    days = (rec["period_end"] - rec["period_start"]).days + 1  # inclusive
    rec["days"] = days
    _chk(checks, "Days (inclusive) agree to invoice", days == rec["days_stated"], f"{days} vs {rec['days_stated']}")
    recalc = rec["loan_balance"] * rec["rate"] * days / DAY_BASIS
    rec["recalculated_amount"] = round(recalc, 2)
    _chk(checks, "Line fee recalculates (limit x rate x days / 365)", abs(recalc - rec["stated_amount"]) <= TOL,
         f"stated {rec['stated_amount']:,.2f} vs recalc {recalc:,.2f}")
    _chk(checks, "Invoice total = sum of calculation lines",
         rec["total"] is None or abs(rec["total"] - rec["stated_amount"]) < 0.005, f"total {rec['total']}")
    lag = (rec["effective_date"] - rec["period_end"]).days
    _chk(checks, "Debit date 1-10 days after period end", 1 <= lag <= 10,
         f"debit {rec['effective_date']} vs period end {rec['period_end']}", warn_only=True)
    _chk(checks, "Line fee rate is 1.40%", abs(rec["rate"] - 0.014) < 1e-9, f"{rec['rate']:.4%}", warn_only=True)
    _chk(checks, "Line fee base = $66.8m facility limit", abs(rec["loan_balance"] - 66_800_000) < 0.5,
         f"{rec['loan_balance']:,.2f}", warn_only=True)


def parse_other(kind, text, rec, checks):
    if kind == "break_cost":
        eff = re.search(r"Effective\s+" + D, text)
        tot = re.search(r"totalling\s+AUD\s+" + AMT, text, re.I)
        rec["effective_date"] = _d(eff.group(1)) if eff else None
        rec["amount"] = _a(tot.group(1)) if tot else None
        rec["summary"] = "Break Cost - one-off fee"
    elif kind == "repayment":
        on = re.search(r"On\s+" + D + r"\s+the following", text)
        tot = re.search(r"totalling:\s*AUD\s+" + AMT, text, re.I)
        intr = re.search(r"Interest Payment\*.*?totalling AUD\s+" + AMT, text, re.S)
        prin = re.search(r"Principal Payment\*.*?totalling AUD\s+" + AMT, text, re.S)
        rec["effective_date"] = _d(on.group(1)) if on else None
        rec["amount"] = _a(tot.group(1)) if tot else None
        rec["interest"] = _a(intr.group(1)) if intr else 0.0
        rec["principal"] = _a(prin.group(1)) if prin else 0.0
        rec["summary"] = f"Loan repayment - principal {rec['principal']:,.2f} + interest {rec['interest']:,.2f}"
    elif kind == "drawdown":
        eff = re.search(r"Effective Date:\s*" + D, text)
        term = re.search(r"Term:\s*" + D + r"\s*-\s*" + D, text)
        rate = re.search(r"all-in rate is:\s*([\d.]+)%", text)
        due = re.search(r"interest due on " + D + r" is AUD\s+" + AMT, text, re.I)
        amt = re.search(r"borrow a total of\s+AUD\s+" + AMT, text, re.S)
        rec["effective_date"] = _d(eff.group(1)) if eff else None
        if term:
            rec["period_start"], rec["period_end"] = _d(term.group(1)), _d(term.group(2))
        rec["rate"] = float(rate.group(1)) / 100 if rate else None
        rec["stated_amount"] = _a(due.group(2)) if due else None
        rec["loan_balance"] = _a(amt.group(1)) if amt else None
        rec["amount"] = rec["loan_balance"]
        rec["summary"] = "Drawdown (new borrowing)"
        if all(rec.get(k) for k in ("period_start", "period_end", "rate", "stated_amount", "loan_balance")):
            days = (rec["period_end"] - rec["period_start"]).days
            recalc = rec["loan_balance"] * rec["rate"] * days / DAY_BASIS
            rec["days"] = days
            rec["recalculated_amount"] = round(recalc, 2)
            _chk(checks, "Drawdown interest recalculates", abs(recalc - rec["stated_amount"]) <= TOL,
                 f"stated {rec['stated_amount']:,.2f} vs recalc {recalc:,.2f}")
    elif kind == "intent":
        eff = re.search(r"Effective Date:\s*" + D, text)
        rec["effective_date"] = _d(eff.group(1)) if eff else None
        ip = _section(text, r"INTEREST payments:", [r"Calculated as follows", r"Borrower "])
        a = re.findall(r"\s{2,}" + AMT + r"\s*$", ip, re.M)
        rec["amount"] = _a(a[0]) if a else None
        rec["summary"] = "Rollover Intent Notice - preliminary, superseded by the final Rollover Notice"


def parse_pdf(path):
    path = Path(path)
    text = pdf_text(path)
    flat = re.sub(r"\s+", " ", text)
    kind = classify(text)
    dd = re.search(r"Date:\s*" + D, text)
    rec = {"file": path.name, "doc_type": kind, "doc_date": _d(dd.group(1)) if dd else None}
    checks = []
    _chk(checks, "Facility is ROSE GRANGE COM BILAT AUD66.8M", FACILITY_TOKEN in flat, "deal ref in letter heading")
    _chk(checks, "Bank reference AUSNA1000072", BANK_REF in flat, "reference line", warn_only=True)
    _chk(checks, "Document date read", rec["doc_date"] is not None, str(rec["doc_date"]))
    _chk(checks, "Document type recognised", kind != "unknown", kind)
    if kind == "rollover":
        parse_rollover(text, rec, checks)
    elif kind == "line_fee":
        parse_line_fee(text, rec, checks)
    elif kind != "unknown":
        parse_other(kind, text, rec, checks)
    rec["is_duplicate_copy"] = "(duplicate)" in path.name.lower()
    rec["parse_checks"] = checks
    rec["parse_ok"] = not any(c["status"] == "FAIL" for c in checks)
    return rec


def _json_default(o):
    if isinstance(o, (date, datetime)):
        return o.isoformat()
    raise TypeError(type(o))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    out = [parse_pdf(p) for p in sys.argv[1:]]
    print(json.dumps(out, default=_json_default, indent=2))
