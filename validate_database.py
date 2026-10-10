#!/usr/bin/env python3
"""
Validate a healthcare pitch deck workbook and, optionally, compare it with another version.

Checks (sheet 'Healthcare Decks'):
  - duplicate company names
  - Website is a URL (hyperlink targets are read) or 'Not verified'
  - Amount Raised uses the form $325K / $54M / $1.6B (or Not verified)
  - Year is a four-digit year between 1990 and the current year (or Not verified)
  - Stage belongs to the standard list
  - the stage named in a summary ("Series D deck") agrees with the Stage column
  - Amount Raised text agrees with the numeric 'Amount (USD M)' column
  - heuristic: amount unusually large for the stage (possible lifetime total instead of round size)
  - unsupported verification wording ("Verified via ...") without a named source
  - promotional wording in the summary column

Usage:
  python validate_database.py DATABASE.xlsx
  python validate_database.py DATABASE.xlsx --compare OTHER.xlsx --report report.xlsx
Exit code: 1 when errors are found, otherwise 0.
"""
import argparse, datetime, re, sys
import pandas as pd
from common import (load_decks, is_missing, URL_RE, AMOUNT_RE, ALLOWED_STAGES, promo_hits, usd_millions,
                    stage_mismatch, amount_conflict)

STAGE_LIMIT_USD_M = {"Pre-Seed": 5, "Crowdfunding": 5, "Seed": 25, "Series A": 100}

def validate(df):
    out = []
    def add(sev, name, field, msg): out.append(dict(Severity=sev, Name=name, Field=field, Issue=msg))
    this_year = datetime.date.today().year
    for name in df["Name"][df["Name"].duplicated(keep=False)].unique():
        add("Error", name, "Name", "Duplicate company name.")
    for _, r in df.iterrows():
        n = r["Name"]
        w = str(r.get("Website", "")).strip()
        if "Website" in df and not is_missing(w) and not URL_RE.match(w):
            add("Error", n, "Website", f"Not a URL: '{w}'.")
        a = str(r.get("Amount Raised", "")).strip()
        if "Amount Raised" in df and not is_missing(a) and not AMOUNT_RE.match(a):
            add("Error", n, "Amount Raised", f"Unrecognised amount format: '{a}'.")
        y = str(r.get("Year", "")).strip()
        if "Year" in df and not is_missing(y):
            y = y[:-2] if y.endswith(".0") else y
            if not (y.isdigit() and 1990 <= int(y) <= this_year):
                add("Error", n, "Year", f"Year not valid: '{y}'.")
        s = str(r.get("Stage", "")).strip()
        if "Stage" in df and s not in ALLOWED_STAGES:
            add("Error", n, "Stage", f"Stage '{s}' is not in the standard list.")
        summ = str(r.get("Company / Deck Summary", ""))
        mm = stage_mismatch(summ, s)
        if mm:
            add("Error", n, "Stage", f"Stage is '{s}' but the summary refers to a {mm} deck.")
        conflict = amount_conflict(a, r.get("Amount (USD M)", ""))
        if conflict is not None:
            add("Warning", n, "Amount Raised", f"Amount Raised text ({a}) differs from the 'Amount (USD M)' column ({conflict:g}).")
        usd = usd_millions(a)
        lim = STAGE_LIMIT_USD_M.get(s)
        if usd is not None and lim and usd > lim:
            add("Warning", n, "Amount Raised", f"{a} is large for stage '{s}' (heuristic). It may be a lifetime total, not the round size.")
        flag = str(r.get("Review Flag", ""))
        if re.match(r"^\s*verified via", flag, re.I):
            add("Warning", n, "Review Flag", "Claims verification without naming a specific source or date.")
        for word in promo_hits(summ):
            add("Warning", n, "Company / Deck Summary", f"Promotional wording: '{word}'.")
    return pd.DataFrame(out, columns=["Severity", "Name", "Field", "Issue"])

def compare(a, b, label_a="A", label_b="B"):
    fields = [c for c in ("Amount Raised", "Year", "Stage", "Website") if c in a and c in b]
    m = a.merge(b, on="Name", how="outer", suffixes=(f" [{label_a}]", f" [{label_b}]"), indicator=True)
    rows = []
    for _, r in m.iterrows():
        if r["_merge"] != "both":
            rows.append(dict(Name=r["Name"], Field="(row)", **{label_a: "present" if r["_merge"] == "left_only" else "missing",
                                                              label_b: "present" if r["_merge"] == "right_only" else "missing"}))
            continue
        for f in fields:
            x, y = str(r[f"{f} [{label_a}]"]).strip(), str(r[f"{f} [{label_b}]"]).strip()
            x = x[:-2] if x.endswith(".0") else x; y = y[:-2] if y.endswith(".0") else y
            if f == "Website": x, y = x.rstrip("/").lower(), y.rstrip("/").lower()
            if x != y: rows.append(dict(Name=r["Name"], Field=f, **{label_a: x, label_b: y}))
    return pd.DataFrame(rows, columns=["Name", "Field", label_a, label_b])

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook"); ap.add_argument("--compare"); ap.add_argument("--report")
    ap.add_argument("--sheet", default="Healthcare Decks")
    a = ap.parse_args()
    df = load_decks(a.workbook, a.sheet)
    res = validate(df)
    print(f"{len(df)} rows checked in {a.workbook}")
    print(f"Errors: {(res.Severity == 'Error').sum()}   Warnings: {(res.Severity == 'Warning').sum()}")
    if len(res):
        print(res.groupby(["Severity", "Field"]).size().to_string())
    diff = None
    if a.compare:
        other = load_decks(a.compare, a.sheet)
        diff = compare(df, other, "This file", "Compared file")
        print(f"\nDifferences against {a.compare}: {len(diff)} cells/rows")
        if len(diff): print(diff.groupby("Field").size().to_string())
    if a.report:
        if a.report.lower().endswith(".xlsx"):
            with pd.ExcelWriter(a.report, engine="openpyxl") as xw:
                res.to_excel(xw, sheet_name="Findings", index=False)
                if diff is not None: diff.to_excel(xw, sheet_name="Differences", index=False)
                for ws in xw.book.worksheets:
                    for col in ws.columns:
                        ws.column_dimensions[col[0].column_letter].width = min(80, max(12, max(len(str(c.value or "")) for c in col) + 2))
                    ws.freeze_panes = "A2"
        else:
            res.to_csv(a.report, index=False)
        print("Report written:", a.report)
    sys.exit(1 if (res.Severity == "Error").any() else 0)

if __name__ == "__main__":
    main()
