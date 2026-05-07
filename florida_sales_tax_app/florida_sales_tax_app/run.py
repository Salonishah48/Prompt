"""
Florida Sales Tax App — Main Runner
====================================

By default, the app treats each file in sales_data/ as a separate filing
period and produces one DR-15 output set per file.

Workflow:
  1. Edit  rules/business_config.json   (your business details — ONCE)
  2. Drop your Shopify tax-export Excel file(s) into  sales_data/
     — one file per filing period (e.g. shopify_Jan_2026.xlsx,
        shopify_Feb_2026.xlsx, shopify_Oct-Dec_2025.xlsx)
  3. Run   python run.py
  4. Review the filled DR-15 files in  output/

Usage:
    python run.py
        Process EACH file separately. One set of outputs per input file.

    python run.py --combine
        Combine all files into one DR-15 (old behavior).

    python run.py --period 2025-07
        Filter every file to only transactions in July 2025.

    python run.py --from 2025-07-01 --to 2025-09-30
        Filter every file to a custom date range.

    python run.py --late
        Mark all returns as late (applies 10% / $50 min penalty).

    python run.py --template
        Produce a BLANK DR-15 template PDF + Excel.
"""

from __future__ import annotations

import argparse
import json
from datetime import date, datetime
from pathlib import Path

from tax_engine import FloridaSalesTaxEngine, DR15Return, return_to_dict
from data_loader import load_all_daily_data, load_per_file
from output_formatter import render_excel, render_pdf

ROOT = Path(__file__).parent
RULES = ROOT / "rules" / "tax_rules.json"
BUSINESS = ROOT / "rules" / "business_config.json"
DATA = ROOT / "sales_data"
OUT = ROOT / "output"
OUT.mkdir(exist_ok=True)


def month_range(period_yyyy_mm: str):
    y, m = map(int, period_yyyy_mm.split("-"))
    start = date(y, m, 1)
    if m == 12:
        end = date(y + 1, 1, 1)
    else:
        end = date(y, m + 1, 1)
    end_inclusive = date.fromordinal(end.toordinal() - 1)
    return start.isoformat(), end_inclusive.isoformat()


def detect_period_from_data(transactions: list) -> tuple:
    """Find the earliest and latest transaction date in the data."""
    dates = []
    for t in transactions:
        d = t.get("date")
        if not d:
            continue
        try:
            dates.append(datetime.fromisoformat(d[:10]).date())
        except ValueError:
            continue
    if not dates:
        return None, None
    return min(dates).isoformat(), max(dates).isoformat()


def format_period_label(start: str, end: str) -> str:
    if not start or not end:
        return "________"
    sy, sm, _ = start.split("-")
    ey, em, _ = end.split("-")
    if sy == ey and sm == em:
        return f"{sy}-{sm}"
    return f"{sy}-{sm} to {ey}-{em}"


def file_tag_from_filename(filename: str, fallback_period: str) -> str:
    """Produce a filename-safe tag for the output files.
    Uses the input filename stem (without extension). Falls back to the
    period label if the stem is empty."""
    stem = Path(filename).stem
    # Strip common Shopify prefixes for a cleaner output name
    for prefix in ("shopify_", "Shopify_", "shopify-"):
        if stem.startswith(prefix):
            stem = stem[len(prefix):]
            break
    # Clean up unsafe characters
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in stem)
    return safe or fallback_period.replace(" ", "_").replace("to", "-")


def build_and_write(
    engine: FloridaSalesTaxEngine,
    transactions: list,
    file_tag: str,
    source_label: str,
    args,
) -> DR15Return:
    """Build a DR-15 for the given transactions and write output files."""
    detected_start, detected_end = detect_period_from_data(transactions)
    period_label = format_period_label(detected_start, detected_end)

    print(f"\n  ── {source_label} ──")
    print(f"     Reporting period : {period_label}  ({detected_start} → {detected_end})")
    print(f"     Transactions     : {len(transactions)}")

    ret = engine.build_return(
        reporting_period=period_label,
        line_a_txns=transactions,
        lawful_deductions=args.lawful_deductions,
        est_tax_paid_last_month=args.est_tax_paid,
        est_tax_due_current_month=args.est_tax_due,
        is_late=args.late,
        scholarship_credits=args.scholarship_credits,
        high_crime_credits=args.high_crime_credits,
        other_credits=args.other_credits,
    )

    print(f"     Gross sales      : ${float(ret.line_a.gross_sales):>14,.2f}")
    print(f"     Exempt sales     : ${float(ret.line_a.exempt_sales):>14,.2f}")
    print(f"     Taxable amount   : ${float(ret.line_a.taxable_amount):>14,.2f}")
    print(f"     State + surtax   : ${float(ret.line_a.tax_due):>14,.2f}")
    if float(ret.shopify_tax_collected) > 0:
        gap = float(ret.tax_gap)
        label = "SHORT" if gap > 0 else ("OVER" if gap < 0 else "match")
        print(f"     Shopify collected: ${float(ret.shopify_tax_collected):>14,.2f}")
        print(f"     Gap              : ${gap:>14,.2f}  ({label})")
    print(f"     Line 14 (remit)  : ${float(ret.line_14_amount_due_with_return):>14,.2f}")

    # Write outputs
    pdf_out = OUT / f"DR15_{file_tag}_filled.pdf"
    xlsx_out = OUT / f"DR15_{file_tag}_filled.xlsx"
    json_out = OUT / f"DR15_{file_tag}_filled.json"

    render_pdf(ret, engine.business, transactions, pdf_out)
    render_excel(ret, engine.business, transactions, xlsx_out)
    json_out.write_text(json.dumps(return_to_dict(ret), indent=2, default=str))

    print(f"     Output → {pdf_out.name}")
    print(f"              {xlsx_out.name}")

    return ret


def blank_template(engine: FloridaSalesTaxEngine) -> DR15Return:
    return DR15Return(
        reporting_period="________",
        county=engine.business["business_info"]["county"],
        surtax_rate=engine.county_surtax_rate(),
    )


def main():
    p = argparse.ArgumentParser(description="Florida DR-15 Sales & Use Tax runner")
    p.add_argument("--combine", action="store_true",
                   help="Combine all files into ONE DR-15 (default is one per file)")
    p.add_argument("--period", help="Single-month reporting period YYYY-MM (filter)")
    p.add_argument("--from", dest="date_from", help="Start date YYYY-MM-DD (inclusive)")
    p.add_argument("--to",   dest="date_to",   help="End date   YYYY-MM-DD (inclusive)")
    p.add_argument("--late", action="store_true", help="Mark return as late (applies penalty)")
    p.add_argument("--lawful-deductions", type=float, default=0)
    p.add_argument("--est-tax-paid", type=float, default=0)
    p.add_argument("--est-tax-due", type=float, default=0)
    p.add_argument("--scholarship-credits", type=float, default=0)
    p.add_argument("--high-crime-credits", type=float, default=0)
    p.add_argument("--other-credits", type=float, default=0)
    p.add_argument("--template", action="store_true",
                   help="Generate a BLANK DR-15 PDF + Excel template and exit")
    args = p.parse_args()

    engine = FloridaSalesTaxEngine(str(RULES), str(BUSINESS))

    if engine.applied_updates:
        print(f"\n  Rules updates applied ({len(engine.applied_updates)}):")
        for u in engine.applied_updates:
            print(f"    ✓ {u}")

    if args.template:
        ret = blank_template(engine)
        pdf_out = OUT / "DR15_blank_template.pdf"
        xlsx_out = OUT / "DR15_blank_template.xlsx"
        render_pdf(ret, engine.business, [], pdf_out)
        render_excel(ret, engine.business, [], xlsx_out)
        print(f"\nBlank templates written:")
        print(f"  {pdf_out}")
        print(f"  {xlsx_out}")
        return

    # Date filter (applies to every file)
    period_start = None
    period_end = None
    if args.period:
        period_start, period_end = month_range(args.period)
    elif args.date_from or args.date_to:
        period_start = args.date_from
        period_end = args.date_to

    print(f"\n▶ Florida Sales Tax App")
    print(f"  Reading from: {DATA}/")
    print(f"  Business county: {engine.business['business_info']['county']} "
          f"(surtax {float(engine.county_surtax_rate())*100:.2f}%)")

    # ------------------ COMBINED MODE (old behavior) ------------------
    if args.combine:
        print(f"  Mode: COMBINED (all files → one DR-15)")
        transactions, files_read = load_all_daily_data(
            str(DATA), period_start, period_end,
        )
        if not transactions:
            print("  No transactions found.")
            return
        for f in files_read:
            print(f"    • {f}")
        build_and_write(
            engine, transactions,
            file_tag="combined",
            source_label="Combined output",
            args=args,
        )
        return

    # ------------------ PER-FILE MODE (default) ------------------
    print(f"  Mode: PER-FILE (one DR-15 per input file)")
    per_file = load_per_file(str(DATA), period_start, period_end)

    if not per_file:
        print("  No supported files found in sales_data/.")
        return

    print(f"  Found {len(per_file)} file(s) to process.")

    total_remit = 0.0
    for filename, transactions in per_file:
        file_tag = file_tag_from_filename(filename, "period")
        ret = build_and_write(
            engine, transactions,
            file_tag=file_tag,
            source_label=filename,
            args=args,
        )
        total_remit += float(ret.line_14_amount_due_with_return)

    # Final summary across all files
    print(f"\n  ══════════════════════════════════════════════")
    print(f"   {len(per_file)} return(s) prepared")
    print(f"   Total to remit across all periods: ${total_remit:,.2f}")
    print(f"  ══════════════════════════════════════════════")


if __name__ == "__main__":
    main()
