"""
report_generator.py
Utilities to compute summary stats and export audit results.
"""

import pandas as pd
import io
from typing import List, Dict, Any


def generate_summary_stats(results: List[Dict[str, Any]]) -> Dict[str, int]:
    total          = len(results)
    duplicates     = sum(1 for r in results if r.get("issue") == "Duplicate")
    mismatches     = sum(1 for r in results if r.get("issue") == "GL Mismatch")
    missing        = sum(1 for r in results if r.get("issue") == "Missing Data")
    uncategorized  = sum(1 for r in results if r.get("issue") == "Uncategorized")
    clean          = sum(1 for r in results if r.get("issue") == "None")

    return {
        "total":         total,
        "duplicates":    duplicates,
        "mismatches":    mismatches,
        "missing":       missing,
        "uncategorized": uncategorized,
        "clean":         clean,
        "flagged":       total - clean,
    }


def generate_excel_report(results: List[Dict[str, Any]]) -> bytes:
    """
    Write audit results to an in-memory Excel workbook with two sheets:
    - Summary  (KPI table)
    - Details  (full results)
    Returns bytes suitable for st.download_button.
    """
    df      = pd.DataFrame(results)
    stats   = generate_summary_stats(results)
    summary = pd.DataFrame([
        {"Metric": "Total Transactions",  "Count": stats["total"]},
        {"Metric": "Duplicates",          "Count": stats["duplicates"]},
        {"Metric": "GL Mismatches",       "Count": stats["mismatches"]},
        {"Metric": "Missing Data",        "Count": stats["missing"]},
        {"Metric": "Uncategorized",       "Count": stats["uncategorized"]},
        {"Metric": "Clean Entries",       "Count": stats["clean"]},
        {"Metric": "Total Flagged",       "Count": stats["flagged"]},
    ])

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        df.to_excel(writer, sheet_name="Audit Details", index=False)
    buf.seek(0)
    return buf.read()
