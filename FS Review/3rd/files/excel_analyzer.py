#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════╗
║           DataLens AI - Excel File Analyzer              ║
║   Analyzes large Excel/CSV files using Claude AI         ║
╚══════════════════════════════════════════════════════════╝

Usage:
    python excel_analyzer.py --files file1.xlsx file2.csv --output report.md
    python excel_analyzer.py --files data.xlsx --prompts prompts.txt
    python excel_analyzer.py --files *.xlsx --chunk-size 200 --output results/
"""

import argparse
import os
import sys
import json
import time
import math
from pathlib import Path
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("❌ Missing dependency. Run: pip install pandas openpyxl anthropic")
    sys.exit(1)

try:
    import anthropic
except ImportError:
    print("❌ Missing dependency. Run: pip install anthropic")
    sys.exit(1)

# ─── CONFIG ──────────────────────────────────────────────────────────────────

DEFAULT_PROMPTS = [
    "Identify key trends and patterns across all files",
    "List all open issues, risks, and action items",
    "Provide strategic recommendations based on the data",
    "Highlight anomalies, outliers, or data quality issues",
    "Summarize the most critical findings",
]

MODEL = "claude-opus-4-6"
MAX_TOKENS = 4096
CHUNK_SIZE = 150          # rows per chunk for large files
MAX_CHUNKS_PER_FILE = 10  # limit total chunks to avoid huge costs
RATE_LIMIT_DELAY = 1.0    # seconds between API calls


# ─── COLORS (ANSI) ───────────────────────────────────────────────────────────

class C:
    RESET  = "\033[0m"
    BOLD   = "\033[1m"
    BLUE   = "\033[94m"
    CYAN   = "\033[96m"
    GREEN  = "\033[92m"
    YELLOW = "\033[93m"
    RED    = "\033[91m"
    PURPLE = "\033[95m"
    DIM    = "\033[2m"
    WHITE  = "\033[97m"

def c(text, color): return f"{color}{text}{C.RESET}"
def header(text): print(f"\n{c('═' * 60, C.DIM)}\n{c(f'  {text}', C.BOLD + C.CYAN)}\n{c('═' * 60, C.DIM)}")
def ok(text):   print(f"  {c('✓', C.GREEN)} {text}")
def warn(text): print(f"  {c('⚠', C.YELLOW)} {text}")
def err(text):  print(f"  {c('✗', C.RED)} {text}")
def info(text): print(f"  {c('·', C.BLUE)} {text}")
def step(n, text): print(f"\n{c(f'[{n}]', C.PURPLE + C.BOLD)} {c(text, C.WHITE)}")


# ─── FILE READING ─────────────────────────────────────────────────────────────

def read_file(path: Path) -> dict:
    """Read an Excel or CSV file and return metadata + chunked previews."""
    ext = path.suffix.lower()
    sheets_data = {}

    try:
        if ext == ".csv":
            df = pd.read_csv(path, encoding="utf-8", on_bad_lines="skip")
            sheets_data["Sheet1"] = df
        elif ext in (".xlsx", ".xls"):
            xl = pd.ExcelFile(path)
            for sheet in xl.sheet_names:
                try:
                    df = xl.parse(sheet)
                    if not df.empty:
                        sheets_data[sheet] = df
                except Exception as e:
                    warn(f"  Skipping sheet '{sheet}': {e}")
        else:
            raise ValueError(f"Unsupported file type: {ext}")
    except Exception as e:
        err(f"Failed to read {path.name}: {e}")
        return None

    total_rows = sum(len(df) for df in sheets_data.values())
    total_cols = max((len(df.columns) for df in sheets_data.values()), default=0)

    ok(f"{path.name}  {c(f'{total_rows:,} rows', C.CYAN)} · {c(f'{total_cols} cols', C.YELLOW)} · {c(f'{len(sheets_data)} sheet(s)', C.BLUE)}")

    return {
        "name": path.name,
        "path": str(path),
        "sheets": sheets_data,
        "total_rows": total_rows,
        "total_cols": total_cols,
        "size_mb": round(path.stat().st_size / 1024 / 1024, 2),
    }


def build_chunks(file_data: dict, chunk_size: int) -> list[str]:
    """Split file data into text chunks suitable for Claude context."""
    chunks = []

    for sheet_name, df in file_data["sheets"].items():
        n_rows = len(df)
        n_chunks = min(MAX_CHUNKS_PER_FILE, math.ceil(n_rows / chunk_size))

        # Always include schema / column stats
        col_info = []
        for col in df.columns:
            dtype = str(df[col].dtype)
            null_pct = round(df[col].isna().mean() * 100, 1)
            if df[col].dtype in ["int64", "float64"]:
                stats = f"min={df[col].min():.2g}, max={df[col].max():.2g}, mean={df[col].mean():.2g}"
            else:
                top = df[col].dropna().value_counts().head(3)
                stats = "top: " + ", ".join(f"'{v}'({c})" for v, c in top.items()) if not top.empty else "—"
            col_info.append(f"  • {col} [{dtype}] nulls={null_pct}%  {stats}")

        schema_block = (
            f"FILE: {file_data['name']}  |  Sheet: {sheet_name}  |  "
            f"{n_rows:,} rows × {len(df.columns)} cols\n"
            f"COLUMNS:\n" + "\n".join(col_info)
        )

        # Sample chunks
        rows_per_chunk = max(1, n_rows // n_chunks)
        for i in range(n_chunks):
            start = i * rows_per_chunk
            end = min(start + rows_per_chunk, n_rows)
            sample = df.iloc[start:end]

            chunk_text = (
                f"{schema_block}\n\n"
                f"DATA SAMPLE (rows {start+1}–{end} of {n_rows:,}):\n"
                f"{sample.to_string(max_rows=chunk_size, max_cols=20, show_dimensions=True)}"
            )
            chunks.append(chunk_text)

        if n_rows > n_chunks * rows_per_chunk:
            warn(f"  '{sheet_name}' has {n_rows:,} rows — sampled {n_chunks} chunks of ~{rows_per_chunk} rows each")

    return chunks


# ─── AI ANALYSIS ─────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are a senior data analyst. Analyze the provided Excel/CSV data and respond using EXACTLY these section headers:

SUMMARY:
(Write 3-4 sentences: what the data is, scope, key takeaways)

RECOMMENDATIONS:
- (Each recommendation on its own line, be specific and actionable)

OPEN POINTS:
- (Each open issue, risk, or required action on its own line)

KEY FINDINGS:
- (Each notable pattern, anomaly, or insight on its own line)

DATA QUALITY:
- (Each data quality issue found: nulls, inconsistencies, outliers)

Rules:
- Reference actual column names, values, and row counts
- Be specific, not generic
- If analyzing a chunk, note it covers rows X–Y of N total
"""


def analyze_chunk(client: anthropic.Anthropic, chunk_text: str, prompts: list[str], chunk_label: str) -> str:
    prompt_list = "\n".join(f"{i+1}. {p}" for i, p in enumerate(prompts))
    user_msg = f"Analyze this data:\n\n{chunk_text}\n\n---\nSpecifically answer:\n{prompt_list}"

    try:
        resp = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_msg}],
        )
        return resp.content[0].text
    except anthropic.RateLimitError:
        warn(f"Rate limit hit on {chunk_label}, waiting 30s...")
        time.sleep(30)
        return analyze_chunk(client, chunk_text, prompts, chunk_label)
    except Exception as e:
        err(f"API error on {chunk_label}: {e}")
        return f"[ERROR analyzing {chunk_label}: {e}]"


def merge_analyses(client: anthropic.Anthropic, analyses: list[dict], prompts: list[str]) -> str:
    """Merge multiple chunk analyses into a single coherent report."""
    info("Merging chunk analyses into final report...")

    combined = "\n\n".join(
        f"=== Analysis of {a['label']} ===\n{a['result']}" for a in analyses
    )
    prompt_list = "\n".join(f"{i+1}. {p}" for i, p in enumerate(prompts))

    merge_msg = (
        f"Below are analyses from multiple data chunks/files. "
        f"Synthesize them into ONE unified report, removing duplicates and reconciling conflicts.\n\n"
        f"{combined}\n\n---\nOriginal questions to answer:\n{prompt_list}"
    )

    try:
        resp = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": merge_msg}],
        )
        return resp.content[0].text
    except Exception as e:
        err(f"Merge API error: {e}")
        return combined  # fallback: return raw combined


# ─── REPORT GENERATION ───────────────────────────────────────────────────────

def parse_sections(raw: str) -> dict:
    sections = {
        "summary": [],
        "recommendations": [],
        "open_points": [],
        "findings": [],
        "data_quality": [],
    }
    current = None
    map_ = {
        "summary": "summary",
        "recommendation": "recommendations",
        "open": "open_points",
        "key finding": "findings",
        "finding": "findings",
        "data quality": "data_quality",
    }
    for line in raw.splitlines():
        stripped = line.strip()
        low = stripped.lower().rstrip(":")
        matched = next((v for k, v in map_.items() if low.startswith(k)), None)
        if matched and stripped.endswith(":"):
            current = matched
            continue
        if current and stripped:
            clean = stripped.lstrip("-•* ").strip()
            if clean:
                sections[current].append(clean)
    return sections


def build_markdown_report(files_meta: list, all_analyses: list, final_report: str,
                           prompts: list[str], sections: dict, elapsed: float) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_table = "\n".join(
        f"| {f['name']} | {f['total_rows']:,} | {f['total_cols']} | {len(f['sheets'])} | {f['size_mb']} MB |"
        for f in files_meta if f
    )

    rec_lines   = "\n".join(f"- {r}" for r in sections["recommendations"]) or "_None identified_"
    open_lines  = "\n".join(f"- ⚠️  {o}" for o in sections["open_points"])    or "_None identified_"
    find_lines  = "\n".join(f"- {x}" for x in sections["findings"])           or "_None identified_"
    dq_lines    = "\n".join(f"- {x}" for x in sections["data_quality"])       or "_No issues found_"

    prompts_md = "\n".join(f"{i+1}. {p}" for i, p in enumerate(prompts))

    return f"""# DataLens AI Analysis Report

> Generated: {now} | Model: {MODEL} | Analysis time: {elapsed:.1f}s

---

## 📁 Files Analyzed

| File | Rows | Columns | Sheets | Size |
|------|------|---------|--------|------|
{file_table}

---

## 🔍 Analysis Prompts

{prompts_md}

---

## 📋 Executive Summary

{chr(10).join(sections["summary"]) or final_report.split("RECOMMENDATIONS")[0].strip()}

---

## ✅ Recommendations

{rec_lines}

---

## ⚠️ Open Points & Action Items

{open_lines}

---

## 💡 Key Findings

{find_lines}

---

## 🧹 Data Quality Issues

{dq_lines}

---

## 📄 Full AI Response

```
{final_report}
```

---
*Report generated by DataLens AI Python Analyzer*
"""


def build_json_report(files_meta, sections, final_report, prompts, elapsed):
    return {
        "metadata": {
            "generated_at": datetime.now().isoformat(),
            "model": MODEL,
            "elapsed_seconds": round(elapsed, 2),
            "files_analyzed": [
                {"name": f["name"], "rows": f["total_rows"], "cols": f["total_cols"], "size_mb": f["size_mb"]}
                for f in files_meta if f
            ],
            "prompts": prompts,
        },
        "summary": " ".join(sections["summary"]),
        "recommendations": sections["recommendations"],
        "open_points": sections["open_points"],
        "findings": sections["findings"],
        "data_quality": sections["data_quality"],
        "raw_ai_response": final_report,
    }


# ─── CLI ─────────────────────────────────────────────────────────────────────

def load_prompts_from_file(path: str) -> list[str]:
    p = Path(path)
    if not p.exists():
        err(f"Prompts file not found: {path}")
        return []
    lines = [l.strip() for l in p.read_text().splitlines() if l.strip() and not l.startswith("#")]
    ok(f"Loaded {len(lines)} prompts from {p.name}")
    return lines


def main():
    parser = argparse.ArgumentParser(
        description="DataLens AI — Analyze large Excel/CSV files with Claude",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python excel_analyzer.py --files data.xlsx
  python excel_analyzer.py --files *.xlsx --output report.md
  python excel_analyzer.py --files a.xlsx b.csv --prompts my_prompts.txt --output results/
  python excel_analyzer.py --files big.xlsx --chunk-size 200 --format both
        """
    )
    parser.add_argument("--files",       nargs="+", required=True,  help="Excel/CSV files to analyze (supports glob)")
    parser.add_argument("--prompts",     type=str,  default=None,   help="Path to .txt file with one prompt per line")
    parser.add_argument("--output",      type=str,  default=None,   help="Output file path (.md / .json) or directory")
    parser.add_argument("--format",      choices=["md", "json", "both"], default="md", help="Output format (default: md)")
    parser.add_argument("--chunk-size",  type=int,  default=CHUNK_SIZE, help=f"Rows per chunk (default: {CHUNK_SIZE})")
    parser.add_argument("--api-key",     type=str,  default=None,   help="Anthropic API key (or set ANTHROPIC_API_KEY env var)")
    parser.add_argument("--no-merge",    action="store_true",       help="Skip final merge step (faster, less coherent)")

    args = parser.parse_args()

    # ── Banner
    print(f"""
{c('╔══════════════════════════════════════════════╗', C.BLUE)}
{c('║', C.BLUE)}  {c('DataLens AI  ·  Excel Analyzer', C.BOLD + C.WHITE)}            {c('║', C.BLUE)}
{c('║', C.BLUE)}  {c('Powered by Claude · github.com/anthropics', C.DIM)}  {c('║', C.BLUE)}
{c('╚══════════════════════════════════════════════╝', C.BLUE)}
""")

    # ── API Key
    api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        err("No API key found. Set ANTHROPIC_API_KEY env var or use --api-key")
        print(f"\n  {c('export ANTHROPIC_API_KEY=sk-ant-...', C.CYAN)}\n")
        sys.exit(1)
    client = anthropic.Anthropic(api_key=api_key)

    # ── Prompts
    if args.prompts:
        prompts = load_prompts_from_file(args.prompts)
        if not prompts:
            prompts = DEFAULT_PROMPTS
    else:
        prompts = DEFAULT_PROMPTS
        info(f"Using {len(prompts)} default prompts (use --prompts to customize)")

    # ── Load Files
    step(1, "Loading files")
    file_paths = []
    for pattern in args.files:
        matches = list(Path(".").glob(pattern)) or [Path(pattern)]
        file_paths.extend(matches)

    file_paths = [p for p in file_paths if p.exists() and p.suffix.lower() in (".xlsx", ".xls", ".csv")]
    if not file_paths:
        err("No valid Excel/CSV files found.")
        sys.exit(1)
    if len(file_paths) > 5:
        warn(f"Found {len(file_paths)} files, using first 5.")
        file_paths = file_paths[:5]

    files_meta = [read_file(p) for p in file_paths]
    files_meta = [f for f in files_meta if f is not None]
    if not files_meta:
        err("No files could be read.")
        sys.exit(1)

    # ── Chunk
    step(2, "Chunking data for analysis")
    all_chunks = []
    for fm in files_meta:
        chunks = build_chunks(fm, args.chunk_size)
        info(f"{fm['name']} → {len(chunks)} chunk(s)")
        for i, chunk in enumerate(chunks):
            all_chunks.append({"label": f"{fm['name']} chunk {i+1}/{len(chunks)}", "text": chunk})

    info(f"Total: {len(all_chunks)} chunk(s) to analyze")

    # ── Analyze
    step(3, f"Analyzing with {c(MODEL, C.CYAN)}")
    start = time.time()
    analyses = []

    for i, chunk in enumerate(all_chunks):
        print(f"  {c(f'[{i+1}/{len(all_chunks)}]', C.DIM)} Analyzing {c(chunk['label'], C.YELLOW)}...", end=" ", flush=True)
        result = analyze_chunk(client, chunk["text"], prompts, chunk["label"])
        analyses.append({"label": chunk["label"], "result": result})
        print(c("done", C.GREEN))
        if i < len(all_chunks) - 1:
            time.sleep(RATE_LIMIT_DELAY)

    # ── Merge
    if len(analyses) > 1 and not args.no_merge:
        step(4, "Merging chunk analyses")
        final_report = merge_analyses(client, analyses, prompts)
    else:
        final_report = analyses[0]["result"] if analyses else ""

    elapsed = time.time() - start

    # ── Parse & Display
    step(5, "Results")
    sections = parse_sections(final_report)

    print(f"\n{c('  SUMMARY', C.BOLD + C.CYAN)}")
    for s in sections["summary"]:
        print(f"  {s}")

    print(f"\n{c('  RECOMMENDATIONS', C.BOLD + C.GREEN)}  ({len(sections['recommendations'])})")
    for r in sections["recommendations"][:5]:
        print(f"  {c('✓', C.GREEN)} {r}")

    print(f"\n{c('  OPEN POINTS', C.BOLD + C.YELLOW)}  ({len(sections['open_points'])})")
    for o in sections["open_points"][:5]:
        print(f"  {c('⚠', C.YELLOW)} {o}")

    print(f"\n{c('  KEY FINDINGS', C.BOLD + C.PURPLE)}  ({len(sections['findings'])})")
    for f in sections["findings"][:5]:
        print(f"  {c('◆', C.PURPLE)} {f}")

    if sections["data_quality"]:
        print(f"\n{c('  DATA QUALITY', C.BOLD + C.RED)}  ({len(sections['data_quality'])})")
        for d in sections["data_quality"][:3]:
            print(f"  {c('!', C.RED)} {d}")

    # ── Save Output
    step(6, "Saving report")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_base = args.output

    if out_base:
        out_path = Path(out_base)
        if out_path.is_dir() or str(out_base).endswith("/"):
            out_path.mkdir(parents=True, exist_ok=True)
            base_name = out_path / f"datalens_report_{ts}"
        else:
            base_name = out_path.with_suffix("")  # strip extension, we'll add it
    else:
        base_name = Path(f"datalens_report_{ts}")

    saved = []
    if args.format in ("md", "both"):
        md = build_markdown_report(files_meta, analyses, final_report, prompts, sections, elapsed)
        md_path = Path(str(base_name) + ".md")
        md_path.write_text(md, encoding="utf-8")
        saved.append(str(md_path))
        ok(f"Markdown report → {c(str(md_path), C.CYAN)}")

    if args.format in ("json", "both"):
        jdata = build_json_report(files_meta, sections, final_report, prompts, elapsed)
        json_path = Path(str(base_name) + ".json")
        json_path.write_text(json.dumps(jdata, indent=2, ensure_ascii=False), encoding="utf-8")
        saved.append(str(json_path))
        ok(f"JSON report     → {c(str(json_path), C.CYAN)}")

    n_recs = len(sections["recommendations"])
    n_open = len(sections["open_points"])
    print(f"\n{c('═' * 60, C.DIM)}")
    print(f"  {c('✓ Analysis complete!', C.BOLD + C.GREEN)}  "
          f"{c(f'{len(all_chunks)} chunks', C.DIM)} · "
          f"{c(f'{elapsed:.1f}s', C.DIM)} · "
          f"{c(f'{n_recs} recs', C.GREEN)} · "
          f"{c(f'{n_open} open pts', C.YELLOW)}")
    print(f"{c('═' * 60, C.DIM)}\n")


if __name__ == "__main__":
    main()
