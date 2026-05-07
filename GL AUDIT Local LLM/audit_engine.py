"""
audit_engine.py
Sends GL transactions in batches to Ollama and parses structured audit results.
Fixed for Ollama v0.2.x - uses /api/chat endpoint with correct message format.
"""

import json
import re
import requests
import pandas as pd

BATCH_SIZE = 3

SYSTEM_PROMPT = """You are an expert Forensic Accountant and Data Auditor.
Your task is to analyze General Ledger (GL) entries for inconsistencies, errors, and fraud markers.

Analyze the provided transaction list and return ONLY a valid JSON object with this exact structure:
{
  "audit_results": [
    {
      "row_id": <integer>,
      "issue": "<Duplicate|GL Mismatch|Missing Data|Uncategorized|None>",
      "confidence": "<High|Medium|Low>",
      "transaction_type": "<Operational|Capital|Payroll|Tax|Travel|Meals|IT|Marketing|Unknown>",
      "suggested_gl": "<suggested GL account or null>",
      "recommendation": "<brief actionable note>"
    }
  ]
}

Rules:
1. Duplicate Entries: Flag transactions with identical amounts, dates, and vendors.
2. GL Inconsistency: Flag if vendor name contradicts the GL account (e.g. "Starbucks" mapped to "Office Equipment" should be "Meals & Entertainment").
3. Missing Data: Flag entries missing vendor names or descriptions.
4. Uncategorized: Flag entries in "Suspense", "Miscellaneous", or "Clearing" accounts.
5. Categorization: Assign transaction_type to EVERY entry.
6. If no issue found, set issue to "None".
7. Return ONLY raw JSON. No markdown. No explanation. No backticks."""


def build_user_prompt(batch_df: pd.DataFrame, offset: int) -> str:
    records = []
    for i, (_, row) in enumerate(batch_df.iterrows()):
        rec = {"row_id": offset + i + 1}
        for col in batch_df.columns:
            val = row[col]
            rec[col] = str(val) if pd.notna(val) else None
        records.append(rec)
    data_json = json.dumps(records, indent=2)
    return f"Analyze these General Ledger transactions:\n\n{data_json}\n\nReturn ONLY the JSON audit_results array as specified."


def call_ollama(prompt: str, model: str, base_url: str) -> str:
    url = f"{base_url.rstrip('/')}/api/chat"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": prompt}
        ],
        "stream": False,
        "options": {"temperature": 0.1}
    }
    resp = requests.post(url, json=payload, timeout=300)
    resp.raise_for_status()
    data = resp.json()
    # v0.2.x returns message.content
    return data["message"]["content"]


def extract_json(raw: str) -> dict:
    cleaned = re.sub(r"```(?:json)?", "", raw).strip().rstrip("`").strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass
    match = re.search(r'\{.*\}', cleaned, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass
    raise ValueError(f"Could not parse JSON from model response:\n{raw[:500]}")


def run_audit(
    batch_df: pd.DataFrame,
    offset: int,
    model: str = "mistral",
    base_url: str = "http://localhost:11434"
) -> list:
    prompt   = build_user_prompt(batch_df, offset)
    raw_resp = call_ollama(prompt, model, base_url)
    parsed   = extract_json(raw_resp)
    results  = parsed.get("audit_results", [])
    clean = []
    for r in results:
        clean.append({
            "row_id":           r.get("row_id", offset + 1),
            "issue":            r.get("issue", "None"),
            "confidence":       r.get("confidence", "Medium"),
            "transaction_type": r.get("transaction_type", "Unknown"),
            "suggested_gl":     r.get("suggested_gl"),
            "recommendation":   r.get("recommendation", ""),
        })
    return clean
