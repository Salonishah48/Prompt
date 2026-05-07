# 🔍 GL Audit Intelligence

A local-first, AI-powered General Ledger forensic audit tool.  
**No data leaves your machine.** Powered by Ollama (Llama 3 / Mistral).

---

## 🚀 Quick Start

### 1. Install Ollama
```bash
curl -fsSL https://ollama.com/install.sh | sh
ollama pull llama3        # or: ollama pull mistral
```

### 2. Install Python Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the App
```bash
streamlit run app.py
```

Open http://localhost:8501 in your browser.

---

## 📁 Project Structure

```
gl_audit_app/
├── app.py                # Streamlit UI
├── audit_engine.py       # LLM batching + prompt logic
├── report_generator.py   # Stats + Excel export
├── requirements.txt
├── sample_gl_data.csv    # Test dataset (20 rows)
└── README.md
```

---

## 🏗️ Architecture

```
Upload CSV/Excel
      │
      ▼
  Pandas (clean + map columns)
      │
      ▼
  Batch Splitter (10–30 rows/batch)
      │
      ▼
  Ollama API  ◄──  System Prompt (Forensic Accountant role)
  (Llama 3 / Mistral)
      │
      ▼
  JSON Parser + Normaliser
      │
      ▼
  Streamlit Dashboard
  (KPIs · Filtered Table · CSV/JSON Export)
```

---

## 🔎 What the LLM Detects

| Issue | Description |
|---|---|
| **Duplicate** | Same vendor + amount + date |
| **GL Mismatch** | Vendor category contradicts GL account |
| **Missing Data** | No vendor name or description |
| **Uncategorized** | Sitting in Suspense / Miscellaneous |
| **None** | Clean entry |

Each entry also gets a **Transaction Type**: Operational · Capital · Payroll · Tax · Travel · IT · Marketing

---

## ⚙️ Configuration

All settings in the sidebar:
- **Ollama Model**: `llama3`, `mistral`, `llama3.1`, `mistral-nemo`
- **Ollama Base URL**: default `http://localhost:11434`
- **Batch Size**: 5–30 rows per LLM call (smaller = more accurate, slower)

---

## 📤 Export

Download results as:
- **CSV** — for Excel analysis
- **JSON** — for downstream systems / APIs

---

## 🔒 Privacy

All processing is local. No GL data is sent to any external API.
Ollama runs entirely on your machine.
