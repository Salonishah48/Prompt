import streamlit as st
import pandas as pd
import json
import time
from audit_engine import run_audit, BATCH_SIZE
from report_generator import generate_summary_stats, generate_excel_report
import io

# ─── Page Config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GL Audit Intelligence",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── Custom CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=Syne:wght@400;700;800&display=swap');

  html, body, [class*="css"] { font-family: 'Syne', sans-serif; }
  code, .stCode { font-family: 'IBM Plex Mono', monospace; }

  .main { background: #0a0d14; }
  .block-container { padding-top: 2rem; }

  /* Sidebar */
  [data-testid="stSidebar"] {
    background: #0e1220;
    border-right: 1px solid #1e2d4a;
  }

  /* Cards */
  .metric-card {
    background: #0e1220;
    border: 1px solid #1e2d4a;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    margin-bottom: 1rem;
  }
  .metric-card h3 { margin: 0; font-size: 2rem; font-weight: 800; }
  .metric-card p  { margin: 0; color: #8899aa; font-size: 0.8rem; letter-spacing: 0.12em; text-transform: uppercase; }

  /* Issue badges */
  .badge-duplicate   { background:#ff4d6d22; color:#ff4d6d; border:1px solid #ff4d6d55; padding:2px 10px; border-radius:999px; font-size:0.75rem; }
  .badge-mismatch    { background:#f7931a22; color:#f7931a; border:1px solid #f7931a55; padding:2px 10px; border-radius:999px; font-size:0.75rem; }
  .badge-missing     { background:#9b59b622; color:#c39bd3; border:1px solid #9b59b655; padding:2px 10px; border-radius:999px; font-size:0.75rem; }
  .badge-uncategorized{ background:#1abc9c22; color:#1abc9c; border:1px solid #1abc9c55; padding:2px 10px; border-radius:999px; font-size:0.75rem; }
  .badge-none        { background:#2ecc7122; color:#2ecc71; border:1px solid #2ecc7155; padding:2px 10px; border-radius:999px; font-size:0.75rem; }

  /* Header */
  .app-header {
    border-bottom: 1px solid #1e2d4a;
    padding-bottom: 1rem;
    margin-bottom: 2rem;
  }
  .app-header h1 { font-size: 2.2rem; font-weight: 800; color: #e8f4fd; margin: 0; }
  .app-header p  { color: #5577aa; font-size: 0.9rem; margin: 0; font-family: 'IBM Plex Mono', monospace; }

  /* Tables */
  .dataframe { background: #0e1220 !important; }

  /* Progress */
  .stProgress > div > div { background: linear-gradient(90deg, #00d4ff, #0099ff); }
</style>
""", unsafe_allow_html=True)

# ─── Header ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <h1>🔍 GL Audit Intelligence</h1>
  <p>// Forensic General Ledger Analysis · Powered by Local LLM (Ollama)</p>
</div>
""", unsafe_allow_html=True)

# ─── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    ollama_model = st.selectbox("Ollama Model", ["phi3:mini", "mistral", "llama3.1", "llama3", "mistral-nemo"], index=0)
    ollama_url   = st.text_input("Ollama Base URL", value="http://localhost:11434")
    batch_size   = st.slider("Batch Size (rows per LLM call)", min_value=5, max_value=30, value=10)

    st.markdown("---")
    st.markdown("### 📋 Issue Filter")
    show_duplicates     = st.checkbox("Duplicates",      value=True)
    show_mismatch       = st.checkbox("GL Mismatch",     value=True)
    show_missing        = st.checkbox("Missing Data",    value=True)
    show_uncategorized  = st.checkbox("Uncategorized",   value=True)
    show_clean          = st.checkbox("Clean Entries",   value=False)

    st.markdown("---")
    st.caption("GL Audit v1.0 · Local-first · No data leaves your machine")

# ─── Upload ─────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload your General Ledger (CSV or Excel)",
    type=["csv", "xlsx", "xls"],
    help="The file should contain columns like Date, Vendor, Amount, GL Account, Description"
)

if uploaded_file:
    # Load data
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ Could not read file: {e}")
        st.stop()

    st.success(f"✅ Loaded **{len(df):,}** transactions from `{uploaded_file.name}`")

    # Preview
    with st.expander("📄 Raw Data Preview", expanded=False):
        st.dataframe(df.head(20), use_container_width=True)

    # Column Mapping
    st.markdown("### 🗂️ Column Mapping")
    cols = ["— select —"] + list(df.columns)
    c1, c2, c3, c4, c5 = st.columns(5)
    col_date    = c1.selectbox("Date",        cols, index=1 if len(cols)>1 else 0)
    col_vendor  = c2.selectbox("Vendor",      cols, index=2 if len(cols)>2 else 0)
    col_amount  = c3.selectbox("Amount",      cols, index=3 if len(cols)>3 else 0)
    col_gl      = c4.selectbox("GL Account",  cols, index=4 if len(cols)>4 else 0)
    col_desc    = c5.selectbox("Description", cols, index=5 if len(cols)>5 else 0)

    # Run Audit
    st.markdown("---")
    if st.button("🚀 Run Forensic Audit", type="primary", use_container_width=True):

        col_map = {
            "date": col_date, "vendor": col_vendor,
            "amount": col_amount, "gl_account": col_gl, "description": col_desc
        }
        # Remove unmapped columns
        col_map = {k: v for k, v in col_map.items() if v != "— select —"}

        # Rename df columns to standard names
        df_work = df.rename(columns={v: k for k, v in col_map.items()})

        progress_bar  = st.progress(0)
        status_text   = st.empty()
        results_store = []

        num_batches = (len(df_work) // batch_size) + (1 if len(df_work) % batch_size else 0)

        for i in range(num_batches):
            batch_df  = df_work.iloc[i*batch_size : (i+1)*batch_size]
            offset    = i * batch_size
            status_text.markdown(f"🔄 Analysing batch **{i+1}/{num_batches}** (rows {offset+1}–{offset+len(batch_df)})…")

            try:
                batch_results = run_audit(
                    batch_df, offset,
                    model=ollama_model,
                    base_url=ollama_url
                )
                results_store.extend(batch_results)
            except Exception as e:
                st.warning(f"⚠️ Batch {i+1} failed: {e}")

            progress_bar.progress((i+1) / num_batches)
            time.sleep(0.1)  # small visual pause

        status_text.markdown("✅ Audit complete!")

        # ─── Store results ────────────────────────────────────────────────
        st.session_state["audit_results"] = results_store
        st.session_state["df_work"]       = df_work

# ─── Results Display ─────────────────────────────────────────────────────────
if "audit_results" in st.session_state:
    results = st.session_state["audit_results"]
    df_work = st.session_state["df_work"]

    stats = generate_summary_stats(results)

    # KPI Row
    st.markdown("### 📊 Audit Summary")
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f'<div class="metric-card"><p>Total Transactions</p><h3 style="color:#00d4ff">{stats["total"]}</h3></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="metric-card"><p>Duplicates</p><h3 style="color:#ff4d6d">{stats["duplicates"]}</h3></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="metric-card"><p>GL Mismatches</p><h3 style="color:#f7931a">{stats["mismatches"]}</h3></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="metric-card"><p>Missing Data</p><h3 style="color:#c39bd3">{stats["missing"]}</h3></div>', unsafe_allow_html=True)
    with k5:
        st.markdown(f'<div class="metric-card"><p>Uncategorized</p><h3 style="color:#1abc9c">{stats["uncategorized"]}</h3></div>', unsafe_allow_html=True)

    # Filter
    issue_filter = []
    if show_duplicates:    issue_filter.append("Duplicate")
    if show_mismatch:      issue_filter.append("GL Mismatch")
    if show_missing:       issue_filter.append("Missing Data")
    if show_uncategorized: issue_filter.append("Uncategorized")
    if show_clean:         issue_filter.append("None")

    filtered = [r for r in results if r.get("issue","None") in issue_filter]

    # Results table
    st.markdown(f"### 🗃️ Detailed Results ({len(filtered)} entries shown)")

    badge_map = {
        "Duplicate":    "badge-duplicate",
        "GL Mismatch":  "badge-mismatch",
        "Missing Data": "badge-missing",
        "Uncategorized":"badge-uncategorized",
        "None":         "badge-none"
    }

    results_df = pd.DataFrame(filtered)
    if not results_df.empty:
        # Pretty display
        def make_badge(issue):
            cls = badge_map.get(issue, "badge-none")
            return f'<span class="{cls}">{issue}</span>'

        display_df = results_df.copy()
        if "issue" in display_df.columns:
            display_df["issue"] = display_df["issue"].apply(make_badge)

        st.write(results_df)  # plain table for interactivity

        # Download
        st.markdown("---")
        st.markdown("### 💾 Export")
        col_dl1, col_dl2 = st.columns(2)

        with col_dl1:
            csv_bytes = results_df.to_csv(index=False).encode()
            st.download_button("⬇️ Download CSV Report", csv_bytes, "gl_audit_results.csv", "text/csv", use_container_width=True)

        with col_dl2:
            json_bytes = json.dumps({"audit_results": filtered}, indent=2).encode()
            st.download_button("⬇️ Download JSON Report", json_bytes, "gl_audit_results.json", "application/json", use_container_width=True)

    else:
        st.info("No results match the current filters.")

    # Raw JSON viewer
    with st.expander("🔩 Raw JSON Output"):
        st.json({"audit_results": results[:50]})

else:
    # Landing state
    st.markdown("""
    <div style="text-align:center; padding: 4rem 2rem; color: #334466;">
      <div style="font-size:4rem">📂</div>
      <h3 style="color:#5577aa; font-weight:700">Upload a GL file to begin</h3>
      <p style="font-family:'IBM Plex Mono',monospace; font-size:0.85rem">
        Supports CSV · XLSX · XLS<br>
        All processing is local — no data is sent to external servers
      </p>
    </div>
    """, unsafe_allow_html=True)
