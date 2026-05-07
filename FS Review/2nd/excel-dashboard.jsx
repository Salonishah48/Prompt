import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── THEME ───────────────────────────────────────────────────────────────────
const T = {
  bg: "#070B14",
  sidebar: "#0C1220",
  panel: "#0F1829",
  card: "#141E30",
  border: "#1E2D45",
  accent: "#3B82F6",
  accentGlow: "#3B82F633",
  teal: "#06B6D4",
  green: "#10B981",
  amber: "#F59E0B",
  red: "#EF4444",
  purple: "#8B5CF6",
  text: "#E2E8F0",
  muted: "#4B6280",
  dim: "#2A3F5F",
};

// ─── ICONS ───────────────────────────────────────────────────────────────────
const Icon = ({ d, size = 16, stroke = 1.8, fill = "none", color = "currentColor" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill={fill} stroke={color} strokeWidth={stroke} strokeLinecap="round" strokeLinejoin="round">
    {Array.isArray(d) ? d.map((p, i) => <path key={i} d={p} />) : <path d={d} />}
  </svg>
);

const Icons = {
  dashboard: <Icon d="M3 3h7v7H3zM14 3h7v7h-7zM3 14h7v7H3zM14 14h7v7h-7z" />,
  upload: <Icon d={["M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4", "M17 8l-5-5-5 5M12 3v12"]} />,
  prompt: <Icon d={["M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"]} />,
  results: <Icon d={["M9 11l3 3L22 4", "M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"]} />,
  file: <Icon d={["M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z", "M14 2v6h6"]} />,
  trash: <Icon d={["M3 6h18", "M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6", "M8 6V4a1 1 0 011-1h6a1 1 0 011 1v2"]} />,
  plus: <Icon d={["M12 5v14", "M5 12h14"]} />,
  close: <Icon d={["M18 6L6 18", "M6 6l12 12"]} />,
  spark: <Icon d="M13 2L3 14h9l-1 8 10-12h-9l1-8z" />,
  copy: <Icon d={["M8 17H5a2 2 0 01-2-2V5a2 2 0 012-2h10a2 2 0 012 2v3", "M9 11h10a2 2 0 012 2v7a2 2 0 01-2 2H9a2 2 0 01-2-2v-7a2 2 0 012-2z"]} />,
  check: <Icon d="M20 6L9 17l-5-5" />,
  warning: <Icon d={["M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z", "M12 9v4M12 17h.01"]} />,
  chart: <Icon d={["M18 20V10", "M12 20V4", "M6 20v-6"]} />,
  list: <Icon d={["M8 6h13", "M8 12h13", "M8 18h13", "M3 6h.01M3 12h.01M3 18h.01"]} />,
  settings: <Icon d={["M12 15a3 3 0 100-6 3 3 0 000 6z", "M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"]} />,
  info: <Icon d={["M12 22a10 10 0 100-20 10 10 0 000 20z", "M12 8h.01M12 12v4"]} />,
};

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const readExcel = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });
      let allData = [], totalRows = 0, totalCols = 0;
      wb.SheetNames.forEach(name => {
        const ws = wb.Sheets[name];
        const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (json.length > 1) {
          const headers = json[0];
          totalCols = Math.max(totalCols, headers.length);
          totalRows += json.length - 1;
          allData.push(`--- Sheet: "${name}" (${json.length - 1} rows, ${headers.length} cols) ---\nColumns: ${headers.join(", ")}\nSample:\n${json.slice(1, 20).map(r => r.join(" | ")).join("\n")}`);
        }
      });
      resolve({ name: file.name, sheets: wb.SheetNames.length, rows: totalRows, cols: totalCols, preview: allData.join("\n\n"), raw: file, id: Date.now() + Math.random() });
    } catch (err) { reject(err); }
  };
  reader.onerror = reject;
  reader.readAsArrayBuffer(file);
});

function parseAIResult(raw) {
  const out = { summary: "", recommendations: [], openPoints: [], findings: [], raw };
  const lines = raw.split("\n").map(l => l.trim()).filter(Boolean);
  let cur = "findings";
  for (const line of lines) {
    const low = line.toLowerCase();
    if (/^#+\s*(recommendation|suggest)/i.test(line) || /^recommendation/i.test(line)) { cur = "rec"; continue; }
    if (/^#+\s*(open|action|issue)/i.test(line) || /^open\s*(point|issue|action)/i.test(line)) { cur = "open"; continue; }
    if (/^#+\s*(summary|overview|executive)/i.test(line) || /^summary/i.test(line)) { cur = "sum"; continue; }
    if (/^#+\s*(finding|insight|key)/i.test(line) || /^(finding|key insight)/i.test(line)) { cur = "findings"; continue; }
    const clean = line.replace(/^[-•*\d.]+\s*/, "").trim();
    if (!clean || clean.length < 5) continue;
    if (cur === "rec") out.recommendations.push(clean);
    else if (cur === "open") out.openPoints.push(clean);
    else if (cur === "sum") out.summary += (out.summary ? " " : "") + clean;
    else out.findings.push(clean);
  }
  if (!out.summary && out.findings.length) { out.summary = out.findings.slice(0, 3).join(" "); out.findings = out.findings.slice(3); }
  return out;
}

// ─── MINI COMPONENTS ─────────────────────────────────────────────────────────
const Pill = ({ color, children }) => (
  <span style={{ background: color + "20", color, border: `1px solid ${color}40`, borderRadius: 12, padding: "2px 8px", fontSize: 10, fontWeight: 700, letterSpacing: "0.05em" }}>{children}</span>
);

const StatCard = ({ label, value, color, icon, delay = 0 }) => (
  <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 12, padding: "16px 18px", display: "flex", gap: 14, alignItems: "center", animation: `slideUp .4s ease ${delay}s both` }}>
    <div style={{ background: color + "18", border: `1px solid ${color}30`, borderRadius: 10, padding: 10, color, flexShrink: 0 }}>{icon}</div>
    <div>
      <div style={{ color: T.muted, fontSize: 11, fontWeight: 600, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 3 }}>{label}</div>
      <div style={{ color: T.text, fontSize: 22, fontWeight: 800, fontFamily: "'Syne', sans-serif" }}>{value}</div>
    </div>
  </div>
);

function ProgressBar({ value, color }) {
  return (
    <div style={{ background: T.border, borderRadius: 4, height: 4, overflow: "hidden" }}>
      <div style={{ width: `${Math.min(100, value)}%`, height: "100%", background: color, borderRadius: 4, transition: "width 1s ease" }} />
    </div>
  );
}

function InsightCard({ item, index, color, icon }) {
  const [visible, setVisible] = useState(false);
  useEffect(() => { setTimeout(() => setVisible(true), index * 80); }, []);
  return (
    <div style={{ display: "flex", gap: 12, padding: "12px 0", borderBottom: `1px solid ${T.border}`, opacity: visible ? 1 : 0, transform: visible ? "none" : "translateX(-8px)", transition: "all .3s ease" }}>
      <div style={{ color, marginTop: 1, flexShrink: 0 }}>{icon}</div>
      <p style={{ color: T.text, fontSize: 13, lineHeight: 1.65, margin: 0 }}>{item}</p>
    </div>
  );
}

// ─── SIDEBAR ─────────────────────────────────────────────────────────────────
const NAV = [
  { id: "upload", label: "Data Sources", icon: Icons.upload },
  { id: "prompts", label: "Prompts", icon: Icons.prompt },
  { id: "results", label: "Results", icon: Icons.results },
];

function Sidebar({ active, setActive, fileCount, hasResults }) {
  return (
    <div style={{ width: 220, background: T.sidebar, borderRight: `1px solid ${T.border}`, display: "flex", flexDirection: "column", flexShrink: 0 }}>
      {/* Logo */}
      <div style={{ padding: "20px 18px 16px", borderBottom: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
          <div style={{ background: `linear-gradient(135deg, ${T.accent}, ${T.teal})`, borderRadius: 8, padding: 7, color: "white", display: "flex" }}>
            {Icons.spark}
          </div>
          <div>
            <div style={{ color: T.text, fontWeight: 800, fontSize: 14, fontFamily: "'Syne', sans-serif", letterSpacing: "-0.3px" }}>DataLens</div>
            <div style={{ color: T.muted, fontSize: 10, letterSpacing: "0.05em" }}>AI ANALYZER</div>
          </div>
        </div>
      </div>

      {/* Nav */}
      <nav style={{ padding: "12px 10px", flex: 1 }}>
        {NAV.map(n => {
          const isActive = active === n.id;
          const disabled = n.id === "results" && !hasResults;
          return (
            <button key={n.id} onClick={() => !disabled && setActive(n.id)} disabled={disabled}
              style={{
                width: "100%", display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", borderRadius: 8,
                background: isActive ? T.accent + "20" : "transparent",
                border: `1px solid ${isActive ? T.accent + "50" : "transparent"}`,
                color: disabled ? T.dim : isActive ? T.accent : T.muted,
                cursor: disabled ? "not-allowed" : "pointer", marginBottom: 3,
                transition: "all .18s", fontFamily: "inherit", fontSize: 13, fontWeight: isActive ? 600 : 400,
                textAlign: "left",
              }}
              onMouseEnter={e => { if (!disabled && !isActive) { e.currentTarget.style.background = T.dim + "30"; e.currentTarget.style.color = T.text; } }}
              onMouseLeave={e => { if (!isActive) { e.currentTarget.style.background = "transparent"; e.currentTarget.style.color = disabled ? T.dim : T.muted; } }}>
              <span style={{ flexShrink: 0 }}>{n.icon}</span>
              {n.label}
              {n.id === "upload" && fileCount > 0 && (
                <span style={{ marginLeft: "auto", background: T.accent, color: "white", borderRadius: 10, padding: "1px 7px", fontSize: 10, fontWeight: 700 }}>{fileCount}</span>
              )}
              {n.id === "results" && hasResults && (
                <span style={{ marginLeft: "auto", width: 6, height: 6, background: T.green, borderRadius: "50%" }} />
              )}
            </button>
          );
        })}
      </nav>

      <div style={{ padding: "12px 16px", borderTop: `1px solid ${T.border}` }}>
        <div style={{ color: T.dim, fontSize: 10, textAlign: "center" }}>Max 5 files · 50 rows/sheet</div>
      </div>
    </div>
  );
}

// ─── TOPBAR ──────────────────────────────────────────────────────────────────
function TopBar({ title, subtitle, action }) {
  return (
    <div style={{ padding: "16px 28px", borderBottom: `1px solid ${T.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", background: T.panel, flexShrink: 0 }}>
      <div>
        <h1 style={{ margin: 0, color: T.text, fontSize: 17, fontWeight: 700, fontFamily: "'Syne', sans-serif" }}>{title}</h1>
        {subtitle && <p style={{ margin: 0, color: T.muted, fontSize: 12, marginTop: 2 }}>{subtitle}</p>}
      </div>
      {action}
    </div>
  );
}

// ─── UPLOAD PANEL ────────────────────────────────────────────────────────────
function UploadPanel({ files, setFiles }) {
  const [drag, setDrag] = useState(false);
  const ref = useRef();

  const handleFiles = useCallback(async (newFiles) => {
    const accepted = [...newFiles].filter(f => /\.(xlsx|xls|csv)$/i.test(f.name));
    const slots = 5 - files.length;
    if (!slots) return;
    const parsed = await Promise.all(accepted.slice(0, slots).map(readExcel));
    setFiles(prev => [...prev, ...parsed]);
  }, [files]);

  return (
    <div style={{ flex: 1, overflow: "auto", padding: 28 }}>
      <style>{`@keyframes slideUp{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}} @keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}} @keyframes spin{to{transform:rotate(360deg)}}`}</style>

      {/* Drop Zone */}
      <div onClick={() => files.length < 5 && ref.current?.click()}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={e => { e.preventDefault(); setDrag(false); handleFiles(e.dataTransfer.files); }}
        style={{
          border: `2px dashed ${drag ? T.accent : T.border}`, borderRadius: 14, padding: "36px 24px",
          textAlign: "center", cursor: files.length < 5 ? "pointer" : "default",
          background: drag ? T.accent + "08" : T.card, transition: "all .2s", marginBottom: 24,
          animation: "slideUp .4s ease",
        }}>
        <div style={{ display: "inline-flex", background: T.accent + "18", borderRadius: 12, padding: 14, color: T.accent, marginBottom: 12 }}>{Icons.upload}</div>
        <div style={{ color: files.length < 5 ? T.text : T.muted, fontWeight: 600, fontSize: 14, marginBottom: 4 }}>
          {files.length < 5 ? "Drop Excel / CSV files here" : "5 files loaded — limit reached"}
        </div>
        <div style={{ color: T.muted, fontSize: 12 }}>Supports .xlsx · .xls · .csv · up to 5 files</div>
        <input ref={ref} type="file" accept=".xlsx,.xls,.csv" multiple style={{ display: "none" }} onChange={e => handleFiles(e.target.files)} />
      </div>

      {/* File List */}
      {files.length > 0 && (
        <div>
          <div style={{ color: T.muted, fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 12 }}>Loaded Files</div>
          <div style={{ display: "grid", gap: 10 }}>
            {files.map((f, i) => (
              <div key={f.id} style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 12, padding: "14px 16px", display: "flex", gap: 14, alignItems: "center", animation: `slideUp .3s ease ${i * .06}s both` }}>
                <div style={{ background: T.accent + "18", borderRadius: 8, padding: 8, color: T.accent, flexShrink: 0 }}>{Icons.file}</div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ color: T.text, fontWeight: 600, fontSize: 13, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{f.name}</div>
                  <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
                    <Pill color={T.teal}>{f.sheets} sheet{f.sheets !== 1 ? "s" : ""}</Pill>
                    <Pill color={T.purple}>{f.rows.toLocaleString()} rows</Pill>
                    <Pill color={T.amber}>{f.cols} cols</Pill>
                  </div>
                </div>
                <button onClick={() => setFiles(p => p.filter(x => x.id !== f.id))}
                  style={{ background: "none", border: "none", color: T.muted, cursor: "pointer", padding: 6, borderRadius: 6, display: "flex", transition: "color .2s" }}
                  onMouseEnter={e => e.currentTarget.style.color = T.red}
                  onMouseLeave={e => e.currentTarget.style.color = T.muted}>
                  {Icons.trash}
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {files.length === 0 && (
        <div style={{ textAlign: "center", padding: "24px 0", color: T.muted, fontSize: 13 }}>
          No files loaded yet. Upload Excel or CSV files above.
        </div>
      )}
    </div>
  );
}

// ─── PROMPTS PANEL ───────────────────────────────────────────────────────────
const DEFAULT_PROMPTS = [
  "Identify key trends and patterns across all files",
  "List all open issues, risks, and action items",
  "Provide strategic recommendations based on the data",
  "Highlight anomalies, outliers, or data quality issues",
];

function PromptsPanel({ prompts, setPrompts, files, onAnalyze, loading }) {
  const [newP, setNewP] = useState("");
  const [error, setError] = useState("");

  const run = () => {
    if (!files.length) { setError("Upload at least one file first."); return; }
    if (!prompts.length) { setError("Add at least one prompt."); return; }
    setError("");
    onAnalyze();
  };

  return (
    <div style={{ flex: 1, overflow: "auto", padding: 28 }}>
      <div style={{ maxWidth: 720 }}>
        {/* Prompts list */}
        <div style={{ marginBottom: 20 }}>
          {prompts.map((p, i) => (
            <div key={i} style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "12px 14px", display: "flex", gap: 10, alignItems: "flex-start", marginBottom: 8, animation: `slideUp .3s ease ${i * .05}s both` }}>
              <div style={{ background: T.accent + "18", borderRadius: 6, padding: "4px 8px", color: T.accent, fontSize: 11, fontWeight: 700, flexShrink: 0, marginTop: 1 }}>{String(i + 1).padStart(2, "0")}</div>
              <p style={{ flex: 1, color: T.text, fontSize: 13, margin: 0, lineHeight: 1.5 }}>{p}</p>
              <button onClick={() => setPrompts(prev => prev.filter((_, j) => j !== i))}
                style={{ background: "none", border: "none", color: T.muted, cursor: "pointer", padding: 4, borderRadius: 4, display: "flex", transition: "color .2s", flexShrink: 0 }}
                onMouseEnter={e => e.currentTarget.style.color = T.red}
                onMouseLeave={e => e.currentTarget.style.color = T.muted}>
                {Icons.close}
              </button>
            </div>
          ))}
        </div>

        {/* Add Prompt */}
        <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 12, padding: 16, marginBottom: 20, animation: "slideUp .4s ease .2s both" }}>
          <div style={{ color: T.muted, fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 10 }}>Add Custom Prompt</div>
          <div style={{ display: "flex", gap: 8 }}>
            <textarea value={newP} onChange={e => setNewP(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey && newP.trim()) { e.preventDefault(); setPrompts(p => [...p, newP.trim()]); setNewP(""); } }}
              placeholder="e.g. Compare performance metrics across all files..."
              rows={2}
              style={{ flex: 1, background: T.panel, border: `1px solid ${T.border}`, borderRadius: 8, padding: "10px 12px", color: T.text, fontSize: 13, outline: "none", resize: "vertical", fontFamily: "inherit", lineHeight: 1.5, transition: "border-color .2s" }}
              onFocus={e => e.target.style.borderColor = T.accent}
              onBlur={e => e.target.style.borderColor = T.border}
            />
            <button onClick={() => { if (newP.trim()) { setPrompts(p => [...p, newP.trim()]); setNewP(""); } }}
              style={{ background: T.accent + "20", border: `1px solid ${T.accent}50`, color: T.accent, borderRadius: 8, padding: "0 14px", cursor: "pointer", display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600, fontFamily: "inherit", transition: "all .2s", flexShrink: 0 }}
              onMouseEnter={e => { e.currentTarget.style.background = T.accent + "35"; }}
              onMouseLeave={e => { e.currentTarget.style.background = T.accent + "20"; }}>
              {Icons.plus} Add
            </button>
          </div>
          <div style={{ color: T.muted, fontSize: 11, marginTop: 7 }}>Press Enter or click Add · Shift+Enter for new line</div>
        </div>

        {error && (
          <div style={{ background: "#1a0808", border: `1px solid ${T.red}40`, borderRadius: 8, padding: "10px 14px", color: T.red, fontSize: 12, display: "flex", gap: 8, alignItems: "center", marginBottom: 16 }}>
            {Icons.warning} {error}
          </div>
        )}

        {/* Run Button */}
        <button onClick={run} disabled={loading}
          style={{
            width: "100%", padding: "15px", borderRadius: 12, border: "none",
            background: loading ? T.card : `linear-gradient(135deg, ${T.accent} 0%, ${T.teal} 100%)`,
            color: loading ? T.muted : "white", fontWeight: 700, fontSize: 14, fontFamily: "'Syne', sans-serif",
            cursor: loading ? "wait" : "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 10,
            transition: "all .25s", letterSpacing: "0.02em",
            boxShadow: loading ? "none" : `0 4px 24px ${T.accent}44`,
          }}>
          {loading ? (
            <>
              <div style={{ width: 16, height: 16, border: `2px solid ${T.dim}`, borderTopColor: T.teal, borderRadius: "50%", animation: "spin 1s linear infinite" }} />
              Analyzing {files.length} file{files.length !== 1 ? "s" : ""}...
            </>
          ) : (
            <>{Icons.spark} Run AI Analysis &nbsp;·&nbsp; {files.length} file{files.length !== 1 ? "s" : ""}, {prompts.length} prompt{prompts.length !== 1 ? "s" : ""}</>
          )}
        </button>
      </div>
    </div>
  );
}

// ─── RESULTS PANEL ───────────────────────────────────────────────────────────
function ResultsPanel({ result, files }) {
  const [tab, setTab] = useState("overview");
  const [copied, setCopied] = useState(false);
  if (!result) return (
    <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 14, color: T.muted, padding: 40 }}>
      <div style={{ opacity: .3 }}>{Icons.chart}</div>
      <div style={{ fontSize: 14 }}>Run an analysis to see results here</div>
    </div>
  );

  const p = result.parsed;
  const total = p.recommendations.length + p.openPoints.length + p.findings.length;

  const TABS = [
    { id: "overview", label: "Overview" },
    { id: "rec", label: `Recommendations (${p.recommendations.length})` },
    { id: "open", label: `Open Points (${p.openPoints.length})` },
    { id: "findings", label: `Findings (${p.findings.length})` },
    { id: "raw", label: "Raw" },
  ];

  return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
      {/* Stats Row */}
      <div style={{ padding: "20px 28px 0", flexShrink: 0 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 20 }}>
          <StatCard label="Files Analyzed" value={files.length} color={T.accent} icon={Icons.file} delay={0} />
          <StatCard label="Recommendations" value={p.recommendations.length} color={T.green} icon={Icons.check} delay={.05} />
          <StatCard label="Open Points" value={p.openPoints.length} color={T.amber} icon={Icons.warning} delay={.1} />
          <StatCard label="Key Findings" value={p.findings.length} color={T.purple} icon={Icons.chart} delay={.15} />
        </div>

        {/* Composition Bar */}
        {total > 0 && (
          <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: "12px 16px", marginBottom: 16, display: "flex", alignItems: "center", gap: 16, animation: "slideUp .4s ease .2s both" }}>
            <div style={{ flex: 1 }}>
              <div style={{ display: "flex", height: 6, borderRadius: 4, overflow: "hidden", gap: 1 }}>
                {[{ v: p.recommendations.length, c: T.green }, { v: p.openPoints.length, c: T.amber }, { v: p.findings.length, c: T.purple }]
                  .map((s, i) => <div key={i} style={{ width: `${(s.v / total) * 100}%`, background: s.c, transition: "width 1s ease" }} />)}
              </div>
            </div>
            <div style={{ display: "flex", gap: 12, flexShrink: 0 }}>
              {[["Recs", T.green, p.recommendations.length], ["Open", T.amber, p.openPoints.length], ["Findings", T.purple, p.findings.length]].map(([l, c, v]) => (
                <div key={l} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11, color: T.muted }}>
                  <span style={{ width: 8, height: 8, background: c, borderRadius: 2, display: "inline-block" }} />{l}: <b style={{ color: c }}>{v}</b>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Tabs */}
        <div style={{ display: "flex", gap: 4, borderBottom: `1px solid ${T.border}`, paddingBottom: 0 }}>
          {TABS.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)}
              style={{
                background: "none", border: "none", borderBottom: `2px solid ${tab === t.id ? T.accent : "transparent"}`,
                color: tab === t.id ? T.accent : T.muted, cursor: "pointer", padding: "8px 14px", fontSize: 12, fontWeight: tab === t.id ? 600 : 400,
                fontFamily: "inherit", transition: "all .2s", marginBottom: -1,
              }}>
              {t.label}
            </button>
          ))}
        </div>
      </div>

      {/* Tab Content */}
      <div style={{ flex: 1, overflow: "auto", padding: "20px 28px" }}>
        {tab === "overview" && (
          <div style={{ animation: "slideUp .3s ease" }}>
            <div style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 12, padding: 20, marginBottom: 16 }}>
              <div style={{ color: T.muted, fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 10 }}>Executive Summary</div>
              <p style={{ color: T.text, fontSize: 14, lineHeight: 1.75, margin: 0 }}>{p.summary || "Analysis complete. See individual tabs for details."}</p>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
              {[
                { title: "Top Recommendations", items: p.recommendations.slice(0, 3), color: T.green, icon: Icons.check },
                { title: "Top Open Points", items: p.openPoints.slice(0, 3), color: T.amber, icon: Icons.warning },
              ].map(({ title, items, color, icon }) => (
                <div key={title} style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 12, padding: 18 }}>
                  <div style={{ color: T.muted, fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 10 }}>{title}</div>
                  {items.length ? items.map((it, i) => (
                    <div key={i} style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                      <span style={{ color, marginTop: 2, flexShrink: 0 }}>{icon}</span>
                      <span style={{ color: T.text, fontSize: 12, lineHeight: 1.55 }}>{it}</span>
                    </div>
                  )) : <span style={{ color: T.muted, fontSize: 12 }}>None identified</span>}
                </div>
              ))}
            </div>
          </div>
        )}

        {tab === "rec" && (
          <div style={{ animation: "slideUp .3s ease" }}>
            {p.recommendations.length ? p.recommendations.map((r, i) => (
              <InsightCard key={i} item={r} index={i} color={T.green} icon={Icons.check} />
            )) : <div style={{ color: T.muted, fontSize: 13, padding: "20px 0" }}>No specific recommendations found.</div>}
          </div>
        )}

        {tab === "open" && (
          <div style={{ animation: "slideUp .3s ease" }}>
            {p.openPoints.length ? p.openPoints.map((o, i) => (
              <InsightCard key={i} item={o} index={i} color={T.amber} icon={Icons.warning} />
            )) : <div style={{ color: T.muted, fontSize: 13, padding: "20px 0" }}>No open points identified.</div>}
          </div>
        )}

        {tab === "findings" && (
          <div style={{ animation: "slideUp .3s ease" }}>
            {p.findings.length ? p.findings.map((f, i) => (
              <InsightCard key={i} item={f} index={i} color={T.purple} icon={Icons.list} />
            )) : <div style={{ color: T.muted, fontSize: 13, padding: "20px 0" }}>No additional findings.</div>}
          </div>
        )}

        {tab === "raw" && (
          <div style={{ animation: "slideUp .3s ease" }}>
            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
              <button onClick={() => { navigator.clipboard?.writeText(result.raw); setCopied(true); setTimeout(() => setCopied(false), 1500); }}
                style={{ background: T.card, border: `1px solid ${T.border}`, color: T.muted, borderRadius: 7, padding: "6px 12px", fontSize: 12, cursor: "pointer", display: "flex", alignItems: "center", gap: 6, fontFamily: "inherit", transition: "all .2s" }}
                onMouseEnter={e => { e.currentTarget.style.color = T.teal; e.currentTarget.style.borderColor = T.teal; }}
                onMouseLeave={e => { e.currentTarget.style.color = T.muted; e.currentTarget.style.borderColor = T.border; }}>
                {copied ? Icons.check : Icons.copy} {copied ? "Copied!" : "Copy"}
              </button>
            </div>
            <pre style={{ background: T.card, border: `1px solid ${T.border}`, borderRadius: 10, padding: 18, color: T.text, fontSize: 12, lineHeight: 1.7, whiteSpace: "pre-wrap", wordBreak: "break-word", margin: 0 }}>
              {result.raw}
            </pre>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── APP ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage] = useState("upload");
  const [files, setFiles] = useState([]);
  const [prompts, setPrompts] = useState([...DEFAULT_PROMPTS]);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);

  const analyze = async () => {
    setLoading(true);
    const dataCtx = files.map(f => `\n=== ${f.name} ===\n${f.preview}`).join("\n\n");
    const promptList = prompts.map((p, i) => `${i + 1}. ${p}`).join("\n");
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          system: `You are an expert data analyst. Structure your response with these exact section headers:\n\nSUMMARY:\n(2-3 sentence executive summary)\n\nRECOMMENDATIONS:\n- (each as a bullet)\n\nOPEN POINTS:\n- (each as a bullet)\n\nKEY FINDINGS:\n- (each as a bullet)\n\nBe specific and reference actual data values.`,
          messages: [{ role: "user", content: `Analyze this data from ${files.length} Excel file(s):\n\n${dataCtx}\n\nAnswer:\n${promptList}` }],
        }),
      });
      const data = await res.json();
      const text = data.content?.map(c => c.text || "").join("") || "";
      setResult({ raw: text, parsed: parseAIResult(text) });
      setPage("results");
    } catch (e) {
      console.error(e);
    }
    setLoading(false);
  };

  const PAGE_TITLES = {
    upload: { title: "Data Sources", subtitle: "Upload up to 5 Excel or CSV files for analysis" },
    prompts: { title: "Analysis Prompts", subtitle: "Define what the AI should investigate in your data" },
    results: { title: "Analysis Results", subtitle: result ? `${files.length} file${files.length !== 1 ? "s" : ""} · ${result.parsed.recommendations.length + result.parsed.openPoints.length + result.parsed.findings.length} insights` : "Run an analysis to see results" },
  };

  return (
    <div style={{ display: "flex", height: "100vh", background: T.bg, color: T.text, fontFamily: "'DM Sans', 'Segoe UI', sans-serif", overflow: "hidden" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&display=swap');
        @keyframes slideUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
        @keyframes spin{to{transform:rotate(360deg)}}
        * { box-sizing: border-box; }
        ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:${T.dim};border-radius:4px}
      `}</style>

      <Sidebar active={page} setActive={setPage} fileCount={files.length} hasResults={!!result} />

      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        <TopBar
          title={PAGE_TITLES[page].title}
          subtitle={PAGE_TITLES[page].subtitle}
          action={
            page !== "results" ? (
              <button onClick={() => setPage(page === "upload" ? "prompts" : "upload")}
                style={{ background: T.card, border: `1px solid ${T.border}`, color: T.muted, borderRadius: 8, padding: "8px 16px", fontSize: 12, cursor: "pointer", fontFamily: "inherit", transition: "all .2s" }}
                onMouseEnter={e => { e.currentTarget.style.color = T.text; e.currentTarget.style.borderColor = T.dim; }}
                onMouseLeave={e => { e.currentTarget.style.color = T.muted; e.currentTarget.style.borderColor = T.border; }}>
                {page === "upload" ? "→ Go to Prompts" : "← Back to Files"}
              </button>
            ) : (
              <button onClick={() => { setResult(null); setPage("upload"); setFiles([]); }}
                style={{ background: T.card, border: `1px solid ${T.border}`, color: T.muted, borderRadius: 8, padding: "8px 16px", fontSize: 12, cursor: "pointer", fontFamily: "inherit", transition: "all .2s" }}
                onMouseEnter={e => { e.currentTarget.style.color = T.red; e.currentTarget.style.borderColor = T.red + "60"; }}
                onMouseLeave={e => { e.currentTarget.style.color = T.muted; e.currentTarget.style.borderColor = T.border; }}>
                New Analysis
              </button>
            )
          }
        />

        {page === "upload" && <UploadPanel files={files} setFiles={setFiles} />}
        {page === "prompts" && <PromptsPanel prompts={prompts} setPrompts={setPrompts} files={files} onAnalyze={analyze} loading={loading} />}
        {page === "results" && <ResultsPanel result={result} files={files} />}
      </div>
    </div>
  );
}
