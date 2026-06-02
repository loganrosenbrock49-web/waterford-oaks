import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Brand Colors (Waterford Oaks: forest green, cream, warm white) ───────
const C = {
  green:      "#1B5E2B",
  greenMid:   "#2E7D3E",
  greenLight: "#4A9B5A",
  greenPale:  "#E8F2EA",
  greenTint:  "#F2F7F3",
  gold:       "#8B6914",
  goldLight:  "#C49A2A",
  goldPale:   "#FBF5E6",
  cream:      "#FDFAF5",
  white:      "#FFFFFF",
  charcoal:   "#1A1A1A",
  dark:       "#2C2C2C",
  mid:        "#555555",
  muted:      "#888888",
  border:     "#D8E4DA",
  borderDark: "#B8CEBC",
  red:        "#B33A3A",
  redPale:    "#FBF0F0",
  blue:       "#1E4D8C",
  bluePale:   "#EEF3FA",
};

const font = `'Lora', 'Georgia', serif`;
const sans = `'Source Sans 3', 'Segoe UI', sans-serif`;

// ─── Persistent storage helpers ───────────────────────────────────────────
const STORE_KEYS = {
  positions: "wo_positions",
  applicants: "wo_applicants",
  hireHistory: "wo_hire_history",
};

function loadStore(key, fallback) {
  try {
    const r = localStorage.getItem(key);
    return r ? JSON.parse(r) : fallback;
  } catch { return fallback; }
}
function saveStore(key, val) {
  try { localStorage.setItem(key, JSON.stringify(val)); } catch {}
}

// ─── Default job positions ─────────────────────────────────────────────────
const DEFAULT_POSITIONS = [
  {
    id: "pos1", title: "Care Staff",
    description: "Provide hands-on personal care to residents including bathing, dressing, grooming, and mobility assistance. Monitor resident health and report changes to nursing staff. Build warm relationships with residents and families.",
    requirements: "CNA certification preferred. 1+ year experience in senior care. Compassionate, reliable, team player.",
    schedule: "Full-time / Part-time", pay: "$16–$20/hr",
  },
  {
    id: "pos2", title: "Medication Aide",
    description: "Administer medications to residents per physician orders. Maintain accurate medication records and observe for side effects. Communicate with nursing staff and pharmacy.",
    requirements: "Certified Medication Aide (CMA) required. 2+ years experience. Detail-oriented.",
    schedule: "Full-time", pay: "$18–$22/hr",
  },
  {
    id: "pos3", title: "Activities Director",
    description: "Plan and execute engaging daily activities for Memory Care and Assisted Living residents. Coordinate volunteers, outings, and family events. Maintain activity documentation.",
    requirements: "Activities certification preferred. Creative, energetic, experience with seniors.",
    schedule: "Full-time", pay: "$20–$26/hr",
  },
  {
    id: "pos4", title: "Dining Staff",
    description: "Serve meals to residents in a warm, dignified manner. Assist with special dietary needs and tray setup. Maintain a clean and welcoming dining environment.",
    requirements: "Food handler's permit. Prior serving or kitchen experience a plus.",
    schedule: "Part-time / Weekends", pay: "$14–$17/hr",
  },
];

const HIRING_STAGES = [
  { id: "applied",     label: "Applied",              color: C.muted,      icon: "📄" },
  { id: "reviewed",    label: "App Reviewed",          color: C.blue,       icon: "🔍" },
  { id: "availability",label: "Interview Requested",   color: C.goldLight,  icon: "📅" },
  { id: "interviewed", label: "Interviewed",           color: C.greenMid,   icon: "🗣️" },
  { id: "fit",         label: "Good Fit ✓",            color: C.green,      icon: "✅" },
  { id: "notfit",      label: "Not a Fit",             color: C.red,        icon: "❌" },
];

// ─── File reading utility ─────────────────────────────────────────────────
async function readFileAsText(file) {
  const name = file.name.toLowerCase();
  // CSV / TSV / TXT / QBO / IIF (QuickBooks export)
  if (name.endsWith(".csv") || name.endsWith(".tsv") || name.endsWith(".txt") ||
      name.endsWith(".qbo") || name.endsWith(".iif") || name.endsWith(".qbx")) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = e => res(e.target.result);
      r.onerror = rej;
      r.readAsText(file);
    });
  }
  // Excel / XLSX / XLS / ODS
  if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".ods")) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = e => {
        try {
          const wb = XLSX.read(e.target.result, { type: "binary" });
          let out = "";
          wb.SheetNames.forEach(sn => {
            out += `=== Sheet: ${sn} ===\n`;
            out += XLSX.utils.sheet_to_csv(wb.Sheets[sn]) + "\n\n";
          });
          res(out);
        } catch(err) { rej(err); }
      };
      r.onerror = rej;
      r.readAsBinaryString(file);
    });
  }
  // PDF — extract text via basic reader
  if (name.endsWith(".pdf")) {
    return new Promise((res) => {
      const r = new FileReader();
      r.onload = e => {
        // Basic PDF text extraction: grab readable strings
        const arr = new Uint8Array(e.target.result);
        let str = "";
        for (let i = 0; i < arr.length; i++) {
          const c = arr[i];
          if (c >= 32 && c < 127) str += String.fromCharCode(c);
          else if (c === 10 || c === 13) str += "\n";
        }
        // Extract readable chunks
        const chunks = str.match(/[A-Za-z0-9 ,.\-:/$%@()\n]{20,}/g) || [];
        res(chunks.join("\n").slice(0, 8000) || "[PDF content extracted — paste text manually if extraction is incomplete]");
      };
      r.readAsArrayBuffer(file);
    });
  }
  // DOCX / DOC — try mammoth
  if (name.endsWith(".docx") || name.endsWith(".doc")) {
    try {
      const mammoth = await import("https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js").catch(() => null);
      if (mammoth) {
        return new Promise((res, rej) => {
          const r = new FileReader();
          r.onload = async e => {
            try {
              const result = await mammoth.extractRawText({ arrayBuffer: e.target.result });
              res(result.value);
            } catch { res("[DOCX detected — paste text manually]"); }
          };
          r.onerror = rej;
          r.readAsArrayBuffer(file);
        });
      }
    } catch {}
    return "[DOCX file detected. For best results, paste the resume text directly.]";
  }
  // JSON
  if (name.endsWith(".json")) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = e => res(JSON.stringify(JSON.parse(e.target.result), null, 2));
      r.onerror = rej;
      r.readAsText(file);
    });
  }
  // Fallback: try as text
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => res(e.target.result);
    r.onerror = () => res("[File format not directly readable — paste content manually]");
    r.readAsText(file);
  });
}

// ─── Claude API ───────────────────────────────────────────────────────────
async function callClaude(system, user, onChunk, max_tokens = 1200) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens,
      stream: true,
      system,
      messages: [{ role: "user", content: user }],
    }),
  });
  const reader = res.body.getReader();
  const dec = new TextDecoder();
  let full = "";
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    for (const line of dec.decode(value).split("\n")) {
      if (line.startsWith("data: ")) {
        try {
          const d = JSON.parse(line.slice(6));
          if (d.type === "content_block_delta" && d.delta?.text) {
            full += d.delta.text;
            onChunk(full);
          }
        } catch {}
      }
    }
  }
  return full;
}

// ─── Shared UI components ─────────────────────────────────────────────────
function Spinner() {
  return (
    <div style={{ display: "flex", gap: 5, padding: "10px 0" }}>
      {[0,1,2].map(i => (
        <div key={i} style={{ width: 8, height: 8, borderRadius: "50%", background: C.greenMid,
          animation: `woPulse 1.2s ease-in-out ${i*0.2}s infinite` }} />
      ))}
    </div>
  );
}

function WOButton({ onClick, children, color, outline, small, disabled, style = {} }) {
  const bg = color || C.green;
  return (
    <button onClick={onClick} disabled={disabled} style={{
      display: "inline-flex", alignItems: "center", gap: 7,
      padding: small ? "7px 16px" : "11px 24px",
      borderRadius: 6,
      border: outline ? `2px solid ${bg}` : "none",
      background: outline ? "transparent" : bg,
      color: outline ? bg : "#fff",
      fontFamily: sans, fontWeight: 700, fontSize: small ? 13 : 14,
      cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.5 : 1,
      letterSpacing: "0.02em", transition: "all 0.15s", ...style,
    }}>{children}</button>
  );
}

function Card({ children, style = {}, green }) {
  return (
    <div style={{
      background: green ? C.greenTint : C.white,
      border: `1px solid ${green ? C.greenLight + "55" : C.border}`,
      borderRadius: 10, padding: 22, ...style,
    }}>{children}</div>
  );
}

function SectionLabel({ children }) {
  return (
    <div style={{ fontFamily: sans, fontSize: 11, fontWeight: 700, letterSpacing: "0.1em",
      textTransform: "uppercase", color: C.green, marginBottom: 8 }}>
      {children}
    </div>
  );
}

function WOTextarea({ value, onChange, placeholder, rows = 5 }) {
  return (
    <textarea value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} rows={rows}
      style={{ width: "100%", boxSizing: "border-box", padding: "11px 14px", fontFamily: sans,
        fontSize: 13.5, lineHeight: 1.65, color: C.dark, background: C.cream,
        border: `1.5px solid ${C.border}`, borderRadius: 7, resize: "vertical", outline: "none" }} />
  );
}

function FileDropZone({ onText, label }) {
  const [dragging, setDragging] = useState(false);
  const [fileName, setFileName] = useState("");
  const [loading, setLoading] = useState(false);
  const inputRef = useRef();

  const handle = async (file) => {
    if (!file) return;
    setFileName(file.name); setLoading(true);
    try {
      const text = await readFileAsText(file);
      onText(text, file.name);
    } catch(e) { onText("[Could not read file — paste content manually]", file.name); }
    setLoading(false);
  };

  return (
    <div
      onDragOver={e => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => { e.preventDefault(); setDragging(false); handle(e.dataTransfer.files[0]); }}
      onClick={() => inputRef.current.click()}
      style={{
        border: `2px dashed ${dragging ? C.green : C.borderDark}`,
        borderRadius: 9, padding: "20px 16px", textAlign: "center",
        background: dragging ? C.greenPale : C.greenTint,
        cursor: "pointer", transition: "all 0.2s",
      }}
    >
      <input ref={inputRef} type="file"
        accept=".pdf,.docx,.doc,.xlsx,.xls,.ods,.csv,.tsv,.txt,.json,.qbo,.iif,.qbx"
        style={{ display: "none" }}
        onChange={e => handle(e.target.files[0])}
      />
      {loading ? <Spinner /> : (
        <>
          <div style={{ fontSize: 28, marginBottom: 6 }}>📁</div>
          <div style={{ fontFamily: sans, fontWeight: 700, color: C.green, fontSize: 14 }}>
            {fileName || label || "Drop file or click to upload"}
          </div>
          <div style={{ fontFamily: sans, fontSize: 11.5, color: C.muted, marginTop: 4 }}>
            PDF, Word, Excel, CSV, QuickBooks (QBO/IIF), JSON, TXT
          </div>
        </>
      )}
    </div>
  );
}

function StreamOutput({ text, loading }) {
  if (!text && !loading) return null;
  return (
    <div style={{ background: C.cream, border: `1px solid ${C.border}`, borderRadius: 8,
      padding: "16px 18px", marginTop: 14 }}>
      {loading && !text && <Spinner />}
      <pre style={{ margin: 0, whiteSpace: "pre-wrap", wordBreak: "break-word",
        fontFamily: sans, fontSize: 13.5, color: C.dark, lineHeight: 1.7 }}>{text}</pre>
      {loading && text && <span style={{ display: "inline-block", width: 9, height: 15,
        background: C.green, marginLeft: 2, animation: "woBlink 1s step-end infinite",
        verticalAlign: "text-bottom" }} />}
    </div>
  );
}

function Badge({ color = C.green, bg, children, small }) {
  return (
    <span style={{ display: "inline-flex", alignItems: "center",
      padding: small ? "2px 8px" : "3px 11px", borderRadius: 20,
      fontSize: small ? 10.5 : 12, fontWeight: 700, letterSpacing: "0.04em",
      background: bg || color + "18", color,
      border: `1px solid ${color}44`, fontFamily: sans }}>
      {children}
    </span>
  );
}

// ─── RECRUITING MODULE ────────────────────────────────────────────────────
function RecruitingModule() {
  const [positions, setPositions] = useState(() => loadStore(STORE_KEYS.positions, DEFAULT_POSITIONS));
  const [applicants, setApplicants] = useState(() => loadStore(STORE_KEYS.applicants, []));
  const [hireHistory, setHireHistory] = useState(() => loadStore(STORE_KEYS.hireHistory, []));

  const [activeView, setActiveView] = useState("analyze"); // analyze | positions | tracker
  const [selectedPos, setSelectedPos] = useState(positions[0]?.id || "");
  const [resumeText, setResumeText] = useState("");
  const [applicantName, setApplicantName] = useState("");
  const [analysis, setAnalysis] = useState("");
  const [email, setEmail] = useState("");
  const [analysisLoading, setAnalysisLoading] = useState(false);
  const [emailLoading, setEmailLoading] = useState(false);
  const [priority, setPriority] = useState(null);

  // Position editor
  const [editingPos, setEditingPos] = useState(null);
  const [newPos, setNewPos] = useState({ title:"", description:"", requirements:"", schedule:"", pay:"" });
  const [showAddPos, setShowAddPos] = useState(false);

  useEffect(() => saveStore(STORE_KEYS.positions, positions), [positions]);
  useEffect(() => saveStore(STORE_KEYS.applicants, applicants), [applicants]);
  useEffect(() => saveStore(STORE_KEYS.hireHistory, hireHistory), [hireHistory]);

  const pos = positions.find(p => p.id === selectedPos);

  // Build learning context from hire history
  const buildLearningContext = () => {
    if (hireHistory.length === 0) return "";
    const good = hireHistory.filter(h => h.outcome === "fit").slice(-10);
    const bad = hireHistory.filter(h => h.outcome === "notfit").slice(-10);
    let ctx = "\n\nLEARNED FROM PREVIOUS HIRES AT WATERFORD OAKS:\n";
    if (good.length) ctx += `Successful hires had these traits: ${good.map(h => h.notes).join("; ")}\n`;
    if (bad.length) ctx += `Poor fits had these issues: ${bad.map(h => h.notes).join("; ")}\n`;
    return ctx;
  };

  const analyzeResume = async () => {
    if (!resumeText.trim() || !pos) return;
    setAnalysisLoading(true); setAnalysis(""); setPriority(null); setEmail("");
    const learn = buildLearningContext();
    const sys = `You are a senior HR recruiter for Waterford Oaks Senior Care, a Memory Care and Assisted Living community in Waterford. You evaluate candidates with care, warmth, and professional rigor.${learn}\n\nAlways end your analysis with exactly one line: PRIORITY: HIGH or PRIORITY: MEDIUM or PRIORITY: LOW`;
    const jd = `Position: ${pos.title}\nDescription: ${pos.description}\nRequirements: ${pos.requirements}\nSchedule: ${pos.schedule}\nPay: ${pos.pay}`;
    const prompt = `${jd}\n\nCandidate: ${applicantName || "Applicant"}\nResume/Application:\n${resumeText}\n\nProvide:\n1. CANDIDATE SUMMARY (2–3 sentences)\n2. STRENGTHS for this role\n3. GAPS or CONCERNS\n4. CULTURE FIT for senior care\n5. RECOMMENDATION\n\nEnd with PRIORITY: HIGH/MEDIUM/LOW`;
    const result = await callClaude(sys, prompt, t => setAnalysis(t));
    setAnalysisLoading(false);
    const p = result.includes("PRIORITY: HIGH") ? "HIGH" : result.includes("PRIORITY: MEDIUM") ? "MEDIUM" : "LOW";
    setPriority(p);
    // Add to tracker
    if (applicantName.trim()) {
      const newApp = {
        id: Date.now().toString(),
        name: applicantName.trim(),
        positionId: selectedPos,
        positionTitle: pos.title,
        stage: "reviewed",
        priority: p,
        date: new Date().toLocaleDateString(),
        notes: result.slice(0, 300),
      };
      setApplicants(prev => [newApp, ...prev]);
    }
  };

  const generateEmail = async () => {
    if (!pos) return;
    setEmailLoading(true); setEmail("");
    const sys = `You are the HR coordinator at Waterford Oaks Senior Care. Write warm, professional emails that reflect the caring culture of a senior living community.`;
    const prompt = `Write a concise interview invitation email for ${applicantName || "the candidate"} who applied for the ${pos.title} position at Waterford Oaks Senior Care. Under 150 words. Warm, professional. Invite them to share interview availability. Sign as "The Waterford Oaks Team".`;
    await callClaude(sys, prompt, t => setEmail(t));
    setEmailLoading(false);
    // Update stage
    setApplicants(prev => prev.map(a =>
      a.name === applicantName && a.positionTitle === pos?.title
        ? { ...a, stage: "availability" } : a
    ));
  };

  const updateStage = (appId, stage) => {
    setApplicants(prev => prev.map(a => a.id === appId ? { ...a, stage } : a));
    const app = applicants.find(a => a.id === appId);
    if ((stage === "fit" || stage === "notfit") && app) {
      setHireHistory(prev => [
        { outcome: stage, role: app.positionTitle, notes: app.notes?.slice(0, 200), date: new Date().toLocaleDateString() },
        ...prev.slice(0, 49),
      ]);
    }
  };

  const deleteApplicant = (id) => setApplicants(prev => prev.filter(a => a.id !== id));

  const savePosition = () => {
    if (!newPos.title.trim()) return;
    const p = { ...newPos, id: "pos_" + Date.now() };
    setPositions(prev => [...prev, p]);
    setNewPos({ title:"", description:"", requirements:"", schedule:"", pay:"" });
    setShowAddPos(false);
  };

  const updatePosition = (id, field, val) => {
    setPositions(prev => prev.map(p => p.id === id ? { ...p, [field]: val } : p));
  };

  const deletePosition = (id) => {
    setPositions(prev => prev.filter(p => p.id !== id));
    if (selectedPos === id) setSelectedPos(positions[0]?.id || "");
  };

  const pColor = priority === "HIGH" ? C.green : priority === "MEDIUM" ? C.goldLight : C.red;

  const tabStyle = (active) => ({
    padding: "8px 18px", borderRadius: "6px 6px 0 0", fontFamily: sans, fontWeight: 700,
    fontSize: 13, cursor: "pointer", border: "none",
    background: active ? C.green : "transparent",
    color: active ? "#fff" : C.mid,
    borderBottom: active ? `2px solid ${C.green}` : "2px solid transparent",
  });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {/* Sub-nav */}
      <div style={{ display: "flex", gap: 4, borderBottom: `2px solid ${C.border}` }}>
        {[["analyze","🔍 Analyze Resume"],["positions","📋 Job Descriptions"],["tracker","📊 Applicant Tracker"]].map(([v,l]) => (
          <button key={v} onClick={() => setActiveView(v)} style={tabStyle(activeView===v)}>{l}</button>
        ))}
      </div>

      {/* ── Analyze Tab ── */}
      {activeView === "analyze" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
            <div style={{ flex: 1, minWidth: 180 }}>
              <SectionLabel>Position</SectionLabel>
              <select value={selectedPos} onChange={e => setSelectedPos(e.target.value)}
                style={{ width: "100%", padding: "10px 12px", borderRadius: 7,
                  border: `1.5px solid ${C.border}`, fontFamily: sans, fontSize: 14,
                  color: C.dark, background: C.cream, outline: "none" }}>
                {positions.map(p => <option key={p.id} value={p.id}>{p.title}</option>)}
              </select>
            </div>
            <div style={{ flex: 1, minWidth: 180 }}>
              <SectionLabel>Applicant Name</SectionLabel>
              <input value={applicantName} onChange={e => setApplicantName(e.target.value)}
                placeholder="Jane Doe"
                style={{ width: "100%", padding: "10px 12px", borderRadius: 7, boxSizing: "border-box",
                  border: `1.5px solid ${C.border}`, fontFamily: sans, fontSize: 14,
                  color: C.dark, background: C.cream, outline: "none" }} />
            </div>
          </div>

          {pos && (
            <Card green>
              <div style={{ fontFamily: sans, fontWeight: 700, color: C.green, marginBottom: 4 }}>{pos.title}</div>
              <div style={{ fontFamily: sans, fontSize: 12.5, color: C.mid, lineHeight: 1.55 }}>{pos.description}</div>
              <div style={{ marginTop: 8, display: "flex", gap: 8, flexWrap: "wrap" }}>
                {pos.schedule && <Badge small>{pos.schedule}</Badge>}
                {pos.pay && <Badge small color={C.gold}>{pos.pay}</Badge>}
              </div>
            </Card>
          )}

          <div>
            <SectionLabel>Upload Resume / Application</SectionLabel>
            <FileDropZone onText={(t) => setResumeText(t)} label="Drop resume file or click to upload" />
          </div>
          <div>
            <SectionLabel>Or Paste Resume Text</SectionLabel>
            <WOTextarea value={resumeText} onChange={setResumeText} placeholder="Paste resume, LinkedIn profile, or application notes…" rows={6} />
          </div>

          <WOButton onClick={analyzeResume} disabled={analysisLoading || !resumeText.trim()}>
            {analysisLoading ? "Analyzing…" : "⚡ Analyze Candidate"}
          </WOButton>

          {(analysis || analysisLoading) && (
            <Card>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <div style={{ fontFamily: font, fontWeight: 700, fontSize: 16 }}>AI Analysis</div>
                {priority && <Badge color={pColor}>{priority} PRIORITY</Badge>}
              </div>
              {hireHistory.length > 0 && (
                <div style={{ fontFamily: sans, fontSize: 11.5, color: C.greenMid, marginBottom: 8,
                  background: C.greenPale, padding: "5px 10px", borderRadius: 5 }}>
                  🧠 Analysis enhanced by {hireHistory.length} previous hiring decisions
                </div>
              )}
              <StreamOutput text={analysis} loading={analysisLoading} />
              {!analysisLoading && analysis && priority !== "LOW" && (
                <div style={{ marginTop: 14 }}>
                  <WOButton onClick={generateEmail} disabled={emailLoading} color={C.greenMid} small>
                    {emailLoading ? "Drafting…" : "✉️ Generate Interview Email"}
                  </WOButton>
                </div>
              )}
            </Card>
          )}

          {(email || emailLoading) && (
            <Card style={{ borderLeft: `4px solid ${C.green}` }}>
              <div style={{ fontFamily: font, fontWeight: 700, fontSize: 15, color: C.green, marginBottom: 8 }}>
                📧 Outreach Email — {applicantName}
              </div>
              <StreamOutput text={email} loading={emailLoading} />
              {email && !emailLoading && (
                <WOButton onClick={() => navigator.clipboard.writeText(email)} outline small color={C.green} style={{ marginTop: 12 }}>
                  Copy Email
                </WOButton>
              )}
            </Card>
          )}
        </div>
      )}

      {/* ── Positions Tab ── */}
      {activeView === "positions" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          <div style={{ display: "flex", justifyContent: "flex-end" }}>
            <WOButton onClick={() => setShowAddPos(!showAddPos)} small color={C.green}>
              {showAddPos ? "Cancel" : "+ Add Position"}
            </WOButton>
          </div>

          {showAddPos && (
            <Card style={{ borderLeft: `4px solid ${C.greenLight}` }}>
              <div style={{ fontFamily: font, fontWeight: 700, fontSize: 15, marginBottom: 14 }}>New Position</div>
              {[["title","Position Title"],["pay","Pay Range"],["schedule","Schedule"]].map(([f,l]) => (
                <div key={f} style={{ marginBottom: 12 }}>
                  <SectionLabel>{l}</SectionLabel>
                  <input value={newPos[f]} onChange={e => setNewPos(p => ({...p, [f]: e.target.value}))}
                    placeholder={l} style={{ width: "100%", boxSizing: "border-box",
                      padding: "9px 12px", borderRadius: 7, border: `1.5px solid ${C.border}`,
                      fontFamily: sans, fontSize: 14, color: C.dark, background: C.cream, outline: "none" }} />
                </div>
              ))}
              <div style={{ marginBottom: 12 }}>
                <SectionLabel>Job Description</SectionLabel>
                <WOTextarea value={newPos.description} onChange={v => setNewPos(p=>({...p,description:v}))} placeholder="Describe responsibilities…" rows={4} />
              </div>
              <div style={{ marginBottom: 14 }}>
                <SectionLabel>Requirements</SectionLabel>
                <WOTextarea value={newPos.requirements} onChange={v => setNewPos(p=>({...p,requirements:v}))} placeholder="Certifications, experience, skills…" rows={3} />
              </div>
              <WOButton onClick={savePosition} color={C.green}>Save Position</WOButton>
            </Card>
          )}

          {positions.map(p => (
            <Card key={p.id} style={{ position: "relative" }}>
              {editingPos === p.id ? (
                <div>
                  {[["title","Title"],["pay","Pay"],["schedule","Schedule"]].map(([f,l]) => (
                    <div key={f} style={{ marginBottom: 10 }}>
                      <SectionLabel>{l}</SectionLabel>
                      <input value={p[f]} onChange={e => updatePosition(p.id, f, e.target.value)}
                        style={{ width: "100%", boxSizing: "border-box", padding: "8px 12px",
                          borderRadius: 6, border: `1.5px solid ${C.border}`, fontFamily: sans,
                          fontSize: 13.5, outline: "none", background: C.cream }} />
                    </div>
                  ))}
                  <div style={{ marginBottom: 10 }}>
                    <SectionLabel>Description</SectionLabel>
                    <WOTextarea value={p.description} onChange={v => updatePosition(p.id, "description", v)} rows={4} />
                  </div>
                  <div style={{ marginBottom: 14 }}>
                    <SectionLabel>Requirements</SectionLabel>
                    <WOTextarea value={p.requirements} onChange={v => updatePosition(p.id, "requirements", v)} rows={3} />
                  </div>
                  <WOButton onClick={() => setEditingPos(null)} small color={C.green}>Done Editing</WOButton>
                </div>
              ) : (
                <>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div style={{ fontFamily: font, fontWeight: 700, fontSize: 16, color: C.green }}>{p.title}</div>
                    <div style={{ display: "flex", gap: 8 }}>
                      <WOButton onClick={() => setEditingPos(p.id)} small outline color={C.green}>Edit</WOButton>
                      <WOButton onClick={() => deletePosition(p.id)} small outline color={C.red}>Delete</WOButton>
                    </div>
                  </div>
                  <div style={{ fontFamily: sans, fontSize: 13, color: C.mid, marginTop: 8, lineHeight: 1.6 }}>{p.description}</div>
                  <div style={{ fontFamily: sans, fontSize: 12.5, color: C.dark, marginTop: 8, fontStyle: "italic" }}>Requirements: {p.requirements}</div>
                  <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap" }}>
                    {p.schedule && <Badge small>{p.schedule}</Badge>}
                    {p.pay && <Badge small color={C.gold}>{p.pay}</Badge>}
                  </div>
                </>
              )}
            </Card>
          ))}
        </div>
      )}

      {/* ── Tracker Tab ── */}
      {activeView === "tracker" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          {/* Stage legend */}
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            {HIRING_STAGES.map(s => (
              <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 5,
                fontFamily: sans, fontSize: 12, color: s.color, fontWeight: 600 }}>
                <span>{s.icon}</span><span>{s.label}</span>
              </div>
            ))}
          </div>

          {applicants.length === 0 && (
            <div style={{ textAlign: "center", padding: "40px 20px", color: C.muted, fontFamily: sans }}>
              No applicants tracked yet. Analyze a resume to add them automatically.
            </div>
          )}

          {applicants.map(app => {
            const stage = HIRING_STAGES.find(s => s.id === app.stage) || HIRING_STAGES[0];
            const pCol = app.priority === "HIGH" ? C.green : app.priority === "MEDIUM" ? C.goldLight : C.red;
            return (
              <Card key={app.id} style={{ borderLeft: `4px solid ${stage.color}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
                  <div>
                    <div style={{ fontFamily: font, fontWeight: 700, fontSize: 15 }}>{app.name}</div>
                    <div style={{ fontFamily: sans, fontSize: 12.5, color: C.mid, marginTop: 2 }}>
                      {app.positionTitle} · Applied {app.date}
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                    {app.priority && <Badge small color={pCol}>{app.priority}</Badge>}
                    <Badge small color={stage.color}>{stage.icon} {stage.label}</Badge>
                  </div>
                </div>

                <div style={{ marginTop: 12 }}>
                  <SectionLabel>Move to Stage</SectionLabel>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {HIRING_STAGES.map(s => (
                      <button key={s.id} onClick={() => updateStage(app.id, s.id)}
                        style={{ padding: "5px 12px", borderRadius: 20, fontFamily: sans, fontWeight: 600,
                          fontSize: 12, cursor: "pointer",
                          border: `1.5px solid ${app.stage === s.id ? s.color : C.border}`,
                          background: app.stage === s.id ? s.color + "20" : "transparent",
                          color: app.stage === s.id ? s.color : C.muted }}>
                        {s.icon} {s.label}
                      </button>
                    ))}
                  </div>
                </div>

                {app.notes && (
                  <div style={{ marginTop: 10, fontFamily: sans, fontSize: 12, color: C.mid,
                    borderTop: `1px solid ${C.border}`, paddingTop: 10, lineHeight: 1.55 }}>
                    {app.notes.slice(0, 200)}{app.notes.length > 200 ? "…" : ""}
                  </div>
                )}

                <div style={{ marginTop: 10 }}>
                  <WOButton onClick={() => deleteApplicant(app.id)} small outline color={C.red}>Remove</WOButton>
                </div>
              </Card>
            );
          })}

          {hireHistory.length > 0 && (
            <Card green>
              <div style={{ fontFamily: sans, fontWeight: 700, color: C.green, marginBottom: 6 }}>
                🧠 AI Learning Summary
              </div>
              <div style={{ fontFamily: sans, fontSize: 13, color: C.mid }}>
                The AI has learned from <strong>{hireHistory.filter(h=>h.outcome==="fit").length}</strong> successful hires
                and <strong>{hireHistory.filter(h=>h.outcome==="notfit").length}</strong> non-fits.
                Future resume analyses will factor these patterns into priority ratings.
              </div>
            </Card>
          )}
        </div>
      )}
    </div>
  );
}

// ─── PAYMENTS MODULE ──────────────────────────────────────────────────────
function PaymentModule() {
  const [reportText, setReportText] = useState("");
  const [tone, setTone] = useState("friendly");
  const [emails, setEmails] = useState([]);
  const [loading, setLoading] = useState(false);

  const generate = async () => {
    if (!reportText.trim()) return;
    setLoading(true); setEmails([]);
    const sys = `You are the billing coordinator at Waterford Oaks Senior Care. Parse payment data and generate reminder emails for unpaid or overdue accounts. Return ONLY a valid JSON array — no markdown, no explanation:\n[{"resident":"Name","family":"Family contact if known","amount":"$X","dueDate":"date","overdue":true/false,"email":"Full email body"}]`;
    const prompt = `Payment report:\n${reportText}\n\nTone: ${tone}. Generate one email per UNPAID or OVERDUE entry. Each email: address family respectfully by name if known, mention resident, state amount due, give 7-day deadline, include Waterford Oaks phone number placeholder, sign as "Waterford Oaks Billing Team".`;
    let raw = "";
    await callClaude(sys, prompt, t => { raw = t; });
    try {
      const clean = raw.replace(/```json|```/g, "").trim();
      setEmails(JSON.parse(clean));
    } catch {
      setEmails([{ resident: "See output", family: "", amount: "", dueDate: "", overdue: false, email: raw }]);
    }
    setLoading(false);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", gap: 14, alignItems: "flex-end", flexWrap: "wrap" }}>
        <div>
          <SectionLabel>Reminder Tone</SectionLabel>
          <select value={tone} onChange={e => setTone(e.target.value)}
            style={{ padding: "10px 12px", borderRadius: 7, border: `1.5px solid ${C.border}`,
              fontFamily: sans, fontSize: 14, color: C.dark, background: C.cream, outline: "none" }}>
            <option value="friendly">Friendly Reminder</option>
            <option value="firm">Firm & Direct</option>
            <option value="urgent">Urgent / Final Notice</option>
          </select>
        </div>
        <WOButton onClick={generate} disabled={loading || !reportText.trim()} color={C.green}>
          {loading ? "Generating…" : "💳 Generate Reminders"}
        </WOButton>
      </div>

      <div>
        <SectionLabel>Upload Payment Report</SectionLabel>
        <FileDropZone onText={t => setReportText(t)} label="Drop payment report (Excel, CSV, PDF, QuickBooks…)" />
      </div>
      <div>
        <SectionLabel>Or Paste Payment Data</SectionLabel>
        <WOTextarea value={reportText} onChange={setReportText} rows={8}
          placeholder={"Resident Name | Due Date | Amount | Status\nMargaret Smith | 06/01/2026 | $4,500 | UNPAID\nRobert Jones | 06/01/2026 | $3,800 | PAID\nDorothy Williams | 05/15/2026 | $4,200 | 30 DAYS OVERDUE"} />
      </div>

      {loading && <Spinner />}

      {emails.length > 0 && (
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <div style={{ fontFamily: font, fontWeight: 700, fontSize: 16, color: C.green }}>
            📬 {emails.length} Reminder{emails.length !== 1 ? "s" : ""} Generated
          </div>
          {emails.map((e, i) => (
            <Card key={i} style={{ borderLeft: `4px solid ${e.overdue ? C.red : C.goldLight}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
                <div>
                  <div style={{ fontFamily: font, fontWeight: 700, fontSize: 15 }}>{e.resident}</div>
                  {e.family && <div style={{ fontFamily: sans, fontSize: 12.5, color: C.mid }}>Family: {e.family}</div>}
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  {e.amount && <Badge color={C.gold}>{e.amount}</Badge>}
                  {e.overdue && <Badge color={C.red}>OVERDUE</Badge>}
                </div>
              </div>
              <pre style={{ margin: 0, whiteSpace: "pre-wrap", fontFamily: sans, fontSize: 13, color: C.mid,
                borderTop: `1px solid ${C.border}`, paddingTop: 12, lineHeight: 1.65 }}>{e.email}</pre>
              <div style={{ marginTop: 10 }}>
                <WOButton onClick={() => navigator.clipboard.writeText(e.email)} small outline color={C.gold}>Copy Email</WOButton>
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── FINANCIALS MODULE ────────────────────────────────────────────────────
function FinancialModule() {
  const [rentRoll, setRentRoll] = useState("");
  const [report, setReport] = useState("");
  const [loading, setLoading] = useState(false);

  const generate = async () => {
    if (!rentRoll.trim()) return;
    setLoading(true); setReport("");
    const sys = `You are a senior living financial analyst specializing in assisted living and memory care communities. Generate clear, structured financial reports with actionable insights.`;
    const prompt = `Analyze this rent roll for Waterford Oaks Senior Care:\n\n${rentRoll}\n\nProvide:\n1. OCCUPANCY SUMMARY (total units, occupied, vacant, occupancy %)\n2. REVENUE SUMMARY (total monthly revenue, avg rent/unit, projected annual, revenue per care level)\n3. COLLECTION STATUS (paid vs pending vs overdue)\n4. CARE LEVEL BREAKDOWN\n5. KEY INSIGHTS (3–4 observations)\n6. RECOMMENDATIONS (2–3 action items)\n\nBe specific with dollar amounts and percentages.`;
    await callClaude(sys, prompt, t => setReport(t));
    setLoading(false);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <WOButton onClick={generate} disabled={loading || !rentRoll.trim()}>
          {loading ? "Analyzing…" : "📊 Generate Report"}
        </WOButton>
      </div>
      <div>
        <SectionLabel>Upload Rent Roll</SectionLabel>
        <FileDropZone onText={t => setRentRoll(t)} label="Drop rent roll file (Excel, CSV, PDF, QuickBooks…)" />
      </div>
      <div>
        <SectionLabel>Or Paste Rent Roll Data</SectionLabel>
        <WOTextarea value={rentRoll} onChange={setRentRoll} rows={10}
          placeholder={"Unit | Resident | Care Level | Monthly Rent | Status | Move-in Date\n101A | Smith, M | Memory Care | $5,200 | Current | 03/15/2024\n102B | Jones, R | Assisted Living | $3,800 | Current | 08/01/2023\n103A | (Vacant) | — | — | Vacant | —"} />
      </div>
      {(report || loading) && (
        <Card style={{ borderLeft: `4px solid ${C.green}` }}>
          <div style={{ fontFamily: font, fontWeight: 700, fontSize: 17, color: C.green }}>📈 Financial Report</div>
          <div style={{ fontFamily: sans, fontSize: 12, color: C.muted, marginBottom: 8 }}>
            Generated {new Date().toLocaleDateString()} · Waterford Oaks Senior Care
          </div>
          <StreamOutput text={report} loading={loading} />
          {report && !loading && (
            <WOButton onClick={() => navigator.clipboard.writeText(report)} outline small color={C.green} style={{ marginTop: 12 }}>
              Copy Report
            </WOButton>
          )}
        </Card>
      )}
    </div>
  );
}

// ─── BLUEPRINT MODULE ─────────────────────────────────────────────────────
const FLOORS = {
  "Memory Care": {
    color: "#7B3F9E",
    rooms: [
      { id:"MC-101", type:"Private Studio", rent:5200, care:"Memory Care", available:true, sqft:340 },
      { id:"MC-102", type:"Private Studio", rent:5200, care:"Memory Care", available:false, resident:"M. Thompson", sqft:340 },
      { id:"MC-103", type:"Companion Suite", rent:3800, care:"Memory Care", available:true, sqft:480, companion:true },
      { id:"MC-104", type:"Private Studio", rent:5400, care:"Memory Care", available:false, resident:"B. Harrington", sqft:380 },
      { id:"MC-105", type:"Private Studio", rent:5200, care:"Memory Care", available:true, sqft:340 },
      { id:"MC-106", type:"Companion Suite", rent:3800, care:"Memory Care", available:false, resident:"L. Crawford", sqft:480, companion:true },
    ]
  },
  "Assisted Living": {
    color: "#1E4D8C",
    rooms: [
      { id:"AL-201", type:"1 Bedroom", rent:3900, care:"Assisted Living", available:false, resident:"R. Martinez", sqft:540 },
      { id:"AL-202", type:"1 Bedroom", rent:3900, care:"Assisted Living", available:true, sqft:540 },
      { id:"AL-203", type:"Studio", rent:3200, care:"Assisted Living", available:true, sqft:360 },
      { id:"AL-204", type:"1 Bed Deluxe", rent:4400, care:"Assisted Living", available:false, resident:"G. Phillips", sqft:680 },
      { id:"AL-205", type:"Studio", rent:3200, care:"Assisted Living", available:true, sqft:360 },
      { id:"AL-206", type:"1 Bedroom", rent:3900, care:"Assisted Living", available:false, resident:"H. Nguyen", sqft:540 },
    ]
  },
  "Independent Living": {
    color: C.green,
    rooms: [
      { id:"IL-301", type:"1 Bedroom", rent:3000, care:"Independent Living", available:true, sqft:600 },
      { id:"IL-302", type:"2 Bedroom", rent:4600, care:"Independent Living", available:false, resident:"K. & W. Donovan", sqft:880, companion:true },
      { id:"IL-303", type:"1 Bedroom", rent:3000, care:"Independent Living", available:false, resident:"P. Wilson", sqft:600 },
      { id:"IL-304", type:"2 Bed Deluxe", rent:5200, care:"Independent Living", available:true, sqft:1020, companion:true },
      { id:"IL-305", type:"Studio", rent:2600, care:"Independent Living", available:true, sqft:420 },
      { id:"IL-306", type:"1 Bedroom", rent:3000, care:"Independent Living", available:false, resident:"A. Bishop", sqft:600 },
    ]
  }
};

function BlueprintModule() {
  const [selected, setSelected] = useState(null);
  const [filter, setFilter] = useState("All");
  const [availOnly, setAvailOnly] = useState(false);
  const [query, setQuery] = useState("");
  const [suggestion, setSuggestion] = useState("");
  const [sugLoading, setSugLoading] = useState(false);

  const allRooms = Object.values(FLOORS).flatMap(f => f.rooms);
  const avail = allRooms.filter(r => r.available).length;

  const getSuggestion = async () => {
    if (!query.trim()) return;
    setSugLoading(true); setSuggestion("");
    const roomList = Object.entries(FLOORS).flatMap(([care, { rooms }]) =>
      rooms.map(r => `${r.id} (${care}): ${r.type}, $${r.rent}/mo, ${r.available ? "AVAILABLE" : `Occupied – ${r.resident}`}, ${r.sqft} sqft${r.companion ? ", shared option" : ""}`)
    ).join("\n");
    const sys = `You are a senior living placement specialist at Waterford Oaks Senior Care. Match incoming residents to available rooms based on care needs, budget, and preferences. Be warm and specific.`;
    const prompt = `Available rooms:\n${roomList}\n\nNew resident inquiry:\n${query}\n\nRecommend the best 2–3 available room options with clear reasoning. Mention rent, care level match, size, and any companion/roommate options.`;
    await callClaude(sys, prompt, t => setSuggestion(t));
    setSugLoading(false);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {/* Stats */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 12 }}>
        {[
          ["Total Units", allRooms.length, C.charcoal],
          ["Available", avail, C.green],
          ["Occupied", allRooms.length - avail, C.blue],
          ["Occupancy", `${Math.round((allRooms.length-avail)/allRooms.length*100)}%`, C.gold],
        ].map(([l,v,c]) => (
          <div key={l} style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 10, padding: "14px 16px" }}>
            <div style={{ fontFamily: font, fontWeight: 800, fontSize: 24, color: c }}>{v}</div>
            <div style={{ fontFamily: sans, fontSize: 12, color: C.muted, marginTop: 2 }}>{l}</div>
          </div>
        ))}
      </div>

      {/* Filters */}
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        {["All", "Memory Care", "Assisted Living", "Independent Living"].map(f => {
          const fc = f === "All" ? C.green : FLOORS[f]?.color || C.green;
          return (
            <button key={f} onClick={() => setFilter(f)} style={{
              padding: "6px 14px", borderRadius: 20, fontFamily: sans, fontWeight: 700,
              fontSize: 12.5, cursor: "pointer",
              border: `2px solid ${filter===f ? fc : C.border}`,
              background: filter===f ? fc+"18" : "transparent",
              color: filter===f ? fc : C.muted,
            }}>{f}</button>
          );
        })}
        <button onClick={() => setAvailOnly(!availOnly)} style={{
          marginLeft: "auto", padding: "6px 14px", borderRadius: 20, fontFamily: sans,
          fontWeight: 700, fontSize: 12.5, cursor: "pointer",
          border: `2px solid ${availOnly ? C.green : C.border}`,
          background: availOnly ? C.greenPale : "transparent",
          color: availOnly ? C.green : C.muted,
        }}>{availOnly ? "✓ " : ""}Available Only</button>
      </div>

      {/* Floor maps */}
      {Object.entries(FLOORS).map(([care, { color, rooms }]) => {
        const filtered = rooms.filter(r =>
          (filter === "All" || r.care === care) && (!availOnly || r.available)
        );
        if (filtered.length === 0) return null;
        return (
          <Card key={care}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <div style={{ fontFamily: font, fontWeight: 700, fontSize: 15, color }}>{care}</div>
              <Badge small color={color}>{rooms.filter(r=>r.available).length} available</Badge>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(120px, 1fr))", gap: 10 }}>
              {filtered.map(room => {
                const isSel = selected?.id === room.id;
                const rc = room.available ? color : C.muted;
                return (
                  <div key={room.id} onClick={() => setSelected(isSel ? null : room)}
                    style={{ padding: "12px 14px", borderRadius: 10, cursor: "pointer",
                      border: `2px solid ${isSel ? color : room.available ? color+"44" : C.border}`,
                      background: isSel ? color+"18" : room.available ? color+"08" : C.cream,
                      transition: "all 0.15s" }}>
                    <div style={{ fontFamily: font, fontWeight: 800, fontSize: 17, color: room.available ? color : C.muted }}>{room.id}</div>
                    <div style={{ fontFamily: sans, fontSize: 11, color: C.muted, marginTop: 1 }}>{room.type}</div>
                    <div style={{ fontFamily: sans, fontSize: 12.5, fontWeight: 700, color: room.available ? C.dark : C.muted, marginTop: 5 }}>
                      ${room.rent.toLocaleString()}/mo
                    </div>
                    <div style={{ fontFamily: sans, fontSize: 10.5, fontWeight: 700, marginTop: 4,
                      color: room.available ? C.green : C.muted }}>
                      {room.available ? "● AVAILABLE" : "○ OCCUPIED"}
                    </div>
                    {room.companion && <div style={{ fontSize: 10, color: color, marginTop: 2, fontFamily: sans }}>👥 Shared Option</div>}
                  </div>
                );
              })}
            </div>
          </Card>
        );
      })}

      {/* Room detail */}
      {selected && (
        <Card style={{ borderLeft: `4px solid ${FLOORS[selected.care]?.color || C.green}` }}>
          <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
            <div>
              <div style={{ fontFamily: font, fontWeight: 800, fontSize: 20 }}>Unit {selected.id}</div>
              <div style={{ fontFamily: sans, color: C.mid, fontSize: 13.5, marginTop: 3 }}>
                {selected.type} · {selected.sqft} sqft · {selected.care}
              </div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontFamily: font, fontWeight: 800, fontSize: 22, color: FLOORS[selected.care]?.color || C.green }}>
                ${selected.rent.toLocaleString()}/mo
              </div>
              {selected.available
                ? <Badge color={C.green} small>Available for Move-In</Badge>
                : <Badge color={C.muted} small>Occupied — {selected.resident}</Badge>
              }
            </div>
          </div>
        </Card>
      )}

      {/* AI Advisor */}
      <Card green>
        <div style={{ fontFamily: font, fontWeight: 700, fontSize: 16, color: C.green, marginBottom: 10 }}>
          🤖 AI Move-In Advisor
        </div>
        <WOTextarea value={query} onChange={setQuery} rows={3}
          placeholder="Describe the incoming resident… e.g. '82-year-old female with mild dementia, budget ~$4,500/mo, prefers private room, enjoys group activities.'" />
        <div style={{ marginTop: 12 }}>
          <WOButton onClick={getSuggestion} disabled={sugLoading || !query.trim()}>
            {sugLoading ? "Matching…" : "🏠 Find Best Room Match"}
          </WOButton>
        </div>
        <StreamOutput text={suggestion} loading={sugLoading} />
      </Card>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────
const TABS = [
  { id: "recruiting", label: "Recruiting",  icon: "👤" },
  { id: "payments",   label: "Payments",    icon: "💳" },
  { id: "financial",  label: "Financials",  icon: "📊" },
  { id: "blueprint",  label: "Blueprint",   icon: "🏢" },
];

export default function App() {
  const [tab, setTab] = useState("recruiting");

  return (
    <div style={{ minHeight: "100vh", background: C.cream, fontFamily: sans, color: C.dark }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Lora:wght@400;600;700;800&family=Source+Sans+3:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: ${C.greenTint}; }
        ::-webkit-scrollbar-thumb { background: ${C.borderDark}; border-radius: 4px; }
        select option { background: #fff; color: ${C.dark}; }
        @keyframes woPulse { 0%,100%{opacity:0.3;transform:scale(0.8)} 50%{opacity:1;transform:scale(1.2)} }
        @keyframes woBlink { 0%,100%{opacity:1} 50%{opacity:0} }
        @keyframes woFade { from{opacity:0;transform:translateY(6px)} to{opacity:1;transform:translateY(0)} }
        textarea:focus, input:focus { border-color: ${C.greenMid} !important; box-shadow: 0 0 0 3px ${C.green}18; }
      `}</style>

      {/* Header */}
      <div style={{ background: C.green, borderBottom: `3px solid ${C.greenMid}`, position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ maxWidth: 980, margin: "0 auto", padding: "0 20px", display: "flex", alignItems: "center", gap: 16, height: 62 }}>
          {/* Logo area */}
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ fontSize: 26 }}>🌳</div>
            <div>
              <div style={{ fontFamily: font, fontWeight: 700, fontSize: 17, color: "#fff", lineHeight: 1.1 }}>
                Waterford Oaks
              </div>
              <div style={{ fontFamily: sans, fontSize: 10.5, color: "#ffffffbb", letterSpacing: "0.12em", textTransform: "uppercase" }}>
                Senior Care · Management Suite
              </div>
            </div>
          </div>

          {/* Nav */}
          <nav style={{ display: "flex", gap: 2, marginLeft: "auto" }}>
            {TABS.map(t => (
              <button key={t.id} onClick={() => setTab(t.id)} style={{
                display: "flex", alignItems: "center", gap: 6,
                padding: "7px 16px", borderRadius: 6, border: "none",
                background: tab === t.id ? "rgba(255,255,255,0.18)" : "transparent",
                color: tab === t.id ? "#fff" : "rgba(255,255,255,0.7)",
                fontFamily: sans, fontWeight: 700, fontSize: 13.5, cursor: "pointer",
                borderBottom: tab === t.id ? "3px solid #fff" : "3px solid transparent",
                transition: "all 0.15s",
              }}>
                <span>{t.icon}</span>
                <span>{t.label}</span>
              </button>
            ))}
          </nav>
        </div>
      </div>

      {/* Green accent bar */}
      <div style={{ height: 4, background: `linear-gradient(90deg, ${C.green}, ${C.greenLight}, ${C.gold})` }} />

      {/* Content */}
      <div style={{ maxWidth: 980, margin: "0 auto", padding: "32px 20px 60px" }}>
        {/* Page title */}
        <div style={{ marginBottom: 24, paddingBottom: 18, borderBottom: `2px solid ${C.border}` }}>
          <h1 style={{ fontFamily: font, fontWeight: 700, fontSize: 24, color: C.green, margin: 0 }}>
            {TABS.find(t=>t.id===tab)?.icon} {TABS.find(t=>t.id===tab)?.label}
          </h1>
          <p style={{ fontFamily: sans, color: C.muted, fontSize: 13.5, marginTop: 5 }}>
            {{
              recruiting: "AI-powered candidate analysis, job description library, and hiring pipeline tracker",
              payments: "Upload any payment report to auto-generate family reminder emails",
              financial: "Drop your rent roll for AI-generated financial insights and recommendations",
              blueprint: "Interactive room map with AI-powered move-in matching",
            }[tab]}
          </p>
        </div>

        <div key={tab} style={{ animation: "woFade 0.25s ease" }}>
          {tab === "recruiting"  && <RecruitingModule />}
          {tab === "payments"    && <PaymentModule />}
          {tab === "financial"   && <FinancialModule />}
          {tab === "blueprint"   && <BlueprintModule />}
        </div>
      </div>

      {/* Footer */}
      <div style={{ borderTop: `2px solid ${C.border}`, background: C.white, padding: "14px 24px",
        textAlign: "center", fontFamily: sans, fontSize: 12, color: C.muted }}>
        Waterford Oaks Senior Care · Memory Care & Assisted Living · waterfordoaksseniorcare.com
      </div>
    </div>
  );
}
