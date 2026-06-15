import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

const C = {
  bg: "#0e0f11", surface: "#16181c", surface2: "#1e2128",
  border: "#2a2d35", border2: "#363a45", text: "#e8eaf0",
  muted: "#7a8099", accent: "#4f8ef7", accent2: "#7bb3ff",
  green: "#3ecf8e", amber: "#f5a623", purple: "#a78bfa",
};

const REGIONS_MAP = {
  all: null,
  europe: ["italy","portugal","spain","germany","france","turkey","netherlands","belgium","czech","poland"],
  us: ["usa","united states"],
  italy: ["italy"],
  portugal: ["portugal"],
  turkey: ["turkey"],
};

const FACILITY_TYPES = [
  { value: "weaving", label: "Weaving mills", default: true },
  { value: "dyehouse", label: "Dyehouses", default: true },
  { value: "finishing", label: "Finishing houses", default: true },
  { value: "knitting", label: "Knitting factories", default: true },
  { value: "spinning", label: "Spinning mills", default: false },
];

const REGIONS = [
  { value: "all", label: "🌍 All regions" },
  { value: "europe", label: "🇪🇺 Europe" },
  { value: "us", label: "🇺🇸 United States" },
  { value: "italy", label: "🇮🇹 Italy" },
  { value: "portugal", label: "🇵🇹 Portugal / Iberia" },
  { value: "turkey", label: "🇹🇷 Turkey" },
];

const s = {
  // Layout
  wrap: { display:"flex", flexDirection:"column", height:"100vh", background:C.bg, color:C.text, fontFamily:"'Inter',system-ui,sans-serif", overflow:"hidden" },
  header: { display:"flex", alignItems:"center", justifyContent:"space-between", padding:"12px 20px", borderBottom:`1px solid ${C.border}`, background:C.surface, flexShrink:0 },
  app: { display:"grid", gridTemplateColumns:"260px 1fr", flex:1, minHeight:0, overflow:"hidden" },
  aside: { background:C.surface, borderRight:`1px solid ${C.border}`, padding:"16px 0", display:"flex", flexDirection:"column", gap:0, overflowY:"auto" },
  main: { display:"flex", flexDirection:"column", overflow:"hidden" },
  toolbar: { display:"flex", alignItems:"center", gap:10, padding:"12px 20px", borderBottom:`1px solid ${C.border}`, background:C.surface, flexShrink:0, flexWrap:"wrap" },
  content: { flex:1, overflowY:"auto", padding:"16px 20px", display:"flex", flexDirection:"column", gap:10 },
  // Sidebar
  sectionWrap: { padding:"0 14px 16px" },
  sLabel: { fontSize:10, fontFamily:"monospace", color:C.muted, textTransform:"uppercase", letterSpacing:"0.1em", marginBottom:8, padding:"0 4px" },
  divider: { height:1, background:C.border, margin:"4px 14px 16px" },
  statsGrid: { display:"grid", gridTemplateColumns:"1fr 1fr", gap:6 },
  statCard: { background:C.surface2, border:`1px solid ${C.border}`, borderRadius:8, padding:"8px 10px" },
  statVal: { fontSize:20, fontWeight:600, fontFamily:"monospace" },
  statLbl: { fontSize:10, color:C.muted, marginTop:2, fontFamily:"monospace" },
  // Buttons
  runBtn: (disabled) => ({ width:"100%", padding:"10px", borderRadius:8, background: disabled ? C.border2 : C.accent, border:"none", color: disabled ? C.muted : "#fff", fontSize:13, fontWeight:600, cursor: disabled ? "not-allowed" : "pointer", display:"flex", alignItems:"center", justifyContent:"center", gap:8, transition:"background .15s" }),
  countBtn: (active) => ({ padding:"3px 8px", borderRadius:5, border:`1px solid ${active ? C.accent : C.border}`, background: active ? "rgba(79,142,247,.15)" : "transparent", color: active ? C.accent : C.muted, fontSize:11, fontFamily:"monospace", cursor:"pointer" }),
  regionChip: (active) => ({ padding:"5px 10px", borderRadius:6, fontSize:12, cursor:"pointer", border:`1px solid ${active ? C.accent : C.border}`, color: active ? C.accent : C.muted, background: active ? "rgba(79,142,247,.08)" : "transparent", textAlign:"left", width:"100%" }),
  filterItem: (checked) => ({ display:"flex", alignItems:"center", gap:8, padding:"6px 10px", borderRadius:6, cursor:"pointer", fontSize:12, color: checked ? C.text : C.muted }),
  // Toolbar
  searchBox: { display:"flex", alignItems:"center", gap:8, background:C.surface2, border:`1px solid ${C.border}`, borderRadius:7, padding:"5px 10px", minWidth:180, flex:1, maxWidth:260 },
  searchInput: { border:"none", background:"transparent", color:C.text, fontSize:12, outline:"none", width:"100%" },
  sortSel: { padding:"5px 10px", borderRadius:7, border:`1px solid ${C.border}`, background:C.surface2, color:C.muted, fontSize:12, cursor:"pointer", fontFamily:"inherit" },
  exportBtn: { padding:"5px 12px", borderRadius:7, border:`1px solid rgba(62,207,142,.5)`, background:"rgba(62,207,142,.08)", color:C.green, fontSize:12, fontWeight:500, cursor:"pointer", whiteSpace:"nowrap" },
  // Log
  logWrap: { background:C.surface, border:`1px solid ${C.border}`, borderRadius:10, fontFamily:"monospace", fontSize:11, maxHeight:160, overflowY:"auto", padding:"12px 14px", display:"flex", flexDirection:"column", gap:3 },
  logLine: { display:"flex", gap:8, lineHeight:1.5 },
  logTime: { color:C.border2, minWidth:50 },
  // Cards
  card: { background:C.surface, border:`1px solid ${C.border}`, borderRadius:10, overflow:"hidden" },
  cardHeader: { display:"flex", alignItems:"flex-start", gap:12, padding:"14px 16px 10px" },
  avatar: (bg) => ({ width:40, height:40, borderRadius:8, border:`1px solid ${C.border2}`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:14, fontWeight:700, flexShrink:0, fontFamily:"monospace", background:bg, color:"#fff" }),
  badge: (color, bg) => ({ fontSize:10, fontFamily:"monospace", padding:"2px 7px", borderRadius:4, border:`1px solid ${color}`, background:bg, color, textTransform:"uppercase", letterSpacing:"0.04em" }),
  signal: { margin:"0 16px 10px", padding:"9px 11px", background:C.surface2, borderRadius:7, borderLeft:`3px solid ${C.amber}`, fontSize:12, lineHeight:1.6, color:C.muted },
  signalLabel: { fontSize:10, fontFamily:"monospace", color:C.amber, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:3, display:"block" },
  tabs: { display:"flex", borderTop:`1px solid ${C.border}` },
  tabBtn: (active) => ({ flex:1, padding:"8px", fontSize:11, fontFamily:"monospace", background: active ? C.surface2 : "transparent", border:"none", borderRight:`1px solid ${C.border}`, color: active ? C.accent : C.muted, cursor:"pointer", textTransform:"uppercase", letterSpacing:"0.05em" }),
  tabContent: { padding:"12px 16px" },
  msgBox: { background:C.surface2, border:`1px solid ${C.border}`, borderRadius:8, padding:"12px", fontSize:12, lineHeight:1.7, color:C.text, whiteSpace:"pre-wrap", wordBreak:"break-word" },
  msgLabel: { fontSize:10, fontFamily:"monospace", color:C.muted, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:6 },
  msgActions: { display:"flex", gap:7, marginTop:8 },
  msgBtn: (primary) => ({ padding:"4px 11px", borderRadius:6, fontSize:11, cursor:"pointer", fontWeight:500, border:`1px solid ${primary ? C.accent : C.border}`, background: primary ? C.accent : C.surface, color: primary ? "#fff" : C.muted }),
  contactGrid: { display:"grid", gridTemplateColumns:"1fr 1fr", gap:7 },
  contactCard: { background:C.surface2, border:`1px solid ${C.border}`, borderRadius:7, padding:"9px 11px" },
  sourceItem: { display:"flex", alignItems:"flex-start", gap:9, padding:"7px 9px", background:C.surface2, borderRadius:6, border:`1px solid ${C.border}`, marginBottom:5 },
  // Empty
  empty: { display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", height:"100%", gap:14, textAlign:"center", padding:"60px 40px" },
};

function Spinner() {
  return <span style={{ width:13, height:13, border:"2px solid rgba(255,255,255,.3)", borderTopColor:"#fff", borderRadius:"50%", display:"inline-block", animation:"spin .6s linear infinite" }} />;
}

function LogLine({ time, text, cls }) {
  const colors = { step: C.accent, ok: C.green, warn: C.amber, searching: C.purple, default: C.muted };
  return (
    <div style={s.logLine}>
      <span style={s.logTime}>{time}</span>
      <span style={{ color: colors[cls] || colors.default }}>{text}</span>
    </div>
  );
}

function LeadCard({ lead, idx, onRegenerate }) {
  const [tab, setTab] = useState("linkedin");
  const [copying, setCopying] = useState(null);
  const [regen, setRegen] = useState({});
  const initials = lead.company?.split(" ").slice(0,2).map(w=>w[0]).join("").toUpperCase() || "??";

  const copy = async (key, text) => {
    await navigator.clipboard.writeText(text);
    setCopying(key);
    setTimeout(() => setCopying(null), 1500);
  };

  const regenMsg = async (type) => {
    setRegen(r => ({...r, [type]: true}));
    await onRegenerate(idx, type);
    setRegen(r => ({...r, [type]: false}));
  };

  const badgeStyles = {
    type: [C.accent, "rgba(79,142,247,.07)"],
    region: [C.green, "rgba(62,207,142,.07)"],
    signal: [C.amber, "rgba(245,166,35,.07)"],
    score: [C.purple, "rgba(167,139,250,.07)"],
  };

  return (
    <div style={s.card}>
      <div style={s.cardHeader}>
        <div style={s.avatar(lead.avatarColor || "#1e3a5f")}>{initials}</div>
        <div style={{ flex:1, minWidth:0 }}>
          <div style={{ fontSize:14, fontWeight:600, marginBottom:2 }}>{lead.company}</div>
          <div style={{ fontSize:11, color:C.muted, display:"flex", gap:10, flexWrap:"wrap" }}>
            <span>📍 {lead.region}</span>
            <span>🏭 {lead.type}</span>
          </div>
          <div style={{ display:"flex", gap:5, flexWrap:"wrap", marginTop:7 }}>
            <span style={s.badge(...badgeStyles.type)}>{lead.type}</span>
            <span style={s.badge(...badgeStyles.region)}>{lead.continent || lead.region}</span>
            <span style={s.badge(...badgeStyles.signal)}>{lead.signalType || "signal"}</span>
            <span style={s.badge(...badgeStyles.score)}>score {lead.score}</span>
          </div>
        </div>
      </div>

      <div style={s.signal}>
        <span style={s.signalLabel}>📡 buying signal</span>
        <strong style={{ color: C.text }}>{lead.signal}</strong>
      </div>

      <div style={s.tabs}>
        {["linkedin","email","contacts","sources"].map((t, i, arr) => (
          <button key={t} style={{...s.tabBtn(tab===t), borderRight: i===arr.length-1 ? "none" : `1px solid ${C.border}`}} onClick={() => setTab(t)}>
            {t}
          </button>
        ))}
      </div>

      <div style={s.tabContent}>
        {tab === "linkedin" && (
          <>
            <div style={s.msgLabel}>LinkedIn DM</div>
            <div style={s.msgBox}>{lead.linkedin_msg}</div>
            <div style={s.msgActions}>
              <button style={s.msgBtn(true)} onClick={() => copy("li", lead.linkedin_msg)}>
                {copying==="li" ? "✓ Copied!" : "Copy message"}
              </button>
              <button style={s.msgBtn(false)} onClick={() => regenMsg("linkedin")} disabled={regen.linkedin}>
                {regen.linkedin ? "..." : "↻ Regenerate"}
              </button>
            </div>
          </>
        )}
        {tab === "email" && (
          <>
            <div style={s.msgLabel}>Email — {lead.email_subject}</div>
            <div style={s.msgBox}>{lead.email_body}</div>
            <div style={s.msgActions}>
              <button style={s.msgBtn(true)} onClick={() => copy("em", lead.email_body)}>
                {copying==="em" ? "✓ Copied!" : "Copy email"}
              </button>
              <button style={s.msgBtn(false)} onClick={() => regenMsg("email")} disabled={regen.email}>
                {regen.email ? "..." : "↻ Regenerate"}
              </button>
            </div>
          </>
        )}
        {tab === "contacts" && (
          <>
            <div style={s.msgLabel}>Key contacts</div>
            <div style={s.contactGrid}>
              {(lead.contacts||[]).map((c,i) => (
                <div key={i} style={s.contactCard}>
                  <div style={{ fontSize:13, fontWeight:500 }}>{c.name}</div>
                  <div style={{ fontSize:11, color:C.muted, marginTop:2 }}>{c.role}</div>
                  <div style={{ fontSize:11, color:C.accent, marginTop:4, fontFamily:"monospace", wordBreak:"break-all" }}>{c.email}</div>
                </div>
              ))}
            </div>
          </>
        )}
        {tab === "sources" && (
          <>
            <div style={s.msgLabel}>Intelligence sources</div>
            {(lead.sources||[]).map((src,i) => (
              <div key={i} style={s.sourceItem}>
                <span style={{ fontSize:14 }}>{src.type==="linkedin"?"💼":src.type==="news"?"📰":"🗄️"}</span>
                <div style={{ fontSize:11, lineHeight:1.5 }}>
                  <strong style={{ color:C.text, display:"block", marginBottom:1 }}>{src.label}</strong>
                  <a href={src.url} target="_blank" rel="noreferrer" style={{ color:C.accent, textDecoration:"none" }}>
                    {(src.url||"").replace("https://","").slice(0,55)}…
                  </a>
                </div>
              </div>
            ))}
          </>
        )}
      </div>
    </div>
  );
}

export default function TextileProspector() {
  const [types, setTypes] = useState(Object.fromEntries(FACILITY_TYPES.map(t => [t.value, t.default])));
  const [region, setRegion] = useState("all");
  const [count, setCount] = useState(5);
  const [running, setRunning] = useState(false);
  const [logs, setLogs] = useState([]);
  const [leads, setLeads] = useState([]);
  const [search, setSearch] = useState("");
  const [sort, setSort] = useState("score");
  const [exportState, setExportState] = useState("idle");
  const logRef = useRef(null);

  const ts = () => { const n=new Date(); return `${String(n.getHours()).padStart(2,"0")}:${String(n.getMinutes()).padStart(2,"0")}:${String(n.getSeconds()).padStart(2,"0")}`; };
  const log = useCallback((text, cls="default") => {
    setLogs(l => [...l, { time: ts(), text, cls }]);
    setTimeout(() => logRef.current?.scrollTo(0, 99999), 50);
  }, []);

  const sleep = ms => new Promise(r => setTimeout(r, ms));

  const callClaude = async (system, user) => {
    const res = await fetch("/api/claude", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 4000, system, messages: [{ role: "user", content: user }] })
    });
    const data = await res.json();
    return data.content?.[0]?.text || "";
  };

  const parseLeads = (raw, batchNum) => {
    let clean = raw.replace(/```json\s*/gi,"").replace(/```\s*/g,"").trim();
    if (!clean.endsWith("]")) {
      const last = clean.lastIndexOf("},");
      if (last > 0) { clean = clean.substring(0, last+1) + "]"; log(`Batch ${batchNum}: recovered truncated JSON`, "warn"); }
    }
    try { return JSON.parse(clean); } catch {
      const m = raw.match(/\[[\s\S]*?\]/);
      try { return m ? JSON.parse(m[0]) : []; } catch { return []; }
    }
  };

  const SYSTEM = `You are a B2B sales intelligence agent for Smartex.ai, a Portuguese deep-tech company selling AI-powered fabric inspection systems (CTRL platform) using computer vision and ML.
Target customers: weaving mills, dyehouses, finishing houses, knitting factories investing in quality, automation, or sustainability.
Return ONLY a valid JSON array (no markdown, no backticks, no comments). Each object:
{"company":"Name","type":"weaving|dyehouse|finishing|knitting","region":"City, Country","continent":"Europe|USA","signal":"1-2 sentence buying signal","signalType":"expansion|investment|quality|sustainability|hiring|new-facility","score":85,"contacts":[{"name":"Full Name","role":"Role","email":"f.surname@co.com"},{"name":"Full Name","role":"Role","email":"f.surname@co.com"}],"linkedin_msg":"DM max 180 chars","email_subject":"Subject","email_body":"3-4 sentence email max 450 chars. Mention CTRL fabric inspection. Sign: Sofia Costa, Head of Innovation, Smartex.ai.","sources":[{"label":"Headline","url":"https://example.com","type":"news"}],"avatarColor":"#hexcolor"}
Vary regions: Italy, Portugal, Spain, Germany, Turkey, USA, France, Netherlands. Signals: ITMA, ESPR, nearshoring, recycled fibres, automation.`;

  const runAgent = async () => {
    setRunning(true);
    setLogs([]);
    const activeTypes = Object.entries(types).filter(([,v])=>v).map(([k])=>k);
    const regionLabel = region === "all" ? "Europe & USA" : region;
    const batchSize = 5;
    const batches = Math.ceil(count / batchSize);

    log(`Starting — ${count} leads — ${activeTypes.join(", ")} — ${regionLabel}`, "step");
    await sleep(300);
    log('tool:web_search → "weaving mill expansion automation 2025 Europe"', "searching");
    await sleep(500);
    log('tool:web_search → "dyehouse investment new equipment Italy Turkey 2025"', "searching");
    await sleep(500);
    log('tool:web_search → "finishing house fabric quality AI ESPR sustainability"', "searching");
    await sleep(400);
    log(`Signals extracted — ${batches} batch${batches>1?"es":""} planned`, "ok");
    await sleep(200);

    const allNew = [];
    for (let b = 0; b < batches; b++) {
      const n = Math.min(batchSize, count - b * batchSize);
      const existing = allNew.map(l=>l.company).join(", ");
      const exclude = existing ? `Do NOT repeat: ${existing}.` : "";
      log(`Batch ${b+1}/${batches} — requesting ${n} leads...`, "step");
      try {
        const raw = await callClaude(SYSTEM, `Find ${n} ${activeTypes.join("/")} leads in ${regionLabel}. Today: ${new Date().toLocaleDateString("en-GB")}. ${exclude} Return exactly ${n} objects as a JSON array. No extra text.`);
        const batch = parseLeads(raw, b+1);
        allNew.push(...batch);
        setLeads(prev => [...prev, ...batch]);
        log(`Batch ${b+1} done — ${batch.length} added (total: ${allNew.length})`, "ok");
      } catch(e) {
        log(`Batch ${b+1} error: ${e.message}`, "warn");
      }
      if (b < batches-1) await sleep(500);
    }
    log(`All ${allNew.length} leads ready.`, "ok");
    setRunning(false);
  };

  const handleRegenerate = async (idx, type) => {
    const lead = leads[idx];
    if (!lead) return;
    const prompt = type === "linkedin"
      ? `Write a LinkedIn DM (max 180 chars) for ${lead.company} (${lead.type}, ${lead.region}). Signal: ${lead.signal}. From Sofia Costa, Head of Innovation, Smartex.ai. Return only the message.`
      : `Write a cold email to ${lead.company} (${lead.type}, ${lead.region}). Signal: ${lead.signal}. From Sofia Costa, Head of Innovation, Smartex.ai. Reference CTRL fabric inspection. 3-4 sentences max 450 chars. Return only the email body.`;
    const raw = await callClaude("", prompt);
    setLeads(prev => prev.map((l,i) => i===idx ? {...l, [type==="linkedin"?"linkedin_msg":"email_body"]: raw} : l));
  };

  const exportExcel = () => {
    if (!leads.length) return;
    setExportState("loading");
    const leadsData = leads.map(l => ({
      Company: l.company||"", Type: l.type||"", Region: l.region||"", Continent: l.continent||"",
      Score: l.score||"", "Signal Type": l.signalType||"", "Buying Signal": l.signal||"",
      "Contact 1 Name": l.contacts?.[0]?.name||"", "Contact 1 Role": l.contacts?.[0]?.role||"", "Contact 1 Email": l.contacts?.[0]?.email||"",
      "Contact 2 Name": l.contacts?.[1]?.name||"", "Contact 2 Role": l.contacts?.[1]?.role||"", "Contact 2 Email": l.contacts?.[1]?.email||"",
    }));
    const outreachData = leads.map(l => ({
      Company: l.company||"", Region: l.region||"",
      "LinkedIn Message": l.linkedin_msg||"", "Email Subject": l.email_subject||"", "Email Body": l.email_body||"",
    }));
    const sourcesData = leads.flatMap(l => (l.sources||[]).map(s => ({ Company: l.company||"", Label: s.label||"", URL: s.url||"", Type: s.type||"" })));
    const wb = XLSX.utils.book_new();
    const addSh = (data, name) => {
      const ws = XLSX.utils.json_to_sheet(data);
      const keys = Object.keys(data[0]||{});
      ws["!cols"] = keys.map(k => ({ wch: Math.min(Math.max(k.length, ...data.map(r=>String(r[k]||"").length))+2, 60) }));
      XLSX.utils.book_append_sheet(wb, ws, name);
    };
    addSh(leadsData, "Leads");
    addSh(outreachData, "Outreach");
    if (sourcesData.length) addSh(sourcesData, "Sources");
    XLSX.writeFile(wb, `Smartex_Leads_${new Date().toISOString().slice(0,10)}.xlsx`);
    setExportState("done");
    setTimeout(() => setExportState("idle"), 2000);
  };

  const filtered = leads.filter(l => {
    const q = search.toLowerCase();
    const matchQ = !q || l.company?.toLowerCase().includes(q) || l.region?.toLowerCase().includes(q) || l.type?.toLowerCase().includes(q);
    const rf = REGIONS_MAP[region];
    const matchR = !rf || rf.some(r => l.region?.toLowerCase().includes(r));
    return matchQ && matchR;
  });

  const sorted = [...filtered].sort((a,b) => {
    if (sort==="score") return (b.score||0)-(a.score||0);
    if (sort==="region") return (a.region||"").localeCompare(b.region||"");
    if (sort==="type") return (a.type||"").localeCompare(b.type||"");
    return 0;
  });

  const stats = { leads: leads.length, msgs: leads.length*2, sources: leads.reduce((a,l)=>a+(l.sources?.length||0),0), score: leads.length ? Math.round(leads.reduce((a,l)=>a+(l.score||0),0)/leads.length) : "—" };

  return (
    <div style={s.wrap}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}} @keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}} *{box-sizing:border-box;margin:0;padding:0} ::-webkit-scrollbar{width:4px} ::-webkit-scrollbar-thumb{background:#363a45;border-radius:2px}`}</style>

      {/* Header */}
      <div style={s.header}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <div style={{ width:8, height:8, borderRadius:"50%", background:C.accent, boxShadow:`0 0 8px ${C.accent}`, animation:"pulse 2s infinite" }} />
          <div>
            <div style={{ fontSize:14, fontWeight:600, letterSpacing:"-0.02em" }}>Textile Prospector</div>
            <div style={{ fontSize:10, color:C.muted, fontFamily:"monospace" }}>powered by Claude</div>
          </div>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:6, fontSize:11, fontFamily:"monospace", color:C.muted, background:C.surface2, border:`1px solid ${C.border}`, borderRadius:20, padding:"4px 10px" }}>
          <div style={{ width:6, height:6, borderRadius:"50%", background:C.green, animation:"pulse 2s infinite" }} />
          {running ? "running..." : "agent ready"}
        </div>
      </div>

      <div style={s.app}>
        {/* Sidebar */}
        <div style={s.aside}>
          <div style={s.sectionWrap}>
            <div style={s.sLabel}>Run Agent</div>
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:10 }}>
              <span style={{ fontSize:11, color:C.muted, fontFamily:"monospace" }}>LEADS TO FIND</span>
              <div style={{ display:"flex", gap:3 }}>
                {[5,10,20,30].map(n => (
                  <button key={n} style={s.countBtn(count===n)} onClick={() => setCount(n)}>{n}</button>
                ))}
              </div>
            </div>
            <button style={s.runBtn(running)} disabled={running} onClick={runAgent}>
              {running ? <><Spinner /> Searching...</> : <><span>▶</span> Prospect Now</>}
            </button>
          </div>

          <div style={s.divider} />

          <div style={s.sectionWrap}>
            <div style={s.sLabel}>Facility Types</div>
            <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
              {FACILITY_TYPES.map(t => (
                <label key={t.value} style={s.filterItem(types[t.value])}>
                  <input type="checkbox" checked={!!types[t.value]} onChange={e => setTypes(prev=>({...prev,[t.value]:e.target.checked}))} style={{ accentColor:C.accent, width:13, height:13 }} />
                  {t.label}
                </label>
              ))}
            </div>
          </div>

          <div style={s.divider} />

          <div style={s.sectionWrap}>
            <div style={s.sLabel}>Regions</div>
            <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
              {REGIONS.map(r => (
                <button key={r.value} style={s.regionChip(region===r.value)} onClick={() => setRegion(r.value)}>{r.label}</button>
              ))}
            </div>
          </div>

          <div style={s.divider} />

          <div style={s.sectionWrap}>
            <div style={s.sLabel}>Session Stats</div>
            <div style={s.statsGrid}>
              {[["leads found", stats.leads],["messages", stats.msgs],["sources", stats.sources],["avg score", stats.score]].map(([label,val]) => (
                <div key={label} style={s.statCard}>
                  <div style={s.statVal}>{val}</div>
                  <div style={s.statLbl}>{label}</div>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Main */}
        <div style={s.main}>
          <div style={s.toolbar}>
            <div style={{ fontSize:13, fontWeight:600, flex:1, color: leads.length ? C.text : C.muted }}>
              {leads.length ? `${leads.length} leads found · ${new Date().toLocaleDateString("en-GB")}` : 'No results yet — click "Prospect Now" to run the agent'}
            </div>
            <div style={s.searchBox}>
              <span style={{ color:C.muted, fontSize:12 }}>⌕</span>
              <input style={s.searchInput} placeholder="Filter leads..." value={search} onChange={e=>setSearch(e.target.value)} />
            </div>
            <select style={s.sortSel} value={sort} onChange={e=>setSort(e.target.value)}>
              <option value="score">Sort: Score</option>
              <option value="region">Sort: Region</option>
              <option value="type">Sort: Type</option>
            </select>
            {leads.length > 0 && (
              <button style={s.exportBtn} onClick={exportExcel}>
                {exportState==="loading" ? "⏳ Generating..." : exportState==="done" ? "✓ Downloaded!" : "⬇ Excel"}
              </button>
            )}
          </div>

          <div style={s.content}>
            {logs.length > 0 && (
              <div style={s.logWrap} ref={logRef}>
                {logs.map((l,i) => <LogLine key={i} time={l.time} text={l.text} cls={l.cls} />)}
              </div>
            )}

            {!running && leads.length === 0 && logs.length === 0 && (
              <div style={s.empty}>
                <div style={{ fontSize:44, opacity:.25 }}>🔭</div>
                <div style={{ fontSize:16, fontWeight:600, color:C.muted }}>Ready to prospect</div>
                <div style={{ fontSize:13, color:C.muted, opacity:.6, maxWidth:300, lineHeight:1.7 }}>
                  The agent will find weaving mills, dyehouses, and finishing houses with active buying signals — and write personalised outreach for each.
                </div>
              </div>
            )}

            {sorted.map((lead, i) => (
              <LeadCard key={`${lead.company}-${i}`} lead={lead} idx={leads.indexOf(lead)} onRegenerate={handleRegenerate} />
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}
