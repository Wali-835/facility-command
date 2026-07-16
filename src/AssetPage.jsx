import { useState, useEffect } from "react";
import { supabase } from "./supabase";
import { t } from "./i18n.js";

const C = {
  bg: "#0d0f12", surface: "#141720", card: "#1a1e2a", border: "#252b3b",
  accent: "#f97316", yellow: "#eab308", green: "#22c55e", red: "#ef4444",
  blue: "#3b82f6", purple: "#a855f7", text: "#e2e8f0", muted: "#64748b", subtle: "#94a3b8",
};

const SEVERITY_COLORS = { Critical: "#ef4444", High: "#f97316", Medium: "#eab308", Low: "#22c55e" };
const statusColor = (s) => ({ Operational: C.green, "Under Maintenance": C.accent, Degraded: C.red, Open: C.accent, "In Progress": C.blue, Completed: C.green, Pending: C.yellow, Acknowledged: C.blue, Resolved: C.green }[s] || C.muted);
const fmtDate = (d) => d ? new Date(d).toLocaleDateString("en-GB") : "—";
const fmtDateTime = (d) => { if (!d) return "—"; try { let s = d.replace(" ","T"); if (!/[Zz]|[+-]\d{2}:?\d{2}$/.test(s)) s += "Z"; const date = new Date(s); if (isNaN(date.getTime())) return "—"; return date.toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit", timeZone: "Africa/Cairo" }); } catch { return "—"; } };
const uid = (p) => `${p}-${Date.now().toString(36).toUpperCase()}${Math.random().toString(36).slice(2,5).toUpperCase()}`;
const TODAY = new Date().toISOString().split("T")[0];

// ─── Shared UI ────────────────────────────────────────────────────────────────
const Input = ({ label, value, onChange, type="text", placeholder }) => (
  <div>
    <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
    <input type={type} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder}
      style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px", color: C.text, fontSize: 15, boxSizing: "border-box" }} />
  </div>
);

const Textarea = ({ label, value, onChange, placeholder }) => (
  <div>
    <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
    <textarea value={value} onChange={e => onChange(e.target.value)} rows={3} placeholder={placeholder}
      style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px", color: C.text, fontSize: 15, boxSizing: "border-box", resize: "vertical" }} />
  </div>
);

const Sel = ({ label, value, onChange, options }) => (
  <div>
    <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
    <select value={value} onChange={e => onChange(e.target.value)}
      style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px", color: C.text, fontSize: 15 }}>
      {options.map(o => <option key={o}>{o}</option>)}
    </select>
  </div>
);

const Btn = ({ children, onClick, color, disabled, secondary, small }) => (
  <button onClick={onClick} disabled={disabled} style={{
    width: "100%", background: secondary ? "transparent" : (color || C.accent),
    color: secondary ? C.muted : "#fff",
    border: secondary ? `1px solid ${C.border}` : "none",
    borderRadius: 10, padding: small ? "10px" : "14px", fontSize: small ? 13 : 15, fontWeight: 700,
    cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.6 : 1, marginBottom: 10,
  }}>{children}</button>
);

const Badge = ({ label, color }) => (
  <span style={{ background: color+"22", color, border: `1px solid ${color}44`, borderRadius: 4, padding: "2px 8px", fontSize: 11, fontWeight: 700, textTransform: "uppercase", whiteSpace: "nowrap" }}>{label}</span>
);

const Banner = ({ msg, color, onDismiss }) => msg ? (
  <div style={{ background: color+"22", border: `1px solid ${color}44`, borderRadius: 8, padding: "12px 16px", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 14, color }}>
    {msg} <button onClick={onDismiss} style={{ background: "none", border: "none", color, cursor: "pointer", fontSize: 18 }}>×</button>
  </div>
) : null;

const SectionHeader = ({ title, onBack }) => (
  <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 20 }}>
    <button onClick={onBack} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, color: C.muted, cursor: "pointer", padding: "8px 12px", fontSize: 16 }}>←</button>
    <div style={{ fontSize: 16, fontWeight: 700, color: C.text }}>{title}</div>
  </div>
);

// ─── CHECKLIST VIEW ───────────────────────────────────────────────────────────
function ChecklistView({ asset, userRole, onDone, onBack }) {
  const [items, setItems] = useState([]);
  const [responses, setResponses] = useState({});
  const [executionId, setExecutionId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [executedBy, setExecutedBy] = useState(userRole?.name || "");
  const [error, setError] = useState(null);
  const [filterFreq, setFilterFreq] = useState("All");
  const FREQ_COLORS = { D: "#22c55e", W: "#3b82f6", F: "#a855f7", M: "#f97316" };
  const FREQ_LABELS = { D: "Daily", W: "Weekly", F: "Bi-weekly", M: "Monthly" };

  useEffect(() => { loadChecklist(); }, []);

  const loadChecklist = async () => {
    setLoading(true);
    const { data: checklists } = await supabase.from("checklists").select("id").limit(1);
    if (!checklists?.length) { setLoading(false); return; }
    const { data: checkItems } = await supabase.from("checklist_items").select("*").eq("checklist_id", checklists[0].id).order("item_number");
    setItems(checkItems || []);
    const now = new Date();
    const { data: existing } = await supabase.from("checklist_executions").select("*").eq("asset_id", asset.id).eq("month", now.getMonth()+1).eq("year", now.getFullYear()).limit(1);
    if (existing?.length) {
      setExecutionId(existing[0].id);
      setExecutedBy(existing[0].executed_by || userRole?.name || "");
      const { data: resps } = await supabase.from("checklist_responses").select("*").eq("execution_id", existing[0].id);
      const respMap = {};
      resps?.forEach(r => { respMap[r.item_id] = r; });
      setResponses(respMap);
    }
    setLoading(false);
  };

  const startExecution = async () => {
    if (!executedBy) { setError("Please enter your name."); return; }
    const now = new Date();
    const execId = uid("EXC");
    const { data: checklists } = await supabase.from("checklists").select("id").limit(1);
    await supabase.from("checklist_executions").insert([{ id: execId, checklist_id: checklists[0].id, asset_id: asset.id, asset_name: asset.name, executed_by: executedBy, execution_date: TODAY, month: now.getMonth()+1, year: now.getFullYear(), status: "In Progress" }]);
    setExecutionId(execId);
  };

  const setResponse = async (itemId, result, notes="") => {
    if (!executionId) return;
    const existing = responses[itemId];
    if (existing) {
      await supabase.from("checklist_responses").update({ result, notes }).eq("id", existing.id);
      setResponses(prev => ({ ...prev, [itemId]: { ...existing, result, notes } }));
    } else {
      const record = { id: uid("RSP"), execution_id: executionId, item_id: itemId, result, notes };
      await supabase.from("checklist_responses").insert([record]);
      setResponses(prev => ({ ...prev, [itemId]: record }));
    }
  };

  const complete = async () => {
    if (!executionId) return;
    setSaving(true);
    await supabase.from("checklist_executions").update({ status: "Completed" }).eq("id", executionId);
    const passCount = Object.values(responses).filter(r => r.result==="PASS").length;
    const failCount = Object.values(responses).filter(r => r.result==="FAIL").length;
    const naCount = Object.values(responses).filter(r => r.result==="N/A").length;
    const defects = items.filter(i => responses[i.id]?.result==="FAIL").map(i => `- ${i.item_en}: ${responses[i.id]?.notes||"No notes"}`).join("\n");
    const description = `CIL Checklist completed by ${executedBy}\n\nResults: ${passCount} PASS · ${failCount} FAIL · ${naCount} N/A\n${defects ? `\nDefects:\n${defects}` : "\nNo defects found."}`;
    const needsApproval = userRole?.role === "maintenance";
    const downtimeStart = brk?.downtime_start ? brk.downtime_start.split("T")[0] : TODAY;
    await supabase.from("maintenance_logs").insert([{ id: uid("LOG"), asset_id: asset.id, asset_name: asset.name, log_type: "Corrective Repair", title: `Breakdown Resolved — ${brk?.severity||""} severity`, description: `BREAKDOWN REPORTED BY: ${brk?.reported_by}\n\nISSUE: ${brk?.description}\n\nMAINTENANCE NOTES: ${resolveForm.notes}`, performed_by: userRole?.name, vendor: resolveForm.vendor==="— None —"?null:resolveForm.vendor||null, start_date: downtimeStart, end_date: TODAY, cost: null, status: "Completed", approval_status: isSupervisor ? "Approved" : "Pending", approved_by: isSupervisor ? userRole?.name : null, approved_at: isSupervisor ? new Date().toISOString() : null, downtime_start: downtimeStart, downtime_end: TODAY, downtime_hours: mins }]);
    const { data: openWOs } = await supabase.from("work_orders").select("id").eq("asset", asset.name).in("status", ["Open","In Progress","Pending"]).ilike("title","PM - %");
    if (openWOs?.length) await supabase.from("work_orders").update({ status: "Completed" }).in("id", openWOs.map(w => w.id));
    onDone();
    setSaving(false);
  };

  const filteredItems = filterFreq === "All" ? items : items.filter(i => i.frequency === filterFreq);
  const answeredCount = filteredItems.filter(i => responses[i.id]).length;
  const failCount = filteredItems.filter(i => responses[i.id]?.result==="FAIL").length;

  return (
    <div>
      <SectionHeader title="📋 CIL Checklist" onBack={onBack} />
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 20 }}>
        <div style={{ marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: C.muted, marginBottom: 6 }}>
            <span>Progress: {answeredCount}/{filteredItems.length}</span>
            <span style={{ color: failCount > 0 ? C.red : C.green }}>{failCount} FAIL</span>
          </div>
          <div style={{ background: C.border, borderRadius: 4, height: 8 }}>
            <div style={{ background: failCount > 0 ? C.accent : C.green, width: `${filteredItems.length > 0 ? (answeredCount/filteredItems.length)*100 : 0}%`, height: 8, borderRadius: 4 }} />
          </div>
        </div>
        <div style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
          {["All","D","W","F","M"].map(f => (
            <button key={f} onClick={() => setFilterFreq(f)} style={{ background: filterFreq===f?(FREQ_COLORS[f]||C.accent):C.surface, color: filterFreq===f?"#fff":C.muted, border: `1px solid ${filterFreq===f?(FREQ_COLORS[f]||C.accent):C.border}`, borderRadius: 6, padding: "4px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
              {f==="All"?"All":FREQ_LABELS[f]}
            </button>
          ))}
        </div>
        {!executionId && (
          <div style={{ background: C.surface, border: `1px solid ${C.accent}44`, borderRadius: 8, padding: 16, marginBottom: 16 }}>
            <Input label="Technician Name *" value={executedBy} onChange={setExecutedBy} />
            {error && <div style={{ color: C.red, fontSize: 12, marginTop: 8 }}>{error}</div>}
            <div style={{ marginTop: 10 }}><Btn onClick={startExecution} color={C.accent}>Start Checklist</Btn></div>
          </div>
        )}
        {loading ? <div style={{ textAlign: "center", color: C.muted, padding: 20 }}>Loading...</div> : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: "50vh", overflowY: "auto" }}>
            {filteredItems.map(item => {
              const resp = responses[item.id];
              const result = resp?.result;
              return (
                <div key={item.id} style={{ background: result==="PASS"?C.green+"11":result==="FAIL"?C.red+"11":C.surface, border: `1px solid ${result==="PASS"?C.green+"44":result==="FAIL"?C.red+"44":C.border}`, borderRadius: 8, padding: 12 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 13, color: C.text, fontWeight: 600 }}>{item.item_en}</div>
                      <div style={{ fontSize: 11, color: C.muted, marginTop: 2, direction: "rtl" }}>{item.item_ar}</div>
                    </div>
                    {executionId && (
                      <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                        <button onClick={() => setResponse(item.id,"PASS")} style={{ background: result==="PASS"?C.green:"transparent", color: result==="PASS"?"#fff":C.green, border: `2px solid ${C.green}`, borderRadius: 6, padding: "5px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✓</button>
                        <button onClick={() => setResponse(item.id,"FAIL")} style={{ background: result==="FAIL"?C.red:"transparent", color: result==="FAIL"?"#fff":C.red, border: `2px solid ${C.red}`, borderRadius: 6, padding: "5px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✗</button>
                        <button onClick={() => setResponse(item.id,"N/A")} style={{ background: result==="N/A"?C.muted:"transparent", color: result==="N/A"?"#fff":C.muted, border: `2px solid ${C.muted}44`, borderRadius: 6, padding: "5px 6px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>N/A</button>
                      </div>
                    )}
                  </div>
                  {result==="FAIL" && executionId && (
                    <textarea placeholder="Describe the issue..." value={resp?.notes||""} onChange={e => setResponse(item.id,"FAIL",e.target.value)} rows={2}
                      style={{ width: "100%", background: C.card, border: `1px solid ${C.red}44`, borderRadius: 6, padding: "8px", color: C.text, fontSize: 12, boxSizing: "border-box", resize: "vertical", marginTop: 8 }} />
                  )}
                </div>
              );
            })}
          </div>
        )}
        {executionId && (
          <div style={{ marginTop: 16 }}>
            <Btn onClick={complete} disabled={saving} color={C.green}>{saving ? "Completing..." : "✓ Complete Checklist"}</Btn>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── MAIN ASSET PAGE ──────────────────────────────────────────────────────────
export default function AssetPage() {
  const [asset, setAsset] = useState(null);
  const [lang, setLang] = useState("en");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [userRole, setUserRole] = useState(null);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [signing, setSigning] = useState(false);
  const [view, setView] = useState("home");
  const [saving, setSaving] = useState(false);
  const [vendors, setVendors] = useState([]);
  const [logs, setLogs] = useState([]);
  const [loadingLogs, setLoadingLogs] = useState(false);
  const [breakdowns, setBreakdowns] = useState([]);
  const [issues, setIssues] = useState([]);
  const [workOrders, setWorkOrders] = useState([]);
  const [loadingBreakdowns, setLoadingBreakdowns] = useState(false);
  const [catalogParts, setCatalogParts] = useState([]);

  const assetId = new URLSearchParams(window.location.search).get("asset");

  // Role helpers
  const isAdmin = userRole?.role === "admin";
  const isSupervisor = userRole?.role === "supervisor" || isAdmin;
  const isMaintenance = userRole?.role === "maintenance" || isSupervisor;
  const isOperations = userRole?.role === "operations";

  useEffect(() => {
    if (!assetId) { setError("No asset ID in URL."); setLoading(false); return; }
    loadAsset();
  }, [assetId]);

  useEffect(() => {
    supabase.auth.getSession().then(async ({ data: { session } }) => {
      if (session) {
        const { data } = await supabase.from("user_roles").select("*").eq("email", session.user.email).single();
        setUserRole(data || { role: "operations", name: session.user.email });
        if (data?.language) setLang(data.language);
        loadVendors();
      }
    });
  }, []);

  const loadAsset = async () => {
    const { data } = await supabase.from("assets").select("*").eq("id", assetId).single();
    if (data) setAsset(data);
    else setError("Asset not found.");
    setLoading(false);
  };

  const loadVendors = async () => {
    const { data } = await supabase.from("vendors").select("*").eq("status", "Active");
    setVendors(data || []);
  };

  const loadLogs = async () => {
    setLoadingLogs(true);
    const { data } = await supabase.from("maintenance_logs").select("*").eq("asset_id", assetId).order("start_date", { ascending: false });
    setLogs(data || []);
    setLoadingLogs(false);
  };

  const loadBreakdownsAndIssues = async () => {
    setLoadingBreakdowns(true);
    const [bRes, iRes, wRes] = await Promise.all([
      supabase.from("breakdown_reports").select("*").eq("asset_id", assetId).order("reported_at", { ascending: false }),
      supabase.from("issue_reports").select("*").eq("asset_id", assetId).order("reported_at", { ascending: false }),
      supabase.from("work_orders").select("*").eq("asset_id", assetId).order("due", { ascending: true }),
    ]);
    setBreakdowns(bRes.data || []);
    setIssues(iRes.data || []);
    setWorkOrders(wRes.data || []);
    setLoadingBreakdowns(false);
  };

  const loadCatalog = async () => {
    if (!asset) return;
    const [apRes, mpRes] = await Promise.all([
      supabase.from("asset_parts").select("*").eq("asset_id", asset.id).order("part_name"),
      asset.model ? supabase.from("model_parts").select("*").eq("model", asset.model).order("part_name") : Promise.resolve({ data: [] }),
    ]);
    const assetParts = apRes.data || [];
    const modelParts = (mpRes.data || []).map(p => ({ ...p, _id: p.id, id: `mdl-${p.id}`, model_part_id: p.id, isModelLevel: true }));
    const merged = [...assetParts];
    modelParts.forEach(mp => { if (!merged.find(ap => ap.part_name === mp.part_name)) merged.push(mp); });
    setCatalogParts(merged);
  };

  const signIn = async () => {
    setSigning(true);
    const { error: err } = await supabase.auth.signInWithPassword({ email, password });
    if (err) { setError(err.message); setSigning(false); return; }
    const { data } = await supabase.from("user_roles").select("*").eq("email", email).single();
    setUserRole(data || { role: "operations", name: email });
    if (data?.language) setLang(data.language);
    loadVendors();
    setSigning(false);
  };

  const toggleLanguage = async () => {
    const newLang = lang === "en" ? "ar" : "en";
    setLang(newLang);
    if (userRole) {
      await supabase.from("user_roles").update({ language: newLang }).eq("email", userRole.email || (await supabase.auth.getSession()).data.session?.user?.email);
    }
  };

  // ─── Breakdown Form ───────────────────────────────────────────────────────
  const [brkForm, setBrkForm] = useState({ reported_by: userRole?.name || "", severity: "High", description: "" });

  useEffect(() => { if (userRole?.name) setBrkForm(p => ({ ...p, reported_by: userRole.name })); }, [userRole]);
  const bf = (k) => (v) => setBrkForm(p => ({ ...p, [k]: v }));

  const submitBreakdown = async () => {
    if (!brkForm.description || !brkForm.reported_by) { setError("Please fill all fields."); return; }
    setSaving(true);
    const now = new Date().toISOString();
    const record = { id: uid("BRK"), asset_id: asset.id, asset_name: asset.name, site: asset.location, reported_by: brkForm.reported_by, reported_at: now, downtime_start: now, description: brkForm.description, severity: brkForm.severity, status: "Open" };
    const { error: err } = await supabase.from("breakdown_reports").insert([record]);
    if (err) { setError(err.message); } else {
      await supabase.from("assets").update({ status: "Under Maintenance" }).eq("id", asset.id);
      setAsset(prev => ({ ...prev, status: "Under Maintenance" }));
      try {
        await fetch("https://evwsdzqgvrwbjusjmrdc.supabase.co/functions/v1/notify-breakdown", { method: "POST", headers: { "Content-Type": "application/json", "Authorization": `Bearer ${import.meta.env.VITE_SUPABASE_ANON_KEY}` }, body: JSON.stringify({ breakdown: record, type: "reported" }) });
      } catch (e) { console.error(e); }
      setSuccess("Breakdown reported! Maintenance team notified.");
      setView("home");
      setBrkForm({ reported_by: "", severity: "High", description: "" });
    }
    setSaving(false);
  };

  // ─── Issue Form ───────────────────────────────────────────────────────────
  const [issForm, setIssForm] = useState({ reported_by: userRole?.name || "", severity: "Medium", description: "" });

  useEffect(() => { if (userRole?.name) setIssForm(p => ({ ...p, reported_by: userRole.name })); }, [userRole]);
  const isf = (k) => (v) => setIssForm(p => ({ ...p, [k]: v }));

  const submitIssue = async () => {
    if (!issForm.description || !issForm.reported_by) { setError("Please fill all fields."); return; }
    setSaving(true);
    const now = new Date().toISOString();
    const issueId = uid("ISS");
    const woId = uid("WO");
    const record = { id: issueId, asset_id: asset.id, asset_name: asset.name, site: asset.location, reported_by: issForm.reported_by, reported_at: now, description: issForm.description, severity: issForm.severity, status: "Open", work_order_id: woId };
    const { error: err } = await supabase.from("issue_reports").insert([record]);
    if (!err) {
      await supabase.from("work_orders").insert([{ id: woId, title: `Issue — ${asset.name}: ${issForm.description.slice(0,50)}`, asset: asset.name, priority: issForm.severity === "Critical" ? "Critical" : "Medium", status: "Open", start_date: now.split("T")[0], due: null, vendor: null }]);
      try {
        await fetch("https://evwsdzqgvrwbjusjmrdc.supabase.co/functions/v1/notify-breakdown", { method: "POST", headers: { "Content-Type": "application/json", "Authorization": `Bearer ${import.meta.env.VITE_SUPABASE_ANON_KEY}` }, body: JSON.stringify({ breakdown: { ...record, type: "issue" }, type: "reported" }) });
      } catch (e) { console.error(e); }
      setSuccess("Issue reported! Work order created and team notified.");
      setView("home");
      setIssForm({ reported_by: "", severity: "Medium", description: "" });
    } else { setError(err.message); }
    setSaving(false);
  };

  // ─── Resolve Breakdown ────────────────────────────────────────────────────
  const [resolveForm, setResolveForm] = useState({ breakdown_id: "", notes: "", vendor: "", downtime_start: "", reported_by: "", description: "", severity: "" });
  const rf = (k) => (v) => setResolveForm(p => ({ ...p, [k]: v }));

  const submitResolve = async () => {
    if (!resolveForm.notes || !resolveForm.breakdown_id) { setError("Please fill all fields."); return; }
    setSaving(true);
    const now = new Date().toISOString();
    const downtimeStartTs = resolveForm.downtime_start;
    const mins = downtimeStartTs ? Math.round((new Date(now) - new Date(downtimeStartTs.endsWith("Z") ? downtimeStartTs : downtimeStartTs+"Z")) / (1000*60)) : null;
    const downtimeStartDate = downtimeStartTs ? downtimeStartTs.split("T")[0] : TODAY;
    const isSupervisorOrAdmin = userRole?.role === "supervisor" || userRole?.role === "admin";
    const newStatus = isSupervisorOrAdmin ? "Pending Operator Confirmation" : "Pending Supervisor Approval";
    await supabase.from("breakdown_reports").update({ status: newStatus, resolved_by: userRole?.name, resolved_at: now, downtime_end: now, downtime_hours: mins, maintenance_notes: resolveForm.notes, supervisor_approved_by: isSupervisorOrAdmin ? userRole?.name : null, supervisor_approved_at: isSupervisorOrAdmin ? now : null }).eq("id", resolveForm.breakdown_id);
    await supabase.from("maintenance_logs").insert([{
      id: uid("LOG"), asset_id: asset.id, asset_name: asset.name,
      log_type: "Corrective Repair",
      title: `Breakdown Repair — ${resolveForm.severity||""} severity`,
      description: `BREAKDOWN REPORTED BY: ${resolveForm.reported_by||""}\n\nISSUE: ${resolveForm.description||""}\n\nMAINTENANCE NOTES: ${resolveForm.notes}`,
      performed_by: userRole?.name,
      vendor: resolveForm.vendor==="— None —"?null:resolveForm.vendor||null,
      start_date: downtimeStartDate,
      end_date: TODAY,
      cost: null,
      status: isSupervisorOrAdmin ? "Completed" : "In Progress",
      approval_status: isSupervisorOrAdmin ? "Approved" : "Pending",
      approved_by: isSupervisorOrAdmin ? userRole?.name : null,
      approved_at: isSupervisorOrAdmin ? new Date().toISOString() : null,
      breakdown_id: resolveForm.breakdown_id,
      downtime_start: downtimeStartDate,
      downtime_end: TODAY,
      downtime_hours: mins,
    }]);
    setSuccess(isSupervisorOrAdmin ? "Breakdown resolved — pending operator confirmation." : "Breakdown resolved — pending supervisor approval.");
    setBreakdowns(prev => prev.map(b => b.id===resolveForm.breakdown_id?{...b,status:newStatus}:b));
    setView("breakdowns");
    setResolveForm({ breakdown_id: "", notes: "", vendor: "", downtime_start: "", reported_by: "", description: "", severity: "" });
    setSaving(false);
  };

  // ─── Maintenance Log Form ─────────────────────────────────────────────────
  const [logForm, setLogForm] = useState({ log_type: "Corrective Repair", title: "", description: "", performed_by: "", cost: "", vendor: "" });
  const lf = (k) => (v) => setLogForm(p => ({ ...p, [k]: v }));

  const submitLog = async () => {
    if (!logForm.title) { setError("Title is required."); return; }
    setSaving(true);
    const needsApproval = userRole?.role === "maintenance";
    const record = { id: uid("LOG"), asset_id: asset.id, asset_name: asset.name, log_type: logForm.log_type, title: logForm.title, description: logForm.description, performed_by: logForm.performed_by || userRole?.name, vendor: logForm.vendor==="— None —"?null:logForm.vendor||null, start_date: TODAY, end_date: TODAY, cost: logForm.cost ? parseFloat(logForm.cost) : null, status: needsApproval ? "In Progress" : "Completed", approval_status: needsApproval ? "Pending" : "Approved", approved_by: needsApproval ? null : userRole?.name, approved_at: needsApproval ? null : new Date().toISOString() };
    const { error: err } = await supabase.from("maintenance_logs").insert([record]);
    if (err) { setError(err.message); } else {
      setSuccess(needsApproval ? "Log saved! Awaiting supervisor approval." : "Maintenance log added!");
      await loadLogs();
      setView("history");
      setLogForm({ log_type: "Corrective Repair", title: "", description: "", performed_by: "", cost: "", vendor: "" });
    }
    setSaving(false);
  };

  // ─── Work Order Status Update ─────────────────────────────────────────────
  const updateWOStatus = async (woId, newStatus, note) => {
    await supabase.from("work_orders").update({ status: newStatus, status_note: note || null }).eq("id", woId);
    setWorkOrders(prev => prev.map(w => w.id===woId?{...w,status:newStatus,status_note:note}:w));
    setSuccess(`Work order updated to ${newStatus}`);
  };

  // ─── Spare Parts Form ─────────────────────────────────────────────────────
  const [partForm, setPartForm] = useState({ log_id: "", part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "", asset_part_id: null, model_part_id: null });
  const pf = (k) => (v) => setPartForm(p => ({ ...p, [k]: v }));

  const submitPart = async () => {
    if (!partForm.part_name || !partForm.log_id) { setError("Part name and log are required."); return; }
    setSaving(true);
    const qty = parseFloat(partForm.quantity) || 1;
    const unitCost = parseFloat(partForm.unit_cost) || 0;
    const rawId = partForm.asset_part_id;
    const isModelLevel = rawId && String(rawId).startsWith("mdl-");
    const cleanAssetPartId = isModelLevel ? null : rawId || null;
    const cleanModelPartId = partForm.model_part_id || null;
    const record = { id: uid("PRT"), log_id: partForm.log_id, asset_id: asset.id, part_name: partForm.part_name, part_number: partForm.part_number, quantity: qty, unit_cost: unitCost, total_cost: qty * unitCost, supplier: partForm.supplier, asset_part_id: cleanAssetPartId, model_part_id: cleanModelPartId };
    const { error: err } = await supabase.from("spare_parts").insert([record]);
    if (err) { setError(err.message); } else {
      // Deduct stock if log is approved
      const log = logs.find(l => l.id === partForm.log_id);
      if (log?.approval_status === "Approved") {
        if (cleanModelPartId) { const { data: mp } = await supabase.from("model_parts").select("stock_quantity").eq("id", cleanModelPartId).single(); if (mp) await supabase.from("model_parts").update({ stock_quantity: Math.max(0,(mp.stock_quantity||0)-qty) }).eq("id", cleanModelPartId); }
        if (cleanAssetPartId) { const { data: ap } = await supabase.from("asset_parts").select("stock_quantity").eq("id", cleanAssetPartId).single(); if (ap) await supabase.from("asset_parts").update({ stock_quantity: Math.max(0,(ap.stock_quantity||0)-qty) }).eq("id", cleanAssetPartId); }
      }
      setSuccess("Spare part added!");
      setPartForm({ log_id: "", part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "", asset_part_id: null, model_part_id: null });
    }
    setSaving(false);
  };

  // ─── Approve Log ─────────────────────────────────────────────────────────
  const approveLog = async (logId) => {
    const now = new Date().toISOString();
    await supabase.from("maintenance_logs").update({ approval_status: "Approved", approved_by: userRole?.name, approved_at: now, status: "Completed" }).eq("id", logId);
    const { data: parts } = await supabase.from("spare_parts").select("*").eq("log_id", logId);
    if (parts?.length) {
      for (const part of parts) {
        if (part.model_part_id) { const { data: mp } = await supabase.from("model_parts").select("stock_quantity").eq("id", part.model_part_id).single(); if (mp) await supabase.from("model_parts").update({ stock_quantity: Math.max(0,(mp.stock_quantity||0)-(part.quantity||1)) }).eq("id", part.model_part_id); }
        if (part.asset_part_id) { const { data: ap } = await supabase.from("asset_parts").select("stock_quantity").eq("id", part.asset_part_id).single(); if (ap) await supabase.from("asset_parts").update({ stock_quantity: Math.max(0,(ap.stock_quantity||0)-(part.quantity||1)) }).eq("id", part.asset_part_id); }
      }
    }
    setLogs(prev => prev.map(l => l.id===logId?{...l,approval_status:"Approved",approved_by:userRole?.name,status:"Completed"}:l));
    setSuccess("Log approved!");
  };

  const rejectLog = async (logId, reason) => {
    await supabase.from("maintenance_logs").update({ approval_status: "Rejected", rejection_notes: reason, status: "In Progress" }).eq("id", logId);
    setLogs(prev => prev.map(l => l.id===logId?{...l,approval_status:"Rejected",rejection_notes:reason}:l));
    setSuccess("Log rejected and sent back.");
  };

  if (loading) return <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", color: C.muted, fontFamily: "Arial, sans-serif" }}>Loading...</div>;
  if (error && !asset) return <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", color: C.red, fontFamily: "Arial, sans-serif", padding: 20 }}>{error}</div>;

  const vendorOptions = ["— None —", ...vendors.map(v => v.name)];
  const openBreakdowns = breakdowns.filter(b => b.status !== "Resolved");
  const openIssues = issues.filter(i => i.status !== "Resolved");
  const activeWOs = workOrders.filter(w => w.status !== "Completed");

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "Arial, sans-serif", color: C.text, maxWidth: 520, margin: "0 auto", padding: 16 }}>

      {/* Header */}
      <div style={{ background: C.accent, borderRadius: 12, padding: 20, marginBottom: 20, textAlign: "center" }}>
        <div style={{ fontSize: 32, marginBottom: 8 }}>🏭</div>
        <div style={{ fontSize: 18, fontWeight: 700, color: "#fff" }}>{asset?.name}</div>
        <div style={{ fontSize: 13, color: "rgba(255,255,255,0.8)", marginTop: 4 }}>{asset?.category} · {asset?.location}</div>
        {asset?.brand && <div style={{ fontSize: 12, color: "rgba(255,255,255,0.7)", marginTop: 2 }}>{asset.brand} {asset.model}</div>}
        <div style={{ marginTop: 10, display: "inline-block", background: statusColor(asset?.status)+"33", color: statusColor(asset?.status), border: `1px solid ${statusColor(asset?.status)}66`, borderRadius: 6, padding: "4px 12px", fontSize: 12, fontWeight: 700, textTransform: "uppercase" }}>
          {asset?.status}
        </div>
      </div>

      <Banner msg={error} color={C.red} onDismiss={() => setError(null)} />
      <Banner msg={success} color={C.green} onDismiss={() => setSuccess(null)} />

      {/* Login */}
      {!userRole ? (
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 20 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.text, marginBottom: 16 }}>Sign in to continue</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 14, marginBottom: 16 }}>
            <Input label="Email" value={email} onChange={setEmail} type="email" />
            <Input label="Password" value={password} onChange={setPassword} type="password" />
          </div>
          <Btn onClick={signIn} disabled={signing}>{signing ? "Signing in..." : "Sign In"}</Btn>
        </div>

      ) : view === "home" ? (
        <>
          {/* Role badge */}
          <div style={{ textAlign: "center", marginBottom: 16, fontSize: 13, color: C.muted }}>
            <strong style={{ color: C.accent }}>{userRole.name}</strong> · <span style={{ textTransform: "capitalize" }}>{userRole.role}</span>
          </div>

          {/* Quick Stats for maintenance+ */}
          {isMaintenance && (
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 16 }}>
              {[["🚨", openBreakdowns.length, "Breakdowns", C.red],["⚠️", openIssues.length, "Issues", C.yellow],["📋", activeWOs.length, "Work Orders", C.blue]].map(([icon,val,label,color]) => (
                <div key={label} onClick={() => { loadBreakdownsAndIssues(); setView(label==="Breakdowns"?"breakdowns":label==="Issues"?"issues":"workorders"); }} style={{ background: C.card, border: `1px solid ${color}33`, borderRadius: 10, padding: "12px 8px", textAlign: "center", cursor: "pointer" }}>
                  <div style={{ fontSize: 20 }}>{icon}</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color, fontFamily: "monospace" }}>{val}</div>
                  <div style={{ fontSize: 10, color: C.muted }}>{label}</div>
                </div>
              ))}
            </div>
          )}

          {/* Action Buttons */}
          <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
            {/* Everyone can report */}
            <Btn onClick={() => setView("breakdown")} color={C.red}>🚨 Report Breakdown</Btn>
            <Btn onClick={() => setView("issue")} color={C.yellow}>⚠️ Report Issue</Btn>

            {/* Operations: view open breakdowns/issues */}
            {isOperations && (
              <Btn onClick={() => { loadBreakdownsAndIssues(); setView("breakdowns"); }} secondary>👁 View Open Breakdowns & Issues</Btn>
            )}

            {/* Maintenance+ */}
            {isMaintenance && (
              <>
                <Btn onClick={() => { loadBreakdownsAndIssues(); setView("breakdowns"); }} color={C.red}>🚨 Breakdowns & Issues</Btn>
                <Btn onClick={() => { loadBreakdownsAndIssues(); setView("workorders"); }} color={C.blue}>📋 Work Orders</Btn>
                <Btn onClick={() => { loadLogs(); loadCatalog(); setView("log"); }} color={C.blue}>🔧 Add Maintenance Log</Btn>
                <Btn onClick={() => setView("checklist")} color={C.green}>📋 Run PM Checklist</Btn>
                <Btn onClick={() => { loadLogs(); setView("parts"); }} color={C.purple}>🔩 Add Spare Parts</Btn>
                <Btn onClick={() => setView("specs")} secondary>📄 Equipment Specs</Btn>
                <Btn onClick={() => { loadLogs(); setView("history"); }} secondary>📜 Maintenance History</Btn>
              </>
            )}
          </div>
        </>

      ) : view === "breakdown" ? (
        <div>
          <SectionHeader title="🚨 Report Breakdown" onBack={() => setView("home")} />
          <div style={{ background: C.card, border: `1px solid ${C.red}44`, borderRadius: 12, padding: 20 }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <div>
                <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>Your Name</div>
                <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px", color: C.text, fontSize: 15 }}>{userRole?.name || "—"}</div>
              </div>
              <Sel label="Severity" value={brkForm.severity} onChange={bf("severity")} options={["Critical","High","Medium","Low"]} />
              <Textarea label="Describe the issue *" value={brkForm.description} onChange={bf("description")} placeholder="What happened? Any error messages?" />
              <div style={{ background: C.yellow+"22", border: `1px solid ${C.yellow}44`, borderRadius: 8, padding: 12, fontSize: 13, color: C.yellow }}>
                ⏱ Downtime starts now: {new Date().toLocaleString("en-GB")}
              </div>
            </div>
            <div style={{ marginTop: 16 }}>
              <Btn onClick={submitBreakdown} disabled={saving} color={C.red}>{saving ? "Reporting..." : "🚨 Report Breakdown"}</Btn>
            </div>
          </div>
        </div>

      ) : view === "issue" ? (
        <div>
          <SectionHeader title="⚠️ Report Issue" onBack={() => setView("home")} />
          <div style={{ background: C.card, border: `1px solid ${C.yellow}44`, borderRadius: 12, padding: 20 }}>
            <div style={{ background: C.green+"11", border: `1px solid ${C.green}33`, borderRadius: 8, padding: 10, fontSize: 13, color: C.green, marginBottom: 14 }}>
              ✅ Equipment stays RUNNING — this creates a work order for maintenance
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <div>
                <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>Your Name</div>
                <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px", color: C.text, fontSize: 15 }}>{userRole?.name || "—"}</div>
              </div>
              <Sel label="Severity" value={issForm.severity} onChange={isf("severity")} options={["Critical","High","Medium","Low"]} />
              <Textarea label="Describe the issue *" value={issForm.description} onChange={isf("description")} placeholder="Noise, leak, warning light, performance drop..." />
            </div>
            <div style={{ marginTop: 16 }}>
              <Btn onClick={submitIssue} disabled={saving} color={C.yellow}>{saving ? "Reporting..." : "⚠️ Report Issue"}</Btn>
            </div>
          </div>
        </div>

      ) : view === "breakdowns" ? (
        <div>
          <SectionHeader title="🚨 Breakdowns & Issues" onBack={() => setView("home")} />
          {loadingBreakdowns ? <div style={{ textAlign: "center", color: C.muted, padding: 20 }}>Loading...</div> : (
            <>
              {/* Breakdowns */}
              <div style={{ fontSize: 13, fontWeight: 700, color: C.red, marginBottom: 10 }}>🚨 Breakdowns ({openBreakdowns.length} open)</div>
              {breakdowns.length === 0 ? <div style={{ color: C.muted, fontSize: 13, marginBottom: 16 }}>No breakdowns.</div> : (
                <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 20 }}>
                  {breakdowns.map(b => (
                    <div key={b.id} style={{ background: C.card, border: `1px solid ${b.status==="Resolved"?C.green+"44":C.red+"44"}`, borderRadius: 10, padding: 14 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                        <div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{b.asset_name}</div>
                        <Badge label={b.status} color={statusColor(b.status)} />
                      </div>
                      <div style={{ fontSize: 12, color: C.muted, marginBottom: 6 }}>{b.reported_by} · {fmtDateTime(b.reported_at)}</div>
                      <div style={{ fontSize: 13, color: C.subtle, marginBottom: 10 }}>{b.description}</div>
                      {/* Maintenance can resolve */}
                      {isMaintenance && (b.status === "Open" || b.status === "Acknowledged") && (
                        <div style={{ marginTop: 8 }}>
                          {b.status === "Open" && (
                            <Btn small onClick={async () => { await supabase.from("breakdown_reports").update({ status: "Acknowledged", acknowledged_by: userRole?.name, acknowledged_at: new Date().toISOString() }).eq("id", b.id); setBreakdowns(prev => prev.map(x => x.id===b.id?{...x,status:"Acknowledged"}:x)); }} color={C.blue}>👁 Acknowledge</Btn>
                          )}
                          <Btn small onClick={() => { setResolveForm({ breakdown_id: b.id, notes: "", vendor: "", downtime_start: b.downtime_start, reported_by: b.reported_by, description: b.description, severity: b.severity }); setView("resolve"); }} color={C.green}>✅ Resolve Breakdown</Btn>
                        </div>
                      )}
                      {b.status === "Pending Supervisor Approval" && isSupervisor && (
                        <QRApprovalStep record={b} table="breakdown_reports" userRole={userRole} onDone={(updated) => setBreakdowns(prev => prev.map(x => x.id===updated.id?updated:x))} />
                      )}
                      {b.status === "Pending Operator Confirmation" && (
                        <QROperatorConfirm record={b} table="breakdown_reports" userRole={userRole} onDone={(updated) => { setBreakdowns(prev => prev.map(x => x.id===updated.id?updated:x)); setAsset(prev => ({ ...prev, status: "Operational" })); }} />
                      )}
                    </div>
                  ))}
                </div>
              )}

              {/* Issues */}
              <div style={{ fontSize: 13, fontWeight: 700, color: C.yellow, marginBottom: 10 }}>⚠️ Issues ({openIssues.length} open)</div>
              {issues.length === 0 ? <div style={{ color: C.muted, fontSize: 13 }}>No issues.</div> : (
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {issues.map(i => (
                    <div key={i.id} style={{ background: C.card, border: `1px solid ${i.status==="Resolved"?C.green+"44":C.yellow+"44"}`, borderRadius: 10, padding: 14 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                        <Badge label={i.severity} color={SEVERITY_COLORS[i.severity]||C.muted} />
                        <Badge label={i.status} color={statusColor(i.status)} />
                      </div>
                      <div style={{ fontSize: 12, color: C.muted, marginBottom: 6 }}>{i.reported_by} · {fmtDateTime(i.reported_at)}</div>
                      <div style={{ fontSize: 13, color: C.subtle, marginBottom: 10 }}>{i.description}</div>
                      {isMaintenance && (i.status === "Open" || i.status === "Acknowledged") && (
                        <div style={{ marginTop: 8 }}>
                          {i.status === "Open" && <Btn small onClick={async () => { await supabase.from("issue_reports").update({ status: "Acknowledged", acknowledged_by: userRole?.name, acknowledged_at: new Date().toISOString() }).eq("id", i.id); setIssues(prev => prev.map(x => x.id===i.id?{...x,status:"Acknowledged"}:x)); }} color={C.blue}>👁 Acknowledge</Btn>}
                          <Btn small onClick={async () => {
                            const now = new Date().toISOString();
                            const isSupervisorOrAdmin = userRole?.role === "supervisor" || userRole?.role === "admin";
                            const newStatus = isSupervisorOrAdmin ? "Pending Operator Confirmation" : "Pending Supervisor Approval";
                            await supabase.from("issue_reports").update({ status: newStatus, resolved_by: userRole?.name, resolved_at: now, supervisor_approved_by: isSupervisorOrAdmin ? userRole?.name : null, supervisor_approved_at: isSupervisorOrAdmin ? now : null }).eq("id", i.id);
                            setIssues(prev => prev.map(x => x.id===i.id?{...x,status:newStatus}:x));
                            setSuccess(isSupervisorOrAdmin ? "Resolved — pending operator confirmation." : "Resolved — pending supervisor approval.");
                          }} color={C.green}>✅ Resolve Issue</Btn>
                        </div>
                      )}
                      {i.status === "Pending Supervisor Approval" && isSupervisor && (
                        <QRApprovalStep record={i} table="issue_reports" userRole={userRole} onDone={(updated) => setIssues(prev => prev.map(x => x.id===updated.id?updated:x))} extraOnApprove={async () => { if (i.work_order_id) await supabase.from("work_orders").update({ status: "Completed" }).eq("id", i.work_order_id); }} />
                      )}
                      {i.status === "Pending Operator Confirmation" && (
                        <QROperatorConfirm record={i} table="issue_reports" userRole={userRole} onDone={(updated) => setIssues(prev => prev.map(x => x.id===updated.id?updated:x))} />
                      )}
                    </div>
                  ))}
                </div>
              )}
            </>
          )}
        </div>

      ) : view === "resolve" ? (
        <div>
          <SectionHeader title="✅ Resolve Breakdown" onBack={() => setView("breakdowns")} />
          <div style={{ background: C.card, border: `1px solid ${C.green}44`, borderRadius: 12, padding: 20 }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <Textarea label="Maintenance Notes *" value={resolveForm.notes} onChange={rf("notes")} placeholder="Root cause, what was done, parts replaced?" />
              <Sel label="Vendor / Contractor Used" value={resolveForm.vendor} onChange={rf("vendor")} options={vendorOptions} />
            </div>
            <div style={{ marginTop: 16 }}>
              <Btn onClick={submitResolve} disabled={saving} color={C.green}>{saving ? "Resolving..." : "✅ Mark as Resolved"}</Btn>
            </div>
          </div>
        </div>

      ) : view === "workorders" ? (
        <div>
          <SectionHeader title="📋 Work Orders" onBack={() => setView("home")} />
          {loadingBreakdowns ? <div style={{ textAlign: "center", color: C.muted, padding: 20 }}>Loading...</div> : workOrders.length === 0 ? (
            <div style={{ textAlign: "center", color: C.muted, padding: 40 }}>No work orders for this asset.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {workOrders.map(wo => (
                <WOCard key={wo.id} wo={wo} isMaintenance={isMaintenance} isSupervisor={isSupervisor} onUpdate={updateWOStatus} />
              ))}
            </div>
          )}
        </div>

      ) : view === "log" ? (
        <div>
          <SectionHeader title="🔧 Add Maintenance Log" onBack={() => setView("home")} />
          <div style={{ background: C.card, border: `1px solid ${C.blue}44`, borderRadius: 12, padding: 20 }}>
            {userRole?.role === "maintenance" && (
              <div style={{ background: C.yellow+"11", border: `1px solid ${C.yellow}33`, borderRadius: 8, padding: 10, fontSize: 13, color: C.yellow, marginBottom: 14 }}>
                ⏳ This log will need supervisor approval before it's closed.
              </div>
            )}
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <Sel label="Type" value={logForm.log_type} onChange={lf("log_type")} options={["Preventive Maintenance","Corrective Repair","Inspection","Overhaul","Part Replacement"]} />
              <Input label="Title *" value={logForm.title} onChange={lf("title")} placeholder="What was done?" />
              <Input label="Performed By" value={logForm.performed_by} onChange={lf("performed_by")} placeholder="Technician name" />
              <Sel label="Vendor" value={logForm.vendor} onChange={lf("vendor")} options={vendorOptions} />
              <Input label="Cost ($)" value={logForm.cost} onChange={lf("cost")} type="number" />
              <Textarea label="Notes" value={logForm.description} onChange={lf("description")} placeholder="Details about the work done..." />
            </div>
            <div style={{ marginTop: 16 }}>
              <Btn onClick={submitLog} disabled={saving} color={C.blue}>{saving ? "Saving..." : "Save Log"}</Btn>
            </div>
          </div>
        </div>

      ) : view === "parts" ? (
        <div>
          <SectionHeader title="🔩 Add Spare Parts" onBack={() => setView("home")} />
          <div style={{ background: C.card, border: `1px solid ${C.purple}44`, borderRadius: 12, padding: 20 }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <div>
                <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase" }}>Link to Maintenance Log *</div>
                <select value={partForm.log_id} onChange={e => setPartForm(p => ({ ...p, log_id: e.target.value }))}
                  style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px", color: C.text, fontSize: 15 }}>
                  <option value="">— Select Log —</option>
                  {loadingLogs ? <option>Loading...</option> : logs.map(l => <option key={l.id} value={l.id}>{l.title} ({l.start_date})</option>)}
                </select>
              </div>
              {catalogParts.length > 0 && (
                <div>
                  <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase" }}>Select from Catalog</div>
                  <select onChange={e => {
                    const found = catalogParts.find(p => p.id === e.target.value);
                    if (found) setPartForm(prev => ({ ...prev, part_name: found.part_name, part_number: found.part_number||"", unit_cost: String(found.unit_cost||""), supplier: found.supplier||"", asset_part_id: found.id, model_part_id: found.model_part_id||null }));
                  }} style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px", color: C.text, fontSize: 15 }}>
                    <option value="">— Select from catalog —</option>
                    {catalogParts.map(p => <option key={p.id} value={p.id}>{p.part_name}{p.part_number?` (${p.part_number})`:""} — ${p.unit_cost||0}</option>)}
                  </select>
                </div>
              )}
              <Input label="Part Name *" value={partForm.part_name} onChange={pf("part_name")} placeholder="e.g. Hydraulic filter" />
              <Input label="Part Number" value={partForm.part_number} onChange={pf("part_number")} />
              <Input label="Quantity" value={partForm.quantity} onChange={pf("quantity")} type="number" />
              <Input label="Unit Cost ($)" value={partForm.unit_cost} onChange={pf("unit_cost")} type="number" />
              <Input label="Supplier" value={partForm.supplier} onChange={pf("supplier")} />
            </div>
            <div style={{ marginTop: 16 }}>
              <Btn onClick={submitPart} disabled={saving} color={C.purple}>{saving ? "Saving..." : "🔩 Add Spare Part"}</Btn>
            </div>
          </div>
        </div>

      ) : view === "checklist" ? (
        <ChecklistView asset={asset} userRole={userRole} onDone={() => { setView("home"); setSuccess("Checklist completed!"); }} onBack={() => setView("home")} />

      ) : view === "specs" ? (
        <div>
          <SectionHeader title="📄 Equipment Specs" onBack={() => setView("home")} />
          <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 20 }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
              {[["Name",asset?.name],["Category",asset?.category],["Site",asset?.location],["Owner",asset?.owner||"—"],["Brand",asset?.brand||"—"],["Model",asset?.model||"—"],["Serial No.",asset?.serial_number||"—"],["Manufacture Date",asset?.manufacture_date?fmtDate(asset.manufacture_date):"—"],["Est. Value",asset?.value||"—"],["PM Every",asset?.pm_frequency?`${asset.pm_frequency} mo.`:"—"],["Last PM",asset?.last_pm_date?fmtDate(asset.last_pm_date):"Never"],["Next Service",asset?.next_service?fmtDate(asset.next_service):"—"]].map(([lbl,val]) => (
                <div key={lbl}>
                  <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase" }}>{lbl}</div>
                  <div style={{ fontSize: 13, color: C.subtle, marginTop: 2 }}>{val}</div>
                </div>
              ))}
            </div>
            {asset?.technical_specs && (
              <div style={{ background: C.surface, borderRadius: 8, padding: 14 }}>
                <div style={{ fontSize: 11, color: C.muted, textTransform: "uppercase", marginBottom: 8 }}>Technical Specifications</div>
                <div style={{ fontSize: 13, color: C.subtle, lineHeight: 1.8 }}>
                  {asset.technical_specs.split("|").map((s,i) => <div key={i}>· {s.trim()}</div>)}
                </div>
              </div>
            )}
            {asset?.pm_task && (
              <div style={{ marginTop: 14, background: C.blue+"11", border: `1px solid ${C.blue}33`, borderRadius: 8, padding: 12 }}>
                <div style={{ fontSize: 11, color: C.muted, textTransform: "uppercase", marginBottom: 6 }}>PM Task</div>
                <div style={{ fontSize: 13, color: C.subtle }}>{asset.pm_task}</div>
              </div>
            )}
          </div>
        </div>

      ) : view === "history" ? (
        <div>
          <SectionHeader title="📜 Maintenance History" onBack={() => setView("home")} />
          {loadingLogs ? <div style={{ textAlign: "center", color: C.muted, padding: 20 }}>Loading...</div> : logs.length === 0 ? (
            <div style={{ textAlign: "center", color: C.muted, padding: 40 }}>No maintenance records yet.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {logs.map(log => (
                <HistoryCard key={log.id} log={log} isSupervisor={isSupervisor} onApprove={approveLog} onReject={rejectLog} />
              ))}
            </div>
          )}
        </div>

      ) : null}

      {/* Sign out */}
      {userRole && (
        <div style={{ textAlign: "center", marginTop: 20 }}>
          <button onClick={async () => { await supabase.auth.signOut(); setUserRole(null); setEmail(""); setPassword(""); setView("home"); }} style={{ background: "transparent", border: `1px solid ${C.border}`, borderRadius: 8, padding: "10px 20px", color: C.muted, cursor: "pointer", fontSize: 13 }}>Sign Out</button>
        </div>
      )}

      {/* Footer */}
      <div style={{ textAlign: "center", marginTop: 24, fontSize: 11, color: C.muted }}>
        Facility Command · EPx Logistics
      </div>
    </div>
  );
}
// ─── QR APPROVAL STEPS ────────────────────────────────────────────────────────
function QRApprovalStep({ record, table, userRole, onDone, extraOnApprove }) {
  const [password, setPassword] = useState("");
  const [show, setShow] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);

  const approve = async () => {
    if (!password) { setError("Enter your password to approve"); return; }
    setSaving(true);
    const { data: sessionData } = await supabase.auth.getSession();
    const email = sessionData?.session?.user?.email;
    const { error: authErr } = await supabase.auth.signInWithPassword({ email, password });
    if (authErr) { setError("Incorrect password. Please try again."); setSaving(false); return; }
    const now = new Date().toISOString();
    await supabase.from(table).update({ status: "Pending Operator Confirmation", supervisor_approved_by: userRole?.name, supervisor_approved_at: now }).eq("id", record.id);
    if (extraOnApprove) await extraOnApprove();
    onDone({ ...record, status: "Pending Operator Confirmation", supervisor_approved_by: userRole?.name, supervisor_approved_at: now });
    setSaving(false);
  };

  return (
    <div style={{ background: C.yellow+"11", border: `1px solid ${C.yellow}44`, borderRadius: 10, padding: 14, marginTop: 10 }}>
      <div style={{ fontSize: 12, color: C.yellow, fontWeight: 700, marginBottom: 10 }}>⏳ Pending Supervisor Approval</div>
      {!show ? (
        <Btn onClick={() => setShow(true)} color={C.green}>✅ Approve</Btn>
      ) : (
        <div>
          <Input label="Confirm Your Password" value={password} type="password" onChange={v => { setPassword(v); setError(null); }} />
          {error && <div style={{ color: C.red, fontSize: 12, marginTop: 6 }}>{error}</div>}
          <div style={{ marginTop: 10 }}>
            <Btn onClick={approve} disabled={saving||!password} color={C.green}>{saving?"Verifying...":"✓ Confirm & Approve"}</Btn>
          </div>
        </div>
      )}
    </div>
  );
}

function QROperatorConfirm({ record, table, userRole, onDone }) {
  const [saving, setSaving] = useState(false);
  const isReporter = userRole?.name === record.reported_by;

  const confirm = async () => {
    setSaving(true);
    const now = new Date().toISOString();
    await supabase.from(table).update({ status: "Resolved", operator_confirmed_by: userRole?.name, operator_confirmed_at: now }).eq("id", record.id);
    if (table === "breakdown_reports") {
      await supabase.from("assets").update({ status: "Operational" }).eq("id", record.asset_id);
      await supabase.from("maintenance_logs").update({ approval_status: "Approved", approved_by: record.supervisor_approved_by || userRole?.name, approved_at: record.supervisor_approved_at || now, status: "Completed" }).eq("breakdown_id", record.id);
    } else {
      await supabase.from("maintenance_logs").update({ approval_status: "Approved", approved_by: record.supervisor_approved_by || userRole?.name, approved_at: record.supervisor_approved_at || now, status: "Completed" }).eq("issue_id", record.id);
    }
    onDone({ ...record, status: "Resolved", operator_confirmed_by: userRole?.name, operator_confirmed_at: now });
    setSaving(false);
  };

  return (
    <div style={{ background: C.blue+"11", border: `1px solid ${C.blue}44`, borderRadius: 10, padding: 14, marginTop: 10 }}>
      <div style={{ fontSize: 12, color: C.blue, fontWeight: 700, marginBottom: 10 }}>👁 Pending Operator Confirmation</div>
      {isReporter ? (
        <Btn onClick={confirm} disabled={saving} color={C.green}>{saving?"Confirming...":"✅ Confirm Equipment is Back to Normal"}</Btn>
      ) : (
        <div style={{ fontSize: 12, color: C.muted }}>Awaiting confirmation from {record.reported_by}</div>
      )}
    </div>
  );
}
// ─── WORK ORDER CARD ──────────────────────────────────────────────────────────
function WOCard({ wo, isMaintenance, isSupervisor, onUpdate }) {
  const [expanded, setExpanded] = useState(false);
  const [showUpdate, setShowUpdate] = useState(false);
  const [newStatus, setNewStatus] = useState(wo.status);
  const [note, setNote] = useState(wo.status_note || "");
  const isOverdue = wo.due && wo.due <= new Date().toISOString().split("T")[0] && wo.status !== "Completed";
  const WO_STATUSES = ["Open","In Progress","Awaiting PO","Awaiting Parts","Awaiting Approval","On Hold","Scheduled","Completed"];

  return (
    <div style={{ background: "#1a1e2a", border: `1px solid ${isOverdue?"#ef444444":wo.status==="Completed"?"#22c55e44":"#252b3b"}`, borderRadius: 10, overflow: "hidden" }}>
      <div onClick={() => setExpanded(!expanded)} style={{ padding: 14, cursor: "pointer" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
          <div style={{ fontSize: 14, fontWeight: 700, color: "#e2e8f0", flex: 1, marginRight: 8 }}>{wo.title}</div>
          <span style={{ background: statusColor(wo.status)+"22", color: statusColor(wo.status), border: `1px solid ${statusColor(wo.status)}44`, borderRadius: 4, padding: "2px 8px", fontSize: 11, fontWeight: 700, whiteSpace: "nowrap" }}>{wo.status}</span>
        </div>
        <div style={{ fontSize: 12, color: "#64748b" }}>
          {wo.category && <span>{wo.category} · </span>}
          {wo.due && <span style={{ color: isOverdue?"#ef4444":"#64748b" }}>Due: {wo.due}</span>}
          {wo.vendor && <span> · {wo.vendor}</span>}
        </div>
        {wo.status_note && <div style={{ marginTop: 6, fontSize: 12, color: "#f97316", background: "#f9731611", borderRadius: 4, padding: "4px 8px" }}>📝 {wo.status_note}</div>}
      </div>
      {expanded && isMaintenance && wo.status !== "Completed" && (
        <div style={{ padding: "0 14px 14px", borderTop: "1px solid #252b3b" }}>
          {!showUpdate ? (
            <button onClick={() => setShowUpdate(true)} style={{ background: "#3b82f622", color: "#3b82f6", border: "1px solid #3b82f644", borderRadius: 8, padding: "10px", width: "100%", fontSize: 14, fontWeight: 700, cursor: "pointer", marginTop: 12 }}>✏️ Update Status</button>
          ) : (
            <div style={{ marginTop: 12 }}>
              <div style={{ marginBottom: 10 }}>
                <div style={{ fontSize: 11, color: "#64748b", marginBottom: 4, textTransform: "uppercase" }}>New Status</div>
                <select value={newStatus} onChange={e => setNewStatus(e.target.value)} style={{ width: "100%", background: "#141720", border: "1px solid #252b3b", borderRadius: 8, padding: "10px", color: "#e2e8f0", fontSize: 14 }}>
                  {WO_STATUSES.map(s => <option key={s}>{s}</option>)}
                </select>
              </div>
              <div style={{ marginBottom: 10 }}>
                <div style={{ fontSize: 11, color: "#64748b", marginBottom: 4, textTransform: "uppercase" }}>Status Note</div>
                <textarea value={note} onChange={e => setNote(e.target.value)} rows={2} placeholder="e.g. Awaiting parts from supplier..." style={{ width: "100%", background: "#141720", border: "1px solid #252b3b", borderRadius: 8, padding: "10px", color: "#e2e8f0", fontSize: 14, boxSizing: "border-box", resize: "vertical" }} />
              </div>
              <button onClick={() => { onUpdate(wo.id, newStatus, note); setShowUpdate(false); }} style={{ background: "#22c55e", color: "#fff", border: "none", borderRadius: 8, padding: "12px", width: "100%", fontSize: 14, fontWeight: 700, cursor: "pointer", marginBottom: 8 }}>Save Update</button>
              <button onClick={() => setShowUpdate(false)} style={{ background: "transparent", color: "#64748b", border: "1px solid #252b3b", borderRadius: 8, padding: "10px", width: "100%", fontSize: 13, cursor: "pointer" }}>Cancel</button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── HISTORY CARD ─────────────────────────────────────────────────────────────
function HistoryCard({ log, isSupervisor, onApprove, onReject }) {
  const [expanded, setExpanded] = useState(false);
  const [showReject, setShowReject] = useState(false);
  const [rejectNote, setRejectNote] = useState("");

  const approvalColor = log.approval_status === "Approved" ? "#22c55e" : log.approval_status === "Rejected" ? "#ef4444" : "#eab308";

  return (
    <div style={{ background: "#1a1e2a", border: `1px solid ${approvalColor}33`, borderRadius: 10, overflow: "hidden" }}>
      <div onClick={() => setExpanded(!expanded)} style={{ padding: 14, cursor: "pointer" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
          <div style={{ fontSize: 14, fontWeight: 700, color: "#e2e8f0", flex: 1, marginRight: 8 }}>{log.title}</div>
          <span style={{ background: approvalColor+"22", color: approvalColor, border: `1px solid ${approvalColor}44`, borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 700, whiteSpace: "nowrap" }}>{log.approval_status||"—"}</span>
        </div>
        <div style={{ fontSize: 12, color: "#64748b" }}>{log.log_type} · {log.start_date} · {log.performed_by}</div>
        {log.cost > 0 && <div style={{ fontSize: 12, color: "#f97316", marginTop: 4 }}>Cost: ${log.cost}</div>}
      </div>
      {expanded && (
        <div style={{ padding: "0 14px 14px", borderTop: "1px solid #252b3b" }}>
          {log.description && <div style={{ marginTop: 10, fontSize: 13, color: "#94a3b8", lineHeight: 1.6 }}>{log.description}</div>}
          {log.approval_status === "Approved" && log.approved_by && (
            <div style={{ marginTop: 10, background: "#22c55e11", border: "1px solid #22c55e33", borderRadius: 8, padding: "10px", fontSize: 12, color: "#22c55e" }}>
              ✅ Approved by {log.approved_by}
            </div>
          )}
          {log.approval_status === "Rejected" && log.rejection_notes && (
            <div style={{ marginTop: 10, background: "#ef444411", border: "1px solid #ef444433", borderRadius: 8, padding: "10px", fontSize: 12, color: "#ef4444" }}>
              ❌ Rejected: {log.rejection_notes}
            </div>
          )}
          {/* Supervisor approval actions */}
          {isSupervisor && log.approval_status === "Pending" && (
            <div style={{ marginTop: 12 }}>
              {!showReject ? (
                <div style={{ display: "flex", gap: 8 }}>
                  <button onClick={() => onApprove(log.id)} style={{ flex: 1, background: "#22c55e", color: "#fff", border: "none", borderRadius: 8, padding: "12px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>✅ Approve</button>
                  <button onClick={() => setShowReject(true)} style={{ flex: 1, background: "#ef444422", color: "#ef4444", border: "1px solid #ef444444", borderRadius: 8, padding: "12px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>❌ Reject</button>
                </div>
              ) : (
                <div>
                  <textarea value={rejectNote} onChange={e => setRejectNote(e.target.value)} rows={2} placeholder="Reason for rejection..." style={{ width: "100%", background: "#141720", border: "1px solid #ef444444", borderRadius: 8, padding: "10px", color: "#e2e8f0", fontSize: 14, boxSizing: "border-box", marginBottom: 8 }} />
                  <button onClick={() => { if (rejectNote) { onReject(log.id, rejectNote); setShowReject(false); } }} style={{ width: "100%", background: "#ef4444", color: "#fff", border: "none", borderRadius: 8, padding: "12px", fontSize: 14, fontWeight: 700, cursor: "pointer", marginBottom: 8 }}>Confirm Reject</button>
                  <button onClick={() => setShowReject(false)} style={{ width: "100%", background: "transparent", color: "#64748b", border: "1px solid #252b3b", borderRadius: 8, padding: "10px", fontSize: 13, cursor: "pointer" }}>Cancel</button>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
