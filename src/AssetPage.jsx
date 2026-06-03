import { useState, useEffect } from "react";
import { supabase } from "./supabase";

const C = {
  bg: "#0d0f12", surface: "#141720", card: "#1a1e2a", border: "#252b3b",
  accent: "#f97316", yellow: "#eab308", green: "#22c55e", red: "#ef4444",
  blue: "#3b82f6", text: "#e2e8f0", muted: "#64748b", subtle: "#94a3b8",
};

const statusColor = (s) => ({ Operational: C.green, "Under Maintenance": C.accent, Degraded: C.red }[s] || C.muted);
const fmtDate = (d) => d ? new Date(d).toLocaleDateString("en-GB") : "—";
const uid = (p) => `${p}-${Date.now().toString(36).toUpperCase()}${Math.random().toString(36).slice(2,5).toUpperCase()}`;
const TODAY = new Date().toISOString().split("T")[0];

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

const Btn = ({ children, onClick, color, disabled, secondary }) => (
  <button onClick={onClick} disabled={disabled} style={{
    width: "100%", background: secondary ? "transparent" : (color || C.accent),
    color: secondary ? C.muted : "#fff",
    border: secondary ? `1px solid ${C.border}` : "none",
    borderRadius: 10, padding: "14px", fontSize: 15, fontWeight: 700,
    cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.6 : 1,
    marginBottom: 10,
  }}>{children}</button>
);

const Banner = ({ msg, color, onDismiss }) => msg ? (
  <div style={{ background: color+"22", border: `1px solid ${color}44`, borderRadius: 8, padding: "12px 16px", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 14, color }}>
    {msg} <button onClick={onDismiss} style={{ background: "none", border: "none", color, cursor: "pointer", fontSize: 18 }}>×</button>
  </div>
) : null;

export default function AssetPage() {
  const [asset, setAsset] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [userRole, setUserRole] = useState(null);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [signing, setSigning] = useState(false);
  const [view, setView] = useState("home"); // home, breakdown, log, workorder, parts
  const [saving, setSaving] = useState(false);
  const [vendors, setVendors] = useState([]);
  const [logs, setLogs] = useState([]);
  const [loadingLogs, setLoadingLogs] = useState(false);

  // Get asset ID from URL
  const assetId = new URLSearchParams(window.location.search).get("asset");

  useEffect(() => {
    if (!assetId) { setError("No asset ID in URL."); setLoading(false); return; }
    loadAsset();
  }, [assetId]);

  useEffect(() => {
    supabase.auth.getSession().then(async ({ data: { session } }) => {
      if (session) {
        const { data } = await supabase.from("user_roles").select("*").eq("email", session.user.email).single();
        setUserRole(data || { role: "operations", name: session.user.email });
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
    const { data } = await supabase.from("maintenance_logs").select("*").eq("asset_id", assetId).order("start_date", { ascending: false }).limit(5);
    setLogs(data || []);
    setLoadingLogs(false);
  };

  const signIn = async () => {
    setSigning(true);
    const { error: err } = await supabase.auth.signInWithPassword({ email, password });
    if (err) { setError(err.message); setSigning(false); return; }
    const { data } = await supabase.from("user_roles").select("*").eq("email", email).single();
    setUserRole(data || { role: "operations", name: email });
    loadVendors();
    setSigning(false);
  };

  // ─── Breakdown Form ───────────────────────────────────────────────────────
  const [brkForm, setBrkForm] = useState({ reported_by: "", severity: "High", description: "" });
  const bf = (k) => (v) => setBrkForm(p => ({ ...p, [k]: v }));

  const submitBreakdown = async () => {
    if (!brkForm.description || !brkForm.reported_by) { setError("Please fill all fields."); return; }
    setSaving(true);
    const now = new Date().toISOString();
    const record = {
      id: uid("BRK"), asset_id: asset.id, asset_name: asset.name, site: asset.location,
      reported_by: brkForm.reported_by, reported_at: now, downtime_start: now,
      description: brkForm.description, severity: brkForm.severity, status: "Open",
    };
    const { error: err } = await supabase.from("breakdown_reports").insert([record]);
    if (err) { setError(err.message); } else {
      await supabase.from("assets").update({ status: "Under Maintenance" }).eq("id", asset.id);
      setAsset(prev => ({ ...prev, status: "Under Maintenance" }));
      // Send email notification
      try {
        await fetch("https://evwsdzqgvrwbjusjmrdc.supabase.co/functions/v1/notify-breakdown", {
          method: "POST",
          headers: { "Content-Type": "application/json", "Authorization": `Bearer ${import.meta.env.VITE_SUPABASE_ANON_KEY}` },
          body: JSON.stringify({ breakdown: record, type: "reported" }),
        });
      } catch (e) { console.error(e); }
      setSuccess("Breakdown reported! Maintenance team notified.");
      setView("home");
      setBrkForm({ reported_by: "", severity: "High", description: "" });
    }
    setSaving(false);
  };

  // ─── Maintenance Log Form ─────────────────────────────────────────────────
  const [logForm, setLogForm] = useState({ log_type: "Preventive Maintenance", title: "", description: "", performed_by: userRole?.name || "", cost: "", status: "Completed" });
  const lf = (k) => (v) => setLogForm(p => ({ ...p, [k]: v }));

  const submitLog = async () => {
    if (!logForm.title) { setError("Title is required."); return; }
    setSaving(true);
    const record = {
      id: uid("LOG"), asset_id: asset.id, asset_name: asset.name,
      log_type: logForm.log_type, title: logForm.title, description: logForm.description,
      performed_by: logForm.performed_by || userRole?.name,
      start_date: TODAY, end_date: TODAY,
      cost: logForm.cost ? parseFloat(logForm.cost) : null,
      status: logForm.status,
    };
    const { error: err } = await supabase.from("maintenance_logs").insert([record]);
    if (err) { setError(err.message); } else {
      setSuccess("Maintenance log added!");
      setView("home");
      setLogForm({ log_type: "Preventive Maintenance", title: "", description: "", performed_by: "", cost: "", status: "Completed" });
    }
    setSaving(false);
  };

  // ─── Work Order Form ──────────────────────────────────────────────────────
  const [woForm, setWoForm] = useState({ title: "", priority: "Medium", due: "", vendor: "" });
  const wf = (k) => (v) => setWoForm(p => ({ ...p, [k]: v }));
  const vendorOptions = ["— None —", ...vendors.map(v => v.name)];

  const submitWO = async () => {
    if (!woForm.title) { setError("Title is required."); return; }
    setSaving(true);
    const vendorName = woForm.vendor === "— None —" ? null : woForm.vendor || null;
    const record = {
      id: uid("WO"), title: woForm.title, asset: asset.name,
      priority: woForm.priority, status: "Open",
      start_date: TODAY, due: woForm.due || null, vendor: vendorName,
    };
    const { error: err } = await supabase.from("work_orders").insert([record]);
    if (err) { setError(err.message); } else {
      setSuccess("Work order created!");
      setView("home");
      setWoForm({ title: "", priority: "Medium", due: "", vendor: "" });
    }
    setSaving(false);
  };

  // ─── Spare Parts Form ─────────────────────────────────────────────────────
  const [partForm, setPartForm] = useState({ log_id: "", part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "" });
  const pf = (k) => (v) => setPartForm(p => ({ ...p, [k]: v }));

  const submitPart = async () => {
    if (!partForm.part_name || !partForm.log_id) { setError("Part name and log are required."); return; }
    setSaving(true);
    const qty = parseFloat(partForm.quantity) || 1;
    const unitCost = parseFloat(partForm.unit_cost) || 0;
    const record = {
      id: uid("PRT"), log_id: partForm.log_id, asset_id: asset.id,
      part_name: partForm.part_name, part_number: partForm.part_number,
      quantity: qty, unit_cost: unitCost, total_cost: qty * unitCost,
      supplier: partForm.supplier,
    };
    const { error: err } = await supabase.from("spare_parts").insert([record]);
    if (err) { setError(err.message); } else {
      setSuccess("Spare part added!");
      setView("home");
      setPartForm({ log_id: "", part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "" });
    }
    setSaving(false);
  };

  if (loading) return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", color: C.muted, fontFamily: "Arial, sans-serif" }}>
      Loading asset...
    </div>
  );

  if (error && !asset) return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", color: C.red, fontFamily: "Arial, sans-serif", padding: 20 }}>
      {error}
    </div>
  );

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "Arial, sans-serif", color: C.text, maxWidth: 500, margin: "0 auto", padding: 16 }}>

      {/* Header */}
      <div style={{ background: C.accent, borderRadius: 12, padding: 20, marginBottom: 20, textAlign: "center" }}>
        <div style={{ fontSize: 32, marginBottom: 8 }}>🏭</div>
        <div style={{ fontSize: 18, fontWeight: 700, color: "#fff" }}>{asset?.name}</div>
        <div style={{ fontSize: 13, color: "rgba(255,255,255,0.8)", marginTop: 4 }}>{asset?.category} · {asset?.location}</div>
        <div style={{ marginTop: 10, display: "inline-block", background: statusColor(asset?.status)+"33", color: statusColor(asset?.status), border: `1px solid ${statusColor(asset?.status)}66`, borderRadius: 6, padding: "4px 12px", fontSize: 12, fontWeight: 700, textTransform: "uppercase" }}>
          {asset?.status}
        </div>
      </div>

      <Banner msg={error} color={C.red} onDismiss={() => setError(null)} />
      <Banner msg={success} color={C.green} onDismiss={() => setSuccess(null)} />

      {/* Login if not authenticated */}
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
          {/* Asset Info */}
          <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 20, marginBottom: 16 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: C.muted, textTransform: "uppercase", marginBottom: 12 }}>Asset Info</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              {[["ID", asset?.id], ["Site", asset?.location], ["PM Every", asset?.pm_frequency ? `${asset.pm_frequency} mo.` : "—"], ["Last PM", asset?.last_pm_date ? fmtDate(asset.last_pm_date) : "Never"], ["Category", asset?.category], ["Value", asset?.value || "—"]].map(([lbl, val]) => (
                <div key={lbl}>
                  <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase" }}>{lbl}</div>
                  <div style={{ fontSize: 13, color: C.subtle, marginTop: 2 }}>{val}</div>
                </div>
              ))}
            </div>
          </div>

          {/* Role badge */}
          <div style={{ textAlign: "center", marginBottom: 16, fontSize: 13, color: C.muted }}>
            Signed in as <strong style={{ color: C.accent }}>{userRole.name}</strong> ({userRole.role})
          </div>

          {/* Action Buttons */}
          <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
            {/* Operations: Report Breakdown */}
            <Btn onClick={() => setView("breakdown")} color={C.red}>🚨 Report Breakdown</Btn>

            {/* Maintenance: Additional actions */}
            {(userRole.role === "maintenance" || userRole.role === "admin") && (
              <>
                <Btn onClick={() => { setView("log"); loadLogs(); }} color={C.blue}>🔧 Add Maintenance Log</Btn>
                <Btn onClick={() => setView("workorder")} color={C.accent}>📋 Open Work Order</Btn>
                <Btn onClick={() => { setView("parts"); loadLogs(); }} color={C.purple || "#a855f7"}>🔩 Add Spare Parts</Btn>
              </>
            )}
          </div>
        </>
      ) : view === "breakdown" ? (
        <div style={{ background: C.card, border: `1px solid ${C.red}44`, borderRadius: 12, padding: 20 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.red, marginBottom: 16 }}>🚨 Report Breakdown</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Input label="Your Name *" value={brkForm.reported_by} onChange={bf("reported_by")} placeholder="Full name" />
            <Sel label="Severity" value={brkForm.severity} onChange={bf("severity")} options={["Critical", "High", "Medium", "Low"]} />
            <Textarea label="Describe the issue *" value={brkForm.description} onChange={bf("description")} placeholder="What happened? Any error messages?" />
            <div style={{ background: C.yellow+"22", border: `1px solid ${C.yellow}44`, borderRadius: 8, padding: 12, fontSize: 13, color: C.yellow }}>
              ⏱ Downtime starts now: {new Date().toLocaleString("en-GB")}
            </div>
          </div>
          <div style={{ marginTop: 16 }}>
            <Btn onClick={submitBreakdown} disabled={saving} color={C.red}>{saving ? "Reporting..." : "🚨 Report Breakdown"}</Btn>
            <Btn onClick={() => setView("home")} secondary>Cancel</Btn>
          </div>
        </div>
      ) : view === "log" ? (
        <div style={{ background: C.card, border: `1px solid ${C.blue}44`, borderRadius: 12, padding: 20 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.blue, marginBottom: 16 }}>🔧 Add Maintenance Log</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Sel label="Type" value={logForm.log_type} onChange={lf("log_type")} options={["Preventive Maintenance", "Corrective Repair", "Inspection", "Overhaul", "Part Replacement"]} />
            <Input label="Title *" value={logForm.title} onChange={lf("title")} placeholder="What was done?" />
            <Input label="Performed By" value={logForm.performed_by} onChange={lf("performed_by")} placeholder="Technician name" />
            <Input label="Cost ($)" value={logForm.cost} onChange={lf("cost")} type="number" />
            <Textarea label="Notes" value={logForm.description} onChange={lf("description")} placeholder="Details about the work done..." />
            <Sel label="Status" value={logForm.status} onChange={lf("status")} options={["Completed", "In Progress", "Pending"]} />
          </div>
          <div style={{ marginTop: 16 }}>
            <Btn onClick={submitLog} disabled={saving} color={C.blue}>{saving ? "Saving..." : "Save Log"}</Btn>
            <Btn onClick={() => setView("home")} secondary>Cancel</Btn>
          </div>
        </div>
      ) : view === "workorder" ? (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 12, padding: 20 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.accent, marginBottom: 16 }}>📋 Open Work Order</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Input label="Title *" value={woForm.title} onChange={wf("title")} placeholder="Describe the work needed" />
            <Sel label="Priority" value={woForm.priority} onChange={wf("priority")} options={["Critical", "High", "Medium", "Low"]} />
            <Input label="Due Date" value={woForm.due} onChange={wf("due")} type="date" />
            <Sel label="Vendor" value={woForm.vendor} onChange={wf("vendor")} options={vendorOptions} />
          </div>
          <div style={{ marginTop: 16 }}>
            <Btn onClick={submitWO} disabled={saving} color={C.accent}>{saving ? "Creating..." : "Create Work Order"}</Btn>
            <Btn onClick={() => setView("home")} secondary>Cancel</Btn>
          </div>
        </div>
      ) : view === "parts" ? (
        <div style={{ background: C.card, border: `1px solid #a855f744`, borderRadius: 12, padding: 20 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: "#a855f7", marginBottom: 16 }}>🔩 Add Spare Parts</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            {/* Select Log */}
            <div>
              <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase" }}>Link to Maintenance Log *</div>
              <select value={partForm.log_id} onChange={e => setPartForm(p => ({ ...p, log_id: e.target.value }))}
                style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px", color: C.text, fontSize: 15 }}>
                <option value="">— Select Log —</option>
                {loadingLogs ? <option>Loading...</option> : logs.map(l => <option key={l.id} value={l.id}>{l.title} ({l.start_date})</option>)}
              </select>
            </div>
            <Input label="Part Name *" value={partForm.part_name} onChange={pf("part_name")} placeholder="e.g. Hydraulic filter" />
            <Input label="Part Number" value={partForm.part_number} onChange={pf("part_number")} placeholder="e.g. HF-2023" />
            <Input label="Quantity" value={partForm.quantity} onChange={pf("quantity")} type="number" />
            <Input label="Unit Cost ($)" value={partForm.unit_cost} onChange={pf("unit_cost")} type="number" />
            <Input label="Supplier" value={partForm.supplier} onChange={pf("supplier")} placeholder="Supplier name" />
          </div>
          <div style={{ marginTop: 16 }}>
            <Btn onClick={submitPart} disabled={saving} color="#a855f7">{saving ? "Saving..." : "Add Spare Part"}</Btn>
            <Btn onClick={() => setView("home")} secondary>Cancel</Btn>
          </div>
        </div>
      ) : null}

      {/* Footer */}
      <div style={{ textAlign: "center", marginTop: 24, fontSize: 11, color: C.muted }}>
        Facility Command · EPx Logistics
      </div>
    </div>
  );
}
