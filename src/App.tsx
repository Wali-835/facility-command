// v2 - vendor work orders
import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase";

const C = {
  bg: "#0d0f12", surface: "#141720", card: "#1a1e2a", border: "#252b3b",
  accent: "#f97316", yellow: "#eab308", green: "#22c55e", red: "#ef4444",
  blue: "#3b82f6", purple: "#a855f7", text: "#e2e8f0", muted: "#64748b", subtle: "#94a3b8",
};

const SITES = [
  "— Select Site —", "Site1", "Site2", "Site3", "Site4", "Site5", "Site6",
  "Site7B", "Site7C", "Site8", "Site9", "Site9A", "Site9B",
  "Site10", "Site10A", "Site10B", "Site11", "Site12", "Site14", "Site14A", "Site14B", "Storage",
];

const LOG_TYPES = ["Preventive Maintenance", "Corrective Repair", "Inspection", "Overhaul", "Part Replacement"];
const LOG_STATUSES = ["Completed", "In Progress", "Pending", "Cancelled"];
const FREQ_COLORS = { D: C.green, W: C.blue, F: C.purple, M: C.accent };
const FREQ_LABELS = { D: "Daily", W: "Weekly", F: "Bi-weekly", M: "Monthly" };
const CAT_ICONS = { Cleaning: "🧹", Inspection: "🔍", Lubrication: "🔧", Safety: "⚠️" };
const SEVERITY_COLORS = { Critical: C.red, High: C.accent, Medium: C.yellow, Low: C.green };

const ADMIN_EMAIL = import.meta.env.VITE_ADMIN_EMAIL;
const TODAY = new Date().toISOString().split("T")[0];
const uid = (p) => `${p}-${Date.now().toString(36).toUpperCase()}${Math.random().toString(36).slice(2,5).toUpperCase()}`;
const priorityColor = (p) => ({ Critical: C.red, High: C.accent, Medium: C.yellow, Low: C.green }[p] || C.muted);
const statusColor = (s) => ({
  Open: C.accent, "In Progress": C.blue, Completed: C.green, Pending: C.yellow,
  Operational: C.green, "Under Maintenance": C.accent, Degraded: C.red,
  Active: C.green, Inactive: C.muted, Cancelled: C.muted, Acknowledged: C.blue, Resolved: C.green,
  "Preventive Maintenance": C.blue, "Corrective Repair": C.red,
  Inspection: C.yellow, Overhaul: C.purple, "Part Replacement": C.accent,
}[s] || C.muted);

const fmt = (n) => n ? `$${Number(n).toLocaleString()}` : "—";
const fmtDate = (d) => d ? new Date(d).toLocaleDateString("en-GB") : "—";
const fmtDateTime = (d) => {
  if (!d) return "—";
  const date = new Date(d.endsWith('Z') ? d : d + 'Z');
  return date.toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit", timeZone: "Africa/Cairo" });
};
const hoursBetween = (a, b) => a && b ? Math.round((new Date(b) - new Date(a)) / (1000 * 60 * 60)) : null;
const minutesBetween = (a, b) => {
  if (!a || !b) return null;
  const dateA = new Date(a.endsWith('Z') ? a : a + 'Z');
  const dateB = new Date(b.endsWith('Z') ? b : b + 'Z');
  return Math.round((dateB - dateA) / (1000 * 60));
};
const formatDowntime = (minutes) => {
  if (minutes === null || minutes === undefined) return "—";
  if (minutes === 0) return "< 1 min";
  if (minutes < 60) return `${minutes} min`;
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return m > 0 ? `${h}h ${m}min` : `${h}h`;
};

// ─── Shared UI ────────────────────────────────────────────────────────────────
const Badge = ({ label, color }) => (
  <span style={{ background: color+"22", color, border: `1px solid ${color}44`, borderRadius: 4, padding: "2px 8px", fontSize: 11, fontWeight: 700, letterSpacing: "0.05em", textTransform: "uppercase", whiteSpace: "nowrap" }}>{label}</span>
);

const Input = ({ label, value, onChange, type="text", placeholder }) => (
  <div>
    <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
    <input type={type} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder}
      style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14, boxSizing: "border-box" }} />
  </div>
);

const Textarea = ({ label, value, onChange, placeholder }) => (
  <div>
    <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
    <textarea value={value} onChange={e => onChange(e.target.value)} rows={3} placeholder={placeholder}
      style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14, boxSizing: "border-box", resize: "vertical" }} />
  </div>
);

const Sel = ({ label, value, onChange, options }) => (
  <div>
    <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
    <select value={value} onChange={e => onChange(e.target.value)}
      style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14 }}>
      {options.map(o => <option key={o}>{o}</option>)}
    </select>
  </div>
);

const Btn = ({ children, onClick, variant, disabled, color, small }) => {
  const isPrimary = variant !== "secondary" && variant !== "danger";
  const isDanger = variant === "danger";
  return (
    <button onClick={onClick} disabled={disabled} style={{
      background: isDanger ? C.red+"22" : isPrimary ? (color||C.accent) : "transparent",
      color: isDanger ? C.red : isPrimary ? "#fff" : C.muted,
      border: isDanger ? `1px solid ${C.red}44` : isPrimary ? "none" : `1px solid ${C.border}`,
      borderRadius: 6, padding: small ? "4px 10px" : "7px 14px", fontSize: small ? 12 : 13, fontWeight: 700,
      cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.5 : 1,
    }}>{children}</button>
  );
};

const StatCard = ({ icon, label, value, sub, color }) => (
  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: "16px 20px", flex: "1 1 140px", borderLeft: `3px solid ${color||C.accent}` }}>
    <div style={{ fontSize: 20, marginBottom: 6 }}>{icon}</div>
    <div style={{ fontSize: 26, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{value}</div>
    <div style={{ fontSize: 13, color: C.subtle, marginTop: 2 }}>{label}</div>
    {sub && <div style={{ fontSize: 11, color: color||C.accent, marginTop: 4 }}>{sub}</div>}
  </div>
);

const Spinner = () => <div style={{ textAlign: "center", padding: 48, color: C.muted, fontSize: 13 }}>Loading...</div>;

const ErrBanner = ({ msg, onDismiss }) => msg ? (
  <div style={{ background: C.red+"22", border: `1px solid ${C.red}44`, borderRadius: 8, padding: "10px 16px", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, color: C.red }}>
    {msg} <button onClick={onDismiss} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 16 }}>x</button>
  </div>
) : null;

const OkBanner = ({ msg, onDismiss }) => msg ? (
  <div style={{ background: C.green+"22", border: `1px solid ${C.green}44`, borderRadius: 8, padding: "10px 16px", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, color: C.green }}>
    {msg} <button onClick={onDismiss} style={{ background: "none", border: "none", color: C.green, cursor: "pointer", fontSize: 16 }}>x</button>
  </div>
) : null;

const StatusSel = ({ value, options, onChange }) => (
  <select value={value} onChange={e => onChange(e.target.value)} style={{
    background: statusColor(value)+"22", color: statusColor(value), border: `1px solid ${statusColor(value)}44`,
    borderRadius: 4, padding: "3px 6px", fontSize: 11, fontWeight: 700, cursor: "pointer", textTransform: "uppercase", letterSpacing: "0.05em",
  }}>
    {options.map(o => <option key={o} value={o}>{o}</option>)}
  </select>
);

const EditModal = ({ title, fields, data, onSave, onClose }) => {
  const [form, setForm] = useState({ ...data });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000aa", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 24, width: "100%", maxWidth: 560, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: C.accent, marginBottom: 20, textTransform: "uppercase" }}>Edit {title}</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
          {fields.map(({ key, label, type, options }) =>
            options ? <Sel key={key} label={label} value={form[key]||""} onChange={f(key)} options={options} />
              : <Input key={key} label={label} value={form[key]||""} onChange={f(key)} type={type||"text"} />
          )}
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
          <Btn onClick={() => onSave(form)}>Save Changes</Btn>
          <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        </div>
      </div>
    </div>
  );
};

const ConfirmDel = ({ name, onConfirm, onClose }) => (
  <div style={{ position: "fixed", inset: 0, background: "#000000aa", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}>
    <div style={{ background: C.card, border: `1px solid ${C.red}44`, borderRadius: 12, padding: 24, width: "100%", maxWidth: 400 }}>
      <div style={{ fontSize: 15, fontWeight: 700, color: C.red, marginBottom: 12 }}>Confirm Delete</div>
      <div style={{ fontSize: 13, color: C.subtle, marginBottom: 20 }}>Delete <strong style={{ color: C.text }}>{name}</strong>? This cannot be undone.</div>
      <div style={{ display: "flex", gap: 8 }}><Btn variant="danger" onClick={onConfirm}>Delete</Btn><Btn variant="secondary" onClick={onClose}>Cancel</Btn></div>
    </div>
  </div>
);

// ─── BREAKDOWN REPORT MODAL (Operations User) ─────────────────────────────────
function BreakdownReportModal({ asset, userRole, onClose, onReported }) {
  const [form, setForm] = useState({ description: "", severity: "High", reported_by: userRole.name || "" });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));

  const submit = async () => {
    if (!form.description) { setError("Please describe the breakdown."); return; }
    if (!form.reported_by) { setError("Please enter your name."); return; }
    setSaving(true); setError(null);
    const now = new Date().toISOString();
    const record = {
      id: uid("BRK"), asset_id: asset.id, asset_name: asset.name, site: asset.location,
      reported_by: form.reported_by, reported_at: now, downtime_start: now,
      description: form.description, severity: form.severity, status: "Open",
    };
    const { error: err } = await supabase.from("breakdown_reports").insert([record]);
    if (err) { setError(err.message); setSaving(false); return; }
    // Update asset status to Under Maintenance
    await supabase.from("assets").update({ status: "Under Maintenance" }).eq("id", asset.id);
    onReported(record);
    onClose();
    setSaving(false);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2000, padding: 16 }}>
      <div style={{ background: C.card, border: `2px solid ${C.red}44`, borderRadius: 12, width: "100%", maxWidth: 500 }}>
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, background: C.red+"11" }}>
          <div style={{ fontSize: 17, fontWeight: 700, color: C.red }}>🚨 Report Breakdown</div>
          <div style={{ fontSize: 13, color: C.muted, marginTop: 4 }}>{asset.name} · {asset.location}</div>
        </div>
        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Input label="Reported By *" value={form.reported_by} onChange={f("reported_by")} placeholder="Your full name" />
            <Sel label="Severity" value={form.severity} onChange={f("severity")} options={["Critical", "High", "Medium", "Low"]} />
            <Textarea label="Breakdown Description *" value={form.description} onChange={f("description")} placeholder="Describe what happened, any error messages, sounds, etc..." />
            <div style={{ background: C.yellow+"11", border: `1px solid ${C.yellow}44`, borderRadius: 8, padding: 12, fontSize: 12, color: C.yellow }}>
              ⏱ Downtime will start from: <strong>{new Date().toLocaleString("en-GB")}</strong>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
            <Btn onClick={submit} disabled={saving} color={C.red}>{saving ? "Reporting..." : "🚨 Report Breakdown"}</Btn>
            <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── BREAKDOWN RESOLUTION MODAL (Maintenance User) ────────────────────────────
function BreakdownResolveModal({ breakdown, userRole, vendors, onClose, onResolved }) {
  const [form, setForm] = useState({ resolved_by: userRole.name || "", maintenance_notes: "", vendor: "" });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const vendorOptions = ["— None —", ...vendors.filter(v => v.status === "Active").map(v => v.name)];

  const downtimeMins = minutesBetween(breakdown.downtime_start, new Date().toISOString());

  const submit = async () => {
    if (!form.resolved_by) { setError("Please enter your name."); return; }
    if (!form.maintenance_notes) { setError("Please add maintenance notes."); return; }
    setSaving(true); setError(null);
    const now = new Date().toISOString();
    const hours = minutesBetween(breakdown.downtime_start, now);

    // Update breakdown report
    await supabase.from("breakdown_reports").update({
      status: "Resolved", resolved_by: form.resolved_by, resolved_at: now,
      downtime_end: now, downtime_hours: hours, maintenance_notes: form.maintenance_notes,
    }).eq("id", breakdown.id);

    // Update asset status back to Operational
    await supabase.from("assets").update({ status: "Operational" }).eq("id", breakdown.asset_id);

    // Auto-create maintenance log
    const logRecord = {
      id: uid("LOG"), asset_id: breakdown.asset_id, asset_name: breakdown.asset_name,
      log_type: "Corrective Repair",
      title: `Breakdown Repair — ${breakdown.severity} severity`,
      description: `BREAKDOWN REPORTED BY: ${breakdown.reported_by}\n\nISSUE: ${breakdown.description}\n\nMAINTENANCE NOTES: ${form.maintenance_notes}`,
      performed_by: form.resolved_by,
      vendor: form.vendor === "— None —" ? null : form.vendor || null,
      start_date: breakdown.downtime_start ? breakdown.downtime_start.split("T")[0] : TODAY,
      end_date: TODAY, cost: null, status: "Completed",
      downtime_start: breakdown.downtime_start ? breakdown.downtime_start.split("T")[0] : null,
      downtime_end: TODAY,
      downtime_hours: hours,
    };
    await supabase.from("maintenance_logs").insert([logRecord]);

    // Close any open work orders for this asset
    const { data: openWOs } = await supabase.from("work_orders").select("id").eq("asset", breakdown.asset_name).in("status", ["Open", "In Progress", "Pending"]);
    if (openWOs?.length) await supabase.from("work_orders").update({ status: "Completed" }).in("id", openWOs.map(w => w.id));

    onResolved({ ...breakdown, status: "Resolved", downtime_end: now, downtime_hours: hours });
    onClose();
    setSaving(false);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2000, padding: 16 }}>
      <div style={{ background: C.card, border: `2px solid ${C.green}44`, borderRadius: 12, width: "100%", maxWidth: 560, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, background: C.green+"11" }}>
          <div style={{ fontSize: 17, fontWeight: 700, color: C.green }}>✅ Resolve Breakdown</div>
          <div style={{ fontSize: 13, color: C.muted, marginTop: 4 }}>{breakdown.asset_name} · {breakdown.site}</div>
        </div>
        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          {/* Breakdown info */}
          <div style={{ background: C.red+"11", border: `1px solid ${C.red}33`, borderRadius: 8, padding: 14, marginBottom: 18 }}>
            <div style={{ fontSize: 12, color: C.muted, marginBottom: 8, textTransform: "uppercase", fontWeight: 600 }}>Breakdown Details</div>
            <div style={{ fontSize: 13, color: C.text, marginBottom: 6 }}><strong>Reported by:</strong> {breakdown.reported_by}</div>
            <div style={{ fontSize: 13, color: C.text, marginBottom: 6 }}><strong>Reported at:</strong> {fmtDateTime(breakdown.reported_at)}</div>
            <div style={{ fontSize: 13, color: C.text, marginBottom: 6 }}><strong>Issue:</strong> {breakdown.description}</div>
            <div style={{ fontSize: 13, color: C.yellow, fontWeight: 700 }}>⏱ Current downtime: {formatDowntime(downtimeMins)}</div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Input label="Resolved By (Technician) *" value={form.resolved_by} onChange={f("resolved_by")} placeholder="Your full name" />
            <Sel label="Vendor / Contractor Used" value={form.vendor} onChange={f("vendor")} options={vendorOptions} />
            <Textarea label="Maintenance Notes *" value={form.maintenance_notes} onChange={f("maintenance_notes")} placeholder="What was the root cause? What was done to fix it? Parts replaced?" />
            <div style={{ background: C.green+"11", border: `1px solid ${C.green}44`, borderRadius: 8, padding: 12, fontSize: 12, color: C.green }}>
              ✅ Equipment will be marked as <strong>Operational</strong> and downtime end will be set to: <strong>{new Date().toLocaleString("en-GB")}</strong>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
            <Btn onClick={submit} disabled={saving} color={C.green}>{saving ? "Resolving..." : "✅ Mark as Resolved"}</Btn>
            <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── BREAKDOWNS TAB ───────────────────────────────────────────────────────────
function Breakdowns({ userRole, assets, setAssets, vendors, workOrders, setWorkOrders }) {
  const [breakdowns, setBreakdowns] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [showReportForm, setShowReportForm] = useState(false);
  const [resolveItem, setResolveItem] = useState(null);
  const [filter, setFilter] = useState("Open");
  const [selectedAsset, setSelectedAsset] = useState(null);

  useEffect(() => { loadBreakdowns(); }, []);

  const loadBreakdowns = async () => {
    setLoading(true);
    const { data } = await supabase.from("breakdown_reports").select("*").order("reported_at", { ascending: false });
    setBreakdowns(data || []);
    setLoading(false);
  };

  const filtered = filter === "All" ? breakdowns : breakdowns.filter(b => b.status === filter);
  const openCount = breakdowns.filter(b => b.status === "Open").length;
  const acknowledgedCount = breakdowns.filter(b => b.status === "Acknowledged").length;
  const resolvedCount = breakdowns.filter(b => b.status === "Resolved").length;
  const totalDowntimeMins = breakdowns.filter(b => b.downtime_hours).reduce((s, b) => s + (b.downtime_hours || 0), 0);

  const onReported = (record) => {
    setBreakdowns(prev => [record, ...prev]);
    setAssets(prev => prev.map(a => a.id === record.asset_id ? { ...a, status: "Under Maintenance" } : a));
    setSuccess("Breakdown reported! Maintenance team has been notified.");
  };

  const onResolved = (updated) => {
    setBreakdowns(prev => prev.map(b => b.id === updated.id ? updated : b));
    setAssets(prev => prev.map(a => a.id === updated.asset_id ? { ...a, status: "Operational" } : a));
    setSuccess("Breakdown resolved! Equipment is back to operation. Maintenance log created automatically.");
  };

  const acknowledge = async (b) => {
    const now = new Date().toISOString();
    const { error: err } = await supabase.from("breakdown_reports").update({
      status: "Acknowledged",
      acknowledged_by: userRole.name || userRole.email || "",
      acknowledged_at: now,
    }).eq("id", b.id);
    if (err) { setError(err.message); return; }
    setBreakdowns(prev => prev.map(x => x.id === b.id ? { ...x, status: "Acknowledged", acknowledged_by: userRole.name, acknowledged_at: now } : x));
    setSuccess("Breakdown acknowledged. You are now responsible for resolving it.");
  };

  return (
    <div>
      {showReportForm && selectedAsset && (
        <BreakdownReportModal asset={selectedAsset} userRole={userRole} onClose={() => { setShowReportForm(false); setSelectedAsset(null); }} onReported={onReported} />
      )}
      {resolveItem && (
        <BreakdownResolveModal breakdown={resolveItem} userRole={userRole} vendors={vendors} onClose={() => setResolveItem(null)} onResolved={onResolved} />
      )}

      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <OkBanner msg={success} onDismiss={() => setSuccess(null)} />

      {/* Stats */}
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        <StatCard icon="🚨" label="Open Breakdowns" value={openCount} sub="needs attention" color={C.red} />
        <StatCard icon="👁" label="Acknowledged" value={acknowledgedCount} sub="being handled" color={C.blue} />
        <StatCard icon="✅" label="Resolved" value={resolvedCount} sub="this period" color={C.green} />
        <StatCard icon="⏱" label="Total Downtime" value={formatDowntime(totalDowntimeMins)} sub="all breakdowns" color={C.yellow} />
        <StatCard icon="🏭" label="Assets Down" value={assets.filter(a => a.status === "Under Maintenance").length} sub="currently offline" color={C.accent} />
      </div>

      {/* Report Breakdown Button */}
      {(userRole.role === "operations" || userRole.role === "admin") && (
        <div style={{ background: C.card, border: `1px solid ${C.red}33`, borderRadius: 10, padding: 20, marginBottom: 20 }}>
          <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 8 }}>🚨 Report Equipment Breakdown</div>
          <div style={{ fontSize: 13, color: C.subtle, marginBottom: 14 }}>Select the equipment that broke down to report it to the maintenance team.</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
            <select onChange={e => setSelectedAsset(assets.find(a => a.id === e.target.value) || null)}
              style={{ flex: 1, minWidth: 200, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14 }}>
              <option value="">— Select Equipment —</option>
              {assets.filter(a => a.status !== "Under Maintenance").map(a => <option key={a.id} value={a.id}>{a.name} ({a.location})</option>)}
            </select>
            <Btn onClick={() => { if (!selectedAsset) { setError("Please select equipment first."); return; } setShowReportForm(true); }} color={C.red}>
              🚨 Report Breakdown
            </Btn>
          </div>
        </div>
      )}

      {/* Filter */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", gap: 6 }}>
          {["All", "Open", "Acknowledged", "Resolved"].map(s => (
            <button key={s} onClick={() => setFilter(s)} style={{ background: filter===s?C.accent:C.card, color: filter===s?"#fff":C.muted, border: `1px solid ${filter===s?C.accent:C.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>{s}</button>
          ))}
        </div>
        <button onClick={loadBreakdowns} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "6px 12px", color: C.muted, cursor: "pointer", fontSize: 12 }}>↻ Refresh</button>
      </div>

      {/* Breakdown List */}
      {loading ? <Spinner /> : filtered.length === 0 ? (
        <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>No breakdowns found.</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {filtered.map(b => {
            const isResolved = b.status === "Resolved";
            const isAcknowledged = b.status === "Acknowledged";
            const currentDowntimeMins = isResolved
  ? Math.round(b.downtime_hours || 0)
  : minutesBetween(b.downtime_start, new Date().toISOString()) || 0;
            return (
              <div key={b.id} style={{ background: C.card, border: `1px solid ${isResolved ? C.green+"44" : isAcknowledged ? C.blue+"44" : C.red+"44"}`, borderRadius: 10, padding: 20 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
                  <div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                      <span style={{ fontSize: 18 }}>{isResolved ? "✅" : isAcknowledged ? "👁" : "🚨"}</span>
                      <div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{b.asset_name}</div>
                    </div>
                    <div style={{ fontSize: 12, color: C.muted }}>{b.site} · Reported by {b.reported_by} · {fmtDateTime(b.reported_at)}</div>
                    {isAcknowledged && b.acknowledged_by && (
                      <div style={{ fontSize: 12, color: C.blue, marginTop: 4 }}>👁 Acknowledged by {b.acknowledged_by} · {fmtDateTime(b.acknowledged_at)}</div>
                    )}
                  </div>
                  <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                    <Badge label={b.severity} color={SEVERITY_COLORS[b.severity] || C.muted} />
                    <Badge label={b.status} color={statusColor(b.status)} />
                  </div>
                </div>

                {/* Downtime Timeline */}
                <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 14 }}>
                  <div style={{ background: C.red+"11", border: `1px solid ${C.red}33`, borderRadius: 8, padding: "10px 16px", flex: "1 1 160px" }}>
                    <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", marginBottom: 4 }}>⛔ Downtime Started</div>
                    <div style={{ fontSize: 13, color: C.red, fontWeight: 700 }}>{fmtDateTime(b.downtime_start)}</div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", color: C.muted, fontSize: 20 }}>→</div>
                  <div style={{ background: isResolved ? C.green+"11" : C.yellow+"11", border: `1px solid ${isResolved ? C.green+"33" : C.yellow+"33"}`, borderRadius: 8, padding: "10px 16px", flex: "1 1 160px" }}>
                    <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", marginBottom: 4 }}>{isResolved ? "✅ Back to Operation" : "⏳ Still Down"}</div>
                    <div style={{ fontSize: 13, color: isResolved ? C.green : C.yellow, fontWeight: 700 }}>{isResolved ? fmtDateTime(b.downtime_end) : "In Progress..."}</div>
                  </div>
                  <div style={{ background: C.purple+"11", border: `1px solid ${C.purple}33`, borderRadius: 8, padding: "10px 16px" }}>
                    <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", marginBottom: 4 }}>⏱ Downtime</div>
                    <div style={{ fontSize: 13, color: C.purple, fontWeight: 700 }}>{formatDowntime(currentDowntimeMins)}</div>
                  </div>
                </div>

                {/* Description */}
                <div style={{ background: C.surface, borderRadius: 8, padding: 12, marginBottom: 14, fontSize: 13, color: C.subtle }}>
                  <strong style={{ color: C.text }}>Issue:</strong> {b.description}
                </div>

                {/* Resolution notes */}
                {isResolved && b.maintenance_notes && (
                  <div style={{ background: C.green+"11", border: `1px solid ${C.green}33`, borderRadius: 8, padding: 12, marginBottom: 14, fontSize: 13, color: C.subtle }}>
                    <strong style={{ color: C.green }}>✅ Resolved by {b.resolved_by}:</strong> {b.maintenance_notes}
                  </div>
                )}

                {/* Actions for maintenance/admin */}
                {!isResolved && (userRole.role === "maintenance" || userRole.role === "admin") && (
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    {b.status === "Open" && (
                      <Btn onClick={() => acknowledge(b)} color={C.blue}>👁 Acknowledge</Btn>
                    )}
                    <Btn onClick={() => setResolveItem(b)} color={C.green}>✅ Mark as Resolved</Btn>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── CIL CHECKLIST MODAL ─────────────────────────────────────────────────────
function ChecklistModal({ asset, workOrderId, onClose }) {
  const [items, setItems] = useState([]);
  const [responses, setResponses] = useState({});
  const [executionId, setExecutionId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [executedBy, setExecutedBy] = useState("");
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [filterFreq, setFilterFreq] = useState("All");

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
      setExecutedBy(existing[0].executed_by || "");
      const { data: resps } = await supabase.from("checklist_responses").select("*").eq("execution_id", existing[0].id);
      const respMap = {};
      resps?.forEach(r => { respMap[r.item_id] = r; });
      setResponses(respMap);
    }
    setLoading(false);
  };

  const startExecution = async () => {
    if (!executedBy) { setError("Please enter your name first."); return; }
    const now = new Date();
    const execId = uid("EXC");
    const { data: checklists } = await supabase.from("checklists").select("id").limit(1);
    await supabase.from("checklist_executions").insert([{ id: execId, checklist_id: checklists[0].id, asset_id: asset.id, asset_name: asset.name, work_order_id: workOrderId || null, executed_by: executedBy, execution_date: TODAY, month: now.getMonth()+1, year: now.getFullYear(), status: "In Progress" }]);
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
    const filtered = filterFreq === "All" ? items : items.filter(i => i.frequency === filterFreq);
    const answered = Object.keys(responses).length;
    if (answered < filtered.length) {
      if (!window.confirm(`${filtered.length - answered} items not yet answered. Complete anyway?`)) { setSaving(false); return; }
    }
    await supabase.from("checklist_executions").update({ status: "Completed" }).eq("id", executionId);
    const passCount = Object.values(responses).filter(r => r.result==="PASS").length;
    const failCount = Object.values(responses).filter(r => r.result==="FAIL").length;
    const naCount = Object.values(responses).filter(r => r.result==="N/A").length;
    const defects = items.filter(i => responses[i.id]?.result==="FAIL").map(i => `- ${i.item_en}: ${responses[i.id]?.notes||"No notes"}`).join("\n");
    const description = `CIL Checklist completed by ${executedBy}\n\nResults: ${passCount} PASS · ${failCount} FAIL · ${naCount} N/A\n${defects ? `\nDefects found:\n${defects}` : "\nNo defects found."}`;
    const logRecord = { id: uid("LOG"), asset_id: asset.id, asset_name: asset.name, log_type: "Preventive Maintenance", title: `CIL Checklist — ${new Date().toLocaleString("default", { month: "long", year: "numeric" })}`, description, performed_by: executedBy, vendor: null, start_date: TODAY, end_date: TODAY, cost: null, status: failCount > 0 ? "In Progress" : "Completed", downtime_start: null, downtime_end: null, downtime_hours: null };
    await supabase.from("maintenance_logs").insert([logRecord]);
    if (workOrderId) await supabase.from("work_orders").update({ status: "Completed" }).eq("id", workOrderId);
    const { data: openWOs } = await supabase.from("work_orders").select("id").eq("asset", asset.name).in("status", ["Open","In Progress","Pending"]).ilike("title","PM - %");
    if (openWOs?.length) await supabase.from("work_orders").update({ status: "Completed" }).in("id", openWOs.map(w => w.id));
    setSuccess(`Checklist completed! ${failCount > 0 ? `⚠️ ${failCount} defect(s) found.` : "✅ All clear — log added to maintenance history."}`);
    setSaving(false);
  };

  const filteredItems = filterFreq === "All" ? items : items.filter(i => i.frequency === filterFreq);
  const answeredCount = filteredItems.filter(i => responses[i.id]).length;
  const failCount = filteredItems.filter(i => responses[i.id]?.result==="FAIL").length;
  const passCount = filteredItems.filter(i => responses[i.id]?.result==="PASS").length;

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 2000, padding: 16, overflowY: "auto" }}>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, width: "100%", maxWidth: 860, marginTop: 20, marginBottom: 20 }}>
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>CIL Checklist — {asset.name}</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{asset.category} · {asset.location} · {fmtDate(TODAY)}</div>
          </div>
          <button onClick={onClose} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, color: C.muted, cursor: "pointer", fontSize: 18, padding: "2px 10px" }}>✕</button>
        </div>
        <div style={{ padding: "14px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 16, flexWrap: "wrap", alignItems: "center" }}>
          <div style={{ flex: 1 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: C.muted, marginBottom: 6 }}>
              <span>Progress: {answeredCount}/{filteredItems.length}</span>
              <span style={{ color: failCount > 0 ? C.red : C.green }}>{failCount} FAIL · {passCount} PASS</span>
            </div>
            <div style={{ background: C.border, borderRadius: 4, height: 8 }}>
              <div style={{ background: failCount > 0 ? C.accent : C.green, width: `${filteredItems.length > 0 ? (answeredCount/filteredItems.length)*100 : 0}%`, height: 8, borderRadius: 4, transition: "width 0.3s" }} />
            </div>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            {["All","D","W","F","M"].map(f => (
              <button key={f} onClick={() => setFilterFreq(f)} style={{ background: filterFreq===f?(FREQ_COLORS[f]||C.accent):C.surface, color: filterFreq===f?"#fff":C.muted, border: `1px solid ${filterFreq===f?(FREQ_COLORS[f]||C.accent):C.border}`, borderRadius: 6, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>{f==="All"?"All":FREQ_LABELS[f]}</button>
            ))}
          </div>
        </div>
        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <OkBanner msg={success} onDismiss={() => setSuccess(null)} />
          {!executionId && (
            <div style={{ background: C.surface, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 20 }}>
              <div style={{ fontSize: 13, color: C.accent, fontWeight: 700, marginBottom: 12, textTransform: "uppercase" }}>Start Checklist</div>
              <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
                <div style={{ flex: 1, minWidth: 200 }}><Input label="Technician Name *" value={executedBy} onChange={setExecutedBy} /></div>
                <Btn onClick={startExecution}>Start</Btn>
              </div>
            </div>
          )}
          {loading ? <Spinner /> : (
            <div>
              {["Cleaning","Safety","Inspection","Lubrication"].map(cat => {
                const catItems = filteredItems.filter(i => i.category===cat);
                if (!catItems.length) return null;
                return (
                  <div key={cat} style={{ marginBottom: 24 }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: C.text, marginBottom: 12, display: "flex", alignItems: "center", gap: 8 }}>
                      <span>{CAT_ICONS[cat]}</span> {cat} <span style={{ fontSize: 11, color: C.muted, fontWeight: 400 }}>({catItems.length} items)</span>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                      {catItems.map(item => {
                        const resp = responses[item.id];
                        const result = resp?.result;
                        return (
                          <div key={item.id} style={{ background: result==="PASS"?C.green+"11":result==="FAIL"?C.red+"11":C.surface, border: `1px solid ${result==="PASS"?C.green+"44":result==="FAIL"?C.red+"44":C.border}`, borderRadius: 8, padding: "12px 16px" }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
                              <div style={{ flex: 1 }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                                  <span style={{ fontSize: 11, color: C.muted, fontFamily: "monospace" }}>#{item.item_number}</span>
                                  <span style={{ background: FREQ_COLORS[item.frequency]+"22", color: FREQ_COLORS[item.frequency], border: `1px solid ${FREQ_COLORS[item.frequency]}44`, borderRadius: 4, padding: "1px 6px", fontSize: 10, fontWeight: 700 }}>{item.frequency_label}</span>
                                </div>
                                <div style={{ fontSize: 14, color: C.text, fontWeight: 600 }}>{item.item_en}</div>
                                <div style={{ fontSize: 12, color: C.muted, marginTop: 2, direction: "rtl", textAlign: "right" }}>{item.item_ar}</div>
                              </div>
                              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                                {executionId ? (
                                  <>
                                    <button onClick={() => setResponse(item.id,"PASS")} style={{ background: result==="PASS"?C.green:"transparent", color: result==="PASS"?"#fff":C.green, border: `2px solid ${C.green}`, borderRadius: 6, padding: "6px 14px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>✓ PASS</button>
                                    <button onClick={() => setResponse(item.id,"FAIL")} style={{ background: result==="FAIL"?C.red:"transparent", color: result==="FAIL"?"#fff":C.red, border: `2px solid ${C.red}`, borderRadius: 6, padding: "6px 14px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>✗ FAIL</button>
                                    <button onClick={() => setResponse(item.id,"N/A")} style={{ background: result==="N/A"?C.muted:"transparent", color: result==="N/A"?"#fff":C.muted, border: `2px solid ${C.muted}44`, borderRadius: 6, padding: "6px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>N/A</button>
                                  </>
                                ) : <span style={{ fontSize: 12, color: C.muted }}>Start checklist to respond</span>}
                              </div>
                            </div>
                            {result==="FAIL" && executionId && (
                              <div style={{ marginTop: 10 }}>
                                <textarea placeholder="Describe the issue / defect..." value={resp?.notes||""} onChange={e => setResponse(item.id,"FAIL",e.target.value)} rows={2} style={{ width: "100%", background: C.card, border: `1px solid ${C.red}44`, borderRadius: 6, padding: "8px 10px", color: C.text, fontSize: 13, boxSizing: "border-box", resize: "vertical" }} />
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
              {executionId && (
                <div style={{ marginTop: 20, display: "flex", justifyContent: "flex-end", gap: 10 }}>
                  {failCount > 0 && <div style={{ fontSize: 13, color: C.red, display: "flex", alignItems: "center", gap: 6 }}>⚠️ {failCount} item(s) FAILED</div>}
                  <Btn onClick={complete} disabled={saving} color={C.green}>{saving ? "Completing..." : "✓ Complete Checklist"}</Btn>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── MAINTENANCE LOG MODAL ────────────────────────────────────────────────────
function MaintenanceModal({ asset, onClose, isAdmin, vendors }) {
  const [logs, setLogs] = useState([]);
  const [loadingLogs, setLoadingLogs] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [showPartForm, setShowPartForm] = useState(null);
  const [showChecklist, setShowChecklist] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [saving, setSaving] = useState(false);
  const [expandedLog, setExpandedLog] = useState(null);
  const [parts, setParts] = useState({});
  const vendorOptions = ["— None —", ...vendors.filter(v => v.status==="Active").map(v => v.name)];
  const [form, setForm] = useState({ log_type: "Preventive Maintenance", title: "", description: "", performed_by: "", vendor: "", start_date: TODAY, end_date: "", cost: "", status: "Completed", downtime_start: "", downtime_end: "" });
  const [partForm, setPartForm] = useState({ part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const pf = (k) => (v) => setPartForm(p => ({ ...p, [k]: v }));

  useEffect(() => { loadLogs(); }, [asset.id]);

  const loadLogs = async () => {
    setLoadingLogs(true);
    const { data } = await supabase.from("maintenance_logs").select("*").eq("asset_id", asset.id).order("start_date", { ascending: false });
    setLogs(data || []);
    setLoadingLogs(false);
  };

  const loadParts = async (logId) => {
    const { data } = await supabase.from("spare_parts").select("*").eq("log_id", logId);
    setParts(prev => ({ ...prev, [logId]: data || [] }));
  };

  const toggleLog = (logId) => {
    setExpandedLog(expandedLog===logId ? null : logId);
    if (!parts[logId]) loadParts(logId);
  };

  const submitLog = async () => {
    if (!form.title) { setError("Title is required."); return; }
    setSaving(true); setError(null);
    const record = {
      id: uid("LOG"), asset_id: asset.id, asset_name: asset.name, log_type: form.log_type, title: form.title, description: form.description, performed_by: form.performed_by, vendor: form.vendor==="— None —"?null:form.vendor||null, start_date: form.start_date||null, end_date: form.end_date||null, cost: form.cost?parseFloat(form.cost):null, status: form.status,
      downtime_start: form.downtime_start||null,
      downtime_end: form.downtime_end||null,
      downtime_hours: (form.downtime_start && form.downtime_end) ? Math.round((new Date(form.downtime_end) - new Date(form.downtime_start)) / (1000 * 60 * 60)) : null,
    };
    const { error: err } = await supabase.from("maintenance_logs").insert([record]);
    if (err) { setError(err.message); } else { setSuccess("Log added!"); setLogs(prev => [record,...prev]); setForm({ log_type: "Preventive Maintenance", title: "", description: "", performed_by: "", vendor: "", start_date: TODAY, end_date: "", cost: "", status: "Completed", downtime_start: "", downtime_end: "" }); setShowForm(false); }
    setSaving(false);
  };

  const submitPart = async (logId) => {
    if (!partForm.part_name) { setError("Part name is required."); return; }
    setSaving(true); setError(null);
    const qty = parseFloat(partForm.quantity)||1;
    const unitCost = parseFloat(partForm.unit_cost)||0;
    const record = { id: uid("PRT"), log_id: logId, asset_id: asset.id, part_name: partForm.part_name, part_number: partForm.part_number, quantity: qty, unit_cost: unitCost, total_cost: qty*unitCost, supplier: partForm.supplier };
    const { error: err } = await supabase.from("spare_parts").insert([record]);
    if (err) { setError(err.message); } else { setSuccess("Part added!"); setParts(prev => ({ ...prev, [logId]: [...(prev[logId]||[]),record] })); setPartForm({ part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "" }); setShowPartForm(null); }
    setSaving(false);
  };

  const deletePart = async (partId, logId) => { await supabase.from("spare_parts").delete().eq("id", partId); setParts(prev => ({ ...prev, [logId]: prev[logId].filter(p => p.id!==partId) })); };
  const deleteLog = async (logId) => { await supabase.from("spare_parts").delete().eq("log_id", logId); await supabase.from("maintenance_logs").delete().eq("id", logId); setLogs(prev => prev.filter(l => l.id!==logId)); };
  const totalCost = logs.reduce((s,l) => s+(l.cost||0), 0);

  return (
    <>
      {showChecklist && <ChecklistModal asset={asset} onClose={() => setShowChecklist(false)} />}
      <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1000, padding: 16, overflowY: "auto" }}>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, width: "100%", maxWidth: 820, marginTop: 20, marginBottom: 20 }}>
          <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
            <div><div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>{asset.name}</div><div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{asset.category} · {asset.location}</div></div>
            <button onClick={onClose} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, color: C.muted, cursor: "pointer", fontSize: 18, padding: "2px 10px" }}>✕</button>
          </div>
          <div style={{ padding: "14px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 12, flexWrap: "wrap" }}>
            {[["📋",logs.length,"Total Logs",C.blue],["💰",fmt(totalCost),"Total Cost",C.accent],["🔧",logs[0]?.start_date?fmtDate(logs[0].start_date):"Never","Last Maintenance",C.green],["⚙️",asset.pm_frequency?`Every ${asset.pm_frequency} mo.`:"—","PM Frequency",C.yellow]].map(([icon,val,label,color]) => (
              <div key={label} style={{ background: C.surface, borderRadius: 8, padding: "10px 16px", flex: "1 1 130px", borderLeft: `3px solid ${color}` }}>
                <div style={{ fontSize: 16 }}>{icon}</div><div style={{ fontSize: 18, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{val}</div><div style={{ fontSize: 11, color: C.muted }}>{label}</div>
              </div>
            ))}
          </div>
          <div style={{ padding: 24 }}>
            <ErrBanner msg={error} onDismiss={() => setError(null)} />
            <OkBanner msg={success} onDismiss={() => setSuccess(null)} />
            <div style={{ display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
              <button onClick={() => setShowChecklist(true)} style={{ background: C.blue+"22", color: C.blue, border: `1px solid ${C.blue}44`, borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>📋 Run CIL Checklist</button>
              <Btn onClick={() => setShowForm(v => !v)}>+ Add Maintenance Log</Btn>
            </div>
            {showForm && (
              <div style={{ background: C.surface, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 20 }}>
                <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>New Maintenance Log</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
                  <Sel label="Type" value={form.log_type} onChange={f("log_type")} options={LOG_TYPES} />
                  <Input label="Title *" value={form.title} onChange={f("title")} />
                  <Input label="Performed By" value={form.performed_by} onChange={f("performed_by")} />
                  <Sel label="Vendor" value={form.vendor} onChange={f("vendor")} options={vendorOptions} />
                  <Input label="Start Date" value={form.start_date} onChange={f("start_date")} type="date" />
                  <Input label="End Date" value={form.end_date} onChange={f("end_date")} type="date" />
                  <Input label="Total Cost ($)" value={form.cost} onChange={f("cost")} type="number" />
                  <Sel label="Status" value={form.status} onChange={f("status")} options={LOG_STATUSES} />
                  <Input label="Downtime Start" value={form.downtime_start} onChange={f("downtime_start")} type="date" />
                  <Input label="Back to Operation" value={form.downtime_end} onChange={f("downtime_end")} type="date" />
                </div>
                <div style={{ marginTop: 12 }}><Textarea label="Description / Notes" value={form.description} onChange={f("description")} /></div>
                <div style={{ display: "flex", gap: 8, marginTop: 14 }}><Btn onClick={submitLog} disabled={saving}>{saving?"Saving...":"Save Log"}</Btn><Btn variant="secondary" onClick={() => setShowForm(false)}>Cancel</Btn></div>
              </div>
            )}
            <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 14 }}>Maintenance History</div>
            {loadingLogs ? <Spinner /> : logs.length===0 ? (
              <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>No maintenance records yet.</div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {logs.map(log => (
                  <div key={log.id} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10, overflow: "hidden" }}>
                    <div onClick={() => toggleLog(log.id)} style={{ padding: "12px 16px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <span style={{ fontSize: 18 }}>{log.log_type==="Preventive Maintenance"?"🔧":log.log_type==="Corrective Repair"?"🔨":log.log_type==="Inspection"?"🔍":log.log_type==="Overhaul"?"⚙️":"🔩"}</span>
                        <div><div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{log.title}</div><div style={{ fontSize: 11, color: C.muted }}>{fmtDate(log.start_date)}{log.performed_by?` · ${log.performed_by}`:""}{log.vendor?` · ${log.vendor}`:""}</div></div>
                      </div>
                      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        <Badge label={log.log_type} color={statusColor(log.log_type)} />
                        <Badge label={log.status} color={statusColor(log.status)} />
                        {log.cost>0 && <span style={{ fontSize: 12, color: C.accent, fontWeight: 700 }}>{fmt(log.cost)}</span>}
                        <span style={{ color: C.muted }}>{expandedLog===log.id?"▲":"▼"}</span>
                      </div>
                    </div>
                    {expandedLog===log.id && (
                      <div style={{ padding: "0 16px 16px", borderTop: `1px solid ${C.border}` }}>
                        {log.description && <div style={{ marginTop: 12, padding: 12, background: C.card, borderRadius: 8, fontSize: 13, color: C.subtle, lineHeight: 1.6 }}>{log.description}</div>}
                        {(log.downtime_start || log.downtime_end) && (
                          <div style={{ marginTop: 12, background: C.red+"11", border: `1px solid ${C.red}33`, borderRadius: 8, padding: "12px 16px", display: "flex", gap: 20, flexWrap: "wrap" }}>
                            <div><div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase" }}>⛔ Equipment Down</div><div style={{ fontSize: 13, color: C.red, fontWeight: 700, marginTop: 2 }}>{fmtDate(log.downtime_start)||"—"}</div></div>
                            <div style={{ fontSize: 20, color: C.muted, alignSelf: "center" }}>→</div>
                            <div><div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase" }}>✅ Back to Operation</div><div style={{ fontSize: 13, color: C.green, fontWeight: 700, marginTop: 2 }}>{fmtDate(log.downtime_end)||"Pending"}</div></div>
                            {log.downtime_hours && <div><div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase" }}>⏱ Total Downtime</div><div style={{ fontSize: 13, color: C.yellow, fontWeight: 700, marginTop: 2 }}>{log.downtime_hours}h ({Math.round(log.downtime_hours/24)}d)</div></div>}
                          </div>
                        )}
                        <div style={{ marginTop: 14 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                            <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>🔩 Spare Parts Used</div>
                            {isAdmin && <Btn small onClick={() => setShowPartForm(showPartForm===log.id?null:log.id)}>+ Add Part</Btn>}
                          </div>
                          {showPartForm===log.id && (
                            <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 8, padding: 14, marginBottom: 12 }}>
                              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10 }}>
                                <Input label="Part Name *" value={partForm.part_name} onChange={pf("part_name")} />
                                <Input label="Part Number" value={partForm.part_number} onChange={pf("part_number")} />
                                <Input label="Qty" value={partForm.quantity} onChange={pf("quantity")} type="number" />
                                <Input label="Unit Cost ($)" value={partForm.unit_cost} onChange={pf("unit_cost")} type="number" />
                                <Input label="Supplier" value={partForm.supplier} onChange={pf("supplier")} />
                              </div>
                              <div style={{ display: "flex", gap: 8, marginTop: 10 }}><Btn small onClick={() => submitPart(log.id)} disabled={saving}>Save</Btn><Btn small variant="secondary" onClick={() => setShowPartForm(null)}>Cancel</Btn></div>
                            </div>
                          )}
                          {parts[log.id]?.length > 0 ? (
                            <div style={{ overflowX: "auto" }}>
                              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                                <thead><tr style={{ borderBottom: `1px solid ${C.border}` }}>{["Part Name","Part No.","Qty","Unit","Total","Supplier",...(isAdmin?[""]:[])].map(h => <th key={h} style={{ textAlign: "left", padding: "6px 10px", color: C.muted, fontWeight: 600, fontSize: 11, textTransform: "uppercase" }}>{h}</th>)}</tr></thead>
                                <tbody>
                                  {parts[log.id].map(part => (
                                    <tr key={part.id} style={{ borderBottom: `1px solid ${C.border}22` }}>
                                      <td style={{ padding: "8px 10px", color: C.text, fontWeight: 600 }}>{part.part_name}</td>
                                      <td style={{ padding: "8px 10px", color: C.subtle, fontFamily: "monospace", fontSize: 11 }}>{part.part_number||"—"}</td>
                                      <td style={{ padding: "8px 10px", color: C.subtle }}>{part.quantity}</td>
                                      <td style={{ padding: "8px 10px", color: C.subtle }}>{part.unit_cost?`$${part.unit_cost}`:"—"}</td>
                                      <td style={{ padding: "8px 10px", color: C.accent, fontWeight: 700 }}>{part.total_cost?`$${part.total_cost}`:"—"}</td>
                                      <td style={{ padding: "8px 10px", color: C.subtle }}>{part.supplier||"—"}</td>
                                      {isAdmin && <td style={{ padding: "8px 10px" }}><Btn small variant="danger" onClick={() => deletePart(part.id, log.id)}>Del</Btn></td>}
                                    </tr>
                                  ))}
                                  <tr style={{ borderTop: `1px solid ${C.border}` }}>
                                    <td colSpan={4} style={{ padding: "8px 10px", color: C.muted, fontSize: 11 }}>Parts Total</td>
                                    <td style={{ padding: "8px 10px", color: C.accent, fontWeight: 700 }}>${parts[log.id].reduce((s,p) => s+(p.total_cost||0),0).toLocaleString()}</td>
                                    <td colSpan={isAdmin?2:1} />
                                  </tr>
                                </tbody>
                              </table>
                            </div>
                          ) : <div style={{ fontSize: 12, color: C.muted }}>No spare parts recorded.</div>}
                        </div>
                        {isAdmin && <div style={{ marginTop: 12, display: "flex", justifyContent: "flex-end" }}><Btn small variant="danger" onClick={() => deleteLog(log.id)}>Delete Log</Btn></div>}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>
    </>
  );
}

function LoginScreen() {
  const [email, setEmail] = useState(""); const [password, setPassword] = useState("");
  const [error, setError] = useState(null); const [loading, setLoading] = useState(false);
  const signIn = async () => {
    if (!email || !password) { setError("Please enter your email and password."); return; }
    setLoading(true); setError(null);
    const { error: err } = await supabase.auth.signInWithPassword({ email, password });
    if (err) setError(err.message);
    setLoading(false);
  };
  return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans', sans-serif", padding: 16 }}>
      <div style={{ width: "100%", maxWidth: 400 }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ background: C.accent, borderRadius: 12, width: 56, height: 56, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 28, margin: "0 auto 16px" }}>🏭</div>
          <div style={{ fontFamily: "monospace", fontSize: 20, letterSpacing: 3, color: C.text, fontWeight: 800 }}>FACILITY COMMAND</div>
          <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>INDUSTRIAL WAREHOUSE MANAGEMENT</div>
        </div>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 24 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.text, marginBottom: 20 }}>Sign in to your account</div>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <Input label="Email" value={email} onChange={setEmail} type="email" />
            <Input label="Password" value={password} onChange={setPassword} type="password" />
            <button onClick={signIn} disabled={loading} style={{ background: C.accent, color: "#fff", border: "none", borderRadius: 6, padding: "12px", fontSize: 15, fontWeight: 700, cursor: loading?"not-allowed":"pointer", opacity: loading?0.7:1 }}>
              {loading ? "Signing in..." : "Sign In"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
function WorkOrderPhotosModal({ workOrder, onClose }) {
  const [photos, setPhotos] = useState([]);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);

  useEffect(() => { loadPhotos(); }, [workOrder.id]);

  const loadPhotos = async () => {
    const { data } = await supabase.storage.from("work-order-photos").list(workOrder.id);
    if (data) {
      const urls = data.map(file => ({
        name: file.name,
        url: supabase.storage.from("work-order-photos").getPublicUrl(`${workOrder.id}/${file.name}`).data.publicUrl,
      }));
      setPhotos(urls);
    }
  };

  const uploadPhoto = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (file.size > 5 * 1024 * 1024) { setError("File too large. Max 5MB."); return; }
    setUploading(true); setError(null);
    const fileName = `${Date.now()}-${file.name}`;
    const { error: err } = await supabase.storage.from("work-order-photos").upload(`${workOrder.id}/${fileName}`, file);
    if (err) { setError(err.message); } else {
      setSuccess("Photo uploaded!");
      await loadPhotos();
    }
    setUploading(false);
  };

  const deletePhoto = async (name) => {
    await supabase.storage.from("work-order-photos").remove([`${workOrder.id}/${name}`]);
    setPhotos(prev => prev.filter(p => p.name !== name));
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1000, padding: 16, overflowY: "auto" }}>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, width: "100%", maxWidth: 720, marginTop: 20, marginBottom: 20 }}>
        
        {/* Header */}
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>📷 Photos — {workOrder.title}</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{workOrder.asset} · {workOrder.id}</div>
          </div>
          <button onClick={onClose} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, color: C.muted, cursor: "pointer", fontSize: 18, padding: "2px 10px" }}>✕</button>
        </div>

        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <OkBanner msg={success} onDismiss={() => setSuccess(null)} />

          {/* Upload Button */}
          <div style={{ marginBottom: 24 }}>
            <label style={{
              display: "inline-flex", alignItems: "center", gap: 8,
              background: C.accent, color: "#fff", borderRadius: 6,
              padding: "10px 20px", fontSize: 13, fontWeight: 700,
              cursor: uploading ? "not-allowed" : "pointer", opacity: uploading ? 0.7 : 1,
            }}>
              {uploading ? "⏳ Uploading..." : "📷 Upload Photo"}
              <input type="file" accept="image/*" onChange={uploadPhoto} style={{ display: "none" }} disabled={uploading} />
            </label>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 8 }}>Max 5MB per photo. JPG, PNG, WEBP supported.</div>
          </div>

          {/* Photos Grid */}
          {photos.length === 0 ? (
            <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13, border: `2px dashed ${C.border}`, borderRadius: 10 }}>
              No photos yet. Upload the first photo for this work order.
            </div>
          ) : (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 14 }}>
              {photos.map(photo => (
                <div key={photo.name} style={{ background: C.surface, borderRadius: 10, overflow: "hidden", border: `1px solid ${C.border}` }}>
                  <img
                    src={photo.url}
                    alt={photo.name}
                    style={{ width: "100%", height: 160, objectFit: "cover", display: "block", cursor: "pointer" }}
                    onClick={() => window.open(photo.url, "_blank")}
                  />
                  <div style={{ padding: "8px 12px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div style={{ fontSize: 11, color: C.muted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 140 }}>{photo.name}</div>
                    <Btn small variant="danger" onClick={() => deletePhoto(photo.name)}>Del</Btn>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
function WorkOrders({ workOrders, setWorkOrders, loading, onAdd, isAdmin, vendors, assets }) {
  const [showForm, setShowForm] = useState(false); const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null); const [filter, setFilter] = useState("All");
  const [editItem, setEditItem] = useState(null); const [deleteItem, setDeleteItem] = useState(null); const [selectedWO, setSelectedWO] = useState(null);
  const [form, setForm] = useState({ title: "", asset: "", priority: "Medium", start_date: "", due: "", vendor: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const filtered = filter==="All" ? workOrders : workOrders.filter(w => w.status===filter);
  const vendorOptions = ["— None —",...vendors.filter(v => v.status==="Active").map(v => v.name)];
  const submit = async () => {
    if (!form.title||!form.asset) { setError("Title and Asset are required."); return; }
    setSaving(true); setError(null);
    const vendorName = form.vendor === "— None —" || !form.vendor ? null : form.vendor;
const record = { id: uid("WO"), title: form.title, asset: form.asset, priority: form.priority, status: "Open", assignee: null, start_date: form.start_date||null, due: form.due||null, vendor: vendorName };
const { error: err } = await supabase.from("work_orders").insert([record]);
if (err) { setError(err.message); } else {
  onAdd(record);
  if (vendorName) {
    const { data: vd } = await supabase.from("vendors").select("id, open_orders").eq("name", vendorName).single();
    if (vd) await supabase.from("vendors").update({ open_orders: (vd.open_orders || 0) + 1 }).eq("id", vd.id).then(() => {});
  }
  setForm({ title: "", asset: "", priority: "Medium", start_date: "", due: "", vendor: "" });
  setShowForm(false);
}
    setSaving(false);
  };
  const updateStatus = async (id, val) => {
  const wo = workOrders.find(w => w.id === id);
  await supabase.from("work_orders").update({ status: val }).eq("id", id);
  setWorkOrders(prev => prev.map(w => w.id === id ? { ...w, status: val } : w));
  // Update vendor open orders when completing
  if (wo?.vendor && val === "Completed" && wo.status !== "Completed") {
    const { data: vendorData } = await supabase.from("vendors").select("id, open_orders").eq("name", wo.vendor).single();
    if (vendorData) await supabase.from("vendors").update({ open_orders: Math.max(0, (vendorData.open_orders || 0) - 1) }).eq("id", vendorData.id);
  }
  // Increase when reopening
  if (wo?.vendor && wo.status === "Completed" && val !== "Completed") {
    const { data: vendorData } = await supabase.from("vendors").select("id, open_orders").eq("name", wo.vendor).single();
    if (vendorData) await supabase.from("vendors").update({ open_orders: (vendorData.open_orders || 0) + 1 }).eq("id", vendorData.id);
  }
  // Auto-create maintenance log when work order is completed
  if (val === "Completed" && wo?.status !== "Completed" && wo?.asset) {
    const { data: assetData } = await supabase.from("assets").select("id").eq("name", wo.asset).single();
    if (assetData) {
      await supabase.from("maintenance_logs").insert([{
        id: uid("LOG"), asset_id: assetData.id, asset_name: wo.asset,
        log_type: wo.title.startsWith("PM -") ? "Preventive Maintenance" : "Corrective Repair",
        title: wo.title,
        description: `Work order completed.\nVendor: ${wo.vendor || "—"}\nPriority: ${wo.priority}`,
        performed_by: wo.assignee || "—",
        vendor: wo.vendor || null,
        start_date: wo.start_date || TODAY,
        end_date: TODAY,
        cost: null, status: "Completed",
        downtime_start: null, downtime_end: null, downtime_hours: null,
      }]);
    }
  }
};
  const updatePriority = async (id,val) => { await supabase.from("work_orders").update({ priority: val }).eq("id",id); setWorkOrders(prev => prev.map(wo => wo.id===id?{...wo,priority:val}:wo)); };
  const saveEdit = async (updated) => { const { error: err } = await supabase.from("work_orders").update(updated).eq("id",updated.id); if (!err) { setWorkOrders(prev => prev.map(wo => wo.id===updated.id?updated:wo)); setEditItem(null); } else setError(err.message); };
  const confirmDelete = async () => { await supabase.from("work_orders").delete().eq("id",deleteItem.id); setWorkOrders(prev => prev.filter(wo => wo.id!==deleteItem.id)); setDeleteItem(null); };
  return (
    <div>
      {selectedWO && <WorkOrderPhotosModal workOrder={selectedWO} onClose={() => setSelectedWO(null)} />}
      {editItem && <EditModal title="Work Order" data={editItem} fields={[{key:"title",label:"Title"},{key:"asset",label:"Asset"},{key:"priority",label:"Priority",options:["Critical","High","Medium","Low"]},{key:"status",label:"Status",options:["Open","In Progress","Pending","Completed"]},{key:"vendor",label:"Vendor",options:vendorOptions},{key:"assignee",label:"Assignee"},{key:"start_date",label:"Start Date",type:"date"},{key:"due",label:"Due Date",type:"date"}]} onSave={saveEdit} onClose={() => setEditItem(null)} />}
      {deleteItem && <ConfirmDel name={deleteItem.title} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {["All","Open","In Progress","Pending","Completed"].map(s => <button key={s} onClick={() => setFilter(s)} style={{ background: filter===s?C.accent:C.card, color: filter===s?"#fff":C.muted, border: `1px solid ${filter===s?C.accent:C.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>{s}</button>)}
        </div>
        <Btn onClick={() => setShowForm(v => !v)}>+ New Work Order</Btn>
      </div>
      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 18 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>New Work Order</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label="Title *" value={form.title} onChange={f("title")} /><div>
  <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>Asset *</div>
  <select value={form.asset} onChange={e => f("asset")(e.target.value)}
    style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14 }}>
    <option value="">— Select Asset —</option>
    {(assets||[]).map(a => <option key={a.id} value={a.name}>{a.name} ({a.location})</option>)}
  </select>
</div>
            <Input label="Start Date" value={form.start_date} onChange={f("start_date")} type="date" /><Input label="Due Date" value={form.due} onChange={f("due")} type="date" />
            <Sel label="Priority" value={form.priority} onChange={f("priority")} options={["Critical","High","Medium","Low"]} /><Sel label="Vendor" value={form.vendor} onChange={f("vendor")} options={vendorOptions} />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}><Btn onClick={submit} disabled={saving}>{saving?"Saving...":"Create"}</Btn><Btn variant="secondary" onClick={() => setShowForm(false)}>Cancel</Btn></div>
        </div>
      )}
      {loading ? <Spinner /> : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 700 }}>
            <thead><tr style={{ borderBottom: `1px solid ${C.border}` }}>{["ID","Title","Asset","Priority","Status","Vendor","Start","Due","📷",...(isAdmin?["Actions"]:[])].map(h => <th key={h} style={{ textAlign: "left", padding: "8px 12px", fontSize: 11, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600 }}>{h}</th>)}</tr></thead>
            <tbody>
              {filtered.map((wo,i) => (
                <tr key={wo.id} style={{ borderBottom: `1px solid ${C.border}22`, background: i%2===0?"transparent":C.surface+"44" }}>
                  <td style={{ padding: "10px 12px", fontSize: 11, color: C.muted, fontFamily: "monospace" }}>{wo.id}</td>
                  <td style={{ padding: "10px 12px", fontSize: 13, color: C.text, fontWeight: 600 }}>{wo.title}</td>
                  <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.asset}</td>
                  <td style={{ padding: "10px 12px" }}><StatusSel value={wo.priority} options={["Critical","High","Medium","Low"]} onChange={val => updatePriority(wo.id,val)} /></td>
                  <td style={{ padding: "10px 12px" }}><StatusSel value={wo.status} options={["Open","In Progress","Pending","Completed"]} onChange={val => updateStatus(wo.id,val)} /></td>
                  <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.vendor||"—"}</td>
                  <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.start_date||"—"}</td>
                  <td style={{ padding: "10px 12px", fontSize: 12, color: wo.due&&wo.due<=TODAY&&wo.status!=="Completed"?C.red:C.subtle }}>{wo.due||"—"}</td>
                  <td style={{ padding: "10px 12px" }}>
  <div style={{ display: "flex", gap: 6 }}>
    <Btn small onClick={() => setSelectedWO(wo)} color={C.purple}>📷</Btn>
    {isAdmin && <><Btn small onClick={() => setEditItem(wo)} color={C.blue}>Edit</Btn><Btn small variant="danger" onClick={() => setDeleteItem(wo)}>Del</Btn></>}
  </div>
</td>
                </tr>
              ))}
              {filtered.length===0 && <tr><td colSpan={isAdmin?10:9} style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>No work orders found.</td></tr>}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function Assets({ assets, setAssets, loading, onAdd, isAdmin, vendors }) {
  const [showForm, setShowForm] = useState(false); const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null); const [editItem, setEditItem] = useState(null);
  const [deleteItem, setDeleteItem] = useState(null); const [selectedAsset, setSelectedAsset] = useState(null);
  const [siteFilter, setSiteFilter] = useState("All"); const [catFilter, setCatFilter] = useState("All"); const [search, setSearch] = useState("");
  const [form, setForm] = useState({ name: "", category: "", location: "— Select Site —", value: "", next_service: "", pm_frequency: "1", pm_task: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const categories = ["All",...new Set(assets.map(a => a.category).filter(Boolean))];
  const filtered = assets.filter(a => (siteFilter==="All"||a.location===siteFilter)&&(catFilter==="All"||a.category===catFilter)&&(!search||a.name.toLowerCase().includes(search.toLowerCase())));
  const submit = async () => {
    if (!form.name) { setError("Asset name is required."); return; }
    if (form.location==="— Select Site —") { setError("Please select a site."); return; }
    setSaving(true); setError(null);
    const record = { id: uid("AST"), name: form.name, category: form.category, location: form.location, value: form.value, status: "Operational", last_service: TODAY, next_service: form.next_service||null, pm_frequency: parseInt(form.pm_frequency)||1, pm_task: form.pm_task||"Scheduled Maintenance", last_pm_date: null };
    const { error: err } = await supabase.from("assets").insert([record]);
    if (err) { setError(err.message); } else { onAdd(record); setForm({ name: "", category: "", location: "— Select Site —", value: "", next_service: "", pm_frequency: "1", pm_task: "" }); setShowForm(false); }
    setSaving(false);
  };
  const updateStatus = async (id,val) => { await supabase.from("assets").update({ status: val }).eq("id",id); setAssets(prev => prev.map(a => a.id===id?{...a,status:val}:a)); };
  const saveEdit = async (updated) => { const { error: err } = await supabase.from("assets").update(updated).eq("id",updated.id); if (!err) { setAssets(prev => prev.map(a => a.id===updated.id?updated:a)); setEditItem(null); } else setError(err.message); };
  const confirmDelete = async () => { await supabase.from("assets").delete().eq("id",deleteItem.id); setAssets(prev => prev.filter(a => a.id!==deleteItem.id)); setDeleteItem(null); };
  return (
    <div>
      {editItem && <EditModal title="Asset" data={editItem} fields={[{key:"name",label:"Asset Name"},{key:"category",label:"Category"},{key:"location",label:"Site",options:SITES},{key:"value",label:"Est. Value"},{key:"status",label:"Status",options:["Operational","Under Maintenance","Degraded"]},{key:"pm_frequency",label:"PM Frequency (months)"},{key:"pm_task",label:"PM Task"},{key:"last_service",label:"Last Service",type:"date"},{key:"next_service",label:"Next Service",type:"date"}]} onSave={saveEdit} onClose={() => setEditItem(null)} />}
      {deleteItem && <ConfirmDel name={deleteItem.name} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      {selectedAsset && <MaintenanceModal asset={selectedAsset} onClose={() => setSelectedAsset(null)} isAdmin={isAdmin} vendors={vendors} />}
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 18 }}>
        <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search assets..." style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 12px", color: C.text, fontSize: 13, flex: "1 1 180px" }} />
        <select value={siteFilter} onChange={e => setSiteFilter(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12 }}>
          <option>All</option>{SITES.filter(s => s!=="— Select Site —").map(s => <option key={s}>{s}</option>)}
        </select>
        <select value={catFilter} onChange={e => setCatFilter(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12 }}>
          {categories.map(c => <option key={c}>{c}</option>)}
        </select>
        <Btn onClick={() => setShowForm(v => !v)}>+ Add Asset</Btn>
      </div>
      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 18 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>Register New Asset</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label="Asset Name *" value={form.name} onChange={f("name")} /><Input label="Category" value={form.category} onChange={f("category")} />
            <Sel label="Site *" value={form.location} onChange={f("location")} options={SITES} /><Input label="Est. Value" value={form.value} onChange={f("value")} />
            <Input label="Next Service Date" value={form.next_service} onChange={f("next_service")} type="date" />
            <Sel label="PM Frequency (months)" value={form.pm_frequency} onChange={f("pm_frequency")} options={["1","2","3","6","12"]} />
            <Input label="PM Task" value={form.pm_task} onChange={f("pm_task")} />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}><Btn onClick={submit} disabled={saving}>{saving?"Saving...":"Register"}</Btn><Btn variant="secondary" onClick={() => setShowForm(false)}>Cancel</Btn></div>
        </div>
      )}
      {loading ? <Spinner /> : (
        <div>
          <div style={{ fontSize: 12, color: C.muted, marginBottom: 12 }}>{filtered.length} assets shown</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 14 }}>
            {filtered.map(a => (
              <div key={a.id} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 18, borderTop: `3px solid ${statusColor(a.status)}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                  <div style={{ flex: 1, marginRight: 8 }}><div style={{ fontSize: 14, fontWeight: 700, color: C.text, lineHeight: 1.3 }}>{a.name}</div><div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>{a.category}</div></div>
                  <StatusSel value={a.status} options={["Operational","Under Maintenance","Degraded"]} onChange={val => updateStatus(a.id,val)} />
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12, marginBottom: 14 }}>
                  {[["Site",a.location],["Value",a.value],["PM Every",a.pm_frequency?`${a.pm_frequency} mo.`:"—"],["Last PM",a.last_pm_date?fmtDate(a.last_pm_date):"Never"]].map(([lbl,val]) => (
                    <div key={lbl}><div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase" }}>{lbl}</div><div style={{ color: C.subtle, marginTop: 2 }}>{val||"—"}</div></div>
                  ))}
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <button onClick={() => setSelectedAsset(a)} style={{ flex: 1, background: C.blue+"22", color: C.blue, border: `1px solid ${C.blue}44`, borderRadius: 6, padding: "7px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>📋 Log & Checklist</button>
                  {isAdmin && <><Btn small onClick={() => setEditItem(a)} color={C.accent}>Edit</Btn><Btn small variant="danger" onClick={() => setDeleteItem(a)}>Del</Btn></>}
                </div>
              </div>
            ))}
            {filtered.length===0 && <div style={{ color: C.muted, fontSize: 13, padding: 32 }}>No assets found.</div>}
          </div>
        </div>
      )}
    </div>
  );
}
function VendorWorkOrdersModal({ vendor, onClose }) {
  const [workOrders, setWorkOrders] = useState([]);
  const [loading, setLoading] = useState(true);
  const [filter, setFilter] = useState("All");

  useEffect(() => { loadWorkOrders(); }, [vendor.id]);

  const loadWorkOrders = async () => {
    setLoading(true);
    const { data } = await supabase.from("work_orders").select("*")
      .eq("vendor", vendor.name).order("due", { ascending: false });
    setWorkOrders(data || []);
    setLoading(false);
  };

  const filtered = filter === "All" ? workOrders : workOrders.filter(w => w.status === filter);
  const open = workOrders.filter(w => w.status !== "Completed").length;
  const completed = workOrders.filter(w => w.status === "Completed").length;

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1000, padding: 16, overflowY: "auto" }}>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, width: "100%", maxWidth: 820, marginTop: 20, marginBottom: 20 }}>
        
        {/* Header */}
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>{vendor.name}</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{vendor.specialty}</div>
          </div>
          <button onClick={onClose} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, color: C.muted, cursor: "pointer", fontSize: 18, padding: "2px 10px" }}>✕</button>
        </div>

        {/* Stats */}
        <div style={{ padding: "14px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 12, flexWrap: "wrap" }}>
          {[
            ["📋", workOrders.length, "Total Work Orders", C.blue],
            ["🔓", open, "Open / In Progress", C.accent],
            ["✅", completed, "Completed", C.green],
          ].map(([icon, val, label, color]) => (
            <div key={label} style={{ background: C.surface, borderRadius: 8, padding: "10px 16px", flex: "1 1 130px", borderLeft: `3px solid ${color}` }}>
              <div style={{ fontSize: 16 }}>{icon}</div>
              <div style={{ fontSize: 22, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{val}</div>
              <div style={{ fontSize: 11, color: C.muted }}>{label}</div>
            </div>
          ))}
        </div>

        <div style={{ padding: 24 }}>
          {/* Filter */}
          <div style={{ display: "flex", gap: 6, marginBottom: 18, flexWrap: "wrap" }}>
            {["All", "Open", "In Progress", "Pending", "Completed"].map(s => (
              <button key={s} onClick={() => setFilter(s)} style={{ background: filter===s?C.accent:C.surface, color: filter===s?"#fff":C.muted, border: `1px solid ${filter===s?C.accent:C.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>{s}</button>
            ))}
          </div>

          {/* Work Orders Table */}
          {loading ? <Spinner /> : filtered.length === 0 ? (
            <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>No work orders found.</div>
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
                <thead>
                  <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                    {["Title", "Asset", "Priority", "Status", "Start", "Due"].map(h => (
                      <th key={h} style={{ textAlign: "left", padding: "8px 12px", fontSize: 11, color: C.muted, textTransform: "uppercase", fontWeight: 600 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((wo, i) => (
                    <tr key={wo.id} style={{ borderBottom: `1px solid ${C.border}22`, background: i%2===0?"transparent":C.surface+"44" }}>
                      <td style={{ padding: "10px 12px", fontSize: 13, color: C.text, fontWeight: 600 }}>{wo.title}</td>
                      <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.asset}</td>
                      <td style={{ padding: "10px 12px" }}><Badge label={wo.priority} color={priorityColor(wo.priority)} /></td>
                      <td style={{ padding: "10px 12px" }}><Badge label={wo.status} color={statusColor(wo.status)} /></td>
                      <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.start_date || "—"}</td>
                      <td style={{ padding: "10px 12px", fontSize: 12, color: wo.due && wo.due <= TODAY && wo.status !== "Completed" ? C.red : C.subtle }}>{wo.due || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
function Vendors({ vendors, setVendors, loading, onAdd, isAdmin }) {
  const [showForm, setShowForm] = useState(false); const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null); const [editItem, setEditItem] = useState(null); const [deleteItem, setDeleteItem] = useState(null); const [selectedVendor, setSelectedVendor] = useState(null);
  const [form, setForm] = useState({ name: "", specialty: "", contact: "", phone: "", email: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const submit = async () => {
    if (!form.name) { setError("Company name is required."); return; }
    setSaving(true); setError(null);
    const record = { id: uid("VND"), ...form, status: "Active", rating: 0, open_orders: 0 };
    const { error: err } = await supabase.from("vendors").insert([record]);
    if (err) { setError(err.message); } else { onAdd(record); setForm({ name: "", specialty: "", contact: "", phone: "", email: "" }); setShowForm(false); }
    setSaving(false);
  };
  const saveEdit = async (updated) => { const { error: err } = await supabase.from("vendors").update(updated).eq("id",updated.id); if (!err) { setVendors(prev => prev.map(v => v.id===updated.id?updated:v)); setEditItem(null); } else setError(err.message); };
  const confirmDelete = async () => { await supabase.from("vendors").delete().eq("id",deleteItem.id); setVendors(prev => prev.filter(v => v.id!==deleteItem.id)); setDeleteItem(null); };
  const Stars = ({ rating }) => <div style={{ display: "flex", gap: 2 }}>{[1,2,3,4,5].map(i => <span key={i} style={{ color: i<=Math.floor(rating)?C.yellow:C.border, fontSize: 14 }}>*</span>)}<span style={{ fontSize: 11, color: C.muted, marginLeft: 4 }}>{rating>0?Number(rating).toFixed(1):"N/A"}</span></div>;
  return (
    <div>
      {selectedVendor && <VendorWorkOrdersModal vendor={selectedVendor} onClose={() => setSelectedVendor(null)} />}
      {editItem && <EditModal title="Vendor" data={editItem} fields={[{key:"name",label:"Company Name"},{key:"specialty",label:"Specialty"},{key:"contact",label:"Contact"},{key:"phone",label:"Phone"},{key:"email",label:"Email"},{key:"status",label:"Status",options:["Active","Inactive"]},{key:"rating",label:"Rating (0-5)"}]} onSave={saveEdit} onClose={() => setEditItem(null)} />}
      {deleteItem && <ConfirmDel name={deleteItem.name} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 18 }}><Btn onClick={() => setShowForm(v => !v)}>+ Add Vendor</Btn></div>
      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 18 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>Register Vendor</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label="Company Name *" value={form.name} onChange={f("name")} /><Input label="Specialty" value={form.specialty} onChange={f("specialty")} />
            <Input label="Contact" value={form.contact} onChange={f("contact")} /><Input label="Phone" value={form.phone} onChange={f("phone")} /><Input label="Email" value={form.email} onChange={f("email")} />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}><Btn onClick={submit} disabled={saving}>{saving?"Saving...":"Register"}</Btn><Btn variant="secondary" onClick={() => setShowForm(false)}>Cancel</Btn></div>
        </div>
      )}
      {loading ? <Spinner /> : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 14 }}>
          {vendors.map(v => (
            <div key={v.id} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
                <div><div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{v.name}</div><div style={{ fontSize: 11, color: C.muted }}>{v.specialty}</div></div>
                <Badge label={v.status} color={statusColor(v.status)} />
              </div>
              <Stars rating={v.rating} />
              <div style={{ marginTop: 12, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12 }}>
                {[["Contact",v.contact],["Phone",v.phone],["Email",v.email]].map(([lbl,val]) => (
  <div key={lbl}><div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase" }}>{lbl}</div><div style={{ color: C.subtle, marginTop: 2 }}>{val||"—"}</div></div>
))}
<div><div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase" }}>Open Orders</div><div style={{ color: (v.open_orders||0)>0?C.accent:C.subtle, marginTop: 2, fontWeight: (v.open_orders||0)>0?700:400 }}>{v.open_orders||0}</div></div>
              </div>
             <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
  <button onClick={() => setSelectedVendor(v)} style={{ flex: 1, background: C.blue+"22", color: C.blue, border: `1px solid ${C.blue}44`, borderRadius: 6, padding: "7px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>📋 Work Orders</button>
  {isAdmin && <><Btn small onClick={() => setEditItem(v)} color={C.accent}>Edit</Btn><Btn small variant="danger" onClick={() => setDeleteItem(v)}>Delete</Btn></>}
</div>
            </div>
          ))}
          {vendors.length===0 && <div style={{ color: C.muted, fontSize: 13 }}>No vendors yet.</div>}
        </div>
      )}
    </div>
  );
}

function PMUpload({ assets, onAssetsImported, onWorkOrdersGenerated }) {
  const [generating, setGenerating] = useState(false); const [uploading, setUploading] = useState(false);
  const [error, setError] = useState(null); const [success, setSuccess] = useState(null);
  const pmDueCount = assets.filter(a => { if (!a.pm_frequency) return false; if (!a.last_pm_date) return true; const now=new Date(); const last=new Date(a.last_pm_date); return (now.getFullYear()-last.getFullYear())*12+(now.getMonth()-last.getMonth())>=a.pm_frequency; }).length;
  const generatePMWorkOrders = async () => {
    setGenerating(true); setError(null);
    const now=new Date(); const due=assets.filter(a => { if (!a.pm_frequency) return false; if (!a.last_pm_date) return true; const last=new Date(a.last_pm_date); return (now.getFullYear()-last.getFullYear())*12+(now.getMonth()-last.getMonth())>=a.pm_frequency; });
    if (!due.length) { setSuccess("No assets due for PM!"); setGenerating(false); return; }
    const dueDate=new Date(now.getFullYear(),now.getMonth()+1,0).toISOString().split("T")[0];
    const newWOs=due.map(a => ({ id: uid("WO"), title: `PM - ${a.name}`, asset: a.name, priority: "Medium", status: "Open", assignee: null, start_date: TODAY, due: dueDate, vendor: null }));
    const { error: err }=await supabase.from("work_orders").insert(newWOs);
    if (err) { setError(err.message); } else { await supabase.from("assets").update({ last_pm_date: TODAY }).in("id",due.map(a => a.id)); setSuccess(`Generated ${newWOs.length} PM work orders!`); onWorkOrdersGenerated(newWOs); }
    setGenerating(false);
  };
  const handleImport = (e) => {
    const file=e.target.files[0]; if (!file) return;
    setUploading(true); setError(null);
    const reader=new FileReader();
    reader.onload=async (evt) => {
      try {
        const data=XLSX.utils.sheet_to_json(XLSX.read(evt.target.result,{type:"binary"}).Sheets[Object.keys(XLSX.read(evt.target.result,{type:"binary"}).Sheets)[0]]);
        const records=data.map(row => ({ id: uid("AST"), name: row["Asset Name"]||row["name"]||"", category: row["Category"]||"", location: row["Site"]||"", value: row["Value"]?String(row["Value"]):"", status: "Operational", last_service: TODAY, next_service: null, pm_frequency: parseInt(row["PM Frequency (months)"]||1), pm_task: row["PM Task"]||"Scheduled Maintenance", last_pm_date: null })).filter(r => r.name);
        const { error: err }=await supabase.from("assets").insert(records);
        if (err) { setError(err.message); } else { setSuccess(`Imported ${records.length} assets!`); onAssetsImported(records); }
      } catch { setError("Failed to parse file."); }
      setUploading(false);
    };
    reader.readAsBinaryString(file);
  };
  return (
    <div>
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <OkBanner msg={success} onDismiss={() => setSuccess(null)} />
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 24, marginBottom: 20 }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 6 }}>Generate PM Work Orders</div>
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 20 }}>
          {[["📋",pmDueCount,"Assets due for PM",C.accent],["🏭",assets.length,"Total assets",C.blue],["📅",new Date().toLocaleString("default",{month:"long",year:"numeric"}),"Current month",C.green]].map(([icon,val,label,color]) => (
            <div key={label} style={{ background: C.surface, borderRadius: 8, padding: "12px 20px", borderLeft: `3px solid ${color}` }}><div style={{ fontSize: 22, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{val}</div><div style={{ fontSize: 12, color: C.muted }}>{label}</div></div>
          ))}
        </div>
        <Btn onClick={generatePMWorkOrders} disabled={generating||pmDueCount===0}>{generating?"Generating...":`Generate ${pmDueCount} PM Work Orders`}</Btn>
      </div>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 24 }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 6 }}>Import Assets from Excel</div>
        <div style={{ fontSize: 13, color: C.subtle, marginBottom: 16 }}>Columns: Asset Name, Site, Category, Value, PM Frequency (months), PM Task</div>
        <label style={{ display: "inline-block", background: C.accent, color: "#fff", borderRadius: 6, padding: "10px 20px", fontSize: 13, fontWeight: 700, cursor: uploading?"not-allowed":"pointer", opacity: uploading?0.7:1 }}>
          {uploading?"Importing...":"📂 Choose Excel File"}
          <input type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display: "none" }} disabled={uploading} />
        </label>
      </div>
    </div>
  );
}

function Overview({ workOrders, assets, vendors, breakdownCount }) {
  const open=workOrders.filter(w => w.status!=="Completed").length;
  const critical=workOrders.filter(w => w.priority==="Critical").length;
  const opAssets=assets.filter(a => a.status==="Operational").length;
  const downAssets=assets.filter(a => a.status==="Under Maintenance").length;
  const activeVendors=vendors.filter(v => v.status==="Active").length;
  const overdue=workOrders.filter(w => w.due&&w.due<=TODAY&&w.status!=="Completed").length;
  const pmDue=assets.filter(a => { if (!a.pm_frequency) return false; if (!a.last_pm_date) return true; const now=new Date(); const last=new Date(a.last_pm_date); return (now.getFullYear()-last.getFullYear())*12+(now.getMonth()-last.getMonth())>=a.pm_frequency; }).length;
  return (
    <div>
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        <StatCard icon="🔧" label="Open Work Orders" value={open} sub={`${critical} critical`} color={C.accent} />
        <StatCard icon="🏭" label="Operational Assets" value={`${opAssets}/${assets.length}`} sub="fleet status" color={C.green} />
        <StatCard icon="🚨" label="Assets Down" value={downAssets} sub="under maintenance" color={C.red} />
        <StatCard icon="⚠️" label="Overdue / At Risk" value={overdue} sub="past due date" color={C.yellow} />
        <StatCard icon="📋" label="PM Due This Month" value={pmDue} sub="preventive maintenance" color={C.blue} />
        <StatCard icon="🤝" label="Active Vendors" value={activeVendors} sub="contractors on file" color={C.purple} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 16 }}>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
          <div style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>Recent Work Orders</div>
          {workOrders.slice(0,5).map(wo => (
            <div key={wo.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${C.border}22`, flexWrap: "wrap", gap: 6 }}>
              <div><div style={{ fontSize: 13, color: C.text, fontWeight: 600 }}>{wo.title}</div><div style={{ fontSize: 11, color: C.muted }}>{wo.asset}</div></div>
              <div style={{ display: "flex", gap: 6 }}><Badge label={wo.priority} color={priorityColor(wo.priority)} /><Badge label={wo.status} color={statusColor(wo.status)} /></div>
            </div>
          ))}
          {workOrders.length===0 && <div style={{ color: C.muted, fontSize: 13 }}>No work orders yet.</div>}
        </div>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
          <div style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>Assets by Site</div>
          {SITES.filter(s => s!=="— Select Site —").map(site => {
            const count=assets.filter(a => a.location===site).length; if (!count) return null;
            const op=assets.filter(a => a.location===site&&a.status==="Operational").length;
            const down=assets.filter(a => a.location===site&&a.status==="Under Maintenance").length;
            return <div key={site} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <span style={{ fontSize: 13, color: C.subtle }}>{site}</span>
              <div style={{ display: "flex", gap: 8 }}>
                <span style={{ fontSize: 11, color: C.green }}>{op} ok</span>
                {down>0 && <span style={{ fontSize: 11, color: C.red }}>{down} down</span>}
                <span style={{ fontSize: 13, color: C.text, fontWeight: 600 }}>{count}</span>
              </div>
            </div>;
          })}
        </div>
      </div>
    </div>
  );
}

// ─── USER MANAGEMENT (Admin only) ────────────────────────────────────────────
function UserManagement() {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [form, setForm] = useState({ email: "", name: "", role: "operations", site: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));

  useEffect(() => { loadUsers(); }, []);
  const loadUsers = async () => {
    setLoading(true);
    const { data } = await supabase.from("user_roles").select("*").order("name");
    setUsers(data || []);
    setLoading(false);
  };

  const submit = async () => {
    if (!form.email || !form.name) { setError("Email and name are required."); return; }
    setSaving(true); setError(null);
    const record = { id: uid("USR"), email: form.email, name: form.name, role: form.role, site: form.site || null };
    const { error: err } = await supabase.from("user_roles").upsert([record], { onConflict: "email" });
    if (err) { setError(err.message); } else { setSuccess("User added/updated!"); setForm({ email: "", name: "", role: "operations", site: "" }); setShowForm(false); loadUsers(); }
    setSaving(false);
  };

  const deleteUser = async (id) => {
    await supabase.from("user_roles").delete().eq("id", id);
    setUsers(prev => prev.filter(u => u.id !== id));
  };

  const roleColor = (r) => ({ admin: C.accent, maintenance: C.blue, operations: C.green }[r] || C.muted);
  const roleIcon = (r) => ({ admin: "★", maintenance: "🔧", operations: "🏭" }[r] || "👤");

  return (
    <div>
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <OkBanner msg={success} onDismiss={() => setSuccess(null)} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 15, fontWeight: 700, color: C.text, marginBottom: 4 }}>User Management</div>
          <div style={{ fontSize: 13, color: C.muted }}>Manage who can access the app and their roles.</div>
        </div>
        <Btn onClick={() => setShowForm(v => !v)}>+ Add User</Btn>
      </div>

      {/* Role descriptions */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12, marginBottom: 24 }}>
        {[["🏭", "Operations", "operations", "Can report breakdowns and view equipment status", C.green],
          ["🔧", "Maintenance", "maintenance", "Can resolve breakdowns, run checklists, add logs", C.blue],
          ["★", "Admin", "admin", "Full access — edit, delete, manage users", C.accent]].map(([icon, title, role, desc, color]) => (
          <div key={role} style={{ background: C.card, border: `1px solid ${color}44`, borderRadius: 8, padding: 14 }}>
            <div style={{ fontSize: 18, marginBottom: 6 }}>{icon}</div>
            <div style={{ fontSize: 13, fontWeight: 700, color }}{...{}}>{title}</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{desc}</div>
          </div>
        ))}
      </div>

      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 20 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>Add / Update User</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label="Email *" value={form.email} onChange={f("email")} type="email" />
            <Input label="Full Name *" value={form.name} onChange={f("name")} />
            <Sel label="Role" value={form.role} onChange={f("role")} options={["operations", "maintenance", "admin"]} />
            <Sel label="Default Site" value={form.site || "— Select Site —"} onChange={f("site")} options={["— Select Site —", ...SITES.filter(s => s !== "— Select Site —")]} />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
            <Btn onClick={submit} disabled={saving}>{saving ? "Saving..." : "Save User"}</Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>Cancel</Btn>
          </div>
        </div>
      )}

      {loading ? <Spinner /> : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 12 }}>
          {users.map(u => (
            <div key={u.id} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 18 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{u.name}</div>
                  <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>{u.email}</div>
                </div>
                <span style={{ background: roleColor(u.role)+"22", color: roleColor(u.role), border: `1px solid ${roleColor(u.role)}44`, borderRadius: 4, padding: "3px 10px", fontSize: 12, fontWeight: 700, textTransform: "uppercase" }}>
                  {roleIcon(u.role)} {u.role}
                </span>
              </div>
              {u.site && <div style={{ fontSize: 12, color: C.muted, marginBottom: 12 }}>📍 Default site: {u.site}</div>}
              <Btn small variant="danger" onClick={() => deleteUser(u.id)}>Remove</Btn>
            </div>
          ))}
          {users.length === 0 && <div style={{ color: C.muted, fontSize: 13 }}>No users registered yet.</div>}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState(null); const [authLoading, setAuthLoading] = useState(true);
  const [userRole, setUserRole] = useState({ role: "operations", name: "", site: "" });
  const [tab, setTab] = useState("Overview");
  const [workOrders, setWorkOrders] = useState([]); const [assets, setAssets] = useState([]); const [vendors, setVendors] = useState([]);
  const [loading, setLoading] = useState({ workOrders: true, assets: true, vendors: true });
  const [globalError, setGlobalError] = useState(null);
  const isAdmin = session?.user?.email === ADMIN_EMAIL || userRole.role === "admin";

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => { setSession(session); setAuthLoading(false); });
    supabase.auth.onAuthStateChange((_e, session) => { setSession(session); setAuthLoading(false); });
  }, []);

  useEffect(() => {
    if (session) {
      supabase.from("user_roles").select("*").eq("email", session.user.email).single().then(({ data }) => {
        if (data) setUserRole(data);
        else setUserRole({ role: session.user.email === ADMIN_EMAIL ? "admin" : "operations", name: session.user.email, site: "" });
      });
    }
  }, [session]);

  const load = useCallback(async () => {
    setLoading({ workOrders: true, assets: true, vendors: true });
    const [woRes, astRes, vndRes] = await Promise.all([
      supabase.from("work_orders").select("*").order("due", { ascending: true }),
      supabase.from("assets").select("*").order("name", { ascending: true }),
      supabase.from("vendors").select("*").order("name", { ascending: true }),
    ]);
    if (woRes.error||astRes.error||vndRes.error) { setGlobalError("Failed to load data."); }
    else { setWorkOrders(woRes.data||[]); setAssets(astRes.data||[]); setVendors(vndRes.data||[]); }
    setLoading({ workOrders: false, assets: false, vendors: false });
  }, []);

  useEffect(() => { if (session) load(); }, [session, load]);
  const signOut = async () => { await supabase.auth.signOut(); setSession(null); };

  if (authLoading) return <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", color: C.muted }}>Loading...</div>;
  if (!session) return <LoginScreen />;

  const roleColor = { admin: C.accent, maintenance: C.blue, operations: C.green }[userRole.role] || C.muted;
  const roleIcon = { admin: "★", maintenance: "🔧", operations: "🏭" }[userRole.role] || "👤";

  // Tabs based on role
  const tabs = [
    "Overview",
    "Breakdowns",
    ...(userRole.role !== "operations" ? ["Work Orders", "Assets", "Vendors", "PM Planner"] : []),
    ...(isAdmin ? ["Users"] : []),
  ];

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'DM Sans', sans-serif", color: C.text }}>
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "0 16px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", minHeight: 58, flexWrap: "wrap", gap: 8, padding: "8px 0" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ background: C.accent, borderRadius: 8, width: 34, height: 34, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>🏭</div>
            <div>
              <div style={{ fontFamily: "monospace", fontSize: 16, letterSpacing: 2, color: C.text, fontWeight: 800 }}>FACILITY COMMAND</div>
              <div style={{ fontSize: 9, color: C.muted }}>INDUSTRIAL WAREHOUSE MANAGEMENT</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <button onClick={load} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 10px", color: C.muted, cursor: "pointer", fontSize: 12 }}>Refresh</button>
            <span style={{ background: roleColor+"22", color: roleColor, border: `1px solid ${roleColor}44`, borderRadius: 4, padding: "3px 8px", fontSize: 11, fontWeight: 700, textTransform: "uppercase" }}>{roleIcon} {userRole.role}</span>
            <div style={{ fontSize: 11, color: C.muted, maxWidth: 130, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{userRole.name || session.user.email}</div>
            <button onClick={signOut} style={{ background: "transparent", border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 10px", color: C.muted, cursor: "pointer", fontSize: 12 }}>Sign Out</button>
          </div>
        </div>
        <div style={{ display: "flex", overflowX: "auto" }}>
          {tabs.map(t => <button key={t} onClick={() => setTab(t)} style={{ background: "transparent", border: "none", padding: "10px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap", color: tab===t?C.accent:C.muted, borderBottom: `2px solid ${tab===t?C.accent:"transparent"}` }}>{t}</button>)}
        </div>
      </div>
      <div style={{ padding: "20px 16px", maxWidth: 1280, margin: "0 auto" }}>
        <ErrBanner msg={globalError} onDismiss={() => setGlobalError(null)} />
        {tab==="Overview" && <Overview workOrders={workOrders} assets={assets} vendors={vendors} />}
        {tab==="Breakdowns" && <Breakdowns userRole={userRole} assets={assets} setAssets={setAssets} vendors={vendors} workOrders={workOrders} setWorkOrders={setWorkOrders} />}
        {tab==="Work Orders" && <WorkOrders workOrders={workOrders} setWorkOrders={setWorkOrders} loading={loading.workOrders} onAdd={r => setWorkOrders(p => [r,...p])} isAdmin={isAdmin} vendors={vendors} />}
        {tab==="Assets" && <Assets assets={assets} setAssets={setAssets} loading={loading.assets} onAdd={r => setAssets(p => [r,...p])} isAdmin={isAdmin} vendors={vendors} />}
        {tab==="Vendors" && <Vendors vendors={vendors} setVendors={setVendors} loading={loading.vendors} onAdd={r => setVendors(p => [r,...p])} isAdmin={isAdmin} />}
        {tab==="PM Planner" && <PMUpload assets={assets} onAssetsImported={r => setAssets(p => [...p,...r])} onWorkOrdersGenerated={r => setWorkOrders(p => [...r,...p])} />}
        {tab==="Users" && isAdmin && <UserManagement />}
      </div>
    </div>
  );
}


