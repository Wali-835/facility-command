// v3 - fully bilingual
import { useState, useEffect, useCallback } from "react";
import { t } from "./i18n.js";
import * as XLSX from "xlsx";
import QRCode from "qrcode";
import jsPDF from "jspdf";
import { applyPlugin } from "jspdf-autotable";
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
  const date = new Date(d.endsWith("Z") ? d : d + "Z");
  return date.toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit", timeZone: "Africa/Cairo" });
};
const minutesBetween = (a, b) => {
  if (!a || !b) return null;
  const dateA = new Date(a.endsWith("Z") ? a : a + "Z");
  const dateB = new Date(b.endsWith("Z") ? b : b + "Z");
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

const Spinner = ({ lang }) => <div style={{ textAlign: "center", padding: 48, color: C.muted, fontSize: 13 }}>{t(lang||"en","loading")}</div>;

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

const EditModal = ({ title, fields, data, onSave, onClose, lang }) => {
  const [form, setForm] = useState({ ...data });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000aa", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 24, width: "100%", maxWidth: 560, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: C.accent, marginBottom: 20, textTransform: "uppercase" }}>{t(lang,"edit")} {title}</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
          {fields.map(({ key, label, type, options }) =>
            options ? <Sel key={key} label={label} value={form[key]||""} onChange={f(key)} options={options} />
              : <Input key={key} label={label} value={form[key]||""} onChange={f(key)} type={type||"text"} />
          )}
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
          <Btn onClick={() => onSave(form)}>{t(lang,"save")}</Btn>
          <Btn variant="secondary" onClick={onClose}>{t(lang,"cancel")}</Btn>
        </div>
      </div>
    </div>
  );
};

const ConfirmDel = ({ name, onConfirm, onClose, lang }) => (
  <div style={{ position: "fixed", inset: 0, background: "#000000aa", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}>
    <div style={{ background: C.card, border: `1px solid ${C.red}44`, borderRadius: 12, padding: 24, width: "100%", maxWidth: 400 }}>
      <div style={{ fontSize: 15, fontWeight: 700, color: C.red, marginBottom: 12 }}>{t(lang,"confirmDelete")}</div>
      <div style={{ fontSize: 13, color: C.subtle, marginBottom: 20 }}>{t(lang,"delete")} <strong style={{ color: C.text }}>{name}</strong>? {t(lang,"deleteConfirmMsg")}</div>
      <div style={{ display: "flex", gap: 8 }}>
        <Btn variant="danger" onClick={onConfirm}>{t(lang,"delete")}</Btn>
        <Btn variant="secondary" onClick={onClose}>{t(lang,"cancel")}</Btn>
      </div>
    </div>
  </div>
);

// ─── BREAKDOWN REPORT MODAL ───────────────────────────────────────────────────
function BreakdownReportModal({ asset, userRole, onClose, onReported, lang }) {
  const [form, setForm] = useState({ description: "", severity: "High", reported_by: userRole.name || "" });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));

  const submit = async () => {
    if (!form.description) { setError(t(lang,"describeIssue")); return; }
    if (!form.reported_by) { setError(t(lang,"yourName")); return; }
    setSaving(true); setError(null);
    const now = new Date().toISOString();
    const record = { id: uid("BRK"), asset_id: asset.id, asset_name: asset.name, site: asset.location, reported_by: form.reported_by, reported_at: now, downtime_start: now, description: form.description, severity: form.severity, status: "Open" };
    const { error: err } = await supabase.from("breakdown_reports").insert([record]);
    if (err) { setError(err.message); setSaving(false); return; }
    await supabase.from("assets").update({ status: "Under Maintenance" }).eq("id", asset.id);
    onReported(record);
    onClose();
    setSaving(false);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2000, padding: 16 }}>
      <div style={{ background: C.card, border: `2px solid ${C.red}44`, borderRadius: 12, width: "100%", maxWidth: 500 }}>
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, background: C.red+"11" }}>
          <div style={{ fontSize: 17, fontWeight: 700, color: C.red }}>🚨 {t(lang,"reportBreakdown")}</div>
          <div style={{ fontSize: 13, color: C.muted, marginTop: 4 }}>{asset.name} · {asset.location}</div>
        </div>
        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Input label={t(lang,"yourName")} value={form.reported_by} onChange={f("reported_by")} />
            <Sel label={t(lang,"severity")} value={form.severity} onChange={f("severity")} options={["Critical","High","Medium","Low"]} />
            <Textarea label={t(lang,"describeIssue")} value={form.description} onChange={f("description")} />
            <div style={{ background: C.yellow+"11", border: `1px solid ${C.yellow}44`, borderRadius: 8, padding: 12, fontSize: 12, color: C.yellow }}>
              {t(lang,"downtimeStartsNow")}: <strong>{new Date().toLocaleString("en-GB")}</strong>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
            <Btn onClick={submit} disabled={saving} color={C.red}>{saving ? t(lang,"reporting") : `🚨 ${t(lang,"reportBreakdown")}`}</Btn>
            <Btn variant="secondary" onClick={onClose}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── BREAKDOWN RESOLUTION MODAL ───────────────────────────────────────────────
function BreakdownResolveModal({ breakdown, userRole, vendors, onClose, onResolved, lang }) {
  const [form, setForm] = useState({ resolved_by: userRole.name || "", maintenance_notes: "", vendor: "" });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const vendorOptions = ["— None —", ...vendors.filter(v => v.status === "Active").map(v => v.name)];
  const downtimeMins = minutesBetween(breakdown.downtime_start, new Date().toISOString());

  const submit = async () => {
    if (!form.resolved_by) { setError(t(lang,"yourName")); return; }
    if (!form.maintenance_notes) { setError(t(lang,"maintenanceNotes")); return; }
    setSaving(true); setError(null);
    const now = new Date().toISOString();
    const hours = minutesBetween(breakdown.downtime_start, now);
    await supabase.from("breakdown_reports").update({ status: "Resolved", resolved_by: form.resolved_by, resolved_at: now, downtime_end: now, downtime_hours: hours, maintenance_notes: form.maintenance_notes }).eq("id", breakdown.id);
    await supabase.from("assets").update({ status: "Operational" }).eq("id", breakdown.asset_id);
    const logRecord = { id: uid("LOG"), asset_id: breakdown.asset_id, asset_name: breakdown.asset_name, log_type: "Corrective Repair", title: `Breakdown Repair — ${breakdown.severity} severity`, description: `BREAKDOWN REPORTED BY: ${breakdown.reported_by}\n\nISSUE: ${breakdown.description}\n\nMAINTENANCE NOTES: ${form.maintenance_notes}`, performed_by: form.resolved_by, vendor: form.vendor === "— None —" ? null : form.vendor || null, start_date: breakdown.downtime_start ? breakdown.downtime_start.split("T")[0] : TODAY, end_date: TODAY, cost: null, status: "Completed", downtime_start: breakdown.downtime_start ? breakdown.downtime_start.split("T")[0] : null, downtime_end: TODAY, downtime_hours: hours };
    await supabase.from("maintenance_logs").insert([logRecord]);
    const { data: openWOs } = await supabase.from("work_orders").select("id").eq("asset", breakdown.asset_name).in("status", ["Open","In Progress","Pending"]);
    if (openWOs?.length) await supabase.from("work_orders").update({ status: "Completed" }).in("id", openWOs.map(w => w.id));
    onResolved({ ...breakdown, status: "Resolved", downtime_end: now, downtime_hours: hours });
    onClose();
    setSaving(false);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2000, padding: 16 }}>
      <div style={{ background: C.card, border: `2px solid ${C.green}44`, borderRadius: 12, width: "100%", maxWidth: 560, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, background: C.green+"11" }}>
          <div style={{ fontSize: 17, fontWeight: 700, color: C.green }}>✅ {t(lang,"markResolved")}</div>
          <div style={{ fontSize: 13, color: C.muted, marginTop: 4 }}>{breakdown.asset_name} · {breakdown.site}</div>
        </div>
        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <div style={{ background: C.red+"11", border: `1px solid ${C.red}33`, borderRadius: 8, padding: 14, marginBottom: 18 }}>
            <div style={{ fontSize: 12, color: C.muted, marginBottom: 8, textTransform: "uppercase", fontWeight: 600 }}>{t(lang,"breakdownDetails")}</div>
            <div style={{ fontSize: 13, color: C.text, marginBottom: 6 }}><strong>{t(lang,"reportedBy")}:</strong> {breakdown.reported_by}</div>
            <div style={{ fontSize: 13, color: C.text, marginBottom: 6 }}><strong>{t(lang,"reportedAt")}:</strong> {fmtDateTime(breakdown.reported_at)}</div>
            <div style={{ fontSize: 13, color: C.text, marginBottom: 6 }}><strong>{t(lang,"issue")}:</strong> {breakdown.description}</div>
            <div style={{ fontSize: 13, color: C.yellow, fontWeight: 700 }}>⏱ {t(lang,"downtime")}: {formatDowntime(downtimeMins)}</div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Input label={t(lang,"yourName")} value={form.resolved_by} onChange={f("resolved_by")} />
            <Sel label={t(lang,"vendorContractor")} value={form.vendor} onChange={f("vendor")} options={vendorOptions} />
            <Textarea label={t(lang,"maintenanceNotes")} value={form.maintenance_notes} onChange={f("maintenance_notes")} />
            <div style={{ background: C.green+"11", border: `1px solid ${C.green}44`, borderRadius: 8, padding: 12, fontSize: 12, color: C.green }}>
              ✅ {t(lang,"operationalStatus")} · {new Date().toLocaleString("en-GB")}
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
            <Btn onClick={submit} disabled={saving} color={C.green}>{saving ? t(lang,"resolving") : `✅ ${t(lang,"markResolved")}`}</Btn>
            <Btn variant="secondary" onClick={onClose}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── BREAKDOWNS TAB ───────────────────────────────────────────────────────────
function Breakdowns({ userRole, assets, setAssets, vendors, workOrders, setWorkOrders, lang }) {
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

  const onReported = async (record) => {
    setBreakdowns(prev => [record, ...prev]);
    setAssets(prev => prev.map(a => a.id === record.asset_id ? { ...a, status: "Under Maintenance" } : a));
    try {
      await fetch("https://evwsdzqgvrwbjusjmrdc.supabase.co/functions/v1/notify-breakdown", {
        method: "POST",
        headers: { "Content-Type": "application/json", "Authorization": `Bearer ${import.meta.env.VITE_SUPABASE_ANON_KEY}` },
        body: JSON.stringify({ breakdown: record, type: "reported" }),
      });
    } catch (e) { console.error("Email error:", e); }
    setSuccess(t(lang,"breakdownReported"));
  };

  const onResolved = async (updated) => {
    setBreakdowns(prev => prev.map(b => b.id === updated.id ? updated : b));
    setAssets(prev => prev.map(a => a.id === updated.asset_id ? { ...a, status: "Operational" } : a));
    try {
      await fetch("https://evwsdzqgvrwbjusjmrdc.supabase.co/functions/v1/notify-breakdown", {
        method: "POST",
        headers: { "Content-Type": "application/json", "Authorization": `Bearer ${import.meta.env.VITE_SUPABASE_ANON_KEY}` },
        body: JSON.stringify({ breakdown: updated, type: "resolved" }),
      });
    } catch (e) { console.error("Email error:", e); }
    setSuccess(t(lang,"breakdownResolved"));
  };

  const acknowledge = async (b) => {
    const now = new Date().toISOString();
    const { error: err } = await supabase.from("breakdown_reports").update({ status: "Acknowledged", acknowledged_by: userRole.name || "", acknowledged_at: now }).eq("id", b.id);
    if (err) { setError(err.message); return; }
    setBreakdowns(prev => prev.map(x => x.id === b.id ? { ...x, status: "Acknowledged", acknowledged_by: userRole.name, acknowledged_at: now } : x));
  };

  return (
    <div>
      {showReportForm && selectedAsset && (
        <BreakdownReportModal asset={selectedAsset} userRole={userRole} lang={lang} onClose={() => { setShowReportForm(false); setSelectedAsset(null); }} onReported={onReported} />
      )}
      {resolveItem && (
        <BreakdownResolveModal breakdown={resolveItem} userRole={userRole} vendors={vendors} lang={lang} onClose={() => setResolveItem(null)} onResolved={onResolved} />
      )}
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <OkBanner msg={success} onDismiss={() => setSuccess(null)} />

      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        <StatCard icon="🚨" label={t(lang,"openBreakdowns")} value={openCount} sub={t(lang,"needsAttention")} color={C.red} />
        <StatCard icon="👁" label={t(lang,"acknowledged")} value={acknowledgedCount} sub={t(lang,"beingHandled")} color={C.blue} />
        <StatCard icon="✅" label={t(lang,"resolved")} value={resolvedCount} sub={t(lang,"thisPeriod")} color={C.green} />
        <StatCard icon="⏱" label={t(lang,"totalDowntime")} value={formatDowntime(totalDowntimeMins)} sub={t(lang,"allBreakdowns")} color={C.yellow} />
        <StatCard icon="🏭" label={t(lang,"assetsDown")} value={assets.filter(a => a.status === "Under Maintenance").length} sub={t(lang,"currentlyOffline")} color={C.accent} />
      </div>

      {(userRole.role === "operations" || userRole.role === "admin") && (
        <div style={{ background: C.card, border: `1px solid ${C.red}33`, borderRadius: 10, padding: 20, marginBottom: 20 }}>
          <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 8 }}>{t(lang,"reportEquipmentBreakdown")}</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
            <select onChange={e => setSelectedAsset(assets.find(a => a.id === e.target.value) || null)}
              style={{ flex: 1, minWidth: 200, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14 }}>
              <option value="">{t(lang,"selectEquipment")}</option>
              {assets.filter(a => a.status !== "Under Maintenance").map(a => <option key={a.id} value={a.id}>{a.name} ({a.location})</option>)}
            </select>
            <Btn onClick={() => { if (!selectedAsset) { setError(t(lang,"selectEquipment")); return; } setShowReportForm(true); }} color={C.red}>
              🚨 {t(lang,"reportBreakdown")}
            </Btn>
          </div>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", gap: 6 }}>
          {["All","Open","Acknowledged","Resolved"].map(s => (
            <button key={s} onClick={() => setFilter(s)} style={{ background: filter===s?C.accent:C.card, color: filter===s?"#fff":C.muted, border: `1px solid ${filter===s?C.accent:C.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>
              {s === "All" ? t(lang,"all") : s === "Open" ? t(lang,"open") : s === "Acknowledged" ? t(lang,"acknowledged") : t(lang,"resolved")}
            </button>
          ))}
        </div>
        <button onClick={loadBreakdowns} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "6px 12px", color: C.muted, cursor: "pointer", fontSize: 12 }}>↻ {t(lang,"refresh")}</button>
      </div>

      {loading ? <Spinner lang={lang} /> : filtered.length === 0 ? (
        <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>{t(lang,"noBreakdownsFound")}</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {filtered.map(b => {
            const isResolved = b.status === "Resolved";
            const isAcknowledged = b.status === "Acknowledged";
            const currentDowntimeMins = isResolved ? Math.round(b.downtime_hours || 0) : minutesBetween(b.downtime_start, new Date().toISOString()) || 0;
            return (
              <div key={b.id} style={{ background: C.card, border: `1px solid ${isResolved?C.green+"44":isAcknowledged?C.blue+"44":C.red+"44"}`, borderRadius: 10, padding: 20 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
                  <div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                      <span style={{ fontSize: 18 }}>{isResolved?"✅":isAcknowledged?"👁":"🚨"}</span>
                      <div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{b.asset_name}</div>
                    </div>
                    <div style={{ fontSize: 12, color: C.muted }}>{b.site} · {t(lang,"reportedBy")} {b.reported_by} · {fmtDateTime(b.reported_at)}</div>
                    {isAcknowledged && b.acknowledged_by && (
                      <div style={{ fontSize: 12, color: C.blue, marginTop: 4 }}>👁 {t(lang,"acknowledged")}: {b.acknowledged_by} · {fmtDateTime(b.acknowledged_at)}</div>
                    )}
                  </div>
                  <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                    <Badge label={b.severity} color={SEVERITY_COLORS[b.severity]||C.muted} />
                    <Badge label={b.status} color={statusColor(b.status)} />
                  </div>
                </div>
                <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 14 }}>
                  <div style={{ background: C.red+"11", border: `1px solid ${C.red}33`, borderRadius: 8, padding: "10px 16px", flex: "1 1 160px" }}>
                    <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", marginBottom: 4 }}>⛔ {t(lang,"downtimeStarted")}</div>
                    <div style={{ fontSize: 13, color: C.red, fontWeight: 700 }}>{fmtDateTime(b.downtime_start)}</div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", color: C.muted, fontSize: 20 }}>→</div>
                  <div style={{ background: isResolved?C.green+"11":C.yellow+"11", border: `1px solid ${isResolved?C.green+"33":C.yellow+"33"}`, borderRadius: 8, padding: "10px 16px", flex: "1 1 160px" }}>
                    <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", marginBottom: 4 }}>{isResolved?`✅ ${t(lang,"backToOperation")}`:t(lang,"stillDown")}</div>
                    <div style={{ fontSize: 13, color: isResolved?C.green:C.yellow, fontWeight: 700 }}>{isResolved?fmtDateTime(b.downtime_end):"..."}</div>
                  </div>
                  <div style={{ background: C.purple+"11", border: `1px solid ${C.purple}33`, borderRadius: 8, padding: "10px 16px" }}>
                    <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", marginBottom: 4 }}>⏱ {t(lang,"downtime")}</div>
                    <div style={{ fontSize: 13, color: C.purple, fontWeight: 700 }}>{formatDowntime(currentDowntimeMins)}</div>
                  </div>
                </div>
                <div style={{ background: C.surface, borderRadius: 8, padding: 12, marginBottom: 14, fontSize: 13, color: C.subtle }}>
                  <strong style={{ color: C.text }}>{t(lang,"issue")}:</strong> {b.description}
                </div>
                {isResolved && b.maintenance_notes && (
                  <div style={{ background: C.green+"11", border: `1px solid ${C.green}33`, borderRadius: 8, padding: 12, marginBottom: 14, fontSize: 13, color: C.subtle }}>
                    <strong style={{ color: C.green }}>✅ {t(lang,"resolvedBy")} {b.resolved_by}:</strong> {b.maintenance_notes}
                  </div>
                )}
                {!isResolved && (userRole.role === "maintenance" || userRole.role === "admin") && (
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    {b.status === "Open" && <Btn onClick={() => acknowledge(b)} color={C.blue}>{t(lang,"acknowledge")}</Btn>}
                    <Btn onClick={() => setResolveItem(b)} color={C.green}>{t(lang,"markResolved")}</Btn>
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
function ChecklistModal({ asset, workOrderId, onClose, lang }) {
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
    if (!executedBy) { setError(t(lang,"technicianName")); return; }
    const now = new Date();
    const execId = uid("EXC");
    const { data: checklists } = await supabase.from("checklists").select("id").limit(1);
    await supabase.from("checklist_executions").insert([{ id: execId, checklist_id: checklists[0].id, asset_id: asset.id, asset_name: asset.name, work_order_id: workOrderId||null, executed_by: executedBy, execution_date: TODAY, month: now.getMonth()+1, year: now.getFullYear(), status: "In Progress" }]);
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
    const description = `CIL Checklist completed by ${executedBy}\n\nResults: ${passCount} PASS · ${failCount} FAIL · ${naCount} N/A\n${defects ? `\nDefects:\n${defects}` : "\nNo defects found."}`;
    const logRecord = { id: uid("LOG"), asset_id: asset.id, asset_name: asset.name, log_type: "Preventive Maintenance", title: `CIL Checklist — ${new Date().toLocaleString("default", { month: "long", year: "numeric" })}`, description, performed_by: executedBy, vendor: null, start_date: TODAY, end_date: TODAY, cost: null, status: failCount > 0 ? "In Progress" : "Completed", downtime_start: null, downtime_end: null, downtime_hours: null };
    await supabase.from("maintenance_logs").insert([logRecord]);
    if (workOrderId) await supabase.from("work_orders").update({ status: "Completed" }).eq("id", workOrderId);
    const { data: openWOs } = await supabase.from("work_orders").select("id").eq("asset", asset.name).in("status", ["Open","In Progress","Pending"]).ilike("title","PM - %");
    if (openWOs?.length) await supabase.from("work_orders").update({ status: "Completed" }).in("id", openWOs.map(w => w.id));
    setSuccess(`${t(lang,"checklistCompleted")} ${failCount > 0 ? `⚠️ ${failCount} FAIL` : "✅"}`);
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
            <div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>{t(lang,"cilChecklist")} — {asset.name}</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{asset.category} · {asset.location} · {fmtDate(TODAY)}</div>
          </div>
          <button onClick={onClose} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, color: C.muted, cursor: "pointer", fontSize: 18, padding: "2px 10px" }}>✕</button>
        </div>
        <div style={{ padding: "14px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 16, flexWrap: "wrap", alignItems: "center" }}>
          <div style={{ flex: 1 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: C.muted, marginBottom: 6 }}>
              <span>{t(lang,"progress")}: {answeredCount}/{filteredItems.length}</span>
              <span style={{ color: failCount > 0 ? C.red : C.green }}>{failCount} FAIL · {passCount} PASS</span>
            </div>
            <div style={{ background: C.border, borderRadius: 4, height: 8 }}>
              <div style={{ background: failCount > 0 ? C.accent : C.green, width: `${filteredItems.length > 0 ? (answeredCount/filteredItems.length)*100 : 0}%`, height: 8, borderRadius: 4, transition: "width 0.3s" }} />
            </div>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            {["All","D","W","F","M"].map(f => (
              <button key={f} onClick={() => setFilterFreq(f)} style={{ background: filterFreq===f?(FREQ_COLORS[f]||C.accent):C.surface, color: filterFreq===f?"#fff":C.muted, border: `1px solid ${filterFreq===f?(FREQ_COLORS[f]||C.accent):C.border}`, borderRadius: 6, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                {f==="All"?t(lang,"all"):FREQ_LABELS[f]}
              </button>
            ))}
          </div>
        </div>
        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <OkBanner msg={success} onDismiss={() => setSuccess(null)} />
          {!executionId && (
            <div style={{ background: C.surface, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 20 }}>
              <div style={{ fontSize: 13, color: C.accent, fontWeight: 700, marginBottom: 12, textTransform: "uppercase" }}>{t(lang,"startChecklist")}</div>
              <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
                <div style={{ flex: 1, minWidth: 200 }}><Input label={t(lang,"technicianName")} value={executedBy} onChange={setExecutedBy} /></div>
                <Btn onClick={startExecution}>{t(lang,"start")}</Btn>
              </div>
            </div>
          )}
          {loading ? <Spinner lang={lang} /> : (
            <div>
              {["Cleaning","Safety","Inspection","Lubrication"].map(cat => {
                const catItems = filteredItems.filter(i => i.category===cat);
                if (!catItems.length) return null;
                return (
                  <div key={cat} style={{ marginBottom: 24 }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: C.text, marginBottom: 12, display: "flex", alignItems: "center", gap: 8 }}>
                      <span>{CAT_ICONS[cat]}</span> {cat} <span style={{ fontSize: 11, color: C.muted, fontWeight: 400 }}>({catItems.length})</span>
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
                                ) : <span style={{ fontSize: 12, color: C.muted }}>{t(lang,"startChecklist")}</span>}
                              </div>
                            </div>
                            {result==="FAIL" && executionId && (
                              <div style={{ marginTop: 10 }}>
                                <textarea value={resp?.notes||""} onChange={e => setResponse(item.id,"FAIL",e.target.value)} rows={2} style={{ width: "100%", background: C.card, border: `1px solid ${C.red}44`, borderRadius: 6, padding: "8px 10px", color: C.text, fontSize: 13, boxSizing: "border-box", resize: "vertical" }} />
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
                  {failCount > 0 && <div style={{ fontSize: 13, color: C.red }}>⚠️ {failCount} FAIL</div>}
                  <Btn onClick={complete} disabled={saving} color={C.green}>{saving ? t(lang,"completing") : `✓ ${t(lang,"completeChecklist")}`}</Btn>
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
function MaintenanceModal({ asset, onClose, isAdmin, vendors, lang }) {
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
    if (!form.title) { setError(t(lang,"title")); return; }
    setSaving(true); setError(null);
    const record = { id: uid("LOG"), asset_id: asset.id, asset_name: asset.name, log_type: form.log_type, title: form.title, description: form.description, performed_by: form.performed_by, vendor: form.vendor==="— None —"?null:form.vendor||null, start_date: form.start_date||null, end_date: form.end_date||null, cost: form.cost?parseFloat(form.cost):null, status: form.status, downtime_start: form.downtime_start||null, downtime_end: form.downtime_end||null, downtime_hours: (form.downtime_start && form.downtime_end) ? Math.round((new Date(form.downtime_end) - new Date(form.downtime_start)) / (1000 * 60 * 60)) : null };
    const { error: err } = await supabase.from("maintenance_logs").insert([record]);
    if (err) { setError(err.message); } else { setSuccess(t(lang,"saving")); setLogs(prev => [record,...prev]); setForm({ log_type: "Preventive Maintenance", title: "", description: "", performed_by: "", vendor: "", start_date: TODAY, end_date: "", cost: "", status: "Completed", downtime_start: "", downtime_end: "" }); setShowForm(false); }
    setSaving(false);
  };

  const submitPart = async (logId) => {
    if (!partForm.part_name) { setError(t(lang,"partName")); return; }
    setSaving(true); setError(null);
    const qty = parseFloat(partForm.quantity)||1;
    const unitCost = parseFloat(partForm.unit_cost)||0;
    const record = { id: uid("PRT"), log_id: logId, asset_id: asset.id, part_name: partForm.part_name, part_number: partForm.part_number, quantity: qty, unit_cost: unitCost, total_cost: qty*unitCost, supplier: partForm.supplier };
    const { error: err } = await supabase.from("spare_parts").insert([record]);
    if (err) { setError(err.message); } else { setSuccess("✓"); setParts(prev => ({ ...prev, [logId]: [...(prev[logId]||[]),record] })); setPartForm({ part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "" }); setShowPartForm(null); }
    setSaving(false);
  };

  const deletePart = async (partId, logId) => { await supabase.from("spare_parts").delete().eq("id", partId); setParts(prev => ({ ...prev, [logId]: prev[logId].filter(p => p.id!==partId) })); };
  const deleteLog = async (logId) => { await supabase.from("spare_parts").delete().eq("log_id", logId); await supabase.from("maintenance_logs").delete().eq("id", logId); setLogs(prev => prev.filter(l => l.id!==logId)); };
  const totalCost = logs.reduce((s,l) => s+(l.cost||0), 0);

  return (
    <>
      {showChecklist && <ChecklistModal asset={asset} lang={lang} onClose={() => setShowChecklist(false)} />}
      <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1000, padding: 16, overflowY: "auto" }}>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, width: "100%", maxWidth: 820, marginTop: 20, marginBottom: 20 }}>
          <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
            <div><div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>{asset.name}</div><div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{asset.category} · {asset.location}</div></div>
            <button onClick={onClose} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, color: C.muted, cursor: "pointer", fontSize: 18, padding: "2px 10px" }}>✕</button>
          </div>
          <div style={{ padding: "14px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 12, flexWrap: "wrap" }}>
            {[["📋",logs.length,t(lang,"totalLogs"),C.blue],["💰",fmt(totalCost),"Total Cost",C.accent],["🔧",logs[0]?.start_date?fmtDate(logs[0].start_date):t(lang,"never"),t(lang,"lastMaintenance"),C.green],["⚙️",asset.pm_frequency?`Every ${asset.pm_frequency} mo.`:"—",t(lang,"pmFrequencyLabel"),C.yellow]].map(([icon,val,label,color]) => (
              <div key={label} style={{ background: C.surface, borderRadius: 8, padding: "10px 16px", flex: "1 1 130px", borderLeft: `3px solid ${color}` }}>
                <div style={{ fontSize: 16 }}>{icon}</div><div style={{ fontSize: 18, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{val}</div><div style={{ fontSize: 11, color: C.muted }}>{label}</div>
              </div>
            ))}
          </div>
          <div style={{ padding: 24 }}>
            <ErrBanner msg={error} onDismiss={() => setError(null)} />
            <OkBanner msg={success} onDismiss={() => setSuccess(null)} />
            <div style={{ display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
              <button onClick={() => setShowChecklist(true)} style={{ background: C.blue+"22", color: C.blue, border: `1px solid ${C.blue}44`, borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>{t(lang,"runCILChecklist")}</button>
              <Btn onClick={() => setShowForm(v => !v)}>{t(lang,"addMaintenanceLog")}</Btn>
            </div>
            {showForm && (
              <div style={{ background: C.surface, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 20 }}>
                <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>{t(lang,"newMaintenanceLog")}</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
                  <Sel label={t(lang,"type")} value={form.log_type} onChange={f("log_type")} options={LOG_TYPES} />
                  <Input label={t(lang,"title")} value={form.title} onChange={f("title")} />
                  <Input label={t(lang,"performedBy")} value={form.performed_by} onChange={f("performed_by")} />
                  <Sel label={t(lang,"vendor")} value={form.vendor} onChange={f("vendor")} options={vendorOptions} />
                  <Input label={t(lang,"startDate")} value={form.start_date} onChange={f("start_date")} type="date" />
                  <Input label={t(lang,"endDate")} value={form.end_date} onChange={f("end_date")} type="date" />
                  <Input label={t(lang,"totalCost")} value={form.cost} onChange={f("cost")} type="number" />
                  <Sel label={t(lang,"status")} value={form.status} onChange={f("status")} options={LOG_STATUSES} />
                  <Input label={t(lang,"downtimeStart")} value={form.downtime_start} onChange={f("downtime_start")} type="date" />
                  <Input label={t(lang,"backToOperationLabel")} value={form.downtime_end} onChange={f("downtime_end")} type="date" />
                </div>
                <div style={{ marginTop: 12 }}><Textarea label={t(lang,"descriptionNotes")} value={form.description} onChange={f("description")} /></div>
                <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
                  <Btn onClick={submitLog} disabled={saving}>{saving?t(lang,"saving"):t(lang,"saveLog")}</Btn>
                  <Btn variant="secondary" onClick={() => setShowForm(false)}>{t(lang,"cancel")}</Btn>
                </div>
              </div>
            )}
            <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 14 }}>{t(lang,"maintenanceHistory")}</div>
            {loadingLogs ? <Spinner lang={lang} /> : logs.length===0 ? (
              <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>{t(lang,"noMaintenanceRecords")}</div>
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
                            <div><div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase" }}>⛔ {t(lang,"downtimeStart")}</div><div style={{ fontSize: 13, color: C.red, fontWeight: 700, marginTop: 2 }}>{fmtDate(log.downtime_start)||"—"}</div></div>
                            <div style={{ fontSize: 20, color: C.muted, alignSelf: "center" }}>→</div>
                            <div><div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase" }}>✅ {t(lang,"backToOperationLabel")}</div><div style={{ fontSize: 13, color: C.green, fontWeight: 700, marginTop: 2 }}>{fmtDate(log.downtime_end)||"—"}</div></div>
                            {log.downtime_hours && <div><div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase" }}>⏱ {t(lang,"totalDowntimeLabel")}</div><div style={{ fontSize: 13, color: C.yellow, fontWeight: 700, marginTop: 2 }}>{log.downtime_hours}h</div></div>}
                          </div>
                        )}
                        <div style={{ marginTop: 14 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                            <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>{t(lang,"spareParts")}</div>
                            {isAdmin && <Btn small onClick={() => setShowPartForm(showPartForm===log.id?null:log.id)}>{t(lang,"addPart")}</Btn>}
                          </div>
                          {showPartForm===log.id && (
                            <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 8, padding: 14, marginBottom: 12 }}>
                              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10 }}>
                                <Input label={t(lang,"partName")} value={partForm.part_name} onChange={pf("part_name")} />
                                <Input label={t(lang,"partNumber")} value={partForm.part_number} onChange={pf("part_number")} />
                                <Input label={t(lang,"quantity")} value={partForm.quantity} onChange={pf("quantity")} type="number" />
                                <Input label={t(lang,"unitCost")} value={partForm.unit_cost} onChange={pf("unit_cost")} type="number" />
                                <Input label={t(lang,"supplier")} value={partForm.supplier} onChange={pf("supplier")} />
                              </div>
                              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                                <Btn small onClick={() => submitPart(log.id)} disabled={saving}>{t(lang,"save")}</Btn>
                                <Btn small variant="secondary" onClick={() => setShowPartForm(null)}>{t(lang,"cancel")}</Btn>
                              </div>
                            </div>
                          )}
                          {parts[log.id]?.length > 0 ? (
                            <div style={{ overflowX: "auto" }}>
                              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                                <thead><tr style={{ borderBottom: `1px solid ${C.border}` }}>
                                  {[t(lang,"partName"),t(lang,"partNumber"),t(lang,"quantity"),"Unit","Total",t(lang,"supplier"),...(isAdmin?[""]:[])].map(h => <th key={h} style={{ textAlign: "left", padding: "6px 10px", color: C.muted, fontWeight: 600, fontSize: 11, textTransform: "uppercase" }}>{h}</th>)}
                                </tr></thead>
                                <tbody>
                                  {parts[log.id].map(part => (
                                    <tr key={part.id} style={{ borderBottom: `1px solid ${C.border}22` }}>
                                      <td style={{ padding: "8px 10px", color: C.text, fontWeight: 600 }}>{part.part_name}</td>
                                      <td style={{ padding: "8px 10px", color: C.subtle, fontFamily: "monospace", fontSize: 11 }}>{part.part_number||"—"}</td>
                                      <td style={{ padding: "8px 10px", color: C.subtle }}>{part.quantity}</td>
                                      <td style={{ padding: "8px 10px", color: C.subtle }}>{part.unit_cost?`$${part.unit_cost}`:"—"}</td>
                                      <td style={{ padding: "8px 10px", color: C.accent, fontWeight: 700 }}>{part.total_cost?`$${part.total_cost}`:"—"}</td>
                                      <td style={{ padding: "8px 10px", color: C.subtle }}>{part.supplier||"—"}</td>
                                      {isAdmin && <td style={{ padding: "8px 10px" }}><Btn small variant="danger" onClick={() => deletePart(part.id, log.id)}>{t(lang,"del")}</Btn></td>}
                                    </tr>
                                  ))}
                                  <tr style={{ borderTop: `1px solid ${C.border}` }}>
                                    <td colSpan={4} style={{ padding: "8px 10px", color: C.muted, fontSize: 11 }}>Total</td>
                                    <td style={{ padding: "8px 10px", color: C.accent, fontWeight: 700 }}>${parts[log.id].reduce((s,p) => s+(p.total_cost||0),0).toLocaleString()}</td>
                                    <td colSpan={isAdmin?2:1} />
                                  </tr>
                                </tbody>
                              </table>
                            </div>
                          ) : <div style={{ fontSize: 12, color: C.muted }}>{t(lang,"noSpareParts")}</div>}
                        </div>
                        {isAdmin && <div style={{ marginTop: 12, display: "flex", justifyContent: "flex-end" }}><Btn small variant="danger" onClick={() => deleteLog(log.id)}>{t(lang,"deleteLog")}</Btn></div>}
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

function LoginScreen({ lang }) {
  const [email, setEmail] = useState(""); const [password, setPassword] = useState("");
  const [error, setError] = useState(null); const [loading, setLoading] = useState(false);
  const signIn = async () => {
    if (!email || !password) { setError(t(lang,"enterEmail")); return; }
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
          <div style={{ fontFamily: "monospace", fontSize: 20, letterSpacing: 3, color: C.text, fontWeight: 800 }}>{t(lang,"appName")}</div>
          <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{t(lang,"appSub")}</div>
        </div>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 24 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.text, marginBottom: 20 }}>{t(lang,"qrSignIn")}</div>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <Input label={t(lang,"enterEmail")} value={email} onChange={setEmail} type="email" />
            <Input label={t(lang,"enterPassword")} value={password} onChange={setPassword} type="password" />
            <button onClick={signIn} disabled={loading} style={{ background: C.accent, color: "#fff", border: "none", borderRadius: 6, padding: "12px", fontSize: 15, fontWeight: 700, cursor: loading?"not-allowed":"pointer", opacity: loading?0.7:1 }}>
              {loading ? t(lang,"signingIn") : t(lang,"signIn")}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function WorkOrderPhotosModal({ workOrder, onClose, lang }) {
  const [photos, setPhotos] = useState([]);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);

  useEffect(() => { loadPhotos(); }, [workOrder.id]);

  const loadPhotos = async () => {
    const { data } = await supabase.storage.from("work-order-photos").list(workOrder.id);
    if (data) {
      const urls = data.map(file => ({ name: file.name, url: supabase.storage.from("work-order-photos").getPublicUrl(`${workOrder.id}/${file.name}`).data.publicUrl }));
      setPhotos(urls);
    }
  };

  const uploadPhoto = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    if (file.size > 5 * 1024 * 1024) { setError("Max 5MB"); return; }
    setUploading(true); setError(null);
    const fileName = `${Date.now()}-${file.name}`;
    const { error: err } = await supabase.storage.from("work-order-photos").upload(`${workOrder.id}/${fileName}`, file);
    if (err) { setError(err.message); } else { setSuccess("✓"); await loadPhotos(); }
    setUploading(false);
  };

  const deletePhoto = async (name) => {
    await supabase.storage.from("work-order-photos").remove([`${workOrder.id}/${name}`]);
    setPhotos(prev => prev.filter(p => p.name !== name));
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1000, padding: 16, overflowY: "auto" }}>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, width: "100%", maxWidth: 720, marginTop: 20, marginBottom: 20 }}>
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div><div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>📷 {t(lang,"photos")} — {workOrder.title}</div><div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{workOrder.asset}</div></div>
          <button onClick={onClose} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, color: C.muted, cursor: "pointer", fontSize: 18, padding: "2px 10px" }}>✕</button>
        </div>
        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <OkBanner msg={success} onDismiss={() => setSuccess(null)} />
          <div style={{ marginBottom: 24 }}>
            <label style={{ display: "inline-flex", alignItems: "center", gap: 8, background: C.accent, color: "#fff", borderRadius: 6, padding: "10px 20px", fontSize: 13, fontWeight: 700, cursor: uploading?"not-allowed":"pointer", opacity: uploading?0.7:1 }}>
              {uploading ? "⏳..." : "📷 Upload"}
              <input type="file" accept="image/*" onChange={uploadPhoto} style={{ display: "none" }} disabled={uploading} />
            </label>
          </div>
          {photos.length === 0 ? (
            <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13, border: `2px dashed ${C.border}`, borderRadius: 10 }}>No photos yet.</div>
          ) : (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 14 }}>
              {photos.map(photo => (
                <div key={photo.name} style={{ background: C.surface, borderRadius: 10, overflow: "hidden", border: `1px solid ${C.border}` }}>
                  <img src={photo.url} alt={photo.name} style={{ width: "100%", height: 160, objectFit: "cover", display: "block", cursor: "pointer" }} onClick={() => window.open(photo.url, "_blank")} />
                  <div style={{ padding: "8px 12px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div style={{ fontSize: 11, color: C.muted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 140 }}>{photo.name}</div>
                    <Btn small variant="danger" onClick={() => deletePhoto(photo.name)}>{t(lang,"del")}</Btn>
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

function WorkOrders({ workOrders, setWorkOrders, loading, onAdd, isAdmin, vendors, assets, lang }) {
  const [showForm, setShowForm] = useState(false); const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null); const [filter, setFilter] = useState("All");
  const [editItem, setEditItem] = useState(null); const [deleteItem, setDeleteItem] = useState(null); const [selectedWO, setSelectedWO] = useState(null);
  const [form, setForm] = useState({ title: "", asset: "", priority: "Medium", start_date: "", due: "", vendor: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const filtered = filter==="All" ? workOrders : workOrders.filter(w => w.status===filter);
  const vendorOptions = ["— None —",...vendors.filter(v => v.status==="Active").map(v => v.name)];

  const submit = async () => {
    if (!form.title||!form.asset) { setError(t(lang,"title")); return; }
    setSaving(true); setError(null);
    const vendorName = form.vendor === "— None —" || !form.vendor ? null : form.vendor;
    const record = { id: uid("WO"), title: form.title, asset: form.asset, priority: form.priority, status: "Open", assignee: null, start_date: form.start_date||null, due: form.due||null, vendor: vendorName };
    const { error: err } = await supabase.from("work_orders").insert([record]);
    if (err) { setError(err.message); } else {
      onAdd(record);
      if (vendorName) {
        const { data: vd } = await supabase.from("vendors").select("id, open_orders").eq("name", vendorName).single();
        if (vd) await supabase.from("vendors").update({ open_orders: (vd.open_orders||0)+1 }).eq("id", vd.id);
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
    if (wo?.vendor && val === "Completed" && wo.status !== "Completed") {
      const { data: vd } = await supabase.from("vendors").select("id, open_orders").eq("name", wo.vendor).single();
      if (vd) await supabase.from("vendors").update({ open_orders: Math.max(0,(vd.open_orders||0)-1) }).eq("id", vd.id);
    }
    if (wo?.vendor && wo.status === "Completed" && val !== "Completed") {
      const { data: vd } = await supabase.from("vendors").select("id, open_orders").eq("name", wo.vendor).single();
      if (vd) await supabase.from("vendors").update({ open_orders: (vd.open_orders||0)+1 }).eq("id", vd.id);
    }
    if (val === "Completed" && wo?.status !== "Completed" && wo?.asset) {
      const { data: assetData } = await supabase.from("assets").select("id").eq("name", wo.asset).single();
      if (assetData) {
        await supabase.from("maintenance_logs").insert([{ id: uid("LOG"), asset_id: assetData.id, asset_name: wo.asset, log_type: wo.title.startsWith("PM -") ? "Preventive Maintenance" : "Corrective Repair", title: wo.title, description: `Work order completed.\nVendor: ${wo.vendor||"—"}\nPriority: ${wo.priority}`, performed_by: wo.assignee||"—", vendor: wo.vendor||null, start_date: wo.start_date||TODAY, end_date: TODAY, cost: null, status: "Completed", downtime_start: null, downtime_end: null, downtime_hours: null }]);
      }
    }
  };

  const updatePriority = async (id,val) => { await supabase.from("work_orders").update({ priority: val }).eq("id",id); setWorkOrders(prev => prev.map(wo => wo.id===id?{...wo,priority:val}:wo)); };
  const saveEdit = async (updated) => { const { error: err } = await supabase.from("work_orders").update(updated).eq("id",updated.id); if (!err) { setWorkOrders(prev => prev.map(wo => wo.id===updated.id?updated:wo)); setEditItem(null); } else setError(err.message); };
  const confirmDelete = async () => { await supabase.from("work_orders").delete().eq("id",deleteItem.id); setWorkOrders(prev => prev.filter(wo => wo.id!==deleteItem.id)); setDeleteItem(null); };

  const filterLabels = { All: t(lang,"all"), Open: t(lang,"open"), "In Progress": "In Progress", Pending: "Pending", Completed: t(lang,"completed") };

  return (
    <div>
      {selectedWO && <WorkOrderPhotosModal workOrder={selectedWO} lang={lang} onClose={() => setSelectedWO(null)} />}
      {editItem && <EditModal lang={lang} title={t(lang,"workOrders")} data={editItem} fields={[{key:"title",label:t(lang,"title")},{key:"asset",label:t(lang,"asset")},{key:"priority",label:t(lang,"priority"),options:["Critical","High","Medium","Low"]},{key:"status",label:t(lang,"status"),options:["Open","In Progress","Pending","Completed"]},{key:"vendor",label:t(lang,"vendor"),options:vendorOptions},{key:"assignee",label:t(lang,"assignee")},{key:"start_date",label:t(lang,"startDate"),type:"date"},{key:"due",label:t(lang,"dueDate"),type:"date"}]} onSave={saveEdit} onClose={() => setEditItem(null)} />}
      {deleteItem && <ConfirmDel lang={lang} name={deleteItem.title} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {["All","Open","In Progress","Pending","Completed"].map(s => <button key={s} onClick={() => setFilter(s)} style={{ background: filter===s?C.accent:C.card, color: filter===s?"#fff":C.muted, border: `1px solid ${filter===s?C.accent:C.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>{filterLabels[s]||s}</button>)}
        </div>
        <Btn onClick={() => setShowForm(v => !v)}>{t(lang,"newWorkOrder")}</Btn>
      </div>
      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 18 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>{t(lang,"newWorkOrder")}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label={t(lang,"title")} value={form.title} onChange={f("title")} />
            <div>
              <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t(lang,"asset")}</div>
              <select value={form.asset} onChange={e => f("asset")(e.target.value)} style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14 }}>
                <option value="">{t(lang,"selectAsset")}</option>
                {(assets||[]).map(a => <option key={a.id} value={a.name}>{a.name} ({a.location})</option>)}
              </select>
            </div>
            <Input label={t(lang,"startDate")} value={form.start_date} onChange={f("start_date")} type="date" />
            <Input label={t(lang,"dueDate")} value={form.due} onChange={f("due")} type="date" />
            <Sel label={t(lang,"priority")} value={form.priority} onChange={f("priority")} options={["Critical","High","Medium","Low"]} />
            <Sel label={t(lang,"vendor")} value={form.vendor} onChange={f("vendor")} options={vendorOptions} />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
            <Btn onClick={submit} disabled={saving}>{saving?t(lang,"saving"):t(lang,"create")}</Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      )}
      {loading ? <Spinner lang={lang} /> : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 700 }}>
            <thead><tr style={{ borderBottom: `1px solid ${C.border}` }}>
              {[t(lang,"id"),t(lang,"title"),t(lang,"asset"),t(lang,"priority"),t(lang,"status"),t(lang,"vendor"),t(lang,"start"),t(lang,"due"),t(lang,"photos"),...(isAdmin?[t(lang,"actions")]:[])].map(h => <th key={h} style={{ textAlign: "left", padding: "8px 12px", fontSize: 11, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600 }}>{h}</th>)}
            </tr></thead>
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
                      {isAdmin && <><Btn small onClick={() => setEditItem(wo)} color={C.blue}>{t(lang,"edit")}</Btn><Btn small variant="danger" onClick={() => setDeleteItem(wo)}>{t(lang,"del")}</Btn></>}
                    </div>
                  </td>
                </tr>
              ))}
              {filtered.length===0 && <tr><td colSpan={isAdmin?10:9} style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>{t(lang,"noWorkOrdersFound")}</td></tr>}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
function AssetEditModal({ data, onSave, onClose, lang, mheModels }) {
  const [form, setForm] = useState({ ...data });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));

 const handleModelSelect = (modelName) => {
    f("model")(modelName);
    const found = mheModels.find(m => m.model === modelName);
    if (found) {
      setForm(p => ({
        ...p,
        model: modelName,
        brand: found.brand || p.brand || "",
        category: p.category || found.subcategory || found.category || "",
        technical_specs: found.technical_specs || p.technical_specs || "",
      }));
    }
  };
  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000aa", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 24, width: "100%", maxWidth: 620, maxHeight: "90vh", overflowY: "auto" }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: C.accent, marginBottom: 20, textTransform: "uppercase" }}>{t(lang,"edit")} — {data.name}</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
          <Input label={t(lang,"assetName")} value={form.name||""} onChange={f("name")} />
          <Input label={t(lang,"category")} value={form.category||""} onChange={f("category")} />
          <Sel label={t(lang,"site")} value={form.location||""} onChange={f("location")} options={SITES} />
          <Input label={t(lang,"owner")} value={form.owner||""} onChange={f("owner")} />
          <Input label={t(lang,"brand")} value={form.brand||""} onChange={f("brand")} />
          <div>
            <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t(lang,"model")}</div>
            <input list="mhe-models-edit" value={form.model||""} onChange={e => handleModelSelect(e.target.value)}
              placeholder="e.g. ETV 216"
              style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14, boxSizing: "border-box" }} />
            <datalist id="mhe-models-edit">
              {mheModels.map(m => <option key={m.model} value={m.model} />)}
            </datalist>
          </div>
          <Input label={t(lang,"serialNumber")} value={form.serial_number||""} onChange={f("serial_number")} />
          <Input label={t(lang,"manufactureDate")} value={form.manufacture_date||""} onChange={f("manufacture_date")} type="date" />
          <Input label={t(lang,"estValue")} value={form.value||""} onChange={f("value")} />
          <Sel label={t(lang,"status")} value={form.status||"Operational"} onChange={f("status")} options={["Operational","Under Maintenance","Degraded"]} />
          <Input label={t(lang,"pmFrequency")} value={form.pm_frequency||""} onChange={f("pm_frequency")} />
          <Input label={t(lang,"nextServiceDate")} value={form.next_service||""} onChange={f("next_service")} type="date" />
          <Input label={t(lang,"pmTask")} value={form.pm_task||""} onChange={f("pm_task")} />
        </div>
        <div style={{ marginTop: 12 }}>
          <Textarea label={t(lang,"technicalSpecs")} value={form.technical_specs||""} onChange={f("technical_specs")} placeholder="Engine specs, capacity, dimensions..." />
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
          <Btn onClick={() => onSave(form)}>{t(lang,"save")}</Btn>
          <Btn variant="secondary" onClick={onClose}>{t(lang,"cancel")}</Btn>
        </div>
      </div>
    </div>
  );
}
function Assets({ assets, setAssets, loading, onAdd, isAdmin, vendors, lang }) {
  const [showForm, setShowForm] = useState(false); const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null); const [editItem, setEditItem] = useState(null);
  const [deleteItem, setDeleteItem] = useState(null); const [selectedAsset, setSelectedAsset] = useState(null);
  const [siteFilter, setSiteFilter] = useState("All"); const [catFilter, setCatFilter] = useState("All"); const [ownerFilter, setOwnerFilter] = useState("All"); const [modelFilter, setModelFilter] = useState("All"); const [search, setSearch] = useState("");
  const [mheModels, setMheModels] = useState([]);

  useEffect(() => {
    supabase.from("mhe_models").select("brand, model, category, subcategory, technical_specs").order("brand").order("model")
      .then(({ data }) => setMheModels(data || []));
  }, []);

  const handleModelSelect = (modelName) => {
    f("model")(modelName);
    const found = mheModels.find(m => m.model === modelName);
    if (found) {
      if (!form.category) f("category")(found.subcategory || found.category || "");
      f("brand")(found.brand || "");
      f("technical_specs")(found.technical_specs || "");
    }
  };
  const [form, setForm] = useState({ name: "", category: "", location: "— Select Site —", value: "", owner: "", brand: "", model: "", serial_number: "", manufacture_date: "", technical_specs: "", next_service: "", pm_frequency: "1", pm_task: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const categories = ["All",...new Set(assets.map(a => a.category).filter(Boolean))];
  const owners = ["All",...new Set(assets.map(a => a.owner).filter(Boolean))];
const modelOptions = ["All",...new Set(assets.map(a => a.model).filter(Boolean))];
const filtered = assets.filter(a =>
  (siteFilter==="All"||a.location===siteFilter) &&
  (catFilter==="All"||a.category===catFilter) &&
  (ownerFilter==="All"||a.owner===ownerFilter) &&
  (modelFilter==="All"||a.model===modelFilter) &&
  (!search||a.name.toLowerCase().includes(search.toLowerCase()))
);

  const submit = async () => {
    if (!form.name) { setError(t(lang,"assetName")); return; }
    if (form.location==="— Select Site —") { setError(t(lang,"site")); return; }
    setSaving(true); setError(null);
    const record = { id: uid("AST"), name: form.name, category: form.category, location: form.location, value: form.value, owner: form.owner||null, brand: form.brand||null, model: form.model||null, serial_number: form.serial_number||null, manufacture_date: form.manufacture_date||null, technical_specs: form.technical_specs||null,
    const { error: err } = await supabase.from("assets").insert([record]);
    if (err) { setError(err.message); } else { onAdd(record); setForm({ name: "", category: "", location: "— Select Site —", value: "", next_service: "", pm_frequency: "1", pm_task: "" }); setShowForm(false); }
    setSaving(false);
  };

  const generateQR = async (asset) => {
    const url = `https://wali-835.github.io/facility-command/?asset=${asset.id}`;
    const canvas = document.createElement("canvas");
    await QRCode.toCanvas(canvas, url, { width: 300, margin: 2, color: { dark: "#000000", light: "#ffffff" } });
    const win = window.open("", "_blank");
    win.document.write(`<html><head><title>QR - ${asset.name}</title></head><body style="font-family:Arial;text-align:center;padding:40px;background:white;"><div style="border:2px solid #f97316;border-radius:12px;padding:30px;max-width:400px;margin:0 auto;"><div style="font-size:14px;color:#666;margin-bottom:8px;text-transform:uppercase;letter-spacing:2px;">EPx Logistics — Facility Command</div><div style="font-size:22px;font-weight:bold;color:#0d0f12;margin-bottom:4px;">${asset.name}</div><div style="font-size:14px;color:#666;margin-bottom:20px;">${asset.category} · ${asset.location}</div><img src="${canvas.toDataURL()}" style="width:250px;height:250px;" /><div style="font-size:12px;color:#999;margin-top:16px;">${t(lang,"qrScanToReport")}</div><div style="font-size:10px;color:#ccc;margin-top:8px;">${asset.id}</div></div><button onclick="window.print()" style="margin-top:20px;background:#f97316;color:white;border:none;padding:12px 24px;border-radius:8px;font-size:14px;cursor:pointer;">🖨️ Print</button></body></html>`);
    win.document.close();
  };

  const updateStatus = async (id,val) => { await supabase.from("assets").update({ status: val }).eq("id",id); setAssets(prev => prev.map(a => a.id===id?{...a,status:val}:a)); };
  const saveEdit = async (updated) => { const { error: err } = await supabase.from("assets").update(updated).eq("id",updated.id); if (!err) { setAssets(prev => prev.map(a => a.id===updated.id?updated:a)); setEditItem(null); } else setError(err.message); };
  const confirmDelete = async () => { await supabase.from("assets").delete().eq("id",deleteItem.id); setAssets(prev => prev.filter(a => a.id!==deleteItem.id)); setDeleteItem(null); };

  return (
    <div>
      {editItem && <AssetEditModal lang={lang} data={editItem} mheModels={mheModels} onSave={saveEdit} onClose={() => setEditItem(null)} />}
      {deleteItem && <ConfirmDel lang={lang} name={deleteItem.name} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      {selectedAsset && <MaintenanceModal asset={selectedAsset} lang={lang} onClose={() => setSelectedAsset(null)} isAdmin={isAdmin} vendors={vendors} />}
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 18 }}>
        <input value={search} onChange={e => setSearch(e.target.value)} placeholder={t(lang,"searchAssets")} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 12px", color: C.text, fontSize: 13, flex: "1 1 180px" }} />
        <select value={siteFilter} onChange={e => setSiteFilter(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12 }}>
          <option>{t(lang,"all")}</option>{SITES.filter(s => s!=="— Select Site —").map(s => <option key={s}>{s}</option>)}
        </select>
       <select value={catFilter} onChange={e => setCatFilter(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12 }}>
          {categories.map(c => <option key={c}>{c}</option>)}
        </select>
        <select value={ownerFilter} onChange={e => setOwnerFilter(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12 }}>
          {owners.map(o => <option key={o}>{o}</option>)}
        </select>
        <select value={modelFilter} onChange={e => setModelFilter(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12 }}>
          {modelOptions.map(m => <option key={m}>{m}</option>)}
        </select>
        {isAdmin && <Btn onClick={() => setShowForm(v => !v)}>{t(lang,"addAsset")}</Btn>}
      </div>
      {showForm && isAdmin && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 18 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>{t(lang,"registerNewAsset")}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label={t(lang,"assetName")} value={form.name} onChange={f("name")} />
            <Input label={t(lang,"category")} value={form.category} onChange={f("category")} />
            <Sel label={t(lang,"site")} value={form.location} onChange={f("location")} options={SITES} />
            <Input label={t(lang,"owner")} value={form.owner} onChange={f("owner")} placeholder="e.g. EPx Logistics" />
            <div>
  <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t(lang,"model")}</div>
  <input list="mhe-models-list" value={form.model} onChange={e => handleModelSelect(e.target.value)}
    placeholder="e.g. ETV 216"
    style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14, boxSizing: "border-box" }} />
  <datalist id="mhe-models-list">
    {mheModels.map(m => <option key={m.model} value={m.model} />)}
  </datalist>
</div>
            <Input label={t(lang,"serialNumber")} value={form.serial_number} onChange={f("serial_number")} />
            <Input label={t(lang,"manufactureDate")} value={form.manufacture_date} onChange={f("manufacture_date")} type="date" />
            <Input label={t(lang,"estValue")} value={form.value} onChange={f("value")} />
            <Input label={t(lang,"nextServiceDate")} value={form.next_service} onChange={f("next_service")} type="date" />
            <Sel label={t(lang,"pmFrequency")} value={form.pm_frequency} onChange={f("pm_frequency")} options={["1","2","3","6","12"]} />
            <Input label={t(lang,"pmTask")} value={form.pm_task} onChange={f("pm_task")} />
          </div>
          <div style={{ marginTop: 12 }}>
            <Textarea label={t(lang,"technicalSpecs")} value={form.technical_specs} onChange={f("technical_specs")} placeholder="Engine specs, capacity, dimensions..." />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
            <Btn onClick={submit} disabled={saving}>{saving?t(lang,"saving"):t(lang,"register")}</Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      )}
          <div style={{ marginTop: 12 }}>
            <Textarea label={t(lang,"technicalSpecs")} value={form.technical_specs} onChange={f("technical_specs")} placeholder="Engine specs, capacity, dimensions..." />
          </div>
      {loading ? <Spinner lang={lang} /> : (
        <div>
          <div style={{ fontSize: 12, color: C.muted, marginBottom: 12 }}>{filtered.length} {t(lang,"assetsShown")}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 14 }}>
            {filtered.map(a => (
              <div key={a.id} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 18, borderTop: `3px solid ${statusColor(a.status)}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                  <div style={{ flex: 1, marginRight: 8 }}><div style={{ fontSize: 14, fontWeight: 700, color: C.text, lineHeight: 1.3 }}>{a.name}</div><div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>{a.category}</div></div>
                  <StatusSel value={a.status} options={["Operational","Under Maintenance","Degraded"]} onChange={val => updateStatus(a.id,val)} />
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12, marginBottom: 14 }}>
                  {[[t(lang,"site"),a.location],[t(lang,"owner"),a.owner||"—"],[t(lang,"brand"),a.brand||"—"],[t(lang,"model"),a.model||"—"],[t(lang,"serialNumber"),a.serial_number||"—"],[t(lang,"value"),a.value||"—"],[t(lang,"pmEvery"),a.pm_frequency?`${a.pm_frequency} mo.`:"—"],[t(lang,"lastPM"),a.last_pm_date?fmtDate(a.last_pm_date):t(lang,"never")],[t(lang,"manufactureDate"),a.manufacture_date?fmtDate(a.manufacture_date):"—"]].map(([lbl,val]) => (
                    <div key={lbl}><div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase" }}>{lbl}</div><div style={{ color: C.subtle, marginTop: 2 }}>{val||"—"}</div></div>
                  ))}
                </div>
                {a.technical_specs && (
                  <div style={{ fontSize: 11, color: C.muted, marginBottom: 10, padding: "8px 10px", background: C.surface, borderRadius: 6, lineHeight: 1.5 }}>
                    <span style={{ color: C.subtle, fontWeight: 600 }}>{t(lang,"technicalSpecs")}: </span>{a.technical_specs}
                  </div>
                )}
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <button onClick={() => setSelectedAsset(a)} style={{ flex: 1, background: C.blue+"22", color: C.blue, border: `1px solid ${C.blue}44`, borderRadius: 6, padding: "7px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{t(lang,"logChecklist")}</button>
                  <button onClick={() => generateQR(a)} style={{ background: C.purple+"22", color: "#a855f7", border: `1px solid #a855f744`, borderRadius: 6, padding: "7px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{t(lang,"qrCode")}</button>
                  {isAdmin && <><Btn small onClick={() => setEditItem(a)} color={C.accent}>{t(lang,"edit")}</Btn><Btn small variant="danger" onClick={() => setDeleteItem(a)}>{t(lang,"del")}</Btn></>}
                </div>
              </div>
            ))}
            {filtered.length===0 && <div style={{ color: C.muted, fontSize: 13, padding: 32 }}>{t(lang,"noAssetsFound")}</div>}
          </div>
        </div>
      )}
    </div>
  );
}

function VendorWorkOrdersModal({ vendor, onClose, lang }) {
  const [wos, setWos] = useState([]); const [loading, setLoading] = useState(true); const [filter, setFilter] = useState("All");
  useEffect(() => { supabase.from("work_orders").select("*").eq("vendor", vendor.name).then(({ data }) => { setWos(data||[]); setLoading(false); }); }, [vendor.name]);
  const filtered = filter==="All" ? wos : wos.filter(w => w.status===filter);
  const open = wos.filter(w => w.status!=="Completed").length;
  const completed = wos.filter(w => w.status==="Completed").length;
  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1000, padding: 16, overflowY: "auto" }}>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, width: "100%", maxWidth: 820, marginTop: 20, marginBottom: 20 }}>
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div><div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>{vendor.name}</div><div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{vendor.specialty}</div></div>
          <button onClick={onClose} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, color: C.muted, cursor: "pointer", fontSize: 18, padding: "2px 10px" }}>✕</button>
        </div>
        <div style={{ padding: "14px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 12, flexWrap: "wrap" }}>
          {[["📋",wos.length,t(lang,"totalWorkOrders"),C.blue],["🔓",open,t(lang,"open"),C.accent],["✅",completed,t(lang,"completed"),C.green]].map(([icon,val,label,color]) => (
            <div key={label} style={{ background: C.surface, borderRadius: 8, padding: "10px 16px", flex: "1 1 100px", borderLeft: `3px solid ${color}` }}>
              <div style={{ fontSize: 16 }}>{icon}</div><div style={{ fontSize: 22, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{val}</div><div style={{ fontSize: 11, color: C.muted }}>{label}</div>
            </div>
          ))}
        </div>
        <div style={{ padding: 24 }}>
          <div style={{ display: "flex", gap: 6, marginBottom: 18, flexWrap: "wrap" }}>
            {["All","Open","In Progress","Pending","Completed"].map(s => <button key={s} onClick={() => setFilter(s)} style={{ background: filter===s?C.accent:C.surface, color: filter===s?"#fff":C.muted, border: `1px solid ${filter===s?C.accent:C.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>{s}</button>)}
          </div>
          {loading ? <Spinner lang={lang} /> : filtered.length===0 ? <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>{t(lang,"noWorkOrdersFound")}</div> : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
                <thead><tr style={{ borderBottom: `1px solid ${C.border}` }}>
                  {[t(lang,"title"),t(lang,"asset"),t(lang,"priority"),t(lang,"status"),t(lang,"start"),t(lang,"due")].map(h => <th key={h} style={{ textAlign: "left", padding: "8px 12px", fontSize: 11, color: C.muted, textTransform: "uppercase", fontWeight: 600 }}>{h}</th>)}
                </tr></thead>
                <tbody>
                  {filtered.map((wo,i) => (
                    <tr key={wo.id} style={{ borderBottom: `1px solid ${C.border}22`, background: i%2===0?"transparent":C.surface+"44" }}>
                      <td style={{ padding: "10px 12px", fontSize: 13, color: C.text, fontWeight: 600 }}>{wo.title}</td>
                      <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.asset}</td>
                      <td style={{ padding: "10px 12px" }}><Badge label={wo.priority} color={priorityColor(wo.priority)} /></td>
                      <td style={{ padding: "10px 12px" }}><Badge label={wo.status} color={statusColor(wo.status)} /></td>
                      <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.start_date||"—"}</td>
                      <td style={{ padding: "10px 12px", fontSize: 12, color: wo.due&&wo.due<=TODAY&&wo.status!=="Completed"?C.red:C.subtle }}>{wo.due||"—"}</td>
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

function Vendors({ vendors, setVendors, loading, onAdd, isAdmin, lang }) {
  const [showForm, setShowForm] = useState(false); const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null); const [editItem, setEditItem] = useState(null); const [deleteItem, setDeleteItem] = useState(null); const [selectedVendor, setSelectedVendor] = useState(null);
  const [form, setForm] = useState({ name: "", specialty: "", contact: "", phone: "", email: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const submit = async () => {
    if (!form.name) { setError(t(lang,"companyName")); return; }
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
      {selectedVendor && <VendorWorkOrdersModal vendor={selectedVendor} lang={lang} onClose={() => setSelectedVendor(null)} />}
      {editItem && <EditModal lang={lang} title={t(lang,"vendors")} data={editItem} fields={[{key:"name",label:t(lang,"companyName")},{key:"specialty",label:t(lang,"specialty")},{key:"contact",label:t(lang,"contact")},{key:"phone",label:t(lang,"phone")},{key:"email",label:t(lang,"email")},{key:"status",label:t(lang,"status"),options:["Active","Inactive"]},{key:"rating",label:t(lang,"rating")}]} onSave={saveEdit} onClose={() => setEditItem(null)} />}
      {deleteItem && <ConfirmDel lang={lang} name={deleteItem.name} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 18 }}><Btn onClick={() => setShowForm(v => !v)}>{t(lang,"addVendor")}</Btn></div>
      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 18 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>{t(lang,"registerVendor")}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label={t(lang,"companyName")} value={form.name} onChange={f("name")} />
            <Input label={t(lang,"specialty")} value={form.specialty} onChange={f("specialty")} />
            <Input label={t(lang,"contact")} value={form.contact} onChange={f("contact")} />
            <Input label={t(lang,"phone")} value={form.phone} onChange={f("phone")} />
            <Input label={t(lang,"email")} value={form.email} onChange={f("email")} />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
            <Btn onClick={submit} disabled={saving}>{saving?t(lang,"saving"):t(lang,"register")}</Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      )}
      {loading ? <Spinner lang={lang} /> : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 14 }}>
          {vendors.map(v => (
            <div key={v.id} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
                <div><div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{v.name}</div><div style={{ fontSize: 11, color: C.muted }}>{v.specialty}</div></div>
                <Badge label={v.status} color={statusColor(v.status)} />
              </div>
              <Stars rating={v.rating} />
              <div style={{ marginTop: 12, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12 }}>
                {[[t(lang,"contact"),v.contact],[t(lang,"phone"),v.phone],[t(lang,"email"),v.email]].map(([lbl,val]) => (
                  <div key={lbl}><div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase" }}>{lbl}</div><div style={{ color: C.subtle, marginTop: 2 }}>{val||"—"}</div></div>
                ))}
                <div><div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase" }}>{t(lang,"openOrders")}</div><div style={{ color: (v.open_orders||0)>0?C.accent:C.subtle, marginTop: 2, fontWeight: (v.open_orders||0)>0?700:400 }}>{v.open_orders||0}</div></div>
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
                <button onClick={() => setSelectedVendor(v)} style={{ flex: 1, background: C.blue+"22", color: C.blue, border: `1px solid ${C.blue}44`, borderRadius: 6, padding: "7px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{t(lang,"workOrdersBtn")}</button>
                {isAdmin && <><Btn small onClick={() => setEditItem(v)} color={C.accent}>{t(lang,"edit")}</Btn><Btn small variant="danger" onClick={() => setDeleteItem(v)}>{t(lang,"delete")}</Btn></>}
              </div>
            </div>
          ))}
          {vendors.length===0 && <div style={{ color: C.muted, fontSize: 13 }}>{t(lang,"noVendors")}</div>}
        </div>
      )}
    </div>
  );
}

function PMUpload({ assets, onAssetsImported, onWorkOrdersGenerated, lang }) {
  const [generating, setGenerating] = useState(false); const [uploading, setUploading] = useState(false);
  const [error, setError] = useState(null); const [success, setSuccess] = useState(null);
  const pmDueCount = assets.filter(a => { if (!a.pm_frequency) return false; if (!a.last_pm_date) return true; const now=new Date(); const last=new Date(a.last_pm_date); return (now.getFullYear()-last.getFullYear())*12+(now.getMonth()-last.getMonth())>=a.pm_frequency; }).length;

  const generatePMWorkOrders = async () => {
    setGenerating(true); setError(null);
    const now=new Date(); const due=assets.filter(a => { if (!a.pm_frequency) return false; if (!a.last_pm_date) return true; const last=new Date(a.last_pm_date); return (now.getFullYear()-last.getFullYear())*12+(now.getMonth()-last.getMonth())>=a.pm_frequency; });
    if (!due.length) { setSuccess(t(lang,"pmDueMonth")); setGenerating(false); return; }
    const dueDate=new Date(now.getFullYear(),now.getMonth()+1,0).toISOString().split("T")[0];
    const newWOs=due.map(a => ({ id: uid("WO"), title: `PM - ${a.name}`, asset: a.name, priority: "Medium", status: "Open", assignee: null, start_date: TODAY, due: dueDate, vendor: null }));
    const { error: err }=await supabase.from("work_orders").insert(newWOs);
    if (err) { setError(err.message); } else { await supabase.from("assets").update({ last_pm_date: TODAY }).in("id",due.map(a => a.id)); setSuccess(`Generated ${newWOs.length}!`); onWorkOrdersGenerated(newWOs); }
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
        if (err) { setError(err.message); } else { setSuccess(`Imported ${records.length}!`); onAssetsImported(records); }
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
        <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 6 }}>{t(lang,"generatePMWorkOrders")}</div>
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 20 }}>
          {[["📋",pmDueCount,t(lang,"assetsDueForPM"),C.accent],["🏭",assets.length,t(lang,"totalAssets"),C.blue],["📅",new Date().toLocaleString("default",{month:"long",year:"numeric"}),t(lang,"currentMonth"),C.green]].map(([icon,val,label,color]) => (
            <div key={label} style={{ background: C.surface, borderRadius: 8, padding: "12px 20px", borderLeft: `3px solid ${color}` }}><div style={{ fontSize: 22, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{val}</div><div style={{ fontSize: 12, color: C.muted }}>{label}</div></div>
          ))}
        </div>
        <Btn onClick={generatePMWorkOrders} disabled={generating||pmDueCount===0}>{generating?t(lang,"generating"):`${t(lang,"generate")} ${pmDueCount}`}</Btn>
      </div>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 24 }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 6 }}>{t(lang,"importAssetsExcel")}</div>
        <label style={{ display: "inline-block", background: C.accent, color: "#fff", borderRadius: 6, padding: "10px 20px", fontSize: 13, fontWeight: 700, cursor: uploading?"not-allowed":"pointer", opacity: uploading?0.7:1 }}>
          {uploading?t(lang,"importing"):t(lang,"chooseExcelFile")}
          <input type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display: "none" }} disabled={uploading} />
        </label>
      </div>
    </div>
  );
}

function Overview({ workOrders, assets, vendors, lang }) {
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
        <StatCard icon="🔧" label={t(lang,"openWorkOrders")} value={open} sub={`${critical} ${t(lang,"critical")}`} color={C.accent} />
        <StatCard icon="🏭" label={t(lang,"operationalAssets")} value={`${opAssets}/${assets.length}`} sub={t(lang,"fleetStatus")} color={C.green} />
        <StatCard icon="🚨" label={t(lang,"assetsDown")} value={downAssets} sub={t(lang,"underMaintenance")} color={C.red} />
        <StatCard icon="⚠️" label={t(lang,"overdueAtRisk")} value={overdue} sub={t(lang,"pastDueDate")} color={C.yellow} />
        <StatCard icon="📋" label={t(lang,"pmDueMonth")} value={pmDue} sub={t(lang,"preventiveMaintenance")} color={C.blue} />
        <StatCard icon="🤝" label={t(lang,"activeVendors")} value={activeVendors} sub={t(lang,"contractorsOnFile")} color={C.purple} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 16 }}>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
          <div style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>{t(lang,"recentWorkOrders")}</div>
          {workOrders.slice(0,5).map(wo => (
            <div key={wo.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${C.border}22`, flexWrap: "wrap", gap: 6 }}>
              <div><div style={{ fontSize: 13, color: C.text, fontWeight: 600 }}>{wo.title}</div><div style={{ fontSize: 11, color: C.muted }}>{wo.asset}</div></div>
              <div style={{ display: "flex", gap: 6 }}><Badge label={wo.priority} color={priorityColor(wo.priority)} /><Badge label={wo.status} color={statusColor(wo.status)} /></div>
            </div>
          ))}
          {workOrders.length===0 && <div style={{ color: C.muted, fontSize: 13 }}>{t(lang,"noWorkOrdersFound")}</div>}
        </div>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
          <div style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>{t(lang,"assetsBySite")}</div>
          {SITES.filter(s => s!=="— Select Site —").map(site => {
            const count=assets.filter(a => a.location===site).length; if (!count) return null;
            const op=assets.filter(a => a.location===site&&a.status==="Operational").length;
            const down=assets.filter(a => a.location===site&&a.status==="Under Maintenance").length;
            return <div key={site} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <span style={{ fontSize: 13, color: C.subtle }}>{site}</span>
              <div style={{ display: "flex", gap: 8 }}>
                <span style={{ fontSize: 11, color: C.green }}>{op} ok</span>
                {down>0 && <span style={{ fontSize: 11, color: C.red }}>{down} {t(lang,"assetsDown")}</span>}
                <span style={{ fontSize: 13, color: C.text, fontWeight: 600 }}>{count}</span>
              </div>
            </div>;
          })}
        </div>
      </div>
    </div>
  );
}

function MaintenanceCalendar({ workOrders, assets, lang }) {
  const now = new Date();
  const [currentMonth, setCurrentMonth] = useState(now.getMonth());
  const [currentYear, setCurrentYear] = useState(now.getFullYear());
  const [selectedDay, setSelectedDay] = useState(null);
  const [filter, setFilter] = useState("All");

  const monthNames = lang === "ar"
    ? ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"]
    : ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const dayNames = lang === "ar"
    ? ["أحد","إثن","ثلا","أرب","خمي","جمع","سبت"]
    : ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];

  const prevMonth = () => { if (currentMonth===0){setCurrentMonth(11);setCurrentYear(y=>y-1);}else setCurrentMonth(m=>m-1); setSelectedDay(null); };
  const nextMonth = () => { if (currentMonth===11){setCurrentMonth(0);setCurrentYear(y=>y+1);}else setCurrentMonth(m=>m+1); setSelectedDay(null); };

  const daysInMonth = new Date(currentYear, currentMonth+1, 0).getDate();
  const firstDayOfMonth = new Date(currentYear, currentMonth, 1).getDay();

  const getEventsForDay = (day) => {
    const dateStr = `${currentYear}-${String(currentMonth+1).padStart(2,"0")}-${String(day).padStart(2,"0")}`;
    const events = [];
    workOrders.forEach(wo => {
      if (wo.due===dateStr&&(filter==="All"||filter==="Work Orders")) events.push({ type: "work_order", title: wo.title, asset: wo.asset, status: wo.status, priority: wo.priority, color: wo.status==="Completed"?C.green:wo.due<=TODAY&&wo.status!=="Completed"?C.red:priorityColor(wo.priority) });
      if (wo.start_date===dateStr&&(filter==="All"||filter==="Work Orders")) events.push({ type: "work_order_start", title: `▶ ${wo.title}`, asset: wo.asset, status: wo.status, color: C.blue });
    });
    assets.forEach(a => { if (a.next_service===dateStr&&(filter==="All"||filter==="PM")) events.push({ type: "pm", title: `PM: ${a.name}`, asset: a.name, color: C.yellow }); });
    return events;
  };

  const selectedEvents = selectedDay ? getEventsForDay(selectedDay) : [];
  const eventCounts = {};
  for (let d=1;d<=daysInMonth;d++) eventCounts[d]=getEventsForDay(d).length;

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20, flexWrap: "wrap", gap: 10 }}>
        <div style={{ fontSize: 18, fontWeight: 700, color: C.text }}>{t(lang,"maintenanceCalendar")}</div>
        <div style={{ display: "flex", gap: 6 }}>
          {["All","Work Orders","PM"].map(f => <button key={f} onClick={() => setFilter(f)} style={{ background: filter===f?C.accent:C.card, color: filter===f?"#fff":C.muted, border: `1px solid ${filter===f?C.accent:C.border}`, borderRadius: 6, padding: "6px 14px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>{f==="All"?t(lang,"all"):f==="Work Orders"?t(lang,"workOrders"):t(lang,"pmPlanner")}</button>)}
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 320px", gap: 20, alignItems: "start" }}>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
          <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <button onClick={prevMonth} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, color: C.text, cursor: "pointer", padding: "6px 14px", fontSize: 16 }}>‹</button>
            <div style={{ fontSize: 16, fontWeight: 700, color: C.text }}>{monthNames[currentMonth]} {currentYear}</div>
            <button onClick={nextMonth} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, color: C.text, cursor: "pointer", padding: "6px 14px", fontSize: 16 }}>›</button>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", borderBottom: `1px solid ${C.border}` }}>
            {dayNames.map(d => <div key={d} style={{ padding: "10px 0", textAlign: "center", fontSize: 11, color: C.muted, fontWeight: 600, textTransform: "uppercase" }}>{d}</div>)}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)" }}>
            {Array.from({ length: firstDayOfMonth }).map((_,i) => <div key={`e-${i}`} style={{ padding: "10px 0", minHeight: 60, borderRight: `1px solid ${C.border}22`, borderBottom: `1px solid ${C.border}22` }} />)}
            {Array.from({ length: daysInMonth }).map((_,i) => {
              const day=i+1;
              const dateStr=`${currentYear}-${String(currentMonth+1).padStart(2,"0")}-${String(day).padStart(2,"0")}`;
              const isToday=dateStr===TODAY; const isSelected=selectedDay===day;
              const count=eventCounts[day]; const events=getEventsForDay(day);
              const hasOverdue=events.some(e => e.color===C.red); const hasPM=events.some(e => e.type==="pm");
              return (
                <div key={day} onClick={() => setSelectedDay(day===selectedDay?null:day)} style={{ padding: "8px 6px", minHeight: 60, borderRight: `1px solid ${C.border}22`, borderBottom: `1px solid ${C.border}22`, cursor: count>0?"pointer":"default", background: isSelected?C.accent+"22":isToday?C.blue+"11":"transparent", transition: "background 0.15s" }}>
                  <div style={{ width: 26, height: 26, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", background: isToday?C.blue:isSelected?C.accent:"transparent", color: isToday||isSelected?"#fff":C.text, fontSize: 13, fontWeight: isToday?700:400, margin: "0 auto 4px" }}>{day}</div>
                  {count>0 && <div style={{ display: "flex", justifyContent: "center", gap: 3, flexWrap: "wrap" }}>
                    {hasOverdue && <div style={{ width: 7, height: 7, borderRadius: "50%", background: C.red }} />}
                    {hasPM && <div style={{ width: 7, height: 7, borderRadius: "50%", background: C.yellow }} />}
                    {!hasOverdue&&!hasPM&&count>0 && <div style={{ width: 7, height: 7, borderRadius: "50%", background: C.accent }} />}
                    {count>2 && <div style={{ fontSize: 9, color: C.muted }}>+{count-1}</div>}
                  </div>}
                </div>
              );
            })}
          </div>
          <div style={{ padding: "12px 20px", borderTop: `1px solid ${C.border}`, display: "flex", gap: 16, flexWrap: "wrap" }}>
            {[["🔴",t(lang,"overdue")],["🟡","PM"],["🟠",t(lang,"workOrders")],["🔵","Today"]].map(([dot,label]) => (
              <div key={label} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: C.muted }}><span>{dot}</span> {label}</div>
            ))}
          </div>
        </div>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 20, position: "sticky", top: 20 }}>
          {selectedDay ? (
            <>
              <div style={{ fontSize: 15, fontWeight: 700, color: C.text, marginBottom: 4 }}>{monthNames[currentMonth]} {selectedDay}, {currentYear}</div>
              <div style={{ fontSize: 12, color: C.muted, marginBottom: 16 }}>{selectedEvents.length} event(s)</div>
              {selectedEvents.length===0 ? <div style={{ textAlign: "center", padding: 24, color: C.muted, fontSize: 13 }}>{t(lang,"noEventsThisDay")}</div> : (
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {selectedEvents.map((ev,i) => (
                    <div key={i} style={{ background: ev.color+"11", border: `1px solid ${ev.color}44`, borderRadius: 8, padding: 12, borderLeft: `3px solid ${ev.color}` }}>
                      <div style={{ fontSize: 13, fontWeight: 700, color: C.text, marginBottom: 4 }}>{ev.title}</div>
                      {ev.asset && <div style={{ fontSize: 11, color: C.muted }}>📍 {ev.asset}</div>}
                      {ev.status && <div style={{ marginTop: 6 }}><Badge label={ev.status} color={statusColor(ev.status)} /></div>}
                    </div>
                  ))}
                </div>
              )}
            </>
          ) : (
            <div style={{ textAlign: "center", padding: 24 }}>
              <div style={{ fontSize: 32, marginBottom: 12 }}>🗓</div>
              <div style={{ fontSize: 13, color: C.muted }}>{t(lang,"clickDayToSee")}</div>
              <div style={{ marginTop: 20 }}>
                <div style={{ fontSize: 12, color: C.muted, marginBottom: 10, textTransform: "uppercase", fontWeight: 600 }}>{t(lang,"thisMonthLabel")}</div>
                {[[t(lang,"workOrdersDue"),workOrders.filter(w => { const d=new Date(w.due||""); return d.getMonth()===currentMonth&&d.getFullYear()===currentYear; }).length,C.accent],[t(lang,"pmScheduled"),assets.filter(a => { const d=new Date(a.next_service||""); return d.getMonth()===currentMonth&&d.getFullYear()===currentYear; }).length,C.yellow],[t(lang,"overdue"),workOrders.filter(w => w.due&&w.due<=TODAY&&w.status!=="Completed").length,C.red]].map(([label,count,color]) => (
                  <div key={label} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${C.border}22` }}>
                    <span style={{ fontSize: 13, color: C.subtle }}>{label}</span>
                    <span style={{ fontSize: 13, fontWeight: 700, color }}>{count}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function Reports({ workOrders, assets, vendors, lang }) {
  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();
    const woData = workOrders.map(wo => ({ "ID": wo.id, "Title": wo.title, "Asset": wo.asset, "Priority": wo.priority, "Status": wo.status, "Vendor": wo.vendor||"—", "Start Date": wo.start_date||"—", "Due Date": wo.due||"—" }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(woData), "Work Orders");
    const assetData = assets.map(a => ({ "ID": a.id, "Name": a.name, "Category": a.category, "Site": a.location, "Status": a.status, "Value": a.value||"—", "PM Frequency": a.pm_frequency||"—", "Last PM": a.last_pm_date||"Never", "Next Service": a.next_service||"—" }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(assetData), "Assets");
    const vendorData = vendors.map(v => ({ "Name": v.name, "Specialty": v.specialty||"—", "Contact": v.contact||"—", "Phone": v.phone||"—", "Email": v.email||"—", "Status": v.status, "Rating": v.rating||"N/A", "Open Orders": v.open_orders||0 }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(vendorData), "Vendors");
    const totalWOs=workOrders.length; const completedWOs=workOrders.filter(w => w.status==="Completed").length; const overdueWOs=workOrders.filter(w => w.due&&w.due<=TODAY&&w.status!=="Completed").length; const operationalAssets=assets.filter(a => a.status==="Operational").length; const pmDue=assets.filter(a => { if (!a.pm_frequency) return false; if (!a.last_pm_date) return true; const now=new Date(); const last=new Date(a.last_pm_date); return (now.getFullYear()-last.getFullYear())*12+(now.getMonth()-last.getMonth())>=a.pm_frequency; }).length;
    const kpiData = [{ "KPI": "Total Work Orders", "Value": totalWOs },{ "KPI": "Completed", "Value": completedWOs },{ "KPI": "Completion Rate", "Value": `${totalWOs>0?Math.round((completedWOs/totalWOs)*100):0}%` },{ "KPI": "Overdue", "Value": overdueWOs },{ "KPI": "Total Assets", "Value": assets.length },{ "KPI": "Operational", "Value": operationalAssets },{ "KPI": "Uptime Rate", "Value": `${assets.length>0?Math.round((operationalAssets/assets.length)*100):0}%` },{ "KPI": "PM Compliance", "Value": `${assets.length>0?Math.round(((assets.length-pmDue)/assets.length)*100):0}%` },{ "KPI": "Active Vendors", "Value": vendors.filter(v => v.status==="Active").length },{ "KPI": "Generated", "Value": new Date().toLocaleString("en-GB") }];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(kpiData), "KPI Summary");
    XLSX.writeFile(wb, `Facility_Report_${TODAY}.xlsx`);
  };

  const exportToPDF = async () => {
    applyPlugin(jsPDF);
    const doc = new jsPDF();
    const totalWOs=workOrders.length; const completedWOs=workOrders.filter(w => w.status==="Completed").length; const overdueWOs=workOrders.filter(w => w.due&&w.due<=TODAY&&w.status!=="Completed").length; const operationalAssets=assets.filter(a => a.status==="Operational").length; const pmDue=assets.filter(a => { if (!a.pm_frequency) return false; if (!a.last_pm_date) return true; const now=new Date(); const last=new Date(a.last_pm_date); return (now.getFullYear()-last.getFullYear())*12+(now.getMonth()-last.getMonth())>=a.pm_frequency; }).length;
    doc.setFillColor(249,115,22); doc.rect(0,0,220,28,"F"); doc.setTextColor(255,255,255); doc.setFontSize(18); doc.setFont("helvetica","bold"); doc.text("FACILITY COMMAND",14,12); doc.setFontSize(10); doc.setFont("helvetica","normal"); doc.text("Industrial Warehouse Management Report",14,20); doc.text(`Generated: ${new Date().toLocaleString("en-GB")}`,14,26);
    doc.setTextColor(0,0,0); doc.setFontSize(13); doc.setFont("helvetica","bold"); doc.text("KPI Summary",14,40);
    doc.autoTable({ startY: 44, head: [["KPI","Value"]], body: [["Total Work Orders",totalWOs],["Completed",completedWOs],["Completion Rate",`${totalWOs>0?Math.round((completedWOs/totalWOs)*100):0}%`],["Overdue",overdueWOs],["Total Assets",assets.length],["Operational",operationalAssets],["Uptime Rate",`${assets.length>0?Math.round((operationalAssets/assets.length)*100):0}%`],["PM Compliance",`${assets.length>0?Math.round(((assets.length-pmDue)/assets.length)*100):0}%`],["Active Vendors",vendors.filter(v => v.status==="Active").length]], headStyles: { fillColor: [249,115,22], textColor: 255 }, alternateRowStyles: { fillColor: [245,245,245] }, margin: { left: 14, right: 14 } });
    doc.addPage(); doc.setFontSize(13); doc.setFont("helvetica","bold"); doc.text("Work Orders",14,20);
    doc.autoTable({ startY: 24, head: [["Title","Asset","Priority","Status","Vendor","Due"]], body: workOrders.slice(0,50).map(wo => [wo.title,wo.asset,wo.priority,wo.status,wo.vendor||"—",wo.due||"—"]), headStyles: { fillColor: [249,115,22], textColor: 255 }, alternateRowStyles: { fillColor: [245,245,245] }, margin: { left: 14, right: 14 }, styles: { fontSize: 9 } });
    doc.addPage(); doc.setFontSize(13); doc.setFont("helvetica","bold"); doc.text("Assets",14,20);
    doc.autoTable({ startY: 24, head: [["Name","Category","Site","Status","PM Every","Last PM"]], body: assets.slice(0,50).map(a => [a.name,a.category||"—",a.location,a.status,a.pm_frequency?`${a.pm_frequency} mo.`:"—",a.last_pm_date?fmtDate(a.last_pm_date):"Never"]), headStyles: { fillColor: [249,115,22], textColor: 255 }, alternateRowStyles: { fillColor: [245,245,245] }, margin: { left: 14, right: 14 }, styles: { fontSize: 9 } });
    doc.addPage(); doc.setFontSize(13); doc.setFont("helvetica","bold"); doc.text("Vendors",14,20);
    doc.autoTable({ startY: 24, head: [["Name","Specialty","Contact","Phone","Status","Open Orders","Rating"]], body: vendors.map(v => [v.name,v.specialty||"—",v.contact||"—",v.phone||"—",v.status,v.open_orders||0,v.rating>0?v.rating.toFixed(1):"N/A"]), headStyles: { fillColor: [249,115,22], textColor: 255 }, alternateRowStyles: { fillColor: [245,245,245] }, margin: { left: 14, right: 14 }, styles: { fontSize: 9 } });
    doc.save(`Facility_Report_${TODAY}.pdf`);
  };

  const [period, setPeriod] = useState("month"); const [breakdowns, setBreakdowns] = useState([]); const [logs, setLogs] = useState([]); const [loading, setLoading] = useState(true);
  useEffect(() => { loadData(); }, [period]);
  const loadData = async () => {
    setLoading(true);
    const now=new Date(); let fromDate;
    if (period==="month") fromDate=new Date(now.getFullYear(),now.getMonth(),1);
    else if (period==="quarter") fromDate=new Date(now.getFullYear(),now.getMonth()-3,1);
    else if (period==="half") fromDate=new Date(now.getFullYear(),now.getMonth()-6,1);
    else fromDate=new Date(2000,0,1);
    const fromStr=fromDate.toISOString();
    const [bRes,lRes]=await Promise.all([supabase.from("breakdown_reports").select("*").gte("reported_at",fromStr),supabase.from("maintenance_logs").select("*").gte("start_date",fromStr.split("T")[0])]);
    setBreakdowns(bRes.data||[]); setLogs(lRes.data||[]); setLoading(false);
  };

  const totalWOs=workOrders.length; const completedWOs=workOrders.filter(w => w.status==="Completed").length; const overdueWOs=workOrders.filter(w => w.due&&w.due<=TODAY&&w.status!=="Completed").length; const completionRate=totalWOs>0?Math.round((completedWOs/totalWOs)*100):0;
  const woByStatus=[{label:t(lang,"open"),count:workOrders.filter(w => w.status==="Open").length,color:C.accent},{label:"In Progress",count:workOrders.filter(w => w.status==="In Progress").length,color:C.blue},{label:"Pending",count:workOrders.filter(w => w.status==="Pending").length,color:C.yellow},{label:t(lang,"completed"),count:completedWOs,color:C.green}];
  const totalAssets=assets.length; const operationalAssets=assets.filter(a => a.status==="Operational").length; const downtimeAssets=assets.filter(a => a.status==="Under Maintenance").length; const degradedAssets=assets.filter(a => a.status==="Degraded").length; const uptimeRate=totalAssets>0?Math.round((operationalAssets/totalAssets)*100):0; const totalDowntimeMins=breakdowns.filter(b => b.downtime_hours).reduce((s,b) => s+(b.downtime_hours||0),0);
  const pmDue=assets.filter(a => { if (!a.pm_frequency) return false; if (!a.last_pm_date) return true; const now=new Date(); const last=new Date(a.last_pm_date); return (now.getFullYear()-last.getFullYear())*12+(now.getMonth()-last.getMonth())>=a.pm_frequency; }).length; const pmCompliance=totalAssets>0?Math.round(((totalAssets-pmDue)/totalAssets)*100):0; const pmLogs=logs.filter(l => l.log_type==="Preventive Maintenance").length; const correctiveLogs=logs.filter(l => l.log_type==="Corrective Repair").length;
  const activeVendors=vendors.filter(v => v.status==="Active").length; const vendorWOs=vendors.map(v => ({ name: v.name, open: workOrders.filter(w => w.vendor===v.name&&w.status!=="Completed").length, completed: workOrders.filter(w => w.vendor===v.name&&w.status==="Completed").length, rating: v.rating })).filter(v => v.open+v.completed>0).sort((a,b) => (b.open+b.completed)-(a.open+a.completed));
  const totalBreakdowns=breakdowns.length; const resolvedBreakdowns=breakdowns.filter(b => b.status==="Resolved").length; const avgDowntime=resolvedBreakdowns>0?Math.round(breakdowns.filter(b => b.downtime_hours).reduce((s,b) => s+(b.downtime_hours||0),0)/resolvedBreakdowns):0;
  const siteData=SITES.filter(s => s!=="— Select Site —").map(site => ({ site, total: assets.filter(a => a.location===site).length, down: assets.filter(a => a.location===site&&a.status==="Under Maintenance").length })).filter(s => s.total>0);

  const KpiCard = ({ icon, label, value, sub, color, percent }) => (
    <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20, borderLeft: `4px solid ${color}` }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div><div style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>{icon} {label}</div><div style={{ fontSize: 32, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{value}</div>{sub && <div style={{ fontSize: 12, color: C.subtle, marginTop: 4 }}>{sub}</div>}</div>
        {percent!==undefined && <div style={{ textAlign: "center" }}><div style={{ fontSize: 22, fontWeight: 800, color, fontFamily: "monospace" }}>{percent}%</div><div style={{ fontSize: 10, color: C.muted }}>rate</div></div>}
      </div>
      {percent!==undefined && <div style={{ marginTop: 12, background: C.border, borderRadius: 4, height: 6 }}><div style={{ background: color, width: `${percent}%`, height: 6, borderRadius: 4, transition: "width 0.6s" }} /></div>}
    </div>
  );

  const BarChart = ({ data, valueKey, labelKey, colorKey, title }) => {
    const max=Math.max(...data.map(d => d[valueKey]),1);
    return (
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.text, marginBottom: 16 }}>{title}</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {data.map((d,i) => (
            <div key={i}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 4 }}><span style={{ color: C.subtle }}>{d[labelKey]}</span><span style={{ color: d[colorKey]||C.accent, fontWeight: 700 }}>{d[valueKey]}</span></div>
              <div style={{ background: C.border, borderRadius: 4, height: 8 }}><div style={{ background: d[colorKey]||C.accent, width: `${(d[valueKey]/max)*100}%`, height: 8, borderRadius: 4, transition: "width 0.6s" }} /></div>
            </div>
          ))}
        </div>
      </div>
    );
  };

  const periodLabels = { month: t(lang,"thisMonth"), quarter: t(lang,"threeMonths"), half: t(lang,"sixMonths"), all: t(lang,"allTime") };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <div style={{ fontSize: 18, fontWeight: 700, color: C.text }}>{t(lang,"reportsKPI")}</div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={exportToExcel} style={{ background: C.green+"22", color: C.green, border: `1px solid ${C.green}44`, borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>{t(lang,"exportExcel")}</button>
            <button onClick={exportToPDF} style={{ background: C.red+"22", color: C.red, border: `1px solid ${C.red}44`, borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>{t(lang,"exportPDF")}</button>
          </div>
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          {["month","quarter","half","all"].map(val => <button key={val} onClick={() => setPeriod(val)} style={{ background: period===val?C.accent:C.card, color: period===val?"#fff":C.muted, border: `1px solid ${period===val?C.accent:C.border}`, borderRadius: 6, padding: "6px 14px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>{periodLabels[val]}</button>)}
        </div>
      </div>

      {loading ? <Spinner lang={lang} /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
          <div>
            <div style={{ fontSize: 13, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>{t(lang,"workOrdersSection")}</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14, marginBottom: 16 }}>
              <KpiCard icon="📋" label={t(lang,"openWorkOrders")} value={totalWOs} sub={`${overdueWOs} ${t(lang,"overdue")}`} color={C.accent} percent={completionRate} />
              <KpiCard icon="✅" label={t(lang,"completed")} value={completedWOs} color={C.green} />
              <KpiCard icon="⚠️" label={t(lang,"overdue")} value={overdueWOs} color={C.red} />
              <KpiCard icon="⏳" label="In Progress" value={workOrders.filter(w => w.status==="In Progress").length} color={C.blue} />
            </div>
            <BarChart title={t(lang,"workOrdersSection")} data={woByStatus} labelKey="label" valueKey="count" colorKey="color" />
          </div>
          <div>
            <div style={{ fontSize: 13, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>{t(lang,"assetPerformance")}</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14, marginBottom: 16 }}>
              <KpiCard icon="✅" label={t(lang,"operationalAssets")} value={`${operationalAssets}/${totalAssets}`} color={C.green} percent={uptimeRate} />
              <KpiCard icon="🚨" label={t(lang,"assetsDown")} value={downtimeAssets} color={C.red} />
              <KpiCard icon="⚠️" label={t(lang,"degradedStatus")} value={degradedAssets} color={C.yellow} />
              <KpiCard icon="⏱" label={t(lang,"totalDowntime")} value={formatDowntime(totalDowntimeMins)} color={C.purple} />
            </div>
            <BarChart title={t(lang,"assetsBySite")} data={siteData.map(s => ({ label: s.site, count: s.total, color: s.down>0?C.accent:C.green }))} labelKey="label" valueKey="count" colorKey="color" />
          </div>
          <div>
            <div style={{ fontSize: 13, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>{t(lang,"pmSection")}</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14 }}>
              <KpiCard icon="📊" label={t(lang,"pmDueMonth")} value={`${totalAssets-pmDue}/${totalAssets}`} color={C.green} percent={pmCompliance} />
              <KpiCard icon="⚠️" label={t(lang,"overdue")} value={pmDue} color={C.red} />
              <KpiCard icon="🔧" label={t(lang,"pmSection")} value={pmLogs} color={C.blue} />
              <KpiCard icon="🔨" label="Corrective" value={correctiveLogs} color={C.accent} />
            </div>
          </div>
          <div>
            <div style={{ fontSize: 13, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>{t(lang,"breakdownAnalysis")}</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14, marginBottom: 16 }}>
              <KpiCard icon="🚨" label={t(lang,"openBreakdowns")} value={totalBreakdowns} color={C.red} />
              <KpiCard icon="✅" label={t(lang,"resolved")} value={resolvedBreakdowns} color={C.green} percent={totalBreakdowns>0?Math.round((resolvedBreakdowns/totalBreakdowns)*100):0} />
              <KpiCard icon="⏱" label={t(lang,"totalDowntime")} value={formatDowntime(avgDowntime)} color={C.yellow} />
              <KpiCard icon="🔓" label={t(lang,"openBreakdowns")} value={totalBreakdowns-resolvedBreakdowns} color={C.accent} />
            </div>
            {breakdowns.length>0 && <BarChart title={t(lang,"breakdownAnalysis")} data={SITES.filter(s => s!=="— Select Site —").map(site => ({ label: site, count: breakdowns.filter(b => b.site===site).length, color: C.red })).filter(s => s.count>0)} labelKey="label" valueKey="count" colorKey="color" />}
          </div>
          <div>
            <div style={{ fontSize: 13, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>{t(lang,"vendorPerformance")}</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14, marginBottom: 16 }}>
              <KpiCard icon="🤝" label={t(lang,"activeVendors")} value={activeVendors} color={C.blue} />
              <KpiCard icon="🔓" label={t(lang,"openOrders")} value={vendors.reduce((s,v) => s+(v.open_orders||0),0)} color={C.accent} />
            </div>
            {vendorWOs.length>0 && (
              <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: C.text, marginBottom: 16 }}>{t(lang,"vendorPerformance")}</div>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead><tr style={{ borderBottom: `1px solid ${C.border}` }}>
                      {[t(lang,"vendors"),t(lang,"open"),t(lang,"completed"),"Total",t(lang,"rating")].map(h => <th key={h} style={{ textAlign: "left", padding: "8px 12px", fontSize: 11, color: C.muted, textTransform: "uppercase", fontWeight: 600 }}>{h}</th>)}
                    </tr></thead>
                    <tbody>
                      {vendorWOs.map((v,i) => (
                        <tr key={i} style={{ borderBottom: `1px solid ${C.border}22` }}>
                          <td style={{ padding: "10px 12px", fontSize: 13, color: C.text, fontWeight: 600 }}>{v.name}</td>
                          <td style={{ padding: "10px 12px" }}><Badge label={String(v.open)} color={v.open>0?C.accent:C.muted} /></td>
                          <td style={{ padding: "10px 12px" }}><Badge label={String(v.completed)} color={C.green} /></td>
                          <td style={{ padding: "10px 12px", fontSize: 13, color: C.text, fontWeight: 700 }}>{v.open+v.completed}</td>
                          <td style={{ padding: "10px 12px", fontSize: 13, color: C.yellow, fontWeight: 700 }}>{v.rating>0?`★ ${Number(v.rating).toFixed(1)}`:"N/A"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

function UserManagement({ lang }) {
  const [users, setUsers] = useState([]); const [loading, setLoading] = useState(true); const [showForm, setShowForm] = useState(false); const [saving, setSaving] = useState(false); const [error, setError] = useState(null); const [success, setSuccess] = useState(null);
  const [form, setForm] = useState({ email: "", name: "", role: "operations", site: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));

  useEffect(() => { loadUsers(); }, []);
  const loadUsers = async () => { setLoading(true); const { data } = await supabase.from("user_roles").select("*").order("name"); setUsers(data||[]); setLoading(false); };
  const submit = async () => {
    if (!form.email||!form.name) { setError(t(lang,"email")); return; }
    setSaving(true); setError(null);
    const record = { id: uid("USR"), email: form.email, name: form.name, role: form.role, site: form.site||null };
    const { error: err } = await supabase.from("user_roles").upsert([record], { onConflict: "email" });
    if (err) { setError(err.message); } else { setSuccess("✓"); setForm({ email: "", name: "", role: "operations", site: "" }); setShowForm(false); loadUsers(); }
    setSaving(false);
  };
  const deleteUser = async (id) => { await supabase.from("user_roles").delete().eq("id",id); setUsers(prev => prev.filter(u => u.id!==id)); };
  const roleColor = (r) => ({ admin: C.accent, maintenance: C.blue, operations: C.green }[r]||C.muted);
  const roleIcon = (r) => ({ admin: "★", maintenance: "🔧", operations: "🏭" }[r]||"👤");

  return (
    <div>
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <OkBanner msg={success} onDismiss={() => setSuccess(null)} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div>
          <div style={{ fontSize: 15, fontWeight: 700, color: C.text, marginBottom: 4 }}>{t(lang,"userManagement")}</div>
          <div style={{ fontSize: 13, color: C.muted }}>{t(lang,"manageUsers")}</div>
        </div>
        <Btn onClick={() => setShowForm(v => !v)}>{t(lang,"addUser")}</Btn>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12, marginBottom: 24 }}>
        {[["🏭",t(lang,"operationsRole"),"operations",t(lang,"operationsDesc"),C.green],["🔧",t(lang,"maintenanceRole"),"maintenance",t(lang,"maintenanceDesc"),C.blue],["★",t(lang,"adminRole"),"admin",t(lang,"adminDesc"),C.accent]].map(([icon,title,role,desc,color]) => (
          <div key={role} style={{ background: C.card, border: `1px solid ${color}44`, borderRadius: 8, padding: 14 }}>
            <div style={{ fontSize: 18, marginBottom: 6 }}>{icon}</div>
            <div style={{ fontSize: 13, fontWeight: 700, color }}>{title}</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{desc}</div>
          </div>
        ))}
      </div>
      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 20 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>{t(lang,"addUpdateUser")}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label={t(lang,"email")} value={form.email} onChange={f("email")} type="email" />
            <Input label={t(lang,"fullName")} value={form.name} onChange={f("name")} />
            <Sel label={t(lang,"role")} value={form.role} onChange={f("role")} options={["operations","maintenance","admin"]} />
            <Sel label={t(lang,"defaultSite")} value={form.site||"— Select Site —"} onChange={f("site")} options={["— Select Site —",...SITES.filter(s => s!=="— Select Site —")]} />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
            <Btn onClick={submit} disabled={saving}>{saving?t(lang,"saving"):t(lang,"saveUser")}</Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      )}
      {loading ? <Spinner lang={lang} /> : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 12 }}>
          {users.map(u => (
            <div key={u.id} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 18 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
                <div><div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{u.name}</div><div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>{u.email}</div></div>
                <span style={{ background: roleColor(u.role)+"22", color: roleColor(u.role), border: `1px solid ${roleColor(u.role)}44`, borderRadius: 4, padding: "3px 10px", fontSize: 12, fontWeight: 700, textTransform: "uppercase" }}>{roleIcon(u.role)} {u.role}</span>
              </div>
              {u.site && <div style={{ fontSize: 12, color: C.muted, marginBottom: 12 }}>📍 {u.site}</div>}
              <Btn small variant="danger" onClick={() => deleteUser(u.id)}>{t(lang,"remove")}</Btn>
            </div>
          ))}
          {users.length===0 && <div style={{ color: C.muted, fontSize: 13 }}>{t(lang,"noUsersRegistered")}</div>}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState(null); const [authLoading, setAuthLoading] = useState(true);
  const [userRole, setUserRole] = useState({ role: "operations", name: "", site: "", language: "en" });
  const [lang, setLang] = useState("en");
  const [tab, setTab] = useState(null);
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
        if (data) { setUserRole(data); const l = data.language || "en"; setLang(l); document.dir = l === "ar" ? "rtl" : "ltr"; }
        else setUserRole({ role: session.user.email === ADMIN_EMAIL ? "admin" : "operations", name: session.user.email, site: "", language: "en" });
      });
    }
  }, [session]);

  const toggleLanguage = async () => {
    const newLang = lang === "en" ? "ar" : "en";
    setLang(newLang);
    setTab(null); // reset tab to prevent mismatch
    document.dir = newLang === "ar" ? "rtl" : "ltr";
    if (session) {
      await supabase.from("user_roles").update({ language: newLang }).eq("email", session.user.email);
      setUserRole(prev => ({ ...prev, language: newLang }));
    }
  };

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
  if (!session) return <LoginScreen lang={lang} />;

  const roleColor = { admin: C.accent, maintenance: C.blue, operations: C.green }[userRole.role] || C.muted;
  const roleIcon = { admin: "★", maintenance: "🔧", operations: "🏭" }[userRole.role] || "👤";
  const roleLabel = { admin: t(lang,"adminRole"), maintenance: t(lang,"maintenanceRole"), operations: t(lang,"operationsRole") }[userRole.role] || userRole.role;

  const tabs = [
    t(lang,"overview"),
    t(lang,"breakdowns"),
    ...(userRole.role !== "operations" ? [t(lang,"workOrders"), t(lang,"assets"), t(lang,"vendors"), t(lang,"pmPlanner"), t(lang,"reports"), t(lang,"calendar")] : []),
    ...(isAdmin ? [t(lang,"users")] : []),
  ];

  const activeTab = tab || tabs[0];

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'DM Sans', sans-serif", color: C.text }}>
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "0 16px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", minHeight: 58, flexWrap: "wrap", gap: 8, padding: "8px 0" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ background: C.accent, borderRadius: 8, width: 34, height: 34, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>🏭</div>
            <div>
              <div style={{ fontFamily: "monospace", fontSize: 16, letterSpacing: 2, color: C.text, fontWeight: 800 }}>{t(lang,"appName")}</div>
              <div style={{ fontSize: 9, color: C.muted }}>{t(lang,"appSub")}</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <button onClick={load} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 10px", color: C.muted, cursor: "pointer", fontSize: 12 }}>{t(lang,"refresh")}</button>
            <button onClick={toggleLanguage} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 10px", color: C.text, cursor: "pointer", fontSize: 12, fontWeight: 700 }}>
              {lang === "en" ? "🇸🇦 العربية" : "🇬🇧 English"}
            </button>
            <span style={{ background: roleColor+"22", color: roleColor, border: `1px solid ${roleColor}44`, borderRadius: 4, padding: "3px 8px", fontSize: 11, fontWeight: 700, textTransform: "uppercase" }}>{roleIcon} {roleLabel}</span>
            <div style={{ fontSize: 11, color: C.muted, maxWidth: 130, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{userRole.name || session.user.email}</div>
            <button onClick={signOut} style={{ background: "transparent", border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 10px", color: C.muted, cursor: "pointer", fontSize: 12 }}>{t(lang,"signOut")}</button>
          </div>
        </div>
        <div style={{ display: "flex", overflowX: "auto" }}>
          {tabs.map(tb => <button key={tb} onClick={() => setTab(tb)} style={{ background: "transparent", border: "none", padding: "10px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap", color: activeTab===tb?C.accent:C.muted, borderBottom: `2px solid ${activeTab===tb?C.accent:"transparent"}` }}>{tb}</button>)}
        </div>
      </div>
      <div style={{ padding: "20px 16px", maxWidth: 1280, margin: "0 auto" }}>
        <ErrBanner msg={globalError} onDismiss={() => setGlobalError(null)} />
        {activeTab===t(lang,"overview") && <Overview workOrders={workOrders} assets={assets} vendors={vendors} lang={lang} />}
        {activeTab===t(lang,"breakdowns") && <Breakdowns userRole={userRole} assets={assets} setAssets={setAssets} vendors={vendors} workOrders={workOrders} setWorkOrders={setWorkOrders} lang={lang} />}
        {activeTab===t(lang,"workOrders") && <WorkOrders workOrders={workOrders} setWorkOrders={setWorkOrders} loading={loading.workOrders} onAdd={r => setWorkOrders(p => [r,...p])} isAdmin={isAdmin} vendors={vendors} assets={assets} lang={lang} />}
        {activeTab===t(lang,"assets") && <Assets assets={assets} setAssets={setAssets} loading={loading.assets} onAdd={r => setAssets(p => [r,...p])} isAdmin={isAdmin} vendors={vendors} lang={lang} />}
        {activeTab===t(lang,"vendors") && <Vendors vendors={vendors} setVendors={setVendors} loading={loading.vendors} onAdd={r => setVendors(p => [r,...p])} isAdmin={isAdmin} lang={lang} />}
        {activeTab===t(lang,"pmPlanner") && <PMUpload assets={assets} onAssetsImported={r => setAssets(p => [...p,...r])} onWorkOrdersGenerated={r => setWorkOrders(p => [...r,...p])} lang={lang} />}
        {activeTab===t(lang,"reports") && <Reports workOrders={workOrders} assets={assets} vendors={vendors} lang={lang} />}
        {activeTab===t(lang,"calendar") && <MaintenanceCalendar workOrders={workOrders} assets={assets} lang={lang} />}
        {activeTab===t(lang,"users") && isAdmin && <UserManagement lang={lang} />}
      </div>
    </div>
  );
}



