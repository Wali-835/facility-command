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

// Fallback only — used if the `sites` table hasn't been migrated/seeded yet or fails to load.
const DEFAULT_SITES = [
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
// Postgres unique-violation (23505) on assets.serial_number -> friendly message; anything else -> raw DB message.
const assetDbErrorMessage = (err, lang) => err?.code === "23505" ? t(lang,"duplicateSerialNumber") : err.message;
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
  try {
    let normalized = d.replace(" ", "T");
    if (!/[Zz]|[+-]\d{2}:?\d{2}$/.test(normalized)) normalized += "Z";
    const date = new Date(normalized);
    if (isNaN(date.getTime())) return "—";
    return date.toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit", timeZone: "Africa/Cairo" });
  } catch { return "—"; }
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
// ─── ISSUE REPORT MODAL ───────────────────────────────────────────────────────
function IssueReportModal({ asset, userRole, onClose, onReported, onWorkOrderCreated, lang }) {
  const [form, setForm] = useState({ description: "", severity: "Medium", reported_by: userRole.name || "" });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));

  const submit = async () => {
    if (!form.description) { setError(t(lang,"issueDescription")); return; }
    if (!form.reported_by) { setError(t(lang,"yourName")); return; }
    setSaving(true); setError(null);
    const now = new Date().toISOString();
    const issueId = uid("ISS");
    const woId = uid("WO");

    // Create issue record
    const record = {
      id: issueId, asset_id: asset.id, asset_name: asset.name, site: asset.location,
      reported_by: form.reported_by, reported_at: now,
      description: form.description, severity: form.severity, status: "Open",
      work_order_id: woId,
    };
    const { error: err } = await supabase.from("issue_reports").insert([record]);
    if (err) { setError(err.message); setSaving(false); return; }

    // Auto-create work order
    const newWO = {
      id: woId, title: `Issue — ${asset.name}: ${form.description.slice(0,50)}`,
      asset: asset.name, priority: form.severity === "Critical" ? "Critical" : form.severity === "High" ? "High" : "Medium",
      status: "Open", assignee: null,
      start_date: now.split("T")[0], due: null, vendor: null,
    };
    await supabase.from("work_orders").insert([newWO]);
    if (onWorkOrderCreated) onWorkOrderCreated(newWO);

    // Send email notification
    try {
      await fetch("https://evwsdzqgvrwbjusjmrdc.supabase.co/functions/v1/notify-breakdown", {
        method: "POST",
        headers: { "Content-Type": "application/json", "Authorization": `Bearer ${import.meta.env.VITE_SUPABASE_ANON_KEY}` },
        body: JSON.stringify({ breakdown: { ...record, type: "issue" }, type: "reported" }),
      });
    } catch (e) { console.error("Email error:", e); }

    onReported(record);
    onClose();
    setSaving(false);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2000, padding: 16 }}>
      <div style={{ background: C.card, border: `2px solid ${C.yellow}44`, borderRadius: 12, width: "100%", maxWidth: 500 }}>
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, background: C.yellow+"11" }}>
          <div style={{ fontSize: 17, fontWeight: 700, color: C.yellow }}>⚠️ {t(lang,"reportIssue")}</div>
          <div style={{ fontSize: 13, color: C.muted, marginTop: 4 }}>{asset.name} · {asset.location}</div>
          <div style={{ fontSize: 12, color: C.yellow, marginTop: 6, background: C.yellow+"22", borderRadius: 6, padding: "6px 10px", display: "inline-block" }}>
            {t(lang,"equipmentStillRunning")}
          </div>
        </div>
        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Input label={t(lang,"yourName")} value={form.reported_by} onChange={f("reported_by")} />
            <Sel label={t(lang,"severity")} value={form.severity} onChange={f("severity")} options={["Critical","High","Medium","Low"]} />
            <Textarea label={t(lang,"issueDescription")} value={form.description} onChange={f("description")} placeholder="Describe the issue — noise, leak, warning light, performance drop..." />
            <div style={{ background: C.green+"11", border: `1px solid ${C.green}44`, borderRadius: 8, padding: 12, fontSize: 12, color: C.green }}>
              ✅ Equipment status will NOT change — it remains Operational
            </div>
            <div style={{ background: C.blue+"11", border: `1px solid ${C.blue}44`, borderRadius: 8, padding: 12, fontSize: 12, color: C.blue }}>
              📋 A work order will be automatically created for maintenance
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
            <Btn onClick={submit} disabled={saving} color={C.yellow}>{saving ? t(lang,"reporting") : `⚠️ ${t(lang,"reportIssue")}`}</Btn>
            <Btn variant="secondary" onClick={onClose}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}
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

// ─── ADD UPDATE MODAL (append a note to an already-open breakdown/issue) ──────
function AddUpdateModal({ item, table, userRole, lang, onClose, onUpdated }) {
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);

  const submit = async () => {
    if (!note.trim()) { setError(t(lang,"addUpdateNote")); return; }
    setSaving(true); setError(null);
    const entry = { note: note.trim(), by: userRole?.name || "—", at: new Date().toISOString() };
    const newUpdates = [...(item.updates||[]), entry];
    const { error: err } = await supabase.from(table).update({ updates: newUpdates }).eq("id", item.id);
    if (err) { setError(err.message); setSaving(false); return; }
    onUpdated({ ...item, updates: newUpdates });
    onClose();
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2000, padding: 16 }}>
      <div style={{ background: C.card, border: `2px solid ${C.blue}44`, borderRadius: 12, width: "100%", maxWidth: 480 }}>
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, background: C.blue+"11" }}>
          <div style={{ fontSize: 17, fontWeight: 700, color: C.blue }}>📝 {t(lang,"addUpdate")}</div>
          <div style={{ fontSize: 13, color: C.muted, marginTop: 4 }}>{item.asset_name} · {item.status}</div>
        </div>
        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <Textarea label={t(lang,"addUpdateNote")} value={note} onChange={setNote} placeholder={t(lang,"addUpdateNotePlaceholder")} />
          <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
            <Btn onClick={submit} disabled={saving} color={C.blue}>{saving ? t(lang,"reporting") : t(lang,"addUpdate")}</Btn>
            <Btn variant="secondary" onClick={onClose}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

const UpdatesLog = ({ updates, lang }) => (!updates || updates.length === 0) ? null : (
  <div style={{ background: C.surface, borderRadius: 8, padding: 12, marginBottom: 14 }}>
    <div style={{ fontSize: 11, color: C.muted, textTransform: "uppercase", marginBottom: 8 }}>📝 {t(lang,"updatesLog")}</div>
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      {updates.map((u,i) => (
        <div key={i} style={{ fontSize: 12, color: C.subtle }}>
          <span style={{ color: C.text, fontWeight: 600 }}>{u.by}</span> · <span style={{ color: C.muted }}>{fmtDateTime(u.at)}</span>
          <div style={{ marginTop: 2 }}>{u.note}</div>
        </div>
      ))}
    </div>
  </div>
);

// ─── PASSWORD RE-AUTH (required before any critical status change) ───────────
function PasswordConfirm({ lang, actionLabel, onConfirmed, color = C.green, disabled }) {
  const [password, setPassword] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [show, setShow] = useState(false);

  const confirm = async () => {
    if (!password) { setError(t(lang,"enterPasswordToApprove")); return; }
    setSaving(true); setError(null);
    const { data: sessionData } = await supabase.auth.getSession();
    const email = sessionData?.session?.user?.email;
    const { error: authErr } = await supabase.auth.signInWithPassword({ email, password });
    if (authErr) { setError(t(lang,"incorrectPassword")); setSaving(false); return; }
    await onConfirmed();
    setSaving(false);
  };

  return !show ? (
    <Btn onClick={() => setShow(true)} disabled={disabled} color={color}>{actionLabel}</Btn>
  ) : (
    <div>
      <Input label={t(lang,"confirmPassword")} value={password} type="password" onChange={v => { setPassword(v); setError(null); }} />
      {error && <div style={{ color: C.red, fontSize: 12, marginTop: 6 }}>{error}</div>}
      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <Btn onClick={confirm} disabled={saving||!password} color={color}>{saving?"Verifying...":actionLabel}</Btn>
        <Btn variant="secondary" onClick={() => { setShow(false); setPassword(""); }}>{t(lang,"cancel")}</Btn>
      </div>
    </div>
  );
}

// ─── BREAKDOWN RESOLUTION MODAL ───────────────────────────────────────────────
function BreakdownResolveModal({ breakdown, userRole, vendors, onClose, onResolved, lang }) {
  const [form, setForm] = useState({ resolved_by: userRole.name || "", maintenance_notes: "", vendor: "" });
  const [error, setError] = useState(null);
  const [confirming, setConfirming] = useState(false);
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const vendorOptions = ["— None —", ...vendors.filter(v => v.status === "Active").map(v => v.name)];
  const downtimeMins = minutesBetween(breakdown.downtime_start, new Date().toISOString());

  const tryProceed = () => {
    if (!form.resolved_by) { setError(t(lang,"yourName")); return; }
    if (!form.maintenance_notes) { setError(t(lang,"maintenanceNotes")); return; }
    setError(null);
    setConfirming(true);
  };

  const submit = async () => {
    const now = new Date().toISOString();
    const hours = minutesBetween(breakdown.downtime_start, now);
    const isSupervisorOrAdmin = userRole?.role === "supervisor" || userRole?.role === "admin";
    const newStatus = isSupervisorOrAdmin ? "Pending Operator Confirmation" : "Pending Supervisor Approval";
    await supabase.from("breakdown_reports").update({ status: newStatus, resolved_by: form.resolved_by, resolved_at: now, downtime_end: now, downtime_hours: hours, maintenance_notes: form.maintenance_notes, supervisor_approved_by: isSupervisorOrAdmin ? userRole?.name : null, supervisor_approved_at: isSupervisorOrAdmin ? now : null }).eq("id", breakdown.id);
    // Asset stays Under Maintenance until Operator gives final confirmation
    const isSupervisorOrAdmin2 = userRole?.role === "supervisor" || userRole?.role === "admin";
    const logRecord = { id: uid("LOG"), asset_id: breakdown.asset_id, asset_name: breakdown.asset_name, log_type: "Corrective Repair", title: `Breakdown Repair — ${breakdown.severity} severity`, description: `BREAKDOWN REPORTED BY: ${breakdown.reported_by}\n\nISSUE: ${breakdown.description}\n\nMAINTENANCE NOTES: ${form.maintenance_notes}`, performed_by: form.resolved_by, vendor: form.vendor === "— None —" ? null : form.vendor || null, start_date: breakdown.downtime_start ? breakdown.downtime_start.split("T")[0] : TODAY, end_date: TODAY, cost: null, status: isSupervisorOrAdmin2 ? "Completed" : "In Progress", approval_status: isSupervisorOrAdmin2 ? "Approved" : "Pending", approved_by: isSupervisorOrAdmin2 ? userRole?.name : null, approved_at: isSupervisorOrAdmin2 ? new Date().toISOString() : null, breakdown_id: breakdown.id, downtime_start: breakdown.downtime_start ? breakdown.downtime_start.split("T")[0] : null, downtime_end: TODAY, downtime_hours: hours };
    await supabase.from("maintenance_logs").insert([logRecord]);
    const { data: openWOs } = await supabase.from("work_orders").select("id").eq("asset", breakdown.asset_name).in("status", ["Open","In Progress","Pending"]);
    if (openWOs?.length) await supabase.from("work_orders").update({ status: "Completed" }).in("id", openWOs.map(w => w.id));
    onResolved({ ...breakdown, status: "Resolved", downtime_end: now, downtime_hours: hours });
    onClose();
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
          <div style={{ marginTop: 20 }}>
            {confirming ? (
              <PasswordConfirm lang={lang} actionLabel={`✅ ${t(lang,"markResolved")}`} onConfirmed={submit} />
            ) : (
              <div style={{ display: "flex", gap: 8 }}>
                <Btn onClick={tryProceed} color={C.green}>✅ {t(lang,"markResolved")}</Btn>
                <Btn variant="secondary" onClick={onClose}>{t(lang,"cancel")}</Btn>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function BreakdownApprovalStep({ breakdown, userRole, lang, onApproved }) {
  const approve = async () => {
    const now = new Date().toISOString();
    await supabase.from("breakdown_reports").update({ status: "Pending Operator Confirmation", supervisor_approved_by: userRole?.name, supervisor_approved_at: now }).eq("id", breakdown.id);
    onApproved({ ...breakdown, status: "Pending Operator Confirmation", supervisor_approved_by: userRole?.name, supervisor_approved_at: now });
  };
  return (
    <div style={{ background: C.yellow+"11", border: `1px solid ${C.yellow}44`, borderRadius: 8, padding: 14, marginBottom: 14 }}>
      <div style={{ fontSize: 12, color: C.yellow, fontWeight: 700, marginBottom: 10 }}>⏳ {t(lang,"pendingSupervisorApproval")}</div>
      <PasswordConfirm lang={lang} actionLabel="✓ Confirm & Approve" onConfirmed={approve} />
    </div>
  );
}

function BreakdownOperatorConfirm({ breakdown, userRole, lang, onConfirmed }) {
  const isReporter = userRole?.name === breakdown.reported_by;

  const confirm = async () => {
    const now = new Date().toISOString();
    await supabase.from("breakdown_reports").update({ status: "Resolved", operator_confirmed_by: userRole?.name, operator_confirmed_at: now }).eq("id", breakdown.id);
    await supabase.from("assets").update({ status: "Operational" }).eq("id", breakdown.asset_id);
    await supabase.from("maintenance_logs").update({ approval_status: "Approved", approved_by: breakdown.supervisor_approved_by || userRole?.name, approved_at: breakdown.supervisor_approved_at || now, status: "Completed" }).eq("breakdown_id", breakdown.id);
    onConfirmed({ ...breakdown, status: "Resolved", operator_confirmed_by: userRole?.name, operator_confirmed_at: now });
  };

  return (
    <div style={{ background: C.blue+"11", border: `1px solid ${C.blue}44`, borderRadius: 8, padding: 14, marginBottom: 14 }}>
      <div style={{ fontSize: 12, color: C.blue, fontWeight: 700, marginBottom: 10 }}>👁 {t(lang,"pendingOperatorConfirmation")}</div>
      {isReporter ? (
        <PasswordConfirm lang={lang} actionLabel={t(lang,"operatorConfirm")} onConfirmed={confirm} />
      ) : (
        <div style={{ fontSize: 12, color: C.muted }}>{t(lang,"awaitingYourConfirmation")} ({breakdown.reported_by})</div>
      )}
    </div>
  );
}

function IssueApprovalStep({ issue, userRole, lang, onApproved }) {
  const approve = async () => {
    const now = new Date().toISOString();
    await supabase.from("issue_reports").update({ status: "Pending Operator Confirmation", supervisor_approved_by: userRole?.name, supervisor_approved_at: now }).eq("id", issue.id);
    if (issue.work_order_id) await supabase.from("work_orders").update({ status: "Completed" }).eq("id", issue.work_order_id);
    onApproved({ ...issue, status: "Pending Operator Confirmation", supervisor_approved_by: userRole?.name, supervisor_approved_at: now });
  };
  return (
    <div style={{ background: C.yellow+"11", border: `1px solid ${C.yellow}44`, borderRadius: 8, padding: 14, marginBottom: 12 }}>
      <div style={{ fontSize: 12, color: C.yellow, fontWeight: 700, marginBottom: 10 }}>⏳ {t(lang,"pendingSupervisorApproval")}</div>
      <PasswordConfirm lang={lang} actionLabel="✓ Confirm & Approve" onConfirmed={approve} />
    </div>
  );
}

function IssueOperatorConfirm({ issue, userRole, lang, onConfirmed }) {
  const isReporter = userRole?.name === issue.reported_by;

  const confirm = async () => {
    const now = new Date().toISOString();
    await supabase.from("issue_reports").update({ status: "Resolved", operator_confirmed_by: userRole?.name, operator_confirmed_at: now }).eq("id", issue.id);
    onConfirmed({ ...issue, status: "Resolved", operator_confirmed_by: userRole?.name, operator_confirmed_at: now });
  };

  return (
    <div style={{ background: C.blue+"11", border: `1px solid ${C.blue}44`, borderRadius: 8, padding: 14, marginBottom: 12 }}>
      <div style={{ fontSize: 12, color: C.blue, fontWeight: 700, marginBottom: 10 }}>👁 {t(lang,"pendingOperatorConfirmation")}</div>
      {isReporter ? (
        <PasswordConfirm lang={lang} actionLabel={t(lang,"operatorConfirm")} onConfirmed={confirm} />
      ) : (
        <div style={{ fontSize: 12, color: C.muted }}>{t(lang,"awaitingYourConfirmation")} ({issue.reported_by})</div>
      )}
    </div>
  );
}

// ─── BREAKDOWNS TAB ───────────────────────────────────────────────────────────
function Breakdowns({ userRole, assets, setAssets, vendors, workOrders, setWorkOrders, lang, setIssuesFromParent, isMaintenance, isSupervisor }) {
  const [breakdowns, setBreakdowns] = useState([]);
  const [issues, setIssues] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [showReportForm, setShowReportForm] = useState(false);
  const [showIssueForm, setShowIssueForm] = useState(false);
  const [resolveItem, setResolveItem] = useState(null);
  const [filter, setFilter] = useState("Open");
  const [activeView, setActiveView] = useState("breakdowns");
  const [selectedAsset, setSelectedAsset] = useState(null);
  const [selectedIssueAsset, setSelectedIssueAsset] = useState(null);
  const [updateTarget, setUpdateTarget] = useState(null); // { item, table }

  useEffect(() => { loadAll(); }, []);

  // An asset can have at most one open (non-Resolved) breakdown or issue at a time.
  const openReportFor = (assetId) => {
    const b = breakdowns.find(x => x.asset_id === assetId && x.status !== "Resolved");
    if (b) return { item: b, table: "breakdown_reports" };
    const i = issues.find(x => x.asset_id === assetId && x.status !== "Resolved");
    if (i) return { item: i, table: "issue_reports" };
    return null;
  };

  const loadAll = async () => {
    setLoading(true);
    const [bRes, iRes] = await Promise.all([
      supabase.from("breakdown_reports").select("*").order("reported_at", { ascending: false }),
      supabase.from("issue_reports").select("*").order("reported_at", { ascending: false }),
    ]);
    setBreakdowns(bRes.data || []);
    setIssues(iRes.data || []);
    setLoading(false);
  };

  const filtered = filter === "All" ? breakdowns : breakdowns.filter(b => b.status === filter);
  const filteredIssues = filter === "All" ? issues : issues.filter(i => i.status === filter);
  const openCount = breakdowns.filter(b => b.status === "Open").length;
  const acknowledgedCount = breakdowns.filter(b => b.status === "Acknowledged").length;
  const resolvedCount = breakdowns.filter(b => b.status === "Resolved").length;
  const totalDowntimeMins = breakdowns.filter(b => b.downtime_hours).reduce((s, b) => s + (b.downtime_hours || 0), 0);
  const openIssues = issues.filter(i => i.status === "Open").length;
  const resolvedIssues = issues.filter(i => i.status === "Resolved").length;

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
const onIssueReported = (record) => {
    setIssues(prev => { const updated = [record, ...prev]; if (setIssuesFromParent) setIssuesFromParent(updated); return updated; });
    setSuccess(t(lang,"issueReported"));
  };

  const resolveIssue = async (issue) => {
    const now = new Date().toISOString();
    const newStatus = userRole?.role === "maintenance" ? "Pending Supervisor Approval" : "Resolved";
    await supabase.from("issue_reports").update({ status: newStatus, resolved_by: userRole.name || "", resolved_at: now, supervisor_approved_by: newStatus === "Resolved" ? userRole?.name : null, supervisor_approved_at: newStatus === "Resolved" ? now : null }).eq("id", issue.id);
    if (issue.work_order_id && newStatus === "Resolved") await supabase.from("work_orders").update({ status: "Completed" }).eq("id", issue.work_order_id);

    await supabase.from("maintenance_logs").insert([{
      id: uid("LOG"),
      asset_id: issue.asset_id,
      asset_name: issue.asset_name,
      log_type: "Corrective Repair",
      title: `Issue Resolved — ${issue.severity} severity`,
      description: `ISSUE REPORTED BY: ${issue.reported_by}\n\nDESCRIPTION: ${issue.description}\n\nRESOLVED BY: ${userRole.name || ""}`,
      performed_by: userRole.name || "",
      vendor: null,
      start_date: issue.reported_at ? issue.reported_at.split("T")[0] : TODAY,
      end_date: TODAY,
      cost: null,
      status: "Completed",
      downtime_start: null,
      downtime_end: null,
      downtime_hours: null,
    }]);

    setIssues(prev => { const updated = prev.map(i => i.id === issue.id ? { ...i, status: newStatus, resolved_by: userRole.name, resolved_at: now } : i); if (setIssuesFromParent) setIssuesFromParent(updated); return updated; });
    setSuccess(newStatus === "Resolved" ? t(lang,"resolved") : t(lang,"pendingSupervisorApproval"));
  };

  const acknowledgeIssue = async (issue) => {
    const now = new Date().toISOString();
    await supabase.from("issue_reports").update({ status: "Acknowledged", acknowledged_by: userRole.name || "", acknowledged_at: now }).eq("id", issue.id);
    setIssues(prev => prev.map(i => i.id === issue.id ? { ...i, status: "Acknowledged", acknowledged_by: userRole.name, acknowledged_at: now } : i));
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
      {showIssueForm && selectedIssueAsset && (
        <IssueReportModal asset={selectedIssueAsset} userRole={userRole} lang={lang} onClose={() => { setShowIssueForm(false); setSelectedIssueAsset(null); }} onReported={onIssueReported} onWorkOrderCreated={wo => setWorkOrders(prev => [wo, ...prev])} />
      )}
      {resolveItem && (
        <BreakdownResolveModal breakdown={resolveItem} userRole={userRole} vendors={vendors} lang={lang} onClose={() => setResolveItem(null)} onResolved={onResolved} />
      )}
      {updateTarget && (
        <AddUpdateModal item={updateTarget.item} table={updateTarget.table} userRole={userRole} lang={lang} onClose={() => setUpdateTarget(null)}
          onUpdated={(updated) => {
            if (updateTarget.table === "breakdown_reports") setBreakdowns(prev => prev.map(x => x.id===updated.id?updated:x));
            else setIssues(prev => { const next = prev.map(x => x.id===updated.id?updated:x); if (setIssuesFromParent) setIssuesFromParent(next); return next; });
          }} />
      )}
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <OkBanner msg={success} onDismiss={() => setSuccess(null)} />

      {/* Stats */}
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        <StatCard icon="🚨" label={t(lang,"openBreakdowns")} value={openCount} sub={t(lang,"needsAttention")} color={C.red} />
        <StatCard icon="⚠️" label={t(lang,"openIssues")} value={openIssues} sub={t(lang,"equipmentStillRunning")} color={C.yellow} />
        <StatCard icon="👁" label={t(lang,"acknowledged")} value={acknowledgedCount} sub={t(lang,"beingHandled")} color={C.blue} />
        <StatCard icon="✅" label={t(lang,"resolved")} value={resolvedCount + resolvedIssues} sub={t(lang,"thisPeriod")} color={C.green} />
        <StatCard icon="⏱" label={t(lang,"totalDowntime")} value={formatDowntime(totalDowntimeMins)} sub={t(lang,"allBreakdowns")} color={C.purple} />
        <StatCard icon="🏭" label={t(lang,"assetsDown")} value={assets.filter(a => a.status === "Under Maintenance").length} sub={t(lang,"currentlyOffline")} color={C.accent} />
      </div>

      {/* Report Buttons */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 20 }}>
        <div style={{ background: C.card, border: `1px solid ${C.red}33`, borderRadius: 10, padding: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: C.red, marginBottom: 8 }}>🚨 {t(lang,"reportEquipmentBreakdown")}</div>
          <div style={{ fontSize: 12, color: C.muted, marginBottom: 10 }}>Equipment is DOWN — downtime starts now</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <select onChange={e => setSelectedAsset(assets.find(a => a.id === e.target.value) || null)}
              style={{ flex: 1, minWidth: 150, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "8px", color: C.text, fontSize: 13 }}>
              <option value="">{t(lang,"selectEquipment")}</option>
              {assets.map(a => <option key={a.id} value={a.id}>{a.name} ({a.location}){openReportFor(a.id) ? ` — ${t(lang,"alreadyReportedTag")}` : ""}</option>)}
            </select>
            <Btn onClick={() => {
              if (!selectedAsset) { setError(t(lang,"selectEquipment")); return; }
              const existing = openReportFor(selectedAsset.id);
              if (existing) { setUpdateTarget(existing); return; }
              setShowReportForm(true);
            }} color={C.red}>🚨 {t(lang,"reportBreakdown")}</Btn>
          </div>
          {selectedAsset && openReportFor(selectedAsset.id) && (
            <div style={{ marginTop: 10, fontSize: 12, color: C.yellow }}>⚠️ {t(lang,"assetHasOpenReport")}</div>
          )}
        </div>
        <div style={{ background: C.card, border: `1px solid ${C.yellow}33`, borderRadius: 10, padding: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: C.yellow, marginBottom: 8 }}>⚠️ {t(lang,"reportEquipmentIssue")}</div>
          <div style={{ fontSize: 12, color: C.muted, marginBottom: 10 }}>Equipment is still RUNNING — issue needs attention</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <select onChange={e => setSelectedIssueAsset(assets.find(a => a.id === e.target.value) || null)}
              style={{ flex: 1, minWidth: 150, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "8px", color: C.text, fontSize: 13 }}>
              <option value="">{t(lang,"selectEquipment")}</option>
              {assets.map(a => <option key={a.id} value={a.id}>{a.name} ({a.location}){openReportFor(a.id) ? ` — ${t(lang,"alreadyReportedTag")}` : ""}</option>)}
            </select>
            <Btn onClick={() => {
              if (!selectedIssueAsset) { setError(t(lang,"selectEquipment")); return; }
              const existing = openReportFor(selectedIssueAsset.id);
              if (existing) { setUpdateTarget(existing); return; }
              setShowIssueForm(true);
            }} color={C.yellow}>⚠️ {t(lang,"reportIssue")}</Btn>
          </div>
          {selectedIssueAsset && openReportFor(selectedIssueAsset.id) && (
            <div style={{ marginTop: 10, fontSize: 12, color: C.yellow }}>⚠️ {t(lang,"assetHasOpenReport")}</div>
          )}
        </div>
      </div>

      {/* View Toggle */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", gap: 6 }}>
          <button onClick={() => setActiveView("breakdowns")} style={{ background: activeView==="breakdowns"?C.red:C.card, color: activeView==="breakdowns"?"#fff":C.muted, border: `1px solid ${activeView==="breakdowns"?C.red:C.border}`, borderRadius: 6, padding: "6px 14px", fontSize: 12, cursor: "pointer", fontWeight: 700 }}>🚨 {t(lang,"breakdowns")} ({breakdowns.filter(b=>b.status==="Open").length})</button>
          <button onClick={() => setActiveView("issues")} style={{ background: activeView==="issues"?C.yellow:C.card, color: activeView==="issues"?"#fff":C.muted, border: `1px solid ${activeView==="issues"?C.yellow:C.border}`, borderRadius: 6, padding: "6px 14px", fontSize: 12, cursor: "pointer", fontWeight: 700 }}>⚠️ {t(lang,"issues")} ({issues.filter(i=>i.status==="Open").length})</button>
        </div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {["All","Open","Acknowledged","Pending Supervisor Approval","Pending Operator Confirmation","Resolved"].map(s => (
            <button key={s} onClick={() => setFilter(s)} style={{ background: filter===s?C.accent:C.card, color: filter===s?"#fff":C.muted, border: `1px solid ${filter===s?C.accent:C.border}`, borderRadius: 6, padding: "6px 10px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>
              {s === "All" ? t(lang,"all") : s === "Open" ? t(lang,"open") : s === "Acknowledged" ? t(lang,"acknowledged") : s === "Resolved" ? t(lang,"resolved") : s === "Pending Supervisor Approval" ? t(lang,"pendingSupervisorApproval") : t(lang,"pendingOperatorConfirmation")}
            </button>
          ))}
          <button onClick={loadAll} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "6px 10px", color: C.muted, cursor: "pointer", fontSize: 12 }}>↻</button>
        </div>
      </div>

      {activeView === "breakdowns" && (loading ? <Spinner lang={lang} /> : filtered.length === 0 ? (
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
                {b.maintenance_notes && (
                  <div style={{ background: C.blue+"11", border: `1px solid ${C.blue}33`, borderRadius: 8, padding: 12, marginBottom: 14, fontSize: 13, color: C.subtle }}>
                    <strong style={{ color: C.blue }}>🔧 {t(lang,"resolvedBy")} {b.resolved_by}:</strong> {b.maintenance_notes}
                  </div>
                )}
                <UpdatesLog updates={b.updates} lang={lang} />
                {b.status !== "Resolved" && (
                  <div style={{ marginBottom: 14 }}>
                    <Btn small variant="secondary" onClick={() => setUpdateTarget({ item: b, table: "breakdown_reports" })}>📝 {t(lang,"addUpdate")}</Btn>
                  </div>
                )}
                {b.status === "Pending Supervisor Approval" && isSupervisor && (
                  <BreakdownApprovalStep breakdown={b} userRole={userRole} lang={lang} onApproved={(updated) => setBreakdowns(prev => prev.map(x => x.id===updated.id?updated:x))} />
                )}
                {b.status === "Pending Operator Confirmation" && (
                  <BreakdownOperatorConfirm breakdown={b} userRole={userRole} lang={lang} onConfirmed={(updated) => setBreakdowns(prev => prev.map(x => x.id===updated.id?updated:x))} />
                )}
                {b.status === "Resolved" && b.supervisor_approved_by && (
                  <div style={{ background: C.green+"11", border: `1px solid ${C.green}33`, borderRadius: 8, padding: 12, marginBottom: 14, fontSize: 12, color: C.green }}>
                    ✅ Supervisor: {b.supervisor_approved_by} {b.operator_confirmed_by && ` · ${t(lang,"confirmedBy")}: ${b.operator_confirmed_by}`}
                  </div>
                )}
                {(b.status === "Open" || b.status === "Acknowledged") && isMaintenance && (
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    {b.status === "Open" && <Btn onClick={() => acknowledge(b)} color={C.blue}>{t(lang,"acknowledge")}</Btn>}
                    <Btn onClick={() => setResolveItem(b)} color={C.green}>{t(lang,"markResolved")}</Btn>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      ))}

      {/* Issues List */}
      {activeView === "issues" && (
        loading ? <Spinner lang={lang} /> : filteredIssues.length === 0 ? (
          <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>{t(lang,"noIssuesFound")}</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {filteredIssues.map(issue => {
              const isResolved = issue.status === "Resolved";
              const isAcknowledged = issue.status === "Acknowledged";
              return (
                <div key={issue.id} style={{ background: C.card, border: `1px solid ${isResolved?C.green+"44":isAcknowledged?C.blue+"44":C.yellow+"44"}`, borderRadius: 10, padding: 20 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10, marginBottom: 12 }}>
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                        <span style={{ fontSize: 18 }}>{isResolved?"✅":isAcknowledged?"👁":"⚠️"}</span>
                        <div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{issue.asset_name}</div>
                      </div>
                      <div style={{ fontSize: 12, color: C.muted }}>{issue.site} · {t(lang,"reportedBy")} {issue.reported_by} · {fmtDateTime(issue.reported_at)}</div>
                      {isAcknowledged && issue.acknowledged_by && (
                        <div style={{ fontSize: 12, color: C.blue, marginTop: 4 }}>👁 {t(lang,"acknowledged")}: {issue.acknowledged_by}</div>
                      )}
                    </div>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <Badge label={issue.severity} color={SEVERITY_COLORS[issue.severity]||C.muted} />
                      <Badge label={issue.status} color={statusColor(issue.status)} />
                      <div style={{ background: C.green+"22", color: C.green, border: `1px solid ${C.green}44`, borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>RUNNING</div>
                    </div>
                  </div>
                  <div style={{ background: C.surface, borderRadius: 8, padding: 12, marginBottom: 12, fontSize: 13, color: C.subtle }}>
                    <strong style={{ color: C.text }}>{t(lang,"issue")}:</strong> {issue.description}
                  </div>
                  {issue.work_order_id && (
                    <div style={{ fontSize: 12, color: C.blue, marginBottom: 12 }}>📋 {t(lang,"workOrders")}: {issue.work_order_id}</div>
                  )}
                  <UpdatesLog updates={issue.updates} lang={lang} />
                  {issue.status !== "Resolved" && (
                    <div style={{ marginBottom: 12 }}>
                      <Btn small variant="secondary" onClick={() => setUpdateTarget({ item: issue, table: "issue_reports" })}>📝 {t(lang,"addUpdate")}</Btn>
                    </div>
                  )}
                  {issue.status === "Pending Supervisor Approval" && isSupervisor && (
                    <IssueApprovalStep issue={issue} userRole={userRole} lang={lang} onApproved={(updated) => setIssues(prev => prev.map(x => x.id===updated.id?updated:x))} />
                  )}
                  {issue.status === "Pending Operator Confirmation" && (
                    <IssueOperatorConfirm issue={issue} userRole={userRole} lang={lang} onConfirmed={(updated) => setIssues(prev => prev.map(x => x.id===updated.id?updated:x))} />
                  )}
                  {issue.status === "Resolved" && issue.supervisor_approved_by && (
                    <div style={{ background: C.green+"11", border: `1px solid ${C.green}33`, borderRadius: 8, padding: 12, marginBottom: 12, fontSize: 12, color: C.green }}>
                      ✅ Supervisor: {issue.supervisor_approved_by}{issue.operator_confirmed_by && ` · ${t(lang,"confirmedBy")}: ${issue.operator_confirmed_by}`}
                    </div>
                  )}
                  {(issue.status === "Open" || issue.status === "Acknowledged") && (userRole.role === "maintenance" || userRole.role === "admin") && (
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      {issue.status === "Open" && <Btn onClick={() => acknowledgeIssue(issue)} color={C.blue}>{t(lang,"acknowledge")}</Btn>}
                      <PasswordConfirm lang={lang} actionLabel={t(lang,"markResolved")} onConfirmed={() => resolveIssue(issue)} />
                    </div>
                  )}
                </div>
              );
            })}
          </div>
       )
      )}
    </div>
  );
}

// ─── CIL CHECKLIST MODAL ─────────────────────────────────────────────────────
function ChecklistModal({ asset, workOrderId, onClose, lang, userRole }) {
  const [items, setItems] = useState([]);
  const [responses, setResponses] = useState({});
  const [executionId, setExecutionId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [executedBy, setExecutedBy] = useState(userRole?.name || userRole?.email || "");
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
    const needsApproval = userRole?.role === "maintenance";
    const logRecord = { id: uid("LOG"), asset_id: asset.id, asset_name: asset.name, log_type: "Preventive Maintenance", title: `CIL Checklist — ${new Date().toLocaleString("default", { month: "long", year: "numeric" })}`, description, performed_by: executedBy, vendor: null, start_date: TODAY, end_date: TODAY, cost: null, status: "In Progress", approval_status: "Pending", work_order_id: workOrderId || null, checklist_execution_id: executionId, downtime_start: null, downtime_end: null, downtime_hours: null };
    await supabase.from("maintenance_logs").insert([logRecord]);
    if (workOrderId) await supabase.from("work_orders").update({ status: "Awaiting Approval" }).eq("id", workOrderId);
    setSuccess(`${t(lang,"checklistCompleted")} ${failCount > 0 ? `⚠️ ${failCount} FAIL` : "✅"} — Sent for supervisor approval.`);
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
                <div style={{ flex: 1, minWidth: 200 }}>
                  <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase" }}>{t(lang,"technicianName")}</div>
                  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14 }}>{executedBy || "—"}</div>
                </div>
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
function ApprovalSection({ log, lang, userRole, onApproved, onRejected }) {
  const [rejectionNotes, setRejectionNotes] = useState("");
  const [showReject, setShowReject] = useState(false);
  const [showSign, setShowSign] = useState(false);
  const [password, setPassword] = useState("");
  const [saving, setSaving] = useState(false);
  const [sigError, setSigError] = useState(null);

  const approve = async () => {
    if (!password) { setSigError(t(lang,"enterPasswordToApprove")); return; }
    setSaving(true);
    const { data: sessionData } = await supabase.auth.getSession();
    const email = sessionData?.session?.user?.email;
    const { error: authErr } = await supabase.auth.signInWithPassword({ email, password });
    if (authErr) { setSigError(t(lang,"incorrectPassword")); setSaving(false); return; }

    await supabase.from("maintenance_logs").update({ approval_status: "Approved", approved_by: userRole?.name || "", approved_at: new Date().toISOString(), approved_signature: userRole?.name, status: "Completed" }).eq("id", log.id);
    const { data: updated } = await supabase.from("maintenance_logs").select("*").eq("id", log.id).single();

    // Close linked work order
    if (log.work_order_id) {
      await supabase.from("work_orders").update({ status: "Completed" }).eq("id", log.work_order_id);
    }

    const { data: parts } = await supabase.from("spare_parts").select("*").eq("log_id", log.id);
    if (parts?.length) {
      for (const part of parts) {
        if (part.model_part_id) {
          const { data: mp } = await supabase.from("model_parts").select("stock_quantity").eq("id", part.model_part_id).single();
          if (mp) await supabase.from("model_parts").update({ stock_quantity: Math.max(0,(mp.stock_quantity||0)-(part.quantity||1)) }).eq("id", part.model_part_id);
        }
        if (part.asset_part_id) {
          const { data: ap } = await supabase.from("asset_parts").select("stock_quantity").eq("id", part.asset_part_id).single();
          if (ap) await supabase.from("asset_parts").update({ stock_quantity: Math.max(0,(ap.stock_quantity||0)-(part.quantity||1)) }).eq("id", part.asset_part_id);
        }
      }
    }
    onApproved(updated || { ...log, approval_status: "Approved", approved_by: userRole?.name, status: "Completed" });
    setSaving(false);
  };

  const reject = async () => {
    if (!rejectionNotes) return;
    setSaving(true);
    await supabase.from("maintenance_logs").update({ approval_status: "Rejected", rejection_notes: rejectionNotes, status: "In Progress" }).eq("id", log.id);
    if (log.work_order_id) {
      await supabase.from("work_orders").update({ status: "In Progress", status_note: `Rejected: ${rejectionNotes}` }).eq("id", log.work_order_id);
    }
    onRejected({ ...log, approval_status: "Rejected", rejection_notes: rejectionNotes, status: "In Progress" });
    setShowReject(false);
    setSaving(false);
  };

  return (
    <div style={{ marginTop: 12, background: C.yellow+"11", border: `1px solid ${C.yellow}44`, borderRadius: 8, padding: 14 }}>
      <div style={{ fontSize: 12, color: C.yellow, fontWeight: 700, marginBottom: 10 }}>⏳ {t(lang,"pendingApproval")}</div>
      {!showReject && !showSign ? (
        <div style={{ display: "flex", gap: 8 }}>
          <Btn onClick={() => setShowSign(true)} color={C.green}>{t(lang,"approveLog")}</Btn>
          <Btn onClick={() => setShowReject(true)} variant="danger">{t(lang,"rejectLog")}</Btn>
        </div>
      ) : showSign ? (
        <div>
          <div style={{ fontSize: 12, color: C.subtle, marginBottom: 8 }}>{t(lang,"enterPasswordToApprove")}:</div>
          <Input label={t(lang,"confirmPassword")} value={password} type="password" onChange={v => { setPassword(v); setSigError(null); }} placeholder="••••••••" />
          {sigError && <div style={{ color: C.red, fontSize: 12, marginTop: 6 }}>{sigError}</div>}
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <Btn onClick={approve} disabled={saving || !password} color={C.green}>{saving ? "Verifying..." : "✓ Confirm & Approve"}</Btn>
            <Btn variant="secondary" onClick={() => { setShowSign(false); setPassword(""); setSigError(null); }}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      ) : (
        <div>
          <Textarea label={t(lang,"rejectionNotes")} value={rejectionNotes} onChange={setRejectionNotes} placeholder="Explain what needs to be corrected..." />
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <Btn onClick={reject} disabled={saving||!rejectionNotes} color={C.red}>{t(lang,"rejectLog")}</Btn>
            <Btn variant="secondary" onClick={() => setShowReject(false)}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      )}
    </div>
  );
}
function MaintenanceModal({ asset, onClose, isAdmin, isSupervisor, isMaintenance, vendors, lang, userRole }) {
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
  const [showCatalog, setShowCatalog] = useState(false);
  const [catalogParts, setCatalogParts] = useState([]);
  const [catalogForm, setCatalogForm] = useState({ part_name: "", part_number: "", supplier: "", unit_cost: "", stock_quantity: "", min_stock_level: "1", notes: "" });
  const [showCatalogForm, setShowCatalogForm] = useState(false);
  const cf = (k) => (v) => setCatalogForm(p => ({ ...p, [k]: v }));

  useEffect(() => { loadCatalog(); }, [asset.id]);

  const loadCatalog = async () => {
    const asset = (assets||[]).find(a => a.name === wo.asset);
    // Load model parts by category/site even if no specific asset found
    const [assetPartsRes, modelPartsRes] = await Promise.all([
      asset ? supabase.from("asset_parts").select("*").eq("asset_id", asset.id).order("part_name") : Promise.resolve({ data: [] }),
      asset?.model
        ? supabase.from("model_parts").select("*").eq("model", asset.model).order("part_name")
        : wo.asset
        ? supabase.from("model_parts").select("*").order("part_name").limit(50)
        : Promise.resolve({ data: [] }),
    ]);
    const assetParts = assetPartsRes.data || [];
    const modelParts = (modelPartsRes.data || []).map(p => ({ ...p, _id: p.id, id: `mdl-${p.id}`, model_part_id: p.id, isModelLevel: true }));
    const merged = [...assetParts];
    modelParts.forEach(mp => { if (!merged.find(ap => ap.part_name === mp.part_name)) merged.push(mp); });
    setCatalogParts(merged);
  };

  const saveCatalogPart = async () => {
    if (!catalogForm.part_name) return;
    const record = { id: uid("PRT"), asset_id: asset.id, asset_name: asset.name, part_name: catalogForm.part_name, part_number: catalogForm.part_number||null, supplier: catalogForm.supplier||null, unit_cost: parseFloat(catalogForm.unit_cost)||0, stock_quantity: parseFloat(catalogForm.stock_quantity)||0, min_stock_level: parseFloat(catalogForm.min_stock_level)||1, notes: catalogForm.notes||null };
    await supabase.from("asset_parts").insert([record]);
    setCatalogParts(prev => [...prev, record]);
    setCatalogForm({ part_name: "", part_number: "", supplier: "", unit_cost: "", stock_quantity: "", min_stock_level: "1", notes: "" });
    setShowCatalogForm(false);
  };

  const deleteCatalogPart = async (id) => {
    await supabase.from("asset_parts").delete().eq("id", id);
    setCatalogParts(prev => prev.filter(p => p.id !== id));
  };

  const updateStock = async (partId, newQty) => {
    await supabase.from("asset_parts").update({ stock_quantity: newQty }).eq("id", partId);
    setCatalogParts(prev => prev.map(p => p.id===partId?{...p,stock_quantity:newQty}:p));
  };
  const vendorOptions = ["— None —", ...vendors.filter(v => v.status==="Active").map(v => v.name)];
  const [form, setForm] = useState({ log_type: "Preventive Maintenance", title: "", description: "", performed_by: "", vendor: "", start_date: TODAY, end_date: "", cost: "", status: "Completed", downtime_start: "", downtime_end: "" });
  const [partForm, setPartForm] = useState({ part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "", asset_part_id: null });
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
    const needsApproval = userRoleRole === "maintenance";
    const record = { id: uid("LOG"), asset_id: asset.id, asset_name: asset.name, log_type: form.log_type, title: form.title, description: form.description, performed_by: form.performed_by, vendor: form.vendor==="— None —"?null:form.vendor||null, start_date: form.start_date||null, end_date: form.end_date||null, cost: form.cost?parseFloat(form.cost):null, status: needsApproval ? "In Progress" : "Completed", approval_status: needsApproval ? "Pending" : "Approved", approved_by: needsApproval ? null : userRole?.name, approved_at: needsApproval ? null : new Date().toISOString(), downtime_start: form.downtime_start||null, downtime_end: form.downtime_end||null, downtime_hours: (form.downtime_start && form.downtime_end) ? Math.round((new Date(form.downtime_end) - new Date(form.downtime_start)) / (1000 * 60 * 60)) : null };
    const { error: err } = await supabase.from("maintenance_logs").insert([record]);
    if (err) { setError(err.message); } else { setSuccess(t(lang,"saving")); setLogs(prev => [record,...prev]); setForm({ log_type: "Preventive Maintenance", title: "", description: "", performed_by: "", vendor: "", start_date: TODAY, end_date: "", cost: "", status: "Completed", downtime_start: "", downtime_end: "" }); setShowForm(false); }
    setSaving(false);
  };

  const submitPart = async (logId) => {
    if (!partForm.part_name) { setError(t(lang,"partName")); return; }
    setSaving(true); setError(null);
    const qty = parseFloat(partForm.quantity)||1;
    const unitCost = parseFloat(partForm.unit_cost)||0;
    // Clean asset_part_id — strip mdl- prefix for model-level parts since they don't exist in asset_parts
    const rawPartId = partForm.asset_part_id;
    const isModelLevel = rawPartId && String(rawPartId).startsWith("mdl-");
    const cleanPartId = isModelLevel ? null : rawPartId || null;
    const record = { id: uid("PRT"), log_id: logId, asset_id: asset.id, part_name: partForm.part_name, part_number: partForm.part_number||null, quantity: qty, unit_cost: unitCost, total_cost: qty*unitCost, supplier: partForm.supplier||null, asset_part_id: cleanPartId };
    const { error: err } = await supabase.from("spare_parts").insert([record]);
    if (err) { setError(err.message); } else { setSuccess("✓"); setParts(prev => ({ ...prev, [logId]: [...(prev[logId]||[]),record] })); setPartForm({ part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "", asset_part_id: null }); setShowPartForm(null); }
    setSaving(false);
  };

  const deletePart = async (partId, logId) => { await supabase.from("spare_parts").delete().eq("id", partId); setParts(prev => ({ ...prev, [logId]: prev[logId].filter(p => p.id!==partId) })); };
  const deleteLog = async (logId) => { await supabase.from("spare_parts").delete().eq("log_id", logId); await supabase.from("maintenance_logs").delete().eq("id", logId); setLogs(prev => prev.filter(l => l.id!==logId)); };
  const totalCost = logs.reduce((s,l) => s+(l.cost||0), 0);

  return (
    <>
      {showChecklist && <ChecklistModal asset={asset} lang={lang} userRole={userRole} onClose={() => setShowChecklist(false)} />}
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
              <div style={{ fontSize: 12, color: C.muted, padding: "8px 0" }}>📋 {t(lang,"maintenanceHistory")} — {t(lang,"readOnly")}</div>
            </div>
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
                        {log.approval_status === "Pending" && <Badge label={t(lang,"pendingApproval")} color={C.yellow} />}
                        {log.approval_status === "Approved" && <Badge label={t(lang,"approved")} color={C.green} />}
                        {log.approval_status === "Rejected" && <Badge label={t(lang,"rejected")} color={C.red} />}
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
                            {log.downtime_hours && <div><div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase" }}>⏱ {t(lang,"totalDowntimeLabel")}</div><div style={{ fontSize: 13, color: C.yellow, fontWeight: 700, marginTop: 2 }}>{formatDowntime(log.downtime_hours)}</div></div>}
                          </div>
                        )}
                        <div style={{ marginTop: 14 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                            <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>{t(lang,"spareParts")}</div>
                            {isSupervisor && <Btn small onClick={() => setShowPartForm(showPartForm===log.id?null:log.id)}>{t(lang,"addPart")}</Btn>}
                          </div>
                          {showPartForm===log.id && (
                            <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 8, padding: 14, marginBottom: 12 }}>
                              {catalogParts.length > 0 && (
                                <div style={{ marginBottom: 10 }}>
                                  <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase" }}>{t(lang,"selectFromCatalog")}</div>
                                  <select onChange={e => {
                                    const found = catalogParts.find(p => p.id === e.target.value);
                                    if (found) setPartForm({ part_name: found.part_name, part_number: found.part_number||"", quantity: "1", unit_cost: String(found.unit_cost||""), supplier: found.supplier||"", asset_part_id: found.id });
                                  }} style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "9px", color: C.text, fontSize: 13 }}>
                                    <option value="">{t(lang,"selectFromCatalog")}</option>
                                    {catalogParts.map(p => <option key={p.id} value={p.id}>{p.part_name} {p.part_number?`(${p.part_number})`:""} — ${p.unit_cost} {p.stock_quantity<=p.min_stock_level?`⚠️ ${p.stock_quantity} left`:""}</option>)}
                                  </select>
                                </div>
                              )}
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
                        {/* Approval section */}
                        {log.approval_status === "Pending" && isSupervisor && (
                          <ApprovalSection log={log} lang={lang} userRole={userRole} onApproved={(updated) => { setLogs(prev => prev.map(l => l.id===updated.id?updated:l)); loadLogs(); }} onRejected={(updated) => { setLogs(prev => prev.map(l => l.id===updated.id?updated:l)); loadLogs(); }} />
                        )}
                        {log.approval_status === "Approved" && log.approved_by && (
                          <div style={{ marginTop: 12, background: C.green+"11", border: `1px solid ${C.green}33`, borderRadius: 8, padding: "10px 14px", fontSize: 12, color: C.green }}>
                            ✅ {t(lang,"approvedBy")} <strong>{log.approved_by}</strong> · {log.approved_at ? fmtDateTime(log.approved_at) : ""}
                          </div>
                        )}
                        {log.approval_status === "Rejected" && log.rejection_notes && (
                          <div style={{ marginTop: 12, background: C.red+"11", border: `1px solid ${C.red}33`, borderRadius: 8, padding: "10px 14px", fontSize: 12, color: C.red }}>
                            ❌ {t(lang,"rejected")}: {log.rejection_notes}
                          </div>
                        )}
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
  const [mode, setMode] = useState("email"); // "email" | "phone"
  const [identifier, setIdentifier] = useState(""); const [password, setPassword] = useState("");
  const [error, setError] = useState(null); const [loading, setLoading] = useState(false);
  const signIn = async () => {
    if (!identifier || !password) { setError(t(lang,"enterEmail")); return; }
    setLoading(true); setError(null);
    const email = mode === "phone" ? `${identifier.replace(/\D/g,"")}@facility-command.local` : identifier;
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
          <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
            {["email","phone"].map(m => (
              <button key={m} onClick={() => { setMode(m); setIdentifier(""); }} style={{ flex: 1, background: mode===m?C.accent:C.surface, color: mode===m?"#fff":C.muted, border: `1px solid ${mode===m?C.accent:C.border}`, borderRadius: 6, padding: "8px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                {m==="email"?t(lang,"email"):t(lang,"phone")}
              </button>
            ))}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {mode === "phone"
              ? <Input label={t(lang,"phone")} value={identifier} onChange={setIdentifier} type="tel" placeholder="01012345678" />
              : <Input label={t(lang,"enterEmail")} value={identifier} onChange={setIdentifier} type="email" />}
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
function WOMaintenanceModal({ wo, onClose, isAdmin, isSupervisor, userRole, lang, vendors, assets }) {
  const [logs, setLogs] = useState([]);
  const [parts, setParts] = useState({});
  const [catalogParts, setCatalogParts] = useState([]);
  const [loadingLogs, setLoadingLogs] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [showPartForm, setShowPartForm] = useState(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [expandedLog, setExpandedLog] = useState(null);
  const [showChecklist, setShowChecklist] = useState(false);
  const vendorOptions = ["— None —",...(vendors||[]).filter(v => v.status==="Active").map(v => v.name)];
  const [form, setForm] = useState({ log_type: "Corrective Repair", title: wo.title, description: "", performed_by: userRole?.name||"", vendor: "", start_date: TODAY, end_date: TODAY, cost: "", downtime_start: "", downtime_end: "" });
  const [partForm, setPartForm] = useState({ part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "", asset_part_id: null, model_part_id: null });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const pf = (k) => (v) => setPartForm(p => ({ ...p, [k]: v }));

  useEffect(() => { loadLogs(); loadCatalog(); }, []);

  const loadLogs = async () => {
    setLoadingLogs(true);
    const { data } = await supabase.from("maintenance_logs").select("*").eq("asset_name", wo.asset).order("start_date", { ascending: false });
    setLogs(data || []);
    setLoadingLogs(false);
  };

  const loadCatalog = async () => {
    const asset = (assets||[]).find(a => a.name === wo.asset);
    if (!asset) return;
    const [assetPartsRes, modelPartsRes] = await Promise.all([
      supabase.from("asset_parts").select("*").eq("asset_id", asset.id).order("part_name"),
      asset.model ? supabase.from("model_parts").select("*").eq("model", asset.model).order("part_name") : Promise.resolve({ data: [] }),
    ]);
    const assetParts = assetPartsRes.data || [];
    const modelParts = (modelPartsRes.data || []).map(p => ({ ...p, _id: p.id, id: `mdl-${p.id}`, model_part_id: p.id, isModelLevel: true }));
    const merged = [...assetParts];
    modelParts.forEach(mp => { if (!merged.find(ap => ap.part_name === mp.part_name)) merged.push(mp); });
    setCatalogParts(merged);
  };

  const loadParts = async (logId) => {
    const { data } = await supabase.from("spare_parts").select("*").eq("log_id", logId);
    setParts(prev => ({ ...prev, [logId]: data || [] }));
  };

  const toggleLog = (logId) => {
    setExpandedLog(expandedLog===logId?null:logId);
    if (!parts[logId]) loadParts(logId);
  };

  const submitLog = async () => {
    if (!form.title) { setError(t(lang,"title")); return; }
    setSaving(true); setError(null);
    const needsApproval = userRole?.role === "maintenance";
    const asset = (assets||[]).find(a => a.name === wo.asset);
    const record = { id: uid("LOG"), asset_id: asset?.id||null, asset_name: wo.asset, log_type: form.log_type, title: form.title, description: form.description, performed_by: form.performed_by, vendor: form.vendor==="— None —"?null:form.vendor||null, start_date: form.start_date||null, end_date: form.end_date||null, cost: form.cost?parseFloat(form.cost):null, status: needsApproval?"In Progress":"Completed", approval_status: needsApproval?"Pending":"Approved", approved_by: needsApproval?null:userRole?.name, approved_at: needsApproval?null:new Date().toISOString(), downtime_start: form.downtime_start||null, downtime_end: form.downtime_end||null, downtime_hours: (form.downtime_start&&form.downtime_end)?Math.round((new Date(form.downtime_end)-new Date(form.downtime_start))/(1000*60*60)):null };
    const { error: err } = await supabase.from("maintenance_logs").insert([record]);
    if (err) { setError(err.message); } else {
      setLogs(prev => [record,...prev]);
      setSuccess(needsApproval ? t(lang,"approvalRequired") : t(lang,"saveLog"));
      setForm({ log_type: "Corrective Repair", title: wo.title, description: "", performed_by: userRole?.name||"", vendor: "", start_date: TODAY, end_date: TODAY, cost: "", downtime_start: "", downtime_end: "" });
      setShowForm(false);
    }
    setSaving(false);
  };

  const submitPart = async (logId) => {
    if (!partForm.part_name) { setError(t(lang,"partName")); return; }
    setSaving(true); setError(null);
    const qty = parseFloat(partForm.quantity)||1;
    const unitCost = parseFloat(partForm.unit_cost)||0;
    const rawId = partForm.asset_part_id;
    const isModelLevel = rawId && String(rawId).startsWith("mdl-");
    const cleanAssetPartId = isModelLevel ? null : rawId || null;
    const cleanModelPartId = partForm.model_part_id || null;
    const assetForPart = (assets||[]).find(a => a.name === wo.asset);
    const record = { id: uid("PRT"), log_id: logId, asset_id: assetForPart?.id||null, part_name: partForm.part_name, part_number: partForm.part_number||null, quantity: qty, unit_cost: unitCost, total_cost: qty*unitCost, supplier: partForm.supplier||null, asset_part_id: cleanAssetPartId, model_part_id: cleanModelPartId };
    const { error: err } = await supabase.from("spare_parts").insert([record]);
    if (err) { setError(err.message); } else {
      setSuccess("✓");
      setParts(prev => ({ ...prev, [logId]: [...(prev[logId]||[]),record] }));
      setPartForm({ part_name: "", part_number: "", quantity: "1", unit_cost: "", supplier: "", asset_part_id: null, model_part_id: null });
      setShowPartForm(null);
      setLogs(prev => prev.map(l => l.id===logId ? { ...l, cost: (l.cost||0) + (qty*unitCost) } : l));
      // Check if the log is already approved — if so deduct stock immediately
      const log = logs.find(l => l.id === logId);
      if (log?.approval_status === "Approved") {
        if (cleanModelPartId) {
          const { data: mp } = await supabase.from("model_parts").select("stock_quantity").eq("id", cleanModelPartId).single();
          if (mp) await supabase.from("model_parts").update({ stock_quantity: Math.max(0,(mp.stock_quantity||0)-qty) }).eq("id", cleanModelPartId);
        }
        if (cleanAssetPartId) {
          const { data: ap } = await supabase.from("asset_parts").select("stock_quantity").eq("id", cleanAssetPartId).single();
          if (ap) await supabase.from("asset_parts").update({ stock_quantity: Math.max(0,(ap.stock_quantity||0)-qty) }).eq("id", cleanAssetPartId);
        }
      }
    }
    setSaving(false);
  };

  const deletePart = async (partId, logId) => { await supabase.from("spare_parts").delete().eq("id", partId); setParts(prev => ({ ...prev, [logId]: prev[logId].filter(p => p.id!==partId) })); };
  const totalCost = logs.reduce((s,l) => s+(l.cost||0),0);

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1000, padding: 16, overflowY: "auto" }}>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, width: "100%", maxWidth: 860, marginTop: 20, marginBottom: 20 }}>
        {/* Header */}
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>{wo.title}</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{wo.asset} · {wo.site} · {wo.category}</div>
          </div>
          <button onClick={onClose} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, color: C.muted, cursor: "pointer", fontSize: 18, padding: "2px 10px" }}>✕</button>
        </div>

        {/* Stats */}
        <div style={{ padding: "14px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 12, flexWrap: "wrap" }}>
          {[["📋",logs.length,t(lang,"totalLogs"),C.blue],["💰",`$${totalCost.toLocaleString()}`,t(lang,"totalCost"),C.accent],["🔧",logs[0]?.start_date?fmtDate(logs[0].start_date):t(lang,"never"),t(lang,"lastMaintenance"),C.green],["⏳",logs.filter(l=>l.approval_status==="Pending").length,t(lang,"pendingApproval"),C.yellow]].map(([icon,val,label,color]) => (
            <div key={label} style={{ background: C.surface, borderRadius: 8, padding: "10px 16px", flex: "1 1 120px", borderLeft: `3px solid ${color}` }}>
              <div style={{ fontSize: 16 }}>{icon}</div><div style={{ fontSize: 18, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{val}</div><div style={{ fontSize: 11, color: C.muted }}>{label}</div>
            </div>
          ))}
        </div>

        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <OkBanner msg={success} onDismiss={() => setSuccess(null)} />

          {/* Add Log Button */}
          <div style={{ marginBottom: 20, display: "flex", gap: 10, flexWrap: "wrap" }}>
            <Btn onClick={() => setShowForm(v => !v)}>{t(lang,"addMaintenanceLog")}</Btn>
            {wo.category === "MHE" && (assets||[]).find(a => a.name === wo.asset) && (
              <button onClick={() => setShowChecklist(true)} style={{ background: C.blue+"22", color: C.blue, border: `1px solid ${C.blue}44`, borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>{t(lang,"runCILChecklist")}</button>
            )}
          </div>

          {/* Log Form */}
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
                <Sel label={t(lang,"status")} value={form.status||"Completed"} onChange={f("status")} options={LOG_STATUSES} />
                <Input label={t(lang,"downtimeStart")} value={form.downtime_start} onChange={f("downtime_start")} type="date" />
                <Input label={t(lang,"backToOperationLabel")} value={form.downtime_end} onChange={f("downtime_end")} type="date" />
              </div>
              <div style={{ marginTop: 12 }}><Textarea label={t(lang,"descriptionNotes")} value={form.description} onChange={f("description")} /></div>
              {userRole?.role === "maintenance" && (
                <div style={{ marginTop: 10, background: C.yellow+"11", border: `1px solid ${C.yellow}33`, borderRadius: 6, padding: "8px 12px", fontSize: 12, color: C.yellow }}>
                  ⏳ {t(lang,"approvalRequired")} — {t(lang,"pendingApproval")}
                </div>
              )}
              <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
                <Btn onClick={submitLog} disabled={saving}>{saving?t(lang,"saving"):t(lang,"saveLog")}</Btn>
                <Btn variant="secondary" onClick={() => setShowForm(false)}>{t(lang,"cancel")}</Btn>
              </div>
            </div>
          )}

          {/* Logs List */}
          <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 14 }}>{t(lang,"maintenanceHistory")}</div>
          {loadingLogs ? <Spinner lang={lang} /> : logs.length===0 ? (
            <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>{t(lang,"noMaintenanceRecords")}</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {logs.map(log => (
                <div key={log.id} style={{ background: C.surface, border: `1px solid ${log.approval_status==="Pending"?C.yellow+"44":log.approval_status==="Rejected"?C.red+"44":C.border}`, borderRadius: 10, overflow: "hidden" }}>
                  <div onClick={() => toggleLog(log.id)} style={{ padding: "12px 16px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <span style={{ fontSize: 18 }}>{log.log_type==="Preventive Maintenance"?"🔧":log.log_type==="Corrective Repair"?"🔨":log.log_type==="Inspection"?"🔍":log.log_type==="Overhaul"?"⚙️":"🔩"}</span>
                      <div><div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{log.title}</div><div style={{ fontSize: 11, color: C.muted }}>{fmtDate(log.start_date)}{log.performed_by?` · ${log.performed_by}`:""}</div></div>
                    </div>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <Badge label={log.log_type} color={statusColor(log.log_type)} />
                      {log.approval_status==="Pending" && <Badge label={t(lang,"pendingApproval")} color={C.yellow} />}
                      {log.approval_status==="Approved" && <Badge label={t(lang,"approved")} color={C.green} />}
                      {log.approval_status==="Rejected" && <Badge label={t(lang,"rejected")} color={C.red} />}
                      {log.cost>0 && <span style={{ fontSize: 12, color: C.accent, fontWeight: 700 }}>${log.cost}</span>}
                      <span style={{ color: C.muted }}>{expandedLog===log.id?"▲":"▼"}</span>
                    </div>
                  </div>
                  {expandedLog===log.id && (
                    <div style={{ padding: "0 16px 16px", borderTop: `1px solid ${C.border}` }}>
                      {log.description && <div style={{ marginTop: 12, padding: 12, background: C.card, borderRadius: 8, fontSize: 13, color: C.subtle, lineHeight: 1.6 }}>{log.description}</div>}

                      {/* Spare Parts */}
                      <div style={{ marginTop: 14 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                          <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>{t(lang,"spareParts")}</div>
                          <Btn small onClick={() => setShowPartForm(showPartForm===log.id?null:log.id)}>{t(lang,"addPart")}</Btn>
                        </div>
                        {showPartForm===log.id && (
                          <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 8, padding: 14, marginBottom: 12 }}>
                            {catalogParts.length>0 && (
                              <div style={{ marginBottom: 10 }}>
                                <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase" }}>{t(lang,"selectFromCatalog")}</div>
                                <select onChange={e => {
                                  const found = catalogParts.find(p => p.id===e.target.value);
                                  if (found) setPartForm({ part_name: found.part_name, part_number: found.part_number||"", quantity: "1", unit_cost: String(found.unit_cost||""), supplier: found.supplier||"", asset_part_id: found.id, model_part_id: found.model_part_id||null });
                                }} style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "9px", color: C.text, fontSize: 13 }}>
                                  <option value="">{t(lang,"selectFromCatalog")}</option>
                                  {catalogParts.map(p => <option key={p.id} value={p.id}>{p.part_name}{p.part_number?` (${p.part_number})`:""} — ${p.unit_cost||0}{p.stock_quantity!==undefined?` | Stock: ${p.stock_quantity}`:""}</option>)}
                                </select>
                              </div>
                            )}
                            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10 }}>
                              <Input label={t(lang,"partName")} value={partForm.part_name} onChange={pf("part_name")} />
                              <Input label={t(lang,"partNumber")} value={partForm.part_number} onChange={pf("part_number")} />
                              <Input label={t(lang,"quantity")} value={partForm.quantity} onChange={pf("quantity")} type="number" />
                              <Input label={t(lang,"unitCostLabel")} value={partForm.unit_cost} onChange={pf("unit_cost")} type="number" />
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
                                {[t(lang,"partName"),t(lang,"partNumber"),t(lang,"quantity"),"Unit","Total",t(lang,"supplier"),""].map(h => <th key={h} style={{ textAlign: "left", padding: "6px 10px", color: C.muted, fontWeight: 600, fontSize: 11, textTransform: "uppercase" }}>{h}</th>)}
                              </tr></thead>
                              <tbody>
                                {parts[log.id].map(part => (
                                  <tr key={part.id} style={{ borderBottom: `1px solid ${C.border}22` }}>
                                    <td style={{ padding: "8px 10px", color: C.text, fontWeight: 600 }}>{part.part_name}</td>
                                    <td style={{ padding: "8px 10px", color: C.subtle, fontFamily: "monospace" }}>{part.part_number||"—"}</td>
                                    <td style={{ padding: "8px 10px", color: C.subtle }}>{part.quantity}</td>
                                    <td style={{ padding: "8px 10px", color: C.subtle }}>{part.unit_cost?`$${part.unit_cost}`:"—"}</td>
                                    <td style={{ padding: "8px 10px", color: C.accent, fontWeight: 700 }}>{part.total_cost?`$${part.total_cost}`:"—"}</td>
                                    <td style={{ padding: "8px 10px", color: C.subtle }}>{part.supplier||"—"}</td>
                                    <td style={{ padding: "8px 10px" }}><Btn small variant="danger" onClick={() => deletePart(part.id, log.id)}>{t(lang,"del")}</Btn></td>
                                  </tr>
                                ))}
                                <tr style={{ borderTop: `1px solid ${C.border}` }}>
                                  <td colSpan={4} style={{ padding: "8px 10px", color: C.muted, fontSize: 11 }}>Total</td>
                                  <td style={{ padding: "8px 10px", color: C.accent, fontWeight: 700 }}>${parts[log.id].reduce((s,p)=>s+(p.total_cost||0),0).toLocaleString()}</td>
                                  <td colSpan={2} />
                                </tr>
                              </tbody>
                            </table>
                          </div>
                        ) : <div style={{ fontSize: 12, color: C.muted }}>{t(lang,"noSpareParts")}</div>}
                      </div>

                      {/* Approval */}
                      {log.approval_status==="Pending" && isSupervisor && (
                        <ApprovalSection log={log} lang={lang} userRole={userRole} onApproved={(updated) => setLogs(prev => prev.map(l => l.id===updated.id?updated:l))} onRejected={(updated) => setLogs(prev => prev.map(l => l.id===updated.id?updated:l))} />
                      )}
                      {log.approval_status==="Approved" && log.approved_by && (
                        <div style={{ marginTop: 12, background: C.green+"11", border: `1px solid ${C.green}33`, borderRadius: 8, padding: "10px 14px", fontSize: 12, color: C.green }}>
                          ✅ {t(lang,"approvedBy")} <strong>{log.approved_by}</strong> · {log.approved_at ? fmtDateTime(log.approved_at) : ""}
                        </div>
                      )}
                      {log.approval_status==="Rejected" && log.rejection_notes && (
                        <div style={{ marginTop: 12, background: C.red+"11", border: `1px solid ${C.red}33`, borderRadius: 8, padding: "10px 14px", fontSize: 12, color: C.red }}>
                          ❌ {t(lang,"rejected")}: {log.rejection_notes}
                        </div>
                      )}
                      {isAdmin && <div style={{ marginTop: 12, display: "flex", justifyContent: "flex-end" }}><Btn small variant="danger" onClick={async () => { await supabase.from("spare_parts").delete().eq("log_id",log.id); await supabase.from("maintenance_logs").delete().eq("id",log.id); setLogs(prev=>prev.filter(l=>l.id!==log.id)); }}>{t(lang,"deleteLog")}</Btn></div>}
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
     </div>
      {showChecklist && (
        <ChecklistModal asset={(assets||[]).find(a => a.name === wo.asset)} workOrderId={wo.id} lang={lang} userRole={userRole} onClose={() => { setShowChecklist(false); loadLogs(); }} />
      )}
    </div>
  );
}
const WO_CATEGORIES = ["MHE","HVAC","Fire Alarm & Suppression","Electrical","Plumbing","Civil & Structural","Security Systems","Lighting","General Maintenance"];
const WO_STATUSES = ["Open","In Progress","Awaiting PO","Awaiting Parts","Awaiting Approval","On Hold","Scheduled","Completed"];
const CATEGORY_ICONS = { "MHE":"🏭","HVAC":"❄️","Fire Alarm & Suppression":"🔥","Electrical":"⚡","Plumbing":"🔧","Civil & Structural":"🏗️","Security Systems":"🔒","Lighting":"💡","General Maintenance":"🔨" };

function WorkOrders({ workOrders, setWorkOrders, loading, onAdd, isAdmin, isSupervisor, isMaintenance, vendors, assets, lang, userRole, sites }) {
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [editItem, setEditItem] = useState(null);
  const [deleteItem, setDeleteItem] = useState(null);
  const [selectedWO, setSelectedWO] = useState(null);
  const [noteItem, setNoteItem] = useState(null);
  const [noteText, setNoteText] = useState("");
  const [logWO, setLogWO] = useState(null);
  const [siteFilter, setSiteFilter] = useState("All");
  const [catFilter, setCatFilter] = useState("All");
  const [statusFilter, setStatusFilter] = useState("Active");
  const [showArchived, setShowArchived] = useState(false);
  const [woType, setWoType] = useState("pm");
  const [archiveMonth, setArchiveMonth] = useState("All");
  const [form, setForm] = useState({ title: "", asset: "", category: "MHE", priority: "Medium", start_date: "", due: "", vendor: "", site: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const vendorOptions = ["— None —",...vendors.filter(v => v.status==="Active").map(v => v.name)];

  // Auto-set site when asset is selected
  const handleAssetSelect = (assetName) => {
    f("asset")(assetName);
    const found = assets.find(a => a.name === assetName);
    if (found) f("site")(found.location);
  };

  // Split active vs archived
  const typeFiltered = workOrders.filter(w => woType === "pm" ? w.title.startsWith("PM -") : !w.title.startsWith("PM -"));
  const activeWOs = typeFiltered.filter(w => w.status !== "Completed");
  const archivedWOs = typeFiltered.filter(w => w.status === "Completed");

  const applyFilters = (list) => list.filter(w =>
    (siteFilter === "All" || w.site === siteFilter) &&
    (catFilter === "All" || w.category === catFilter) &&
    (statusFilter === "All" || statusFilter === "Active" || w.status === statusFilter)
  );

  const filteredActive = applyFilters(activeWOs);
  const filteredArchived = applyFilters(archivedWOs).filter(w => {
    if (archiveMonth === "All") return true;
    const d = new Date(w.updated_at || w.start_date || "");
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}` === archiveMonth;
  });

  // Get unique months from archived WOs for filter
  const archiveMonths = ["All", ...new Set(archivedWOs.map(w => {
    const d = new Date(w.updated_at || w.start_date || "");
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
  }).filter(Boolean)).values()].sort().reverse();

  // Group by site
  const groupBySite = (list) => {
    const groups = {};
    list.forEach(wo => {
      const site = wo.site || "Unassigned";
      if (!groups[site]) groups[site] = [];
      groups[site].push(wo);
    });
    return groups;
  };

  const activeGroups = siteFilter === "All" ? groupBySite(filteredActive) : { [siteFilter]: filteredActive };
  const archivedGroups = siteFilter === "All" ? groupBySite(filteredArchived) : { [siteFilter]: filteredArchived };

  // Per-site KPIs
  const siteKPIs = sites.filter(s => s !== "— Select Site —").map(site => ({
    site,
    open: workOrders.filter(w => w.site === site && w.status === "Open").length,
    inProgress: workOrders.filter(w => w.site === site && w.status === "In Progress").length,
    overdue: workOrders.filter(w => w.site === site && w.due && w.due <= TODAY && w.status !== "Completed").length,
    total: workOrders.filter(w => w.site === site && w.status !== "Completed").length,
  })).filter(s => s.total > 0);

  const submit = async () => {
    if (!form.title || !form.asset) { setError(t(lang,"title")); return; }
    setSaving(true); setError(null);
    const vendorName = form.vendor === "— None —" || !form.vendor ? null : form.vendor;
    const found = assets.find(a => a.name === form.asset);
    const record = { id: uid("WO"), title: form.title, asset: form.asset, asset_id: found?.id || null, category: form.category, priority: form.priority, status: "Open", assignee: null, start_date: form.start_date||null, due: form.due||null, vendor: vendorName, site: found?.location || form.site || null };
    const { error: err } = await supabase.from("work_orders").insert([record]);
    if (err) { setError(err.message); } else {
      onAdd(record);
      if (vendorName) {
        const { data: vd } = await supabase.from("vendors").select("id, open_orders").eq("name", vendorName).single();
        if (vd) await supabase.from("vendors").update({ open_orders: (vd.open_orders||0)+1 }).eq("id", vd.id);
      }
      setForm({ title: "", asset: "", category: "MHE", priority: "Medium", start_date: "", due: "", vendor: "", site: "" });
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
        await supabase.from("maintenance_logs").insert([{ id: uid("LOG"), asset_id: assetData.id, asset_name: wo.asset, log_type: wo.title.startsWith("PM -") ? "Preventive Maintenance" : "Corrective Repair", title: wo.title, description: `Work order completed on ${TODAY}.\nCategory: ${wo.category||"—"}\nVendor: ${wo.vendor||"—"}\nPriority: ${wo.priority}`, performed_by: wo.assignee||"—", vendor: wo.vendor||null, start_date: wo.start_date||TODAY, end_date: TODAY, cost: null, status: "Completed", downtime_start: null, downtime_end: null, downtime_hours: null }]);
        // Update last_pm_date to actual completion date for PM work orders
        if (wo.title.startsWith("PM -")) {
          await supabase.from("assets").update({ last_pm_date: TODAY }).eq("id", assetData.id);
        }
      }
    }
  };

  const updatePriority = async (id, val) => { await supabase.from("work_orders").update({ priority: val }).eq("id",id); setWorkOrders(prev => prev.map(wo => wo.id===id?{...wo,priority:val}:wo)); };
  const saveEdit = async (updated) => { const { error: err } = await supabase.from("work_orders").update(updated).eq("id",updated.id); if (!err) { setWorkOrders(prev => prev.map(wo => wo.id===updated.id?updated:wo)); setEditItem(null); } else setError(err.message); };
  const saveNote = async () => {
    if (!noteItem) return;
    await supabase.from("work_orders").update({ status_note: noteText }).eq("id", noteItem.id);
    setWorkOrders(prev => prev.map(w => w.id === noteItem.id ? { ...w, status_note: noteText } : w));
    setNoteItem(null);
    setNoteText("");
  };
  const confirmDelete = async () => { await supabase.from("work_orders").delete().eq("id",deleteItem.id); setWorkOrders(prev => prev.filter(wo => wo.id!==deleteItem.id)); setDeleteItem(null); };

  const WOTable = ({ wos }) => (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 700 }}>
        <thead><tr style={{ borderBottom: `1px solid ${C.border}` }}>
          {[t(lang,"id"),t(lang,"title"),t(lang,"asset"),t(lang,"category"),t(lang,"priority"),t(lang,"status"),t(lang,"lastUpdate"),t(lang,"vendor"),t(lang,"due"),t(lang,"photos"),...(isAdmin?[t(lang,"actions")]:[])].map(h => <th key={h} style={{ textAlign: "left", padding: "8px 12px", fontSize: 11, color: C.muted, textTransform: "uppercase", fontWeight: 600 }}>{h}</th>)}
        </tr></thead>
        <tbody>
          {wos.map((wo,i) => (
            <tr key={wo.id} style={{ borderBottom: `1px solid ${C.border}22`, background: i%2===0?"transparent":C.surface+"44" }}>
              <td style={{ padding: "10px 12px", fontSize: 11, color: C.muted, fontFamily: "monospace" }}>{wo.id}</td>
              <td style={{ padding: "10px 12px", fontSize: 13, color: C.text, fontWeight: 600 }}>{wo.title}</td>
              <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.asset}</td>
              <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{CATEGORY_ICONS[wo.category]||"🔧"} {wo.category||"—"}</td>
              <td style={{ padding: "10px 12px" }}><StatusSel value={wo.priority} options={["Critical","High","Medium","Low"]} onChange={val => updatePriority(wo.id,val)} /></td>
              <td style={{ padding: "10px 12px" }}><StatusSel value={wo.status} options={isAdmin || isSupervisor ? WO_STATUSES : WO_STATUSES.filter(s => s !== "Completed")} onChange={val => updateStatus(wo.id,val)} /></td>
                <td style={{ padding: "10px 12px", maxWidth: 200 }}>
                {wo.status_note ? (
                  <div style={{ fontSize: 11, color: C.subtle, background: C.surface, borderRadius: 4, padding: "4px 8px" }}>{wo.status_note}</div>
                ) : (
                  isAdmin && <button onClick={() => { setNoteItem(wo); setNoteText(wo.status_note||""); }} style={{ fontSize: 11, color: C.muted, background: "transparent", border: `1px dashed ${C.border}`, borderRadius: 4, padding: "3px 8px", cursor: "pointer" }}>+ {t(lang,"addUpdate")}</button>
                )}
                {wo.status_note && isAdmin && (
                  <button onClick={() => { setNoteItem(wo); setNoteText(wo.status_note||""); }} style={{ fontSize: 10, color: C.muted, background: "transparent", border: "none", cursor: "pointer", display: "block", marginTop: 2 }}>✏️ {t(lang,"edit")}</button>
                )}
              </td>
              <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.vendor||"—"}</td>
              <td style={{ padding: "10px 12px", fontSize: 12, color: wo.due&&wo.due<=TODAY&&wo.status!=="Completed"?C.red:C.subtle }}>{wo.due||"—"}</td>
              <td style={{ padding: "10px 12px" }}>
                <div style={{ display: "flex", gap: 6 }}>
                  <Btn small onClick={() => setSelectedWO(wo)} color={C.purple}>📷</Btn>
                  {isMaintenance && <Btn small onClick={() => setLogWO(wo)} color={C.green}>🔧</Btn>}
                  {isAdmin && <><Btn small onClick={() => setEditItem(wo)} color={C.blue}>{t(lang,"edit")}</Btn><Btn small variant="danger" onClick={() => setDeleteItem(wo)}>{t(lang,"del")}</Btn></>}
                </div>
              </td>
            </tr>
          ))}
          {wos.length===0 && <tr><td colSpan={isAdmin?10:9} style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>{t(lang,"noWorkOrdersForSite")}</td></tr>}
        </tbody>
      </table>
    </div>
  );

  return (
    <div>
      {selectedWO && <WorkOrderPhotosModal workOrder={selectedWO} lang={lang} onClose={() => setSelectedWO(null)} />}
      {logWO && <WOMaintenanceModal wo={logWO} onClose={() => setLogWO(null)} isAdmin={isAdmin} isSupervisor={isSupervisor} userRole={userRole} lang={lang} vendors={vendors} assets={assets} />}
      {noteItem && (
        <div style={{ position: "fixed", inset: 0, background: "#000000aa", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}>
          <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 24, width: "100%", maxWidth: 460 }}>
            <div style={{ fontSize: 15, fontWeight: 700, color: C.accent, marginBottom: 4 }}>{t(lang,"statusUpdate")}</div>
            <div style={{ fontSize: 12, color: C.muted, marginBottom: 16 }}>{noteItem.title}</div>
            <textarea value={noteText} onChange={e => setNoteText(e.target.value)}
              placeholder="e.g. Awaiting PO approval from procurement team..."
              rows={3} style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14, boxSizing: "border-box", resize: "vertical", marginBottom: 16 }} />
            <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
              {["Awaiting PO","Awaiting Parts","Awaiting Approval","On Hold","Scheduled"].map(s => (
                <button key={s} onClick={() => setNoteText(s)} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 10px", fontSize: 11, color: C.subtle, cursor: "pointer" }}>{s}</button>
              ))}
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <Btn onClick={saveNote}>{t(lang,"save")}</Btn>
              <Btn variant="secondary" onClick={() => { setNoteItem(null); setNoteText(""); }}>{t(lang,"cancel")}</Btn>
            </div>
          </div>
        </div>
      )}
      {editItem && <EditModal lang={lang} title={t(lang,"workOrders")} data={editItem} fields={[{key:"title",label:t(lang,"title")},{key:"asset",label:t(lang,"asset")},{key:"category",label:t(lang,"category"),options:WO_CATEGORIES},{key:"site",label:t(lang,"site"),options:sites},{key:"priority",label:t(lang,"priority"),options:["Critical","High","Medium","Low"]},{key:"status",label:t(lang,"status"),options:WO_STATUSES},{key:"status_note",label:t(lang,"statusUpdate")},{key:"vendor",label:t(lang,"vendor"),options:vendorOptions},{key:"assignee",label:t(lang,"assignee")},{key:"start_date",label:t(lang,"startDate"),type:"date"},{key:"due",label:t(lang,"dueDate"),type:"date"}]} onSave={saveEdit} onClose={() => setEditItem(null)} />}
      {deleteItem && <ConfirmDel lang={lang} name={deleteItem.title} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      <ErrBanner msg={error} onDismiss={() => setError(null)} />

      {/* Per-Site KPIs */}
      {siteKPIs.length > 0 && (
        <div style={{ marginBottom: 24 }}>
          <div style={{ fontSize: 13, color: C.muted, textTransform: "uppercase", fontWeight: 600, marginBottom: 12 }}>{t(lang,"siteKPIs")}</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            {siteKPIs.map(s => (
              <div key={s.site} onClick={() => setSiteFilter(s.site)} style={{ background: siteFilter===s.site?C.accent+"22":C.card, border: `1px solid ${siteFilter===s.site?C.accent:C.border}`, borderRadius: 10, padding: "12px 16px", cursor: "pointer", minWidth: 130 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: siteFilter===s.site?C.accent:C.text, marginBottom: 6 }}>{s.site}</div>
                <div style={{ display: "flex", gap: 10, fontSize: 12 }}>
                  <span style={{ color: C.accent }}>{s.open} {t(lang,"open")}</span>
                  {s.overdue > 0 && <span style={{ color: C.red }}>⚠️ {s.overdue}</span>}
                  {s.inProgress > 0 && <span style={{ color: C.blue }}>{s.inProgress} WIP</span>}
                </div>
              </div>
            ))}
            {siteFilter !== "All" && (
              <div onClick={() => setSiteFilter("All")} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: "12px 16px", cursor: "pointer", display: "flex", alignItems: "center" }}>
                <span style={{ fontSize: 12, color: C.muted }}>✕ {t(lang,"all")}</span>
              </div>
            )}
          </div>
        </div>
      )}
{/* PM vs Other Toggle */}
      <div style={{ display: "flex", gap: 8, marginBottom: 18 }}>
        <button onClick={() => setWoType("pm")} style={{ background: woType==="pm"?C.blue:C.card, color: woType==="pm"?"#fff":C.muted, border: `1px solid ${woType==="pm"?C.blue:C.border}`, borderRadius: 6, padding: "8px 16px", fontSize: 13, cursor: "pointer", fontWeight: 700 }}>📅 Scheduled PM</button>
        <button onClick={() => setWoType("other")} style={{ background: woType==="other"?C.accent:C.card, color: woType==="other"?"#fff":C.muted, border: `1px solid ${woType==="other"?C.accent:C.border}`, borderRadius: 6, padding: "8px 16px", fontSize: 13, cursor: "pointer", fontWeight: 700 }}>🔧 Other Work Orders</button>
      </div>
      {/* Filters */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <select value={siteFilter} onChange={e => setSiteFilter(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12 }}>
            <option value="All">{t(lang,"all")} Sites</option>
            {sites.filter(s => s !== "— Select Site —").map(s => <option key={s}>{s}</option>)}
          </select>
          <select value={catFilter} onChange={e => setCatFilter(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12 }}>
            <option value="All">{t(lang,"all")} {t(lang,"category")}</option>
            {WO_CATEGORIES.map(c => <option key={c}>{c}</option>)}
          </select>
          <select value={statusFilter} onChange={e => setStatusFilter(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12 }}>
            <option value="Active">Active</option>
            <option value="All">{t(lang,"all")}</option>
            {WO_STATUSES.filter(s => s !== "Completed").map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <Btn onClick={() => setShowForm(v => !v)}>{t(lang,"newWorkOrder")}</Btn>
      </div>

      {/* New WO Form */}
      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 18 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, textTransform: "uppercase" }}>{t(lang,"newWorkOrder")}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label={t(lang,"title")} value={form.title} onChange={f("title")} />
            <div>
              <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t(lang,"asset")}</div>
              <select value={form.asset} onChange={e => handleAssetSelect(e.target.value)} style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14 }}>
                <option value="">{t(lang,"selectAsset")}</option>
                {(assets||[]).map(a => <option key={a.id} value={a.name}>{a.name} ({a.location})</option>)}
              </select>
            </div>
            <Sel label={t(lang,"category")} value={form.category} onChange={f("category")} options={WO_CATEGORIES} />
            <Sel label={t(lang,"site")} value={form.site||"— Select Site —"} onChange={f("site")} options={sites} />
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

      {/* Active Work Orders grouped by site */}
      {loading ? <Spinner lang={lang} /> : (
        <div>
          {Object.entries(activeGroups).map(([site, wos]) => wos.length === 0 ? null : (
            <div key={site} style={{ marginBottom: 28 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                <div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>📍 {site}</div>
                <div style={{ fontSize: 12, color: C.muted }}>{wos.length} {t(lang,"workOrders")}</div>
                {wos.filter(w => w.due && w.due <= TODAY).length > 0 && (
                  <Badge label={`⚠️ ${wos.filter(w => w.due && w.due <= TODAY).length} overdue`} color={C.red} />
                )}
              </div>
              <WOTable wos={wos} />
            </div>
          ))}
          {Object.values(activeGroups).every(g => g.length === 0) && (
            <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>{t(lang,"noWorkOrdersFound")}</div>
          )}

          {/* Archived Section */}
          <div style={{ marginTop: 24, borderTop: `1px solid ${C.border}`, paddingTop: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ fontSize: 14, fontWeight: 700, color: C.muted }}>📦 {t(lang,"archived")}</div>
                <Badge label={`${archivedWOs.length}`} color={C.muted} />
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                {showArchived && (
                  <select value={archiveMonth} onChange={e => setArchiveMonth(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 10px", color: C.text, fontSize: 12 }}>
                    {archiveMonths.map(m => <option key={m} value={m}>{m === "All" ? t(lang,"all") : m}</option>)}
                  </select>
                )}
                <button onClick={() => setShowArchived(v => !v)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "6px 12px", color: C.muted, cursor: "pointer", fontSize: 12 }}>
                  {showArchived ? t(lang,"hideArchived") : t(lang,"showArchived")}
                </button>
              </div>
            </div>
            {showArchived && (
              <div>
                {Object.entries(archivedGroups).map(([site, wos]) => wos.length === 0 ? null : (
                  <div key={site} style={{ marginBottom: 20 }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: C.muted, marginBottom: 10 }}>📍 {site}</div>
                    <WOTable wos={wos} />
                  </div>
                ))}
                {Object.values(archivedGroups).every(g => g.length === 0) && (
                  <div style={{ textAlign: "center", padding: 20, color: C.muted, fontSize: 13 }}>{t(lang,"noWorkOrdersFound")}</div>
                )}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
function AssetEditModal({ data, onSave, onClose, lang, mheModels, sites, error }) {
  const [form, setForm] = useState({ ...data });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const brands = [...new Set(mheModels.map(m => m.brand).filter(Boolean))];
  const modelsForBrand = form.brand ? mheModels.filter(m => m.brand === form.brand) : mheModels;

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
        {error && <div style={{ background: C.red+"22", border: `1px solid ${C.red}44`, borderRadius: 8, padding: "10px 16px", marginBottom: 16, fontSize: 13, color: C.red }}>{error}</div>}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
          <Input label={t(lang,"assetName")} value={form.name||""} onChange={f("name")} />
          <Input label={t(lang,"category")} value={form.category||""} onChange={f("category")} />
          <Sel label={t(lang,"site")} value={form.location||""} onChange={f("location")} options={sites} />
          <Input label={t(lang,"owner")} value={form.owner||""} onChange={f("owner")} />
          <div>
            <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t(lang,"brand")}</div>
            <input list="brands-edit" value={form.brand||""} onChange={e => f("brand")(e.target.value)}
              placeholder="e.g. Jungheinrich"
              style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14, boxSizing: "border-box" }} />
            <datalist id="brands-edit">
              {brands.map(b => <option key={b} value={b} />)}
            </datalist>
          </div>
          <div>
            <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t(lang,"model")}</div>
            <input list="mhe-models-edit" value={form.model||""} onChange={e => handleModelSelect(e.target.value)}
              placeholder="e.g. ETV 216"
              style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14, boxSizing: "border-box" }} />
            <datalist id="mhe-models-edit">
              {modelsForBrand.map(m => <option key={m.model} value={m.model} />)}
            </datalist>
          </div>
          <Input label={t(lang,"serialNumber")} value={form.serial_number||""} onChange={f("serial_number")} />
          <Input label={t(lang,"manufactureDate")} value={form.manufacture_date||""} onChange={f("manufacture_date")} type="date" />
          <Input label={t(lang,"estValue")} value={form.value||""} onChange={f("value")} />
          <Sel label={t(lang,"status")} value={form.status||"Operational"} onChange={f("status")} options={["Operational","Under Maintenance","Degraded"]} />
          <Input label={t(lang,"pmFrequency")} value={form.pm_frequency||""} onChange={f("pm_frequency")} />
          <Input label={t(lang,"nextServiceDate")} value={form.next_service||""} onChange={f("next_service")} type="date" />
          <Input label={t(lang,"pmTask")} value={form.pm_task||""} onChange={f("pm_task")} />
          <Input label={t(lang,"invoiceNumber")} value={form.invoice_number||""} onChange={f("invoice_number")} />
          <Input label={t(lang,"poNumber")} value={form.po_number||""} onChange={f("po_number")} />
          <Input label={t(lang,"purchaseDate")} value={form.purchase_date||""} onChange={f("purchase_date")} type="date" />
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
function Assets({ assets, setAssets, loading, onAdd, isAdmin, isSupervisor, isMaintenance, vendors, lang, userRole, sites }) {
  const [showForm, setShowForm] = useState(false); const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null); const [editItem, setEditItem] = useState(null);
  const [deleteItem, setDeleteItem] = useState(null); const [selectedAsset, setSelectedAsset] = useState(null);
  const [siteFilter, setSiteFilter] = useState("All"); const [catFilter, setCatFilter] = useState("All"); const [ownerFilter, setOwnerFilter] = useState("All"); const [modelFilter, setModelFilter] = useState("All"); const [search, setSearch] = useState("");
  const [mheModels, setMheModels] = useState([]);
  useEffect(() => {
    supabase.from("mhe_models").select("brand, model, category, subcategory, technical_specs").order("brand").order("model")
      .then(({ data, error }) => {
        if (error) console.error("mhe_models load error:", error);
        setMheModels(data || []);
      });
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
  const [form, setForm] = useState({ name: "", category: "", location: "— Select Site —", value: "", owner: "", brand: "", model: "", serial_number: "", manufacture_date: "", technical_specs: "", next_service: "", pm_frequency: "1", pm_task: "", invoice_number: "", po_number: "", purchase_date: "" });
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
    if (form.serial_number && assets.some(a => a.serial_number === form.serial_number)) { setError(t(lang,"duplicateSerialNumber")); return; }
    setSaving(true); setError(null);
    const record = { id: uid("AST"), name: form.name, category: form.category, location: form.location, value: form.value, owner: form.owner||null, brand: form.brand||null, model: form.model||null, serial_number: form.serial_number||null, manufacture_date: form.manufacture_date||null, technical_specs: form.technical_specs||null, status: "Operational", last_service: TODAY, next_service: form.next_service||null, pm_frequency: parseInt(form.pm_frequency)||1, pm_task: form.pm_task||"Scheduled Maintenance", last_pm_date: null, invoice_number: form.invoice_number||null, po_number: form.po_number||null, purchase_date: form.purchase_date||null };
    const { error: err } = await supabase.from("assets").insert([record]);
    if (err) { setError(assetDbErrorMessage(err, lang)); } else { onAdd(record); setForm({ name: "", category: "", location: "— Select Site —", value: "", owner: "", brand: "", model: "", serial_number: "", manufacture_date: "", technical_specs: "", next_service: "", pm_frequency: "1", pm_task: "", invoice_number: "", po_number: "", purchase_date: "" }); setShowForm(false); }
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
  const saveEdit = async (updated) => {
    if (updated.serial_number && assets.some(a => a.id !== updated.id && a.serial_number === updated.serial_number)) { setError(t(lang,"duplicateSerialNumber")); return; }
    const { error: err } = await supabase.from("assets").update(updated).eq("id",updated.id);
    if (!err) { setAssets(prev => prev.map(a => a.id===updated.id?updated:a)); setEditItem(null); setError(null); } else setError(assetDbErrorMessage(err, lang));
  };
  const confirmDelete = async () => { await supabase.from("assets").delete().eq("id",deleteItem.id); setAssets(prev => prev.filter(a => a.id!==deleteItem.id)); setDeleteItem(null); };

  return (
    <div>
      {editItem && <AssetEditModal lang={lang} data={editItem} mheModels={mheModels} sites={sites} error={error} onSave={saveEdit} onClose={() => { setEditItem(null); setError(null); }} />}
      {deleteItem && <ConfirmDel lang={lang} name={deleteItem.name} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      {selectedAsset && <MaintenanceModal asset={selectedAsset} lang={lang} onClose={() => setSelectedAsset(null)} isAdmin={isAdmin} isSupervisor={isSupervisor} isMaintenance={isMaintenance} vendors={vendors} userRole={userRole} />}
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 18 }}>
        <input value={search} onChange={e => setSearch(e.target.value)} placeholder={t(lang,"searchAssets")} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 12px", color: C.text, fontSize: 13, flex: "1 1 180px" }} />
        <select value={siteFilter} onChange={e => setSiteFilter(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "7px 10px", color: C.text, fontSize: 12 }}>
          <option>{t(lang,"all")}</option>{sites.filter(s => s!=="— Select Site —").map(s => <option key={s}>{s}</option>)}
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
            <Sel label={t(lang,"site")} value={form.location} onChange={f("location")} options={sites} />
            <Input label={t(lang,"owner")} value={form.owner} onChange={f("owner")} placeholder="e.g. EPx Logistics" />
            <div>
            <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t(lang,"brand")}</div>
            <input list="brands-list" value={form.brand} onChange={e => f("brand")(e.target.value)}
              placeholder="e.g. Jungheinrich"
              style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14, boxSizing: "border-box" }} />
            <datalist id="brands-list">
              {[...new Set(mheModels.map(m => m.brand).filter(Boolean))].map(b => <option key={b} value={b} />)}
            </datalist>
          </div>
          <div>
            <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t(lang,"model")}</div>
            <input list="mhe-models-list" value={form.model} onChange={e => handleModelSelect(e.target.value)}
              placeholder="e.g. ETV 216"
              style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14, boxSizing: "border-box" }} />
            <datalist id="mhe-models-list">
              {(form.brand ? mheModels.filter(m => m.brand === form.brand) : mheModels).map(m => <option key={m.model} value={m.model} />)}
            </datalist>
          </div>
            <Input label={t(lang,"serialNumber")} value={form.serial_number} onChange={f("serial_number")} />
            <Input label={t(lang,"manufactureDate")} value={form.manufacture_date} onChange={f("manufacture_date")} type="date" />
            <Input label={t(lang,"estValue")} value={form.value} onChange={f("value")} />
            <Input label={t(lang,"nextServiceDate")} value={form.next_service} onChange={f("next_service")} type="date" />
            <Sel label={t(lang,"pmFrequency")} value={form.pm_frequency} onChange={f("pm_frequency")} options={["1","2","3","6","12"]} />
            <Input label={t(lang,"pmTask")} value={form.pm_task} onChange={f("pm_task")} />
            <Input label={t(lang,"invoiceNumber")} value={form.invoice_number} onChange={f("invoice_number")} />
            <Input label={t(lang,"poNumber")} value={form.po_number} onChange={f("po_number")} />
            <Input label={t(lang,"purchaseDate")} value={form.purchase_date} onChange={f("purchase_date")} type="date" />
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
                  {[[t(lang,"site"),a.location],[t(lang,"owner"),a.owner||"—"],[t(lang,"brand"),a.brand||"—"],[t(lang,"model"),a.model||"—"],[t(lang,"serialNumber"),a.serial_number||"—"],[t(lang,"value"),a.value||"—"],[t(lang,"pmEvery"),a.pm_frequency?`${a.pm_frequency} mo.`:"—"],[t(lang,"lastPM"),a.last_pm_date?fmtDate(a.last_pm_date):t(lang,"never")],[t(lang,"manufactureDate"),a.manufacture_date?fmtDate(a.manufacture_date):"—"],[t(lang,"invoiceNumber"),a.invoice_number||"—"],[t(lang,"poNumber"),a.po_number||"—"],[t(lang,"purchaseDate"),a.purchase_date?fmtDate(a.purchase_date):"—"]].map(([lbl,val]) => (
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

function PMUpload({ assets, onAssetsImported, onWorkOrdersGenerated, lang, sites }) {
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
    if (err) { setError(err.message); } else { setSuccess(`Generated ${newWOs.length}!`); onWorkOrdersGenerated(newWOs); }
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
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 24, marginBottom: 20 }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 6 }}>{t(lang,"importAssetsExcel")}</div>
        <div style={{ fontSize: 13, color: C.subtle, marginBottom: 16 }}>Columns: Asset Name, Site, Category, Value, PM Frequency (months), PM Task</div>
        <label style={{ display: "inline-block", background: C.accent, color: "#fff", borderRadius: 6, padding: "10px 20px", fontSize: 13, fontWeight: 700, cursor: uploading?"not-allowed":"pointer", opacity: uploading?0.7:1 }}>
          {uploading?"Importing...":"📂 Choose Excel File"}
          <input type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display: "none" }} disabled={uploading} />
        </label>
      </div>

      <AnnualPMPlanUpload assets={assets} lang={lang} sites={sites} />
    </div>
  );
}

function AnnualPMPlanUpload({ assets, lang, sites }) {
  const [importing, setImporting] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [siteCounts, setSiteCounts] = useState([]);
  const [uploadSite, setUploadSite] = useState("— Select Site —");
  const [genSite, setGenSite] = useState("All Sites");
  const [manageSite, setManageSite] = useState("");
  const [showManage, setShowManage] = useState(false);

  useEffect(() => { loadSiteCounts(); }, []);

  const loadSiteCounts = async () => {
    const { data } = await supabase.from("maintenance_plans").select("site").eq("active", true);
    const counts = {};
    (data||[]).forEach(r => { counts[r.site] = (counts[r.site]||0) + 1; });
    setSiteCounts(Object.entries(counts).map(([site,count]) => ({ site, count })).sort((a,b) => b.count-a.count));
  };

  const handleImport = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    if (uploadSite === "— Select Site —") { setError("Please select a site before uploading."); return; }
    const { count } = await supabase.from("maintenance_plans").select("id", { count: "exact", head: true }).eq("site", uploadSite);
    if (count > 0) {
      if (!window.confirm(`${uploadSite} already has ${count} plan entries. Replace them with this new upload?`)) {
        e.target.value = "";
        return;
      }
      await supabase.from("maintenance_plans").delete().eq("site", uploadSite);
    }
    setImporting(true); setError(null);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        const headerRowIdx = rows.findIndex(r => r[0] === "Equipment Name");
        if (headerRowIdx === -1) { setError("Could not find header row. Expected 'Equipment Name' in column A."); setImporting(false); return; }
        const weekHeaderRow = rows[headerRowIdx];
        const weekCols = [];
        for (let c = 7; c < weekHeaderRow.length; c++) {
          if (weekHeaderRow[c]) weekCols.push(c);
        }
        const dataRows = rows.slice(headerRowIdx + 1).filter(r => r[0]);

        const records = [];
        for (const r of dataRows) {
          const name = String(r[0] || "").trim();
          const code = String(r[1] || "").trim();
          const category = String(r[2] || "").trim();
          const site = uploadSite;
          if (!name) continue;

          // Match by unique code first (names like "Forklift" repeat across many assets)
          let found = code ? assets.find(a => a.serial_number === code) : assets.find(a => a.name.toLowerCase() === name.toLowerCase());
          if (!found) {
            const newAssetId = uid("AST");
            const newAsset = { id: newAssetId, name, category, location: site, status: "Operational", serial_number: code || null, last_service: TODAY, pm_frequency: 1, pm_task: "Scheduled Maintenance" };
            const { error: assetErr } = await supabase.from("assets").insert([newAsset]);
            if (!assetErr) { found = newAsset; assets.push(newAsset); }
          }
          // Find which codes appear in which week columns for this row
          const codesUsed = {};
          weekCols.forEach((c, weekIdx) => {
            const val = String(r[c] || "").trim().toUpperCase();
            if (val) {
              if (!codesUsed[val]) codesUsed[val] = weekIdx + 1;
            }
          });

          if (Object.keys(codesUsed).length === 0) {
            // default to Monthly if nothing marked
            codesUsed["M"] = 1;
          }

          Object.entries(codesUsed).forEach(([freq, firstWeek]) => {
            const startMonth = Math.ceil(firstWeek / 4.33) || 1;
            records.push({
              id: uid("PLN"),
              asset_id: found ? found.id : null,
              asset_name: name,
              equipment_code: code || null,
              category: category || null,
              site,
              task: "Scheduled Maintenance",
              frequency: freq,
              start_month: startMonth,
              start_week: firstWeek,
              active: true,
            });
          });
        }

        if (records.length === 0) { setError("No valid rows found."); setImporting(false); return; }
        const { error: err } = await supabase.from("maintenance_plans").insert(records);
        if (err) { setError(err.message); } else {
          setSuccess(`Imported ${records.length} maintenance plan entries for ${uploadSite}!`);
          loadSiteCounts();
        }
      } catch (ex) { setError("Failed to parse file: " + ex.message); }
      setImporting(false);
    };
    reader.readAsBinaryString(file);
  };

  const generateNow = async () => {
    setGenerating(true); setError(null);
    const siteParam = genSite === "All Sites" ? null : genSite;
    const { data, error: err } = await supabase.rpc("generate_due_pm_work_orders", { target_site: siteParam });
    if (err) { setError(err.message); } else {
      const created = data?.[0]?.created_count ?? 0;
      setSuccess(`Generated ${created} work order(s) for ${genSite}.`);
    }
    setGenerating(false);
  };

  const totalPlans = siteCounts.reduce((s,x) => s+x.count, 0);

  return (
    <div style={{ background: C.card, border: `1px solid ${C.purple}44`, borderRadius: 10, padding: 24 }}>
      <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 6 }}>📅 Annual Maintenance Plan</div>
      <div style={{ fontSize: 13, color: C.subtle, marginBottom: 16 }}>
        Upload your 52-week PM plan per site (Equipment Name, Code, Category, then W1-W52 columns marked W/M/Q/S/A). Work orders auto-generate weekly, or trigger manually per site below.
      </div>
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <OkBanner msg={success} onDismiss={() => setSuccess(null)} />

      {/* Per-site breakdown */}
      <div style={{ display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
        <div style={{ background: C.surface, borderRadius: 8, padding: "12px 20px", borderLeft: `3px solid ${C.purple}` }}>
          <div style={{ fontSize: 22, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{totalPlans}</div>
          <div style={{ fontSize: 12, color: C.muted }}>Total plan entries</div>
        </div>
        {siteCounts.map(s => (
          <div key={s.site} style={{ background: C.surface, borderRadius: 8, padding: "12px 20px", borderLeft: `3px solid ${C.blue}` }}>
            <div style={{ fontSize: 18, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{s.count}</div>
            <div style={{ fontSize: 12, color: C.muted }}>{s.site}</div>
          </div>
        ))}
      </div>

      {/* Upload */}
      <div style={{ background: C.surface, borderRadius: 8, padding: 16, marginBottom: 16 }}>
        <div style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", marginBottom: 10 }}>Upload Plan For Site</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <select value={uploadSite} onChange={e => setUploadSite(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "9px 12px", color: C.text, fontSize: 13, minWidth: 160 }}>
            {sites.map(s => <option key={s}>{s}</option>)}
          </select>
          <label style={{ display: "inline-block", background: C.purple, color: "#fff", borderRadius: 6, padding: "10px 20px", fontSize: 13, fontWeight: 700, cursor: importing?"not-allowed":"pointer", opacity: importing?0.7:1 }}>
            {importing ? "Importing..." : "📂 Upload Annual Plan"}
            <input type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display: "none" }} disabled={importing} />
          </label>
        </div>
      </div>
      {/* Manage */}
      <div style={{ background: C.surface, borderRadius: 8, padding: 16, marginBottom: 16 }}>
        <div style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", marginBottom: 10 }}>Manage Existing Plan</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <select value={manageSite} onChange={e => setManageSite(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "9px 12px", color: C.text, fontSize: 13, minWidth: 160 }}>
            <option value="">— Select Site —</option>
            {siteCounts.map(s => <option key={s.site}>{s.site}</option>)}
          </select>
          <Btn onClick={() => setShowManage(true)} disabled={!manageSite} color={C.blue}>📋 View / Edit Plan</Btn>
        </div>
      </div>
      {/* Generate */}
      <div style={{ background: C.surface, borderRadius: 8, padding: 16 }}>
        <div style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", marginBottom: 10 }}>Generate Work Orders</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <select value={genSite} onChange={e => setGenSite(e.target.value)} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "9px 12px", color: C.text, fontSize: 13, minWidth: 160 }}>
            <option>All Sites</option>
            {siteCounts.map(s => <option key={s.site}>{s.site}</option>)}
          </select>
          <Btn onClick={generateNow} disabled={generating} color={C.green}>{generating ? "Generating..." : "⚡ Generate Now"}</Btn>
        </div>
      </div>

      {showManage && <PlanManageModal site={manageSite} onClose={() => { setShowManage(false); loadSiteCounts(); }} lang={lang} />}
    </div>
  );
}

function PlanManageModal({ site, onClose, lang }) {
  const [plans, setPlans] = useState([]);
  const [loading, setLoading] = useState(true);
  const [editItem, setEditItem] = useState(null);
  const [showAdd, setShowAdd] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [form, setForm] = useState({ asset_name: "", equipment_code: "", category: "", task: "Scheduled Maintenance", frequency: "M", start_month: "1", start_week: "1" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));

  useEffect(() => { loadPlans(); }, []);

  const loadPlans = async () => {
    setLoading(true);
    const { data } = await supabase.from("maintenance_plans").select("*").eq("site", site).order("asset_name");
    setPlans(data || []);
    setLoading(false);
  };

  const savePlan = async () => {
    if (!form.asset_name) { setError("Asset name required."); return; }
    setSaving(true);
    const record = { id: uid("PLN"), asset_id: null, asset_name: form.asset_name, equipment_code: form.equipment_code||null, category: form.category||null, site, task: form.task, frequency: form.frequency, start_month: parseInt(form.start_month)||1, start_week: parseInt(form.start_week)||1, active: true };
    const { error: err } = await supabase.from("maintenance_plans").insert([record]);
    if (err) { setError(err.message); } else { setPlans(prev => [...prev, record]); setForm({ asset_name: "", equipment_code: "", category: "", task: "Scheduled Maintenance", frequency: "M", start_month: "1", start_week: "1" }); setShowAdd(false); }
    setSaving(false);
  };

  const updatePlan = async (updated) => {
    const { error: err } = await supabase.from("maintenance_plans").update(updated).eq("id", updated.id);
    if (!err) { setPlans(prev => prev.map(p => p.id===updated.id?updated:p)); setEditItem(null); } else setError(err.message);
  };

  const deletePlan = async (id) => {
    await supabase.from("maintenance_plans").delete().eq("id", id);
    setPlans(prev => prev.filter(p => p.id !== id));
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const rows = plans.map(p => ({
      "Asset Name": p.asset_name, "Equipment Code": p.equipment_code||"", "Category": p.category||"",
      "Task": p.task, "Frequency": p.frequency, "Start Month": p.start_month, "Start Week": p.start_week,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, "PM Plan");
    XLSX.writeFile(wb, `${site}_PM_Plan_Export.xlsx`);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#000000cc", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 1000, padding: 16, overflowY: "auto" }}>
      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, width: "100%", maxWidth: 1000, marginTop: 20, marginBottom: 20 }}>
        <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 700, color: C.text }}>📋 {site} — Maintenance Plan</div>
            <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{plans.length} entries</div>
          </div>
          <button onClick={onClose} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, color: C.muted, cursor: "pointer", fontSize: 18, padding: "2px 10px" }}>✕</button>
        </div>
        <div style={{ padding: 24 }}>
          <ErrBanner msg={error} onDismiss={() => setError(null)} />
          <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
            <Btn small onClick={() => setShowAdd(v => !v)}>+ Add Entry</Btn>
            <Btn small onClick={exportExcel} color={C.green}>📥 Export Excel</Btn>
          </div>

          {showAdd && (
            <div style={{ background: C.surface, border: `1px solid ${C.accent}44`, borderRadius: 8, padding: 16, marginBottom: 16 }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 10 }}>
                <Input label="Asset Name" value={form.asset_name} onChange={f("asset_name")} />
                <Input label="Code" value={form.equipment_code} onChange={f("equipment_code")} />
                <Input label="Category" value={form.category} onChange={f("category")} />
                <Input label="Task" value={form.task} onChange={f("task")} />
                <Sel label="Frequency" value={form.frequency} onChange={f("frequency")} options={["W","M","Q","S","A"]} />
                <Input label="Start Month" value={form.start_month} onChange={f("start_month")} type="number" />
                <Input label="Start Week" value={form.start_week} onChange={f("start_week")} type="number" />
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                <Btn small onClick={savePlan} disabled={saving}>Save</Btn>
                <Btn small variant="secondary" onClick={() => setShowAdd(false)}>Cancel</Btn>
              </div>
            </div>
          )}

          {loading ? <Spinner lang={lang} /> : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead><tr style={{ borderBottom: `1px solid ${C.border}` }}>
                  {["Asset Name","Code","Category","Task","Freq","Start Mo.","Start Wk","Last Gen","Actions"].map(h => <th key={h} style={{ textAlign: "left", padding: "8px 10px", fontSize: 11, color: C.muted, fontWeight: 600, textTransform: "uppercase" }}>{h}</th>)}
                </tr></thead>
                <tbody>
                  {plans.map((p,i) => (
                    <tr key={p.id} style={{ borderBottom: `1px solid ${C.border}22`, background: i%2===0?"transparent":C.surface+"44" }}>
                      <td style={{ padding: "8px 10px", color: C.text, fontWeight: 600 }}>{p.asset_name}</td>
                      <td style={{ padding: "8px 10px", color: C.subtle, fontFamily: "monospace", fontSize: 11 }}>{p.equipment_code||"—"}</td>
                      <td style={{ padding: "8px 10px", color: C.subtle }}>{p.category||"—"}</td>
                      <td style={{ padding: "8px 10px", color: C.subtle }}>{p.task}</td>
                      <td style={{ padding: "8px 10px" }}><Badge label={p.frequency} color={C.accent} /></td>
                      <td style={{ padding: "8px 10px", color: C.subtle }}>{p.start_month}</td>
                      <td style={{ padding: "8px 10px", color: C.subtle }}>{p.start_week}</td>
                      <td style={{ padding: "8px 10px", color: C.muted, fontSize: 11 }}>{p.last_generated_date||"—"}</td>
                      <td style={{ padding: "8px 10px" }}>
                        <div style={{ display: "flex", gap: 6 }}>
                          <Btn small onClick={() => setEditItem(p)} color={C.blue}>Edit</Btn>
                          <Btn small variant="danger" onClick={() => deletePlan(p.id)}>Del</Btn>
                        </div>
                      </td>
                    </tr>
                  ))}
                  {plans.length===0 && <tr><td colSpan={9} style={{ padding: 32, textAlign: "center", color: C.muted }}>No plan entries.</td></tr>}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
      {editItem && (
        <EditModal lang={lang} title="Plan Entry" data={editItem} fields={[{key:"asset_name",label:"Asset Name"},{key:"equipment_code",label:"Code"},{key:"category",label:"Category"},{key:"task",label:"Task"},{key:"frequency",label:"Frequency",options:["W","M","Q","S","A"]},{key:"start_month",label:"Start Month"},{key:"start_week",label:"Start Week"}]} onSave={updatePlan} onClose={() => setEditItem(null)} />
      )}
    </div>
  );
}
function LowStockAlerts({ lang, isSupervisor }) {
  const [alerts, setAlerts] = useState([]);
  useEffect(() => {
    Promise.all([
      supabase.from("asset_parts").select("*"),
      supabase.from("model_parts").select("*"),
    ]).then(([apRes, mpRes]) => {
      const ap = (apRes.data||[]).filter(p => (p.stock_quantity||0) <= (p.min_stock_level||1)).map(p => ({ ...p, source: "asset" }));
      const mp = (mpRes.data||[]).filter(p => (p.stock_quantity||0) <= (p.min_stock_level||1)).map(p => ({ ...p, asset_name: p.model, source: "model" }));
      setAlerts([...ap,...mp]);
    });
  }, []);
  if (!alerts.length || !isSupervisor) return null;
  return (
    <div style={{ background: C.yellow+"11", border: `1px solid ${C.yellow}44`, borderRadius: 10, padding: 16, marginBottom: 20 }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: C.yellow, marginBottom: 10 }}>⚠️ {t(lang,"lowStockAlerts")} ({alerts.length})</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
        {alerts.slice(0,5).map(a => (
          <div key={a.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: C.subtle }}>
            <span><strong style={{ color: C.text }}>{a.part_name}</strong> — {a.asset_name}</span>
            <span style={{ color: a.stock_quantity===0?C.red:C.yellow, fontWeight: 700 }}>{a.stock_quantity===0?t(lang,"outOfStock"):`${a.stock_quantity} left (min: ${a.min_stock_level})`}</span>
          </div>
        ))}
        {alerts.length > 5 && <div style={{ fontSize: 11, color: C.muted }}>+{alerts.length-5} more</div>}
      </div>
    </div>
  );
}
function Overview({ workOrders, assets, vendors, lang, isSupervisor }) {
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
      <LowStockAlerts lang={lang} isSupervisor={isSupervisor} />
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
          {[...new Set(assets.map(a => a.location).filter(Boolean))].sort().map(site => {
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

function Reports({ workOrders, assets, vendors, lang, issues }) {
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
  const totalIssues=(issues||[]).length; const resolvedIssues=(issues||[]).filter(i => i.status==="Resolved").length; const openIssues=(issues||[]).filter(i => i.status==="Open").length;
  const siteData=[...new Set(assets.map(a => a.location).filter(Boolean))].sort().map(site => ({ site, total: assets.filter(a => a.location===site).length, down: assets.filter(a => a.location===site&&a.status==="Under Maintenance").length })).filter(s => s.total>0);

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
            <div style={{ fontSize: 13, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14, marginTop: 16 }}>⚠️ {t(lang,"issues")}</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14, marginBottom: 16 }}>
              <KpiCard icon="⚠️" label={t(lang,"openIssues")} value={openIssues} color={C.yellow} />
              <KpiCard icon="✅" label={t(lang,"resolved")} value={resolvedIssues} color={C.green} percent={totalIssues>0?Math.round((resolvedIssues/totalIssues)*100):0} />
              <KpiCard icon="📋" label={t(lang,"allIssues")} value={totalIssues} color={C.blue} />
              <KpiCard icon="👁" label={t(lang,"acknowledged")} value={(issues||[]).filter(i=>i.status==="Acknowledged").length} color={C.purple} />
            </div>
            {breakdowns.length>0 && <BarChart title={t(lang,"breakdownAnalysis")} data={[...new Set(breakdowns.map(b => b.site).filter(Boolean))].sort().map(site => ({ label: site, count: breakdowns.filter(b => b.site===site).length, color: C.red })).filter(s => s.count>0)} labelKey="label" valueKey="count" colorKey="color" />}
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
function PartsCatalogMgmt({ lang }) {
  const [models, setModels] = useState([]);
  const [selectedModel, setSelectedModel] = useState(null);
  const [parts, setParts] = useState([]);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [importing, setImporting] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ part_name: "", part_number: "", supplier: "", unit_cost: "", notes: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));

  useEffect(() => {
    supabase.from("mhe_models").select("id, brand, model, category, subcategory").order("brand").order("model")
      .then(({ data }) => setModels(data || []));
  }, []);

  useEffect(() => {
    if (selectedModel) loadParts();
    else setParts([]);
  }, [selectedModel]);

  const loadParts = async () => {
    setLoading(true);
    const { data } = await supabase.from("model_parts").select("*").eq("model_id", selectedModel.id).order("part_name");
    setParts(data || []);
    setLoading(false);
  };

  const savePart = async () => {
    if (!form.part_name || !selectedModel) return;
    setSaving(true);
    const record = { id: uid("MPT"), model_id: selectedModel.id, model: selectedModel.model, brand: selectedModel.brand, part_name: form.part_name, part_number: form.part_number||null, supplier: form.supplier||null, unit_cost: parseFloat(form.unit_cost)||0, notes: form.notes||null };
    const { error: err } = await supabase.from("model_parts").insert([record]);
    if (err) { setError(err.message); } else { setParts(prev => [...prev, record]); setForm({ part_name: "", part_number: "", supplier: "", unit_cost: "", notes: "" }); setShowForm(false); setSuccess("Part added!"); }
    setSaving(false);
  };

  const deletePart = async (id) => {
    await supabase.from("model_parts").delete().eq("id", id);
    setParts(prev => prev.filter(p => p.id !== id));
  };

  const handleImport = (e) => {
    const file = e.target.files[0]; if (!file) return;
    setImporting(true); setError(null);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "binary" });
        const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        const records = [];
        for (const row of data) {
          const modelName = row["Model"] || row["model"] || "";
          const brand = row["Brand"] || row["brand"] || "";
          if (!modelName || !row["Part Name"]) continue;
          // Find model in DB
          const found = models.find(m => m.model.toLowerCase() === modelName.toLowerCase());
          if (!found) continue;
          records.push({ id: uid("MPT"), model_id: found.id, model: found.model, brand: found.brand || brand, part_name: row["Part Name"] || "", part_number: row["Part Number"] || null, supplier: row["Supplier"] || null, unit_cost: parseFloat(row["Unit Cost"]) || 0, notes: row["Notes"] || null });
        }
        if (records.length === 0) { setError("No valid rows found. Check column names."); setImporting(false); return; }
        const { error: err } = await supabase.from("model_parts").insert(records);
        if (err) { setError(err.message); } else { setSuccess(`${t(lang,"partImported")} (${records.length} parts)`); if (selectedModel) loadParts(); }
      } catch { setError("Failed to parse file."); }
      setImporting(false);
    };
    reader.readAsBinaryString(file);
  };

  // Group models by brand
  const brandGroups = models.reduce((acc, m) => {
    if (!acc[m.brand]) acc[m.brand] = [];
    acc[m.brand].push(m);
    return acc;
  }, {});

  return (
    <div>
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <OkBanner msg={success} onDismiss={() => setSuccess(null)} />

      {/* Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20, flexWrap: "wrap", gap: 10 }}>
        <div style={{ fontSize: 18, fontWeight: 700, color: C.text }}>🔩 {t(lang,"partsCatalogMgmt")}</div>
        <div style={{ display: "flex", gap: 10 }}>
          <label style={{ display: "inline-flex", alignItems: "center", gap: 8, background: C.green+"22", color: C.green, border: `1px solid ${C.green}44`, borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: importing?"not-allowed":"pointer", opacity: importing?0.7:1 }}>
            {importing ? t(lang,"importing") : `📥 ${t(lang,"importPartsExcel")}`}
            <input type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display: "none" }} disabled={importing} />
          </label>
        </div>
      </div>

      {/* Excel format hint */}
      <div style={{ background: C.blue+"11", border: `1px solid ${C.blue}33`, borderRadius: 8, padding: "10px 14px", fontSize: 12, color: C.blue, marginBottom: 20 }}>
        📋 {t(lang,"excelFormat")}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "280px 1fr", gap: 20, alignItems: "start" }}>
        {/* Model List */}
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
          <div style={{ padding: "14px 16px", borderBottom: `1px solid ${C.border}`, fontSize: 13, fontWeight: 700, color: C.text }}>📦 {t(lang,"selectModel")}</div>
          <div style={{ maxHeight: 600, overflowY: "auto" }}>
            {Object.entries(brandGroups).map(([brand, brandModels]) => (
              <div key={brand}>
                <div style={{ padding: "8px 16px", fontSize: 11, color: C.muted, fontWeight: 700, textTransform: "uppercase", background: C.surface, letterSpacing: "0.07em" }}>{brand}</div>
                {brandModels.map(m => {
                  const isSelected = selectedModel?.id === m.id;
                  return (
                    <div key={m.id} onClick={() => setSelectedModel(m)} style={{ padding: "10px 16px", cursor: "pointer", background: isSelected?C.accent+"22":"transparent", borderLeft: `3px solid ${isSelected?C.accent:"transparent"}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div>
                        <div style={{ fontSize: 13, fontWeight: isSelected?700:400, color: isSelected?C.accent:C.text }}>{m.model}</div>
                        <div style={{ fontSize: 11, color: C.muted }}>{m.subcategory || m.category}</div>
                      </div>
                    </div>
                  );
                })}
              </div>
            ))}
          </div>
        </div>

        {/* Parts Panel */}
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 20 }}>
          {!selectedModel ? (
            <div style={{ textAlign: "center", padding: 60, color: C.muted, fontSize: 13 }}>
              <div style={{ fontSize: 40, marginBottom: 12 }}>🔩</div>
              {t(lang,"noModelSelected")}
            </div>
          ) : (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <div>
                  <div style={{ fontSize: 16, fontWeight: 700, color: C.text }}>{selectedModel.brand} {selectedModel.model}</div>
                  <div style={{ fontSize: 12, color: C.muted }}>{selectedModel.subcategory} · {parts.length} {t(lang,"partsCount")}</div>
                </div>
                <Btn onClick={() => setShowForm(v => !v)}>{t(lang,"addPartToModel")}</Btn>
              </div>

              {showForm && (
                <div style={{ background: C.surface, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 16, marginBottom: 16 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 10 }}>
                    <Input label={t(lang,"partName")} value={form.part_name} onChange={f("part_name")} />
                    <Input label={t(lang,"partNumber")} value={form.part_number} onChange={f("part_number")} />
                    <Input label={t(lang,"supplier")} value={form.supplier} onChange={f("supplier")} />
                    <Input label={t(lang,"unitCostLabel")} value={form.unit_cost} onChange={f("unit_cost")} type="number" />
                    <Input label={t(lang,"descriptionNotes")} value={form.notes} onChange={f("notes")} />
                  </div>
                  <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                    <Btn onClick={savePart} disabled={saving}>{saving?t(lang,"saving"):t(lang,"save")}</Btn>
                    <Btn variant="secondary" onClick={() => setShowForm(false)}>{t(lang,"cancel")}</Btn>
                  </div>
                </div>
              )}

              {loading ? <Spinner lang={lang} /> : parts.length === 0 ? (
                <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>{t(lang,"noPartsInCatalog")}</div>
              ) : (
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead><tr style={{ borderBottom: `1px solid ${C.border}` }}>
                      {[t(lang,"partName"),t(lang,"partNumber"),t(lang,"unitCostLabel"),t(lang,"currentStock"),t(lang,"minStockLevel"),t(lang,"supplier"),t(lang,"descriptionNotes"),""].map(h => <th key={h} style={{ textAlign: "left", padding: "8px 12px", fontSize: 11, color: C.muted, fontWeight: 600, textTransform: "uppercase" }}>{h}</th>)}
                    </tr></thead>
                    <tbody>
                      {parts.map((part, i) => (
                        <tr key={part.id} style={{ borderBottom: `1px solid ${C.border}22`, background: i%2===0?"transparent":C.surface+"44" }}>
                          <td style={{ padding: "10px 12px", color: C.text, fontWeight: 600 }}>{part.part_name}</td>
                          <td style={{ padding: "10px 12px", color: C.subtle, fontFamily: "monospace", fontSize: 11 }}>{part.part_number||"—"}</td>
                          <td style={{ padding: "10px 12px", color: C.accent, fontWeight: 700 }}>{part.unit_cost?`$${part.unit_cost}`:"—"}</td>
                          <td style={{ padding: "10px 12px" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                              <input type="number" value={part.stock_quantity||0} onChange={async e => {
                                const qty = parseFloat(e.target.value)||0;
                                await supabase.from("model_parts").update({ stock_quantity: qty }).eq("id", part.id);
                                setParts(prev => prev.map(p => p.id===part.id?{...p,stock_quantity:qty}:p));
                              }} style={{ width: 60, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 4, padding: "3px 6px", color: C.text, fontSize: 12 }} />
                              {(part.stock_quantity||0)===0 && <span style={{ fontSize: 10, color: C.red, fontWeight: 700 }}>{t(lang,"outOfStock")}</span>}
                              {(part.stock_quantity||0)>0 && (part.stock_quantity||0)<=(part.min_stock_level||1) && <span style={{ fontSize: 10, color: C.yellow, fontWeight: 700 }}>{t(lang,"lowStock")}</span>}
                            </div>
                          </td>
                          <td style={{ padding: "10px 12px" }}>
                            <input type="number" value={part.min_stock_level||1} onChange={async e => {
                              const qty = parseFloat(e.target.value)||1;
                              await supabase.from("model_parts").update({ min_stock_level: qty }).eq("id", part.id);
                              setParts(prev => prev.map(p => p.id===part.id?{...p,min_stock_level:qty}:p));
                            }} style={{ width: 50, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 4, padding: "3px 6px", color: C.text, fontSize: 12 }} />
                          </td>
                          <td style={{ padding: "10px 12px", color: C.subtle }}>{part.supplier||"—"}</td>
                          <td style={{ padding: "10px 12px", color: C.muted, fontSize: 12 }}>{part.notes||"—"}</td>
                          <td style={{ padding: "10px 12px" }}><Btn small variant="danger" onClick={() => deletePart(part.id)}>{t(lang,"del")}</Btn></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
function MySubmissions({ userRole, lang }) {
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [expandedLog, setExpandedLog] = useState(null);

  useEffect(() => { loadMine(); }, []);

  const loadMine = async () => {
    setLoading(true);
    const { data } = await supabase.from("maintenance_logs").select("*").eq("performed_by", userRole.name).order("start_date", { ascending: false }).limit(50);
    setLogs(data || []);
    setLoading(false);
  };

  const pending = logs.filter(l => l.approval_status === "Pending");
  const approved = logs.filter(l => l.approval_status === "Approved");
  const rejected = logs.filter(l => l.approval_status === "Rejected");

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div style={{ fontSize: 18, fontWeight: 700, color: C.text }}>📝 {t(lang,"mySubmissions")}</div>
        <button onClick={loadMine} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "6px 12px", color: C.muted, cursor: "pointer", fontSize: 12 }}>↻ {t(lang,"refresh")}</button>
      </div>

      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        <StatCard icon="⏳" label={t(lang,"pendingApproval")} value={pending.length} color={C.yellow} />
        <StatCard icon="✅" label={t(lang,"approved")} value={approved.length} color={C.green} />
        <StatCard icon="❌" label={t(lang,"rejected")} value={rejected.length} color={C.red} />
      </div>

      {loading ? <Spinner lang={lang} /> : logs.length === 0 ? (
        <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>{t(lang,"noSubmissions")}</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {logs.map(log => {
            const color = log.approval_status === "Approved" ? C.green : log.approval_status === "Rejected" ? C.red : C.yellow;
            return (
              <div key={log.id} style={{ background: C.card, border: `1px solid ${color}44`, borderRadius: 10, overflow: "hidden" }}>
                <div onClick={() => setExpandedLog(expandedLog===log.id?null:log.id)} style={{ padding: 14, cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                  <div>
                    <div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{log.title}</div>
                    <div style={{ fontSize: 11, color: C.muted }}>{log.asset_name} · {fmtDate(log.start_date)}</div>
                  </div>
                  <Badge label={log.approval_status||"—"} color={color} />
                </div>
                {expandedLog===log.id && (
                  <div style={{ padding: "0 14px 14px", borderTop: `1px solid ${C.border}` }}>
                    {log.description && <div style={{ marginTop: 10, fontSize: 13, color: C.subtle, lineHeight: 1.6, background: C.surface, borderRadius: 8, padding: 12 }}>{log.description}</div>}
                    {log.approval_status === "Approved" && log.approved_by && (
                      <div style={{ marginTop: 10, background: C.green+"11", border: `1px solid ${C.green}33`, borderRadius: 8, padding: "10px 14px", fontSize: 12, color: C.green }}>
                        ✅ Approved by <strong>{log.approved_by}</strong> · {log.approved_at ? fmtDateTime(log.approved_at) : ""}
                        {log.approved_signature && <div style={{ marginTop: 4, fontStyle: "italic" }}>Signed: "{log.approved_signature}"</div>}
                      </div>
                    )}
                    {log.approval_status === "Rejected" && log.rejection_notes && (
                      <div style={{ marginTop: 10, background: C.red+"11", border: `1px solid ${C.red}33`, borderRadius: 8, padding: "10px 14px", fontSize: 12, color: C.red }}>
                        ❌ {log.rejection_notes}
                      </div>
                    )}
                    {log.approval_status === "Pending" && (
                      <div style={{ marginTop: 10, background: C.yellow+"11", border: `1px solid ${C.yellow}33`, borderRadius: 8, padding: "10px 14px", fontSize: 12, color: C.yellow }}>
                        ⏳ Waiting for supervisor approval
                      </div>
                    )}
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
function PendingApprovals({ userRole, isAdmin, lang, assets, vendors, onJumpToBreakdowns }) {
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [expandedLog, setExpandedLog] = useState(null);
  const [parts, setParts] = useState({});

  useEffect(() => { loadPending(); }, []);

  const loadPending = async () => {
    setLoading(true);
    const { data } = await supabase.from("maintenance_logs").select("*").eq("approval_status", "Pending").order("start_date", { ascending: false });
    setLogs(data || []);
    setLoading(false);
  };

  const loadParts = async (logId) => {
    const { data } = await supabase.from("spare_parts").select("*").eq("log_id", logId);
    setParts(prev => ({ ...prev, [logId]: data || [] }));
  };

  const toggleLog = (logId) => {
    setExpandedLog(expandedLog===logId?null:logId);
    if (!parts[logId]) loadParts(logId);
  };

  // Determine scope
  const supervisedSites = userRole.supervised_sites;
  const supervisedCategories = userRole.supervised_categories;

  const isInScope = (log) => {
    if (isAdmin) return true;
    const asset = (assets||[]).find(a => a.id === log.asset_id) || (assets||[]).find(a => a.name === log.asset_name);
    const site = asset?.location;
    const category = asset?.category;
    const siteMatch = !supervisedSites?.length || (site && supervisedSites.includes(site));
    const categoryMatch = !supervisedCategories?.length || (category && supervisedCategories.includes(category));
    return siteMatch && categoryMatch;
  };

  const inScopeLogs = logs.filter(isInScope);
  const outOfScopeLogs = logs.filter(l => !isInScope(l));

  const LogCard = ({ log, dimmed }) => (
    <div style={{ background: C.card, border: `1px solid ${dimmed?C.border:C.yellow+"44"}`, borderRadius: 10, overflow: "hidden", opacity: dimmed?0.6:1 }}>
      <div onClick={() => toggleLog(log.id)} style={{ padding: "14px 18px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 18 }}>{log.log_type==="Preventive Maintenance"?"🔧":log.log_type==="Corrective Repair"?"🔨":"🔍"}</span>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{log.title}</div>
            <div style={{ fontSize: 11, color: C.muted }}>{log.asset_name} · {fmtDate(log.start_date)} · {log.performed_by}</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {dimmed && <Badge label={t(lang,"outOfScope")} color={C.muted} />}
          <Badge label={t(lang,"pendingApproval")} color={C.yellow} />
          {log.cost>0 && <span style={{ fontSize: 12, color: C.accent, fontWeight: 700 }}>${log.cost}</span>}
          <span style={{ color: C.muted }}>{expandedLog===log.id?"▲":"▼"}</span>
        </div>
      </div>
      {expandedLog===log.id && (
        <div style={{ padding: "0 18px 18px", borderTop: `1px solid ${C.border}` }}>
          {log.description && <div style={{ marginTop: 12, padding: 12, background: C.surface, borderRadius: 8, fontSize: 13, color: C.subtle, lineHeight: 1.6 }}>{log.description}</div>}
          {parts[log.id]?.length > 0 && (
            <div style={{ marginTop: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.text, marginBottom: 8 }}>{t(lang,"spareParts")}</div>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <tbody>
                  {parts[log.id].map(p => (
                    <tr key={p.id} style={{ borderBottom: `1px solid ${C.border}22` }}>
                      <td style={{ padding: "6px 8px", color: C.text }}>{p.part_name}</td>
                      <td style={{ padding: "6px 8px", color: C.subtle }}>x{p.quantity}</td>
                      <td style={{ padding: "6px 8px", color: C.accent, fontWeight: 700 }}>${p.total_cost||0}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          {log.breakdown_id || log.issue_id ? (
            <div style={{ marginTop: 12, background: C.blue+"11", border: `1px solid ${C.blue}33`, borderRadius: 8, padding: 12, fontSize: 12, color: C.blue }}>
              <div style={{ marginBottom: 8 }}>ℹ️ This log is linked to a {log.breakdown_id ? "breakdown" : "issue"}. Approve it from the <strong>Breakdowns & Issues</strong> tab so the operator confirmation step isn't skipped.</div>
              <Btn small onClick={onJumpToBreakdowns} color={C.blue}>{t(lang,"goToBreakdownsTab")}</Btn>
            </div>
          ) : (isAdmin || !dimmed) && (
            <ApprovalSection log={log} lang={lang} userRole={userRole} onApproved={() => setLogs(prev => prev.filter(l => l.id!==log.id))} onRejected={() => setLogs(prev => prev.filter(l => l.id!==log.id))} />
          )}
        </div>
      )}
    </div>
  );

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <div style={{ fontSize: 18, fontWeight: 700, color: C.text }}>⏳ {t(lang,"pendingApprovalsSection")}</div>
        <button onClick={loadPending} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "6px 12px", color: C.muted, cursor: "pointer", fontSize: 12 }}>↻ {t(lang,"refresh")}</button>
      </div>

      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        <StatCard icon="⏳" label={t(lang,"pendingApprovalsSection")} value={logs.length} color={C.yellow} />
        <StatCard icon="📍" label={t(lang,"supervisedSites")} value={supervisedSites?.length || t(lang,"allSitesScope")} color={C.blue} />
        <StatCard icon="🔧" label={t(lang,"supervisedCategories")} value={supervisedCategories?.length || t(lang,"allCategoriesScope")} color={C.purple} />
      </div>

      {loading ? <Spinner lang={lang} /> : logs.length === 0 ? (
        <div style={{ textAlign: "center", padding: 40, color: C.muted, fontSize: 13 }}>{t(lang,"noPendingApprovals")}</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {inScopeLogs.map(log => <LogCard key={log.id} log={log} dimmed={false} />)}
          {outOfScopeLogs.length > 0 && (
            <>
              <div style={{ fontSize: 12, color: C.muted, marginTop: 10, marginBottom: 4, textTransform: "uppercase" }}>{t(lang,"outOfScope")}</div>
              {outOfScopeLogs.map(log => <LogCard key={log.id} log={log} dimmed={true} />)}
            </>
          )}
        </div>
      )}
    </div>
  );
}
function UserManagement({ lang, sites }) {
  const [users, setUsers] = useState([]); const [loading, setLoading] = useState(true); const [showForm, setShowForm] = useState(false); const [saving, setSaving] = useState(false); const [error, setError] = useState(null); const [success, setSuccess] = useState(null);
  const [identifierType, setIdentifierType] = useState("email"); // "email" | "phone"
  const [form, setForm] = useState({ email: "", phone: "", password: "", name: "", role: "operations", site: "", supervised_sites: [], supervised_categories: [] });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const resetForm = () => { setForm({ email: "", phone: "", password: "", name: "", role: "operations", site: "", supervised_sites: [], supervised_categories: [] }); setIdentifierType("email"); };

  useEffect(() => { loadUsers(); }, []);
  const loadUsers = async () => { setLoading(true); const { data } = await supabase.from("user_roles").select("*").order("name"); setUsers(data||[]); setLoading(false); };

  const submit = async () => {
    if (!form.name) { setError(t(lang,"fullName")); return; }
    if (identifierType === "email" && !form.email) { setError(t(lang,"email")); return; }
    if (identifierType === "phone" && !form.phone) { setError(t(lang,"phone")); return; }
    setSaving(true); setError(null);
    const scopedSites = form.role==="supervisor" ? (form.supervised_sites.length?form.supervised_sites:null) : null;
    const scopedCategories = form.role==="supervisor" ? (form.supervised_categories.length?form.supervised_categories:null) : null;

    if (form.password) {
      // Brand-new login — must go through the Edge Function since only it holds the service-role key.
      const { data: sessionData } = await supabase.auth.getSession();
      const token = sessionData?.session?.access_token;
      try {
        const res = await fetch("https://evwsdzqgvrwbjusjmrdc.supabase.co/functions/v1/create-user", {
          method: "POST",
          headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
          body: JSON.stringify({
            email: identifierType === "email" ? form.email : null,
            phone: identifierType === "phone" ? form.phone : null,
            password: form.password, name: form.name, role: form.role, site: form.site || null,
            supervised_sites: scopedSites, supervised_categories: scopedCategories,
          }),
        });
        const result = await res.json();
        if (!res.ok) { setError(result.error || "Failed to create account."); setSaving(false); return; }
      } catch { setError("Failed to reach account creation service."); setSaving(false); return; }
    } else {
      // No password — just updating an existing account's role/permissions.
      let record;
      if (identifierType === "phone") {
        const existing = users.find(u => u.phone === form.phone);
        if (!existing) { setError(t(lang,"noAccountForPhone")); setSaving(false); return; }
        record = { ...existing, name: form.name, role: form.role, site: form.site||null, supervised_sites: scopedSites, supervised_categories: scopedCategories };
      } else {
        record = { id: uid("USR"), email: form.email, phone: form.phone||null, name: form.name, role: form.role, site: form.site||null, supervised_sites: scopedSites, supervised_categories: scopedCategories };
      }
      const { error: err } = await supabase.from("user_roles").upsert([record], { onConflict: "email" });
      if (err) { setError(err.message); setSaving(false); return; }
    }
    setSuccess("✓"); resetForm(); setShowForm(false); loadUsers();
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
        {[["🏭",t(lang,"operationsRole"),"operations",t(lang,"operationsDesc"),C.green],["🔧",t(lang,"maintenancePersonnelRole"),"maintenance",t(lang,"maintenancePersonnelDesc"),C.blue],["👁",t(lang,"supervisorRole"),"supervisor",t(lang,"supervisorDesc"),C.purple],["★",t(lang,"adminRole"),"admin",t(lang,"adminDesc"),C.accent]].map(([icon,title,role,desc,color]) => (
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
          <div style={{ marginBottom: 12 }}>
            <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{t(lang,"identifierType")}</div>
            <div style={{ display: "flex", gap: 8 }}>
              {["email","phone"].map(m => (
                <button key={m} onClick={() => setIdentifierType(m)} style={{ background: identifierType===m?C.accent:C.surface, color: identifierType===m?"#fff":C.muted, border: `1px solid ${identifierType===m?C.accent:C.border}`, borderRadius: 6, padding: "8px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                  {m==="email"?t(lang,"email"):t(lang,"phone")}
                </button>
              ))}
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            {identifierType === "phone"
              ? <Input label={t(lang,"phone")} value={form.phone} onChange={f("phone")} type="tel" placeholder="01012345678" />
              : <Input label={t(lang,"email")} value={form.email} onChange={f("email")} type="email" />}
            <Input label={t(lang,"fullName")} value={form.name} onChange={f("name")} />
            <Sel label={t(lang,"role")} value={form.role} onChange={f("role")} options={["operations","maintenance","supervisor","admin"]} />
            <Sel label={t(lang,"defaultSite")} value={form.site||"— Select Site —"} onChange={f("site")} options={["— Select Site —",...sites.filter(s => s !== "— Select Site —")]} />
          </div>
          <div style={{ marginTop: 12 }}>
            <Input label={t(lang,"password")} value={form.password} onChange={f("password")} type="password" />
            <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>{t(lang,"newAccountPasswordHint")}</div>
          </div>
          {form.role === "supervisor" && (
            <div style={{ marginTop: 14, background: C.surface, border: `1px solid ${C.purple}33`, borderRadius: 8, padding: 14 }}>
              <div style={{ fontSize: 12, color: C.purple, fontWeight: 700, marginBottom: 4, textTransform: "uppercase" }}>{t(lang,"supervisedSites")} / {t(lang,"supervisedCategories")}</div>
              <div style={{ fontSize: 11, color: C.muted, marginBottom: 10 }}>{t(lang,"scopeNote")}</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                <div>
                  <div style={{ fontSize: 11, color: C.muted, marginBottom: 6 }}>{t(lang,"supervisedSites")}</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4, maxHeight: 150, overflowY: "auto" }}>
                    {sites.filter(s => s !== "— Select Site —").map(s => (
                      <label key={s} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: C.subtle, cursor: "pointer" }}>
                        <input type="checkbox" checked={form.supervised_sites.includes(s)} onChange={e => {
                          setForm(p => ({ ...p, supervised_sites: e.target.checked ? [...p.supervised_sites, s] : p.supervised_sites.filter(x => x !== s) }));
                        }} />
                        {s}
                      </label>
                    ))}
                  </div>
                </div>
                <div>
                  <div style={{ fontSize: 11, color: C.muted, marginBottom: 6 }}>{t(lang,"supervisedCategories")}</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4, maxHeight: 150, overflowY: "auto" }}>
                    {["MHE","HVAC","Fire Alarm & Suppression","Electrical","Plumbing","Civil & Structural","Security Systems","Lighting","General Maintenance"].map(c => (
                      <label key={c} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: C.subtle, cursor: "pointer" }}>
                        <input type="checkbox" checked={form.supervised_categories.includes(c)} onChange={e => {
                          setForm(p => ({ ...p, supervised_categories: e.target.checked ? [...p.supervised_categories, c] : p.supervised_categories.filter(x => x !== c) }));
                        }} />
                        {c}
                      </label>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          )}
          <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
            <Btn onClick={submit} disabled={saving}>{saving?(form.password?t(lang,"creatingAccount"):t(lang,"saving")):t(lang,"saveUser")}</Btn>
            <Btn variant="secondary" onClick={() => { setShowForm(false); resetForm(); }}>{t(lang,"cancel")}</Btn>
          </div>
        </div>
      )}
      {loading ? <Spinner lang={lang} /> : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 12 }}>
          {users.map(u => (
            <div key={u.id} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 18 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
                <div><div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{u.name}</div><div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>{u.phone ? `📱 ${u.phone}` : u.email}</div></div>
                <span style={{ background: roleColor(u.role)+"22", color: roleColor(u.role), border: `1px solid ${roleColor(u.role)}44`, borderRadius: 4, padding: "3px 10px", fontSize: 12, fontWeight: 700, textTransform: "uppercase" }}>{roleIcon(u.role)} {u.role}</span>
              </div>
              {u.site && <div style={{ fontSize: 12, color: C.muted, marginBottom: 8 }}>📍 {u.site}</div>}
              {u.role === "supervisor" && (
                <div style={{ fontSize: 11, color: C.purple, marginBottom: 12 }}>
                  {u.supervised_sites?.length ? `🏭 ${u.supervised_sites.join(", ")}` : `🏭 ${t(lang,"allSitesScope")}`}<br/>
                  {u.supervised_categories?.length ? `🔧 ${u.supervised_categories.join(", ")}` : `🔧 ${t(lang,"allCategoriesScope")}`}
                </div>
              )}
              <Btn small variant="danger" onClick={() => deleteUser(u.id)}>{t(lang,"remove")}</Btn>
            </div>
          ))}
          {users.length===0 && <div style={{ color: C.muted, fontSize: 13 }}>{t(lang,"noUsersRegistered")}</div>}
        </div>
      )}
    </div>
  );
}
function SitesManagement({ sites, setSites, lang }) {
  const [newName, setNewName] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [renameId, setRenameId] = useState(null);
  const [renameValue, setRenameValue] = useState("");

  const addSite = async () => {
    const name = newName.trim();
    if (!name) return;
    setSaving(true); setError(null);
    const { data, error: err } = await supabase.from("sites").insert([{ name }]).select().single();
    if (err) { setError(err.code === "23505" ? t(lang,"duplicateSiteName") : err.message); }
    else { setSites(prev => [...prev, data].sort((a,b) => a.name.localeCompare(b.name))); setNewName(""); }
    setSaving(false);
  };

  const toggleActive = async (site) => {
    const { error: err } = await supabase.from("sites").update({ active: !site.active }).eq("id", site.id);
    if (!err) setSites(prev => prev.map(s => s.id === site.id ? { ...s, active: !site.active } : s));
  };

  const saveRename = async (site) => {
    const name = renameValue.trim();
    if (!name || name === site.name) { setRenameId(null); return; }
    const { error: err } = await supabase.from("sites").update({ name }).eq("id", site.id);
    if (err) { setError(err.code === "23505" ? t(lang,"duplicateSiteName") : err.message); }
    else { setSites(prev => prev.map(s => s.id === site.id ? { ...s, name } : s).sort((a,b) => a.name.localeCompare(b.name))); setRenameId(null); }
  };

  return (
    <div>
      <ErrBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ marginBottom: 20 }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: C.text, marginBottom: 4 }}>{t(lang,"sitesManagement")}</div>
        <div style={{ fontSize: 13, color: C.muted }}>{t(lang,"manageSites")}</div>
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
        <input value={newName} onChange={e => setNewName(e.target.value)} placeholder={t(lang,"siteNamePlaceholder")}
          onKeyDown={e => e.key === "Enter" && addSite()}
          style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "9px 12px", color: C.text, fontSize: 13, flex: "1 1 220px" }} />
        <Btn onClick={addSite} disabled={saving || !newName.trim()}>{t(lang,"addSite")}</Btn>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))", gap: 10 }}>
        {sites.map(s => (
          <div key={s.id} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, opacity: s.active ? 1 : 0.55 }}>
            {renameId === s.id ? (
              <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
                <input value={renameValue} onChange={e => setRenameValue(e.target.value)} onKeyDown={e => e.key === "Enter" && saveRename(s)}
                  style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "6px 8px", color: C.text, fontSize: 13, flex: 1 }} autoFocus />
              </div>
            ) : (
              <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 8 }}>{s.name}</div>
            )}
            <div style={{ fontSize: 11, color: C.muted, marginBottom: 10 }}>
              {s.active ? t(lang,"active") : t(lang,"inactiveStatus")}
            </div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {renameId === s.id ? (
                <>
                  <Btn small onClick={() => saveRename(s)} color={C.blue}>{t(lang,"save")}</Btn>
                  <Btn small variant="secondary" onClick={() => setRenameId(null)}>{t(lang,"cancel")}</Btn>
                </>
              ) : (
                <Btn small onClick={() => { setRenameId(s.id); setRenameValue(s.name); }} color={C.blue}>{t(lang,"edit")}</Btn>
              )}
              <Btn small variant={s.active ? "danger" : "secondary"} onClick={() => toggleActive(s)}>
                {s.active ? t(lang,"deactivate") : t(lang,"activate")}
              </Btn>
            </div>
          </div>
        ))}
        {sites.length===0 && <div style={{ color: C.muted, fontSize: 13 }}>{t(lang,"noSitesRegistered")}</div>}
      </div>
      <div style={{ fontSize: 11, color: C.muted, marginTop: 16 }}>
        ⚠️ Renaming a site only affects new entries — assets, work orders, and reports already assigned to the old name will keep showing the old name.
      </div>
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState(null); const [authLoading, setAuthLoading] = useState(true);
  const [userRole, setUserRole] = useState({ role: "operations", name: "", site: "", language: "en" });
  const [lang, setLang] = useState("en");
  const [tab, setTab] = useState(null);
  const [workOrders, setWorkOrders] = useState([]); const [assets, setAssets] = useState([]); const [vendors, setVendors] = useState([]); const [issues, setIssues] = useState([]);
  const [sites, setSites] = useState([]);
  const [loading, setLoading] = useState({ workOrders: true, assets: true, vendors: true });
  const [globalError, setGlobalError] = useState(null);
  const isAdmin = session?.user?.email === ADMIN_EMAIL || userRole.role === "admin";
  const isSupervisor = userRole.role === "supervisor" || isAdmin;
  const isMaintenance = userRole.role === "maintenance" || userRole.role === "supervisor" || isAdmin;

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
    const [woRes, astRes, vndRes, issRes, siteRes] = await Promise.all([
      supabase.from("work_orders").select("*").order("due", { ascending: true }),
      supabase.from("assets").select("*").order("name", { ascending: true }),
      supabase.from("vendors").select("*").order("name", { ascending: true }),
      supabase.from("issue_reports").select("*").order("reported_at", { ascending: false }),
      supabase.from("sites").select("*").order("name", { ascending: true }),
    ]);
    if (woRes.error||astRes.error||vndRes.error) { setGlobalError("Failed to load data."); }
    else { setWorkOrders(woRes.data||[]); setAssets(astRes.data||[]); setVendors(vndRes.data||[]); setIssues(issRes.data||[]); }
    // sites table may not exist yet if the Phase 1 migration hasn't been run — fall back silently.
    setSites(siteRes.error ? [] : (siteRes.data||[]));
    setLoading({ workOrders: false, assets: false, vendors: false });
  }, []);

  useEffect(() => { if (session) load(); }, [session, load]);
  const signOut = async () => { await supabase.auth.signOut(); setSession(null); };

  if (authLoading) return <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", color: C.muted }}>Loading...</div>;
  if (!session) return <LoginScreen lang={lang} />;

  const siteNames = sites.length ? ["— Select Site —", ...sites.filter(s => s.active).map(s => s.name)] : DEFAULT_SITES;

  const roleColor = { admin: C.accent, maintenance: C.blue, operations: C.green }[userRole.role] || C.muted;
  const roleIcon = { admin: "★", maintenance: "🔧", operations: "🏭" }[userRole.role] || "👤";
  const roleLabel = { admin: t(lang,"adminRole"), maintenance: t(lang,"maintenancePersonnelRole"), supervisor: t(lang,"supervisorRole"), operations: t(lang,"operationsRole") }[userRole.role] || userRole.role;

  const tabs = [
    t(lang,"overview"),
    t(lang,"breakdownsAndIssues"),
    ...(isSupervisor ? [t(lang,"pendingApprovalsSection")] : []),
    ...(userRole.role === "maintenance" ? [t(lang,"mySubmissions")] : []),
    ...(isMaintenance ? [t(lang,"workOrders"), t(lang,"assets"), t(lang,"vendors"), t(lang,"pmPlanner"), t(lang,"reports"), t(lang,"calendar")] : []),
    ...(isAdmin ? [t(lang,"partsCatalogMgmt"), t(lang,"users"), t(lang,"sitesManagement")] : []),
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
        {activeTab===t(lang,"overview") && <Overview workOrders={workOrders} assets={assets} vendors={vendors} lang={lang} isSupervisor={isSupervisor} />}
        {activeTab===t(lang,"pendingApprovalsSection") && <PendingApprovals userRole={userRole} isAdmin={isAdmin} lang={lang} assets={assets} vendors={vendors} onJumpToBreakdowns={() => setTab(t(lang,"breakdownsAndIssues"))} />}
        {activeTab===t(lang,"mySubmissions") && <MySubmissions userRole={userRole} lang={lang} />}
        {activeTab===t(lang,"breakdownsAndIssues") && <Breakdowns userRole={userRole} assets={assets} setAssets={setAssets} vendors={vendors} workOrders={workOrders} setWorkOrders={setWorkOrders} lang={lang} setIssuesFromParent={setIssues} isMaintenance={isMaintenance} isSupervisor={isSupervisor} />}
        {activeTab===t(lang,"workOrders") && <WorkOrders workOrders={workOrders} setWorkOrders={setWorkOrders} loading={loading.workOrders} onAdd={r => setWorkOrders(p => [r,...p])} isAdmin={isAdmin} isSupervisor={isSupervisor} isMaintenance={isMaintenance} vendors={vendors} assets={assets} lang={lang} userRole={userRole} sites={siteNames} />}
        {activeTab===t(lang,"assets") && <Assets assets={assets} setAssets={setAssets} loading={loading.assets} onAdd={r => setAssets(p => [r,...p])} isAdmin={isAdmin} isSupervisor={isSupervisor} isMaintenance={isMaintenance} vendors={vendors} lang={lang} userRole={userRole} sites={siteNames} />}
        {activeTab===t(lang,"vendors") && <Vendors vendors={vendors} setVendors={setVendors} loading={loading.vendors} onAdd={r => setVendors(p => [r,...p])} isAdmin={isAdmin} lang={lang} />}
        {activeTab===t(lang,"pmPlanner") && <PMUpload assets={assets} onAssetsImported={r => setAssets(p => [...p,...r])} onWorkOrdersGenerated={r => setWorkOrders(p => [...r,...p])} lang={lang} sites={siteNames} />}
        {activeTab===t(lang,"reports") && <Reports workOrders={workOrders} assets={assets} vendors={vendors} lang={lang} issues={issues} isSupervisor={isSupervisor} />}
        {activeTab===t(lang,"calendar") && <MaintenanceCalendar workOrders={workOrders} assets={assets} lang={lang} />}
        {activeTab===t(lang,"partsCatalogMgmt") && isAdmin && <PartsCatalogMgmt lang={lang} />}
        {activeTab===t(lang,"users") && isAdmin && <UserManagement lang={lang} sites={siteNames} />}
        {activeTab===t(lang,"sitesManagement") && isAdmin && <SitesManagement sites={sites} setSites={setSites} lang={lang} />}
      </div>
    </div>
  );
}



