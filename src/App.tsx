import { useState, useEffect, useCallback } from "react";
import { supabase } from "./supabase";

const C = {
  bg: "#0d0f12", surface: "#141720", card: "#1a1e2a", border: "#252b3b",
  accent: "#f97316", accentDim: "#7c3b12", yellow: "#eab308",
  green: "#22c55e", red: "#ef4444", blue: "#3b82f6",
  text: "#e2e8f0", muted: "#64748b", subtle: "#94a3b8",
};

const SITES = [
  "— Select Site —",
  "Site1", "Site2", "Site3", "Site4", "Site5", "Site6",
  "Site7B", "Site7C", "Site8", "Site9A", "Site9B",
  "Site10A", "Site10B", "Site11", "Site12", "Site14A", "Site14B",
];

const ADMIN_EMAIL = import.meta.env.VITE_ADMIN_EMAIL;
const TODAY = new Date().toISOString().split("T")[0];
const uid = (prefix) => `${prefix}-${Date.now().toString(36).toUpperCase()}`;
const priorityColor = (p) => ({ Critical: C.red, High: C.accent, Medium: C.yellow, Low: C.green }[p] || C.muted);
const statusColor = (s) => ({
  Open: C.accent, "In Progress": C.blue, Completed: C.green,
  Pending: C.yellow, Operational: C.green, "Under Maintenance": C.accent,
  Degraded: C.red, Active: C.green, Inactive: C.muted,
}[s] || C.muted);

const Badge = ({ label, color }) => (
  <span style={{
    background: color + "22", color, border: `1px solid ${color}44`,
    borderRadius: 4, padding: "2px 8px", fontSize: 11, fontWeight: 700,
    letterSpacing: "0.05em", textTransform: "uppercase", whiteSpace: "nowrap",
  }}>{label}</span>
);

const Input = ({ label, value, onChange, type = "text" }) => (
  <div>
    <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
    <input type={type} value={value} onChange={e => onChange(e.target.value)}
      style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14, boxSizing: "border-box" }} />
  </div>
);

const SelectInput = ({ label, value, onChange, options }) => (
  <div>
    <div style={{ fontSize: 11, color: C.muted, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
    <select value={value} onChange={e => onChange(e.target.value)}
      style={{ width: "100%", background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: "10px", color: C.text, fontSize: 14 }}>
      {options.map(o => <option key={o}>{o}</option>)}
    </select>
  </div>
);

const Btn = ({ children, onClick, variant, disabled, color }) => {
  const isPrimary = variant !== "secondary" && variant !== "danger";
  const isDanger = variant === "danger";
  return (
    <button onClick={onClick} disabled={disabled} style={{
      background: isDanger ? C.red + "22" : isPrimary ? (color || C.accent) : "transparent",
      color: isDanger ? C.red : isPrimary ? "#fff" : C.muted,
      border: isDanger ? `1px solid ${C.red}44` : isPrimary ? "none" : `1px solid ${C.border}`,
      borderRadius: 6, padding: "7px 14px", fontSize: 13, fontWeight: 700,
      cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.5 : 1,
    }}>{children}</button>
  );
};

const StatCard = ({ icon, label, value, sub, color }) => (
  <div style={{
    background: C.card, border: `1px solid ${C.border}`, borderRadius: 10,
    padding: "16px 20px", flex: "1 1 140px", borderLeft: `3px solid ${color || C.accent}`,
  }}>
    <div style={{ fontSize: 20, marginBottom: 6 }}>{icon}</div>
    <div style={{ fontSize: 26, fontWeight: 800, color: C.text, fontFamily: "monospace" }}>{value}</div>
    <div style={{ fontSize: 13, color: C.subtle, marginTop: 2 }}>{label}</div>
    {sub && <div style={{ fontSize: 11, color: color || C.accent, marginTop: 4 }}>{sub}</div>}
  </div>
);

const Spinner = () => (
  <div style={{ textAlign: "center", padding: 48, color: C.muted, fontSize: 13 }}>Loading...</div>
);

const ErrorBanner = ({ msg, onDismiss }) => msg ? (
  <div style={{ background: C.red + "22", border: `1px solid ${C.red}44`, borderRadius: 8, padding: "10px 16px", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, color: C.red }}>
    {msg}
    <button onClick={onDismiss} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 16 }}>x</button>
  </div>
) : null;

const StatusSelect = ({ value, options, onChange }) => (
  <select value={value} onChange={e => onChange(e.target.value)} style={{
    background: statusColor(value) + "22", color: statusColor(value),
    border: `1px solid ${statusColor(value)}44`, borderRadius: 4,
    padding: "3px 6px", fontSize: 11, fontWeight: 700, cursor: "pointer",
    textTransform: "uppercase", letterSpacing: "0.05em",
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
        <div style={{ fontSize: 15, fontWeight: 700, color: C.accent, marginBottom: 20, textTransform: "uppercase", letterSpacing: "0.08em" }}>Edit {title}</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
          {fields.map(({ key, label, type, options }) =>
            options ? (
              <SelectInput key={key} label={label} value={form[key] || ""} onChange={f(key)} options={options} />
            ) : (
              <Input key={key} label={label} value={form[key] || ""} onChange={f(key)} type={type || "text"} />
            )
          )}
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 20, flexWrap: "wrap" }}>
          <Btn onClick={() => onSave(form)}>Save Changes</Btn>
          <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
        </div>
      </div>
    </div>
  );
};

const ConfirmDelete = ({ name, onConfirm, onClose }) => (
  <div style={{ position: "fixed", inset: 0, background: "#000000aa", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }}>
    <div style={{ background: C.card, border: `1px solid ${C.red}44`, borderRadius: 12, padding: 24, width: "100%", maxWidth: 400 }}>
      <div style={{ fontSize: 15, fontWeight: 700, color: C.red, marginBottom: 12 }}>Confirm Delete</div>
      <div style={{ fontSize: 13, color: C.subtle, marginBottom: 20 }}>Are you sure you want to delete <strong style={{ color: C.text }}>{name}</strong>? This cannot be undone.</div>
      <div style={{ display: "flex", gap: 8 }}>
        <Btn variant="danger" onClick={onConfirm}>Delete</Btn>
        <Btn variant="secondary" onClick={onClose}>Cancel</Btn>
      </div>
    </div>
  </div>
);

function LoginScreen() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

  const signIn = async () => {
    if (!email || !password) { setError("Please enter your email and password."); return; }
    setLoading(true); setError(null);
    const { error: err } = await supabase.auth.signInWithPassword({ email, password });
    if (err) setError(err.message);
    setLoading(false);
  };

  return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans', 'Helvetica Neue', sans-serif", padding: 16 }}>
      <div style={{ width: "100%", maxWidth: 400 }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ background: C.accent, borderRadius: 12, width: 56, height: 56, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 28, margin: "0 auto 16px" }}>🏭</div>
          <div style={{ fontFamily: "monospace", fontSize: 20, letterSpacing: 3, color: C.text, fontWeight: 800 }}>FACILITY COMMAND</div>
          <div style={{ fontSize: 12, color: C.muted, marginTop: 4, letterSpacing: "0.1em" }}>INDUSTRIAL WAREHOUSE MANAGEMENT</div>
        </div>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 24 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.text, marginBottom: 20 }}>Sign in to your account</div>
          <ErrorBanner msg={error} onDismiss={() => setError(null)} />
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <Input label="Email" value={email} onChange={setEmail} type="email" />
            <Input label="Password" value={password} onChange={setPassword} type="password" />
            <button onClick={signIn} disabled={loading} style={{
              background: C.accent, color: "#fff", border: "none", borderRadius: 6,
              padding: "12px", fontSize: 15, fontWeight: 700,
              cursor: loading ? "not-allowed" : "pointer", opacity: loading ? 0.7 : 1, marginTop: 4,
            }}>
              {loading ? "Signing in..." : "Sign In"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function WorkOrders({ workOrders, setWorkOrders, loading, onAdd, isAdmin, vendors }) {
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [filter, setFilter] = useState("All");
  const [editItem, setEditItem] = useState(null);
  const [deleteItem, setDeleteItem] = useState(null);
  const [form, setForm] = useState({ title: "", asset: "", priority: "Medium", start_date: "", due: "", vendor: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));
  const filtered = filter === "All" ? workOrders : workOrders.filter(w => w.status === filter);
  const vendorOptions = ["— None —", ...vendors.filter(v => v.status === "Active").map(v => v.name)];

  const submit = async () => {
    if (!form.title || !form.asset) { setError("Title and Asset are required."); return; }
    setSaving(true); setError(null);
    const record = {
      id: uid("WO"), title: form.title, asset: form.asset,
      priority: form.priority, status: "Open", assignee: null,
      start_date: form.start_date || null,
      due: form.due || null,
      vendor: form.vendor === "— None —" ? null : form.vendor || null,
    };
    const { error: err } = await supabase.from("work_orders").insert([record]);
    if (err) { setError(err.message); } else {
      onAdd(record);
      setForm({ title: "", asset: "", priority: "Medium", start_date: "", due: "", vendor: "" });
      setShowForm(false);
    }
    setSaving(false);
  };

  const updateStatus = async (id, newStatus) => {
    const { error: err } = await supabase.from("work_orders").update({ status: newStatus }).eq("id", id);
    if (!err) setWorkOrders(prev => prev.map(wo => wo.id === id ? { ...wo, status: newStatus } : wo));
  };

  const updatePriority = async (id, newPriority) => {
    const { error: err } = await supabase.from("work_orders").update({ priority: newPriority }).eq("id", id);
    if (!err) setWorkOrders(prev => prev.map(wo => wo.id === id ? { ...wo, priority: newPriority } : wo));
  };

  const saveEdit = async (updated) => {
    const { error: err } = await supabase.from("work_orders").update(updated).eq("id", updated.id);
    if (!err) { setWorkOrders(prev => prev.map(wo => wo.id === updated.id ? updated : wo)); setEditItem(null); }
    else setError(err.message);
  };

  const confirmDelete = async () => {
    const { error: err } = await supabase.from("work_orders").delete().eq("id", deleteItem.id);
    if (!err) { setWorkOrders(prev => prev.filter(wo => wo.id !== deleteItem.id)); setDeleteItem(null); }
    else setError(err.message);
  };

  return (
    <div>
      {editItem && (
        <EditModal
          title="Work Order"
          data={editItem}
          fields={[
            { key: "title", label: "Title" },
            { key: "asset", label: "Asset / Location" },
            { key: "priority", label: "Priority", options: ["Critical", "High", "Medium", "Low"] },
            { key: "status", label: "Status", options: ["Open", "In Progress", "Pending", "Completed"] },
            { key: "vendor", label: "Vendor", options: vendorOptions },
            { key: "assignee", label: "Assignee" },
            { key: "start_date", label: "Start Date", type: "date" },
            { key: "due", label: "Due Date", type: "date" },
          ]}
          onSave={saveEdit}
          onClose={() => setEditItem(null)}
        />
      )}
      {deleteItem && <ConfirmDelete name={deleteItem.title} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      <ErrorBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {["All", "Open", "In Progress", "Pending", "Completed"].map(s => (
            <button key={s} onClick={() => setFilter(s)} style={{ background: filter === s ? C.accent : C.card, color: filter === s ? "#fff" : C.muted, border: `1px solid ${filter === s ? C.accent : C.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>{s}</button>
          ))}
        </div>
        <Btn onClick={() => setShowForm(v => !v)}>+ New Work Order</Btn>
      </div>
      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 18 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, letterSpacing: "0.08em", textTransform: "uppercase" }}>New Work Order</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label="Title *" value={form.title} onChange={f("title")} />
            <Input label="Asset *" value={form.asset} onChange={f("asset")} />
            <Input label="Start Date" value={form.start_date} onChange={f("start_date")} type="date" />
            <Input label="Due Date" value={form.due} onChange={f("due")} type="date" />
            <SelectInput label="Priority" value={form.priority} onChange={f("priority")} options={["Critical", "High", "Medium", "Low"]} />
            <SelectInput label="Vendor" value={form.vendor} onChange={f("vendor")} options={vendorOptions} />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
            <Btn onClick={submit} disabled={saving}>{saving ? "Saving..." : "Create"}</Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>Cancel</Btn>
          </div>
        </div>
      )}
      {loading ? <Spinner /> : (
        <div style={{ overflowX: "auto", WebkitOverflowScrolling: "touch" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 700 }}>
            <thead>
              <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                {["ID", "Title", "Asset", "Priority", "Status", "Vendor", "Start", "Due", ...(isAdmin ? ["Actions"] : [])].map(h => (
                  <th key={h} style={{ textAlign: "left", padding: "8px 12px", fontSize: 11, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((wo, i) => (
                <tr key={wo.id} style={{ borderBottom: `1px solid ${C.border}22`, background: i % 2 === 0 ? "transparent" : C.surface + "44" }}>
                  <td style={{ padding: "10px 12px", fontSize: 11, color: C.muted, fontFamily: "monospace" }}>{wo.id}</td>
                  <td style={{ padding: "10px 12px", fontSize: 13, color: C.text, fontWeight: 600 }}>{wo.title}</td>
                  <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.asset}</td>
                  <td style={{ padding: "10px 12px" }}>
                    <StatusSelect value={wo.priority} options={["Critical", "High", "Medium", "Low"]} onChange={(val) => updatePriority(wo.id, val)} />
                  </td>
                  <td style={{ padding: "10px 12px" }}>
                    <StatusSelect value={wo.status} options={["Open", "In Progress", "Pending", "Completed"]} onChange={(val) => updateStatus(wo.id, val)} />
                  </td>
                  <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.vendor || "—"}</td>
                  <td style={{ padding: "10px 12px", fontSize: 12, color: C.subtle }}>{wo.start_date || "—"}</td>
                  <td style={{ padding: "10px 12px", fontSize: 12, color: wo.due && wo.due <= TODAY && wo.status !== "Completed" ? C.red : C.subtle }}>{wo.due || "—"}</td>
                  {isAdmin && (
                    <td style={{ padding: "10px 12px" }}>
                      <div style={{ display: "flex", gap: 6 }}>
                        <Btn onClick={() => setEditItem(wo)} color={C.blue}>Edit</Btn>
                        <Btn variant="danger" onClick={() => setDeleteItem(wo)}>Delete</Btn>
                      </div>
                    </td>
                  )}
                </tr>
              ))}
              {filtered.length === 0 && <tr><td colSpan={isAdmin ? 9 : 8} style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>No work orders found.</td></tr>}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function Assets({ assets, setAssets, loading, onAdd, isAdmin }) {
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [editItem, setEditItem] = useState(null);
  const [deleteItem, setDeleteItem] = useState(null);
  const [form, setForm] = useState({ name: "", category: "", location: "— Select Site —", value: "", next_service: "" });
  const f = (k) => (v) => setForm(p => ({ ...p, [k]: v }));

  const submit = async () => {
    if (!form.name) { setError("Asset name is required."); return; }
    if (form.location === "— Select Site —") { setError("Please select a site."); return; }
    setSaving(true); setError(null);
    const record = {
      id: uid("AST"), name: form.name, category: form.category,
      location: form.location, value: form.value,
      status: "Operational", last_service: TODAY,
      next_service: form.next_service || null,
    };
    const { error: err } = await supabase.from("assets").insert([record]);
    if (err) { setError(err.message); } else {
      onAdd(record);
      setForm({ name: "", category: "", location: "— Select Site —", value: "", next_service: "" });
      setShowForm(false);
    }
    setSaving(false);
  };

  const updateStatus = async (id, newStatus) => {
    const { error: err } = await supabase.from("assets").update({ status: newStatus }).eq("id", id);
    if (!err) setAssets(prev => prev.map(a => a.id === id ? { ...a, status: newStatus } : a));
  };

  const saveEdit = async (updated) => {
    const { error: err } = await supabase.from("assets").update(updated).eq("id", updated.id);
    if (!err) { setAssets(prev => prev.map(a => a.id === updated.id ? updated : a)); setEditItem(null); }
    else setError(err.message);
  };

  const confirmDelete = async () => {
    const { error: err } = await supabase.from("assets").delete().eq("id", deleteItem.id);
    if (!err) { setAssets(prev => prev.filter(a => a.id !== deleteItem.id)); setDeleteItem(null); }
    else setError(err.message);
  };

  return (
    <div>
      {editItem && (
        <EditModal
          title="Asset"
          data={editItem}
          fields={[
            { key: "name", label: "Asset Name" },
            { key: "category", label: "Category" },
            { key: "location", label: "Site / Location", options: SITES },
            { key: "value", label: "Est. Value" },
            { key: "status", label: "Status", options: ["Operational", "Under Maintenance", "Degraded"] },
            { key: "last_service", label: "Last Service", type: "date" },
            { key: "next_service", label: "Next Service", type: "date" },
          ]}
          onSave={saveEdit}
          onClose={() => setEditItem(null)}
        />
      )}
      {deleteItem && <ConfirmDelete name={deleteItem.name} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      <ErrorBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 18 }}>
        <Btn onClick={() => setShowForm(v => !v)}>+ Add Asset</Btn>
      </div>
      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 18 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, letterSpacing: "0.08em", textTransform: "uppercase" }}>Register New Asset</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label="Asset Name *" value={form.name} onChange={f("name")} />
            <Input label="Category" value={form.category} onChange={f("category")} />
            <SelectInput label="Site / Location *" value={form.location} onChange={f("location")} options={SITES} />
            <Input label="Est. Value" value={form.value} onChange={f("value")} />
            <Input label="Next Service Date" value={form.next_service} onChange={f("next_service")} type="date" />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
            <Btn onClick={submit} disabled={saving}>{saving ? "Saving..." : "Register"}</Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>Cancel</Btn>
          </div>
        </div>
      )}
      {loading ? <Spinner /> : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 14 }}>
          {assets.map(a => (
            <div key={a.id} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 18, borderTop: `3px solid ${statusColor(a.status)}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{a.name}</div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>{a.id} · {a.category}</div>
                </div>
                <StatusSelect value={a.status} options={["Operational", "Under Maintenance", "Degraded"]} onChange={(val) => updateStatus(a.id, val)} />
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12 }}>
                {[["Site", a.location], ["Value", a.value], ["Last Service", a.last_service], ["Next Service", a.next_service]].map(([lbl, val]) => (
                  <div key={lbl}>
                    <div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.06em" }}>{lbl}</div>
                    <div style={{ color: C.subtle, marginTop: 2 }}>{val || "—"}</div>
                  </div>
                ))}
              </div>
              {isAdmin && (
                <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
                  <Btn onClick={() => setEditItem(a)} color={C.blue}>Edit</Btn>
                  <Btn variant="danger" onClick={() => setDeleteItem(a)}>Delete</Btn>
                </div>
              )}
            </div>
          ))}
          {assets.length === 0 && <div style={{ color: C.muted, fontSize: 13, padding: 32 }}>No assets registered yet.</div>}
        </div>
      )}
    </div>
  );
}

function Vendors({ vendors, setVendors, loading, onAdd, isAdmin }) {
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [editItem, setEditItem] = useState(null);
  const [deleteItem, setDeleteItem] = useState(null);
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

  const saveEdit = async (updated) => {
    const { error: err } = await supabase.from("vendors").update(updated).eq("id", updated.id);
    if (!err) { setVendors(prev => prev.map(v => v.id === updated.id ? updated : v)); setEditItem(null); }
    else setError(err.message);
  };

  const confirmDelete = async () => {
    const { error: err } = await supabase.from("vendors").delete().eq("id", deleteItem.id);
    if (!err) { setVendors(prev => prev.filter(v => v.id !== deleteItem.id)); setDeleteItem(null); }
    else setError(err.message);
  };

  const Stars = ({ rating }) => (
    <div style={{ display: "flex", alignItems: "center", gap: 2 }}>
      {[1, 2, 3, 4, 5].map(i => <span key={i} style={{ color: i <= Math.floor(rating) ? C.yellow : C.border, fontSize: 14 }}>*</span>)}
      <span style={{ fontSize: 11, color: C.muted, marginLeft: 4 }}>{rating > 0 ? Number(rating).toFixed(1) : "N/A"}</span>
    </div>
  );

  return (
    <div>
      {editItem && (
        <EditModal
          title="Vendor"
          data={editItem}
          fields={[
            { key: "name", label: "Company Name" },
            { key: "specialty", label: "Specialty" },
            { key: "contact", label: "Contact Person" },
            { key: "phone", label: "Phone" },
            { key: "email", label: "Email" },
            { key: "status", label: "Status", options: ["Active", "Inactive"] },
            { key: "rating", label: "Rating (0-5)" },
          ]}
          onSave={saveEdit}
          onClose={() => setEditItem(null)}
        />
      )}
      {deleteItem && <ConfirmDelete name={deleteItem.name} onConfirm={confirmDelete} onClose={() => setDeleteItem(null)} />}
      <ErrorBanner msg={error} onDismiss={() => setError(null)} />
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 18 }}>
        <Btn onClick={() => setShowForm(v => !v)}>+ Add Vendor</Btn>
      </div>
      {showForm && (
        <div style={{ background: C.card, border: `1px solid ${C.accent}44`, borderRadius: 10, padding: 20, marginBottom: 18 }}>
          <div style={{ color: C.accent, fontWeight: 700, marginBottom: 14, fontSize: 13, letterSpacing: "0.08em", textTransform: "uppercase" }}>Register Vendor / Contractor</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <Input label="Company Name *" value={form.name} onChange={f("name")} />
            <Input label="Specialty" value={form.specialty} onChange={f("specialty")} />
            <Input label="Contact Person" value={form.contact} onChange={f("contact")} />
            <Input label="Phone" value={form.phone} onChange={f("phone")} />
            <Input label="Email" value={form.email} onChange={f("email")} />
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
            <Btn onClick={submit} disabled={saving}>{saving ? "Saving..." : "Register"}</Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>Cancel</Btn>
          </div>
        </div>
      )}
      {loading ? <Spinner /> : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 14 }}>
          {vendors.map(v => (
            <div key={v.id} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{v.name}</div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>{v.specialty}</div>
                </div>
                <Badge label={v.status} color={statusColor(v.status)} />
              </div>
              <Stars rating={v.rating} />
              <div style={{ marginTop: 12, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12 }}>
                {[["Contact", v.contact], ["Phone", v.phone], ["Email", v.email]].map(([lbl, val]) => (
                  <div key={lbl}>
                    <div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase" }}>{lbl}</div>
                    <div style={{ color: C.subtle, marginTop: 2 }}>{val || "—"}</div>
                  </div>
                ))}
                <div>
                  <div style={{ color: C.muted, fontSize: 10, textTransform: "uppercase" }}>Open Orders</div>
                  <div style={{ color: v.open_orders > 0 ? C.accent : C.subtle, marginTop: 2, fontWeight: v.open_orders > 0 ? 700 : 400 }}>{v.open_orders}</div>
                </div>
              </div>
              {isAdmin && (
                <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
                  <Btn onClick={() => setEditItem(v)} color={C.blue}>Edit</Btn>
                  <Btn variant="danger" onClick={() => setDeleteItem(v)}>Delete</Btn>
                </div>
              )}
            </div>
          ))}
          {vendors.length === 0 && <div style={{ color: C.muted, fontSize: 13, padding: 32 }}>No vendors registered yet.</div>}
        </div>
      )}
    </div>
  );
}

function Overview({ workOrders, assets, vendors }) {
  const open = workOrders.filter(w => w.status !== "Completed").length;
  const critical = workOrders.filter(w => w.priority === "Critical").length;
  const opAssets = assets.filter(a => a.status === "Operational").length;
  const activeVendors = vendors.filter(v => v.status === "Active").length;
  const overdue = workOrders.filter(w => w.due && w.due <= TODAY && w.status !== "Completed").length;

  return (
    <div>
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        <StatCard icon="🔧" label="Open Work Orders" value={open} sub={`${critical} critical`} color={C.accent} />
        <StatCard icon="🏭" label="Operational Assets" value={`${opAssets}/${assets.length}`} sub="fleet status" color={C.green} />
        <StatCard icon="🤝" label="Active Vendors" value={activeVendors} sub="contractors on file" color={C.blue} />
        <StatCard icon="⚠️" label="Overdue / At Risk" value={overdue} sub="past due date" color={C.red} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 16 }}>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
          <div style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>Recent Work Orders</div>
          {workOrders.slice(0, 5).map(wo => (
            <div key={wo.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${C.border}22`, flexWrap: "wrap", gap: 6 }}>
              <div>
                <div style={{ fontSize: 13, color: C.text, fontWeight: 600 }}>{wo.title}</div>
                <div style={{ fontSize: 11, color: C.muted }}>{wo.asset} {wo.vendor ? `· ${wo.vendor}` : ""}</div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <Badge label={wo.priority} color={priorityColor(wo.priority)} />
                <Badge label={wo.status} color={statusColor(wo.status)} />
              </div>
            </div>
          ))}
          {workOrders.length === 0 && <div style={{ color: C.muted, fontSize: 13 }}>No work orders yet.</div>}
        </div>
        <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 20 }}>
          <div style={{ fontSize: 12, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 14 }}>Asset Status Breakdown</div>
          {[["Operational", C.green], ["Under Maintenance", C.accent], ["Degraded", C.red]].map(([status, color]) => {
            const count = assets.filter(a => a.status === status).length;
            const pct = assets.length > 0 ? Math.round((count / assets.length) * 100) : 0;
            return (
              <div key={status} style={{ marginBottom: 12 }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4, fontSize: 12 }}>
                  <span style={{ color: C.subtle }}>{status}</span>
                  <span style={{ color, fontWeight: 700 }}>{count} ({pct}%)</span>
                </div>
                <div style={{ background: C.border, borderRadius: 4, height: 6 }}>
                  <div style={{ background: color, width: `${pct}%`, height: 6, borderRadius: 4 }} />
                </div>
              </div>
            );
          })}
          <div style={{ marginTop: 20, fontSize: 12, color: C.muted, textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 600, marginBottom: 10 }}>Top Vendors by Rating</div>
          {[...vendors].filter(v => v.rating > 0).sort((a, b) => b.rating - a.rating).slice(0, 3).map(v => (
            <div key={v.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <span style={{ fontSize: 13, color: C.subtle }}>{v.name}</span>
              <span style={{ fontSize: 13, color: C.yellow, fontWeight: 700 }}>* {Number(v.rating).toFixed(1)}</span>
            </div>
          ))}
          {vendors.filter(v => v.rating > 0).length === 0 && <div style={{ color: C.muted, fontSize: 13 }}>No rated vendors yet.</div>}
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [tab, setTab] = useState("Overview");
  const [workOrders, setWorkOrders] = useState([]);
  const [assets, setAssets] = useState([]);
  const [vendors, setVendors] = useState([]);
  const [loading, setLoading] = useState({ workOrders: true, assets: true, vendors: true });
  const [globalError, setGlobalError] = useState(null);

  const isAdmin = session?.user?.email === ADMIN_EMAIL;

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session); setAuthLoading(false);
    });
    supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session); setAuthLoading(false);
    });
  }, []);

  const load = useCallback(async () => {
    setLoading({ workOrders: true, assets: true, vendors: true });
    const [woRes, astRes, vndRes] = await Promise.all([
      supabase.from("work_orders").select("*").order("due", { ascending: true }),
      supabase.from("assets").select("*").order("name", { ascending: true }),
      supabase.from("vendors").select("*").order("name", { ascending: true }),
    ]);
    if (woRes.error || astRes.error || vndRes.error) {
      setGlobalError("Failed to load data from Supabase.");
    } else {
      setWorkOrders(woRes.data || []);
      setAssets(astRes.data || []);
      setVendors(vndRes.data || []);
    }
    setLoading({ workOrders: false, assets: false, vendors: false });
  }, []);

  useEffect(() => { if (session) load(); }, [session, load]);

  const signOut = async () => {
    await supabase.auth.signOut();
    setSession(null);
  };

  if (authLoading) return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", color: C.muted, fontFamily: "monospace" }}>
      Loading...
    </div>
  );

  if (!session) return <LoginScreen />;

  const tabs = ["Overview", "Work Orders", "Assets", "Vendors"];

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'DM Sans', 'Helvetica Neue', sans-serif", color: C.text }}>
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "0 16px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", minHeight: 58, flexWrap: "wrap", gap: 8, padding: "8px 0" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ background: C.accent, borderRadius: 8, width: 34, height: 34, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>🏭</div>
            <div>
              <div style={{ fontFamily: "monospace", fontSize: 16, letterSpacing: 2, color: C.text, fontWeight: 800 }}>FACILITY COMMAND</div>
              <div style={{ fontSize: 9, color: C.muted, letterSpacing: "0.08em" }}>
                INDUSTRIAL WAREHOUSE MANAGEMENT
                {isAdmin && <span style={{ marginLeft: 8, color: C.accent }}>★ ADMIN</span>}
              </div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            <button onClick={load} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 10px", color: C.muted, cursor: "pointer", fontSize: 12 }}>Refresh</button>
            <div style={{ fontSize: 11, color: C.muted, maxWidth: 150, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{session.user.email}</div>
            <button onClick={signOut} style={{ background: "transparent", border: `1px solid ${C.border}`, borderRadius: 6, padding: "5px 10px", color: C.muted, cursor: "pointer", fontSize: 12 }}>Sign Out</button>
          </div>
        </div>
        <div style={{ display: "flex", overflowX: "auto", WebkitOverflowScrolling: "touch" }}>
          {tabs.map(t => (
            <button key={t} onClick={() => setTab(t)} style={{ background: "transparent", border: "none", padding: "10px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap", color: tab === t ? C.accent : C.muted, borderBottom: `2px solid ${tab === t ? C.accent : "transparent"}` }}>{t}</button>
          ))}
        </div>
      </div>
      <div style={{ padding: "20px 16px", maxWidth: 1280, margin: "0 auto" }}>
        <ErrorBanner msg={globalError} onDismiss={() => setGlobalError(null)} />
        {tab === "Overview" && <Overview workOrders={workOrders} assets={assets} vendors={vendors} />}
        {tab === "Work Orders" && <WorkOrders workOrders={workOrders} setWorkOrders={setWorkOrders} loading={loading.workOrders} onAdd={r => setWorkOrders(p => [r, ...p])} isAdmin={isAdmin} vendors={vendors} />}
        {tab === "Assets" && <Assets assets={assets} setAssets={setAssets} loading={loading.assets} onAdd={r => setAssets(p => [r, ...p])} isAdmin={isAdmin} />}
        {tab === "Vendors" && <Vendors vendors={vendors} setVendors={setVendors} loading={loading.vendors} onAdd={r => setVendors(p => [r, ...p])} isAdmin={isAdmin} />}
      </div>
    </div>
  );
}

