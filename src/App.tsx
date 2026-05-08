import { useState, useEffect, useCallback } from 'react';

// ─── Theme ───────────────────────────────────────────────────────────────────
const C = {
  bg: '#0d0f12',
  surface: '#141720',
  card: '#1a1e2a',
  border: '#252b3b',
  accent: '#f97316',
  accentDim: '#7c3b12',
  yellow: '#eab308',
  green: '#22c55e',
  red: '#ef4444',
  blue: '#3b82f6',
  text: '#e2e8f0',
  muted: '#64748b',
  subtle: '#94a3b8',
};

// ─── Helpers ─────────────────────────────────────────────────────────────────
const TODAY = new Date().toISOString().split('T')[0];

const uid = (prefix) => `${prefix}-${Date.now().toString(36).toUpperCase()}`;

const priorityColor = (p) =>
  ({ Critical: C.red, High: C.accent, Medium: C.yellow, Low: C.green }[p] ||
  C.muted);

const statusColor = (s) =>
  ({
    Open: C.accent,
    'In Progress': C.blue,
    Completed: C.green,
    Pending: C.yellow,
    Operational: C.green,
    'Under Maintenance': C.accent,
    Degraded: C.red,
    Active: C.green,
    Inactive: C.muted,
  }[s] || C.muted);

// ─── Shared UI ───────────────────────────────────────────────────────────────
const Badge = ({ label, color }) => (
  <span
    style={{
      background: color + '22',
      color,
      border: `1px solid ${color}44`,
      borderRadius: 4,
      padding: '2px 8px',
      fontSize: 11,
      fontWeight: 700,
      letterSpacing: '0.05em',
      textTransform: 'uppercase',
      whiteSpace: 'nowrap',
    }}
  >
    {label}
  </span>
);

const Input = ({ label, value, onChange, type = 'text', style = {} }) => (
  <div style={style}>
    <div
      style={{
        fontSize: 11,
        color: C.muted,
        marginBottom: 4,
        textTransform: 'uppercase',
        letterSpacing: '0.06em',
      }}
    >
      {label}
    </div>
    <input
      type={type}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{
        width: '100%',
        background: C.surface,
        border: `1px solid ${C.border}`,
        borderRadius: 6,
        padding: '7px 10px',
        color: C.text,
        fontSize: 13,
        boxSizing: 'border-box',
      }}
    />
  </div>
);

const Select = ({ label, value, onChange, options }) => (
  <div>
    <div
      style={{
        fontSize: 11,
        color: C.muted,
        marginBottom: 4,
        textTransform: 'uppercase',
        letterSpacing: '0.06em',
      }}
    >
      {label}
    </div>
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{
        width: '100%',
        background: C.surface,
        border: `1px solid ${C.border}`,
        borderRadius: 6,
        padding: '7px 10px',
        color: C.text,
        fontSize: 13,
      }}
    >
      {options.map((o) => (
        <option key={o}>{o}</option>
      ))}
    </select>
  </div>
);

const Btn = ({ children, onClick, variant, disabled }) => {
  const isPrimary = variant !== 'secondary';
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      style={{
        background: isPrimary ? C.accent : 'transparent',
        color: isPrimary ? '#fff' : C.muted,
        border: isPrimary ? 'none' : `1px solid ${C.border}`,
        borderRadius: 6,
        padding: '7px 18px',
        fontSize: 13,
        fontWeight: 700,
        cursor: disabled ? 'not-allowed' : 'pointer',
        opacity: disabled ? 0.5 : 1,
      }}
    >
      {children}
    </button>
  );
};

const StatCard = ({ icon, label, value, sub, color }) => (
  <div
    style={{
      background: C.card,
      border: `1px solid ${C.border}`,
      borderRadius: 10,
      padding: '18px 22px',
      flex: 1,
      minWidth: 140,
      borderLeft: `3px solid ${color || C.accent}`,
    }}
  >
    <div style={{ fontSize: 22, marginBottom: 6 }}>{icon}</div>
    <div
      style={{
        fontSize: 28,
        fontWeight: 800,
        color: C.text,
        fontFamily: 'monospace',
        letterSpacing: 1,
      }}
    >
      {value}
    </div>
    <div style={{ fontSize: 13, color: C.subtle, marginTop: 2 }}>{label}</div>
    {sub && (
      <div style={{ fontSize: 11, color: color || C.accent, marginTop: 4 }}>
        {sub}
      </div>
    )}
  </div>
);

const Spinner = () => (
  <div
    style={{ textAlign: 'center', padding: 48, color: C.muted, fontSize: 13 }}
  >
    ⏳ Loading data from Supabase...
  </div>
);

const ErrorBanner = ({ msg, onDismiss }) =>
  msg ? (
    <div
      style={{
        background: C.red + '22',
        border: `1px solid ${C.red}44`,
        borderRadius: 8,
        padding: '10px 16px',
        marginBottom: 16,
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        fontSize: 13,
        color: C.red,
      }}
    >
      ⚠️ {msg}
      <button
        onClick={onDismiss}
        style={{
          background: 'none',
          border: 'none',
          color: C.red,
          cursor: 'pointer',
          fontSize: 16,
        }}
      >
        ✕
      </button>
    </div>
  ) : null;

// ─── WORK ORDERS ─────────────────────────────────────────────────────────────
function WorkOrders({ workOrders, loading, onAdd }) {
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [filter, setFilter] = useState('All');
  const [form, setForm] = useState({
    title: '',
    asset: '',
    priority: 'Medium',
    due: '',
    vendor: '',
  });

  const f = (k) => (v) => setForm((p) => ({ ...p, [k]: v }));

  const filtered =
    filter === 'All'
      ? workOrders
      : workOrders.filter((w) => w.status === filter);

  const submit = async () => {
    if (!form.title || !form.asset) {
      setError('Title and Asset are required.');
      return;
    }
    setSaving(true);
    setError(null);
    const record = {
      id: uid('WO'),
      title: form.title,
      asset: form.asset,
      priority: form.priority,
      status: 'Open',
      assignee: null,
      due: form.due || null,
      vendor: form.vendor || null,
    };
    const { error: err } = await supabase.from('work_orders').insert([record]);
    if (err) {
      setError(err.message);
    } else {
      onAdd(record);
      setForm({
        title: '',
        asset: '',
        priority: 'Medium',
        due: '',
        vendor: '',
      });
      setShowForm(false);
    }
    setSaving(false);
  };

  return (
    <div>
      <ErrorBanner msg={error} onDismiss={() => setError(null)} />
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          marginBottom: 18,
        }}
      >
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          {['All', 'Open', 'In Progress', 'Pending', 'Completed'].map((s) => (
            <button
              key={s}
              onClick={() => setFilter(s)}
              style={{
                background: filter === s ? C.accent : C.card,
                color: filter === s ? '#fff' : C.muted,
                border: `1px solid ${filter === s ? C.accent : C.border}`,
                borderRadius: 6,
                padding: '5px 14px',
                fontSize: 12,
                cursor: 'pointer',
                fontWeight: 600,
              }}
            >
              {s}
            </button>
          ))}
        </div>
        <Btn onClick={() => setShowForm((v) => !v)}>+ New Work Order</Btn>
      </div>

      {showForm && (
        <div
          style={{
            background: C.card,
            border: `1px solid ${C.accent}44`,
            borderRadius: 10,
            padding: 20,
            marginBottom: 18,
          }}
        >
          <div
            style={{
              color: C.accent,
              fontWeight: 700,
              marginBottom: 14,
              fontSize: 13,
              letterSpacing: '0.08em',
              textTransform: 'uppercase',
            }}
          >
            New Work Order
          </div>
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(3, 1fr)',
              gap: 12,
            }}
          >
            <Input label="Title *" value={form.title} onChange={f('title')} />
            <Input
              label="Asset / Location *"
              value={form.asset}
              onChange={f('asset')}
            />
            <Input
              label="Due Date"
              value={form.due}
              onChange={f('due')}
              type="date"
            />
            <Select
              label="Priority"
              value={form.priority}
              onChange={f('priority')}
              options={['Critical', 'High', 'Medium', 'Low']}
            />
            <Input
              label="Vendor (optional)"
              value={form.vendor}
              onChange={f('vendor')}
            />
          </div>
          <div style={{ display: 'flex', gap: 8, marginTop: 14 }}>
            <Btn onClick={submit} disabled={saving}>
              {saving ? 'Saving…' : 'Create'}
            </Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>
              Cancel
            </Btn>
          </div>
        </div>
      )}

      {loading ? (
        <Spinner />
      ) : (
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead>
              <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                {[
                  'ID',
                  'Title',
                  'Asset',
                  'Priority',
                  'Status',
                  'Assignee',
                  'Vendor',
                  'Due',
                ].map((h) => (
                  <th
                    key={h}
                    style={{
                      textAlign: 'left',
                      padding: '8px 12px',
                      fontSize: 11,
                      color: C.muted,
                      textTransform: 'uppercase',
                      letterSpacing: '0.07em',
                      fontWeight: 600,
                    }}
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((wo, i) => (
                <tr
                  key={wo.id}
                  style={{
                    borderBottom: `1px solid ${C.border}22`,
                    background: i % 2 === 0 ? 'transparent' : C.surface + '44',
                  }}
                >
                  <td
                    style={{
                      padding: '10px 12px',
                      fontSize: 12,
                      color: C.muted,
                      fontFamily: 'monospace',
                    }}
                  >
                    {wo.id}
                  </td>
                  <td
                    style={{
                      padding: '10px 12px',
                      fontSize: 13,
                      color: C.text,
                      fontWeight: 600,
                    }}
                  >
                    {wo.title}
                  </td>
                  <td
                    style={{
                      padding: '10px 12px',
                      fontSize: 12,
                      color: C.subtle,
                    }}
                  >
                    {wo.asset}
                  </td>
                  <td style={{ padding: '10px 12px' }}>
                    <Badge
                      label={wo.priority}
                      color={priorityColor(wo.priority)}
                    />
                  </td>
                  <td style={{ padding: '10px 12px' }}>
                    <Badge label={wo.status} color={statusColor(wo.status)} />
                  </td>
                  <td
                    style={{
                      padding: '10px 12px',
                      fontSize: 12,
                      color: C.subtle,
                    }}
                  >
                    {wo.assignee || '—'}
                  </td>
                  <td
                    style={{
                      padding: '10px 12px',
                      fontSize: 12,
                      color: C.subtle,
                    }}
                  >
                    {wo.vendor || '—'}
                  </td>
                  <td
                    style={{
                      padding: '10px 12px',
                      fontSize: 12,
                      color:
                        wo.due && wo.due <= TODAY && wo.status !== 'Completed'
                          ? C.red
                          : C.subtle,
                    }}
                  >
                    {wo.due || '—'}
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr>
                  <td
                    colSpan={8}
                    style={{
                      padding: 32,
                      textAlign: 'center',
                      color: C.muted,
                      fontSize: 13,
                    }}
                  >
                    No work orders found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── ASSETS ──────────────────────────────────────────────────────────────────
function Assets({ assets, loading, onAdd }) {
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [form, setForm] = useState({
    name: '',
    category: '',
    location: '',
    value: '',
    next_service: '',
  });

  const f = (k) => (v) => setForm((p) => ({ ...p, [k]: v }));

  const submit = async () => {
    if (!form.name) {
      setError('Asset name is required.');
      return;
    }
    setSaving(true);
    setError(null);
    const record = {
      id: uid('AST'),
      name: form.name,
      category: form.category,
      location: form.location,
      value: form.value,
      status: 'Operational',
      last_service: TODAY,
      next_service: form.next_service || null,
    };
    const { error: err } = await supabase.from('assets').insert([record]);
    if (err) {
      setError(err.message);
    } else {
      onAdd(record);
      setForm({
        name: '',
        category: '',
        location: '',
        value: '',
        next_service: '',
      });
      setShowForm(false);
    }
    setSaving(false);
  };

  return (
    <div>
      <ErrorBanner msg={error} onDismiss={() => setError(null)} />
      <div
        style={{
          display: 'flex',
          justifyContent: 'flex-end',
          marginBottom: 18,
        }}
      >
        <Btn onClick={() => setShowForm((v) => !v)}>+ Add Asset</Btn>
      </div>

      {showForm && (
        <div
          style={{
            background: C.card,
            border: `1px solid ${C.accent}44`,
            borderRadius: 10,
            padding: 20,
            marginBottom: 18,
          }}
        >
          <div
            style={{
              color: C.accent,
              fontWeight: 700,
              marginBottom: 14,
              fontSize: 13,
              letterSpacing: '0.08em',
              textTransform: 'uppercase',
            }}
          >
            Register New Asset
          </div>
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(3, 1fr)',
              gap: 12,
            }}
          >
            <Input
              label="Asset Name *"
              value={form.name}
              onChange={f('name')}
            />
            <Input
              label="Category"
              value={form.category}
              onChange={f('category')}
            />
            <Input
              label="Location / Zone"
              value={form.location}
              onChange={f('location')}
            />
            <Input
              label="Est. Value"
              value={form.value}
              onChange={f('value')}
            />
            <Input
              label="Next Service Date"
              value={form.next_service}
              onChange={f('next_service')}
              type="date"
            />
          </div>
          <div style={{ display: 'flex', gap: 8, marginTop: 14 }}>
            <Btn onClick={submit} disabled={saving}>
              {saving ? 'Saving…' : 'Register'}
            </Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>
              Cancel
            </Btn>
          </div>
        </div>
      )}

      {loading ? (
        <Spinner />
      ) : (
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))',
            gap: 14,
          }}
        >
          {assets.map((a) => (
            <div
              key={a.id}
              style={{
                background: C.card,
                border: `1px solid ${C.border}`,
                borderRadius: 10,
                padding: 18,
                borderTop: `3px solid ${statusColor(a.status)}`,
              }}
            >
              <div
                style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'flex-start',
                  marginBottom: 10,
                }}
              >
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>
                    {a.name}
                  </div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>
                    {a.id} · {a.category}
                  </div>
                </div>
                <Badge label={a.status} color={statusColor(a.status)} />
              </div>
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr',
                  gap: 8,
                  fontSize: 12,
                }}
              >
                {[
                  ['📍 Location', a.location],
                  ['💰 Value', a.value],
                  ['🔧 Last Service', a.last_service],
                  ['📅 Next Service', a.next_service],
                ].map(([lbl, val]) => (
                  <div key={lbl}>
                    <div
                      style={{
                        color: C.muted,
                        fontSize: 10,
                        textTransform: 'uppercase',
                        letterSpacing: '0.06em',
                      }}
                    >
                      {lbl}
                    </div>
                    <div style={{ color: C.subtle, marginTop: 2 }}>
                      {val || '—'}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ))}
          {assets.length === 0 && (
            <div style={{ color: C.muted, fontSize: 13, padding: 32 }}>
              No assets registered yet.
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── VENDORS ─────────────────────────────────────────────────────────────────
function Vendors({ vendors, loading, onAdd }) {
  const [showForm, setShowForm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);
  const [form, setForm] = useState({
    name: '',
    specialty: '',
    contact: '',
    phone: '',
    email: '',
  });

  const f = (k) => (v) => setForm((p) => ({ ...p, [k]: v }));

  const submit = async () => {
    if (!form.name) {
      setError('Company name is required.');
      return;
    }
    setSaving(true);
    setError(null);
    const record = {
      id: uid('VND'),
      ...form,
      status: 'Active',
      rating: 0,
      open_orders: 0,
    };
    const { error: err } = await supabase.from('vendors').insert([record]);
    if (err) {
      setError(err.message);
    } else {
      onAdd(record);
      setForm({ name: '', specialty: '', contact: '', phone: '', email: '' });
      setShowForm(false);
    }
    setSaving(false);
  };

  const Stars = ({ rating }) => (
    <div style={{ display: 'flex', alignItems: 'center', gap: 2 }}>
      {[1, 2, 3, 4, 5].map((i) => (
        <span
          key={i}
          style={{
            color: i <= Math.floor(rating) ? C.yellow : C.border,
            fontSize: 13,
          }}
        >
          ★
        </span>
      ))}
      <span style={{ fontSize: 11, color: C.muted, marginLeft: 4 }}>
        {rating > 0 ? Number(rating).toFixed(1) : 'N/A'}
      </span>
    </div>
  );

  return (
    <div>
      <ErrorBanner msg={error} onDismiss={() => setError(null)} />
      <div
        style={{
          display: 'flex',
          justifyContent: 'flex-end',
          marginBottom: 18,
        }}
      >
        <Btn onClick={() => setShowForm((v) => !v)}>+ Add Vendor</Btn>
      </div>

      {showForm && (
        <div
          style={{
            background: C.card,
            border: `1px solid ${C.accent}44`,
            borderRadius: 10,
            padding: 20,
            marginBottom: 18,
          }}
        >
          <div
            style={{
              color: C.accent,
              fontWeight: 700,
              marginBottom: 14,
              fontSize: 13,
              letterSpacing: '0.08em',
              textTransform: 'uppercase',
            }}
          >
            Register Vendor / Contractor
          </div>
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(3, 1fr)',
              gap: 12,
            }}
          >
            <Input
              label="Company Name *"
              value={form.name}
              onChange={f('name')}
            />
            <Input
              label="Specialty"
              value={form.specialty}
              onChange={f('specialty')}
            />
            <Input
              label="Contact Person"
              value={form.contact}
              onChange={f('contact')}
            />
            <Input label="Phone" value={form.phone} onChange={f('phone')} />
            <Input label="Email" value={form.email} onChange={f('email')} />
          </div>
          <div style={{ display: 'flex', gap: 8, marginTop: 14 }}>
            <Btn onClick={submit} disabled={saving}>
              {saving ? 'Saving…' : 'Register'}
            </Btn>
            <Btn variant="secondary" onClick={() => setShowForm(false)}>
              Cancel
            </Btn>
          </div>
        </div>
      )}

      {loading ? (
        <Spinner />
      ) : (
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))',
            gap: 14,
          }}
        >
          {vendors.map((v) => (
            <div
              key={v.id}
              style={{
                background: C.card,
                border: `1px solid ${C.border}`,
                borderRadius: 10,
                padding: 20,
              }}
            >
              <div
                style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'flex-start',
                  marginBottom: 12,
                }}
              >
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>
                    {v.name}
                  </div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>
                    {v.specialty}
                  </div>
                </div>
                <Badge label={v.status} color={statusColor(v.status)} />
              </div>
              <Stars rating={v.rating} />
              <div
                style={{
                  marginTop: 12,
                  display: 'grid',
                  gridTemplateColumns: '1fr 1fr',
                  gap: 8,
                  fontSize: 12,
                }}
              >
                {[
                  ['👤 Contact', v.contact],
                  ['📞 Phone', v.phone],
                  ['✉️ Email', v.email],
                ].map(([lbl, val]) => (
                  <div key={lbl}>
                    <div
                      style={{
                        color: C.muted,
                        fontSize: 10,
                        textTransform: 'uppercase',
                      }}
                    >
                      {lbl}
                    </div>
                    <div style={{ color: C.subtle, marginTop: 2 }}>
                      {val || '—'}
                    </div>
                  </div>
                ))}
                <div>
                  <div
                    style={{
                      color: C.muted,
                      fontSize: 10,
                      textTransform: 'uppercase',
                    }}
                  >
                    📋 Open Orders
                  </div>
                  <div
                    style={{
                      color: v.open_orders > 0 ? C.accent : C.subtle,
                      marginTop: 2,
                      fontWeight: v.open_orders > 0 ? 700 : 400,
                    }}
                  >
                    {v.open_orders}
                  </div>
                </div>
              </div>
            </div>
          ))}
          {vendors.length === 0 && (
            <div style={{ color: C.muted, fontSize: 13, padding: 32 }}>
              No vendors registered yet.
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── OVERVIEW ────────────────────────────────────────────────────────────────
function Overview({ workOrders, assets, vendors }) {
  const open = workOrders.filter((w) => w.status !== 'Completed').length;
  const critical = workOrders.filter((w) => w.priority === 'Critical').length;
  const opAssets = assets.filter((a) => a.status === 'Operational').length;
  const activeVendors = vendors.filter((v) => v.status === 'Active').length;
  const overdue = workOrders.filter(
    (w) => w.due && w.due <= TODAY && w.status !== 'Completed'
  ).length;

  return (
    <div>
      <div
        style={{ display: 'flex', gap: 14, flexWrap: 'wrap', marginBottom: 24 }}
      >
        <StatCard
          icon="🔧"
          label="Open Work Orders"
          value={open}
          sub={`${critical} critical`}
          color={C.accent}
        />
        <StatCard
          icon="🏭"
          label="Operational Assets"
          value={`${opAssets}/${assets.length}`}
          sub="fleet status"
          color={C.green}
        />
        <StatCard
          icon="🤝"
          label="Active Vendors"
          value={activeVendors}
          sub="contractors on file"
          color={C.blue}
        />
        <StatCard
          icon="⚠️"
          label="Overdue / At Risk"
          value={overdue}
          sub="past due date"
          color={C.red}
        />
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <div
          style={{
            background: C.card,
            border: `1px solid ${C.border}`,
            borderRadius: 10,
            padding: 20,
          }}
        >
          <div
            style={{
              fontSize: 12,
              color: C.muted,
              textTransform: 'uppercase',
              letterSpacing: '0.07em',
              fontWeight: 600,
              marginBottom: 14,
            }}
          >
            Recent Work Orders
          </div>
          {workOrders.slice(0, 5).map((wo) => (
            <div
              key={wo.id}
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                padding: '8px 0',
                borderBottom: `1px solid ${C.border}22`,
              }}
            >
              <div>
                <div style={{ fontSize: 13, color: C.text, fontWeight: 600 }}>
                  {wo.title}
                </div>
                <div style={{ fontSize: 11, color: C.muted }}>{wo.asset}</div>
              </div>
              <div style={{ display: 'flex', gap: 6 }}>
                <Badge label={wo.priority} color={priorityColor(wo.priority)} />
                <Badge label={wo.status} color={statusColor(wo.status)} />
              </div>
            </div>
          ))}
          {workOrders.length === 0 && (
            <div style={{ color: C.muted, fontSize: 13 }}>
              No work orders yet.
            </div>
          )}
        </div>

        <div
          style={{
            background: C.card,
            border: `1px solid ${C.border}`,
            borderRadius: 10,
            padding: 20,
          }}
        >
          <div
            style={{
              fontSize: 12,
              color: C.muted,
              textTransform: 'uppercase',
              letterSpacing: '0.07em',
              fontWeight: 600,
              marginBottom: 14,
            }}
          >
            Asset Status Breakdown
          </div>
          {[
            ['Operational', C.green],
            ['Under Maintenance', C.accent],
            ['Degraded', C.red],
          ].map(([status, color]) => {
            const count = assets.filter((a) => a.status === status).length;
            const pct =
              assets.length > 0 ? Math.round((count / assets.length) * 100) : 0;
            return (
              <div key={status} style={{ marginBottom: 12 }}>
                <div
                  style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    marginBottom: 4,
                    fontSize: 12,
                  }}
                >
                  <span style={{ color: C.subtle }}>{status}</span>
                  <span style={{ color, fontWeight: 700 }}>
                    {count} ({pct}%)
                  </span>
                </div>
                <div
                  style={{ background: C.border, borderRadius: 4, height: 6 }}
                >
                  <div
                    style={{
                      background: color,
                      width: `${pct}%`,
                      height: 6,
                      borderRadius: 4,
                    }}
                  />
                </div>
              </div>
            );
          })}

          <div
            style={{
              marginTop: 20,
              fontSize: 12,
              color: C.muted,
              textTransform: 'uppercase',
              letterSpacing: '0.07em',
              fontWeight: 600,
              marginBottom: 10,
            }}
          >
            Top Vendors by Rating
          </div>
          {[...vendors]
            .filter((v) => v.rating > 0)
            .sort((a, b) => b.rating - a.rating)
            .slice(0, 3)
            .map((v) => (
              <div
                key={v.id}
                style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  marginBottom: 8,
                }}
              >
                <span style={{ fontSize: 13, color: C.subtle }}>{v.name}</span>
                <span
                  style={{ fontSize: 13, color: C.yellow, fontWeight: 700 }}
                >
                  ★ {Number(v.rating).toFixed(1)}
                </span>
              </div>
            ))}
          {vendors.filter((v) => v.rating > 0).length === 0 && (
            <div style={{ color: C.muted, fontSize: 13 }}>
              No rated vendors yet.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── ROOT APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState('Overview');
  const [workOrders, setWorkOrders] = useState([]);
  const [assets, setAssets] = useState([]);
  const [vendors, setVendors] = useState([]);
  const [loading, setLoading] = useState({
    workOrders: true,
    assets: true,
    vendors: true,
  });
  const [globalError, setGlobalError] = useState(null);

  // ── Load all data on mount ──
  const load = useCallback(async () => {
    setLoading({ workOrders: true, assets: true, vendors: true });

    const [woRes, astRes, vndRes] = await Promise.all([
      supabase
        .from('work_orders')
        .select('*')
        .order('due', { ascending: true }),
      supabase.from('assets').select('*').order('name', { ascending: true }),
      supabase.from('vendors').select('*').order('name', { ascending: true }),
    ]);

    if (woRes.error || astRes.error || vndRes.error) {
      setGlobalError(
        'Failed to load data from Supabase. Check your environment variables and table setup.'
      );
    } else {
      setWorkOrders(woRes.data || []);
      setAssets(astRes.data || []);
      setVendors(vndRes.data || []);
    }
    setLoading({ workOrders: false, assets: false, vendors: false });
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  const tabs = ['Overview', 'Work Orders', 'Assets', 'Vendors'];

  return (
    <div
      style={{
        minHeight: '100vh',
        background: C.bg,
        fontFamily: "'DM Sans', 'Helvetica Neue', sans-serif",
        color: C.text,
      }}
    >
      {/* Header */}
      <div
        style={{
          background: C.surface,
          borderBottom: `1px solid ${C.border}`,
          padding: '0 28px',
        }}
      >
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            height: 58,
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
            <div
              style={{
                background: C.accent,
                borderRadius: 8,
                width: 34,
                height: 34,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: 18,
              }}
            >
              🏭
            </div>
            <div>
              <div
                style={{
                  fontFamily: 'monospace',
                  fontSize: 18,
                  letterSpacing: 3,
                  color: C.text,
                  fontWeight: 800,
                }}
              >
                FACILITY COMMAND
              </div>
              <div
                style={{
                  fontSize: 10,
                  color: C.muted,
                  letterSpacing: '0.1em',
                  marginTop: -2,
                }}
              >
                INDUSTRIAL WAREHOUSE MANAGEMENT
              </div>
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <button
              onClick={load}
              title="Refresh data"
              style={{
                background: C.card,
                border: `1px solid ${C.border}`,
                borderRadius: 6,
                padding: '5px 12px',
                color: C.muted,
                cursor: 'pointer',
                fontSize: 13,
              }}
            >
              ↻ Refresh
            </button>
            <div style={{ fontSize: 12, color: C.muted }}>
              {new Date().toDateString()}
            </div>
            <div
              style={{
                width: 32,
                height: 32,
                borderRadius: '50%',
                background: C.accentDim,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: 14,
                fontWeight: 700,
                color: C.accent,
              }}
            >
              FM
            </div>
          </div>
        </div>
        <div style={{ display: 'flex' }}>
          {tabs.map((t) => (
            <button
              key={t}
              onClick={() => setTab(t)}
              style={{
                background: 'transparent',
                border: 'none',
                padding: '10px 20px',
                fontSize: 13,
                fontWeight: 600,
                cursor: 'pointer',
                color: tab === t ? C.accent : C.muted,
                borderBottom: `2px solid ${
                  tab === t ? C.accent : 'transparent'
                }`,
              }}
            >
              {t}
            </button>
          ))}
        </div>
      </div>

      {/* Content */}
      <div style={{ padding: 28, maxWidth: 1280, margin: '0 auto' }}>
        <ErrorBanner msg={globalError} onDismiss={() => setGlobalError(null)} />
        {tab === 'Overview' && (
          <Overview workOrders={workOrders} assets={assets} vendors={vendors} />
        )}
        {tab === 'Work Orders' && (
          <WorkOrders
            workOrders={workOrders}
            loading={loading.workOrders}
            onAdd={(r) => setWorkOrders((p) => [r, ...p])}
          />
        )}
        {tab === 'Assets' && (
          <Assets
            assets={assets}
            loading={loading.assets}
            onAdd={(r) => setAssets((p) => [r, ...p])}
          />
        )}
        {tab === 'Vendors' && (
          <Vendors
            vendors={vendors}
            loading={loading.vendors}
            onAdd={(r) => setVendors((p) => [r, ...p])}
          />
        )}
      </div>
    </div>
  );
}
