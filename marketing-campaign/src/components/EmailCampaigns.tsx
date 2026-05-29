import { useState, useEffect, useCallback } from 'react';
import type { Contact, Campaign } from '../types';
import { EMAIL_TEMPLATES } from '../data/emailTemplates';

const STORAGE_CONTACTS = 'fiorentini_contacts';
const STORAGE_CAMPAIGNS = 'fiorentini_campaigns';

function uid() {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function loadFromStorage<T>(key: string, fallback: T): T {
  try {
    const raw = localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
  }
}

type SubTab = 'contacts' | 'campaigns';

// ─── Contact form state ───
const EMPTY_CONTACT = {
  firstName: '', lastName: '', email: '', company: '',
  jobTitle: '', phone: '', tags: '',
};

// ─── Campaign form state ───
const EMPTY_CAMPAIGN = {
  name: '', templateId: '', contactIds: [] as string[],
  senderName: '', senderEmail: '', customSubject: '',
};

export default function EmailCampaigns() {
  const [subTab, setSubTab] = useState<SubTab>('contacts');
  const [contacts, setContacts] = useState<Contact[]>(() =>
    loadFromStorage<Contact[]>(STORAGE_CONTACTS, [])
  );
  const [campaigns, setCampaigns] = useState<Campaign[]>(() =>
    loadFromStorage<Campaign[]>(STORAGE_CAMPAIGNS, [])
  );

  // Contact modal
  const [showContactModal, setShowContactModal] = useState(false);
  const [editingContact, setEditingContact] = useState<Contact | null>(null);
  const [contactForm, setContactForm] = useState(EMPTY_CONTACT);
  const [contactSearch, setContactSearch] = useState('');

  // Campaign modal
  const [showCampaignModal, setShowCampaignModal] = useState(false);
  const [campaignForm, setCampaignForm] = useState(EMPTY_CAMPAIGN);
  const [previewOpen, setPreviewOpen] = useState(false);
  const [campaignSearch, setCampaignSearch] = useState('');

  // Delete confirm
  const [deleteConfirm, setDeleteConfirm] = useState<{ type: 'contact' | 'campaign'; id: string } | null>(null);

  // Toast
  const [toast, setToast] = useState<{ msg: string; kind: 'success' | 'info' } | null>(null);

  const showToast = useCallback((msg: string, kind: 'success' | 'info' = 'success') => {
    setToast({ msg, kind });
    setTimeout(() => setToast(null), 3000);
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_CONTACTS, JSON.stringify(contacts));
  }, [contacts]);

  useEffect(() => {
    localStorage.setItem(STORAGE_CAMPAIGNS, JSON.stringify(campaigns));
  }, [campaigns]);

  // ── Contacts ──
  const openAddContact = () => {
    setEditingContact(null);
    setContactForm(EMPTY_CONTACT);
    setShowContactModal(true);
  };

  const openEditContact = (c: Contact) => {
    setEditingContact(c);
    setContactForm({
      firstName: c.firstName, lastName: c.lastName,
      email: c.email, company: c.company,
      jobTitle: c.jobTitle, phone: c.phone,
      tags: c.tags.join(', '),
    });
    setShowContactModal(true);
  };

  const saveContact = () => {
    if (!contactForm.firstName.trim() || !contactForm.email.trim()) return;
    const tags = contactForm.tags.split(',').map(t => t.trim()).filter(Boolean);
    if (editingContact) {
      setContacts(prev => prev.map(c =>
        c.id === editingContact.id
          ? { ...c, ...contactForm, tags, phone: contactForm.phone }
          : c
      ));
      showToast('Contact updated');
    } else {
      const newContact: Contact = {
        id: uid(),
        firstName: contactForm.firstName.trim(),
        lastName: contactForm.lastName.trim(),
        email: contactForm.email.trim(),
        company: contactForm.company.trim(),
        jobTitle: contactForm.jobTitle.trim(),
        phone: contactForm.phone.trim(),
        tags,
        addedAt: new Date().toISOString(),
      };
      setContacts(prev => [newContact, ...prev]);
      showToast('Contact added');
    }
    setShowContactModal(false);
  };

  const deleteItem = () => {
    if (!deleteConfirm) return;
    if (deleteConfirm.type === 'contact') {
      setContacts(prev => prev.filter(c => c.id !== deleteConfirm.id));
      showToast('Contact deleted', 'info');
    } else {
      setCampaigns(prev => prev.filter(c => c.id !== deleteConfirm.id));
      showToast('Campaign deleted', 'info');
    }
    setDeleteConfirm(null);
  };

  const filteredContacts = contacts.filter(c => {
    const q = contactSearch.toLowerCase();
    return (
      `${c.firstName} ${c.lastName}`.toLowerCase().includes(q) ||
      c.email.toLowerCase().includes(q) ||
      c.company.toLowerCase().includes(q)
    );
  });

  // ── Campaigns ──
  const openAddCampaign = () => {
    setCampaignForm(EMPTY_CAMPAIGN);
    setPreviewOpen(false);
    setShowCampaignModal(true);
  };

  const selectedTemplate = EMAIL_TEMPLATES.find(t => t.id === campaignForm.templateId);

  const getPreviewBody = () => {
    if (!selectedTemplate) return '';
    const firstName = campaignForm.contactIds.length > 0
      ? (contacts.find(c => c.id === campaignForm.contactIds[0])?.firstName ?? 'there')
      : 'there';
    return selectedTemplate.body
      .replace(/{{firstName}}/g, firstName)
      .replace(/{{senderName}}/g, campaignForm.senderName || '[Sender Name]');
  };

  const saveCampaign = (status: 'draft' | 'sent') => {
    if (!campaignForm.name.trim() || !campaignForm.templateId) return;
    const newCampaign: Campaign = {
      id: uid(),
      name: campaignForm.name.trim(),
      templateId: campaignForm.templateId,
      contactIds: campaignForm.contactIds,
      status,
      senderName: campaignForm.senderName.trim(),
      senderEmail: campaignForm.senderEmail.trim(),
      customSubject: campaignForm.customSubject.trim() || (selectedTemplate?.subject ?? ''),
      createdAt: new Date().toISOString(),
      sentAt: status === 'sent' ? new Date().toISOString() : undefined,
    };
    setCampaigns(prev => [newCampaign, ...prev]);
    setShowCampaignModal(false);
    showToast(status === 'sent' ? 'Campaign marked as sent' : 'Campaign saved as draft');
  };

  const toggleCampaignContact = (id: string) => {
    setCampaignForm(prev => ({
      ...prev,
      contactIds: prev.contactIds.includes(id)
        ? prev.contactIds.filter(x => x !== id)
        : [...prev.contactIds, id],
    }));
  };

  const filteredCampaigns = campaigns.filter(c =>
    c.name.toLowerCase().includes(campaignSearch.toLowerCase())
  );

  const templateName = (id: string) =>
    EMAIL_TEMPLATES.find(t => t.id === id)?.name ?? '—';

  return (
    <div>
      {/* Toast */}
      {toast && (
        <div style={{
          position: 'fixed', bottom: 24, right: 24, zIndex: 999,
          background: toast.kind === 'success' ? 'var(--success)' : 'var(--text-2)',
          color: '#fff', padding: '10px 18px', borderRadius: 8,
          boxShadow: 'var(--shadow-lg)', fontWeight: 500, fontSize: 14,
          animation: 'fadeIn 0.2s ease',
        }}>{toast.msg}</div>
      )}

      {/* Sub-tab nav */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 24 }}>
        <div style={{ display: 'flex', gap: 2, background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, padding: 3 }}>
          {(['contacts', 'campaigns'] as SubTab[]).map(tab => (
            <button
              key={tab}
              onClick={() => setSubTab(tab)}
              style={{
                padding: '7px 18px',
                borderRadius: 6,
                fontWeight: 500,
                fontSize: 14,
                transition: 'all 0.15s',
                background: subTab === tab ? 'var(--primary)' : 'transparent',
                color: subTab === tab ? '#fff' : 'var(--text-2)',
              }}
            >
              {tab.charAt(0).toUpperCase() + tab.slice(1)}
              {tab === 'contacts' && contacts.length > 0 && (
                <span style={{
                  marginLeft: 6,
                  background: subTab === tab ? 'rgba(255,255,255,0.25)' : 'var(--border)',
                  color: subTab === tab ? '#fff' : 'var(--text-2)',
                  borderRadius: 10,
                  padding: '1px 7px',
                  fontSize: 12,
                }}>{contacts.length}</span>
              )}
              {tab === 'campaigns' && campaigns.length > 0 && (
                <span style={{
                  marginLeft: 6,
                  background: subTab === tab ? 'rgba(255,255,255,0.25)' : 'var(--border)',
                  color: subTab === tab ? '#fff' : 'var(--text-2)',
                  borderRadius: 10,
                  padding: '1px 7px',
                  fontSize: 12,
                }}>{campaigns.length}</span>
              )}
            </button>
          ))}
        </div>
        <button
          onClick={subTab === 'contacts' ? openAddContact : openAddCampaign}
          style={{
            display: 'flex', alignItems: 'center', gap: 7,
            padding: '9px 18px', borderRadius: 8,
            background: 'var(--primary)', color: '#fff',
            fontWeight: 600, fontSize: 14,
            boxShadow: '0 1px 3px rgba(37,99,235,0.25)',
            transition: 'background 0.15s',
          }}
        >
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
            <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
          </svg>
          {subTab === 'contacts' ? 'Add Contact' : 'New Campaign'}
        </button>
      </div>

      {/* ──────── CONTACTS ──────── */}
      {subTab === 'contacts' && (
        <div style={{ background: 'var(--surface)', borderRadius: 'var(--radius-lg)', border: '1px solid var(--border)', boxShadow: 'var(--shadow-sm)' }}>
          <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--border)', display: 'flex', alignItems: 'center', gap: 12 }}>
            <SearchInput value={contactSearch} onChange={setContactSearch} placeholder="Search contacts..." />
            <span style={{ fontSize: 13, color: 'var(--text-3)', whiteSpace: 'nowrap' }}>
              {filteredContacts.length} of {contacts.length}
            </span>
          </div>

          {filteredContacts.length === 0 ? (
            <EmptyState
              icon={<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="var(--text-3)" strokeWidth="1.5"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>}
              title={contacts.length === 0 ? 'No contacts yet' : 'No results'}
              subtitle={contacts.length === 0 ? 'Add your first contact to get started' : 'Try a different search term'}
            />
          ) : (
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                  <tr style={{ background: 'var(--surface-2)', borderBottom: '1px solid var(--border)' }}>
                    {['Name', 'Company', 'Job Title', 'Email', 'Tags', ''].map(h => (
                      <th key={h} style={{ padding: '10px 16px', textAlign: 'left', fontSize: 12, fontWeight: 600, color: 'var(--text-3)', letterSpacing: '0.04em', textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredContacts.map((c, i) => (
                    <tr key={c.id} style={{ borderBottom: i < filteredContacts.length - 1 ? '1px solid var(--border)' : 'none', transition: 'background 0.1s' }}>
                      <td style={{ padding: '12px 16px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                          <Avatar name={`${c.firstName} ${c.lastName}`} />
                          <div>
                            <div style={{ fontWeight: 600, color: 'var(--text)', fontSize: 14 }}>{c.firstName} {c.lastName}</div>
                            {c.phone && <div style={{ fontSize: 12, color: 'var(--text-3)' }}>{c.phone}</div>}
                          </div>
                        </div>
                      </td>
                      <td style={{ padding: '12px 16px', color: 'var(--text-2)' }}>{c.company || '—'}</td>
                      <td style={{ padding: '12px 16px', color: 'var(--text-2)' }}>{c.jobTitle || '—'}</td>
                      <td style={{ padding: '12px 16px' }}>
                        <a href={`mailto:${c.email}`} style={{ color: 'var(--primary)', textDecoration: 'none', fontSize: 13 }}>{c.email}</a>
                      </td>
                      <td style={{ padding: '12px 16px' }}>
                        <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                          {c.tags.map(tag => (
                            <span key={tag} style={{ background: 'var(--primary-light)', color: 'var(--primary)', borderRadius: 4, padding: '2px 8px', fontSize: 12, fontWeight: 500 }}>{tag}</span>
                          ))}
                        </div>
                      </td>
                      <td style={{ padding: '12px 16px' }}>
                        <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                          <IconButton title="Edit" onClick={() => openEditContact(c)}>
                            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </IconButton>
                          <IconButton title="Delete" danger onClick={() => setDeleteConfirm({ type: 'contact', id: c.id })}>
                            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/></svg>
                          </IconButton>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* ──────── CAMPAIGNS ──────── */}
      {subTab === 'campaigns' && (
        <div>
          {/* Stats row */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 16, marginBottom: 24 }}>
            {[
              { label: 'Total Campaigns', value: campaigns.length, icon: '📊' },
              { label: 'Sent', value: campaigns.filter(c => c.status === 'sent').length, icon: '✅' },
              { label: 'Drafts', value: campaigns.filter(c => c.status === 'draft').length, icon: '📝' },
            ].map(stat => (
              <div key={stat.label} style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 'var(--radius-lg)', padding: '18px 20px', boxShadow: 'var(--shadow-sm)' }}>
                <div style={{ fontSize: 22, marginBottom: 4 }}>{stat.icon}</div>
                <div style={{ fontSize: 28, fontWeight: 700, color: 'var(--text)' }}>{stat.value}</div>
                <div style={{ fontSize: 13, color: 'var(--text-3)', fontWeight: 500 }}>{stat.label}</div>
              </div>
            ))}
          </div>

          <div style={{ background: 'var(--surface)', borderRadius: 'var(--radius-lg)', border: '1px solid var(--border)', boxShadow: 'var(--shadow-sm)' }}>
            <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--border)' }}>
              <SearchInput value={campaignSearch} onChange={setCampaignSearch} placeholder="Search campaigns..." />
            </div>

            {filteredCampaigns.length === 0 ? (
              <EmptyState
                icon={<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="var(--text-3)" strokeWidth="1.5"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>}
                title={campaigns.length === 0 ? 'No campaigns yet' : 'No results'}
                subtitle={campaigns.length === 0 ? 'Create your first email campaign' : 'Try a different search term'}
              />
            ) : (
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr style={{ background: 'var(--surface-2)', borderBottom: '1px solid var(--border)' }}>
                      {['Campaign', 'Template', 'Recipients', 'Subject', 'Status', 'Created', ''].map(h => (
                        <th key={h} style={{ padding: '10px 16px', textAlign: 'left', fontSize: 12, fontWeight: 600, color: 'var(--text-3)', letterSpacing: '0.04em', textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredCampaigns.map((c, i) => (
                      <tr key={c.id} style={{ borderBottom: i < filteredCampaigns.length - 1 ? '1px solid var(--border)' : 'none' }}>
                        <td style={{ padding: '12px 16px', fontWeight: 600, color: 'var(--text)' }}>{c.name}</td>
                        <td style={{ padding: '12px 16px', color: 'var(--text-2)' }}>{templateName(c.templateId)}</td>
                        <td style={{ padding: '12px 16px', color: 'var(--text-2)' }}>{c.contactIds.length}</td>
                        <td style={{ padding: '12px 16px', color: 'var(--text-2)', maxWidth: 220 }}>
                          <span style={{ display: 'block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{c.customSubject}</span>
                        </td>
                        <td style={{ padding: '12px 16px' }}>
                          <StatusBadge status={c.status} />
                        </td>
                        <td style={{ padding: '12px 16px', color: 'var(--text-3)', fontSize: 13, whiteSpace: 'nowrap' }}>
                          {new Date(c.createdAt).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}
                        </td>
                        <td style={{ padding: '12px 16px' }}>
                          <IconButton title="Delete" danger onClick={() => setDeleteConfirm({ type: 'campaign', id: c.id })}>
                            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/></svg>
                          </IconButton>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ──────── CONTACT MODAL ──────── */}
      {showContactModal && (
        <Modal title={editingContact ? 'Edit Contact' : 'Add Contact'} onClose={() => setShowContactModal(false)} width={500}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
            <Field label="First Name *">
              <input value={contactForm.firstName} onChange={e => setContactForm(p => ({ ...p, firstName: e.target.value }))} placeholder="Jane" style={inputStyle} />
            </Field>
            <Field label="Last Name">
              <input value={contactForm.lastName} onChange={e => setContactForm(p => ({ ...p, lastName: e.target.value }))} placeholder="Smith" style={inputStyle} />
            </Field>
            <Field label="Email *">
              <input type="email" value={contactForm.email} onChange={e => setContactForm(p => ({ ...p, email: e.target.value }))} placeholder="jane@company.com" style={inputStyle} />
            </Field>
            <Field label="Phone">
              <input value={contactForm.phone} onChange={e => setContactForm(p => ({ ...p, phone: e.target.value }))} placeholder="+1 234 567 8900" style={inputStyle} />
            </Field>
            <Field label="Company">
              <input value={contactForm.company} onChange={e => setContactForm(p => ({ ...p, company: e.target.value }))} placeholder="Acme Logistics" style={inputStyle} />
            </Field>
            <Field label="Job Title">
              <input value={contactForm.jobTitle} onChange={e => setContactForm(p => ({ ...p, jobTitle: e.target.value }))} placeholder="Facility Manager" style={inputStyle} />
            </Field>
            <Field label="Tags (comma separated)" style={{ gridColumn: '1 / -1' }}>
              <input value={contactForm.tags} onChange={e => setContactForm(p => ({ ...p, tags: e.target.value }))} placeholder="warehouse, logistics, priority" style={inputStyle} />
            </Field>
          </div>
          <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 20 }}>
            <SecondaryBtn onClick={() => setShowContactModal(false)}>Cancel</SecondaryBtn>
            <PrimaryBtn onClick={saveContact} disabled={!contactForm.firstName.trim() || !contactForm.email.trim()}>
              {editingContact ? 'Save Changes' : 'Add Contact'}
            </PrimaryBtn>
          </div>
        </Modal>
      )}

      {/* ──────── CAMPAIGN MODAL ──────── */}
      {showCampaignModal && (
        <Modal title="New Campaign" onClose={() => setShowCampaignModal(false)} width={680}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
              <Field label="Campaign Name *" style={{ gridColumn: '1 / -1' }}>
                <input value={campaignForm.name} onChange={e => setCampaignForm(p => ({ ...p, name: e.target.value }))} placeholder="Q3 Warehouse Managers Outreach" style={inputStyle} />
              </Field>
              <Field label="Email Template *">
                <select value={campaignForm.templateId} onChange={e => setCampaignForm(p => ({ ...p, templateId: e.target.value, customSubject: '' }))} style={inputStyle}>
                  <option value="">Select a template…</option>
                  {EMAIL_TEMPLATES.map(t => (
                    <option key={t.id} value={t.id}>{t.name}</option>
                  ))}
                </select>
              </Field>
              <Field label="Custom Subject (optional)">
                <input value={campaignForm.customSubject} onChange={e => setCampaignForm(p => ({ ...p, customSubject: e.target.value }))}
                  placeholder={selectedTemplate?.subject ?? 'Leave blank to use template subject'}
                  style={inputStyle} />
              </Field>
              <Field label="Sender Name">
                <input value={campaignForm.senderName} onChange={e => setCampaignForm(p => ({ ...p, senderName: e.target.value }))} placeholder="Your Name" style={inputStyle} />
              </Field>
              <Field label="Sender Email">
                <input type="email" value={campaignForm.senderEmail} onChange={e => setCampaignForm(p => ({ ...p, senderEmail: e.target.value }))} placeholder="you@company.com" style={inputStyle} />
              </Field>
            </div>

            {/* Contact selector */}
            <Field label={`Select Recipients (${campaignForm.contactIds.length} selected)`}>
              {contacts.length === 0 ? (
                <div style={{ padding: '12px 14px', background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 6, color: 'var(--text-3)', fontSize: 13 }}>
                  No contacts yet — add contacts first.
                </div>
              ) : (
                <div style={{ maxHeight: 180, overflowY: 'auto', border: '1px solid var(--border)', borderRadius: 6 }}>
                  {contacts.map((c, i) => (
                    <label key={c.id} style={{
                      display: 'flex', alignItems: 'center', gap: 10,
                      padding: '9px 14px',
                      borderBottom: i < contacts.length - 1 ? '1px solid var(--border)' : 'none',
                      cursor: 'pointer',
                      background: campaignForm.contactIds.includes(c.id) ? 'var(--primary-light)' : 'var(--surface)',
                      transition: 'background 0.1s',
                    }}>
                      <input
                        type="checkbox"
                        checked={campaignForm.contactIds.includes(c.id)}
                        onChange={() => toggleCampaignContact(c.id)}
                        style={{ accentColor: 'var(--primary)', width: 15, height: 15 }}
                      />
                      <Avatar name={`${c.firstName} ${c.lastName}`} size={26} />
                      <div style={{ flex: 1 }}>
                        <span style={{ fontWeight: 500, fontSize: 13 }}>{c.firstName} {c.lastName}</span>
                        {c.company && <span style={{ color: 'var(--text-3)', fontSize: 12 }}> · {c.company}</span>}
                      </div>
                      <span style={{ fontSize: 12, color: 'var(--text-3)' }}>{c.email}</span>
                    </label>
                  ))}
                </div>
              )}
            </Field>

            {/* Preview toggle */}
            {selectedTemplate && (
              <div>
                <button onClick={() => setPreviewOpen(p => !p)} style={{
                  display: 'flex', alignItems: 'center', gap: 6,
                  fontSize: 13, fontWeight: 500, color: 'var(--primary)',
                  background: 'none', padding: 0,
                }}>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                    <circle cx="12" cy="12" r="3"/>
                  </svg>
                  {previewOpen ? 'Hide preview' : 'Preview email'}
                </button>
                {previewOpen && (
                  <div style={{ marginTop: 10, background: 'var(--surface-2)', border: '1px solid var(--border)', borderRadius: 6, padding: 16 }}>
                    <div style={{ marginBottom: 8 }}>
                      <span style={{ fontSize: 12, fontWeight: 600, color: 'var(--text-3)', textTransform: 'uppercase', letterSpacing: '0.04em' }}>Subject: </span>
                      <span style={{ fontSize: 13, fontWeight: 600, color: 'var(--text)' }}>{campaignForm.customSubject || selectedTemplate.subject}</span>
                    </div>
                    <pre style={{ whiteSpace: 'pre-wrap', fontSize: 13, color: 'var(--text-2)', lineHeight: 1.7, fontFamily: 'inherit', margin: 0 }}>
                      {getPreviewBody()}
                    </pre>
                  </div>
                )}
              </div>
            )}
          </div>

          <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 20 }}>
            <SecondaryBtn onClick={() => setShowCampaignModal(false)}>Cancel</SecondaryBtn>
            <SecondaryBtn onClick={() => saveCampaign('draft')} disabled={!campaignForm.name.trim() || !campaignForm.templateId}>
              Save Draft
            </SecondaryBtn>
            <PrimaryBtn onClick={() => saveCampaign('sent')} disabled={!campaignForm.name.trim() || !campaignForm.templateId}>
              Mark as Sent
            </PrimaryBtn>
          </div>
        </Modal>
      )}

      {/* ──────── DELETE CONFIRM ──────── */}
      {deleteConfirm && (
        <Modal title="Confirm Delete" onClose={() => setDeleteConfirm(null)} width={400}>
          <p style={{ color: 'var(--text-2)', lineHeight: 1.6 }}>
            Are you sure you want to delete this {deleteConfirm.type}? This action cannot be undone.
          </p>
          <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 20 }}>
            <SecondaryBtn onClick={() => setDeleteConfirm(null)}>Cancel</SecondaryBtn>
            <button onClick={deleteItem} style={{ padding: '9px 18px', borderRadius: 7, background: 'var(--danger)', color: '#fff', fontWeight: 600, fontSize: 14 }}>
              Delete
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── Small shared components ───

function Avatar({ name, size = 32 }: { name: string; size?: number }) {
  const initials = name.split(' ').map(p => p[0]).slice(0, 2).join('').toUpperCase();
  const colors = ['#2563eb', '#7c3aed', '#059669', '#d97706', '#dc2626', '#0891b2'];
  const idx = name.charCodeAt(0) % colors.length;
  return (
    <div style={{
      width: size, height: size, borderRadius: '50%',
      background: colors[idx], color: '#fff',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      fontSize: size * 0.38, fontWeight: 700, flexShrink: 0,
    }}>{initials}</div>
  );
}

function StatusBadge({ status }: { status: 'draft' | 'sent' }) {
  const map = {
    draft: { bg: '#f1f5f9', color: '#64748b', label: 'Draft' },
    sent: { bg: 'var(--success-light)', color: 'var(--success)', label: 'Sent' },
  };
  const s = map[status];
  return (
    <span style={{ background: s.bg, color: s.color, borderRadius: 5, padding: '3px 10px', fontSize: 12, fontWeight: 600 }}>
      {s.label}
    </span>
  );
}

function SearchInput({ value, onChange, placeholder }: { value: string; onChange: (v: string) => void; placeholder: string }) {
  return (
    <div style={{ position: 'relative', flex: 1 }}>
      <svg style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', pointerEvents: 'none' }} width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="var(--text-3)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
        <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
      </svg>
      <input value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder}
        style={{ ...inputStyle, paddingLeft: 34, minWidth: 220 }} />
    </div>
  );
}

function EmptyState({ icon, title, subtitle }: { icon: React.ReactNode; title: string; subtitle: string }) {
  return (
    <div style={{ padding: '48px 24px', textAlign: 'center', color: 'var(--text-3)' }}>
      <div style={{ marginBottom: 12, display: 'flex', justifyContent: 'center' }}>{icon}</div>
      <div style={{ fontWeight: 600, fontSize: 15, color: 'var(--text-2)', marginBottom: 4 }}>{title}</div>
      <div style={{ fontSize: 13 }}>{subtitle}</div>
    </div>
  );
}

function Modal({ title, onClose, width = 500, children }: { title: string; onClose: () => void; width?: number; children: React.ReactNode }) {
  return (
    <div style={{
      position: 'fixed', inset: 0, zIndex: 200,
      background: 'rgba(15,23,42,0.45)', backdropFilter: 'blur(2px)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
    }} onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{
        background: 'var(--surface)', borderRadius: 'var(--radius-lg)',
        boxShadow: 'var(--shadow-lg)', width: '100%', maxWidth: width,
        maxHeight: '90vh', overflowY: 'auto',
      }}>
        <div style={{ padding: '18px 22px', borderBottom: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <h2 style={{ fontSize: 16, fontWeight: 700, color: 'var(--text)' }}>{title}</h2>
          <button onClick={onClose} style={{ background: 'none', padding: 4, color: 'var(--text-3)', borderRadius: 4, lineHeight: 0 }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
            </svg>
          </button>
        </div>
        <div style={{ padding: '20px 22px' }}>{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children, style }: { label: string; children: React.ReactNode; style?: React.CSSProperties }) {
  return (
    <div style={style}>
      <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: 'var(--text-2)', marginBottom: 5, letterSpacing: '0.02em' }}>{label}</label>
      {children}
    </div>
  );
}

function PrimaryBtn({ onClick, disabled, children }: { onClick: () => void; disabled?: boolean; children: React.ReactNode }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{
      padding: '9px 18px', borderRadius: 7,
      background: disabled ? 'var(--border-2)' : 'var(--primary)',
      color: disabled ? 'var(--text-3)' : '#fff',
      fontWeight: 600, fontSize: 14, cursor: disabled ? 'not-allowed' : 'pointer',
    }}>{children}</button>
  );
}

function SecondaryBtn({ onClick, disabled, children }: { onClick: () => void; disabled?: boolean; children: React.ReactNode }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{
      padding: '9px 18px', borderRadius: 7,
      background: 'var(--surface)', border: '1px solid var(--border)',
      color: disabled ? 'var(--text-3)' : 'var(--text-2)',
      fontWeight: 500, fontSize: 14, cursor: disabled ? 'not-allowed' : 'pointer',
    }}>{children}</button>
  );
}

function IconButton({ onClick, title, danger, children }: { onClick: () => void; title: string; danger?: boolean; children: React.ReactNode }) {
  return (
    <button onClick={onClick} title={title} style={{
      background: 'none', padding: 5, borderRadius: 5, lineHeight: 0,
      color: danger ? 'var(--danger)' : 'var(--text-3)',
      transition: 'background 0.1s, color 0.1s',
    }}>{children}</button>
  );
}

const inputStyle: React.CSSProperties = {
  width: '100%', padding: '8px 12px',
  border: '1px solid var(--border)',
  borderRadius: 6, fontSize: 14,
  background: 'var(--surface)',
  color: 'var(--text)',
  transition: 'border-color 0.15s',
};
