import { useState, useEffect, useCallback } from 'react';
import type { LinkedInPost } from '../types';
import { MACHINES, POST_TYPES, TONES, generateLinkedInPost } from '../data/linkedinTemplates';

const STORAGE_POSTS = 'fiorentini_linkedin_posts';

function uid() {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function loadFromStorage<T>(key: string, fallback: T): T {
  try {
    const raw = localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch { return fallback; }
}

export default function LinkedInContent() {
  const [postType, setPostType] = useState('product-spotlight');
  const [machineId, setMachineId] = useState('ride-on-scrubber');
  const [tone, setTone] = useState('professional');
  const [customPoint, setCustomPoint] = useState('');
  const [generated, setGenerated] = useState('');
  const [savedPosts, setSavedPosts] = useState<LinkedInPost[]>(() =>
    loadFromStorage<LinkedInPost[]>(STORAGE_POSTS, [])
  );
  const [copied, setCopied] = useState(false);
  const [savedAnim, setSavedAnim] = useState(false);
  const [toast, setToast] = useState<string | null>(null);
  const [subTab, setSubTab] = useState<'generator' | 'library'>('generator');
  const [deleteId, setDeleteId] = useState<string | null>(null);

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2800);
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_POSTS, JSON.stringify(savedPosts));
  }, [savedPosts]);

  const generate = () => {
    const post = generateLinkedInPost({ postType, machineId, tone, customPoint: customPoint.trim() || undefined });
    setGenerated(post);
  };

  const copyToClipboard = async () => {
    if (!generated) return;
    await navigator.clipboard.writeText(generated);
    setCopied(true);
    showToast('Copied to clipboard');
    setTimeout(() => setCopied(false), 2000);
  };

  const savePost = () => {
    if (!generated) return;
    const machineLabel = MACHINES.find(m => m.id === machineId)?.label ?? machineId;
    const typeLabel = POST_TYPES.find(p => p.id === postType)?.label ?? postType;
    const post: LinkedInPost = {
      id: uid(),
      postType: typeLabel,
      machine: machineLabel,
      tone,
      content: generated,
      savedAt: new Date().toISOString(),
    };
    setSavedPosts(prev => [post, ...prev]);
    setSavedAnim(true);
    showToast('Post saved to library');
    setTimeout(() => setSavedAnim(false), 600);
  };

  const deletePost = (id: string) => {
    setSavedPosts(prev => prev.filter(p => p.id !== id));
    setDeleteId(null);
    showToast('Post deleted');
  };

  const charCount = generated.length;

  return (
    <div>
      {/* Toast */}
      {toast && (
        <div style={{
          position: 'fixed', bottom: 24, right: 24, zIndex: 999,
          background: 'var(--success)', color: '#fff',
          padding: '10px 18px', borderRadius: 8,
          boxShadow: 'var(--shadow-lg)', fontWeight: 500, fontSize: 14,
        }}>{toast}</div>
      )}

      {/* Sub-tabs */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 24 }}>
        <div style={{ display: 'flex', gap: 2, background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 8, padding: 3 }}>
          {(['generator', 'library'] as const).map(tab => (
            <button key={tab} onClick={() => setSubTab(tab)} style={{
              padding: '7px 18px', borderRadius: 6, fontWeight: 500, fontSize: 14, transition: 'all 0.15s',
              background: subTab === tab ? 'var(--primary)' : 'transparent',
              color: subTab === tab ? '#fff' : 'var(--text-2)',
            }}>
              {tab.charAt(0).toUpperCase() + tab.slice(1)}
              {tab === 'library' && savedPosts.length > 0 && (
                <span style={{ marginLeft: 6, background: subTab === tab ? 'rgba(255,255,255,0.25)' : 'var(--border)', color: subTab === tab ? '#fff' : 'var(--text-2)', borderRadius: 10, padding: '1px 7px', fontSize: 12 }}>
                  {savedPosts.length}
                </span>
              )}
            </button>
          ))}
        </div>
      </div>

      {/* ──────── GENERATOR ──────── */}
      {subTab === 'generator' && (
        <div style={{ display: 'grid', gridTemplateColumns: '340px 1fr', gap: 24, alignItems: 'start' }}>
          {/* Left: Controls */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <Card>
              <SectionTitle>Content Settings</SectionTitle>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
                <Field label="Post Type">
                  <select value={postType} onChange={e => setPostType(e.target.value)} style={inputStyle}>
                    {POST_TYPES.map(p => <option key={p.id} value={p.id}>{p.label}</option>)}
                  </select>
                </Field>
                <Field label="Machine / Product">
                  <select value={machineId} onChange={e => setMachineId(e.target.value)} style={inputStyle}>
                    {MACHINES.map(m => <option key={m.id} value={m.id}>{m.label}</option>)}
                  </select>
                </Field>
                <Field label="Tone">
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 6 }}>
                    {TONES.map(t => (
                      <button key={t.id} onClick={() => setTone(t.id)} style={{
                        padding: '7px 4px', borderRadius: 6, fontSize: 12, fontWeight: 500,
                        border: tone === t.id ? '2px solid var(--primary)' : '1px solid var(--border)',
                        background: tone === t.id ? 'var(--primary-light)' : 'var(--surface)',
                        color: tone === t.id ? 'var(--primary)' : 'var(--text-2)',
                        transition: 'all 0.1s',
                      }}>{t.label}</button>
                    ))}
                  </div>
                </Field>
                <Field label="Custom Talking Point (optional)">
                  <textarea
                    value={customPoint}
                    onChange={e => setCustomPoint(e.target.value)}
                    placeholder="e.g. We offer 12-month interest-free financing..."
                    rows={3}
                    style={{ ...inputStyle, resize: 'vertical', lineHeight: 1.5 }}
                  />
                </Field>
              </div>
              <button onClick={generate} style={{
                width: '100%', marginTop: 16, padding: '11px 0',
                borderRadius: 8, background: 'var(--primary)', color: '#fff',
                fontWeight: 700, fontSize: 15, letterSpacing: '-0.01em',
                boxShadow: '0 1px 3px rgba(37,99,235,0.25)',
                transition: 'background 0.15s',
                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
              }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>
                </svg>
                Generate Post
              </button>
            </Card>

            {/* Machine Info Card */}
            {machineId && (
              <Card>
                <SectionTitle>Machine Snapshot</SectionTitle>
                {(() => {
                  const m = MACHINES.find(x => x.id === machineId)!;
                  return (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13 }}>
                        <span style={{ color: 'var(--text-3)', fontWeight: 500 }}>Coverage</span>
                        <span style={{ color: 'var(--text)', fontWeight: 600 }}>{m.coverage}</span>
                      </div>
                      <div style={{ fontSize: 13 }}>
                        <div style={{ color: 'var(--text-3)', fontWeight: 500, marginBottom: 4 }}>Best for</div>
                        <div style={{ color: 'var(--text-2)' }}>{m.bestFor}</div>
                      </div>
                      <div style={{ fontSize: 13 }}>
                        <div style={{ color: 'var(--text-3)', fontWeight: 500, marginBottom: 4 }}>Key benefits</div>
                        <ul style={{ paddingLeft: 16, margin: 0, color: 'var(--text-2)', display: 'flex', flexDirection: 'column', gap: 3 }}>
                          {m.keyBenefits.slice(0, 3).map((b, i) => (
                            <li key={i} style={{ fontSize: 12 }}>{b}</li>
                          ))}
                        </ul>
                      </div>
                    </div>
                  );
                })()}
              </Card>
            )}
          </div>

          {/* Right: Preview */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <Card>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
                <SectionTitle style={{ marginBottom: 0 }}>Post Preview</SectionTitle>
                {generated && (
                  <div style={{ display: 'flex', gap: 8 }}>
                    <span style={{ fontSize: 12, color: charCount > 3000 ? 'var(--danger)' : 'var(--text-3)', fontWeight: 500, alignSelf: 'center' }}>
                      {charCount} / 3000
                    </span>
                    <ActionBtn onClick={copyToClipboard} icon={copied
                      ? <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
                      : <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                    }>{copied ? 'Copied!' : 'Copy'}</ActionBtn>
                    <ActionBtn onClick={savePost} primary icon={
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: savedAnim ? 'scale(1.3)' : 'scale(1)', transition: 'transform 0.2s' }}>
                        <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>
                      </svg>
                    }>Save</ActionBtn>
                  </div>
                )}
              </div>

              {!generated ? (
                <div style={{
                  minHeight: 400, display: 'flex', flexDirection: 'column',
                  alignItems: 'center', justifyContent: 'center',
                  background: 'var(--surface-2)', borderRadius: 8,
                  border: '2px dashed var(--border)', gap: 12,
                }}>
                  <LinkedInIcon size={40} />
                  <div style={{ color: 'var(--text-3)', fontWeight: 500, fontSize: 15 }}>Your post will appear here</div>
                  <div style={{ color: 'var(--text-3)', fontSize: 13 }}>Configure the settings and click Generate Post</div>
                </div>
              ) : (
                <div>
                  {/* LinkedIn-style post card */}
                  <div style={{
                    border: '1px solid #e0dfdf', borderRadius: 8,
                    background: '#fff', overflow: 'hidden',
                    fontFamily: '-apple-system, system-ui, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
                  }}>
                    {/* Profile row */}
                    <div style={{ padding: '12px 16px 0', display: 'flex', gap: 10, alignItems: 'center' }}>
                      <div style={{
                        width: 48, height: 48, borderRadius: '50%',
                        background: 'linear-gradient(135deg, #2563eb 0%, #7c3aed 100%)',
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        color: '#fff', fontWeight: 700, fontSize: 16, flexShrink: 0,
                      }}>F</div>
                      <div>
                        <div style={{ fontWeight: 600, fontSize: 14, color: '#000' }}>Fiorentini</div>
                        <div style={{ fontSize: 12, color: '#666' }}>Industrial Cleaning Solutions</div>
                        <div style={{ fontSize: 11, color: '#999', display: 'flex', alignItems: 'center', gap: 3 }}>
                          Just now · <LinkedInGlobeIcon />
                        </div>
                      </div>
                    </div>
                    {/* Post body */}
                    <div style={{ padding: '12px 16px 16px' }}>
                      <pre style={{
                        whiteSpace: 'pre-wrap', fontSize: 14, color: '#000',
                        lineHeight: 1.6, fontFamily: 'inherit', margin: 0,
                      }}>{generated}</pre>
                    </div>
                    {/* Engagement row */}
                    <div style={{ borderTop: '1px solid #e0dfdf', padding: '4px 16px', display: 'flex', gap: 0 }}>
                      {['👍 Like', '💬 Comment', '🔁 Repost', '📤 Send'].map(action => (
                        <div key={action} style={{ flex: 1, textAlign: 'center', padding: '8px 0', fontSize: 13, color: '#666', fontWeight: 500 }}>{action}</div>
                      ))}
                    </div>
                  </div>

                  {/* Editable textarea */}
                  <div style={{ marginTop: 12 }}>
                    <div style={{ fontSize: 12, color: 'var(--text-3)', marginBottom: 5 }}>Edit post text:</div>
                    <textarea
                      value={generated}
                      onChange={e => setGenerated(e.target.value)}
                      rows={12}
                      style={{ ...inputStyle, resize: 'vertical', lineHeight: 1.6, fontSize: 13 }}
                    />
                  </div>
                </div>
              )}
            </Card>
          </div>
        </div>
      )}

      {/* ──────── LIBRARY ──────── */}
      {subTab === 'library' && (
        <div>
          {savedPosts.length === 0 ? (
            <div style={{
              background: 'var(--surface)', border: '1px solid var(--border)',
              borderRadius: 'var(--radius-lg)', padding: '64px 24px',
              textAlign: 'center', boxShadow: 'var(--shadow-sm)',
            }}>
              <LinkedInIcon size={40} style={{ margin: '0 auto 12px' }} />
              <div style={{ fontWeight: 600, fontSize: 15, color: 'var(--text-2)', marginBottom: 4 }}>No saved posts yet</div>
              <div style={{ fontSize: 13, color: 'var(--text-3)' }}>Generate a post and click Save to build your library</div>
              <button onClick={() => setSubTab('generator')} style={{ marginTop: 16, padding: '8px 20px', borderRadius: 7, background: 'var(--primary)', color: '#fff', fontWeight: 600, fontSize: 14 }}>
                Go to Generator
              </button>
            </div>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(400px, 1fr))', gap: 20 }}>
              {savedPosts.map(post => (
                <div key={post.id} style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 'var(--radius-lg)', boxShadow: 'var(--shadow-sm)', overflow: 'hidden' }}>
                  {/* Card header */}
                  <div style={{ padding: '12px 16px', borderBottom: '1px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: 'var(--surface-2)' }}>
                    <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                      <span style={{ background: 'var(--primary-light)', color: 'var(--primary)', borderRadius: 4, padding: '2px 8px', fontSize: 11, fontWeight: 600 }}>{post.postType}</span>
                      <span style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 4, padding: '2px 8px', fontSize: 11, color: 'var(--text-3)', fontWeight: 500 }}>{post.machine}</span>
                    </div>
                    <div style={{ fontSize: 12, color: 'var(--text-3)' }}>
                      {new Date(post.savedAt).toLocaleDateString('en-GB', { day: '2-digit', month: 'short' })}
                    </div>
                  </div>
                  {/* Content */}
                  <div style={{ padding: '14px 16px' }}>
                    <pre style={{ whiteSpace: 'pre-wrap', fontSize: 13, color: 'var(--text-2)', lineHeight: 1.6, fontFamily: 'inherit', margin: 0, maxHeight: 180, overflowY: 'auto' }}>
                      {post.content}
                    </pre>
                  </div>
                  {/* Actions */}
                  <div style={{ padding: '10px 16px', borderTop: '1px solid var(--border)', display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                    <ActionBtn onClick={async () => {
                      await navigator.clipboard.writeText(post.content);
                      showToast('Copied to clipboard');
                    }} icon={
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                    }>Copy</ActionBtn>
                    <ActionBtn onClick={() => setDeleteId(post.id)} danger icon={
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/></svg>
                    }>Delete</ActionBtn>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Delete confirm modal */}
      {deleteId && (
        <div style={{
          position: 'fixed', inset: 0, zIndex: 200,
          background: 'rgba(15,23,42,0.45)', backdropFilter: 'blur(2px)',
          display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20,
        }} onClick={e => e.target === e.currentTarget && setDeleteId(null)}>
          <div style={{ background: 'var(--surface)', borderRadius: 'var(--radius-lg)', boxShadow: 'var(--shadow-lg)', width: '100%', maxWidth: 380, padding: '22px 24px' }}>
            <h3 style={{ fontSize: 16, fontWeight: 700, marginBottom: 8 }}>Delete post?</h3>
            <p style={{ color: 'var(--text-2)', fontSize: 14, lineHeight: 1.6 }}>This post will be permanently removed from your library.</p>
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 20 }}>
              <button onClick={() => setDeleteId(null)} style={{ padding: '9px 18px', borderRadius: 7, background: 'var(--surface)', border: '1px solid var(--border)', color: 'var(--text-2)', fontWeight: 500, fontSize: 14 }}>Cancel</button>
              <button onClick={() => deletePost(deleteId)} style={{ padding: '9px 18px', borderRadius: 7, background: 'var(--danger)', color: '#fff', fontWeight: 600, fontSize: 14 }}>Delete</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Small shared components ───

function Card({ children }: { children: React.ReactNode }) {
  return (
    <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: 'var(--radius-lg)', padding: '20px', boxShadow: 'var(--shadow-sm)' }}>
      {children}
    </div>
  );
}

function SectionTitle({ children, style }: { children: React.ReactNode; style?: React.CSSProperties }) {
  return (
    <div style={{ fontSize: 13, fontWeight: 700, color: 'var(--text)', marginBottom: 14, textTransform: 'uppercase', letterSpacing: '0.05em', ...style }}>
      {children}
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <label style={{ display: 'block', fontSize: 12, fontWeight: 600, color: 'var(--text-2)', marginBottom: 5 }}>{label}</label>
      {children}
    </div>
  );
}

function ActionBtn({ onClick, children, icon, primary, danger }: { onClick: () => void; children: React.ReactNode; icon?: React.ReactNode; primary?: boolean; danger?: boolean }) {
  return (
    <button onClick={onClick} style={{
      display: 'flex', alignItems: 'center', gap: 5,
      padding: '6px 12px', borderRadius: 6, fontSize: 13, fontWeight: 500,
      border: primary ? 'none' : '1px solid var(--border)',
      background: primary ? 'var(--primary)' : danger ? 'transparent' : 'var(--surface)',
      color: primary ? '#fff' : danger ? 'var(--danger)' : 'var(--text-2)',
      transition: 'all 0.1s',
    }}>
      {icon}{children}
    </button>
  );
}

function LinkedInIcon({ size = 24, style }: { size?: number; style?: React.CSSProperties }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="#0a66c2" style={style}>
      <path d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 0 1-2.063-2.065 2.064 2.064 0 1 1 2.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z"/>
    </svg>
  );
}

function LinkedInGlobeIcon() {
  return (
    <svg width="12" height="12" viewBox="0 0 24 24" fill="#666">
      <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z"/>
    </svg>
  );
}

const inputStyle: React.CSSProperties = {
  width: '100%', padding: '8px 12px',
  border: '1px solid var(--border)',
  borderRadius: 6, fontSize: 14,
  background: 'var(--surface)',
  color: 'var(--text)',
};
