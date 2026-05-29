import { useState } from 'react';
import EmailCampaigns from './components/EmailCampaigns';
import LinkedInContent from './components/LinkedInContent';

type Tab = 'email' | 'linkedin';

export default function App() {
  const [activeTab, setActiveTab] = useState<Tab>('email');

  return (
    <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>
      {/* Header */}
      <header style={{
        background: 'var(--surface)',
        borderBottom: '1px solid var(--border)',
        boxShadow: 'var(--shadow-sm)',
        position: 'sticky',
        top: 0,
        zIndex: 100,
      }}>
        <div style={{
          maxWidth: 1200,
          margin: '0 auto',
          padding: '0 24px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          height: 64,
        }}>
          {/* Logo / Brand */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <div style={{
              width: 36,
              height: 36,
              background: 'var(--primary)',
              borderRadius: 8,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
            }}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>
                <polyline points="9 22 9 12 15 12 15 22"/>
              </svg>
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 15, color: 'var(--text)', lineHeight: 1.2 }}>
                Fiorentini
              </div>
              <div style={{ fontSize: 11, color: 'var(--text-3)', fontWeight: 500, letterSpacing: '0.05em', textTransform: 'uppercase' }}>
                Marketing Hub
              </div>
            </div>
          </div>

          {/* Tab Navigation */}
          <nav style={{ display: 'flex', gap: 4 }}>
            <TabButton
              active={activeTab === 'email'}
              onClick={() => setActiveTab('email')}
              icon={
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
                  <polyline points="22,6 12,13 2,6"/>
                </svg>
              }
              label="Email Campaigns"
            />
            <TabButton
              active={activeTab === 'linkedin'}
              onClick={() => setActiveTab('linkedin')}
              icon={
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M16 8a6 6 0 0 1 6 6v7h-4v-7a2 2 0 0 0-2-2 2 2 0 0 0-2 2v7h-4v-7a6 6 0 0 1 6-6z"/>
                  <rect x="2" y="9" width="4" height="12"/>
                  <circle cx="4" cy="4" r="2"/>
                </svg>
              }
              label="LinkedIn Content"
            />
          </nav>
        </div>
      </header>

      {/* Main */}
      <main style={{ flex: 1, maxWidth: 1200, width: '100%', margin: '0 auto', padding: '28px 24px' }}>
        {activeTab === 'email' && <EmailCampaigns />}
        {activeTab === 'linkedin' && <LinkedInContent />}
      </main>
    </div>
  );
}

function TabButton({
  active,
  onClick,
  icon,
  label,
}: {
  active: boolean;
  onClick: () => void;
  icon: React.ReactNode;
  label: string;
}) {
  return (
    <button
      onClick={onClick}
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: 7,
        padding: '7px 14px',
        borderRadius: 6,
        fontWeight: 500,
        fontSize: 14,
        transition: 'all 0.15s',
        background: active ? 'var(--primary-light)' : 'transparent',
        color: active ? 'var(--primary)' : 'var(--text-2)',
        border: active ? '1px solid #bfdbfe' : '1px solid transparent',
      }}
    >
      {icon}
      {label}
    </button>
  );
}
