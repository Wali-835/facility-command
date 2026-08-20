-- Facility Command — Site document archive (operational license, approved
-- layouts, permits, etc.)
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).
--
-- Separate from site_layouts (the interactive floor-plan image used to pin
-- a ticket's location) — this is a general per-site document library, same
-- shape as asset_documents, with optional issue/expiry dates so licenses
-- and permits can be flagged as expiring/expired.

create table if not exists site_documents (
  id text primary key,
  site text not null,
  document_type text not null, -- 'Operational License' | 'Approved Layout' | 'Fire Safety Certificate' | 'Civil Defense Permit' | 'Environmental Permit' | 'Other'
  file_name text not null,
  file_path text not null,
  issue_date date,
  expiry_date date,
  notes text,
  uploaded_by text,
  uploaded_at timestamptz not null default now()
);
create index if not exists idx_site_documents_site on site_documents(site);
alter table site_documents enable row level security;
drop policy if exists "site_documents_all_authenticated" on site_documents;
create policy "site_documents_all_authenticated" on site_documents for all to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
