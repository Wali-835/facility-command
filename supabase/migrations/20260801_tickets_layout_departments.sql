-- Facility Command — Ticketing system upgrade
-- Adds: a department tag on users, vendor assignment + precise location on
-- tickets, department stamping on ticket events, and site floor-plan
-- layouts with named pins (room/bathroom/cafeteria) so a ticket's location
-- can be picked visually instead of just "site".
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

alter table user_roles add column if not exists department text;

alter table tickets add column if not exists vendor text;
alter table tickets add column if not exists location_detail text;

alter table ticket_events add column if not exists department text;

-- Floor plan image per site (reuses the existing public asset-documents
-- bucket under a site-layouts/ prefix — no new bucket needed).
create table if not exists site_layouts (
  id text primary key,
  site text not null unique,
  file_name text not null,
  file_path text not null,
  uploaded_by text,
  uploaded_at timestamptz not null default now()
);
alter table site_layouts enable row level security;
drop policy if exists "site_layouts_all_authenticated" on site_layouts;
create policy "site_layouts_all_authenticated" on site_layouts for all to authenticated using (true) with check (true);

-- Named pins placed on a layout image, stored as percentage coordinates
-- (0-100) so they stay correctly positioned regardless of render size.
create table if not exists site_layout_points (
  id text primary key,
  site_layout_id text not null,
  label text not null,
  x_pct numeric not null,
  y_pct numeric not null
);
create index if not exists idx_slp_layout on site_layout_points(site_layout_id);
alter table site_layout_points enable row level security;
drop policy if exists "site_layout_points_all_authenticated" on site_layout_points;
create policy "site_layout_points_all_authenticated" on site_layout_points for all to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
