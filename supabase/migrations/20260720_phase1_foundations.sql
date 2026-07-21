-- Facility Command — Phase 1: Foundations
-- Run this once in the Supabase SQL Editor (Project > SQL Editor > New query).
-- Safe to re-run: every statement is idempotent (IF NOT EXISTS / ON CONFLICT DO NOTHING).

-- ============================================================
-- 1. Sites as data (replaces the hardcoded list in App.tsx)
-- ============================================================
create table if not exists sites (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  active boolean not null default true,
  created_at timestamptz not null default now()
);

alter table sites enable row level security;

drop policy if exists "sites_select_authenticated" on sites;
create policy "sites_select_authenticated" on sites for select to authenticated using (true);

drop policy if exists "sites_write_authenticated" on sites;
create policy "sites_write_authenticated" on sites for all to authenticated using (true) with check (true);

-- Seed with the sites currently hardcoded in App.tsx, so nothing breaks on switchover.
insert into sites (name) values
  ('Site1'), ('Site2'), ('Site3'), ('Site4'), ('Site5'), ('Site6'),
  ('Site7B'), ('Site7C'), ('Site8'), ('Site9'), ('Site9A'), ('Site9B'),
  ('Site10'), ('Site10A'), ('Site10B'), ('Site11'), ('Site12'),
  ('Site14'), ('Site14A'), ('Site14B'), ('Storage')
on conflict (name) do nothing;

-- ============================================================
-- 2. Unique asset codes
-- Reuses assets.serial_number as the canonical unique code (already treated
-- that way elsewhere in the app for asset matching). Partial index so
-- multiple assets without a code (null / empty) don't collide.
--
-- If this CREATE UNIQUE INDEX fails, you already have duplicate codes in
-- production data. Find them first with:
--   select serial_number, count(*) from assets
--   where serial_number is not null and serial_number <> ''
--   group by serial_number having count(*) > 1;
-- and resolve (rename or clear) before re-running this migration.
-- ============================================================
create unique index if not exists assets_serial_number_unique
  on assets (serial_number)
  where serial_number is not null and serial_number <> '';

-- ============================================================
-- 3. Invoice / PO traceability on assets
-- ============================================================
alter table assets add column if not exists invoice_number text;
alter table assets add column if not exists po_number text;
alter table assets add column if not exists purchase_date date;
