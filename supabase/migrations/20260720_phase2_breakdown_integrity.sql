-- Facility Command — Phase 2: Breakdown/Issue integrity
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

-- Lets a maintenance/operations user add a note to an already-open breakdown
-- or issue instead of filing a second, duplicate report on the same asset.
alter table breakdown_reports add column if not exists updates jsonb not null default '[]'::jsonb;
alter table issue_reports add column if not exists updates jsonb not null default '[]'::jsonb;
