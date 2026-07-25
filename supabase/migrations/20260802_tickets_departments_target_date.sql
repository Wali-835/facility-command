-- Facility Command — Tickets: multi-department assignment + target date
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

alter table tickets add column if not exists departments text[];
alter table tickets add column if not exists target_date date;

notify pgrst, 'reload schema';
