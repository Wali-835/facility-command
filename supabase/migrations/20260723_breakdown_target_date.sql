-- Facility Command — Target resolution date on breakdowns
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

alter table breakdown_reports add column if not exists target_date date;

notify pgrst, 'reload schema';
