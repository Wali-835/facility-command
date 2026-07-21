-- Facility Command — Work order enhancements
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

-- Action Plan: what's the plan to fix it, and by when.
alter table work_orders add column if not exists action_plan text;
alter table work_orders add column if not exists target_date date;
