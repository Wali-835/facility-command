-- Facility Command — Automated monthly PM work order generation + auto-close
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).
--
-- pm_month tags a work order as an auto-generated recurring PM task for a
-- given calendar month ('YYYY-MM'), so the scheduled job can (a) tell which
-- work orders are its own recurring PM tasks vs. everything else, and
-- (b) never generate two for the same asset in the same month. Manually
-- created work orders (including the existing "Generate PM Work Orders"
-- button) leave this null.

alter table work_orders add column if not exists pm_month text;
create unique index if not exists idx_wo_pm_asset_month on work_orders(asset_id, pm_month) where pm_month is not null;

notify pgrst, 'reload schema';
