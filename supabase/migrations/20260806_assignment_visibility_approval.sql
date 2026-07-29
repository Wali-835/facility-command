-- Facility Command — Assignment-based visibility + supervisor self-assign escalation
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

-- Breakdowns/issues can now be assigned to a specific maintenance person
-- directly (previously only the work order spawned from them could be).
alter table breakdown_reports add column if not exists assignee text;
alter table issue_reports add column if not exists assignee text;

-- Tracks which role actually performed/logged the work, so approval can be
-- routed correctly: a maintenance-submitted log still needs Supervisor+
-- approval, but a Supervisor-submitted log (e.g. a supervisor who assigned
-- a work order to themselves) now needs Engineer+ approval instead, so a
-- supervisor can never approve their own work.
alter table maintenance_logs add column if not exists performer_role text;

notify pgrst, 'reload schema';
