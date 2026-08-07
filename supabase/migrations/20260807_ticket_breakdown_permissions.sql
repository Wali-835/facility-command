-- Facility Command — Per-user permission to open tickets / report breakdowns
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).
--
-- Previously any logged-in user could report a breakdown/issue or open a
-- ticket. These two flags let an admin restrict either capability per
-- individual user (e.g. a requester who should only file tickets, or an
-- operations user who should only report breakdowns). Both default to
-- true so existing users keep working exactly as before until an admin
-- explicitly narrows someone's access.

alter table user_roles add column if not exists can_report_breakdowns boolean not null default true;
alter table user_roles add column if not exists can_open_tickets boolean not null default true;

notify pgrst, 'reload schema';
