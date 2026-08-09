-- Facility Command — Ticket hierarchy (main ticket + sub-tickets)
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).
--
-- An Engineer or Admin can split a main ticket into sub-tickets, each
-- routed to its own department(s)/technician/vendor. The main ticket
-- stays open until every sub-ticket is Resolved/Closed, at which point
-- the original requester (or an admin) gets the authority to close it.

alter table tickets add column if not exists parent_ticket_id text;
create index if not exists idx_tickets_parent on tickets(parent_ticket_id);

notify pgrst, 'reload schema';
