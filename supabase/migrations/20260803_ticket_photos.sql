-- Facility Command — Ticket photo attachments
-- Lets anyone viewing a ticket (including a Requester) attach photos, both
-- at creation and afterward. Reuses the existing public asset-documents
-- storage bucket under a ticket-photos/ prefix — no new bucket needed.
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

create table if not exists ticket_photos (
  id text primary key,
  ticket_id text not null,
  file_name text not null,
  file_path text not null,
  uploaded_by text,
  uploaded_at timestamptz not null default now()
);
create index if not exists idx_ticket_photos_ticket on ticket_photos(ticket_id);
alter table ticket_photos enable row level security;
drop policy if exists "ticket_photos_all_authenticated" on ticket_photos;
create policy "ticket_photos_all_authenticated" on ticket_photos for all to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
