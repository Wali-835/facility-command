-- Facility Command — Phase 3: Auth & permissions
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

-- Lets technicians without email sign in by phone number. Their actual
-- Supabase Auth login is a synthetic email derived from the phone digits
-- (see the create-user Edge Function) — this column is what the UI shows
-- and matches against, not the real login identifier.
alter table user_roles add column if not exists phone text;

create unique index if not exists user_roles_phone_unique
  on user_roles (phone)
  where phone is not null and phone <> '';
