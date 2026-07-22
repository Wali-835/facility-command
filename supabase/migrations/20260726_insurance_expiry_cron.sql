-- Facility Command — Daily schedule for check-insurance-expiry
-- Run this once in the Supabase SQL Editor. Safe to re-run.
--
-- Your project doesn't show a Cron tab on the Edge Function page, so this
-- sets up the same thing via pg_cron + pg_net directly in Postgres instead.
--
-- BEFORE RUNNING: replace YOUR_ANON_KEY_HERE below with your project's
-- anon/public key. Find it at: Project Settings -> API -> Project API keys
-- -> "anon" / "public" (NOT the service_role key — that one should never
-- leave the Supabase dashboard).

create extension if not exists pg_cron with schema extensions;
create extension if not exists pg_net with schema extensions;

select cron.unschedule('check-insurance-expiry-daily')
where exists (select 1 from cron.job where jobname = 'check-insurance-expiry-daily');

select cron.schedule(
  'check-insurance-expiry-daily',
  '0 6 * * *', -- 6:00 AM UTC every day — adjust the hour to your timezone
  $$
  select net.http_post(
    url := 'https://evwsdzqgvrwbjusjmrdc.supabase.co/functions/v1/check-insurance-expiry',
    headers := jsonb_build_object(
      'Content-Type', 'application/json',
      'Authorization', 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImV2d3NkenFndnJ3Ymp1c2ptcmRjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzgxOTc1OTcsImV4cCI6MjA5Mzc3MzU5N30.Cw2m6A6tUFy-oS0Eg3xkeNCIUWNRM19n_jdCWrTJiDo'
    ),
    body := '{}'::jsonb
  );
  $$
);

-- Verify it's scheduled:
-- select * from cron.job where jobname = 'check-insurance-expiry-daily';
