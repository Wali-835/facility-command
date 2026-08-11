-- Facility Command — Daily schedule for run-monthly-pm
-- Run this once in the Supabase SQL Editor. Safe to re-run.
--
-- Same pg_cron + pg_net approach as 20260726_insurance_expiry_cron.sql.
-- The function itself decides what to do each day: on the 1st of the
-- month it generates that month's PM work orders for every due asset; on
-- the last day of the month it auto-closes (status "Missed") any of that
-- month's PM work orders still open. Every other day it's a no-op.
--
-- Uses the same project anon key as 20260726_insurance_expiry_cron.sql
-- (Project Settings -> API -> Project API keys -> "anon" / "public" —
-- NOT the service_role key, which must never leave the dashboard).

create extension if not exists pg_cron with schema extensions;
create extension if not exists pg_net with schema extensions;

select cron.unschedule('run-monthly-pm-daily')
where exists (select 1 from cron.job where jobname = 'run-monthly-pm-daily');

select cron.schedule(
  'run-monthly-pm-daily',
  '0 2 * * *', -- 2:00 AM UTC every day — adjust the hour to your timezone
  $$
  select net.http_post(
    url := 'https://evwsdzqgvrwbjusjmrdc.supabase.co/functions/v1/run-monthly-pm',
    headers := jsonb_build_object(
      'Content-Type', 'application/json',
      'Authorization', 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImV2d3NkenFndnJ3Ymp1c2ptcmRjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzgxOTc1OTcsImV4cCI6MjA5Mzc3MzU5N30.Cw2m6A6tUFy-oS0Eg3xkeNCIUWNRM19n_jdCWrTJiDo'
    ),
    body := '{}'::jsonb
  );
  $$
);

-- Verify it's scheduled:
-- select * from cron.job where jobname = 'run-monthly-pm-daily';
