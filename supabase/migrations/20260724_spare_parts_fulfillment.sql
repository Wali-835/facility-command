-- Facility Command — Spare parts: single-stage approval + stock/order fulfillment
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).
--
-- Spare parts approval is now single-stage: maintenance or supervisor
-- requests a part (from catalog/stock, or a brand new uncatalogued part),
-- and an engineer or admin gives the final approval — choosing whether it's
-- fulfilled from existing stock (deducts stock_quantity) or needs to be
-- ordered (no stock to deduct yet). The two-stage "Supervisor Approved"
-- columns from an earlier migration are no longer used by the app but are
-- left in place — harmless if already applied.

alter table spare_parts add column if not exists fulfillment_type text; -- 'Stock' | 'Order'

notify pgrst, 'reload schema';
