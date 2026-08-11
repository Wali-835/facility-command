// Facility Command — run-monthly-pm Edge Function
//
// Runs daily (see the pg_cron schedule in
// 20260809_pm_automation_cron.sql). On the 1st of the month it generates
// this month's recurring PM work orders for every asset that's due (same
// "due" rule the manual Generate PM Work Orders button uses: no PM yet, or
// months since last_pm_date >= pm_frequency). On the last day of the
// month it auto-closes (status "Missed") any of this month's PM work
// orders that are still open, so a missed PM never silently rolls over —
// it shows up in the monthly report and a fresh one gets generated for
// the new month.
//
// Deploy: supabase functions deploy run-monthly-pm

import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const uid = (p: string) => `${p}-${Date.now().toString(36).toUpperCase()}${Math.random().toString(36).slice(2, 5).toUpperCase()}`;

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  try {
    const supabaseUrl = Deno.env.get("SUPABASE_URL");
    const supabaseKey = Deno.env.get("DB_SERVICE_KEY");
    const dbHeaders = { apikey: supabaseKey!, Authorization: `Bearer ${supabaseKey}`, "Content-Type": "application/json" };

    const now = new Date();
    const y = now.getUTCFullYear(), m = now.getUTCMonth();
    const todayStr = now.toISOString().split("T")[0];
    const currentMonthStr = `${y}-${String(m + 1).padStart(2, "0")}`;
    const lastDayOfMonth = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
    const isFirstDay = now.getUTCDate() === 1;
    const isLastDay = now.getUTCDate() === lastDayOfMonth;

    let closed = 0;
    let generated = 0;

    // Auto-close this month's PM work orders that are still open once the month ends.
    if (isLastDay) {
      const openRes = await fetch(
        `${supabaseUrl}/rest/v1/work_orders?pm_month=eq.${currentMonthStr}&status=neq.Completed&status=neq.Missed&select=id`,
        { headers: dbHeaders }
      );
      const openPmWOs = await openRes.json();
      if (Array.isArray(openPmWOs) && openPmWOs.length) {
        const ids = openPmWOs.map((w: { id: string }) => w.id);
        await fetch(`${supabaseUrl}/rest/v1/work_orders?id=in.(${ids.join(",")})`, {
          method: "PATCH",
          headers: { ...dbHeaders, Prefer: "return=minimal" },
          body: JSON.stringify({ status: "Missed", status_note: `Auto-closed: monthly PM not completed by ${todayStr}.` }),
        });
        closed = ids.length;
      }
    }

    // Generate this month's PM work orders for every asset that's due.
    if (isFirstDay) {
      const assetsRes = await fetch(
        `${supabaseUrl}/rest/v1/assets?select=id,name,category,location,pm_frequency,pm_task,last_pm_date`,
        { headers: dbHeaders }
      );
      const assets = await assetsRes.json();
      const due = (Array.isArray(assets) ? assets : []).filter((a: { pm_frequency: number | null; last_pm_date: string | null }) => {
        if (!a.pm_frequency) return false;
        if (!a.last_pm_date) return true;
        const last = new Date(a.last_pm_date);
        return (y - last.getUTCFullYear()) * 12 + (m - last.getUTCMonth()) >= a.pm_frequency;
      });

      if (due.length) {
        // Don't double-generate for an asset that already has this month's PM WO
        // (e.g. someone used the manual button already, or the cron reran).
        const existingRes = await fetch(
          `${supabaseUrl}/rest/v1/work_orders?pm_month=eq.${currentMonthStr}&select=asset_id`,
          { headers: dbHeaders }
        );
        const existing = await existingRes.json();
        const existingAssetIds = new Set((Array.isArray(existing) ? existing : []).map((w: { asset_id: string }) => w.asset_id));
        const toCreate = due.filter((a: { id: string }) => !existingAssetIds.has(a.id));

        if (toCreate.length) {
          const dueDate = new Date(Date.UTC(y, m + 1, 0)).toISOString().split("T")[0];
          const rows = toCreate.map((a: { id: string; name: string; category: string | null; location: string | null }) => ({
            id: uid("WO"),
            title: `PM - ${a.name}`,
            asset: a.name,
            asset_id: a.id,
            category: a.category || null,
            site: a.location || null,
            priority: "Medium",
            status: "Open",
            assignee: null,
            start_date: todayStr,
            due: dueDate,
            vendor: null,
            pm_month: currentMonthStr,
          }));
          await fetch(`${supabaseUrl}/rest/v1/work_orders`, {
            method: "POST",
            headers: { ...dbHeaders, Prefer: "return=minimal" },
            body: JSON.stringify(rows),
          });
          generated = rows.length;
        }
      }
    }

    return new Response(JSON.stringify({ ok: true, date: todayStr, isFirstDay, isLastDay, generated, closed }), { status: 200, headers: corsHeaders });
  } catch (e) {
    return new Response(JSON.stringify({ error: e instanceof Error ? e.message : "Unexpected error" }), { status: 500, headers: corsHeaders });
  }
});
