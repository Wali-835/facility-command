import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

Deno.serve(async (req) => {
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const { breakdown, type } = await req.json();
    const resendKey = Deno.env.get("RESEND_API_KEY");
    const supabaseUrl = Deno.env.get("SUPABASE_URL");
    const supabaseKey = Deno.env.get("DB_SERVICE_KEY");


    // Get notification recipients
    const res = await fetch(`${supabaseUrl}/rest/v1/notification_settings?select=email`, {
      headers: { apikey: supabaseKey, Authorization: `Bearer ${supabaseKey}` },
    });
    const notifEmail = Deno.env.get("NOTIFICATION_EMAIL") || "ahmedwali835@gmail.com";
const recipients = [{ email: notifEmail }];


    const isResolved = type === "resolved";
    const subject = isResolved
      ? `✅ Breakdown Resolved — ${breakdown.asset_name}`
      : `🚨 New Breakdown — ${breakdown.asset_name} (${breakdown.severity})`;

    const html = isResolved ? `
      <div style="font-family: Arial, sans-serif; max-width: 600px;">
        <div style="background: #22c55e; padding: 20px; border-radius: 8px 8px 0 0;">
          <h2 style="color: white; margin: 0;">✅ Breakdown Resolved</h2>
        </div>
        <div style="background: #f9fafb; padding: 20px; border-radius: 0 0 8px 8px;">
          <p><strong>Equipment:</strong> ${breakdown.asset_name}</p>
          <p><strong>Site:</strong> ${breakdown.site}</p>
          <p><strong>Resolved by:</strong> ${breakdown.resolved_by}</p>
          <p><strong>Notes:</strong> ${breakdown.maintenance_notes || "—"}</p>
          <a href="https://wali-835.github.io/facility-command/" style="background: #22c55e; color: white; padding: 10px 20px; border-radius: 6px; text-decoration: none; display: inline-block; margin-top: 10px;">View in App</a>
        </div>
      </div>
    ` : `
      <div style="font-family: Arial, sans-serif; max-width: 600px;">
        <div style="background: #ef4444; padding: 20px; border-radius: 8px 8px 0 0;">
          <h2 style="color: white; margin: 0;">🚨 New Breakdown Reported</h2>
        </div>
        <div style="background: #f9fafb; padding: 20px; border-radius: 0 0 8px 8px;">
          <p><strong>Equipment:</strong> ${breakdown.asset_name}</p>
          <p><strong>Site:</strong> ${breakdown.site}</p>
          <p><strong>Reported by:</strong> ${breakdown.reported_by}</p>
          <p><strong>Severity:</strong> ${breakdown.severity}</p>
          <p><strong>Issue:</strong> ${breakdown.description}</p>
          <p><strong>Time:</strong> ${new Date(breakdown.reported_at).toLocaleString("en-GB")}</p>
          <a href="https://wali-835.github.io/facility-command/" style="background: #ef4444; color: white; padding: 10px 20px; border-radius: 6px; text-decoration: none; display: inline-block; margin-top: 10px;">View in App</a>
        </div>
      </div>
    `;

    await Promise.all(recipients.map(r =>
      fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${resendKey}` },
        body: JSON.stringify({
          from: "Facility Command <onboarding@resend.dev>",
          to: r.email,
          subject,
          html,
        }),
      })
    ));

    return new Response("Notifications sent!", { status: 200, headers: corsHeaders });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), { status: 500, headers: corsHeaders });
  }
});
