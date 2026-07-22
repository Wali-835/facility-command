import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const resendKey = Deno.env.get("RESEND_API_KEY");
    const supabaseUrl = Deno.env.get("SUPABASE_URL");
    const supabaseKey = Deno.env.get("DB_SERVICE_KEY");
    const notifEmail = Deno.env.get("NOTIFICATION_EMAIL") || "ahmedwali835@gmail.com";

    const soon = new Date();
    soon.setDate(soon.getDate() + 30);
    const soonStr = soon.toISOString().split("T")[0];
    const todayStr = new Date().toISOString().split("T")[0];

    const res = await fetch(
      `${supabaseUrl}/rest/v1/insurance_policies?select=*&expiry_date=lte.${soonStr}&expiry_date=not.is.null&order=expiry_date.asc`,
      { headers: { apikey: supabaseKey, Authorization: `Bearer ${supabaseKey}` } }
    );
    const policies = await res.json();

    if (!Array.isArray(policies) || policies.length === 0) {
      return new Response("No policies expiring soon — nothing sent.", { status: 200, headers: corsHeaders });
    }

    const rows = policies.map((p) => {
      const expired = p.expiry_date < todayStr;
      return `<tr>
        <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${p.asset_name || "—"}</td>
        <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${p.provider || "—"} ${p.policy_number ? `(${p.policy_number})` : ""}</td>
        <td style="padding:8px;border-bottom:1px solid #e5e7eb;color:${expired ? "#ef4444" : "#f59e0b"};font-weight:bold;">${p.expiry_date}${expired ? " (EXPIRED)" : ""}</td>
      </tr>`;
    }).join("");

    const subject = `🛡️ Insurance Alert — ${policies.length} polic${policies.length === 1 ? "y" : "ies"} expiring or expired`;
    const html = `
      <div style="font-family: Arial, sans-serif; max-width: 640px;">
        <div style="background: #ef4444; padding: 20px; border-radius: 8px 8px 0 0;">
          <h2 style="color: white; margin: 0;">🛡️ Insurance Expiry Alert</h2>
        </div>
        <div style="background: #f9fafb; padding: 20px; border-radius: 0 0 8px 8px;">
          <p>The following insurance policies are expired or expiring within 30 days:</p>
          <table style="width:100%;border-collapse:collapse;font-size:13px;">
            <thead><tr style="text-align:left;background:#f3f4f6;">
              <th style="padding:8px;">Asset</th><th style="padding:8px;">Provider / Policy #</th><th style="padding:8px;">Expiry</th>
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
          <a href="https://wali-835.github.io/facility-command/" style="background: #ef4444; color: white; padding: 10px 20px; border-radius: 6px; text-decoration: none; display: inline-block; margin-top: 16px;">View in App</a>
        </div>
      </div>
    `;

    await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${resendKey}` },
      body: JSON.stringify({
        from: "Facility Command <onboarding@resend.dev>",
        to: notifEmail,
        subject,
        html,
      }),
    });

    return new Response(`Sent digest for ${policies.length} polic${policies.length === 1 ? "y" : "ies"}.`, { status: 200, headers: corsHeaders });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), { status: 500, headers: corsHeaders });
  }
});
