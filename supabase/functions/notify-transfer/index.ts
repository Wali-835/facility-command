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
    const { asset, oldLocation, newLocation, changedBy } = await req.json();
    const resendKey = Deno.env.get("RESEND_API_KEY");
    const notifEmail = Deno.env.get("NOTIFICATION_EMAIL") || "ahmedwali835@gmail.com";
    const recipients = [{ email: notifEmail }];

    const subject = `🔄 Asset Transfer — ${asset.name} (${oldLocation || "—"} → ${newLocation || "—"})`;
    const html = `
      <div style="font-family: Arial, sans-serif; max-width: 600px;">
        <div style="background: #f97316; padding: 20px; border-radius: 8px 8px 0 0;">
          <h2 style="color: white; margin: 0;">🔄 Asset Transfer / Security Exit</h2>
        </div>
        <div style="background: #f9fafb; padding: 20px; border-radius: 0 0 8px 8px;">
          <p><strong>Asset:</strong> ${asset.name} (${asset.id})</p>
          <p><strong>Category:</strong> ${asset.category || "—"}</p>
          <p><strong>From Site:</strong> ${oldLocation || "—"}</p>
          <p><strong>To Site:</strong> ${newLocation || "—"}</p>
          <p><strong>Authorized by:</strong> ${changedBy || "—"}</p>
          <p><strong>Time:</strong> ${new Date().toLocaleString("en-GB")}</p>
          <a href="https://wali-835.github.io/facility-command/" style="background: #f97316; color: white; padding: 10px 20px; border-radius: 6px; text-decoration: none; display: inline-block; margin-top: 10px;">View in App</a>
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

    return new Response("Notification sent!", { status: 200, headers: corsHeaders });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), { status: 500, headers: corsHeaders });
  }
});
