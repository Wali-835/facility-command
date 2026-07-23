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
    const { to, subject, filename, base64, mimeType, message } = await req.json();
    const resendKey = Deno.env.get("RESEND_API_KEY");

    if (!Array.isArray(to) || to.length === 0) {
      return new Response(JSON.stringify({ error: "No recipients." }), { status: 400, headers: corsHeaders });
    }

    const html = `
      <div style="font-family: Arial, sans-serif; max-width: 600px;">
        <div style="background: #f97316; padding: 20px; border-radius: 8px 8px 0 0;">
          <h2 style="color: white; margin: 0;">📊 Facility Command Report</h2>
        </div>
        <div style="background: #f9fafb; padding: 20px; border-radius: 0 0 8px 8px;">
          <p>${message || "A filtered report is attached."}</p>
          <a href="https://wali-835.github.io/facility-command/" style="background: #f97316; color: white; padding: 10px 20px; border-radius: 6px; text-decoration: none; display: inline-block; margin-top: 10px;">View in App</a>
        </div>
      </div>
    `;

    const emailRes = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${resendKey}` },
      body: JSON.stringify({
        from: "Facility Command <onboarding@resend.dev>",
        to,
        subject,
        html,
        attachments: base64 ? [{ filename, content: base64 }] : undefined,
      }),
    });

    if (!emailRes.ok) {
      const detail = await emailRes.text();
      return new Response(JSON.stringify({ error: `Resend error: ${detail}` }), { status: 502, headers: corsHeaders });
    }

    return new Response("Report sent!", { status: 200, headers: corsHeaders });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), { status: 500, headers: corsHeaders });
  }
});
