exports.handler = async function() {
  const supabaseUrl = process.env.VITE_SUPABASE_URL;
  const supabaseKey = process.env.VITE_SUPABASE_ANON_KEY;
  const resendKey = process.env.RESEND_API_KEY;
  const notifyEmail = process.env.NOTIFY_EMAIL;

  const res = await fetch(
    `${supabaseUrl}/rest/v1/work_orders?select=*&status=neq.Completed`,
    {
      headers: {
        apikey: supabaseKey,
        Authorization: `Bearer ${supabaseKey}`,
      },
    }
  );

  const orders = await res.json();
  const today = new Date().toISOString().split("T")[0];
  const overdue = orders.filter(o => o.due && o.due <= today);

  if (overdue.length === 0) {
    return { statusCode: 200, body: "No overdue orders." };
  }

  const list = overdue
    .map(o => `- ${o.title} (${o.asset}) — Due: ${o.due}`)
    .join("\n");

  await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${resendKey}`,
    },
    body: JSON.stringify({
      from: "Facility Command <onboarding@resend.dev>",
      to: notifyEmail,
      subject: `⚠️ ${overdue.length} Overdue Work Order(s)`,
      text: `The following work orders are overdue:\n\n${list}\n\nPlease log in to take action.\n\nhttps://fmteam.netlify.app`,
    }),
  });

  return { statusCode: 200, body: `Notified for ${overdue.length} overdue orders.` };
};
