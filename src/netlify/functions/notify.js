const nodemailer = require("nodemailer");

exports.handler = async function () {
  const supabaseUrl = process.env.VITE_SUPABASE_URL;
  const supabaseKey = process.env.VITE_SUPABASE_ANON_KEY;

  const res = await fetch(`${supabaseUrl}/rest/v1/work_orders?select=*&status=neq.Completed`, {
    headers: {
      apikey: supabaseKey,
      Authorization: `Bearer ${supabaseKey}`,
    },
  });

  const orders = await res.json();
  const today = new Date().toISOString().split("T")[0];
  const overdue = orders.filter(o => o.due && o.due <= today);

  if (overdue.length === 0) {
    return { statusCode: 200, body: "No overdue orders." };
  }

  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: process.env.GMAIL_USER,
      pass: process.env.GMAIL_PASS,
    },
  });

  const list = overdue.map(o => `- ${o.title} (${o.asset}) — Due: ${o.due}`).join("\n");

  await transporter.sendMail({
    from: process.env.GMAIL_USER,
    to: process.env.NOTIFY_EMAIL,
    subject: `⚠️ ${overdue.length} Overdue Work Order(s) — Facility Command`,
    text: `The following work orders are overdue:\n\n${list}\n\nPlease log in to take action.`,
  });

  return { statusCode: 200, body: `Sent notification for ${overdue.length} overdue orders.` };
};
