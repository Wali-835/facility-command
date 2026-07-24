// Facility Command — create-user Edge Function
//
// Provisions a real Supabase Auth login (email or phone-derived synthetic
// email) plus the matching user_roles row. This has to run server-side
// with the service-role key — that key must never be shipped to the
// browser, so account creation can't happen directly from the React app.
//
// Deploy: supabase functions deploy create-user
// Called by: UserManagement in src/App.tsx, authenticated as an admin.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY")!;

const digitsOnly = (s: string) => (s || "").replace(/\D/g, "");
const uid = (p: string) => `${p}-${Date.now().toString(36).toUpperCase()}${Math.random().toString(36).slice(2, 5).toUpperCase()}`;

function json(body: unknown, status: number) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
  });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS_HEADERS });
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);

  try {
    const callerToken = (req.headers.get("Authorization") || "").replace("Bearer ", "");
    if (!callerToken) return json({ error: "Missing Authorization header" }, 401);

    // Verify the caller with their own JWT (anon client), then use the
    // service-role client only for the privileged operations below.
    const callerClient = createClient(SUPABASE_URL, ANON_KEY, {
      global: { headers: { Authorization: `Bearer ${callerToken}` } },
    });
    const { data: callerData, error: callerErr } = await callerClient.auth.getUser();
    if (callerErr || !callerData?.user?.email) return json({ error: "Invalid session" }, 401);

    const admin = createClient(SUPABASE_URL, SERVICE_ROLE_KEY);

    const { data: callerRole } = await admin.from("user_roles").select("role").eq("email", callerData.user.email).single();
    if (callerRole?.role !== "admin") return json({ error: "Admin access required" }, 403);

    const body = await req.json();
    const { email, phone, password, name, role, site, department, supervised_sites, supervised_categories, language } = body || {};

    if (!name) return json({ error: "Name is required" }, 400);
    if (!password || password.length < 6) return json({ error: "Password must be at least 6 characters" }, 400);
    if (!email && !phone) return json({ error: "Email or phone is required" }, 400);

    const phoneDigits = digitsOnly(phone || "");
    if (!email && phoneDigits.length < 6) return json({ error: "Enter a valid phone number" }, 400);
    const loginEmail = email || `${phoneDigits}@facility-command.local`;

    const { error: createErr } = await admin.auth.admin.createUser({
      email: loginEmail,
      password,
      email_confirm: true,
    });
    if (createErr) return json({ error: createErr.message }, 400);

    const { error: upsertErr } = await admin.from("user_roles").upsert([{
      id: uid("USR"),
      email: loginEmail,
      phone: phone || null,
      name,
      role: role || "operations",
      site: site || null,
      department: department || null,
      supervised_sites: supervised_sites?.length ? supervised_sites : null,
      supervised_categories: supervised_categories?.length ? supervised_categories : null,
      language: language || "en",
    }], { onConflict: "email" });
    if (upsertErr) return json({ error: upsertErr.message }, 400);

    return json({ success: true, email: loginEmail }, 200);
  } catch (e) {
    return json({ error: e instanceof Error ? e.message : "Unexpected error" }, 500);
  }
});
