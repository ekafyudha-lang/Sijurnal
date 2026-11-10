// Si Jurnal - API backend (Supabase Edge Function), pemuat.
// Logika backend (lihat supabase/backend/) disimpan di tabel privat public.app_code dan dimuat di sini,
// sehingga bisa diperbarui tanpa deploy ulang fungsi. Fungsi ini publik (verify_jwt = false):
// otentikasi memakai token sesi aplikasi sendiri, diperiksa oleh setiap fungsi backend.
import postgres from "npm:postgres@3.4.7";
import { createHash, randomUUID } from "node:crypto";
import { Buffer } from "node:buffer";

const sql = postgres(Deno.env.get("SUPABASE_DB_URL")!, { prepare: false, max: 4, idle_timeout: 20, connect_timeout: 10 });
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "content-type, authorization, apikey, x-client-info",
  "Access-Control-Max-Age": "86400",
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store", ...CORS } });
const sha = (text: string) => createHash("sha256").update(text, "utf8").digest("hex");
const build = (code: string) => new Function("__env", code)({ sql, env: (k: string) => Deno.env.get(k), createHash, randomUUID, Buffer });

let handler: ((body: unknown) => Promise<unknown>) | null = null;
let version: string | null = null;

async function getHandler() {
  const v = await sql`select version from public.app_code where id = 1`;
  if (!v.length) throw new Error("Kode backend belum dipasang.");
  if (!handler || version !== v[0].version) {
    const r = await sql`select code, version from public.app_code where id = 1`;
    handler = build(r[0].code);
    version = r[0].version;
  }
  return handler!;
}

// Pasang / perbarui kode backend. Hanya aktif selama properti install_key_hash ada di app_properties.
async function install(secret: unknown, code: unknown) {
  const r = await sql`select value from public.app_properties where key = 'install_key_hash'`;
  if (!r.length || !secret || r[0].value !== sha(String(secret))) return { success: false, message: "Tidak diizinkan." };
  const text = String(code || "");
  build(text); // tolak bila kode tidak bisa dijalankan
  const ver = sha(text);
  await sql`insert into public.app_code (id, code, version, updated_at) values (1, ${text}, ${ver}, now())
            on conflict (id) do update set code = excluded.code, version = excluded.version, updated_at = now()`;
  return { success: true, version: ver, bytes: text.length };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS });
  if (req.method !== "POST") return json({ success: false, message: "Gunakan metode POST." }, 405);
  try {
    const body = await req.json();
    if (body && body.fn === "__install") return json(await install(body.args?.[0], body.args?.[1]));
    return json(await (await getHandler())(body));
  } catch (err) {
    console.error("ERROR: " + ((err as Error)?.stack || err));
    return json({ success: false, message: "Kesalahan server: " + ((err as Error)?.message || err) });
  }
});
