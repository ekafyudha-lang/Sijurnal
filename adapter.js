// Adapter Postgres + Storage Supabase dan penangan permintaan API Si Jurnal.
import { TABLES, PG_TYPE } from './schema.js';
import { execute } from './runtime.js';
import { createApp } from './app.js';

// sql = klien postgres.js ; env(nama) = pembaca variabel lingkungan
export function makeHandler(sql, env) {
const API_FUNCTIONS = new Set([
  'bulkSave', 'calculateShu', 'cancelTradePayment', 'changeMyPassword', 'deleteAccount', 'deleteEvidence',
  'deleteGarden', 'deleteGardenProgress', 'deleteItem', 'deleteJournalEntry', 'deleteMember', 'deleteMemberAccount',
  'deleteMemberTrans', 'deletePurchase', 'deleteRole', 'deleteSale', 'deleteSaving', 'deleteStaffUser',
  'deleteSupplier', 'getAssets', 'getCOAList', 'getCashFlow', 'getClosingInfo', 'getCompanySettings',
  'getComparativeReport', 'getCoopProfile', 'getDashboardSummary', 'getEquityChanges', 'getEvidenceFile', 'getEvidenceIndex',
  'getEvidenceList', 'getFinancialStatements', 'getGardenProgress', 'getGardens', 'getItems', 'getJournals',
  'getLedger', 'getMemberTrans', 'getMembers', 'getMyPerms', 'getMyPortal', 'getPurchases',
  'getRoles', 'getSales', 'getSavings', 'getShuSettings', 'getStaffUsers', 'getSupplierNames',
  'getSupplierOptions', 'getSuppliers', 'getTradePayments', 'getYieldSettings', 'importCoaCSV', 'importCoaExcel',
  'importMembers', 'loginUser', 'payTrade', 'postDepreciationJournal', 'postInvoiceTax', 'previewClosing',
  'renameAccount', 'runClosing', 'saveAccount', 'saveAsset', 'saveCompanySettings', 'saveCoopProfile',
  'saveGarden', 'saveGardenProgress', 'saveItem', 'saveJournal', 'saveMember', 'saveMemberAccount',
  'saveMemberTrans', 'savePurchase', 'saveRole', 'saveSale', 'saveSaving', 'saveShuSettings',
  'saveStaffUser', 'saveSupplier', 'saveYieldCurve', 'setupCoopCoa', 'updateSaldoAwalBulk', 'uploadEvidence',
]);

const BUCKET = 'si-jurnal';
const LOCK_KEY = 742026111;

const q = (s) => '"' + s + '"';
const SELECT = {}, RECORD = {};
for (const [name, def] of Object.entries(TABLES)) {
  SELECT[name] = 'select id, ' + def.cols.map((c) => c.type === 'n' ? q(c.name) + '::float8 as ' + q(c.name) : c.type === 'd' ? q(c.name) + '::text as ' + q(c.name) : q(c.name)).join(', ') +
    ' from public.' + q(def.table) + ' order by id';
  RECORD[name] = def.cols.map((c) => q(c.name) + ' ' + PG_TYPE[c.type]).join(', ');
}

// ---------- Storage ----------
function storageHeaders(extra) {
  let key = '';
  try { key = JSON.parse(env('SUPABASE_SECRET_KEYS') || '{}')['default'] || ''; } catch (_e) { /* pakai kunci lama */ }
  const h = Object.assign({}, extra || {});
  if (key) h.apikey = key;
  else { key = env('SUPABASE_SERVICE_ROLE_KEY') || ''; h.apikey = key; h.Authorization = 'Bearer ' + key; }
  return h;
}
const objUrl = (path) => env('SUPABASE_URL') + '/storage/v1/object/' + BUCKET + '/' + path.split('/').map(encodeURIComponent).join('/');
async function storageUpload(u) {
  const res = await fetch(objUrl(u.path), { method: 'POST', headers: storageHeaders({ 'Content-Type': u.mime, 'x-upsert': 'true' }), body: u.bytes });
  if (!res.ok) throw new Error('Gagal menyimpan file ke Storage (' + res.status + '): ' + (await res.text()).slice(0, 200));
  await res.body?.cancel();
}
async function storageDownload(path) {
  const res = await fetch(objUrl(path), { headers: storageHeaders() });
  if (res.status === 404 || res.status === 400) { await res.body?.cancel(); return null; }
  if (!res.ok) throw new Error('Gagal membaca file dari Storage (' + res.status + ').');
  return { bytes: new Uint8Array(await res.arrayBuffer()), mime: res.headers.get('content-type') || 'application/octet-stream' };
}
async function storageDelete(paths) {
  for (const p of paths) { const res = await fetch(objUrl(p), { method: 'DELETE', headers: storageHeaders() }); await res.body?.cancel(); }
}
async function storageSignedUrl(path, seconds) {
  const res = await fetch(env('SUPABASE_URL') + '/storage/v1/object/sign/' + BUCKET + '/' + path.split('/').map(encodeURIComponent).join('/'),
    { method: 'POST', headers: storageHeaders({ 'Content-Type': 'application/json' }), body: JSON.stringify({ expiresIn: seconds }) });
  if (!res.ok) throw new Error('Gagal membuat tautan unduh (' + res.status + ').');
  const j = await res.json();
  return env('SUPABASE_URL') + '/storage/v1' + (j.signedURL || j.signedUrl);
}

// ---------- Adapter Postgres ----------
function txApi(t) {
  return {
    async loadSheet(name) {
      const def = TABLES[name];
      const rows = await t.unsafe(SELECT[name]);
      return rows.map((r) => ({ id: Number(r.id), vals: def.cols.map((c) => { const v = r[c.name]; return v === null || v === undefined ? '' : v; }) }));
    },
    async loadProps() { const m = new Map(); (await t`select key, value from public.app_properties`).forEach((r) => m.set(r.key, r.value)); return m; },
    async loadCache() { const m = new Map(); (await t`select key, value from public.app_cache where expires_at > now()`).forEach((r) => m.set(r.key, r.value)); return m; },
    loadFile: storageDownload,
    upload: storageUpload,
    async flush(ch) {
      for (const s of ch.sheets) {
        const def = TABLES[s.name], tbl = 'public.' + q(def.table), cols = def.cols.map((c) => q(c.name));
        if (s.deletes.length) await t.unsafe('delete from ' + tbl + " where id = any(string_to_array($1::text, ',')::bigint[])", [s.deletes.join(',')]);
        if (s.updates.length) {
          await t.unsafe('update ' + tbl + ' as t set ' + cols.map((c) => c + ' = x.' + c).join(', ') +
            ' from jsonb_to_recordset($1::text::jsonb) as x(id bigint, ' + RECORD[s.name] + ') where t.id = x.id', [JSON.stringify(s.updates)]);
        }
        if (s.inserts.length) {
          await t.unsafe('insert into ' + tbl + ' (' + cols.join(', ') + ') select ' + cols.map((c) => 'x.' + c).join(', ') +
            ' from jsonb_array_elements($1::text::jsonb) with ordinality as e(obj, ord) cross join lateral jsonb_to_record(e.obj) as x(' + RECORD[s.name] + ') order by e.ord', [JSON.stringify(s.inserts)]);
        }
      }
      for (const [k, v] of ch.props) {
        if (v === null) await t`delete from public.app_properties where key = ${k}`;
        else await t`insert into public.app_properties (key, value) values (${k}, ${v}) on conflict (key) do update set value = excluded.value`;
      }
      for (const c of ch.cachePuts) {
        await t`insert into public.app_cache (key, value, expires_at) values (${c.key}, ${c.value}, now() + ${Math.round(c.ttl)}::int * interval '1 second')
                on conflict (key) do update set value = excluded.value, expires_at = excluded.expires_at`;
      }
      for (const k of ch.cacheRemoves) await t`delete from public.app_cache where key = ${k}`;
      if (ch.cachePuts.length) await t`delete from public.app_cache where expires_at < now()`;
    },
  };
}
const adapter = {
  // Fungsi yang menulis memegang kunci transaksi: semua penulisan berjalan bergantian, seperti LockService di versi lama.
  transaction: (lock, fn) => sql.begin(async (t) => { if (lock) await t`select pg_advisory_xact_lock(${LOCK_KEY})`; return fn(txApi(t)); }),
  deleteFiles: storageDelete,
};


  // body = { fn, args } -> hasil fungsi (objek JSON)
  return async function handle(body) {
    const fn = String((body && body.fn) || '');
    const args = Array.isArray(body && body.args) ? body.args : [];
    const log = (m) => console.log('[' + fn + '] ' + m);
    try {
      // Tautan unduh cadangan tutup buku (khusus Administrator; berlaku 5 menit)
      if (fn === 'getBackupLink') {
        const info = await execute(adapter, createApp, 'getClosingInfo', [args[0]], log);
        if (!info || !info.success) return info;
        const path = String(args[1] || '');
        if (!/^arsip\/[A-Za-z0-9._-]+\.json$/.test(path)) return { success: false, message: 'Cadangan tidak valid.' };
        return { success: true, url: await storageSignedUrl(path, 300) };
      }
      if (!API_FUNCTIONS.has(fn)) return { success: false, message: 'Fungsi "' + fn + '" tidak tersedia di backend.' };
      const out = await execute(adapter, createApp, fn, args, log);
      return out === undefined ? null : out;
    } catch (err) {
      console.error('[' + fn + '] ERROR: ' + ((err && err.stack) || err));
      return { success: false, message: 'Kesalahan server: ' + ((err && err.message) || err) };
    }
  };
}
