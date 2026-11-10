// Lapisan kompatibilitas: logika bisnis Si Jurnal (app.ts) ditulis untuk layanan Google Apps Script
// (SpreadsheetApp, PropertiesService, CacheService, LockService, DriveApp, Utilities).
// Di sini layanan itu ditiru di atas Postgres + Storage Supabase:
//  - setiap "sheet" = satu tabel; baris dimuat ke memori, diubah di memori, lalu selisihnya ditulis dalam SATU transaksi;
//  - logika berjalan sinkron, jadi bila ia menyentuh data yang belum dimuat, eksekusi diulang setelah data dimuat
//    (aman karena tidak ada yang ditulis ke database sebelum eksekusi selesai tanpa kekurangan data).
import { createHash, randomUUID } from 'node:crypto';
import { Buffer } from 'node:buffer';
import { TABLES } from './schema.js';

const TZ = 'Asia/Jakarta';
class Need extends Error { constructor(m) { super('Data belum dimuat: ' + m); } }

function tzParts(date) {
  const p = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(date).forEach((x) => { p[x.type] = x.value; });
  p.MMM = new Intl.DateTimeFormat('en-US', { timeZone: TZ, month: 'short' }).format(date);
  return p;
}
function formatDate(date, _tz, fmt) {
  const p = tzParts(new Date(date));
  return String(fmt).replace(/yyyy|MMM|MM|dd|HH|mm|ss/g, (t) => ({ yyyy: p.year, MMM: p.MMM, MM: p.month, dd: p.day, HH: p.hour, mm: p.minute, ss: p.second }[t]));
}
const isBlank = (vals) => vals.every((v) => v === '' || v === null || v === undefined);
const same = (a, b) => (a instanceof Date && b instanceof Date) ? a.getTime() === b.getTime() : a === b;
const safeName = (s) => String(s).normalize('NFKD').replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 100) || 'file';

// Nilai memori -> nilai JSON untuk kolom database
function toDb(v, col, sheetName) {
  if (v === undefined || v === null) v = '';
  const where = sheetName + '.' + col.name;
  if (col.type === 'n') {
    if (v === '') return null;
    const n = typeof v === 'number' ? v : Number(String(v).trim());
    if (!isFinite(n)) throw new Error('Nilai "' + v + '" bukan angka (' + where + ').');
    return n;
  }
  if (col.type === 'd') {
    if (v === '') return null;
    if (v instanceof Date) return formatDate(v, TZ, 'yyyy-MM-dd');
    const s = String(v).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s) || isNaN(new Date(s).getTime())) throw new Error('Tanggal "' + v + '" tidak valid, gunakan format yyyy-MM-dd (' + where + ').');
    return s;
  }
  if (col.type === 's') {
    if (v === '') return null;
    return v instanceof Date ? v.toISOString() : String(v);
  }
  return v instanceof Date ? v.toISOString() : String(v);
}

class Range {
  constructor(sheet, r, c, nr, nc) { this.s = sheet; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.s._get(this.r + i, this.c + j)); out.push(row); }
    return out;
  }
  getValue() { return this.s._get(this.r, this.c); }
  setValues(vals) {
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.s._set(this.r + i, this.c + j, vals[i][j]);
    return this;
  }
  setValue(v) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.s._set(this.r + i, this.c + j, v); return this; }
  clearContent() { return this.setValue(''); }
  setNumberFormat() { return this; }
}

class Sheet {
  constructor(name, baseRows) {
    this.name = name;
    this.def = TABLES[name];
    this.headers = this.def.cols.map((c) => c.name);
    this.n = this.headers.length;
    this.rows = baseRows.map((r) => ({ id: r.id, vals: r.vals.slice(), dirty: false }));
    this.deleted = [];
  }
  getName() { return this.name; }
  getLastColumn() { return this.n; }
  getMaxRows() { return 10000000; }
  insertRowsAfter() { return this; }
  setFrozenRows() { return this; }
  getLastRow() { let i = this.rows.length - 1; while (i >= 0 && isBlank(this.rows[i].vals)) i--; return i + 2; }
  getDataRange() { return new Range(this, 1, 1, this.getLastRow(), this.n); }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr === undefined ? 1 : nr, nc === undefined ? 1 : nc); }
  appendRow(arr) { const r = this.getLastRow() + 1; for (let j = 0; j < this.n; j++) this._set(r, j + 1, arr[j]); return this; }
  deleteRow(r) {
    if (r < 2 || r - 2 >= this.rows.length) return this;
    const row = this.rows.splice(r - 2, 1)[0];
    if (row.id !== null) this.deleted.push(row.id);
    return this;
  }
  // Pertahankan hanya baris yang lolos keepFn(nilaiBaris); kembalikan jumlah yang dihapus.
  __filterRows(keepFn) {
    const keep = [];
    let removed = 0;
    this.rows.forEach((row) => {
      if (isBlank(row.vals) || keepFn(row.vals.slice())) keep.push(row);
      else { removed++; if (row.id !== null) this.deleted.push(row.id); }
    });
    this.rows = keep;
    return removed;
  }
  _get(r, c) {
    if (c < 1 || c > this.n) return '';
    if (r === 1) return this.headers[c - 1];
    const row = this.rows[r - 2];
    return row ? row.vals[c - 1] : '';
  }
  _set(r, c, v) {
    if (r < 2 || c < 1 || c > this.n) return;   // baris judul tidak bisa diubah
    if (v === undefined || v === null) v = '';
    while (this.rows.length < r - 1) {
      if (v === '') return;                       // jangan membuat baris kosong
      this.rows.push({ id: null, vals: new Array(this.n).fill(''), dirty: true });
    }
    const row = this.rows[r - 2];
    if (!same(row.vals[c - 1], v)) { row.vals[c - 1] = v; row.dirty = true; }
  }
  changes() {
    const cols = this.def.cols, name = this.name;
    const conv = (vals) => { const o = {}; cols.forEach((col, j) => { o[col.name] = toDb(vals[j], col, name); }); return o; };
    const inserts = [], updates = [], deletes = this.deleted.slice();
    this.rows.forEach((row) => {
      if (row.id === null) { if (!isBlank(row.vals)) inserts.push(conv(row.vals)); }
      else if (isBlank(row.vals)) deletes.push(row.id);
      else if (row.dirty) updates.push(Object.assign({ id: row.id }, conv(row.vals)));
    });
    return (inserts.length || updates.length || deletes.length) ? { name, inserts, updates, deletes } : null;
  }
}

// base = data yang sudah dimuat dari database untuk permintaan ini.
// { sheets: Map(nama -> [{id, vals}]), props: Map, cache: Map, files: Map(id -> {bytes,mime} | null), locked, allLoaded }
class Runtime {
  constructor(base, log) {
    const rt = this;
    this.base = base;
    this.needs = { sheets: new Set(), files: new Set(), lock: false, all: false };
    this.sheets = new Map();
    this.props = new Map(base.props); this.propChanges = new Map();
    this.cache = new Map(base.cache); this.cachePuts = []; this.cacheRemoves = [];
    this.uploads = []; this.trash = [];

    const sheet = (name) => {
      if (!TABLES[name]) return null;
      if (!rt.sheets.has(name)) {
        if (!base.sheets.has(name)) { rt.needs.sheets.add(name); throw new Need(name); }
        rt.sheets.set(name, new Sheet(name, base.sheets.get(name)));
      }
      return rt.sheets.get(name);
    };
    const ss = {
      getSheetByName: sheet,
      insertSheet(name) { throw new Error('Tabel "' + name + '" belum ada di database.'); },
      getName() { return 'Si Jurnal (Supabase)'; },
      getId() { return 'supabase'; },
    };
    const fileObj = (id) => ({
      getId: () => id,
      getUrl: () => 'file:' + id,
      setTrashed() { rt.trash.push(id); return this; },
      getBlob() {
        const up = rt.uploads.filter((u) => u.path === id)[0];
        let f = up || null;
        if (!f) {
          if (!base.files.has(id)) { rt.needs.files.add(id); throw new Need('file ' + id); }
          f = base.files.get(id);
          if (!f) throw new Error('File tidak ditemukan di penyimpanan.');
        }
        return { getBytes: () => f.bytes, getContentType: () => f.mime };
      },
    });
    const folderObj = (id) => ({
      getId: () => id,
      isTrashed: () => false,
      createFile(blob) {
        const path = id + '/' + randomUUID().slice(0, 8) + '-' + safeName(blob.name);
        rt.uploads.push({ path, bytes: blob.bytes, mime: blob.mime || 'application/octet-stream' });
        return fileObj(path);
      },
    });

    this.globals = {
      SpreadsheetApp: {
        getActiveSpreadsheet: () => ss,
        flush() {},
        // Cadangan tutup buku: potret seluruh tabel (sebelum diubah) sebagai file JSON di Storage.
        // File diunggah lebih dulu; bila unggah gagal, transaksi database dibatalkan.
        __createBackup(title) {
          if (!base.allLoaded) { rt.needs.all = true; throw new Need('semua tabel'); }
          const tables = {};
          Object.keys(TABLES).forEach((name) => {
            const sh = sheet(name);
            tables[name] = { columns: sh.headers, rows: sh.rows.filter((r) => !isBlank(r.vals)).map((r) => r.vals.slice()) };
          });
          const props = {}; rt.props.forEach((v, k) => { props[k] = v; });
          const bytes = new Uint8Array(Buffer.from(JSON.stringify({ title, created_at: new Date().toISOString(), tables, properties: props }), 'utf8'));
          const path = 'arsip/' + safeName(title) + '-' + randomUUID().slice(0, 8) + '.json';
          rt.uploads.push({ path, bytes, mime: 'application/json' });
          return { id: path, url: 'backup:' + path };
        },
      },
      PropertiesService: {
        getDocumentProperties: () => ({
          getProperty: (k) => (rt.props.has(k) ? rt.props.get(k) : null),
          setProperty(k, v) { rt.props.set(k, String(v)); rt.propChanges.set(k, String(v)); return this; },
          deleteProperty(k) { rt.props.delete(k); rt.propChanges.set(k, null); return this; },
        }),
      },
      CacheService: {
        getScriptCache: () => ({
          get: (k) => (rt.cache.has(k) ? rt.cache.get(k) : null),
          put(k, v, ttl) { rt.cache.set(k, String(v)); rt.cachePuts.push({ key: k, value: String(v), ttl: Number(ttl) || 600 }); },
          remove(k) { rt.cache.delete(k); rt.cacheRemoves.push(k); },
        }),
      },
      LockService: {
        getScriptLock: () => ({
          tryLock() { if (!base.locked) { rt.needs.lock = true; throw new Need('kunci'); } return true; },
          releaseLock() {},
        }),
      },
      DriveApp: {
        createFolder: (name) => folderObj(/arsip/i.test(name) ? 'arsip' : 'bukti'),
        getFolderById: (id) => folderObj(String(id)),
        getFileById: (id) => fileObj(String(id)),
      },
      Utilities: {
        DigestAlgorithm: { SHA_256: 'sha256' },
        Charset: { UTF_8: 'utf8' },
        getUuid: () => randomUUID(),
        computeDigest: (_alg, text) => Array.from(createHash('sha256').update(String(text), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b)),
        formatDate,
        base64Encode: (bytes) => Buffer.from(bytes).toString('base64'),
        base64Decode: (str) => new Uint8Array(Buffer.from(String(str), 'base64')),
        newBlob: (bytes, mime, name) => ({ bytes, mime, name, getBytes: () => bytes }),
      },
      Session: { getScriptTimeZone: () => TZ },
      Logger: { log: (m) => log && log(String(m)) },
    };
  }
  changes() {
    const sheets = [];
    this.sheets.forEach((sh) => { const c = sh.changes(); if (c) sheets.push(c); });
    return { sheets, props: this.propChanges, cachePuts: this.cachePuts, cacheRemoves: this.cacheRemoves };
  }
}

const HINTS = new Map();   // nama fungsi -> data yang dibutuhkan (dipelajari saat berjalan, agar panggilan berikutnya langsung memuat semuanya)
const RESTART = Symbol('restart');

// adapter: { transaction(lock, fn(tx)), deleteFiles(paths) } ; tx: { loadSheet, loadProps, loadCache, loadFile, flush, upload }
export async function execute(adapter, makeApp, fnName, args, log) {
  let hints = HINTS.get(fnName);
  if (!hints) { hints = { sheets: new Set(), lock: false, all: false }; HINTS.set(fnName, hints); }
  const allNames = Object.keys(TABLES);
  for (let attempt = 0; attempt < 3; attempt++) {
    const out = await adapter.transaction(hints.lock, async (tx) => {
      const base = { sheets: new Map(), files: new Map(), locked: hints.lock, allLoaded: false };
      const load = async (names) => {
        const todo = names.filter((n) => !base.sheets.has(n));
        const res = await Promise.all(todo.map((n) => tx.loadSheet(n)));
        todo.forEach((n, i) => base.sheets.set(n, res[i]));
      };
      [base.props, base.cache] = await Promise.all([tx.loadProps(), tx.loadCache()]);
      await load(hints.all ? allNames : Array.from(hints.sheets));
      base.allLoaded = hints.all;
      for (let i = 0; i < 80; i++) {
        const rt = new Runtime(base, log);
        const app = makeApp(rt.globals);
        let result, err = null;
        try { result = app[fnName].apply(null, args); } catch (e) { err = e; }
        const n = rt.needs;
        if (n.lock && !base.locked) { hints.lock = true; return RESTART; }   // ulangi dalam transaksi baru yang memegang kunci
        if (n.all && !base.allLoaded) { hints.all = true; await load(allNames); base.allLoaded = true; continue; }
        if (n.sheets.size || n.files.size) {
          n.sheets.forEach((s) => hints.sheets.add(s));
          await load(Array.from(n.sheets));
          for (const id of n.files) base.files.set(id, await tx.loadFile(id));
          continue;
        }
        if (err) throw err;
        // Fungsi yang menjawab gagal tidak boleh meninggalkan perubahan setengah jalan: semua perubahannya dibuang.
        if (result && result.success === false) return { result, trash: [] };
        for (const u of rt.uploads) await tx.upload(u);
        await tx.flush(rt.changes());
        return { result, trash: rt.trash };
      }
      throw new Error('Terlalu banyak pengulangan saat memuat data untuk "' + fnName + '".');
    });
    if (out === RESTART) continue;
    if (out.trash.length) { try { await adapter.deleteFiles(out.trash); } catch (e) { log && log('hapus file gagal: ' + e.message); } }
    return out.result;
  }
  throw new Error('Gagal mendapatkan kunci transaksi.');
}
