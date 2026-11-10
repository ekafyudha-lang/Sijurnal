// Menggabungkan logika backend menjadi satu berkas (dist/backend.js) yang dipasang ke tabel public.app_code.
// Jalankan: node build.mjs
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
const dir = new URL('.', import.meta.url);
const strip = (f) => readFileSync(new URL(f, dir), 'utf8').split('\n')
  .filter((l) => !/^import .* from '.*';$/.test(l)).map((l) => l.replace(/^export /, '')).join('\n');
const out = "const { sql, env, createHash, randomUUID, Buffer } = __env;\n" +
  ['schema.js', 'runtime.js', 'app.js', 'adapter.js'].map(strip).join('\n') + "\nreturn makeHandler(sql, env);\n";
new Function('__env', out);   // periksa sintaks
mkdirSync(new URL('dist/', dir), { recursive: true });
writeFileSync(new URL('dist/backend.js', dir), out);
console.log('dist/backend.js', out.length, 'byte');
