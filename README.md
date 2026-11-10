# Si Jurnal — Firebase Hosting + Supabase

- Frontend: `public/index.html` (Firebase Hosting). Backend: Supabase Edge Function `api` (`supabase/functions/api/index.ts`).
- Database: Postgres project `si-jurnal` (ref `dbncsklfqcnctzozhqho`); skema di `supabase/schema.sql`. Berkas bukti & cadangan tutup buku di Storage bucket privat `si-jurnal`.
- Logika bisnis ada di `supabase/backend/*.js`, digabung jadi satu bundel: `node supabase/backend/build.mjs` -> `dist/backend.js`.

## Deploy frontend
1. Isi project ID di `.firebaserc`.
2. `npm i -g firebase-tools && firebase login`
3. `firebase deploy --only hosting`

## Memperbarui backend
1. `node supabase/backend/build.mjs`
2. Aktifkan sementara: `insert into app_properties(key,value) values ('install_key_hash', '<sha256 dari kunci rahasia>');`
3. POST `{"fn":"__install","args":["<kunci rahasia>","<isi dist/backend.js>"]}` ke URL fungsi.
4. Hapus baris `install_key_hash` lagi.

Login awal: `admin` — segera ganti kata sandinya.
