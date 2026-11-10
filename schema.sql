-- Skema database Si Jurnal (Supabase / Postgres).
-- Semua tabel hanya diakses oleh Edge Function `api` lewat koneksi database langsung.
-- RLS aktif tanpa policy: kunci publik (anon/authenticated) tidak bisa membaca atau menulis apa pun.

create table public."users" (
  id bigint generated always as identity primary key,
  "username" text not null default '',
  "password_hash" text not null default '',
  "role" text not null default '',
  "member_id" text not null default ''
);
alter table public."users" enable row level security;
revoke all on public."users" from anon, authenticated;

create table public."coa" (
  id bigint generated always as identity primary key,
  "account_code" text not null default '',
  "account_name" text not null default '',
  "category" text not null default '',
  "normal_balance" text not null default '',
  "status" text not null default '',
  "opening_debit" numeric,
  "opening_credit" numeric
);
alter table public."coa" enable row level security;
revoke all on public."coa" from anon, authenticated;

create table public."journal" (
  id bigint generated always as identity primary key,
  "date" date,
  "ref_no" text not null default '',
  "account_code" text not null default '',
  "account_name" text not null default '',
  "description" text not null default '',
  "debit" numeric,
  "credit" numeric,
  "timestamp" timestamptz,
  "unit" text not null default ''
);
alter table public."journal" enable row level security;
revoke all on public."journal" from anon, authenticated;

create table public."assets" (
  id bigint generated always as identity primary key,
  "asset_id" text not null default '',
  "asset_name" text not null default '',
  "purchase_date" date,
  "cost" numeric,
  "salvage_value" numeric,
  "useful_life_years" numeric,
  "accumulated_dep" numeric,
  "last_dep_date" date
);
alter table public."assets" enable row level security;
revoke all on public."assets" from anon, authenticated;

create table public."members" (
  id bigint generated always as identity primary key,
  "member_id" text not null default '',
  "member_no" text not null default '',
  "name" text not null default '',
  "join_date" date,
  "phone" text not null default '',
  "address" text not null default '',
  "status" text not null default '',
  "exit_date" date
);
alter table public."members" enable row level security;
revoke all on public."members" from anon, authenticated;

create table public."savings" (
  id bigint generated always as identity primary key,
  "tx_id" text not null default '',
  "date" date,
  "member_id" text not null default '',
  "type" text not null default '',
  "direction" text not null default '',
  "amount" numeric,
  "ref_no" text not null default '',
  "note" text not null default '',
  "timestamp" timestamptz
);
alter table public."savings" enable row level security;
revoke all on public."savings" from anon, authenticated;

create table public."member_trans" (
  id bigint generated always as identity primary key,
  "tx_id" text not null default '',
  "date" date,
  "member_id" text not null default '',
  "kind" text not null default '',
  "amount" numeric,
  "note" text not null default '',
  "timestamp" timestamptz
);
alter table public."member_trans" enable row level security;
revoke all on public."member_trans" from anon, authenticated;

create table public."items" (
  id bigint generated always as identity primary key,
  "item_id" text not null default '',
  "code" text not null default '',
  "name" text not null default '',
  "unit" text not null default '',
  "sell_price" numeric,
  "stock" numeric,
  "avg_cost" numeric,
  "status" text not null default '',
  "item_type" text not null default ''
);
alter table public."items" enable row level security;
revoke all on public."items" from anon, authenticated;

create table public."purchases" (
  id bigint generated always as identity primary key,
  "purchase_id" text not null default '',
  "date" date,
  "ref_no" text not null default '',
  "supplier" text not null default '',
  "total" numeric,
  "paid" numeric,
  "cash_account_code" text not null default '',
  "note" text not null default '',
  "timestamp" timestamptz,
  "vehicle_no" text not null default '',
  "driver" text not null default '',
  "kind" text not null default '',
  "unit" text not null default '',
  "supplier_id" text not null default ''
);
alter table public."purchases" enable row level security;
revoke all on public."purchases" from anon, authenticated;

create table public."purchase_lines" (
  id bigint generated always as identity primary key,
  "line_id" text not null default '',
  "purchase_id" text not null default '',
  "item_id" text not null default '',
  "item_name" text not null default '',
  "qty" numeric,
  "unit_cost" numeric,
  "subtotal" numeric,
  "gross" numeric,
  "tare" numeric,
  "cut_pct" numeric,
  "cut_kg" numeric
);
alter table public."purchase_lines" enable row level security;
revoke all on public."purchase_lines" from anon, authenticated;

create table public."sales" (
  id bigint generated always as identity primary key,
  "sale_id" text not null default '',
  "date" date,
  "ref_no" text not null default '',
  "customer" text not null default '',
  "total" numeric,
  "paid" numeric,
  "cost_total" numeric,
  "cash_account_code" text not null default '',
  "note" text not null default '',
  "timestamp" timestamptz,
  "vehicle_no" text not null default '',
  "driver" text not null default '',
  "receivable_account_code" text not null default '',
  "kind" text not null default '',
  "source_purchase_id" text not null default '',
  "unit" text not null default '',
  "tax_ppn" numeric,
  "tax_pph22" numeric,
  "tax_pph23" numeric,
  "tax_ref" text not null default ''
);
alter table public."sales" enable row level security;
revoke all on public."sales" from anon, authenticated;

create table public."sale_lines" (
  id bigint generated always as identity primary key,
  "line_id" text not null default '',
  "sale_id" text not null default '',
  "item_id" text not null default '',
  "item_name" text not null default '',
  "qty" numeric,
  "unit_price" numeric,
  "unit_cost" numeric,
  "subtotal" numeric,
  "gross" numeric,
  "tare" numeric,
  "cut_pct" numeric,
  "cut_kg" numeric
);
alter table public."sale_lines" enable row level security;
revoke all on public."sale_lines" from anon, authenticated;

create table public."suppliers" (
  id bigint generated always as identity primary key,
  "supplier_id" text not null default '',
  "code" text not null default '',
  "name" text not null default '',
  "group" text not null default '',
  "phone" text not null default '',
  "address" text not null default '',
  "village" text not null default '',
  "member_no" text not null default '',
  "status" text not null default '',
  "join_date" date,
  "note" text not null default '',
  "timestamp" timestamptz,
  "unit" text not null default ''
);
alter table public."suppliers" enable row level security;
revoke all on public."suppliers" from anon, authenticated;

create table public."gardens" (
  id bigint generated always as identity primary key,
  "garden_id" text not null default '',
  "supplier_id" text not null default '',
  "name" text not null default '',
  "location" text not null default '',
  "area_ha" numeric,
  "planting_year" numeric,
  "trees" numeric,
  "variety" text not null default '',
  "land_status" text not null default '',
  "lat" numeric,
  "lng" numeric,
  "status" text not null default '',
  "note" text not null default '',
  "timestamp" timestamptz,
  "plantings" text not null default ''
);
alter table public."gardens" enable row level security;
revoke all on public."gardens" from anon, authenticated;

create table public."garden_progress" (
  id bigint generated always as identity primary key,
  "progress_id" text not null default '',
  "garden_id" text not null default '',
  "supplier_id" text not null default '',
  "date" date,
  "stage" text not null default '',
  "progress_pct" numeric,
  "activity" text not null default '',
  "production_kg" numeric,
  "note" text not null default '',
  "by" text not null default '',
  "timestamp" timestamptz
);
alter table public."garden_progress" enable row level security;
revoke all on public."garden_progress" from anon, authenticated;

create table public."trade_payments" (
  id bigint generated always as identity primary key,
  "pay_id" text not null default '',
  "date" date,
  "kind" text not null default '',
  "ref_id" text not null default '',
  "ref_no" text not null default '',
  "amount" numeric,
  "cash_account_code" text not null default '',
  "journal_ref" text not null default '',
  "timestamp" timestamptz
);
alter table public."trade_payments" enable row level security;
revoke all on public."trade_payments" from anon, authenticated;

create table public."evidence" (
  id bigint generated always as identity primary key,
  "att_id" text not null default '',
  "ref_no" text not null default '',
  "file_id" text not null default '',
  "file_name" text not null default '',
  "mime" text not null default '',
  "size" numeric,
  "uploaded_by" text not null default '',
  "timestamp" timestamptz
);
alter table public."evidence" enable row level security;
revoke all on public."evidence" from anon, authenticated;

create table public."closing_log" (
  id bigint generated always as identity primary key,
  "closed_at" timestamptz,
  "cutoff" date,
  "backup_url" text not null default '',
  "backup_id" text not null default '',
  "journal_removed" numeric,
  "trade_removed" numeric,
  "profit" numeric,
  "by" text not null default '',
  "status" text not null default ''
);
alter table public."closing_log" enable row level security;
revoke all on public."closing_log" from anon, authenticated;

create table public.app_properties (key text primary key, value text not null);
alter table public.app_properties enable row level security;
revoke all on public.app_properties from anon, authenticated;

-- Sesi login (pengganti CacheService): token -> pengguna/role, kedaluwarsa otomatis.
create table public.app_cache (key text primary key, value text not null, expires_at timestamptz not null);
alter table public.app_cache enable row level security;
revoke all on public.app_cache from anon, authenticated;
create index app_cache_expires_at_idx on public.app_cache (expires_at);

create index journal_ref_no_idx on public.journal (ref_no);
create index journal_date_idx on public.journal ("date");
create unique index users_username_key on public.users (lower(username));
create unique index coa_account_code_key on public.coa (account_code);

-- Bucket privat untuk bukti transaksi, logo, dan arsip tutup buku.
insert into storage.buckets (id, name, public) values ('si-jurnal', 'si-jurnal', false) on conflict (id) do nothing;

-- Kode backend yang dibaca Edge Function "api" (diisi lewat __install)
create table if not exists public.app_code (
  id int primary key,
  code text not null,
  version text not null,
  updated_at timestamptz not null default now()
);
alter table public.app_code enable row level security;
revoke all on public.app_code from anon, authenticated;
