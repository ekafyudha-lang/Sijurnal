// Struktur tabel Si Jurnal. Urutan kolom HARUS sama dengan urutan kolom "sheet" pada logika asal (app.ts memakai nomor kolom).
// Tipe: t = text, n = numeric, d = date (yyyy-MM-dd), s = timestamptz
function T(table, spec) {
  return { table, cols: spec.trim().split(/\s+/).map((c) => { const [name, type] = c.split(':'); return { name, type: type || 't' }; }) };
}
export const TABLES = {
  Users: T('users', 'username password_hash role member_id'),
  COA: T('coa', 'account_code account_name category normal_balance status opening_debit:n opening_credit:n'),
  Journal: T('journal', 'date:d ref_no account_code account_name description debit:n credit:n timestamp:s unit'),
  Assets: T('assets', 'asset_id asset_name purchase_date:d cost:n salvage_value:n useful_life_years:n accumulated_dep:n last_dep_date:d'),
  Members: T('members', 'member_id member_no name join_date:d phone address status exit_date:d'),
  Savings: T('savings', 'tx_id date:d member_id type direction amount:n ref_no note timestamp:s'),
  MemberTrans: T('member_trans', 'tx_id date:d member_id kind amount:n note timestamp:s'),
  Items: T('items', 'item_id code name unit sell_price:n stock:n avg_cost:n status item_type'),
  Purchases: T('purchases', 'purchase_id date:d ref_no supplier total:n paid:n cash_account_code note timestamp:s vehicle_no driver kind unit supplier_id'),
  PurchaseLines: T('purchase_lines', 'line_id purchase_id item_id item_name qty:n unit_cost:n subtotal:n gross:n tare:n cut_pct:n cut_kg:n'),
  Sales: T('sales', 'sale_id date:d ref_no customer total:n paid:n cost_total:n cash_account_code note timestamp:s vehicle_no driver receivable_account_code kind source_purchase_id unit tax_ppn:n tax_pph22:n tax_pph23:n tax_ref'),
  SaleLines: T('sale_lines', 'line_id sale_id item_id item_name qty:n unit_price:n unit_cost:n subtotal:n gross:n tare:n cut_pct:n cut_kg:n'),
  Suppliers: T('suppliers', 'supplier_id code name group phone address village member_no status join_date:d note timestamp:s unit'),
  Gardens: T('gardens', 'garden_id supplier_id name location area_ha:n planting_year:n trees:n variety land_status lat:n lng:n status note timestamp:s plantings'),
  GardenProgress: T('garden_progress', 'progress_id garden_id supplier_id date:d stage progress_pct:n activity production_kg:n note by timestamp:s'),
  TradePayments: T('trade_payments', 'pay_id date:d kind ref_id ref_no amount:n cash_account_code journal_ref timestamp:s'),
  Evidence: T('evidence', 'att_id ref_no file_id file_name mime size:n uploaded_by timestamp:s'),
  ClosingLog: T('closing_log', 'closed_at:s cutoff:d backup_url backup_id journal_removed:n trade_removed:n profit:n by status'),
};
export const PG_TYPE = { t: 'text', n: 'numeric', d: 'date', s: 'timestamptz' };
