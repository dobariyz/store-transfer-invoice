#!/usr/bin/env node
const fs = require('fs');
const path = require('path');

function readJson(p) {
  try {
    return JSON.parse(fs.readFileSync(p, 'utf8'));
  } catch (e) {
    console.error('Failed to read/parse', p, e.message);
    process.exit(1);
  }
}

function csvEscape(v) {
  if (v === null || v === undefined) return '';
  const s = typeof v === 'string' ? v : JSON.stringify(v);
  return '"' + s.replace(/"/g, '""') + '"';
}

function sqlEscape(v) {
  if (v === null || v === undefined) return 'NULL';
  if (typeof v === 'number') return v;
  return ("'" + String(v).replace(/'/g, "''") + "'");
}

const prodFile = process.argv[2] || 'products.json';
const transFile = process.argv[3] || 'transfers.json';
const outDir = process.argv[4] || '.';

const products = readJson(path.resolve(prodFile));
const transfers = readJson(path.resolve(transFile));

// Collect stores
const storesSet = new Set();
transfers.forEach(t => { if (t.fromStore) storesSet.add(t.fromStore); if (t.toStore) storesSet.add(t.toStore); });
const stores = Array.from(storesSet);

// Write CSVs
function writeCSV(filename, header, rows) {
  const out = [header.join(',')];
  for (const r of rows) {
    out.push(r.map(csvEscape).join(','));
  }
  fs.writeFileSync(path.join(outDir, filename), out.join('\n'), 'utf8');
  console.log('Wrote', filename);
}

// products.csv
const prodRows = products.map(p => [p.sku || '', p.name || '', p.category || '', p.subcategory || '', p.unitPrice || '', p.priceType || '', p.packQty || '', p.packPrice || '', JSON.stringify(p.metadata || {}), p.createdAt || '']);
writeCSV('products.csv', ['sku','name','category','subcategory','unit_price','price_type','pack_qty','pack_price','metadata','created_at'], prodRows);

// stores.csv
const storeRows = stores.map(s => [s, s, '{}']);
writeCSV('stores.csv', ['name','display_name','metadata'], storeRows);

// transfers.csv and transfer_items.csv
const transRows = [];
const itemRows = [];
for (const t of transfers) {
  transRows.push([t.id || '', t.date || '', t.fromStore || '', t.toStore || '', t.notes || '', t.subtotal || '', t.hst || '', t.total || '', t.createdAt || '', JSON.stringify(t.reconciledWith || null), JSON.stringify(t.raw || t)]);
  (t.items || []).forEach(it => {
    itemRows.push([t.id || '', it.sku || '', it.name || '', it.category || '', it.qty || 0, it.unitPrice || 0, it.lineTotal || 0, JSON.stringify(it.metadata || {})]);
  });
}
writeCSV('transfers.csv', ['id','date','from_store','to_store','notes','subtotal','hst','total','created_at','reconciled_with','raw'], transRows);
writeCSV('transfer_items.csv', ['transfer_id','sku','name','category','qty','unit_price','line_total','metadata'], itemRows);

// SQL INSERT generation
function writeSQL() {
  const out = [];
  out.push('-- Inserts for products');
  for (const p of products) {
    out.push(`INSERT INTO products (sku,name,category,subcategory,unit_price,price_type,pack_qty,pack_price,metadata,created_at) VALUES (${sqlEscape(p.sku)}, ${sqlEscape(p.name)}, ${sqlEscape(p.category)}, ${sqlEscape(p.subcategory)}, ${p.unitPrice === undefined ? 'NULL' : p.unitPrice}, ${sqlEscape(p.priceType)}, ${p.packQty === undefined ? 'NULL' : p.packQty}, ${p.packPrice === undefined ? 'NULL' : p.packPrice}, ${sqlEscape(JSON.stringify(p.metadata || {}))}, ${sqlEscape(p.createdAt || new Date().toISOString())});`);
  }
  out.push('\n-- Inserts for stores');
  for (const s of stores) {
    out.push(`INSERT INTO stores (name, display_name, metadata) VALUES (${sqlEscape(s)}, ${sqlEscape(s)}, ${sqlEscape('{}')});`);
  }
  out.push('\n-- Inserts for transfers');
  for (const t of transfers) {
    out.push(`INSERT INTO transfers (id,date,from_store,to_store,notes,subtotal,hst,total,created_at,reconciled_with,raw) VALUES (${sqlEscape(t.id)}, ${sqlEscape(t.date)}, ${sqlEscape(t.fromStore)}, ${sqlEscape(t.toStore)}, ${sqlEscape(t.notes)}, ${t.subtotal === undefined ? 'NULL' : t.subtotal}, ${t.hst === undefined ? 'NULL' : t.hst}, ${t.total === undefined ? 'NULL' : t.total}, ${sqlEscape(t.createdAt || new Date().toISOString())}, ${sqlEscape(JSON.stringify(t.reconciledWith || null))}, ${sqlEscape(JSON.stringify(t.raw || t))});`);
  }
  out.push('\n-- Inserts for transfer_items');
  for (const t of transfers) {
    for (const it of (t.items || [])) {
      out.push(`INSERT INTO transfer_items (transfer_id, sku, name, category, qty, unit_price, line_total, metadata) VALUES (${sqlEscape(t.id)}, ${sqlEscape(it.sku)}, ${sqlEscape(it.name)}, ${sqlEscape(it.category)}, ${it.qty === undefined ? 0 : it.qty}, ${it.unitPrice === undefined ? 'NULL' : it.unitPrice}, ${it.lineTotal === undefined ? 'NULL' : it.lineTotal}, ${sqlEscape(JSON.stringify(it.metadata || {}))});`);
    }
  }
  fs.writeFileSync(path.join(outDir, 'inserts.sql'), out.join('\n'), 'utf8');
  console.log('Wrote inserts.sql');
}

writeSQL();

console.log('\nDone. Files generated: products.csv, stores.csv, transfers.csv, transfer_items.csv, inserts.sql');
console.log('Usage: node scripts/export_to_supabase.js [products.json] [transfers.json] [outDir]');
