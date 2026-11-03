/**
 * VFW Post Bar Inventory Scanner backend (Google Apps Script)
 * Step-by-step setup: https://vfw9509.github.io/inventory/setup.html
 *
 * Paste this whole file into Extensions > Apps Script of the inventory Google Sheet.
 * 1. Set the two PINs below (the setup page fills them in for you).
 * 2. Run `setup` once (Run button, pick "setup", approve permissions).
 * 3. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 *
 * Staff PIN: scan, receive, remove, count, add new items.
 * Manager PIN: everything, plus edit par/cost/POS names, delete items, import sales.
 * To change a PIN later: edit it here, Save, then Deploy > Manage deployments > edit > New version > Deploy.
 */

const STAFF_PIN = 'SET-STAFF-PIN';      // bartenders
const MANAGER_PIN = 'SET-MANAGER-PIN';  // quartermaster / bar manager

const ITEMS_SHEET = 'Items';
const LOG_SHEET = 'Log';

// Columns the script writes. Two formula columns ("To Order", "Value") follow these.
const COLS = ['ID', 'Name', 'Category', 'Unit', 'UPC', 'Case UPC', 'Case Qty', 'On Hand', 'Par',
  'Cost', 'POS Names', 'Pours Per Unit', 'Last Counted', 'Updated', 'Updated By'];
const KEYS = ['id', 'name', 'category', 'unit', 'upc', 'caseUpc', 'caseQty', 'onHand', 'par',
  'cost', 'posNames', 'poursPerUnit', 'lastCounted', 'updated', 'updatedBy'];
const LOG_COLS = ['Timestamp', 'User', 'Action', 'Item ID', 'Item', 'Change', 'On Hand After', 'Note', 'Op ID'];

// Fields a staff PIN may fill on an item that already exists (only when blank).
const STAFF_LINK_FIELDS = ['upc', 'caseUpc', 'caseQty'];

/* ---------------------------------------------------------------- setup */

function setup() {
  const ss = SpreadsheetApp.getActive();
  let items = ss.getSheetByName(ITEMS_SHEET) || ss.insertSheet(ITEMS_SHEET, 0);
  if (items.getLastRow() === 0) {
    items.getRange(1, 1, 1, COLS.length).setValues([COLS]);
  }
  const n = COLS.length;
  const col = c => colLetter(COLS.indexOf(c) + 1);
  const oh = col('On Hand'), par = col('Par'), cost = col('Cost');
  items.getRange(1, n + 1).setFormula(
    `={"To Order";ARRAYFORMULA(IF(A2:A="","",IF(${par}2:${par}="","",IF(${par}2:${par}-${oh}2:${oh}>0,${par}2:${par}-${oh}2:${oh},0))))}`);
  items.getRange(1, n + 2).setFormula(
    `={"Value";ARRAYFORMULA(IF(A2:A="","",N(${oh}2:${oh})*N(${cost}2:${cost})))}`);
  // Keep IDs and barcodes as text so leading zeros survive.
  ['ID', 'UPC', 'Case UPC'].forEach(c => items.getRange(`${col(c)}:${col(c)}`).setNumberFormat('@'));
  items.getRange(`${cost}:${cost}`).setNumberFormat('$0.00');
  items.getRange(`${colLetter(n + 2)}:${colLetter(n + 2)}`).setNumberFormat('$0.00');
  items.getRange(1, 1, 1, n + 2).setFontWeight('bold').setBackground('#1f3a68').setFontColor('#ffffff');
  items.setFrozenRows(1);
  items.setFrozenColumns(2);

  // Highlight rows that are below par.
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($A2<>"",$${par}2<>"",$${oh}2<$${par}2)`)
    .setBackground('#fde2e1')
    .setRanges([items.getRange(2, 1, items.getMaxRows() - 1, n + 2)])
    .build();
  items.setConditionalFormatRules([rule]);

  let log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, LOG_COLS.length).setValues([LOG_COLS])
      .setFontWeight('bold').setBackground('#1f3a68').setFontColor('#ffffff');
    log.setFrozenRows(1);
  }
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(blank);
}

/* ---------------------------------------------------------------- web entry points */

function doGet(e) {
  return handle_(e.parameter || {});
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  return handle_(req);
}

function handle_(req) {
  try {
    const role = roleFor_(String(req.pin || ''));
    if (!role) return json_({ ok: false, error: 'Wrong PIN', code: 'pin' });
    const user = String(req.user || '').slice(0, 40) || role;
    switch (req.action) {
      case 'ping': return json_({ ok: true, role });
      case 'items': return json_({ ok: true, role, items: readItems_().items });
      case 'log': return json_({ ok: true, entries: readLog_(Number(req.limit) || 100, req.itemId) });
      case 'lookup': return json_({ ok: true, product: lookupUpc_(String(req.upc || '')) });
      case 'saveItem': return withLock_(() => json_(saveItem_(req.item || {}, role, user)));
      case 'importItems': return withLock_(() => json_(importItems_(req.items || [], role, user)));
      case 'deleteItem': return withLock_(() => json_(deleteItem_(String(req.id || ''), role, user)));
      case 'adjust': return withLock_(() => json_(runOps_([req], role, user, null)));
      case 'batch': return withLock_(() => json_(runOps_(req.ops || [], role, user, req)));
      default: return json_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function roleFor_(pin) {
  if (/^SET-/.test(STAFF_PIN) || /^SET-/.test(MANAGER_PIN) || STAFF_PIN === MANAGER_PIN) {
    throw new Error('PINs are not set up. Put a staff PIN and a different manager PIN at the top of the script, then deploy a new version.');
  }
  if (pin && pin === MANAGER_PIN) return 'manager';
  if (pin && pin === STAFF_PIN) return 'staff';
  return null;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------------------------------------------------------- items */

function itemsSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ITEMS_SHEET);
  if (!sh) throw new Error('Items sheet missing: run setup in Apps Script');
  return sh;
}

// The formula columns fill every row, so getLastRow() can't be trusted on Items: look at column A.
function lastItemRow_(sh) {
  const ids = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
  for (let r = ids.length - 1; r >= 0; r--) if (ids[r][0] !== '') return r + 1;
  return 1;
}

// Write rows after `afterRow`, growing the sheet first if needed.
function writeRows_(sh, afterRow, rows) {
  if (!rows.length) return;
  const need = afterRow + rows.length - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need + 100);
  sh.getRange(afterRow + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function readItems_() {
  const sh = itemsSheet_();
  const last = lastItemRow_(sh);
  const header = sh.getRange(1, 1, 1, COLS.length).getValues()[0];
  COLS.forEach((c, i) => {
    if (header[i] !== c) throw new Error(`Items sheet column ${colLetter(i + 1)} should be "${c}" — don't move columns`);
  });
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, COLS.length).getValues() : [];
  const items = [];
  const rowOf = {};
  rows.forEach((r, i) => {
    if (!r[0]) return;
    const it = rowToItem_(r);
    items.push(it);
    rowOf[it.id] = i + 2;
  });
  return { sh, items, rowOf, last };
}

function rowToItem_(r) {
  const it = {};
  KEYS.forEach((k, i) => {
    let v = r[i];
    if (v instanceof Date) v = v.toISOString();
    it[k] = v;
  });
  it.id = String(it.id);
  it.upc = String(it.upc || '');
  it.caseUpc = String(it.caseUpc || '');
  ['onHand', 'par', 'cost', 'caseQty', 'poursPerUnit'].forEach(k => {
    it[k] = it[k] === '' || it[k] === null ? '' : Number(it[k]);
  });
  return it;
}

function itemToRow_(it) {
  return KEYS.map(k => (it[k] === undefined || it[k] === null) ? '' : it[k]);
}

function cleanItem_(src) {
  const out = {};
  const str = (v, n) => String(v === undefined || v === null ? '' : v).trim().slice(0, n);
  const num = v => (v === '' || v === undefined || v === null || isNaN(Number(v))) ? '' : Number(v);
  if ('name' in src) out.name = str(src.name, 120);
  if ('category' in src) out.category = str(src.category, 40);
  if ('unit' in src) out.unit = str(src.unit, 30);
  if ('upc' in src) out.upc = normUpc_(src.upc);
  if ('caseUpc' in src) out.caseUpc = normUpc_(src.caseUpc);
  if ('caseQty' in src) out.caseQty = num(src.caseQty);
  if ('onHand' in src) out.onHand = num(src.onHand);
  if ('par' in src) out.par = num(src.par);
  if ('cost' in src) out.cost = num(src.cost);
  if ('posNames' in src) out.posNames = str(src.posNames, 500);
  if ('poursPerUnit' in src) out.poursPerUnit = num(src.poursPerUnit);
  return out;
}

function saveItem_(src, role, user) {
  const data = readItems_();
  const fields = cleanItem_(src);
  const now = new Date();
  const dupe = (code, selfId) => code && data.items.find(i => i.id !== selfId && (i.upc === code || i.caseUpc === code));

  if (src.id && data.rowOf[src.id]) {
    const it = data.items.find(i => i.id === src.id);
    const before = it.onHand;
    if (role === 'manager') {
      Object.assign(it, fields);
    } else {
      // Staff may only attach barcodes to an item that doesn't have them yet.
      STAFF_LINK_FIELDS.forEach(k => { if (k in fields && (it[k] === '' || it[k] === null)) it[k] = fields[k]; });
    }
    for (const code of [it.upc, it.caseUpc]) {
      const other = dupe(code, it.id);
      if (other) return { ok: false, error: `Barcode ${code} is already on "${other.name}"` };
    }
    it.updated = now; it.updatedBy = user;
    data.sh.getRange(data.rowOf[it.id], 1, 1, COLS.length).setValues([itemToRow_(it)]);
    const change = (Number(it.onHand) || 0) - (Number(before) || 0);
    appendLog_([[now, user, 'EDIT', it.id, it.name, change || '', it.onHand, src.note || '', '']]);
    return { ok: true, item: rowToItem_(itemToRow_(it)) };
  }

  if (!fields.name) return { ok: false, error: 'Item needs a name' };
  for (const code of [fields.upc, fields.caseUpc]) {
    const other = dupe(code, null);
    if (other) return { ok: false, error: `Barcode ${code} is already on "${other.name}"` };
  }
  const it = Object.assign({ category: '', unit: 'each', upc: '', caseUpc: '', caseQty: '', onHand: 0, par: '', cost: '',
    posNames: '', poursPerUnit: 1, lastCounted: '' }, fields);
  if (role !== 'manager') { it.cost = ''; it.posNames = ''; }
  it.id = newId_();
  it.updated = now; it.updatedBy = user;
  if (it.onHand === '') it.onHand = 0;
  writeRows_(data.sh, data.last, [itemToRow_(it)]);
  appendLog_([[now, user, 'NEW ITEM', it.id, it.name, it.onHand || '', it.onHand, src.note || '', '']]);
  return { ok: true, item: rowToItem_(itemToRow_(it)) };
}

function importItems_(list, role, user) {
  if (role !== 'manager') return { ok: false, error: 'Manager PIN required to import items' };
  if (!Array.isArray(list) || list.length > 500) return { ok: false, error: 'Send 1–500 items at a time' };
  const data = readItems_();
  const names = {};
  const codes = {};
  data.items.forEach(i => {
    names[i.name.toLowerCase()] = true;
    if (i.upc) codes[i.upc] = true;
    if (i.caseUpc) codes[i.caseUpc] = true;
  });
  const now = new Date();
  const made = [];
  list.forEach(src => {
    const f = cleanItem_(src);
    if (!f.name || names[f.name.toLowerCase()]) return;
    if (f.upc && codes[f.upc]) f.upc = '';
    if (f.caseUpc && codes[f.caseUpc]) f.caseUpc = '';
    const it = Object.assign({ category: '', unit: 'each', upc: '', caseUpc: '', caseQty: '', onHand: 0, par: '', cost: '',
      posNames: '', poursPerUnit: 1, lastCounted: '' }, f, { id: newId_(), updated: now, updatedBy: user });
    if (it.onHand === '') it.onHand = 0;
    if (it.poursPerUnit === '') it.poursPerUnit = 1;
    names[it.name.toLowerCase()] = true;
    if (it.upc) codes[it.upc] = true;
    if (it.caseUpc) codes[it.caseUpc] = true;
    made.push(it);
  });
  if (made.length) {
    writeRows_(data.sh, data.last, made.map(itemToRow_));
    appendLog_([[now, user, 'IMPORT', '', made.length + ' items', '', '', '', '']]);
  }
  return { ok: true, created: made.length, items: made.map(it => rowToItem_(itemToRow_(it))) };
}

function deleteItem_(id, role, user) {
  if (role !== 'manager') return { ok: false, error: 'Manager PIN required to delete' };
  const data = readItems_();
  const row = data.rowOf[id];
  if (!row) return { ok: false, error: 'Item not found' };
  const it = data.items.find(i => i.id === id);
  data.sh.deleteRow(row);
  appendLog_([[new Date(), user, 'DELETE', id, it.name, '', '', '', '']]);
  return { ok: true };
}

/* ---------------------------------------------------------------- stock changes */

/**
 * ops: [{opId, id, mode: 'add'|'remove'|'set', qty, note, source}]
 * source 'sales' (SkyTab import) needs the manager PIN.
 */
function runOps_(ops, role, user, req) {
  if (!Array.isArray(ops) || !ops.length) return { ok: true, results: [], items: [] };
  if (ops.length > 1000) return { ok: false, error: 'Too many changes in one batch' };
  if (ops.some(o => o.source === 'sales') && role !== 'manager') {
    return { ok: false, error: 'Manager PIN required to import sales' };
  }
  const props = PropertiesService.getScriptProperties();
  const importKey = req && req.importKey ? 'import_' + String(req.importKey).slice(0, 200) : null;
  if (importKey && props.getProperty(importKey) && !req.force) {
    return { ok: false, code: 'duplicate_import', error: 'This sales report was already imported on ' + props.getProperty(importKey) };
  }

  const cache = CacheService.getScriptCache();
  const seen = cache.getAll(ops.filter(o => o.opId).map(o => 'op_' + o.opId));
  const data = readItems_();
  const byId = {};
  data.items.forEach(i => { byId[i.id] = i; });
  const changed = {};
  const logRows = [];
  const results = [];
  const now = new Date();

  ops.forEach(o => {
    if (o.opId && seen['op_' + o.opId]) { results.push({ opId: o.opId, ok: true, duplicate: true }); return; }
    const it = byId[String(o.id)];
    const qty = Number(o.qty);
    if (!it) { results.push({ opId: o.opId, ok: false, error: 'Item not found' }); return; }
    if (!isFinite(qty) || qty < 0 || qty > 100000) { results.push({ opId: o.opId, ok: false, error: 'Bad quantity' }); return; }
    const before = Number(it.onHand) || 0;
    let after;
    if (o.mode === 'add') after = before + qty;
    else if (o.mode === 'remove') after = before - qty;
    else if (o.mode === 'set') { after = qty; it.lastCounted = now; }
    else { results.push({ opId: o.opId, ok: false, error: 'Bad mode' }); return; }
    after = Math.round(after * 100) / 100;
    it.onHand = after; it.updated = now; it.updatedBy = user;
    changed[it.id] = true;
    const action = { add: 'RECEIVED', remove: o.source === 'sales' ? 'SOLD (SkyTab)' : 'REMOVED', set: 'COUNT' }[o.mode];
    const when = o.at ? new Date(o.at) : now;
    logRows.push([isNaN(when) ? now : when, user, action, it.id, it.name, Math.round((after - before) * 100) / 100,
      after, String(o.note || '').slice(0, 200), String(o.opId || '')]);
    results.push({ opId: o.opId, ok: true, id: it.id, onHand: after });
  });

  Object.keys(changed).forEach(id => {
    data.sh.getRange(data.rowOf[id], 1, 1, COLS.length).setValues([itemToRow_(byId[id])]);
  });
  appendLog_(logRows);
  const done = {};
  ops.forEach(o => { if (o.opId) done['op_' + o.opId] = '1'; });
  if (Object.keys(done).length) cache.putAll(done, 21600);
  if (importKey) props.setProperty(importKey, Utilities.formatDate(now, Session.getScriptTimeZone(), 'M/d/yyyy h:mm a'));

  return {
    ok: true,
    results,
    items: Object.keys(changed).map(id => rowToItem_(itemToRow_(byId[id])))
  };
}

/* ---------------------------------------------------------------- log */

function appendLog_(rows) {
  if (!rows.length) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
  if (!sh) throw new Error('Log sheet missing: run setup in Apps Script');
  writeRows_(sh, sh.getLastRow(), rows);
}

function readLog_(limit, itemId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const take = Math.min(last - 1, itemId ? 3000 : Math.min(limit, 500));
  const rows = sh.getRange(last - take + 1, 1, take, LOG_COLS.length).getValues().reverse();
  return rows
    .filter(r => !itemId || String(r[3]) === String(itemId))
    .slice(0, limit)
    .map(r => ({
      at: r[0] instanceof Date ? r[0].toISOString() : r[0], user: r[1], action: r[2], id: String(r[3]),
      name: r[4], change: r[5], onHand: r[6], note: r[7]
    }));
}

/* ---------------------------------------------------------------- UPC lookup */

function lookupUpc_(upc) {
  upc = normUpc_(upc);
  if (!upc) return null;
  const cache = CacheService.getScriptCache();
  const hit = cache.get('upc_' + upc);
  if (hit) return JSON.parse(hit);
  let product = null;
  try {
    const off = UrlFetchApp.fetch('https://world.openfoodfacts.org/api/v2/product/' + upc +
      '.json?fields=product_name,brands,quantity', { muteHttpExceptions: true, headers: { 'User-Agent': 'VFW9509-Inventory/1.0' } });
    if (off.getResponseCode() === 200) {
      const p = JSON.parse(off.getContentText()).product;
      if (p && p.product_name) product = { name: p.product_name, brand: p.brands || '', size: p.quantity || '', source: 'Open Food Facts' };
    }
  } catch (e) { /* try next source */ }
  if (!product) {
    try {
      const r = UrlFetchApp.fetch('https://api.upcitemdb.com/prod/trial/lookup?upc=' + upc, { muteHttpExceptions: true });
      if (r.getResponseCode() === 200) {
        const it = (JSON.parse(r.getContentText()).items || [])[0];
        if (it && it.title) product = { name: it.title, brand: it.brand || '', size: it.size || '', source: 'UPCitemdb' };
      }
    } catch (e) { /* no match */ }
  }
  cache.put('upc_' + upc, JSON.stringify(product), 21600);
  return product;
}

/* ---------------------------------------------------------------- helpers */

function normUpc_(v) {
  let s = String(v === undefined || v === null ? '' : v).replace(/\D/g, '');
  if (s.length === 13 && s[0] === '0') s = s.slice(1); // EAN-13 with a leading 0 is a UPC-A
  return s;
}

function newId_() {
  return 'I' + Utilities.getUuid().replace(/-/g, '').slice(0, 9).toUpperCase();
}

function colLetter(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
