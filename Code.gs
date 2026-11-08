/**
 * Creekside Lions PTO · Spirit Store Tracker · Google Sheets backend
 *
 * The Google Sheet is the master copy. The web page reads from it and saves to it through this script.
 * Setup (one time):
 *   1. Upload "Creekside Spirit Store.xlsx" to Google Drive and open it as a Google Sheet.
 *   2. In the Google Sheet: Extensions > Apps Script. Delete what is there, paste this whole file, Save.
 *   3. Reload the Google Sheet. A "Spirit Store" menu appears. Run Spirit Store > Set PINs.
 *   4. In Apps Script: Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Deploy.
 *   5. Copy the Web app URL into the API_URL line at the top of index.html.
 * After you change this code later, use Deploy > Manage deployments > Edit > New version, so the same URL keeps working.
 */

/* ===== Shared rules: used by both the web page and the Google Apps Script backend ===== */
var SCHEMA = {
  items: { tab: 'Items', cols: [['id','ID'],['name','Item'],['category','Category'],['price','Sell Price'],['unitCost','Cost per Unit (if no order)'],['onHand','On Hand Before'],['countedAfter','Counted After'],['notes','Notes']] },
  purchases: { tab: 'Orders', cols: [['id','ID'],['date','Order Date'],['store','Store'],['item','Item'],['packs','Packs'],['packCost','Pack Cost'],['unitsPerPack','Units / Pack'],['source','Cost Source'],['sale','For Sale'],['notes','Notes']] },
  moneyIn: { tab: 'Money In', cols: [['id','ID'],['date','Date'],['event','Event'],['amount','Amount'],['notes','Notes']] },
  moneyOut: { tab: 'Money Out', cols: [['id','ID'],['date','Date'],['order','Order'],['amount','Amount'],['link','Auto'],['notes','Notes']] },
  sales: { tab: 'Sale History', cols: [['id','ID'],['date','Sale Date'],['theme','Theme'],['spent','Spent on Store Items'],['sales','Spirit Store Sales'],['other','Other Fundraising'],['what','What it was'],['link','Auto']] },
  cash: { tab: 'Cash Count', cols: [['id','ID'],['label','Line'],['value','Value']] },
  settings: { tab: 'Settings', cols: [['id','Setting'],['value','Value']] },
  log: { tab: 'Change Log', cols: [['time','Time'],['name','Who'],['role','Role'],['table','Section'],['rowId','Row'],['field','Field'],['old','Old Value'],['new','New Value']] }
};
var NUMERIC = {
  items: ['price','unitCost','onHand','countedAfter'],
  purchases: ['packs','packCost','unitsPerPack'],
  moneyIn: ['amount'], moneyOut: ['amount'],
  sales: ['spent','sales','other']
};
/* What a volunteer PIN may change. The admin PIN may change everything. */
var VOLUNTEER_CAN = {
  items: ['onHand','countedAfter','notes'],
  moneyIn: ['amount','notes'],
  cash: ['value']
};
var CASH_NUMERIC_IDS = ['start','end','cashapp'];
var NONNEGATIVE = ['price','unitCost','onHand','countedAfter','packs','packCost','unitsPerPack'];

function fieldsOf(table) { return SCHEMA[table].cols.map(function (c) { return c[0]; }); }
function isNumeric(table, field, id) {
  if (table === 'cash' && field === 'value') return CASH_NUMERIC_IDS.indexOf(String(id)) >= 0;
  if (table === 'settings' && field === 'value') return String(id) === 'startCash';
  return (NUMERIC[table] || []).indexOf(field) >= 0;
}
function cleanVal(table, field, v, id) {
  if (v === null || v === undefined) return '';
  if (isNumeric(table, field, id)) {
    var s = String(v).replace(/[$,\s]/g, '');
    if (s === '') return '';
    if (/^\(.*\)$/.test(s)) s = '-' + s.slice(1, -1);
    var n = Number(s);
    if (!isFinite(n)) return '';
    if (n < 0 && NONNEGATIVE.indexOf(field) >= 0) return '';
    return n;
  }
  return String(v).trim();
}
function normVal(table, field, v, id) {
  var c = cleanVal(table, field, v, id);
  if (c === '') return '';
  if (typeof c === 'number') return String(Math.round(c * 10000) / 10000);
  return c;
}
function canDo(role, op) {
  if (role === 'admin' || role === 'super') return true; // super admin can do everything an admin can
  if (role !== 'volunteer') return false;
  // the upload screen's remembered answers live in one settings row that volunteers may update
  if (op.table === 'settings' && (op.op === 'set' ? (op.id === 'learned' && op.field === 'value') : (op.op === 'add' && op.row && op.row.id === 'learned'))) return true;
  return op.op === 'set' && (VOLUNTEER_CAN[op.table] || []).indexOf(op.field) >= 0;
}
function labelOf(table, row) {
  if (!row) return '';
  var k = { items: 'name', purchases: 'item', moneyIn: 'event', moneyOut: 'order', sales: 'theme', cash: 'label', settings: 'id' }[table];
  return String(row[k] || row.id || '');
}
/* Apply a batch of edits to in-memory tables. Each "set" carries the value the editor started from (prev);
   if someone else changed that value in the meantime, the edit is held back as a conflict instead of overwriting. */
function applyOps(t, ops, role, who, nowIso, atomic) {
  if (atomic) {
    // all or nothing: dry-run on a copy first, apply only if every edit is clean
    var trial = applyOps(JSON.parse(JSON.stringify(t)), ops, role, who, nowIso, false);
    if (trial.conflicts.length) return { applied: 0, conflicts: trial.conflicts, log: [], rejected: true };
  }
  var res = { applied: 0, conflicts: [], log: [] };
  for (var i = 0; i < ops.length; i++) {
    var op = ops[i];
    if (!op || !SCHEMA[op.table] || op.table === 'log' || !t[op.table]) { res.conflicts.push({ op: op, reason: 'unknown section' }); continue; }
    if (!canDo(role, op)) { res.conflicts.push({ op: op, reason: 'not allowed with this PIN' }); continue; }
    var table = t[op.table], fields = fieldsOf(op.table);
    if (op.op === 'set') {
      if (fields.indexOf(op.field) < 0 || op.field === 'id') { res.conflicts.push({ op: op, reason: 'unknown field' }); continue; }
      var row = null;
      for (var j = 0; j < table.length; j++) if (String(table[j].id) === String(op.id)) { row = table[j]; break; }
      if (!row) { res.conflicts.push({ op: op, reason: 'row was removed' }); continue; }
      var cur = normVal(op.table, op.field, row[op.field], op.id);
      var nv = cleanVal(op.table, op.field, op.value, op.id);
      if (cur === normVal(op.table, op.field, nv, op.id)) continue;
      if (cur !== normVal(op.table, op.field, op.prev, op.id)) { res.conflicts.push({ op: op, reason: 'changed', current: row[op.field] }); continue; }
      res.log.push({ time: nowIso, name: who, role: role, table: op.table, rowId: String(op.id), field: op.field, old: row[op.field] === undefined ? '' : String(row[op.field]), 'new': String(nv) });
      row[op.field] = nv; res.applied++;
    } else if (op.op === 'add') {
      var src = op.row || {}, nr = {};
      for (var f = 0; f < fields.length; f++) nr[fields[f]] = cleanVal(op.table, fields[f], src[fields[f]], src.id);
      if (!nr.id) nr.id = op.table.slice(0, 2) + Date.now().toString(36) + Math.floor(Math.random() * 1e4);
      var dup = false;
      for (var k = 0; k < table.length; k++) if (String(table[k].id) === String(nr.id)) dup = true;
      if (dup) { res.conflicts.push({ op: op, reason: 'duplicate' }); continue; }
      table.push(nr);
      res.log.push({ time: nowIso, name: who, role: role, table: op.table, rowId: String(nr.id), field: '(added)', old: '', 'new': labelOf(op.table, nr) });
      res.applied++;
    } else if (op.op === 'del') {
      var idx = -1;
      for (var m = 0; m < table.length; m++) if (String(table[m].id) === String(op.id)) idx = m;
      if (idx < 0) continue;
      res.log.push({ time: nowIso, name: who, role: role, table: op.table, rowId: String(op.id), field: '(removed)', old: labelOf(op.table, table[idx]), 'new': '' });
      table.splice(idx, 1); res.applied++;
    }
  }
  return res;
}

/* ===== Sale numbers that follow Money In =====
   A closed sale's store sales and other fundraising should be the money that actually came in for it. Every Money In
   line belongs to a sale by its date and its name, and counts as store money unless its name says it is something else
   (class sponsors, face painting, donations and the like). Within reason the sale follows Money In by itself. A big
   difference is asked about instead of changed, and so is a sale whose Money In is still waiting on an amount.
   The page and the backend both use this, so the report, the History tab and the Google Sheet always agree. */
var OTHER_MONEY_RE = /sponsor|face ?paint|donat|grant\b|raffle|auction|bake sale|concession|book ?fair|membership|\bdues\b|\(other\)/i;
var NOT_STORE_NOTE_RE = /\bnot (the )?(spirit )?store\b/i;
var MS_GENERIC = { spirit: 1, store: 1, stores: 1, sale: 1, sales: 1, shop: 1, school: 1, pto: 1, creekside: 1, lion: 1, lions: 1, the: 1, and: 1, of: 1, a: 1, cash: 1, app: 1, deposit: 1, deposits: 1, zeffy: 1, zelle: 1, venmo: 1, paypal: 1, total: 1 };
function msNorm(s) { return String(s === undefined || s === null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }
function msBase(ev) { var s = String(ev === undefined || ev === null ? '' : ev), prev; do { prev = s; s = s.replace(/\s*\([^()]*\)\s*$/, ''); } while (s !== prev); return msNorm(s); }
function msBlank(v) { return v === undefined || v === null || String(v).trim() === ''; }
function msNum(v) { var n = Number(String(v === undefined || v === null ? '' : v).replace(/[$,\s]/g, '')); return isFinite(n) ? n : 0; }
function msR2(n) { return Math.round(n * 100) / 100; }
function msDay(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') return isNaN(d.getTime()) ? null : Math.round(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 864e5);
  var t = String(d === undefined || d === null ? '' : d).trim(), m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return Math.round(Date.UTC(+m[1], +m[2] - 1, +m[3]) / 864e5);
  m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // typed straight into the Sheet as 10/22/2026
  return m ? Math.round(Date.UTC(+m[3], +m[1] - 1, +m[2]) / 864e5) : null;
}
/* how well a line's name matches a sale's theme: 0 = not at all; spec = the match has a real word, not just "spirit store" */
function msMatch(line, sale) {
  var th = msNorm(sale.theme); if (!th) return { n: 0, spec: false };
  var ev = ' ' + msNorm(line.event) + ' ', base = msBase(line.event), hit = '';
  if (ev.indexOf(' ' + th + ' ') >= 0) hit = th;
  else if (base.length >= 4 && (' ' + th + ' ').indexOf(' ' + base + ' ') >= 0) hit = base;
  if (!hit) return { n: 0, spec: false };
  return { n: hit.length, spec: hit.split(' ').some(function (w) { return !MS_GENERIC[w]; }) };
}
/* other fundraising: its name says so (a word the sale's own name does not have: money called "Book Fair Night" at a
   sale called Book Fair Night is that sale's store money), or its note says it is not store money */
function msIsOther(line, sale) {
  if (NOT_STORE_NOTE_RE.test(String(line.notes || ''))) return true;
  var re = new RegExp(OTHER_MONEY_RE.source, 'gi'), ev = String(line.event || ''), th = String((sale && sale.theme) || '').toLowerCase(), m;
  while ((m = re.exec(ev))) if (th.indexOf(m[0].toLowerCase()) < 0) return true;
  return false;
}
/* a line still waiting on its amount: blank, or $0 with the "Enter the ..." note the close-out left on it */
function msOpen(line) { return msBlank(line.amount) || (msNum(line.amount) === 0 && /^\s*enter\b/i.test(String(line.notes || ''))); }
/* which sale a Money In line is for, or null. unsure = two sales that day and the name doesn't say which */
function msSaleFor(line, sales) {
  var d = msDay(line.date); if (d === null) return { sale: null };
  var best = function (list, needSpec) {
    var top = null, topN = 0, tie = false;
    list.forEach(function (s) { var m = msMatch(line, s); if (!m.n || (needSpec && !m.spec)) return; if (m.n > topN) { top = s; topN = m.n; tie = false; } else if (m.n === topN) tie = true; });
    return tie ? 'tie' : top;
  };
  var same = sales.filter(function (s) { return msDay(s.date) === d; });
  var b = best(same, false);                       // 1. a sale that day with the line's name
  if (b && b !== 'tie') return { sale: b };
  if (b === 'tie') return { sale: null, unsure: same };
  var near = sales.filter(function (s) { var sd = msDay(s.date); return sd !== null && d - sd >= -3 && d - sd <= 14; });
  b = best(near, true);                            // 2. a sale a few days around it with its own name in the line (a deposit made later)
  if (b && b !== 'tie') return { sale: b };
  if (same.length === 1) return { sale: same[0] }; // 3. the only sale that day
  if (same.length > 1) return { sale: null, unsure: same };
  return { sale: null };
}
function reconcileSales(sales, moneyIn, keep) {
  sales = sales || []; moneyIn = moneyIn || []; keep = keep && typeof keep === 'object' ? keep : {};
  var out = { sales: {}, lines: {}, unsure: [], ops: [] };
  sales.forEach(function (s) { out.sales[String(s.id)] = { store: { sum: 0, n: 0, open: [], lines: [] }, other: { sum: 0, n: 0, open: [], lines: [] } }; });
  moneyIn.forEach(function (m) {
    var f = msSaleFor(m, sales);
    if (!f.sale) { if (f.unsure) out.unsure.push({ line: m, sales: f.unsure }); return; }
    var kind = msIsOther(m, f.sale) ? 'other' : 'store', g = out.sales[String(f.sale.id)][kind];
    out.lines[String(m.id)] = { sale: String(f.sale.id), kind: kind };
    g.lines.push(m);
    if (msOpen(m)) g.open.push(m); else { g.sum += msNum(m.amount); g.n++; }
  });
  sales.forEach(function (s) {
    var r = out.sales[String(s.id)], live = s.link === 'current';
    ['store', 'other'].forEach(function (kind) {
      var g = r[kind], field = kind === 'store' ? 'sales' : 'other', stored = msBlank(s[field]) ? null : msNum(s[field]);
      var mi = msR2(g.sum), k = keep[String(s.id)] && keep[String(s.id)][field];
      g.stored = stored; g.mi = g.n ? mi : null; g.v = stored;
      if (live && kind === 'store') { g.status = 'live'; return; } // the current sale uses the cash box until it is closed out
      if (!g.n && !g.open.length) { g.status = 'none'; return; }
      if (g.open.length) { g.status = 'open'; if (stored === null && g.n) g.v = mi; return; }
      var diff = stored === null ? null : msR2(mi - stored);
      if (diff !== null && Math.abs(diff) < 0.005) { g.status = 'ok'; g.v = mi; return; }
      if (k !== undefined && k !== null && Math.abs(msNum(k) - mi) < 0.005) { g.status = 'kept'; return; }
      var ok = stored === null || (kind === 'store' ? Math.abs(diff) <= Math.max(100, 0.25 * Math.max(Math.abs(stored), Math.abs(mi))) : diff > 0);
      if (!ok) { g.status = 'ask'; return; }
      g.status = 'fix'; g.v = mi;
      out.ops.push({ op: 'set', table: 'sales', id: s.id, field: field, prev: msBlank(s[field]) ? '' : s[field], value: mi });
    });
  });
  return out;
}
function salesKeepFrom(settingsRows) {
  var r = (settingsRows || []).filter(function (x) { return String(x.id) === 'salesKeep'; })[0];
  if (!r || msBlank(r.value)) return {};
  try { var o = JSON.parse(String(r.value)); return o && typeof o === 'object' && !Array.isArray(o) ? o : {}; } catch (e) { return {}; }
}


var SEED = null;

/* ------------------------------ Sheet menu ------------------------------ */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Spirit Store')
    .addItem('Set PINs', 'setPins')
    .addItem('Set super admin PIN', 'setSuperPin')
    .addItem('Require an access code to view (optional)', 'setViewCode')
    .addItem('Set AI key (for importing order screenshots)', 'setAiKey')
    .addItem('Test the AI order reader', 'selfTestAi')
    .addSeparator()
    .addItem('Match sale numbers to Money In', 'fixSalesFromMoneyIn')
    .addItem('Rebuild any missing tabs', 'setupTabs')
    .addToUi();
}

function setupTabs() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(SCHEMA).forEach(function (t) {
    var sh = ss.getSheetByName(SCHEMA[t].tab) || ss.insertSheet(SCHEMA[t].tab);
    if (sh.getLastRow() > 0) return; // never overwrite a tab that already has data
    var headers = SCHEMA[t].cols.map(function (c) { return c[1]; });
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#0C482D').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    var rows = ((SEED && SEED[t]) || []).map(function (r) { return SCHEMA[t].cols.map(function (c) { var v = r[c[0]]; return v === undefined || v === null ? '' : v; }); });
    if (rows.length) { ensureRows(sh, rows.length + 1); sh.getRange(2, 1, rows.length, headers.length).setValues(rows); }
    sh.autoResizeColumns(1, headers.length);
  });
  var first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  SpreadsheetApp.getUi().alert('Tabs are ready. Next: Spirit Store > Set PINs.');
}

function setPins() {
  var ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  var v = ui.prompt('Volunteer PIN', 'Volunteers can enter counts, deposits and the cash box. 4 to 8 digits works well.', ui.ButtonSet.OK_CANCEL);
  if (v.getSelectedButton() !== ui.Button.OK) return;
  var a = ui.prompt('Admin PIN', 'Admins can change everything: items, prices, orders, history. Keep this one to yourself.', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var vp = v.getResponseText().trim(), ap = a.getResponseText().trim();
  if (!vp || !ap || vp === ap) { ui.alert('Both PINs are needed and they must be different. Nothing was changed.'); return; }
  props.setProperties({ VOLUNTEER_PIN: vp, ADMIN_PIN: ap });
  ui.alert('PINs saved. They take effect right away, no redeploy needed.');
}

/* The super admin PIN unlocks everything an admin can do, plus the Usage page (who opens the tracker and what they do).
   It can only be set here, by the owner of the Google Sheet, never from the website. */
function setSuperPin() {
  var ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  var r = ui.prompt('Super admin PIN', 'Unlocks everything plus the Usage page. Pick one only you know (8 or more digits works well). Leave it blank and tap OK to turn super admin off.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var sp = r.getResponseText().trim();
  if (!sp) { props.deleteProperty('SUPER_PIN'); ui.alert('Super admin is off.'); return; }
  if (sp.length < 6) { ui.alert('Use at least 6 characters. Nothing was changed.'); return; }
  if (sp === props.getProperty('VOLUNTEER_PIN') || sp === props.getProperty('ADMIN_PIN')) { ui.alert('That matches another PIN. Pick a different one. Nothing was changed.'); return; }
  props.setProperty('SUPER_PIN', sp);
  ui.alert('Super admin PIN saved. Unlock the tracker with it to see the Usage tab.');
}

function setViewCode() {
  var ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  var r = ui.prompt('Access code to view', 'Leave blank so anyone with the link can view (they still need a PIN to edit). Or type a code people must enter to see the numbers.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var code = r.getResponseText().trim();
  if (code) props.setProperty('VIEW_CODE', code); else props.deleteProperty('VIEW_CODE');
  ui.alert(code ? 'Access code saved.' : 'Access code removed. Anyone with the link can view.');
}

/* ------------------------------ Web API ------------------------------ */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!codeOk(p.code)) return json({ ok: false, error: 'code' });
  if (p.action === 'photos') return json({ ok: true, photos: readPhotos() }); // anyone who can view can print the price sheet
  return ContentService.createTextOutput('{"ok":true,"data":' + dataJson() + '}').setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Bad request.' }); }
  if (!codeOk(body.code)) return json({ ok: false, error: 'code' });
  if (body.action === 'ping') return json(ping(body));
  var role = roleFor(body.pin);
  if (role.error) { if (body.action === 'auth') logVisit(body, 'wrongpin', '', ''); return json({ ok: false, error: role.error }); }
  var isSuper = role.role === 'super';
  if (isSuper) { role = { role: 'admin' }; body.me = true; } // a super admin can do everything an admin can, and the owner's own use stays out of the counts
  if (body.action === 'auth') { logVisit(body, 'unlock', role.role, role.role); return json({ ok: true, role: isSuper ? 'super' : role.role }); }
  if (body.action === 'usage') return json(isSuper ? usage() : { ok: false, error: 'Only the super admin PIN can see usage.' });
  if (body.action === 'setPins') return json(setPinsFromWeb(body, role.role));
  if (body.action === 'importRead') return json(importRead(body, role.role));
  if (body.action === 'readDoc') return json(readDoc(body, role.role));
  if (body.action === 'ask') return json(ask(body, role.role, false, isSuper));
  if (body.action === 'setPhoto') return json(setPhoto(body, role.role));
  if (body.action === 'keepUpload') return json(upSafe(keepUpload, body, role.role));
  if (body.action === 'markUpload') return json(upSafe(markUpload, body, role.role));
  if (body.action === 'uploads') return json(upSafe(listUploads, body, role.role));
  if (body.action === 'getUpload') return json(upSafe(getUpload, body, role.role));
  if (body.action === 'delUpload') return json(upSafe(delUpload, body, role.role));
  if (body.action !== 'save' || !Array.isArray(body.ops)) return json({ ok: false, error: 'Bad request.' });
  if (body.ops.length > 20000) return json({ ok: false, error: 'Too many changes at once. Save in smaller batches.' });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json({ ok: false, error: 'The tracker is busy saving someone else\'s changes. Try again in a moment.' });
  try {
    var touched = {}, tables = {};
    body.ops.forEach(function (op) { if (op && SCHEMA[op.table] && op.table !== 'log') touched[op.table] = true; });
    Object.keys(touched).forEach(function (t) { tables[t] = readTable(t).rows; });
    var who = String(body.name || 'Someone').slice(0, 40);
    var now = new Date().toISOString();
    var res = applyOps(tables, body.ops, role.role, who, now, !!body.atomic);
    if (res.applied) {
      res.log = res.log.concat(autoFixSales(tables, touched, now)); // closed sales follow Money In, every save
      Object.keys(touched).forEach(function (t) { writeTable(t, tables[t]); });
      appendLog(res.log);
      if (res.log && res.log.length) logVisit(body, 'save', String(res.log.length), role.role);
    }
    return ContentService.createTextOutput('{"ok":true,"applied":' + res.applied + ',"conflicts":' + JSON.stringify(res.conflicts) + ',"data":' + dataJson() + '}').setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/* After every save, each closed sale's store sales and other fundraising are brought in line with its Money In lines,
   within reason (reconcileSales in the shared rules), so the Sheet stays right and each change shows in Activity as
   "Tracker auto-fix". Bigger differences, and sales whose Money In is still waiting on an amount, are left for the
   admin to settle on the page. It reads only what the save did not already read, and it never stops a save. */
function autoFixSales(tables, touched, now) {
  try {
    ['sales', 'moneyIn', 'settings'].forEach(function (t) { if (!tables[t]) tables[t] = readTable(t).rows; });
    var rc = reconcileSales(tables.sales, tables.moneyIn, salesKeepFrom(tables.settings));
    if (!rc.ops.length) return [];
    var fx = applyOps(tables, rc.ops, 'admin', 'Tracker auto-fix', now, false);
    if (!fx.applied) return [];
    touched.sales = true;
    fx.log.forEach(function (e) { e.role = 'auto'; });
    return fx.log;
  } catch (e) { return []; }
}
/* Run from the menu (Spirit Store > Match sale numbers to Money In) or the script editor any time. Safe to run again. */
function fixSalesFromMoneyIn() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy. Try again in a minute.');
  try {
    var tables = {}, touched = {}, log = autoFixSales(tables, touched, new Date().toISOString());
    if (touched.sales) { writeTable('sales', tables.sales); appendLog(log); }
    var name = {}; tables.sales.forEach(function (s) { name[String(s.id)] = s.theme || s.id; });
    var said = log.map(function (e) { return name[e.rowId] + ' ' + (e.field === 'sales' ? 'store sales' : 'other fundraising') + ' ' + e.old + ' to ' + e['new']; });
    var rc = reconcileSales(tables.sales, tables.moneyIn, salesKeepFrom(tables.settings)), left = [];
    tables.sales.forEach(function (s) {
      var r = rc.sales[String(s.id)];
      ['store', 'other'].forEach(function (k) {
        var g = r[k], lab = k === 'store' ? 'store sales' : 'other fundraising';
        if (g.status === 'open') left.push(name[String(s.id)] + ' ' + lab + ': waiting on ' + g.open.map(function (m) { return m.event; }).join(' and '));
        if (g.status === 'ask') left.push(name[String(s.id)] + ' ' + lab + ': the sale says ' + g.stored + ', Money In shows ' + g.mi);
      });
    });
    var msg = (said.length ? 'Fixed ' + said.length + ': ' + said.join('; ') + '.' : 'Nothing needed fixing.') + (left.length ? ' Left for the admin to settle on the page: ' + left.join('; ') + '.' : ' Every closed sale matches Money In.');
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
    return msg;
  } finally { lock.releaseLock(); }
}

/* Admins can change PINs and the access code from the website. The values are never sent back out. */
function setPinsFromWeb(body, role) {
  if (role !== 'admin') return { ok: false, error: 'Only the admin PIN can change PINs.' };
  var props = PropertiesService.getScriptProperties();
  var vp = String(body.newVolunteerPin || '').trim(), ap = String(body.newAdminPin || '').trim();
  var curV = props.getProperty('VOLUNTEER_PIN'), curA = props.getProperty('ADMIN_PIN');
  var nextV = vp || curV, nextA = ap || curA;
  if (vp && vp.length < 4) return { ok: false, error: 'The volunteer PIN needs at least 4 characters.' };
  if (ap && ap.length < 4) return { ok: false, error: 'The admin PIN needs at least 4 characters.' };
  if (nextV === nextA) return { ok: false, error: 'The volunteer and admin PINs have to be different.' };
  var spin = props.getProperty('SUPER_PIN');
  if (spin && (nextV === spin || nextA === spin)) return { ok: false, error: 'That PIN is already taken. Pick a different one.' };
  var mode = body.viewMode || 'keep', code = String(body.viewCode || '').trim();
  if (mode === 'set' && !code) return { ok: false, error: 'Type the new access code, or choose to remove it.' };
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'The tracker is busy. Try again in a moment.' };
  try {
    var changed = [];
    if (vp && vp !== curV) { props.setProperty('VOLUNTEER_PIN', vp); changed.push('volunteer PIN'); }
    if (ap && ap !== curA) { props.setProperty('ADMIN_PIN', ap); changed.push('admin PIN'); }
    if (mode === 'set') { props.setProperty('VIEW_CODE', code); changed.push('access code'); }
    if (mode === 'clear' && props.getProperty('VIEW_CODE')) { props.deleteProperty('VIEW_CODE'); changed.push('access code removed'); }
    if (changed.length) appendLog([{ time: new Date().toISOString(), name: String(body.name || 'Someone').slice(0, 40), role: role, table: 'settings', rowId: 'security', field: '(security)', old: '', 'new': 'Changed: ' + changed.join(', ') }]);
    return { ok: true, changed: changed };
  } finally { lock.releaseLock(); }
}

/* ------------------------------ AI order reader ------------------------------ */
var AI_MODELS = ['claude-sonnet-5-5', 'claude-sonnet-4-5', 'claude-haiku-5-5'];
/* Anthropic retires older models every year or two. When every model above is gone, the tracker asks Anthropic
   which models this key can use, picks the newest Sonnet (or the newest of anything), remembers it in the Script
   Properties as AI_MODEL, and carries on. A model that works after the first one failed is remembered the same way. */
function aiModelList() {
  var saved = PropertiesService.getScriptProperties().getProperty('AI_MODEL');
  return (saved ? [saved] : []).concat(AI_MODELS.filter(function (m) { return m !== saved; }));
}
function aiFindModel(key, tried) {
  try {
    var r = UrlFetchApp.fetch('https://api.anthropic.com/v1/models?limit=100', { method: 'get', muteHttpExceptions: true, headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' } });
    if (r.getResponseCode() !== 200) return '';
    var ids = ((JSON.parse(r.getContentText()) || {}).data || []).map(function (m) { return String((m && m.id) || ''); })
      .filter(function (id) { return /^claude-[a-z0-9.-]+$/.test(id) && tried.indexOf(id) < 0; }); // newest first, as Anthropic lists them
    return ids.filter(function (id) { return /sonnet/.test(id); })[0] || ids.filter(function (id) { return /opus|haiku/.test(id); })[0] || ids[0] || '';
  } catch (e) { return ''; }
}

function setAiKey() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('AI key', 'Paste the Anthropic API key (it starts with sk-ant-). It is stored inside this script only and never shown on the website.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var key = r.getResponseText().trim();
  if (!/^sk-ant-/.test(key)) { ui.alert('That does not look like an Anthropic API key. Nothing was saved.'); return; }
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', key);
  var test = aiExtract([], 'Amazon order placed October 1, 2026. Pink ribbon bracelets, 24 Pack. Qty: 2. $11.99 each. Order total $23.98', ['Pink ribbon bracelets (24-pack)']);
  ui.alert(test.ok ? 'AI key saved and tested. The order reader works.' : 'Key saved, but the test failed: ' + test.error);
}

function selfTestAi() {
  var t = aiExtract([], 'Amazon order placed October 1, 2026. Halloween pumpkin squishies 12 Pack, Qty 3, $18.99. Glow sticks bulk 50 pack, Qty 2, $12.50 each.', ['Halloween pumpkin squishies (12-pack)']);
  Logger.log(JSON.stringify(t));
  try { SpreadsheetApp.getUi().alert(t.ok ? 'The order reader works. It found ' + t.result.lines.length + ' lines.' : 'Problem: ' + t.error); } catch (e) {}
  return t;
}

/* One upload for everything: works out what kind of document it is, then reads it. */
function cleanFiles(body) {
  return (Array.isArray(body.images) ? body.images : []).slice(0, 8).filter(function (im) {
    return im && /^(image\/(jpeg|png|webp|gif)|application\/pdf)$/.test(im.media_type) && typeof im.data === 'string' && im.data.length < 12000000;
  });
}
function readDoc(body, role) {
  if (role !== 'admin' && role !== 'volunteer') return { ok: false, error: 'Unlock with a PIN first.' };
  var files = cleanFiles(body);
  if (!files.length) return { ok: false, error: 'Add a photo, screenshot or PDF first.' };
  var names = (Array.isArray(body.names) ? body.names : []).slice(0, 2500).map(function (n) { return String(n).slice(0, 80); });
  var rows = (Array.isArray(body.rows) ? body.rows : []).slice(0, 2500).map(function (r) {
    return { n: Number(r.n) || 0, name: String(r.name || '').slice(0, 80), price: String(r.price || '').slice(0, 12), printedBefore: r.before === '' || r.before === null || r.before === undefined ? null : Number(r.before) };
  });
  var prompt = [
    'You are helping a school PTO spirit store tracker. First decide what these pages are:',
    '"order" = an online store cart, order confirmation, invoice or receipt for things bought to sell (usually Amazon).',
    '"countSheet" = the printed "Inventory Count Sheet" with numbered rows that volunteers filled in by hand.',
    '"deposit" = a bank deposit slip, deposit receipt, or bank record showing money put into the PTO account.',
    '"bags" = a handwritten list for mystery bags (or grab bags): how many bags at each price, and the products used to fill them with how many of each.',
    '"other" = anything else.',
    'Then return ONLY one JSON object, no other text, matching the type:',
    'For "order": {"type":"order","store":"Amazon","orderDate":"YYYY-MM-DD or empty","orderTotal":number or null,"lines":[{"title":"product title as shown","name":"short clean item name with pack size, e.g. Pink ribbon bracelets (24-pack)","match":"exact existing item name if the same product, else null","packs":integer,"packCost":price of ONE pack,"unitsPerPack":pieces in one pack or null if unclear,"note":""}]}',
    'Existing store items (exact names, for match): ' + JSON.stringify(names),
    'Order rules: one line per distinct product; skip Saved for later, recommendations, ads, shipping and tax; never invent products; if only a line total is shown divide by quantity.',
    'For "countSheet": {"type":"countSheet","countedBeforeBy":"","countedAfterBy":"","cash":{"start":number or null,"end":number or null,"cashApp":number or null},"notes":"","rows":[{"n":row number or null,"name":"item as written","before":number or null,"after":number or null,"sold":number or null,"note":"","unclear":true or false}]}',
    'Rows printed on the count sheet (n = row number): ' + JSON.stringify(rows),
    'Count sheet rules: include a row only if something is handwritten on it; for before give the final correct number (a handwritten correction beats a crossed-out printed number); null for blanks; never guess, mark unclear instead; include extra handwritten lines with n = null.',
    'For "deposit": {"type":"deposit","date":"YYYY-MM-DD or empty","amount":number,"what":"short description, e.g. Cash deposit or Checks","note":""}',
    'For "bags": {"type":"bags","name":"the list title as written, e.g. Halloween Mystery Bags","bags":[{"price":number,"count":integer}],"items":[{"name":"product as written","count":integer,"match":"exact existing item name from the list above if it is clearly the same product, else null"}],"more":"title of any other bag list on the page, else empty"}',
    'Bag list rules: "17 - $5" means 17 bags at $5; items are the totals used in all the bags together; read only the first complete list and name any other list in "more"; never invent numbers.',
    'For "other": {"type":"other","what":"one short sentence saying what it looks like"}'
  ].join('\n');
  return callClaude(files, prompt);
}

/* ------------------------------ Ask the tracker (admins only) ------------------------------
   A plain-language answer about the live numbers or how to do something. The page sends a short summary of
   its numbers with the question. Each question and answer is kept on the Ask Log tab of the Google Sheet, so answers can be checked. */
var HELP_GUIDE = [
  'How the tracker works (use these exact button and tab names):',
  '- One link for everyone. Looking needs only the access code. Changing anything needs a name and a PIN (button at the top: "Unlock to edit"). Volunteer PIN: counts, deposits, cash box. Admin PIN: everything, including orders, Cash App, sale details, PINs & access.',
  '- Tabs: Overview, Count, Money, Cash Box, Orders, History, Activity. Admins also see Uploads.',
  '- Overview: "Real profit to date" = money in minus spent plus stock on the shelf (valued at cost). The "This sale" card shows Items sold, Sold by count (after-sale counts times prices), Money taken in (cash taken in plus Cash App, once the cash box is counted), Profit (money taken in minus what the sold items cost). After unlocking, the number on the Overview tab and the "Needs your input" list show what that person can do next (viewers do not see it).',
  '- Sold by count and Money taken in rarely match exactly. Small gaps are normal: Cash App fees (if Cash App was entered after fees), pom pom pairs sold 2 for a deal, miscounts. A big short usually means an item was miscounted or money is missing; a big over usually means an item was missed on the Count tab.',
  '- Count tab: each item has BEFORE (leftovers on hand plus this sale\'s orders) and AFTER (what is left after the sale). Sold = before minus after. Items with nothing on hand do not need an after count.',
  '- Bundles: a bundle (like a Red Ribbon Week pre-order set) is an item sold as one thing that is made of other items. On the Count tab an admin taps Bundles, picks the bundle item and lists what goes in it. Each bundle sold is then taken out of those items, so a loose item shows "+N in bundles" and only the rest counts as sold on its own at its own price; the bundle earns its bundle price. Count bundles like any item: how many before (for pre-orders, how many were ordered) and how many are left after (0 once all are picked up). A red "Check count" tag means bundles used more of an item than went missing.',
  '- Cash Box tab: Starting cash (change put in the box, usually $150), Ending cash (EVERYTHING in the box at closing, including the starting cash), Cash App total for this sale. Cash taken in = ending minus starting. Over/(short) compares cash taken in plus Cash App with Sold by count. If ending cash looks like it leaves out the starting cash, a yellow "Check the ending cash" box offers a one-tap fix.',
  '- Money tab: Money in (bank deposits, Cash App totals) and Money out (orders and other costs). Cash deposits are net of the starting cash.',
  '- Orders tab (admin): what was bought for this sale. Earlier orders are kept below.',
  '- Buttons only show for people who can use them. Upload shows after unlocking with a PIN; Ask, PINs & access, Edit sale details, Add and Remove show only for admins.',
  '- Print button (top, anyone can use it) offers three PDFs: Price sheet (every item and its price with pictures, grouped like the table, from the tracker prices; "Leave off items with none left" is on by default), Count sheet (BEFORE and AFTER columns for volunteers), and Sales report (every sale with totals, the four headline numbers, and every money in and money out line). Print report buttons on the Overview card and the History tab make the sales report too.',
  '- Price sheet editing (admin): Print, Price sheet, "Edit names, groups and photos". Change the name a customer sees, the group, special price text like "1 for $2, 2 for $3", hide an item, or add a photo. Two items with the same name in the same group print as one box. A new item shows up on its own; to add something like a hat that is not in the tracker yet, add it as an item on the Count tab first.',
  '- Upload button (top, after unlocking): add a photo of the count sheet (each page), a deposit slip photo, an Amazon order screenshot or PDF, or the Cash App export file (admin). It reads it, asks one-tap questions about anything that does not add up, and nothing saves until every question is answered and Save is tapped. The green Saved screen has "Undo this upload".',
  '- Count sheet (Print button, then Count sheet): makes a PDF. After the after-sale counts are in, it asks which numbers go in BEFORE: "What\'s left now" (for the next sale) or the current sale\'s starting numbers. "Leave off items with none left" hides zero rows; row numbers never change.',
  '- Typing in a field: a bar appears at the bottom with "Submit changes" (saves for everyone) and "Discard".',
  '- History tab: past sales with spent, sales and profit. To finish a sale: History tab, "Done with <sale name>?" card, "Close out". It locks in the sale (money taken in, or sold by count if the cash box is empty), carries the after counts forward as leftovers, clears the cash box and starts the next sale.',
  '- Activity tab: every saved change with name and time. Admins can Undo a change there.',
  '- Sale numbers follow Money In: once a sale is closed, its store sales and other fundraising come from its lines on the Money tab (lines belong to a sale by date and name; class sponsors, face painting, donations and the like count as other fundraising, everything else as store sales). Small differences fix themselves on the next save and show in Activity as "Tracker auto-fix". A difference over $100 and over a quarter of the sale shows as a to-do ("Check <sale> store sales") where the admin picks Use or Keep. A Money In line still at $0 with the note "Enter the ..." keeps the sale from being final: type the real amount, or tap "It was $0" on that line. In History those numbers say "from Money In"; to change one, fix the Money In line.',
  '- Mystery bags (Count tab, admins): for bags filled with a mix of items (Halloween mystery bags, spirit grab bags). Tap Mystery bags, give them a name, enter each bag price with how many sold and how many were left over, then list how many of each item went into ALL the bags together (the totals off the list; nobody has to say what went in each bag). The tracker makes one item per bag price, takes the items inside out of their own sales (they show "+N in mystery bags" in the Sold column), counts the bags at the bag price, and spreads what the items inside cost over the bags by price, so profit is right. If the list says more of an item went in than the count shows gone, the item gets a "Check count" tag. A photo of a handwritten bag list can be uploaded with Upload and it fills in Mystery bags to check. Close-out marks the bags done for that sale. Without this, the items inside count as sold at full price, which makes the sale look short (Oct 30, 2026: $123.92 short before the Halloween bags were entered). If a sale is well short when closing out, the close-out screen asks about mystery bags and bundles first.',
  '- Cut up or split (Count tab, admins): for packs cut into pieces sold one at a time, like sticker sheets cut into single paws. Pick the item, say how many were cut up (blank means all), how many pieces there are now (count them; cutting rarely comes out even), the price for each piece, and a name. What the pack cost is spread over the pieces so profit stays right. All of it: the item changes in place, and this sale\'s order for it changes with it. Some of it: the pieces become their own item and the rest stay. It can not be done once the item has an after count for this sale; close out the sale first. The Saved screen has Undo. Print a new count sheet or price sheet afterward if one was already printed.',
  '- Uploads tab (admins only): every count sheet photo, deposit slip, Amazon order and Cash App file uploaded since Nov 7, 2026 is kept there, newest first, with who uploaded it, when, for which sale, and whether it was saved. Tap one to see it full size (tap the picture to zoom), then "Save or share" to keep a copy or print it, or "Delete this upload". Deleting only removes the picture; the numbers it saved stay. Uploads from before Nov 7, 2026 were not kept. "See past uploads" on the Upload screen opens the same page.',
  '- "Edit sale details" on the Overview card (admin) changes the sale name, date and color.',
  '- PINs & access button (admin) changes the PINs and the access code.'
].join('\n');
var ASK_TABS = ['overview', 'count', 'money', 'cash', 'orders', 'history', 'activity'];
/* Run from the script editor to check the Ask button end to end with the real AI (nothing is saved). */
function selfTestAsk() {
  var ctx = { sale: { name: 'Pink Out Game', soldByCount: 1157.5, moneyTakenIn: 1149.28, afterCounts: '47 of 48 items counted' },
              cashBox: { startingCash: 150, endingCash: 1113.93, cashApp: 185.35, cashTakenIn: 963.93, overShort: -8.22 } };
  var r = ask({ q: 'Why is Sold by count higher than the money we took in?', name: 'Self test', ctx: ctx }, 'admin', true);
  Logger.log(JSON.stringify(r));
  return r;
}
function ask(body, role, test, quiet) {
  if (role !== 'admin') return { ok: false, error: 'Ask is for admins. Unlock with the admin PIN.' };
  var q = String(body.q || '').trim().slice(0, 600);
  if (!q) return { ok: false, error: 'Type a question first.' };
  var c = CacheService.getScriptCache(), n = Number(c.get('askN') || 0);
  if (n >= 40) return { ok: false, error: 'That is a lot of questions for one hour. Try again a little later.' };
  try { c.put('askN', String(n + 1), 3600); } catch (e) {}
  var ctx = '';
  try { ctx = JSON.stringify(body.ctx || {}).slice(0, 16000); } catch (e) {}
  var prompt = [
    'You are the "Ask" button inside the Creekside Lions PTO spirit store tracker (Creekside Elementary, Livingston ISD). The person asking is an admin volunteer, not technical, usually on a phone.',
    HELP_GUIDE,
    'Live numbers from the tracker right now (JSON): ' + ctx,
    'Question from ' + String(body.name || 'an admin').slice(0, 40) + ': ' + q,
    'Answer rules: 1 to 5 short sentences in plain words. Use the live numbers and show the simple math when it helps (for example "$963.93 + $185.35 = $1,149.28"). For how-to questions give the taps in order, using the exact button and tab names. Never invent numbers; if something is not in the numbers, say so. No dashes as punctuation, no markdown, no bullet symbols.',
    'If the question asks for a change to the tracker itself (a new feature, something broken, something worded confusingly), say that would be a change to the tracker and to text Hunter the idea.',
    'Return ONLY one JSON object, no other text: {"answer":"...","goTo":"one of overview, count, money, cash, orders, history, activity, or empty if no tab helps"}'
  ].join('\n\n');
  var r = callClaude([], prompt, { maxTokens: 1200, rawOk: true });
  if (!r.ok) return r;
  var o = r.result || {};
  var answer = String(o.answer || '').trim().slice(0, 2500);
  if (!answer) return { ok: false, error: "Couldn't get an answer. Try asking another way." };
  var goTo = ASK_TABS.indexOf(String(o.goTo || '').toLowerCase()) >= 0 ? String(o.goTo).toLowerCase() : '';
  if (!test) { logAsk(body.name, q, answer); if (!quiet) logVisit(body, 'ask', '', role); }
  return { ok: true, answer: answer, goTo: goTo };
}
var ASK_TAB = 'Ask Log', ASK_COLS = ['Time', 'Who', 'Question', 'Answer'];
function cellText(v, max) { var t = String(v || '').trim().slice(0, max); return /^[=+\-@']/.test(t) ? "'" + t : t; }
function logAsk(who, q, answer) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(ASK_TAB);
    if (!sh) { sh = ss.insertSheet(ASK_TAB); sh.getRange(1, 1, 1, ASK_COLS.length).setValues([ASK_COLS]).setFontWeight('bold'); sh.setFrozenRows(1); }
    ensureRows(sh, sh.getLastRow() + 1);
    var rng = sh.getRange(sh.getLastRow() + 1, 1, 1, ASK_COLS.length);
    rng.setNumberFormat('@');
    rng.setValues([[new Date().toISOString(), cellText(who || 'Someone', 40), cellText(q, 1000), cellText(answer, 2500)]]);
  } catch (e) {} finally { lock.releaseLock(); }
}

/* ------------------------------ Usage (super admin only) ------------------------------
   One row per thing that happens: someone opens the tracker, unlocks it, looks around, prints, saves, asks.
   "Visit" groups one sitting; "Device" is a random ID the page keeps on that phone or computer. No locations,
   no IP addresses: just the name people typed when they unlocked, if they ever did. */
var VISIT_TAB = 'Visits', VISIT_COLS = ['Time', 'Visit', 'Device ID', 'Device', 'Name', 'Role', 'What', 'Details'];
var VISIT_KINDS = { open: 1, tabs: 1, print: 1, unlock: 1, wrongpin: 1, save: 1, ask: 1 };
function visitSheet() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(VISIT_TAB);
  if (!sh) { sh = ss.insertSheet(VISIT_TAB); sh.getRange(1, 1, 1, VISIT_COLS.length).setValues([VISIT_COLS]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function logVisit(body, kind, details, role) {
  try {
    if (!VISIT_KINDS[kind] || body.me) return; // the owner's own devices stay out of the counts
    var sh = visitSheet();
    sh.appendRow([new Date().toISOString(), cellText(body.sid, 24), cellText(body.dev, 24), cellText(body.device, 40), cellText(body.name, 40), cellText(role || '', 12), kind, cellText(details, 200)]);
    if (sh.getLastRow() > 6000) sh.deleteRows(2, 2000); // keep the tab small: the newest 4,000 or so
  } catch (e) {}
}
function ping(body) {
  var c = CacheService.getScriptCache(), k = 'ping_' + String(body.dev || 'x').slice(0, 24), n = Number(c.get(k) || 0);
  if (n > 80) return { ok: true }; // one device can't flood the log
  try { c.put(k, String(n + 1), 3600); } catch (e) {}
  var evs = Array.isArray(body.ev) ? body.ev.slice(0, 6) : [];
  var role = /^(volunteer|admin)$/.test(String(body.role)) ? String(body.role) : ''; // just a label for the usage page; nothing is unlocked by it
  evs.forEach(function (e) { if (e && (e.k === 'open' || e.k === 'tabs' || e.k === 'print')) logVisit(body, e.k, e.d, role); });
  return { ok: true };
}
function usage() {
  var sh = SpreadsheetApp.getActive().getSheetByName(VISIT_TAB);
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };
  var last = sh.getLastRow(), first = Math.max(2, last - 3999);
  var since = new Date(Date.now() - 62 * 864e5).toISOString();
  var rows = sh.getRange(first, 1, last - first + 1, VISIT_COLS.length).getValues().filter(function (r) { return String(r[0]) >= since; })
    .map(function (r) { return r.map(function (v) { return String(v); }); });
  return { ok: true, rows: rows };
}

/* ------------------------------ Price sheet photos ------------------------------
   One small JPEG per item on a "Photos" tab (item ID, picture). The starter pictures ship with the website;
   a photo saved here replaces the starter one, and "none" means print that item without a picture. */
var PHOTO_TAB = 'Photos';
function readPhotos() {
  var sh = SpreadsheetApp.getActive().getSheetByName(PHOTO_TAB), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0]) out[String(r[0])] = String(r[1]); });
  return out;
}
function setPhoto(body, role) {
  if (role !== 'admin') return { ok: false, error: 'Photos are for admins. Unlock with the admin PIN.' };
  var id = String(body.id || ''), pic = String(body.data || '');
  if (!/^[A-Za-z0-9_-]{1,40}$/.test(id)) return { ok: false, error: 'Bad request.' };
  if (pic && pic !== 'none' && !(/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(pic) && pic.length < 48000)) return { ok: false, error: 'That picture is too big or not a photo. Try another one.' };
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'The tracker is busy. Try again in a moment.' };
  try {
    var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(PHOTO_TAB);
    if (!sh) { sh = ss.insertSheet(PHOTO_TAB); sh.getRange(1, 1, 1, 2).setValues([['Item ID', 'Photo (used on the price sheet)']]).setFontWeight('bold'); sh.setFrozenRows(1); }
    var n = sh.getLastRow(), ids = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    var at = ids.indexOf(id), had = at >= 0;
    if (had) { sh.getRange(at + 2, 1, 1, 2).setValues([[id, pic]]); }
    else { ensureRows(sh, n + 1); var rng = sh.getRange(n + 1, 1, 1, 2); rng.setNumberFormat('@'); rng.setValues([[id, pic]]); }
    appendLog([{ time: new Date().toISOString(), name: String(body.name || 'Someone').slice(0, 40), role: role, table: 'photos', rowId: id, field: 'photo',
                 old: had ? 'photo' : '', new: pic === 'none' ? 'no photo' : pic ? 'new photo' : 'starter photo' }]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ------------------------------ Past uploads ------------------------------
   Every photo, PDF and Cash App file that gets uploaded is kept, so the paper is never the only copy.
   They go in a second Google Sheet the tracker makes on its own ("Creekside Spirit Store Uploads"), so this
   one stays quick. One row per file: who, when, what it was, a small picture for the list, and the file itself
   split across the cells to the right (one cell holds 50,000 characters). When that Sheet gets big, the tracker
   starts another one and keeps reading the old ones. Anyone with a PIN adds to it; only admins open or delete. */
var UP_COLS = ['File ID', 'Upload ID', 'Time', 'Who', 'Role', 'What', 'Sale', 'Page', 'Pages', 'Type', 'File name', 'Summary', 'Status', 'Size', 'Small picture', 'Parts'];
var UP_C = { id: 0, up: 1, t: 2, who: 3, role: 4, kind: 5, sale: 6, page: 7, pages: 8, type: 9, name: 10, sum: 11, st: 12, size: 13, thumb: 14, parts: 15 };
var UP_META = UP_COLS.length, UP_CHUNK = 40000, UP_BOOK_LIMIT = 20000000, UP_FILE_MAX = 3200000;
var UP_KINDS = { countSheet: 'Count sheet', order: 'Amazon order', deposit: 'Deposit slip', cashapp: 'Cash App file', bags: 'Mystery bag list' };
var UP_TYPES = { 'image/jpeg': 1, 'image/png': 1, 'image/webp': 1, 'image/gif': 1, 'application/pdf': 1, 'text/csv': 1 };
var UP_ID = /^[a-z0-9]{4,20}$/;

function upSafe(fn, body, role) {
  try { return fn(body, role); }
  catch (e) { return { ok: false, error: 'Past uploads are not working right now. Try again in a minute. (' + String(e && e.message || e).slice(0, 80) + ')' }; }
}
function upText(v, max) { return String(v === undefined || v === null ? '' : v).replace(/[\u0000-\u001f\u007f]+/g, ' ').trim().slice(0, max); }
function upBooks() {
  try { var a = JSON.parse(PropertiesService.getScriptProperties().getProperty('UPLOAD_BOOKS') || '[]'); return Array.isArray(a) ? a.filter(function (x) { return typeof x === 'string' && x; }) : []; }
  catch (e) { return []; }
}
function upSheet(id) { var ss = SpreadsheetApp.openById(id); return ss.getSheetByName('Uploads') || ss.getSheets()[0]; }
function upNewBook(n) {
  var ss = SpreadsheetApp.create('Creekside Spirit Store Uploads' + (n > 1 ? ' ' + n : ''));
  var sh = ss.getSheets()[0];
  sh.setName('Uploads');
  sh.getRange(1, 1, 1, UP_META + 1).setValues([UP_COLS.concat(['The file, split across the cells to the right. Open it from the Uploads tab in the tracker.'])]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return ss.getId();
}
/* the rows of one upload, looking in the newest Sheet first */
function upFind(uid, maxBooks) {
  var books = upBooks().reverse().slice(0, maxBooks || 999);
  for (var b = 0; b < books.length; b++) {
    var sh; try { sh = upSheet(books[b]); } catch (e) { continue; }
    var last = sh.getLastRow(), out = [];
    if (last < 2) continue;
    var ids = sh.getRange(2, UP_C.up + 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === uid) out.push({ sh: sh, row: i + 2, meta: sh.getRange(i + 2, 1, 1, UP_META).getValues()[0] });
    if (out.length) return out;
  }
  return [];
}

function keepUpload(body, role) {
  if (role !== 'admin' && role !== 'volunteer') return { ok: false, error: 'Unlock with a PIN first.' };
  var uid = String(body.uid || ''), kind = String(body.kind || '');
  if (!UP_ID.test(uid) || !Object.prototype.hasOwnProperty.call(UP_KINDS, kind)) return { ok: false, error: 'Bad request.' };
  var files = (Array.isArray(body.files) ? body.files : []).slice(0, 8).filter(function (f) {
    return f && UP_TYPES[f.type] === 1 && typeof f.data === 'string' && f.data.length > 0 && f.data.length <= UP_FILE_MAX && /^[A-Za-z0-9+\/]+={0,2}$/.test(f.data)
      && (!f.thumb || (typeof f.thumb === 'string' && f.thumb.length <= 40000 && /^data:image\/jpeg;base64,[A-Za-z0-9+\/]+={0,2}$/.test(f.thumb)));
  });
  if (!files.length) return { ok: false, error: 'Nothing to keep.' };
  var who = upText(body.name || 'Someone', 40), sale = upText(body.sale, 80), sum = upText(body.sum, 160), stamp;
  var chars = 0, maxParts = 0;
  files.forEach(function (f) { chars += f.data.length + (f.thumb || '').length; maxParts = Math.max(maxParts, Math.ceil(f.data.length / UP_CHUNK)); });
  var sh, first, lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'The tracker is busy. Try again in a moment.' };
  try {
    var props = PropertiesService.getScriptProperties(), books = upBooks(), used = Number(props.getProperty('UPLOAD_CHARS')) || 0;
    if (books.length) { try { sh = upSheet(books[books.length - 1]); } catch (e) { sh = null; } }
    if (!sh || (used > 0 && used + chars > UP_BOOK_LIMIT)) { // first upload, the Sheet is full, or someone deleted it
      books.push(upNewBook(books.length + 1)); used = 0;
      props.setProperty('UPLOAD_BOOKS', JSON.stringify(books));
      sh = upSheet(books[books.length - 1]);
    }
    var last = sh.getLastRow();
    if (last > 1) { // the same upload sent twice is kept once
      var lo = Math.max(2, last - 199);
      if (sh.getRange(lo, UP_C.up + 1, last - lo + 1, 1).getValues().some(function (r) { return String(r[0]) === uid; })) return { ok: true, kept: 0 };
    }
    // every upload gets its own time, a millisecond apart at least, so paging through the list never skips one
    var now = new Date(), lastT = props.getProperty('UPLOAD_LAST_T') || '';
    if (lastT && now.toISOString() <= lastT) now = new Date(Date.parse(lastT) + 1);
    stamp = now.toISOString();
    props.setProperty('UPLOAD_LAST_T', stamp);
    first = last + 1;
    // rows and columns are only ever added at the end, and only while holding the lock, so nothing already saved moves
    ensureRows(sh, first + files.length - 1);
    ensureCols(sh, UP_META + maxParts);
    var meta = sh.getRange(first, 1, files.length, UP_META);
    meta.setNumberFormat('@');
    meta.setValues(files.map(function (f, i) {
      return [uid + '-' + (i + 1), uid, stamp, who, role, kind, sale, String(i + 1), String(files.length), f.type, upText(f.name, 80), sum, 'writing', String(f.data.length), f.thumb || '', '0'];
    }));
    props.setProperty('UPLOAD_CHARS', String(used + chars));
  } finally { lock.releaseLock(); }
  // the big part is written after letting go of the lock, so a save never waits behind a photo
  files.forEach(function (f, i) {
    var parts = [];
    for (var p = 0; p < f.data.length; p += UP_CHUNK) parts.push(f.data.slice(p, p + UP_CHUNK));
    var d = sh.getRange(first + i, UP_META + 1, 1, parts.length);
    d.setNumberFormat('@');
    d.setValues([parts]);
    sh.getRange(first + i, UP_C.st + 1, 1, UP_META - UP_C.st).setValues([['kept', String(f.data.length), f.thumb || '', String(parts.length)]]);
  });
  return { ok: true, kept: files.length };
}

/* after the upload is saved (or undone), so the list can say so */
function markUpload(body, role) {
  if (role !== 'admin' && role !== 'volunteer') return { ok: false, error: 'Unlock with a PIN first.' };
  var uid = String(body.uid || ''), st = String(body.st || '');
  if (!UP_ID.test(uid) || (st !== 'saved' && st !== 'undone')) return { ok: false, error: 'Bad request.' };
  var rows = upFind(uid, 2).filter(function (r) { return String(r.meta[UP_C.st]) !== 'writing'; });
  rows.forEach(function (r) { r.sh.getRange(r.row, UP_C.st + 1, 1, 1).setValues([[st]]); });
  return { ok: true, marked: rows.length };
}

/* newest first, 30 uploads at a time (all pages of an upload stay together) */
function listUploads(body, role) {
  if (role !== 'admin') return { ok: false, error: 'Only the admin PIN can open past uploads.' };
  var before = String(body.before || ''), want = 30, files = [], seen = {}, nUp = 0, more = false;
  var books = upBooks().reverse();
  outer:
  for (var b = 0; b < books.length; b++) {
    var sh; try { sh = upSheet(books[b]); } catch (e) { continue; }
    var last = sh.getLastRow();
    if (last < 2) continue;
    var rows = sh.getRange(2, 1, last - 1, UP_META).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      var r = rows[i], up = String(r[UP_C.up]), t = String(r[UP_C.t]), st = String(r[UP_C.st]);
      if (!r[UP_C.id] || st === 'writing' || (before && !(t < before))) continue;
      if (!seen[up]) { if (nUp >= want) { more = true; break outer; } seen[up] = 1; nUp++; }
      files.push({ id: String(r[UP_C.id]), up: up, t: t, who: String(r[UP_C.who]), role: String(r[UP_C.role]), kind: String(r[UP_C.kind]), sale: String(r[UP_C.sale]),
        page: Number(r[UP_C.page]) || 1, pages: Number(r[UP_C.pages]) || 1, type: String(r[UP_C.type]), name: String(r[UP_C.name]), sum: String(r[UP_C.sum]),
        st: st, size: Number(r[UP_C.size]) || 0, thumb: String(r[UP_C.thumb]) });
    }
  }
  files.sort(function (x, y) { return x.t < y.t ? 1 : x.t > y.t ? -1 : x.up < y.up ? 1 : x.up > y.up ? -1 : x.page - y.page; });
  return { ok: true, files: files, more: more };
}

function getUpload(body, role) {
  if (role !== 'admin') return { ok: false, error: 'Only the admin PIN can open past uploads.' };
  var id = String(body.id || ''), m = id.match(/^([a-z0-9]{4,20})-(\d{1,2})$/);
  if (!m) return { ok: false, error: 'Bad request.' };
  var rows = upFind(m[1]);
  for (var k = 0; k < rows.length; k++) {
    var r = rows[k].meta;
    if (String(r[UP_C.id]) !== id) continue;
    var parts = Number(r[UP_C.parts]) || 0, size = Number(r[UP_C.size]) || 0;
    var data = parts && String(r[UP_C.st]) !== 'writing' ? rows[k].sh.getRange(rows[k].row, UP_META + 1, 1, parts).getValues()[0].map(String).join('') : '';
    if (!data || data.length !== size) return { ok: false, error: 'This file did not finish saving, so it can not be opened.' };
    return { ok: true, type: String(r[UP_C.type]), name: String(r[UP_C.name]), data: data };
  }
  return { ok: false, error: 'That upload is not there anymore.' };
}

/* deletes every page of one upload; the numbers it saved stay in the tracker */
function delUpload(body, role, quiet) {
  if (role !== 'admin') return { ok: false, error: 'Only the admin PIN can delete uploads.' };
  var uid = String(body.uid || '');
  if (!UP_ID.test(uid)) return { ok: false, error: 'Bad request.' };
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'The tracker is busy. Try again in a moment.' };
  try {
    var rows = upFind(uid);
    if (!rows.length) return { ok: true, deleted: 0 };
    var m = rows[0].meta;
    rows.forEach(function (r) { r.sh.getRange(r.row, 1, 1, r.sh.getMaxColumns()).clearContent(); }); // cleared, not removed, so no other row moves
    if (!quiet) {
      var when = Utilities.formatDate(new Date(String(m[UP_C.t])), sheetTz(), 'MMM d, yyyy h:mm a');
      appendLog([{ time: new Date().toISOString(), name: upText(body.name || 'Someone', 40), role: role, table: 'uploads', rowId: uid, field: '(removed)',
                   old: (UP_KINDS[String(m[UP_C.kind])] || 'Upload') + (rows.length > 1 ? ' (' + rows.length + ' pages)' : '') + ' from ' + when + ' by ' + String(m[UP_C.who]), new: '' }]);
    }
    return { ok: true, deleted: rows.length };
  } finally { lock.releaseLock(); }
}

/* Run from the script editor to check past uploads from start to finish: keeps a test file big enough to
   need several cells, finds it, opens it, checks every character came back, then removes it. */
function selfTestUploads() {
  var big = '';
  while (big.length < 95000) big += 'QUJDREVGR0hJSktMTU5PUFFSU1RVVldYWVowMTIzNDU2Nzg5K/+';
  big = big.slice(0, 95000);
  var uid = 'selftest' + Date.now().toString(36).slice(-6);
  var k = keepUpload({ uid: uid, kind: 'deposit', name: 'Self test', sale: 'Self test', sum: 'Self test', files: [{ type: 'image/jpeg', name: 'test.jpg', data: big, thumb: '' }] }, 'admin');
  var l = listUploads({}, 'admin'), found = !!(l.ok && l.files.some(function (f) { return f.up === uid; }));
  var g = getUpload({ id: uid + '-1' }, 'admin'), same = !!(g.ok && g.data === big);
  var d = delUpload({ uid: uid }, 'admin', true);
  var gone = !listUploads({}, 'admin').files.some(function (f) { return f.up === uid; });
  var good = k.ok && k.kept === 1 && found && same && d.ok && d.deleted === 1 && gone;
  var msg = good ? 'Past uploads work. A test file of ' + big.length + ' characters (' + Math.ceil(big.length / UP_CHUNK) + ' cells) was kept, listed, opened in full and removed.'
                 : 'Problem: ' + JSON.stringify({ keep: k, found: found, same: same, del: d, gone: gone }).slice(0, 400);
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

/* One-time update, Oct 2026: adds the items on the paper count sheet that were not in the tracker yet.
   Run it from the script editor. Safe to run again: it skips anything already there. */
function addCountSheetItems() {
  var want = [
    { key: /cotton.*tee/i, name: 'Cotton spirit tees (all designs)', price: 20, onHand: 9, notes: 'From the paper count sheet' },
    { key: /dri.?fit/i, name: 'Dri-fit spirit shirts (all designs)', price: 25, onHand: 3, notes: 'From the paper count sheet' },
    { key: /\bhat\b/i, name: 'Lion spirit hat', price: 10, onHand: 1, notes: 'From the paper count sheet' },
    { key: /croc/i, name: 'Croc shoe charms', price: 3, onHand: '', notes: 'Count before opening (more brought in)' },
    { key: /pink pom ?poms?|pom ?poms?\W+pink/i, name: 'Pink pom poms', price: 2, onHand: '', notes: 'Count, or ask who ordered them. 1 for $2, 2 for $3' }
  ];
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy. Try again in a minute.');
  try {
    var tables = { items: readTable('items').rows }, ops = [], skipped = [];
    want.forEach(function (w, i) {
      if (tables.items.some(function (r) { return w.key.test(String(r.name)); })) { skipped.push(w.name); return; }
      ops.push({ op: 'add', table: 'items', row: { id: 'it' + Date.now().toString(36) + i, name: w.name, category: 'Carry-over', price: w.price, unitCost: '', onHand: w.onHand, countedAfter: '', notes: w.notes } });
    });
    var paint = tables.items.filter(function (r) { return /lion paint kit/i.test(String(r.name)); })[0];
    if (paint && !String(paint.notes || '').trim()) ops.push({ op: 'set', table: 'items', id: paint.id, field: 'notes', prev: paint.notes, value: 'Recount before the sale (more brought in)' });
    var res = applyOps(tables, ops, 'admin', 'Count sheet update', new Date().toISOString(), true);
    if (res.applied) { writeTable('items', tables.items); appendLog(res.log); }
    var msg = 'Added ' + ops.filter(function (o) { return o.op === 'add'; }).length + ' items' + (skipped.length ? '. Already there: ' + skipped.join(', ') : '') + (res.conflicts.length ? '. Problems: ' + JSON.stringify(res.conflicts) : '');
    Logger.log(msg);
    return msg;
  } finally { lock.releaseLock(); }
}

/* One-time correction, Nov 7, 2026: the pink paw print glitter eye stickers came as 120 pairs and were cut into
   single paws. Brittany counted 196 and sells each one for 50 cents. Same result as Count > "Cut up or split":
   the count becomes 196, the name says single paws, and the $13.99 they cost is spread over the 196 pieces.
   Run it once from the script editor. It checks the numbers first and changes nothing if they already moved. */
function fixPawStickersNov7() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy. Try again in a minute.');
  try {
    var tables = { items: readTable('items').rows };
    var it = tables.items.filter(function (r) { return String(r.id) === 'it16'; })[0], msg;
    if (!it) msg = 'Not changed: the pink paw sticker item (it16) is not in the Items tab.';
    else if (String(it.name) !== 'Pink paw print glitter eye stickers (120 pairs)' || Number(it.onHand) !== 120 || String(it.countedAfter) !== '')
      msg = 'Not changed, because it no longer matches what was checked on Nov 7 (name "' + it.name + '", on hand ' + it.onHand + ', after count "' + it.countedAfter + '"). Use Count > Cut up or split instead.';
    else {
      var cost = Number(it.unitCost) > 0 ? Math.round(Number(it.unitCost) * 120 / 196 * 10000) / 10000 : it.unitCost;
      var ops = [
        { op: 'set', table: 'items', id: 'it16', field: 'onHand', prev: it.onHand, value: 196 },
        { op: 'set', table: 'items', id: 'it16', field: 'name', prev: it.name, value: 'Pink paw print glitter eye stickers (single paws)' },
        { op: 'set', table: 'items', id: 'it16', field: 'unitCost', prev: it.unitCost, value: cost },
        { op: 'set', table: 'items', id: 'it16', field: 'notes', prev: it.notes, value: (String(it.notes || '').trim() ? String(it.notes).trim() + ' · ' : '') + 'Cut up Nov 7, 2026: 120 pairs became 196 single paws' }
      ];
      if (Number(it.price) !== 0.5) ops.push({ op: 'set', table: 'items', id: 'it16', field: 'price', prev: it.price, value: 0.5 });
      var res = applyOps(tables, ops, 'admin', 'Hunter (correction)', new Date().toISOString(), true);
      if (res.applied) { writeTable('items', tables.items); appendLog(res.log); }
      msg = res.applied ? 'Done. Pink paw print glitter eye stickers (single paws): 196 on hand, 50 cents each, cost ' + cost + ' each.' : 'Not changed: ' + JSON.stringify(res.conflicts);
    }
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
    return msg;
  } finally { lock.releaseLock(); }
}

/* One-time entry, Nov 8, 2026: the Oct 30 Halloween mystery bags from Brittany's list. 17 bags at $5 and 12 at $3,
   filled with 114 items. Where her list had more of an item than the count showed gone, the matching similar item
   that sold out makes up the rest (5 mini mochi, 3 pumpkin paint kits, 2 glitter cubes). The 4-in-1 pens stay at 21
   as written, so the tracker flags that count to check. Same result as Count > Mystery bags. Run once from the
   script editor. It checks everything first and changes nothing if the tracker no longer matches. */
var OCT30_BAG_PARTS = [
  ['itmvg56r7ju14', 'Halloween mochi squishies (60-pack)', 24], ['it37', 'Mini mochi squishies', 5],
  ['it15', 'Stretchy skeletons (48-pack)', 12], ['itmvkgke2wu16', 'Halloween paint kits', 6], ['itmvab30u7u1', 'Pumpkin paint kit $4', 3],
  ['itmvg56r7ju18', 'Bone skeleton pens (100-pack)', 6], ['itmvg56r7ju11', 'Witch broom pencils (24-pack)', 6],
  ['it12', 'Halloween 4-in-1 pens (50-pack)', 21], ['itmvg56r7ju5', 'Halloween cube squishies (12-pack)', 29], ['it27', 'Glitter cube squishies', 2]
];
function fixMysteryBagsOct30() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy. Try again in a minute.');
  try {
    var tables = { items: readTable('items').rows, settings: readTable('settings').rows }, msg = '';
    var byId = {}; tables.items.forEach(function (r) { byId[String(r.id)] = r; });
    var setting = function (id) { var r = tables.settings.filter(function (x) { return String(x.id) === id; })[0]; return r ? r : null; };
    var ev = setting('eventName'), bun = setting('bundles'), raw = {};
    try { raw = bun && String(bun.value).trim() ? JSON.parse(String(bun.value)) : {}; } catch (e) { raw = null; }
    if (!ev || String(ev.value) !== 'Spirit Store 10/30/2026') msg = 'Not changed: the current sale is "' + (ev ? ev.value : '') + '", not Spirit Store 10/30/2026.';
    else if (!raw || typeof raw !== 'object' || Array.isArray(raw)) msg = 'Not changed: the bundles setting could not be read.';
    else if (raw.m && raw.m.mbhalloween1030) msg = 'Not changed: the Oct 30 Halloween bags are already entered.';
    else if (byId.itmb1030d5 || byId.itmb1030d3 || tables.items.some(function (r) { return /^halloween mystery bag/i.test(String(r.name)); })) msg = 'Not changed: there is already a Halloween mystery bag item. Use Count > Mystery bags instead.';
    else {
      var missing = OCT30_BAG_PARTS.filter(function (p) { return !byId[p[0]] || String(byId[p[0]].name) !== p[1]; });
      if (missing.length) msg = 'Not changed: these items are not as expected: ' + missing.map(function (p) { return p[1]; }).join(', ') + '. Use Count > Mystery bags instead.';
    }
    if (!msg) {
      raw.m = raw.m && typeof raw.m === 'object' && !Array.isArray(raw.m) ? raw.m : {};
      raw.m.mbhalloween1030 = { name: 'Halloween mystery bags', bags: ['itmb1030d5', 'itmb1030d3'], parts: OCT30_BAG_PARTS.map(function (p) { return [p[0], p[2]]; }), done: '', at: '2026-10-30' };
      var ops = [
        { op: 'add', table: 'items', row: { id: 'itmb1030d5', name: 'Halloween mystery bag ($5)', category: 'New', price: 5, unitCost: '', onHand: 17, countedAfter: 0, notes: 'Mystery bag. 17 sold at $5 (Brittany\'s list, Oct 30, 2026)' } },
        { op: 'add', table: 'items', row: { id: 'itmb1030d3', name: 'Halloween mystery bag ($3)', category: 'New', price: 3, unitCost: '', onHand: 12, countedAfter: 0, notes: 'Mystery bag. 12 sold at $3 (Brittany\'s list, Oct 30, 2026)' } },
        bun ? { op: 'set', table: 'settings', id: 'bundles', field: 'value', prev: bun.value, value: JSON.stringify(raw) } : { op: 'add', table: 'settings', row: { id: 'bundles', value: JSON.stringify(raw) } }
      ];
      var res = applyOps(tables, ops, 'admin', 'Hunter (correction)', new Date().toISOString(), true);
      if (res.applied) { writeTable('items', tables.items); writeTable('settings', tables.settings); appendLog(res.log); }
      msg = res.applied === ops.length ? 'Done. Halloween mystery bags for Oct 30: 17 at $5 and 12 at $3, with the 114 items from the list taken out of their own sales.' : 'Not changed: ' + JSON.stringify(res.conflicts);
    }
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
    return msg;
  } finally { lock.releaseLock(); }
}

/* Run from the script editor to test reading a filled-in count sheet photo end to end. */
function selfTestCountImage() {
  var blob = UrlFetchApp.fetch('https://raw.githubusercontent.com/creekside-pto/spirit-store/main/sample-count.png').getBlob();
  var rows = JSON.parse(UrlFetchApp.fetch('https://raw.githubusercontent.com/creekside-pto/spirit-store/main/sample-count-rows.json').getContentText());
  var names = readTable('items').rows.map(function (r) { return String(r.name); });
  var t = readDoc({ images: [{ media_type: 'image/png', data: Utilities.base64Encode(blob.getBytes()) }], names: names, rows: rows }, 'admin');
  Logger.log(JSON.stringify(t));
  return t;
}

/* Run from the script editor to test reading a real screenshot end to end. */
function selfTestAiImage() {
  var blob = UrlFetchApp.fetch('https://raw.githubusercontent.com/creekside-pto/spirit-store/main/sample-order.png').getBlob();
  var names = readTable('items').rows.map(function (r) { return String(r.name); });
  var t = aiExtract([{ media_type: 'image/png', data: Utilities.base64Encode(blob.getBytes()) }], '', names);
  Logger.log(JSON.stringify(t));
  return t;
}

function importRead(body, role) {
  if (role !== 'admin') return { ok: false, error: 'Only the admin PIN can import orders.' };
  var images = (Array.isArray(body.images) ? body.images : []).slice(0, 8).filter(function (im) {
    return im && /^image\/(jpeg|png|webp|gif)$/.test(im.media_type) && typeof im.data === 'string' && im.data.length < 8000000;
  });
  var text = String(body.text || '').slice(0, 20000);
  if (!images.length && !text.trim()) return { ok: false, error: 'Add a screenshot or paste the order text first.' };
  var names = readTable('items').rows.map(function (r) { return String(r.name); });
  return aiExtract(images, text, names);
}

function aiExtract(images, text, names) {
  var instructions = [
    'These are screenshots or text from an Amazon cart, order confirmation, or receipt for a school PTO spirit store.',
    'Existing store items (exact names): ' + JSON.stringify(names),
    'Return ONLY a JSON object, no other text, shaped like:',
    '{"store":"Amazon","orderDate":"YYYY-MM-DD or empty","orderTotal":number or null,"lines":[{"title":"product title as shown","name":"short clean item name with pack size, e.g. Pink ribbon bracelets (24-pack)","match":"exact existing item name if this is the same product, else null","packs":integer quantity ordered,"packCost":price of ONE pack as a number,"unitsPerPack":number of sellable pieces in one pack (24 Pack = 24, 120 pairs = 120, a single item = 1) or null if unclear,"note":"short note if anything is uncertain, else empty"}]}',
    'Rules: one line per distinct product. Skip anything under Saved for later, recommendations, ads, shipping, and tax. Never invent products. If only a line total is shown, divide by quantity to get packCost. Only set match when you are confident it is the same product.'
  ].join('\n');
  var r = callClaude(images, instructions + (text ? '\n\nOrder text:\n' + text : ''));
  if (!r.ok) return r;
  if (!Array.isArray(r.result.lines)) r.result.lines = [];
  return r;
}

function callClaude(images, prompt, opts) {
  opts = opts || {};
  var key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) return { ok: false, error: "The AI reader isn't set up yet. In the Google Sheet: Spirit Store > Set AI key." };
  var content = images.map(function (im) {
    if (im.media_type === 'application/pdf') return { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: im.data } };
    return { type: 'image', source: { type: 'base64', media_type: im.media_type, data: im.data } };
  });
  content.push({ type: 'text', text: prompt });
  var last = '', models = aiModelList(), looked = false;
  for (var i = 0; i < models.length; i++) {
    var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: models[i], max_tokens: opts.maxTokens || 8000, messages: [{ role: 'user', content: content }] })
    });
    var code = resp.getResponseCode(), raw = resp.getContentText();
    if (code === 404 || (code === 400 && /model/i.test(raw) && !/credit/i.test(raw))) {
      last = raw;
      if (i === models.length - 1 && !looked) { looked = true; var found = aiFindModel(key, models); if (found) models.push(found); }
      continue;
    }
    if (code === 200 && i > 0) { try { PropertiesService.getScriptProperties().setProperty('AI_MODEL', models[i]); } catch (e) {} }
    if (code === 401 || code === 403) return { ok: false, error: 'The AI key was rejected. Set a new one: Spirit Store > Set AI key.' };
    if (code === 429 || code === 529 || code >= 500) return { ok: false, error: 'The AI service is busy right now. Try again in a minute.' };
    if (code === 400 && /credit/i.test(raw)) return { ok: false, error: 'The AI account is out of credit. Add credit in the Claude Console.' };
    if (code !== 200) return { ok: false, error: 'The AI reader returned an error (' + code + ').' };
    var out = JSON.parse(raw), txt = (out.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
    var a = txt.indexOf('{'), b = txt.lastIndexOf('}');
    if (a < 0 || b < a) return opts.rawOk && txt.trim() ? { ok: true, result: { answer: txt.trim() }, model: models[i] } : { ok: false, error: opts.rawOk ? "Couldn't get an answer. Try asking another way." : "Couldn't find anything to read. Try a clearer picture." };
    var result;
    try { result = JSON.parse(txt.slice(a, b + 1)); } catch (e) { return opts.rawOk ? { ok: true, result: { answer: txt.trim() }, model: models[i] } : { ok: false, error: "Couldn't read that. Try again or use a clearer picture." }; }
    return { ok: true, result: result, model: models[i] };
  }
  return { ok: false, error: 'No AI model was available. ' + String(last).slice(0, 120) };
}

/* ------------------------------ helpers ------------------------------ */
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function codeOk(code) {
  var need = PropertiesService.getScriptProperties().getProperty('VIEW_CODE');
  return !need || String(code || '').trim() === need;
}

function roleFor(pin) {
  var props = PropertiesService.getScriptProperties();
  var vp = props.getProperty('VOLUNTEER_PIN'), ap = props.getProperty('ADMIN_PIN');
  if (!vp || !ap) return { error: 'PINs are not set yet. In the Google Sheet use Spirit Store > Set PINs.' };
  var cache = CacheService.getScriptCache();
  if (Number(cache.get('fails') || 0) >= 30) return { error: 'Too many wrong PINs. Wait 10 minutes and try again.' };
  pin = String(pin || '').trim();
  var sp = props.getProperty('SUPER_PIN');
  if (pin && sp && pin === sp) return { role: 'super' };
  if (pin && pin === ap) return { role: 'admin' };
  if (pin && pin === vp) return { role: 'volunteer' };
  var lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) { try { cache.put('fails', String(Number(cache.get('fails') || 0) + 1), 600); } finally { lock.releaseLock(); } }
  return { error: 'That PIN did not work.' };
}

var SHEET_TZ = null;
function sheetTz() { return SHEET_TZ || (SHEET_TZ = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone()); }
function fmtCell(v, t) {
  if (v instanceof Date) return t === 'log' ? v.toISOString() : Utilities.formatDate(v, sheetTz(), 'yyyy-MM-dd');
  return v;
}

function readTable(t, tail) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SCHEMA[t].tab);
  if (!sh) throw new Error('Missing tab "' + SCHEMA[t].tab + '". Run Spirit Store > Set up tabs.');
  var values;
  if (tail) { // only the newest rows (the change log grows forever)
    var last = sh.getLastRow(), lc = Math.max(sh.getLastColumn(), 1), from = Math.max(2, last - tail + 1);
    values = sh.getRange(1, 1, 1, lc).getValues().concat(last >= 2 ? sh.getRange(from, 1, last - from + 1, lc).getValues() : []);
  } else values = sh.getDataRange().getValues();
  var header = normHeader(values[0] || []), idx = {};
  SCHEMA[t].cols.forEach(function (c) {
    idx[c[0]] = header.indexOf(c[1].toLowerCase());
    if (idx[c[0]] < 0) throw new Error('The "' + SCHEMA[t].tab + '" tab is missing its "' + c[1] + '" column. Put that header back in row 1.');
  });
  var rows = [], fixed = false;
  for (var r = 1; r < values.length; r++) {
    var row = {}, any = false;
    SCHEMA[t].cols.forEach(function (c) {
      var i = idx[c[0]], v = i >= 0 ? fmtCell(values[r][i], t) : '';
      row[c[0]] = v === null || v === undefined ? '' : v;
      if (row[c[0]] !== '') any = true;
    });
    if (!any) continue;
    if (t !== 'log' && row.id === '') { row.id = t.slice(0, 2) + Date.now().toString(36) + r; fixed = true; }
    rows.push(row);
  }
  if (t === 'settings') rows = joinParts(rows);
  return { rows: rows, sheet: sh, idx: idx, fixed: fixed };
}

/* One cell holds 50,000 characters. A setting that grows past that over the years (the price sheet setup lists
   every item) is split across extra rows named like "priceSheet~part2" and put back together when read, so the
   rest of the tracker only ever sees one value. Parts are close to equal in size so none is a short scrap. */
var PART_MAX = 40000, PART_TAG = '~part';
function splitParts(rows) {
  var out = [];
  rows.forEach(function (r) {
    var v = r.value, s = typeof v === 'string' ? v : String(v === undefined || v === null ? '' : v);
    if (String(r.id).indexOf(PART_TAG) >= 0) return; // never write an old part back on its own
    if (typeof v !== 'string' || s.length <= PART_MAX) { out.push(r); return; }
    var n = Math.ceil(s.length / PART_MAX), size = Math.ceil(s.length / n);
    for (var k = 0; k < n; k++) out.push({ id: k === 0 ? r.id : r.id + PART_TAG + (k + 1), value: s.slice(k * size, (k + 1) * size) });
  });
  return out;
}
function joinParts(rows) {
  var parts = {}, out = [];
  rows.forEach(function (r) {
    var m = String(r.id).match(/^(.*)~part(\d+)$/);
    if (m) (parts[m[1]] = parts[m[1]] || []).push([Number(m[2]), String(r.value)]); else out.push(r);
  });
  out.forEach(function (r) {
    var p = parts[String(r.id)];
    if (p) r.value = String(r.value) + p.sort(function (a, b) { return a[0] - b[0]; }).map(function (x) { return x[1]; }).join('');
  });
  return out;
}

function normHeader(h) { return h.map(function (x) { return String(x).trim().toLowerCase(); }); }
/* Keep text as text: stop Sheets from turning notes like "10/16" into dates and "=..." into formulas. */
function sheetSafe(t, field, v, id) {
  if (typeof v !== 'string' || v === '') return v;
  if (isNumeric(t, field, id)) return v;
  if (field === 'date' || (t === 'settings' && String(id) === 'eventDate')) return v;
  if (/^[=+\-@']/.test(v) || /^[\d\s\/.:,$()%-]+$/.test(v)) return "'" + v;
  return v;
}
function writeTable(t, rows) {
  dropCache();
  if (t === 'settings') rows = splitParts(rows);
  var sh = SpreadsheetApp.getActive().getSheetByName(SCHEMA[t].tab);
  var header = normHeader(sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]);
  var height = Math.max(sh.getLastRow() - 1, rows.length, 1);
  ensureRows(sh, height + 1);
  SCHEMA[t].cols.forEach(function (c) {
    var col = header.indexOf(c[1].toLowerCase()) + 1;
    if (col < 1) throw new Error('The "' + SCHEMA[t].tab + '" tab is missing its "' + c[1] + '" column.');
    var out = [];
    for (var r = 0; r < height; r++) { var v = r < rows.length ? rows[r][c[0]] : ''; v = v === undefined || v === null ? '' : v; out.push([sheetSafe(t, c[0], v, r < rows.length ? rows[r].id : '')]); }
    sh.getRange(2, col, height, 1).setValues(out);
  });
}

/* A tab has a set number of rows (1,000 when it is new) and writing below the last one fails,
   which would stop saves once the Change Log fills up. So rows are added first, 500 spare at a time. */
function ensureRows(sh, lastRowNeeded) {
  var max = sh.getMaxRows();
  if (lastRowNeeded > max) sh.insertRowsAfter(max, lastRowNeeded - max + 500);
}
function ensureCols(sh, lastColNeeded) {
  var max = sh.getMaxColumns();
  if (lastColNeeded > max) sh.insertColumnsAfter(max, lastColNeeded - max);
}

function appendLog(entries) {
  if (!entries || !entries.length) return;
  dropCache();
  var sh = SpreadsheetApp.getActive().getSheetByName(SCHEMA.log.tab);
  var rows = entries.map(function (e) { return SCHEMA.log.cols.map(function (c) { var v = e[c[0]] === undefined ? '' : String(e[c[0]]); return v.length > PART_MAX ? v.slice(0, PART_MAX) + ' (cut short)' : v; }); });
  // a new tab has 26 columns but the log uses 8; the empty ones still count toward Google's 10 million cells per Sheet
  var cols = rows[0].length;
  if (sh.getMaxColumns() > cols + 2 && sh.getLastColumn() <= cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
  ensureRows(sh, sh.getLastRow() + rows.length);
  var rng = sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length);
  rng.setNumberFormat('@'); // keep log entries as plain text so dates and numbers are not reinterpreted
  rng.setValues(rows);
}

var CASH_ROWS = [['start', 'Starting cash in box (change)'], ['end', 'Ending cash in box'], ['cashapp', 'Cash App total for this sale'], ['countedBy', 'Counted by'], ['notes', 'Notes']];
/* Fast loads: the whole dataset is kept in the script cache for a minute and dropped on every save,
   and on hand edits to the Google Sheet (onEdit below). Chunked because one cache entry holds 100KB at most. */
var CACHE_KEY = 'data2';
function dataJson() {
  var c = CacheService.getScriptCache();
  try {
    var n = Number(c.get(CACHE_KEY + 'n'));
    if (n > 0) {
      var keys = [];
      for (var i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
      var got = c.getAll(keys), s = '';
      for (var j = 0; j < n; j++) { if (got[keys[j]] === undefined || got[keys[j]] === null) { s = ''; break; } s += got[keys[j]]; }
      if (s) return s;
    }
  } catch (e) {}
  var ver = null;
  try { ver = c.get(CACHE_KEY + 'ver'); } catch (e) {}
  var out = JSON.stringify(loadAll());
  try {
    if (c.get(CACHE_KEY + 'ver') === ver) { // nothing was saved while this was being read
      var parts = {}, k = 0;
      for (var p = 0; p < out.length; p += 20000) parts[CACHE_KEY + '_' + (k++)] = out.slice(p, p + 20000);
      c.putAll(parts, 60); c.put(CACHE_KEY + 'n', String(k), 60);
    }
  } catch (e) {}
  return out;
}
function dropCache() {
  try { var c = CacheService.getScriptCache(); c.put(CACHE_KEY + 'ver', String(Date.now()) + Math.random(), 21600); c.remove(CACHE_KEY + 'n'); } catch (e) {}
}
function onEdit() { dropCache(); }

function loadAll() {
  var out = {};
  Object.keys(SCHEMA).forEach(function (t) {
    if (t === 'log') return;
    var r = readTable(t);
    if (t === 'cash') CASH_ROWS.forEach(function (c) {
      if (!r.rows.some(function (x) { return String(x.id) === c[0]; })) { r.rows.push({ id: c[0], label: c[1], value: '' }); r.fixed = true; }
    });
    if (r.fixed) { // rows typed straight into the sheet get an ID; missing cash lines get added
      var lock = LockService.getScriptLock();
      if (lock.tryLock(10000)) { try { r = readTable(t); if (t === 'cash') CASH_ROWS.forEach(function (c) { if (!r.rows.some(function (x) { return String(x.id) === c[0]; })) r.rows.push({ id: c[0], label: c[1], value: '' }); }); writeTable(t, r.rows); } finally { lock.releaseLock(); } }
    }
    out[t] = r.rows;
  });
  out.settings = out.settings.filter(function (s) { return !/pin|code/i.test(String(s.id)); });
  out.log = readTable('log', 80).rows.reverse();
  return out;
}
