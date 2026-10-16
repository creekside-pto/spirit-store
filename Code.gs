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
  if (role === 'admin') return true;
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


var SEED = null;

/* ------------------------------ Sheet menu ------------------------------ */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Spirit Store')
    .addItem('Set PINs', 'setPins')
    .addItem('Require an access code to view (optional)', 'setViewCode')
    .addItem('Set AI key (for importing order screenshots)', 'setAiKey')
    .addItem('Test the AI order reader', 'selfTestAi')
    .addSeparator()
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
    if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
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
  return ContentService.createTextOutput('{"ok":true,"data":' + dataJson() + '}').setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Bad request.' }); }
  if (!codeOk(body.code)) return json({ ok: false, error: 'code' });
  var role = roleFor(body.pin);
  if (role.error) return json({ ok: false, error: role.error });
  if (body.action === 'auth') return json({ ok: true, role: role.role });
  if (body.action === 'setPins') return json(setPinsFromWeb(body, role.role));
  if (body.action === 'importRead') return json(importRead(body, role.role));
  if (body.action === 'readDoc') return json(readDoc(body, role.role));
  if (body.action === 'ask') return json(ask(body, role.role));
  if (body.action === 'request') return json(addRequest(body, role.role));
  if (body.action === 'requests') return json(listRequests(body, role.role));
  if (body.action === 'requestDone') return json(requestDone(body, role.role));
  if (body.action !== 'save' || !Array.isArray(body.ops)) return json({ ok: false, error: 'Bad request.' });
  if (body.ops.length > 3000) return json({ ok: false, error: 'Too many changes at once. Save in smaller batches.' });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json({ ok: false, error: 'The tracker is busy saving someone else\'s changes. Try again in a moment.' });
  try {
    var touched = {}, tables = {};
    body.ops.forEach(function (op) { if (op && SCHEMA[op.table] && op.table !== 'log') touched[op.table] = true; });
    Object.keys(touched).forEach(function (t) { tables[t] = readTable(t).rows; });
    var who = String(body.name || 'Someone').slice(0, 40);
    var res = applyOps(tables, body.ops, role.role, who, new Date().toISOString(), !!body.atomic);
    if (res.applied) {
      Object.keys(touched).forEach(function (t) { writeTable(t, tables[t]); });
      appendLog(res.log);
    }
    return ContentService.createTextOutput('{"ok":true,"applied":' + res.applied + ',"conflicts":' + JSON.stringify(res.conflicts) + ',"data":' + dataJson() + '}').setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
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
  var names = (Array.isArray(body.names) ? body.names : []).slice(0, 400).map(function (n) { return String(n).slice(0, 80); });
  var rows = (Array.isArray(body.rows) ? body.rows : []).slice(0, 400).map(function (r) {
    return { n: Number(r.n) || 0, name: String(r.name || '').slice(0, 80), price: String(r.price || '').slice(0, 12), printedBefore: r.before === '' || r.before === null || r.before === undefined ? null : Number(r.before) };
  });
  var prompt = [
    'You are helping a school PTO spirit store tracker. First decide what these pages are:',
    '"order" = an online store cart, order confirmation, invoice or receipt for things bought to sell (usually Amazon).',
    '"countSheet" = the printed "Inventory Count Sheet" with numbered rows that volunteers filled in by hand.',
    '"deposit" = a bank deposit slip, deposit receipt, or bank record showing money put into the PTO account.',
    '"other" = anything else.',
    'Then return ONLY one JSON object, no other text, matching the type:',
    'For "order": {"type":"order","store":"Amazon","orderDate":"YYYY-MM-DD or empty","orderTotal":number or null,"lines":[{"title":"product title as shown","name":"short clean item name with pack size, e.g. Pink ribbon bracelets (24-pack)","match":"exact existing item name if the same product, else null","packs":integer,"packCost":price of ONE pack,"unitsPerPack":pieces in one pack or null if unclear,"note":""}]}',
    'Existing store items (exact names, for match): ' + JSON.stringify(names),
    'Order rules: one line per distinct product; skip Saved for later, recommendations, ads, shipping and tax; never invent products; if only a line total is shown divide by quantity.',
    'For "countSheet": {"type":"countSheet","countedBeforeBy":"","countedAfterBy":"","cash":{"start":number or null,"end":number or null,"cashApp":number or null},"notes":"","rows":[{"n":row number or null,"name":"item as written","before":number or null,"after":number or null,"sold":number or null,"note":"","unclear":true or false}]}',
    'Rows printed on the count sheet (n = row number): ' + JSON.stringify(rows),
    'Count sheet rules: include a row only if something is handwritten on it; for before give the final correct number (a handwritten correction beats a crossed-out printed number); null for blanks; never guess, mark unclear instead; include extra handwritten lines with n = null.',
    'For "deposit": {"type":"deposit","date":"YYYY-MM-DD or empty","amount":number,"what":"short description, e.g. Cash deposit or Checks","note":""}',
    'For "other": {"type":"other","what":"one short sentence saying what it looks like"}'
  ].join('\n');
  return callClaude(files, prompt);
}

/* ------------------------------ Ask the tracker (admins only) ------------------------------
   A plain-language answer about the live numbers or how to do something. The page sends a short summary of
   its numbers with the question. Every question, and every "Send to Hunter" request, is kept on the Requests tab. */
var HELP_GUIDE = [
  'How the tracker works (use these exact button and tab names):',
  '- One link for everyone. Looking needs only the access code. Changing anything needs a name and a PIN (button at the top: "Unlock to edit"). Volunteer PIN: counts, deposits, cash box. Admin PIN: everything, including orders, Cash App, sale details, PINs & access.',
  '- Tabs: Overview, Count, Money, Cash Box, Orders, History, Activity.',
  '- Overview: "Real profit to date" = money in minus spent plus stock on the shelf (valued at cost). The "This sale" card shows Items sold, Sold by count (after-sale counts times prices), Money taken in (cash taken in plus Cash App, once the cash box is counted), Profit (money taken in minus what the sold items cost). The red number on the Overview tab is the to-do list.',
  '- Sold by count and Money taken in rarely match exactly. Small gaps are normal: Cash App fees (if Cash App was entered after fees), pom pom pairs sold 2 for a deal, miscounts. A big short usually means an item was miscounted or money is missing; a big over usually means an item was missed on the Count tab.',
  '- Count tab: each item has BEFORE (leftovers on hand plus this sale\'s orders) and AFTER (what is left after the sale). Sold = before minus after. Items with nothing on hand do not need an after count.',
  '- Cash Box tab: Starting cash (change put in the box, usually $150), Ending cash (EVERYTHING in the box at closing, including the starting cash), Cash App total for this sale. Cash taken in = ending minus starting. Over/(short) compares cash taken in plus Cash App with Sold by count. If ending cash looks like it leaves out the starting cash, a yellow "Check the ending cash" box offers a one-tap fix.',
  '- Money tab: Money in (bank deposits, Cash App totals) and Money out (orders and other costs). Cash deposits are net of the starting cash.',
  '- Orders tab (admin): what was bought for this sale. Earlier orders are kept below.',
  '- Upload button (top): add a photo of the count sheet (each page), a deposit slip photo, an Amazon order screenshot or PDF, or the Cash App export file (admin). It reads it, asks one-tap questions about anything that does not add up, and nothing saves until every question is answered and Save is tapped. The green Saved screen has "Undo this upload".',
  '- Print count sheet button (top): makes a PDF. After the after-sale counts are in, it asks which numbers go in BEFORE: "What\'s left now" (for the next sale) or the current sale\'s starting numbers. "Leave off items with none left" hides zero rows; row numbers never change.',
  '- Typing in a field: a bar appears at the bottom with "Submit changes" (saves for everyone) and "Discard".',
  '- History tab: past sales with spent, sales and profit. To finish a sale: History tab, "Done with <sale name>?" card, "Close out". It locks in the sale (money taken in, or sold by count if the cash box is empty), carries the after counts forward as leftovers, clears the cash box and starts the next sale.',
  '- Activity tab: every saved change with name and time. Admins can Undo a change there.',
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
function ask(body, role, test) {
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
    'If the question asks for a change to the tracker itself (a new feature, something broken, something worded confusingly), say Hunter can add it and put a one sentence summary of the change in "request". Otherwise "request" is empty.',
    'Return ONLY one JSON object, no other text: {"answer":"...","goTo":"one of overview, count, money, cash, orders, history, activity, or empty if no tab helps","request":""}'
  ].join('\n\n');
  var r = callClaude([], prompt, { maxTokens: 1200, rawOk: true });
  if (!r.ok) return r;
  var o = r.result || {};
  var answer = String(o.answer || '').trim().slice(0, 2500);
  if (!answer) return { ok: false, error: "Couldn't get an answer. Try asking another way." };
  var goTo = ASK_TABS.indexOf(String(o.goTo || '').toLowerCase()) >= 0 ? String(o.goTo).toLowerCase() : '';
  var request = String(o.request || '').trim().slice(0, 400);
  if (!test) saveRequest('Question', body.name, q, answer);
  return { ok: true, answer: answer, goTo: goTo, request: request };
}
var REQ_TAB = 'Requests', REQ_COLS = ['Time', 'Who', 'Kind', 'Text', 'Reply', 'Status'];
function reqSheet(make) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(REQ_TAB);
  if (!sh && make) {
    sh = ss.insertSheet(REQ_TAB);
    sh.getRange(1, 1, 1, REQ_COLS.length).setValues([REQ_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function reqText(v, max) { var t = String(v || '').trim().slice(0, max); return /^[=+\-@']/.test(t) ? "'" + t : t; }
function saveRequest(kind, who, text, reply) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return false;
  try {
    var sh = reqSheet(true);
    var rng = sh.getRange(sh.getLastRow() + 1, 1, 1, REQ_COLS.length);
    rng.setNumberFormat('@');
    rng.setValues([[new Date().toISOString(), reqText(who || 'Someone', 40), kind, reqText(text, 1000), reqText(reply, 2500), kind === 'Request' ? 'Open' : '']]);
    return true;
  } finally { lock.releaseLock(); }
}
function addRequest(body, role) {
  if (role !== 'admin') return { ok: false, error: 'Requests are for admins. Unlock with the admin PIN.' };
  var t = String(body.text || '').trim();
  if (!t) return { ok: false, error: 'Type what you would like changed first.' };
  if (!saveRequest('Request', body.name, t, String(body.context || ''))) return { ok: false, error: 'The tracker is busy. Try again in a moment.' };
  return { ok: true };
}
function listRequests(body, role) {
  if (role !== 'admin') return { ok: false, error: 'Unlock with the admin PIN.' };
  var sh = reqSheet(false);
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };
  var last = sh.getLastRow(), first = Math.max(2, last - 39);
  var vals = sh.getRange(first, 1, last - first + 1, REQ_COLS.length).getValues(), out = [];
  for (var i = vals.length - 1; i >= 0; i--) {
    var v = vals[i];
    out.push({ row: first + i, time: String(v[0]), who: String(v[1]), kind: String(v[2]), text: String(v[3]), reply: String(v[4]), status: String(v[5]) });
  }
  return { ok: true, rows: out };
}
function requestDone(body, role) {
  if (role !== 'admin') return { ok: false, error: 'Unlock with the admin PIN.' };
  var sh = reqSheet(false), row = Number(body.row);
  if (!sh || !(row >= 2) || row > sh.getLastRow()) return { ok: false, error: 'That request is not there anymore.' };
  var v = sh.getRange(row, 1, 1, REQ_COLS.length).getValues()[0];
  if (String(v[0]) !== String(body.time)) return { ok: false, error: 'That list changed. Close it and open it again.' };
  sh.getRange(row, 6).setValue(body.done === false ? 'Open' : 'Done');
  return { ok: true };
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
  var last = '';
  for (var i = 0; i < AI_MODELS.length; i++) {
    var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: AI_MODELS[i], max_tokens: opts.maxTokens || 8000, messages: [{ role: 'user', content: content }] })
    });
    var code = resp.getResponseCode(), raw = resp.getContentText();
    if (code === 404 || (code === 400 && /model/i.test(raw) && !/credit/i.test(raw))) { last = raw; continue; }
    if (code === 401 || code === 403) return { ok: false, error: 'The AI key was rejected. Set a new one: Spirit Store > Set AI key.' };
    if (code === 429 || code === 529 || code >= 500) return { ok: false, error: 'The AI service is busy right now. Try again in a minute.' };
    if (code === 400 && /credit/i.test(raw)) return { ok: false, error: 'The AI account is out of credit. Add credit in the Claude Console.' };
    if (code !== 200) return { ok: false, error: 'The AI reader returned an error (' + code + ').' };
    var out = JSON.parse(raw), txt = (out.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
    var a = txt.indexOf('{'), b = txt.lastIndexOf('}');
    if (a < 0 || b < a) return opts.rawOk && txt.trim() ? { ok: true, result: { answer: txt.trim() }, model: AI_MODELS[i] } : { ok: false, error: opts.rawOk ? "Couldn't get an answer. Try asking another way." : "Couldn't find anything to read. Try a clearer picture." };
    var result;
    try { result = JSON.parse(txt.slice(a, b + 1)); } catch (e) { return opts.rawOk ? { ok: true, result: { answer: txt.trim() }, model: AI_MODELS[i] } : { ok: false, error: "Couldn't read that. Try again or use a clearer picture." }; }
    return { ok: true, result: result, model: AI_MODELS[i] };
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
  return { rows: rows, sheet: sh, idx: idx, fixed: fixed };
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
  var sh = SpreadsheetApp.getActive().getSheetByName(SCHEMA[t].tab);
  var header = normHeader(sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]);
  var height = Math.max(sh.getLastRow() - 1, rows.length, 1);
  SCHEMA[t].cols.forEach(function (c) {
    var col = header.indexOf(c[1].toLowerCase()) + 1;
    if (col < 1) throw new Error('The "' + SCHEMA[t].tab + '" tab is missing its "' + c[1] + '" column.');
    var out = [];
    for (var r = 0; r < height; r++) { var v = r < rows.length ? rows[r][c[0]] : ''; v = v === undefined || v === null ? '' : v; out.push([sheetSafe(t, c[0], v, r < rows.length ? rows[r].id : '')]); }
    sh.getRange(2, col, height, 1).setValues(out);
  });
}

function appendLog(entries) {
  if (!entries || !entries.length) return;
  dropCache();
  var sh = SpreadsheetApp.getActive().getSheetByName(SCHEMA.log.tab);
  var rows = entries.map(function (e) { return SCHEMA.log.cols.map(function (c) { return e[c[0]] === undefined ? '' : String(e[c[0]]); }); });
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
