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


var SEED = {"items":[{"id":"it01","name":"Butter squishies (8-pack)","category":"New","price":5,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it02","name":"Cheese squishies, slow-rising (12-pack)","category":"New","price":4,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it03","name":"Green tinsel pom pom headbands (24-pack)","category":"New","price":3,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it04","name":"Pink glitter hair tinsel (12-pack)","category":"New","price":3,"unitCost":"","onHand":2,"countedAfter":"","notes":""},{"id":"it05","name":"Pink ribbon bracelets (24-pack)","category":"New","price":3,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it06","name":"Halloween pumpkin squishies (12-pack)","category":"New","price":3,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it07","name":"Halloween pumpkin glitter squishies (8-pack)","category":"New","price":3,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it08","name":"Mini glitter ghost squishies (12-pack)","category":"New","price":3,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it09","name":"Glitter heart squishies (24-pack)","category":"New","price":3,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it10","name":"Green cheer pom poms","category":"Restock","price":2,"unitCost":"","onHand":0,"countedAfter":"","notes":""},{"id":"it11","name":"Slow-rising stress cubes (24-pack)","category":"New","price":2,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it12","name":"Halloween 4-in-1 pens (50-pack)","category":"New","price":2,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it13","name":"Mini turtle squishies (24-pack)","category":"New","price":1,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it14","name":"Pink ribbon & football tattoos (20 sheets)","category":"New","price":1,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it15","name":"Stretchy skeletons (48-pack)","category":"New","price":1,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it16","name":"Pink paw print glitter eye stickers (120 pairs)","category":"New","price":1,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it17","name":"Pink glitter eye black stickers (120 pairs)","category":"New","price":1,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it18","name":"Green glitter eye black stickers (120 pairs)","category":"New","price":1,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it19","name":"Green paw print bookmarks (75-pack)","category":"New","price":0.5,"unitCost":"","onHand":"","countedAfter":"","notes":""},{"id":"it20","name":"Gold football charm bracelets","category":"Carry-over","price":3,"unitCost":0.7496,"onHand":16,"countedAfter":"","notes":""},{"id":"it21","name":"Green paw print bracelets","category":"Carry-over","price":3,"unitCost":0.7496,"onHand":39,"countedAfter":"","notes":""},{"id":"it22","name":"Black FOOTBALL clay bracelets","category":"Carry-over","price":3,"unitCost":0.6663,"onHand":15,"countedAfter":"","notes":""},{"id":"it23","name":"Football candy bracelets","category":"Carry-over","price":3,"unitCost":0.4033,"onHand":22,"countedAfter":"","notes":""},{"id":"it24","name":"Green cheer bracelets","category":"Carry-over","price":3,"unitCost":0.5995,"onHand":16,"countedAfter":"","notes":""},{"id":"it25","name":"Green tinsel hair extensions","category":"Carry-over","price":3,"unitCost":0.555,"onHand":21,"countedAfter":"","notes":""},{"id":"it26","name":"Glitter bear squishies","category":"Carry-over","price":3,"unitCost":1.1246,"onHand":0,"countedAfter":"","notes":""},{"id":"it27","name":"Glitter cube squishies","category":"Carry-over","price":3,"unitCost":1.2496,"onHand":0,"countedAfter":"","notes":""},{"id":"it28","name":"Slow-rise squishy stress balls","category":"Carry-over","price":3,"unitCost":1.2494,"onHand":0,"countedAfter":"","notes":""},{"id":"it29","name":"Football #1 keychains","category":"Carry-over","price":3,"unitCost":0.9156,"onHand":5,"countedAfter":"","notes":""},{"id":"it30","name":"Green pom pom keychains","category":"Carry-over","price":2,"unitCost":0.9995,"onHand":0,"countedAfter":"","notes":"Not on leftover count, so set to 0"},{"id":"it31","name":"Mini foam footballs","category":"Carry-over","price":2,"unitCost":0.5798,"onHand":32,"countedAfter":"","notes":""},{"id":"it32","name":"Lion 4-color pens","category":"Carry-over","price":2,"unitCost":0.8325,"onHand":0,"countedAfter":"","notes":"No leftovers counted; 12 new"},{"id":"it33","name":"Football fidget spinners","category":"Carry-over","price":1,"unitCost":0.3956,"onHand":6,"countedAfter":"","notes":""},{"id":"it34","name":"Eye black stickers (pair)","category":"Carry-over","price":1,"unitCost":0.0418,"onHand":135,"countedAfter":"","notes":""},{"id":"it35","name":"Cross eye black stickers (pair)","category":"Carry-over","price":1,"unitCost":0.1873,"onHand":24,"countedAfter":"","notes":""},{"id":"it36","name":"Balloon animal erasers","category":"Carry-over","price":1,"unitCost":0.4612,"onHand":0,"countedAfter":"","notes":"No leftovers counted; 24 new"},{"id":"it37","name":"Mini mochi squishies","category":"Carry-over","price":1,"unitCost":0.1999,"onHand":18,"countedAfter":"","notes":""},{"id":"it38","name":"Football pens, rainbow bendy (30-pack)","category":"Carry-over","price":1,"unitCost":0.4663,"onHand":3,"countedAfter":"","notes":""},{"id":"it39","name":"Football pens, green bendy (50-pack)","category":"Carry-over","price":1,"unitCost":0.3998,"onHand":50,"countedAfter":"","notes":"Arrived late last week; none sold yet"},{"id":"it40","name":"Green rubber \"I am a Creekside Star\" bracelets","category":"Carry-over","price":0.5,"unitCost":0,"onHand":33,"countedAfter":"","notes":"Older stock: add cost if known"},{"id":"it41","name":"Air freshener","category":"Carry-over","price":1,"unitCost":0,"onHand":1,"countedAfter":"","notes":"Older stock: add cost if known"},{"id":"it42","name":"Stickers (regular)","category":"Carry-over","price":0.5,"unitCost":0,"onHand":"","countedAfter":"","notes":"Count before Friday; add cost if known"},{"id":"it43","name":"Fuzzy monster pens","category":"Carry-over","price":2,"unitCost":0.7495,"onHand":20,"countedAfter":"","notes":"From Sept 26 order (was saved for Christmas)"},{"id":"it44","name":"Bear stackable pencils","category":"Carry-over","price":1,"unitCost":0.463,"onHand":30,"countedAfter":"","notes":"From Sept 26 order (was saved for Christmas)"},{"id":"it45","name":"Fidget snake cubes","category":"Carry-over","price":1,"unitCost":0.5297,"onHand":30,"countedAfter":"","notes":"From Sept 26 order (was saved for Christmas)"},{"id":"it46","name":"Football lollipop rings","category":"Donated","price":1,"unitCost":0,"onHand":54,"countedAfter":"","notes":"3 donated boxes of 18"},{"id":"it47","name":"Green spirit tattoos","category":"Donated","price":1,"unitCost":0,"onHand":0,"countedAfter":"","notes":"Not on leftover count, so set to 0"},{"id":"it48","name":"Green bead necklaces","category":"Donated","price":1,"unitCost":0,"onHand":121,"countedAfter":"","notes":""},{"id":"it49","name":"Green heart sunglasses","category":"Donated","price":2,"unitCost":0,"onHand":17,"countedAfter":"","notes":""},{"id":"it50","name":"Lion paint kits","category":"Donated","price":4,"unitCost":0,"onHand":8,"countedAfter":"","notes":""}],"purchases":[{"id":"pu01","date":"2026-10-12","store":"Amazon","item":"Butter squishies (8-pack)","packs":1,"packCost":24.99,"unitsPerPack":8,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu02","date":"2026-10-12","store":"Amazon","item":"Cheese squishies, slow-rising (12-pack)","packs":4,"packCost":25.64,"unitsPerPack":12,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu03","date":"2026-10-12","store":"Amazon","item":"Green tinsel pom pom headbands (24-pack)","packs":1,"packCost":25.99,"unitsPerPack":24,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu04","date":"2026-10-12","store":"Amazon","item":"Pink glitter hair tinsel (12-pack)","packs":1,"packCost":7.99,"unitsPerPack":12,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu05","date":"2026-10-12","store":"Amazon","item":"Pink ribbon bracelets (24-pack)","packs":1,"packCost":11.99,"unitsPerPack":24,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu06","date":"2026-10-12","store":"Amazon","item":"Halloween pumpkin squishies (12-pack)","packs":3,"packCost":18.99,"unitsPerPack":12,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu07","date":"2026-10-12","store":"Amazon","item":"Halloween pumpkin glitter squishies (8-pack)","packs":2,"packCost":15.99,"unitsPerPack":8,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu08","date":"2026-10-12","store":"Amazon","item":"Mini glitter ghost squishies (12-pack)","packs":1,"packCost":26.72,"unitsPerPack":12,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu09","date":"2026-10-12","store":"Amazon","item":"Glitter heart squishies (24-pack)","packs":2,"packCost":19.99,"unitsPerPack":24,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu10","date":"2026-10-12","store":"Amazon","item":"Green cheer pom poms","packs":1,"packCost":22.99,"unitsPerPack":36,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu11","date":"2026-10-12","store":"Amazon","item":"Slow-rising stress cubes (24-pack)","packs":3,"packCost":19.99,"unitsPerPack":24,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu12","date":"2026-10-12","store":"Amazon","item":"Halloween 4-in-1 pens (50-pack)","packs":1,"packCost":27.99,"unitsPerPack":50,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu13","date":"2026-10-12","store":"Amazon","item":"Mini turtle squishies (24-pack)","packs":3,"packCost":9.99,"unitsPerPack":24,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu14","date":"2026-10-12","store":"Amazon","item":"Pink ribbon & football tattoos (20 sheets)","packs":1,"packCost":9.99,"unitsPerPack":160,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu15","date":"2026-10-12","store":"Amazon","item":"Stretchy skeletons (48-pack)","packs":1,"packCost":17.09,"unitsPerPack":48,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu16","date":"2026-10-12","store":"Amazon","item":"Pink paw print glitter eye stickers (120 pairs)","packs":1,"packCost":13.99,"unitsPerPack":120,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu17","date":"2026-10-12","store":"Amazon","item":"Pink glitter eye black stickers (120 pairs)","packs":1,"packCost":8.79,"unitsPerPack":120,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu18","date":"2026-10-12","store":"Amazon","item":"Green glitter eye black stickers (120 pairs)","packs":1,"packCost":7.79,"unitsPerPack":120,"source":"Actual","sale":"Pink Out Game","notes":""},{"id":"pu19","date":"2026-10-12","store":"Amazon","item":"Green paw print bookmarks (75-pack)","packs":1,"packCost":28.45,"unitsPerPack":75,"source":"Actual","sale":"Pink Out Game","notes":""}],"moneyIn":[{"id":"mi01","date":"2026-09-25","event":"Spirit store","amount":647.7,"notes":"Bank deposit slip. Tally sheet math doesn't match ($641.70 net); deposit used"},{"id":"mi02","date":"2026-10-02","event":"Spirit store","amount":1108.4,"notes":"Matches tally sheet"},{"id":"mi03","date":"2026-10-02","event":"Class sponsor deposit (face painting)","amount":119,"notes":"Separate deposit slip. Not spirit store sales"},{"id":"mi04","date":"2026-09-25","event":"Cash App - spirit store","amount":"","notes":"Enter Cash App total for this day"},{"id":"mi05","date":"2026-10-02","event":"Cash App - spirit store","amount":"","notes":"Enter Cash App total for this day"},{"id":"mi06","date":"2026-10-02","event":"Cash App - class sponsors","amount":"","notes":"Enter Cash App total for this day"},{"id":"mi07","date":"2026-10-16","event":"Pink Out spirit store (cash deposit)","amount":"","notes":"Enter Friday's deposit"},{"id":"mi08","date":"2026-10-16","event":"Pink Out spirit store (Cash App)","amount":"","notes":"Enter Friday's Cash App total"}],"moneyOut":[{"id":"mo01","date":"","order":"Stock for Sept 25 sale (earlier orders)","amount":"","link":"","notes":"Enter if known"},{"id":"mo02","date":"2026-09-26","order":"Amazon order (Sept 26)","amount":519.55,"link":"","notes":"Order 114-5952805-0437038. Spirit store $436.31 + $44.77 saved for Christmas + $38.47 supplies"},{"id":"mo03","date":"2026-10-12","order":"Amazon order (Pink Out)","amount":"","link":"purchases","notes":"Fills in automatically from the Orders tab"},{"id":"mo04","date":"","order":"Fall festival order","amount":"","link":"","notes":"Enter when ordered (about $1,006.97)"}],"sales":[{"id":"sh01","date":"2026-09-25","theme":"Spirit store","spent":"","sales":647.7,"other":"","what":"","link":""},{"id":"sh02","date":"2026-10-02","theme":"Football / spirit","spent":436.31,"sales":1108.4,"other":119,"what":"Class sponsor deposit (face painting), separate slip","link":""},{"id":"sh03","date":"2026-10-16","theme":"Pink Out","spent":"","sales":"","other":"","what":"","link":"current"}],"cash":[{"id":"start","label":"Starting cash in box (change)","value":0},{"id":"end","label":"Ending cash in box","value":""},{"id":"countedBy","label":"Counted by","value":""},{"id":"notes","label":"Notes","value":""}],"settings":[{"id":"orgName","value":"Creekside Lions PTO"},{"id":"schoolLine","value":"Creekside Elementary \u00b7 Livingston ISD"},{"id":"eventName","value":"Pink Out Game"},{"id":"eventDate","value":"2026-10-16"},{"id":"startCash","value":150}],"log":[]};

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
    var rows = (SEED[t] || []).map(function (r) { return SCHEMA[t].cols.map(function (c) { var v = r[c[0]]; return v === undefined || v === null ? '' : v; }); });
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
  return json({ ok: true, data: loadAll() });
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
  if (body.action !== 'save' || !Array.isArray(body.ops)) return json({ ok: false, error: 'Bad request.' });
  if (body.ops.length > 3000) return json({ ok: false, error: 'Too many changes at once. Save in smaller batches.' });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json({ ok: false, error: 'The tracker is busy saving someone else\'s changes. Try Submit again.' });
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
    return json({ ok: true, applied: res.applied, conflicts: res.conflicts, data: loadAll() });
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

function callClaude(images, prompt) {
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
      payload: JSON.stringify({ model: AI_MODELS[i], max_tokens: 8000, messages: [{ role: 'user', content: content }] })
    });
    var code = resp.getResponseCode(), raw = resp.getContentText();
    if (code === 404 || (code === 400 && /model/i.test(raw) && !/credit/i.test(raw))) { last = raw; continue; }
    if (code === 401 || code === 403) return { ok: false, error: 'The AI key was rejected. Set a new one: Spirit Store > Set AI key.' };
    if (code === 429 || code === 529 || code >= 500) return { ok: false, error: 'The AI service is busy right now. Try again in a minute.' };
    if (code === 400 && /credit/i.test(raw)) return { ok: false, error: 'The AI account is out of credit. Add credit in the Claude Console.' };
    if (code !== 200) return { ok: false, error: 'The AI reader returned an error (' + code + ').' };
    var out = JSON.parse(raw), txt = (out.content || []).filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('');
    var a = txt.indexOf('{'), b = txt.lastIndexOf('}');
    if (a < 0 || b < a) return { ok: false, error: "Couldn't find anything to read. Try a clearer picture." };
    var result;
    try { result = JSON.parse(txt.slice(a, b + 1)); } catch (e) { return { ok: false, error: "Couldn't read that. Try again or use a clearer picture." }; }
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

function readTable(t) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SCHEMA[t].tab);
  if (!sh) throw new Error('Missing tab "' + SCHEMA[t].tab + '". Run Spirit Store > Set up tabs.');
  var values = sh.getDataRange().getValues();
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
  var sh = SpreadsheetApp.getActive().getSheetByName(SCHEMA.log.tab);
  var rows = entries.map(function (e) { return SCHEMA.log.cols.map(function (c) { return e[c[0]] === undefined ? '' : String(e[c[0]]); }); });
  var rng = sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length);
  rng.setNumberFormat('@'); // keep log entries as plain text so dates and numbers are not reinterpreted
  rng.setValues(rows);
}

var CASH_ROWS = [['start', 'Starting cash in box (change)'], ['end', 'Ending cash in box'], ['cashapp', 'Cash App total for this sale'], ['countedBy', 'Counted by'], ['notes', 'Notes']];
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
  var lg = readTable('log').rows;
  out.log = lg.slice(-80).reverse();
  return out;
}
