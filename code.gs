/**
 * VD Horala – Charge list : Google Apps Script
 * Shows every tab of the spreadsheet, lets you edit any cell and add rows.
 *
 * Setup: open the Google Sheet → Extensions → Apps Script → paste this file
 * → Deploy → New deployment → Web app → Execute as: Me → Access: Anyone.
 */

// Tabs that are NEVER sent to the web page (they hold passwords / other web-app links).
// Hidden tabs are always skipped too.
var EXCLUDE_TABS = ['Credential', 'webapp', 'vaccinator code'];

// Optional passcode. If set, the page asks for it. '' = no passcode. (Recommended.)
var ACCESS_CODE = '';

// Default layout of a tab: row 1 = title, rows 2-3 = column headings, data below.
var DEFAULT_LAYOUT = { titleRow: 1, headerRows: [2, 3], firstDataRow: 4 };
// Tabs with a different layout (only needed if you remove them from EXCLUDE_TABS):
var LAYOUTS = {
  'Credential':      { titleRow: 0, headerRows: [1], firstDataRow: 2 },
  'webapp':          { titleRow: 0, headerRows: [],  firstDataRow: 1 },
  'vaccinator code': { titleRow: 0, headerRows: [],  firstDataRow: 1 }
};

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (ACCESS_CODE && String(body.code || '') !== ACCESS_CODE) return json_({ error: 'AUTH' });

    if (body.action === 'tabs') return json_({ tabs: listTabs_() });
    if (body.action === 'listAll') {      // every allowed tab with its rows, in one call
      var all = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(allowed_).map(function (sh) {
        var d = readTab_(sh);
        d.name = sh.getName();
        d.label = d.title || sh.getName().trim();
        return d;
      });
      return json_({ tabs: all });
    }

    var sheet = getTab_(body.tab);
    if (!sheet) return json_({ error: 'Tab not found or not allowed: ' + body.tab });
    if (body.action === 'list')   return json_(readTab_(sheet));
    if (body.action === 'update') return json_(updateRow_(sheet, body));
    if (body.action === 'add')    return json_(addRow_(sheet, body));
    return json_({ error: 'Unknown action: ' + body.action });
  } catch (err) {
    return json_({ error: String(err) });
  }
}

function doGet() { return json_({ status: 'ok', message: 'VD Horala API running. Use POST.' }); }

/* ---------- helpers ---------- */

function s_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  if (typeof v === 'number') return String(v);
  return String(v).trim();
}

function layoutFor_(name) { return LAYOUTS[name] || DEFAULT_LAYOUT; }

function allowed_(sheet) {
  return !sheet.isSheetHidden() && EXCLUDE_TABS.indexOf(sheet.getName()) < 0;
}

function getTab_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(name || ''));
  return sh && allowed_(sh) ? sh : null;
}

function titleOf_(sheet) {
  var L = layoutFor_(sheet.getName());
  if (!L.titleRow || sheet.getLastColumn() < 1) return { title: '', sub: '' };
  var vals = sheet.getRange(L.titleRow, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(s_).filter(function (x) { return x; });
  return { title: vals[0] || '', sub: vals.slice(1).join(' · ') };
}

function listTabs_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(allowed_)
    .map(function (sh) {
      var t = titleOf_(sh);
      return { name: sh.getName(), label: t.title || sh.getName().trim() };
    });
}

/** Column labels: last header row with text wins (so row 3 refines row 2). */
function headersOf_(sheet, L) {
  var maxCol = Math.max(sheet.getLastColumn(), 1);
  var hv = L.headerRows.map(function (r) { return sheet.getRange(r, 1, 1, maxCol).getValues()[0]; });
  var labels = [], last = 0;
  for (var c = 0; c < maxCol; c++) {
    var label = '';
    for (var k = hv.length - 1; k >= 0 && !label; k--) label = s_(hv[k][c]);
    labels.push(label);
    if (label) last = c + 1;
  }
  return { labels: labels.slice(0, last), count: last };
}

function colLetter_(i) { return String.fromCharCode(65 + i); }

function readTab_(sheet) {
  var L = layoutFor_(sheet.getName());
  var t = titleOf_(sheet);
  var h = headersOf_(sheet, L);
  var lastRow = sheet.getLastRow();
  var nCols = h.count;
  if (!nCols) nCols = Math.min(sheet.getLastColumn(), 26) || 1;
  var labels = h.labels.length ? h.labels : [];
  for (var i = 0; i < nCols; i++) if (!labels[i]) labels[i] = labels[i] || colLetter_(i);

  var rows = [], serialCol = -1, formulaCols = [];
  if (lastRow >= L.firstDataRow) {
    var n = lastRow - L.firstDataRow + 1;
    var range = sheet.getRange(L.firstDataRow, 1, n, nCols);
    var vals = range.getValues(), fm = range.getFormulas();
    var lastSerialIdx = -1;
    for (var r = 0; r < n; r++) {
      var v = vals[r].map(s_), any = false;
      for (var c = 0; c < nCols; c++) if (v[c] !== '') { any = true; break; }
      if (!any) continue;
      var f = [];
      for (c = 0; c < nCols; c++) if (fm[r][c]) f.push(c);
      var row = { r: L.firstDataRow + r, v: v };
      if (f.length) row.f = f;
      rows.push(row);
      if (/^\d+$/.test(v[0])) lastSerialIdx = r;
    }
    if (lastSerialIdx >= 0) {
      serialCol = 0;
      for (c = 0; c < nCols; c++) if (fm[lastSerialIdx][c]) formulaCols.push(c);
    }
  }
  return { title: t.title, sub: t.sub, headers: labels, rows: rows, serialCol: serialCol, formulaCols: formulaCols };
}

function writeCell_(cell, v) {
  v = String(v === null || v === undefined ? '' : v).trim();
  if (v === '') { cell.clearContent(); return; }
  var cur = cell.getValue();
  if (cur instanceof Date) cell.setValue(v);                       // let Sheets parse dates
  else if (/^-?\d+(\.\d+)?$/.test(v) && !/^0\d/.test(v)) cell.setValue(Number(v));
  else { cell.setNumberFormat('@'); cell.setValue(v); }            // keep text exactly as typed
}

/* ---------- edit ---------- */

function updateRow_(sheet, body) {
  var L = layoutFor_(sheet.getName());
  var row = Number(body.row);
  var changes = body.changes || {}, old = body.old || {};
  if (!(row >= L.firstDataRow && row <= sheet.getLastRow())) return { error: 'Row out of range' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var cols = Object.keys(changes).map(Number);
    for (var i = 0; i < cols.length; i++) {
      var cell = sheet.getRange(row, cols[i] + 1);
      if (cell.getFormula()) return { error: 'Column ' + colLetter_(cols[i]) + ' is a formula and cannot be edited' };
      if (s_(cell.getValue()) !== String(old[cols[i]] || '')) {
        return { error: 'This row was changed by someone else. Please refresh and try again.' };
      }
    }
    for (i = 0; i < cols.length; i++) writeCell_(sheet.getRange(row, cols[i] + 1), changes[cols[i]]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- add ---------- */

/** New row goes right after the last numbered row (before any "taken over / handed over" notes). */
function addRow_(sheet, body) {
  var L = layoutFor_(sheet.getName());
  var h = headersOf_(sheet, L);
  var nCols = h.count || Math.min(sheet.getLastColumn(), 26) || 1;
  var values = body.values || [];
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var lastRow = sheet.getLastRow();
    var prev = L.firstDataRow - 1, hasSerial = false, lastAny = L.firstDataRow - 1;
    if (lastRow >= L.firstDataRow) {
      var vals = sheet.getRange(L.firstDataRow, 1, lastRow - L.firstDataRow + 1, nCols).getValues();
      for (var r = 0; r < vals.length; r++) {
        if (/^\d+$/.test(s_(vals[r][0]))) { prev = L.firstDataRow + r; hasSerial = true; }
        for (var c = 0; c < nCols; c++) if (s_(vals[r][c]) !== '') { lastAny = L.firstDataRow + r; break; }
      }
    }
    if (!hasSerial) prev = lastAny;
    var any = false;
    for (c = 0; c < nCols; c++) if (c !== 0 || !hasSerial) if (String(values[c] || '').trim() !== '') any = true;
    if (!any) return { error: 'Enter at least one value' };

    var prevSerial = hasSerial ? Number(s_(sheet.getRange(prev, 1).getValue())) : 0;
    var prevFormulas = prev >= L.firstDataRow ? sheet.getRange(prev, 1, 1, nCols).getFormulasR1C1()[0] : [];
    sheet.insertRowAfter(Math.max(prev, 1));
    var nr = Math.max(prev, 1) + 1;
    for (c = 0; c < nCols; c++) {
      var cell = sheet.getRange(nr, c + 1);
      if (prevFormulas[c]) { cell.setFormulaR1C1(prevFormulas[c]); continue; }   // keep auto columns (e.g. Sl No)
      if (c === 0 && hasSerial) { cell.setValue(prevSerial + 1); continue; }
      writeCell_(cell, values[c]);
    }
    return { ok: true, row: nr };
  } finally { lock.releaseLock(); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
