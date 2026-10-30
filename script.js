/**
 * ============================================================
 *  CONFIG – put your Apps Script Web App URL here
 * ============================================================
 */
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzurPKVxPDrAS0ifptOj6bOIhMgBpi2XUamvotpxiUX9pl5ENoDDFrWQCclBzGdVaXLPw/exec';
/* ============================================================ */

let tabs = [];            // [{name,label}]
let active = null;        // tab name
const cache = {};         // name -> {headers, rows, serialCol, formulaCols, title, sub}
let mode = 'edit';        // 'edit' | 'add'
let editing = null;       // row index in ctxTab
let ctxTab = null;        // tab the open modal belongs to
let connected = false;

function $(id) { return document.getElementById(id); }
function cur() { return cache[ctxTab || active]; }

function toast(msg, isError) {
  const t = $('toast');
  t.textContent = msg;
  t.className = 'toast show' + (isError ? ' error' : '');
  setTimeout(() => t.className = 'toast', 2800);
}
function setStatus(text, type) {
  $('statusText').textContent = text;
  $('liveDot').className = 'dot' + (type === 'ok' ? ' live' : type === 'err' ? ' err' : '');
}
function setLoading(on) { $('tableLoading').classList.toggle('show', on); }
function setEnabled(yes) {
  connected = yes;
  ['btnRefresh', 'btnAdd', 'searchInput'].forEach(id => $(id).disabled = !yes);
}
function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

async function callScript(action, payload, retried) {
  if (!SCRIPT_URL || SCRIPT_URL.startsWith('PASTE_')) {
    throw new Error('Open script.js and set your Web App URL in SCRIPT_URL');
  }
  const body = Object.assign({ action: action, code: sessionStorage.getItem('vdCode') || '' }, payload || {});
  const res = await fetch(SCRIPT_URL, {
    method: 'POST', mode: 'cors',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error('Network error HTTP ' + res.status);
  const json = await res.json();
  if (json.error === 'AUTH') {
    if (retried) throw new Error('Wrong passcode');
    const code = prompt('Enter passcode');
    if (code === null) throw new Error('Passcode required');
    sessionStorage.setItem('vdCode', code);
    return callScript(action, payload, true);
  }
  if (json.error) throw new Error(json.error);
  return json;
}

/* ---------- tabs ---------- */
function renderTabs() {
  $('tabs').innerHTML = tabs.map(function (t, i) {
    return '<button class="tab' + (t.name === active ? ' active' : '') + '" title="' + esc(t.label) +
      '" onclick="selectTab(' + i + ')">' + esc(t.label) + '</button>';
  }).join('');
}

async function loadAll() {
  const res = await callScript('listAll');
  tabs = (res.tabs || []).map(t => ({ name: t.name, label: t.label }));
  (res.tabs || []).forEach(t => { cache[t.name] = t; });
  if (!tabs.length) throw new Error('No tabs available');
  if (!tabs.some(t => t.name === active)) active = tabs[0].name;
}

async function init() {
  setLoading(true);
  setStatus('Loading…');
  try {
    await loadAll();
    renderTabs();
    render();
    setEnabled(true);
  } catch (e) {
    setEnabled(false);
    setStatus('Error: ' + e.message, 'err');
    toast(e.message, true);
  } finally { setLoading(false); }
}

async function loadTab(name) {
  cache[name] = await callScript('list', { tab: name });
  cache[name].name = name;
  const t = tabs.find(x => x.name === name);
  cache[name].label = t ? t.label : name;
}

function selectTab(i) {
  active = tabs[i].name;
  $('searchInput').value = '';      // jumping to a tab leaves search mode
  renderTabs();
  render();
}

async function refresh() {            // reloads every tab (also picks up newly added tabs)
  setLoading(true);
  setStatus('Refreshing…');
  try {
    await loadAll();
    renderTabs();
    render();
    toast('Refreshed');
  } catch (e) {
    setStatus('Error: ' + e.message, 'err');
    toast(e.message, true);
  } finally { setLoading(false); }
}

/* ---------- table ---------- */
function isAuto(c) {            // columns filled automatically when adding a row
  const d = cur();
  return d.formulaCols.indexOf(c) >= 0 || (d.serialCol === c);
}

function terms() { return $('searchInput').value.trim().toLowerCase().split(/\s+/).filter(Boolean); }

/** Escape + highlight search terms */
function hl(text, ts) {
  text = String(text);
  if (!ts.length) return esc(text);
  const re = new RegExp('(' + ts.map(t => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|') + ')', 'gi');
  return text.split(re).map((p, i) => i % 2 ? '<mark>' + esc(p) + '</mark>' : esc(p)).join('');
}

function tableHtml(name, list, ts) {
  const d = cache[name];
  const head = '<tr>' + d.headers.map(h => '<th>' + esc(h) + '</th>').join('') + '<th style="width:70px">Edit</th></tr>';
  const body = list.map(function (item) {
    const r = item.row;
    const note = d.serialCol >= 0 && !/^\d+$/.test(r.v[0]);
    return '<tr' + (note ? ' class="note-row"' : '') + '>' +
      r.v.map(x => '<td>' + hl(x, ts) + '</td>').join('') +
      '<td class="edit-cell" style="white-space:nowrap;min-width:76px"><button class="btn btn-outline btn-sm" style="white-space:nowrap" onclick="openEditModal(' + JSON.stringify(name).replace(/"/g, '&quot;') + ',' + item.idx + ')">Edit</button></td></tr>';
  }).join('');
  return '<div class="table-scroll"><table><thead>' + head + '</thead><tbody>' + body + '</tbody></table></div>';
}

function render() {
  const ts = terms();
  const searching = ts.length > 0;
  $('btnAdd').disabled = !connected || searching;
  $('btnAdd').title = searching ? 'Clear the search, then pick a tab to add a row' : '';
  renderTabs();
  const out = [];
  let shown = 0, total = 0;

  if (!searching) {
    const d = cache[active];
    if (!d) return;
    $('tabTitle').innerHTML = esc(d.title || '') + (d.sub ? '<span class="tab-sub">' + esc(d.sub) + '</span>' : '');
    const list = d.rows.map((row, idx) => ({ row: row, idx: idx }));
    total = d.rows.length; shown = list.length;
    if (shown) out.push(tableHtml(active, list, ts));
  } else {
    $('tabTitle').textContent = 'Search results – all tabs';
    tabs.forEach(function (t) {
      const d = cache[t.name];
      if (!d) return;
      total += d.rows.length;
      const list = [];
      d.rows.forEach(function (row, idx) {
        const hay = row.v.join(' ').toLowerCase();
        if (ts.every(x => hay.includes(x))) list.push({ row: row, idx: idx });
      });
      if (!list.length) return;
      shown += list.length;
      out.push('<div class="section-title" onclick="goTab(' + JSON.stringify(t.name).replace(/"/g, '&quot;') + ')">' +
        esc(t.label) + ' <span class="count">' + list.length + '</span></div>' + tableHtml(t.name, list, ts));
    });
  }
  $('tableArea').innerHTML = out.join('');
  $('totalCount').textContent = total;
  $('showCount').textContent = shown;
  $('emptyState').style.display = shown === 0 ? 'block' : 'none';
  setStatus('Live · ' + tabs.length + ' tabs', 'ok');
}

function goTab(name) { selectTab(tabs.findIndex(t => t.name === name)); }

/* ---------- modal ---------- */
function buildFields(values, lockedFn) {
  const d = cur();
  return d.headers.map(function (h, c) {
    const locked = lockedFn(c);
    return '<div class="form-group"><label>' + esc(h) + (locked === 'auto' ? ' <span class="auto">(auto)</span>' : '') +
      '</label><input type="text" id="f' + c + '" value="' + esc(values[c] || '') + '"' +
      (locked ? ' disabled class="ro"' : '') + ' autocomplete="off"></div>';
  }).join('');
}

function openEditModal(name, idx) {
  if (!connected) return;
  mode = 'edit'; editing = idx; ctxTab = name;
  const r = cur().rows[idx], fset = r.f || [];
  $('modalTitle').textContent = 'Edit · ' + (cache[name].label || name.trim()) + ' · row ' + r.r;
  $('modalFields').innerHTML = buildFields(r.v, c => fset.indexOf(c) >= 0 ? 'auto' : '');
  $('saveBtn').textContent = 'Update Sheet';
  $('modal').classList.add('open');
  const first = cur().headers.findIndex((_, c) => fset.indexOf(c) < 0);
  if ($('f' + first)) $('f' + first).focus();
}

function openAddModal() {
  if (!connected) return;
  mode = 'add'; editing = null; ctxTab = active;
  const d = cur();
  $('modalTitle').textContent = 'Add to ' + (d.title || active.trim());
  $('modalFields').innerHTML = buildFields([], c => isAuto(c) ? 'auto' : '');
  $('saveBtn').textContent = 'Add to Sheet';
  $('modal').classList.add('open');
  const first = d.headers.findIndex((_, c) => !isAuto(c));
  if ($('f' + first)) $('f' + first).focus();
}

function closeModal() { $('modal').classList.remove('open'); }

async function saveItem() {
  if (!connected) return;
  const d = cur();
  const btn = $('saveBtn');
  const original = btn.textContent;
  const vals = d.headers.map((_, c) => $('f' + c).value.trim());
  let payload, action;

  if (mode === 'edit') {
    const r = d.rows[editing], fset = r.f || [];
    const changes = {}, old = {};
    vals.forEach(function (v, c) {
      if (fset.indexOf(c) < 0 && v !== r.v[c]) { changes[c] = v; old[c] = r.v[c]; }
    });
    if (!Object.keys(changes).length) { closeModal(); return; }
    action = 'update';
    payload = { tab: ctxTab, row: r.r, changes: changes, old: old };
  } else {
    if (!vals.some((v, c) => v && !isAuto(c))) { toast('Enter at least one value', true); return; }
    action = 'add';
    payload = { tab: ctxTab, values: vals };
  }

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Saving…';
  try {
    await callScript(action, payload);
    const wasAdd = mode === 'add', tabName = ctxTab;
    closeModal();
    await loadTab(tabName);            // row numbers shift after an insert, so reload that tab
    render();
    toast(wasAdd ? 'Row added' : 'Sheet updated');
  } catch (e) {
    toast(e.message, true);
  } finally {
    btn.disabled = false;
    btn.textContent = original;
  }
}

$('searchInput').addEventListener('input', render);
$('modal').addEventListener('click', e => { if (e.target.id === 'modal') closeModal(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

init();
