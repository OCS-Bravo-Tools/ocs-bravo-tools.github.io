/** @OnlyCurrentDoc */
/**
 * Peer evaluation · Collector v2, the shared sheet for cadre mode (Google Apps Script, free).
 * One Google Sheet serves every cadre. Its web app address is built into the site, never into
 * a link. @OnlyCurrentDoc limits it to this one sheet; appsscript.json asks for nothing more.
 *
 * Set up once by the maintainer, about 15 minutes (the steps are in the site's docs):
 *   new Sheet > Extensions > Apps Script > paste this file and appsscript.json > Deploy >
 *   New deployment > Web app > Execute as: Me > Who has access: Anyone > Deploy > Allow.
 *   Later changes: Deploy > Manage deployments > Edit > New version (the address stays).
 *
 * What it keeps: sealed data only. The board (names and squads) is sealed with a key that
 * rides after # in the links and never reaches this sheet; results are sealed for the
 * cadre's key. It can't read either.
 *
 *   index tab   one row per evaluation: id, created, write-code hash, tabs, started (0, 1, or
 *               the start mark of the started board), board rev, first answer, last answer,
 *               sealed board, row counts, removed
 *   month tabs  2026-11, 2026-11 b, ...: one row per pick or result of the evaluations made
 *               that month (received, evaluation, oc, version, sealed envelope)
 *
 * Calls (POST, JSON as text/plain):
 *   new    {e, hw}             make an evaluation: an index row, no tab yet
 *   board  {e, w, brev, board, st} replace the sealed board (write code, newer rev only).
 *                             st: the start mark of a started board. Once started, only boards
 *                             with that mark go in; anything else gets "old" and the page
 *                             reads the started board
 *   start  {e, w}              close the squad picker (write code)
 *   state  {e}                 brev, started, used: for phones waiting for Start (cached)
 *   get    {e}                 the sealed board (cached)
 *   rows   {e, w, cur}         rows after the cursor (write code)
 *   forget {e, w}              remove the evaluation (write code)
 *   an envelope (no op)        a pick or a result, for an evaluation that has a board
 * Removing clears an evaluation's rows (all but the time) and never deletes a row of a month
 * tab, so other evaluations' row numbers never move. A month tab goes once no evaluation
 * uses it. Caps: 20 new evaluations an hour and 100 a day for the whole site, 3,000 rows an
 * evaluation, 30 per OC id, one board write per 2 s; the same envelope twice is stored once.
 */
var VERSION = 3; // 3: a started board can't be replaced by another
var MAX_BODY = 50000; // one sheet cell holds 50,000 characters
var MAX_ROWS = 3000; // rows per evaluation
var MAX_PER_OC = 30; // rows per OC id per evaluation
var NEW_PER_HOUR = 20;
var NEW_PER_DAY = 100;
var TAB_ROWS = 50000; // a month tab past this goes on in "<month> b", "<month> c", ...
var UNUSED_DAYS = 30;
var KEEP_DAYS = 365;
var DAY = 86400000;
var CACHE_S = 21600; // 6 h, the longest the script cache keeps anything
var BOARD_GAP = 2000; // at most one board write per 2 s per evaluation
var READ_ROWS = 2000; // rows scanned per rows call
var IDX = 'index';
var COLS = ['id', 'created', 'hw', 'tabs', 'started', 'brev', 'first', 'last', 'board', 'counts', 'removed'];
var ID_RE = /^[A-Za-z0-9_-]{8}$/;
var HW_RE = /^[A-Za-z0-9_-]{43}$/;

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.ping) return reply({ ok: true, app: 'peer-eval', v: VERSION });
  return reply({ ok: false, error: 'bad request' });
}

function doPost(e) {
  var body = e && e.postData ? e.postData.contents : '';
  if (!body || body.length > MAX_BODY) return reply({ ok: false, error: 'bad size' });
  var p;
  try {
    p = JSON.parse(body);
  } catch (err) {
    return reply({ ok: false, error: 'bad request' });
  }
  if (!p || typeof p !== 'object') return reply({ ok: false, error: 'bad request' });
  if (!p.op) return reply(post(p, body));
  if (!ID_RE.test(String(p.e))) return reply({ ok: false, error: 'bad request' });
  if (p.op === 'state') return reply(state(p.e));
  if (p.op === 'get') return reply(getBoard(p.e));
  if (p.op === 'new') return reply(make(p));
  if (p.op === 'board') return reply(putBoard(p));
  if (p.op === 'start') return reply(start(p));
  if (p.op === 'rows') return reply(rows(p));
  if (p.op === 'forget') return reply(forget(p));
  return reply({ ok: false, error: 'bad request' });
}

/* ---------- the index ---------- */

function book() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function index() {
  var sh = book().getSheetByName(IDX);
  if (!sh) {
    sh = book().insertSheet(IDX, 0);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    trim(sh, COLS.length);
  }
  return sh;
}

function trim(sh, cols) {
  if (sh.getMaxColumns() > cols) sh.deleteColumns(cols + 1, sh.getMaxColumns() - cols);
}

/** The index row number of an evaluation, or 0. */
function rowOf(id) {
  var cache = CacheService.getScriptCache();
  var hit = Number(cache.get('r:' + id) || 0);
  var sh = index();
  if (hit && String(sh.getRange(hit, 1).getValue()) === id) return hit;
  if (sh.getLastRow() < 2) return 0;
  var found = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).matchCase(true).findNext();
  if (!found) return 0;
  cache.put('r:' + id, String(found.getRow()), CACHE_S);
  return found.getRow();
}

function readRow(n) {
  var v = index().getRange(n, 1, 1, COLS.length).getValues()[0];
  var o = {};
  COLS.forEach(function (c, i) {
    o[c] = v[i];
  });
  o.tabs = o.tabs ? JSON.parse(o.tabs) : [];
  o.counts = o.counts ? JSON.parse(o.counts) : { n: 0, r: {} };
  var st = Number(o.started) || 0;
  o.stok = st > 1 ? st : 0; // the start mark of the started board on the sheet
  o.started = st >= 1;
  o.brev = Number(o.brev) || 0;
  o.created = Number(o.created) || 0;
  o.first = Number(o.first) || 0;
  o.last = Number(o.last) || 0;
  o.removed = Number(o.removed) || 0;
  return o;
}

function writeRow(n, o) {
  index().getRange(n, 1, 1, COLS.length).setValues([[o.id, o.created, o.hw, JSON.stringify(o.tabs), o.stok || (o.started ? 1 : 0), o.brev, o.first || '', o.last || '', o.board || '', JSON.stringify(o.counts), o.removed || '']]);
  cacheState(o);
}

function stateOf(o) {
  if (o.removed) return { ok: false, error: 'removed', removedAt: o.removed, created: o.created };
  return { ok: true, brev: o.brev, started: o.started, used: !!o.first, created: o.created, last: o.last || 0 };
}

function cacheState(o) {
  var cache = CacheService.getScriptCache();
  cache.put('s:' + o.id, JSON.stringify(stateOf(o)), CACHE_S);
  if (o.board && !o.removed) cache.put('g:' + o.id, o.board, CACHE_S);
  else cache.remove('g:' + o.id);
}

function hash(text) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)).replace(/=+$/, '');
}

function locked(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'busy' };
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** The evaluation and its index row, after checking the write code. */
function owned(p) {
  var n = rowOf(String(p.e));
  if (!n) return { err: { ok: false, error: 'none' } };
  var o = readRow(n);
  if (o.removed) return { err: stateOf(o) };
  if (!p.w || hash(String(p.w)) !== o.hw) return { err: { ok: false, error: 'denied' } };
  return { n: n, o: o };
}

/* ---------- calls ---------- */

function state(id) {
  var hit = CacheService.getScriptCache().get('s:' + id);
  if (hit) return JSON.parse(hit);
  var n = rowOf(id);
  if (!n) return { ok: false, error: 'none' };
  var o = readRow(n);
  cacheState(o);
  return stateOf(o);
}

function getBoard(id) {
  var cache = CacheService.getScriptCache();
  var s = state(id);
  if (!s.ok) return s;
  var b = cache.get('g:' + id);
  if (!b) {
    var o = readRow(rowOf(id));
    b = o.board;
    if (b) cache.put('g:' + id, b, CACHE_S);
  }
  if (!b) return { ok: false, error: 'none' };
  return { ok: true, brev: s.brev, started: s.started, used: s.used, created: s.created, last: s.last, board: JSON.parse(b) };
}

function make(p) {
  // an id Sheets would store as a number ("00012345", "1234E567") could not be found again
  if (!HW_RE.test(String(p.hw)) || !isNaN(Number(p.e))) return { ok: false, error: 'bad request' };
  return locked(function () {
    var now = new Date();
    var props = PropertiesService.getScriptProperties();
    var hour = Utilities.formatDate(now, 'Etc/UTC', 'yyyyMMddHH');
    var day = hour.slice(0, 8);
    var h = String(props.getProperty('nhour') || '').split(':');
    var d = String(props.getProperty('nday') || '').split(':');
    var inHour = h[0] === hour ? Number(h[1]) || 0 : 0;
    var inDay = d[0] === day ? Number(d[1]) || 0 : 0;
    if (inHour >= NEW_PER_HOUR || inDay >= NEW_PER_DAY) return { ok: false, error: 'limit' };
    if (rowOf(String(p.e))) return { ok: false, error: 'taken' };
    var o = { id: String(p.e), created: now.getTime(), hw: String(p.hw), tabs: [], started: false, brev: 0, first: 0, last: 0, board: '', counts: { n: 0, r: {} }, removed: 0 };
    var sh = index();
    sh.appendRow([o.id, o.created, o.hw, '[]', 0, 0, '', '', '', JSON.stringify(o.counts), '']);
    CacheService.getScriptCache().put('r:' + o.id, String(sh.getLastRow()), CACHE_S);
    cacheState(o);
    props.setProperty('nhour', hour + ':' + (inHour + 1));
    props.setProperty('nday', day + ':' + (inDay + 1));
    sweep(5);
    return { ok: true, created: o.created };
  });
}

function putBoard(p) {
  var b = p.board;
  if (!b || typeof b !== 'object' || !b.iv || !b.ct) return { ok: false, error: 'bad request' };
  var text = JSON.stringify(b);
  if (text.length > MAX_BODY) return { ok: false, error: 'bad size' };
  var cache = CacheService.getScriptCache();
  return locked(function () {
    var a = owned(p);
    if (a.err) return a.err;
    var st = Number(p.st) > 1 ? Math.floor(Number(p.st)) : 0;
    // once started, no other board: not a sign-up board (pages from before the mark send none),
    // not a second start frozen on another device. "old" makes the page read the started one
    if (a.o.started && (!st || (a.o.stok && st !== a.o.stok))) return { ok: false, error: 'old', brev: a.o.brev };
    if (!(Number(p.brev) > a.o.brev)) return { ok: false, error: 'old', brev: a.o.brev };
    var gap = Number(cache.get('bw:' + p.e) || 0);
    if (gap && Date.now() - gap < BOARD_GAP) return { ok: false, error: 'busy' };
    a.o.brev = Number(p.brev);
    a.o.board = text;
    if (st) {
      a.o.started = true; // a started board in is the start: phones open from it
      a.o.stok = a.o.stok || st;
    }
    writeRow(a.n, a.o);
    cache.put('bw:' + p.e, String(Date.now()), 60);
    return { ok: true, brev: a.o.brev };
  });
}

function start(p) {
  return locked(function () {
    var a = owned(p);
    if (a.err) return a.err;
    a.o.started = true;
    writeRow(a.n, a.o);
    return { ok: true };
  });
}

/** Rows after the cursor [tab index, row], across the evaluation's tabs in order. */
function rows(p) {
  var a = owned(p);
  if (a.err) return a.err;
  var cur = Array.isArray(p.cur) ? p.cur : [0, 0];
  var items = [];
  var t = Math.max(0, Number(cur[0]) || 0);
  var r = Number(cur[1]) || 0;
  var scanned = 0;
  while (t < a.o.tabs.length) {
    var sh = book().getSheetByName(a.o.tabs[t][0]);
    var from = Math.max(r + 1, Number(a.o.tabs[t][1]) || 2);
    var last = sh ? sh.getLastRow() : 0;
    if (sh && last >= from) {
      var count = Math.min(last - from + 1, READ_ROWS - scanned);
      var vals = sh.getRange(from, 1, count, 5).getValues();
      vals.forEach(function (v) {
        if (String(v[1]) === a.o.id && v[4]) items.push(String(v[4]));
      });
      scanned += count;
      r = from + count - 1;
      if (from + count - 1 < last) return { ok: true, items: items, cur: [t, r], more: true };
    }
    if (t === a.o.tabs.length - 1) return { ok: true, items: items, cur: [t, Math.max(r, last)], more: false };
    t++;
    r = 0;
    if (scanned >= READ_ROWS) return { ok: true, items: items, cur: [t, 0], more: true };
  }
  return { ok: true, items: items, cur: [t, r], more: false };
}

function forget(p) {
  return locked(function () {
    var a = owned(p);
    if (a.err) return a.err;
    remove(a.n, a.o);
    return { ok: true };
  });
}

/** A pick or a result. */
function post(env, body) {
  if (env.v !== 1 || !ID_RE.test(String(env.e)) || typeof env.r !== 'number' || !env.ct || !env.mac || !env.iv || !env.epk) return { ok: false, error: 'bad envelope' };
  var cache = CacheService.getScriptCache();
  var seen = 'd:' + hash(body);
  if (cache.get(seen)) return { ok: true, again: true }; // the same envelope sent twice is stored once
  var s = state(String(env.e));
  if (!s.ok) return s;
  if (!s.brev) return { ok: false, error: 'none' };
  return locked(function () {
    if (cache.get(seen)) return { ok: true, again: true };
    var n = rowOf(String(env.e));
    var o = readRow(n);
    if (o.removed) return stateOf(o);
    var mine = Number(o.counts.r[env.r] || 0);
    if (o.counts.n >= MAX_ROWS || mine >= MAX_PER_OC) return { ok: false, error: 'full' };
    var tab = tabFor(o);
    tab.sheet.appendRow([new Date(), o.id, env.r, Number(env.rev) || 1, body]);
    var row = tab.sheet.getLastRow();
    if (!o.tabs.length || o.tabs[o.tabs.length - 1][0] !== tab.name) o.tabs.push([tab.name, row]);
    o.counts.n += 1;
    o.counts.r[env.r] = mine + 1;
    if (!o.first) o.first = Date.now();
    o.last = Date.now();
    writeRow(n, o);
    cache.put(seen, '1', CACHE_S);
    return { ok: true };
  });
}

/** The month tab this evaluation writes to now: its month's, or the next "b", "c" past TAB_ROWS. */
function tabFor(o) {
  var base = o.tabs.length ? o.tabs[o.tabs.length - 1][0] : Utilities.formatDate(new Date(o.created), 'Etc/UTC', 'yyyy-MM');
  var name = base;
  for (var i = 0; i < 26; i++) {
    var sh = book().getSheetByName(name);
    if (!sh) {
      sh = book().insertSheet(name);
      sh.getRange(1, 1, 1, 5).setValues([['received', 'evaluation', 'oc', 'version', 'sealed']]);
      trim(sh, 5);
      return { name: name, sheet: sh };
    }
    if (sh.getLastRow() < TAB_ROWS) return { name: name, sheet: sh };
    var month = base.replace(/ [a-z]$/, '');
    name = month + ' ' + String.fromCharCode(98 + i); // b, c, ...
  }
  throw new Error('month full');
}

/* ---------- removing ---------- */

/**
 * Clear an evaluation's rows, keep a tombstone, drop month tabs nobody uses. A cleared row
 * keeps its time in column A, so it is never the sheet's last row again only because it was
 * cleared: new rows still go below every row number another evaluation has already read.
 */
function remove(n, o) {
  o.tabs.forEach(function (t) {
    var sh = book().getSheetByName(t[0]);
    if (!sh || sh.getLastRow() < t[1]) return;
    var ids = sh.getRange(t[1], 2, sh.getLastRow() - t[1] + 1, 1).getValues();
    var list = [];
    ids.forEach(function (v, i) {
      if (String(v[0]) === o.id) list.push('B' + (t[1] + i) + ':E' + (t[1] + i));
    });
    for (var i = 0; i < list.length; i += 500) sh.getRangeList(list.slice(i, i + 500)).clearContent();
  });
  var tabs = o.tabs.map(function (t) {
    return t[0];
  });
  o.removed = Date.now();
  o.board = '';
  o.counts = { n: 0, r: {} };
  writeRow(n, o);
  CacheService.getScriptCache().remove('g:' + o.id);
  var used = {};
  var sh = index();
  var cnt = sh.getLastRow() - 1;
  if (cnt >= 1) {
    // only the tabs (D) and removed (K) columns, never the boards, so the lock stays short
    var gone = sh.getRange(2, 11, cnt, 1).getValues();
    sh.getRange(2, 4, cnt, 1).getValues().forEach(function (v, i) {
      if (gone[i][0]) return;
      (v[0] ? JSON.parse(v[0]) : []).forEach(function (t) {
        used[t[0]] = true;
      });
    });
  }
  tabs.forEach(function (name) {
    var tab = book().getSheetByName(name);
    if (tab && !used[name]) book().deleteSheet(tab);
  });
}

/** Up to max removals per call: unused for 30 days, or a year after the last answer; tombstones a year on. */
function sweep(max) {
  var sh = index();
  var last = sh.getLastRow();
  if (last < 2) return;
  var props = PropertiesService.getScriptProperties();
  var from = Number(props.getProperty('sweep') || 2);
  if (from > last) from = 2;
  var count = Math.min(200, last - from + 1);
  // only id and created (A:B), first and last (G:H), removed (K); never the boards
  var head = sh.getRange(from, 1, count, 2).getValues();
  var when = sh.getRange(from, 7, count, 2).getValues();
  var gone = sh.getRange(from, 11, count, 1).getValues();
  var now = Date.now();
  var done = 0;
  var dead = [];
  for (var i = 0; i < count && done < max; i++) {
    var created = Number(head[i][1]) || 0;
    var firstAt = Number(when[i][0]) || 0;
    var lastAt = Number(when[i][1]) || 0;
    var removed = Number(gone[i][0]) || 0;
    if (removed) {
      if (now - removed > KEEP_DAYS * DAY) dead.push(i);
      continue;
    }
    var expired = firstAt ? now - lastAt > KEEP_DAYS * DAY : now - created > UNUSED_DAYS * DAY;
    if (expired) {
      remove(from + i, readRow(from + i));
      done++;
    }
  }
  // tombstones go last, bottom up, so the row numbers above stay right
  var cache = CacheService.getScriptCache();
  dead.reverse().forEach(function (j) {
    cache.remove('r:' + String(head[j][0]));
    sh.deleteRow(from + j);
  });
  // next time: on from the first row not looked at, or back to the top once the last row was
  // looked at (each new evaluation adds its row just before this runs, so it must wrap)
  var next = from + i;
  props.setProperty('sweep', String(next > last ? 2 : next - dead.length));
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
