/**
 * Admin console: one page, four steps.
 *   1 Roster   drop or paste the sheets you have (a class roster, each squad's workbook, a
 *              list of names): names, PLT and SQD are filled in from them, merged by name;
 *              only what is missing is asked (a tap per OC, a question for a near-miss
 *              spelling), or the OCs pick their own squad from one link per PLT (Ask OCs);
 *              the results sheet card (Google, set up once, or skip); settings
 *   2 Links    one personal link per OC; QR slips by squad, 8 per page
 *   3 Collect  results arrive from the sheet by themselves; paste, drop files or scan as fallbacks
 *   4 Results  per-squad order of merit with the sheet's formula; Export xlsx
 *
 * Storage: the event (keys, tokens, results) is never kept in the browser, hosted
 * or opened from a file. Other pages on the same site, or every local page on
 * file://, could read it there. The event file is the only copy. It downloads at Create
 * links (with Ask OCs, a sign-up file first, which the event file replaces); results from
 * the sheet come back each time it is opened, so only results added by hand need another
 * save. Only the unfinished setup (the roster read so far with the admin's taps, settings,
 * sheet address, read key) is kept in this tab's sessionStorage, so a reload does not lose
 * it. During a sign-up the admin's work also goes to the sheet, encrypted, so reopening the
 * sign-up file anywhere brings it back. Dropped files are read here and never uploaded,
 * and only last name, first name, PLT, SQD and the file's name are kept from them.
 */
import { dealSquad } from './lib/deal.js';
import { scoreSquad } from './lib/scoring.js';
import { generateAdminKeys, randomToken, randomBytes, hmac } from './lib/crypto.js';
import { packPacket, openResult, findEnvelopes, sealWork, openMarked } from './lib/envelope.js';
import { readRoster, csvText, fileUnits, mergeRoster, numberRoster, ORD } from './lib/roster.js';
import { buildXlsx, readXlsx } from './lib/xlsx.js';
import { makeAssembler } from './lib/qrparts.js';
import { buildLinkPdf } from './lib/linkpdf.js';
import { resultsXlsx, answersXlsx, XLSX_TYPE } from './lib/report.js';

const OLD_STORE = 'pe:admin:v1';
const DRAFT = 'pe:admin:draft';
const AUTO_PULL_MS = 120000;
const SIGNUP_PULL_MS = 30000; // while a sign-up runs: the OCs' picks, every 30 s
const WORK_MS = 30000; // the admin's work during a sign-up goes to the sheet at most this often

const app = document.getElementById('admin');
let E = null;
/* The roster being made: every file or paste read (sources) and the admin's decisions (dec): squads,
   PLTs and names tapped, questions answered, Add or Ignore, a 5th squad. See mergeRoster. */
const D = { sources: [], dec: {}, seq: 0, signup: null, name: 'Peer evaluation', collectorUrl: '', readKey: '', appUrl: defaultAppUrl(), sheetSkipped: false }; // signup: a sign-up started from this tab
const ui = {
  open: new Set(), busy: '', unsaved: false, lastPull: '', scanner: null, connected: false, refused: false, sheet: '', sheetOpen: false, scriptCopied: '',
  edit: '', editLast: '', editFirst: '', held: false, // the name being edited; Create links pressed once with something to check
  ans: {}, extra: [], signupResults: false, // the OCs' sign-up picks; results already in the sheet for a sign-up file
  work: '', workAt: '', workEnv: null, workSent: 0 // the admin's sign-up work: '' | 'pending' | 'saving' | 'saved'
};
let workTimer = 0;

/* The results sheet: the collector script version this console expects, and the script
   itself for the card's copy buttons. tools/build.mjs puts Collector.gs into COLLECTOR_GS in the
   single file; the hosted console fetches it. */
const COLLECTOR_V = 1;
const COLLECTOR_GS = null;
const SHEET_URL = /^https:\/\/script\.google\.com\/(macros|a\/macros\/[^/]+)\/s\/[A-Za-z0-9_-]+\/exec$/;
const LOCAL_URL = /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?\//; // the stand-in collector used by the tests
const SHEET_TEXT = {
  checking: 'Checking…',
  ok: '✓ Sheet ready',
  reused: '✓ Sheet ready: the same sheet as your last evaluation',
  used: 'This sheet belongs to an earlier evaluation. Open that evaluation\'s event file and press New evaluation, or set up a new sheet (3 min).',
  dev: 'That is the test URL (/dev), which only works for you. Copy the Web app URL that ends in /exec.',
  editor: 'That is the script editor. Deploy › Manage deployments, then copy the Web app URL (it ends in /exec).',
  sheet: 'That is the sheet itself. Do steps 2 and 3, then paste the Web app URL (it ends in /exec).',
  wrongurl: 'Paste the Web app URL from step 3: https://script.google.com/macros/s/…/exec',
  noanswer: 'No answer. Set Who has access to Anyone: Deploy › Manage deployments › Edit. If this network blocks Google, skip the sheet.',
  notscript: 'That web app is not the results script. Paste the script (step 2), Save, then Deploy › Manage deployments › Edit › Version: New version.',
  old: 'This sheet runs an older script. Press Copy again (step 2), paste it over everything, Save, then Deploy › Manage deployments › Edit › Version: New version.'
};
const SHEET_BLOCKS = ['used', 'dev', 'editor', 'sheet', 'wrongurl', 'notscript', 'old'];
const SCRIPT_TEXT = { ok: '✓ Script copied', no: 'Not copied: press Copy again' };
let sheetTimer = 0;
let scriptText = COLLECTOR_GS;

const esc = (s) => String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = (x, d) => (x === null || x === undefined || !isFinite(x) ? '—' : Number(x).toFixed(d));
const memberOf = (id) => E.members.find((m) => m.id === id); // ids fixed at Ask OCs can leave gaps
const nameOf = (id) => (memberOf(id) ? memberOf(id).name : '?');
const squadOf = (m) => E.squads.find((s) => s.key === m.squad);
const squadResults = (sq) => sq.ids.map((id) => E.results[id]).filter(Boolean).map((r) => r.payload);
const live = () => !!E && E.stage !== 'signup'; // links exist (a sign-up file has none yet)
const PLT = (n) => ORD(n) + ' PLT';
const letters = (s) => String(s || '').toUpperCase().replace(/[^\p{L}]/gu, '');
const fullName = (o) => o.last + (o.first ? ', ' + o.first : '');
const byName = (a, b) => a.last.localeCompare(b.last) || a.first.localeCompare(b.first);
const ICON_PEN = '<svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#4A4F57" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20 L8 19 L19 8 L16 5 L5 16 Z"/></svg>';

function defaultAppUrl() {
  if (!/^https?:/.test(location.protocol)) return '';
  return new URL('./', location.href).href;
}

/* ---------- storage ---------- */

/** Mark the event changed (Save event file *) and keep the setup draft for this tab. */
function persist() {
  ui.unsaved = true;
  try {
    sessionStorage.setItem(DRAFT, JSON.stringify({ sources: D.sources, dec: D.dec, seq: D.seq, signup: D.signup, name: D.name, collectorUrl: D.collectorUrl, readKey: D.readKey, appUrl: D.appUrl, sheetSkipped: D.sheetSkipped }));
  } catch (err) {
    // only a convenience; the event file is the real copy
  }
}

function restore() {
  try {
    localStorage.removeItem(OLD_STORE);
  } catch (err) {
    // nothing to clean up
  }
  try {
    const d = JSON.parse(sessionStorage.getItem(DRAFT) || 'null') || {};
    const old = typeof d.roster === 'string' && d.roster.trim() && !d.sources ? d : null; // a Build 09 tab: its pasted roster becomes a list
    ['roster', 'rosterPlt', 'rosterSqd', 'parsed'].forEach((k) => delete d[k]);
    Object.assign(D, d);
    if (old) D.sources = [sourceOf(readRoster(old.roster), 'Pasted list', { plt: Number(old.rosterPlt) || null, sqd: Number(old.rosterSqd) || null })];
  } catch (err) {
    // start empty
  }
}

/**
 * A file or paste as the roster keeps it: the rows read (last, first, PLT, SQD only) and
 * what was not used, under a new source id. PLT and SQD from the file name fill rows with none.
 */
function sourceOf(r, name, units) {
  D.seq = (D.seq || 0) + 1;
  return {
    id: 's' + D.seq, name,
    rows: r.rows.map((x) => ({ last: x.last, first: x.first || '', plt: x.plt || null, plts: x.plts || null, sqd: x.sqd || null, two: !!x.two })),
    unread: r.unread || [], warn: r.warn || [], columns: r.columns || [], empty: r.empty || 0, titles: r.titles || 0, tabs: r.tabs || 0, tab: r.tab || '',
    plt: units ? units.plt : null, sqd: units ? units.sqd : null
  };
}

/** The roster as it stands: every source merged, with the admin's decisions and the OCs' sign-up picks. */
const rosterNow = () => mergeRoster(D.sources, D.dec, Object.assign({ extra: ui.extra }, ui.ans), E && E.stage === 'signup' ? E.ids : {}, E && E.stage === 'signup' ? E.listed : null);

function log(text) {
  if (!E) return;
  E.log = E.log || [];
  E.log.unshift({ at: new Date().toISOString(), text });
  E.log = E.log.slice(0, 200);
}

/* ---------- helpers ---------- */

function toast(text) {
  document.querySelectorAll('.toast').forEach((x) => x.remove());
  const t = document.createElement('div');
  t.className = 'toast';
  t.setAttribute('role', 'status');
  t.textContent = text;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 4000);
}

function download(name, data, type) {
  const url = URL.createObjectURL(new Blob([data], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

/** @returns {Promise<boolean>} false when the browser refused and the text was shown to copy by hand */
async function copy(text, what) {
  try {
    await navigator.clipboard.writeText(text);
    toast('Copied ' + what + '.');
    return true;
  } catch (err) {
    const dlg = document.createElement('dialog');
    dlg.innerHTML = '<p><b>Copy ' + esc(what) + '</b></p><textarea readonly style="min-height:240px"></textarea><div class="row" style="margin-top:10px"><button class="btn primary">Done</button></div>';
    document.body.appendChild(dlg);
    const box = dlg.querySelector('textarea');
    box.value = text;
    dlg.querySelector('button').addEventListener('click', () => {
      dlg.close();
      dlg.remove();
    });
    dlg.showModal();
    box.focus();
    box.select();
    return false;
  }
}

const slug = (s) => String(s).replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '');

/* ---------- event ---------- */

function packetFor(m) {
  const sq = squadOf(m);
  return {
    v: 1, e: E.id, n: E.name, g: sq.label, me: m.id, nm: m.name,
    p: sq.ids.filter((id) => id !== m.id).map((id) => [id, memberOf(id).name]),
    a: m.hand, k: E.keys.pub, t: m.token, u: E.collector.url || ''
  };
}

async function buildLinks() {
  const base = E.appUrl.replace(/#.*$/, '');
  for (const m of E.members) {
    m.link = squadOf(m).ids.length > 1 ? base + '#' + (await packPacket(packetFor(m))) : '';
  }
}

/**
 * Every OC's personal link. After Ask OCs the evaluation already exists: its id and keys,
 * every OC's id and token and the deal seed were fixed then, so the same roster always gives
 * the same links. OCs added since get new ids, and tokens derived from the sign-up's own key.
 */
async function createEvent(R) {
  const appUrl = String(D.appUrl || '').trim();
  if (appUrl && !/^https?:\/\//.test(appUrl)) {
    toast('The OC page address must start with https:// (or leave it empty).');
    return;
  }
  ui.busy = 'Creating links…';
  render();
  const signup = E && E.stage === 'signup' ? E : null;
  if (signup && ui.workEnv && ui.work !== 'saved' && signup.collector.url) {
    // the roster these links are made from goes to the sheet first, so the sign-up file opened anywhere gives the same links
    clearTimeout(workTimer);
    workTimer = 0;
    try {
      await fetch(signup.collector.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(ui.workEnv), redirect: 'follow' });
      ui.work = 'saved';
    } catch (err) {
      // the sign-up file saved below still holds it
    }
  }
  const keys = signup ? signup.keys : await generateAdminKeys();
  const seed = signup ? signup.seed : new DataView(randomBytes(4).buffer).getUint32(0) % 1000000;
  const n = numberRoster(R, signup ? signup.ids : {});
  const members = [];
  for (const m of n.members) {
    const token = !signup ? randomToken(16) : signup.tokens[m.id] || (await hmac(signup.join.work, 'oc ' + m.id)).slice(0, 22);
    members.push(Object.assign({}, m, { token, hand: [] }));
  }
  const squads = n.squads.map((sq, si) => {
    const d = dealSquad(sq.ids.length, seed + si * 101);
    sq.ids.forEach((id, i) => {
      members.find((m) => m.id === id).hand = d.hands[i] || [];
    });
    return Object.assign({}, sq, { k: d.k, seed: d.seed });
  });
  const url = D.sheetSkipped ? '' : D.collectorUrl.trim();
  E = Object.assign(signup || {}, {
    v: 1, stage: 'links', id: signup ? signup.id : randomToken(6), name: D.name || 'Peer evaluation', created: signup ? signup.created : new Date().toISOString(),
    appUrl, collector: { url, readKey: url ? (signup && signup.collector.readKey) || D.readKey || randomToken(18) : '' }, keys, seed, members, squads, results: {}, log: (signup && signup.log) || []
  });
  delete E.roster; // the members now hold the roster
  D.signup = null;
  await buildLinks();
  log('Created ' + members.length + ' links in ' + squads.length + ' squads');
  ui.busy = '';
  ui.connected = false;
  ui.refused = false;
  ui.held = false;
  persist();
  render();
  saveEventFile();
  const replaces = signup ? ' This event file replaces the sign-up file: keep this one, and a second copy where only you can open it.' : '';
  if (url) await connectCollector('Links created.' + replaces + ' ');
  else toast('Links created. Keep the event file and a second copy where only you can open them: it is the only key to the results.' + replaces);
}

/** The event file (…-event-date.json), or during a sign-up the sign-up file (…-signup-date.json) with the roster work so far. */
function saveEventFile() {
  if (!E) return;
  E.savedAt = new Date().toISOString();
  if (E.stage === 'signup') E.roster = { sources: D.sources, dec: D.dec, seq: D.seq };
  download('peer-eval-' + slug(E.name) + '-' + E.id + (E.stage === 'signup' ? '-signup-' : '-event-') + E.savedAt.slice(0, 10) + '.json', JSON.stringify(E), 'application/json');
  ui.unsaved = false;
  persist();
  ui.unsaved = false;
  render();
}

function openEventFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const obj = JSON.parse(String(reader.result));
      if (!obj || obj.v !== 1 || !obj.id || !obj.keys || !obj.keys.priv || !Array.isArray(obj.members)) throw new Error('bad');
      if (E && ui.unsaved && !window.confirm('Open "' + obj.name + '"? Changes to "' + E.name + '" since you last saved its event file are lost.')) return;
      E = obj;
      E.results = E.results || {};
      if (E.stage === 'signup') {
        // the sign-up file: the roster work as it was saved; the sheet may hold newer work and the OCs' picks
        const r = E.roster || {};
        Object.assign(D, { sources: r.sources || [], dec: r.dec || {}, seq: r.seq || (r.sources || []).length, name: E.name, appUrl: E.appUrl });
        const a = E.answers || {};
        Object.assign(ui, { ans: a.ans || {}, extra: a.extra || [], signupResults: false, work: '', workEnv: null });
      }
      D.collectorUrl = E.collector.url || '';
      D.readKey = E.collector.readKey || '';
      D.sheetSkipped = !E.collector.url;
      persist();
      Object.assign(ui, { unsaved: false, connected: false, refused: false, lastPull: '', held: false, edit: '' });
      render();
      toast(E.stage === 'signup' ? 'Opened the sign-up for ' + E.name + '. Picks and your later work come back from the sheet.' : 'Opened ' + E.name + ' (' + E.members.length + ' OCs).');
      if (E.collector.url && E.collector.readKey) pull(true);
    } catch (err) {
      toast('That is not an event file from this app.');
    }
  };
  reader.readAsText(file);
}

/**
 * Start the next evaluation from the open one: same roster (ready to edit) and the same
 * results sheet with its read key, so nothing needs doing on the Google side. With
 * newSheet the sheet is left out and its card opens. The new evaluation gets its own id,
 * so its event file never replaces the old one.
 */
function newEvaluation(newSheet) {
  if (!E) return;
  if (ui.unsaved && !window.confirm('Start a new evaluation? Results you added by hand to "' + E.name + '" since you last saved its event file are lost.')) return;
  D.sources = [];
  D.dec = {};
  D.seq = 0;
  const rows = E.members.map((m) => ({ last: m.last, first: m.first || '', plt: m.plt, sqd: m.sqd }));
  D.sources.push(sourceOf({ rows }, 'Last evaluation (' + E.name + ')', null));
  D.signup = null;
  D.name = 'Peer evaluation ' + new Date().toISOString().slice(0, 10);
  D.collectorUrl = newSheet ? '' : E.collector.url || '';
  D.readKey = newSheet ? '' : E.collector.readKey || '';
  D.sheetSkipped = !newSheet && !E.collector.url;
  E = null;
  Object.assign(ui, { unsaved: false, connected: false, refused: false, lastPull: '', sheet: '', sheetOpen: !!newSheet, ans: {}, extra: [], signupResults: false, held: false, edit: '', work: '', workEnv: null });
  persist();
  render();
  if (D.collectorUrl) checkSheet();
  toast(newSheet ? 'Set up a new results sheet, then press Create links.' : 'Same roster and results sheet. Edit the roster if it changed, then press Create links.');
}

/* ---------- collector ---------- */

/** A read from the sheet. Two quiet retries, after 1 s and 3 s (P3-2): Google answers an error now and then. */
async function callCollector(url, params) {
  const sep = url.indexOf('?') >= 0 ? '&' : '?';
  let last = null;
  for (const wait of [0, 1000, 3000]) {
    if (wait) await new Promise((r) => setTimeout(r, wait));
    try {
      const res = await fetch(url + sep + new URLSearchParams(params).toString(), { cache: 'no-store', redirect: 'follow' });
      return JSON.parse(await res.text());
    } catch (err) {
      last = err; // no answer, or Google's own error page
    }
  }
  throw last;
}

/**
 * Before links exist: check a pasted sheet address by itself, with one line per mistake.
 * Only a ping is sent, so nothing is set and closing the tab loses nothing. A sheet that
 * already has a key is reused only when this console holds that key (opened from the
 * last event file); asking with it changes nothing on a sheet that has a key.
 */
function urlProblem(url) {
  if (SHEET_URL.test(url) || LOCAL_URL.test(url)) return '';
  if (/^https:\/\/script\.google\.com\/.*\/dev\/?$/.test(url)) return 'dev';
  if (/^https:\/\/docs\.google\.com\/spreadsheets\//.test(url)) return 'sheet';
  if (/^https:\/\/script\.google\.com\//.test(url)) return 'editor';
  return 'wrongurl';
}

async function checkSheet() {
  const url = D.collectorUrl.trim();
  const show = (state) => {
    if (D.collectorUrl.trim() !== url) return; // pasted again meanwhile
    ui.sheet = state;
    if (state === 'ok' || state === 'reused') {
      ui.sheetOpen = false;
      render();
      return;
    }
    const el = document.getElementById('sheet-status');
    if (el) {
      el.className = 'sheetstatus' + (state && state !== 'checking' ? ' bad' : '');
      el.textContent = SHEET_TEXT[state] || '';
    }
  };
  if (!url) return show('');
  const bad = urlProblem(url);
  if (bad) return show(bad);
  show('checking');
  let r;
  try {
    r = await callCollector(url, { ping: 1 });
  } catch (err) {
    return show('noanswer'); // no network, a login page (Who has access is not Anyone), or no script
  }
  if (!r || r.app !== 'peer-eval') return show('notscript');
  if (!r.v || r.v < COLLECTOR_V) return show('old');
  if (!r.keyed) return show('ok');
  if (!D.readKey) return show('used');
  try {
    const s = await callCollector(url, { setup: D.readKey });
    return show(s.ok ? 'reused' : 'used');
  } catch (err) {
    return show('noanswer');
  }
}

/** The collector script: inside the single file; fetched once by the hosted console, ahead of the click. */
async function loadScript() {
  if (!scriptText) {
    const res = await fetch('apps-script/Collector.gs');
    if (!res.ok) throw new Error(String(res.status));
    scriptText = await res.text();
  }
  return scriptText;
}

/**
 * Loaded ahead, the clipboard write starts inside the click, before the new sheet's tab
 * takes the focus. The card then says whether it worked, for when the admin comes back.
 */
async function copyScript() {
  let text = '';
  try {
    text = scriptText || (await loadScript());
  } catch (err) {
    toast('Could not load the script here. Use PeerEval_Admin.html, which carries it.');
  }
  ui.scriptCopied = text && (await copy(text, 'the script. Paste it over everything in Apps Script, then Save')) ? 'ok' : 'no';
  const el = document.getElementById('script-status');
  if (el) {
    el.className = 'sheetstatus' + (ui.scriptCopied === 'ok' ? '' : ' bad');
    el.textContent = SCRIPT_TEXT[ui.scriptCopied];
  }
}

/**
 * After links exist: set the evaluation's read key on the sheet and send a test result
 * through. The key is already in the event file when the sheet first sees it, and the
 * sheet keeps the first key it gets, so nobody holding an OC link (which carries the
 * sheet's address) can change it. Until this passes, the Links section says not to hand
 * out links; it is tried again on every pull (every 2 minutes, and when the event file is
 * opened) and by Try again.
 * @returns {Promise<'ok'|'refused'|'unreachable'>}
 */
async function confirmSheet() {
  const url = E.collector.url;
  const key = E.collector.readKey;
  ui.connected = false;
  try {
    const r = await callCollector(url, { setup: key });
    ui.refused = !r.ok;
    if (!r.ok) return 'refused';
    const probe = 'ping-' + randomToken(4);
    const post = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ v: 1, e: probe, r: -1, rev: 1, z: 0, epk: 'x', iv: 'x', ct: 'x', mac: 'x' }), redirect: 'follow' });
    const posted = JSON.parse(await post.text());
    const back = await callCollector(url, { k: key, e: probe });
    ui.connected = !!(posted.ok && back.ok && Array.isArray(back.items) && back.items.length >= 1);
  } catch (err) {
    // no answer: tried again later
  }
  return ui.connected ? 'ok' : 'unreachable';
}

async function connectCollector(prefix) {
  if (!E || !E.collector.url) return;
  ui.busy = 'Checking the sheet…';
  render();
  const s = await confirmSheet();
  ui.busy = '';
  render();
  toast((prefix || '') + (s === 'ok' ? 'Connected. Test result went in and came back.'
    : s === 'refused' ? 'This sheet belongs to another evaluation. Don\'t hand out these links.'
      : 'The test result did not go through. Don\'t hand out links until the warning under Links goes away.'));
}

async function pull(quiet) {
  if (!E || !E.collector.url || !E.collector.readKey) return;
  if (!ui.connected) {
    const s = await confirmSheet();
    if (s !== 'ok') {
      render();
      if (!quiet) toast(s === 'refused' ? 'This sheet belongs to another evaluation, so it won\'t release these results.' : 'Can\'t get new answers right now. It tries again by itself in 2 minutes.');
      return;
    }
  }
  ui.busy = quiet ? ui.busy : 'Pulling…';
  if (!quiet) render();
  try {
    const r = await callCollector(E.collector.url, { k: E.collector.readKey, e: E.id });
    if (r && r.error === 'unauthorized') {
      ui.connected = false; // checked again on the next pull
      if (!quiet) toast('The sheet did not accept this evaluation\'s read key. It is checked again on the next pull.');
      ui.busy = '';
      render();
      return;
    }
    if (!r.ok || !Array.isArray(r.items)) throw new Error(r.error || 'bad reply');
    ui.connected = true;
    const envs = [];
    r.items.forEach((item) => {
      try {
        envs.push(typeof item === 'string' ? JSON.parse(item) : item);
      } catch (err) {
        // skip damaged rows
      }
    });
    ui.lastPull = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    if (E.stage === 'signup') await importSignup(envs, quiet);
    else await importEnvelopes(envs, 'sheet', quiet);
  } catch (err) {
    if (!quiet) toast('Can\'t get new answers right now. It tries again by itself in 2 minutes.');
  }
  ui.busy = '';
  render();
}

/* ---------- import ---------- */

async function importEnvelopes(envs, source, quiet) {
  if (!E) return;
  const c = { added: 0, replaced: 0, same: 0, older: 0, rejected: 0 };
  const reasons = {};
  for (const env of envs) {
    if (String(env.e).indexOf('ping-') === 0 || env.j) continue; // a squad pick or the admin's sign-up work: not a result
    const o = await openResult(E, env);
    if (!o.ok) {
      c.rejected++;
      reasons[o.reason] = (reasons[o.reason] || 0) + 1;
      continue;
    }
    const cur = E.results[env.r];
    const dv = String(o.payload.dv || '');
    const devs = cur ? (cur.devs || []).slice() : [];
    if (dv && devs.indexOf(dv) < 0) devs.push(dv);
    if (cur) cur.devs = devs; // every version counts, older ones too: one OC from two devices gets flagged
    if (cur && cur.hash === o.hash) {
      c.same++;
      continue;
    }
    if (cur && (cur.rev > env.rev || (cur.rev === env.rev && String(o.payload.at) <= String(cur.at)))) {
      c.older++;
      continue;
    }
    E.results[env.r] = { rev: env.rev, at: o.payload.at, hash: o.hash, source, payload: o.payload, devs };
    if (cur) c.replaced++;
    else c.added++;
  }
  const parts = [];
  if (c.added) parts.push(c.added + ' new');
  if (c.replaced) parts.push(c.replaced + ' newer versions');
  if (c.same) parts.push(c.same + ' already in');
  if (c.older) parts.push(c.older + ' older ignored');
  if (c.rejected) parts.push(c.rejected + ' rejected (' + Object.keys(reasons).map((k) => k + ' ×' + reasons[k]).join(', ') + ')');
  const text = (source === 'sheet' ? 'Pulled ' : 'Imported ') + envs.length + ': ' + (parts.join(', ') || 'nothing new');
  if (c.added || c.replaced || c.rejected || !quiet) log(text);
  if ((c.added || c.replaced) && source !== 'sheet') persist(); // the sheet's results come back when the file is opened
  if (!quiet) toast(text);
}

/**
 * The sign-up so far, from the sheet: each phone's newest pick (by the phone's own rev, then
 * by arrival), one answer per OC (two phones with two squads for one OC are flagged), names
 * typed in by OCs not on the list, and the admin's own newest work from any computer. A result
 * for this evaluation means its links were made from another file: Create links stays off here.
 */
async function importSignup(envs, quiet) {
  const phones = {};
  let work = null;
  const known = new Set(Object.values(E.ids || {}).map(Number));
  ui.signupResults = envs.some((env) => !env.j && String(env.e) === E.id);
  for (let i = 0; i < envs.length; i++) {
    const env = envs[i];
    if (!env.j || String(env.e) !== E.id) continue;
    const o = await openMarked(E, env);
    if (!o.ok) continue;
    const p = o.payload;
    if (env.j === 2) {
      if (!work || p.rev >= work.rev) work = p;
      continue;
    }
    const sqdOk = Number.isInteger(p.sqd) && p.sqd >= 1 && p.sqd <= 20;
    const who = p.r >= 0 ? known.has(p.r) : Array.isArray(p.grp) ? p.grp.every((id) => known.has(id)) && !!p.first : typeof p.last === 'string' && /\p{L}/u.test(p.last) && p.last.length <= 40;
    if (!sqdOk || !who) continue;
    const dv = String(p.dv || 'x' + i);
    if (!phones[dv] || p.rev >= phones[dv].rev) phones[dv] = Object.assign({}, p, { plt: o.plt });
  }
  const byId = {};
  const groups = {};
  const extra = [];
  Object.keys(phones).forEach((dv) => {
    const p = phones[dv];
    if (p.r >= 0) (byId[p.r] = byId[p.r] || []).push(p);
    else if (p.grp) (groups[p.grp.join(',')] = groups[p.grp.join(',')] || []).push(p);
    else extra.push({ key: 'j:' + dv, last: String(p.last).trim().toUpperCase(), first: '', plt: p.plt, sqd: p.sqd });
  });
  // OCs sharing a last name with no first names: each pick, in first-name order, goes to the next of them
  Object.keys(groups).forEach((g) => {
    const ids = g.split(',').map(Number);
    groups[g].sort((a, b) => String(a.first).localeCompare(String(b.first))).forEach((p, i) => {
      if (ids[i] !== undefined) (byId[ids[i]] = byId[ids[i]] || []).push(Object.assign({}, p, { first: String(p.first).trim() }));
    });
  });
  const ans = {};
  Object.keys(byId).forEach((id) => {
    const list = byId[id];
    const sqds = list.map((p) => p.sqd).filter((x, k, a) => a.indexOf(x) === k).sort((a, b) => a - b);
    const p = list[list.length - 1];
    ans[id] = sqds.length > 1 ? { twice: sqds, plt: p.plt } : { sqd: p.sqd, plt: p.plt, first: p.first || '' };
  });
  ui.ans = ans;
  ui.extra = extra;
  E.answers = { ans, extra };
  if (work && work.rev > (E.join.wrev || 0) && Array.isArray(work.sources)) {
    Object.assign(D, { sources: work.sources, dec: work.dec || {}, seq: Math.max(D.seq || 0, ...work.sources.map((s) => Number(String(s.id).slice(1)) || 0)) });
    E.join.wrev = work.rev;
    persist();
    if (!quiet) toast('Your later work on this sign-up came back from the sheet.');
  }
  if (!quiet) toast(Object.keys(ans).length + extra.length + ' sign-up answers so far.');
}

/** The admin's work during a sign-up, sealed for the sheet right away; it goes out at most every 30 s, and once more on leaving. */
async function markWork() {
  if (!E || E.stage !== 'signup' || !E.collector.url) return;
  E.join.wrev = (E.join.wrev || 0) + 1;
  const payload = { v: 1, j: 2, e: E.id, r: -2, rev: E.join.wrev, at: new Date().toISOString(), sources: D.sources, dec: D.dec };
  ui.workEnv = await sealWork(E, payload);
  ui.work = 'pending';
  showWork();
  if (!workTimer) workTimer = setTimeout(flushWork, Math.max(0, WORK_MS - (Date.now() - ui.workSent)));
}

async function flushWork() {
  workTimer = 0;
  const env = ui.workEnv;
  if (!env || ui.work !== 'pending' || !E || !E.collector.url) return;
  ui.work = 'saving';
  showWork();
  ui.workSent = Date.now();
  let ok = false;
  try {
    const res = await fetch(E.collector.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(env), redirect: 'follow' });
    ok = !!JSON.parse(await res.text()).ok;
  } catch (err) {
    // tried again below
  }
  if (ui.workEnv === env) {
    ui.work = ok ? 'saved' : 'pending';
    if (ok) ui.workAt = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  } else ui.work = 'pending'; // changed meanwhile: the newer one goes next
  showWork();
  if (ui.work === 'pending' && !workTimer) workTimer = setTimeout(flushWork, WORK_MS);
}

const workText = () => (ui.work === 'saved' ? 'Your own taps and answers are saved in your sheet too, encrypted · last saved ' + ui.workAt + '.'
  : ui.work ? 'Saving…' : 'Your own taps and answers are saved in your sheet too, encrypted.');

function showWork() {
  const el = document.getElementById('work-status');
  if (el) el.textContent = workText();
}

/** After the admin changed the roster: keep the draft, save the sign-up work, draw again. */
function changed() {
  ui.held = false;
  persist();
  markWork();
  render();
}

async function importText(text, source) {
  const envs = findEnvelopes(text);
  if (!envs.length) {
    toast('No results found in that text.');
    return;
  }
  await importEnvelopes(envs, source);
  render();
}

async function importFiles(files) {
  let text = '';
  for (const f of files) text += '\n' + (await f.text());
  await importText(text, 'file');
}

/* ---------- scanner ---------- */

async function openScanner() {
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia || typeof jsQR !== 'function') {
    toast('This computer has no camera access. Use paste or files.');
    return;
  }
  const dlg = document.createElement('dialog');
  dlg.innerHTML = '<p><b>Scan a phone</b></p><p class="muted">Hold the OC\'s Done screen up to the camera until all codes are read.</p>' +
    '<video playsinline muted></video><p id="scan-status" class="muted">Starting camera…</p>' +
    '<div class="row" style="margin-top:10px"><button class="btn">Close</button></div>';
  document.body.appendChild(dlg);
  const video = dlg.querySelector('video');
  const status = dlg.querySelector('#scan-status');
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  const asm = makeAssembler();
  let stream = null;
  let loop = 0;
  const stop = () => {
    clearInterval(loop);
    if (stream) stream.getTracks().forEach((t) => t.stop());
    dlg.close();
    dlg.remove();
    render();
  };
  dlg.querySelector('button').addEventListener('click', stop);
  dlg.addEventListener('cancel', stop);
  dlg.showModal();
  try {
    stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' }, audio: false });
  } catch (err) {
    status.textContent = 'The camera is blocked on this computer. Use paste or files instead.';
    return;
  }
  video.srcObject = stream;
  await video.play().catch(() => {});
  status.textContent = 'Looking for codes…';
  loop = setInterval(async () => {
    if (!video.videoWidth) return;
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    ctx.drawImage(video, 0, 0);
    const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const code = jsQR(img.data, img.width, img.height, { inversionAttempts: 'dontInvert' });
    if (!code || !code.data) return;
    const r = asm.add(code.data);
    if (!r.ok) return;
    if (!r.text) {
      status.textContent = 'Read ' + r.have + ' of ' + r.total + ' codes…';
      return;
    }
    const envs = findEnvelopes(r.text);
    if (!envs.length) return;
    await importEnvelopes(envs, 'qr', true);
    const m = memberOf(envs[0].r);
    status.textContent = (m ? m.name : 'Result') + ' ✓ received. Next phone.';
  }, 150);
}

/* ---------- export ---------- */

function exportXlsx() {
  const f = resultsXlsx(E);
  download(f.name, f.data, XLSX_TYPE);
  log('Exported xlsx');
}

/**
 * The answers so far, as each OC sent them: an Index tab (every OC: in or not yet), then
 * one tab per OC whose result is in (the newest version), read straight from the opened
 * results with no scoring and no formulas. It can be saved at any time during collection,
 * so a readable copy exists even if a later step fails. Like Export xlsx it holds names,
 * scores and comments: keep it where only you can open it.
 */
function saveAnswers() {
  const c = counts();
  const f = answersXlsx(E, c, new Date());
  download(f.name, f.data, XLSX_TYPE);
  log('Saved the answers so far (' + c.done + ' of ' + c.total + ' in)');
}

/* ---------- printing ---------- */

function slipHtml(m) {
  const qr = qrcode(0, 'L');
  qr.addData(m.link, 'Byte');
  qr.make();
  return '<div class="slip">' + qr.createSvgTag({ cellSize: 2, margin: 0, scalable: true }) +
    '<div><b>' + esc(m.name) + '</b><span class="sq">' + esc(squadOf(m).label) + ' · ' + esc(E.name) + '</span>' +
    'Scan with your phone camera; open in Safari or Chrome. Personal: do not share. Open it at least once a week until you submit.</div></div>';
}

/** The link list: one PDF with every OC's link, to upload once where all OCs can open it. */
function linkList() {
  if (!E.appUrl) {
    toast('The link list needs an OC page address, set before creating links. Send each OC their message instead.');
    return;
  }
  const groups = E.squads.filter((sq) => sq.ids.length > 1).map((sq) => ({ label: sq.label, rows: sq.ids.map((id) => ({ name: memberOf(id).name, link: memberOf(id).link })) }));
  const pdf = buildLinkPdf({
    title: 'Peer evaluation links · ' + E.name,
    notes: [
      'Find your name and tap it: your own evaluation opens. Use Safari or Chrome.',
      'Use only your own link. Each link counts as that OC\'s evaluation, and the admin',
      'can see when one OC\'s evaluation comes from more than one device.',
      'Answers save on your phone; open your link at least once a week until you submit.'
    ],
    groups
  });
  download('peer-eval-links-' + slug(E.name) + '-' + E.id + '.pdf', pdf, 'application/pdf');
  log('Saved the link list');
  toast('Saved the link list. Upload this one file where your OCs can open it: Teams, a shared drive or the group chat.');
}

function printSlips(keys) {
  if (!E.appUrl) {
    toast('QR slips need an OC page address, set before creating links. Send each OC their message instead.');
    return;
  }
  const list = E.members.filter((m) => m.link && (!keys || keys.indexOf(m.squad) >= 0));
  let html = '';
  for (let i = 0; i < list.length; i += 8) html += '<div class="page">' + list.slice(i, i + 8).map(slipHtml).join('') + '</div>';
  document.getElementById('print').innerHTML = html;
  window.print();
}

function messageFor(m) {
  if (!E.appUrl) {
    return 'OC ' + m.name + ', your peer evaluation (personal, do not share). Open PeerEval_OC.html and paste this whole message into it:\n' + m.link;
  }
  return 'OC ' + m.name + ', your peer evaluation link (personal, do not share):\n' + m.link +
    '\nOpen it in Safari or Chrome. Answers save on your phone; open it at least once a week until you submit.';
}

/* ---------- rendering ---------- */

function group(id, summary, body, openDefault) {
  const open = ui.open.has(id) || (openDefault && !ui.open.has('closed:' + id));
  return '<details class="group" data-gid="' + esc(id) + '"' + (open ? ' open' : '') + '><summary>' + summary + '</summary><div class="body">' + body + '</div></details>';
}

function counts() {
  if (!live()) return { done: 0, total: 0 };
  const raters = E.members.filter((m) => squadOf(m).ids.length > 1);
  return { done: raters.filter((m) => E.results[m.id]).length, total: raters.length };
}

function renderTop() {
  const c = counts();
  const pct = c.total ? Math.round((100 * c.done) / c.total) : 0;
  return '<header class="topbar"><div class="in">' +
    '<div class="left"><h1>Peer Eval · Admin</h1><span class="muted">' + (E ? esc(E.name) + (live() ? '' : ' · squad sign-up') : 'No evaluation yet') + '</span></div>' +
    '<div class="right">' +
    (live() ? '<div class="progress" title="Results in"><div class="bar"><i style="width:' + pct + '%"></i></div><b>' + c.done + ' / ' + c.total + ' in</b></div>' : '') +
    (E ? '<button class="btn small" data-act="save-file">' + (live() ? 'Save event file' : 'Save sign-up file') + (ui.unsaved ? ' *' : '') + '</button>' : '') +
    '<button class="btn small" data-act="open-file">Open event file</button>' +
    '</div></div></header>';
}

function renderSettings() {
  const locked = !!E;
  const sheetLine = locked ? '' : '<div class="row"><span>Results sheet: ' + (D.sheetSkipped ? 'off, results by file and text' : D.collectorUrl.trim() ? 'your Google Sheet' : 'not set up') + '</span>' +
    '<button class="btn small" data-act="sheet-setup">' + (D.collectorUrl.trim() && !D.sheetSkipped ? 'Change' : 'Set up') + '</button></div>';
  return '<div class="settings">' +
    '<label class="field"><span>Evaluation name</span><input type="text" data-field="name" value="' + esc(locked ? E.name : D.name) + '"' + (locked ? ' disabled' : '') + '></label>' +
    '<label class="field"><span>OC page address (where the OC app is hosted; leave empty to send codes for PeerEval_OC.html)</span><input type="url" data-field="appUrl" placeholder="https://…/peer/" value="' + esc(locked ? E.appUrl : D.appUrl) + '"' + (locked ? ' disabled' : '') + '></label>' +
    '</div>' + sheetLine +
    '<p class="muted">Attributes per OC: 6 for squads of 10 or more, more for smaller squads, so every OC is rated on all 30.</p>';
}

/** The results sheet, set up once: four steps, two of them in this console. Hidden once ready or skipped. */
function renderSheetCard() {
  const st = ui.sheet;
  return '<section class="card sheetcard"><div class="head"><h2>Results sheet</h2><span class="muted">Google, free, set up once (about 3 minutes)</span></div><ol class="steps">' +
    '<li><a class="btn small" data-act="open-sheet" href="https://sheets.new" target="_blank" rel="noopener noreferrer">Copy the script and open a new sheet</a> In the new sheet: Extensions › Apps Script.</li>' +
    '<li>In Apps Script, select everything, paste (the script is already copied), then Save. <button class="btn small" data-act="copy-script">Copy again</button>' +
    '<span id="script-status" class="sheetstatus' + (ui.scriptCopied === 'no' ? ' bad' : '') + '">' + esc(SCRIPT_TEXT[ui.scriptCopied] || '') + '</span></li>' +
    '<li>Deploy › New deployment › ⚙ › Web app.<br>Who has access: <b>Anyone</b> › Deploy.<br>Authorize access › Allow. If Google says it hasn\'t verified the app: Advanced › Go to … (unsafe) › Allow. It is your own script.</li>' +
    '<li><label class="field"><span>Copy the Web app URL and paste it here.</span><input type="url" data-field="collectorUrl" placeholder="https://script.google.com/macros/s/…/exec" autocomplete="off" spellcheck="false" value="' + esc(D.collectorUrl) + '"></label>' +
    '<span id="sheet-status" class="sheetstatus' + (st && st !== 'checking' && st !== 'ok' && st !== 'reused' ? ' bad' : '') + '">' + esc(SHEET_TEXT[st] || '') + '</span></li></ol>' +
    '<p class="muted">To change the script later: Deploy › Manage deployments › Edit › Version: New version. That keeps the URL; a new deployment would break every link already sent.</p>' +
    '<div class="row"><button class="btn small" data-act="skip-sheet">Skip: results by file and text only</button></div></section>';
}

function renderRoster() {
  if (live()) {
    const plts = new Set(E.squads.map((s) => s.plt)).size;
    let body = '';
    E.squads.forEach((sq) => {
      body += group('roster:' + sq.key, '<span class="sl">' + esc(sq.label) + '</span><span class="muted">' + sq.ids.length + ' OCs · ' + sq.k + ' attributes each</span>',
        '<div class="chips">' + sq.ids.map((id) => '<span class="chip">' + esc(nameOf(id)) + '</span>').join('') + '</div>');
    });
    return '<section class="card"><div class="head"><h2 class="step"><span class="num ok">✓</span>Roster</h2>' +
      '<span class="okline">' + E.members.length + ' OCs · ' + plts + ' PLT · ' + E.squads.length + ' SQD</span></div>' +
      '<div class="groups">' + body + '</div>' +
      group('settings', '<span class="sl">Settings</span><span class="muted">' + (E.collector.url ? 'Results go to your Google Sheet' : 'Results by file and text') + '</span>', renderSettings()) +
      '<div class="row"><button class="btn" data-act="new-eval">New evaluation</button><span class="muted">Same roster and results sheet; you can edit the roster first.</span></div></section>';
  }
  const R = rosterNow();
  const sheetReady = !D.sheetSkipped && !!D.collectorUrl.trim() && (ui.sheet === 'ok' || ui.sheet === 'reused');
  const showCard = !E && (ui.sheetOpen || (!D.sheetSkipped && !sheetReady));
  return (showCard ? renderSheetCard() : '') + (E ? renderSignup(R) : '') + renderRosterCard(R, sheetReady, showCard);
}

/** Squad buttons in a PLT: 4 (B10-1), more when a squad above 4 is already in use or a 5th was added. */
function squadCount(R, plt) {
  const linked = E && E.join ? (E.join.plts.find((p) => p.plt === plt) || {}).sq || 0 : 0;
  return Math.max(4, linked, (D.dec.five || {})[plt] ? 5 : 0, ...R.ocs.filter((o) => o.plt === plt).map((o) => o.sqd || 0));
}

/** PLT buttons for an OC with none: the PLTs in the roster and one more. */
function pltChoices(R) {
  const ps = R.ocs.map((o) => o.plt).filter((p, i, a) => p && a.indexOf(p) === i).sort((a, b) => a - b);
  return ps.length ? ps.concat(ps[ps.length - 1] < 20 ? [ps[ps.length - 1] + 1] : []) : [1, 2, 3];
}

const pick = (act, key, v, on, label, who) => '<button class="pick" data-act="' + act + '" data-key="' + esc(key) + '" data-v="' + v + '" aria-pressed="' + on + '" aria-label="' + esc(who + ': ' + label) + '">' + esc(label) + '</button>';

/** One OC: its name (tap to edit), what is known about it, and the PLT and squad buttons it needs. */
function ocRow(o, R) {
  const name = ui.edit === o.key
    ? '<span class="redit"><input type="text" data-field="editLast" aria-label="Last name" value="' + esc(ui.editLast) + '"><input type="text" data-field="editFirst" aria-label="First name" placeholder="First name" value="' + esc(ui.editFirst) + '">' +
      '<button class="btn small primary" data-act="edit-save" data-key="' + esc(o.key) + '">Save</button><button class="btn small" data-act="edit-cancel">Cancel</button></span>'
    : '<button class="rname" data-act="edit" data-key="' + esc(o.key) + '" aria-label="Edit the name ' + esc(fullName(o)) + '"><b>' + esc(fullName(o)) + '</b>' + ICON_PEN + '</button>';
  let picks = '';
  if (!o.plt) picks += '<span class="picks" role="group" aria-label="PLT for ' + esc(o.last) + '">' + (o.plts || pltChoices(R)).map((p) => pick('plt', o.key, p, false, PLT(p), o.last)).join('') + '</span>';
  const n = o.plt ? squadCount(R, o.plt) : 4;
  picks += '<span class="picks" role="group" aria-label="Squad for ' + esc(o.last) + '">' + Array.from({ length: n }, (x, i) => pick('sq', o.key, i + 1, o.sqd === i + 1, ORD(i + 1), o.last)).join('') + '</span>';
  const done = !!(o.plt && o.sqd);
  const tag = o.tags.join(' · ') || (done ? '' : o.plt ? 'needs a squad' : 'needs a PLT and a squad');
  const bad = o.tags.some((t) => /^(you set|the file says|the files say|picked twice)/.test(t));
  return '<div class="rrow' + (done ? '' : ' need') + '">' + name + '<span class="rtag' + (bad ? ' bad' : '') + '">' + esc(tag) + '</span>' + picks + '</div>';
}

function summaryOf(R) {
  const plts = R.ocs.map((o) => o.plt).filter((p, i, a) => p && a.indexOf(p) === i).length;
  const q = R.questions.filter((x) => x.kind !== 'add').length;
  const parts = [R.ocs.length + ' OCs', plts + ' PLT'];
  if (q) parts.push(q + (q === 1 ? ' question' : ' questions'));
  if (R.need) parts.push(R.need + ' need a squad');
  if (!q && !R.need && R.ocs.length) parts.push(R.squads.length + ' SQD', R.block.length ? 'see below' : 'ready');
  return parts.join(' · ');
}

/** The questions only the admin can answer: Same person?, Which one?, Add or Ignore. */
function questionsBox(R) {
  if (!R.questions.length) return '';
  const row = (html, btns) => '<div class="qrow"><span>' + html + '</span><span class="row">' +
    btns.map(([act, id, v, label, not]) => '<button class="btn small" data-act="' + act + '" data-q="' + esc(id) + '" data-v="' + esc(v) + '"' + (not ? ' data-not="' + esc(not) + '"' : '') + '>' + esc(label) + '</button>').join('') + '</span></div>';
  const where = (q) => (q.plt ? ', ' + PLT(q.plt) : '') + (q.sqd ? ' ' + ORD(q.sqd) + ' SQD' : '');
  return '<div class="questions"><b>Answer these before the links</b>' + R.questions.map((q) => {
    if (q.kind === 'same') {
      return row('Same person? <b>' + esc(q.a.name) + '</b> in ' + esc(q.a.src) + ' and <b>' + esc(q.b.name) + '</b> in ' + esc(q.b.src) + (q.plt ? ', both ' + PLT(q.plt) : '') + '.',
        [['same', q.id, 'same', 'Same person'], ['same', q.id, 'two', 'Two people']]);
    }
    if (q.kind === 'which') {
      return row('Which ' + esc(q.name) + '? ' + esc(q.src) + ' has ' + esc(q.name) + where(q) + '; the roster has ' + q.options.length + ' who could be this OC.',
        q.options.map((o) => ['which', q.id, o.key, o.name + (o.plt && !q.plt ? ' (' + PLT(o.plt) + ')' : '')]).concat([['which', q.id, 'none', 'Someone else', q.options.map((o) => o.key).join(' ')]]));
    }
    const intro = q.src === 'the sign-up' ? 'Not on the roster: <b>' + esc(q.name) + '</b> picked' + esc(where(q).replace(/^, /, ' ')) + '.' : 'Only in ' + esc(q.src) + ': <b>' + esc(q.name) + '</b>' + esc(where(q)) + '. A bigger roster covers that PLT.';
    return row(intro + (q.hint ? ' ' + esc(q.hint) : ''), [['add', q.id, 'add', 'Add'], ['add', q.id, 'ignore', 'Ignore']]);
  }).join('') + '</div>';
}

/** Each file or paste: what it gave, PLT and SQD for its names that have none, and Remove. */
function sourceChips() {
  return D.sources.map((s) => {
    const ch = Object.assign({ plt: s.plt, sqd: s.sqd }, (D.dec.src || {})[s.id] || {});
    const info = [s.tab, s.rows.length + (s.rows.length === 1 ? ' name' : ' names')];
    if (s.plt || s.sqd) info.push([s.plt ? PLT(s.plt) : '', s.sqd ? ORD(s.sqd) + ' SQD' : ''].filter(Boolean).join(' ') + ' from the file name');
    const loose = s.rows.some((r) => (!r.plt && !r.plts) || !r.sqd);
    const sel = (f, v) => '<select data-src="' + esc(s.id) + '" data-f="' + f + '" aria-label="' + f.toUpperCase() + ' for the names in ' + esc(s.name) + '"><option value="">' + f.toUpperCase() + '</option>' +
      Array.from({ length: 20 }, (x, i) => '<option value="' + (i + 1) + '"' + (v === i + 1 ? ' selected' : '') + '>' + ORD(i + 1) + '</option>').join('') + '</select>';
    return '<span class="srcchip"><b>' + esc(s.name) + '</b> · ' + esc(info.filter(Boolean).join(' · ')) +
      (loose ? '<span class="srcset">names without them: ' + sel('plt', ch.plt) + sel('sqd', ch.sqd) + '</span>' : '') +
      '<button class="x" data-act="drop-src" data-src="' + esc(s.id) + '" aria-label="Remove ' + esc(s.name) + '">×</button></span>';
  }).join('');
}

function notUsedLine() {
  const sum = (k) => D.sources.reduce((n, s) => n + (s[k] || 0), 0);
  const cols = D.sources.reduce((a, s) => a.concat(s.columns.filter((c) => a.indexOf(c) < 0)), []);
  const parts = [];
  if (cols.length) parts.push('the ' + cols.join(', ') + (cols.length === 1 ? ' column' : ' columns'));
  if (sum('titles')) parts.push(sum('titles') + (sum('titles') === 1 ? ' title row' : ' title rows'));
  if (sum('empty')) parts.push(sum('empty') + (sum('empty') === 1 ? ' empty place' : ' empty places') + ' (N/A, Unoccupied, …)');
  if (sum('tabs')) parts.push(sum('tabs') + (sum('tabs') === 1 ? ' other tab' : ' other tabs'));
  return '<p class="muted">' + (parts.length ? 'Not used: ' + esc(parts.join(', ')) + '. ' : '') + 'Only last name, first name, PLT and SQD are kept.</p>';
}

function renderRosterCard(R, sheetReady, showCard) {
  const head = '<div class="head"><h2 class="step"><span class="num">1</span>Roster</h2>' +
    (D.sources.length ? '<span class="' + (R.block.length ? 'badline' : 'okline') + '" role="status">' + esc(summaryOf(R)) + '</span>' : '') + '</div>';
  const settings = group('settings', '<span class="sl">Settings</span><span class="muted">' + esc(D.name) + (D.appUrl ? '' : ' · codes for PeerEval_OC.html') + '</span>', renderSettings()) +
    (showCard || E ? '' : '<p class="' + (sheetReady ? 'okline' : 'muted') + '">' + (sheetReady ? '✓ Results go to your Google Sheet.' : 'Results come in by file and text (no sheet).') + '</p>');
  if (!D.sources.length) {
    return '<section class="card">' + head +
      '<div class="drop rosterdrop" data-drop="roster"><b>Drop your roster files here</b>' +
      '<span>Any sheet works: a class roster, a squad\'s Peer Evaluations workbook, a list of names. Add as many as you like.</span>' +
      '<span class="row"><button class="btn primary" data-act="pick-roster">Choose files</button><button class="btn" data-act="paste-roster">Paste a list</button><span>or paste with Ctrl+V (⌘V on a Mac)</span></span>' +
      '<span>Old .xls, password-protected and Numbers files can\'t be read: save them as .xlsx, or copy and paste.</span></div>' +
      '<div class="row"><button class="linkbtn" data-act="blank-sheet">Download a blank roster sheet</button></div>' + settings + createRow(R) + '</section>';
  }
  const byPlt = {};
  R.ocs.forEach((o) => {
    const k = o.plt ? String(o.plt) : o.plts ? 'unclear' : 'none';
    (byPlt[k] = byPlt[k] || []).push(o);
  });
  let body = '';
  Object.keys(byPlt).filter((k) => /^\d+$/.test(k)).sort((a, b) => a - b).forEach((k) => {
    const plt = Number(k);
    const ocs = byPlt[k];
    const need = ocs.filter((o) => !o.sqd).sort(byName);
    const five = (D.dec.five || {})[plt];
    const sq = squadCount(R, plt);
    const fiveBtn = five ? '<button class="linkbtn" data-act="five" data-plt="' + plt + '">Remove the 5th squad</button>'
      : sq === 4 ? '<button class="linkbtn" data-act="five" data-plt="' + plt + '">' + (E ? 'Add a 5th (only you can pick it)' : 'Add a 5th squad') + '</button>' : '';
    body += '<div class="plt"><div class="plthead"><h3>' + PLT(plt) + '</h3><span class="muted">' + ocs.length + ' OCs · ' + (need.length ? need.length + ' need a squad' : 'every OC has a squad') + ' · ' + sq + ' squads</span>' + fiveBtn + '</div>' +
      need.map((o) => ocRow(o, R)).join('') +
      '<div class="groups">' + R.squads.filter((s) => s.plt === plt).map((s) => group('rsq:' + s.key, '<span class="sl">' + esc(s.label) + '</span><span class="muted">' + s.ocs.length + ' OCs</span>', s.ocs.map((o) => ocRow(o, R)).join(''))).join('') + '</div></div>';
  });
  if (byPlt.unclear) body += '<div class="plt"><div class="plthead"><h3>PLT unclear</h3><span class="muted">' + byPlt.unclear.length + ' OCs: tap their PLT and squad</span></div>' + byPlt.unclear.sort(byName).map((o) => ocRow(o, R)).join('') + '</div>';
  if (byPlt.none) body += '<div class="plt"><div class="plthead"><h3>No PLT yet</h3><span class="muted">' + byPlt.none.length + ' OCs: set PLT and SQD for a whole list on its chip above, or tap each</span></div>' + byPlt.none.sort(byName).map((o) => ocRow(o, R)).join('') + '</div>';
  const running = !E && D.signup ? '<div class="warn"><b>A squad sign-up is running for this roster</b> (started ' + esc(D.signup.at) + '). To keep its answers, open its sign-up file (…-signup-date.json) with Open event file. Ask OCs again would make new links, and the ones already sent would stop counting.</div>' : '';
  const helper = E ? '' : running + '<div class="helper"><span>' + (R.need ? 'Tap a squad on each yellow row, or let the OCs pick:' : 'Every OC has a PLT and a squad.') + '</span><span class="row">' +
    (R.need ? '<button class="btn primary" data-act="ask">Ask OCs to pick their squad</button>' : '') + '<button class="linkbtn" data-act="sheet-out">Download as a sheet</button></span></div>';
  return '<section class="card" data-drop="roster">' + head +
    '<div class="row srcs">' + sourceChips() + '<button class="btn small" data-act="pick-roster">Add a file</button><button class="btn small" data-act="paste-roster">Paste a list</button></div>' +
    notUsedLine() + questionsBox(R) + helper + '<div class="plts">' + body + '</div>' + settings + createRow(R) + '</section>';
}

/** What still stops Create links, what it stops once for, what is only listed; then the button. */
function createRow(R) {
  const li = (t) => '<li>' + esc(t) + '</li>';
  const off = !!ui.busy || !R.ocs.length || R.block.length > 0 || (E && ui.signupResults);
  const notes = (R.block.length ? '<div class="warn"><b>Before Create links:</b><ul>' + R.block.map(li).join('') + '</ul></div>' : '') +
    (R.hold.length ? '<div class="warn soft">' + (ui.held ? 'Going ahead with these on the next press:' : 'Check these: Create links stops once to show them.') + '<ul>' + R.hold.map(li).join('') + '</ul></div>' : '') +
    (R.list.length ? '<p class="muted">' + esc(R.list.join(' · ')) + '</p>' : '');
  const note = ui.busy || (E
    ? 'Create links saves the event file (…-event-date.json), which replaces the sign-up file: keep it, and a second copy where only you can open it. Never email or share it.'
    : 'Nothing is kept in this browser. The event file downloads when you create links: it is the only key to the results. Keep a second copy where only you can open it (for example your OneDrive, not the Google account with the sheet). Never email or share it.');
  return notes + '<div class="row"><button class="btn primary" data-act="create"' + (off ? ' disabled' : '') + '>Create links</button><span class="muted">' + esc(note) + '</span></div>';
}

/** While a sign-up runs: the PLT links, how many have picked, and the admin's own work saved in the sheet. */
function renderSignup(R) {
  const answered = Object.keys(ui.ans).length + ui.extra.length;
  const inPlt = (o, plt) => o.plt === plt || (!o.plt && (o.plts || []).indexOf(plt) >= 0);
  const links = E.join.plts.map((p) => {
    const left = R.ocs.filter((o) => inPlt(o, p.plt) && o.id !== undefined && !ui.ans[o.id]).length;
    return '<div class="linkrow"><span class="t"><b>' + PLT(p.plt) + ' link</b><span class="u">' + (left ? left + ' still to pick' : 'everyone has picked') + ' · ' + p.sq + ' squads</span></span>' +
      '<span class="row"><button class="btn small" data-act="copy-join" data-plt="' + p.plt + '">Copy link</button><button class="btn small" data-act="qr-join" data-plt="' + p.plt + '">Show QR</button></span></div>';
  }).join('');
  const warn = ui.signupResults ? '<div class="warn"><b>Links were already created for this evaluation.</b> This is its sign-up file. Open the event file from Create links (…-event-date.json) instead; Create links stays off here.</div>' : '';
  return '<section class="card"><div class="head"><h2>Squad sign-up</h2><span class="okline" role="status">' + answered + ' of ' + E.join.count + ' answered' + (ui.lastPull ? ' · updated ' + esc(ui.lastPull) : '') + ' · checks every 30 s</span></div>' + warn +
    '<p class="muted">Each PLT has its own link with its squad buttons, fixed from now on. OCs tap their name, then their squad; answers arrive here by themselves. A squad you tap yourself always wins.</p>' +
    '<p class="okline" id="work-status">' + esc(workText()) + '</p><div class="links">' + links + '</div>' +
    '<div class="row"><button class="btn small" data-act="remind">Copy a reminder</button><span class="muted">The sign-up file (…-signup-date.json) downloaded at Ask OCs. It holds the sign-up only; the event file from Create links replaces it.' +
    (R.ocs.some((o) => o.plts) ? ' OCs with an unclear PLT are in each of their PLT links.' : '') + '</span></div></section>';
}

function renderLinks() {
  if (!live()) return '<section class="card off"><div class="head"><h2 class="step"><span class="num">2</span>Links</h2></div><p class="muted">Appears after you create links.</p></section>';
  const off = ui.refused && E.collector.url ? ' disabled' : ''; // links whose sheet refused the key are never handed out
  let body = '';
  E.squads.forEach((sq) => {
    const rows = sq.ids.map((id) => {
      const m = memberOf(id);
      if (!m.link) return '<div class="linkrow"><span class="t"><b>' + esc(m.name) + '</b><span class="u">Nobody to rate in this squad</span></span></div>';
      return '<div class="linkrow"><span class="t"><b>' + esc(m.name) + '</b><span class="u">' + esc(m.link.slice(0, 60)) + '…</span></span>' +
        '<button class="btn small" data-act="copy-link" data-id="' + id + '" aria-label="Copy the message for ' + esc(m.name) + '"' + off + '>Copy</button></div>';
    }).join('');
    body += group('links:' + sq.key, '<span class="sl">' + esc(sq.label) + '</span><span class="muted">' + sq.ids.length + ' links</span>',
      '<div class="row"><button class="btn small" data-act="print-squad" data-key="' + esc(sq.key) + '"' + off + '>Print QR slips</button>' +
      '<button class="btn small" data-act="copy-squad" data-key="' + esc(sq.key) + '"' + off + '>Copy all messages for this squad</button></div><div class="links">' + rows + '</div>');
  });
  const warn = !E.collector.url ? ''
    : ui.refused ? '<div class="warn">Don\'t hand out these links: the sheet belongs to another evaluation and won\'t release their results. <button class="btn small" data-act="new-sheet">Make the links again with a new sheet</button></div>'
      : !ui.connected && !counts().done ? '<div class="warn">Don\'t hand out links yet: the sheet hasn\'t passed its check for this evaluation. It tries again by itself. <button class="btn small" data-act="connect">Try again</button></div>' : '';
  return '<section class="card"><div class="head"><h2 class="step"><span class="num ok">✓</span>Links</h2>' +
    '<div class="row"><button class="btn primary" data-act="link-list"' + off + '>Link list: one PDF for everyone</button><button class="btn" data-act="print-all"' + off + '>Print all QR slips (8 per page)</button></div></div>' +
    '<p class="muted">Upload the link list once where every OC can open it, hand out the slips, or send each OC their own message. A link counts as the evaluation of the OC it belongs to.</p>' + warn +
    '<div class="groups">' + body + '</div></section>';
}

function renderCollect() {
  if (!live()) return '<section class="card off"><div class="head"><h2 class="step"><span class="num">3</span>Collect</h2></div><p class="muted">Appears after you create links.</p></section>';
  const c = counts();
  let body = '';
  E.squads.forEach((sq) => {
    const raters = sq.ids.length > 1 ? sq.ids : [];
    const done = raters.filter((id) => E.results[id]).length;
    const chips = raters.map((id) => (E.results[id]
      ? '<span class="chip in' + ((E.results[id].devs || []).length > 1 ? ' flag' : '') + '">✓ ' + esc(nameOf(id)) + (E.results[id].rev > 1 ? ' (v' + E.results[id].rev + ')' : '') + ((E.results[id].devs || []).length > 1 ? ' · ' + E.results[id].devs.length + ' devices' : '') + '</span>'
      : '<span class="chip missing">' + esc(nameOf(id)) + '</span>')).join('');
    body += group('collect:' + sq.key, '<span class="sl">' + esc(sq.label) + '</span><span class="' + (done === raters.length ? 'okline' : 'muted') + '">' + done + ' / ' + raters.length + ' in</span>', '<div class="chips">' + chips + '</div>');
  });
  const sheet = E.collector.url
    ? '<div class="row"><span class="grow">Results arrive by themselves from your Google Sheet' + (ui.lastPull ? ' · last pulled ' + esc(ui.lastPull) : '') + '.</span>' +
      '<button class="btn primary" data-act="pull">Pull results</button></div>'
    : '<p class="muted">No Google Sheet collector: results come in as files or text from OCs.</p>';
  const logHtml = (E.log || []).slice(0, 30).map((l) => '<div>' + esc(new Date(l.at).toLocaleString()) + '  ' + esc(l.text) + '</div>').join('');
  const twice = E.members.filter((m) => E.results[m.id] && (E.results[m.id].devs || []).length > 1).map((m) => m.name);
  return '<section class="card"><div class="head"><h2 class="step"><span class="num' + (c.done === c.total && c.total ? ' ok' : '') + '">' + (c.done === c.total && c.total ? '✓' : '3') + '</span>Collect</h2>' +
    '<b>' + c.done + ' / ' + c.total + ' in</b></div>' + sheet +
    (c.done ? '<div class="row"><button class="btn small" data-act="save-answers">Save answers so far (xlsx)</button><span class="muted">One tab per OC, as sent, not scored; it opens in Excel on its own. Keep it where only you can open it.</span></div>' : '') +
    (twice.length ? '<div class="warn">Sent from more than one device: ' + esc(twice.join(', ')) + '. Ask each of them whether the newest result is their own; if not, have them submit again.</div>' : '') +
    '<div class="drop" data-drop="1">Drop result files here, or <button class="btn small" data-act="pick-files">choose files</button>, ' +
    '<button class="btn small" data-act="paste">paste text</button> or <button class="btn small" data-act="scan">scan a phone</button></div>' +
    '<div class="groups">' + body + '</div>' +
    (logHtml ? group('log', '<span class="sl">Import log</span>', '<div class="log">' + logHtml + '</div>') : '') + '</section>';
}

function renderResults() {
  if (!live()) return '<section class="card off"><div class="head"><h2 class="step"><span class="num">4</span>Results</h2></div><p class="muted">Appears after you create links.</p></section>';
  let body = '';
  E.squads.forEach((sq) => {
    const results = squadResults(sq);
    const { rows, raterChecks } = scoreSquad(sq.ids, results);
    const table = '<div class="tablewrap"><table><thead><tr><th>Rank</th><th>Candidate</th><th>Avg attribute</th><th>Total attribute</th><th>Peer score</th><th>Overall</th><th>Raters</th><th>Look closer</th></tr></thead><tbody>' +
      rows.map((r) => '<tr><td class="n">' + (r.merit || '—') + '</td><td><b>' + esc(nameOf(r.id)) + '</b></td><td class="n">' + fmt(r.A, 2) + '</td><td class="n">' + fmt(r.T, 1) + '</td><td class="n">' + fmt(r.Q, 1) + '</td><td class="n"><b>' + fmt(r.overall, 1) + '</b></td><td class="n">' + r.raters + ' / ' + (sq.ids.length - 1) + '</td><td>' + esc(r.flags.join(' · ') || '—') + '</td></tr>').join('') +
      '</tbody></table></div>';
    const checks = raterChecks.filter((x) => x.flags.length).map((x) => '<li><b>' + esc(nameOf(x.r)) + '</b>: ' + esc(x.flags.join(' · ')) + '</li>').join('');
    body += group('results:' + sq.key, '<span class="sl">' + esc(sq.label) + '</span><span class="muted">' + results.length + ' of ' + (sq.ids.length > 1 ? sq.ids.length : 0) + ' raters in</span>',
      table + (checks ? '<div class="warn">How raters filled it in:<ul>' + checks + '</ul></div>' : ''));
  });
  return '<section class="card"><div class="head"><h2 class="step"><span class="num">4</span>Results</h2>' +
    '<button class="btn primary" data-act="export">Export xlsx</button></div>' +
    '<p class="muted">Same formula as the sheet: total attribute score (30 × the average) plus #Yes minus rank, averaged across raters. Flags ask for a closer look; they never change a score.</p>' +
    '<div class="groups">' + body + '</div></section>';
}

function render() {
  app.innerHTML = renderTop() + '<main>' + renderRoster() + renderLinks() + renderCollect() + renderResults() + '</main>' +
    '<input type="file" id="file-event" accept=".json,application/json" hidden><input type="file" id="file-results" multiple hidden>' +
    '<input type="file" id="file-roster" accept=".xlsx,.csv,.tsv,.txt,.xls,.numbers,.ods" multiple hidden>';
  document.getElementById('file-event').addEventListener('change', (e) => {
    if (e.target.files[0]) openEventFile(e.target.files[0]);
  });
  document.getElementById('file-results').addEventListener('change', (e) => {
    if (e.target.files.length) importFiles(Array.from(e.target.files));
  });
  document.getElementById('file-roster').addEventListener('change', (e) => {
    if (e.target.files.length) addFiles(Array.from(e.target.files));
  });
}

/* ---------- roster: files, pastes, Ask OCs ---------- */

/**
 * Add what was read from a file or paste. Rows from this console's own Download as a sheet
 * go back to their OCs by ID (a spelling fixed there asks nothing); any other rows join the
 * roster as a new source, merged by name.
 */
function addSource(r, name, units) {
  if (!r.rows.length && !r.unread.length) {
    toast(name + ': no names found.' + (r.columns && r.columns.length ? ' Its header row has no Last or Name column.' : ''));
    return;
  }
  const byKey = {};
  rosterNow().ocs.forEach((o) => o.keys.forEach((k) => (byKey[k] = o)));
  const back = r.rows.filter((x) => x.ref && byKey[x.ref]);
  let changes = 0;
  const set = (field, key, v) => {
    D.dec[field] = Object.assign({}, D.dec[field] || {}, { [key]: v });
    changes++;
  };
  back.forEach((x) => {
    const o = byKey[x.ref];
    const typed = (x.first || '') !== (o.first || ''); // the first name only where the sheet changed it, with the one shown before the first fix
    if (letters(x.last) !== letters(o.last) || typed) set('name', o.key, Object.assign(typed ? { saw: o.first || '' } : {}, (D.dec.name || {})[o.key], { last: x.last }, typed ? { first: x.first || '' } : {}));
    if (x.plt && x.plt !== o.plt) set('plt', o.key, x.plt);
    if (x.sqd && x.sqd !== o.sqd) set('sq', o.key, x.sqd);
  });
  const rest = Object.assign({}, r, { rows: r.rows.filter((x) => !(x.ref && byKey[x.ref])) });
  const taken = D.sources.filter((s) => s.name === name || s.name.indexOf(name + ' (') === 0).length;
  if (rest.rows.length || (!back.length && rest.unread.length)) D.sources.push(sourceOf(rest, taken ? name + ' (' + (taken + 1) + ')' : name, units));
  changed();
  toast(back.length ? name + ': ' + back.length + ' OCs matched by ID, ' + changes + (changes === 1 ? ' change' : ' changes') + '.' + (rest.rows.length ? ' ' + rest.rows.length + ' other rows added.' : '')
    : 'Read ' + name + ': ' + r.rows.length + (r.rows.length === 1 ? ' name.' : ' names.'));
}

const addText = (text, name) => addSource(readRoster(text), name, null);

/** Roster files, read in this browser only: .xlsx (the first tab with names), .csv, .tsv, .txt. */
async function addFiles(files) {
  for (const f of files) {
    const name = f.name || 'file';
    const ext = ((/\.([A-Za-z0-9]+)$/.exec(name) || [])[1] || '').toLowerCase();
    try {
      if (ext === 'xlsx' || ext === 'xlsm') {
        const tabs = await readXlsx(new Uint8Array(await f.arrayBuffer()));
        let found = null;
        tabs.some((t, i) => {
          const r = readRoster(t.cells.map((row) => row.join('\t')).join('\n'));
          if (r.rows.length || r.unread.length) found = Object.assign(r, { tabs: tabs.length - 1, tab: tabs.length > 1 ? 'tab ' + (i + 1) + ' of ' + tabs.length : '' });
          return !!found;
        });
        if (found) addSource(found, name, fileUnits(name));
        else toast(name + ': no names found in any tab.');
      } else if (ext === 'csv' || ext === 'tsv' || ext === 'txt') {
        const text = await f.text();
        addSource(readRoster(ext === 'csv' ? csvText(text) : text), name, fileUnits(name));
      } else {
        toast(name + ' can\'t be read here. Save it as .xlsx (in Excel: File › Save As), or copy the names and paste them.');
      }
    } catch (err) {
      toast(name + (err && err.message === 'locked' ? ' is an old .xls or has a password: save it as .xlsx without one, or copy and paste the names.' : ' can\'t be read: save it as .xlsx, or copy and paste the names.'));
    }
  }
}

function openRosterPaste() {
  const dlg = document.createElement('dialog');
  dlg.innerHTML = '<p><b>Paste a list</b></p><p class="muted">From any sheet, header row and all, or names alone. Several lists can go in one after another.</p>' +
    '<textarea style="min-height:240px" aria-label="Roster to add"></textarea><div class="row" style="margin-top:10px"><button class="btn primary" data-go="1">Add</button><button class="btn">Cancel</button></div>';
  document.body.appendChild(dlg);
  const box = dlg.querySelector('textarea');
  dlg.querySelectorAll('button').forEach((btn) => btn.addEventListener('click', () => {
    const go = btn.dataset.go;
    const text = box.value;
    dlg.close();
    dlg.remove();
    if (go && text.trim()) addText(text, 'Pasted list');
  }));
  dlg.showModal();
  box.focus();
}

/** The roster as a sheet to hand to someone: every OC with its ID, blanks in yellow; dropped back, it goes in by ID. */
function downloadSheet() {
  const R = rosterNow();
  const rows = [['Peer evaluation roster · ' + D.name], ['Fill in the yellow cells. Keep the ID column, then drop this file back into the console.'], ['ID', 'Last', 'First', 'PLT', 'SQD']];
  const fill = new Set();
  R.ocs.slice().sort((a, b) => (a.plt || 99) - (b.plt || 99) || (a.sqd || 99) - (b.sqd || 99) || byName(a, b)).forEach((o) => {
    if (!o.plt) fill.add(rows.length + ',3');
    if (!o.sqd) fill.add(rows.length + ',4');
    rows.push([o.key, o.last, o.first, o.plt || '', o.sqd || '']);
  });
  download('peer-eval-roster-' + slug(D.name) + '.xlsx', buildXlsx([{ name: 'Roster', rows, widths: [10, 22, 18, 8, 8], header: false, bold: [2], fill }]), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  toast('Saved the roster sheet. It holds names: keep it where only you and whoever fills it can open it.');
}

function downloadBlank() {
  download('peer-eval-roster-blank.xlsx', buildXlsx([{ name: 'Roster', rows: [['Last', 'First', 'PLT', 'SQD']], widths: [22, 18, 8, 8] }]), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

/**
 * Ask OCs: the evaluation starts now (id, keys, every OC's id and token, the deal seed), so
 * the OCs' picks can be sealed for it. One link per PLT lists that PLT's OCs; the sign-up
 * file downloads, and the sheet gets this evaluation's read key.
 */
async function askOCs() {
  if (D.signup && !window.confirm('A squad sign-up is already running for this roster (started ' + D.signup.at + '). Start a new one? The sign-up links already sent would stop counting. To keep it, press Cancel and open its sign-up file.')) return;
  const url = D.collectorUrl.trim();
  if (D.sheetSkipped || !url || !(ui.sheet === 'ok' || ui.sheet === 'reused')) {
    ui.sheetOpen = true;
    D.sheetSkipped = false;
    render();
    toast('Ask OCs needs the Google results sheet, so the picks can reach you. Set it up first, or tap each squad yourself.');
    return;
  }
  const appUrl = String(D.appUrl || '').trim();
  if (!/^https?:\/\//.test(appUrl)) {
    toast('Ask OCs needs the OC page address (Settings), so the sign-up links open on phones.');
    return;
  }
  const R = rosterNow();
  if (!R.ocs.length) return;
  ui.busy = 'Making the sign-up links…';
  render();
  const order = R.ocs.slice().sort((a, b) => (a.plt || 99) - (b.plt || 99) || byName(a, b));
  const ids = {};
  const tokens = {};
  const listed = {}; // the first name each link lists: an OC split later keeps its id on the half with it
  order.forEach((o, i) => {
    o.keys.forEach((k) => (ids[k] = i));
    tokens[i] = randomToken(16);
    listed[i] = o.first || '';
  });
  const keys = await generateAdminKeys();
  E = {
    v: 1, stage: 'signup', id: randomToken(6), name: D.name || 'Peer evaluation', created: new Date().toISOString(), appUrl,
    collector: { url, readKey: D.readKey || randomToken(18) }, keys, seed: new DataView(randomBytes(4).buffer).getUint32(0) % 1000000,
    ids, listed, tokens, join: { work: randomToken(16), wrev: 0, count: order.length, plts: [] }, answers: {}, members: [], squads: [], results: {}, log: []
  };
  const plts = order.reduce((a, o) => a.concat((o.plt ? [o.plt] : o.plts || []).filter((p) => a.indexOf(p) < 0)), []).sort((a, b) => a - b);
  for (const plt of plts) {
    const token = randomToken(16);
    const ns = order.filter((o) => o.plt === plt || (!o.plt && (o.plts || []).indexOf(plt) >= 0)).map((o) => (o.first ? [ids[o.key], o.last, o.first] : [ids[o.key], o.last]));
    const sq = squadCount(R, plt);
    const link = appUrl.replace(/#.*$/, '') + '#' + (await packPacket({ v: 1, j: 1, e: E.id, n: E.name, g: PLT(plt), pl: plt, sq, ns, k: keys.pub, t: token, u: url }));
    E.join.plts.push({ plt, sq, token, link });
  }
  Object.assign(ui, { busy: '', ans: {}, extra: [], signupResults: false, held: false, connected: false, refused: false, work: '', workEnv: null });
  D.signup = { id: E.id, at: new Date().toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }) };
  D.readKey = E.collector.readKey; // this tab keeps the sheet's key, as New evaluation does
  log('Asked the OCs to pick their squad: ' + plts.length + ' sign-up links');
  saveEventFile(); // the sign-up file
  const none = R.ocs.filter((o) => !o.plt && !o.plts).length;
  await connectCollector('Sign-up links ready' + (none ? ' (' + none + ' OCs with no PLT are in no link: tap theirs yourself)' : '') + '. ');
  pull(true);
}

function showJoinQr(plt) {
  const p = E.join.plts.find((x) => x.plt === plt);
  if (!p || typeof qrcode !== 'function') return;
  const qr = qrcode(0, 'L');
  qr.addData(p.link, 'Byte');
  qr.make();
  const dlg = document.createElement('dialog');
  dlg.className = 'qrfull';
  dlg.innerHTML = '<p><b>' + esc(PLT(plt)) + ' · squad sign-up</b></p><div class="qrbig" role="img" aria-label="Sign-up code for ' + esc(PLT(plt)) + '">' + qr.createSvgTag({ cellSize: 4, margin: 2, scalable: true }) + '</div>' +
    '<p class="muted">Scan with the phone camera from 1 to 2 m, then open it in Safari or Chrome. It is a dense code: show it full screen, and never print it small.</p><div class="row"><button class="btn primary">Close</button></div>';
  document.body.appendChild(dlg);
  const close = () => {
    dlg.close();
    dlg.remove();
  };
  dlg.querySelector('button').addEventListener('click', close);
  dlg.addEventListener('cancel', close);
  dlg.showModal();
}

/** For the PLT chats: who has not picked yet, with each PLT's link. */
function reminderText() {
  const R = rosterNow();
  return E.join.plts.map((p) => {
    const left = R.ocs.filter((o) => (o.plt === p.plt || (!o.plt && (o.plts || []).indexOf(p.plt) >= 0)) && o.id !== undefined && !ui.ans[o.id]).map(fullName);
    return PLT(p.plt) + ' squad sign-up' + (left.length ? '. Not picked yet: ' + left.join('; ') : ': everyone has picked') + '\n' + p.link;
  }).join('\n\n') + '\n\nOpen your PLT\'s link in Safari or Chrome, tap your name, then your squad.';
}

async function onCreate() {
  const R = rosterNow();
  const useSheet = !D.sheetSkipped && !!D.collectorUrl.trim();
  if (!E && useSheet && (!ui.sheet || ui.sheet === 'checking')) await checkSheet();
  if (!R.ocs.length || R.block.length) {
    render();
    toast(R.ocs.length ? 'Fix what is listed above Create links first.' : 'Add the roster first: drop a file or paste a list.');
  } else if (E && ui.signupResults) {
    toast('Links were already created for this evaluation: open its event file instead.');
  } else if (!E && useSheet && SHEET_BLOCKS.indexOf(ui.sheet) >= 0) {
    ui.sheetOpen = true;
    render();
    toast('Fix the results sheet first (its last step says how), or skip it.');
  } else if (R.hold.length && !ui.held) {
    ui.held = true;
    render();
    toast('Check what is listed above Create links. Press Create links again to go ahead.');
  } else {
    await createEvent(R);
  }
}

/* ---------- events ---------- */

app.addEventListener('toggle', (e) => {
  const d = e.target;
  if (!d.matches || !d.matches('details[data-gid]')) return;
  const id = d.dataset.gid;
  if (d.open) {
    ui.open.add(id);
    ui.open.delete('closed:' + id);
  } else {
    ui.open.delete(id);
    ui.open.add('closed:' + id);
  }
}, true);

app.addEventListener('input', (e) => {
  const f = e.target.dataset && e.target.dataset.field;
  if (!f) return;
  if (f === 'editLast' || f === 'editFirst') {
    ui[f] = e.target.value;
    return;
  }
  D[f] = e.target.value;
  persist();
  if (f === 'collectorUrl') {
    clearTimeout(sheetTimer);
    sheetTimer = setTimeout(checkSheet, 400);
  }
});

/* PLT and SQD for the names of one file or paste that have none (its chip) */
app.addEventListener('change', (e) => {
  const src = e.target.dataset && e.target.dataset.src;
  if (!src || !e.target.dataset.f) return;
  const cur = Object.assign({}, (D.dec.src || {})[src] || {}, { [e.target.dataset.f]: Number(e.target.value) || null });
  D.dec.src = Object.assign({}, D.dec.src || {}, { [src]: cur });
  changed();
});

app.addEventListener('click', async (e) => {
  const b = e.target.closest('[data-act]');
  if (!b) return;
  const act = b.dataset.act;
  const dec = (field, k, v) => {
    D.dec[field] = Object.assign({}, D.dec[field] || {}, { [k]: v });
    changed();
  };
  if (act === 'create') {
    await onCreate();
  } else if (act === 'pick-roster') {
    document.getElementById('file-roster').click();
  } else if (act === 'paste-roster') {
    openRosterPaste();
  } else if (act === 'blank-sheet') {
    downloadBlank();
  } else if (act === 'sheet-out') {
    downloadSheet();
  } else if (act === 'drop-src') {
    D.sources = D.sources.filter((s) => s.id !== b.dataset.src);
    changed();
  } else if (act === 'edit') {
    const o = rosterNow().ocs.find((x) => x.key === b.dataset.key);
    if (!o) return;
    Object.assign(ui, { edit: o.key, editLast: o.last, editFirst: o.first, editWas: { last: o.last, first: o.first || '' } });
    render();
    const box = app.querySelector('[data-field=editLast]');
    if (box) box.focus();
  } else if (act === 'edit-save') {
    const last = ui.editLast.trim();
    const first = String(ui.editFirst || '').trim();
    const was = ui.editWas || {};
    const typed = first !== was.first; // the first name only where it was typed (a fix of the last name alone keeps the OC's own), with the one shown before the first fix
    ui.edit = '';
    if (last && /\p{L}/u.test(last) && (last.toUpperCase() !== was.last || typed)) dec('name', b.dataset.key, Object.assign(typed ? { saw: was.first } : {}, (D.dec.name || {})[b.dataset.key], { last: last.toUpperCase() }, typed ? { first } : {}));
    else render(); // nothing changed: nothing fixed
  } else if (act === 'edit-cancel') {
    ui.edit = '';
    render();
  } else if (act === 'sq') {
    dec('sq', b.dataset.key, Number(b.dataset.v));
  } else if (act === 'plt') {
    dec('plt', b.dataset.key, Number(b.dataset.v));
  } else if (act === 'same' || act === 'which' || act === 'add') {
    if (b.dataset.not) D.dec.not = Object.assign({}, D.dec.not || {}, { [b.dataset.q]: [...new Set(((D.dec.not || {})[b.dataset.q] || []).concat(b.dataset.not.split(' ')))] }); // Someone else: none of the OCs offered (nor the ones before)
    dec(act, b.dataset.q, b.dataset.v);
  } else if (act === 'five') {
    dec('five', b.dataset.plt, !(D.dec.five || {})[b.dataset.plt]);
  } else if (act === 'ask') {
    await askOCs();
  } else if (act === 'copy-join') {
    const p = E.join.plts.find((x) => x.plt === Number(b.dataset.plt));
    await copy(p.link, 'the ' + PLT(p.plt) + ' sign-up link');
  } else if (act === 'qr-join') {
    showJoinQr(Number(b.dataset.plt));
  } else if (act === 'remind') {
    await copy(reminderText(), 'a reminder for the PLT chats');
  } else if (act === 'save-file') {
    saveEventFile();
  } else if (act === 'open-file') {
    document.getElementById('file-event').click();
  } else if (act === 'new-eval') {
    newEvaluation(false);
  } else if (act === 'new-sheet') {
    newEvaluation(true);
  } else if (act === 'open-sheet') {
    await copyScript(); // the link itself opens the new sheet, after the clipboard write has started
  } else if (act === 'copy-script') {
    await copyScript();
  } else if (act === 'skip-sheet') {
    D.sheetSkipped = true;
    ui.sheetOpen = false;
    persist();
    render();
  } else if (act === 'sheet-setup') {
    D.sheetSkipped = false;
    ui.sheetOpen = true;
    persist();
    render();
    if (D.collectorUrl.trim()) checkSheet();
  } else if (act === 'connect') {
    await connectCollector();
  } else if (act === 'pull') {
    await pull(false);
  } else if (act === 'pick-files') {
    document.getElementById('file-results').click();
  } else if (act === 'paste') {
    openPaste();
  } else if (act === 'scan') {
    await openScanner();
  } else if (act === 'export') {
    exportXlsx();
  } else if (act === 'save-answers') {
    saveAnswers();
  } else if (act === 'link-list') {
    linkList();
  } else if (act === 'print-all') {
    printSlips(null);
  } else if (act === 'print-squad') {
    printSlips([b.dataset.key]);
  } else if (act === 'copy-link') {
    const m = memberOf(Number(b.dataset.id));
    await copy(messageFor(m), 'the message for ' + m.name);
  } else if (act === 'copy-squad') {
    const sq = E.squads.find((s) => s.key === b.dataset.key);
    await copy(sq.ids.map(memberOf).filter((m) => m.link).map(messageFor).join('\n\n'), 'all messages for ' + sq.label);
  }
});

function openPaste() {
  const dlg = document.createElement('dialog');
  dlg.innerHTML = '<p><b>Paste results</b></p><p class="muted">Paste one or more results from emails or files. Extra text around them is fine.</p>' +
    '<textarea style="min-height:240px"></textarea><div class="row" style="margin-top:10px"><button class="btn primary" data-go="1">Import</button><button class="btn">Cancel</button></div>';
  document.body.appendChild(dlg);
  const box = dlg.querySelector('textarea');
  dlg.querySelectorAll('button').forEach((btn) => btn.addEventListener('click', async () => {
    const go = btn.dataset.go;
    const text = box.value;
    dlg.close();
    dlg.remove();
    if (go) await importText(text, 'text');
  }));
  dlg.showModal();
  box.focus();
}

app.addEventListener('dragover', (e) => {
  const zone = e.target.closest('[data-drop]');
  if (!zone) return;
  e.preventDefault();
  zone.classList.add('over');
});
app.addEventListener('dragleave', (e) => {
  const zone = e.target.closest('[data-drop]');
  if (zone) zone.classList.remove('over');
});
app.addEventListener('drop', (e) => {
  const zone = e.target.closest('[data-drop]');
  if (!zone) return;
  e.preventDefault();
  zone.classList.remove('over');
  if (!e.dataTransfer.files.length) return;
  if (zone.dataset.drop === 'roster') addFiles(Array.from(e.dataTransfer.files));
  else importFiles(Array.from(e.dataTransfer.files));
});

/* A list pasted anywhere on the page, outside a box, joins the roster (before links exist) */
document.addEventListener('paste', (e) => {
  if (live() || (e.target.closest && e.target.closest('input, textarea, select, dialog, [contenteditable]'))) return;
  const text = e.clipboardData && e.clipboardData.getData('text/plain');
  if (!text || !text.trim() || /^\s*https?:\/\/\S+\s*$/.test(text)) return; // a web address alone (the sheet's, a link) is not a roster
  e.preventDefault();
  addText(text, 'Pasted list');
});

/* Leaving during a sign-up: the admin's last changes go to the sheet once more */
window.addEventListener('pagehide', () => {
  if (!E || E.stage !== 'signup' || !E.collector.url || !ui.workEnv || ui.work === 'saved') return;
  try {
    fetch(E.collector.url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(ui.workEnv), keepalive: true }).catch(() => {});
  } catch (err) {
    // nothing more to do on the way out
  }
});

window.addEventListener('beforeunload', (e) => {
  if (E && ui.unsaved && (live() || ui.work !== 'saved')) {
    e.preventDefault();
    e.returnValue = '';
  }
});

restore();
render();
loadScript().catch(() => {}); // ready before "Copy the script and open a new sheet" is pressed
if (D.collectorUrl.trim() && !D.sheetSkipped) checkSheet();
setInterval(() => {
  if (live() && E.collector.url && E.collector.readKey && document.visibilityState === 'visible' && !ui.busy) pull(true);
}, AUTO_PULL_MS);
setInterval(() => {
  if (E && E.stage === 'signup' && E.collector.readKey && document.visibilityState === 'visible' && !ui.busy) pull(true);
}, SIGNUP_PULL_MS);
