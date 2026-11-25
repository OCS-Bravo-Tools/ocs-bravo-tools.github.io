/**
 * Cadre page (cadre mode, the design on the canvas, page 05, v0.3).
 *
 * A cadre drops a roster, gives one QR code to every OC and reads the results. No Google
 * steps and no files to keep: the evaluation's id and keys are made here, the shared sheet
 * (Collector v2) holds only sealed data, this browser keeps the keys when it is the cadre's
 * own ("Is this your own phone?"), and page 2 of the printout (the cadre link) brings the
 * evaluation back on any phone or computer.
 *
 * Screens: home, open, roster, read, set, own, saving, qr, lobby (the sign-up, evaluations made before 17c),
 * progress, results, again, edit (change who's in it), removed, all (Home), more, name,
 * import. The main path has no technical words; they live in More only (T19).
 * Step 1: Results is a plain ranked list with Download Excel. The page per OC is step 2.
 */
import { readRoster, csvText, fileUnits, mergeRoster, typedName, nearMiss, firstCase, firstOk, sameOc } from './lib/roster.js';
import { readXlsx, xlsxInZip } from './lib/xlsx.js';
import { openResult, openMarked, findEnvelopes, sealResult } from './lib/envelope.js';
import { scoreSquad } from './lib/scoring.js';
import { resultsXlsx, answersXlsx, blankBooks, rosterXlsx, rowNotes, formChecks, feedbackBooks, lowScoreGap, readForms, matchName, formAnswers, feedbackPages, QUESTION_LABEL, XLSX_TYPE } from './lib/report.js';
import { ATTRIBUTES, CATEGORIES } from './lib/attributes.js';
import { newKeys, privOf, ocLink, cadreLink, parseCadre, writeToken, writeHash, sealBoard, openBoard, mergeBoard, eventOf, unusedUntil, keptUntil, dayLabel, fullName, ord, squadsOf, handsOf, packetOf } from './lib/board.js';
import { call, newEval, putBoard, startEval, getState, getBoard, getRows, forgetEval } from './lib/collector2.js';
import { collectorUrl, testQuery } from './lib/config.js';

const STORE = 'pe:cadre:v1';
const OWN = 'pe:cadre:own';
const POLL_MS = 30000;
const ERR_MS = 120000;
const BOARD_MS = 30000;
const ROW_KEEP = 12;
const LINK = 'pe:cadre:link'; // this tab's cadre link (sessionStorage) once it leaves the address bar
const AFTER = 'pe:cadre:after'; // the screen (and toast) a reload for a new version opens on (sessionStorage)
const PE = /peer[\s-]*eval/i;

const root = document.getElementById('cadre');
const printEl = document.getElementById('print');
let own = false;
let list = [];
let R = null; // the evaluation on screen, as kept on this device
let X = null; // its working state, in memory only
const ui = { screen: 'home', from: '', sheet: null, tab: 0, resTab: 0, draft: null, again: null, saveProblem: '', paste: false, kept: true, nm: null };
const timers = { poll: 0, toast: 0, retry: 0, after: 0 };
const starting = new Set(); // evaluations (by id) whose frozen board is on its way: no squad moves, no second Start
const waits = new Set(); // working states whose Start waits: for the squads read, for its freeze, or to try again (a reload for a new version waits for them too)
const importing = new Set(); // added results (by mac) on their way to the sheet
let drawn = false; // the first screen is up: a reload for a new version waits for it (a page still opening its evaluation is not on Home)

const esc = (s) => String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const device = () => (matchMedia('(pointer: coarse)').matches || /iPhone|iPad|Android/i.test(navigator.userAgent) ? 'phone' : 'computer');
const site = () => new URL('./', location.href).href.replace(/[?#].*$/, '');
const ocUrl = (r) => ocLink(site() + testQuery(), r.e, r.k);
const cadreUrl = (r) => cadreLink(site() + 'cadre.html' + testQuery(), r);
const listOf = (a, max) => (max && a.length > max ? a.slice(0, max).join(', ') + ' and ' + (a.length - max) + ' more' : a.length > 1 ? a.slice(0, -1).join(', ') + ' and ' + a[a.length - 1] : a.join(''));
const plural = (n, one, many) => n + ' ' + (n === 1 ? one : many);
const dateOf = (r) => dayLabel(new Date(r.created));
const titleOf = (r) => r.name + ' · ' + dateOf(r);
/** A name that already says peer evaluation isn't said twice. */
const peName = (name, sep) => (PE.test(name) ? name : 'Peer evaluation' + sep + name);
/** A frozen board's start mark: the sheet takes no other board once a started one is in. */
const startMark = () => 2 + (crypto.getRandomValues(new Uint32Array(1))[0] % 1000000000);
/** The write token for the sheet (kept in memory for the page-leave send). */
async function wt(r) {
  if (!r.wt || r.wtFor !== r.e) {
    r.wt = await writeToken(r.w, r.e);
    r.wtFor = r.e;
  }
  return r.wt;
}
const ICON_BACK = '<svg width="22" height="22" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#2F4A27" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M15 5 L8 12 L15 19"/></svg>';
const ICON_FILE = '<svg width="40" height="40" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#3F5F33" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 3 H14 L19 8 V21 H6 Z"/><path d="M14 3 V8 H19"/><path d="M9 13 H16"/><path d="M9 17 H16"/></svg>';
const ICON_PEOPLE = '<svg width="34" height="34" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#3F5F33" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="8" r="3.2"/><path d="M3.5 19.5c.6-3.2 2.8-5 5.5-5s4.9 1.8 5.5 5"/><circle cx="16.5" cy="9" r="2.6"/><path d="M15.2 14.6c2.6-.2 4.6 1.4 5.3 4.4"/></svg>';
const ICON_OK = '<svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#3F5F33" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5 L10 17 L19 7"/></svg>';

/* ---------- keeping evaluations on this device ---------- */

function persist() {
  if (!own) return;
  try {
    localStorage.setItem(STORE, JSON.stringify({ v: 1, list }));
  } catch (err) {
    // storage full or blocked: the cadre link still brings it back
  }
}

function restore() {
  try {
    own = localStorage.getItem(OWN) === '1';
    const s = own ? JSON.parse(localStorage.getItem(STORE) || 'null') : null;
    return s && Array.isArray(s.list) ? s.list : [];
  } catch (err) {
    return [];
  }
}

function keepHere() {
  try {
    localStorage.setItem(OWN, '1');
  } catch (err) {
    return toast('This browser can\'t keep it. Keep page 2 of the printout.');
  }
  own = true;
  if (R && list.indexOf(R) < 0) list.unshift(R);
  ui.kept = true;
  persist();
  toast('This ' + device() + ' keeps it now.');
  render();
}

/** This tab has left the evaluation its cadre link opened: a reload no longer opens it. */
function dropLink() {
  try {
    sessionStorage.removeItem(LINK);
  } catch (err) {
    // nothing kept
  }
}

function forgetHere(r) {
  list = list.filter((x) => x !== r);
  if (own) {
    try {
      if (!list.length) localStorage.removeItem(STORE);
      else persist();
    } catch (err) {
      // nothing kept
    }
  }
}

/* ---------- small helpers ---------- */

/** A short line at the bottom; with label and fn, one button (Undo) that runs fn. */
function toast(text, label, fn) {
  clearTimeout(timers.toast);
  let t = document.querySelector('.toast');
  if (!t) {
    t = document.createElement('div');
    t.className = 'toast';
    t.setAttribute('role', 'status');
    document.body.appendChild(t);
  }
  t.textContent = text;
  if (label) {
    const b = document.createElement('button');
    b.className = 'link';
    b.textContent = label;
    b.addEventListener('click', () => {
      clearTimeout(timers.toast);
      t.remove();
      fn();
    });
    t.appendChild(b);
  }
  timers.toast = setTimeout(() => t.remove(), label ? 6000 : 3200);
}

function download(name, data, type) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([data], { type }));
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    URL.revokeObjectURL(a.href);
    a.remove();
  }, 1000);
}

async function copy(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch (err) {
    const t = document.createElement('textarea');
    t.value = text;
    document.body.appendChild(t);
    t.select();
    const ok = document.execCommand && document.execCommand('copy');
    t.remove();
    return !!ok;
  }
}

function qrSvg(text) {
  const qr = qrcode(0, 'M');
  qr.addData(text, 'Byte');
  qr.make();
  return qr.createSvgTag({ cellSize: 4, margin: 0, scalable: true });
}

/* ---------- the roster, before the code ---------- */

/** added: OCs the cadre added (inserts), read like names typed in a sign-up and never matched to a file's; touched: rows fixed here (the Fixed tag). */
function newDraft() {
  return { sources: [], dec: {}, seq: 0, pltSel: 0, pltAll: 0, mode: 'set', setMap: {}, name: '', sq: 4, added: [], touched: {} }; // 17c, the user: squads not in the file are the cadre's to set, never the OCs' to choose
}

function sourceOf(r, name, units) {
  const D = ui.draft;
  D.seq += 1;
  return {
    id: 's' + D.seq, name,
    rows: r.rows.map((x) => ({ last: x.last, first: x.first || '', plt: x.plt || null, plts: x.plts || null, sqd: x.sqd || null, two: !!x.two })),
    unread: r.unread || [], warn: r.warn || [], columns: r.columns || [], empty: r.empty || 0, titles: r.titles || 0, tabs: r.tabs || 0, tab: r.tab || '',
    plt: units ? units.plt : null, sqd: units ? units.sqd : null
  };
}

function addSource(r, name, units) {
  if (!r.rows.length && !r.unread.length) return toast(name + ': no names found in it.');
  const D = ui.draft;
  const taken = D.sources.filter((s) => s.name === name || s.name.indexOf(name + ' (') === 0).length;
  D.sources.push(sourceOf(r, taken ? name + ' (' + (taken + 1) + ')' : name, units));
  if (!D.name) D.name = name === 'Pasted list' || /class|roster/i.test(name) ? 'Peer evaluation' : name.replace(/\.[A-Za-z0-9]+$/, '').replace(/[_-]+/g, ' ').slice(0, 40);
  go('read');
}

/** Zips (a Classroom or Drive download) opened into their Excel files; anything else as it is. */
async function expandFiles(files) {
  const out = [];
  for (const f of files) {
    if (!/\.zip$/i.test(f.name || '')) {
      out.push(f);
      continue;
    }
    try {
      (await xlsxInZip(new Uint8Array(await f.arrayBuffer()))).forEach((x) => out.push(new File([x.data], x.name, { lastModified: x.time || 0 }))); // 0: no date (the order in the upload decides)
    } catch (err) {
      toast((f.name || 'That zip') + ' can\'t be opened.');
    }
  }
  return out;
}

/** Every name on the form tabs of a workbook: each tab's rater (E4) and peer (E5), as "LAST, First". */
function formNames(tabs) {
  const at = (c, ref) => String(((c[Number(ref.slice(1)) - 1] || [])[ref.charCodeAt(0) - 65]) ?? '').trim();
  const out = [];
  tabs.forEach((t) => {
    if (!/^select your last name$/i.test(at(t.cells || [], 'D4'))) return;
    ['E4', 'E5'].forEach((ref) => {
      const v = at(t.cells, ref);
      if (v && out.indexOf(v) < 0) out.push(v);
    });
  });
  return out;
}

async function addFiles(files) {
  files = await expandFiles(files);
  const filled = []; // 17j: filled-in forms from Classroom: their squads make the roster, and they go in as results once the evaluation is made
  const squads = {};
  for (const f of files) {
    const name = f.name || 'file';
    const ext = ((/\.([A-Za-z0-9]+)$/.exec(name) || [])[1] || '').toLowerCase();
    try {
      if (ext === 'xlsx' || ext === 'xlsm') {
        const tabs = await readXlsx(new Uint8Array(await f.arrayBuffer()));
        const forms = readForms(tabs);
        if (forms.length) {
          const u = fileUnits(name);
          const plt = forms[0].plt || u.plt || 0;
          const sqd = forms[0].sqd || u.sqd || 0;
          const s = (squads[plt + '|' + sqd] = squads[plt + '|' + sqd] || { plt, sqd, names: [], n: 0 });
          formNames(tabs).forEach((x) => {
            if (s.names.indexOf(x) < 0) s.names.push(x);
          });
          s.n++;
          filled.push(f);
          continue;
        }
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
        toast(name + ' can\'t be read here. Save it as .xlsx in Excel, or copy the names and paste them.');
      }
    } catch (err) {
      toast(name + (err && err.message === 'locked' ? ' has a password or is an old .xls: save it as .xlsx without one, or paste the names.' : ' can\'t be read: save it as .xlsx, or paste the names.'));
    }
  }
  if (!filled.length) return;
  ui.draft.forms = (ui.draft.forms || []).concat(filled);
  ui.draft.name = ui.draft.name || 'Peer evaluation'; // not the first squad's label
  Object.values(squads).sort((a, b) => a.plt - b.plt || a.sqd - b.sqd).forEach((s) => {
    const text = ['Last\tFirst\tPLT\tSQD'].concat(s.names.map((x) => {
      const i = x.indexOf(',');
      return (i < 0 ? x : x.slice(0, i)).trim() + '\t' + (i < 0 ? '' : x.slice(i + 1).trim()) + '\t' + (s.plt || '') + '\t' + (s.sqd || '');
    })).join('\n');
    addSource(readRoster(text), (s.plt ? s.plt + 'PLT ' : '') + (s.sqd ? ord(s.sqd) + ' SQD' : 'Forms') + ' · ' + plural(s.n, 'filled form', 'filled forms'), { plt: s.plt || null, sqd: s.sqd || null });
  });
}

/** The merged roster for the draft: the Build 10 rules, with the squads left to the OCs where none are given. */
function draftRoster() {
  const D = ui.draft;
  const sources = D.sources.map((s) => (D.pltAll && !s.plt && s.rows.every((x) => !x.plt && !(x.plts || []).length) ? Object.assign({}, s, { plt: D.pltAll }) : s));
  const m = mergeRoster(sources, D.dec, { extra: D.added.map((x) => (D.pltAll && !x.plt ? Object.assign({}, x, { plt: D.pltAll }) : x)) }, {});
  const plts = [...new Set(m.ocs.map((o) => o.plt).filter(Boolean))].sort((a, b) => a - b);
  if (D.pltSel && plts.indexOf(D.pltSel) < 0) D.pltSel = 0; // its platoon emptied (a fix moved them, or all were left out): All again
  const inPlt = (o) => !D.pltSel || o.plt === D.pltSel || (!o.plt && (!o.plts || o.plts.indexOf(D.pltSel) >= 0)); // no clear platoon: it stays, with its buttons
  const ocs = m.ocs; // the evaluation is always the whole company (the user, 11-24): a platoon tab only narrows the list on screen
  const shown = m.ocs.filter(inPlt);
  const noPlt = ocs.filter((o) => !o.plt);
  const namesOnly = m.ocs.length > 0 && !plts.length;
  // two OCs with one name and no first names are fine while OCs choose: each phone asks the first name; with a first name it's one OC twice
  const block = m.block.filter((t) => !/ no squad yet\.$| no PLT yet\.$| question(s)? to answer first\.$| is listed twice: /.test(t));
  const twice = ocs.find((o, i) => o.first && ocs.some((x, j) => j < i && sameOc(x, o)));
  if (twice) block.unshift(fullName(twice.last, twice.first) + ' is on the list twice. Tap one and leave it out, or fix its name if they are two OCs.');
  const sqd = ocs.filter((o) => o.sqd).length;
  const allSquads = ocs.length > 0 && sqd === ocs.length;
  const twin = twinIn(ocs); // the cadre sets every squad and the evaluation opens with the QR code, so no phone asks a first name
  if (twin) block.unshift('Two OCs named ' + twin.last + ' with no first names in ' + (twin.plt ? ord(twin.plt) + ' PLT' : 'one platoon') + '. Tap each one and add a first name.');
  const questions = m.questions;
  return { m, ocs, shown, outs: m.outs.concat(m.drops).filter(inPlt), plts, noPlt, namesOnly, block, questions, allSquads, someSquads: sqd > 0 }; // a copy taken off is under Left out too, to put back
}

/** The squad tapped on Tap each OC's squad, by any of the OC's rows: a bigger file added later doesn't lose it. */
const setOf = (o) => o.keys.map((k) => ui.draft.setMap[k]).find(Boolean);

/** The board's list from the draft: [id, LAST, First, plt, sqd]; sqd from the file, or the cadre's taps. */
function draftList(dr) {
  const D = ui.draft;
  const ocs = dr.ocs.slice().sort((a, b) => (a.plt || 0) - (b.plt || 0) || (a.sqd || 0) - (b.sqd || 0) || a.last.localeCompare(b.last) || (a.first || '').localeCompare(b.first || ''));
  return ocs.map((o, i) => [i, o.last, o.first || '', o.plt || 0, D.mode === 'set' && !dr.allSquads ? setOf(o) || o.sqd || 0 : o.sqd || 0]);
}

/* ---------- the evaluation's state ---------- */

function fresh() {
  return { E: null, picks: {}, trail: {}, results: {}, cursor: [0, 0], err: false, polling: null, extra: [], sealed: null, dirty: false, opened: false, loaded: false, boardTimer: 0, startTimer: 0, seen: 0, base: null };
}

/** The board is frozen for Start but not known to be in yet: no squad moves meanwhile. */
const frozen = () => starting.has(R.e) || (R.board.started && !R.started);

const ocOf = (id) => R.board.oc.find((o) => o[0] === id);
const placedByCadre = (id) => !!(R.fixed && R.fixed[id]);

/** A typed name is a board row's name: the same last name, and the same first name when both have one. */
const sameTyped = (o, last, first) => o[1] === last && (!first || !o[2] || String(o[2]).toLowerCase() === String(first).toLowerCase());
/** A board row has this very name: the last name, and the first name (none on both, or the same). */
const sameName = (o, last, first) => o[1] === last && String(o[2] || '').toLowerCase() === String(first || '').toLowerCase();
/** Where a typed name's Leave out is kept: by phone and name, so another name typed on that phone still shows. */
const ignKey = (x) => x.dv + '|' + x.last + (x.first ? '|' + x.first.toLowerCase() : '');
/** The letters of a last name, to see two OCs with one last name. */
const lastKey = (s) => String(s || '').toUpperCase().replace(/[^\p{L}]/gu, '');
/** Two first names a typo or a short form apart (Jhon, John; Mike, Michael), or one missing. */
const firstNear = (a, b) => !lastKey(a) || !lastKey(b) || lastKey(a)[0] === lastKey(b)[0] || nearMiss(a, b);
/** The name a phone's pick is for: a roster row, a twin with the first name it sent, or a typed name. */
const pickName = (p) => {
  const g = p.r >= 0 ? ocOf(p.r) : Array.isArray(p.grp) ? ocOf(p.grp[0]) : null;
  if (p.r >= 0) return g ? [g[1], g[2] || ''] : ['', ''];
  return [g ? g[1] : String(p.last || '').toUpperCase().slice(0, 30), String(p.first || '').trim().slice(0, 30)];
};
/** A pick of an added row's own phone is for that row: its id, or the name typed (as the board has it, or as the cadre linked it). */
const ownPick = (o, p) => {
  if (p.r === o[0]) return true;
  if (p.r !== -1 || p.grp) return false;
  const last = String(p.last || '').toUpperCase().slice(0, 30);
  const first = String(p.first || '').trim().slice(0, 30);
  return (p.me === 0 ? sameName : sameTyped)(o, last, first) || (R.board.ln || {})[ignKey({ dv: o[5], last, first })] === o[0]; // he said someone else: only this very name (a row with no first name isn't his)
};
/**
 * An added newcomer's phone since he was added: '' while its newest pick is still his row; 'same' when it signed up a
 * name close to his since (a fixed spelling, or his own roster name: the same OC, so his row waits for the cadre),
 * 'lent' when it signed up someone else (a lent phone: his row keeps the squad that phone's last pick for him gave it).
 */
function awayOf(o) {
  const p = o && o[5] && X.picks[o[5]];
  if (!p || ownPick(o, p)) return '';
  if (p.me === 1 || p.me === 0) return p.me ? 'same' : 'lent'; // the person said, on the phone: no guess from the names
  const [last, first] = pickName(p);
  return (lastKey(last) === lastKey(o[1]) || nearMiss(last, o[1])) && firstNear(first, o[2]) ? 'same' : 'lent';
}

/** Each phone's newest pick; one phone counts once. Typed names (not on the list) wait for Add or Ignore; once added, the row goes by its phone's pick, as any OC's. */
function picks() {
  const byId = {};
  const extra = [];
  const firsts = {};
  const fdv = {}; // the phone that sent each first name
  const claimed = {};
  const ign = R.ignored || {};
  const nd = R.board.nd || {}; // twin rows a phone gave its first name to: that phone's id
  Object.entries(X.picks).sort((a, b) => String(a[1].at).localeCompare(String(b[1].at))).forEach(([dv, p]) => {
    let id = p.r;
    if (id === -1 && Array.isArray(p.grp)) { // twins with one last name: the row this phone named (a new spelling is theirs to make), or the row with this first name (this OC's other phone took it)
      const f = String(p.first || '').trim().toLowerCase();
      const named = (g) => String((ocOf(g) || [])[2] || firsts[g] || '').trim().toLowerCase();
      const own = p.grp.find((g) => ocOf(g) && nd[g] === dv); // the row this phone named
      const mine = own !== undefined && (p.me === 1 || (p.me !== 0 && firstNear(named(own), f))) ? own : undefined; // a re-spelling of the name it gave (or he said it's him); another name from that phone is the other twin (a lent phone)
      const like = f ? p.grp.filter((g) => ocOf(g) && !claimed[g] && named(g) && named(g) !== f && (nearMiss(named(g), f) || firstOk(named(g), f))) : []; // one row a typo or a middle initial apart (Jhon, Mike Q); Mary Ann and Mary Beth are two
      id = mine !== undefined ? mine : f ? p.grp.find((g) => ocOf(g) && named(g) === f) : undefined;
      if (id === undefined && like.length === 1 && ocOf(like[0])[2] && !nd[like[0]]) id = like[0]; // the cadre named it so (Michael, for the twin who types Mike)
      if (id === undefined) { // a new first name: a free row (never a newcomer the cadre added: his phone has it), first the one the file (or the cadre) put in the squad this OC tapped, then one nobody placed; one left out comes last
        const free = p.grp.filter((g) => !claimed[g] && ocOf(g) && !ocOf(g)[5] && !named(g) && !nd[g]);
        id = free.find((g) => placedByCadre(g) && ocOf(g)[4] === p.sqd);
        if (id === undefined) id = free.find((g) => !placedByCadre(g));
        if (id === undefined) id = free.find((g) => ocOf(g)[4]);
        if (id === undefined) id = free[0];
      }
      if (id === undefined && like.length === 1) id = like[0]; // no twin row free: the one a typo apart
      if (id === undefined && p.me !== 0) id = own; // no other twin row free: the row this phone named, under another name (Tony, Anthony); never when he said someone else
      if (id === undefined) return;
      claimed[id] = true;
      if (p.first && (mine !== undefined || !firsts[id])) {
        firsts[id] = firstCase(String(p.first)).slice(0, 30); // tony from an older phone: Tony, as the cadre's fields write it
        fdv[id] = dv;
      }
    }
    if (id === -1) { // a typed name: once the cadre added it (on any of the cadre's devices: the board has it), that row's pick (a newer squad moves it); or left out
      const last = String(p.last || '').toUpperCase().slice(0, 30);
      const first = String(p.first || '').trim().slice(0, 30);
      const x = { dv, last, first, plt: p.plt, sqd: p.sqd, told: !!p.told, me: p.me }; // told: the phone asked "X is on the list. You?" and the OC said someone else; me: his answer to "This phone signed up X. Is that you?"
      const ln = (R.board.ln || {})[ignKey(x)];
      const row = R.board.oc.find((o) => o[5] === dv && (p.me === 0 ? sameName : sameTyped)(o, last, first)) || (ln !== undefined ? ocOf(ln) : undefined); // "someone else": this very name only, never the phone's nameless row // its own phone's, or the row the cadre said is this OC (Same OC: his roster name, or added from his other browser)
      if (!row) {
        if (last && !ign[ignKey(x)]) extra.push(x);
        return;
      }
      id = row[0];
    }
    const o = ocOf(id);
    if (!o) return;
    if (p.r >= 0 && p.first && (nd[id] === dv || (!o[2] && !firsts[id]))) { // no first name next to another OC with that last name: the one its phone gave (that phone may fix its spelling)
      firsts[id] = firstCase(String(p.first)).slice(0, 30);
      fdv[id] = dv;
    }
    (byId[id] = byId[id] || []).push({ dv, sqd: p.sqd });
  });
  Object.entries(X.picks).filter(([, p]) => p.me === 0).forEach(([dv, p]) => { // "No, I'm someone else" on a lent phone: each OC it signed up before keeps his pick, back along the chain while each newer one said someone else
    const who = (q) => (q.r >= 0 ? 'r' + q.r : (Array.isArray(q.grp) ? 'g' : 't' + String(q.last || '').toUpperCase()) + '|' + String(q.first || '').trim().toLowerCase());
    const trail = (X.trail[dv] || []).slice().sort((a, b) => b.rev - a.rev); // newest first
    const kept = new Set([who(p)]);
    let cur = p;
    while (cur && cur.me === 0) {
      const prev = trail.find((q) => q.rev < cur.rev && !kept.has(who(q)));
      if (!prev) break;
      kept.add(who(prev));
      const id = prev.r >= 0 ? prev.r : Array.isArray(prev.grp) ? prev.grp.find((g) => ocOf(g) && nd[g] === dv) : undefined; // a roster row, or the twin row it named (an added row keeps his squad as a lent row)
      if (id !== undefined && ocOf(id) && !ocOf(id)[5] && !(byId[id] || []).some((x) => x.dv === dv)) (byId[id] = byId[id] || []).push({ dv, sqd: prev.sqd });
      cur = trail.filter((q) => who(q) === who(prev)).sort((a, b) => a.rev - b.rev)[0]; // that OC's first pick here: someone else too (me 0)? then the one before him keeps his as well
    }
  });
  return { byId, extra, firsts, fdv };
}

/** An added newcomer of this nameless OC's last name, while this OC (placed by the file or the cadre) hasn't opened the link: maybe the same OC. */
const besideOf = (o, pk) => (!o[2] && !o[5] && !((pk || picks()).byId[o[0]] || []).length ? R.board.oc.find((y) => y[5] && y[0] !== o[0] && lastKey(y[1]) === lastKey(o[1]) && (!y[3] || !o[3] || y[3] === o[3])) : undefined);
/** Where an OC is now: a squad number, 0 = out of this evaluation, null = not in yet (or in a squad above the squads per platoon), -1 = two phones disagree. */
function placeOf(id, pk) {
  const o = ocOf(id);
  if (R.started) return o[4] || 0;
  const away = awayOf(o);
  if (away === 'same') return null; // his phone signed up his own name since (a fixed spelling, his roster name): this row waits for the cadre, even where placed
  let q;
  if (placedByCadre(id)) {
    q = o[4] || 0;
    if (q && besideOf(o, pk)) return null; // maybe the newcomer added beside him: Start asks
  } else if (away === 'lent') { // a lent phone: his row keeps the squad its last pick for him gave it, or follows a newer pick for him on another phone
    const mine = (X.trail[o[5]] || []).filter((p) => ownPick(o, p)).sort((a, b) => b.rev - a.rev)[0];
    const last = [mine].concat(((pk || picks()).byId[id] || []).map((x) => X.picks[x.dv])).filter(Boolean).sort((a, b) => String(b.at).localeCompare(String(a.at)))[0];
    q = last ? last.sqd : o[4] || null; // no pick read yet: the board's
  } else {
    const p = (pk || picks()).byId[id];
    if (!p || !p.length) return null;
    const sq = [...new Set(p.map((x) => x.sqd))];
    q = sq.length > 1 ? -1 : sq[0];
  }
  return q > (R.sq || 4) ? null : q;
}

async function refreshEvent() {
  X.E = await eventOf(R.board, R.k, privOf(R.d, R.pub));
  Object.assign(X.E.results, X.results);
}

/** Seal the board now (so leaving the page can still send it) and send it within 30 s. */
async function boardChanged(r = R, x = X) {
  x.dirty = true;
  r.brev = (r.brev || 0) + 1;
  const brev = r.brev;
  r.board.at = r.board.at || r.created;
  r.board.fx = Object.keys(r.fixed || {}).map(Number); // the cadre's own placements, a Leave out too: another device of the cadre's keeps them
  if (r.board.started && !r.board.st) r.board.st = startMark(); // a board started before the start mark gets one
  const st = r.board.started ? r.board.st : 0;
  const plain = JSON.stringify(r.board); // once it is in, the board the sheet has: this device's next changes are the ones made since
  const board = await sealBoard(r.k, r.board);
  if (x.sealed && x.sealed.brev > brev) return; // a newer change was sealed meanwhile: that one goes
  x.sealed = { brev, st, board, plain };
  persist();
  clearTimeout(x.boardTimer);
  x.boardTimer = setTimeout(() => sendBoard(r, x), BOARD_MS);
}

/** The cadre's own placements on a board from the sheet: a squad on the sign-up board, and the ones in fx (a Leave out is 0 there). */
function takeFixed(r, b) {
  r.fixed = {};
  b.oc.forEach((o) => {
    if ((o[4] && !o[5]) || b.mode === 'set') r.fixed[o[0]] = true; // an added newcomer's squad on the board is his phone's pick (for older pages): only fx fixes him
  });
  (b.fx || []).forEach((id) => {
    r.fixed[id] = true;
  });
}

/** Send r's sealed board; it stays tied to r even if the cadre opens another evaluation meanwhile. */
async function sendBoard(r, x) {
  if (!x || !x.sealed) return true;
  const s = x.sealed;
  if (s.st && R === r && X !== x) { // frozen in a working state the cadre left by opening the evaluation again: it never goes in later
    x.sealed = null;
    clearTimeout(x.boardTimer);
    return true;
  }
  if (!s.st && (await adoptStarted(r, x))) return true; // a sign-up board never goes over a start made on another device
  if (x.sealed !== s) return true;
  let res;
  try {
    res = await putBoard(r.e, await wt(r), s.brev, s.board, s.st);
  } catch (err) {
    res = null;
  }
  if (res && res.ok) {
    if (s.plain) x.base = { brev: s.brev, board: s.plain }; // the board the sheet has now
    if (x.sealed === s) {
      x.dirty = false;
      x.sealed = null;
    }
    if (s.st && !r.started) { // a frozen board went in (on any send): that is the start, with the board that went in
      try {
        r.board = await openBoard(r.k, r.e, s.board);
        r.brev = Math.max(r.brev || 0, s.brev);
      } catch (err) {
        // the board here is that board
      }
      await begun(r);
    }
    return true;
  }
  if (res && res.error === 'old') {
    if (!x.sealed || x.sealed.brev < r.brev) { // it went in on another send, or r has a newer board since: never send an older one over it
      x.sealed = null;
      return true;
    }
    if (await adoptStarted(r, x)) return true; // started (here or on another device): the sheet's board stays
    if (x.sealed !== s) return true; // another cadre device's board came in, with this change over it: that one goes next
    r.brev = Math.max(r.brev, Number(res.brev) || 0) + 1;
    x.sealed = Object.assign({}, s, { brev: r.brev });
  }
  if (res && res.error === 'removed') return markRemoved(r);
  clearTimeout(x.boardTimer);
  x.boardTimer = setTimeout(() => sendBoard(r, x), res && res.error === 'busy' ? 2500 : 20000);
  return false;
}

/** If the sheet's board is a started one, take it (the change made here is dropped); a newer sign-up board is taken in too, with this device's changes over it. @returns {Promise<boolean>} true when it took a started one */
async function adoptStarted(r, x) {
  let g = null;
  let b = null;
  try {
    g = await getBoard(r.e);
    if (g && g.ok) b = await openBoard(r.k, r.e, g.board);
  } catch (err) {
    b = null;
  }
  if (g && g.ok) x.seen = Math.max(x.seen, g.brev);
  if (b && !b.started) await takeNewer(r, x, g, b);
  if (!b || !b.started) return false; // started = a frozen board on the sheet; its flag alone over a sign-up board is not (as on the phones)
  await takeStarted(r, x, g, b);
  return true;
}

/** A sign-up board another cadre device sent since the one this device last had from the sheet: the board here now, with this device's own changes since over it (they go in next). */
async function takeNewer(r, x, g, b) {
  if (!x.base || g.brev <= x.base.brev || r.board.started || (R === r && X !== x)) return; // the board known here; a Start's freeze here (it goes in as it is); a working state the cadre left
  const m = mergeBoard(b, JSON.parse(x.base.board), r.board);
  x.base = { brev: g.brev, board: JSON.stringify(b) };
  r.board = m.board;
  r.name = r.board.n || r.name;
  r.sq = r.board.sq || r.sq;
  takeFixed(r, r.board);
  r.brev = Math.max(r.brev || 0, g.brev);
  if (m.own) await boardChanged(r, x);
  else {
    clearTimeout(x.boardTimer);
    x.sealed = null; // nothing of this device's own to send over it
    x.dirty = false;
    persist();
  }
  if (R !== r || X !== x) return;
  await refreshEvent();
  if ((ui.screen === 'lobby' || ui.screen === 'qr') && !ui.sheet) {
    ui.keepScroll = true;
    render();
  }
}

/** The sheet holds a started board: take it, drop the change made here, show the progress. */
async function takeStarted(r, x, g, b) {
  x.sealed = null;
  r.board = b;
  r.brev = g.brev;
  r.name = b.n || r.name;
  r.sq = b.sq || r.sq;
  r.started = true;
  persist();
  if (R === r) {
    await refreshEvent();
    if (ui.screen === 'lobby') go('progress');
  }
  return true;
}

/** The started board is on the sheet and phones open from it: this is the start. */
async function begun(r) {
  if (r.started) return;
  r.started = true;
  persist();
  if (R !== r) return;
  await refreshEvent();
  if (ui.screen === 'lobby') go('progress');
  poll(true);
}

/* ---------- making the QR code ---------- */

async function makeCode(oc, mode, name, sq, from, seed) {
  dropLink();
  const K = await newKeys();
  const fixed = {};
  oc.forEach((o) => {
    if (o[4] || mode === 'set') fixed[o[0]] = true;
  });
  const top = Math.max(sq || 4, ...oc.map((o) => o[4] || 0)); // never fewer squads than the roster has
  R = {
    e: K.e, k: K.k, w: K.w, d: K.d, pub: K.pub, name: name || 'Peer evaluation', created: new Date().toISOString(), sq: top, mode, brev: 0,
    started: mode === 'set', page2: false, shared: false, used: false, removed: false, fixed, out: {}, ignored: {}, from: from || '',
    board: { v: 2, e: K.e, n: name || 'Peer evaluation', at: '', pub: K.pub, seed: seed || newSeed(), dl: 3, sq: top, mode, started: mode === 'set', oc } // the draft's seed if its blank sheets went out: the phones get the same attributes; dl 3: each squad dealt by its own platoon and number, 6 different attributes for each rater and peer (17e)
  };
  if (mode === 'set') R.board.st = startMark();
  X = fresh();
  ui.saveProblem = '';
  go('saving');
  await saveNew();
}

const newSeed = () => crypto.getRandomValues(new Uint32Array(1))[0] % 1000000;

/**
 * The original paper form, one workbook per squad (17i): each OC picks his name on it, and the attributes fill in by
 * formula. To fill in if the app has trouble, or with no app at all. oc = rows [id, LAST, First, plt, sqd] as placed now;
 * none = OCs with no squad yet.
 */
function saveBlankSheets(oc, seed, name, none) {
  const board = { oc, seed };
  const squads = squadsOf(board).map((sq) => ({ plt: sq.plt, sqd: sq.sqd, ocs: sq.ids.map((id) => {
    const o = oc.find((x) => x[0] === id);
    return { id, last: o[1], first: o[2] || '' };
  }) }));
  const f = blankBooks(name, squads, seed);
  if (!f) return toast('No squads yet. Come back once your OCs are in their squads.');
  download(f.name, f.data, /\.zip$/.test(f.name) ? 'application/zip' : XLSX_TYPE);
  return toast('Saved: one Excel workbook per squad. Each OC picks their name in it, and their attributes fill in.' + (none ? ' Not in it: ' + plural(none, 'OC', 'OCs') + ' with no squad yet.' : ''));
}

/** What blank sheets say, by squad: its OCs and the attributes they get. { 'plt-sqd': text } for squads of 2 or more. */
function sheetsSig(oc, seed) {
  const board = { oc, seed, dl: 3 };
  const hands = handsOf(board);
  const sig = {};
  squadsOf(board).filter((sq) => sq.ids.length > 1).forEach((sq) => (sig[sq.key] = sq.ids.map((id) => {
    const o = oc.find((x) => x[0] === id);
    return o[1] + ',' + o[2] + '=' + JSON.stringify(hands[id]);
  }).join(';')));
  return sig;
}

/** On Who's in it? and Tap each OC's squad: a squad's names or attributes changed after its blank sheets were saved (a new squad is only not in them). */
function staleLine(dr) {
  const D = ui.draft;
  if (!D.sheetsSig) return '';
  const now = sheetsSig(draftList(dr), D.seed);
  if (Object.keys(D.sheetsSig).every((k) => !(k in now) || now[k] === D.sheetsSig[k])) return '';
  return '<div class="notice">Blank sheets saved before this no longer match: save them again. ' + link('sheets-draft', 'Blank sheets for OCs (Excel)') + '</div>';
}

/** new, then the first board, then (squads known) start: retried quietly until they go in. */
async function saveNew() {
  clearTimeout(timers.retry);
  const r = R;
  let res;
  try {
    if (!r.made) {
      res = await newEval(r.e, await writeHash(await wt(r)));
      if (res && res.error === 'taken') {
        const K = await newKeys();
        Object.assign(r, { e: K.e, k: K.k, w: K.w, d: K.d, pub: K.pub });
        Object.assign(r.board, { e: K.e, pub: K.pub });
        return saveNew();
      }
      if (res && res.error === 'limit') return stuck('limit', 180000);
      if (!res || !res.ok) return stuck('reach', 20000);
      r.made = true;
      r.created = new Date(res.created || Date.now()).toISOString();
      r.board.at = r.created;
    }
    if (!r.brev) {
      const b = await sealBoard(r.k, r.board);
      res = await putBoard(r.e, await wt(r), 1, b, r.board.started ? r.board.st : 0);
      if (!res || !(res.ok || res.error === 'old')) return stuck('reach', 20000);
      r.brev = 1;
      if (R === r && X) X.base = { brev: 1, board: JSON.stringify(r.board) };
    }
    if (r.mode === 'set' && !r.startSent) {
      res = await startEval(r.e, await wt(r));
      if (!res || !res.ok) return stuck('reach', 20000);
      r.startSent = true;
    }
  } catch (err) {
    return stuck('reach', 20000);
  }
  if (R !== r) return undefined;
  if (list.indexOf(r) < 0) list.unshift(r);
  persist();
  await refreshEvent();
  ui.saveProblem = '';
  if (paperForms && paperForms.length) { // 17j: forms collected from Classroom: no QR code to hand out; they go in now
    const f = paperForms;
    paperForms = null;
    go('progress');
    importFiles(f);
  } else go('qr');
  schedulePoll(POLL_MS);
  return undefined;
}

function stuck(kind, wait) {
  ui.saveProblem = kind;
  if (ui.screen === 'saving') render();
  timers.retry = setTimeout(saveNew, wait);
}

/* ---------- reading the sheet ---------- */

function schedulePoll(ms) {
  clearTimeout(timers.poll);
  timers.poll = setTimeout(() => poll(false), ms === undefined ? (X && X.err ? ERR_MS : POLL_MS) : ms);
}

async function poll(force) {
  if (!R || !X || R.removed || !R.brev) return undefined;
  if (document.hidden && !force) return schedulePoll();
  if (X.polling) return X.polling; // a read on its way: its answer is this one's too (Start waits for it)
  const x = X;
  const p = readSheet(R, x);
  x.polling = p;
  const done = () => {
    if (x.polling === p) x.polling = null;
  };
  p.then(done, done);
  return p;
}

/** One read of the sheet: the rows since the last read, then (not started) its state. */
async function readSheet(r, x) { // an answer counts only for the working state that asked for it
  const errBefore = x.err;
  let changed = false;
  let thawed = false;
  try {
    for (let i = 0; i < 25; i++) {
      const res = await getRows(r.e, await wt(r), x.cursor);
      if (R !== r || X !== x) return undefined;
      if (!res.ok) {
        if (res.error === 'removed') return markRemoved(r);
        throw new Error(res.error);
      }
      x.cursor = res.cur;
      for (const item of res.items) if (await takeRow(item, r, x)) changed = true;
      const macs = res.items.map((it) => { // results added here that the sheet has now (sent from here, or by their phones) need no keeping
        try {
          return (typeof it === 'string' ? JSON.parse(it) : it).mac;
        } catch (err) {
          return '';
        }
      });
      if ((r.imported || []).some((env) => macs.indexOf(env.mac) >= 0)) {
        r.imported = r.imported.filter((env) => macs.indexOf(env.mac) < 0);
        persist();
      }
      if (!res.more) break;
    }
    x.loaded = true; // the picks are in: Start can count them
    if ((r.imported || []).length) sendImported(r); // the sheet answers again: the ones it doesn't have yet go now
    if (!r.started && !starting.has(r.e)) { // another cadre device may have started it: then its board shows here too
      const s = await getState(r.e);
      if (R !== r || X !== x) return undefined;
      const stale = r.board.started && !x.sealed; // frozen by a Start no longer on its way (the page was reloaded, or opened again)
      const took = s.ok && (s.started || stale || s.brev > Math.max(r.brev, x.seen)) && (await adoptStarted(r, x));
      if (s.ok && !s.started && stale && !took && R === r && X === x && !r.started && !x.sealed && !starting.has(r.e)) {
        thaw(r); // that Start never went in: the sign-up goes on
        persist();
        thawed = true;
      }
    }
    x.err = false;
  } catch (err) {
    x.err = true;
  }
  if (R !== r || X !== x) return undefined;
  if (changed) {
    if (!r.used) {
      r.used = true;
      persist();
    }
    if (!r.started && !r.board.started && ui.screen === 'lobby' && noteNames()) boardChanged().then(() => sendBoard(r, x)); // what the phones' picks changed goes at once: a cadre page on 16p reads a lent row's squad from the board
  }
  // redraw only when something changed, so a tap is never lost to a redraw
  if (r.started && ui.screen === 'progress' && allIn()) go('results');
  else if ((changed || thawed || X.err !== errBefore) && ['qr', 'lobby', 'progress', 'results'].indexOf(ui.screen) >= 0 && !ui.sheet) {
    ui.keepScroll = true;
    render();
  }
  schedulePoll();
  return undefined;
}

/** One row of the sheet: a squad pick or a result, for r's working state x. Rows that don't open are ignored. */
async function takeRow(item, r = R, x = X) {
  let env;
  try {
    env = typeof item === 'string' ? JSON.parse(item) : item;
  } catch (err) {
    return false;
  }
  if (!env || String(env.e) !== r.e) return false;
  if (env.j === 1) {
    const o = await openMarked(x.E, env);
    if (!o.ok) return false;
    const p = o.payload;
    const dv = String(p.dv || 'unknown');
    const cur = x.picks[dv];
    const pick = { r: p.r, rev: p.rev, at: p.at, plt: p.plt, sqd: p.sqd, first: p.first, last: p.last, grp: p.grp, told: p.told, me: p.me };
    x.trail[dv] = (x.trail[dv] || []).filter((q) => q.rev !== p.rev).concat([pick]); // each phone's picks: a lent phone's earlier one still places the OC it signed up first
    if (cur && cur.rev >= p.rev) return false;
    x.picks[dv] = pick;
    return true;
  }
  if (env.j) return false;
  const o = await openResult(x.E, env);
  if (!o.ok) return false;
  const cur = x.results[env.r];
  const dv = String(o.payload.dv || '');
  const devs = cur ? (cur.devs || []).slice() : [];
  if (dv && devs.indexOf(dv) < 0) devs.push(dv);
  if (cur) cur.devs = devs;
  if (cur && (cur.hash === o.hash || cur.rev > env.rev || (cur.rev === env.rev && String(o.payload.at) <= String(cur.at)))) return true;
  x.results[env.r] = { rev: env.rev, at: o.payload.at, hash: o.hash, source: 'sheet', payload: o.payload, devs };
  x.E.results[env.r] = x.results[env.r];
  r.last = new Date().toISOString();
  return true;
}

/** First names the phones sent for two OCs who share a last name go into the board. @returns {boolean} true when one went in: the caller sends the board */
function noteNames() {
  const pk = picks();
  let moved = false;
  R.board.oc.forEach((o) => { // a lent phone's newcomer: his squad on the board too, for a cadre page on 16p (it places a lent row by the board; this page by the picks)
    if (!o[5] || placedByCadre(o[0]) || awayOf(o) !== 'lent') return;
    const q = Math.max(0, placeOf(o[0], pk) || 0);
    if (q === (o[4] || 0)) return;
    o[4] = q;
    moved = true;
  });
  Object.entries(pk.firsts).forEach(([id, first]) => {
    const o = ocOf(Number(id));
    const nd = R.board.nd || {};
    if (o && o[2] && !o[5] && !nd[id] && pk.fdv[id]) { // a twin row the cadre named (Michael), taken by a phone's pick (Mike): that phone's row, so it opens it after Start
      R.board.nd = Object.assign({}, nd, { [id]: pk.fdv[id] });
      moved = true;
      return;
    }
    if (!o || (o[2] && !(nd[id] && nd[id] === pk.fdv[id] && o[2] !== first))) return; // a first name once, or a new spelling from the phone that gave it
    o[2] = first;
    R.board.nd = Object.assign({}, nd, { [id]: pk.fdv[id] }); // that phone's row from now on (on every cadre device: it's on the board)
    moved = true;
  });
  return moved;
}

function markRemoved(r) {
  r.removed = true;
  r.removedAt = new Date().toISOString();
  persist();
  if (R !== r) return undefined;
  clearTimeout(timers.poll);
  go('removed');
  return undefined;
}

const inEval = () => (R && R.started ? R.board.oc.filter((o) => o[4]) : []);
const counts = () => {
  const ids = X && X.E ? X.E.squads.filter((sq) => sq.ids.length > 1).flatMap((sq) => sq.ids) : [];
  return { done: ids.filter((id) => X.results[id]).length, total: ids.length };
};
const allIn = () => {
  const c = counts();
  return c.total > 0 && c.done >= c.total;
};

/* ---------- opening an evaluation ---------- */

async function openEval(r, screen) {
  if (X && X.sealed && R !== r) sendBoard(R, X); // a change not sent yet goes now, for the evaluation it was made on
  if (R !== r || !X) X = fresh(); // one working state per evaluation: opening the same one again keeps it
  R = r;
  clearTimeout(timers.poll);
  await refreshEvent();
  for (const env of r.imported || []) await takeRow(env, r, X); // results added here from a file or text, not on the sheet yet
  sendImported(r);
  go(screen || screenOf(r));
  if (r.removed) return undefined; // nothing to load; it leaves this device when the cadre moves on from it
  const shown = ui.screen;
  if (!r.brev) return saveNew();
  let took = false;
  try {
    const g = await getBoard(r.e);
    if (R !== r) return undefined;
    if (!g.ok && g.error === 'removed') return markRemoved(r);
    if (g.ok) {
      const b = await openBoard(r.k, r.e, g.board);
      if (R !== r) return undefined;
      const flying = starting.has(r.e) || !!(X.sealed && X.sealed.st); // this page's Start on its way: its frozen board goes in over any sign-up board
      if (!b.started && X.base && !flying && !r.board.started) await takeNewer(r, X, g, b); // opened again on this page: another device's changes come in, this one's stay over them
      else if (b.started || (g.brev > (r.brev || 0) && !flying)) { // started (a frozen board: the flag alone over a sign-up board isn't), or newer: the sheet's board counts
        took = true;
        r.board = b;
        r.brev = g.brev;
        r.name = b.n || r.name;
        r.sq = b.sq || r.sq;
        r.started = !!b.started;
        takeFixed(r, b); // what the cadre placed on another device holds here too
        await refreshEvent();
        if (b.started && !g.started) startSend(r, X); // the frozen board went in, the start didn't: it goes now
      } else if (r.board.started && !r.started && !flying) {
        thaw(r); // a Start that never went in (the page closed meanwhile): the sign-up goes on
      }
      if (X.base === null && !b.started) X.base = { brev: g.brev, board: JSON.stringify(b) }; // the board the sheet has: this device's changes from here are its own
      if (g.used) r.used = true;
      persist();
    }
  } catch (err) {
    // no answer now: the page works from what it has and checks again with the rows
  }
  X.opened = true;
  // the right screen once the board is in, unless the cadre has moved on meanwhile
  if (!screen && R === r && ui.screen === shown && ui.screen !== screenOf(r)) go(screenOf(r));
  else if (took && R === r && ui.screen === shown && !ui.sheet) { // the board from the sheet, drawn: no chip of the old one stays
    ui.keepScroll = true;
    render();
  }
  return poll(true);
}

function screenOf(r) {
  if (r.removed) return 'removed';
  if (!r.brev) return 'saving';
  if (r.mode === 'oc' && !r.started) return 'lobby';
  return 'progress';
}

async function openFromLink(c) {
  const hit = list.find((x) => x.e === c.e);
  if (hit) return openEval(hit);
  const r = { e: c.e, k: c.k, w: c.w, d: c.d, pub: c.pub, name: 'Peer evaluation', created: new Date().toISOString(), sq: 4, mode: 'oc', brev: 0, started: false, page2: true, shared: false, used: false, removed: false, fixed: {}, out: {}, ignored: {}, made: true, board: null };
  ui.kept = false;
  R = r;
  X = fresh();
  go('saving');
  let g;
  try {
    g = await getBoard(c.e);
  } catch (err) {
    ui.saveProblem = 'link';
    return render();
  }
  if (!g.ok) {
    if (g.error === 'removed') {
      r.board = { v: 2, e: c.e, n: 'Peer evaluation', at: new Date(g.created || Date.now()).toISOString(), pub: c.pub, seed: 0, sq: 4, mode: 'oc', started: false, oc: [] };
      r.created = r.board.at;
      return markRemoved(r);
    }
    ui.saveProblem = 'gone';
    return render();
  }
  try {
    r.board = await openBoard(c.k, c.e, g.board);
  } catch (err) {
    ui.saveProblem = 'gone';
    return render();
  }
  Object.assign(r, { name: r.board.n, created: new Date(g.created || r.board.at).toISOString(), sq: r.board.sq || 4, mode: r.board.mode, brev: g.brev, started: !!r.board.started, used: !!g.used });
  takeFixed(r, r.board);
  return openEval(r);
}

/* ---------- Start the evaluation (OCs chose their squads) ---------- */

/** Start's checks on the picks now. @returns {boolean} true when they hold it: a sheet asks first (a retry resends the board it froze) */
function held(force, retry) {
  const pk = picks();
  const red = R.board.oc.filter((o) => placeOf(o[0], pk) === -1);
  if (red.length) {
    if (retry) startSend(R, X); // a retry keeps the board it froze
    else openSheet('fix', { names: red.map((o) => fullName(o[1], o[2])) });
    return true;
  }
  const bare = (o) => !o[2] && !pk.firsts[o[0]];
  const twins = R.board.oc.filter((o) => bare(o) && placeOf(o[0], pk) > 0 && R.board.oc.some((y) => y[0] !== o[0] && y[3] === o[3] && bare(y) && lastKey(y[1]) === lastKey(o[1]) && placeOf(y[0], pk) === placeOf(o[0], pk))); // raters couldn't tell them apart
  if (twins.length) {
    if (retry) startSend(R, X);
    else openSheet('twins', { names: [...new Set(twins.map((o) => o[1]))] });
    return true;
  }
  const wait = R.board.oc.filter((o) => placeOf(o[0], pk) === null).map((o) => fullName(o[1], o[2])).concat(pk.extra.map((e) => fullName(e.last, e.first))); // typed names not added are not in either
  if (!wait.length || force) return false;
  openSheet('start', { names: wait });
  return true;
}

/** A Start that never went in: the sign-up goes on, without the freeze's squads (the cadre's own placements stay), so no phone is held to them. */
function thaw(r) {
  r.board.started = false;
  delete r.board.st;
  r.board.oc.forEach((o) => {
    if (!(r.fixed && r.fixed[o[0]])) o[4] = 0;
  });
}

async function startNow(force, retry) {
  if (starting.has(R.e)) return toast('Starting the evaluation… one moment.');
  if (!X.loaded) { // no picks yet: everyone would look not in. They're read now, and Start goes on once they are
    const x = X;
    toast(x.err ? 'Can\'t reach the sheet right now. Trying again…' : 'Loading the squads… one moment.');
    waits.add(x);
    try {
      await poll(true);
    } finally {
      waits.delete(x);
    }
    if (X !== x || !x.loaded) return X === x && x.err ? toast('Can\'t reach the sheet right now. Tap Start again in a minute.') : undefined;
    return startNow(force, retry);
  }
  const r = R;
  const x = X;
  if (held(force, retry)) return undefined;
  if (!retry || (ui.screen === 'lobby' && !ui.sheet)) { // a retry on another screen, or under a sheet the cadre has open, goes on quietly
    closeSheet();
    toast('Starting the evaluation…');
  }
  starting.add(r.e);
  const done = await adoptStarted(r, x); // started already (on another device, or by a Start whose answer was lost): that board stays
  starting.delete(r.e);
  if (!done && retry && R !== r) return startSend(r, x); // a retry the cadre left for another evaluation: its frozen board goes on trying, as the retry's timer does
  if (!done && retry && X !== x) return notStarted(r); // opened again meanwhile: that Start is over, and the page says so
  if (done || R !== r || X !== x) return undefined; // or the cadre left this evaluation meanwhile, even to come back: this Start is over
  if (held(force, retry)) return undefined; // a pick that came in during the read gets the same checks: the freeze uses these picks
  const pk = picks();
  const nd = r.board.nd || {};
  r.board.oc = r.board.oc.map((o) => {
    const p = placeOf(o[0], pk);
    const first = pk.firsts[o[0]] && nd[o[0]] && pk.fdv[o[0]] === nd[o[0]] ? pk.firsts[o[0]] : o[2] || pk.firsts[o[0]] || ''; // the phone that named a twin row may have fixed its spelling since
    return [o[0], o[1], first, o[3], p > 0 ? p : 0].concat(o[5] ? [o[5]] : []); // an added newcomer keeps their phone's id: it finds its row by it
  });
  r.board.started = true;
  r.board.st = startMark(); // this freeze's own mark: once it is in, the sheet takes no other board
  waits.add(x);
  await boardChanged();
  return startSend(r, x);
}

/** A Start that is over without going in (the cadre opened the evaluation again meanwhile): the sign-up goes on at once, not frozen until the next read. */
function notStarted(r) {
  if (r.board.started && !r.started) {
    thaw(r);
    persist();
    if (R === r && ui.screen === 'lobby' && !ui.sheet) render();
  }
  return toast('Not started. Tap Start the evaluation again.');
}

/** The frozen board, then start: for r only, even if the cadre opens another evaluation meanwhile. */
async function startSend(r, x) {
  clearTimeout(x.startTimer);
  waits.delete(x);
  starting.add(r.e); // no squad moves and no second Start while the frozen board is on its way
  for (let i = 0; i < 6 && !r.started && !r.removed; i++) {
    await sendBoard(r, x); // started only once the sheet holds a started board: this one went in (begun), or one is there (takeStarted)
    if (!r.started) await new Promise((ok) => setTimeout(ok, 2500));
  }
  starting.delete(r.e);
  if (r.removed) return undefined;
  if (!r.started) { // a retry still on the sign-up reads the sheet first, then freezes again from the latest picks
    clearTimeout(x.boardTimer); // only this retry sends it now, never the first frozen board on its own
    if (R === r && X !== x) return notStarted(r); // the cadre opened it again meanwhile: that Start is over, and nothing retries it
    if (R === r && ui.screen === 'lobby') toast('Can\'t start it yet. It tries again by itself.');
    waits.add(x);
    x.startTimer = setTimeout(() => {
      waits.delete(x);
      if (R === r && X !== x) return r.started ? undefined : notStarted(r); // the cadre opened it again since: that Start is over
      return R === r ? startNow(true, true) : startSend(r, x); // on any screen, freeze again from the latest picks
    }, 20000);
    return undefined;
  }
  let res = null;
  try {
    res = await startEval(r.e, await wt(r));
  } catch (err) {
    res = null;
  }
  if (res && res.error === 'removed') return markRemoved(r);
  if (!res || !res.ok) { // quietly: only the sheet's own flag is left
    waits.add(x);
    x.startTimer = setTimeout(() => startSend(r, x), 20000);
  }
  return undefined;
}

/* ---------- Again with the same OCs and squads ---------- */

function againFrom(r) {
  const oc = r.board.oc.filter((o) => (r.started ? o[4] : o[4] || r.mode === 'oc'));
  return { name: r.name, sq: Math.max(r.sq || 4, ...r.board.oc.map((o) => o[4] || 0)), from: r.e, oc: r.board.oc.map((o) => o.slice()), out: r.board.oc.filter((o) => r.started && !o[4]).map((o) => o[0]), keep: oc.map((o) => o[0]) };
}

/** Where an OC is on the Again board: a squad, 0 = left out, null = needs a squad (none yet, or above the squads per platoon). */
const placeAgain = (o) => (ui.again.out.indexOf(o[0]) >= 0 ? 0 : o[4] > 0 && o[4] <= ui.again.sq ? o[4] : null);

function makeAgain() {
  const a = ui.again;
  if (a.oc.some((o) => placeAgain(o) === null)) return toast('Put every OC in a squad, or leave them out, first.');
  const twin = twinIn(a.oc.filter((o) => placeAgain(o) > 0).map((o) => ({ last: o[1], first: o[2], plt: o[3], sqd: o[4] })));
  if (twin) return toast('Two OCs named ' + twin.last + ' with no first names are in one platoon. Tap one and fix the name to add a first name.');
  const oc = a.oc.filter((o) => placeAgain(o) > 0).map((o, i) => [i, o[1], o[2], o[3], o[4]]);
  if (!oc.length) return toast('Put at least two OCs in a squad first.');
  return makeCode(oc, 'set', a.name, a.sq, a.from);
}

/* ---------- printing ---------- */

function printPages(only2) {
  const r = R;
  const p1 = '<div class="page p1"><span class="kicker">' + esc(peName(titleOf(r), ' · ')) + '</span><h1>Scan with your phone camera</h1>' +
    '<div class="code">' + qrSvg(ocUrl(r)) + '</div><p class="then">Then tap your name.</p>' +
    '<div class="lines"><p>About 25 minutes. It saves as you go.</p><p>Only cadre see your answers.</p><p>No signal? Open the link again when you have signal.</p></div>' +
    '<p class="foot1">Open it in Safari or Chrome, not inside GroupMe or Signal.</p></div>';
  const p2 = '<div class="page p2"><div class="band">CADRE ONLY · DO NOT POST</div><h2>Your results for ' + esc(titleOf(r)) + '</h2>' +
    '<div class="two"><div class="code">' + qrSvg(cadreUrl(r)) + '</div><div class="words"><b>This code opens the results on any phone or computer.</b>' +
    '<span>Keep this page. Don\'t post it, and don\'t send it to the platoon.</span><span>If you lose this page and your browser is cleared, the results are gone.</span></div></div>' +
    '<div class="pdf">Saved this page as a PDF? Click to open the results: <a href="' + esc(cadreUrl(r)) + '">Open the results</a></div>' +
    '<p>The page to post is page 1: one big code that says Scan with your phone camera.</p><div class="band2">CADRE ONLY</div></div>';
  printEl.innerHTML = only2 ? p2 : p1 + p2;
  window.print();
  r.page2 = true;
  persist();
}

/**
 * Print feedback for everyone (17j, step 3 of the paper process: "给每一个人打印出来"): a page or two per OC, by squad, with
 * what their Excel feedback holds (strengths, to work on, by category against the squad, Yes/No counts, every comment,
 * no names). From the print dialog, Save as PDF or print. Ranks and the order of merit are not in it.
 */
function printFeedback() {
  const pages = feedbackPages(X.E, new Date());
  if (!pages.length) return toast('No feedback yet: no one has been rated.');
  const one = (v) => (typeof v === 'number' ? (Math.round(v * 10) / 10).toFixed(1) : '–');
  const w = (v) => (typeof v === 'number' ? Math.round((v / 5) * 100) : 0);
  printEl.innerHTML = pages.map((pg) => {
    const named = (list) => list.map((ai) => esc(ATTRIBUTES[ai].name) + ' (' + one(pg.r.attrMean[ai]) + ')').join(' · ') || '–';
    const cats = CATEGORIES.map((c, i) => '<tr><td>' + esc(c) + '</td><td class="num">' + one(pg.r.cat[c]) + '</td><td class="num">' + one(pg.squadCat[i]) + '</td><td class="barc"><i style="width:' + w(pg.r.cat[c]) + '%"></i><b style="left:' + w(pg.squadCat[i]) + '%"></b></td></tr>').join('');
    const asks = pg.ask.map((x) => '<tr><td>' + esc(QUESTION_LABEL[x.q.id]) + '</td><td class="num">' + x.yes + '</td><td class="num">' + x.no + '</td></tr>').join('');
    const said = pg.said.map(({ q, items }) => (items.length ? '<h4>' + esc(QUESTION_LABEL[q.id]) + '</h4><ul>' + items.map((x) => '<li><b>' + (x.a === 'Y' ? 'Yes' : x.a === 'N' ? 'No' : '') + '</b> ' + esc(x.t) + '</li>').join('') + '</ul>' : '')).join('');
    const notes = pg.notes.length ? '<h4>On your attributes</h4><ul>' + pg.notes.map((t) => '<li>' + esc(t) + '</li>').join('') + '</ul>' : '';
    return '<div class="fb"><div class="fbhead"><span>' + esc(peName(titleOf(R), ' · ')) + ' · ' + esc(pg.sq.label) + ' · ' + esc(pg.day) + '</span><span>For this OC only</span></div>' +
      '<h1>' + esc(fullName(pg.m.last, pg.m.first)) + '</h1><p class="from">Peer feedback from ' + pg.from + (pg.from === 1 ? ' squadmate' : ' squadmates') + '. No names; the comments are in mixed order.</p>' +
      '<div class="sw"><div><b>Strengths</b>' + named(pg.top) + '</div><div><b>To work on</b>' + named(pg.low) + '</div></div>' +
      '<h3>By category, 0 to 5 <small>(bar: you · red line: your squad)</small></h3><table class="cat"><tr><th></th><th class="num">You</th><th class="num">Squad</th><th></th></tr>' + cats + '</table>' +
      '<h3>Peer questions</h3><table class="ask"><tr><th></th><th class="num">Yes</th><th class="num">No</th></tr>' + asks + '</table>' +
      '<h3>What your squadmates said</h3>' + said + notes + '</div>';
  }).join('');
  window.print();
  return toast('Printing ' + plural(pages.length, 'page set', 'page sets') + ', one per OC. Choose Save as PDF to keep a copy.');
}

function afterCode() {
  return R.mode === 'oc' && !R.started ? 'lobby' : 'progress';
}

async function shareLink() {
  const r = R;
  const url = ocUrl(r);
  if (navigator.share && device() === 'phone') {
    try {
      await navigator.share({ title: 'Peer evaluation', text: peName(r.name, ', ') + '. Open it and tap your name:', url });
    } catch (err) {
      return undefined; // the cadre closed the share sheet
    }
  } else if (await copy(url)) toast('Copied the link for OCs.');
  else return toast('Copy didn\'t work here. Print the 2 pages instead.');
  r.shared = true;
  persist();
  clearTimeout(timers.after);
  timers.after = setTimeout(() => ui.screen === 'qr' && go(afterCode()), 1200); // not if the cadre has moved on
  return undefined;
}

/* ---------- screens ---------- */

function head(back, label) {
  const b = back ? '<button class="hbtn" data-act="back" data-to="' + esc(back) + '">' + ICON_BACK + '<span>' + esc(label || 'Back') + '</span></button>' : '<span class="spacer"></span>';
  const more = ui.screen === 'more' ? '<button class="hbtn" data-act="close-more">Close</button>' : '<button class="hbtn" data-act="more">More</button>';
  return '<div class="head">' + b + more + '</div>';
}

const foot = (inner) => '<div class="foot">' + inner + '</div>';
const primary = (act, label, extra) => '<button class="primary" data-act="' + act + '"' + (extra || '') + '>' + esc(label) + '</button>';
const links = (...items) => '<div class="links">' + items.filter(Boolean).join('') + '</div>'; // a footer's secondary taps share a line when they fit (a shorter footer on a phone)
const link = (act, label, extra) => '<button class="link" data-act="' + act + '"' + (extra || '') + '>' + esc(label) + '</button>';
const keptLine = () => (!ui.kept && R && !own ? '<p class="notice">This ' + device() + ' won\'t keep this evaluation. ' + link('keep-here', 'Keep it here') + '</p>' : '');

function sHome() {
  const step = (n, t) => '<li><span class="n">' + n + '</span><span>' + t + '</span></li>';
  return head() + '<div class="main"><div class="hero"><span class="mark" aria-hidden="true">' + ICON_PEOPLE + '</span><h1 class="big">Peer evaluation</h1></div>' + // 30 words at most on this first screen
    '<ol class="steps">' + step(1, 'Add your roster') + step(2, 'OCs scan one QR code and rate') + step(3, 'Get the rankings and Excel files') + '</ol></div>' +
    foot(primary('start', 'Start a peer evaluation') + link('open', 'Open one you started'));
}

function sOpen() {
  return head('home') + '<div class="main"><h1>Open one you started</h1>' +
    '<p class="body">Scan the code on page 2 of your printout with your phone camera. Saved page 2 as a PDF? Click the link on it.</p></div>';
}

function sAll() {
  const rows = list.map((r, i) => {
    const st = r.removed ? 'Removed: no one used it by ' + dayLabel(unusedUntil(r.created)) : !r.started && r.mode === 'oc' ? 'Squad sign-up' : 'Evaluation open';
    return '<button class="option" data-act="open-eval" data-i="' + i + '"><span><b>' + esc(titleOf(r)) + '</b><span class="sub">' + esc(st) + '</span></span></button>';
  }).join('');
  const again = list.find((r) => r.started && !r.removed);
  return head() + '<div class="main"><h1 class="big">Peer evaluation</h1>' + primary('start', 'Start a peer evaluation') +
    (rows ? '<span class="sub">On this ' + device() + '</span>' + rows : '') +
    (again ? link('again', 'Again with the same OCs and squads', ' data-e="' + esc(again.e) + '"') : '') + '</div>';
}

function sRoster() {
  const drop = device() === 'phone' ? '<span><b>Your roster file</b></span><span class="sub">Excel or CSV, from email or Files</span>' : '<span><b>Drop your roster file here</b></span><span class="sub">Excel or CSV, straight from email or a folder</span>';
  return head(list.length ? 'all' : 'home') + '<div class="main"><h1>Who\'s in it?</h1>' +
    '<div class="drop" data-drop="roster" data-act="choose" role="button" tabindex="0">' + ICON_FILE + drop + '</div>' +
    '<p class="body">Names are enough. Platoons and squads too, if the file has them.</p>' +
    '<p class="body">Collecting the filled-in forms from Classroom? Drop them, or their zip, here: the names come from the forms, and their answers go in.</p>' +
    (ui.paste ? pasteBox() : '') + '</div>' +
    foot(primary('choose', 'Choose file') + link('paste', 'Paste from Excel'));
}

function pasteBox() {
  return '<label class="sub" for="paste">Paste the names here</label><textarea id="paste" class="input" spellcheck="false"></textarea>' + link('paste-add', 'Add these names');
}

function sRead() {
  const D = ui.draft;
  const dr = draftRoster();
  const files = D.sources.map((s) => '<div class="file">' + ICON_OK + '<span>' + esc(s.name) + (s.tab ? ' (' + esc(s.tab) + ')' : '') + '</span></div>').join('');
  const n = dr.ocs.length;
  const sqds = [...new Set(dr.ocs.map((o) => o.sqd).filter(Boolean))];
  const onePlt = dr.plts.length === 1 ? dr.plts[0] : 0; // the count is the whole company, whatever tab is on
  const pltsLine = dr.namesOnly ? '' : onePlt ? ord(onePlt) + ' PLT' + (dr.allSquads && sqds.length === 1 ? ' · ' + ord(sqds[0]) + ' SQD' : '') : plural(dr.plts.length, 'platoon', 'platoons');
  const count = plural(n, 'OC', 'OCs') + (pltsLine ? ' · ' + pltsLine : '');
  const line = dr.allSquads ? (D.sources.some((s) => s.sqd) ? 'The platoon and squad come from the file name.' : 'Squads are in the file.') + ' Check the names now: the evaluation opens with the QR code.' : 'Squads aren\'t in the file. Next, tap each OC\'s squad.';
  let h = '<h1>Who\'s in it?</h1>' + files + '<p class="count">' + esc(count) + '</p>' + (n ? '<p class="body">' + esc(line) + '</p>' : '');
  if (dr.plts.length > 1) {
    h += '<span class="sub">Show</span><div class="seg" style="--n:' + (dr.plts.length + 1) + '">' +
      [0].concat(dr.plts).map((p) => '<button data-act="plt-sel" data-v="' + p + '" aria-pressed="' + (D.pltSel === p) + '">' + (p ? ord(p) : 'All') + '</button>').join('') + '</div>';
  }
  if (dr.namesOnly) {
    h += '<span class="sub">Which platoon is this?</span><div class="seg">' + [1, 2, 3, 4].map((p) => '<button data-act="plt-all" data-v="' + p + '" aria-pressed="' + (D.pltAll === p) + '">' + ord(p) + '</button>').join('') + '</div>';
  } else if (dr.noPlt.length) {
    h += '<div class="q"><b>' + plural(dr.noPlt.length, 'OC has', 'OCs have') + ' no clear platoon. Tap one for each.</b>' +
      dr.noPlt.map((o) => '<div><span>' + esc(fullName(o.last, o.first)) + '</span><div class="btns">' + (o.plts && o.plts.length ? o.plts : dr.plts.length ? dr.plts : [1, 2, 3, 4]).map((p) =>
        '<button data-act="oc-plt" data-key="' + esc(o.key) + '" data-v="' + p + '">' + ord(p) + '</button>').join('') + '</div></div>').join('') + '</div>';
  }
  dr.questions.forEach((q) => {
    if (q.kind === 'same') {
      h += '<div class="q"><span>Same person? <b>' + esc(q.a.name) + '</b> (' + esc(q.a.src) + ') and <b>' + esc(q.b.name) + '</b> (' + esc(q.b.src) + ')</span><div class="btns">' +
        '<button data-act="q-same" data-q="' + esc(q.id) + '" data-v="same">Same person</button><button data-act="q-same" data-q="' + esc(q.id) + '" data-v="two">Two people</button></div></div>';
    } else if (q.kind === 'which') {
      h += '<div class="q"><span>Which <b>' + esc(q.name) + '</b> is it?</span><div class="btns">' + q.options.map((o) => '<button data-act="q-which" data-q="' + esc(q.id) + '" data-v="' + esc(o.key) + '">' + esc(o.name) + '</button>').join('') +
        '<button data-act="q-which" data-q="' + esc(q.id) + '" data-v="none" data-not="' + esc(q.options.map((o) => o.key).join(' ')) + '">Someone else</button></div></div>';
    } else {
      h += '<div class="q"><span>Only in ' + esc(q.src) + ': <b>' + esc(q.name) + '</b>. Add them?</span><div class="btns">' +
        '<button data-act="q-add" data-q="' + esc(q.id) + '" data-v="add">Add</button><button data-act="q-add" data-q="' + esc(q.id) + '" data-v="ignore">Leave out</button></div></div>';
    }
  });
  if (dr.block.length) h += '<p class="notice">' + esc(dr.block[0]) + (dr.block.length > 1 ? ' (and ' + (dr.block.length - 1) + ' more)' : '') + '</p>';
  h += staleLine(dr) + nameList(dr);
  if (ui.paste) h += pasteBox();
  return head('roster') + '<div class="main">' + h + '</div>' + foot(primary('read-next', 'Next') + links(link('choose', 'Add another file'), n ? link('roster-xlsx', 'Save the roster (Excel)') : '', n && dr.allSquads ? link('sheets-draft', 'Blank sheets for OCs (Excel)') : ''));
}

/** Who's in it?: Find a name, Add an OC (an insert) and the ones left out on top, then every name, to fix or leave out (a recycle). */
function nameList(dr) {
  const D = ui.draft;
  const q = String(ui.find || '').trim().toLowerCase();
  const many = !D.pltSel && dr.plts.length > 1;
  const ocs = dr.shown.slice().sort((a, b) => (a.plt || 0) - (b.plt || 0) || a.last.localeCompare(b.last) || (a.first || '').localeCompare(b.first || ''));
  const hit = (o) => !q || fullName(o.last, o.first).toLowerCase().indexOf(q) >= 0;
  let h = '<div class="names"><p class="sub">Tap a name to fix it or leave it out.</p>' +
    (ocs.length > 12 || q ? '<input id="ofind" class="input" type="search" placeholder="Find a name" aria-label="Find a name" autocomplete="off" autocorrect="off" spellcheck="false" enterkeyhint="search" value="' + esc(ui.find || '') + '">' : '') +
    '<button class="link left" data-act="oc-new">+ Add an OC</button>';
  if (dr.outs.length) {
    h += '<span class="sub plt">Left out · ' + dr.outs.length + '</span>' +
      dr.outs.map((o) => '<button class="nrow out" data-act="oc-back" data-key="' + esc(o.key) + '"><span>' + esc(fullName(o.last, o.first)) + '</span><span class="tag">Put back</span></button>').join('');
  }
  let plt = -1;
  ocs.forEach((o) => {
    if (many && (o.plt || 0) !== plt) {
      plt = o.plt || 0;
      h += '<span class="sub plt" data-plt' + (q ? ' hidden' : '') + '>' + (plt ? ord(plt) + ' PLT' : 'No platoon yet') + ' · ' + ocs.filter((x) => (x.plt || 0) === plt).length + '</span>';
    }
    const name = fullName(o.last, o.first);
    const tag = o.srcs[0] === 'j' ? 'Added' : o.renamed !== undefined || o.keys.some((k) => D.touched[k]) ? 'Fixed' : '';
    h += '<button class="nrow" data-act="oc-edit" data-key="' + esc(o.key) + '" data-name="' + esc(name.toLowerCase()) + '"' + (hit(o) ? '' : ' hidden') + '><span>' + esc(name) + '</span><span class="right">' +
      (dr.someSquads && o.sqd ? '<span class="sub">' + ord(o.sqd) + ' SQD</span>' : '') + (tag ? '<span class="tag">' + tag + '</span>' : '') + '</span></button>';
  });
  return h + '<p class="sub" id="onone"' + (q && !ocs.some(hit) ? '' : ' hidden') + '>No name like that. Tap Add an OC to add it.</p></div>';
}

function sSet() {
  const D = ui.draft;
  const dr = draftRoster();
  const plts = [...new Set(dr.ocs.map((o) => o.plt || 0))].sort((a, b) => a - b);
  const tab = plts.indexOf(ui.tab) >= 0 ? ui.tab : plts[0];
  const tabs = plts.length > 1 ? '<div class="tabs">' + plts.map((p) => {
    const inP = dr.ocs.filter((o) => (o.plt || 0) === p);
    return '<button data-act="tab" data-v="' + p + '" aria-pressed="' + (p === tab) + '">' + (p ? ord(p) : 'All') + ' · ' + inP.filter((o) => o.sqd || setOf(o)).length + '/' + inP.length + '</button>';
  }).join('') + '</div>' : '';
  const rows = dr.ocs.filter((o) => (o.plt || 0) === tab).sort((a, b) => a.last.localeCompare(b.last)).map((o) => '<div class="row"><b>' + esc(fullName(o.last, o.first)) + '</b><div class="seg" style="--n:' + D.sq + '">' +
    Array.from({ length: D.sq }, (x, i) => i + 1).map((q) => '<button data-act="set-sq" data-key="' + esc(o.key) + '" data-v="' + q + '" aria-pressed="' + ((setOf(o) || o.sqd) === q) + '">' + ord(q) + '</button>').join('') + '</div></div>').join('');
  return head('read') + '<div class="main"><h1>Tap each OC\'s squad</h1>' + staleLine(dr) + tabs + rows + '</div>' + foot(primary('set-next', 'Next') + links(link('roster-xlsx', 'Save the roster (Excel)'), link('sheets-draft', 'Blank sheets for OCs (Excel)')));
}

function sOwn() {
  const d = device();
  return head(ui.from || 'read') + '<div class="main center"><h1>Is this your own ' + d + '?</h1>' +
    '<p class="body">If yes, this ' + d + ' keeps the evaluation, so it opens right here next time. If no, nothing stays on it, and page 2 of your printout brings it back.</p></div>' +
    foot(primary('own-yes', 'Yes, it\'s mine') + '<button class="second" data-act="own-no">No, it\'s shared</button>');
}

function sSaving() {
  const p = ui.saveProblem;
  const title = !p ? 'Making your QR code…' : p === 'limit' ? 'Can\'t make a new evaluation right now' : p === 'link' ? 'Can\'t open it yet' : p === 'gone' ? 'This QR code has expired' : 'Can\'t save it yet';
  const line = !p ? '' : p === 'limit' ? 'Too many new evaluations on this site today. It tries again by itself every few minutes, and your roster stays here.'
    : p === 'link' ? 'No answer from the sheet. Open this link again when you have signal.'
      : p === 'gone' ? 'This evaluation isn\'t on the sheet any more.'
        : 'It keeps trying by itself and shows the QR code here as soon as it\'s saved. Nothing for you to do.';
  return head() + '<div class="main center" style="text-align:center">' + (p === 'gone' ? '' : '<div class="spin" aria-hidden="true"></div>') + '<h1>' + esc(title) + '</h1>' + (line ? '<p class="body">' + esc(line) + '</p>' : '') +
    (p === 'gone' ? primary('start', 'Start a peer evaluation') : '') + '</div>';
}

function sQr() {
  const r = R;
  const d = device();
  const n = r.board.oc.filter((o) => !r.started || o[4]).length;
  const shared = !own;
  const need = !shared && r.shared && !r.page2;
  const box = need
    ? '<div class="box warm"><b>Save page 2 too.</b><span class="body">It\'s your way back if this ' + d + ' is cleared.</span>' + link('save2', 'Save page 2', ' style="align-self:flex-start;padding:0"') + '</div>'
    : '<div class="box"><b>Page 2 is only for you.</b><span class="body">' + (shared ? 'This ' + d + ' won\'t keep it. Print page 2 now: its code is the only way back to the results.'
      : 'Its code opens the results on any phone or computer. If you lose it and this ' + d + ' is cleared, the results are gone.') + '</span></div>';
  const expiry = !r.used ? '<p class="sub">If no OC answers by ' + esc(dayLabel(unusedUntil(r.created))) + ', this QR code stops working.</p>' : '';
  return head('all', 'Home') + keptLine() + '<div class="main"><div><h1>Give this to your OCs</h1><span class="sub">' + esc(plural(n, 'OC', 'OCs')) + ' · one QR code for all</span></div>' +
    '<div class="qr" role="img" aria-label="QR code for the OCs">' + qrSvg(ocUrl(r)) + '</div>' + expiry + box + '</div>' +
    foot(primary('print', 'Print 2 pages') + links(link('share', d === 'phone' && navigator.share ? 'Share link' : 'Copy link'), link('sheets', 'Blank sheets for OCs (Excel)')));
}

function remindLine() {
  return own && R.shared && !R.page2 ? '<div class="notice slim">Save page 2 too: it\'s your way back if this ' + device() + ' is cleared. ' + link('save2', 'Save page 2') + '</div>' : '';
}

/** The squad board: the sign-up (lobby), or Change who's in it before Again (edit). */
function sBoard(edit) {
  const pk = edit ? null : picks();
  const oc = edit ? ui.again.oc : R.board.oc;
  const place = (o) => (edit ? placeAgain(o) : placeOf(o[0], pk));
  const plts = [...new Set(oc.map((o) => o[3] || 0))].sort((a, b) => a - b);
  const tab = plts.indexOf(ui.tab) >= 0 ? ui.tab : plts[0];
  const placed = oc.filter((o) => place(o) > 0).length;
  const tabs = plts.length > 1 ? '<div class="tabs">' + plts.map((p) => {
    const inP = oc.filter((o) => (o[3] || 0) === p);
    return '<button data-act="tab" data-v="' + p + '" aria-pressed="' + (p === tab) + '">' + (p ? ord(p) : 'All') + ' · ' + inP.filter((o) => place(o) > 0).length + '/' + inP.length + '</button>';
  }).join('') + '</div>' : '';
  const inP = oc.filter((o) => (o[3] || 0) === tab).sort((a, b) => a[1].localeCompare(b[1]));
  const chip = (o, kind) => '<button class="chip' + (kind ? ' ' + kind : '') + '" data-act="move" data-id="' + o[0] + '">' + esc(fullName(o[1], o[2] || (pk && pk.firsts[o[0]]) || '')) + (kind === 'red' ? ' · 2 phones' : '') + '</button>';
  const sq = edit ? ui.again.sq : R.sq || 4;
  let groups = '';
  for (let q = 1; q <= sq; q++) {
    const m = inP.filter((o) => place(o) === q);
    groups += '<div class="group"><b>' + ord(q) + ' SQD · ' + m.length + '</b><div class="chips">' + m.map((o) => chip(o, '')).join('') + '</div></div>';
  }
  const red = inP.filter((o) => place(o) === -1);
  const wait = inP.filter((o) => place(o) === null);
  const out = inP.filter((o) => place(o) === 0);
  groups = (red.length ? '<div class="group"><b class="err">Two phones disagree · ' + red.length + ' · tap to choose</b><div class="chips">' + red.map((o) => chip(o, 'red')).join('') + '</div></div>' : '') + // what needs the cadre first, then the squads
    (pk && pk.extra.length ? '<div class="group"><b>Not on the list · ' + pk.extra.length + '</b><div class="chips">' + pk.extra.map((x) => '<button class="chip wait" data-act="extra" data-dv="' + esc(x.dv) + '">' + esc(fullName(x.last, x.first)) + '</button>').join('') + '</div></div>' : '') +
    (wait.length ? '<div class="group"><b>Not in yet · ' + wait.length + '</b><div class="chips">' + wait.map((o) => chip(o, 'wait')).join('') + '</div></div>' : '') + groups +
    (out.length ? '<div class="group"><b>' + (edit ? 'Not in this one' : 'Left out') + ' · ' + out.length + '</b><div class="chips">' + out.map((o) => chip(o, 'wait')).join('') + '</div></div>' : '');
  const anyRed = oc.some((o) => place(o) === -1);
  const redHint = anyRed && !red.length; // red names in another platoon's tab: this one's group can't say it
  const hint = edit ? 'Tap a name to move it, fix it, add it or leave it out.' : anyRed ? (redHint ? 'A name in red: two phones picked different squads. Tap it to choose.' : '') : oc.some((o) => place(o) === null) || (pk && pk.extra.length) ? 'Tap a name to move it.' : 'Everyone\'s in. Start when you\'re ready.';
  const title = edit ? 'Who\'s in it' : 'Squads';
  const err = !edit && X.err ? '<p class="notice">Can\'t reach the sheet right now, so new sign-ups may be missing. It tries again by itself.</p>' : '';
  return head(edit ? 'again' : 'all', edit ? 'Back' : 'Home') + keptLine() + '<div class="main"><div style="display:flex;align-items:baseline;justify-content:space-between;gap:12px"><h1>' + title + '</h1><span class="count" style="color:#3F5F33">' + placed + ' of ' + oc.length + ' in</span></div>' +
    (edit ? '' : '<p class="sub">' + esc(titleOf(R)) + '</p>' + remindLine()) + err + tabs + groups + '</div>' + foot('<p class="' + (redHint ? 'err' : 'sub') + '" role="status">' + esc(hint) + '</p>' +
      (edit ? primary('make-again', 'Make the QR code') + link('oc-new', '+ Add an OC') : primary('start-eval', 'Start the evaluation') + link('show-qr', 'Show the QR code')));
}

function sProgress() {
  const c = counts();
  const rows = (X.E ? X.E.squads : []).map((sq) => {
    const left = sq.ids.filter((id) => !X.results[id]).map((id) => X.E.members.find((m) => m.id === id).name).sort();
    return { sq, left, html: '<div class="row"><div class="line"><span>' + esc(sq.label) + '</span><span class="' + (left.length ? '' : 'ok') + '">' + (sq.ids.length - left.length) + ' of ' + sq.ids.length + (left.length ? '' : ' ✓') + '</span></div>' +
      (left.length ? '<span class="sub">Not yet: ' + esc(left.join(', ')) + '</span>' : '') + '</div>' };
  }).sort((a, b) => (b.left.length > 0) - (a.left.length > 0)).map((x) => x.html).join('');
  const outs = R.board.oc.filter((o) => R.started && !o[4]).map((o) => fullName(o[1], o[2]));
  const status = X.err ? '<p class="notice">Can\'t get new answers right now. It tries again by itself in 2 minutes.</p>' : '<p class="sub">Updates by itself.</p>';
  return head('all', 'Home') + keptLine() + '<div class="main">' + remindLine() + '<p class="huge">' + c.done + ' of ' + c.total + ' done</p>' +
    '<div class="pbar" role="progressbar" aria-label="Done" aria-valuemin="0" aria-valuemax="' + c.total + '" aria-valuenow="' + c.done + '"><i style="width:' + (c.total ? Math.round((100 * c.done) / c.total) : 0) + '%"></i></div>' + status + (X.E ? lowNotice(X.E.squads) : '') + rows +
    (outs.length ? '<p class="note">Not in this evaluation: ' + esc(listOf(outs)) + '. Their phones tell them to see you.</p>' : '') + '</div>' +
    foot(primary('copy-names', 'Copy names not done') + links(link('results', 'Results so far'), link('import', 'Add filled forms (Excel)')));
}

/** Plain notes from the scoring flags, no technical words, no rater names. */
function sResults() {
  const sqs = X.E ? X.E.squads.filter((sq) => sq.ids.length > 1) : [];
  const plts = [...new Set(sqs.map((sq) => sq.plt))].sort((a, b) => a - b);
  const tab = plts.indexOf(ui.resTab) >= 0 ? ui.resTab : plts[0];
  const tabs = plts.length > 1 ? '<div class="tabs">' + plts.map((p) => '<button data-act="res-tab" data-v="' + p + '" aria-pressed="' + (p === tab) + '">' + ord(p) + ' PLT</button>').join('') + '</div>' : '';
  const body = sqs.filter((sq) => sq.plt === tab).map((sq) => {
    const res = sq.ids.map((id) => X.results[id]).filter(Boolean).map((r) => r.payload);
    const { rows, raterChecks } = scoreSquad(sq.ids, res);
    const name = (id) => X.E.members.find((m) => m.id === id).name;
    return '<div class="group"><div class="ghead"><b>' + esc(ord(sq.sqd)) + ' SQD · ' + res.length + ' of ' + sq.ids.length + ' in</b><span>Overall</span></div><div>' + rows.map((r) => {
      const note = rowNotes(r, res).join(' · ');
      const marks = formChecks(X.results[r.id], raterChecks, r.id, res).join(' · '); // how this OC filled in his own form
      return '<div class="rank"><span class="n">' + (r.merit || '–') + '</span><span class="who">' + esc(name(r.id)) + (note ? '<span class="flag">' + esc(note) + '</span>' : '') + (marks ? '<span class="mark">' + esc(marks) + '</span>' : '') + '</span><span class="score">' + (r.overall === null ? '' : r.overall.toFixed(1)) + '</span></div>';
    }).join('') + '</div></div>';
  }).join('');
  return head('all', 'Home') + keptLine() + '<div class="main"><h1>Results</h1>' + tabs + lowNotice(sqs.filter((sq) => sq.plt === tab)) + (body || '<p class="sub">No results in yet.</p>') + '</div>' +
    foot(primary('download', 'Download Excel') + links(link('print-feedback', 'Print feedback for everyone (PDF)'), link('feedback', 'Feedback for each OC (Excel)'), link('import', 'Add filled forms (Excel)'), link('again', 'Again with the same OCs and squads', ' data-e="' + esc(R.e) + '"')));
}

/** A rater who scored the squad far lower than the others did (a point or more on average): named, for the cadre to look closer. Nothing changes. */
function lowNotice(sqs) {
  const low = [];
  sqs.forEach((sq) => {
    const res = sq.ids.map((id) => X.results[id]).filter(Boolean).map((r) => r.payload);
    sq.ids.forEach((id) => {
      const g = lowScoreGap(res, id);
      if (g !== null && g <= -1) low.push(X.E.members.find((m) => m.id === id).name);
    });
  });
  return low.length ? '<p class="notice"><b>Look closer:</b> ' + esc(listOf(low)) + ' scored others far lower than the squad did. Nothing in the scores is changed.</p>' : '';
}

function sAgain() {
  const a = ui.again;
  const keep = a.oc.filter((o) => placeAgain(o) > 0);
  const need = a.oc.filter((o) => placeAgain(o) === null).map((o) => fullName(o[1], o[2]));
  const squads = new Set(keep.map((o) => o[3] + '-' + o[4])).size;
  const outs = a.oc.filter((o) => a.out.indexOf(o[0]) >= 0).map((o) => fullName(o[1], o[2]));
  return head('all', 'Home') + '<div class="main"><h1>Again with the same OCs and squads</h1><p class="count">' + plural(keep.length, 'OC', 'OCs') + ' · ' + plural(squads, 'squad', 'squads') + '</p>' +
    '<p class="body">The same people and squads as ' + esc(a.title) + ', with a new QR code.</p>' +
    (outs.length ? '<p class="notice">' + plural(outs.length, 'was', 'were').replace(/^(\d+) /, '$1 ') + ' out last time: ' + esc(listOf(outs, 5)) + '. They stay out unless you add them.</p>' : '') +
    (need.length ? '<p class="notice">' + esc(plural(need.length, 'OC needs', 'OCs need') + ' a squad: ' + listOf(need, 5) + '. Tap Change who\'s in it.') + '</p>' : '') + '</div>' +
    foot(primary('make-again', 'Make the QR code') + link('edit-again', 'Change who\'s in it'));
}

function sRemoved() {
  // page 2 on a device that never kept it has no names to make a new code from
  return head('all', 'Home') + '<div class="main center"><h1>This QR code was removed</h1>' +
    '<p class="body">No one used this QR code by ' + esc(dayLabel(unusedUntil(R.created))) + ', so it was removed. Start a new one.</p></div>' +
    foot(R.board.oc.length ? primary('remake', 'Make a new QR code') + link('start', 'Start from a new roster') : primary('start', 'Start from a new roster'));
}

function sMore() {
  const items = [];
  const btn = (act, label) => items.push('<button data-act="' + act + '">' + esc(label) + '</button>');
  const d = device();
  const again = ui.moreFrom === 'again' || ui.moreFrom === 'edit';
  // only a screen that shows the evaluation acts on it: not Home, Open, a code still saving, or Again (a new one)
  if (R && R.brev && !R.removed && !again && ['home', 'all', 'open', 'saving'].indexOf(ui.moreFrom) < 0) {
    btn('show-qr', 'Show the QR code again');
    btn('copy-cadre', 'Copy my cadre link (only for you)');
    if (own && list.indexOf(R) >= 0) btn('forget', 'Forget this ' + d);
    else btn('keep-here', 'Keep it on this ' + d);
    btn('name', 'Name and squads per platoon');
    btn('import', 'Add results: a file, text or Excel form');
    if (R.started) btn('answers', 'Save answers so far (Excel)');
    btn('remove-sheet', 'Remove this evaluation from the sheet');
    items.push('<p class="info">' + (R.used ? 'Kept on the sheet until ' + esc(dayLabel(keptUntil(R.last || R.created))) + ', a year after the last answer. Download Excel keeps your own copy.' : 'If no OC answers by ' + esc(dayLabel(unusedUntil(R.created))) + ', the QR code stops working.') + '</p>');
  } else if (again) btn('name', 'Name and squads per platoon'); // for the new QR code
  items.push('<a href="manual.pdf" target="_blank" rel="noopener">Cadre manual (PDF)</a>'); // online only: not in sw.js
  items.push('<button data-act="admin">Use my own Google sheet, open an event file, links and QR slips for each OC (the admin console)</button>');
  items.push('<p class="info">For Army computers: the two one-file versions come from the person who keeps this site.</p>');
  return head() + '<div class="main"><h1>More</h1><div class="menu">' + items.join('') + '</div></div>';
}

/** What Name and squads per platoon changes: the new evaluation on Again, else the one on screen (or the draft). */
const nameTarget = () => (ui.moreFrom === 'again' || ui.moreFrom === 'edit' ? ui.again : R || ui.draft);

function sName() {
  const n = ui.nm; // what's typed and tapped here; only Save keeps it
  return head(ui.moreFrom || 'more') + '<div class="main"><h1>Name and squads per platoon</h1><label class="sub" for="ename">Name</label><input id="ename" class="input" maxlength="40" value="' + esc(n.name) + '">' +
    '<span class="sub">Squads per platoon</span><div class="seg">' + [3, 4, 5, 6].map((q) => '<button data-act="sq-count" data-v="' + q + '" aria-pressed="' + (n.sq === q) + '">' + q + '</button>').join('') + '</div></div>' +
    foot(primary('name-save', 'Save'));
}

function sImport() {
  return head('more') + '<div class="main"><h1>Add results</h1><p class="body">An OC without signal can send their answers as a file or text. Filled-in Excel forms work too: Blank sheets for OCs, or an OC\'s own Excel. Drop the files here, or paste the text.</p>' +
    '<div class="drop" data-drop="results" data-act="import-files" role="button" tabindex="0">' + ICON_FILE + '<span><b>Result files or Excel forms</b></span></div><textarea id="imp" class="input" spellcheck="false"></textarea></div>' +
    foot(primary('import-go', 'Add them') + link('import-files', 'Choose files'));
}

const SCREENS = { home: sHome, open: sOpen, all: sAll, roster: sRoster, read: sRead, set: sSet, own: sOwn, saving: sSaving, qr: sQr, lobby: () => sBoard(false), edit: () => sBoard(true), progress: sProgress, results: sResults, again: sAgain, removed: sRemoved, more: sMore, name: sName, import: sImport };

function render() {
  const fn = SCREENS[ui.screen] || sHome;
  const scroll = window.scrollY;
  root.innerHTML = fn();
  if (ui.sheet) renderSheet();
  document.title = ui.screen === 'home' || ui.screen === 'all' || !R ? 'Peer evaluation' : peName(R.name || '', ' · ');
  if (ui.keepScroll) window.scrollTo(0, scroll);
  ui.keepScroll = false;
}

function go(screen) {
  if (screen === 'lobby' && R && R.started) screen = 'progress'; // it started meanwhile
  // ui.from: the screen before this one (Back on Own); ui.moreFrom: the one More was opened from
  const sub = (s) => s === 'more' || s === 'name' || s === 'import';
  if (sub(screen)) {
    if (!sub(ui.screen)) ui.moreFrom = ui.screen;
  } else if (!sub(ui.screen)) ui.from = ui.screen;
  ui.screen = screen;
  drawn = true;
  if (screen === 'all' || screen === 'home' || screen === 'roster') dropLink();
  closeSheet(); // a sheet never stays over the next screen
  render();
  window.scrollTo(0, 0);
  reloadIfIdle();
}

/* ---------- bottom sheets ---------- */

function openSheet(kind, data) {
  const t = document.querySelector('.toast'); // a sheet's buttons are never under a toast
  if (t) t.remove();
  ui.sheet = Object.assign({ kind }, data || {});
  ui.keepScroll = true;
  render();
}

function closeSheet() {
  ui.sheet = null;
  const s = document.querySelector('.scrim');
  if (s) s.remove();
}

function renderSheet() {
  const s = ui.sheet;
  const old = document.querySelector('.scrim');
  if (old) old.remove();
  let h = '';
  const btn = (act, label, cls, extra) => '<button class="' + (cls || 'second') + '" data-act="' + act + '"' + (extra || '') + '>' + esc(label) + '</button>';
  if (s.kind === 'move') {
    const edit = ui.screen === 'edit';
    const oc = (edit ? ui.again.oc : R.board.oc).find((o) => o[0] === s.id);
    if (!oc) { // gone from the board since it was drawn
      ui.sheet = null;
      return render();
    }
    const now = edit ? placeAgain(oc) : placeOf(oc[0]);
    const sq = edit ? ui.again.sq : R.sq || 4;
    const away = edit ? '' : awayOf(oc); // a newcomer the cadre added, whose phone signed up another name since
    const other = away ? fullName(...pickName(X.picks[oc[5]])) : '';
    const was = 'The phone that signed up ' + fullName(oc[1], oc[2]) + ' signed up ' + other + ' since';
    const beside = !edit && now === null && placedByCadre(oc[0]) && oc[4] ? besideOf(oc) : undefined;
    const picked = now === -1 ? [...new Set((picks().byId[oc[0]] || []).map((x) => (X.picks[x.dv] || {}).sqd).filter(Boolean))].sort((a, b) => a - b) : []; // two phones disagree: the squads they picked
    const line = beside ? fullName(beside[1], beside[2]) + ' was added beside this one, who hasn\'t opened the link: the same OC? Leave this one out. Two OCs? Type this one\'s first name, then the squad.'
      : away === 'same' ? was + '. The same OC? Leave this one out.'
      : now === -1 ? 'Phones picked ' + listOf(picked.map((q) => ord(q) + ' SQD')) + '. Ask ' + oc[1] + ' which, then tap it.' : (now ? 'Now in ' + ord(now) + ' SQD.' : now === 0 ? 'Not in this evaluation.' : 'Not in a squad yet.') + (away === 'lent' ? ' ' + was + ': this one keeps the squad picked before.' : '');
    const pk = edit ? null : picks();
    const twin = !edit && !oc[2] && !pk.firsts[oc[0]] && R.board.oc.some((y) => y[0] !== oc[0] && (!y[3] || !oc[3] || y[3] === oc[3]) && lastKey(y[1]) === lastKey(oc[1])); // another OC of this last name: raters tell them apart by the first name
    h = '<h2>' + esc(fullName(oc[1], oc[2]) + (oc[3] ? ' · ' + ord(oc[3]) + ' PLT' : '')) + '</h2><p class="body">' + esc(line) + '</p>' +
      (twin ? '<label class="sub" for="mfirst">Another ' + esc(oc[1]) + ' is in this evaluation. This one\'s first name, then the squad:</label><input id="mfirst" class="input" maxlength="30" autocomplete="off" autocorrect="off" autocapitalize="words" spellcheck="false" enterkeyhint="done"><p class="err" id="mhint" role="status"></p>' : '') +
      Array.from({ length: sq }, (x, i) => i + 1).map((q) => btn('place', ord(q) + ' SQD' + (now === q ? ' · now' : picked.indexOf(q) >= 0 ? ' · picked' : ''), now === q ? 'primary' : 'second', ' data-id="' + oc[0] + '" data-v="' + q + '"')).join('') +
      btn('place', 'Leave out of this evaluation', 'second', ' data-id="' + oc[0] + '" data-v="0"') +
      (edit ? btn('oc-fix', 'Fix the name', 'link', ' data-id="' + oc[0] + '"') : '') + btn('close-sheet', 'Cancel', 'link');
  } else if (s.kind === 'oc') { // Fix the name / Add an OC: the taps first, so the keyboard never hides them
    const edit = ui.screen === 'edit';
    const seg = (f, vals) => '<div class="seg" style="--n:' + vals.length + '">' + vals.map((v) => '<button data-act="oc-pick" data-f="' + f + '" data-v="' + v + '" aria-pressed="' + (s[f] === v) + '">' + ord(v) + '</button>').join('') + '</div>';
    const sqs = Array.from({ length: edit ? ui.again.sq : Math.max(ui.draft.sq || 4, ...draftRoster().ocs.map((o) => o.sqd || 0)) }, (x, i) => i + 1); // a 5th SQD in the file is offered too
    h = '<h2>' + (s.isNew ? 'Add an OC' : 'Fix the name') + '</h2>' +
      (s.plts ? '<span class="sub">Platoon</span>' + seg('plt', s.plts) : '') + (s.askSq ? '<span class="sub">Squad</span>' + seg('sqd', sqs) : '') +
      '<label class="sub" for="olast">Last name</label><input id="olast" class="input" maxlength="30" autocomplete="off" autocorrect="off" autocapitalize="characters" spellcheck="false" enterkeyhint="next" value="' + esc(s.last) + '">' +
      '<label class="sub" for="ofirst">First name</label><input id="ofirst" class="input" maxlength="30" autocomplete="off" autocorrect="off" autocapitalize="words" spellcheck="false" enterkeyhint="done" value="' + esc(s.first) + '">' +
      '<p class="err" id="odup" role="status"></p>' + btn('oc-save', s.isNew ? 'Add' : 'Save', 'primary', ' id="osave"') +
      (!s.isNew && !edit ? btn('oc-out', 'Leave out of this evaluation', 'link danger') : '') + btn('close-sheet', 'Cancel', 'link');
  } else if (s.kind === 'extra') {
    const pk = picks();
    const x = pk.extra.find((e) => e.dv === s.dv);
    if (!x) { // added or left out since it was drawn (on another device too)
      ui.sheet = null;
      return render();
    }
    const name = fullName(x.last, x.first);
    const plts = x.plt ? [x.plt] : [...new Set(R.board.oc.map((o) => o[3]).filter(Boolean))].sort((m, n) => m - n);
    const before = R.board.oc.filter((o) => o[5] === x.dv).pop(); // this phone added before by another name: the name fixed, or a borrowed phone; only the cadre can tell
    const fixing = !!before && (x.me === 1 || (x.me !== 0 && (lastKey(before[1]) === lastKey(x.last) || nearMiss(before[1], x.last)) && firstNear(before[2], x.first))); // he said it's him (or, with no answer, a close name: likely his fix)
    const match = R.board.oc.filter((o) => o[5] !== x.dv && sameName(o, x.last, x.first));
    const same = !before && match.length === 1 ? match[0] : null; // this very name is in: a roster name the OC missed, or added from his other browser
    const named = (o) => o[2] || pk.firsts[o[0]];
    const nameless = same ? [] : R.board.oc.filter((o) => o[5] !== x.dv && lastKey(o[1]) === lastKey(x.last) && !named(o) && (!x.plt || !o[3] || o[3] === x.plt) && placeOf(o[0], pk) !== 0); // in this evaluation and platoon, with this last name and no first name
    const lone = R.board.oc.filter((o) => o[5] !== x.dv && lastKey(o[1]) === lastKey(x.last)).length === 1; // the one OC of this last name (no twin, no named one)
    const bare = !before && !x.told && x.first && lone && nameless.length === 1 && !nameless[0][5] && !placedByCadre(nameless[0][0]) && !(pk.byId[nameless[0][0]] || []).length ? nameless[0] : null; // nobody placed it or picked it, and the phone didn't say someone else: maybe this OC, who typed his first name too
    const other = nameless.filter((o) => (pk.byId[o[0]] || []).length); // signed up with no first name: theirs, if the cadre knows it (one not signed up yet is asked on his phone)
    const needFirst = match.length > 0 || (!x.first && (R.board.oc.some((o) => lastKey(o[1]) === lastKey(x.last)) || pk.extra.some((e) => e.dv !== x.dv && lastKey(e.last) === lastKey(x.last)))); // another OC with this last name: raters tell them apart by the first name
    const ask = before ? ' This phone signed up ' + fullName(before[1], before[2]) + ' before' + (x.me === 1 ? ': the OC says it\'s him, the name fixed.' : x.me === 0 ? ': the OC says someone else (a lent phone).' : ': the same OC with the name fixed, or someone else?')
      : same ? ' ' + name + (same[5] ? ' is in already, from another phone.' : ' is on the roster: a name they may have missed.') + ' The same OC?'
        : bare ? ' ' + bare[1] + (bare[3] ? ' (' + ord(bare[3]) + ' PLT)' : '') + ' on the roster has no first name and hasn\'t signed up: the same OC?' : ' Add them?';
    const who = before || same || bare ? 'Someone else: a' : 'A';
    const firstAsk = match.length ? 'Someone else? Their first name, so raters can tell them apart (a middle initial too):'
      : R.board.oc.some((o) => lastKey(o[1]) === lastKey(x.last)) ? 'Another ' + x.last + ' is in this evaluation. Their first name, so raters can tell them apart:' : 'Another phone typed ' + x.last + ' too. Their first name, so raters can tell them apart:';
    const field = (id, label, more) => '<label class="sub" for="' + id + '">' + esc(label) + '</label><input id="' + id + '" class="input" maxlength="30" autocomplete="off" autocorrect="off" autocapitalize="words" spellcheck="false" enterkeyhint="done" value="' + esc(s[id] || '') + '"' + (more || '') + '>';
    h = '<h2>' + esc(name) + '</h2><p class="body">' + (same || bare ? '' : 'Not on the roster. ') + 'Picked ' + (x.plt ? ord(x.plt) + ' PLT, ' : '') + ord(x.sqd) + ' SQD.' + esc(ask) + '</p>' +
      (same || bare ? btn('extra-link', 'Same OC', 'primary', ' data-dv="' + esc(x.dv) + '" data-id="' + (same || bare)[0] + '"') : '') +
      (needFirst ? field('xfirst', firstAsk) : '') +
      (other.length === 1 ? field('xother', 'The ' + x.last + ' in already has no first name. Theirs, if you know it, so raters can tell them apart:', ' data-id="' + other[0][0] + '"') : '') +
      (needFirst || other.length === 1 ? '<p class="err" id="xhint" role="status"></p>' : '') +
      (before ? btn('extra-same', 'Same OC: rename ' + fullName(before[1], before[2]) + ' to ' + name, fixing ? 'primary' : 'second', ' data-dv="' + esc(x.dv) + '" data-id="' + before[0] + '"') : '') +
      (plts.length > 1 ? plts.map((pl) => btn('extra-add', who + 'dd to ' + ord(pl) + ' PLT', 'second', ' data-dv="' + esc(x.dv) + '" data-plt="' + pl + '"')).join('') : btn('extra-add', who + 'dd ' + x.last, fixing || same || bare ? 'second' : 'primary', ' data-dv="' + esc(x.dv) + '" data-plt="' + (plts[0] || 0) + '"')) +
      btn('extra-ignore', before || same || bare ? 'Someone else: leave out' : 'Leave out', 'second', ' data-dv="' + esc(x.dv) + '"') + btn('close-sheet', 'Cancel', 'link');
  } else if (s.kind === 'start') {
    h = '<h2>Start without ' + plural(s.names.length, 'OC', 'OCs') + '?</h2><p class="body">' + esc(listOf(s.names, 5)) + (s.names.length === 1 ? ' isn\'t' : ' aren\'t') + ' in a squad. They won\'t be in this evaluation, and their phones will tell them to see you.</p>' +
      btn('start-force', 'Start without them', 'primary') + btn('close-sheet', 'Not yet', 'link');
  } else if (s.kind === 'twins') {
    h = '<h2>One thing to fix first</h2><p class="body">' + esc(listOf(s.names.map((n) => 'two ' + n), 5)) + ' with no first names are in one squad: raters couldn\'t tell them apart. Tap each one and type their first name.</p>' + btn('close-sheet', 'OK', 'primary');
  } else if (s.kind === 'fix') {
    h = '<h2>One thing to fix first</h2><p class="body">Two phones put ' + esc(listOf(s.names, 5)) + ' in different squads. Tap the red name to choose the squad.</p>' + btn('close-sheet', 'OK', 'primary');
  } else if (s.kind === 'save2first') {
    h = '<h2>Save page 2 first</h2><p class="body">This ' + device() + ' keeps nothing. Page 2 is your only way back to the results.</p>' + btn('save2-then-share', 'Save page 2', 'primary') + btn('close-sheet', 'Cancel', 'link');
  } else if (s.kind === 'forget') {
    const kept = (R.imported || []).length; // added here from a file or text, and not on the sheet yet
    h = '<h2>Forget it on this ' + device() + '?</h2><p class="body">This evaluation leaves this ' + device() + '. Page 2 of your printout still opens it.' + (kept ? ' The ' + plural(kept, 'result', 'results') + ' you added here that the sheet doesn\'t have yet go with it.' : '') + '</p>' + btn('forget-go', 'Forget', 'primary') + btn('close-sheet', 'Cancel', 'link');
  } else if (s.kind === 'remove') {
    h = '<h2>Remove it from the sheet?</h2><p class="body">Every answer to this evaluation is cleared from the sheet, and its QR code stops working. Download Excel first to keep your own copy. This can\'t be undone.</p>' + btn('remove-go', 'Remove it', 'primary') + btn('close-sheet', 'Cancel', 'link');
  }
  const el = document.createElement('div');
  el.className = 'scrim';
  el.innerHTML = '<div class="sheet" role="dialog" aria-modal="true">' + h + '</div>';
  el.addEventListener('click', (e) => {
    if (e.target === el && !el.querySelector('input')) closeSheet(); // a name being typed goes only with Cancel
  });
  el.addEventListener('input', (e) => {
    if (e.target.id === 'xfirst' || e.target.id === 'xother') s[e.target.id] = e.target.value;
  });
  document.body.appendChild(el);
  if (s.kind === 'oc') {
    el.addEventListener('input', checkOc);
    el.addEventListener('keydown', (e) => { // Return: Last name, then First name, then Save
      if (e.key !== 'Enter' || !/^(olast|ofirst)$/.test(e.target.id)) return;
      e.preventDefault();
      if (e.target.id === 'olast') document.getElementById('ofirst').focus();
      else saveOc();
    });
    checkOc();
    if (s.isNew && !s.last) document.getElementById('olast').focus();
  }
}

/* ---------- actions ---------- */

function chooseFiles(accept, onFiles) {
  const inp = document.createElement('input');
  inp.type = 'file';
  inp.multiple = true;
  inp.accept = accept;
  inp.id = 'file-pick';
  inp.className = 'vh';
  inp.addEventListener('change', () => {
    onFiles([...inp.files]);
    inp.remove();
  });
  document.body.appendChild(inp);
  inp.click();
}

/**
 * Filled-in Excel forms (17h): each OC's form becomes that OC's result, sealed the way their phone would send it, then
 * added like a result file. Names are matched within the form's squad. 17m (the user: "Yes, newest Excel wins"): a form
 * replaces an earlier Excel form (a higher version, so every device and the sheet take it); an OC who answered in the app
 * keeps those answers; two or more forms of one OC in an upload: only the newest (the file's or the zip entry's date; with
 * no date, the last in the upload), never mixed.
 * @returns {Promise<{envs: object[], notes: string[]}>}
 */
async function formEnvelopes(files) {
  const E = X.E;
  const people = (ids) => E.members.filter((m) => ids.indexOf(m.id) >= 0);
  const byOc = {};
  const unknown = [];
  let noName = 0;
  let any = 0;
  for (let i = 0; i < files.length; i++) {
    let forms = [];
    try {
      forms = readForms(await readXlsx(new Uint8Array(await files[i].arrayBuffer())));
    } catch (err) {
      forms = []; // not a workbook this page can read: said below if nothing came of it
    }
    any += forms.length;
    const groups = {};
    forms.forEach((f) => {
      const k = f.plt + '|' + f.sqd + '|' + f.rater.toUpperCase();
      (groups[k] = groups[k] || []).push(f);
    });
    for (const tabs of Object.values(groups)) {
      const f = tabs[0];
      if (!f.rater) {
        noName++;
        continue;
      }
      let me = null;
      let sq = null;
      for (const s of E.squads.filter((x) => (!f.plt || x.plt === f.plt) && (!f.sqd || x.sqd === f.sqd))) {
        const m = matchName(f.rater, people(s.ids));
        if (m) [me, sq] = [m, s];
      }
      if (!me) {
        if (unknown.indexOf(f.rater) < 0) unknown.push(f.rater);
        continue;
      }
      (byOc[me.id] = byOc[me.id] || []).push({ i, time: Number(files[i].lastModified) || 0, tabs, me, sq });
    }
  }
  const envs = [];
  const replaced = [];
  const kept = [];
  const twice = [];
  for (const list of Object.values(byOc)) {
    const dated = list.every((c) => c.time > 0);
    const c = list.slice().sort((x, y) => (dated ? x.time - y.time : 0) || x.i - y.i).pop(); // the newest; with no usable dates, the last in the upload
    const cur = X.results[c.me.id];
    if (cur && String((cur.payload || {}).dv) !== 'excel-form') {
      kept.push(c.me.name);
      continue;
    }
    if (list.length > 1) twice.push([c.me.name, list.length]);
    const a = formAnswers(c.tabs, people(c.sq.ids.filter((id) => id !== c.me.id)));
    if (!a.n) {
      unknown.push(c.me.name + '\'s squadmates');
      continue;
    }
    const Pk = await packetOf(R.board, R.k, c.me.id, '');
    envs.push(await sealResult(Pk, { v: 1, e: R.e, r: c.me.id, rev: cur ? cur.rev + 1 : 1, at: new Date().toISOString(), ans: a.ans, com: a.com, sc: a.sc, why: a.why, rank: a.rank, dv: 'excel-form' }));
    if (cur) replaced.push(c.me.name);
  }
  const notes = [];
  if (replaced.length) notes.push(replaced.length === 1 ? 'Replaced the earlier Excel form for ' + replaced[0] + '.' : 'Replaced the earlier Excel forms for ' + replaced.length + ' OCs: ' + listOf(replaced, 3) + '.');
  if (kept.length) notes.push(kept.length === 1 ? kept[0] + ' answered in the app; the Excel form was not used.' : kept.length + ' OCs answered in the app; their Excel forms were not used: ' + listOf(kept, 3) + '.');
  if (twice.length) notes.push(twice.length === 1 ? twice[0][0] + ' sent ' + twice[0][1] + ' forms in this upload; the newest was used.' : twice.length + ' OCs sent more than one form in this upload; the newest of each was used: ' + listOf(twice.map((x) => x[0]), 3) + '.');
  if (unknown.length) notes.push('Not on this evaluation\'s list: ' + listOf(unknown, 3) + '.');
  if (noName) notes.push(plural(noName, 'form has', 'forms have') + ' answers but no name picked in Rater Name (B2): pick it and add them again.');
  if (files.length && !any) notes.push('No filled-in forms in that.');
  return { envs, notes };
}

/** Files on Add results: result files and text as before, filled-in Excel forms (17h) read as forms. */
async function importFiles(files) {
  files = await expandFiles(files);
  const xl = files.filter((f) => /\.xls[xm]$/i.test(f.name));
  let text = '';
  for (const f of files.filter((x) => xl.indexOf(x) < 0)) text += (await f.text()) + '\n';
  return importText(text, xl.length ? await formEnvelopes(xl) : null);
}

let paperForms = null; // 17j: filled-in forms dropped on Who's in it?: they go in as results once the evaluation is made

async function importText(text, more) {
  const envs = findEnvelopes(String(text || '')).concat(more ? more.envs : []);
  const r = R;
  let n = 0;
  for (const env of envs) {
    if (!(await takeRow(env))) continue;
    n++;
    if (!env.j) r.imported = (r.imported || []).filter((x) => JSON.stringify(x) !== JSON.stringify(env)).concat([env]); // kept with the evaluation: a reload, or opening it again, has it too
  }
  toast((n ? 'Added ' + plural(n, 'result', 'results') + '.' : 'No new results for this evaluation in that.') + (more && more.notes.length ? ' ' + more.notes.join(' ') : ''));
  if (n) {
    R.used = true;
    persist();
    sendImported(r);
    ui.keepScroll = true;
    render(); // the count and the lists show them at once
  }
}

/** Results added here from a file or text go to the sheet too, as their phones would have sent them: every cadre device reads them then. This device keeps each one until the sheet has it. */
async function sendImported(r) {
  await Promise.all((r.imported || []).filter((env) => !importing.has(env.mac)).map(async (env) => { // each on its own way, never twice at once
    importing.add(env.mac);
    let res = null;
    try {
      res = await call(env);
    } catch (err) {
      res = null;
    }
    importing.delete(env.mac);
    if (!res || !res.ok) return; // tried again after the next good read, or the next time it opens
    r.imported = (r.imported || []).filter((x) => x.mac !== env.mac);
    persist();
  }));
}

const act = {
  back: (b) => {
    if (ui.screen === 'removed' && R) forgetHere(R); // seen: it leaves this device as the cadre moves on
    go(b.dataset.to || 'home');
  },
  more: () => go('more'),
  'close-more': () => go(ui.moreFrom || (R ? screenOf(R) : 'home')),
  start: () => {
    if (ui.screen === 'removed' && R) forgetHere(R);
    clearTimeout(timers.poll);
    R = null;
    X = null;
    ui.draft = newDraft();
    ui.paste = false;
    ui.find = '';
    go('roster');
  },
  open: () => go('open'),
  'open-eval': (b) => openEval(list[Number(b.dataset.i)]),
  choose: () => chooseFiles('.xlsx,.xlsm,.csv,.tsv,.txt,.zip', addFiles),
  'roster-xlsx': () => { // Who's in it? as a sheet: the names, platoons and squads the evaluation would get, to edit in Excel and drop back in later
    const f = rosterXlsx(ui.draft.name || 'Peer evaluation', draftList(draftRoster()).map(([, last, first, plt, sqd]) => ({ last, first, plt, sqd })), new Date());
    download(f.name, f.data, XLSX_TYPE);
    toast('Saved ' + f.name + '. Edit it in Excel if you like; next time, drop it here.');
  },
  paste: () => {
    ui.paste = true;
    render();
    const t = document.getElementById('paste');
    if (t) t.focus();
  },
  'paste-add': () => {
    const t = document.getElementById('paste');
    const text = t ? t.value : '';
    if (!text.trim()) return toast('Paste the names first.');
    ui.paste = false;
    return addSource(readRoster(text), 'Pasted list', null);
  },
  'plt-sel': (b) => {
    ui.draft.pltSel = Number(b.dataset.v);
    render();
  },
  'plt-all': (b) => {
    ui.draft.pltAll = Number(b.dataset.v);
    render();
  },
  'oc-plt': (b) => {
    const D = ui.draft;
    D.dec.plt = Object.assign({}, D.dec.plt || {}, { [b.dataset.key]: Number(b.dataset.v) });
    ui.keepScroll = true;
    render();
  },
  'q-same': (b) => decide('same', b),
  'q-which': (b) => decide('which', b),
  'q-add': (b) => decide('add', b),
  'oc-edit': (b) => { // a name on Who's in it?: fix it or leave it out
    const dr = draftRoster();
    const o = dr.m.ocs.find((x) => x.key === b.dataset.key);
    if (!o) return undefined;
    const plts = dr.namesOnly ? null : !o.plt ? o.plts || dr.plts : dr.plts.length > 1 ? dr.plts : null; // a one-platoon roster asks nothing
    return openSheet('oc', { ref: o.key, keys: o.keys, added: o.srcs[0] === 'j', last: o.last, first: o.first || '', plt: o.plt || 0, sqd: o.sqd || 0, was: { last: o.last, first: o.first || '', plt: o.plt || 0, sqd: o.sqd || 0 }, plts: plts && plts.length ? plts : null, askSq: dr.allSquads });
  },
  'oc-new': () => {
    if (ui.screen === 'edit') { // Again's Who's in it: an insert joins the next evaluation, in a squad
      const plts = [...new Set(ui.again.oc.map((o) => o[3] || 0))].sort((m, n) => m - n);
      return openSheet('oc', { isNew: true, last: '', first: '', plt: plts.indexOf(ui.tab) >= 0 ? ui.tab : plts[0], sqd: 0, plts: plts.length > 1 ? plts : null, askSq: true }); // the platoon on screen
    }
    const dr = draftRoster();
    const one = ui.draft.pltSel || (dr.plts.length === 1 ? dr.plts[0] : 0); // under All with several platoons, nothing is picked for the cadre
    const typed = String(ui.find || '').trim();
    const none = typed && !dr.ocs.some((o) => fullName(o.last, o.first).toLowerCase().indexOf(typed.toLowerCase()) >= 0);
    return openSheet('oc', { isNew: true, last: none ? typed : '', first: '', plt: one, sqd: 0, plts: !dr.namesOnly && !one && dr.plts.length > 1 ? dr.plts : null, askSq: dr.allSquads });
  },
  'oc-fix': (b) => {
    const o = ui.again.oc.find((x) => x[0] === Number(b.dataset.id));
    return o ? openSheet('oc', { ref: o[0], last: o[1], first: o[2] || '', plt: o[3] || 0, sqd: o[4] || 0 }) : undefined;
  },
  'oc-pick': (b) => {
    const s = ui.sheet;
    Object.assign(s, { last: fieldOf('olast'), first: fieldOf('ofirst') }); // the redraw keeps what's typed
    s[b.dataset.f] = Number(b.dataset.v);
    if (b.dataset.f === 'plt' && s.askSq) s.sqd = 0; // another platoon: tap its squad again
    renderSheet();
  },
  'oc-save': () => saveOc(),
  'oc-out': () => leaveOut(),
  'oc-back': (b) => putBack(b.dataset.key),
  'read-next': () => {
    const D = ui.draft;
    const dr = draftRoster();
    if (!dr.ocs.length) return toast(dr.outs.length ? 'An evaluation needs at least two OCs.' : 'Add a roster first.');
    if (dr.namesOnly && !D.pltAll) return toast('Tap which platoon this is first.');
    if (dr.noPlt.length && !dr.namesOnly) return toast('Tap a platoon for each OC without one first.');
    if (dr.questions.some((q) => q.kind !== 'add')) return toast('Answer the questions about the names first.');
    if (dr.questions.length) return toast('Choose Add or Leave out for each name first.');
    if (dr.block.length) return toast(dr.block[0]);
    if (dr.ocs.length < 2) return toast('An evaluation needs at least two OCs.');
    D.mode = 'set'; // 17c: squads not in the file are tapped by the cadre; OCs never choose
    return dr.allSquads ? toOwnOrSave('read') : go('set');
  },
  tab: (b) => {
    ui.tab = Number(b.dataset.v);
    render();
  },
  'set-sq': (b) => {
    ui.draft.setMap[b.dataset.key] = Number(b.dataset.v);
    ui.keepScroll = true;
    render();
  },
  'set-next': () => {
    const dr = draftRoster();
    const left = dr.ocs.filter((o) => !o.sqd && !setOf(o)).length;
    if (left) return toast('Every OC needs a squad first: ' + left + ' to go.');
    const twin = twinIn(dr.ocs.map((o) => ({ last: o.last, first: o.first, plt: o.plt, sqd: setOf(o) || o.sqd })));
    if (twin) return toast('Two OCs named ' + twin.last + ' with no first names are in one platoon. Go back and tap each name to add a first name.');
    return toOwnOrSave('set');
  },
  'own-yes': () => {
    try {
      localStorage.setItem(OWN, '1');
      own = true;
    } catch (err) {
      own = false;
    }
    return saveDraft();
  },
  'own-no': () => {
    own = false;
    ui.askedOwn = true;
    return saveDraft();
  },
  print: () => {
    printPages(false);
    toast('Printing 2 pages… Post page 1 in each platoon.');
    clearTimeout(timers.after);
    timers.after = setTimeout(() => ui.screen === 'qr' && go(afterCode()), 1400); // not if the cadre has moved on
  },
  share: () => {
    if (!own && !R.page2) return openSheet('save2first');
    return shareLink();
  },
  'save2-then-share': () => {
    closeSheet();
    printPages(true);
    toast('Page 2 is ready to print or save as a PDF.');
    setTimeout(shareLink, 600);
  },
  save2: () => {
    printPages(true);
    toast('Page 2 is ready to print or save as a PDF.');
    render();
  },
  'show-qr': () => go('qr'),
  'print-feedback': () => printFeedback(),
  move: (b) => openSheet('move', { id: Number(b.dataset.id) }),
  place: (b) => placeOc(Number(b.dataset.id), Number(b.dataset.v)),
  extra: (b) => openSheet('extra', { dv: b.dataset.dv }),
  'extra-add': (b) => addExtra(b.dataset.dv, Number(b.dataset.plt) || 0),
  'extra-same': (b) => { // the OC this phone signed up, the name fixed: that row takes the new spelling (its pick then places it, as the phone says "You're in …"), and the phone finds it by its id and name
    if (frozen()) return toast('Starting the evaluation… one moment.');
    const x = picks().extra.find((e) => e.dv === b.dataset.dv);
    const o = ocOf(Number(b.dataset.id));
    if (!x || !o) return undefined;
    const first = extraFirst(x, o[0]); // a first name typed on this sheet counts here too
    if (first === null) return undefined;
    o[1] = x.last;
    if (first) o[2] = first;
    if (x.first && first && first.toLowerCase() !== x.first.toLowerCase()) R.board.ln = Object.assign({}, R.board.ln || {}, { [ignKey(x)]: o[0] }); // another first name than the phone typed (Mike Q): the phone and its pick find the row by the phone
    if (x.sqd && placedByCadre(o[0])) o[4] = x.sqd; // a row placed by hand (or added before 16o): the squad the phone picked last
    closeSheet();
    render();
    boardChanged();
    return refreshEvent();
  },
  'extra-link': (b) => { // Same OC: the name typed on this phone is that OC's (his roster name, or added from his other browser): his pick counts there, and the phone opens it
    if (frozen()) return toast('Starting the evaluation… one moment.');
    const x = picks().extra.find((e) => e.dv === b.dataset.dv);
    const o = ocOf(Number(b.dataset.id));
    if (!x || !o) return undefined;
    R.board.ln = Object.assign({}, R.board.ln || {}, { [ignKey(x)]: o[0] });
    if (!o[2] && x.first) o[2] = x.first; // his first name, typed on the phone
    closeSheet();
    render();
    return boardChanged();
  },
  'extra-ignore': (b) => {
    const x = picks().extra.find((e) => e.dv === b.dataset.dv);
    if (x) R.ignored = Object.assign({}, R.ignored || {}, { [ignKey(x)]: true }); // this name from this phone: another name typed on it still shows
    persist();
    closeSheet();
    render();
  },
  'close-sheet': () => {
    closeSheet();
  },
  'start-eval': () => startNow(false),
  'start-force': () => startNow(true),
  'copy-names': async () => {
    const lines = (X.E ? X.E.squads : []).map((sq) => {
      const left = sq.ids.filter((id) => !X.results[id]).map((id) => X.E.members.find((m) => m.id === id).name).sort();
      return left.length ? sq.label + ': ' + left.join(', ') : '';
    }).filter(Boolean);
    if (!lines.length) return toast('Everyone is done.');
    const text = 'Peer evaluation: still need\n' + lines.join('\n') + '\nOpen it in Safari or Chrome on the phone you started on (not inside this chat app) and finish: ' + ocUrl(R) +
      '\nYou\'re done when it says "Sent to your cadre". If your phone says "Saved on this device", open it again the same way where you have signal, or send it to your cadre as a file.';
    return toast((await copy(text)) ? 'Copied. Paste it in your group chat.' : 'Copy didn\'t work here.');
  },
  results: () => go('results'),
  'res-tab': (b) => {
    ui.resTab = Number(b.dataset.v);
    render();
  },
  download: () => {
    if (!X.E) return undefined;
    const f = resultsXlsx(X.E);
    return download(f.name, f.data, XLSX_TYPE);
  },
  answers: () => {
    const f = answersXlsx(X.E, counts(), new Date());
    download(f.name, f.data, XLSX_TYPE);
  },
  feedback: () => { // each OC's own feedback, anonymous: when everyone is in, or when the cadre ends it here
    if (!X.E) return undefined;
    const c = counts();
    if (c.done < c.total && !window.confirm(c.done + ' of ' + c.total + ' OCs have sent theirs. End here and make the feedback with what\'s in? Those not in yet are in no one\'s feedback.')) return undefined;
    const f = feedbackBooks(X.E, new Date());
    if (!f) return toast('No results in yet.');
    download(f.name, f.data, /\.zip$/.test(f.name) ? 'application/zip' : XLSX_TYPE);
    return toast('Saved: one feedback file per OC, with no names in it. Give each OC their own.');
  },
  sheets: () => { // the board as placed now, dealt as the phones will be (left out: on no sheet; not in yet: no squad to put him in)
    const pk = R.started ? null : picks();
    let none = 0;
    const oc = R.board.oc.map((o) => {
      const q = placeOf(o[0], pk);
      if (!(q > 0) && q !== 0) none += 1;
      return [o[0], o[1], o[2] || (pk && pk.firsts[o[0]]) || '', o[3] || 0, q > 0 ? q : 0];
    });
    return saveBlankSheets(oc, R.board.seed, R.board.n, none, R.board.dl);
  },
  'sheets-draft': () => { // before the QR code: the squads in the file or tapped here; the seed is kept, so the phones later get the same attributes
    const D = ui.draft;
    D.seed = D.seed || newSeed();
    const oc = draftList(draftRoster());
    const was = !!D.sheetsSig;
    const sig = sheetsSig(oc, D.seed);
    if (Object.keys(sig).length) D.sheetsSig = sig; // the sheets went out: a later name or squad change says "save them again"
    if (was) {
      ui.keepScroll = true;
      render();
    }
    return saveBlankSheets(oc, D.seed, D.name || 'Peer evaluation', oc.filter((o) => !o[4]).length, 3);
  },
  again: (b) => {
    const r = list.find((x) => x.e === b.dataset.e) || R;
    if (!r || !r.board) return undefined;
    ui.again = Object.assign(againFrom(r), { title: titleOf(r) });
    return go('again');
  },
  'edit-again': () => {
    ui.tab = 0;
    go('edit');
  },
  'make-again': () => makeAgain(),
  remake: () => {
    const r = R;
    forgetHere(r);
    ui.again = Object.assign(againFrom(r), { title: titleOf(r) });
    if (r.mode === 'oc' && !r.started) {
      const oc = r.board.oc.map((o, i) => [i, o[1], o[2], o[3], r.fixed && r.fixed[o[0]] ? o[4] : 0]);
      return makeCode(oc, 'oc', r.name, r.sq, r.e);
    }
    return makeAgain();
  },
  'copy-cadre': async () => toast((await copy(cadreUrl(R))) ? 'Copied. Anyone with this link sees the results.' : 'Copy didn\'t work here.'),
  forget: () => openSheet('forget'),
  'forget-go': () => {
    forgetHere(R);
    clearTimeout(timers.poll);
    R = null;
    X = null;
    toast('Forgotten on this ' + device() + '.');
    go(list.length ? 'all' : 'home');
  },
  'keep-here': () => keepHere(),
  'remove-sheet': () => openSheet('remove'),
  'remove-go': async () => {
    closeSheet();
    const r = R;
    let res = null;
    try {
      res = await forgetEval(r.e, await wt(r));
    } catch (err) {
      res = null;
    }
    if (!res || !res.ok) return toast('Can\'t remove it right now. Try again in a minute.');
    forgetHere(r);
    toast('Removed from the sheet.');
    if (R !== r) return ui.screen === 'all' ? render() : undefined; // the cadre opened another one meanwhile: the Home list loses this row
    R = null;
    X = null;
    return go(list.length ? 'all' : 'home');
  },
  name: () => {
    const t = nameTarget();
    ui.nm = { name: t.name, sq: t.sq || 4 };
    go('name');
  },
  'sq-count': (b) => {
    const el = document.getElementById('ename');
    if (el) ui.nm.name = el.value; // the redraw keeps what's typed
    ui.nm.sq = Number(b.dataset.v);
    render();
  },
  'name-save': () => {
    const v = String((document.getElementById('ename') || {}).value || '').trim().slice(0, 40) || 'Peer evaluation';
    const t = nameTarget();
    if (t === R && frozen()) return toast('Starting the evaluation… one moment.'); // a frozen board is never sealed again by a name change
    t.name = v;
    t.sq = ui.nm.sq;
    if (t === R) {
      R.board.n = v;
      R.board.sq = R.sq;
      boardChanged();
      refreshEvent(); // Download Excel uses the new name
    }
    toast('Saved.');
    go(ui.moreFrom || 'more');
  },
  import: () => go('import'),
  'import-go': () => importText((document.getElementById('imp') || {}).value),
  'import-files': () => chooseFiles('.txt,.json,.peval,.xlsx,.xlsm,.zip', importFiles),
  admin: () => {
    location.href = site() + 'admin.html' + testQuery();
  }
};

function decide(field, b) {
  const D = ui.draft;
  D.dec[field] = Object.assign({}, D.dec[field] || {}, { [b.dataset.q]: b.dataset.v });
  if (b.dataset.not) D.dec.not = Object.assign({}, D.dec.not || {}, { [b.dataset.q]: [...new Set(((D.dec.not || {})[b.dataset.q] || []).concat(b.dataset.not.split(' ')))] }); // Someone else: none of the OCs offered (nor the ones before): one never offered can still be this OC
  ui.keepScroll = true;
  render();
}

function toOwnOrSave(from) {
  ui.from = from;
  if (own || ui.askedOwn) return saveDraft();
  return go('own');
}

function saveDraft() {
  const D = ui.draft;
  const dr = draftRoster();
  const oc = draftList(dr);
  paperForms = (D.forms || []).slice();
  return makeCode(oc, 'set', D.name || 'Peer evaluation', D.sq, '', D.seed); // 17c: always set by the cadre, open from the QR code on
}

function placeOc(id, q) {
  if (ui.screen === 'edit') {
    const a = ui.again;
    const o = a.oc.find((x) => x[0] === id);
    a.out = a.out.filter((x) => x !== id);
    if (q) o[4] = q;
    else a.out.push(id);
    closeSheet();
    ui.keepScroll = true;
    return render();
  }
  if (frozen()) return toast('Starting the evaluation… one moment.');
  const o = ocOf(id);
  const typed = (typedName(o[1], fieldOf('mfirst')) || {}).first; // a first name typed on the sheet (another OC has this last name)
  if (q > 0 && !typed && besideOf(o)) { // two OCs, the newcomer and this one: this one's first name first (the same last name needs one)
    const el = document.getElementById('mhint');
    if (el) el.textContent = 'Type this one\'s first name first, or leave this one out.';
    return undefined;
  }
  noteNames(); // a twin named so far only by his phone: his name goes in with the move, so the row stays his
  if (typed && !o[2]) o[2] = typed;
  o[4] = q;
  R.fixed = Object.assign({}, R.fixed || {}, { [id]: true });
  closeSheet();
  ui.keepScroll = true;
  render();
  return boardChanged();
}

/** A field of the name sheet, as typed. */
const fieldOf = (id) => String((document.getElementById(id) || {}).value || '');

/** The name sheet, updated in place as the cadre types: Save stays off until it is a name, not one OC twice, with its platoon and squad tapped. */
function checkOc() {
  const s = ui.sheet;
  if (!s || s.kind !== 'oc') return;
  const v = typedName(fieldOf('olast'), fieldOf('ofirst'));
  let note = '';
  if (v) {
    const me = { last: v.last, first: v.first, plt: s.plt || 0 };
    let twin = null;
    let out = false;
    if (ui.screen === 'edit') { // Not in this one too: that OC comes back by a tap on the board, not as a second copy
      twin = ui.again.oc.find((o) => o[0] !== s.ref && sameOc({ last: o[1], first: o[2], plt: o[3] }, me));
      out = !!twin && ui.again.out.indexOf(twin[0]) >= 0;
    } else {
      const m = draftRoster().m;
      twin = m.ocs.find((o) => o.key !== s.ref && sameOc(o, me));
      if (!twin) {
        twin = m.outs.find((o) => o.key !== s.ref && sameOc(o, me));
        out = !!twin;
      }
    }
    if (twin) note = !v.first ? 'Two people? Add a first name.' : !out ? 'Already on the list.' : ui.screen === 'edit' ? 'Not in this one: tap it on the board and pick a squad.' : 'Left out: tap it under Left out to put it back.';
  }
  const save = document.getElementById('osave');
  const line = document.getElementById('odup');
  if (save) save.disabled = !v || !!note || (!!s.plts && !s.plt) || (!!s.askSq && !s.sqd);
  if (line) line.textContent = note;
}

/** Save or Add in the name sheet: the draft keeps a fix as a decision on the OC's rows and an insert as an added name; Again changes its own list. */
function saveOc() {
  const s = ui.sheet;
  const v = typedName(fieldOf('olast'), fieldOf('ofirst'));
  const save = document.getElementById('osave');
  if (!v || !save || save.disabled) return undefined;
  if (ui.screen === 'edit') {
    const a = ui.again;
    if (s.isNew) a.oc.push([Math.max(-1, ...a.oc.map((o) => o[0])) + 1, v.last, v.first, s.plt || 0, s.sqd]);
    else {
      const o = a.oc.find((x) => x[0] === s.ref);
      o[1] = v.last;
      o[2] = v.first;
    }
  } else {
    const D = ui.draft;
    const set = (field, key, value) => {
      D.dec[field] = Object.assign({}, D.dec[field] || {}, { [key]: value });
    };
    if (s.isNew) {
      D.seq += 1;
      D.added.push({ key: 'c' + D.seq, last: v.last, first: v.first, plt: s.plt || null, sqd: s.sqd || null });
      set('add', 'c' + D.seq, 'add');
    } else if (s.added) {
      Object.assign(D.added.find((x) => x.key === s.ref), { last: v.last, first: v.first, plt: s.plt || null, sqd: s.sqd || null });
    } else {
      const named = fieldOf('olast') !== s.was.last || fieldOf('ofirst') !== s.was.first; // what the cadre typed: Save with the name as it was fixes nothing
      if (named || s.plt !== s.was.plt || s.sqd !== s.was.sqd) { // Fixed only when something changed
        const typed = fieldOf('ofirst') !== s.was.first; // the first name only where it was typed: a later file can still fill an empty one
        if (named) set('name', s.ref, Object.assign(typed ? { saw: s.was.first } : {}, (D.dec.name || {})[s.ref], { last: v.last }, typed ? { first: v.first } : {})); // with the one shown before the first fix: a later file's row with it still finds this OC
        if (s.plt !== s.was.plt) set('plt', s.ref, s.plt);
        if (s.sqd !== s.was.sqd) set('sq', s.ref, s.sqd);
        D.touched[s.ref] = true;
      }
    }
    ui.find = '';
  }
  closeSheet();
  ui.keepScroll = true;
  render();
  const D = ui.draft;
  if (ui.screen === 'read' && D.pltSel && s.plt && s.plt !== D.pltSel) toast(fullName(v.last, v.first) + ' is in ' + ord(s.plt) + ' PLT now, so not on this list.'); // only while that platoon is still the one shown
  return undefined;
}

/**
 * Leave out of this evaluation (a recycle): the OC goes under Left out; Undo on the toast, or Put back there, brings it back.
 * One OC listed twice: only the copy's row in the file that lists it twice goes (another file's row stays with the one kept,
 * and a later file matches that one); the copy is under Left out too, to put back if they were two OCs.
 */
function leaveOut() {
  const s = ui.sheet;
  const D = ui.draft;
  const m = draftRoster().m;
  const me = m.ocs.find((o) => o.key === s.ref);
  const twin = me && me.first ? m.ocs.find((o) => o.key !== me.key && sameOc(o, me)) : null;
  const copy = twin ? me.keys.filter((k, i) => twin.srcs.indexOf(me.srcs[i]) >= 0) : [];
  const field = copy.length ? 'drop' : 'out';
  const keys = copy.length ? copy : s.keys;
  const mark = (on) => {
    D.dec[field] = Object.assign({}, D.dec[field] || {});
    keys.forEach((k) => {
      if (on) D.dec[field][k] = true;
      else delete D.dec[field][k];
    });
  };
  mark(true);
  closeSheet();
  ui.keepScroll = true;
  render();
  toast('Left out ' + (copy.length ? 'the second ' : '') + fullName(s.was.last, s.was.first) + '.', 'Undo', () => {
    if (ui.screen !== 'read' || ui.draft !== D) return;
    mark(false);
    ui.keepScroll = true;
    render();
  });
}

/** Put back a left-out OC, unless that would put one OC on the list twice. A copy taken off always comes back: then the list says it's there twice. */
function putBack(key) {
  const D = ui.draft;
  if ((D.dec.drop || {})[key]) {
    D.dec.drop = Object.assign({}, D.dec.drop);
    delete D.dec.drop[key];
    ui.keepScroll = true;
    return render();
  }
  const m = draftRoster().m;
  const o = m.outs.find((x) => x.key === key);
  if (!o) return undefined;
  if (o.first && m.ocs.some((x) => sameOc(x, o))) return toast(fullName(o.last, o.first) + ' is on the list already.');
  D.dec.out = Object.assign({}, D.dec.out || {});
  o.keys.forEach((k) => delete D.dec.out[k]);
  ui.keepScroll = true;
  return render();
}

/** Two OCs with one last name and no first names in one platoon: no phone could tell them apart (it lists the platoon's names). @returns the second of them, or undefined */
function twinIn(oc) {
  const seen = {};
  return oc.find((o) => {
    if (o.first) return false;
    const k = (o.plt || 0) + '|' + o.last;
    if (seen[k]) return true;
    seen[k] = true;
    return false;
  });
}

/**
 * The first names typed on a Not on the list sheet: this OC's (another OC has the last name), and the one in already with
 * that last name and no first name (it goes on the board). @returns {string|null} this OC's first name, or null: a hint says what's missing
 */
function extraFirst(x, row) {
  const say = (el, t) => {
    document.getElementById('xhint').textContent = t;
    el.focus();
    return null;
  };
  const typed = document.getElementById('xfirst');
  const first = (typed && (typedName(x.last, typed.value) || {}).first) || x.first || '';
  if (typed && !first) return say(typed, 'Type their first name first.'); // another OC with this last name: raters tell them apart by the first name
  if (typed && R.board.oc.some((o) => o[0] !== row && o[5] !== x.dv && sameName(o, x.last, first))) return say(typed, fullName(x.last, first) + ' is in already. Type a first name that tells them apart.');
  const ot = document.getElementById('xother'); // the one in already with this last name and no first name: theirs too, if the cadre knows it
  const otFirst = ot ? (typedName(x.last, ot.value) || {}).first || '' : '';
  if (otFirst && otFirst.toLowerCase() === first.toLowerCase()) return say(ot, 'Type two first names that tell them apart.');
  const was = ot ? ocOf(Number(ot.dataset.id)) : null;
  if (was && !was[2] && otFirst) was[2] = otFirst;
  return first;
}

function addExtra(dv, plt) {
  if (frozen()) return toast('Starting the evaluation… one moment.');
  const x = picks().extra.find((e) => e.dv === dv);
  if (!x) return undefined;
  const first = extraFirst(x);
  if (first === null) return undefined;
  const id = Math.max(-1, ...R.board.oc.map((o) => o[0])) + 1;
  R.board.oc.push([id, x.last, first, x.plt || plt || 0, 0, dv]); // the phone that typed it finds this row by its own id and the name, and its pick places it (a newer squad too), as any OC's; the board shows it's added
  if (x.first && first.toLowerCase() !== x.first.toLowerCase()) R.board.ln = Object.assign({}, R.board.ln || {}, { [ignKey(x)]: id }); // another first name than the phone typed (Mike Q): the phone and its pick find the row by the phone
  closeSheet();
  render();
  boardChanged();
  return refreshEvent();
}

/* ---------- start ---------- */

root.addEventListener('click', (e) => {
  const b = e.target.closest('[data-act]');
  if (!b) return;
  const fn = act[b.dataset.act];
  if (fn) fn(b);
});
root.addEventListener('keydown', (e) => { // a card that works as a button (the file drops): Return or Space taps it
  if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('[role=button][data-act]')) {
    e.preventDefault();
    e.target.click();
  }
});
root.addEventListener('input', (e) => { // Find a name on Who's in it?: the list narrows in place as the cadre types
  if (e.target.id !== 'ofind') return;
  ui.find = e.target.value;
  const q = ui.find.trim().toLowerCase();
  let any = false;
  root.querySelectorAll('.nrow[data-name]').forEach((b) => {
    b.hidden = !!q && b.dataset.name.indexOf(q) < 0;
    any = any || !b.hidden;
  });
  root.querySelectorAll('.names [data-plt]').forEach((x) => {
    x.hidden = !!q;
  });
  const none = document.getElementById('onone');
  if (none) none.hidden = any || !q;
});
if (window.visualViewport) { // the keyboard: a sheet stays above it, as on the OC page
  const vv = window.visualViewport;
  const kb = () => document.documentElement.style.setProperty('--kb', Math.max(0, window.innerHeight - vv.height - vv.offsetTop) + 'px');
  vv.addEventListener('resize', kb);
  vv.addEventListener('scroll', kb);
}
document.body.addEventListener('click', (e) => {
  const b = e.target.closest('.scrim [data-act]');
  if (!b) return;
  const fn = act[b.dataset.act];
  if (fn) fn(b);
});
root.addEventListener('dragover', (e) => {
  const d = e.target.closest('[data-drop]');
  if (!d) return;
  e.preventDefault();
  d.classList.add('over');
});
root.addEventListener('dragleave', (e) => {
  const d = e.target.closest('[data-drop]');
  if (d) d.classList.remove('over');
});
root.addEventListener('drop', (e) => {
  const d = e.target.closest('[data-drop]');
  if (!d) return;
  e.preventDefault();
  const files = [...(e.dataTransfer ? e.dataTransfer.files : [])];
  if (d.dataset.drop === 'roster') addFiles(files);
  else importFiles(files);
});
document.addEventListener('paste', (e) => {
  if (ui.screen !== 'roster' && ui.screen !== 'read') return;
  if (e.target && /^(INPUT|TEXTAREA)$/.test(e.target.tagName)) return;
  const text = e.clipboardData ? e.clipboardData.getData('text/plain') : '';
  if (!text || /^\s*https?:\/\/\S+\s*$/.test(text)) return;
  e.preventDefault();
  addSource(readRoster(text), 'Pasted list', null);
});
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible' && R && X) poll(true);
});
window.addEventListener('pagehide', () => {
  (R && !R.removed ? R.imported || [] : []).forEach((env) => { // added results not on the sheet yet: a shared device keeps nothing
    try {
      fetch(collectorUrl(), { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(env), keepalive: true }).catch(() => {});
    } catch (err) {
      // sent the next time it opens here
    }
  });
  if (!R || !X || !X.sealed) return;
  try {
    fetch(collectorUrl(), { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ op: 'board', e: R.e, w: R.wt, brev: X.sealed.brev, board: X.sealed.board, st: X.sealed.st || 0 }), keepalive: true }).catch(() => {});
  } catch (err) {
    // the board is sent the next time this page opens
  }
});
window.addEventListener('hashchange', () => location.reload());

/** What only this tab holds: evaluations a shared device doesn't keep, one not saved yet, a board not sent, a Start on its way or waiting (to read the squads, or to try again), on any evaluation. */
const inHand = () => (!own && list.length > 0) || starting.size > 0 || waits.size > 0 || !!(R && !R.brev) || !!(X && X.sealed) || !!(R && (R.imported || []).length);

/** A new version took over: it loads on Home or the list, once nothing is in hand (else at the next visit there). It opens on that screen again, with its toast. */
function reloadIfIdle() {
  if (!ui.update || !drawn || (ui.screen !== 'home' && ui.screen !== 'all') || inHand()) return;
  try {
    const t = document.querySelector('.toast');
    sessionStorage.setItem(AFTER, JSON.stringify({ screen: ui.screen, toast: t && t.firstChild ? t.firstChild.textContent : '' }));
  } catch (err) {
    // no session storage: it opens as on any visit
  }
  location.reload();
}

async function boot() {
  if ('serviceWorker' in navigator && /^https?:/.test(location.protocol)) {
    const before = navigator.serviceWorker.controller; // the version this page runs
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (!before) return; // the first install: this page is that version
      ui.update = true;
      reloadIfIdle();
    });
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  list = restore();
  if (!collectorUrl()) {
    root.innerHTML = head() + '<div class="main center"><h1>Peer evaluation</h1><p class="body">This site isn\'t connected to its sheet yet. The person who keeps the site sets it up once.</p>' + link('admin', 'Use my own Google sheet instead') + '</div>';
    return undefined;
  }
  let c = parseCadre(location.hash);
  try {
    if (c) sessionStorage.setItem(LINK, location.hash);
    else if ((performance.getEntriesByType('navigation')[0] || {}).type === 'reload') c = parseCadre(sessionStorage.getItem(LINK) || ''); // a reload of this tab opens it again
    else sessionStorage.removeItem(LINK); // opened afresh: the tab has left that link
  } catch (err) {
    // no session storage: a reload starts on the first screen
  }
  let after = null;
  try {
    after = JSON.parse(sessionStorage.getItem(AFTER) || 'null');
    sessionStorage.removeItem(AFTER); // read once, on any open: a link opened meanwhile never leaves it for a later visit
  } catch (err) {
    after = null;
  }
  if (c) {
    history.replaceState(null, '', location.pathname + location.search); // out of the address bar: it can't print in a page footer or show on a projector
    return openFromLink(c);
  }
  if (after) { // reloaded for a new version on Home or the list: there again
    go(after.screen === 'all' && list.length ? 'all' : 'home');
    if (after.toast) toast(after.toast);
    return undefined;
  }
  if (list.length === 1) return openEval(list[0]); // removed or not: one evaluation opens straight in
  if (list.length) return go('all');
  return go('home');
}

boot();
