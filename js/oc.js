/**
 * OC (rater) app.
 *
 * Flow: 4 Yes/No screens -> one screen per dealt attribute -> Rank (= Submit) -> Done.
 * Every screen moves on by itself 0.4 s after its last answer. Every tap is saved
 * in localStorage; reopening the link resumes at the first unfinished screen.
 * Submit seals the result for the admin and sends it to the collector when online;
 * otherwise it waits and retries, with a rotating QR, a file and plain text as
 * fallbacks that also work on locked-down networks.
 *
 * A squad sign-up link (Ask OCs, one per PLT) opens two screens instead: tap your name,
 * tap your squad. The pick is sealed and sent like a result; one phone counts once, by
 * its newest pick.
 *
 * Cadre mode's one link (#o2.<id>.<link key>) opens the sealed board instead: tap your name
 * (and your squad while OCs choose), wait for the cadre's Start, then the evaluation as above,
 * with the packet worked out on this phone from the board (lib/board.js).
 */
import { QUESTIONS, ATTRIBUTES, SCALE_SHORT } from './lib/attributes.js';
import { readPacket, packPacket, sealResult, sealJoin, envelopeToText } from './lib/envelope.js';
import { splitParts } from './lib/qrparts.js';
import { randomToken } from './lib/crypto.js';
import { parseOc, openBoard, packetOf, joinToken, fullName, ord as nth } from './lib/board.js';
import { getBoard, getState } from './lib/collector2.js';
import { collectorUrl, testQuery } from './lib/config.js';
import { squadBook, marksOf, readForms, matchName, formAnswers, XLSX_TYPE, QUESTION_LABEL } from './lib/report.js';
import { readXlsx } from './lib/xlsx.js';

const ADVANCE_MS = 400;
const root = document.getElementById('app');
let screenEl;
let sheetEl;
let kbEl;

let P = null;
let S = null;
let KEY = '';
let storageOk = true;
const ui = { sheet: null, sel: -1, pending: false, opening: false, mine: null }; // opening: a reason card opens by itself in a moment; mine: the answers sent, in memory only (Save my answers)
const timers = { adv: 0, open: 0, retry: 0, qr: 0 };
let sending = false;
let attempt = 0;

const IN_APP = /FBAN|FBAV|Instagram|Line\/|GSA\/|Snapchat|LinkedInApp|GroupMe|Twitter|KAKAOTALK|; wv\)/i.test(navigator.userAgent);

const ICON_BACK = '<svg width="20" height="20" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#17191C" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 5 L8 12 L15 19"/></svg>';
const MARK_COMMENT = '<svg class="mk" width="16" height="16" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#3F5F33" stroke-width="2" stroke-linejoin="round"><path d="M4 5 H20 V16 H10 L5 20 V16 H4 Z"/></svg>';
const ICON_SHEET = '<svg width="20" height="20" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#3F5F33" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4 V15 M7.5 10.5 L12 15 L16.5 10.5 M5 19.5 H19"/></svg>';
const ICON_CHECK = '<svg width="30" height="30" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#FFFFFF" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5 L10 17 L19 7"/></svg>';
const ICON_CLOCK = '<svg width="30" height="30" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#17191C" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="8"/><path d="M12 7.5 V12 L15 14"/></svg>';

const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const nameOf = (id) => (P.p.find((x) => x[0] === id) || [0, '?'])[1];
const peerOf = (i) => P.p[i][0];
/** A round badge with a peer's initials (first and last), toned by the name so each peer keeps one color. */
function avatar(nm) {
  const [last, first] = String(nm).split(', ');
  const ini = (first ? first.charAt(0) + last.charAt(0) : last.slice(0, 2)).toUpperCase();
  let h = 0;
  for (const c of String(nm)) h = (h * 31 + c.charCodeAt(0)) >>> 0;
  return '<span class="av tone' + (h % 6) + '" aria-hidden="true">' + esc(ini) + '</span>';
}
const ord = (n) => n + (['th', 'st', 'nd', 'rd'][(n % 100 - 20) % 10] || ['th', 'st', 'nd', 'rd'][n % 100] || 'th');

/* ---------- state ---------- */

function fresh() {
  return { v: 1, step: 'q0', seen: {}, ans: [{}, {}, {}, {}], com: [{}, {}, {}, {}], sc: {}, why: {}, rank: null, rev: 0, outbox: null };
}

/**
 * A random id for this browser, kept apart from the answers and sent inside the sealed
 * result: the admin sees when one OC's results come from two devices (someone else
 * using their link). It says nothing about the device itself.
 */
function deviceId() {
  try {
    let id = localStorage.getItem('pe:device');
    if (!id) {
      id = randomToken(9);
      localStorage.setItem('pe:device', id);
    }
    return id;
  } catch (err) {
    return '';
  }
}

function load() {
  try {
    const raw = localStorage.getItem(KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (err) {
    storageOk = false;
    return null;
  }
}

function save() {
  try {
    localStorage.setItem(KEY, JSON.stringify(S));
  } catch (err) {
    storageOk = false;
  }
}

/** By item (the default): one question or attribute, every peer, per screen. By person: one peer, everything, per screen. The answers are the same either way. */
function steps() {
  const mid = S.mode === 'person' ? P.p.map((x, i) => 'p' + i) : ['q0', 'q1', 'q2', 'q3'].concat(Array.from({ length: kOf() }, (x, j) => 'a' + j));
  return mid.concat(['rank', 'done']);
}

const stepIndex = () => steps().indexOf(S.step);
/** 17e on (the user, 11-24): a set of different attributes for each peer (P.a = { peer id: list }); before, one list for every peer. */
const pairs = () => !Array.isArray(P.a);
/** The OC's attributes for a peer: the dealt ones (P.a), with any swapped for another (S.hand, the same shape). */
function handFor(id) {
  if (!pairs()) return Array.isArray(S.hand) && S.hand.length === P.a.length ? S.hand : P.a;
  const dealt = P.a[id] || [];
  const own = S.hand && !Array.isArray(S.hand) ? S.hand[id] : null;
  return Array.isArray(own) && own.length === dealt.length ? own : dealt;
}
/** Attributes per peer (6 from 17e): the screens a0, a1, … by item. */
const kOf = () => (pairs() ? Math.max(0, ...P.p.map(([id]) => handFor(id).length)) : P.a.length);
const attrOf = (j, id) => handFor(id)[j];
/** Every hand as it is now, to keep with the answers: one list, or { peer id: list }. */
const handsNow = () => (pairs() ? Object.fromEntries(P.p.map(([id]) => [id, handFor(id).slice()])) : handFor().slice());
const scOf = (a) => (S.sc[a] = S.sc[a] || {});
const whyOf = (a) => (S.why[a] = S.why[a] || {});
const answered = (q, id) => S.ans[q][id] === 'Y' || S.ans[q][id] === 'N';
const hasWhy = (box, id) => !!String(box[id] || '').trim();
/* Every answer has a reason: every score (11-13, "打多少分都要写") and every Yes and No (11-24, "最开始的yes no也必须写理由"). */
const qAnswered = (q) => P.p.every(([id]) => answered(q, id));
const qDone = (q) => P.p.every(([id]) => answered(q, id) && hasWhy(S.com[q], id));
const allScored = (j) => P.p.every(([id]) => !!scOf(attrOf(j, id))[id]);
const aDone = (j) => P.p.every(([id]) => !!scOf(attrOf(j, id))[id] && !!String(whyOf(attrOf(j, id))[id] || '').trim());
const pAnswered = (id) => QUESTIONS.every((q, i) => answered(i, id)) && handFor(id).every((a) => !!scOf(a)[id]);
const pDone = (id) => QUESTIONS.every((q, i) => answered(i, id) && hasWhy(S.com[i], id)) && handFor(id).every((a) => !!scOf(a)[id] && !!String(whyOf(a)[id] || '').trim());

function stepDone(step) {
  if (step[0] === 'q') return qDone(Number(step.slice(1)));
  if (step[0] === 'a') return aDone(Number(step.slice(1)));
  if (step[0] === 'p') return pDone(peerOf(Number(step.slice(1))));
  return false;
}

function firstOpen() {
  const list = steps();
  for (let i = 0; i < list.length - 2; i++) if (!stepDone(list[i])) return list[i];
  return 'rank';
}

function clearTimers() {
  clearTimeout(timers.adv);
  clearTimeout(timers.open);
  ui.pending = false;
}

/* ---------- navigation ---------- */

function go(step) {
  clearTimers();
  closeSheet();
  ui.sel = -1;
  S.step = step;
  save();
  render();
  window.scrollTo(0, 0);
}

/** On to the next screen; from the last one, back to the first screen still unfinished (One peer: a swap leaves the peers done before without the new attribute), Rank only when all are done. */
const goNext = () => {
  const nx = steps()[stepIndex() + 1];
  go(nx === 'rank' && firstOpen() !== 'rank' ? firstOpen() : nx);
};
const goBack = () => go(steps()[Math.max(0, stepIndex() - 1)]);

/** After an answer: advance once, the first time the screen is complete. */
function afterChange() {
  const st = S.step;
  if (stepDone(st) && !S.seen[st]) {
    S.seen[st] = true;
    save();
    ui.pending = true;
    render();
    clearTimeout(timers.adv);
    timers.adv = setTimeout(() => {
      ui.pending = false;
      if (S.step === st) goNext();
    }, ADVANCE_MS);
    return;
  }
  render();
}

/* ---------- rendering ---------- */

function top() {
  const list = steps();
  const i = stepIndex();
  const n = list.length;
  const back = i > 0 && S.step !== 'done'
    ? '<button class="back" data-act="back" aria-label="Back">' + ICON_BACK + '</button>'
    : '<span class="spacer44" aria-hidden="true"></span>';
  return '<div class="top">' + back +
    '<div class="bar" role="progressbar" aria-label="Progress" aria-valuemin="0" aria-valuemax="' + n + '" aria-valuenow="' + (i + 1) + '">' +
    '<i style="width:' + (((i + 1) / n) * 100).toFixed(1) + '%"></i></div>' +
    '<button class="tool" data-act="xlsx" aria-label="Save my answers so far (Excel)">' + ICON_SHEET + '</button></div>'; // any time, on the original paper form
}

function banners() {
  let h = '';
  if (!storageOk) h += '<p class="banner">This browser can\'t save your answers, so finish in one sitting. A normal Safari, Chrome or Edge window works best.</p>';
  if (IN_APP) h += '<p class="banner">You opened this inside another app. Open the link in Safari or Chrome so your answers are kept.</p>';
  return h;
}

/** The first screen's line: how many, how long, who sees it. */
function intro() {
  if (stepIndex() !== 0) return '';
  const n = P.p.length;
  const k = kOf();
  const min = Math.max(5, Math.round(((4 * n + n * k) * 20 + (8 * n + 2 * n * k + 5) * 1.5) / 300) * 5); // a line (~20 s) for every answer and score, plus the taps
  return '<p class="sub">' + n + ' peers · about ' + min + ' min · saves as you go · only cadre see your answers</p>' +
    '<button class="link left" data-act="load-xlsx">Filled in the Excel form? Load it</button>';
}

/** The OC's choice, any time: everyone on one item per screen, or one person per screen. */
function modeSwitch() {
  const p = S.mode === 'person';
  return '<div class="mode" role="group" aria-label="How to rate"><button data-act="mode" data-v="item" aria-pressed="' + !p + '">All peers</button><button data-act="mode" data-v="person" aria-pressed="' + p + '">One peer</button></div>';
}

function renderQ(q) {
  const st = S.step;
  let h = top() + banners() + modeSwitch() + intro();
  h += '<p class="kicker">Question ' + (q + 1) + ' of ' + QUESTIONS.length + '</p><h1>' + esc(QUESTIONS[q].text) + '</h1>';
  const all = qAnswered(q);
  if (q === 0) h += '<p class="sub">Each Yes or No needs a one-line reason.</p>';
  if (q === 0 && P.one && !all) h += '<p class="me">You\'re ' + esc(P.nm) + '. <button class="link" data-act="one-notme">Not you? Go back</button></p>';
  h += '<div class="list">';
  P.p.forEach(([id, nm]) => {
    const a = S.ans[q][id];
    const c = hasWhy(S.com[q], id);
    const needs = all && !c;
    h += '<div class="row">' +
      '<button class="name' + (needs ? ' needs' : '') + '" data-act="qname" data-id="' + esc(id) + '" aria-label="' + esc(nm + (needs ? ', needs a reason' : '') + '. Open the reason card') + '">' + avatar(nm) + '<span class="t">' + esc(nm) + '</span>' + (c ? MARK_COMMENT : '') + '</button>' +
      '<span class="yn2"><button class="yn yes" data-act="yes" data-id="' + esc(id) + '" aria-pressed="' + (a === 'Y') + '" aria-label="' + esc(nm) + ': Yes">Yes</button>' +
      '<button class="yn no" data-act="no" data-id="' + esc(id) + '" aria-pressed="' + (a === 'N') + '" aria-label="' + esc(nm) + ': No">No</button></span>' +
      '</div>';
  });
  h += '</div>';
  return h + stepFoot(qDone(q), all);
}

/** Next once every answer has its reason; Say why while one is missing and the card is closed (a reason for every answer). */
function stepFoot(done, all) {
  if (done && S.seen[S.step] && !ui.pending) return '<button class="primary" data-act="next">Next</button>';
  if (all && !done && !ui.sheet && !ui.opening) return '<button class="primary" data-act="why-next">Say why · ' + cardsHere().filter(needsWhy).length + ' left</button>';
  return '';
}

function renderA(j) {
  const st = S.step;
  const each = pairs(); // 17e: every peer has their own attribute in this slot, named under their name
  const scored = allScored(j);
  let h = top() + banners() + modeSwitch() + '<p class="kicker">Attribute ' + (j + 1) + ' of ' + kOf() + '</p>';
  if (each) h += '<h1>Score each peer</h1><p class="sub">Each peer has their own attribute, under their name.</p>';
  else {
    const attr = ATTRIBUTES[attrOf(j)];
    h += '<h1>' + esc(attr.name) + '</h1><p class="sub">' + esc(attr.def) + '</p><button class="link swap" data-act="swap" data-j="' + j + '">Swap for another attribute</button>';
  }
  h += '<div class="list">' + LEGEND;
  P.p.forEach(([id, nm]) => {
    const a = attrOf(j, id);
    const what = each ? nm + ', ' + ATTRIBUTES[a].name : nm;
    const v = scOf(a)[id];
    const needs = !!v && !String(whyOf(a)[id] || '').trim() && scored;
    h += '<div class="row arow ia">' + // by item: the name wraps beside the dots, no badge, so a row stays one line
      '<button class="name' + (needs ? ' needs' : '') + '" data-act="aname" data-id="' + esc(id) + '" aria-label="' + esc(what + (needs ? ', needs a reason' : '') + '. Open the reason card') + '">' +
      (each ? '<span class="lab"><span class="t">' + esc(nm) + '</span><small>' + esc(ATTRIBUTES[a].name) + '</small></span>' : '<span class="t">' + esc(nm) + '</span>') + '</button>' +
      '<div class="dots">';
    for (let x = 1; x <= 5; x++) {
      h += '<button class="dot" data-act="score" data-id="' + esc(id) + '" data-v="' + x + '" aria-pressed="' + (v === x) + '" aria-label="' + esc(what + ', ' + x + ', ' + SCALE_SHORT[x]) + '"><span>' + x + '</span></button>';
    }
    h += '</div></div>';
  });
  h += '</div>';
  return h + stepFoot(aDone(j), scored);
}

/** By person: one peer's four questions and every attribute dealt, on one screen; a reason card for each score, as by item. */
function renderP(i) {
  const st = S.step;
  const [id, nm] = P.p[i];
  const scored = pAnswered(id); // the 4 questions answered and every attribute scored: the card walks the reasons
  let h = top() + banners() + modeSwitch() + intro() +
    '<p class="kicker">Person ' + (i + 1) + ' of ' + P.p.length + '</p><div class="who">' + avatar(nm) + '<h1>' + esc(nm) + '</h1></div>';
  if (i === 0 && P.one && !scored) h += '<p class="me">You\'re ' + esc(P.nm) + '. <button class="link" data-act="one-notme">Not you? Go back</button></p>';
  h += '<p class="sub">Each answer and score needs a one-line reason.</p><div class="list">';
  QUESTIONS.forEach((q, qi) => {
    const a = S.ans[qi][id];
    const c = hasWhy(S.com[qi], id);
    const needs = scored && !c;
    h += '<div class="row">' +
      '<button class="name qtext' + (needs ? ' needs' : '') + '" data-act="qname" data-q="' + qi + '" aria-label="' + esc(q.text + (needs ? ', needs a reason' : '') + '. Open the reason card') + '"><span class="t">' + esc(q.text) + '</span>' + (c ? MARK_COMMENT : '') + '</button>' +
      '<span class="yn2"><button class="yn yes" data-act="yes" data-q="' + qi + '" aria-pressed="' + (a === 'Y') + '" aria-label="' + esc(q.text + ' Yes') + '">Yes</button>' +
      '<button class="yn no" data-act="no" data-q="' + qi + '" aria-pressed="' + (a === 'N') + '" aria-label="' + esc(q.text + ' No') + '">No</button></span>' +
      '</div>';
  });
  h += '</div><div class="list">' + LEGEND;
  handFor(id).forEach((a, j) => {
    const v = scOf(a)[id];
    const needs = !!v && !String(whyOf(a)[id] || '').trim() && scored;
    const attr = ATTRIBUTES[a];
    h += '<div class="row arow">' +
      '<button class="name' + (needs ? ' needs' : '') + '" data-act="aname" data-j="' + j + '" aria-label="' + esc(attr.name + (needs ? ', needs a reason' : '') + '. Open the reason card') + '"><span class="lab"><span class="t">' + esc(attr.name) + '</span><small>' + esc(attr.def) + '</small></span></button>' +
      '<div class="dots">';
    for (let x = 1; x <= 5; x++) {
      h += '<button class="dot" data-act="score" data-j="' + j + '" data-v="' + x + '" aria-pressed="' + (v === x) + '" aria-label="' + esc(attr.name + ', ' + x + ', ' + SCALE_SHORT[x]) + '"><span>' + x + '</span></button>';
    }
    h += '</div></div>';
  });
  h += '</div><button class="link swap" data-act="swap">Swap an attribute</button>';
  return h + stepFoot(pDone(id), scored);
}

/** The scale over the dots, the first row of the score card: far below over 1, meets over 3, far exceeds to the edge of 5. */
const LEGEND = '<div class="legend" aria-hidden="true"><div><span>far below</span><span>meets</span><span>far exceeds</span></div></div>';

function totalOf(id) {
  return handFor(id).reduce((s, a) => s + (Number(scOf(a)[id]) || 0), 0);
}

function yesOf(id) {
  return S.ans.reduce((s, m) => s + (m[id] === 'Y' ? 1 : 0), 0);
}

function ensureRank() {
  const ids = P.p.map((x) => x[0]);
  const ok = Array.isArray(S.rank) && S.rank.length === ids.length && ids.every((id) => S.rank.indexOf(id) >= 0);
  if (ok && S.moved) return; // the OC's own order stays; until he moves someone, it follows his ratings (a score changed via Back re-sorts it)
  S.tb = S.tb || {}; // one random draw per peer, kept: ties are never sorted A-Z, which favored early names
  ids.forEach((id) => {
    if (typeof S.tb[id] !== 'number') S.tb[id] = Math.random();
  });
  const next = ids.slice().sort((x, y) => totalOf(y) - totalOf(x) || yesOf(y) - yesOf(x) || S.tb[x] - S.tb[y]);
  if (!ok || next.join() !== S.rank.join()) {
    S.rank = next;
    save();
  }
}

/** The ratings behind the pre-sort, to tell when they changed after the OC's own order. */
const rankSig = () => P.p.map(([id]) => totalOf(id) + ':' + yesOf(id)).join(',');

function renderRank() {
  ensureRank();
  const tieAt = (j) => (j > 0 && totalOf(S.rank[j - 1]) === totalOf(S.rank[j])) || (j < S.rank.length - 1 && totalOf(S.rank[j + 1]) === totalOf(S.rank[j]));
  const hint = ui.sel >= 0 ? 'Now tap the new place for ' + esc(nameOf(S.rank[ui.sel])) + '.'
    : S.moved && S.rankSig !== rankSig() ? '1 = best. Your order (scores changed since: check it).'
      : S.rank.some((id, j) => tieAt(j)) ? '1 = best. Put the ties (=) in order: tap one, then its place.' : '1 = best. Pre-sorted by your ratings. Tap to move.';
  let h = top() + banners() + '<p class="kicker">Last step</p><h1>Rank your squad</h1><p class="sub">' + hint + '</p><div class="list">';
  S.rank.forEach((id, j) => {
    const t = totalOf(id);
    const tie = (j > 0 && totalOf(S.rank[j - 1]) === t) || (j < S.rank.length - 1 && totalOf(S.rank[j + 1]) === t);
    let aria = 'Place ' + (j + 1) + ', ' + nameOf(id) + (tie ? ', tied with a neighbor on your ratings' : '') + '. Tap to move';
    if (ui.sel >= 0) aria = ui.sel === j ? 'Cancel moving ' + nameOf(id) : 'Put ' + nameOf(S.rank[ui.sel]) + ' in place ' + (j + 1);
    h += '<div class="row"><button class="rankrow' + (ui.sel === j ? ' sel' : '') + '" data-act="rk" data-j="' + j + '" aria-label="' + esc(aria) + '">' +
      '<span class="pos">' + (j + 1) + '</span><span class="t">' + esc(nameOf(id)) + (tie ? ' <span class="tie">=&nbsp;tie</span>' : '') + '</span>' +
      '</button></div>';
  });
  h += '</div><button class="primary gap24" data-act="submit">Submit</button>';
  return h;
}

function renderDone() {
  const ob = S.outbox;
  const again = S.cleared
    ? '<button class="link" data-act="restart">Start over</button>'
    : '<button class="link" data-act="change">Change answers</button>';
  const xlsx = mine() ? '<button class="link" data-act="xlsx">Save my answers (Excel)</button>' : '';
  if (ob && ob.sentAt) {
    return banners() + '<div class="done"><span class="badge">' + ICON_CHECK + '</span><h1>Sent to your cadre</h1><p class="sub">You\'re done. You can close this page.</p></div>' +
      '<div class="links">' + xlsx + again + '</div>';
  }
  let line = 'Send it as a file or copy it as text to your cadre, or show them this code.';
  if (P.u && ob && ob.blocked) line = 'The sheet did not take it. Send it as a file or copy it as text to your cadre.';
  else if (P.u) line = (navigator.onLine === false ? 'It sends by itself when you have signal.' : 'Not sent yet. It keeps trying by itself.') + ' Don\'t close this page until it says Sent to your cadre.';
  return banners() + '<div class="done"><span class="badge wait">' + ICON_CLOCK + '</span>' +
    '<h1>' + (P.u ? 'Saved on this device' : 'Ready for your cadre') + '</h1><p class="sub">' + line + '</p>' +
    '<div class="qr" id="qr" role="img" aria-label="Rotating QR code with your encrypted answers"></div>' +
    '<p class="sub" id="qrpart"></p></div>' +
    '<div class="links"><button class="link" data-act="file">Send as a file</button>' +
    '<button class="link" data-act="copy">Copy as text</button>' + xlsx + again + '</div>';
}

function render() {
  clearInterval(timers.qr);
  if (ONE && ONE.screen) {
    screenEl.innerHTML = renderOne();
    return;
  }
  if (P.j) {
    screenEl.innerHTML = renderJoin();
    return;
  }
  const st = S.step;
  let h;
  if (st[0] === 'q') h = renderQ(Number(st.slice(1)));
  else if (st[0] === 'a') h = renderA(Number(st.slice(1)));
  else if (st[0] === 'p') h = renderP(Number(st.slice(1)));
  else if (st === 'rank') h = renderRank();
  else h = renderDone();
  screenEl.innerHTML = h;
  if (st === 'done' && S.outbox && !S.outbox.sentAt) startQr();
}

function startQr() {
  const el = document.getElementById('qr');
  const label = document.getElementById('qrpart');
  if (!el || typeof qrcode !== 'function') return;
  const parts = splitParts(envelopeToText(S.outbox.env));
  let i = 0;
  const show = () => {
    const qr = qrcode(0, 'L');
    qr.addData(parts[i], 'Byte');
    qr.make();
    el.innerHTML = qr.createSvgTag({ cellSize: 4, margin: 2, scalable: true });
    if (label) label.textContent = parts.length > 1 ? 'Code ' + (i + 1) + ' of ' + parts.length : '';
    i = (i + 1) % parts.length;
  };
  show();
  if (parts.length > 1) timers.qr = setInterval(show, 450);
}

/* ---------- sheets ---------- */

function closeSheet() {
  clearTimeout(timers.open);
  ui.opening = false;
  if (!ui.sheet) return;
  ui.sheet = null;
  sheetEl.hidden = true;
  sheetEl.innerHTML = '';
}

/** A reason still being typed, kept with the answers at every key: a page closed mid-sentence loses nothing (the user, 11-24: "Yes, keep it"). */
function keepDraft() {
  const s = ui.sheet;
  if (!s || (s.kind !== 'q' && s.kind !== 'a')) return;
  const text = document.getElementById('sheet-input').value;
  S.draft = text.trim() ? { kind: s.kind, k: s.k, id: s.id, text } : null;
  save();
}
const draftFor = (s) => (S.draft && S.draft.kind === s.kind && S.draft.k === s.k && S.draft.id === s.id ? S.draft.text : null);

/** Back on the page with a reason half typed: its card opens again, the words in it. */
function resumeDraft() {
  const d = S.draft;
  if (!d || S.step === 'done' || ui.sheet) return;
  if (cardsHere().some((c) => c.kind === d.kind && c.k === d.k && c.id === d.id)) openCard({ kind: d.kind, k: d.k, id: d.id });
}

function showSheet(html) {
  sheetEl.innerHTML = '<button class="scrim" data-act="sheet-close" aria-label="Close"></button>' + html;
  sheetEl.hidden = false;
  const input = document.getElementById('sheet-input');
  if (input) {
    input.addEventListener('input', () => {
      keepDraft();
      updateSaveState();
    });
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        onSave();
      }
    });
    input.focus({ preventScroll: true });
  }
}

/**
 * Swap an attribute the app dealt for another: which one first (By person), then for which. Its scores and
 * reasons so far go with it. Before 17e it changes for every peer; from 17e (a set per peer) for this peer only, in One peer.
 */
function openSwap(j) {
  clearTimers();
  const id = pairs() ? peerOf(Number(S.step.slice(1))) : null;
  const hand = handFor(id);
  ui.sheet = { kind: 'swap', j, id };
  render();
  const who = id === null ? 'It changes for every peer.' : 'It changes for ' + nameOf(id) + ' only.';
  const opt = (act, v, a) => '<button class="opt" data-act="' + act + '" data-v="' + v + '"><b>' + esc(ATTRIBUTES[a].name) + '</b><span class="sub">' + esc(ATTRIBUTES[a].def) + '</span></button>';
  const h = j === null
    ? '<div class="head"><b>Swap which attribute?</b></div><p class="sub">' + esc(who) + '</p><div class="opts">' + hand.map((a, jj) => opt('swap-from', jj, a)).join('') + '</div>'
    : '<div class="head"><b>Swap ' + esc(ATTRIBUTES[hand[j]].name) + ' for</b></div><p class="sub">' + ((id === null ? P.p.some(([x]) => scOf(hand[j])[x]) : scOf(hand[j])[id]) ? 'Its scores and reasons so far are cleared. ' : '') + esc(who) + '</p>' +
      '<div class="opts">' + ATTRIBUTES.map((x, i) => i).filter((i) => hand.indexOf(i) < 0).map((i) => opt('swap-to', i, i)).join('') + '</div>';
  showSheet('<div class="card swapcard" role="dialog" aria-label="Swap an attribute">' + h + '<button class="link" data-act="sheet-close">Cancel</button></div>');
}

/** Swap slot j of the hand for attribute a: for every peer, or (17e) for the peer the swap card was opened for. */
function swapTo(j, a) {
  const id = ui.sheet && ui.sheet.id !== undefined ? ui.sheet.id : null;
  const hand = handFor(id).slice();
  const old = hand[j];
  hand[j] = a;
  if (id === null) {
    S.hand = hand;
    delete S.sc[old];
    delete S.why[old];
    delete S.seen['a' + j];
  } else {
    S.hand = Object.assign({}, S.hand && !Array.isArray(S.hand) ? S.hand : {}, { [id]: hand });
    delete scOf(old)[id];
    delete whyOf(old)[id];
    delete S.seen[S.step];
  }
  save();
  closeSheet();
  render();
  toast(ATTRIBUTES[old].name + ' swapped for ' + ATTRIBUTES[a].name + (id === null ? '.' : ' for ' + nameOf(id) + '.'));
}

/* ---------- the reason card: one per answer, kind q (a question) or a (an attribute), k its index, one peer ---------- */

const whyBox = (kind, k, id) => (kind === 'q' ? S.com[k] : whyOf(attrOf(k, id)));
const answerOf = (kind, k, id) => (kind === 'q' ? (answered(k, id) ? S.ans[k][id] : '') : scOf(attrOf(k, id))[id] || 0);
const needsWhy = (c) => !!answerOf(c.kind, c.k, c.id) && !hasWhy(whyBox(c.kind, c.k, c.id), c.id);

/** This screen's cards in walk order: by item its peers; by person the peer's 4 questions, then the attributes. */
function cardsHere() {
  const st = S.step;
  const n = Number(st.slice(1));
  if (st[0] === 'q' || st[0] === 'a') return P.p.map(([id]) => ({ kind: st[0], k: n, id }));
  if (st[0] !== 'p') return [];
  const id = peerOf(n);
  return QUESTIONS.map((q, i) => ({ kind: 'q', k: i, id })).concat(handFor(id).map((a, j) => ({ kind: 'a', k: j, id })));
}

/** The next card after c (the first when c is null) with an answer still without a reason. */
function nextCard(c) {
  const list = cardsHere();
  const at = c ? list.findIndex((x) => x.kind === c.kind && x.k === c.k && x.id === c.id) : -1;
  for (let d = 1; d <= list.length; d++) {
    const x = list[(at + d + list.length) % list.length];
    if (needsWhy(x)) return x;
  }
  return null;
}

function openCard(c) {
  clearTimeout(timers.open);
  ui.sheet = { kind: c.kind, k: c.k, id: c.id };
  showSheet('<div class="card" role="dialog" aria-labelledby="card-title">' +
    '<div class="head"><b id="card-title"></b><span id="card-pos"></span></div>' +
    '<div class="seg" id="card-seg" role="group"></div>' +
    '<label class="sub" for="sheet-input" id="sheet-label"></label>' +
    '<input id="sheet-input" type="text" enterkeyhint="next" autocomplete="off" maxlength="300" placeholder="One example">' +
    '<p class="sub tip" id="sheet-tip">' + MIC_TIP + '</p>' +
    '<button class="primary" id="sheet-save" data-act="sheet-save"></button></div>');
  fillCard(true);
}

/** Update the reason card in place so the input keeps focus (and the keyboard stays up). */
function fillCard(fresh) {
  const s = ui.sheet;
  const q = s.kind === 'q';
  const v = answerOf(s.kind, s.k, s.id);
  const list = cardsHere();
  document.getElementById('card-title').textContent = nameOf(s.id) + ' · ' + (q ? QUESTION_LABEL[QUESTIONS[s.k].id] : ATTRIBUTES[attrOf(s.k, s.id)].name);
  document.getElementById('card-pos').textContent = list.filter((x) => answerOf(x.kind, x.k, x.id) && hasWhy(whyBox(x.kind, x.k, x.id), x.id)).length + ' of ' + list.length + ' done';
  const seg = document.getElementById('card-seg');
  seg.className = 'seg' + (q ? ' two' : '');
  seg.setAttribute('aria-label', q ? 'Answer' : 'Score');
  seg.innerHTML = q ? [['Y', 'Yes'], ['N', 'No']].map(([x, t]) => '<button data-act="seg" data-v="' + x + '" aria-pressed="' + (v === x) + '">' + t + '</button>').join('')
    : [1, 2, 3, 4, 5].map((x) => '<button data-act="seg" data-v="' + x + '" aria-pressed="' + (v === x) + '" aria-label="Score ' + x + ', ' + SCALE_SHORT[x] + '">' + x + '</button>').join('');
  document.getElementById('sheet-label').textContent = !v ? (q ? 'Pick Yes or No, then say why' : 'Pick a score, then say why') : q ? 'Why ' + (v === 'Y' ? 'Yes' : 'No') + '?' : 'Why a ' + v + ' (' + SCALE_SHORT[v] + ')?';
  if (fresh) document.getElementById('sheet-input').value = draftFor(s) !== null ? draftFor(s) : String(whyBox(s.kind, s.k, s.id)[s.id] || ''); // the words typed last, saved or not
  updateSaveState();
}

const MIC_TIP = 'Tip: tap the mic on your keyboard and say it.';

/** The tip line under a reason: red with what's missing after an empty Save, the mic tip again on the next keystroke or answer. */
function sheetTip(text) {
  const tip = document.getElementById('sheet-tip');
  if (!tip) return;
  tip.textContent = text || MIC_TIP;
  tip.classList.toggle('err', !!text);
  if (text) document.getElementById('sheet-input').focus({ preventScroll: true }); // the keyboard stays up
}

function updateSaveState() {
  const s = ui.sheet;
  if (!s) return;
  const input = document.getElementById('sheet-input');
  const btn = document.getElementById('sheet-save');
  const text = input.value.trim();
  sheetTip('');
  btn.disabled = false; // never greyed out: an empty Save says what's missing
  const box = whyBox(s.kind, s.k, s.id);
  const valid = !!answerOf(s.kind, s.k, s.id) && !!text;
  const prev = box[s.id];
  box[s.id] = text || '-'; // the others only: is there a next one after this card
  const more = nextCard(s) !== null;
  box[s.id] = prev;
  btn.textContent = more ? 'Save & next' : 'Save';
  input.setAttribute('enterkeyhint', more ? 'next' : 'done');
}

/** Save from the button or Return: with an answer and a line it saves and walks on; otherwise the card stays and says what's missing. */
function onSave() {
  const s = ui.sheet;
  if (!s) return;
  const text = document.getElementById('sheet-input').value.trim();
  if (!answerOf(s.kind, s.k, s.id)) return sheetTip(s.kind === 'q' ? 'Pick Yes or No first.' : 'Pick a score first.');
  if (!text) return sheetTip('Type a few words first.');
  whyBox(s.kind, s.k, s.id)[s.id] = text;
  S.draft = null;
  save();
  const next = nextCard(s);
  if (next !== null) {
    Object.assign(s, next);
    fillCard(true);
    render();
    return;
  }
  closeSheet();
  afterChange();
}

/** Tap outside: keep what was typed if it is valid. */
function onClose() {
  const s = ui.sheet;
  if (!s) return;
  if (s.kind === 'swap') {
    closeSheet();
    render();
    return;
  }
  const text = document.getElementById('sheet-input').value.trim();
  if (answerOf(s.kind, s.k, s.id) && text) {
    whyBox(s.kind, s.k, s.id)[s.id] = text;
    S.draft = null;
    save();
    closeSheet();
    afterChange();
    return;
  }
  closeSheet();
  render();
}

/* ---------- submit and send ---------- */

function buildPayload() {
  const ans = {};
  const com = {};
  const sc = {};
  const why = {};
  P.p.forEach(([id]) => {
    ans[id] = QUESTIONS.map((q, i) => S.ans[i][id] || '-').join('');
    const c = {};
    QUESTIONS.forEach((q, i) => {
      const t = String(S.com[i][id] || '').trim();
      if (t) c[q.id] = t;
    });
    if (Object.keys(c).length) com[id] = c;
    sc[id] = {};
    why[id] = {};
    handFor(id).forEach((a) => {
      const v = scOf(a)[id];
      if (v) sc[id][a] = v;
      const w = String(whyOf(a)[id] || '').trim();
      if (w) why[id][a] = w;
    });
  });
  return { v: 1, e: P.e, r: P.me, rev: S.rev, at: new Date().toISOString(), ans, com, sc, why, rank: (S.rank || []).slice(), dv: deviceId() };
}

async function submit() {
  clearTimers();
  ui.sel = -1;
  ensureRank();
  S.rev = (S.rev || 0) + 1;
  let env;
  const p = buildPayload();
  try {
    env = await sealResult(P, p);
  } catch (err) {
    S.rev -= 1;
    toast('Could not prepare your answers on this browser. Try Safari, Chrome or Edge.');
    return;
  }
  S.outbox = { env, rev: S.rev, sentAt: null };
  S.step = 'done';
  ui.mine = { key: KEY, p, hand: handsNow() };
  save();
  render();
  window.scrollTo(0, 0);
  attempt = 0;
  trySend();
}

async function trySend() {
  clearTimeout(timers.retry);
  if (!S || !S.outbox || S.outbox.sentAt || S.outbox.blocked || !P.u || sending) return;
  sending = true;
  let ok = false;
  let refused = false;
  try {
    const res = await fetch(P.u, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(S.outbox.env),
      cache: 'no-store',
      redirect: 'follow'
    });
    const data = JSON.parse(await res.text());
    ok = !!(data && data.ok);
    refused = !ok && /^bad (size|json|envelope)$/.test(String(data && data.error));
  } catch (err) {
    ok = false;
  }
  sending = false;
  if (refused) {
    S.outbox.blocked = true;
    save();
    if (S.step === 'done' || P.j) render();
    return;
  }
  if (ok) {
    S.outbox.sentAt = new Date().toISOString();
    if (P.j) {
      save();
      render();
    } else if (S.step === 'done') {
      const rev = S.outbox.rev;
      S = Object.assign(fresh(), { step: 'done', cleared: true, rev: Math.max(S.rev, rev), outbox: { rev, sentAt: S.outbox.sentAt } });
      closeSheet();
      save();
      render();
    } else {
      save();
    }
    attempt = 0;
    return;
  }
  attempt += 1;
  const delay = Math.min(60000, 2000 * Math.pow(2, Math.min(attempt, 5))) * (0.5 + Math.random());
  timers.retry = setTimeout(trySend, delay);
}

/**
 * After the result went out as a file or text, keep only the sealed copy: it can be
 * sent again, and with a collector it still sends by itself later.
 */
function clearPlain() {
  if (S.cleared || !S.outbox) return;
  S = Object.assign(fresh(), { step: 'done', cleared: true, rev: S.rev, outbox: S.outbox });
  save();
  render();
}

function fileName() {
  return 'peer-eval-' + P.e + '-' + String(P.nm).replace(/[^A-Za-z0-9]+/g, '-') + '-v' + S.outbox.rev + '.txt';
}

/** The phone's share sheet for a file (Mail, Messages, Files…), or a download where it has none; false if the OC backed out. */
async function shareFile(name, data, type, tip) {
  try {
    const file = new File([data], name, { type });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], title: 'Peer evaluation' });
      return true;
    }
  } catch (err) {
    if (err && err.name === 'AbortError') return false;
  }
  const url = URL.createObjectURL(new Blob([data], { type }));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
  toast('Saved as ' + name + '. ' + tip);
  return true;
}

async function sendFile() {
  if (await shareFile(fileName(), envelopeToText(S.outbox.env) + '\n', 'text/plain', 'Email it to your cadre.')) clearPlain();
}

/** This OC's own answers in plain: from Submit until the page closes (the device keeps only the sealed copy once sent), or still on it before then. */
const mine = () => (S && !S.cleared && S.step !== 'done' ? { p: buildPayload(), hand: handsNow() } : null) // editing: what's on the phone now
  || (ui.mine && ui.mine.key === KEY ? ui.mine : null) || (S && !S.cleared ? { p: buildPayload(), hand: handsNow() } : null);

/** Save my answers (Excel): the original paper form of this squad with this OC's answers, to keep or give the cadre if the app has trouble. */
async function saveSheet() {
  const m = mine();
  if (!m) return render();
  const p = m.p;
  const num = (re) => Number((re.exec(P.g || '') || [])[1]) || 0;
  const split = (nm) => { const i = String(nm).indexOf(', '); return i < 0 ? { last: String(nm), first: '' } : { last: nm.slice(0, i), first: nm.slice(i + 2) }; };
  const ocs = P.p.map(([id, nm]) => Object.assign(split(nm), { id, m: marksOf(p, id) })); // id: each tab its own attributes (17e)
  const note = 'From the peer evaluation app: ' + P.nm + '\'s answers, ' + new Date(p.at).toLocaleString() + (S.step === 'done' || S.cleared ? '.' : ', not finished yet.');
  const f = squadBook({ plt: num(/(\d+)\w*\s+PLT/), sqd: num(/(\d+)\w*\s+SQD/) }, ocs, split(P.nm), note, m.hand);
  await shareFile(f.name, f.data, XLSX_TYPE, 'Keep it: your cadre may ask for it.');
}

/**
 * Load answers from this OC's filled-in Excel form (17h): the app's own Excel, or the blank sheet the cadre gave out.
 * The tabs for his squadmates fill the app (attributes swapped on the form too); he checks them, then Submits.
 */
function loadSheet() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  input.onchange = async () => {
    const file = input.files && input.files[0];
    if (!file) return;
    let forms = [];
    try {
      forms = readForms(await readXlsx(new Uint8Array(await file.arrayBuffer())));
    } catch (err) {
      forms = [];
    }
    const split = (nm) => { const i = String(nm).indexOf(', '); return i < 0 ? { last: String(nm), first: '' } : { last: nm.slice(0, i), first: nm.slice(i + 2) }; };
    const me = Object.assign({ id: 'me' }, split(P.nm));
    const theirs = forms.find((f) => f.rater && !matchName(f.rater, [me]));
    const mine = forms.filter((f) => !f.rater || matchName(f.rater, [me])); // no name picked: on his own phone, his
    const peers = P.p.map(([id, nm]) => Object.assign({ id }, split(nm)));
    const a = formAnswers(mine, peers);
    if (!a.n) return toast(theirs && !mine.length ? 'That form is ' + theirs.rater + '\'s, not yours.' : 'No answers for your squadmates in that file.');
    Object.keys(a.ans).forEach((key) => {
      const id = Number(key);
      String(a.ans[id]).split('').forEach((v, q) => {
        if (v === 'Y' || v === 'N') S.ans[q][id] = v;
      });
      QUESTIONS.forEach((q, i) => {
        if ((a.com[id] || {})[q.id]) S.com[i][id] = a.com[id][q.id];
      });
      const list = Object.keys(a.sc[id]).map(Number);
      if (pairs() && list.length === (P.a[id] || []).length) S.hand = Object.assign({}, S.hand && !Array.isArray(S.hand) ? S.hand : {}, { [id]: list }); // the form's own six, swaps included
      handFor(id).forEach((at) => {
        if (a.sc[id][at]) scOf(at)[id] = a.sc[id][at];
        if (a.why[id][at]) whyOf(at)[id] = a.why[id][at];
      });
    });
    if (mine.some((f) => f.rank)) {
      S.rank = a.rank;
      S.moved = true;
    }
    S.draft = null;
    save();
    go(firstOpen());
    toast('Loaded ' + a.n + ' of ' + P.p.length + ' squadmates from the Excel. Check them, then Submit.');
  };
  input.click();
}

async function copyText() {
  const text = envelopeToText(S.outbox.env);
  try {
    await navigator.clipboard.writeText(text);
    toast('Copied. Paste it into an email to your cadre.');
    clearPlain();
  } catch (err) {
    ui.sheet = { kind: 'text' };
    sheetEl.innerHTML = '<button class="scrim" data-act="text-close" aria-label="Close"></button>' +
      '<div class="card" role="dialog" aria-label="Your answers as text"><b>Copy this into an email to your cadre</b>' +
      '<textarea class="textbox" id="copybox" readonly></textarea><button class="primary" data-act="text-close">Done</button></div>';
    sheetEl.hidden = false;
    const box = document.getElementById('copybox');
    box.value = text;
    box.focus();
    box.select();
    if (!S.cleared) {
      S = Object.assign(fresh(), { step: 'done', cleared: true, rev: S.rev, outbox: S.outbox });
      save();
    }
  }
}

function toast(text) {
  document.querySelectorAll('.toast').forEach((x) => x.remove());
  const t = document.createElement('div');
  t.className = 'toast';
  t.setAttribute('role', 'status');
  t.textContent = text;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 3500);
}

/* ---------- input ---------- */

/* ---------- squad sign-up ---------- */

/** The names to pick from: "LAST, First" where the roster has a first name; OCs sharing a last name with no first names show once. */
function joinEntries() {
  const out = [];
  const seen = {};
  P.ns.forEach(([id, last, first]) => {
    const twins = P.ns.filter((x) => x[1] === last && !x[2]);
    if (!first && twins.length > 1) {
      if (!seen[last]) out.push({ ids: twins.map((x) => x[0]), label: last, note: twins.length + ' OCs' });
      seen[last] = true;
    } else out.push({ ids: [id], label: last + (first ? ', ' + first : ''), note: '' });
  });
  return out;
}

function joinList() {
  const q = String(S.q || '').trim().toLowerCase();
  const rows = joinEntries().map((e, i) => ({ e, i })).filter((x) => !q || x.e.label.toLowerCase().indexOf(q) >= 0);
  return rows.map((x) => '<button class="jname" data-act="jname" data-i="' + x.i + '"><span class="t">' + esc(x.e.label) + '</span>' +
    (x.e.note ? '<span class="sub">' + esc(x.e.note) + '</span>' : '') + '</button>').join('') ||
    '<p class="sub jnone">No name like that. Check the spelling, or tap Not on the list.</p>';
}

const joinWho = () => S.label + (S.grp && String(S.first || '').trim() ? ', ' + String(S.first).trim() : '');

function renderJoin() {
  const sc = S.screen || 'names';
  const back = '<div class="top"><button class="back" data-act="jback" aria-label="Back to the names">' + ICON_BACK + '</button></div>';
  if (sc === 'names') {
    return '<div class="jhead"><span class="sub">' + esc(P.g) + ' · squad sign-up</span><h1>Tap your name</h1></div>' +
      '<label class="jfind"><span class="sub">Find your name</span><input id="jsearch" type="search" autocomplete="off" value="' + esc(S.q || '') + '"></label>' +
      '<div class="jlist" id="jlist">' + joinList() + '</div><button class="link" data-act="jnew">Not on the list?</button>';
  }
  if (sc === 'new') {
    return back + '<h1>Not on the list</h1>' +
      '<label class="jfind"><span class="sub">Your last name</span><input id="jlast" type="text" autocomplete="family-name" value="' + esc(S.newLast || '') + '"></label>' +
      '<p class="sub">Your cadre sees this name and adds you.</p><button class="primary" data-act="jnext">Next</button>';
  }
  if (sc === 'squad') {
    const who = joinWho();
    const btns = Array.from({ length: P.sq }, (x, i) => '<button class="jsq" data-act="jsq" data-v="' + (i + 1) + '">' + ord(i + 1) + ' SQD</button>').join('');
    const replaced = S.sentFor && S.sentFor !== who
      ? '<p class="sub">This phone already picked for ' + esc(S.sentFor) + ': sending this replaces it. Using ' + esc(S.sentFor) + '\'s phone? Ask your cadre to set your squad instead.</p>' : '';
    return back + '<div class="jhead"><span class="sub">' + esc(P.g) + '</span><h1>' + esc(S.label) + ', which squad are you in?</h1></div>' +
      (S.grp ? '<label class="jfind"><span class="sub">There are ' + S.grp.length + ' ' + esc(S.label) + ' in ' + esc(P.g) + ' and the roster has no first names. Your first name:</span>' +
        '<input id="jfirst" type="text" autocomplete="given-name" value="' + esc(S.first || '') + '"></label>' : '') +
      '<div class="jsqs">' + btns + '</div>' + replaced +
      '<p class="sub" id="jhint" role="status">' + esc(S.hint || 'Tapping a squad sends it to your cadre.') + '</p>';
  }
  const sent = !!(S.outbox && S.outbox.sentAt);
  const line = navigator.onLine === false ? 'No signal. It sends by itself when you have signal.' : 'Not sent yet. It keeps trying by itself.';
  return '<div class="done"><span class="badge' + (sent ? '' : ' wait') + '">' + (sent ? ICON_CHECK : ICON_CLOCK) + '</span>' +
    '<h1>' + (sent ? 'Sent' : 'Saved on this phone') + '</h1><p><b>' + esc(S.sentFor || '') + ' · ' + esc(P.g) + ' · ' + ord(S.sqd || 1) + ' SQD</b></p>' +
    '<p class="sub">' + (sent ? 'Your cadre has it. You can close this page.' : line) + '</p>' +
    '<button class="link" data-act="jchange">Change</button><p class="sub">Change goes back to the names, in case you tapped the wrong one.</p></div>';
}

/** Seal the pick and send it; one phone counts once, by its newest pick (rev). */
async function joinSend(sqd) {
  if (S.grp && !String(S.first || '').trim()) {
    S.hint = 'Type your first name first.';
    const el = document.getElementById('jhint');
    if (el) el.textContent = S.hint;
    return;
  }
  const who = joinWho();
  S.rev = (S.rev || 0) + 1;
  const payload = { v: 1, j: 1, e: P.e, r: S.pick, rev: S.rev, at: new Date().toISOString(), plt: P.pl, sqd, dv: deviceId() };
  if (S.grp) {
    payload.grp = S.grp;
    payload.first = String(S.first).trim();
  }
  if (S.typed) payload.last = S.typed;
  let env;
  try {
    env = await sealJoin(P, payload);
  } catch (err) {
    S.rev -= 1;
    toast('Could not prepare your answer on this browser. Try Safari, Chrome or Edge.');
    return;
  }
  Object.assign(S, { outbox: { env, rev: S.rev, sentAt: null }, sqd, sentFor: who, screen: 'sent', hint: '' });
  save();
  render();
  window.scrollTo(0, 0);
  attempt = 0;
  trySend();
}

function onJoinClick(act, b) {
  const show = (screen, extra) => {
    Object.assign(S, extra || {}, { screen });
    save();
    render();
    window.scrollTo(0, 0);
  };
  if (act === 'jname') {
    const e = joinEntries()[Number(b.dataset.i)];
    if (e) show('squad', { pick: e.ids.length > 1 ? -1 : e.ids[0], grp: e.ids.length > 1 ? e.ids : null, label: e.label, typed: '', first: '', hint: '' });
  } else if (act === 'jnew') show('new');
  else if (act === 'jnext') {
    const t = String(S.newLast || '').trim().toUpperCase();
    if (!t) return toast('Type your last name first.');
    show('squad', { pick: -1, grp: null, label: t, typed: t, first: '', hint: '' });
  } else if (act === 'jback') show('names', { hint: '' });
  else if (act === 'jchange') show('names', { q: '' });
  else if (act === 'jsq') return joinSend(Number(b.dataset.v));
  return undefined;
}

function onJoinInput(e) {
  if (ONE && ONE.screen) return onOneInput(e);
  if (!P || !P.j || !S) return;
  const id = e.target.id;
  if (id === 'jsearch') {
    S.q = e.target.value;
    const list = document.getElementById('jlist');
    if (list) list.innerHTML = joinList();
  } else if (id === 'jfirst' || id === 'jlast') {
    S[id === 'jfirst' ? 'first' : 'newLast'] = e.target.value;
    save();
  }
}

function onClick(e) {
  const b = e.target.closest('[data-act]');
  if (b && ONE && ONE.screen) return onOneClick(b.dataset.act, b);
  if (!b || !P) return;
  const act = b.dataset.act;
  if (act === 'one-notme') return oneNotMe();
  if (P.j) return onJoinClick(act, b);
  const id = b.dataset.id !== undefined ? Number(b.dataset.id) : null;
  const st = S.step;
  if (act === 'back') return goBack();
  if (act === 'next') return goNext();
  const person = st[0] === 'p'; // by person: the row is a question or an attribute of this screen's peer
  const pid = person ? peerOf(Number(st.slice(1))) : id;
  const q = person ? Number(b.dataset.q) : Number(st.slice(1));
  if (act === 'mode') {
    if ((S.mode || 'item') === b.dataset.v) return undefined;
    S.mode = b.dataset.v === 'person' ? 'person' : 'item';
    return go(firstOpen());
  }
  if (act === 'yes' || act === 'no' || act === 'score') { // the answer only: once the screen is all answered, the reason card opens by itself
    const j = person ? Number(b.dataset.j) : Number(st.slice(1));
    const full = () => (person ? pAnswered(pid) : st[0] === 'q' ? qAnswered(q) : allScored(j));
    const before = full();
    const kind = act === 'score' ? 'a' : 'q';
    const was = answerOf(kind, kind === 'a' ? j : q, pid);
    if (act === 'score') scOf(attrOf(j, pid))[pid] = Number(b.dataset.v);
    else S.ans[q][pid] = act === 'yes' ? 'Y' : 'N';
    save();
    if (was && was !== answerOf(kind, kind === 'a' ? j : q, pid) && hasWhy(whyBox(kind, kind === 'a' ? j : q, pid), pid)) return openCard({ kind, k: kind === 'a' ? j : q, id: pid }); // a changed answer: its card shows the reason written for the old one, to confirm or edit
    if (!before && full()) {
      const first = nextCard(null);
      if (first !== null) {
        kbEl.focus({ preventScroll: true });
        ui.opening = true; // the card opens by itself in a moment: no Say why button meanwhile
        render();
        clearTimeout(timers.open);
        timers.open = setTimeout(() => {
          ui.opening = false;
          if (S.step !== st || ui.sheet) return;
          openCard(first);
        }, ADVANCE_MS);
        return;
      }
    }
    return afterChange();
  }
  if (act === 'load-xlsx') return loadSheet();
  if (act === 'swap') return openSwap(b.dataset.j !== undefined ? Number(b.dataset.j) : null);
  if (act === 'swap-from') return openSwap(Number(b.dataset.v));
  if (act === 'swap-to') return swapTo(ui.sheet.j, Number(b.dataset.v));
  if (act === 'qname') return openCard({ kind: 'q', k: q, id: pid });
  if (act === 'aname') return openCard({ kind: 'a', k: person ? Number(b.dataset.j) : Number(st.slice(1)), id: pid });
  if (act === 'why-next') { // Say why: the first answer on this screen still without a reason
    const f = nextCard(null);
    return f ? openCard(f) : render();
  }
  if (act === 'seg') { // the answer on the card: Yes / No, or a score
    const s = ui.sheet;
    if (s.kind === 'q') S.ans[s.k][s.id] = b.dataset.v === 'N' ? 'N' : 'Y';
    else scOf(attrOf(s.k, s.id))[s.id] = Number(b.dataset.v);
    save();
    fillCard(false);
    render();
    document.getElementById('sheet-input').focus({ preventScroll: true });
    return;
  }
  if (act === 'sheet-save') return onSave();
  if (act === 'sheet-close') return onClose();
  if (act === 'text-close') {
    ui.sheet = null;
    sheetEl.hidden = true;
    sheetEl.innerHTML = '';
    return render();
  }
  if (act === 'rk') {
    const j = Number(b.dataset.j);
    if (ui.sel < 0) ui.sel = j;
    else if (ui.sel === j) ui.sel = -1;
    else {
      const moved = S.rank.splice(ui.sel, 1)[0];
      S.rank.splice(j, 0, moved);
      ui.sel = -1;
      S.moved = true;
      S.rankSig = rankSig();
      save();
    }
    return render();
  }
  if (act === 'submit') return firstOpen() !== 'rank' ? go(firstOpen()) : submit(); // a screen still unfinished goes first
  if (act === 'change') {
    ui.mine = null; // editing again: the copy submitted before is no longer what's on the phone
    return S.cleared ? render() : go('rank');
  }
  if (act === 'restart') {
    if (!S.cleared) return go('rank');
    const how = S.outbox && S.outbox.sentAt ? 'once your cadre had them' : 'once you saved or copied them';
    const ok = window.confirm('For privacy, your answers were cleared from this device ' + how + '. Start again from the first question? Your new answers will replace the old ones.');
    if (!ok) return undefined;
    const queued = S.outbox && !S.outbox.sentAt ? S.outbox : null;
    ui.mine = null;
    S = Object.assign(fresh(), { rev: S.rev, outbox: queued });
    save();
    return go('q0');
  }
  if (act === 'file') return sendFile();
  if (act === 'xlsx') return saveSheet();
  if (act === 'copy') return copyText();
  return undefined;
}

/* ---------- one link (cadre mode) ---------- */

let ONE = null; // { e, k, board, brev, st: this phone's own state, screen, wait, timer }
const oneKey = (e) => 'pe:one:' + e;
const boardKey = (e) => 'pe:b2:' + e;

function oneLoad(e) {
  try {
    return JSON.parse(localStorage.getItem(oneKey(e)) || 'null') || { me: null, rev: 0 };
  } catch (err) {
    storageOk = false;
    return { me: null, rev: 0 };
  }
}

function oneSave() {
  try {
    localStorage.setItem(oneKey(ONE.e), JSON.stringify(ONE.st));
  } catch (err) {
    storageOk = false;
  }
}

/** Open the one link: the board kept on this phone first (no signal needed), then the sheet's. */
async function startOne(code) {
  const l = parseOc(code);
  if (!l) return renderStart('This link looks damaged. Ask your cadre for the QR code again.');
  ONE = { e: l.e, k: l.k, board: null, brev: 0, st: oneLoad(l.e), screen: 'loading', wait: 30000, timer: 0, sendTimer: 0, sending: false, attempt: 0, q: '' };
  document.addEventListener('visibilitychange', () => {
    if (!ONE || ONE.screen !== 'waiting') return;
    if (document.visibilityState === 'visible') oneWatch(true);
    else clearTimeout(ONE.timer);
  });
  window.addEventListener('online', () => ONE && oneSend());
  try {
    const kept = JSON.parse(localStorage.getItem(boardKey(l.e)) || 'null');
    if (kept && kept.board) {
      ONE.board = await openBoard(l.k, l.e, kept.board);
      ONE.brev = kept.brev;
    }
  } catch (err) {
    ONE.board = null;
  }
  if (ONE.board) await oneShow();
  else render();
  const got = await oneFetch();
  if (got === 'gone') return oneScreen('expired');
  if (!ONE.board) return oneScreen('nosignal');
  if (ONE.screen) return oneShow();
  return undefined;
}

/** The sheet's board: 'ok', 'gone' (removed or never there), or 'none' (no answer). */
async function oneFetch(opts) {
  let g = null;
  try {
    g = await getBoard(ONE.e, opts);
  } catch (err) {
    return 'none';
  }
  if (!g || !g.ok) return g && (g.error === 'removed' || g.error === 'none') ? 'gone' : 'none';
  try {
    ONE.board = await openBoard(ONE.k, ONE.e, g.board);
  } catch (err) {
    return ONE.board ? 'ok' : 'gone';
  }
  ONE.brev = g.brev;
  try {
    localStorage.setItem(boardKey(ONE.e), JSON.stringify({ brev: g.brev, board: g.board }));
  } catch (err) {
    storageOk = false;
  }
  return 'ok';
}

/** Open only from a frozen board (a set-mode board is frozen from the start). The sheet's flag over a sign-up board keeps the phone waiting and checking, never "not in a squad". */
const oneOpen = () => !!ONE.board && (ONE.board.started || ONE.board.mode === 'set');
const oneRow = (id) => ONE.board.oc.find((o) => o[0] === id);
/** A squad set for this row: by the file or the cadre (a newcomer's own pick, kept on the board for older pages, is still his to change). */
const setSq = (o) => (o && o[4] && (!o[5] || (ONE.board.fx || []).indexOf(o[0]) >= 0) ? o[4] : 0);
const onePlt = (n) => (n ? nth(n) + ' PLT' : '');
/** The letters of a last name, to see two OCs with one last name. */
const lastKey = (s) => String(s || '').toUpperCase().replace(/[^\p{L}]/gu, '');
/** Who this phone signs up now: the row, the twins with the first name, or the name typed. */
const whoKey = () => {
  const st = ONE.st;
  if (st.me !== null && st.me !== undefined) return 'r:' + st.me;
  if (st.grp) return 'g:' + st.grp.join(',') + '|' + String(st.first || '').trim().toLowerCase();
  return 't:' + st.typed + '|' + String(st.typedFirst || '').toLowerCase();
};
/** A typed first name as the cadre's fields write it: each part's first letter up (Jean-Luc), never after an apostrophe (D'andre). */
const capFirst = (s) => String(s || '').trim().replace(/(^|[\s-])(\p{Ll})/gu, (m, a, c) => a + c.toUpperCase()).slice(0, 30);

/** This phone's own row on the board: its roster row, the twin row with its first name, or the row added for its typed name; undefined when none is there yet. */
function ownRow() {
  const st = ONE.st;
  if (st.me !== null && st.me !== undefined) return st.me;
  if (st.grp) {
    const f = String(st.first || '').trim().toLowerCase();
    const name = (id) => String(oneRow(id)[2] || '').toLowerCase();
    const rows = st.grp.filter((id) => oneRow(id));
    const hit = rows.find((id) => name(id) === f) // his first name on the board
      ?? rows.find((id) => (ONE.board.nd || {})[id] === deviceId() && !!f && name(id)[0] === f[0]); // else the row this phone's pick took, named a near spelling (the cadre's Michael, for Mike)
    if (hit !== undefined) return hit;
  }
  if (st.typed) { // a name typed in: the row the cadre added from this phone, or the one the cadre said is this OC (Same OC: his roster name, or added from his other browser); never another OC by its name alone
    const f = String(st.typedFirst || '').toLowerCase();
    const link = (ONE.board.ln || {})[deviceId() + '|' + st.typed + (f ? '|' + f : '')];
    const other = st.meFor === whoKey() && st.isMe === 0; // he said someone else: never the nameless row this phone signed up before
    const hit = ONE.board.oc.find((o) => o[5] === deviceId() && o[1] === st.typed && (other ? String(o[2] || '').toLowerCase() === f : !f || !o[2] || String(o[2]).toLowerCase() === f)) || (link !== undefined ? oneRow(link) : null);
    if (hit) return hit[0];
  }
  return undefined;
}

/** Where this phone stands: its name picked, its squad, the evaluation open. */
async function oneShow() {
  const st = ONE.st;
  if (oneOpen()) {
    if (st.me === null) {
      const hit = ownRow();
      if (hit !== undefined) st.me = hit;
    }
    if (st.me !== null && st.me !== undefined && oneRow(st.me)) return oneBegin();
    if (st.grp || st.typed) return oneScreen('out');
    return oneScreen('names');
  }
  if (st.me !== null && st.me !== undefined && oneRow(st.me)) {
    const set = setSq(oneRow(st.me));
    if (set && !st.ask && !st.outbox) return onePick(set); // signed in on a set squad before this version: the cadre sees it now
    return oneScreen((set && !st.ask) || st.sqd ? 'waiting' : 'squad');
  }
  if (st.grp || st.typed) return oneScreen(st.sqd ? 'waiting' : 'squad');
  return oneScreen('names');
}

function oneScreen(screen) {
  ONE.screen = screen;
  render();
  window.scrollTo(0, 0);
  if (screen === 'waiting' || (screen === 'squad' && ONE.st.late)) oneWatch(true); // asked his first name late, he is signed up already: the Start still opens it
  else clearTimeout(ONE.timer);
}

/** The evaluation itself: this OC's packet from the board, then the screens as before. */
async function oneBegin() {
  const packet = await packetOf(ONE.board, ONE.k, ONE.st.me, collectorUrl());
  if (!packet) return oneScreen('out');
  clearTimeout(ONE.timer);
  oneSave();
  ONE.screen = null;
  P = packet;
  return start();
}

function oneNotMe() {
  if (!ONE) return;
  Object.assign(ONE.st, { me: null, grp: null, ask: false, typed: '', sqd: 0 });
  oneSave();
  P = null;
  S = null;
  oneScreen('names');
}

/** The names to tap: by platoon; while OCs choose, two OCs with one last name and no first names show once. */
function oneEntries() {
  const out = [];
  const seen = {};
  const open = oneOpen();
  const oc = ONE.board.oc.slice().sort((a, b) => (a[3] || 0) - (b[3] || 0) || a[1].localeCompare(b[1]) || String(a[2]).localeCompare(String(b[2])));
  oc.forEach(([id, last, first, plt, sqd]) => {
    const twins = oc.filter((x) => x[3] === plt && x[1] === last && !x[2]);
    if (!first && twins.length > 1 && !open) {
      if (!seen[plt + last]) out.push({ ids: twins.map((x) => x[0]), plt, label: last, note: twins.length + ' OCs' });
      seen[plt + last] = true;
    } else out.push({ ids: [id], plt, label: fullName(last, first), note: !first && twins.length > 1 && sqd ? nth(sqd) + ' SQD' : '' });
  });
  return out;
}

function oneList() {
  const q = String(ONE.q || '').trim().toLowerCase();
  let plt = -1;
  let h = '';
  oneEntries().forEach((x, i) => {
    if (q && x.label.toLowerCase().indexOf(q) < 0) return;
    if (x.plt !== plt) {
      plt = x.plt;
      if (plt) h += '<p class="jplt">' + esc(onePlt(plt)) + '</p>';
    }
    h += '<button class="jname" data-act="o-name" data-i="' + i + '"><span class="t">' + esc(x.label) + '</span>' + (x.note ? '<span class="sub">' + esc(x.note) + '</span>' : '') + '</button>';
  });
  return h || '<p class="sub jnone">No name like that. Check the spelling' + (oneOpen() ? '.' : ', or tap Not on the list.') + '</p>';
}

function oneWho() {
  const st = ONE.st;
  return st.label + ((st.grp || st.ask) && String(st.first || '').trim() ? ', ' + String(st.first).trim() : '');
}

const ONE_ICON_GONE = '<svg width="30" height="30" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="#17191C" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 6 V13"/><path d="M12 17.5 V18"/></svg>';

function renderOne() {
  const sc = ONE.screen;
  const st = ONE.st;
  const back = '<div class="top"><button class="back" data-act="o-back" aria-label="Back to the names">' + ICON_BACK + '</button></div>';
  if (sc === 'loading') return '<div class="done"><p class="sub">Opening…</p></div>';
  if (sc === 'nosignal') {
    return banners() + '<div class="done"><span class="badge wait">' + ICON_CLOCK + '</span><h1>No signal</h1><p class="sub">Open this link again when you have signal. Once it opens, you can fill it in without signal.</p>' +
      '<button class="link" data-act="o-retry">Try again</button></div>';
  }
  if (sc === 'expired') return '<div class="done"><span class="badge wait">' + ICON_CLOCK + '</span><h1>This QR code has expired</h1><p class="sub">Ask your cadre for a new one.</p></div>';
  if (sc === 'out') {
    return '<div class="done"><span class="badge wait">' + ONE_ICON_GONE + '</span><h1>The evaluation has started</h1><p class="sub">You\'re not in a squad for it. Tell your cadre.</p>' +
      '<button class="link" data-act="o-change">Not you? Go back</button></div>';
  }
  if (sc === 'names') {
    return banners() + '<div class="jhead"><span class="sub">' + esc(/peer[\s-]*eval/i.test(ONE.board.n) ? ONE.board.n : ONE.board.n + ' · peer evaluation') + '</span><h1>Tap your name</h1></div>' +
      '<label class="jfind"><span class="sub">Find your name</span><input id="osearch" type="search" autocomplete="off" value="' + esc(ONE.q || '') + '"></label>' +
      '<div class="jlist" id="olist">' + oneList() + '</div>' + (oneOpen() ? '' : '<button class="link" data-act="o-new">Not on the list?</button>');
  }
  if (sc === 'new') {
    return back + '<h1>Not on the list</h1>' +
      '<label class="jfind"><span class="sub">Your last name</span><input id="olast" type="text" autocomplete="family-name" value="' + esc(st.newLast || '') + '"></label>' +
      '<label class="jfind"><span class="sub">Your first name</span><input id="onewfirst" type="text" autocomplete="given-name" value="' + esc(st.newFirst || '') + '"></label>' +
      '<p class="sub" id="onewhint" role="status">Your cadre sees this name and adds you.</p><button class="primary" data-act="o-newnext">Next</button>';
  }
  if (sc === 'isyou') {
    return back + '<div class="jhead"><h1>This phone signed up ' + esc(st.sentFor) + '. Is that you?</h1></div>' +
      '<div class="jsqs"><button class="primary" data-act="o-me" data-v="1">Yes, fix my name</button><button class="jsq" data-act="o-me" data-v="0">No, I\'m someone else</button></div>' +
      '<p class="sub">Tapped the wrong name before? Tap Yes.</p>';
  }
  if (sc === 'squad') {
    const o = st.me !== null && st.me !== undefined ? oneRow(st.me) : null;
    const plt = o ? o[3] : st.plt;
    const set = setSq(o) || (st.late ? st.sqd : 0); // the cadre or the file set his squad, or he picked it before another OC of his last name came: only his first name is asked
    const btns = set ? '<button class="primary" data-act="o-sq" data-v="' + set + '">Send</button>' : Array.from({ length: ONE.board.sq || 4 }, (x, i) => '<button class="jsq" data-act="o-sq" data-v="' + (i + 1) + '">' + nth(i + 1) + ' SQD</button>').join('');
    return back + '<div class="jhead"><span class="sub">' + esc(st.label + (plt ? ' · ' + onePlt(plt) : '')) + '</span><h1>' + (set ? 'You\'re in ' + nth(set) + ' squad' : 'Which squad are you in?') + '</h1></div>' +
      (st.grp || st.ask ? '<label class="jfind"><span class="sub">' + (st.grp ? 'There are ' + st.grp.length + ' ' + esc(st.label) + (plt ? ' in ' + esc(onePlt(plt)) : '') + ' and the roster has no first names. Your first name:' : 'Another OC is named ' + esc(st.label) + '. Your first name, so raters can tell you apart:') + '</span>' +
        '<input id="ofirst" type="text" autocomplete="given-name" value="' + esc(st.first || '') + '"></label>' : '') +
      '<div class="jsqs">' + btns + '</div>' + '<p class="sub" id="ohint" role="status">' + esc(st.hint || (set ? 'Send gives your cadre your first name.' : 'Tapping a squad sends it to your cadre.')) + '</p>';
  }
  // waiting for the cadre's Start: the squad on this phone's own row when the cadre set or moved it (a twin row, an added newcomer too)
  const own = ownRow();
  const o = own !== null && own !== undefined ? oneRow(own) : null;
  const sqd = setSq(o) || st.sqd || (o && o[4]);
  const sent = !st.outbox || !!st.outbox.sentAt;
  const line = st.outbox && st.outbox.blocked ? 'The sheet did not take your squad. Tell your cadre which squad you are in.'
    : !sent ? (navigator.onLine === false ? 'Saved on this phone. It sends by itself when you have signal.' : 'Sending… It keeps trying by itself.') : '';
  return banners() + '<div class="done"><span class="badge' + (sent ? '' : ' wait') + '">' + (sent ? ICON_CHECK : ICON_CLOCK) + '</span>' +
    '<p class="whoami">' + esc(oneWho()) + '</p><h1>You\'re in ' + nth(sqd) + ' squad</h1><p><b>The evaluation opens here when your cadre starts it.</b></p>' +
    '<p class="sub">Keep this page open, or come back to the same link.</p>' + (line ? '<p class="sub">' + esc(line) + '</p>' : '') +
    '<button class="link" data-act="o-change">' + (setSq(o) ? 'Wrong name? Change' : 'Wrong name or squad? Change') + '</button></div>';
}

/** Check for Start lightly: now, then every 30 s ± 10 s while the page shows; double the wait after an error, up to 2 min. */
function oneWatch(now) {
  clearTimeout(ONE.timer);
  if (document.hidden) return;
  ONE.timer = setTimeout(oneCheck, now ? 0 : ONE.wait * (2 / 3 + Math.random() * (2 / 3)));
}

async function oneCheck() {
  if (!ONE || !(ONE.screen === 'waiting' || (ONE.screen === 'squad' && ONE.st.late))) return undefined;
  let s = null;
  try {
    s = await getState(ONE.e, { retries: 0 });
  } catch (err) {
    s = null;
  }
  if (s && s.ok) {
    ONE.wait = 30000;
    if (s.started || s.brev > ONE.brev) {
      const got = await oneFetch();
      if (got === 'gone') return oneScreen('expired');
      if (oneOpen()) return oneShow();
      const st = ONE.st;
      const row = st.me !== null && st.me !== undefined ? oneRow(st.me) : null;
      if (ONE.screen === 'waiting' && row && !row[2] && !String(st.first || '').trim() && ONE.board.oc.some((y) => y[0] !== row[0] && lastKey(y[1]) === lastKey(row[1]))) { // another OC of this last name since: his first name, so raters tell them apart
        Object.assign(st, { ask: true, late: true });
        oneSave();
        return oneScreen('squad');
      }
      if (ONE.screen === 'waiting') render();
    }
  } else if (s && (s.error === 'removed' || s.error === 'none')) {
    return oneScreen('expired');
  } else {
    ONE.wait = Math.min(120000, ONE.wait * 2);
  }
  oneWatch(false);
  return undefined;
}

/** Seal the squad pick and send it; one phone counts once, by its newest pick (rev). */
async function onePick(sqd) {
  const st = ONE.st;
  if ((st.grp || st.ask) && !String(st.first || '').trim()) {
    st.hint = 'Type your first name first.';
    const el = document.getElementById('ohint');
    if (el) el.textContent = st.hint;
    return undefined;
  }
  const o = st.me !== null && st.me !== undefined ? oneRow(st.me) : null;
  if (st.grp || st.ask) st.first = capFirst(st.first);
  const key = whoKey();
  if (st.sentFor && String(st.sentFor).toLowerCase() !== oneWho().toLowerCase() && (!st.sentKey || st.sentKey !== key) && st.meFor !== key) { // this phone signed up another OC before: the person tells the cadre who he is
    st.askMe = sqd;
    oneSave();
    return oneScreen('isyou');
  }
  st.rev = (st.rev || 0) + 1;
  const payload = { v: 1, j: 1, e: ONE.e, r: o ? o[0] : -1, rev: st.rev, at: new Date().toISOString(), plt: o ? o[3] : st.plt || 0, sqd, dv: deviceId() };
  if (st.grp) payload.grp = st.grp;
  if (st.grp || st.ask) payload.first = st.first; // a twin, or a name without a first name next to another OC's: the cadre's board takes it (tony goes as Tony)
  if (st.meFor === key) payload.me = st.isMe; // "Is that you?": 1 = the same OC, his name fixed; 0 = someone else on a lent phone
  if (st.typed) payload.last = st.typed;
  if (st.typed && st.typedFirst) payload.first = st.typedFirst;
  if (st.typed && st.told) payload.told = 1; // "X is on the list. You?": he said someone else
  let env;
  try {
    env = await sealJoin({ k: ONE.board.pub, t: await joinToken(ONE.k) }, payload);
  } catch (err) {
    st.rev -= 1;
    return toast('Could not prepare your answer on this browser. Try Safari, Chrome or Edge.');
  }
  Object.assign(st, { outbox: { env, sentAt: null }, sqd, sentFor: oneWho(), sentKey: key, hint: '', late: false, askMe: null });
  oneSave();
  oneScreen('waiting');
  ONE.attempt = 0;
  return oneSend();
}

async function oneSend() {
  clearTimeout(ONE.sendTimer);
  const st = ONE.st;
  if (!st.outbox || st.outbox.sentAt || st.outbox.blocked || ONE.sending) return;
  ONE.sending = true;
  let ok = false;
  let refused = false;
  try {
    const res = await fetch(collectorUrl(), { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(st.outbox.env), cache: 'no-store', redirect: 'follow' });
    const data = JSON.parse(await res.text());
    ok = !!(data && data.ok);
    refused = !ok && /^(bad (size|envelope)|removed|none|full)$/.test(String(data && data.error));
  } catch (err) {
    ok = false;
  }
  ONE.sending = false;
  if (ok || refused) {
    if (ok) st.outbox.sentAt = new Date().toISOString();
    else st.outbox.blocked = true;
    oneSave();
    if (ONE.screen === 'waiting') render();
    ONE.attempt = 0;
    return;
  }
  ONE.attempt += 1;
  ONE.sendTimer = setTimeout(oneSend, Math.min(60000, 2000 * Math.pow(2, Math.min(ONE.attempt, 5))) * (0.5 + Math.random()));
}

function onOneClick(act, b) {
  const st = ONE.st;
  if (act === 'o-name') {
    const x = oneEntries()[Number(b.dataset.i)];
    if (!x) return undefined;
    if (oneOpen()) {
      Object.assign(st, { me: x.ids[0], grp: null, typed: '', label: x.label });
      return oneBegin();
    }
    if (x.ids.length > 1) Object.assign(st, { me: null, grp: x.ids, ask: false, typed: '', label: x.label, first: '', plt: x.plt, sqd: 0, hint: '' });
    else {
      const row = oneRow(x.ids[0]);
      const ask = !!row && !row[2] && ONE.board.oc.some((y) => y[0] !== row[0] && lastKey(y[1]) === lastKey(row[1])); // no first name, and another OC has this last name: raters tell them apart by it
      Object.assign(st, { me: x.ids[0], grp: null, ask, typed: '', label: x.label, first: ask ? '' : st.first, plt: x.plt, sqd: 0, hint: '' });
    }
    oneSave();
    const o = st.me !== null ? oneRow(st.me) : null;
    if (setSq(o) && !st.ask) return onePick(setSq(o)); // his squad is set: the pick only tells the cadre he signed in
    return oneScreen('squad');
  }
  if (act === 'o-sq') return onePick(Number(b.dataset.v));
  if (act === 'o-me') { // "Is that you?": the answer, then the pick it was asked for
    Object.assign(st, { isMe: Number(b.dataset.v), meFor: whoKey() });
    oneSave();
    return onePick(st.askMe);
  }
  if (act === 'o-back' || act === 'o-change') {
    ONE.q = '';
    return oneScreen('names');
  }
  if (act === 'o-new') return oneScreen('new');
  if (act === 'o-newnext') return oneNewNext();
  if (act === 'o-retry') return location.reload();
  return undefined;
}

/** Not on the list › Next: against the board as it is now (a name the cadre added since this page opened counts too). */
async function oneNewNext() {
  const st = ONE.st;
  if (!String(st.newLast || '').trim()) return toast('Type your last name first.');
  if (ONE.checking) return undefined; // the check on its way reads the names as typed when it ends
  const hint = (text) => {
    const el = document.getElementById('onewhint');
    if (el) el.textContent = text;
  };
  ONE.checking = true;
  hint('Checking the list…');
  const got = navigator.onLine === false ? 'none' : await oneFetch({ retries: 0 }); // no signal: the board kept on this phone
  ONE.checking = false;
  if (got === 'gone') return oneScreen('expired');
  if (ONE.screen !== 'new') return undefined;
  const t = String(st.newLast || '').trim().toUpperCase().slice(0, 30);
  const f = capFirst(st.newFirst);
  if (!t) return hint('Type your last name first.');
  const typed = () => {
    Object.assign(st, { me: null, grp: null, ask: false, typed: t, typedFirst: f, label: f ? t + ', ' + f : t, plt: 0, sqd: 0, hint: '', told: ONE.told === t + '|' + f });
    oneSave();
  };
  const same = ONE.board.oc.filter((o) => lastKey(o[1]) === lastKey(t)); // (started meanwhile too: his own name on the list still opens his evaluation)
  if (!f && same.length) return hint(t + ' is on the list. You? Go back and tap your name. Someone else? Type your first name too.'); // the first name tells them apart
  const listed = same.find((o) => String(o[2] || '').toLowerCase() === f.toLowerCase()); // his own name, missed on the list
  if (listed) {
    ONE.q = listed[1]; // as the list spells it (ZR-DORSET, O'DORSET): the search finds it
    oneScreen('names');
    return toast(fullName(listed[1], listed[2]) + ' is on the list. Tap your name.');
  }
  const bare = same.find((o) => !o[2] && !((ONE.board.fx || []).indexOf(o[0]) >= 0 && !o[4])); // one with this last name and no first name (one the cadre left out isn't asked about): maybe him
  if (bare && ONE.told !== t + '|' + f) {
    ONE.told = t + '|' + f;
    const plts = new Set(ONE.board.oc.map((o) => o[3]).filter(Boolean)).size > 1; // one code for several platoons: his is named
    return hint(bare[1] + (plts && bare[3] ? ' in ' + onePlt(bare[3]) : '') + ' is on the list with no first name. You? Go back and tap ' + bare[1] + '. Someone else? Tap Next again.');
  }
  typed();
  if (oneOpen()) return oneShow(); // started meanwhile: no more sign-ups; the row the cadre gave this name, else "not in a squad"
  return oneScreen('squad');
}

function onOneInput(e) {
  const id = e.target.id;
  if (id === 'osearch') {
    ONE.q = e.target.value;
    const el = document.getElementById('olist');
    if (el) el.innerHTML = oneList();
  } else if (id === 'ofirst') {
    ONE.st.first = e.target.value;
    oneSave();
  } else if (id === 'olast') {
    ONE.st.newLast = e.target.value;
  } else if (id === 'onewfirst') {
    ONE.st.newFirst = e.target.value;
  }
}

/* ---------- start ---------- */

function renderStart(message) {
  // the site's front page (no link, hosted) also leads cadre to their page
  const cadre = !message && /^https?:/.test(location.protocol) ? '<a class="link" id="to-cadre" href="cadre.html' + esc(testQuery()) + '" style="display:flex;align-items:center;justify-content:center;margin-top:20px">For cadre: run a peer evaluation</a>' : '';
  screenEl.innerHTML = '<div class="start" style="display:flex;flex-direction:column;gap:14px;padding-top:24px">' +
    '<h1>Open your peer evaluation</h1>' +
    '<p class="sub">' + esc(message || 'Paste the message from your cadre, or just the link in it.') + '</p>' +
    '<label class="sub" for="code">Message or link</label>' +
    '<textarea id="code" autocomplete="off" spellcheck="false"></textarea>' +
    '<button class="primary" id="open-code">Open</button>' + cadre + '</div>';
  document.getElementById('open-code').addEventListener('click', async () => {
    const text = document.getElementById('code').value;
    try {
      P = await readPacket(text);
    } catch (err) {
      return renderStart('No working link found in that text. Paste the whole message from your cadre again.');
    }
    try {
      history.replaceState(null, '', '#' + (await packPacket(P)));
    } catch (err) {
      // file:// pages may refuse; the app still works for this visit
    }
    return start();
  });
}

function start() {
  KEY = P.j ? 'pe:join:' + P.e + ':' + P.pl : 'pe:oc:' + P.e + ':' + P.me + ':' + String(P.t).slice(0, 6);
  try {
    localStorage.setItem('pe:probe', '1');
    localStorage.removeItem('pe:probe');
  } catch (err) {
    storageOk = false;
  }
  if (P.j) {
    S = load() || { v: 1, screen: 'names', rev: 0 };
    if (S.screen === 'squad' && !S.label) S.screen = 'names';
    save();
    document.title = 'Squad sign-up · ' + P.g;
    render();
    trySend();
    return;
  }
  S = load() || fresh();
  if (!Array.isArray(S.ans) || S.ans.length !== 4) S = fresh();
  if (S.cleared) S.step = 'done';
  else if (!S.outbox) S.step = firstOpen();
  if (steps().indexOf(S.step) < 0) S.step = S.outbox ? 'done' : firstOpen();
  save();
  document.title = 'Peer evaluation · ' + P.nm;
  render();
  resumeDraft();
  trySend();
}

function setupViewport() {
  const vv = window.visualViewport;
  if (!vv) return;
  const update = () => {
    const kb = Math.max(0, window.innerHeight - vv.height - vv.offsetTop);
    document.documentElement.style.setProperty('--kb', kb + 'px');
  };
  vv.addEventListener('resize', update);
  vv.addEventListener('scroll', update);
  update();
}

async function boot() {
  root.innerHTML = '<div id="screen" class="screen"></div><div id="sheet" class="sheet" hidden></div>' +
    '<input id="kbproxy" class="kbproxy" aria-hidden="true" tabindex="-1" autocomplete="off">';
  screenEl = document.getElementById('screen');
  sheetEl = document.getElementById('sheet');
  kbEl = document.getElementById('kbproxy');
  root.addEventListener('click', onClick);
  root.addEventListener('input', onJoinInput);
  setupViewport();
  window.addEventListener('online', () => {
    attempt = 0;
    trySend();
  });
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') trySend();
  });
  // Another link opened in this tab changes only the # (no page load), so start again from it.
  // The app itself never sets the # (pasting uses replaceState), so this cannot loop.
  window.addEventListener('hashchange', () => location.reload());
  if ('serviceWorker' in navigator && /^https?:/.test(location.protocol)) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  const code = location.hash.slice(1);
  if (!code) return renderStart();
  if (/^o2\./.test(code)) return startOne(code);
  try {
    P = await readPacket(code);
  } catch (err) {
    return renderStart('This link looks damaged. Ask your cadre for a new one, or paste your code here.');
  }
  return start();
}

boot();
