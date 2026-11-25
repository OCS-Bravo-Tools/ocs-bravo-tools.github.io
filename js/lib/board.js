/**
 * Cadre mode: one link per evaluation, and the sealed board behind it.
 *
 * OC link     <site>/#o2.<id>.<link key>                                   about 80 characters
 * Cadre link  <site>/cadre.html#c2.<id>.<link key>.<write code>.<d>.<pub>   about 240 characters
 *   id          8 characters (6 random bytes), also the sheet's name for the evaluation
 *   link key    16 random bytes. It opens the board. Each OC's signing token is
 *               HMAC(link key, 'oc ' + id) and the sign-up's is HMAC(link key, 'join'):
 *               one shared link can't prove more than that it was used
 *   write code  16 random bytes. The page sends HMAC(write code, 'write ' + id), and the
 *               sheet keeps only that token's SHA-256. It changes the board, starts the
 *               evaluation, reads rows and removes the evaluation
 *   d, pub      the cadre's P-256 key pair (private scalar, raw public key). Results are
 *               sealed for pub, as before; only d opens them
 * Nothing after # is ever sent to a server.
 *
 * The board, sealed with the link key (AES-GCM, the id as associated data):
 *   { v: 2, e, n: name, at: created, pub, seed, sq: squads per platoon, mode: 'oc' | 'set',
 *     started, st: start mark, oc: [[id, LAST, First, plt, sqd, dv?]] }   sqd 0 = not in a squad yet;
 *     dv only on a row the cadre added for a "Not on the list" name: the id of the phone that typed it
 *   st: a random number made when the board is frozen at Start. It also goes to the sheet in
 *   the clear, so the sheet takes no other board once a started one is in
 *   dl: the deal. None: one list per rater, squads seeded by their place; 2 (17b): seeded by
 *   platoon and squad number; 3 (17e): PAIR_K different attributes for each rater and peer
 * The cadre's page and every phone work out the squads and each OC's deal from it with
 * the same code (eventOf), so a phone needs nothing but the board and its own name.
 */
import { b64url, fromB64url, utf8, fromUtf8, deflate, inflate, canCompress } from './codec.js';
import { hmac, sha256, randomToken, randomBytes, generateAdminKeys } from './crypto.js';
import { dealSquad, dealPairs, PAIR_K } from './deal.js';

const subtle = globalThis.crypto.subtle;
export const UNUSED_DAYS = 30; // an evaluation with no OC answer is removed after this
export const KEEP_DAYS = 365; // one with answers is removed this long after the last one
const DAY = 86400000;

const ID_RE = '[A-Za-z0-9_-]{8}';
const KEY_RE = '[A-Za-z0-9_-]{22}';
const OC_HASH = new RegExp('^o2\\.(' + ID_RE + ')\\.(' + KEY_RE + ')$');
const CADRE_HASH = new RegExp('^c2\\.(' + ID_RE + ')\\.(' + KEY_RE + ')\\.(' + KEY_RE + ')\\.([A-Za-z0-9_-]{43})\\.([A-Za-z0-9_-]{87})$');

export const ord = (n) => n + (['th', 'st', 'nd', 'rd'][(n % 100 - 20) % 10] || ['th', 'st', 'nd', 'rd'][n % 100] || 'th');
export const fullName = (last, first) => last + (first ? ', ' + first : '');

/** A new evaluation's id and every key, all made here. */
export async function newKeys() {
  const pair = await generateAdminKeys();
  let e = randomToken(6);
  while (!isNaN(Number(e)) || e.indexOf('-') >= 0) e = randomToken(6); // Sheets would store "00012345" or "1234E567" as a number, "1-2-2026" as a date
  return { e, k: randomToken(16), w: randomToken(16), d: pair.priv.d, pub: pair.pub };
}

/** The private JWK from the scalar and the raw public key (0x04 || X || Y). */
export function privOf(d, pub) {
  const raw = fromB64url(pub);
  return { kty: 'EC', crv: 'P-256', d, x: b64url(raw.subarray(1, 33)), y: b64url(raw.subarray(33, 65)), ext: true };
}

export const ocLink = (base, e, k) => base + '#o2.' + e + '.' + k;
export const cadreLink = (base, r) => base + '#c2.' + [r.e, r.k, r.w, r.d, r.pub].join('.');

/** @returns {{e: string, k: string} | null} */
export function parseOc(hash) {
  const m = OC_HASH.exec(String(hash || '').replace(/^#/, ''));
  return m ? { e: m[1], k: m[2] } : null;
}

/** @returns {{e, k, w, d, pub} | null} */
export function parseCadre(hash) {
  const m = CADRE_HASH.exec(String(hash || '').replace(/^#/, ''));
  return m ? { e: m[1], k: m[2], w: m[3], d: m[4], pub: m[5] } : null;
}

export const ocToken = async (k, id) => (await hmac(k, 'oc ' + id)).slice(0, 22);
export const joinToken = async (k) => (await hmac(k, 'join')).slice(0, 22);
/** What the page sends for the write code: HMAC(write code, 'write ' + id), so the code itself never leaves the device. */
export const writeToken = async (w, e) => hmac(w, 'write ' + e);
/** What the sheet keeps: the SHA-256 of that token. */
export const writeHash = (wt) => sha256(wt);

const boardKey = (k) => subtle.importKey('raw', fromB64url(k), 'AES-GCM', false, ['encrypt', 'decrypt']);

/** Seal the board with the link key. @returns {Promise<{z: number, iv: string, ct: string}>} */
export async function sealBoard(k, board) {
  const raw = utf8(JSON.stringify(board));
  const z = canCompress ? 1 : 0;
  const body = z ? await deflate(raw) : raw;
  const iv = randomBytes(12);
  const ct = new Uint8Array(await subtle.encrypt({ name: 'AES-GCM', iv, additionalData: utf8('pe-board ' + board.e) }, await boardKey(k), body));
  return { z, iv: b64url(iv), ct: b64url(ct) };
}

/** Open a sealed board. Throws if the key is wrong or anything was changed. */
export async function openBoard(k, e, sealed) {
  const pt = new Uint8Array(await subtle.decrypt({ name: 'AES-GCM', iv: fromB64url(sealed.iv), additionalData: utf8('pe-board ' + e) }, await boardKey(k), fromB64url(sealed.ct)));
  const board = JSON.parse(fromUtf8(sealed.z ? await inflate(pt) : pt));
  if (!board || board.v !== 2 || board.e !== e || !Array.isArray(board.oc)) throw new Error('Not a board');
  return board;
}

/**
 * Two cadre devices on one sign-up board: this device's own changes since base (the board it last
 * had from the sheet) go over b, the newer board the other device sent. A row changed here goes
 * over b's; a row added here is added (the next free id if b has one by that id, not at all if b
 * has that phone's newcomer already); twin names, Same OC links, the cadre's placements, the name and the squads
 * per platoon changed here go in too.
 * @returns {{board: object, own: boolean}} own: this device changed something since base
 */
export function mergeBoard(b, base, mine) {
  const out = Object.assign({}, b, { oc: b.oc.map((o) => o.slice()), nd: Object.assign({}, b.nd || {}), fx: (b.fx || []).slice() });
  const was = {};
  base.oc.forEach((o) => {
    was[o[0]] = JSON.stringify(o);
  });
  const ids = {}; // a row added here: its id on the merged board
  let own = false;
  mine.oc.forEach((o) => {
    if (was[o[0]] === JSON.stringify(o)) return;
    own = true;
    const i = out.oc.findIndex((x) => x[0] === o[0]);
    if (was[o[0]] !== undefined) {
      if (i >= 0) out.oc[i] = o.slice();
      return;
    }
    const same = out.oc.find((x) => x[5] && x[5] === o[5] && x[1] === o[1] && (x[2] || '') === (o[2] || ''));
    if (same) {
      ids[o[0]] = same[0];
      return;
    }
    ids[o[0]] = i < 0 ? o[0] : Math.max(...out.oc.map((x) => x[0]), ...mine.oc.map((x) => x[0])) + 1;
    out.oc.push([ids[o[0]]].concat(o.slice(1)));
  });
  const idOf = (k) => (ids[k] !== undefined ? ids[k] : Number(k));
  Object.keys(mine.nd || {}).forEach((k) => {
    if ((base.nd || {})[k] === mine.nd[k]) return;
    own = true;
    out.nd[idOf(k)] = mine.nd[k];
  });
  (mine.fx || []).forEach((k) => {
    if ((base.fx || []).indexOf(k) < 0) own = true;
    if (out.fx.indexOf(idOf(k)) < 0) out.fx.push(idOf(k));
  });
  Object.keys(mine.ln || {}).forEach((k) => { // Same OC: a typed name tied to a row
    if ((base.ln || {})[k] === mine.ln[k]) return;
    own = true;
    out.ln = Object.assign({}, out.ln || {}, { [k]: idOf(mine.ln[k]) });
  });
  ['n', 'sq'].forEach((f) => {
    if (mine[f] === base[f]) return;
    own = true;
    out[f] = mine[f];
  });
  return { board: out, own };
}

/** The squads on a board: OCs placed in a squad, by platoon then squad, ids in order. */
export function squadsOf(board) {
  const by = {};
  board.oc.forEach(([id, , , plt, sqd]) => {
    if (!sqd) return;
    const key = plt + '-' + sqd;
    (by[key] = by[key] || { key, plt, sqd, label: (plt ? ord(plt) + ' PLT · ' : '') + ord(sqd) + ' SQD', ids: [] }).ids.push(id);
  });
  return Object.values(by).sort((a, b) => a.plt - b.plt || a.sqd - b.sqd).map((sq) => Object.assign(sq, { ids: sq.ids.slice().sort((a, b) => a - b) }));
}

/** A squad's deal seed: by its platoon and number on boards from 17b (dl 2), so another squad filling up never re-deals it; by its place among the squads before. */
const squadSeed = (board, sq, si) => board.seed + (board.dl >= 2 ? (sq.plt || 0) * 1009 + sq.sqd * 101 : si * 101);

/**
 * One squad's deal: { k, seed, hands: { id: hand } }. A hand is one list of attribute indexes for
 * every peer, or on boards from 17e (dl 3) { peer id: PAIR_K different attribute indexes }.
 */
function squadDeal(board, sq, si) {
  const seed = squadSeed(board, sq, si);
  const hands = {};
  if (board.dl >= 3) {
    const d = dealPairs(sq.ids.length, seed);
    sq.ids.forEach((id, i) => {
      hands[id] = {};
      sq.ids.forEach((p, j) => {
        if (j !== i) hands[id][p] = d[i][j];
      });
    });
    return { k: PAIR_K, seed, hands };
  }
  const d = dealSquad(sq.ids.length, seed);
  sq.ids.forEach((id, i) => (hands[id] = d.hands[i] || []));
  return { k: d.k, seed: d.seed, hands };
}

/** Each OC's attributes as this board deals them, the same as eventOf: { id: hand } (see squadDeal). */
export function handsOf(board) {
  const hands = {};
  squadsOf(board).forEach((sq, si) => Object.assign(hands, squadDeal(board, sq, si).hands));
  return hands;
}

/**
 * The evaluation as the admin console keeps it (members with tokens and hands, squads with
 * their deal), so results open, check and score with the same code as before.
 * @param {object} board
 * @param {string} k link key
 * @param {object|null} priv the cadre's private JWK, or null on a phone
 */
export async function eventOf(board, k, priv) {
  const hand = {};
  const home = {};
  const squads = squadsOf(board).map((sq, si) => {
    const d = squadDeal(board, sq, si);
    Object.assign(hand, d.hands);
    sq.ids.forEach((id) => (home[id] = sq.key));
    return Object.assign({}, sq, { k: d.k, seed: d.seed });
  });
  const members = [];
  for (const [id, last, first, plt, sqd] of board.oc) {
    members.push({ id, last, first: first || '', name: fullName(last, first), plt, sqd, squad: home[id] || null, token: await ocToken(k, id), hand: hand[id] || [] });
  }
  return {
    v: 2, id: board.e, name: board.n, created: board.at, keys: { pub: board.pub, priv }, seed: board.seed,
    members, squads,
    join: { work: '', plts: [{ plt: 0, token: await joinToken(k) }] }, results: {}
  };
}

/** One OC's packet, the same one a personal link carried, built on the phone from the board. */
export async function packetOf(board, k, id, url) {
  const ev = await eventOf(board, k, null);
  const me = ev.members.find((m) => m.id === id);
  if (!me || !me.squad) return null;
  const sq = ev.squads.find((s) => s.key === me.squad);
  const peers = sq.ids.filter((x) => x !== id).map((x) => [x, ev.members.find((m) => m.id === x).name]);
  return { v: 1, e: board.e, n: board.n, g: sq.label, me: id, nm: me.name, p: peers, a: me.hand, k: board.pub, t: me.token, u: url, one: 1 };
}

/** The day an unused evaluation is removed, from its created time (ISO). */
export const unusedUntil = (created) => new Date(new Date(created).getTime() + UNUSED_DAYS * DAY);
/** The day a used evaluation is removed, from its last answer (ISO). */
export const keptUntil = (last) => new Date(new Date(last).getTime() + KEEP_DAYS * DAY);
/** "15 Dec", or "15 Dec 2027" when the year differs from now. */
export function dayLabel(d, now) {
  const n = now || new Date();
  const opt = { day: 'numeric', month: 'short' };
  if (d.getFullYear() !== n.getFullYear()) opt.year = 'numeric';
  return d.toLocaleDateString('en-GB', opt);
}
