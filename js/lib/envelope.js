/**
 * Personal links (packets) and sealed results (envelopes).
 *
 * Packet, carried after # in each OC's link (never sent to a server):
 *   { v, e: eventId, n: eventName, g: squad label, me: ocId, nm: display name,
 *     p: [[peerId, name]...], a: [attrIds], k: admin public key, t: token,
 *     u: collector URL ('' = file and paste only) }
 *
 * Envelope, what leaves the phone (the only thing a collector ever stores):
 *   { v, e, r: ocId, rev, z, epk, iv, ct, mac }
 * The payload inside ct is the result JSON described in scoring.js.
 *
 * Squad sign-up (Ask OCs): one link per PLT, after # like a personal link:
 *   { v, j: 1, e, n: eventName, g: PLT label, pl: PLT, sq: squad buttons,
 *     ns: [[ocId, LAST, First]...], k: admin public key, t: link token, u: collector URL }
 * An OC's pick is an envelope like a result, marked j: 1 in the clear, r = the picked OC's id
 * (-1 for a name typed in), signed with the link token. The admin's own work during the
 * sign-up is one more envelope, j: 2 and r = -2, signed with a token that only the admin's
 * files hold, so no link can fake it. Results import skips both.
 */
import { packJSON, unpackJSON, utf8, fromUtf8, deflate, inflate, b64url, fromB64url, canCompress } from './codec.js';
import { sealFor, openWith, hmac, sha256 } from './crypto.js';
import { ATTRIBUTE_COUNT } from './attributes.js';

export const FORMAT = 1;

const CODE = /^[A-Za-z0-9_-]{1,128}$/;
const isCode = (x) => typeof x === 'string' && CODE.test(x);
const isId = (x) => Number.isSafeInteger(x) && x >= 0;
const isText = (x) => typeof x === 'string';

/**
 * Anyone can type a packet, so every field is checked before the page shows or
 * stores any of it: ids are whole numbers, codes are base64url, names are text
 * and every attribute is one of the 30.
 */
function validPacket(p) {
  const isHand = (a) => Array.isArray(a) && a.every((x) => Number.isInteger(x) && x >= 0 && x < ATTRIBUTE_COUNT);
  return !!p && p.v === FORMAT && isCode(p.e) && isId(p.me) && isCode(p.k) && isCode(p.t) &&
    isText(p.n) && isText(p.g) && isText(p.nm) && isText(p.u) &&
    Array.isArray(p.p) && p.p.every((x) => Array.isArray(x) && isId(x[0]) && isText(x[1])) &&
    (isHand(p.a) || (!!p.a && typeof p.a === 'object' && Object.keys(p.a).length === p.p.length && p.p.every((x) => isHand(p.a[x[0]])))); // one list, or a set per peer (17e)
}

/** A sign-up link: the same checks, and at most 300 names, each an id and text. */
function validJoin(p) {
  const n = (x) => Number.isInteger(x) && x >= 1 && x <= 20;
  return !!p && p.v === FORMAT && p.j === 1 && isCode(p.e) && isCode(p.k) && isCode(p.t) &&
    isText(p.n) && isText(p.g) && isText(p.u) && n(p.pl) && n(p.sq) &&
    Array.isArray(p.ns) && p.ns.length <= 300 && p.ns.every((x) => Array.isArray(x) && isId(x[0]) && isText(x[1]) && (x[2] === undefined || isText(x[2])));
}

/** @returns {Promise<string>} the URL-safe code that goes after # */
export const packPacket = (packet) => packJSON(packet);

/**
 * Possible link codes in pasted text, most likely first. After '#': the run of
 * base64url characters there, then that run joined with the next whole lines in
 * case a mail client wrapped the link. Without '#': the longest run.
 * Text around the link (the admin's whole message) is ignored.
 */
function codeCandidates(text) {
  const s = String(text || '').replace(/%23/gi, '#');
  const i = s.indexOf('#');
  if (i < 0) {
    const runs = s.match(/[A-Za-z0-9_-]+/g) || [];
    return runs.length ? [runs.reduce((a, b) => (b.length > a.length ? b : a))] : [];
  }
  const out = [];
  let code = '';
  for (const part of s.slice(i + 1).split(/\s+/).slice(0, 12)) {
    const m = /^[A-Za-z0-9_-]+/.exec(part);
    if (!m) break;
    code += m[0];
    out.push(code);
    if (m[0].length < part.length) break;
  }
  return out;
}

/**
 * Read and check a packet from a link, a hash, a pasted code, or a whole message
 * that contains the link.
 * @param {string} text
 * @returns {Promise<object>}
 */
export async function readPacket(text) {
  for (const code of codeCandidates(text)) {
    try {
      const p = await unpackJSON(code);
      if (validPacket(p) || validJoin(p)) return p;
    } catch (err) {
      // not this candidate
    }
  }
  throw new Error('Not a peer evaluation link');
}

const macText = (env) => [env.v, env.e, env.r, env.rev, env.z, env.epk, env.iv, env.ct].join('.');

/**
 * Seal a result payload for the admin.
 * @param {object} packet the OC's packet
 * @param {object} payload result JSON (must carry rev)
 * @returns {Promise<object>} envelope
 */
export async function sealResult(packet, payload) {
  const raw = utf8(JSON.stringify(payload));
  const z = canCompress ? 1 : 0;
  const body = z ? await deflate(raw) : raw;
  const sealed = await sealFor(packet.k, body);
  const env = { v: FORMAT, e: packet.e, r: packet.me, rev: payload.rev, z, epk: sealed.epk, iv: sealed.iv, ct: sealed.ct };
  env.mac = await hmac(packet.t, macText(env));
  return env;
}

/**
 * Open an envelope with the event's keys and tokens.
 * @param {object} event admin event (id, keys.priv, members[].token)
 * @param {object} env
 * @returns {Promise<{ok: boolean, reason?: string, payload?: object, hash?: string}>}
 */
export async function openResult(event, env) {
  if (!env || env.v !== FORMAT || !env.ct || !env.mac) return { ok: false, reason: 'Not a result' };
  if (env.e !== event.id) return { ok: false, reason: 'From another evaluation' };
  const member = event.members.find((m) => m.id === env.r);
  if (!member) return { ok: false, reason: 'Unknown OC' };
  const mac = await hmac(member.token, macText(env));
  if (mac !== env.mac) return { ok: false, reason: 'Signature does not match' };
  let payload;
  try {
    const body = await openWith(event.keys.priv, env);
    payload = JSON.parse(fromUtf8(env.z ? await inflate(body) : body));
  } catch (err) {
    return { ok: false, reason: 'Cannot be decrypted' };
  }
  if (payload.e !== env.e || payload.r !== env.r || payload.rev !== env.rev) return { ok: false, reason: 'Contents do not match' };
  return { ok: true, payload, hash: await sha256(env.ct) };
}

/** Seal a sign-up pick (j 1) or the admin's work (j 2): like a result, marked in the clear. */
async function sealMarked(pub, token, payload) {
  const raw = utf8(JSON.stringify(payload));
  const z = canCompress ? 1 : 0;
  const body = z ? await deflate(raw) : raw;
  const sealed = await sealFor(pub, body);
  const env = { v: FORMAT, e: payload.e, r: payload.r, rev: payload.rev, z, epk: sealed.epk, iv: sealed.iv, ct: sealed.ct, j: payload.j };
  env.mac = await hmac(token, macText(env));
  return env;
}

/** An OC's pick, from a sign-up link. payload: { v, j: 1, e, r, rev, at, plt, sqd, dv, first?, last?, grp? } */
export const sealJoin = (packet, payload) => sealMarked(packet.k, packet.t, payload);

/** The admin's own work during a sign-up. payload: { v, j: 2, e, r: -2, rev, at, sources, dec } */
export const sealWork = (event, payload) => sealMarked(event.keys.pub, event.join.work, payload);

/**
 * Open a sign-up pick or the admin's work with the event's keys. A pick counts for the PLT
 * whose link token signed it.
 * @returns {Promise<{ok: boolean, payload?: object, plt?: number}>}
 */
export async function openMarked(event, env) {
  if (!env || env.v !== FORMAT || !env.ct || !env.mac || (env.j !== 1 && env.j !== 2) || !event.join || env.e !== event.id) return { ok: false };
  const signers = env.j === 2 ? [{ token: event.join.work, plt: 0 }] : event.join.plts;
  let signer = null;
  for (const s of signers) {
    if ((await hmac(s.token, macText(env))) === env.mac) {
      signer = s;
      break;
    }
  }
  if (!signer) return { ok: false };
  let payload;
  try {
    const body = await openWith(event.keys.priv, env);
    payload = JSON.parse(fromUtf8(env.z ? await inflate(body) : body));
  } catch (err) {
    return { ok: false };
  }
  if (!payload || payload.e !== env.e || payload.r !== env.r || payload.rev !== env.rev || payload.j !== env.j) return { ok: false };
  return { ok: true, payload, plt: signer.plt };
}

const TEXT_TAG = 'PEVAL1:';

/**
 * Envelope as plain text that survives email: a tag plus base64url only, so
 * line wrapping or extra spaces added by a mail client do no harm.
 * @param {object} env
 * @returns {string}
 */
export function envelopeToText(env) {
  return TEXT_TAG + b64url(utf8(JSON.stringify(env))) + '.';
}

/**
 * Find envelopes in any text: PEVAL1 text blocks, JSON objects, a .txt or
 * .peval file, a pasted email (also wrapped or quoted with > in a reply), or the
 * CSV downloaded from the collector sheet.
 * @param {string} text
 * @returns {object[]}
 */
export function findEnvelopes(text) {
  const out = [];
  const tagged = /PEVAL1:([A-Za-z0-9_\-\s>]+)\./g;
  let t;
  while ((t = tagged.exec(String(text || '')))) {
    try {
      const obj = JSON.parse(fromUtf8(fromB64url(t[1].replace(/[\s>]+/g, ''))));
      if (obj && obj.ct && obj.mac) out.push(obj);
    } catch (err) {
      // damaged block; skip
    }
  }
  return out.concat(findJsonEnvelopes(text));
}

function findJsonEnvelopes(text) {
  let s = String(text || '');
  if (s.indexOf('""v""') >= 0 || s.indexOf('""ct""') >= 0) s = s.replace(/""/g, '"');
  const out = [];
  const re = /\{"v":\d+,"e":"[^"]*"/g;
  let m;
  while ((m = re.exec(s))) {
    let depth = 0;
    let end = -1;
    for (let i = m.index; i < s.length; i++) {
      const c = s[i];
      if (c === '"') {
        i++;
        while (i < s.length && s[i] !== '"') {
          if (s[i] === '\\') i++;
          i++;
        }
      } else if (c === '{') {
        depth++;
      } else if (c === '}') {
        depth--;
        if (depth === 0) {
          end = i;
          break;
        }
      }
    }
    if (end < 0) break;
    try {
      const obj = JSON.parse(s.slice(m.index, end + 1));
      if (obj && obj.ct && obj.mac) out.push(obj);
    } catch (err) {
      // not an envelope; skip
    }
    re.lastIndex = end + 1;
  }
  return out;
}

export { b64url, fromB64url };
