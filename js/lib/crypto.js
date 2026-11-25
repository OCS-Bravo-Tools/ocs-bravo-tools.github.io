/**
 * WebCrypto helpers (browser and Node 18+).
 *
 * Results are sealed for the admin with ECDH P-256 (a fresh key per result)
 * + HKDF-SHA256 + AES-GCM-256. Each OC's link carries a random token; the
 * HMAC-SHA256 of the envelope with that token proves which link sent it.
 */
import { b64url, fromB64url, utf8 } from './codec.js';

const subtle = globalThis.crypto.subtle;
const CURVE = { name: 'ECDH', namedCurve: 'P-256' };
const INFO = utf8('peer-eval v1');

/** @param {number} n @returns {Uint8Array} */
export function randomBytes(n) {
  const b = new Uint8Array(n);
  globalThis.crypto.getRandomValues(b);
  return b;
}

/** Random URL-safe string from n bytes. */
export const randomToken = (n = 16) => b64url(randomBytes(n));

/**
 * New admin key pair for an event.
 * @returns {Promise<{pub: string, priv: object}>} raw public key (base64url), private JWK
 */
export async function generateAdminKeys() {
  const kp = await subtle.generateKey(CURVE, true, ['deriveBits']);
  const pub = new Uint8Array(await subtle.exportKey('raw', kp.publicKey));
  const priv = await subtle.exportKey('jwk', kp.privateKey);
  return { pub: b64url(pub), priv };
}

async function aesKey(privateKey, publicKey, salt) {
  const bits = await subtle.deriveBits({ name: 'ECDH', public: publicKey }, privateKey, 256);
  const hk = await subtle.importKey('raw', bits, 'HKDF', false, ['deriveKey']);
  return subtle.deriveKey(
    { name: 'HKDF', hash: 'SHA-256', salt, info: INFO },
    hk,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
}

/**
 * Encrypt bytes so only the admin's private key can open them.
 * @param {string} adminPub base64url raw public key
 * @param {Uint8Array} bytes
 * @returns {Promise<{epk: string, iv: string, ct: string}>}
 */
export async function sealFor(adminPub, bytes) {
  const pub = await subtle.importKey('raw', fromB64url(adminPub), CURVE, false, []);
  const eph = await subtle.generateKey(CURVE, true, ['deriveBits']);
  const epkRaw = new Uint8Array(await subtle.exportKey('raw', eph.publicKey));
  const iv = randomBytes(12);
  const key = await aesKey(eph.privateKey, pub, epkRaw);
  const ct = new Uint8Array(await subtle.encrypt({ name: 'AES-GCM', iv }, key, bytes));
  return { epk: b64url(epkRaw), iv: b64url(iv), ct: b64url(ct) };
}

/**
 * Decrypt with the admin's private JWK. Throws if anything was changed.
 * @param {object} privJwk
 * @param {{epk: string, iv: string, ct: string}} sealed
 * @returns {Promise<Uint8Array>}
 */
export async function openWith(privJwk, sealed) {
  const priv = await subtle.importKey('jwk', privJwk, CURVE, false, ['deriveBits']);
  const epkRaw = fromB64url(sealed.epk);
  const epk = await subtle.importKey('raw', epkRaw, CURVE, false, []);
  const key = await aesKey(priv, epk, epkRaw);
  const pt = await subtle.decrypt({ name: 'AES-GCM', iv: fromB64url(sealed.iv) }, key, fromB64url(sealed.ct));
  return new Uint8Array(pt);
}

/**
 * HMAC-SHA256 of text with a base64url token.
 * @returns {Promise<string>} base64url
 */
export async function hmac(token, text) {
  const key = await subtle.importKey('raw', fromB64url(token), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  return b64url(new Uint8Array(await subtle.sign('HMAC', key, utf8(text))));
}

/** SHA-256 of text, base64url. */
export async function sha256(text) {
  return b64url(new Uint8Array(await subtle.digest('SHA-256', utf8(text))));
}
