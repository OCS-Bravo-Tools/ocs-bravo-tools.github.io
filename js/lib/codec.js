/**
 * Byte and text helpers: UTF-8, base64url, deflate, and packed JSON.
 * Works in browsers and in Node 18+ (both have CompressionStream).
 */

const enc = new TextEncoder();
const dec = new TextDecoder();

export const utf8 = (s) => enc.encode(s);
export const fromUtf8 = (b) => dec.decode(b);

/** @param {Uint8Array} bytes @returns {string} base64url without padding */
export function b64url(bytes) {
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

/** @param {string} s base64url @returns {Uint8Array} */
export function fromB64url(s) {
  const b64 = s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** True when this browser can deflate (Chrome/Edge 103+, Safari 16.4+, Firefox 113+). */
export const canCompress = (() => {
  try {
    new CompressionStream('deflate-raw');
    new DecompressionStream('deflate-raw');
    return true;
  } catch (err) {
    return false;
  }
})();

/** Push the bytes in and read the result out at the same time (no Blob, no backpressure stall). */
async function pipe(bytes, stream) {
  const writer = stream.writable.getWriter();
  writer.write(bytes).catch(() => {});
  writer.close().catch(() => {});
  return new Uint8Array(await new Response(stream.readable).arrayBuffer());
}

/** @param {Uint8Array} bytes @returns {Promise<Uint8Array>} */
export const deflate = (bytes) => pipe(bytes, new CompressionStream('deflate-raw'));

/** @param {Uint8Array} bytes @returns {Promise<Uint8Array>} */
export const inflate = (bytes) => pipe(bytes, new DecompressionStream('deflate-raw'));

/**
 * Pack an object into a short URL-safe string.
 * The first character says how: 'z' deflated, 'u' plain.
 * @param {object} obj
 * @returns {Promise<string>}
 */
export async function packJSON(obj) {
  const raw = utf8(JSON.stringify(obj));
  if (canCompress) return 'z' + b64url(await deflate(raw));
  return 'u' + b64url(raw);
}

/**
 * @param {string} s output of packJSON
 * @returns {Promise<object>}
 */
export async function unpackJSON(s) {
  const kind = s[0];
  const bytes = fromB64url(s.slice(1));
  if (kind === 'z') return JSON.parse(fromUtf8(await inflate(bytes)));
  if (kind === 'u') return JSON.parse(fromUtf8(bytes));
  throw new Error('Unknown packing');
}
