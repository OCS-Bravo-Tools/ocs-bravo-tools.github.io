/**
 * Split long text into several QR codes shown in rotation, and put them back
 * together on the admin side in any order.
 *
 * Part format: "PQ1|<id>|<i>/<n>|<chunk>", id = first 6 characters of a hash
 * of the whole text, so parts from two phones never mix.
 */

function shortId(text) {
  let h = 2166136261 >>> 0;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 16777619) >>> 0;
  }
  return h.toString(36).slice(0, 6);
}

/**
 * @param {string} text
 * @param {number} size characters per part (about 0.6 KB keeps each code easy to scan)
 * @returns {string[]}
 */
export function splitParts(text, size = 600) {
  const id = shortId(text);
  const n = Math.max(1, Math.ceil(text.length / size));
  const parts = [];
  for (let i = 0; i < n; i++) {
    parts.push('PQ1|' + id + '|' + (i + 1) + '/' + n + '|' + text.slice(i * size, (i + 1) * size));
  }
  return parts;
}

/**
 * Collects scanned parts; add() returns progress, and text when complete.
 */
export function makeAssembler() {
  const groups = {};
  return {
    add(raw) {
      const m = /^PQ1\|([0-9a-z]+)\|(\d+)\/(\d+)\|([\s\S]*)$/.exec(String(raw || ''));
      if (!m) return { ok: false };
      const id = m[1];
      const i = Number(m[2]);
      const n = Number(m[3]);
      const g = (groups[id] = groups[id] || { n, parts: {} });
      g.parts[i] = m[4];
      const have = Object.keys(g.parts).length;
      if (have < g.n) return { ok: true, id, have, total: g.n };
      let text = '';
      for (let k = 1; k <= g.n; k++) text += g.parts[k];
      delete groups[id];
      return { ok: true, id, have, total: g.n, text };
    }
  };
}
