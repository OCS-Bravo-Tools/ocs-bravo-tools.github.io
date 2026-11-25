/**
 * Balanced, seeded dealing of attributes to the raters of one squad.
 *
 * Every rater gets k attributes. Attributes are dealt like cards from decks
 * that spread the six ALRM categories evenly, so each attribute is held by
 * floor(n*k/30) or ceil(n*k/30) raters. With k >= ceil(60/n) every attribute
 * is held by at least two raters, so every OC is still rated on all 30 after
 * their own hand is left out.
 */
import { ATTRIBUTES, ATTRIBUTE_COUNT } from './attributes.js';

/**
 * Smallest k that keeps full coverage for a squad of n OCs (6 to 30).
 * @param {number} n squad size
 * @returns {number}
 */
export function minK(n) {
  if (n <= 2) return ATTRIBUTE_COUNT;
  return Math.min(ATTRIBUTE_COUNT, Math.max(6, Math.ceil((2 * ATTRIBUTE_COUNT) / n)));
}

/**
 * Small deterministic PRNG (mulberry32).
 * @param {number} seed
 * @returns {() => number} values in [0, 1)
 */
export function makeRng(seed) {
  let a = (seed * 2654435761) >>> 0;
  return function () {
    a = (a + 0x6D2B79F5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function shuffle(list, rnd) {
  const a = list.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(rnd() * (i + 1));
    const t = a[i];
    a[i] = a[j];
    a[j] = t;
  }
  return a;
}

/** One shuffled pass of all 30 attributes with categories spread evenly. */
function onePass(rnd) {
  const byCat = {};
  ATTRIBUTES.forEach((attr, id) => {
    (byCat[attr.cat] = byCat[attr.cat] || []).push(id);
  });
  const keyed = [];
  Object.keys(byCat).forEach((cat) => {
    const ids = shuffle(byCat[cat], rnd);
    const off = rnd();
    ids.forEach((id, r) => keyed.push({ id, key: (r + off) / ids.length }));
  });
  keyed.sort((p, q) => p.key - q.key);
  return keyed.map((p) => p.id);
}

/**
 * Deal k attributes to each of n raters.
 * @param {number} n raters in the squad
 * @param {number} k attributes per rater
 * @param {number} seed
 * @returns {number[][]} hands, one sorted list of attribute ids per rater
 */
export function deal(n, k, seed) {
  const rnd = makeRng(seed);
  const passes = Math.ceil((n * k) / ATTRIBUTE_COUNT) + 1;
  let pile = [];
  for (let p = 0; p < passes; p++) pile = pile.concat(onePass(rnd));
  const hands = [];
  let pos = 0;
  for (let r = 0; r < n; r++) {
    const hand = [];
    while (hand.length < k) {
      const passEnd = (Math.floor(pos / ATTRIBUTE_COUNT) + 1) * ATTRIBUTE_COUNT;
      let j = pos;
      while (j < passEnd && hand.indexOf(pile[j]) >= 0) j++;
      if (j >= passEnd) {
        j = pos;
        while (j < pile.length - 1 && hand.indexOf(pile[j]) >= 0) j++;
      }
      const t = pile[pos];
      pile[pos] = pile[j];
      pile[j] = t;
      hand.push(pile[pos]);
      pos++;
    }
    hands.push(hand.sort((x, y) => x - y));
  }
  return hands;
}

/** Attributes a rater scores for each peer on boards from 17e (dl 3). */
export const PAIR_K = 6;

/**
 * 17e (the user, 11-24: "每个oc应该只随机到6个…每一个oc对另一个oc的是随机的…不能是比如说两个fitness"):
 * every rater gets PAIR_K different attributes for each peer, dealt at random for that pair. For each
 * peer, the raters' sets are cut in a random rater order from shuffled passes of all 30 (categories
 * spread), so they differ as far as they can: with 5 raters or fewer no attribute is given twice, and
 * in a squad of 6 every OC is rated once on each of the 30. A set never spans two passes (30 = 5 x 6),
 * so it never holds one attribute twice.
 * @param {number} n squad size
 * @param {number} seed
 * @returns {number[][][]} hands[r][p]: sorted attribute ids rater r scores for peer p ([] for r = p)
 */
export function dealPairs(n, seed) {
  const rnd = makeRng(seed);
  const hands = Array.from({ length: n }, () => Array.from({ length: n }, () => []));
  for (let p = 0; p < n; p++) {
    const raters = shuffle(Array.from({ length: n }, (x, i) => i).filter((r) => r !== p), rnd);
    let pile = [];
    while (pile.length < raters.length * PAIR_K) pile = pile.concat(onePass(rnd));
    raters.forEach((r, i) => {
      hands[r][p] = pile.slice(i * PAIR_K, (i + 1) * PAIR_K).sort((x, y) => x - y);
    });
  }
  return hands;
}

/**
 * Check that every ratee is covered on every attribute by someone else.
 * @param {number[][]} hands
 * @returns {{ok: boolean, gaps: number, minHolders: number, maxHolders: number}}
 */
export function coverage(hands) {
  const n = hands.length;
  const holders = Array.from({ length: ATTRIBUTE_COUNT }, () => []);
  hands.forEach((hand, r) => hand.forEach((id) => holders[id].push(r)));
  let gaps = 0;
  if (n >= 2) {
    for (let j = 0; j < n; j++) {
      holders.forEach((h) => {
        if (!h.some((r) => r !== j)) gaps++;
      });
    }
  }
  const counts = holders.map((h) => h.length);
  return { ok: gaps === 0, gaps, minHolders: Math.min(...counts), maxHolders: Math.max(...counts) };
}

/**
 * Deal for one squad and guarantee coverage: retry seeds, then raise k.
 * @param {number} n squad size
 * @param {number} seed base seed
 * @returns {{k: number, seed: number, hands: number[][]}}
 */
export function dealSquad(n, seed) {
  if (n < 2) return { k: 0, seed, hands: n === 1 ? [[]] : [] };
  for (let k = minK(n); k <= ATTRIBUTE_COUNT; k++) {
    for (let tries = 0; tries < 200; tries++) {
      const s = seed + tries;
      const hands = deal(n, k, s);
      if (coverage(hands).ok) return { k, seed: s, hands };
    }
  }
  const all = Array.from({ length: ATTRIBUTE_COUNT }, (x, i) => i);
  return { k: ATTRIBUTE_COUNT, seed, hands: Array.from({ length: n }, () => all.slice()) };
}
