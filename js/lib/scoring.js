/**
 * Scoring, the same formula as the original sheet, averaged across raters.
 *
 * One rater, one peer (the sheet):
 *   total   = sum of the attribute ratings (30-150 when all 30 are rated)
 *   peerQ   = #Yes of the 4 questions - rank given (1 = best)
 *   overall = total + peerQ
 *
 * All raters, one OC (this app):
 *   A       = mean over attributes of (mean rating that attribute got)
 *   T       = 30 x A, back on the sheet's 30-150 scale
 *   Q       = mean over raters of (#Yes - rank)
 *   overall = T + Q
 * With one rater and all 30 attributes, T, Q and overall equal the sheet.
 *
 * Pure functions, no DOM. Input results are decrypted payloads:
 *   { r: raterId, ans: {peerId: 'YNYY'}, com: {peerId: {lead,need,combat,team}},
 *     sc: {peerId: {attrId: 1..5}}, why: {peerId: {attrId: text}}, rank: [peerId...] }
 */
import { ATTRIBUTES, ATTRIBUTE_COUNT, CATEGORIES, QUESTIONS } from './attributes.js';

const mean = (xs) => (xs.length ? xs.reduce((s, x) => s + x, 0) / xs.length : null);

function sd(xs) {
  if (xs.length < 2) return 0;
  const m = mean(xs);
  return Math.sqrt(xs.reduce((s, x) => s + (x - m) * (x - m), 0) / (xs.length - 1));
}

/**
 * Score one OC from the results of the other raters in the squad.
 * @param {number} id the ratee
 * @param {object[]} results decrypted payloads
 * @returns {object} row
 */
export function scoreOne(id, results) {
  const key = String(id);
  const raters = results.filter((res) => res && res.r !== id && res.sc && res.sc[key]);
  const attrSum = new Array(ATTRIBUTE_COUNT).fill(0);
  const attrN = new Array(ATTRIBUTE_COUNT).fill(0);
  raters.forEach((res) => {
    Object.keys(res.sc[key]).forEach((a) => {
      const v = Number(res.sc[key][a]);
      if (v >= 1 && v <= 5) {
        attrSum[a] += v;
        attrN[a] += 1;
      }
    });
  });
  const attrMean = attrSum.map((s, a) => (attrN[a] ? s / attrN[a] : null));
  const means = attrMean.filter((m) => m !== null);
  const covered = means.length;
  const sumMeans = means.reduce((s, x) => s + x, 0);
  const A = covered ? sumMeans / covered : null;
  const T = covered ? (ATTRIBUTE_COUNT * sumMeans) / covered : null;

  const qs = [];
  const yes = {};
  const no = {};
  QUESTIONS.forEach((q) => {
    yes[q.id] = 0;
    no[q.id] = 0;
  });
  raters.forEach((res) => {
    const ans = String((res.ans && res.ans[key]) || '');
    let y = 0;
    QUESTIONS.forEach((q, i) => {
      if (ans[i] === 'Y') {
        y++;
        yes[q.id]++;
      } else if (ans[i] === 'N') {
        no[q.id]++;
      }
    });
    const rk = Array.isArray(res.rank) ? res.rank.indexOf(id) + 1 : 0;
    if (rk > 0) qs.push({ r: res.r, yes: y, rk });
  });
  const ranks = qs.map((x) => x.rk);
  const Q = qs.length ? mean(qs.map((x) => x.yes - x.rk)) : null;
  const meanRank = mean(ranks);
  const rankSD = sd(ranks);
  const overall = T !== null && Q !== null ? T + Q : null;

  const cat = {};
  CATEGORIES.forEach((c) => {
    const ms = attrMean.filter((m, a) => m !== null && ATTRIBUTES[a].cat === c);
    cat[c] = mean(ms);
  });

  const flags = [];
  const trustNo = no.lead + no.need + no.combat;
  if (trustNo > 0) flags.push(trustNo + (trustNo === 1 ? ' No on trust or combat' : ' Nos on trust or combat'));
  if (no.team > 0) flags.push(no.team + ' No on team player');
  if (qs.length >= 3 && rankSD >= 2.5) flags.push('Split opinion on rank');
  if (qs.length >= 3) {
    const outliers = qs.filter((x) => {
      const others = qs.filter((o) => o !== x).map((o) => o.rk);
      return Math.abs(x.rk - mean(others)) >= 4;
    }).length;
    if (outliers) flags.push(outliers + ' outlier rater' + (outliers > 1 ? 's' : ''));
  }
  if (raters.length && covered < ATTRIBUTE_COUNT) flags.push(ATTRIBUTE_COUNT - covered + ' attributes unrated');

  return { id, raters: raters.length, A, T, Q, overall, meanRank, rankSD, yes, no, attrMean, cat, flags };
}

/**
 * Order of merit: overall high to low; ties go to the better mean rank,
 * then the higher attribute index.
 */
export function meritCompare(x, y) {
  if (y.overall !== x.overall) return y.overall - x.overall;
  if (x.meanRank !== y.meanRank) return x.meanRank - y.meanRank;
  return (y.A || 0) - (x.A || 0);
}

/**
 * Checks on how each rater filled the form (admin-side flags only).
 * @param {object[]} results
 * @returns {object[]} [{r, flags}]
 */
export function raterChecks(results) {
  return results.filter(Boolean).map((res) => {
    const flags = [];
    const scores = [];
    const reasons = [];
    Object.keys(res.sc || {}).forEach((p) => {
      Object.keys(res.sc[p]).forEach((a) => {
        scores.push(Number(res.sc[p][a]));
        const text = String(((res.why || {})[p] || {})[a] || '').trim();
        reasons.push(text);
      });
    });
    if (scores.length >= 10) {
      const counts = {};
      scores.forEach((s) => (counts[s] = (counts[s] || 0) + 1));
      const top = Math.max(...Object.values(counts));
      if (top / scores.length >= 0.9) flags.push('Same score on ' + Math.round((100 * top) / scores.length) + '% of ratings');
    }
    const seen = {};
    reasons.filter((t) => t).forEach((t) => {
      const k = t.toLowerCase().replace(/\s+/g, ' ');
      seen[k] = (seen[k] || 0) + 1;
    });
    const pasted = Object.values(seen).filter((c) => c >= 3).length;
    if (pasted) flags.push('Copy-paste reasons (' + pasted + ' repeated)');
    const short = reasons.filter((t) => t && t.split(/\s+/).length <= 2).length;
    if (short) flags.push(short + ' very short reason' + (short > 1 ? 's' : ''));
    const missing = reasons.filter((t) => !t).length;
    if (missing) flags.push(missing + ' score' + (missing > 1 ? 's' : '') + ' without a reason');
    return { r: res.r, flags };
  });
}

/**
 * Score a whole squad.
 * @param {number[]} memberIds everyone in the squad
 * @param {object[]} results decrypted payloads from raters in this squad
 * @returns {{rows: object[], raterChecks: object[]}} rows in order of merit
 */
export function scoreSquad(memberIds, results) {
  const rows = memberIds.map((id) => scoreOne(id, results));
  const scored = rows.filter((row) => row.overall !== null).sort(meritCompare);
  scored.forEach((row, i) => {
    row.merit = i + 1;
  });
  const unscored = rows.filter((row) => row.overall === null);
  unscored.forEach((row) => {
    row.merit = null;
  });
  return { rows: scored.concat(unscored), raterChecks: raterChecks(results) };
}

/**
 * Every comment and reason one OC received, without rater names.
 * @param {number} id the ratee
 * @param {object[]} results
 * @returns {{where: string, score: number|null, text: string}[]}
 */
export function commentsFor(id, results) {
  const key = String(id);
  const out = [];
  results.forEach((res) => {
    if (!res || res.r === id) return;
    const com = (res.com || {})[key] || {};
    QUESTIONS.forEach((q, i) => {
      const text = String(com[q.id] || '').trim();
      if (text) {
        const a = String((res.ans || {})[key] || '')[i];
        out.push({ where: q.id, answer: a === 'N' ? 'No' : 'Yes', score: null, text });
      }
    });
    const sc = (res.sc || {})[key] || {};
    const why = (res.why || {})[key] || {};
    Object.keys(sc).forEach((a) => {
      const text = String(why[a] || '').trim();
      if (text) out.push({ where: ATTRIBUTES[a].name, answer: null, score: Number(sc[a]), text });
    });
  });
  return out;
}
