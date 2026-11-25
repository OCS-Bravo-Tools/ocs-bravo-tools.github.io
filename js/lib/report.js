/**
 * The Excel files, shared by the admin console, the cadre page and the OC page so all give exactly
 * the same workbook (the first two moved here unchanged from admin.js, Build 10):
 *   resultsXlsx   Export: one ranked tab per squad, then Comments and Raw
 *   answersXlsx   Save answers so far: an Index tab, then one tab per OC whose result is in
 *   squadBook     the original paper form of one squad (1PLT_2nd_Squad_Peer_Evaluations.xlsx) for one rater,
 *                 with that OC's own attributes: blank to fill in (blankBooks, one per OC), or with their answers (16u, 16z)
 *   rosterXlsx    the cadre's roster (Who's in it?), to edit in Excel and drop back in later (16v)
 *   feedbackBooks each OC's own feedback, anonymous, with charts of strengths and what to work on (16z)
 * All hold names, scores and comments: they stay on the cadre's or the OC's own device.
 * E is the evaluation as the console keeps it: { name, members, squads, results }.
 */
import { ATTRIBUTES, QUESTIONS, SCALE, CATEGORIES } from './attributes.js';
import { scoreSquad, commentsFor } from './scoring.js';
import { buildXlsx, zip, STYLE } from './xlsx.js';
import { ord } from './board.js';

export const QUESTION_LABEL = { lead: 'Trust to lead soldiers', need: 'Trust in a time of need', combat: 'Follow into combat', team: 'Team player' };
export const SHEET_BOX = { lead: 'Trust', need: 'Trust', combat: 'Follow', team: 'Additional comments' };
export const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const slug = (s) => String(s).replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '');
const memberOf = (E, id) => E.members.find((m) => m.id === id);
const nameOf = (E, id) => (memberOf(E, id) ? memberOf(E, id).name : '?');
const squadResults = (E, sq) => sq.ids.map((id) => E.results[id]).filter(Boolean).map((r) => r.payload);

const TRUST_NO = /^(\d+) Nos? on trust or combat$/;
const TEAM_NO = /^(\d+) No on team player$/;

/**
 * One OC's notes on the results, in raters, not answers: "1 of 5 said No on trust or follow" (one rater's
 * three Nos count once), "Raters disagree on rank". Excel (all) keeps the number of Nos and the other flags.
 * r = the scored row (scoring.js), res = the squad's payloads.
 */
export function rowNotes(r, res, all) {
  const key = String(r.id);
  const raters = res.filter((p) => p && p.r !== r.id && p.sc && p.sc[key]); // the raters scoring.js counts
  const ans = (p) => String((p.ans || {})[key] || '');
  const out = [];
  (r.flags || []).forEach((f) => {
    const m = TRUST_NO.exec(f);
    if (m) return out.push(raters.filter((p) => /N/.test(ans(p).slice(0, 3))).length + ' of ' + r.raters + ' said No on trust or follow' + (all ? ' (' + m[1] + (m[1] === '1' ? ' No)' : ' Nos)') : ''));
    if (TEAM_NO.test(f)) return out.push(raters.filter((p) => ans(p)[3] === 'N').length + ' of ' + r.raters + ' said No on team player');
    if (f === 'Split opinion on rank') return out.push('Raters disagree on rank');
    return all ? out.push(f) : 0;
  });
  return out;
}

/** How much lower (negative) or higher this rater scored each peer than the rest of the squad did, on average; null under 3 peers to compare. */
export function lowScoreGap(res, id) {
  const me = res.find((p) => p && p.r === id);
  if (!me || !me.sc) return null;
  const avg = (xs) => xs.reduce((s, x) => s + x, 0) / xs.length;
  const gaps = [];
  Object.keys(me.sc).forEach((peer) => {
    const mine = Object.values(me.sc[peer] || {}).map(Number).filter(Boolean);
    const others = [].concat(...res.filter((p) => p && p.r !== id && p.sc && p.sc[peer]).map((p) => Object.values(p.sc[peer]).map(Number).filter(Boolean)));
    if (mine.length && others.length) gaps.push(avg(mine) - avg(others));
  });
  return gaps.length >= 3 ? avg(gaps) : null;
}

/** How this OC filled in the form, flagged in plain words (never in Overall or the order): far lower scores than the squad, two devices, one score nearly everywhere, one reason reused. */
export function formChecks(result, checks, id, res) {
  const out = [];
  const gap = res ? lowScoreGap(res, id) : null;
  if (gap !== null && gap <= -1) out.push('Scored others far lower than the squad did (' + (-gap).toFixed(1) + ' lower on average)');
  const devs = ((result && result.devs) || []).length;
  if (devs > 1) out.push('Sent from ' + devs + ' devices');
  const c = (checks || []).find((x) => x.r === id);
  (c ? c.flags : []).forEach((f) => {
    if (/^Same score on /.test(f)) out.push('Gave almost every rating the same score');
  });
  const p = result && result.payload;
  const seen = {}; // every reason this OC wrote, the Yes/No ones too (16y)
  (p ? [p.why, p.com] : []).forEach((box) => Object.values(box || {}).forEach((x) => Object.values(x || {}).forEach((t) => {
    const k = String(t).trim().toLowerCase().replace(/\s+/g, ' ');
    if (k) seen[k] = (seen[k] || 0) + 1;
  })));
  if (Object.values(seen).some((n) => n >= 3)) out.push('Reused one reason 3+ times');
  return out;
}

/** Export: { name, data } of the results workbook. */
export function resultsXlsx(E) {
  const sheets = [];
  E.squads.forEach((sq) => {
    const res = squadResults(E, sq);
    const { rows, raterChecks } = scoreSquad(sq.ids, res);
    const data = [['Rank', 'Candidate', 'Average attribute score', 'Total attribute score', 'Peer questions score', 'Overall', 'Raters', 'Look closer']];
    rows.forEach((r) => data.push([
      r.merit || '', nameOf(E, r.id),
      r.A === null ? '' : Number(r.A.toFixed(3)), r.T === null ? '' : Number(r.T.toFixed(2)),
      r.Q === null ? '' : Number(r.Q.toFixed(2)), r.overall === null ? '' : Number(r.overall.toFixed(2)),
      r.raters, rowNotes(r, res, true).concat(formChecks(E.results[r.id], raterChecks, r.id, res)).join('; ')
    ]));
    sheets.push({ name: sq.label.replace(' · ', ' '), rows: data, widths: [6, 22, 14, 14, 14, 10, 8, 44] });
  });
  const com = [['Squad', 'Candidate', 'Question or attribute', 'Answer or score', 'Sheet box', 'Comment']];
  const raw = [['Squad', 'Rater', 'Candidate', 'Lead', 'Need', 'Combat', 'Team', 'Rank given'].concat(ATTRIBUTES.map((a) => a.name))];
  E.squads.forEach((sq) => {
    const res = squadResults(E, sq);
    sq.ids.forEach((id) => {
      commentsFor(id, res).forEach((c) => {
        const q = QUESTION_LABEL[c.where];
        com.push([sq.label, nameOf(E, id), q || c.where, q ? c.answer : c.score, q ? SHEET_BOX[c.where] : 'Attribute comment', c.text]);
      });
    });
    res.forEach((p) => {
      sq.ids.filter((id) => id !== p.r).forEach((id) => {
        const ans = String((p.ans || {})[id] || '');
        const sc = (p.sc || {})[id] || {};
        raw.push([sq.label, nameOf(E, p.r), nameOf(E, id)].concat(QUESTIONS.map((q, i) => (ans[i] === 'Y' ? 'Yes' : ans[i] === 'N' ? 'No' : '')), [(p.rank || []).indexOf(id) + 1 || ''], ATTRIBUTES.map((a, ai) => sc[ai] || '')));
      });
    });
  });
  sheets.push({ name: 'Comments', rows: com, widths: [18, 20, 26, 12, 18, 80] });
  sheets.push({ name: 'Raw', rows: raw });
  return { name: 'peer-eval-' + slug(E.name) + '-' + new Date().toISOString().slice(0, 10) + '.xlsx', data: buildXlsx(sheets) };
}

/**
 * Save answers so far: { name, data }. c = { done, total } results in; now = the save time.
 * Read straight from the opened results, with no scoring and no formulas.
 */
export function answersXlsx(E, c, now) {
  const pad = (n) => String(n).padStart(2, '0');
  const full = (m) => m.last + (m.first ? ', ' + m.first : '');
  const yn = (ans, i) => (ans[i] === 'Y' ? 'Yes' : ans[i] === 'N' ? 'No' : '');
  const index = [
    ['Answers so far · ' + E.name],
    [c.done + ' of ' + c.total + ' results in · saved ' + now.toLocaleString()],
    ['Each tab is one OC\'s evaluation, as they sent it (the newest version). Scores: ' + SCALE.slice(1).map((x, i) => (i + 1) + ' ' + x).join(', ') + '. Not scored: the order of merit is in Export xlsx.'],
    [],
    ['Squad', 'OC', 'In', 'Sent', 'Version']
  ];
  const sheets = [{ name: 'Index', rows: index, widths: [20, 26, 8, 24, 10], bold: [4] }];
  E.squads.filter((sq) => sq.ids.length > 1).forEach((sq) => {
    sq.ids.forEach((rid) => {
      const m = memberOf(E, rid);
      const r = E.results[rid];
      index.push([sq.label, full(m), r ? 'Yes' : 'Not yet', r ? new Date(r.at).toLocaleString() : '', r ? r.rev : '']);
      if (!r) return;
      const p = r.payload;
      const peers = sq.ids.filter((id) => id !== rid);
      const rated = [...new Set([].concat(...Object.values(p.sc || {}).map((x) => Object.keys(x).map(Number))))].sort((x, y) => x - y);
      const dealt = Array.isArray(m.hand) ? m.hand : [...new Set([].concat(...Object.values(m.hand || {})))].sort((x, y) => x - y); // one list, or a set per peer (17e)
      const hand = rated.length ? rated : dealt; // what the OC rated: the deal, or attributes they swapped in
      const devs = (r.devs || []).length;
      const rows = [
        ['Evaluation by ' + full(m)],
        ['Squad', sq.label],
        ['Sent', new Date(r.at).toLocaleString() + (r.rev > 1 ? ' · version ' + r.rev : '') + (devs > 1 ? ' · sent from ' + devs + ' devices' : '')],
        [],
        ['Candidate'].concat(QUESTIONS.map((q) => QUESTION_LABEL[q.id]), ['Rank given'], hand.map((a) => ATTRIBUTES[a].name))
      ];
      peers.forEach((id) => {
        const ans = String((p.ans || {})[id] || '');
        const sc = (p.sc || {})[id] || {};
        rows.push([nameOf(E, id)].concat(QUESTIONS.map((q, i) => yn(ans, i)), [(p.rank || []).indexOf(id) + 1 || ''], hand.map((a) => sc[a] || '')));
      });
      const notes = rows.length + 1;
      rows.push([], ['Candidate', 'Question or attribute', 'Answer or score', 'Comment or reason']);
      peers.forEach((id) => {
        const ans = String((p.ans || {})[id] || '');
        const sc = (p.sc || {})[id] || {};
        const com = (p.com || {})[id] || {};
        const why = (p.why || {})[id] || {};
        QUESTIONS.forEach((q, i) => {
          if (com[q.id]) rows.push([nameOf(E, id), QUESTION_LABEL[q.id], yn(ans, i), com[q.id]]);
        });
        hand.forEach((a) => {
          if (why[a]) rows.push([nameOf(E, id), ATTRIBUTES[a].name, sc[a] || '', why[a]]);
        });
      });
      sheets.push({ name: sq.plt + '-' + sq.sqd + ' ' + m.name, rows, widths: [22, 24, 22, 22, 22, 10].concat(hand.map(() => 18)), bold: [4, notes] });
    });
  });
  const stamp = now.getFullYear() + '-' + pad(now.getMonth() + 1) + '-' + pad(now.getDate()) + '-' + pad(now.getHours()) + pad(now.getMinutes());
  return { name: 'peer-eval-' + slug(E.name) + '-answers-' + stamp + '.xlsx', data: buildXlsx(sheets) };
}

/* ---------- the original paper form ---------- */

const LEVELS = SCALE.slice(1); // the sheet's rating dropdown, 1 to 5
const DASH = 'SQD_Summary_Dashboard';
const cap = (s) => String(s).charAt(0).toUpperCase() + String(s).slice(1).toLowerCase();
const fullOf = (o) => o.last + (o.first ? ', ' + o.first : '');

/** One rater's answers for one OC (a sent payload), in the boxes of the paper form. */
export function marksOf(p, id) {
  const ans = String((p.ans || {})[id] || '');
  const com = (p.com || {})[id] || {};
  return {
    yes: QUESTIONS.map((q, i) => (ans[i] === 'Y' ? 'Yes' : ans[i] === 'N' ? 'No' : '')),
    trust: ['lead', 'need'].filter((q) => com[q]).map((q) => QUESTION_LABEL[q] + ': ' + com[q]).join('\n'),
    follow: com.combat || '',
    more: com.team ? QUESTION_LABEL.team + ': ' + com.team : '',
    rank: (p.rank || []).indexOf(id) + 1 || '',
    sc: (p.sc || {})[id] || {},
    why: (p.why || {})[id] || {}
  };
}

/**
 * Filled-in paper forms read back (17h, the user: "OC导出excel，最后再导入。以及cadre页面也支持导入excel"): every Eval tab
 * of a workbook from readXlsx (this app's form, or the original squad sheet, same cells) as
 * { rater, plt, sqd, peer, ans: ['Y'|'N'|''] x4, com: { lead, need, combat, team }, rank, sc: { attr: 1-5 }, why: { attr: text } }.
 * Names stay as written (matchName finds them on a roster). A tab with nothing filled in is left out; an attribute
 * not on the list, or one written twice, counts once (the first).
 */
export function readForms(book) {
  const at = (c, ref) => String(((c[Number(ref.slice(1)) - 1] || [])[ref.charCodeAt(0) - 65]) ?? '').trim();
  const num = (s) => Number((/(\d+)/.exec(s) || [])[1]) || 0;
  const yn = (s) => (/^y/i.test(s) ? 'Y' : /^n/i.test(s) ? 'N' : '');
  const byName = {};
  ATTRIBUTES.forEach((a, i) => (byName[a.name.toLowerCase()] = i));
  const level = (s) => {
    const i = SCALE.findIndex((x, k) => k > 0 && x.toLowerCase() === s.toLowerCase());
    return i > 0 ? i : /^[1-5]$/.test(s) ? Number(s) : 0;
  };
  const labels = Object.entries(QUESTION_LABEL);
  const split = (text) => { // "Trust to lead soldiers: …" lines as the app writes them; other text goes to _ (one box, both questions)
    const out = {};
    let cur = '_';
    String(text).split(/\r?\n/).forEach((line) => {
      const hit = labels.find(([, l]) => line.toLowerCase().startsWith(l.toLowerCase() + ':'));
      if (hit) {
        cur = hit[0];
        out[cur] = line.slice(hit[1].length + 1).trim();
      } else if (line.trim()) out[cur] = ((out[cur] || '') + ' ' + line.trim()).trim();
    });
    return out;
  };
  const forms = [];
  book.forEach((t) => {
    const c = t.cells || [];
    if (!/^select your last name$/i.test(at(c, 'D4'))) return;
    const rater = at(c, 'E4');
    const peer = at(c, 'E5');
    if (!peer) return; // a filled tab with no name picked stays, rater '' (said so on import)
    const trust = split(at(c, 'D10'));
    const follow = split(at(c, 'F10'));
    const more = split(at(c, 'H8'));
    const com = {};
    [['lead', trust.lead || trust._], ['need', trust.need || trust._], ['combat', follow.combat || follow._], ['team', more.team || more._]].forEach(([q, v]) => {
      if (v) com[q] = v;
    });
    const sc = {};
    const why = {};
    for (let r = 23; r < 23 + ATTRIBUTES.length; r++) {
      const ai = byName[at(c, 'D' + r).toLowerCase()];
      const v = level(at(c, 'E' + r));
      if (ai === undefined || sc[ai] || !v) continue;
      sc[ai] = v;
      if (at(c, 'F' + r)) why[ai] = at(c, 'F' + r);
    }
    const ans = [at(c, 'E7'), at(c, 'G7'), at(c, 'G9'), at(c, 'G13')].map(yn);
    if (!ans.some(Boolean) && !Object.keys(sc).length) return; // a blank tab
    forms.push({ rater, plt: num(at(c, 'E2')), sqd: num(at(c, 'E3')), peer, ans, com, rank: num(at(c, 'E13')), sc, why });
  });
  return forms;
}

/** A name as written on a form ("LAST, First", "LAST", any case) to one of members [{ id, last, first }], or null if none or two fit. */
export function matchName(written, members) {
  const n = (s) => String(s || '').toUpperCase().replace(/[^A-Z]/g, '');
  const w = String(written || '');
  const i = w.indexOf(',');
  const l = i < 0 ? w : w.slice(0, i);
  const f = i < 0 ? '' : w.slice(i + 1);
  const exact = members.filter((m) => n(m.last) === n(l) && n(m.first) === n(f));
  if (exact.length === 1) return exact[0];
  const near = members.filter((m) => n(m.last) === n(l) && (!n(f) || !n(m.first) || n(m.first)[0] === n(f)[0]));
  return near.length === 1 ? near[0] : null;
}

/** Answers of one rater, from their form tabs matched to peers [{ id, last, first }]: the payload's ans, com, sc, why and rank. */
export function formAnswers(tabs, peers) {
  const out = { ans: {}, com: {}, sc: {}, why: {}, rank: [], n: 0 };
  const ranked = [];
  tabs.forEach((f) => {
    const m = matchName(f.peer, peers);
    if (!m) return;
    out.n++;
    out.ans[m.id] = f.ans.map((x) => x || '-').join('');
    if (Object.keys(f.com).length) out.com[m.id] = f.com;
    out.sc[m.id] = f.sc;
    out.why[m.id] = f.why;
    if (f.rank) ranked.push([m.id, f.rank]);
  });
  out.rank = ranked.sort((x, y) => x[1] - y[1]).map((x) => x[0]); // only ranks written on the form: one never given is not made up, and scoring leaves that peer's Yes/No-and-rank part out for this rater
  return out;
}

/**
 * The original paper workbook of one squad, as in 1PLT_2nd_Squad_Peer_Evaluations.xlsx: the
 * SQD_Summary_Dashboard tab, then one Eval tab per OC, with the same cells, dropdowns and
 * formulas (no DODID). sq = { plt, sqd } (0: none); ocs = [{ id, last, first, m }], m = marksOf(...),
 * or none for a blank form; rater = { last, first } or null; note = a line under the dashboard;
 * attrs = the rater's attributes: { peer id: list } (17e, one set per peer), one list for every peer, or none (all 30);
 * pick = { seed } (17i): one workbook for the whole squad. The rater picks his name once (the dashboard's Rater Name, a
 * dropdown of the squad); every tab's E4 follows it, and each tab's 6 attributes are formulas of his place and the
 * peer's in the squad: the same name always gives the same 6, never one twice on a tab, and per peer the raters' sixes
 * differ as far as they can (a peer's own order of the 30, MOD(a*x+b,30) with a coprime to 30, block (r-p) mod n).
 * @returns {{ name: string, data: Uint8Array }}
 */
export function squadBook(sq, ocs, rater, note, attrs, pick) {
  const S = STYLE;
  const ALL = ATTRIBUTES.map((a, i) => i);
  if (pick) attrs = Object.fromEntries(ocs.map((o) => [o.id, ALL.slice(0, 6)])); // 6 rows a tab; the names come from the formulas
  const listOf = (o) => { // this rater's attributes on that peer's tab (the deal, with swaps), or all 30
    const x = attrs && !Array.isArray(attrs) ? attrs[o.id] : attrs;
    return x && x.length ? x : ALL;
  };
  const end = 22 + (ocs.length ? Math.max(...ocs.map((o) => listOf(o).length)) : ALL.length); // the last attribute row
  const all = rater ? ocs.concat([rater]) : ocs;
  const twin = (o) => all.filter((x) => x.last.toLowerCase() === o.last.toLowerCase()).length > 1; // the rater counts: Drew's tab on Casey's form says Drew
  const used = {};
  const tabs = ocs.map((o) => { // Eval_<Last> as in the original; letters only, so a formula can quote it plainly
    const base = ('Eval_' + cap(o.last.replace(/[^A-Za-z0-9]/g, '')) + (twin(o) && o.first ? '_' + cap(o.first.replace(/[^A-Za-z0-9]/g, '')) : '')).slice(0, 28);
    let t = base;
    for (let i = 2; used[t.toLowerCase()]; i++) t = base + '_' + i;
    used[t.toLowerCase()] = true;
    return t;
  });
  const grid = (h, w) => Array.from({ length: h }, () => Array(w).fill(null));
  const at = (ref) => { const m = /^([A-Z])(\d+)$/.exec(ref); return [Number(m[2]) - 1, m[1].charCodeAt(0) - 65]; };
  const put = (g, ref, cell) => { const [r, c] = at(ref); g[r][c] = cell; };
  const merged = (g, list) => list.forEach((rg) => { // every cell of a merged range styled, so the border goes all round
    const [a, b] = rg.split(':').map(at);
    for (let r = a[0]; r <= b[0]; r++) for (let c = a[1]; c <= b[1]; c++) if (!g[r][c]) g[r][c] = { s: g[a[0]][a[1]].s };
  });
  const plt = sq.plt ? ord(sq.plt) + ' PLT' : '';
  const sqd = sq.sqd ? ord(sq.sqd) + ' SQD' : '';
  const sheets = [];
  const calc = ocs.map((o) => { // every formula's result, worked out here too: a phone's preview shows them without calculating
    const m = o.m || { yes: [], sc: {} };
    const pts = listOf(o).map((ai) => Number(m.sc[ai]) || 0);
    const total = pts.reduce((x, y) => x + y, 0);
    const n = pts.filter(Boolean).length;
    const q = (m.yes || []).filter((y) => y === 'Yes').length - (Number(m.rank) || 0);
    return { pts, total, avg: n ? total / n : '', q, overall: total + q };
  });
  const last = 3 + ocs.length; // the dashboard's last name row
  const dash = grid(last + 2, 5);
  put(dash, 'A1', { v: (sq.plt ? sq.plt + 'PLT ' : '') + (sq.sqd ? ord(sq.sqd) + ' ' : '') + 'SQUAD PERFORMANCE SUMMARY DASHBOARD', s: S.title });
  put(dash, 'A2', { v: 'Rater Name', s: S.rater });
  put(dash, 'B2', { v: rater ? fullOf(rater) : '', s: pick ? 2 : S.val }); // 2: the yellow of a cell to fill in
  ['Rank', 'Candidate Name(s)', 'Average Score Attribute Score', 'Total Attribute Score', 'Overall Ranking'].forEach((t, c) => (dash[2][c] = { v: t, s: S.head }));
  ocs.forEach((o, i) => {
    const r = 4 + i;
    const t = "'" + tabs[i] + "'!";
    const c = calc[i];
    dash[r - 1] = [
      { f: 'RANK(E' + r + ',$E$4:$E$' + last + ',0)', v: 1 + calc.filter((x) => x.overall > c.overall).length, s: S.mid },
      { v: fullOf(o), s: S.val },
      { f: 'IFERROR(AVERAGE(' + t + 'J23:J' + end + '),"")', v: c.avg, s: S.mid },
      { f: 'SUM(' + t + 'J23:J' + end + ')', v: c.total, s: S.mid },
      { f: 'SUM(' + t + '$J$23:$J$' + end + ')+' + t + '$G$14', v: c.overall, s: S.mid }
    ];
  });
  const notes = note ? ['A' + (last + 2) + ':E' + (last + 2)] : [];
  if (note) put(dash, 'A' + (last + 2), { v: note, s: S.box });
  merged(dash, ['A1:E1'].concat(notes));
  sheets.push({ name: DASH, rows: dash, widths: [10.71, 24, 26.57, 18, 14.29], header: false, rowHeight: 15, heights: Object.assign({ 0: 21 }, note ? { [last + 1]: pick ? 78 : 48 } : {}), merge: ['A1:E1'].concat(notes),
    lists: pick ? [{ ref: 'B2', range: '$B$4:$B$' + last, promptTitle: 'Your name', prompt: 'Pick your own name first: every tab then shows the 6 attributes you rate for that squadmate.' }] : [] });
  const ME = "'" + DASH + "'!$B$2";
  const MATCH_ME = 'MATCH(' + ME + ",'" + DASH + "'!$B$4:$B$" + last + ',0)-1'; // the rater's place in the squad, from 0
  const COPRIME = [1, 7, 11, 13, 17, 19, 23, 29];
  ocs.forEach((o, i) => {
    const g = grid(end, 10);
    const m = o.m || { yes: [], sc: {}, why: {} };
    const cell = (ref, v, s) => put(g, ref, { v, s });
    cell('D1', 'Assessment Details', S.band);
    cell('D2', 'Platoon', S.label); cell('E2', plt, S.val); cell('F2', 'INTENTIONALLY LEFT BLANK', S.mid);
    cell('D3', 'Squad', S.label); cell('E3', sqd, S.val);
    cell('D4', 'Select Your Last Name', S.label);
    if (pick) put(g, 'E4', { f: 'IF(' + ME + '="","",' + ME + ')', v: '', s: S.val }); // the name picked on the dashboard
    else cell('E4', rater ? fullOf(rater) : '', S.val);
    cell('D5', 'Selected Peer to Rate', S.label); cell('E5', fullOf(o), S.val);
    cell('D6', 'Peer Questions', S.band);
    cell('D7', QUESTIONS[0].text, S.ask); cell('E7', m.yes[0] || '', S.mid);
    cell('F7', QUESTIONS[1].text, S.ask); cell('G7', m.yes[1] || '', S.mid);
    cell('H7', 'Additional Comments', S.ask); cell('H8', m.more || '', S.box);
    cell('D9', 'Explain why you would / would not trust this OC in the comment box below.', S.ask);
    cell('F9', 'Would you follow this OC into combat? Check Yes/No on the right then Explain your answer below.', S.ask); cell('G9', m.yes[2] || '', S.mid);
    cell('D10', m.trust || '', S.box); cell('F10', m.follow || '', S.box);
    cell('D13', 'How would you rank this OC in your SQUAD [1 = Best]?', S.ask); cell('E13', m.rank || '', S.mid);
    cell('F13', QUESTIONS[3].text, S.ask); cell('G13', m.yes[3] || '', S.mid);
    cell('F14', 'Peer Questions Score Total:', S.label);
    put(g, 'G14', { f: 'SUM(IF(E7="Yes",1,0),IF(G7="Yes",1,0),IF(G9="Yes",1,0),IF(G13="Yes",1,0))-E13', v: calc[i].q, s: S.mid });
    cell('D20', 'Attribute Rating Table', S.band);
    cell('D21', 'INSTRUCTIONS: Select attribute ratings from the dropdown in Column E, Leave your comments regarding the Rated Candidate in Column F. Be True and Honest. To swap an attribute, pick another in Column ' + (pick ? 'C (Swap to)' : 'D') + ', never one already on this tab.', S.note);
    ['Attributes', 'Ratings', 'Comments', 'Below', 'Meets', 'Exceeds', 'Score'].forEach((t, c) => cell(String.fromCharCode(68 + c) + '22', t, S.head));
    if (pick) cell('C22', 'Swap to', S.head); // a pick here wins over D's formula: a dropdown and a formula in one cell is lost in Numbers
    listOf(o).forEach((ai, r0) => {
      const a = ATTRIBUTES[ai];
      const r = 23 + r0;
      const e = 'E' + r;
      const pt = calc[i].pts[r0];
      if (pick) { // 17i: this rater's attribute for this peer, from the name picked on the dashboard
        const k = 'MOD(' + MATCH_ME + '-' + i + ',' + ocs.length + ')';
        const ak = COPRIME[(pick.seed + i * 5) % 8];
        const bk = (pick.seed * 7 + i * 11) % 30;
        cell('C' + r, '', S.val);
        put(g, 'D' + r, { f: 'IF(C' + r + '<>"",C' + r + ',IFERROR(IF(' + k + '=0,"' + (r0 ? '' : 'This is you: leave it blank') + '",INDEX(\'Attributes\'!$A$2:$A$' + (ATTRIBUTES.length + 1) + ',MOD(' + ak + '*MOD(6*(' + k + '-1)+' + r0 + ',30)+' + bk + ',30)+1)),"' + (r0 ? '' : 'Pick your name first') + '"))', v: r0 ? '' : 'Pick your name first', s: S.val });
      } else cell('D' + r, a.name, S.val);
      cell(e, m.sc[ai] ? LEVELS[m.sc[ai] - 1] : '', S.val);
      cell('F' + r, m.why[ai] || '', S.val);
      put(g, 'G' + r, { f: 'IF(OR(' + e + '="' + LEVELS[0] + '",' + e + '="' + LEVELS[1] + '"),1,"")', v: pt && pt <= 2 ? 1 : '', s: S.mid });
      put(g, 'H' + r, { f: 'IF(' + e + '="' + LEVELS[2] + '",1,"")', v: pt === 3 ? 1 : '', s: S.mid });
      put(g, 'I' + r, { f: 'IF(OR(' + e + '="' + LEVELS[3] + '",' + e + '="' + LEVELS[4] + '"),1,"")', v: pt >= 4 ? 1 : '', s: S.mid });
      put(g, 'J' + r, { f: LEVELS.reduceRight((rest, l, li) => 'IF(' + e + '="' + l + '",' + (li + 1) + ',' + rest + ')', '""'), v: pt || '', s: S.mid }); // the original's nested IF: 1 to 5, blank when not rated
    });
    const merge = ['D1:J1', 'F2:J5', 'D6:J6', 'D7:D8', 'E7:E8', 'F7:F8', 'G7:G8', 'H7:J7', 'H8:J14', 'D9:E9', 'D10:E12', 'F10:G12', 'D13:D14', 'E13:E14', 'D20:J20', 'D21:J21'];
    merged(g, merge);
    const heights = { 8: 40.5, 11: 34.5, 12: 28.5, 20: 26 };
    for (let r = 22; r < end; r++) heights[r] = 30;
    sheets.push({
      name: tabs[i], rows: g, header: false, rowHeight: 15, heights, merge, landscape: true,
      widths: [8.86, 8.86, pick ? 22 : 8.86, 25.29, 21.71, 51.14, 8, 6.29, 7.71, 5.71],
      lists: [
        { ref: 'E2', items: Array.from({ length: Math.max(3, sq.plt || 0) }, (x, k) => ord(k + 1) + ' PLT') },
        { ref: 'E3', items: Array.from({ length: Math.max(4, sq.sqd || 0) }, (x, k) => ord(k + 1) + ' SQD') },
        pick ? null : { ref: 'E4', range: "'" + DASH + "'!$B$" + (rater ? '2' : '4:$B$' + last) }, // the rater's own name (the summary's B2), so Excel takes the name filled in; with no rater, the squad's names; picked once on the dashboard (17i)
        { ref: 'E7 G7 G9 G13', items: ['Yes', 'No'] },
        { ref: 'E23:E' + end, items: LEVELS },
        { ref: (pick ? 'C23:C' : 'D23:D') + end, range: "'Attributes'!$A$2:$A$" + (ATTRIBUTES.length + 1), promptTitle: 'Swap an attribute', prompt: 'Pick another from the list if you like, but never one already on this tab: each attribute only once.' } // swap an attribute: pick another name
      ],
      dupes: ['D23:D' + end] // the same attribute twice on a tab turns red
    });
  });
  sheets.push({ name: 'Attributes', rows: [['Attribute', 'Category', 'What it means']].concat(ATTRIBUTES.map((a) => [a.name, a.cat, a.def])), widths: [40, 14, 60] });
  const name = (sq.plt ? sq.plt + 'PLT_' : '') + (sq.sqd ? ord(sq.sqd) + '_' : '') + 'Squad_Peer_Evaluations' + (rater ? '_' + slug(rater.last) : '') + '.xlsx';
  return { name, data: buildXlsx(sheets) };
}

/**
 * Blank paper forms for OCs to fill in if the app has trouble (17i, the user: "…每一个sqd一个表"): one workbook per
 * squad, zipped when there are several. Each OC picks his own name on it once; the attributes then fill in by formula
 * (the same name, the same six; they may differ from the app's). squads = [{ plt, sqd, ocs: [{ id, last, first }] }].
 */
export function blankBooks(name, squads, seed) {
  const note = 'First pick your own name in B2 (Rater Name): every tab then shows the 6 attributes you rate for that squadmate, and your own tab stays blank. On each tab: Yes or No with a reason, your rank, then a rating and a reason for each attribute. To swap an attribute, pick another in its name cell, never one already on that tab (a repeat turns red). Then send this file to your cadre.';
  const books = [];
  squads.forEach((sq) => {
    if (sq.ocs.length < 2) return;
    const b = squadBook(sq, sq.ocs.map((o, i) => Object.assign({ id: i }, o, { id: o.id === undefined ? i : o.id })), null, note, null, { seed: (Number(seed) || 0) + (sq.plt || 0) * 31 + (sq.sqd || 0) * 7 });
    books.push({ name: b.name, data: b.data });
  });
  if (!books.length) return null;
  return books.length === 1 ? books[0] : { name: 'peer-eval-' + slug(name) + '-blank-sheets.zip', data: zip(books) };
}

/**
 * The cadre's roster as a sheet, the console's Download as a sheet without its ID column: edit it
 * in Excel, keep it, or drop it on Who's in it? later; blanks in yellow. ocs = [{ last, first, plt, sqd }].
 */
export function rosterXlsx(title, ocs, now) {
  const pad = (n) => String(n).padStart(2, '0');
  const rows = [['Peer evaluation roster · ' + title], ['Fix names, platoons and squads here if you like, one OC a row. To use it, drop this file on Who\'s in it? in the peer evaluation app.'], ['Last', 'First', 'PLT', 'SQD']];
  const fill = new Set();
  ocs.forEach((o) => {
    if (!o.plt) fill.add(rows.length + ',2');
    if (!o.sqd) fill.add(rows.length + ',3');
    rows.push([o.last, o.first || '', o.plt || '', o.sqd || '']);
  });
  const day = now.getFullYear() + '-' + pad(now.getMonth() + 1) + '-' + pad(now.getDate());
  return { name: 'peer-eval-roster-' + day + '.xlsx', data: buildXlsx([{ name: 'Roster', rows, widths: [22, 18, 8, 8], header: false, bold: [2], fill }]) }; // no title in the file name: a "2nd SQD" in it would read as the squad of every row
}

/* ---------- each OC's feedback ---------- */

/**
 * Feedback for each OC, anonymous (no names; the comments in mixed order): how the squad rated them against the
 * squad average, by category and by attribute (bars in the cells and a chart), their strengths and what to work
 * on, the Yes/No questions, then every comment. One workbook per OC, zipped, for the cadre to give each OC theirs.
 * @returns {{ name: string, data: Uint8Array } | null} null when no OC has a result about them yet
 */
export function feedbackPages(E, now) {
  const pages = [];
  const day = now.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
  const mean = (xs) => { const ok = xs.filter((x) => typeof x === 'number'); return ok.length ? ok.reduce((t, x) => t + x, 0) / ok.length : null; };
  const mixed = (a) => { // the order of the comments says nothing about who wrote them
    for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
    return a;
  };
  E.squads.filter((sq) => sq.ids.length > 1).forEach((sq) => {
    const res = squadResults(E, sq);
    const { rows } = scoreSquad(sq.ids, res);
    const squadCat = CATEGORIES.map((c) => mean(rows.map((r) => r.cat[c])));
    const squadAttr = ATTRIBUTES.map((a, ai) => mean(rows.map((r) => r.attrMean[ai])));
    sq.ids.forEach((id) => {
      const m = memberOf(E, id);
      const r = rows.find((x) => x.id === id);
      const key = String(id);
      const from = res.filter((p) => p && p.r !== id && ((p.sc || {})[key] || (p.ans || {})[key]));
      if (!r || !from.length) return;
      const attrs = ATTRIBUTES.map((a, ai) => ai).filter((ai) => typeof r.attrMean[ai] === 'number')
        .sort((x, y) => r.attrMean[y] - r.attrMean[x] || (r.attrMean[y] - (squadAttr[y] || 0)) - (r.attrMean[x] - (squadAttr[x] || 0)));
      const top = attrs.slice(0, 3);
      const low = attrs.slice(-3).reverse().filter((ai) => top.indexOf(ai) < 0);
      const ask = QUESTIONS.map((q, i) => ({ q, yes: from.filter((p) => String((p.ans || {})[key] || '')[i] === 'Y').length, no: from.filter((p) => String((p.ans || {})[key] || '')[i] === 'N').length }));
      const said = QUESTIONS.map((q, i) => ({ q, items: mixed(from.map((p) => ({ a: String((p.ans || {})[key] || '')[i], t: String(((p.com || {})[key] || {})[q.id] || '').trim() })).filter((x) => x.a === 'Y' || x.a === 'N' || x.t)) }));
      const notes = mixed([].concat(...attrs.map((ai) => from.map((p) => String(((p.why || {})[key] || {})[ai] || '').trim()))).filter(Boolean)); // one or two raters per attribute: with its name or score a reason could point to its writer (17b)
      pages.push({ sq, m, day, from: from.length, r, attrs, top, low, ask, said, notes, squadCat, squadAttr });
    });
  });
  return pages;
}

/** Feedback for each OC as Excel (see feedbackPages): one workbook per OC, zipped, for the cadre to give each OC theirs. */
export function feedbackBooks(E, now) {
  const S = STYLE;
  const books = [];
  const used = {};
  const one = (v) => (v === null ? '' : Math.round(v * 10) / 10);
  const bar = (v) => (v === null ? '' : '█'.repeat(Math.max(1, Math.round(v * 3)))); // 3 blocks a point: a 5 is 15, the column's width
  feedbackPages(E, now).forEach(({ sq, m, day, from: nFrom, r, attrs, top, low, ask, said: saidQ, notes, squadCat, squadAttr }) => {
      const named = (list) => list.map((ai) => ATTRIBUTES[ai].name + ' (' + one(r.attrMean[ai]) + ')').join(' · ');
      const sheet = [];
      const merge = [];
      const line = (cells, mergeTo) => {
        if (mergeTo) merge.push('A' + (sheet.length + 1) + ':' + mergeTo + (sheet.length + 1));
        sheet.push(cells);
      };
      line([{ v: 'Peer feedback · ' + m.last + (m.first ? ', ' + m.first : ''), s: S.title }, { s: S.title }, { s: S.title }, { s: S.title }], 'D');
      line([{ v: E.name + ' · ' + sq.label + ' · ' + day, s: S.val }, { s: S.val }, { s: S.val }, { s: S.val }], 'D');
      line([{ v: 'From ' + nFrom + (nFrom === 1 ? ' squadmate' : ' squadmates') + '. No names are in it, and the comments are in mixed order.', s: S.note }, { s: S.note }, { s: S.note }, { s: S.note }], 'D');
      line([]);
      line([{ v: 'Your strengths', s: S.label }, { v: named(top), s: S.box }, { s: S.box }, { s: S.box }]);
      merge.push('B' + sheet.length + ':D' + sheet.length);
      line([{ v: 'To work on', s: S.label }, { v: named(low), s: S.box }, { s: S.box }, { s: S.box }]);
      merge.push('B' + sheet.length + ':D' + sheet.length);
      line([]);
      line([{ v: 'By category', s: S.band }, { s: S.band }, { s: S.band }, { s: S.band }], 'D');
      line(['Category', 'You', 'Squad', 'You, 0 to 5'].map((t) => ({ v: t, s: S.head })));
      const catAt = sheet.length + 1;
      CATEGORIES.forEach((c, i) => line([{ v: c, s: S.val }, { v: one(r.cat[c]), s: S.mid }, { v: one(squadCat[i]), s: S.mid }, { v: bar(r.cat[c]), s: S.bar }]));
      line([]);
      line([{ v: 'By attribute', s: S.band }, { s: S.band }, { s: S.band }, { s: S.band }], 'D');
      line(['Attribute', 'You', 'Squad', 'You, 0 to 5'].map((t) => ({ v: t, s: S.head })));
      const attrAt = sheet.length + 1;
      attrs.forEach((ai) => line([{ v: ATTRIBUTES[ai].name, s: S.val }, { v: one(r.attrMean[ai]), s: S.mid }, { v: one(squadAttr[ai]), s: S.mid }, { v: bar(r.attrMean[ai]), s: S.bar }]));
      line([]);
      line([{ v: 'Peer questions', s: S.band }, { s: S.band }, { s: S.band }, { s: S.band }], 'D');
      line(['Question', 'Yes', 'No', ''].map((t) => ({ v: t, s: S.head })));
      ask.forEach((x) => line([{ v: QUESTION_LABEL[x.q.id], s: S.val }, { v: x.yes, s: S.mid }, { v: x.no, s: S.mid }, { v: '', s: S.mid }]));
      const ref = (col, at, n) => "'Summary'!$" + col + '$' + at + ':$' + col + '$' + (at + n - 1);
      const catVals = CATEGORIES.map((c) => one(r.cat[c]));
      const charts = [{ title: 'By category: you and your squad', from: [5, 6], to: [13, 22], max: 5, cats: CATEGORIES, catRef: ref('A', catAt, CATEGORIES.length),
        series: [{ name: 'You', ref: ref('B', catAt, CATEGORIES.length), values: catVals, color: '3F5F33', labels: true }, { name: 'Squad', ref: ref('C', catAt, CATEGORIES.length), values: squadCat.map(one), color: 'C2D69B' }] }];
      if (attrs.length) {
        charts.push({ title: 'By attribute, strongest first', from: [5, 23], to: [13, 23 + Math.max(14, Math.ceil(attrs.length * 1.4))], max: 5, cats: attrs.map((ai) => ATTRIBUTES[ai].name), catRef: ref('A', attrAt, attrs.length),
          series: [{ name: 'You', ref: ref('B', attrAt, attrs.length), values: attrs.map((ai) => one(r.attrMean[ai])), color: '3F5F33', labels: true }, { name: 'Squad', ref: ref('C', attrAt, attrs.length), values: attrs.map((ai) => one(squadAttr[ai])), color: 'C2D69B' }] });
      }
      const said = [[{ v: 'What your squadmates said', s: S.title }, { s: S.title }, { s: S.title }], [{ v: 'No names, mixed order. Every Yes, No and score came with a reason.', s: S.note }, { s: S.note }, { s: S.note }], [], ['Question', 'Answer', 'What they said'].map((t) => ({ v: t, s: S.head }))];
      const merge2 = ['A1:C1', 'A2:C2'];
      saidQ.forEach(({ q, items }) => items.forEach((x) => said.push([{ v: QUESTION_LABEL[q.id], s: S.val }, { v: x.a === 'Y' ? 'Yes' : x.a === 'N' ? 'No' : '', s: S.mid }, { v: x.t, s: S.box }]))); // every rater answers every question: the question can stay
      if (notes.length) {
        said.push([]);
        merge2.push('A' + (said.length + 1) + ':C' + (said.length + 1));
        said.push([{ v: 'On your attributes (no names, no scores, mixed order)', s: S.band }, { s: S.band }, { s: S.band }]);
        notes.forEach((t) => {
          merge2.push('A' + (said.length + 1) + ':C' + (said.length + 1));
          said.push([{ v: t, s: S.box }, { s: S.box }, { s: S.box }]);
        });
      }
      const xlsx = buildXlsx([
        { name: 'Summary', rows: sheet, widths: [30, 8, 8, 24, 2], header: false, rowHeight: 18, heights: { 0: 24, 4: 34, 5: 34 }, merge, charts },
        { name: 'Comments', rows: said, widths: [30, 18, 80], header: false, merge: merge2 } // rows grow with long comments
      ]);
      const twin = sq.ids.filter((x) => memberOf(E, x).last.toLowerCase() === m.last.toLowerCase()).length > 1;
      let file = (sq.plt ? sq.plt + 'PLT ' : '') + ord(sq.sqd) + ' SQD/Feedback_' + slug(m.last) + (twin && m.first ? '_' + slug(m.first) : '') + '.xlsx';
      for (let n = 2; used[file]; n++) file = file.replace(/(_\d+)?\.xlsx$/, '_' + n + '.xlsx'); // VAN DYKE and VAN-DYKE: one file name each
      used[file] = true;
      books.push({ name: file, data: xlsx });
  });
  if (!books.length) return null;
  return books.length === 1 ? { name: books[0].name.split('/').pop(), data: books[0].data } : { name: 'peer-eval-' + slug(E.name) + '-feedback.zip', data: zip(books) };
}
