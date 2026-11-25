/**
 * Rosters: read any sheet, then merge several into one.
 *
 * readRoster(text): one sheet (a paste, a .csv, or an .xlsx tab written out as tab-separated
 * lines) into rows of { last, first, plt, plts, sqd, two, ref }. It reads:
 *   a header row (title rows above it are skipped): its Last, First, Name, PLT, SQD and ID
 *     columns, in any order. Several name columns (a room roster: Resident #1 to #3) give
 *     one OC per cell. An unlabeled column of 1st, 2nd, 3rd ... is the PLT column, and
 *     "1st /2nd" there is an unclear PLT with both choices. Every other column, score
 *     columns too, is listed in `columns` and never kept.
 *   "LAST, First" in one cell; "OC" or "OC/C" in front of a name, as a word of its own
 *     (OCHOA stays OCHOA); two words with no comma are one last name (DE LEON), tagged.
 *   placeholders (the whole name is N/A, Unoccupied, Vacant, TBD or a dash): counted as
 *     empty and listed, never an OC;
 *   heading lines ("1st PLT 2nd SQD", or "1st PLT" then "2nd SQD"): the names under them;
 *   names alone, or with no PLT or SQD: rows the admin completes in the console.
 * Nothing is dropped without a line saying so: a line with a name that cannot be read is in
 * `unread`, and blocks Create links until it is fixed.
 *
 * mergeRoster(sources, dec, ans, ids): one roster from every source, the admin's decisions
 * and the OCs' sign-up answers, with the questions to ask and what stops Create links.
 * Only last name, first name, PLT, SQD and the source survive; nothing else is kept.
 */

const ORD = (n) => {
  const s = ['th', 'st', 'nd', 'rd'];
  const v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
};

const MAX_SQUAD = 20;
const PLACEHOLDER = /^(n\/a|unoccupied|vacant|tbd|[-–—]+)$/i;
const OC_CELL = /^oc(\/c)?\.?$/i;
const OC_WORD = /^oc(\/c)?\.?\s+(?=\S)/i;
const HEADER = {
  last: ['last', 'lastname', 'surname', 'familyname', 'lname'],
  first: ['first', 'firstname', 'givenname', 'fname'],
  name: ['name', 'names', 'fullname', 'ocname', 'candidate', 'candidatename', 'candidatenames', 'cadet', 'student', 'resident', 'occupant', 'lastfirst', 'lastnamefirstname', 'namelastfirst'],
  oc: ['oc'], // a Name column, unless the header has one: then a rank column
  plt: ['plt', 'platoon', 'pltno', 'pltnumber', 'platoonno', 'platoonnumber'],
  sqd: ['sqd', 'squad', 'sqdno', 'sqdnumber', 'squadno', 'squadnumber'],
  id: ['id'], // the console's own Download as a sheet: rows go back by it
  score: ['avg', 'average', 'total', 'overall', 'merit', 'orderofmerit', 'oml', 'raters', 'lookcloser'], // and any "...score"
  other: ['rank', 'room', 'roomnumber', 'roomno', 'gender', 'roomgender', 'sex', 'phone', 'phonenumber', 'cell', 'email', 'company', 'co', 'team', 'mi', 'middle', 'middlename', 'middleinitial', 'unit', 'notes', 'note', 'status', 'remarks', 'bunk', 'bed']
};
const CORE = ['last', 'first', 'name', 'oc', 'plt', 'sqd', 'score'];
const SEP = '[\\s,;:·./|–—-]*';
const HEADING = new RegExp('^' + SEP +
  '(?:(\\d+)(?:st|nd|rd|th)?\\s*(?:plt|platoon)|(?:plt|platoon)\\s*#?\\s*(\\d+))?' + SEP +
  '(?:(\\d+)(?:st|nd|rd|th)?\\s*(?:sqd|squad)|(?:sqd|squad)\\s*#?\\s*(\\d+))?' + SEP + '$', 'i');
const ORDINAL = '\\d{1,2}(?:st|nd|rd|th)?(?:\\s*(?:plt|platoon))?';
const PLT_CELL = new RegExp('^' + ORDINAL + '(?:\\s*[/&,]\\s*' + ORDINAL + ')*$', 'i'); // 1st, 2nd PLT, 1st /2nd
const TWO_WORDS = /^\p{L}[\p{L}'’.-]*(\s+\p{L}[\p{L}'’.-]*)+$/u;

const numbers = (text) => (String(text || '').match(/\d+/g) || []).map(Number);

/** The one number in a PLT or SQD cell; a cell with two numbers ("1st /2nd") is not read. */
function unit(text) {
  const all = numbers(text);
  return all.length === 1 ? all[0] : null;
}

const inRange = (n) => n !== null && n >= 1 && n <= MAX_SQUAD;
const clean = (s) => String(s || '').replace(/\s+/g, ' ').trim();
/** What two spellings must share to be one OC: the letters, any case. */
const letters = (s) => String(s || '').toUpperCase().replace(/[^\p{L}]/gu, '');

function splitLine(line) {
  if (line.indexOf('\t') >= 0) return line.split('\t');
  if (line.indexOf(',') >= 0) return line.split(',');
  return line.split(/\s{2,}/);
}

/** The column a header cell names, matched as a whole word ("Last Name", "PLT #", "Resident #2"), or ''. */
function roleOf(cell) {
  const k = cell.toLowerCase().replace(/[^a-z]/g, '');
  if (/scores?$/.test(k)) return 'score'; // "Peer score", "Total attribute score"; not SCORESBY
  return Object.keys(HEADER).find((r) => HEADER[r].indexOf(k) >= 0) || '';
}

/**
 * The header row: the first line (of the first 30) with no numbers and two or more known
 * column names, one of them a name, PLT, SQD or score column, so title rows above it are
 * skipped; else a first line that is one column name alone ("Name"). -1 when there is none.
 */
function findHeader(lines) {
  const roles = (li) => (/\d/.test(lines[li].replace(/#\s*\d+/g, '')) ? [] : splitLine(lines[li]).map((c) => roleOf(c.trim())).filter(Boolean)); // "Resident #2" is a column name
  let firstLi = -1;
  let seen = 0;
  for (let li = 0; li < lines.length && seen < 30; li++) {
    if (!lines[li].trim()) continue;
    seen++;
    if (firstLi < 0) firstLi = li;
    const r = roles(li);
    if (r.length >= 2 && r.some((x) => CORE.indexOf(x) >= 0)) return li;
  }
  if (firstLi < 0) return -1;
  const one = splitLine(lines[firstLi]).map((c) => c.trim()).filter(Boolean);
  return one.length === 1 && CORE.indexOf(roleOf(one[0])) >= 0 && !/\d/.test(one[0].replace(/#\s*\d+/g, '')) ? firstLi : -1;
}

/** Text from a file or a paste as lines; Excel's CSV quotes ("SMITH, John") become tab-separated cells. */
export function csvText(text) {
  const out = [];
  let row = [];
  let cell = '';
  let quoted = false;
  const s = String(text || '');
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (quoted) {
      if (c === '"' && s[i + 1] === '"') {
        cell += '"';
        i++;
      } else if (c === '"') quoted = false;
      else cell += c;
    } else if (c === '"' && !cell) quoted = true;
    else if (c === ',') {
      row.push(cell);
      cell = '';
    } else if (c === '\n' || c === '\r') {
      if (c === '\r' && s[i + 1] === '\n') i++;
      row.push(cell);
      out.push(row.map((x) => x.replace(/\t/g, ' ')).join('\t'));
      row = [];
      cell = '';
    } else cell += c;
  }
  if (cell || row.length) {
    row.push(cell);
    out.push(row.map((x) => x.replace(/\t/g, ' ')).join('\t'));
  }
  return out.join('\n');
}

/**
 * PLT and SQD from a file name: only a number joined to PLT/Platoon or SQD/Squad counts
 * (1PLT_2nd_Squad_Peer_Evaluations). A class number (001-27), a date, a browser's "(1)" or
 * "- Copy" never does.
 */
export function fileUnits(fileName) {
  const base = String(fileName || '').replace(/\.[A-Za-z0-9]{1,5}$/, '');
  const pick = (word) => {
    const a = new RegExp('(?:^|[^A-Za-z0-9])(\\d{1,2})(?:st|nd|rd|th)?[\\s_-]*(?:' + word + ')(?![A-Za-z])', 'i').exec(base);
    const b = new RegExp('(?:^|[^A-Za-z])(?:' + word + ')[\\s_#-]*(\\d{1,2})(?!\\d)', 'i').exec(base);
    const n = a ? Number(a[1]) : b ? Number(b[1]) : null;
    return inRange(n) ? n : null;
  };
  return { plt: pick('plt|platoon'), sqd: pick('sqd|squad') };
}

/**
 * Read one sheet. See the top of this file.
 * @param {string} text
 * @returns {{rows: object[], unread: string[], warn: string[], skipped: string[], columns: string[], empty: number, titles: number}}
 *   rows: { last, first, plt, plts, sqd, two, ref, line }; unread blocks Create links;
 *   warn asks for a look once; skipped and columns are listed, never kept
 */
export function readRoster(text) {
  const rows = [];
  const bad = []; // [line index, text]
  const skipped = [];
  const pairs = [];
  let columns = [];
  let empty = 0;
  let head = null; // the last heading: {plt, sqd}
  const at = (li, t) => 'Line ' + (li + 1) + ': ' + t;
  // copies from Teams, Outlook and web pages carry no-break and zero-width spaces
  const lines = String(text || '').replace(/[​-‍⁠﻿]/g, '').replace(/ /g, ' ').split(/\r?\n/);
  const headLi = findHeader(lines);
  const above = lines.slice(0, Math.max(headLi, 0)).map((l, i) => (l.trim() ? i + 1 : 0)).filter(Boolean);
  if (above.length) skipped.push((above.length > 1 ? 'Lines ' + above[0] + '–' + above[above.length - 1] : 'Line ' + above[0]) + ': above the header');
  let headCells = [];
  let roles = [];
  let c = null; // column of each single field, from the header row
  let ocCols = []; // one OC per cell in each of these
  if (headLi >= 0) {
    skipped.push(at(headLi, 'header row')); // listed, so no line goes unseen
    headCells = splitLine(lines[headLi]).map((x) => x.trim());
    roles = headCells.map(roleOf);
    const hasName = roles.some((r) => r === 'last' || r === 'name');
    roles = roles.map((r) => (r === 'oc' ? (hasName ? 'other' : 'name') : r));
    if (roles.indexOf('plt') < 0) {
      // an unlabeled column of 1st, 2nd, 3rd under the header is the PLT column
      const body = lines.slice(headLi + 1).filter((l) => l.trim()).map((l) => splitLine(l).map((x) => x.trim()));
      const width = Math.max(headCells.length, ...body.map((r) => r.length));
      for (let i = 0; i < width; i++) {
        if (headCells[i]) continue;
        const vals = body.map((r) => r[i] || '').filter(Boolean);
        if (vals.length && vals.every((v) => PLT_CELL.test(v))) {
          roles[i] = 'plt';
          break;
        }
      }
    }
    const idx = (r) => roles.map((x, i) => (x === r ? i : -1)).filter((i) => i >= 0);
    const lasts = idx('last');
    ocCols = lasts.length ? lasts : idx('name');
    c = {};
    ['first', 'plt', 'sqd', 'id'].forEach((r) => {
      const i = roles.indexOf(r);
      if (i >= 0) c[r] = i;
    });
    if (ocCols.length !== 1) delete c.first; // first names go with one name column only
    columns = headCells.filter((x, i) => x && ocCols.indexOf(i) < 0 && [c.first, c.plt, c.sqd, c.id].indexOf(i) < 0);
    if (!ocCols.length && (c.plt !== undefined || c.sqd !== undefined)) {
      const t = at(headLi, 'The header row has no Last or Name column.');
      return { rows: [], unread: [t], warn: [], skipped, columns: [], empty: 0, titles: above.length };
    }
    if (!ocCols.length) c = null; // a score or room column alone: read line by line
  }
  /** One name cell: the OC's last and first name, or '' for a placeholder. */
  const nameOf = (raw, li) => {
    let name = clean(raw).replace(OC_WORD, '');
    let given = '';
    const two = name.indexOf(',') < 0 && TWO_WORDS.test(name);
    const comma = name.indexOf(',');
    if (comma > 0) {
      given = name.slice(comma + 1);
      name = name.slice(0, comma);
    }
    name = clean(name);
    given = clean(given);
    if (PLACEHOLDER.test(name)) {
      empty++;
      return null;
    }
    if (!name) return given ? (bad.push([li, at(li, 'a first name with no last name')]), null) : null;
    if (!/\p{L}/u.test(name)) {
      bad.push([li, at(li, '"' + name + '" is not a name.' + (/\d/.test(name) ? ' This looks like a score table; paste last name, PLT and SQD.' : ' Delete the line or put the name in.'))]);
      return null;
    }
    return { last: name.toUpperCase(), first: given, two };
  };
  /** PLT from its cell: one number, or the choices of an unclear "1st /2nd". */
  const pltOf = (raw, li) => {
    const all = numbers(raw);
    if (!all.length) return { plt: null, plts: null };
    if (!all.every(inRange)) {
      bad.push([li, at(li, 'PLT and SQD must be 1 to 20')]);
      return { plt: null, plts: null, bad: true };
    }
    return all.length === 1 ? { plt: all[0], plts: null } : { plt: null, plts: all.filter((n, i) => all.indexOf(n) === i) };
  };
  lines.forEach((line, li) => {
    if (li <= headLi || !line.trim()) return;
    // under a one-column header ("Name"), a line with no tab is one cell: "MORGAN, Alex" is one OC
    const cells = (c && headCells.length === 1 && line.indexOf('\t') < 0 ? [line] : splitLine(line)).map((x) => x.trim());
    const h = HEADING.exec(line);
    if (h && (h[1] || h[2] || h[3] || h[4])) {
      const plt = h[1] || h[2];
      head = { plt: plt ? Number(plt) : head ? head.plt : null, sqd: h[3] || h[4] ? Number(h[3] || h[4]) : null };
      if ((plt && !inRange(head.plt)) || (head.sqd !== null && !inRange(head.sqd))) {
        bad.push([li, at(li, 'PLT and SQD must be 1 to 20')]);
        head = null; // the names under it wait for their PLT and SQD
      }
      return;
    }
    if (c) {
      const loose = cells.find((x, i) => x && !headCells[i] && roles[i] !== 'plt' && /\p{L}/u.test(x));
      if (loose) return bad.push([li, at(li, 'a name in a column with no header. Put each OC on a line of their own, or give the column a header.')]);
      const get = (r) => (c[r] === undefined ? '' : cells[c[r]] || '');
      const p = pltOf(get('plt'), li);
      const sqdRaw = get('sqd');
      const s = sqdRaw ? unit(sqdRaw) : null;
      if (sqdRaw && !inRange(s)) bad.push([li, at(li, numbers(sqdRaw).length > 1 ? 'two numbers in the SQD cell' : 'PLT and SQD must be 1 to 20')]);
      if (p.bad || (sqdRaw && !inRange(s))) return undefined;
      let any = false;
      ocCols.forEach((ci) => {
        const n = nameOf(cells[ci] || '', li);
        if (!n) return;
        any = true;
        rows.push({ last: n.last, first: c.first !== undefined ? clean(get('first')) || n.first : n.first, plt: p.plt || (head && !p.plts ? head.plt : null), plts: p.plts, sqd: s || (head ? head.sqd : null), two: n.two, ref: clean(get('id')), line: li });
      });
      if (!any && !ocCols.some((ci) => PLACEHOLDER.test(clean(cells[ci] || '')))) skipped.push(at(li, 'no name'));
      return undefined;
    }
    while (cells.length && !cells[cells.length - 1]) cells.pop();
    while (cells.length && !cells[0]) cells.shift();
    if (cells.length > 1 && OC_CELL.test(cells[0])) cells.shift();
    const own = !(cells.length <= 2 && !/\d/.test(cells.join('')));
    // a comma in either cell, or a placeholder in the second (a room roster: "DAVIS, Unoccupied"), is two names on one line
    if (!own && cells.length === 2 && (cells.join('').indexOf(',') >= 0 || PLACEHOLDER.test(clean(cells[1])))) return bad.push([li, at(li, 'two names on one line? Put each OC on a line of their own.')]);
    if (own && cells.length < 3) {
      if (PLACEHOLDER.test(clean(cells[0]).replace(OC_WORD, ''))) return skipped.push(at(li, '"' + clean(cells[0]) + '"'));
      return bad.push([li, at(li, 'needs last name, PLT and SQD')]);
    }
    const n = nameOf(cells[0] || '', li);
    if (!n) {
      if (PLACEHOLDER.test(clean(cells[0]))) skipped.push(at(li, '"' + clean(cells[0]) + '"'));
      return undefined;
    }
    if (!own) {
      const given = clean(cells[1] || '') || n.first;
      // LAST<TAB>First, or two OCs? "MORGAN<TAB>Alex" is the usual layout; the same letter case in both cells ("Morgan<TAB>Vance") is worth a look
      if (cells.length === 2 && line.indexOf(',') < 0 && (cells[0] === cells[0].toUpperCase()) === (cells[1] === cells[1].toUpperCase())) pairs.push(n.last + ', ' + given);
      return rows.push({ last: n.last, first: given, plt: head ? head.plt : null, plts: null, sqd: head ? head.sqd : null, two: n.two, ref: '', line: li });
    }
    const p = unit(cells[cells.length - 2]);
    const s = unit(cells[cells.length - 1]);
    if (p === null || s === null) return bad.push([li, at(li, 'cannot read the name, PLT or SQD')]);
    if (!inRange(p) || !inRange(s)) return bad.push([li, at(li, 'PLT and SQD must be 1 to 20 (this line has ' + p + ' and ' + s + ')')]);
    return rows.push({ last: n.last, first: cells.length >= 4 ? clean(cells[1]) : n.first, plt: p, plts: null, sqd: s, two: n.two, ref: '', line: li });
  });
  // many lines with the same message read as one entry ("Lines 5, 6, 7 and 60 more: ...")
  const key = (b) => b[1].replace(/^Line \d+: /, '');
  const lineSets = {};
  bad.sort((a, b) => a[0] - b[0]).forEach((b) => (lineSets[key(b)] = lineSets[key(b)] || []).push(b[0] + 1));
  const unread = [];
  bad.forEach((b) => {
    const ls = lineSets[key(b)];
    if (ls.length <= 3) unread.push(b[1]);
    else if (ls[0] === b[0] + 1) unread.push('Lines ' + ls.slice(0, 3).join(', ') + ' and ' + (ls.length - 3) + ' more: ' + key(b));
  });
  const warn = pairs.length ? ['The second column is read as first names: ' + pairs.slice(0, 3).join(' · ') + (pairs.length > 3 ? ' · …' : '') + '. If it holds other OCs, put each OC on a line of their own.'] : [];
  return { rows, unread, warn, skipped, columns, empty, titles: above.length };
}

/** Letters apart, for near misses (MORGEN and MORGAN). */
function distance(a, b) {
  const d = Array.from({ length: b.length + 1 }, (x, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let prev = d[0];
    d[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const t = d[j];
      d[j] = Math.min(d[j] + 1, d[j - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = t;
    }
  }
  return d[b.length];
}

/** Two spellings that may be one OC: 1 letter apart (2 for 8 or more letters), or one extra word (DE LEON, LEON). */
export function nearMiss(x, y) {
  const a = letters(x);
  const b = letters(y);
  if (!a || !b || a === b) return false;
  const wa = clean(x).toUpperCase().split(' ').map(letters).filter(Boolean);
  const wb = clean(y).toUpperCase().split(' ').map(letters).filter(Boolean);
  const [few, many] = wa.length < wb.length ? [wa, wb] : [wb, wa];
  if (many.length === few.length + 1 && few.every((w) => many.indexOf(w) >= 0)) return true;
  return distance(a, b) <= (Math.max(a.length, b.length) >= 8 ? 2 : 1);
}

const fullOf = (o) => o.last + (o.first ? ', ' + o.first : '');

/** First names that can be one OC: one is missing, or the words fit, a lone initial to its word (Jory, Jory Q, Jory Quinn). Two names that differ (Sam, Steve; Mary Ann, Mary Beth) are asked instead (Same person?). */
export const firstOk = (a, b) => {
  const x = String(a || '').trim().split(/\s+/).map(letters).filter(Boolean);
  const y = String(b || '').trim().split(/\s+/).map(letters).filter(Boolean);
  const fits = (p, q) => (p.length < 2 || q.length < 2 ? p[0] === q[0] : p === q);
  return !x.length || !y.length || (fits(x[0], y[0]) && (x.length < 2 || y.length < 2 || fits(x[1], y[1])));
};

/** Squad label and member names: same last name in one squad gets the first initial, or the first name. */
function nameSquad(sq, block) {
  sq.ocs.sort((a, b) => a.last.localeCompare(b.last) || a.first.localeCompare(b.first));
  sq.ocs.forEach((m) => {
    const same = sq.ocs.filter((o) => letters(o.last) === letters(m.last));
    let name = m.last;
    if (same.length > 1) {
      const initial = m.first ? m.first[0].toUpperCase() + '.' : '';
      const sameInitial = same.filter((o) => (o.first ? o.first[0].toUpperCase() : '') === initial[0]);
      name = m.last + (sameInitial.length > 1 && m.first ? ', ' + m.first : initial ? ' ' + initial : '');
    }
    m.name = name;
  });
  const told = {};
  sq.ocs.forEach((m) => {
    const same = sq.ocs.filter((o) => letters(o.last) === letters(m.last));
    if (same.length > 1 && same.some((o) => !o.first) && !told[letters(m.last)]) {
      told[letters(m.last)] = true;
      block(sq.label + ': two ' + m.last + ' need first names to tell apart. Tap a name to add one.');
    }
  });
}

/**
 * One roster from every source. See the top of this file.
 * @param {object[]} sources [{ id, name, rows, plt, sqd, unread, warn }] in the order they came in;
 *   plt and sqd fill rows that have none (from the file name, or set on its chip)
 * @param {object} [dec] the admin's decisions, each keyed by a row key ("s1:12"):
 *   sq (squad), plt, name ({last, first, saw}; first only where it was typed, saw: the first name shown before the first fix), same (per
 *   question: 'same' | 'two'), which (per question: the chosen OC's key, or 'none': someone else), not (per question: the
 *   keys of the OCs offered when Someone else was said: never those), add ('add' | 'ignore'), src (per
 *   source id: {plt, sqd}), five (per PLT: a 5th squad button), out (true: the OC is left out,
 *   a recycle; it still matches, so a later file's row of the same OC stays out, and questions
 *   about it say so), drop (true: a copy of one OC listed twice; the row is not read at all,
 *   and is in drops to put back). name and plt count before matching, and the file's own
 *   spelling, first name and platoon still match too, so a later file finds the fixed OC
 * @param {object} [ans] the OCs' sign-up answers by OC id: { sqd, plt, first, twice: [sqd, sqd] },
 *   plus ans.extra: names typed in by OCs not on the list [{ key, last, first, plt, sqd }]
 * @param {object} [ids] OC ids fixed at Ask OCs, by row key
 * @param {object} [asked] the first name each Ask OCs link listed, by OC id: an OC split since keeps its id on the half with that name
 * @returns {{ocs: object[], outs: object[], drops: object[], squads: object[], questions: object[], block: string[], hold: string[], list: string[], need: number}}
 */
export function mergeRoster(sources, dec, ans, ids, asked) {
  dec = dec || {};
  ans = ans || {};
  ids = ids || {};
  asked = asked || {};
  const of = (k) => dec[k] || {};
  const outTag = (o) => ((o.keys || [o.key]).some((k) => of('out')[k]) ? ' (left out)' : ''); // an OC, or one row
  const ocs = [];
  const drops = [];
  const questions = [];
  const block = [];
  const hold = [];
  const list = [];
  const order = sources.map((s, i) => ({ s, i })).sort((a, b) => b.s.rows.length - a.s.rows.length || a.i - b.i).map((x) => x.s);
  const pltOk = (o, plt, plts) => {
    const mine = o.plt ? [o.plt] : o.plts || [];
    const theirs = plt ? [plt] : plts || [];
    return !mine.length || !theirs.length || mine.some((p) => theirs.indexOf(p) >= 0);
  };
  const initialOk = (a, b) => !letters(a) || !letters(b) || letters(a)[0] === letters(b)[0];
  // a fixed OC goes by its fix and by the file's own spelling, first name and platoon: a later file of that OC finds it either way
  // (a first name added where the file had none matches only the new one: the other OC of that last name never takes the fix)
  const same = (a, b) => !!a && !!b && letters(a) === letters(b);
  const any = (as, bs, fn) => as.some((a) => bs.some((b) => fn(a, b)));
  const lastsOf = (x) => [x.last, x.orig || x.renamed].filter(Boolean); // a row's file spelling is orig, an OC's is renamed
  const firstsOf = (x) => [x.long || x.first || ''].concat(x.origFirst ? [x.origFirst] : []); // an OC's fullest name from the files (long, where the cadre's is shorter), and the file's own before a fix
  const full = (s) => letters(String(s || '').trim().split(/\s+/)[0]).length > 1; // a first name, not blank or an initial
  const lastOk = (o, r) => any(lastsOf(o), lastsOf(r), same);
  const lastNear = (o, r) => any(lastsOf(o), lastsOf(r), nearMiss);
  const pltFits = (o, r) => pltOk(o, r.plt, r.plts) || (!!r.origPlt && pltOk(o, r.origPlt, null)) || (!!o.origPlt && (pltOk({ plt: o.origPlt }, r.plt, r.plts) || o.origPlt === r.origPlt));
  const cadreSq = (o) => o.keys.map((k) => of('sq')[k]).find(Boolean);
  const sqOf = (o) => cadreSq(o) || o.sqd; // where it is: the cadre's own squad counts
  const sqOk = (o, r) => { // a row's squad agrees with one the files or the cadre gave this OC (or either has none)
    const sqs = [o.sqd, cadreSq(o)].concat(o.opts || []).filter(Boolean);
    return !r.sqd || !sqs.length || sqs.indexOf(r.sqd) >= 0;
  };
  const askName = (o) => fullOf(o) + (o.long && letters(o.long) !== letters(o.first) ? ' (' + o.long + ')' : ''); // in a question: the cadre's name and the fuller one a file gave
  /** Someone else, said to a Which: never one of the OCs offered then (by their keys); an OC never offered can still be this one. */
  const offered = (qid, o) => of('which')[qid] === 'none' && (of('not')[qid] || []).some((k) => o.keys.indexOf(k) >= 0);
  const whichQ = (r, src, opts, qid) => ({ kind: 'which', id: qid, name: fullOf(r) + outTag(r), src: src.name, plt: r.plt, sqd: r.sqd, options: opts.map((o) => ({ key: o.key, name: askName(o) + (sqOf(o) && opts.some((x) => x !== o && askName(x) === askName(o)) ? ' · ' + ORD(sqOf(o)) + ' SQD' : '') + outTag(o), plt: o.plt })) }); // two options of one name: their squads tell them apart
  /** PLTs a source holds whole: rows there with no squad, or with two or more squads (a class or platoon roster, not one squad's list). */
  const covered = (src) => {
    const by = {};
    src.rows.forEach((r) => {
      const plt = r.plt || src.chipPlt;
      if (!plt) return;
      by[plt] = by[plt] || new Set();
      by[plt].add(r.sqd || src.chipSqd || 0);
    });
    return Object.keys(by).filter((p) => by[p].has(0) || by[p].size > 1).map(Number);
  };
  const merge = (o, src, key, r) => {
    o.keys.push(key);
    o.srcs.push(src.id);
    if (r.was) o.was = (o.was || []).concat(r.was);
    if (o.from.indexOf(src.name) < 0) o.from.push(src.name);
    if (r.orig !== undefined) { // a row the cadre fixed: the OC goes by the fix from here on, and its file spelling still finds it
      o.renamed = o.renamed || r.orig;
      o.last = r.last;
      o.origFirst = o.origFirst || r.origFirst;
    }
    if (r.firstFixed) { // the cadre's first name (a fix of the last name alone keeps the OC's own)
      if (!o.firstFixed) { // a bigger file's row of this OC holds a newer one: a fix goes on the OC's first row
        if (firstOk(o.first, r.first) && letters(o.first).length > letters(r.first || '').length) o.long = o.first; // a fuller name the files gave still finds later rows (Mary Ann, fixed to Mary)
        else if (o.first && letters(o.first) !== letters(r.first)) o.was = (o.was || []).concat([o.first]); // the name it replaces still finds the OC for a question (Tony, fixed to Anthony)
        o.first = r.first || '';
      }
      o.firstFixed = true; // no file changes it
    } else if (r.first) { // the fuller name (M, Mary, Mary Ann): another Mary B. then no longer fits
      const fuller = (a) => firstOk(a, r.first) && letters(r.first).length > letters(a || '').length;
      if (fuller(o.long || o.first)) {
        o.was = (o.was || []).concat(o.long || o.first ? [o.long || o.first] : []); // the shorter one still finds the OC for a question (M, Mike: a Michael is asked)
        if (o.firstFixed) o.long = r.first; // the cadre's name stays on the list; later rows match the fuller one (Mary, then Mary Ann: a Mary Beth is asked)
        else o.first = r.first;
      }
      if (o.origFirst && fuller(o.origFirst)) { // the file's own before a fix, fuller now too
        o.was = (o.was || []).concat([o.origFirst]);
        o.origFirst = r.first;
      }
    }
    if (r.origPlt) { // a transfer: the new platoon from here on
      o.origPlt = o.origPlt || r.origPlt;
      o.plt = r.plt;
      o.plts = null;
    } else if (!o.plt && r.plt) o.plt = r.plt;
    if (r.sqd) {
      if (!o.sqd && !o.opts) o.sqd = r.sqd;
      else if (o.sqd && o.sqd !== r.sqd) {
        o.opts = [o.sqd, r.sqd];
        o.sqd = null; // two files disagree: both values offered as buttons
      } else if (o.opts && o.opts.indexOf(r.sqd) < 0) o.opts.push(r.sqd);
    }
    if (letters(o.last) !== letters(r.last) && o.also.indexOf(r.last) < 0) o.also.push(r.last);
  };
  const add = (src, key, r) => {
    const o = { key, keys: [key], srcs: [src.id], from: [src.name], last: r.last, first: r.first || '', plt: r.plt || null, plts: r.plt ? null : r.plts || null, sqd: r.sqd || null, opts: null, two: !!r.two, also: [], renamed: r.orig, origPlt: r.origPlt, origFirst: r.origFirst, was: r.was, firstFixed: !!r.firstFixed, long: r.firstFixed && r.origFirst && firstOk(r.first, r.origFirst) && letters(r.origFirst).length > letters(r.first).length ? r.origFirst : undefined }; // Mary Ann shortened to Mary: a Mary Beth is asked
    ocs.push(o);
    return o;
  };
  /** A copy of one OC listed twice, taken off (not read): kept to put back. */
  const dropped = (src, r) => drops.push({ key: r.key, keys: [r.key], srcs: [src.id], from: [src.name], last: r.last, first: r.first || '', plt: r.plt || null, plts: r.plt ? null : r.plts || null });
  const seenCover = [];
  order.forEach((src) => {
    const chip = Object.assign({ plt: src.plt || null, sqd: src.sqd || null }, of('src')[src.id] || {});
    src.chipPlt = chip.plt;
    src.chipSqd = chip.sqd;
    const rows = src.rows.map((r, ri) => {
      // a fix (name, platoon) counts before matching: a later file then matches the fixed OC, never someone the file spells alike
      const key = src.id + ':' + ri;
      const nm = of('name')[key];
      const filePlt = r.plt || (r.plts ? null : chip.plt);
      const plt = of('plt')[key] || filePlt;
      const row = Object.assign({}, r, { plt, plts: plt ? null : r.plts, sqd: r.sqd || (filePlt === chip.plt || !r.plt ? chip.sqd : null), key });
      if (filePlt && plt !== filePlt) row.origPlt = filePlt; // a transfer: a later file still lists the old platoon
      if (nm) {
        const first = nm.first === undefined ? r.first || '' : clean(nm.first);
        // the cadre's first name: typed (a case-only fix too), or one the file had, cleared
        Object.assign(row, { orig: r.last, last: clean(nm.last).toUpperCase() || r.last, first, firstFixed: nm.first !== undefined && (nm.saw !== undefined || !!first || first !== clean(r.first)) });
        if (r.first && letters(r.first) !== letters(first)) row.origFirst = r.first; // a first name changed: the file's own still finds the OC
        if (nm.saw && letters(nm.saw) !== letters(first) && letters(nm.saw) !== letters(r.first)) row.was = [nm.saw]; // the one the cadre saw (another file's): a later row with it is asked about this OC, never given to another quietly
      }
      return row;
    });
    rows.filter((r) => of('drop')[r.key]).forEach((r) => dropped(src, r));
    const later = [];
    rows.filter((r) => !of('drop')[r.key]).forEach((r) => { // one OC listed twice: the copy dropped is not read at all
      const qid = 'w:' + r.key;
      const fits = (o) => o.srcs.indexOf(src.id) < 0 && lastOk(o, r) && pltFits(o, r) && !offered(qid, o);
      const said = of('which')[qid]; // the cadre's Which answer stands, whatever names were fixed since
      const pick = said && said !== 'none' && ocs.find((o) => o.srcs.indexOf(src.id) < 0 && o.keys.indexOf(said) >= 0);
      if (pick) return merge(pick, src, r.key, r);
      if (said === 'none' && !of('not')[qid]) return later.push(r); // Someone else, said before the OCs offered were kept: on to the questions below
      let cands = ocs.filter((o) => fits(o) && firstOk(o.long || o.first, r.first));
      // a first name the cadre fixed (Tony, Anthony) still finds its OC by the file's own, but only when no other OC fits and the squads agree (or one has none): else it's asked;
      // a blank or an initial the file had never merges alone once the cadre typed a full first name (T, typed Tony: a Tamsin is asked)
      if (!cands.length) cands = ocs.filter((o) => fits(o) && any(firstsOf(o).filter((f) => full(f) || !full(o.long || o.first)), firstsOf(r), firstOk) && sqOk(o, r));
      // only a blank or an initial fits: an OC that fits by a name it had (the file's own before a fix, M before Mike) could be this one too, so it's asked
      else if (!cands.some((o) => full(o.long || o.first) && full(r.first))) cands = cands.concat(ocs.filter((o) => cands.indexOf(o) < 0 && fits(o) && any(firstsOf(o).concat(o.was || []), firstsOf(r), firstOk)));
      if (cands.length > 1 && r.sqd) { // the squad picks: a squad list's Tony is that squad's ZRDORSET; one with no squad yet may be him too; none agree: all are asked
        const agree = cands.filter((o) => sqOk(o, r));
        if (agree.length) cands = agree;
      }
      const one = cands.length === 1 ? cands[0] : null;
      // two names in this file fit its blank or initial (Tamsin, Tony for T): which one it is is asked, never the first quietly
      if (one && !full(one.long || one.first) && full(r.first) && rows.some((r2) => r2 !== r && !of('drop')[r2.key] && full(r2.first) && letters(r2.first) !== letters(r.first) && lastOk(one, r2) && pltFits(one, r2) && firstOk(one.long || one.first, r2.first) && sqOk(one, r2))) return questions.push(whichQ(r, src, [one], qid));
      if (one && one.firstFixed && !one.long && r.first && firstOk(one.first, r.first) && letters(r.first).length > letters(one.first).length) { // the cadre's short name (Mary, M, none) meets a fuller one: asked once (Mary Ann, or another OC's Mary Beth)
        const sq = 'n:' + one.key + '|' + r.key;
        if (of('same')[sq] === 'same') return merge(one, src, r.key, r);
        if (!of('same')[sq]) return questions.push({ kind: 'same', id: sq, a: { name: askName(one) + outTag(one), src: one.from[0] }, b: { name: fullOf(r) + outTag(r), src: src.name }, plt: r.plt });
        return later.push(r); // Two people: on to Add below
      }
      if (one) return merge(one, src, r.key, r);
      if (cands.length > 1) return questions.push(whichQ(r, src, cands, qid));
      return later.push(r);
    });
    const cover = seenCover.slice();
    later.forEach((r) => {
      const qid = 'w:' + r.key;
      const said = of('which')[qid];
      const old = said === 'none' && !of('not')[qid]; // Someone else, said before the OCs offered were kept: not asked again
      // a spelling one letter apart, or the same last name with another first name of one initial (Mike, Michael), by the fix, the file's own or a name it had (Tony, fixed to Anthony): asked (never about an OC Someone else was said to)
      const near = ocs.filter((o) => o.srcs.indexOf(src.id) < 0 && !offered(qid, o) && pltFits(o, r) && (lastOk(o, r) || lastNear(o, r)) && (any(firstsOf(o), firstsOf(r), initialOk) || any(o.was || [], firstsOf(r), (a, b) => firstOk(a, b) || nearMiss(a, b))))
        .sort((a, b) => distance(letters(a.last), letters(r.last)) - distance(letters(b.last), letters(r.last)));
      if (near.length > 1) { // two who could be this OC (Sam and Steve, for Stephen): which one, or someone else
        const pick = near.find((o) => o.keys.indexOf(said) >= 0);
        if (pick) return merge(pick, src, r.key, r);
        if (!old) return questions.push(whichQ(r, src, near, qid));
      } else if (near.length && !old) {
        const qid = 'n:' + near[0].key + '|' + r.key;
        const said = of('same')[qid];
        if (said === 'same') return merge(near[0], src, r.key, r);
        if (!said) return questions.push({ kind: 'same', id: qid, a: { name: askName(near[0]) + outTag(near[0]), src: near[0].from[0] }, b: { name: fullOf(r) + outTag(r), src: src.name }, plt: r.plt }); // the row's own platoon: a row with none is asked whatever platoon the evaluation runs
      }
      const covers = !outTag(r) && cover.some((c) => (r.plt ? c.plts.indexOf(r.plt) >= 0 : c.plts.length > 1)); // one the cadre left out isn't asked to be added
      if (covers) {
        const said = of('add')[r.key];
        if (said === 'add') return add(src, r.key, r);
        if (said === 'ignore') return list.push(src.name + ': ' + r.last + ' ignored (not in the bigger roster)');
        return questions.push({ kind: 'add', id: r.key, name: r.last + (r.first ? ', ' + r.first : ''), src: src.name, plt: r.plt, sqd: r.sqd, hint: '' });
      }
      return add(src, r.key, r);
    });
    seenCover.push({ src, plts: covered(src) });
    (src.unread || []).forEach((t) => block.push(src.name + ': ' + t));
    (src.warn || []).forEach((t) => hold.push(src.name + ': ' + t));
  });
  (ans.extra || []).forEach((x) => {
    if (of('drop')[x.key]) return dropped({ id: 'j', name: 'sign-up' }, x); // an added copy of one OC listed twice can be dropped too
    const said = of('add')[x.key];
    const like = ocs.find((o) => pltOk(o, x.plt, null) && (letters(o.last) === letters(x.last) || nearMiss(o.last, x.last)));
    if (said === 'add') return add({ id: 'j', name: 'sign-up' }, x.key, x);
    if (said === 'ignore') return undefined;
    return questions.push({ kind: 'add', id: x.key, name: x.last + (x.first ? ', ' + x.first : ''), src: 'the sign-up', plt: x.plt, sqd: x.sqd, hint: like ? 'Looks like ' + like.last + (like.first ? ', ' + like.first : '') + ' in ' + ORD(like.plt || x.plt) + ' PLT.' : '' });
  });

  // the admin's decisions, then the OCs' answers
  const firstOf = (o, field) => {
    for (const k of o.keys) if (of(field)[k] !== undefined && of(field)[k] !== null && of(field)[k] !== '') return of(field)[k];
    return undefined;
  };
  const held = {}; // ids fixed at Ask OCs already given: an OC split off later (Someone else, Two people) gets a new one: the id
  // (and the phone's pick) stays with the half that has the first name the link listed, else with the half the admin didn't rename
  const link = (o) => o.keys.some((k) => ids[k] !== undefined && !!asked[ids[k]] && letters(asked[ids[k]]) === letters(o.first));
  const rank = (o) => (link(o) ? 0 : 2) + (firstOf(o, 'name') ? 1 : 0);
  ocs.slice().sort((a, b) => rank(a) - rank(b)).forEach((o) => {
    o.id = o.keys.map((k) => ids[k]).find((id) => id !== undefined && !held[id]);
    if (o.id !== undefined) held[o.id] = true;
  });
  ocs.forEach((o) => {
    const nm = firstOf(o, 'name');
    if (nm) {
      const read = o.renamed !== undefined; // a file row's fix, read before matching: its first name counted only where the fix changed it
      o.renamed = o.renamed || o.last;
      o.last = clean(nm.last).toUpperCase() || o.last;
      if (!read) o.first = nm.first === undefined ? o.first : clean(nm.first);
    }
    const a = o.id !== undefined ? ans[o.id] : null;
    const myPlt = firstOf(o, 'plt');
    if (myPlt) o.plt = myPlt;
    else if (!o.plt && a && a.plt && (!o.plts || o.plts.indexOf(a.plt) >= 0)) o.plt = a.plt;
    if (o.plt) o.plts = null;
    if (a && a.first && !o.first) o.first = a.first;
    const mine = firstOf(o, 'sq');
    const file = o.sqd;
    o.tags = [];
    if (o.two) o.tags.push('two words, read as one last name');
    if (o.also.length) o.tags.push('also written ' + o.also.join(', '));
    if (o.renamed) o.tags.push('renamed from ' + o.renamed);
    if (mine) {
      o.sqd = mine;
      o.by = 'you';
      if (a && a.sqd && a.sqd !== mine) o.tags.push('you set ' + ORD(mine) + ', the OC picked ' + ORD(a.sqd));
    } else if (a && a.twice) {
      o.sqd = null;
      o.by = '';
      o.twice = a.twice;
      o.tags.push('picked twice from two phones: ' + a.twice.map(ORD).join(' and '));
    } else if (file) {
      o.by = 'file';
      if (a && a.sqd && a.sqd !== file) o.tags.push('the file says ' + ORD(file) + ', the OC picked ' + ORD(a.sqd));
    } else if (a && a.sqd) {
      o.sqd = a.sqd;
      o.by = 'oc';
    } else {
      o.by = '';
      if (o.opts) o.tags.push('the files say ' + o.opts.map(ORD).join(' or '));
    }
    if (o.by === 'oc') o.tags.unshift('picked by the OC');
    else if (o.by === 'you') o.tags.unshift('set by you');
  });

  // the OCs left out (recycles) leave the list here, before every check below: those run on who is in
  const outs = [];
  for (let i = ocs.length - 1; i >= 0; i--) if (outTag(ocs[i])) outs.unshift(ocs.splice(i, 1)[0]);

  // the same OC twice in one file (after the admin's fixes and the first names phones sent)
  const listed = {};
  ocs.forEach((o) => {
    const k = o.srcs[0] + '|' + letters(o.last) + '|' + letters(o.first) + '|' + (o.plt || (o.plts || []).join('/')) + '|' + (o.sqd || '');
    const src = sources.find((x) => x.id === o.srcs[0]);
    if (listed[k] === 1) block.push((src ? src.name : 'The sign-up') + ': ' + fullOf(o) + ' is listed twice: delete one, or tap a name to add a middle initial if they are two people.');
    listed[k] = (listed[k] || 0) + 1;
  });
  const squads = [];
  const byKey = {};
  ocs.filter((o) => o.plt && o.sqd).forEach((o) => {
    const key = o.plt + '-' + o.sqd;
    if (!byKey[key]) {
      byKey[key] = { key, plt: o.plt, sqd: o.sqd, label: ORD(o.plt) + ' PLT · ' + ORD(o.sqd) + ' SQD', ocs: [] };
      squads.push(byKey[key]);
    }
    byKey[key].ocs.push(o);
  });
  squads.sort((a, b) => a.plt - b.plt || a.sqd - b.sqd);
  squads.forEach((sq) => {
    nameSquad(sq, (t) => block.push(t));
    if (sq.ocs.length > MAX_SQUAD) block.push(sq.label + ': ' + sq.ocs.length + ' OCs is too many for one squad (20 at most).');
    else if (sq.ocs.length > 16) list.push(sq.label + ': ' + sq.ocs.length + ' OCs is a lot for one squad');
    if (sq.ocs.length < 2) hold.push(sq.label + ': only one OC, nobody to rate');
  });
  ocs.filter((o) => !(o.plt && o.sqd)).forEach((o) => {
    o.name = o.last;
  });
  const noPlt = ocs.filter((o) => !o.plt).length;
  const noSqd = ocs.filter((o) => o.plt && !o.sqd).length;
  const twice = ocs.filter((o) => o.twice).map((o) => o.last);
  if (twice.length) block.unshift('Picked into two squads from two phones: ' + twice.join(', ') + '. Tap the right squad.');
  if (noSqd) block.unshift(noSqd + (noSqd === 1 ? ' OC has' : ' OCs have') + ' no squad yet.');
  if (noPlt) block.unshift(noPlt + (noPlt === 1 ? ' OC has' : ' OCs have') + ' no PLT yet.');
  const open = questions.filter((q) => q.kind !== 'add').length;
  if (open) block.unshift(open + (open === 1 ? ' question' : ' questions') + ' to answer first.');
  const waiting = questions.filter((q) => q.kind === 'add');
  if (waiting.length) hold.unshift(waiting.length + (waiting.length === 1 ? ' name waits' : ' names wait') + ' for Add or Ignore: ' + waiting.map((q) => q.name).join(', ') + '.');
  return { ocs, outs, drops, squads, questions, block, hold, list, need: noPlt + noSqd };
}

/**
 * A name typed in by hand, cleaned like a file's: no zero-width or no-break spaces, one space
 * between words, no OC in front, "LAST, First" in the last name split, the last name in capitals,
 * 30 characters each. null when it is no name: empty, a placeholder (N/A, TBD, -) or no letters.
 */
export function typedName(last, first) {
  const tidy = (s) => clean(String(s || '').replace(/[​-‍⁠﻿]/g, '').replace(/ /g, ' ')).replace(OC_WORD, '');
  let l = tidy(last);
  let f = tidy(first);
  const comma = l.indexOf(',');
  if (comma >= 0) {
    f = f || clean(l.slice(comma + 1));
    l = clean(l.slice(0, comma));
  }
  f = clean(f.replace(/,/g, ' '));
  if (!letters(l) || PLACEHOLDER.test(l)) return null;
  return { last: l.toUpperCase().slice(0, 30), first: firstCase(f).slice(0, 30) };
}

/** A first name's case: each part's first letter up (Jean-Luc), never after an apostrophe: D'andre and Ma'ayan stay as typed. */
export const firstCase = (s) => String(s || '').replace(/(^|[\s-])(\p{Ll})/gu, (m, a, c) => a + c.toUpperCase());

/** One OC twice: the same letters in the last and the first name, in the same platoon. */
export const sameOc = (a, b) => letters(a.last) === letters(b.last) && letters(a.first) === letters(b.first) && (a.plt || 0) === (b.plt || 0);

/**
 * A roster as the event keeps it: every OC in a squad gets an id (in squad order, or the one
 * fixed at Ask OCs; then new ones), and each squad lists its ids.
 * @returns {{members: object[], squads: object[]}}
 */
export function numberRoster(merged, ids) {
  let next = Math.max(-1, ...Object.values(ids || {}).map(Number)) + 1;
  const members = [];
  const squads = merged.squads.map((sq) => {
    sq.ocs.forEach((o) => {
      const id = o.id !== undefined ? o.id : next++;
      members.push({ id, key: o.key, last: o.last, first: o.first, plt: o.plt, sqd: o.sqd, name: o.name, squad: sq.key });
    });
    return { key: sq.key, plt: sq.plt, sqd: sq.sqd, label: sq.label, ids: sq.ocs.map((o) => members.find((m) => m.key === o.key).id) };
  });
  members.sort((a, b) => a.id - b.id);
  return { members, squads };
}

/**
 * A pasted roster read and merged on its own (the Build 09 shape, for the tests): one
 * squad's PLT and SQD for rows that have none.
 */
export function parseRoster(text, one) {
  const r = readRoster(text);
  const src = Object.assign({ id: 'p', name: 'Pasted list', plt: one && inRange(unit(one.plt)) ? unit(one.plt) : null, sqd: one && inRange(unit(one.sqd)) ? unit(one.sqd) : null }, r);
  const m = mergeRoster([src], {});
  const n = numberRoster(m, {});
  const unread = r.unread.concat(m.block.filter((t) => r.unread.every((u) => t !== src.name + ': ' + u))); // the lines as read, then the roster's own blocks
  return { members: n.members, squads: n.squads, problems: unread.concat(m.hold.map((t) => t.replace(src.name + ': ', '')), m.list), unread, namesOnly: r.rows.filter((x) => !x.plt && !x.sqd).length, skipped: r.skipped, columns: r.columns, need: m.need };
}

export { ORD };
