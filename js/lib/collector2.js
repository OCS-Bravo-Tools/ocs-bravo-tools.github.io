/**
 * Talking to Collector v2, cadre mode's shared sheet. Every call is one POST of JSON sent as
 * text/plain, as with Collector v1, so browsers send it without a CORS preflight.
 *
 * Calls retry twice by themselves, quietly, after 1 s and 3 s (P3-2): Google answers an
 * error now and then. An answer from the sheet (good or bad) is final, except "busy" (the
 * script lock), which is retried too. Only after that does the page say anything.
 *
 * Answers: { ok: true, ... } or { ok: false, error } with error one of
 *   none      no evaluation with this id        removed   it was removed (removedAt)
 *   denied    wrong write code                   old       a newer board, or a started one, is already there
 *   taken     the id exists (make another)        limit     too many new evaluations today
 *   full      the evaluation's row caps            bad size / bad envelope / bad request
 */
import { collectorUrl } from './config.js';

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const WAITS = [1000, 3000];

/** One call with the quiet retries. Throws only when no answer came at all. */
export async function call(body, opts = {}) {
  const url = opts.url || collectorUrl();
  if (!url) throw new Error('no collector');
  const tries = opts.retries === undefined ? WAITS.length : opts.retries;
  let last = null;
  let busy = null;
  for (let i = 0; i <= tries; i++) {
    try {
      const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body), cache: 'no-store', redirect: 'follow' });
      let data = null;
      try {
        data = JSON.parse(await res.text());
      } catch (err) {
        data = null; // Google's own HTML error page
      }
      if (data && typeof data === 'object' && (data.ok || data.error !== 'busy')) return data;
      if (data) busy = data;
      else last = new Error('HTTP ' + res.status);
    } catch (err) {
      last = err;
    }
    if (i < tries) await sleep(WAITS[Math.min(i, WAITS.length - 1)]);
  }
  if (busy) return busy; // the sheet answered, but stayed busy: the page tries again later
  throw last || new Error('no answer');
}

export const newEval = (e, hw) => call({ op: 'new', e, hw });
/** st: the start mark of a started board (0 for a sign-up board). Once one is in, the sheet takes no other. */
export const putBoard = (e, w, brev, board, st) => call({ op: 'board', e, w, brev, board, st: st || 0 });
export const startEval = (e, w) => call({ op: 'start', e, w });
export const getState = (e, opts) => call({ op: 'state', e }, opts);
export const getBoard = (e, opts) => call({ op: 'get', e }, opts);
export const getRows = (e, w, cur) => call({ op: 'rows', e, w, cur });
export const forgetEval = (e, w) => call({ op: 'forget', e, w });
