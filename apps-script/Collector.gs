/** @OnlyCurrentDoc */
/**
 * Peer evaluation collector (Google Apps Script, free, nothing to maintain).
 * @OnlyCurrentDoc above limits it to the one sheet it is pasted into.
 *
 * Setup, once, about 3 minutes (the admin console shows the same 4 steps):
 *   1. Open a new Google Sheet. Extensions > Apps Script.
 *   2. Replace the code with this file (the console's "Copy the script"). Save.
 *   3. Deploy > New deployment > Web app. Who has access: Anyone. Deploy, then Allow.
 *   4. Paste the web app URL (it ends in /exec) into the console; it checks it by itself.
 *
 * What it stores: one row per submission, the envelope exactly as sent. The
 * envelope is encrypted for the admin, so this sheet holds no readable names,
 * answers or scores. Reading rows back needs the read key. The sheet keeps the first
 * key it is given. The console sets it right after creating links, with the key
 * already saved in the event file, so it is never lost; and OC links, which carry
 * this script's address, cannot change it. A later evaluation started with the
 * console's "New evaluation" reuses the same key. To reset: delete READ_KEY under
 * Project Settings > Script Properties.
 *
 * To change this code later, use Deploy > Manage deployments > Edit > Version: New
 * version, so the URL stays the same. A new deployment gets a new URL, and every link
 * already sent would point at the old one.
 */
var VERSION = 1;
var SHEET_NAME = 'results';
var MAX_BODY = 50000; // one sheet cell holds 50,000 characters; bigger results go by file or text

function doPost(e) {
  var body = e && e.postData ? e.postData.contents : '';
  if (!body || body.length > MAX_BODY) return reply({ ok: false, error: 'bad size' });
  var env;
  try {
    env = JSON.parse(body);
  } catch (err) {
    return reply({ ok: false, error: 'bad json' });
  }
  if (!env || env.v !== 1 || !env.e || env.r === undefined || !env.ct || !env.mac) {
    return reply({ ok: false, error: 'bad envelope' });
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return reply({ ok: false, error: 'busy' });
  try {
    sheet().appendRow([new Date(), String(env.e), String(env.r), Number(env.rev) || 1, body]);
  } finally {
    lock.releaseLock();
  }
  return reply({ ok: true });
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('READ_KEY');
  if (p.setup) {
    if (!key) {
      props.setProperty('READ_KEY', p.setup);
      return reply({ ok: true, setup: true });
    }
    return reply({ ok: key === p.setup, setup: false });
  }
  if (p.ping) return reply({ ok: true, app: 'peer-eval', v: VERSION, keyed: !!key });
  if (!key || p.k !== key) return reply({ ok: false, error: 'unauthorized' });
  var rows = sheet().getDataRange().getValues();
  var items = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(p.e)) items.push(String(rows[i][4]));
  }
  return reply({ ok: true, items: items });
}

function sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['received', 'event', 'oc', 'version', 'envelope (encrypted)']);
  }
  return sh;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
