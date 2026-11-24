/**
 * Learner Gate Pass backend.
 * Setup: open a Google Sheet > Extensions > Apps Script > paste this file.
 * 1) Change ADMIN_PASSWORD below.  2) Run setup() once.
 * 3) Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Optional: Project Settings > Script properties: SMS_API_KEY, SMS_SENDER (Semaphore).
 */
const ADMIN_USER = 'superadmin';
const ADMIN_PASSWORD = 'CHANGE-ME-123';
const HEAD = {
  Users: ['id', 'name', 'user', 'role', 'salt', 'hash'],
  Learners: ['id', 'lrn', 'name', 'grade', 'pn', 'pp', 'info', 'status', 'lastTs'],
  Logs: ['id', 'lid', 'type', 't', 'by'],
  Msgs: ['id', 'to', 'parent', 'body', 't', 'status']
};
const CAN = { super: ['admin', 'teacher', 'gate'], admin: ['teacher', 'gate'] };
const TZ = Session.getScriptTimeZone();
const sh = n => SpreadsheetApp.getActive().getSheetByName(n);
const rid = () => Utilities.getUuid().slice(0, 8);

function setup() {
  Object.keys(HEAD).forEach(n => {
    const s = sh(n) || SpreadsheetApp.getActive().insertSheet(n);
    s.getRange('A:Z').setNumberFormat('@'); // plain text: keeps 09xx numbers, blocks formulas
    if (s.getLastRow() === 0) { s.appendRow(HEAD[n]); s.setFrozenRows(1); }
  });
  if (!rows('Users').length) addRow('Users', mkUser('Super Admin', ADMIN_USER, 'super', ADMIN_PASSWORD));
}

/* ---------- sheet helpers ---------- */
function rows(n) {
  const v = sh(n).getDataRange().getValues(), h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, String(r[i])])));
}
function addRow(n, o) { sh(n).appendRow(HEAD[n].map(k => o[k] === undefined ? '' : String(o[k]))); }
function findIdx(n, id) { return sh(n).getDataRange().getValues().findIndex((r, j) => j > 0 && r[0] === id); }
function setRow(n, id, o) {
  const i = findIdx(n, id); if (i < 0) return false;
  const s = sh(n), cur = s.getRange(i + 1, 1, 1, HEAD[n].length).getValues()[0];
  s.getRange(i + 1, 1, 1, HEAD[n].length).setValues([HEAD[n].map((k, c) => o[k] === undefined ? cur[c] : String(o[k]))]);
  return true;
}
function delRow(n, id) { const i = findIdx(n, id); if (i > 0) sh(n).deleteRow(i + 1); }

/* ---------- auth ---------- */
const hash = (salt, p) => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + p));
function mkUser(name, user, role, pass) { const salt = Utilities.getUuid(); return { id: rid(), name, user, role, salt, hash: hash(salt, pass) }; }
function session(t) {
  const s = t && CacheService.getScriptCache().get('s_' + t);
  if (!s) throw new Error('auth');
  return JSON.parse(s);
}
function login(q) {
  const c = CacheService.getScriptCache(), k = 'f_' + q.user, n = +c.get(k) || 0;
  if (n >= 5) throw new Error('locked');
  const u = rows('Users').find(x => x.user === String(q.user));
  if (!u || u.hash !== hash(u.salt, String(q.pass))) { c.put(k, n + 1, 900); throw new Error('bad'); }
  c.remove(k);
  const me = { id: u.id, name: u.name, role: u.role }, token = Utilities.getUuid() + Utilities.getUuid();
  c.put('s_' + token, JSON.stringify(me), 21600); // 6 hours
  return { token, me };
}

/* ---------- API ---------- */
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return out({ ok: true, msg: 'Gate pass API' }); }
function doPost(e) {
  try {
    const q = JSON.parse(e.postData.contents);
    return out(Object.assign({ ok: true }, route(q)));
  } catch (x) { return out({ ok: false, error: String(x.message || x) }); }
}
const pubL = l => ({ id: l.id, lrn: l.lrn, name: l.name, grade: l.grade, pn: l.pn, pp: l.pp, info: l.info });

function route(q) {
  if (q.action === 'login') return login(q);
  const me = session(q.token);
  const need = (...r) => { if (!r.includes(me.role)) throw new Error('forbidden'); };
  switch (q.action) {
    case 'logout': CacheService.getScriptCache().remove('s_' + q.token); return {};
    case 'scan': return scan(me, q.code);
    case 'list_learners': need('super', 'admin', 'teacher'); return { rows: rows('Learners').map(pubL) };
    case 'save_learner': need('super', 'admin', 'teacher'); return saveLearner(q.d);
    case 'list_users': need('super', 'admin');
      return { rows: rows('Users').filter(u => CAN[me.role].includes(u.role)).map(u => ({ id: u.id, name: u.name, user: u.user, role: u.role })) };
    case 'save_user': need('super', 'admin'); return saveUser(me, q.d);
    case 'delete_user': {
      need('super', 'admin');
      const u = rows('Users').find(x => x.id === q.id);
      if (!u || !CAN[me.role].includes(u.role)) throw new Error('forbidden');
      delRow('Users', q.id); return {};
    }
    case 'change_password': {
      const p = String(q.pass || ''), salt = Utilities.getUuid();
      if (p.length < 6) throw new Error('bad');
      setRow('Users', me.id, { salt, hash: hash(salt, p) }); return {};
    }
    case 'list_logs': need('super', 'admin', 'teacher');
      return { rows: rows('Logs').slice(-300).map(x => Object.assign(x, { t: Number(x.t) })) };
    case 'list_msgs': need('super', 'admin');
      return { rows: rows('Msgs').slice(-200).map(x => Object.assign(x, { t: Number(x.t) })) };
  }
  throw new Error('bad');
}

function saveLearner(d) {
  if (!d || !d.name || !d.pp) throw new Error('bad');
  if (d.lrn && rows('Learners').some(x => x.lrn === d.lrn && x.id !== d.id)) throw new Error('dup');
  const o = { lrn: d.lrn, name: d.name, grade: d.grade, pn: d.pn, pp: d.pp, info: d.info };
  if (d.id) { if (!setRow('Learners', d.id, o)) throw new Error('bad'); }
  else addRow('Learners', Object.assign({ id: 'L' + rid(), status: 'OUT', lastTs: 0 }, o));
  return {};
}

function saveUser(me, d) {
  if (!d || !d.name || !d.user || String(d.pass || '').length < 6 || !CAN[me.role].includes(d.role)) throw new Error('forbidden');
  if (rows('Users').some(u => u.user === d.user)) throw new Error('taken');
  addRow('Users', mkUser(d.name, d.user, d.role, d.pass));
  return {};
}

/* ---------- gate scan ---------- */
function scan(me, code) {
  code = String(code || '').trim().replace(/^GATEPASS:/, '');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  let l, now, type, view;
  try {
    l = rows('Learners').find(x => x.id === code || x.lrn === code);
    if (!l) throw new Error('unknown');
    now = Date.now();
    view = me.role === 'gate' ? { id: l.id, name: l.name, grade: l.grade } : pubL(l);
    if (now - Number(l.lastTs || 0) < 10000) return { duplicate: true, learner: view, type: l.status, t: Number(l.lastTs) };
    type = l.status === 'IN' ? 'OUT' : 'IN'; // toggles on every scan
    setRow('Learners', l.id, { status: type, lastTs: now });
    addRow('Logs', { id: rid(), lid: l.id, type, t: now, by: me.name });
  } finally { lock.releaseLock(); }
  const d = new Date(now);
  const body = l.name + ' has ' + (type === 'IN' ? 'entered' : 'left') + ' the school campus at ' +
    Utilities.formatDate(d, TZ, 'h:mm a') + ' on ' + Utilities.formatDate(d, TZ, 'MMM d, yyyy') + '.';
  const status = notify(l.pp, body);
  addRow('Msgs', { id: rid(), to: l.pp, parent: l.pn, body, t: now, status });
  return { type, t: now, learner: view, sms: status };
}

/* SMS via Semaphore (PH). Swap the URL/payload for Twilio or another gateway if you prefer. */
function notify(to, body) {
  const P = PropertiesService.getScriptProperties(), key = P.getProperty('SMS_API_KEY');
  if (!key) return 'Queued (no SMS key set)';
  try {
    const pl = { apikey: key, number: to, message: body }, sn = P.getProperty('SMS_SENDER');
    if (sn) pl.sendername = sn;
    const r = UrlFetchApp.fetch('https://api.semaphore.co/api/v4/messages', { method: 'post', payload: pl, muteHttpExceptions: true });
    return r.getResponseCode() < 300 ? 'Sent' : 'Failed (' + r.getResponseCode() + ')';
  } catch (x) { return 'Failed'; }
}
