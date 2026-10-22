/* ============================================================
   APLIKASI SANTRI - Backend Google Apps Script
   Data disimpan di Google Spreadsheet tempat skrip ini dipasang.
   Panduan lengkap ada di file PANDUAN-PEMASANGAN.pdf
   ============================================================ */

// >>> WAJIB DIUBAH sebelum deploy <<<
var ADMIN_USER = 'admin';
var ADMIN_PASS = 'GANTI-PASSWORD-INI';
var SESI_HARI = 30;           // lama login tetap aktif (hari)

/* ---------- Pintu masuk ---------- */
function doGet() { return out({ ok: true, info: 'API Aplikasi Santri aktif' }); }

function doPost(e) {
  var lock = LockService.getScriptLock(), got = false;
  try {
    lock.waitLock(25000); got = true;
    var q = JSON.parse(e.postData.contents);
    return out(route(q));
  } catch (err) {
    return out({ ok: false, err: String(err && err.message ? err.message : err) });
  } finally {
    if (got) lock.releaseLock();
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function route(q) {
  if (q.a === 'login') return login(q.u, q.p);
  var s = auth(q.tok);
  if (!s) return { ok: false, err: 'Sesi berakhir. Silakan masuk lagi.' };
  var db = readDB();
  if (s.role === 'santri') {
    var me = (db.S || []).filter(function (x) { return String(x.id) === String(s.sid); })[0];
    if (!me || me.aktif === false) return { ok: false, err: 'Akun tidak aktif.' };
  }
  switch (q.a) {
    case 'load':
      if (s.role === 'admin') return { ok: true, role: 'admin', db: db.S ? db : null, log: getLog(), seen: getSeen() };
      return { ok: true, role: 'santri', sid: s.sid, db: forSantri(db, s.sid), log: [] };
    case 'act': logAct(s, q.t); return { ok: true };
    case 'logout': dropTok(q.tok); return { ok: true };
  }
  if (s.role !== 'admin') return { ok: false, err: 'Khusus admin.' };
  switch (q.a) {
    case 'save': writeDB(q.db); mirror(q.db); return { ok: true };
    case 'mkak': var r = mkOne(q.sid, q.name); return { ok: true, username: r.username, password: r.password };
    case 'mkall': return { ok: true, list: mkAll(q.list || []) };
    case 'delak': delAkun(q.sid); return { ok: true };
  }
  return { ok: false, err: 'Aksi tidak dikenal.' };
}

/* ---------- Login & sesi ---------- */
function login(u, p) {
  u = String(u || '').trim().toLowerCase(); p = String(p || '');
  var cache = CacheService.getScriptCache(), key = 'f_' + u, n = +(cache.get(key) || 0);
  if (n >= 5) return { ok: false, err: 'Terlalu banyak percobaan. Coba lagi 10 menit lagi.' };
  function gagal() { cache.put(key, String(n + 1), 600); return { ok: false, err: 'Username atau kata sandi salah.' }; }

  if (u === ADMIN_USER.toLowerCase()) {
    if (ADMIN_PASS === 'GANTI-PASSWORD-INI') return { ok: false, err: 'Ubah ADMIN_PASS di Code.gs terlebih dahulu.' };
    if (p !== ADMIN_PASS) return gagal();
    return { ok: true, tok: newTok('admin', 0, 'Admin') };
  }
  var r = findAkun(u);
  if (!r || hash(p, r.salt) !== r.hash) return gagal();
  var db = readDB(), me = (db.S || []).filter(function (x) { return String(x.id) === String(r.sid); })[0];
  if (!me) return { ok: false, err: 'Data santri tidak ditemukan.' };
  if (me.aktif === false) return { ok: false, err: 'Akun dinonaktifkan oleh admin.' };
  logAct({ name: me.n, role: 'santri' }, 'Masuk ke aplikasi');
  return { ok: true, tok: newTok('santri', r.sid, me.n) };
}

function hash(p, salt) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + p, Utilities.Charset.UTF_8));
}
function rid() { return Utilities.getUuid().replace(/-/g, ''); }

function newTok(role, sid, name) {
  var P = PropertiesService.getScriptProperties(), all = P.getProperties(), now = Date.now(), k;
  for (k in all) {
    if (k.indexOf('t_') === 0) { try { if (JSON.parse(all[k]).exp < now) P.deleteProperty(k); } catch (e) { P.deleteProperty(k); } }
  }
  var t = rid() + rid().slice(0, 8);
  P.setProperty('t_' + t, JSON.stringify({ role: role, sid: sid, name: name, exp: now + SESI_HARI * 864e5 }));
  return t;
}
function auth(t) {
  if (!t) return null;
  var P = PropertiesService.getScriptProperties(), v = P.getProperty('t_' + t);
  if (!v) return null;
  var o = JSON.parse(v);
  if (o.exp < Date.now()) { P.deleteProperty('t_' + t); return null; }
  return o;
}
function dropTok(t) { if (t) PropertiesService.getScriptProperties().deleteProperty('t_' + t); }

/* ---------- Akun santri (tab "Akun", tersembunyi) ---------- */
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(n, hide) {
  var s = ss().getSheetByName(n);
  if (!s) { s = ss().insertSheet(n); if (hide) s.hideSheet(); }
  return s;
}
function akunRows() {
  var s = sheet('Akun', true), n = s.getLastRow();
  if (n < 1) { s.appendRow(['Username', 'Salt', 'Hash', 'ID Santri', 'Nama', 'Dibuat']); return []; }
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, 6).getValues();
}
function findAkun(u) {
  var rows = akunRows();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === u) return { username: rows[i][0], salt: String(rows[i][1]), hash: String(rows[i][2]), sid: rows[i][3] };
  }
  return null;
}
function slug(n) { return String(n).toLowerCase().replace(/[^a-z0-9]+/g, '').slice(0, 16) || 'santri'; }
function mkOne(sid, name) {
  var s = sheet('Akun', true), rows = akunRows(), pw = rid().slice(0, 10), salt = rid(), h = hash(pw, salt);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][3]) === String(sid)) {
      s.getRange(i + 2, 2, 1, 2).setValues([[salt, h]]);
      return { username: rows[i][0], password: pw };
    }
  }
  var base = slug(name), un = base, k = 1, used = {};
  rows.forEach(function (r) { used[String(r[0]).toLowerCase()] = 1; });
  used[ADMIN_USER.toLowerCase()] = 1;
  while (used[un]) { k++; un = base + k; }
  s.appendRow([un, salt, h, sid, name, new Date()]);
  return { username: un, password: pw };
}
function mkAll(list) {
  var have = {}; akunRows().forEach(function (r) { have[String(r[3])] = 1; });
  var out = [];
  list.forEach(function (x) {
    if (have[String(x.id)]) return;
    var r = mkOne(x.id, x.name); have[String(x.id)] = 1;
    out.push({ id: x.id, username: r.username, password: r.password });
  });
  return out;
}
function delAkun(sid) {
  var s = sheet('Akun', true), rows = akunRows();
  for (var i = rows.length - 1; i >= 0; i--) if (String(rows[i][3]) === String(sid)) s.deleteRow(i + 2);
}

/* ---------- Penyimpanan data (tab "_Data", tersembunyi) ---------- */
function readDB() {
  var s = sheet('_Data', true), n = s.getLastRow();
  if (n < 1) return {};
  var txt = s.getRange(1, 1, n, 1).getValues().map(function (r) { return r[0]; }).join('');
  try { return JSON.parse(txt) || {}; } catch (e) { return {}; }
}
function writeDB(db) {
  var s = sheet('_Data', true), j = JSON.stringify(db), parts = [];
  for (var i = 0; i < j.length; i += 40000) parts.push([j.substr(i, 40000)]);
  if (!parts.length) parts.push(['{}']);
  s.clear();
  var r = s.getRange(1, 1, parts.length, 1);
  r.setNumberFormat('@');
  r.setValues(parts);
}
function forSantri(db, sid) {
  return {
    org: db.org, O: [], nid: db.nid,
    S: (db.S || []).filter(function (x) { return String(x.id) === String(sid); }),
    P: db.P || [], M: [], G: db.G || [], N: db.N || [], log: []
  };
}

/* ---------- Log aktivitas (tab "Log") ---------- */
function logAct(s, t) {
  var sh = sheet('Log');
  if (sh.getLastRow() < 1) sh.appendRow(['Waktu', 'Siapa', 'Aktivitas']);
  sh.appendRow([new Date(), s.name || s.role, String(t || '').slice(0, 300)]);
}
function getLog() {
  var sh = sheet('Log'), n = sh.getLastRow();
  if (n < 2) return [];
  var st = Math.max(2, n - 99), v = sh.getRange(st, 1, n - st + 1, 3).getValues();
  return v.reverse().map(function (r) { return { t: new Date(r[0]).getTime(), w: r[1], a: r[2] }; });
}

function getSeen() {
  var sh = sheet('Log'), n = sh.getLastRow();
  if (n < 2) return {};
  var st = Math.max(2, n - 499), v = sh.getRange(st, 1, n - st + 1, 3).getValues(), m = {};
  v.forEach(function (r) { m[r[1]] = new Date(r[0]).getTime(); });
  return m;
}

/* ---------- Cermin ke tab Spreadsheet yang mudah dibaca ---------- */
function put(name, head, rows) {
  var s = sheet(name);
  s.clear();
  s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) s.getRange(2, 1, rows.length, head.length).setValues(rows);
  s.setFrozenRows(1);
}
function mirror(db) {
  var S = db.S || [], t = function (x) { return x ? new Date(x) : ''; };
  put('Santri', ['ID', 'Nama', 'Tempat, Tanggal Lahir', 'Alamat', 'Kelas', 'Kamar', 'Status', 'Username'],
    S.map(function (x) { return [x.id, x.n, x.ttl, x.al, x.kl, x.km, x.aktif === false ? 'Nonaktif' : 'Aktif', x.un || '']; }));
  var pg = [];
  (db.O || []).forEach(function (o) { (o.m || []).forEach(function (m) { pg.push([o.n, m.n, m.j, m.k, m.r]); }); });
  put('Pengurus', ['Organisasi', 'Nama', 'Jabatan', 'Kelas', 'Kamar'], pg);
  var nl = [], ab = [], hf = [], tb = [];
  S.forEach(function (x) {
    (x.nilai || []).forEach(function (q) { nl.push([x.n, x.kl, q.m, q.v, q.t || '']); });
    (x.absen || []).forEach(function (q) { ab.push([q.d, x.n, x.kl, q.s]); });
    (x.haf || []).forEach(function (q) { hf.push([x.n, x.kl, q.s, q.a, q.d]); });
    (x.tab || []).forEach(function (q) { tb.push([t(q.t), x.n, q.j, q.k, q.n]); });
  });
  ab.sort(function (a, b) { return a[0] < b[0] ? 1 : a[0] > b[0] ? -1 : 0; });
  put('Nilai', ['Santri', 'Kelas', 'Mata Pelajaran', 'Nilai', 'Semester'], nl);
  put('Absen', ['Tanggal', 'Santri', 'Kelas', 'Status'], ab);
  put('Hafalan', ['Santri', 'Kelas', 'Surat', 'Ayat', 'Tanggal'], hf);
  put('Tabungan', ['Waktu', 'Santri', 'Jenis', 'Keperluan', 'Nominal'], tb);
  put('Agenda', ['Tanggal', 'Judul', 'Jam', 'Tempat'], (db.G || []).map(function (g) { return [g.d, g.j, g.w || '', g.p || '']; }));
  put('Pengumuman', ['Waktu', 'Judul', 'Isi'], (db.N || []).map(function (n) { return [t(n.t), n.j, n.i]; }));
  put('Perpustakaan', ['Judul', 'Tautan Google Drive', 'Ditambahkan'], (db.P || []).map(function (x) { return [x.n, x.u || '', x.d || '']; }));
  put('Materi', ['Judul', 'Tautan Google Drive', 'Ditambahkan'], (db.M || []).map(function (x) { return [x.n, x.u || '', x.d || '']; }));
}
