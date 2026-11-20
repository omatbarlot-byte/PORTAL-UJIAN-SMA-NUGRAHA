/**
 * =====================================================================
 *  CBT INFORMATIKA SMA NUGRAHA - BACKEND (Google Apps Script Web App)
 *  Versi performa tinggi (untuk 500+ peserta bersamaan)
 * =====================================================================
 *  Deploy: Deploy > New deployment > Web app
 *          Execute as: Me | Who has access: Anyone
 *  Setelah menempel kode ini: jalankan fungsi setup() SEKALI
 *  (membuat sheet pendukung + memasang pemicu otomatis), lalu
 *  Deploy > Manage deployments > pensil > New version > Deploy.
 *
 *  Cara kerja agar cepat:
 *   - Data siswa/guru/ujian dibaca dari CacheService (bukan dari Sheet) pada tiap permintaan.
 *   - Status sesi ujian (pelanggaran, kunci, waktu) disimpan di cache dan
 *     disalin ke Sheet "Sesi Ujian" tiap 1 menit oleh pemicu flushSessions.
 *   - Sinyal "heartbeat" siswa hanya menyentuh cache (tanpa kunci global).
 *   - Token login bertanda tangan (HMAC), tidak disimpan di server.
 * =====================================================================
 */

var SS_ID = '1Xh-Bqr-lOCrUCopS-VB0e5xCTvWIF-17lXDd-CFsU6c';

var SHEET_NAMES = {
  SISWA: ['Siswa'],
  GURU: ['Guru / Admin', 'Guru/Admin', 'Guru', 'Admin'],
  EXAM: ['Pengaturan Ujian'],
  SESI: ['Sesi Ujian'],
  LOG: ['Log Pelanggaran']
};
var EXAM_HEAD = ['ID Ujian', 'Nama Sekolah', 'Nama Ujian', 'Link Google Form', 'Batas Pelanggaran', 'Durasi (Menit)', 'Status (Aktif/Nonaktif)'];
var SESI_HEAD = ['Sesi ID', 'Nama Sekolah', 'Nomor Peserta', 'Nama Siswa', 'Kelas', 'ID Ujian', 'Nama Ujian', 'Mulai (ms)', 'Berakhir (ms)', 'Pelanggaran', 'Status', 'Terakhir Aktif (ms)', 'Waktu Selesai', 'Catatan'];
var LOG_HEAD = ['Waktu', 'Sesi ID', 'Nama Sekolah', 'Nomor Peserta', 'Nama Siswa', 'ID Ujian', 'Jenis', 'Detail', 'Pelanggaran Ke'];

var ST = { RUN: 'Sedang Mengerjakan', LOCK: 'Terkunci', DONE: 'Selesai' };
var TOKEN_TTL = 6 * 60 * 60;        // detik
var SES_TTL = 6 * 60 * 60;          // detik (maks CacheService)
var ONLINE_WINDOW_MS = 100 * 1000;  // online jika heartbeat < 100 detik
var MAX_LOGIN_FAIL = 5;
var GRACE_MS = 120 * 1000;          // masa tenggang setelah waktu habis (harus > CONFIG.SUBMIT_GRACE_SEC di index.html)
var TTL_STUDENTS = 600, TTL_GURU = 600, TTL_EXAM = 15, TTL_SCHOOLS = 300, TTL_MON = 6;

/* ----------------------------- Router ------------------------------ */

var ACTIONS = {
  getSchools: getSchools_, loginSiswa: loginSiswa_, loginAdmin: loginAdmin_, getExams: getExams_,
  startExam: startExam_, heartbeat: heartbeat_, logViolation: logViolation_, finishExam: finishExam_,
  saveExam: saveExam_, deleteExam: deleteExam_, getMonitor: getMonitor_, resetSession: resetSession_, getLogs: getLogs_
};
// Hanya operasi admin yang jarang yang memakai kunci global. Jalur siswa TANPA kunci.
var LOCKED_ACTIONS = ['saveExam', 'deleteExam'];

function doGet(e) {
  if (e && e.parameter && e.parameter.debug === '1') {
    var info = {};
    try {
      ss_().getSheets().forEach(function (sh) {
        var cols = Math.max(1, sh.getLastColumn());
        info[sh.getName()] = { baris_data: Math.max(0, sh.getLastRow() - 1), header: sh.getLastRow() > 0 ? sh.getRange(1, 1, 1, cols).getValues()[0] : [] };
      });
    } catch (err) { info = { error: String(err && err.message ? err.message : err) }; }
    return json_({ ok: true, debug: info });
  }
  return json_({ ok: true, service: 'CBT Informatika SMA Nugraha API', time: Date.now() });
}

function doPost(e) {
  var lock = null, got = false;
  try {
    var req = JSON.parse(e.postData.contents);
    var fn = ACTIONS[req.action];
    if (!fn) return json_({ ok: false, error: 'Aksi tidak dikenal.' });
    if (LOCKED_ACTIONS.indexOf(req.action) > -1) {
      lock = LockService.getScriptLock();
      lock.waitLock(15000);
      got = true;
    }
    var res = fn(req) || {};
    res.ok = true;
    return json_(res);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    if (got) { try { lock.releaseLock(); } catch (x) {} }
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ----------------------------- Helpers ----------------------------- */

var SS_MEMO_ = null;
function ss_() { if (!SS_MEMO_) SS_MEMO_ = SpreadsheetApp.openById(SS_ID); return SS_MEMO_; }

function sheet_(names, create, head) {
  var ss = ss_();
  var list = [].concat(names);
  for (var i = 0; i < list.length; i++) {
    var s = ss.getSheetByName(list[i]);
    if (s) return s;
  }
  var all = ss.getSheets();
  for (var i = 0; i < list.length; i++) {
    var hit = all.filter(function (x) { return x.getName().trim().toLowerCase() === list[i].toLowerCase(); })[0];
    if (hit) return hit;
  }
  if (create) {
    var s = ss.insertSheet(list[0]);
    s.appendRow(head);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#4F46E5').setFontColor('#FFFFFF');
    return s;
  }
  throw new Error('Sheet "' + list[0] + '" tidak ditemukan di spreadsheet.');
}

function norm_(s) { return String(s).trim().toLowerCase().replace(/\s+/g, ' '); }
function s_(v) { return String(v == null ? '' : v).trim(); }

var KNOWN_ = { 'nama sekolah': 1, 'nama siswa': 1, 'nomor peserta': 1, 'kelas': 1, 'password': 1, 'nama guru': 1, 'username / nip': 1,
  'id ujian': 1, 'nama ujian': 1, 'link google form': 1, 'batas pelanggaran': 1, 'durasi (menit)': 1, 'status (aktif/nonaktif)': 1 };

/** Samakan variasi nama header ke nama baku. */
function canon_(h) {
  var x = norm_(h);
  if (!x) return '';
  if (x.indexOf('sekolah') > -1) return 'nama sekolah';
  if (x.indexOf('nama guru') > -1) return 'nama guru';
  if (x.indexOf('nama siswa') > -1) return 'nama siswa';
  if (x.indexOf('nomor') > -1 || x.indexOf('peserta') > -1) return 'nomor peserta';
  if (x.indexOf('username') > -1 || x.indexOf('nip') > -1) return 'username / nip';
  if (x.indexOf('kelas') > -1) return 'kelas';
  if (x.indexOf('password') > -1 || x.indexOf('sandi') > -1) return 'password';
  if (x.indexOf('id ujian') > -1) return 'id ujian';
  if (x.indexOf('nama ujian') > -1) return 'nama ujian';
  if (x.indexOf('link') > -1 || x.indexOf('form') > -1) return 'link google form';
  if (x.indexOf('batas') > -1) return 'batas pelanggaran';
  if (x.indexOf('durasi') > -1) return 'durasi (menit)';
  if (x.indexOf('status') > -1) return 'status (aktif/nonaktif)';
  return x;
}

/** Baca sheet menjadi array objek; baris header dicari otomatis di 6 baris pertama. */
function table_(sh) {
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var hi = 0, best = -1, r, c;
  for (r = 0; r < Math.min(v.length, 6); r++) {
    var sc = 0;
    for (c = 0; c < v[r].length; c++) { if (KNOWN_[canon_(v[r][c])]) sc++; }
    if (sc > best) { best = sc; hi = r; }
  }
  var h = v[hi].map(canon_);
  var out = [];
  for (var i = hi + 1; i < v.length; i++) {
    var o = { _row: i + 1 }, filled = false;
    for (var j = 0; j < h.length; j++) {
      if (!h[j]) continue;
      if (o[h[j]] === undefined) o[h[j]] = v[i][j];
      if (s_(v[i][j]) !== '') filled = true;
    }
    if (filled) out.push(o);
  }
  return out;
}

function isWildcard_(school) {
  var a = s_(school).toLowerCase();
  return a === 'semua' || a === '*' || a === 'all';
}
function canSchool_(admin, school) {
  return isWildcard_(admin.school) || s_(admin.school).toLowerCase() === s_(school).toLowerCase();
}

function throttle_(key) {
  var c = CacheService.getScriptCache();
  var n = Number(c.get(key) || 0);
  if (n >= MAX_LOGIN_FAIL) throw new Error('Terlalu banyak percobaan gagal. Coba lagi dalam 5 menit.');
  return {
    fail: function () { c.put(key, String(n + 1), 300); },
    ok: function () { c.remove(key); }
  };
}

/** Ambil src jika admin menempelkan kode <iframe>, pastikan ?embedded=true. */
function toEmbed_(raw) {
  var url = s_(raw);
  var m = url.match(/src=["']([^"']+)["']/i);
  if (m) url = m[1].replace(/&amp;/g, '&');
  if (!/^https:\/\/docs\.google\.com\/forms\//i.test(url)) {
    throw new Error('Link harus berupa Google Form (https://docs.google.com/forms/...). Link pendek forms.gle tidak bisa di-embed.');
  }
  if (url.indexOf('embedded=true') === -1) url += (url.indexOf('?') > -1 ? '&' : '?') + 'embedded=true';
  return url;
}


/* ------------------------- Cache util ------------------------- */

function cache_() { return CacheService.getScriptCache(); }

function getAllCache_(keys) {
  var c = cache_(), out = {};
  for (var i = 0; i < keys.length; i += 100) {
    var part = c.getAll(keys.slice(i, i + 100));
    for (var k in part) out[k] = part[k];
  }
  return out;
}
/** Nilai besar dipecah (batas CacheService 100KB per item). */
function putBig_(key, str, ttl) {
  var c = cache_(), o = {}, n = 0;
  for (var i = 0; i < str.length; i += 30000) { o[key + ':' + n] = str.substr(i, 30000); n++; }
  o[key + ':n'] = String(n);
  c.putAll(o, ttl);
}
function getBig_(key) {
  var c = cache_(), n = c.get(key + ':n');
  if (!n) return null;
  n = Number(n);
  var ks = [], i;
  for (i = 0; i < n; i++) ks.push(key + ':' + i);
  var m = c.getAll(ks), s = '';
  for (i = 0; i < n; i++) { var p = m[key + ':' + i]; if (p == null) return null; s += p; }
  return s;
}
function delBig_(key) {
  var c = cache_(), n = c.get(key + ':n'), ks = [key + ':n'];
  if (n) for (var i = 0; i < Number(n); i++) ks.push(key + ':' + i);
  c.removeAll(ks);
}

/** Baca tabel dari cache; jika kosong, bangun ulang dari Sheet (hanya satu eksekusi yang membangun). */
function tableCached_(tag, names, ttl, force) {
  var key = 'tbl_' + tag, s;
  if (!force) { s = getBig_(key); if (s) return JSON.parse(s); }
  var lock = LockService.getScriptLock(), got = false;
  try {
    got = lock.tryLock(8000);
    if (!force) { s = getBig_(key); if (s) return JSON.parse(s); }
    var rows = table_(sheet_(names));
    putBig_(key, JSON.stringify(rows), ttl);
    return rows;
  } finally { if (got) { try { lock.releaseLock(); } catch (x) {} } }
}
function studentsRows_() { return tableCached_('siswa', SHEET_NAMES.SISWA, TTL_STUDENTS); }
function guruRows_() { return tableCached_('guru', SHEET_NAMES.GURU, TTL_GURU); }
function examRows_() { return tableCached_('exam', SHEET_NAMES.EXAM, TTL_EXAM); }

/* ------------------- Token bertanda tangan (stateless) ------------------- */

function secret_() {
  var c = cache_().get('sec');
  if (c) return c;
  var p = PropertiesService.getScriptProperties();
  var s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  cache_().put('sec', s, SES_TTL);
  return s;
}
function sign_(data) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(data, secret_())); }
function issueToken_(user) {
  var u = JSON.parse(JSON.stringify(user));
  u.exp = Date.now() + TOKEN_TTL * 1000;
  var body = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(u)).getBytes());
  return body + '.' + sign_(body);
}
function auth_(req, role) {
  var t = s_(req.token);
  if (!t) throw new Error('Token tidak ada. Silakan login.');
  var i = t.indexOf('.');
  var bad = 'Sesi login berakhir. Silakan login ulang.';
  if (i < 1) throw new Error(bad);
  var body = t.substring(0, i);
  if (t.substring(i + 1) !== sign_(body)) throw new Error(bad);
  var u = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(body)).getDataAsString());
  if (!u.exp || u.exp < Date.now()) throw new Error(bad);
  if (role && u.role !== role) throw new Error('Akses ditolak.');
  return u;
}

/* ----------------------------- Ujian ----------------------------- */

function examOut_(r, withLink) {
  var o = {
    id: s_(r['id ujian']),
    school: s_(r['nama sekolah']),
    name: s_(r['nama ujian']),
    max: Number(r['batas pelanggaran']) || 3,
    duration: Number(r['durasi (menit)']) || 60,
    status: /^aktif$/i.test(s_(r['status (aktif/nonaktif)'])) ? 'Aktif' : 'Nonaktif'
  };
  if (withLink) o.link = s_(r['link google form']);
  return o;
}
function findExam_(id) {
  var r = examRows_().filter(function (x) { return s_(x['id ujian']) === s_(id); })[0];
  if (!r) throw new Error('Ujian tidak ditemukan.');
  return r;
}
function freshExamRows_() { return table_(sheet_(SHEET_NAMES.EXAM)); }

/* ----------------------------- Sesi ----------------------------- */

function sesiSheet_() { return sheet_(SHEET_NAMES.SESI, true, SESI_HEAD); }
function logSheet_() { return sheet_(SHEET_NAMES.LOG, true, LOG_HEAD); }

function sid_(school, nomor, examId) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s_(school).toLowerCase() + '|' + s_(nomor) + '|' + s_(examId));
  return d.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').substring(0, 12);
}
function nowStr_() { return Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss'); }
function getSes_(id) { var raw = cache_().get('ses:' + id); return raw ? JSON.parse(raw) : null; }
function putSes_(s) { cache_().put('ses:' + s.id, JSON.stringify(s), SES_TTL); }

function sesFromRow_(v) {
  return {
    id: s_(v[0]), school: s_(v[1]), nomor: s_(v[2]), nama: s_(v[3]), kelas: s_(v[4]), examId: s_(v[5]), examName: s_(v[6]),
    start: Number(v[7]) || 0, end: Number(v[8]) || 0, violations: Number(v[9]) || 0, status: s_(v[10]),
    last: Number(v[11]) || 0, finishedAt: s_(v[12]), note: s_(v[13])
  };
}
/** Cache dulu; jika tidak ada (mis. cache terhapus) ambil dari Sheet. */
function loadSes_(id) {
  var s = getSes_(id);
  if (s) return s;
  if (!id) return null;
  var sh = sesiSheet_();
  var f = sh.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!f) return null;
  s = sesFromRow_(sh.getRange(f.getRow(), 1, 1, SESI_HEAD.length).getValues()[0]);
  putSes_(s);
  return s;
}
function ownSes_(req, u) {
  var s = loadSes_(s_(req.sessionId));
  if (!s || s.nomor !== u.nomor || s.school !== u.school) throw new Error('Sesi tidak valid.');
  return s;
}
function expire_(s) {
  if (s.status === ST.RUN && Date.now() > s.end + GRACE_MS) {
    s.status = ST.DONE; s.finishedAt = nowStr_(); s.note = 'Waktu habis';
    putSes_(s);
  }
  return s;
}
function stateOf_(s, max) {
  return { id: s.id, status: s.status, violations: s.violations, max: max || s.max || 3, endsAt: s.end, serverNow: Date.now(), note: s.note };
}
function appendRetry_(sh, row) {
  for (var i = 0; i < 3; i++) {
    try { sh.appendRow(row); return; }
    catch (e) { if (i === 2) throw e; Utilities.sleep(250 * (i + 1)); }
  }
}

/** Semua sesi (Sheet digabung status terbaru dari cache). */
function allSesMerged_() {
  var sh = sesiSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, SESI_HEAD.length).getValues();
  var ids = [], i;
  for (i = 0; i < rows.length; i++) { var id = s_(rows[i][0]); if (id) ids.push('ses:' + id); }
  var m = getAllCache_(ids), seen = {}, out = [];
  for (i = 0; i < rows.length; i++) {
    var sid = s_(rows[i][0]);
    if (!sid || seen[sid]) continue;
    seen[sid] = 1;
    out.push(m['ses:' + sid] ? JSON.parse(m['ses:' + sid]) : sesFromRow_(rows[i]));
  }
  return out;
}

/** Salin status dari cache ke Sheet secara massal (dipanggil pemicu tiap 1 menit). */
function flushSessions_() {
  var lock = LockService.getScriptLock(), got = false;
  try {
    got = lock.tryLock(8000);
    if (!got) return;
    var sh = sesiSheet_(), last = sh.getLastRow();
    if (last < 2) return;
    var vals = sh.getRange(2, 1, last - 1, SESI_HEAD.length).getValues();
    var keys = vals.map(function (v) { return 'ses:' + s_(v[0]); });
    var m = getAllCache_(keys);
    var out = [], changed = false;
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i], raw = m[keys[i]];
      var cur = [v[8], v[9], v[10], v[11], v[12], v[13]];
      if (raw) {
        var s = JSON.parse(raw);
        var nw = [s.end, s.violations, s.status, s.last || 0, s.finishedAt ? "'" + s.finishedAt : '', s.note || ''];
        if (Number(cur[0]) !== Number(s.end) || Number(cur[1]) !== Number(s.violations) || s_(cur[2]) !== s_(s.status) ||
            Math.abs((Number(cur[3]) || 0) - (Number(s.last) || 0)) >= 30000 || s_(cur[4]) !== s_(s.finishedAt) || s_(cur[5]) !== s_(s.note)) {
          changed = true; out.push(nw); continue;
        }
      }
      out.push(cur);
    }
    if (changed) sh.getRange(2, 9, out.length, 6).setValues(out);
  } finally { if (got) { try { lock.releaseLock(); } catch (x) {} } }
}
function flushSessions() { flushSessions_(); }

function warmCache() {
  tableCached_('siswa', SHEET_NAMES.SISWA, TTL_STUDENTS, true);
  tableCached_('guru', SHEET_NAMES.GURU, TTL_GURU, true);
  tableCached_('exam', SHEET_NAMES.EXAM, TTL_EXAM, true);
  cache_().remove('schools');
  getSchools_();
}

/* ----------------------------- Actions ----------------------------- */

function getSchools_() {
  var c = cache_().get('schools');
  if (c) return { schools: JSON.parse(c) };
  var set = {};
  function add(rows) {
    rows.forEach(function (r) { var n = s_(r['nama sekolah']); if (n && !isWildcard_(n)) set[n] = 1; });
  }
  try { add(studentsRows_()); } catch (e) {}
  if (!Object.keys(set).length) {
    try { add(examRows_()); } catch (e1) {}
    try { add(guruRows_()); } catch (e2) {}
  }
  var list = Object.keys(set).sort();
  if (list.length) cache_().put('schools', JSON.stringify(list), TTL_SCHOOLS);
  return { schools: list };
}

function loginSiswa_(req) {
  var school = s_(req.school), nomor = s_(req.nomor), pass = s_(req.password);
  if (!school || !nomor || !pass) throw new Error('Lengkapi sekolah, nomor peserta, dan password.');
  var th = throttle_('fail_s_' + school + '_' + nomor);
  var r = studentsRows_().filter(function (x) {
    return s_(x['nama sekolah']).toLowerCase() === school.toLowerCase() && s_(x['nomor peserta']) === nomor;
  })[0];
  if (!r || s_(r['password']) !== pass) { th.fail(); throw new Error('Nomor peserta atau password salah.'); }
  th.ok();
  var user = { role: 'siswa', school: s_(r['nama sekolah']), nomor: nomor, nama: s_(r['nama siswa']), kelas: s_(r['kelas']) };
  return { token: issueToken_(user), user: user, exams: listExamsForStudent_(user) };
}

function loginAdmin_(req) {
  var username = s_(req.username), pass = s_(req.password);
  if (!username || !pass) throw new Error('Lengkapi username dan password.');
  var th = throttle_('fail_a_' + username);
  var r = guruRows_().filter(function (x) { return s_(x['username / nip']) === username; })[0];
  if (!r || s_(r['password']) !== pass) { th.fail(); throw new Error('Username/NIP atau password salah.'); }
  th.ok();
  var user = { role: 'admin', school: s_(r['nama sekolah']), nama: s_(r['nama guru']), username: username };
  user.wildcard = isWildcard_(user.school);
  return { token: issueToken_(user), user: user, exams: listExamsForAdmin_(user), schools: getSchools_().schools };
}

function listExamsForStudent_(user) {
  var rows = examRows_().filter(function (r) {
    return s_(r['nama sekolah']).toLowerCase() === user.school.toLowerCase() && /^aktif$/i.test(s_(r['status (aktif/nonaktif)']));
  });
  var keys = rows.map(function (r) { return 'ses:' + sid_(user.school, user.nomor, s_(r['id ujian'])); });
  var m = getAllCache_(keys), now = Date.now();
  return rows.map(function (r, i) {
    var e = examOut_(r, false), raw = m[keys[i]], st = null;
    if (raw) { var s = JSON.parse(raw); st = s.status; if (st === ST.RUN && now > s.end + GRACE_MS) st = ST.DONE; }
    e.myStatus = st;
    return e;
  });
}
function listExamsForAdmin_(user) {
  return freshExamRows_()
    .filter(function (r) { return canSchool_(user, r['nama sekolah']); })
    .map(function (r) { return examOut_(r, true); });
}

function getExams_(req) {
  var u = auth_(req);
  return { exams: u.role === 'admin' ? listExamsForAdmin_(u) : listExamsForStudent_(u) };
}

function startExam_(req) {
  var u = auth_(req, 'siswa');
  var ex = examOut_(findExam_(req.examId), true);
  if (ex.school.toLowerCase() !== u.school.toLowerCase()) throw new Error('Ujian ini bukan untuk sekolah Anda.');
  if (ex.status !== 'Aktif') throw new Error('Ujian sedang tidak aktif.');

  var id = sid_(u.school, u.nomor, ex.id), now = Date.now();
  var s = loadSes_(id);
  if (!s) {
    s = { id: id, school: u.school, nomor: u.nomor, nama: u.nama, kelas: u.kelas, examId: ex.id, examName: ex.name,
      start: now, end: now + ex.duration * 60000, violations: 0, status: ST.RUN, last: now, finishedAt: '', note: 'Mulai', max: ex.max };
    appendRetry_(sesiSheet_(), [id, u.school, "'" + u.nomor, u.nama, u.kelas, ex.id, ex.name, now, s.end, 0, ST.RUN, now, '', 'Mulai']);
  } else {
    expire_(s);
  }
  if (s.status === ST.DONE) throw new Error('Anda sudah menyelesaikan ujian ini.');
  s.last = now; s.max = ex.max;
  putSes_(s);

  var out = { session: stateOf_(s, ex.max), exam: { id: ex.id, name: ex.name, school: ex.school, duration: ex.duration, max: ex.max } };
  if (s.status === ST.RUN) out.exam.link = toEmbed_(ex.link);
  return out;
}

function heartbeat_(req) {
  var u = auth_(req, 'siswa');
  var s = ownSes_(req, u);
  expire_(s);
  var now = Date.now();
  if (now - (s.last || 0) > 20000) { s.last = now; putSes_(s); }
  var out = { session: stateOf_(s) };
  if (s.status === ST.RUN && req.wantLink) out.link = toEmbed_(examOut_(findExam_(s.examId), true).link);
  return out;
}

function logViolation_(req) {
  var u = auth_(req, 'siswa');
  var s = ownSes_(req, u);
  expire_(s);
  var max = s.max || 3;
  try { max = examOut_(findExam_(s.examId), false).max; } catch (e) {}
  if (s.status === ST.RUN) {
    var n = s.violations + 1;
    s.violations = n; s.max = max; s.last = Date.now();
    if (n >= max) { s.status = ST.LOCK; s.note = 'Terkunci otomatis: melebihi batas pelanggaran'; }
    putSes_(s);
    try { appendRetry_(logSheet_(), [new Date(), s.id, s.school, "'" + s.nomor, s.nama, s.examId, s_(req.type), s_(req.detail).substring(0, 200), n]); } catch (e2) {}
  }
  return { session: stateOf_(s, max) };
}

function finishExam_(req) {
  var u = auth_(req, 'siswa');
  var s = ownSes_(req, u);
  if (s.status !== ST.DONE) {
    s.status = ST.DONE; s.finishedAt = nowStr_(); s.last = Date.now();
    s.note = s_(req.reason) === 'timeout' ? 'Waktu habis' : 'Selesai oleh siswa';
    putSes_(s);
  }
  return { session: stateOf_(s) };
}

function saveExam_(req) {
  var u = auth_(req, 'admin');
  var e = req.exam || {};
  var school = s_(e.school), name = s_(e.name);
  var max = parseInt(e.max, 10), dur = parseInt(e.duration, 10);
  var status = /^aktif$/i.test(s_(e.status)) ? 'Aktif' : 'Nonaktif';
  if (!school || !name) throw new Error('Nama sekolah dan nama ujian wajib diisi.');
  if (!canSchool_(u, school)) throw new Error('Anda tidak berhak mengelola sekolah "' + school + '".');
  if (!(max >= 1 && max <= 50)) throw new Error('Batas pelanggaran harus 1-50.');
  if (!(dur >= 1 && dur <= 600)) throw new Error('Durasi harus 1-600 menit.');
  var link = toEmbed_(e.link);

  var sh = sheet_(SHEET_NAMES.EXAM);
  var id = s_(e.id);
  if (id) {
    var f = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (!f) throw new Error('Ujian tidak ditemukan.');
    var old = freshExamRows_().filter(function (x) { return s_(x['id ujian']) === id; })[0];
    if (!old || !canSchool_(u, old['nama sekolah'])) throw new Error('Akses ditolak.');
    sh.getRange(f.getRow(), 1, 1, 7).setValues([[id, school, name, link, max, dur, status]]);
  } else {
    id = 'UJ-' + Date.now().toString(36).toUpperCase();
    sh.appendRow([id, school, name, link, max, dur, status]);
  }
  delBig_('tbl_exam');
  return { id: id, exams: listExamsForAdmin_(u) };
}

function deleteExam_(req) {
  var u = auth_(req, 'admin');
  var r = freshExamRows_().filter(function (x) { return s_(x['id ujian']) === s_(req.id); })[0];
  if (!r) throw new Error('Ujian tidak ditemukan.');
  if (!canSchool_(u, r['nama sekolah'])) throw new Error('Akses ditolak.');
  sheet_(SHEET_NAMES.EXAM).deleteRow(r._row);
  delBig_('tbl_exam');
  return { exams: listExamsForAdmin_(u) };
}

function monKey_(u, examId) { return 'mon_' + (isWildcard_(u.school) ? '*' : s_(u.school)) + '_' + (examId || 'all'); }
function clearMon_(u, examId) {
  delBig_(monKey_(u, '')); delBig_(monKey_(u, examId || ''));
  delBig_(monKey_({ school: '*' }, '')); delBig_(monKey_({ school: '*' }, examId || ''));
}

function getMonitor_(req) {
  var u = auth_(req, 'admin');
  var key = monKey_(u, s_(req.examId));
  var cached = getBig_(key);
  if (cached) return JSON.parse(cached);
  var now = Date.now();
  var list = allSesMerged_().filter(function (s) {
    return canSchool_(u, s.school) && (!req.examId || s.examId === req.examId);
  }).map(function (s) {
    var st = s.status, note = s.note;
    if (st === ST.RUN && now > s.end + GRACE_MS) { st = ST.DONE; note = 'Waktu habis'; }
    return {
      id: s.id, school: s.school, nomor: s.nomor, nama: s.nama, kelas: s.kelas, examId: s.examId, examName: s.examName,
      violations: s.violations, status: st, last: s.last, end: s.end, note: note,
      online: st === ST.RUN && now - (s.last || 0) < ONLINE_WINDOW_MS
    };
  });
  list.sort(function (a, b) { return b.last - a.last; });
  var out = { serverNow: now, total: list.length, sessions: list.slice(0, 800) };
  putBig_(key, JSON.stringify(out), TTL_MON);
  return out;
}

function resetSession_(req) {
  var u = auth_(req, 'admin');
  var who = u.nama + ' (' + u.username + ')';
  var now = Date.now();
  var extra = Math.max(0, Math.min(240, parseInt(req.extraMinutes, 10) || 0));
  var reset = req.resetViolations !== false;
  var maxCache = {};

  function maxOf(s) {
    if (maxCache[s.examId] == null) {
      try { maxCache[s.examId] = examOut_(findExam_(s.examId), false).max; } catch (e) { maxCache[s.examId] = s.max || 3; }
    }
    return maxCache[s.examId];
  }
  function unlockOne(s) {
    var newEnd = Math.max(s.end, now) + extra * 60000;
    if (newEnd - now < 30000) throw new Error('Waktu ujian ' + s.nama + ' sudah habis. Isi "Tambah waktu" minimal 1 menit.');
    s.status = ST.RUN; s.violations = reset ? 0 : Math.max(0, maxOf(s) - 1);
    s.last = now; s.end = newEnd; s.finishedAt = ''; s.note = 'Dibuka proktor: ' + who;
    putSes_(s);
  }

  if (req.op === 'unlockAll') {
    var ok = 0, skipped = 0;
    allSesMerged_().forEach(function (x) {
      if (x.status !== ST.LOCK || !canSchool_(u, x.school)) return;
      if (req.examId && x.examId !== req.examId) return;
      try { unlockOne(x); ok++; } catch (e) { skipped++; }
    });
    flushSessions_(); clearMon_(u, s_(req.examId));
    return { unlocked: ok, skipped: skipped };
  }

  var s = loadSes_(s_(req.sessionId));
  if (!s) throw new Error('Sesi tidak ditemukan.');
  if (!canSchool_(u, s.school)) throw new Error('Akses ditolak.');
  if (req.op === 'unlock') {
    unlockOne(s);
  } else if (req.op === 'finish') {
    s.status = ST.DONE; s.finishedAt = nowStr_(); s.note = 'Diselesaikan proktor: ' + who; putSes_(s);
  } else if (req.op === 'lock') {
    s.status = ST.LOCK; s.note = 'Dikunci proktor: ' + who; putSes_(s);
  } else {
    throw new Error('Operasi tidak dikenal.');
  }
  flushSessions_(); clearMon_(u, s.examId);
  return { ok: true };
}

function getLogs_(req) {
  var u = auth_(req, 'admin');
  var lg = logSheet_();
  var last = lg.getLastRow();
  if (last < 2) return { logs: [] };
  var start = Math.max(2, last - 299);
  var vals = lg.getRange(start, 1, last - start + 1, LOG_HEAD.length).getValues();
  var logs = vals.filter(function (v) { return canSchool_(u, v[2]) && (!req.sessionId || s_(v[1]) === s_(req.sessionId)); })
    .map(function (v) {
      return { time: v[0] instanceof Date ? v[0].getTime() : 0, sessionId: s_(v[1]), school: s_(v[2]), nomor: s_(v[3]), nama: s_(v[4]), examId: s_(v[5]), type: s_(v[6]), detail: s_(v[7]), n: v[8] };
    }).reverse();
  return { logs: logs };
}

/* ------------------- Jalankan SEKALI untuk inisialisasi ------------------- */

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'flushSessions' || f === 'warmCache') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('flushSessions').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('warmCache').timeBased().everyMinutes(5).create();
}

function setup() {
  sesiSheet_();
  logSheet_();
  try { sheet_(SHEET_NAMES.EXAM); } catch (e) { sheet_(SHEET_NAMES.EXAM, true, EXAM_HEAD); }
  secret_();
  installTriggers();
  warmCache();
  Logger.log('Setup selesai. Pemicu flushSessions (1 menit) dan warmCache (5 menit) terpasang.');
}