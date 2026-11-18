/**
 * =====================================================================
 *  CBT INFORMATIKA SMA NUGRAHA - BACKEND (Google Apps Script Web App)
 * =====================================================================
 *  Deploy: Extensions > Apps Script > tempel file ini > Deploy > New deployment
 *          Type: Web app | Execute as: Me | Who has access: Anyone
 *  Lalu salin URL "/exec" ke CONFIG.API_URL pada index.html.
 *
 *  Prinsip keamanan:
 *   - Password TIDAK PERNAH dikirim ke browser. Validasi dilakukan di server.
 *   - Link Google Form hanya dikirim saat siswa menekan "Mulai Ujian".
 *   - Hitungan pelanggaran, kunci, dan waktu berakhir dijaga SERVER (bukan browser),
 *     sehingga refresh / buka tab baru tidak mereset pelanggaran.
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
var TOKEN_TTL = 6 * 60 * 60;      // detik
var ONLINE_WINDOW_MS = 60 * 1000; // dianggap online jika heartbeat < 60 detik
var MAX_LOGIN_FAIL = 5;

/* ----------------------------- Router ------------------------------ */

var ACTIONS = {
  getSchools: getSchools_,
  loginSiswa: loginSiswa_,
  loginAdmin: loginAdmin_,
  getExams: getExams_,
  startExam: startExam_,
  heartbeat: heartbeat_,
  logViolation: logViolation_,
  finishExam: finishExam_,
  saveExam: saveExam_,
  deleteExam: deleteExam_,
  getMonitor: getMonitor_,
  resetSession: resetSession_,
  getLogs: getLogs_
};
var WRITE_ACTIONS = ['startExam', 'heartbeat', 'logViolation', 'finishExam', 'saveExam', 'deleteExam', 'resetSession'];

function doGet(e) {
  // Diagnosa: buka URL /exec?debug=1 untuk melihat nama sheet, header, dan jumlah baris
  if (e && e.parameter && e.parameter.debug === '1') {
    var info = {};
    try {
      SpreadsheetApp.openById(SS_ID).getSheets().forEach(function (sh) {
        var cols = Math.max(1, sh.getLastColumn());
        info[sh.getName()] = { baris_data: Math.max(0, sh.getLastRow() - 1), header: sh.getLastRow() > 0 ? sh.getRange(1, 1, 1, cols).getValues()[0] : [] };
      });
    } catch (err) { info = { error: String(err && err.message ? err.message : err) }; }
    return json_({ ok: true, debug: info });
  }
  return json_({ ok: true, service: 'CBT Informatika SMA Nugraha API', time: Date.now() });
}

function doPost(e) {
  var lock = null;
  try {
    var req = JSON.parse(e.postData.contents);
    var fn = ACTIONS[req.action];
    if (!fn) return json_({ ok: false, error: 'Aksi tidak dikenal.' });
    if (WRITE_ACTIONS.indexOf(req.action) > -1) {
      lock = LockService.getScriptLock();
      lock.waitLock(15000);
    }
    var res = fn(req) || {};
    res.ok = true;
    return json_(res);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (x) {} }
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ----------------------------- Helpers ----------------------------- */

function sheet_(names, create, head) {
  var ss = SpreadsheetApp.openById(SS_ID);
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

function auth_(req, role) {
  if (!req.token) throw new Error('Token tidak ada. Silakan login.');
  var raw = CacheService.getScriptCache().get('tok_' + req.token);
  if (!raw) throw new Error('Sesi login berakhir. Silakan login ulang.');
  var u = JSON.parse(raw);
  if (role && u.role !== role) throw new Error('Akses ditolak.');
  return u;
}

function issueToken_(user) {
  var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('tok_' + token, JSON.stringify(user), TOKEN_TTL);
  return token;
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

function examRows_() { return table_(sheet_(SHEET_NAMES.EXAM)); }

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

/* ---- Sesi ---- */
function sesiSheet_() { return sheet_(SHEET_NAMES.SESI, true, SESI_HEAD); }
function logSheet_() { return sheet_(SHEET_NAMES.LOG, true, LOG_HEAD); }

function rowToSesi_(row, v) {
  return {
    _row: row, id: s_(v[0]), school: s_(v[1]), nomor: s_(v[2]), nama: s_(v[3]), kelas: s_(v[4]),
    examId: s_(v[5]), examName: s_(v[6]), start: Number(v[7]), end: Number(v[8]),
    violations: Number(v[9]) || 0, status: s_(v[10]), last: Number(v[11]) || 0, finishedAt: s_(v[12]), note: s_(v[13])
  };
}
function findSesi_(sh, id) {
  if (!id) throw new Error('Sesi tidak valid.');
  var f = sh.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!f) throw new Error('Sesi tidak ditemukan.');
  var row = f.getRow();
  return rowToSesi_(row, sh.getRange(row, 1, 1, SESI_HEAD.length).getValues()[0]);
}
function allSesi_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, SESI_HEAD.length).getValues();
  return vals.map(function (v, i) { return rowToSesi_(i + 2, v); }).filter(function (x) { return x.id; });
}
function setSesi_(sh, s, patch) {
  Object.keys(patch).forEach(function (k) {
    var col = { violations: 10, status: 11, last: 12, finishedAt: 13, note: 14, end: 9 }[k];
    if (col) { sh.getRange(s._row, col).setValue(patch[k]); s[k] = patch[k]; }
  });
}
function stateOf_(s, max) {
  return { id: s.id, status: s.status, violations: s.violations, max: max, endsAt: s.end, serverNow: Date.now(), note: s.note };
}
/** Tutup sesi otomatis jika waktu habis. */
function expire_(sh, s) {
  if (s.status === ST.RUN && Date.now() > s.end + 3000) {
    setSesi_(sh, s, { status: ST.DONE, finishedAt: new Date(), note: 'Waktu habis' });
  }
  return s;
}
function ownSesi_(sh, req, user) {
  var s = findSesi_(sh, req.sessionId);
  if (s.nomor !== user.nomor || s.school !== user.school) throw new Error('Akses ditolak.');
  return s;
}

/* ----------------------------- Actions ----------------------------- */

function getSchools_() {
  var set = {};
  function add(names) {
    try {
      table_(sheet_(names)).forEach(function (r) {
        var n = s_(r['nama sekolah']);
        if (n && !isWildcard_(n)) set[n] = 1;
      });
    } catch (e) {}
  }
  add(SHEET_NAMES.SISWA);
  if (!Object.keys(set).length) { add(SHEET_NAMES.EXAM); add(SHEET_NAMES.GURU); }
  return { schools: Object.keys(set).sort() };
}

function loginSiswa_(req) {
  var school = s_(req.school), nomor = s_(req.nomor), pass = s_(req.password);
  if (!school || !nomor || !pass) throw new Error('Lengkapi sekolah, nomor peserta, dan password.');
  var th = throttle_('fail_s_' + school + '_' + nomor);
  var r = table_(sheet_(SHEET_NAMES.SISWA)).filter(function (x) {
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
  var r = table_(sheet_(SHEET_NAMES.GURU)).filter(function (x) { return s_(x['username / nip']) === username; })[0];
  if (!r || s_(r['password']) !== pass) { th.fail(); throw new Error('Username/NIP atau password salah.'); }
  th.ok();
  var user = { role: 'admin', school: s_(r['nama sekolah']), nama: s_(r['nama guru']), username: username };
  user.wildcard = isWildcard_(user.school);
  return { token: issueToken_(user), user: user, exams: listExamsForAdmin_(user), schools: getSchools_().schools };
}

function listExamsForStudent_(user) {
  var sh = sesiSheet_();
  var mine = allSesi_(sh).filter(function (s) { return s.nomor === user.nomor && s.school === user.school; });
  return examRows_()
    .filter(function (r) { return s_(r['nama sekolah']).toLowerCase() === user.school.toLowerCase() && /^aktif$/i.test(s_(r['status (aktif/nonaktif)'])); })
    .map(function (r) {
      var e = examOut_(r, false);
      var s = mine.filter(function (x) { return x.examId === e.id; })[0];
      e.myStatus = s ? s.status : null;
      return e;
    });
}
function listExamsForAdmin_(user) {
  return examRows_()
    .filter(function (r) { return canSchool_(user, r['nama sekolah']); })
    .map(function (r) { return examOut_(r, true); });
}

function getExams_(req) {
  var u = auth_(req);
  return { exams: u.role === 'admin' ? listExamsForAdmin_(u) : listExamsForStudent_(u) };
}

function startExam_(req) {
  var u = auth_(req, 'siswa');
  var r = findExam_(req.examId);
  var ex = examOut_(r, true);
  if (ex.school.toLowerCase() !== u.school.toLowerCase()) throw new Error('Ujian ini bukan untuk sekolah Anda.');
  if (ex.status !== 'Aktif') throw new Error('Ujian sedang tidak aktif.');

  var sh = sesiSheet_();
  var s = allSesi_(sh).filter(function (x) { return x.nomor === u.nomor && x.school === u.school && x.examId === ex.id; })[0];
  var now = Date.now();

  if (!s) {
    var id = Utilities.getUuid().substring(0, 8);
    var end = now + ex.duration * 60000;
    sh.getRange(sh.getLastRow() + 1, 1, 1, SESI_HEAD.length).setValues([[id, u.school, "'" + u.nomor, u.nama, u.kelas, ex.id, ex.name, now, end, 0, ST.RUN, now, '', 'Mulai']]);
    s = findSesi_(sh, id);
  } else {
    expire_(sh, s);
  }
  if (s.status === ST.DONE) throw new Error('Anda sudah menyelesaikan ujian ini.');
  setSesi_(sh, s, { last: now });

  var out = { session: stateOf_(s, ex.max), exam: { id: ex.id, name: ex.name, school: ex.school, duration: ex.duration, max: ex.max } };
  if (s.status === ST.RUN) out.exam.link = toEmbed_(ex.link); // link hanya diberikan jika tidak terkunci
  return out;
}

function heartbeat_(req) {
  var u = auth_(req, 'siswa');
  var sh = sesiSheet_();
  var s = ownSesi_(sh, req, u);
  expire_(sh, s);
  setSesi_(sh, s, { last: Date.now() });
  var max = examOut_(findExam_(s.examId), false).max;
  var out = { session: stateOf_(s, max) };
  if (s.status === ST.RUN && req.wantLink) out.link = toEmbed_(examOut_(findExam_(s.examId), true).link);
  return out;
}

function logViolation_(req) {
  var u = auth_(req, 'siswa');
  var sh = sesiSheet_();
  var s = ownSesi_(sh, req, u);
  expire_(sh, s);
  var ex = examOut_(findExam_(s.examId), false);
  if (s.status === ST.RUN) {
    var n = s.violations + 1;
    var patch = { violations: n, last: Date.now() };
    if (n >= ex.max) { patch.status = ST.LOCK; patch.note = 'Terkunci otomatis: melebihi batas pelanggaran'; }
    setSesi_(sh, s, patch);
    var lg = logSheet_();
    lg.appendRow([new Date(), s.id, s.school, "'" + s.nomor, s.nama, s.examId, s_(req.type), s_(req.detail).substring(0, 200), n]);
  }
  return { session: stateOf_(s, ex.max) };
}

function finishExam_(req) {
  var u = auth_(req, 'siswa');
  var sh = sesiSheet_();
  var s = ownSesi_(sh, req, u);
  if (s.status !== ST.DONE) {
    setSesi_(sh, s, { status: ST.DONE, finishedAt: new Date(), last: Date.now(), note: s_(req.reason) === 'timeout' ? 'Waktu habis' : 'Selesai oleh siswa' });
  }
  return { session: stateOf_(s, examOut_(findExam_(s.examId), false).max) };
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
    var old = findExam_(id);
    if (!canSchool_(u, old['nama sekolah'])) throw new Error('Akses ditolak.');
    sh.getRange(f.getRow(), 1, 1, 7).setValues([[id, school, name, link, max, dur, status]]);
  } else {
    id = 'UJ-' + Date.now().toString(36).toUpperCase();
    sh.appendRow([id, school, name, link, max, dur, status]);
  }
  return { id: id, exams: listExamsForAdmin_(u) };
}

function deleteExam_(req) {
  var u = auth_(req, 'admin');
  var r = findExam_(req.id);
  if (!canSchool_(u, r['nama sekolah'])) throw new Error('Akses ditolak.');
  sheet_(SHEET_NAMES.EXAM).deleteRow(r._row);
  return { exams: listExamsForAdmin_(u) };
}

function getMonitor_(req) {
  var u = auth_(req, 'admin');
  var now = Date.now();
  var sh = sesiSheet_();
  var list = allSesi_(sh).filter(function (s) { return canSchool_(u, s.school); });
  if (req.examId) list = list.filter(function (s) { return s.examId === req.examId; });
  list.forEach(function (s) { if (s.status === ST.RUN && now > s.end + 3000) expire_(sh, s); });
  list.sort(function (a, b) { return b.last - a.last; });
  return {
    serverNow: now,
    sessions: list.slice(0, 400).map(function (s) {
      return {
        id: s.id, school: s.school, nomor: s.nomor, nama: s.nama, kelas: s.kelas, examId: s.examId, examName: s.examName,
        violations: s.violations, status: s.status, last: s.last, end: s.end, note: s.note,
        online: s.status === ST.RUN && now - s.last < ONLINE_WINDOW_MS
      };
    })
  };
}

function resetSession_(req) {
  var u = auth_(req, 'admin');
  var sh = sesiSheet_();
  var who = u.nama + ' (' + u.username + ')';
  var now = Date.now();
  var extra = Math.max(0, Math.min(240, parseInt(req.extraMinutes, 10) || 0));
  var reset = req.resetViolations !== false;

  // Buka kunci satu sesi: status kembali "Sedang Mengerjakan", waktu bisa ditambah
  function unlockOne(s) {
    var newEnd = Math.max(s.end, now) + extra * 60000;
    if (newEnd - now < 30000) {
      throw new Error('Waktu ujian ' + s.nama + ' sudah habis. Isi "Tambah waktu" minimal 1 menit.');
    }
    var max = examOut_(findExam_(s.examId), false).max;
    setSesi_(sh, s, {
      status: ST.RUN,
      violations: reset ? 0 : Math.max(0, max - 1),
      last: now,
      end: newEnd,
      finishedAt: '',
      note: 'Dibuka proktor: ' + who
    });
  }

  if (req.op === 'unlockAll') {
    var ok = 0, skipped = 0;
    allSesi_(sh).forEach(function (x) {
      if (x.status !== ST.LOCK || !canSchool_(u, x.school)) return;
      if (req.examId && x.examId !== req.examId) return;
      try { unlockOne(x); ok++; } catch (e) { skipped++; }
    });
    return { unlocked: ok, skipped: skipped };
  }

  var s = findSesi_(sh, req.sessionId);
  if (!canSchool_(u, s.school)) throw new Error('Akses ditolak.');
  if (req.op === 'unlock') {
    unlockOne(s);
  } else if (req.op === 'finish') {
    setSesi_(sh, s, { status: ST.DONE, finishedAt: new Date(), note: 'Diselesaikan proktor: ' + who });
  } else if (req.op === 'lock') {
    setSesi_(sh, s, { status: ST.LOCK, note: 'Dikunci proktor: ' + who });
  } else {
    throw new Error('Operasi tidak dikenal.');
  }
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
function setup() {
  sesiSheet_();
  logSheet_();
  try { sheet_(SHEET_NAMES.EXAM); } catch (e) { sheet_(SHEET_NAMES.EXAM, true, EXAM_HEAD); }
  Logger.log('Setup selesai. Pastikan sheet Siswa & Guru / Admin sudah terisi.');
}
