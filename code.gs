/**
 * dikiLabs Learning Hub — Google Apps Script Web App
 * ---------------------------------------------------
 * Struktur:
 *   index.html         -> landing page hub (pilih SkillCheck / Excel Logic Academy)
 *   logicacademy.html  -> landing page Excel Logic Academy (dulu index.html)
 *   soal.html          -> modul interaktif ELA (peta level, lab spreadsheet, kuis, sertifikat)
 *   skillcheck.html    -> modul SkillCheck (assessment interaktif, seluruh bank soal &
 *                          aturan skoring di-hardcode di dalam JS file tersebut — TIDAK
 *                          memakai Sheet/DB. Sheet di bawah hanya dipakai untuk mencatat
 *                          RINGKASAN hasil, bukan sebagai sumber soal.)
 *   practice.html      -> modul Practice (lesson + micro-challenge per skill, Data
 *                          Detective Mode, Mini Project checklist). Sengaja dipisah
 *                          dari skillcheck.html supaya menambah/mengubah materi latihan
 *                          cukup edit file ini saja. Progres disimpan di localStorage
 *                          browser pengguna (bukan Sheet/DB). Menerima query ?skill=<id>
 *                          untuk deep-link dari rekomendasi SkillCheck, misal:
 *                          ?page=practice&skill=sql
 *   commstyle.html     -> modul Communication Style Assessment (24 soal situasional,
 *                          4 gaya: RED/YELLOW/GREEN/BLUE). Seluruh bank soal, kunci
 *                          warna, dan aturan skoring di-hardcode di JS file tersebut,
 *                          TIDAK memakai Sheet/DB sebagai sumber soal.
 *
 * doGet menentukan file mana yang ditampilkan lewat parameter URL ?page=
 *   (kosong)          -> index.html (hub)
 *   ?page=logicacademy -> logicacademy.html
 *   ?page=soal         -> soal.html
 *   ?page=skillcheck   -> skillcheck.html
 *   ?page=practice      -> practice.html
 *   ?page=commstyle     -> commstyle.html
 *   page lain/tidak dikenal -> tetap index.html (hub)
 */

/* Link donasi/tip kopi (Saweria, Trakteer, dll) yang ditampilkan opsional
   di layar sertifikat soal.html. Kosongkan ('') untuk menyembunyikan
   tombolnya. Disimpan di sini (server) supaya gampang diganti tanpa
   menyentuh file HTML. */
var COFFEE_TIP_URL = 'https://tiptap.gg/dikilabs';

/* Peta page -> nama file HtmlService. Tambahkan entri baru di sini kalau
   nanti ada halaman lain (mis. skillcheck-assessment, skillcheck-hasil, dst). */
var PAGE_FILES = {
  index: 'index',
  logicacademy: 'logicacademy',
  soal: 'soal',
  skillcheck: 'skillcheck',
  practice: 'practice',
  commstyle: 'commstyle',
  leader: 'leader'
};

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : 'index';
  var fileName = PAGE_FILES[page] || 'index';

  var template = HtmlService.createTemplateFromFile(fileName);
  // scriptUrl dipakai di dalam HTML (lewat <?= scriptUrl ?>) untuk membuat
  // tautan antar halaman tetap valid setelah aplikasi di-deploy.
  template.scriptUrl = ScriptApp.getService().getUrl();
  // coffeeTipUrl dipakai di soal.html (lewat <?= coffeeTipUrl ?>) untuk
  // menampilkan tombol tip kopi opsional di layar sertifikat.
  template.coffeeTipUrl = COFFEE_TIP_URL;

var titles = {
  index: 'Jadi Bisa',
  logicacademy: 'Excel Logic Academy — Belajar Excel dari Basic sampai Advanced',
  soal: 'Excel Skill Test & Practice — Excel Logic Academy',
  skillcheck: 'Skill Assessment — Kenali Skill & Skill Gap Kamu | SkillCheck',
  practice: 'Skill Practice — Latihan Excel, SQL & Data Skills | SkillCheck',
  commstyle: 'Communication Style Assessment — Kenali Gaya Komunikasimu | SkillCheck',
  leader: 'Are You a Leader? — Kenali Gaya Kepemimpinanmu | SkillCheck'
};

  return template.evaluate()
    .setTitle(titles[fileName] || 'dikiLabs')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper opsional — sediakan kalau nanti CSS/JS ingin dipecah jadi file
 * terpisah (mis. "styles.html") dan disisipkan lewat <?!= include('styles'); ?>
 * Tidak dipakai secara default karena semua file HTML sudah mandiri.
 */
function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

/* ============================= PENYIMPANAN HASIL — EXCEL LOGIC ACADEMY =============================
 * Dipanggil dari soal.html (lewat google.script.run) setiap kali seseorang
 * membuat sertifikat di layar terakhir. Menulis satu baris berisi:
 *   datetime | nama | nilaifinal | pesan | tips
 * ke Google Sheet dengan ID di bawah. Sheet (tab) dibuat otomatis kalau
 * belum ada, lengkap dengan header di baris pertama.
 * ==================================================================================================== */

var RESULT_SPREADSHEET_ID = '18dsEgSIO9_tRbbvADdo8zkFh-kewAvWx2g4NYRGr_Dk';
var RESULT_SHEET_NAME = 'Hasil Excel Logic Academy';
var RESULT_HEADERS = ['datetime', 'nama', 'nilaifinal', 'pesan', 'tips'];

/**
 * Mengambil sheet tujuan di dalam spreadsheet RESULT_SPREADSHEET_ID.
 * Kalau sheet-nya belum ada, buat baru sekaligus tulis header di baris 1.
 * Kalau sheet sudah ada tapi baris headernya kosong/tidak lengkap, header
 * akan ditulis ulang supaya urutan kolom tetap konsisten.
 */
function getOrCreateResultSheet_() {
  return getOrCreateSheet_(RESULT_SHEET_NAME, RESULT_HEADERS);
}

/**
 * Fungsi yang dipanggil dari client (soal.html) lewat:
 *   google.script.run.saveCertificateResult({ nama, nilaifinal, pesan, tips })
 *
 * data: {
 *   nama:       string  -> nama yang diketik di layar sertifikat
 *   nilaifinal: number  -> skor akhir gabungan seluruh level (0-100)
 *   pesan:      string  -> isi jurnal refleksi dari Level terakhir
 *   tips:       string  -> tips otomatis berdasarkan skor akhir
 * }
 * Mengembalikan { ok:true } jika berhasil, atau { ok:false, error:'...' } jika gagal
 * (client tidak boleh menganggap kegagalan ini sebagai error fatal — lihat soal.html).
 */
function saveCertificateResult(data) {
  try {
    data = data || {};
    var sheet = getOrCreateResultSheet_();

    var nama = data.nama ? String(data.nama).trim() : '';
    var nilaifinal = (typeof data.nilaifinal === 'number' && !isNaN(data.nilaifinal)) ? data.nilaifinal : '';
    var pesan = data.pesan ? String(data.pesan).trim() : '';
    var tips = data.tips ? String(data.tips).trim() : '';

    sheet.appendRow([new Date(), nama, nilaifinal, pesan, tips]);

    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

/* ============================= PENYIMPANAN HASIL — SKILLCHECK =============================
 * Dipanggil dari skillcheck.html (lewat google.script.run) setiap kali
 * seseorang menyelesaikan assessment. Bank soal, skoring, skill profile,
 * skill gap, dan recommendation SEMUA dihitung di client (hardcoded di JS
 * skillcheck.html) — fungsi ini hanya mencatat RINGKASAN hasil akhir ke
 * Google Sheet untuk keperluan analytics admin (lihat PRD §23 Analytics),
 * bukan sebagai sumber data soal maupun logika penilaian.
 *
 * Menulis satu baris:
 *   datetime | nama | status | careerInterest | experience | overallScore | skillScoresJson
 * ============================================================================================ */

var SKILLCHECK_SHEET_NAME = 'Hasil SkillCheck';
var SKILLCHECK_HEADERS = ['datetime', 'nama', 'status', 'careerInterest', 'experience', 'overallScore', 'skillScoresJson'];

function getOrCreateSkillCheckSheet_() {
  return getOrCreateSheet_(SKILLCHECK_SHEET_NAME, SKILLCHECK_HEADERS);
}

/**
 * Dipanggil dari client:
 *   google.script.run.saveSkillCheckResult({
 *     nama, status, careerInterest, experience, overallScore, skillScores (JSON string)
 *   })
 *
 * Mengembalikan { ok:true } jika berhasil, atau { ok:false, error:'...' } jika gagal.
 * Kegagalan di sini TIDAK boleh dianggap fatal oleh client — hasil assessment
 * sudah selesai dihitung & ditampilkan sepenuhnya di browser sebelum fungsi
 * ini dipanggil (fire-and-forget, lihat skillcheck.html: sendResultToBackend).
 */
function saveSkillCheckResult(data) {
  try {
    data = data || {};
    var sheet = getOrCreateSkillCheckSheet_();

    var nama = data.nama ? String(data.nama).trim() : '';
    var status = data.status ? String(data.status).trim() : '';
    var careerInterest = data.careerInterest ? String(data.careerInterest).trim() : '';
    var experience = data.experience ? String(data.experience).trim() : '';
    var overallScore = (typeof data.overallScore === 'number' && !isNaN(data.overallScore)) ? data.overallScore : '';
    var skillScoresJson = data.skillScores ? String(data.skillScores) : '';

    sheet.appendRow([new Date(), nama, status, careerInterest, experience, overallScore, skillScoresJson]);

    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

/* ============================= PENYIMPANAN HASIL — COMMUNICATION STYLE =============================
 * Dipanggil dari commstyle.html (lewat google.script.run) setiap kali
 * seseorang menyelesaikan 24 pertanyaan. Bank soal, kunci warna, dan
 * scoring SEMUA dihitung di client (hardcoded di JS commstyle.html),
 * fungsi ini hanya mencatat RINGKASAN skor akhir per warna ke Google
 * Sheet untuk keperluan analytics admin, bukan sebagai sumber soal.
 *
 * Menulis satu baris:
 *   datetime | nama | red | yellow | green | blue | primary | secondary
 * ==================================================================================================== */

var COMMSTYLE_SHEET_NAME = 'Hasil Communication Style';
var COMMSTYLE_HEADERS = ['datetime', 'nama', 'red', 'yellow', 'green', 'blue', 'primary', 'secondary'];

function getOrCreateCommStyleSheet_() {
  return getOrCreateSheet_(COMMSTYLE_SHEET_NAME, COMMSTYLE_HEADERS);
}

/**
 * Dipanggil dari client:
 *   google.script.run.saveCommStyleResult({
 *     nama, red, yellow, green, blue, primary, secondary
 *   })
 *
 * Mengembalikan { ok:true } jika berhasil, atau { ok:false, error:'...' } jika gagal.
 * Kegagalan di sini TIDAK boleh dianggap fatal oleh client — hasil sudah
 * selesai dihitung & ditampilkan sepenuhnya di browser sebelum fungsi ini
 * dipanggil (fire-and-forget, lihat commstyle.html: sendResultToBackend).
 */
function saveCommStyleResult(data) {
  try {
    data = data || {};
    var sheet = getOrCreateCommStyleSheet_();

    var nama = data.nama ? String(data.nama).trim() : '';
    var red = (typeof data.red === 'number' && !isNaN(data.red)) ? data.red : '';
    var yellow = (typeof data.yellow === 'number' && !isNaN(data.yellow)) ? data.yellow : '';
    var green = (typeof data.green === 'number' && !isNaN(data.green)) ? data.green : '';
    var blue = (typeof data.blue === 'number' && !isNaN(data.blue)) ? data.blue : '';
    var primary = data.primary ? String(data.primary).trim() : '';
    var secondary = data.secondary ? String(data.secondary).trim() : '';

    sheet.appendRow([new Date(), nama, red, yellow, green, blue, primary, secondary]);

    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

/* ============================= HELPER SHEET BERSAMA =============================
 * Dipakai oleh ketiga fitur penyimpanan hasil di atas (ELA, SkillCheck &
 * Communication Style) supaya logika "buat sheet kalau belum ada + pastikan
 * header konsisten" tidak diduplikasi.
 * ================================================================================= */
function getOrCreateSheet_(sheetName, headers) {
  var ss = SpreadsheetApp.openById(RESULT_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var headerOk = headers.every(function (h, i) { return firstRow[i] === h; });
  if (!headerOk) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  return sheet;
}