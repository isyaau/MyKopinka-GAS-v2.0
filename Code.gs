// =====================================================
// KONFIGURASI DATABASE PUSAT
// =====================================================
var SHEET_USERS = "Users";
var SHEET_VOUCHERS = "Voucher";
var SHEET_LAPORAN = "Mutasi";
var SHEET_POTONGAN = "Potongan";
var SHEET_NOTIF = "Notifikasi"; 
var SHEET_KREDIT_TOKO = "Piutang";
var SHEET_BAYAR_PIUTANG = "PembayaranPiutang";
var SHEET_UNIT = "Unit";
var SHEET_SETTINGS = "Settings";

var FOLDER_ID = "18kCGyT26jGqXshHhaARqi9rmun_rC-vB"; 
var FOLDER_NOTIF_ID = "1UmdX6k_bTjAq6Ag9x8YCE2E-cUF1KxdT"; 
var FOLDER_KREDIT_TOKO_ID = "1zFKB6N4vjpJtnY1k_G6c1hutIhshLYyW";

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('Portal Kopinka v2.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) { 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

// =====================================================
// FUNGSI UTILITAS FORMAT DATA
// =====================================================
function _formatRp(angka) {
  var num = Number(angka) || 0;
  return Math.floor(num).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function _parseDate(dateVal) {
  if (!dateVal || dateVal === "") return null;
  if (Object.prototype.toString.call(dateVal) === '[object Date]') return isNaN(dateVal.getTime()) ? null : new Date(dateVal.getTime());
  var d = null;
  if (typeof dateVal === 'string') {
    var s = dateVal.trim();
    // Format ISO 8601: yyyy-MM-ddTHH:mm:ss(.sss)Z / dengan offset / yyyy-MM-dd
    var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?(?:\.(\d+))?)?(Z|[+-]\d{2}:?\d{2})?$/);
    if (iso) {
      var base = Date.UTC(+iso[1], +iso[2] - 1, +iso[3], +(iso[4] || 0), +(iso[5] || 0), +(iso[6] || 0), +(iso[7] || 0));
      if (iso[8] === 'Z') d = new Date(base);
      else if (/^[+-]/.test(iso[8])) {
        var off = iso[8].replace(':', '');
        var oh = +off.slice(1, 3), om = +off.slice(3, 5);
        var msOff = (oh * 60 + om) * 60000;
        d = new Date(base - (off[0] === '+' ? msOff : -msOff));
      } else d = new Date(base);
    } else {
      var parts = s.split(/[-/]/);
      if (parts.length === 3 && /^\d+$/.test(parts[0]) && /^\d+$/.test(parts[1]) && /^\d+$/.test(parts[2])) {
        if (parts[2].length === 4) d = new Date(+parts[2], +parts[1] - 1, +parts[0]);
        else if (parts[0].length === 4) d = new Date(+parts[0], +parts[1] - 1, +parts[2]);
      } else {
        d = new Date(s);
      }
    }
  } else {
    d = new Date(dateVal);
  }
  return isNaN(d.getTime()) ? null : d;
}

function _formatTglIndo(tgl) {
  if (!tgl || tgl === "" || tgl === "-") return "-";
  if (Object.prototype.toString.call(tgl) === '[object Date]') return Utilities.formatDate(tgl, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return String(tgl);
}

function _getSetting(key) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getSheetByName(SHEET_SETTINGS);
    if (!s) return null;
    var data = s.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === key) return data[i][1];
    }
  } catch (e) { return null; }
  return null;
}

// =====================================================
// CACHE UTIL (CacheService - chunked, karena max 100KB/entry)
// =====================================================
var _CACHE_CHUNK = 90000; // 90KB per entry
var _CACHE_TTL_PIUTANG = 180; // 3 menit
var _CACHE_TTL_USERS = 600;   // 10 menit
var _CACHE_TTL_PAY = 180;     // 3 menit

function _cacheSet(key, obj, ttl) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(obj);
    var n = Math.max(1, Math.ceil(json.length / _CACHE_CHUNK));
    cache.put(key + '_meta', JSON.stringify({ n: n }), ttl || 180);
    for (var i = 0; i < n; i++) {
      cache.put(key + '_' + i, json.substr(i * _CACHE_CHUNK, _CACHE_CHUNK), ttl || 180);
    }
  } catch (e) {}
}

function _cacheGet(key) {
  try {
    var cache = CacheService.getScriptCache();
    var metaJson = cache.get(key + '_meta');
    if (!metaJson) return null;
    var meta = JSON.parse(metaJson);
    var parts = [];
    for (var i = 0; i < meta.n; i++) {
      var p = cache.get(key + '_' + i);
      if (p === null || p === undefined) return null;
      parts.push(p);
    }
    return JSON.parse(parts.join(''));
  } catch (e) { return null; }
}

function _cacheRemove(key) {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove(key + '_meta');
    for (var i = 0; i < 200; i++) {
      if (cache.get(key + '_' + i) === null) break;
      cache.remove(key + '_' + i);
    }
  } catch (e) {}
}

function _invalidateKreditCaches() {
  _cacheRemove('kp_piutang');
  _cacheRemove('kp_paymap');
  _cacheRemove('kp_payrows');
  _cacheRemove('kp_users');
}

// Data mentah sheet Piutang (semua baris)
function _getCachedPiutangData() {
  var cached = _cacheGet('kp_piutang');
  if (cached) return cached;
  var data = _ensureKreditTokoSheet().getDataRange().getValues();
  _cacheSet('kp_piutang', data, _CACHE_TTL_PIUTANG);
  return data;
}

// Map pembayaran per anggota (dari sheet PembayaranPiutang)
function _getCachedMemberPaymentMap() {
  var cached = _cacheGet('kp_paymap');
  if (cached) return cached;
  var map = _getMemberPaymentMap();
  _cacheSet('kp_paymap', map, _CACHE_TTL_PAY);
  return map;
}

// Baris pembayaran mentah (dengan tanggal) untuk simulasi biaya jasa bulanan
function _getCachedPaymentRows() {
  var cached = _cacheGet('kp_payrows');
  if (cached) return cached;
  var rows = [];
  try {
    var s = _ensurePembayaranPiutangSheet();
    var data = s.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var noAng = String(data[i][1] || '').replace(/'/g, '').trim().toLowerCase();
      var jml = Number(data[i][2]) || 0;
      var tglObj = _parseDate(data[i][0]);
      if (!noAng || jml <= 0 || !tglObj) continue;
      rows.push({ noAng: noAng, t: tglObj.getTime(), jml: jml });
    }
  } catch (e) {}
  _cacheSet('kp_payrows', rows, _CACHE_TTL_PAY);
  return rows;
}

// Data mentah sheet Users
function _getCachedUsersData() {
  var cached = _cacheGet('kp_users');
  if (cached) return cached;
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS).getDataRange().getValues();
  _cacheSet('kp_users', data, _CACHE_TTL_USERS);
  return data;
}

// _getRed = _getSetting; sinkronkan cache saat setting berubah
function _getSettingCached(key, ttl) {
  var cached = _cacheGet('kp_set_' + key);
  if (cached !== null && cached !== undefined) return cached;
  var v = _getSetting(key);
  _cacheSet('kp_set_' + key, v === null || v === undefined ? '' : v, ttl || 300);
  return v;
}

// Fitur limit kredit aktif/nonaktif. Default aktif agar perilaku lama tetap berjalan.
function _isLimitKreditAktif() {
  var v = _getSettingCached('limit_kredit_aktif');
  if (v === null || v === undefined || v === '') return true;
  var s = String(v).toLowerCase();
  return s === 'true' || s === '1' || s === 'aktif' || s === 'yes' || s === 'y';
}

function _getGlobalLimitCached() {
  if (!_isLimitKreditAktif()) return -1;
  return Number(_getSettingCached('limit_piutang')) || 0;
}

function _clearSettingCache(key) {
  _cacheRemove('kp_set_' + (key || ''));
}

function _isKreditTokoBlocked(noAnggota) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sUsers = ss.getSheetByName(SHEET_USERS);
    var uData = sUsers.getDataRange().getValues();
    var q = String(noAnggota).trim().toLowerCase();
    for (var i = 1; i < uData.length; i++) {
      if (String(uData[i][0]).trim().toLowerCase() === q) {
        var blokir = String(uData[i][23] || '').trim().toLowerCase();
        return blokir === 'y' || blokir === 'yes' || blokir === '1' || blokir === 'blokir';
      }
    }
  } catch(e) {}
  return false;
}

function _isPiutangBlocked(noAnggota) {
  return _isKreditTokoBlocked(noAnggota);
}

function _getKreditTokoBulanIni(noAnggota) {
  try {
    var data = _getCachedPiutangData();
    var total = 0;
    var now = new Date();
    var curMonth = now.getMonth();
    var curYear = now.getFullYear();
    for (var i = 1; i < data.length; i++) {
      var rowAnggota = String(data[i][6]).replace(/'/g,'').trim();
      if (rowAnggota === String(noAnggota).trim()) {
        var tgl = _parseDate(data[i][0]);
        if (tgl && tgl.getMonth() === curMonth && tgl.getFullYear() === curYear) {
          total += (Number(data[i][5]) || 0);
        }
      }
    }
    return total;
  } catch (e) { return 0; }
}

function _ensurePembayaranPiutangSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_BAYAR_PIUTANG);
  if (!s) {
    s = ss.insertSheet(SHEET_BAYAR_PIUTANG);
    s.getRange(1, 1, 1, 6).setValues([['Tgl Bayar', 'No Anggota', 'Jumlah Bayar', 'Nota/Keterangan', 'Metode', 'Status']]);
  }
  return s;
}

// Mapping No Anggota -> total pembayaran (akumulatif dari semua baris PembayaranPiutang)
function _getMemberPaymentMap() {
  var map = {};
  try {
    var s = _ensurePembayaranPiutangSheet();
    var data = s.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var noAng = String(data[i][1] || "").replace(/'/g, "").trim().toLowerCase();
      var jumlah = Number(data[i][2]) || 0;
      if (!noAng) continue;
      if (!map[noAng]) map[noAng] = 0;
      map[noAng] += jumlah;
    }
  } catch (e) {}
  return map;
}

// Batas kredit bulan ini = limit global - sisa tagihan (outstanding) dari periode sebelumnya.
// Sisa tagihan = total (Nilai - Pembayaran) untuk semua kredit bulan LALU (bulan sebelum bulan berjalan).
// Tagihan lama yang belum lunas terus terakumulasi dan mengurangi limit sampai limit = 0 (diblokir).
function _getBatasKreditAnggota(noAnggota) {
  var global = _getGlobalLimitCached();
  var outstanding = 0;
  var now = new Date();
  var curMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  var q = String(noAnggota).trim().toLowerCase();
  try {
    var memberBayar = _getCachedMemberPaymentMap();
    var totalDibayar = memberBayar[q] || 0;
    var totalPiutang = 0;
    var data = _getCachedPiutangData();
    for (var i = 1; i < data.length; i++) {
      var nM = String(data[i][6]).replace(/'/g, '').trim().toLowerCase();
      if (nM !== q) continue;
      var tgl = _parseDate(data[i][0]);
      if (!tgl) continue;
      if (tgl.getTime() >= curMonthStart.getTime()) continue;
      totalPiutang += Number(data[i][5]) || 0;
    }
    outstanding = totalPiutang - totalDibayar;
    if (outstanding < 0) outstanding = 0;
  } catch (e) {}
  if (!_isLimitKreditAktif()) {
    return { global: -1, outstanding: outstanding, effective: -1 };
  }
  var effective = global - outstanding;
  if (effective < 0) effective = 0;
  return { global: global, outstanding: outstanding, effective: effective };
}

// Hitung limit kredit untuk SEMUA anggota dalam 1 kali scan sheet (hemat request ke Google Sheets).
function _getBatasKreditAllMembers(dataPiutangOpt) {
  var global = _getGlobalLimitCached();
  var now = new Date();
  var curMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);

  var piutangLalu = {};
  try {
    var data = dataPiutangOpt || _getCachedPiutangData();
    for (var i = 1; i < data.length; i++) {
      var q = String(data[i][6]).replace(/'/g, '').trim().toLowerCase();
      if (!q) continue;
      var tgl = _parseDate(data[i][0]);
      if (!tgl) continue;
      if (tgl.getTime() < curMonthStart.getTime()) {
        if (!piutangLalu[q]) piutangLalu[q] = 0;
        piutangLalu[q] += Number(data[i][5]) || 0;
      }
    }
  } catch (e) {}

  var memberBayar = _getCachedMemberPaymentMap();
  if (!_isLimitKreditAktif()) {
    var outLimits = {};
    for (var kk in piutangLalu) {
      var totPay = memberBayar[kk] || 0;
      var out = piutangLalu[kk] - totPay;
      if (out < 0) out = 0;
      outLimits[kk] = { global: -1, outstanding: out, effective: -1 };
    }
    return outLimits;
  }
  var limits = {};
  for (var k in piutangLalu) {
    var totalDibayar = memberBayar[k] || 0;
    var outstanding = piutangLalu[k] - totalDibayar;
    if (outstanding < 0) outstanding = 0;
    var effective = global - outstanding;
    if (effective < 0) effective = 0;
    limits[k] = { global: global, outstanding: outstanding, effective: effective };
  }
  return limits;
}

function importPembayaranPiutang(dataArray) {
  try {
    if (!dataArray || dataArray.length < 2) return { status: 'error', msg: 'Data kosong.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Kumpulkan data piutang per No Anggota: { noAng -> { totalNilai, rows[] } }
    var sPi = _ensureKreditTokoSheet();
    var dataPi = sPi.getDataRange().getValues();
    var piutangByMember = {};
    for (var i = 1; i < dataPi.length; i++) {
      var nA = String(dataPi[i][6]).replace(/'/g, '').trim().toLowerCase();
      if (!nA) continue;
      var nilai = Number(dataPi[i][5]) || 0;
      if (!piutangByMember[nA]) piutangByMember[nA] = { totalNilai: 0 };
      piutangByMember[nA].totalNilai += nilai;
    }

    // Kumpulkan daftar No Anggota valid dari sheet Users
    var sUs = ss.getSheetByName(SHEET_USERS);
    var dataUs = sUs.getDataRange().getValues();
    var memberMap = {};
    for (var j = 1; j < dataUs.length; j++) {
      var mn = String(dataUs[j][0]).trim();
      if (mn) memberMap[mn.toLowerCase()] = true;
    }

    // Hitung total pembayaran per No Anggota yang sudah ada di PembayaranPiutang
    var memberBayar = _getMemberPaymentMap();

    // Deteksi duplikat yang sudah ada
    var sBayar = _ensurePembayaranPiutangSheet();
    var exData = sBayar.getDataRange().getValues();
    var existKeys = {};
    for (var k = 1; k < exData.length; k++) {
      var kAng = String(exData[k][1] || '').replace(/'/g, '').trim().toLowerCase();
      var kJml = Number(exData[k][2]) || 0;
      var kTgl = String(exData[k][0] || '').trim();
      if (kAng) existKeys[kAng + '|' + kJml + '|' + kTgl] = true;
    }

    var toInsert = [];
    var err = [];
    var skipped = 0;
    var tz = Session.getScriptTimeZone();

    for (var r = 1; r < dataArray.length; r++) {
      var row = dataArray[r];
      var tgl = String(row[0] || '').trim();
      var noAng = String(row[1] || '').trim();
      if (noAng.charAt(0) === "'") noAng = noAng.substring(1);
      var jumlah = Number(String(row[2] || '').replace(/[^0-9\-\\.]/g, '')) || 0;
      var nota = String(row[3] || '').trim() || '-';
      var metode = String(row[4] || '').trim() || '-';

      if (noAng === '' && jumlah <= 0) continue;

      var dateObj = _parseDate(tgl);
      if (!tgl || !dateObj) { err.push('Baris ' + (r + 1) + ': Tgl Bayar tidak valid (' + tgl + ')'); continue; }
      var tglNorm = Utilities.formatDate(dateObj, tz, "dd/MM/yyyy");
      if (jumlah <= 0) { err.push('Baris ' + (r + 1) + ': Jumlah Bayar harus lebih dari 0'); continue; }
      if (!noAng) { err.push('Baris ' + (r + 1) + ': No Anggota wajib diisi'); continue; }
      if (!memberMap[noAng.toLowerCase()]) { err.push('Baris ' + (r + 1) + ': No Anggota tidak ditemukan (' + noAng + ')'); continue; }
      if (!piutangByMember[noAng.toLowerCase()]) { err.push('Baris ' + (r + 1) + ': Tidak ada piutang untuk No Anggota (' + noAng + ')'); continue; }

      var totalPiutang = piutangByMember[noAng.toLowerCase()].totalNilai;
      var totalBayarExisting = memberBayar[noAng.toLowerCase()] || 0;
      var sisa = totalPiutang - totalBayarExisting;
      if (sisa <= 0) { err.push('Baris ' + (r + 1) + ': Piutang anggota ' + noAng + ' sudah lunas'); continue; }
      if (jumlah > sisa) { err.push('Baris ' + (r + 1) + ': Jumlah bayar Rp ' + _formatRp(jumlah) + ' melebihi sisa piutang Rp ' + _formatRp(sisa) + ' (anggota ' + noAng + ')'); continue; }

      var key = noAng.toLowerCase() + '|' + jumlah + '|' + tglNorm;
      if (existKeys[key]) { skipped++; continue; }
      existKeys[key] = true;
      memberBayar[noAng.toLowerCase()] = totalBayarExisting + jumlah;
      toInsert.push([tglNorm, noAng, jumlah, nota, metode, 'Masuk']);
    }

    if (toInsert.length > 0) {
      sBayar.getRange(sBayar.getLastRow() + 1, 1, toInsert.length, toInsert[0].length).setValues(toInsert);
      _invalidateKreditCaches();
    }
    var msg = 'Berhasil import ' + toInsert.length + ' baris pembayaran.';
    if (skipped > 0) msg += ' (' + skipped + ' baris duplikat dilewati).';
    if (err.length > 0) msg += ' ' + err.length + ' baris tidak valid: ' + err.join('; ');
    return { status: toInsert.length > 0 ? 'sukses' : 'error', msg: msg };
  } catch (e) {
    return { status: 'error', msg: 'Gagal import: ' + e.toString() };
  }
}

// Opsi form tambah pembayaran manual: daftar anggota + daftar piutang belum lunas
function getOpsiPembayaranManual() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sUs = ss.getSheetByName(SHEET_USERS);
    var dataUs = sUs.getDataRange().getValues();
    var members = [];
    for (var j = 1; j < dataUs.length; j++) {
      var un = String(dataUs[j][0] || '').trim();
      var role = String(dataUs[j][1] || '');
      if (!un || role !== 'Anggota') continue;
      members.push({ username: un, nama: String(dataUs[j][2] || '-') });
    }
    var bayarMap = _getMemberPaymentMap();
    var sPi = _ensureKreditTokoSheet();
    var dataPi = sPi.getDataRange().getValues();
    var piutangByMember = {};
    for (var i = 1; i < dataPi.length; i++) {
      var nA = String(dataPi[i][6]).replace(/'/g, '').trim().toLowerCase();
      if (!nA) continue;
      if (!piutangByMember[nA]) piutangByMember[nA] = 0;
      piutangByMember[nA] += Number(dataPi[i][5]) || 0;
    }
    var piutang = [];
    for (var key in piutangByMember) {
      if (!piutangByMember.hasOwnProperty(key)) continue;
      var nilai = piutangByMember[key];
      var dibayar = bayarMap[key] || 0;
      var sisa = nilai - dibayar;
      if (sisa <= 0) continue;
      piutang.push({ anggota: key, nilai: nilai, dibayar: dibayar, sisa: sisa });
    }
    return { status: 'sukses', members: members, piutang: piutang };
  } catch (e) {
    return { status: 'error', msg: "Server Error: " + e.toString() };
  }
}

// Simpan pembayaran piutang manual langsung ke sheet PembayaranPiutang
function simpanPembayaranManual(obj) {
  try {
    var p = obj || {};
    var tgl = String(p.tgl || '').trim();
    var noAng = String(p.noAnggota || '').trim();
    var jumlah = Number(String(p.jumlah || '').replace(/[^0-9\-\\.]/g, '')) || 0;
    var nota = String(p.nota || '').trim() || '-';
    var metode = String(p.metode || '').trim() || '-';

    var tz = Session.getScriptTimeZone();
    var dateObj = _parseDate(tgl);
    if (!dateObj) return { status: 'error', msg: 'Tanggal bayar tidak valid.' };

    // Hitung total piutang anggota
    var sPi = _ensureKreditTokoSheet();
    var dataPi = sPi.getDataRange().getValues();
    var totalPiutang = 0;
    for (var i = 1; i < dataPi.length; i++) {
      if (String(dataPi[i][6]).replace(/'/g, '').trim().toLowerCase() === noAng.toLowerCase()) {
        totalPiutang += Number(dataPi[i][5]) || 0;
      }
    }
    if (totalPiutang <= 0) return { status: 'error', msg: 'Tidak ada piutang untuk anggota ' + noAng + '.' };

    var bayarMap = _getMemberPaymentMap();
    var dibayar = bayarMap[noAng.toLowerCase()] || 0;
    var sisa = totalPiutang - dibayar;
    if (sisa <= 0) return { status: 'error', msg: 'Piutang anggota ' + noAng + ' sudah lunas.' };
    if (jumlah <= 0) return { status: 'error', msg: 'Jumlah bayar harus lebih dari 0.' };
    if (jumlah > sisa) return { status: 'error', msg: 'Jumlah bayar (Rp ' + _formatRp(jumlah) + ') melebihi sisa tagihan (Rp ' + _formatRp(sisa) + ').' };

    var sBayar = _ensurePembayaranPiutangSheet();
    var tglNorm = Utilities.formatDate(dateObj, tz, "dd/MM/yyyy");
    sBayar.getRange(sBayar.getLastRow() + 1, 1, 1, 6).setValues([[tglNorm, noAng, jumlah, nota, metode, 'Masuk']]);
    _invalidateKreditCaches();
    return { status: 'sukses', msg: 'Pembayaran Rp ' + _formatRp(jumlah) + ' untuk anggota ' + noAng + ' berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', msg: "Gagal simpan: " + e.toString() };
  }
}

// =====================================================
// MANAJEMEN BERKAS DRIVE
// =====================================================
function saveFileToDrive(base64Data, fileName) {
  if (!base64Data || base64Data.trim() === "") return "";
  try {
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var splitBase = base64Data.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var rawBase64 = splitBase[1].replace(/[^A-Za-z0-9\+\/\=]/g, "");
    while (rawBase64.length % 4 !== 0) rawBase64 += "="; 
    var blob = Utilities.newBlob(Utilities.base64Decode(rawBase64), type, fileName);
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w500";
  } catch(e) { throw new Error("Gagal simpan ke Drive: " + e.message); }
}

function saveNotifFileToDrive(base64Data, fileName) {
  if (!base64Data || base64Data.trim() === "") return "";
  try {
    var folder = DriveApp.getFolderById(FOLDER_NOTIF_ID);
    var splitBase = base64Data.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var rawBase64 = splitBase[1].replace(/[^A-Za-z0-9\+\/\=]/g, "");
    while (rawBase64.length % 4 !== 0) rawBase64 += "="; 
    var blob = Utilities.newBlob(Utilities.base64Decode(rawBase64), type, fileName);
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return file.getId(); 
  } catch(e) { throw new Error("Gagal simpan lampiran ke folder Drive: " + e.message); }
}

function deleteFileFromDrive(urlOrId) {
  if (!urlOrId || urlOrId.trim() === "") return;
  try {
    var fileId = urlOrId;
    if (urlOrId.indexOf('id=') !== -1) { fileId = urlOrId.split('id=')[1].split('&')[0]; }
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {}
}

// =====================================================
// FUNGSI NOTIFIKASI BARU (DENGAN READ RECEIPTS)
// =====================================================
function getNotifikasiUser(role, kodeToko, kelompok, username) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NOTIF);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var notifs = [];
    var tz = Session.getScriptTimeZone();
    
    for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][5]).toLowerCase() !== 'aktif') continue; 
        
        var tipe = String(data[i][1]).toLowerCase().trim();
        var detail = String(data[i][2]).toLowerCase().trim();
        var isMatch = false;

        if (tipe === 'semua') { isMatch = true; } 
        else if (tipe === 'toko' && (role === 'Kasir' || role === 'Admin Potongan')) { if (detail === String(kodeToko).toLowerCase().trim()) isMatch = true; } 
        else if (tipe === 'kelompok' && role === 'Anggota') { if (detail === String(kelompok).toLowerCase().trim()) isMatch = true; } 
        else if (tipe === 'personal') { if (detail === String(username).toLowerCase().trim()) isMatch = true; }

        if (isMatch) {
            // Cek apakah user ini sudah membaca
            var dibacaOleh = String(data[i][7] || ""); // Kolom H (Index 7)
            var isRead = false;
            if(dibacaOleh) {
                var arrReaders = dibacaOleh.split(',');
                if(arrReaders.indexOf(String(username).trim()) !== -1) isRead = true;
            }

            notifs.push({
                rowIdx: i + 1,
                waktu: data[i][0] ? Utilities.formatDate(_parseDate(data[i][0]), tz, "dd/MM/yyyy HH:mm") : "-",
                judul: String(data[i][3]),
                pesan: String(data[i][4]),
                lampiranId: String(data[i][6] || ""),
                isRead: isRead
            });
        }
    }
    return notifs;
  } catch (e) { return []; }
}

function markNotifRead(rowIdx, username) {
  try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NOTIF);
      var currentVal = String(sheet.getRange(rowIdx, 8).getValue() || "").trim();
      var arr = currentVal ? currentVal.split(',') : [];
      if(arr.indexOf(String(username).trim()) === -1) {
          arr.push(String(username).trim());
          sheet.getRange(rowIdx, 8).setValue(arr.join(','));
      }
      return true;
  } catch(e) { return false; }
}

function getSemuaNotifikasiAdmin() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NOTIF);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var res = [];
    
    for (var i = data.length - 1; i >= 1; i--) {
       var dibacaOleh = String(data[i][7] || "");
       var readCount = dibacaOleh ? dibacaOleh.split(',').length : 0;

       res.push({
           rowIdx: i + 1,
           waktu: data[i][0] ? Utilities.formatDate(_parseDate(data[i][0]), tz, "dd/MM/yyyy HH:mm") : "-",
           tipe: String(data[i][1]),
           detail: String(data[i][2] || "-"),
           judul: String(data[i][3]),
           pesan: String(data[i][4]),
           status: String(data[i][5]),
           lampiranId: String(data[i][6] || ""),
           readCount: readCount
       });
    }
    return res;
  } catch(e) { return []; }
}

function saveNotifikasi(tipe, detail, judul, pesan, fileBase64, fileName) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NOTIF);
    if(!sheet) return {status: 'error', msg: 'Sheet Notifikasi belum dibuat!'};
    
    var fileIdBaru = "";
    if (fileBase64 && fileBase64.length > 50) {
        var cleanFileName = fileName || ("Lampiran_" + new Date().getTime());
        fileIdBaru = saveNotifFileToDrive(fileBase64, cleanFileName);
    }
    
    var waktu = new Date();
    sheet.appendRow([waktu, tipe, detail, judul, pesan, "Aktif", fileIdBaru, ""]);
    return {status: 'sukses', msg: 'Notifikasi berhasil dikirim.'};
  } catch(e) { return {status: 'error', msg: e.toString()}; }
}

function hapusAtauArsipNotif(rowIdx, aksi) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NOTIF);
    if(aksi === 'hapus') {
        var fileIdLama = String(sheet.getRange(rowIdx, 7).getValue());
        if(fileIdLama) deleteFileFromDrive(fileIdLama);
        sheet.deleteRow(rowIdx);
        return {status: 'sukses', msg: 'Notifikasi & Lampirannya dihapus permanen.'};
    } else {
        sheet.getRange(rowIdx, 6).setValue("Arsip");
        return {status: 'sukses', msg: 'Notifikasi diarsipkan (Tidak tampil ke user).'};
    }
  } catch(e) { return {status: 'error', msg: e.toString()}; }
}

function uploadNotifExcel(dataArray) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NOTIF);
    if(!sheet) return {status: 'error', msg: 'Sheet Notifikasi belum dibuat!'};
    if (dataArray.length < 2) return { status: 'error', msg: "Data Excel Kosong" };

    var newDataToInsert = [];
    var waktuStr = new Date();

    for (var j = 1; j < dataArray.length; j++) {
        var row = dataArray[j];
        var tipe = String(row[0] || "").trim();
        var detail = String(row[1] || "").trim();
        var judul = String(row[2] || "").trim();
        var pesan = String(row[3] || "").trim();
        
        if (!tipe || !judul || !pesan) continue; 
        
        // Col G (Lampiran) & Col H (Dibaca) diset kosong
        newDataToInsert.push([waktuStr, tipe, detail, judul, pesan, "Aktif", "", ""]);
    }

    if (newDataToInsert.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newDataToInsert.length, 8).setValues(newDataToInsert);
      return { status: 'sukses', msg: "Berhasil broadcast " + newDataToInsert.length + " notifikasi via Excel." };
    } else {
      return { status: 'error', msg: "Format Excel salah atau data kosong." };
    }
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

// =====================================================
// MANAJEMEN PENGGUNA (LOGIN & PROFIL)
// =====================================================
function loginUser(u, p) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    if (!sheet) return { status: 'error', msg: 'Sheet Users tidak ditemukan!' };
    var data = sheet.getDataRange().getValues();
    var inputU = String(u).trim().toLowerCase();
    var inputP = String(p).trim();

    for (var i = 1; i < data.length; i++) {
      var rowU = data[i][0];
      if (!rowU) continue;
      var sheetU = String(rowU).trim().toLowerCase();
      var sheetP = String(data[i][1]).trim();
      var matchUser = (sheetU === inputU);

      if (!matchUser && inputU !== "") {
         var numSheet = parseInt(sheetU, 10);
         var numInput = parseInt(inputU, 10);
         if (!isNaN(numSheet) && !isNaN(numInput) && numSheet === numInput) matchUser = true;
      }

      if (matchUser) {
        var statusAnggota = String(data[i][14] || 'Aktif').trim();
        if (statusAnggota.toLowerCase() === 'nonaktif') {
            return { status: 'error', msg: 'Akun Anda telah di-NONAKTIF-kan. Silakan hubungi Admin.' };
        }

        if (sheetP === inputP) {
          var role = String(data[i][3]||'Anggota');
          var kelompok = String(data[i][4]||'-').trim();
          var displayToko = kelompok;

          // Ambil Nama Asli Unit dari Sheet Unit berdasarkan Kode Kelompok (Unit)
          if (kelompok !== '-' && kelompok !== '') {
            try {
              var unitSheet = _ensureUnitSheet();
              var unitData = unitSheet.getDataRange().getValues();
              for(var j=1; j<unitData.length; j++) {
                if (String(unitData[j][0]).trim().toUpperCase() === kelompok.toUpperCase()) {
                  displayToko = String(unitData[j][1]); // Kolom B: Nama Unit
                  break;
                }
              }
            } catch(e) {}
          }

          var batas = _getBatasKreditAnggota(data[i][0]);
          var limitGlobal = batas.effective;
          var hutangBulanIni = _getKreditTokoBulanIni(data[i][0]);

          return { 
            status: 'sukses', user: String(data[i][0]), nama: String(data[i][2]||'-'), 
            role: role, kelompok: kelompok, 
            toko: displayToko,
            limitPiutang: limitGlobal,
outstanding: batas.outstanding,
            foto: String(data[i][5]||''), nip: String(data[i][6]||'-'), 
            tglMasuk: _formatTglIndo(data[i][7]), ktp: String(data[i][8]||'-'), 
            email: String(data[i][9]||'-'), noHp: String(data[i][10]||'-'), 
            rekBank: String(data[i][11]||'-'), namaBank: String(data[i][12]||'-'), 
            atasNama: String(data[i][13]||'-'), statusAnggota: statusAnggota, 
            jabatan: String(data[i][15]||'Anggota Kelompok'), 
            alamatKtp: String(data[i][16]||'-'), alamatDomisili: String(data[i][17]||'-'),
            kodeToko: kelompok.toUpperCase(),
            npwp: String(data[i][18]||'-'), 
            jk: String(data[i][19]||'-'),
            tempatLahir: String(data[i][20]||'-'),
            tglLahir: _formatTglIndo(data[i][21]),
            kota: String(data[i][22]||'-'),
            blokirPiutang: String(data[i][23]||'').trim()
          };
        } else {
          return { status: 'error', msg: 'Password SALAH untuk pengguna: ' + String(data[i][0]) };
        }
      }
    }
    return { status: 'error', msg: 'Nomor Anggota tidak terdaftar!' };
  } catch (e) { return { status: 'error', msg: 'System Error: ' + e.message }; }
}

function getUserInfoByUsername(username) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    if (!sheet) return { status: 'error', msg: 'Sheet Users tidak ditemukan!' };
    var data = sheet.getDataRange().getValues();
    var query = String(username).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      var rowU = String(data[i][0] || '').trim().toLowerCase();
      if (!rowU) continue;
      var match = (rowU === query);
      if (!match && query !== "") {
        var numRow = parseInt(rowU, 10);
        var numQuery = parseInt(query, 10);
        if (!isNaN(numRow) && !isNaN(numQuery) && numRow === numQuery) match = true;
      }
      if (match) {
        var batas = _getBatasKreditAnggota(data[i][0]);
        var limitGlobal = batas.effective;
        var hutangBulanIni = _getKreditTokoBulanIni(data[i][0]);
        
        return { status: 'sukses', data: {
          username: String(data[i][0] || ''),
          nama: String(data[i][2] || '-'),
          kelompok: String(data[i][4] || '-'),
          statusAnggota: String(data[i][14] || 'Aktif'),
          noHp: String(data[i][10] || '-'),
          email: String(data[i][9] || '-'),
          foto: String(data[i][5] || ''),
          limit: limitGlobal,
          sisaLimit: limitGlobal - hutangBulanIni,
          outstanding: batas.outstanding,
          blokirPiutang: String(data[i][23]||'').trim()
        } };
      }
    }
    return { status: 'error', msg: 'Nomor anggota tidak ditemukan.' };
  } catch (e) { return { status: 'error', msg: 'System Error: ' + e.message }; }
}

function getAllUsersFull() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS).getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    users.push({ 
      username: String(r[0]||''), pass: String(r[1]||''), nama: String(r[2]||''), role: String(r[3]||''), 
      kelompok: String(r[4]||''), foto: String(r[5]||''), nip: String(r[6]||''), 
      tglMasuk: _formatTglIndo(r[7]), ktp: String(r[8]||''), email: String(r[9]||''), hp: String(r[10]||''), 
      rek: String(r[11]||''), namaBank: String(r[12]||''), atasNama: String(r[13]||''), 
      status: String(r[14]||'Aktif'), jabatan: String(r[15]||'Anggota Kelompok'),
      alamatKtp: String(r[16]||'-'), alamatDomisili: String(r[17]||'-'),
      npwp: String(r[18]||'-'), jk: String(r[19]||'-'),
      tempatLahir: String(r[20]||'-'), tglLahir: _formatTglIndo(r[21]),
      kota: String(r[22]||'-'),
      blokirPiutang: String(r[23]||'').trim()
    });
  }
  return users;
}

function getAnggotaListKasir() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS).getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var role = String(r[3] || '');
    if (role !== 'Anggota') continue;
    users.push({
      username: String(r[0] || ''),
      nama: String(r[2] || '-'),
      kelompok: String(r[4] || '-'),
      foto: String(r[5] || ''),
      nip: String(r[6] || '-'),
      tglMasuk: _formatTglIndo(r[7]),
      ktp: String(r[8] || '-'),
      email: String(r[9] || '-'),
      hp: String(r[10] || '-'),
      rek: String(r[11] || '-'),
      namaBank: String(r[12] || '-'),
      atasNama: String(r[13] || '-'),
      status: String(r[14] || 'Aktif'),
      jabatan: String(r[15] || 'Anggota Kelompok'),
      alamatKtp: String(r[16] || '-'),
      alamatDomisili: String(r[17] || '-'),
      npwp: String(r[18] || '-'),
      jk: String(r[19] || '-'),
      tempatLahir: String(r[20] || '-'),
      tglLahir: _formatTglIndo(r[21]),
      kota: String(r[22] || '-'),
      blokirPiutang: String(r[23] || '').trim()
    });
  }
  return users;
}

function updateProfilMandiri(username, passLama, passBaru, fotoBase64, nip, ktp, email, hp, rek, bank, atasnama, alamatKtp, alamatDomisili, nama, npwp, jk, tempatLahir, tglLahir, kota) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(username).trim()) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return { status: 'error', msg: 'User tidak ditemukan.' };
    if (String(data[rowIndex-1][1]).trim() !== String(passLama).trim()) return { status: 'error', msg: 'Password saat ini salah!' };

    var finalPass = (passBaru && passBaru.trim() !== "") ? passBaru : data[rowIndex-1][1];
    var linkFoto = data[rowIndex-1][5]; 

    if (fotoBase64 && fotoBase64.length > 50) {
      var namaFile = "Foto_" + username + "_" + new Date().getTime() + ".jpg";
      var urlBaru = saveFileToDrive(fotoBase64, namaFile);
      if (urlBaru) { deleteFileFromDrive(linkFoto); linkFoto = urlBaru; }
    }

    sheet.getRange(rowIndex, 2).setValue(finalPass);
    sheet.getRange(rowIndex, 3).setValue(nama);
    sheet.getRange(rowIndex, 6).setValue(linkFoto);  
    sheet.getRange(rowIndex, 7).setValue("'" + nip);
    sheet.getRange(rowIndex, 9).setValue("'" + ktp);
    sheet.getRange(rowIndex, 10).setValue(email || '-');
    sheet.getRange(rowIndex, 11).setValue("'" + hp);
    sheet.getRange(rowIndex, 12).setValue("'" + rek);
    sheet.getRange(rowIndex, 13).setValue(bank || '-');
    sheet.getRange(rowIndex, 14).setValue(atasnama || '-');
    sheet.getRange(rowIndex, 17).setValue(alamatKtp || '-'); 
    sheet.getRange(rowIndex, 18).setValue(alamatDomisili || '-'); 
    sheet.getRange(rowIndex, 19).setValue("'" + npwp);
    sheet.getRange(rowIndex, 20).setValue(jk);
    sheet.getRange(rowIndex, 21).setValue(tempatLahir);
    sheet.getRange(rowIndex, 22).setValue(tglLahir);
    sheet.getRange(rowIndex, 23).setValue(kota);

    return { status: 'sukses', msg: 'Profil & Foto Diperbarui!', url: linkFoto };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function saveUserAdmin(mode, oldUsername, obj) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    var data = sheet.getDataRange().getValues();
    var linkFotoBaru = "";
    
    if (obj.foto_base64 && obj.foto_base64.length > 50) {
      var namaFile = "Foto_" + obj.username + "_" + new Date().getTime() + ".jpg";
      linkFotoBaru = saveFileToDrive(obj.foto_base64, namaFile);
    }

    var rowData = [ 
      "'" + obj.username, obj.pass, obj.nama, obj.role, obj.kelompok, linkFotoBaru, 
      "'" + obj.nip, obj.tglMasuk, "'" + obj.ktp, obj.email, "'" + obj.hp, "'" + obj.rek, 
      obj.namaBank, obj.atasNama, obj.status, obj.jabatan, obj.alamatKtp, obj.alamatDomisili,
      "'" + (obj.npwp || "-"), obj.jk || "-", obj.tempatLahir || "-", obj.tglLahir || "-", obj.kota || "-",
      obj.blokirPiutang || ""
    ];

    if (mode === 'add') {
      for (var i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === String(obj.username).trim()) return { status: 'error', msg: 'User sudah ada!' }; }
      sheet.appendRow(rowData); _cacheRemove('kp_users'); return { status: 'sukses', msg: 'User ditambahkan.' };
    } else {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(oldUsername).trim()) {
          if (!linkFotoBaru) { rowData[5] = data[i][5]; } else { deleteFileFromDrive(data[i][5]); }
          sheet.getRange(i + 1, 1, 1, 24).setValues([rowData]);
          _cacheRemove('kp_users');
          return { status: 'sukses', msg: 'Data diperbarui.' };
        }
      }
      return { status: 'error', msg: 'User tidak ditemukan.' };
    }
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function deleteUserAdmin(u) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(u).trim()) { 
      deleteFileFromDrive(data[i][5]); sheet.deleteRow(i + 1); 
      _cacheRemove('kp_users');
      return { status: 'sukses', msg: 'User dihapus.' }; 
    }
  }
}

function toggleBlokirPiutangUser(username) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(username).trim()) {
        var current = String(data[i][23] || '').trim().toLowerCase();
        var newVal = (current === 'y' || current === 'yes' || current === '1' || current === 'blokir') ? '' : 'blokir';
        sheet.getRange(i + 1, 24).setValue(newVal);
        _cacheRemove('kp_users');
        return { status: 'sukses', msg: newVal === 'blokir' ? 'Piutang anggota diblokir.' : 'Blokir piutang dicabut.', blokir: newVal === 'blokir' };
      }
    }
    return { status: 'error', msg: 'User tidak ditemukan.' };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

// =====================================================
// MANAJEMEN UNIT (TOKO)
// =====================================================
function _ensureUnitSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_UNIT);
  if (!s) {
    s = ss.insertSheet(SHEET_UNIT);
    s.getRange(1,1,1,4).setValues([["Kode Unit", "Nama Unit", "No WA", "Alamat"]]);
  }
  return s;
}

function getSemuaUnit() {
  try {
    var s = _ensureUnitSheet();
    var data = s.getDataRange().getValues();
    var res = [];
    for (var i = 1; i < data.length; i++) {
      res.push({
        kode: String(data[i][0] || ""),
        nama: String(data[i][1] || ""),
        wa: String(data[i][2] || ""),
        alamat: String(data[i][3] || "")
      });
    }
    return res;
  } catch(e) { return []; }
}

function saveUnit(oldKode, data) {
  try {
    var s = _ensureUnitSheet();
    var sheetData = s.getDataRange().getValues();
    var rowData = ["'" + data.kode, data.nama, "'" + data.wa, data.alamat];
    
    if (!oldKode) {
      for (var i = 1; i < sheetData.length; i++) {
        if (String(sheetData[i][0]).trim().toUpperCase() === data.kode.toUpperCase()) {
          return { status: "error", msg: "Kode Unit sudah terdaftar!" };
        }
      }
      s.appendRow(rowData);
      return { status: "sukses", msg: "Unit berhasil ditambahkan." };
    } else {
      for (var i = 1; i < sheetData.length; i++) {
        if (String(sheetData[i][0]).trim().toUpperCase() === oldKode.toUpperCase()) {
          s.getRange(i + 1, 1, 1, 4).setValues([rowData]);
          return { status: "sukses", msg: "Data Unit diperbarui." };
        }
      }
      return { status: "error", msg: "Unit tidak ditemukan." };
    }
  } catch (e) { return { status: "error", msg: e.toString() }; }
}

function deleteUnit(kode) {
  try {
    var s = _ensureUnitSheet();
    var data = s.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === kode.toUpperCase()) {
        s.deleteRow(i + 1);
        return { status: "sukses", msg: "Unit dihapus." };
      }
    }
  } catch (e) { return { status: "error", msg: e.toString() }; }
}

function getTemanKelompok(kelompok) {
  if (!kelompok || kelompok === '-' || kelompok === '') return { status: 'error', msg: 'Data kelompok kosong.' };
  try {
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS).getDataRange().getValues();
    var teman = []; var sk = String(kelompok).toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][4]).toLowerCase().trim() === sk && String(data[i][3]).toLowerCase().trim() === 'anggota') {
        teman.push({ nama: String(data[i][2] || '-').trim(), hp: String(data[i][10] || '-').trim(), jabatan: String(data[i][15] || 'Anggota Kelompok').trim(), foto: String(data[i][5] || '').trim() });
      }
    }
    return { status: 'sukses', data: teman };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function getDaftarToko() {
  try {
    var units = getSemuaUnit();
    var toko = units.map(function(u) {
      return { username: u.kode, nama: u.nama, wa: u.wa };
    });
    return { status: 'sukses', data: toko };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function updateWaToko(username, noWa) {
  try {
    var s = _ensureUnitSheet();
    var data = s.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === String(username).trim().toUpperCase()) {
        s.getRange(i + 1, 3).setValue("'" + noWa); // Kolom C (index 3) adalah Nomor WA di sheet Unit
        return { status: 'sukses', msg: 'Nomor WA Delivery Unit berhasil diperbarui.' };
      }
    }
    return { status: 'error', msg: 'Unit tidak ditemukan.' };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

// =====================================================
// MANAJEMEN VOUCHER (VERSI TURBO & KOMPRESI JSON)
// =====================================================
function getAdminDataUtama() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS);
    if (!sheet) return JSON.stringify({ status: 'sukses', stats: { totalVoucher: 0, totalActive: 0, totalUsed: 0, totalBlokir: 0 }, rekap: [] });
    
    var data = sheet.getDataRange().getValues();
    var rekap = [];
    var tV = 0, tA = 0, tU = 0, tB = 0, tE = 0;
    
    var today = new Date();
    today.setHours(0,0,0,0);
    
    function formatTglCepat(d) {
      if (!d || isNaN(d.getTime())) return "-";
      var dd = String(d.getDate()).padStart(2, '0');
      var mm = String(d.getMonth() + 1).padStart(2, '0');
      var yyyy = d.getFullYear();
      return dd + '/' + mm + '/' + yyyy;
    }
    
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; 
      tV++;
      
      var rawStatus = String(data[i][6] || "").trim();
      var st = rawStatus; 
      var ex = _parseDate(data[i][7]);
      var startD = _parseDate(data[i][8]);

      if(ex) ex.setHours(23,59,59,999);
      if(startD) startD.setHours(0,0,0,0);

      if (rawStatus.toLowerCase() === 'active') {
          if (ex && today > ex) st = 'Expired'; 
          else if (startD && today < startD) st = 'Belum Aktif';
      }

      var finalSt = st.toLowerCase();
      if (finalSt === 'active') { tA++; } 
      else if (finalSt === 'used') { tU++; } 
      else if (finalSt === 'expired') { tE++; }
      else { tB++; } 
      
      // Dikompres menjadi Array agar payload 60% lebih ringan
      rekap.push([
        String(data[i][0]).trim(), // 0: Kode
        String(data[i][2]),        // 1: Nama
        String(data[i][1]).trim(), // 2: No Anggota
        String(data[i][4]),        // 3: Kelompok
        st,                        // 4: Status
        formatTglCepat(startD),    // 5: Start Date
        formatTglCepat(ex),        // 6: Exp Date
        Number(data[i][5]) || 0    // 7: Nilai
      ]);
    }
    
    // Dikirim sebagai STRING agar Google tidak error "Payload Too Large"
    return JSON.stringify({ 
        status: 'sukses',
        stats: { totalVoucher: tV, totalActive: tA, totalUsed: tU, totalBlokir: tB, totalExpired: tE }, 
        rekap: rekap 
    });
  } catch(e) { 
      return JSON.stringify({ status: 'error', msg: e.toString() }); 
  }
}

// Perbarui masa aktif satu voucher (expired), status dikembalikan ke Active
function updateMasaAktifVoucher(kode, tglMulai, tglExp) {
  try {
    var sV = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS);
    var dV = sV.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var mObj = _parseDate(tglMulai);
    var eObj = _parseDate(tglExp);
    if (!mObj || !eObj) return { status: 'error', msg: 'Format tanggal tidak valid.' };
    if (eObj.getTime() <= mObj.getTime()) return { status: 'error', msg: 'Tanggal Expired harus setelah Tanggal Mulai.' };
    var row = -1;
    for (var i = 1; i < dV.length; i++) {
      if (String(dV[i][0]).trim().toLowerCase() === String(kode).trim().toLowerCase()) { row = i + 1; break; }
    }
    if (row === -1) return { status: 'error', msg: 'Voucher ' + kode + ' tidak ditemukan.' };
    var mDate = Utilities.formatDate(mObj, tz, "dd/MM/yyyy");
    var eDate = Utilities.formatDate(eObj, tz, "dd/MM/yyyy");
    sV.getRange(row, 9).setValue(mDate);   // Start Date
    sV.getRange(row, 8).setValue(eDate);   // Exp Date
    sV.getRange(row, 7).setValue('Active'); // Status
    return { status: 'sukses', msg: 'Masa aktif voucher ' + kode + ' diperbarui: ' + mDate + ' s/d ' + eDate + '.' };
  } catch (e) { return { status: 'error', msg: "Gagal: " + e.toString() }; }
}

// Perpanjang sekaligus semua voucher yang berstatus Expired
function perpanjangSemuaExpired(tglMulai, tglExp) {
  try {
    var sV = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS);
    var dV = sV.getDataRange().getValues();
    if (dV.length < 2) return { status: 'error', msg: 'Tidak ada data voucher.' };
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var tz = Session.getScriptTimeZone();
    var mObj = _parseDate(tglMulai);
    var eObj = _parseDate(tglExp);
    if (!mObj || !eObj) return { status: 'error', msg: 'Format tanggal tidak valid.' };
    if (eObj.getTime() <= mObj.getTime()) return { status: 'error', msg: 'Tanggal Expired harus setelah Tanggal Mulai.' };
    var mDate = Utilities.formatDate(mObj, tz, "dd/MM/yyyy");
    var eDate = Utilities.formatDate(eObj, tz, "dd/MM/yyyy");
    var count = 0;
    for (var i = 1; i < dV.length; i++) {
      if (!String(dV[i][0] || '').trim()) continue;
      var rawStatus = String(dV[i][6] || '').trim().toLowerCase();
      if (rawStatus === 'used') continue;
      var ex = _parseDate(dV[i][7]);
      var isExpired = (rawStatus === 'expired') || (ex && today > ex);
      if (!isExpired) continue;
      sV.getRange(i + 1, 9).setValue(mDate);
      sV.getRange(i + 1, 8).setValue(eDate);
      sV.getRange(i + 1, 7).setValue('Active');
      count++;
    }
    return { status: count > 0 ? 'sukses' : 'error', msg: count > 0 ? count + ' voucher expired berhasil diperpanjang.' : 'Tidak ada voucher expired yang ditemukan.' };
  } catch (e) { return { status: 'error', msg: "Gagal: " + e.toString() }; }
}

// Perpanjang semua voucher Expired milik satu No. Anggota
function perpanjangVoucherAnggota(noAnggota, tglMulai, tglExp) {
  try {
    var sV = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS);
    var dV = sV.getDataRange().getValues();
    if (dV.length < 2) return { status: 'error', msg: 'Tidak ada data voucher.' };
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var tz = Session.getScriptTimeZone();
    var mObj = _parseDate(tglMulai);
    var eObj = _parseDate(tglExp);
    if (!mObj || !eObj) return { status: 'error', msg: 'Format tanggal tidak valid.' };
    if (eObj.getTime() <= mObj.getTime()) return { status: 'error', msg: 'Tanggal Expired harus setelah Tanggal Mulai.' };
    var mDate = Utilities.formatDate(mObj, tz, "dd/MM/yyyy");
    var eDate = Utilities.formatDate(eObj, tz, "dd/MM/yyyy");
    var count = 0;
    var kodeList = [];
    for (var i = 1; i < dV.length; i++) {
      if (String(dV[i][1]).trim() !== String(noAnggota).trim()) continue;
      var rawStatus = String(dV[i][6] || '').trim().toLowerCase();
      if (rawStatus === 'used') continue;
      var ex = _parseDate(dV[i][7]);
      var isExpired = (rawStatus === 'expired') || (ex && today > ex);
      if (!isExpired) continue;
      sV.getRange(i + 1, 9).setValue(mDate);
      sV.getRange(i + 1, 8).setValue(eDate);
      sV.getRange(i + 1, 7).setValue('Active');
      count++;
      kodeList.push(String(dV[i][0]).trim());
    }
    if (count === 0) return { status: 'error', msg: 'Tidak ada voucher expired untuk anggota ' + noAnggota + '.' };
    return { status: 'sukses', msg: count + ' voucher expired anggota ' + noAnggota + ' diperpanjang: ' + kodeList.join(', ') + '.', kode: kodeList };
  } catch (e) { return { status: 'error', msg: "Gagal: " + e.toString() }; }
}

function cekMember(noAnggota) {
  if (!noAnggota) return { status: 'error', msg: 'No. Anggota kosong!' };
  try {
    var dV = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS).getDataRange().getValues();
    var found = false; var nama = ""; var res = []; var tA = 0, tB = 0, tU = 0, tBl = 0;
    var today = new Date();
    today.setHours(0,0,0,0);

    for (var i = 1; i < dV.length; i++) {
      if (String(dV[i][1]).trim() === String(noAnggota).trim()) { 
        found = true; 
        nama = String(dV[i][2]); 
        
        var st = String(dV[i][6]).trim();
        var ex = _parseDate(dV[i][7]);
        var startD = _parseDate(dV[i][8]);

        if(ex) ex.setHours(23,59,59,999);
        if(startD) startD.setHours(0,0,0,0);

        var isExp = (ex && today > ex);
        var isEarly = (startD && today < startD);

        if (st === 'Active' && isExp) st = 'Expired';
        else if (st === 'Active' && isEarly) st = 'Belum Aktif';

        res.push({ kode: String(dV[i][0]), nilai: "Rp " + _formatRp(dV[i][5]), status: st });
        
        if (st === 'Active') tA++; 
        else if (st === 'Belum Aktif') tB++; 
        else if (st === 'Used') tU++; 
        else tBl++; 
      }
    }
    if (!found) return { status: 'error', msg: 'Pelanggan tidak memiliki voucher.' };
    return { status: 'sukses', nama: nama, total: res.length, aktif: tA, belumAktif: tB, terpakai: tU, blokir: tBl, vouchers: res };
  } catch(e) { return { status: 'error', msg: e.toString() }; }
}

function ubahStatusMember(noAnggota, targetStatus) {
  try {
    var sV = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS);
    var dV = sV.getDataRange().getValues();
    var count = 0;
    for (var i = 1; i < dV.length; i++) {
      if (String(dV[i][1]).trim() === String(noAnggota).trim() && String(dV[i][6]) !== 'Used') { 
          sV.getRange(i + 1, 7).setValue(targetStatus); 
          count++; 
      }
    }
    if (count === 0) return { status: 'error', msg: 'Tidak ada voucher yg bisa diubah.' };
    return { status: 'sukses', msg: count + ' voucher diubah.' };
  } catch(e) { return { status: 'error', msg: e.toString() }; }
}

function blokirVoucher(kode) {
  var sV = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS);
  var dV = sV.getDataRange().getValues();
  for (var i = 1; i < dV.length; i++) { 
      if (String(dV[i][0]).trim() === String(kode).trim()) { 
          sV.getRange(i + 1, 7).setValue('Diblokir'); 
          return 'sukses'; 
      } 
  }
}

function aktifkanVoucher(kode) {
  var sV = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS);
  var dV = sV.getDataRange().getValues();
  for (var i = 1; i < dV.length; i++) { 
      if (String(dV[i][0]).trim() === String(kode).trim()) { 
          sV.getRange(i + 1, 7).setValue('Active'); 
          return 'sukses'; 
      } 
  }
}

function prosesVoucher(kode, mode, userToko, notaToko) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sVoucher = ss.getSheetByName(SHEET_VOUCHERS);
    var dVoucher = sVoucher.getDataRange().getValues();
    var sLaporan = ss.getSheetByName(SHEET_LAPORAN);
    var vRow = -1; var vData = null;
    
    for (var i = 1; i < dVoucher.length; i++) { 
        if (String(dVoucher[i][0]).trim() === String(kode).trim()) { vRow = i + 1; vData = dVoucher[i]; break; } 
    }
    if (vRow === -1) return { status: 'error', msg: '❌ Voucher Tidak Ditemukan!' };

    var tz = Session.getScriptTimeZone();
    var st = String(vData[6]).trim(); 
    var ex = _parseDate(vData[7]); 
    var startD = _parseDate(vData[8]); 
    var today = new Date(); 
    
    today.setHours(0,0,0,0);
    if(ex) ex.setHours(23,59,59,999);
    if(startD) startD.setHours(0,0,0,0);

    var isExp = (ex && today > ex) ? true : false;
    var isEarly = (startD && today < startD) ? true : false; 
    var sisaLimit = 0;
    var limitGlobal = 0;
    if (vData[1]) {
      var batas = _getBatasKreditAnggota(vData[1]);
      limitGlobal = batas.effective;
      var hutangBulanIni = _getKreditTokoBulanIni(vData[1]);
      sisaLimit = limitGlobal - hutangBulanIni;
    }
    var isBlocked = _isPiutangBlocked(vData[1]);

    var info = { 
        kode: vData[0], noAnggota: vData[1], nama: vData[2], kelompok: vData[4], status: st, nilai: "Rp " + _formatRp(vData[5]), 
        masaBerlaku: (startD ? Utilities.formatDate(startD, tz, "dd/MM/yyyy") : "-") + " s/d " + (ex ? Utilities.formatDate(ex, tz, "dd/MM/yyyy") : "-"),
        sisaLimit: sisaLimit, isBlocked: isBlocked
    };

    if (mode === 'cek') {
      if (st === 'Used') {
        var dLap = sLaporan.getDataRange().getValues();
        for (var j = dLap.length - 1; j >= 1; j--) {
          if (String(dLap[j][3]).trim() === String(kode).trim()) {
              return { status: 'sudah_pakai_struk', nota: { toko: dLap[j][4], waktu: Utilities.formatDate(_parseDate(dLap[j][0]), tz, "dd/MM/yyyy HH:mm"), notaToko: dLap[j][2], idSystem: dLap[j][1], kasir: dLap[j][5], kode: dLap[j][3], nama: dLap[j][7], noAnggota: dLap[j][8], kelompok: dLap[j][9], masaBerlaku: info.masaBerlaku, nilai: "Rp " + _formatRp(dLap[j][6]) } };
          }
        }
        return { status: 'used_error_no_history', nota: { toko: "-", waktu: "-", notaToko: "-", idSystem: "-", kasir: "-", kode: info.kode, nama: info.nama, noAnggota: info.noAnggota, kelompok: info.kelompok, masaBerlaku: info.masaBerlaku, nilai: info.nilai } };
      }
      if (st === 'Active') { 
          if (isExp) return { status: 'invalid_expired', nota: info }; 
          if (isEarly) return { status: 'invalid_belum_aktif', nota: info }; 
          return { status: 'cek_valid', nota: info }; 
      }
      if (st === 'Belum Aktif') return { status: 'invalid_belum_aktif', nota: info };
      return { status: 'invalid_blokir', nota: info };
    } 

    if (mode === 'redeem') {
      if (st === 'Used') return { status: 'sudah_pakai_struk' }; 
      if (st !== 'Active' || isExp || isEarly) return { status: 'error', msg: '❌ Voucher tidak dapat digunakan (Belum Masuk Masa Aktif / Kadaluarsa).' };
      
      var timestamp = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
      var idSys = "TRX-" + new Date().getTime().toString().slice(-6);
      sVoucher.getRange(vRow, 7).setValue('Used'); 
      sLaporan.appendRow([timestamp, idSys, "'" + notaToko, "'" + info.kode, userToko.nama, "'" + userToko.user, vData[5], info.nama, "'" + info.noAnggota, info.kelompok]);
      return { status: 'sukses', nota: { toko: userToko.nama, waktu: timestamp, notaToko: notaToko, idSystem: idSys, kasir: userToko.user, kode: info.kode, nama: info.nama, noAnggota: info.noAnggota, kelompok: info.kelompok, masaBerlaku: info.masaBerlaku, nilai: info.nilai } };
    }
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function cariVoucherMulti(noAnggota) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sUsers = ss.getSheetByName(SHEET_USERS);
    var uData = sUsers.getDataRange().getValues();
    var userInfo = null;
    var uQuery = String(noAnggota).trim().toLowerCase();

    for (var k = 1; k < uData.length; k++) {
      var rowUser = String(uData[k][0]).trim().toLowerCase();
      if (rowUser === uQuery) {
        var batas = _getBatasKreditAnggota(uData[k][0]);
        var limitGlobal = batas.effective;
        var hutangBulanIni = _getKreditTokoBulanIni(uData[k][0]);
        userInfo = {
          username: String(uData[k][0]),
          nama: String(uData[k][2] || "-"),
          kelompok: String(uData[k][4] || "-"),
          foto: String(uData[k][5] || ""),
          limit: limitGlobal,
          sisaLimit: limitGlobal - hutangBulanIni,
          outstanding: batas.outstanding,
          blokirPiutang: String(uData[k][23] || '').trim()
        };
        break;
      }
    }

    var data = ss.getSheetByName(SHEET_VOUCHERS).getDataRange().getValues();
    var res = []; var tz = Session.getScriptTimeZone(); var today = new Date();
    today.setHours(0,0,0,0);
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === uQuery) { 
        var ex = _parseDate(data[i][7]);
        var startD = _parseDate(data[i][8]);
        
        if(ex) ex.setHours(23,59,59,999);
        if(startD) startD.setHours(0,0,0,0);
        
        var st = String(data[i][6]).trim();
        var isExp = (ex && today > ex);
        var isEarly = (startD && today < startD);
        
        if (st === 'Active' && isExp) st = 'Expired';
        else if (st === 'Active' && isEarly) st = 'Belum Aktif';
        
        res.push({ 
            kode: String(data[i][0]), 
            nama: String(data[i][2]), 
            nilai: Number(data[i][5]) || 0, 
            startDate: startD ? Utilities.formatDate(startD, tz, "dd/MM/yyyy") : "-", 
            expDate: ex ? Utilities.formatDate(ex, tz, "dd/MM/yyyy") : "-", 
            status: st, 
            bisaRedeem: (st === 'Active' && !isEarly && !isExp) 
        });
      }
    }
    if (res.length === 0) return { status: 'error', msg: 'Anggota tidak ditemukan / tidak ada voucher.' };
    return { status: 'sukses', data: res, info: userInfo };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

// Ringkasan pemakaian per bulan (12 bulan tahun berjalan) untuk kartu kasir/anggota.
function getPemakaianBulananMember(noAnggota) {
  try {
    var tahun = new Date().getFullYear();
    var data = _getCachedPiutangData();
    var totals = {};
    var q = String(noAnggota).trim();
    for (var i = 1; i < data.length; i++) {
      var nM = String(data[i][6]).replace(/'/g, '').trim();
      if (nM !== q) continue;
      var tgl = _parseDate(data[i][0]);
      if (!tgl || tgl.getFullYear() !== tahun) continue;
      var m = tgl.getMonth() + 1;
      totals[m] = (Number(totals[m]) || 0) + (Number(data[i][5]) || 0);
    }
    var months = [];
    for (var m = 1; m <= 12; m++) months.push({ m: m, total: Math.round(Number(totals[m]) || 0) });
    return { status: 'sukses', tahun: tahun, currentMonth: new Date().getMonth() + 1, months: months };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function redeemMassal(arrKode, userToko, notaToko) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sVoucher = ss.getSheetByName(SHEET_VOUCHERS);
    var dVoucher = sVoucher.getDataRange().getValues();
    var sLaporan = ss.getSheetByName(SHEET_LAPORAN);
    var tz = Session.getScriptTimeZone(); var today = new Date();
    today.setHours(0,0,0,0);
    
    var realtime = new Date(); 
    var timestamp = Utilities.formatDate(realtime, tz, "dd/MM/yyyy HH:mm:ss");
    var idSys = "TRX-" + realtime.getTime().toString().slice(-6);
    var tNilai = 0; var kodes = []; var vName = "-"; var vAngg = "-"; var vKel = "-";
    
    for (var k = 0; k < arrKode.length; k++) {
      var targetKode = String(arrKode[k]).trim();
      for (var i = 1; i < dVoucher.length; i++) {
        if (String(dVoucher[i][0]).trim() === targetKode) {
            var st = String(dVoucher[i][6] || "").trim().toLowerCase();
            var ex = _parseDate(dVoucher[i][7]);
            var startD = _parseDate(dVoucher[i][8]);
            if(ex) ex.setHours(23,59,59,999);
            if(startD) startD.setHours(0,0,0,0);
            
            var isExp = (ex && today > ex);
            var isEarly = (startD && today < startD);
            
            if (st === 'active' && !isExp && !isEarly) {
              sVoucher.getRange(i + 1, 7).setValue('Used');
              tNilai += (Number(dVoucher[i][5]) || 0); kodes.push(dVoucher[i][0]);
              vName = dVoucher[i][2]; vAngg = dVoucher[i][1]; vKel = dVoucher[i][4];
              sLaporan.appendRow([timestamp, idSys + "-" + k, "'" + notaToko, "'" + dVoucher[i][0], userToko.nama, "'" + userToko.user, dVoucher[i][5], vName, "'" + vAngg, vKel]);
            }
            break;
        }
      }
    }
    if (kodes.length === 0) return { status: 'error', msg: 'Semua voucher gagal diredeem karena belum masuk masa aktif atau sudah kedaluwarsa.' };
    return { status: 'sukses', nota: { toko: userToko.nama, waktu: timestamp, notaToko: notaToko, idSystem: idSys + "-MULTI", kasir: userToko.user, kode: kodes.join(", "), nama: vName, noAnggota: vAngg, kelompok: vKel, masaBerlaku: "Sesuai Detail", nilai: "Rp " + _formatRp(tNilai) } };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function redeemMassalPiutang(arrKode, userToko, notaToko, base64Photo, fileName, massNominal) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sVoucher = ss.getSheetByName(SHEET_VOUCHERS);
    var dVoucher = sVoucher.getDataRange().getValues();
    var sLaporan = ss.getSheetByName(SHEET_LAPORAN);
    var sPi = _ensureKreditTokoSheet();
    var tz = Session.getScriptTimeZone(); 
    var today = new Date();
    today.setHours(0,0,0,0);
    
    var realtime = new Date(); 
    var timestamp = Utilities.formatDate(realtime, tz, "dd/MM/yyyy HH:mm:ss");
    var idSys = "TRX-" + (realtime.getTime() + 1000).toString().slice(-8); // ID Transaksi Voucher
    var idPi = 'PIU-' + (realtime.getTime() + 2000).toString().slice(-8); // ID Transaksi Piutang
    var tNilai = 0; 
    var kodes = []; 
    var vName = "-"; 
    var vAngg = "-"; 
    var vKel = "-";
    
    var limitGlobal = 0;
    var curHutang = 0; // dihitung ulang per voucher di dalam loop

    // Proses simpan foto jika ada
    var fileId = '';
    if (base64Photo && base64Photo.length > 50) {
      var fn = fileName || ('piu_massal_' + idPi + '.jpg');
      fileId = saveKreditTokoFileToDrive(base64Photo, fn);
    }
    
    // Proses setiap voucher
    var memberNo = "";
    for (var k = 0; k < arrKode.length; k++) {
      var targetKode = String(arrKode[k]).trim();
      for (var i = 1; i < dVoucher.length; i++) {
        if (String(dVoucher[i][0]).trim() === targetKode) {
            var st = String(dVoucher[i][6] || "").trim().toLowerCase();
            var ex = _parseDate(dVoucher[i][7]);
            var startD = _parseDate(dVoucher[i][8]);
            if(ex) ex.setHours(23,59,59,999);
            if(startD) startD.setHours(0,0,0,0);
            
            var isExp = (ex && today > ex);
            var isEarly = (startD && today < startD);
            
            if (st === 'active' && !isExp && !isEarly) {
              var valNum = (massNominal !== undefined && massNominal !== null && String(massNominal).toString().trim() !== "") ? Number(massNominal) || 0 : Number(dVoucher[i][5]) || 0;
              if (_isLimitKreditAktif()) {
                var batasM = _getBatasKreditAnggota(dVoucher[i][1]);
                var currentMemberDebt = _getKreditTokoBulanIni(dVoucher[i][1]);
                if (batasM.effective <= 0) {
                  continue; // Sisa tagihan sudah membebani limit, kredit diblokir
                }
                if (currentMemberDebt + valNum > batasM.effective) {
                  // Jika salah satu voucher menyebabkan overload, skip atau hentikan? Biasanya skip atau tolak semua.
                  continue; 
                }
              }

              // Tandai voucher sebagai Used
              sVoucher.getRange(i + 1, 7).setValue('Used');
              
              // Akumulasi untuk nota
              tNilai += valNum; 
              kodes.push(dVoucher[i][0]);
              vName = dVoucher[i][2]; 
              vAngg = dVoucher[i][1]; 
              vKel = dVoucher[i][4];

              // Append ke laporan mutasi untuk setiap voucher yang berhasil diredeem
              sLaporan.appendRow([timestamp, (idSys + "-" + k) + " / " + idPi, "'" + notaToko, "'" + dVoucher[i][0], userToko.nama, "'" + userToko.user, valNum, vName, "'" + vAngg, vKel]);
            }
            break;
        }
      }
    }
    
    // Catat piutang satu kali saja (setelah akumulasi nominal)
    if (kodes.length > 0) {
      sPi.appendRow([timestamp, idPi, notaToko || '-', userToko ? userToko.nama : '-', userToko ? userToko.user : '-', tNilai, "'" + vAngg, fileId]);
      _invalidateKreditCaches();
    }

    if (kodes.length === 0) return { status: 'error', msg: 'Semua voucher gagal diredeem karena belum masuk masa aktif atau sudah kedaluwarsa.' };
    return { status: 'sukses', nota: { toko: userToko.nama, waktu: timestamp, notaToko: notaToko, idSystem: idSys + " / " + idPi, kasir: userToko.user, kode: kodes.join(", "), nama: vName, noAnggota: vAngg, kelompok: vKel, masaBerlaku: "Sesuai Detail", nilai: "Rp " + _formatRp(tNilai), piutangId: idPi } };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function catatPiutangVoucher(kode, userToko, notaToko, base64Photo, fileName, nominalOverride) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sVoucher = ss.getSheetByName(SHEET_VOUCHERS);
    var dVoucher = sVoucher.getDataRange().getValues();
    var vRow = -1; var vData = null;
    for (var i = 1; i < dVoucher.length; i++) {
      if (String(dVoucher[i][0]).trim() === String(kode).trim()) { vRow = i + 1; vData = dVoucher[i]; break; }
    }
    if (vRow === -1) return { status: 'error', msg: '❌ Voucher Tidak Ditemukan!' };

    var sPi = _ensureKreditTokoSheet();
    var tz = Session.getScriptTimeZone();
    var timestamp = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
    var idPi = 'PIU-' + new Date().getTime().toString().slice(-8);
    var idSys = 'TRX-' + (new Date().getTime() + 500).toString().slice(-8);
    var valNum = Number(nominalOverride !== undefined && nominalOverride !== null ? nominalOverride : dVoucher[vRow-1][5]) || 0;
    var fileId = '';
    if (base64Photo && base64Photo.length > 50) {
      var fn = fileName || ('piu_' + idPi + '.jpg');
      fileId = saveKreditTokoFileToDrive(base64Photo, fn);
    }
    // Simpan ke mutasi
    var sLaporan = ss.getSheetByName(SHEET_LAPORAN);
    sLaporan.appendRow([timestamp, idSys + " / " + idPi, "'" + notaToko, "'" + dVoucher[vRow-1][0], userToko.nama, "'" + userToko.user, valNum, dVoucher[vRow-1][2], "'" + dVoucher[vRow-1][1], dVoucher[vRow-1][4]]);
    
    // Ubah status voucher menjadi Used agar tidak muncul active lagi
    sVoucher.getRange(vRow, 7).setValue('Used');
    
    // Cek blokir piutang
    if (_isPiutangBlocked(dVoucher[vRow-1][1])) {
      return { status: 'error', msg: '❌ Anggota ' + dVoucher[vRow-1][1] + ' (' + dVoucher[vRow-1][2] + ') sedang diblokir untuk transaksi piutang.' };
    }
    
    sPi.appendRow([timestamp, idPi, notaToko || '-', userToko ? userToko.nama : '-', userToko ? userToko.user : '-', valNum, "'" + dVoucher[vRow-1][1], fileId]);
    _invalidateKreditCaches();
    return { status: 'sukses', nota: { toko: userToko.nama, waktu: timestamp, notaToko: notaToko, idSystem: idSys + " / " + idPi, kasir: userToko.user, kode: dVoucher[vRow-1][0], nama: dVoucher[vRow-1][2], noAnggota: dVoucher[vRow-1][1], kelompok: dVoucher[vRow-1][4], masaBerlaku: "Sesuai Detail", nilai: "Rp " + _formatRp(valNum) } };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function catatPiutangMassal(arrKode, userToko, notaToko, base64Photo, fileName, massNominal) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sVoucher = ss.getSheetByName(SHEET_VOUCHERS);
    var dVoucher = sVoucher.getDataRange().getValues();
    var sPi = _ensureKreditTokoSheet();
    var tz = Session.getScriptTimeZone();
    var timestamp = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
    var idSys = 'TRX-' + (new Date().getTime() + 500).toString().slice(-8);
    var idPi = 'PIU-' + new Date().getTime().toString().slice(-8);
    var fileId = '';
    if (base64Photo && base64Photo.length > 50) {
      var fn = fileName || ('piu_massal_' + idPi + '.jpg');
      fileId = saveKreditTokoFileToDrive(base64Photo, fn);
    }
    var kodes = []; var tNilai = 0; var vName = "-"; var vAngg = "-"; var vKel = "-";
    var sLaporan = ss.getSheetByName(SHEET_LAPORAN);
    
    // Ambil No Anggota dari voucher pertama untuk validasi limit total
    for (var i = 1; i < dVoucher.length; i++) {
        if (String(dVoucher[i][0]).trim() === String(arrKode[0]).trim()) {
            vAngg = String(dVoucher[i][1]).trim();
            break;
        }
    }
    
    if (vAngg && _isLimitKreditAktif()) {
        var batasM = _getBatasKreditAnggota(vAngg);
        if (batasM.effective <= 0) {
          return { status: 'error', msg: '❌ Kredit Toko diblokir: sisa tagihan bulan lalu (Rp ' + _formatRp(batasM.outstanding) + ') sudah mencapai limit. Segera bayar tagihan agar dapat kredit baru.' };
        }
        var hutangBulanIni = _getKreditTokoBulanIni(vAngg);
        if (hutangBulanIni + Number(massNominal) > batasM.effective) {
            return { status: 'error', msg: 'Total piutang melebihi sisa limit anggota! Batas (limit global ' + _formatRp(batasM.global) + ' - sisa tagihan ' + _formatRp(batasM.outstanding) + '): Rp ' + _formatRp(batasM.effective) + '. Sisa: Rp ' + _formatRp(batasM.effective - hutangBulanIni) };
        }
    }
    
    // Cek blokir piutang untuk anggota voucher pertama (semua voucher harus anggota sama)
    if (vAngg && _isPiutangBlocked(vAngg)) {
      return { status: 'error', msg: '❌ Anggota ' + vAngg + ' sedang diblokir untuk transaksi piutang.' };
    }

    for (var k = 0; k < arrKode.length; k++) {
      var targetKode = String(arrKode[k]).trim();
      for (var i = 1; i < dVoucher.length; i++) {
        if (String(dVoucher[i][0]).trim() === targetKode) {
          var isOverride = (massNominal !== undefined && massNominal !== null && String(massNominal).toString().trim() !== "");
          var valNum = isOverride ? (Number(massNominal) / arrKode.length) : (Number(dVoucher[i][5]) || 0);

          tNilai += valNum;
          kodes.push(dVoucher[i][0]);
          vName = dVoucher[i][2]; vAngg = dVoucher[i][1]; vKel = dVoucher[i][4];

          // Simpan ke mutasi per voucher
          sLaporan.appendRow([timestamp, (idSys + "-" + k) + " / " + idPi, "'" + notaToko, "'" + dVoucher[i][0], userToko.nama, "'" + userToko.user, valNum, vName, "'" + vAngg, vKel]);
          
          // Ubah status voucher menjadi Used
          sVoucher.getRange(i + 1, 7).setValue('Used');
          break;
        }
      }
    }
    // Catat piutang massal satu baris
    if (kodes.length > 0) {
      sPi.appendRow([timestamp, idPi, notaToko || '-', userToko ? userToko.nama : '-', userToko ? userToko.user : '-', tNilai, "'" + vAngg, fileId]);
      _invalidateKreditCaches();
    }

    if (kodes.length === 0) return { status: 'error', msg: 'Tidak ada voucher valid untuk dicatat piutang.' };
    return { status: 'sukses', nota: { toko: userToko.nama, waktu: timestamp, notaToko: notaToko, idSystem: idSys + " / " + idPi, kasir: userToko.user, kode: kodes.join(", "), nama: vName, noAnggota: vAngg, kelompok: vKel, masaBerlaku: "Sesuai Detail", nilai: "Rp " + _formatRp(tNilai), piutangId: idPi } };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function cariVoucherUsed(query) {
  try {
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS).getDataRange().getValues();
    var res = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][6]).trim() === 'Used' && (String(data[i][0]).trim() === String(query).trim() || String(data[i][1]).trim() === String(query).trim())) {
        res.push({ kode: String(data[i][0]), nama: String(data[i][2]), noAnggota: String(data[i][1]), nilai: Number(data[i][5]) || 0, status: 'Used' });
      }
    }
    return { status: 'sukses', data: res };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function batalkanRedeem(kode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sV = ss.getSheetByName(SHEET_VOUCHERS);
    var dV = sV.getDataRange().getValues();
    var sL = ss.getSheetByName(SHEET_LAPORAN);
    var dL = sL.getDataRange().getValues();
    for (var i = 1; i < dV.length; i++) {
      if (String(dV[i][0]).trim() === String(kode).trim()) {
        if (String(dV[i][6]).trim() !== 'Used') return { status: 'error', msg: 'Voucher tidak berstatus Used.' };
        sV.getRange(i + 1, 7).setValue('Active');
        for (var j = dL.length - 1; j >= 1; j--) { 
            if (String(dL[j][3]).trim() === String(kode).trim()) { 
                sL.deleteRow(j + 1); 
                break; 
            } 
        }
        return { status: 'sukses', msg: 'Berhasil dibatalkan!' };
      }
    }
    return { status: 'error', msg: 'Kode tidak ditemukan.' };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function batalkanRedeemMassal(arrKode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sV = ss.getSheetByName(SHEET_VOUCHERS);
    var dV = sV.getDataRange().getValues();
    var sL = ss.getSheetByName(SHEET_LAPORAN);
    var count = 0;
    for (var k = 0; k < arrKode.length; k++) {
      for (var i = 1; i < dV.length; i++) {
        if (String(dV[i][0]).trim() === String(arrKode[k]).trim() && String(dV[i][6]).trim() === 'Used') {
          sV.getRange(i + 1, 7).setValue('Active');
          var dL = sL.getDataRange().getValues();
          for (var j = dL.length - 1; j >= 1; j--) { 
              if (String(dL[j][3]).trim() === String(arrKode[k]).trim()) { 
                  sL.deleteRow(j + 1); 
                  break; 
              } 
          }
          count++; break;
        }
      }
    }
    return { status: 'sukses', msg: count + ' Transaksi berhasil dibatalkan.' };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function getLaporanData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LAPORAN);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var lap = []; var tz = Session.getScriptTimeZone();
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      lap.push([ 
          data[i][0] ? Utilities.formatDate(_parseDate(data[i][0]), tz, "yyyy-MM-dd HH:mm:ss") : "-", 
          String(data[i][1]||"-"), String(data[i][2]||"-"), String(data[i][3]||"-"), 
          String(data[i][4]||"-"), String(data[i][5]||"-"), Number(data[i][6]||0), 
          String(data[i][7]||"-"), String(data[i][8]||"-"), String(data[i][9]||"-") 
      ]);
    }
    return lap;
  } catch (e) { return []; }
}

function getAnggotaData(noAnggota) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dVoucher = ss.getSheetByName(SHEET_VOUCHERS).getDataRange().getValues();
    var dLaporan = ss.getSheetByName(SHEET_LAPORAN).getDataRange().getValues();
    var tz = Session.getScriptTimeZone(); var today = new Date();
    today.setHours(0,0,0,0);
    
    var myVouchers = []; var myRiwayat = [];
    
    for (var i = 1; i < dVoucher.length; i++) {
      if (String(dVoucher[i][1]).trim() === String(noAnggota).trim()) {
        var ex = dVoucher[i][7] ? _parseDate(dVoucher[i][7]) : null;
        var startD = _parseDate(dVoucher[i][8]);
        
        if(ex) ex.setHours(23,59,59,999);
        if(startD) startD.setHours(0,0,0,0);

        var st = String(dVoucher[i][6]).trim();
        var isExp = (ex && today > ex);
        var isEarly = (startD && today < startD);
        
        if (st === 'Active' && isExp) st = 'Expired';
        else if (st === 'Active' && isEarly) st = 'Belum Aktif';

        myVouchers.push({ 
            kode: String(dVoucher[i][0]), 
            kelompok: String(dVoucher[i][4]), 
            nilai: Number(dVoucher[i][5]) || 0, 
            status: st, 
            startDate: startD ? Utilities.formatDate(startD, tz, "dd/MM/yyyy") : "-", 
            expDate: ex ? Utilities.formatDate(ex, tz, "dd/MM/yyyy") : "-" 
        });
      }
    }
    
    for (var j = dLaporan.length - 1; j >= 1; j--) {
      if (String(dLaporan[j][8]).trim() === String(noAnggota).trim()) {
        myRiwayat.push({ 
            waktu: dLaporan[j][0] ? Utilities.formatDate(_parseDate(dLaporan[j][0]), tz, "dd/MM/yyyy HH:mm") : "-", 
            idSystem: String(dLaporan[j][1]), 
            notaToko: String(dLaporan[j][2]), 
            kode: String(dLaporan[j][3]), 
            toko: String(dLaporan[j][4]), 
            kasir: String(dLaporan[j][5]), 
            nilai: "Rp " + _formatRp(dLaporan[j][6]) 
        });
      }
    }
    return { status: 'sukses', vouchers: myVouchers, riwayat: myRiwayat };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

// =====================================================
// DATA POTONGAN GAJI
// =====================================================
function getSemuaPotongan() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_POTONGAN);
    if (!sheet) return JSON.stringify([]);
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify([]); 
    
    var data = sheet.getRange(2, 1, lastRow - 1, 38).getValues();
    var res = [];
    var tz = Session.getScriptTimeZone();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var noAnggota = row[0];
      if (!noAnggota || String(noAnggota).trim() === "") continue;
      
      var tot = Number(row[34]) || 0;
      if (tot === 0) { 
          tot += (Number(row[8])||0) + (Number(row[9])||0) + (Number(row[10])||0) + (Number(row[11])||0) + 
                 (Number(row[12])||0) + (Number(row[14])||0) + (Number(row[16])||0) + (Number(row[18])||0) + 
                 (Number(row[20])||0) + (Number(row[22])||0) + (Number(row[24])||0) + (Number(row[26])||0) + 
                 (Number(row[27])||0) + (Number(row[28])||0) + (Number(row[29])||0) + (Number(row[30])||0) + 
                 (Number(row[31])||0) + (Number(row[32])||0) + (Number(row[33])||0);
      }
      
      var cleanData = new Array(38);
      for(var c=0; c<38; c++){
         var cellVal = row[c];
         cleanData[c] = (cellVal instanceof Date) ? Utilities.formatDate(cellVal, tz, "dd/MM/yyyy") : (cellVal == null ? "" : String(cellVal));
      }
      res.push({ 
          rowIdx: i + 2, 
          noAnggota: String(noAnggota).trim(), 
          tahun: parseInt(row[1], 10) || 0,     
          bulan: parseInt(row[2], 10) || 0,     
          total: tot, 
          d: cleanData 
      });
    }
    return JSON.stringify(res);
  } catch (e) { 
    return JSON.stringify({error: e.toString()}); 
  }
}

function getSlipPotongan(noAnggota, bulan, tahun) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_POTONGAN);
    var data = sheet.getDataRange().getValues();
    
    var uNama = "-"; var uNip = "-";
    try {
      var sUsers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS).getDataRange().getValues();
      for(var j=1; j<sUsers.length; j++){
         if(String(sUsers[j][0]).trim() === String(noAnggota).trim()){ 
             uNama = String(sUsers[j][2] || "-").trim(); 
             uNip = String(sUsers[j][6] || "-").trim(); 
             break; 
         }
      }
    } catch(e) {}

    var tz = Session.getScriptTimeZone();
    var targetBulan = parseInt(bulan, 10);
    var targetTahun = parseInt(tahun, 10);

    for (var i = 1; i < data.length; i++) {
      var sheetBulan = parseInt(data[i][2], 10);
      var sheetTahun = parseInt(data[i][1], 10);
      var sheetNo = String(data[i][0]).trim();

      if (sheetNo === String(noAnggota).trim() && sheetTahun === targetTahun && sheetBulan === targetBulan) {
        var cleanData = [];
        for(var c=0; c<38; c++){
           var item = data[i][c];
           if (Object.prototype.toString.call(item) === '[object Date]') {
               cleanData.push(Utilities.formatDate(item, tz, "dd/MM/yyyy"));
           } else {
               cleanData.push(item === undefined || item === null ? "" : String(item));
           }
        }
        return { status: 'sukses', data: cleanData, nama: uNama, nip: uNip };
      }
    }
    return { status: 'error', msg: 'Data slip untuk Bulan ' + targetBulan + ' Tahun ' + targetTahun + ' belum tersedia.' };
  } catch (e) { 
      return { status: 'error', msg: e.toString() }; 
  }
}

function savePotonganManual(rowIdx, arrData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_POTONGAN);
    if (rowIdx > 0) { 
        sheet.getRange(rowIdx, 1, 1, 38).setValues([arrData]); 
        return { status: 'sukses', msg: 'Berhasil diupdate.' }; 
    } else { 
        sheet.appendRow(arrData); 
        return { status: 'sukses', msg: 'Data ditambahkan.' }; 
    }
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function deletePotonganSingle(rowIdx) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_POTONGAN).deleteRow(rowIdx);
}

function deletePotonganMassal(bulan, tahun) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_POTONGAN);
    var data = sheet.getDataRange().getValues();
    var deleted = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]).trim() === String(tahun).trim() && String(data[i][2]).trim() === String(bulan).trim()) { 
          sheet.deleteRow(i + 1); 
          deleted++; 
      }
    }
    return { status: 'sukses', msg: deleted + " data dihapus." };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function uploadPotonganGaji(dataArray) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_POTONGAN);
    if (dataArray.length < 2) return { status: 'error', msg: "Data Kosong" };

    var existingData = sheet.getDataRange().getValues();
    var existingKeys = {};
    
    for (var i = 1; i < existingData.length; i++) {
        var noAngg = String(existingData[i][0]).trim();
        var thn = parseInt(existingData[i][1], 10) || 0;
        var bln = parseInt(existingData[i][2], 10) || 0;
        if(noAngg) existingKeys[noAngg + "_" + thn + "_" + bln] = true;
    }

    var newDataToInsert = [];
    var skippedCount = 0;

    for (var j = 1; j < dataArray.length; j++) {
        var row = dataArray[j];
        var noAngg = String(row[0]).trim();
        var thn = parseInt(row[1], 10) || 0;
        var bln = parseInt(row[2], 10) || 0;

        if (!noAngg) continue; 

        var key = noAngg + "_" + thn + "_" + bln;
        
        if (existingKeys[key]) {
            skippedCount++;
        } else {
            newDataToInsert.push(row);
            existingKeys[key] = true; 
        }
    }

    if (newDataToInsert.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newDataToInsert.length, newDataToInsert[0].length).setValues(newDataToInsert);
      var msg = "Berhasil upload " + newDataToInsert.length + " baris data.";
      if (skippedCount > 0) msg += " (" + skippedCount + " baris dilewati karena duplikat/sudah ada).";
      return { status: 'sukses', msg: msg };
    } else {
      return { status: 'error', msg: "Semua data dalam file sudah pernah diupload (Duplikat). Tidak ada data baru ditambahkan." };
    }
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

// =====================================================
// MANAJEMEN KREDIT TOKO (Sederhana)
// =====================================================
function _ensureKreditTokoSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_KREDIT_TOKO);
    if (!s) {
    s = ss.insertSheet(SHEET_KREDIT_TOKO);
    // Header: Waktu, ID System, Nota Toko, Toko, Petugas, Nominal, No Anggota, Verifikasi, Status Notif
    s.getRange(1,1,1,9).setValues([['Waktu','ID System','Nota Toko','Toko','Petugas','Nominal','No Anggota','Verifikasi','Status Notif']]);
  }
  return s;
}

function saveKreditTokoFileToDrive(base64Data, fileName) {
  if (!base64Data || base64Data.trim() === "") return "";
  try {
    var folder = DriveApp.getFolderById(FOLDER_KREDIT_TOKO_ID);
    var splitBase = base64Data.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var rawBase64 = splitBase[1].replace(/[^A-Za-z0-9\+\/\=]/g, "");
    while (rawBase64.length % 4 !== 0) rawBase64 += "=";
    var blob = Utilities.newBlob(Utilities.base64Decode(rawBase64), type, fileName);
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return file.getId();
  } catch(e) { throw new Error("Gagal simpan lampiran kredit toko ke Drive: " + e.message); }
}

function addPiutangManual(noAnggota, notaToko, nilai, userToko, base64Photo, fileName) {
  try {
    if (!noAnggota) return { status: 'error', msg: 'Nomor anggota kosong.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sUsers = ss.getSheetByName(SHEET_USERS);
    var uData = sUsers.getDataRange().getValues();
    var memberNama = noAnggota;
    var memberKelompok = '-';
    var uLower = String(noAnggota).trim().toLowerCase();
    for (var ui = 1; ui < uData.length; ui++) {
      if (String(uData[ui][0]).trim().toLowerCase() === uLower) {
        memberNama = String(uData[ui][2] || noAnggota);
        memberKelompok = String(uData[ui][4] || '-');
        // Cek blokir piutang (kolom 24 / indeks 23)
        var blokir = String(uData[ui][23] || '').trim().toLowerCase();
        if (blokir === 'y' || blokir === 'yes' || blokir === '1' || blokir === 'blokir') {
          return { status: 'error', msg: '❌ Anggota ' + noAnggota + ' (' + memberNama + ') sedang diblokir untuk transaksi piutang.' };
        }
        break;
      }
    }
    
    var sPi = _ensureKreditTokoSheet();
    var tz = Session.getScriptTimeZone();
    var waktu = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
    var idSys = 'PIU-' + new Date().getTime().toString().slice(-8);
    var valNum = Number(String(nilai).replace(/[^0-9\-\\.]/g,'')) || 0;

    if (_isLimitKreditAktif()) {
      var batas = _getBatasKreditAnggota(noAnggota);
      var hutangBulanIni = _getKreditTokoBulanIni(noAnggota);
      if (batas.effective <= 0) {
        return { status: 'error', msg: '❌ Kredit Toko diblokir: sisa tagihan bulan lalu (Rp ' + _formatRp(batas.outstanding) + ') sudah mencapai limit. Segera bayar tagihan agar dapat kredit baru.' };
      }
      if (hutangBulanIni + valNum > batas.effective) {
        return { status: 'error', msg: 'Limit piutang bulan ini tidak mencukupi! Batas (limit global ' + _formatRp(batas.global) + ' - sisa tagihan ' + _formatRp(batas.outstanding) + '): Rp ' + _formatRp(batas.effective) + '. Sisa: Rp ' + _formatRp(batas.effective - hutangBulanIni) };
      }
    }

    var fileId = '';
    if (base64Photo && base64Photo.length > 50) {
      var fn = fileName || ('piutang_' + idSys + '.jpg');
      fileId = saveKreditTokoFileToDrive(base64Photo, fn);
    }
    // Append in order: Waktu, ID System, Nota Toko, Toko, Petugas, Nominal, No Anggota, Verifikasi
    sPi.appendRow([waktu, idSys, notaToko || '-', userToko ? userToko.nama : '-', userToko ? userToko.user : '-', valNum, "'" + noAnggota, fileId]);
    _invalidateKreditCaches();
    var nota = {
      toko: userToko ? userToko.nama : '-',
      waktu: waktu,
      notaToko: notaToko || '-',
      idSystem: idSys,
      kode: idSys,
      nama: memberNama,
      noAnggota: noAnggota,
      nilai: 'Rp ' + valNum.toLocaleString('id-ID'),
      kasir: userToko ? userToko.user : '-',
      kelompok: memberKelompok,
      piutangId: idSys
    };
    return { status: 'sukses', msg: 'Piutang dicatat.', nota: nota };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

// Import Excel tagihan piutang (bulk) - admin. Nota Toko opsional (isi '-'), Waktu bisa dari bulan lalu.
function importPiutangTagihanAdmin(dataArray) {
  try {
    if (!dataArray || dataArray.length < 2) return { status: 'error', msg: 'File kosong / tidak ada baris data.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sPi = _ensureKreditTokoSheet();
    var sUs = ss.getSheetByName(SHEET_USERS);
    var uData = sUs ? sUs.getDataRange().getValues() : [];
    var userMap = {};
    for (var ui = 1; ui < uData.length; ui++) {
      var k = String(uData[ui][0]).replace(/'/g, '').trim();
      if (k) userMap[k] = String(uData[ui][2] || k);
    }
    var tz = Session.getScriptTimeZone();
    var imported = 0;
    var skipped = [];
    for (var i = 1; i < dataArray.length; i++) {
      var r = dataArray[i];
      if (!r) continue;
      var noAnggota = String(r[1] || '').replace(/'/g, '').replace(/[^0-9]/g, '').trim();
      if (!noAnggota) { skipped.push({ row: i + 1, reason: 'No Anggota kosong' }); continue; }
      if (!userMap[noAnggota]) { skipped.push({ row: i + 1, reason: 'Anggota ' + noAnggota + ' tidak ditemukan di Users' }); continue; }
      var nilai = Number(String(r[2] || '').replace(/[^0-9\-\.]/g, '')) || 0;
      if (nilai <= 0) { skipped.push({ row: i + 1, reason: 'Nilai <= 0 / tidak valid' }); continue; }
      var tglObj = _parseDate(String(r[0] || ''));
      if (!tglObj) { skipped.push({ row: i + 1, reason: 'Tanggal tidak valid (pakai dd/MM/yyyy)' }); continue; }
      var nota = String(r[3] || '').trim() || '-';
      var toko = String(r[4] || '').trim() || '-';
      var petugas = String(r[5] || '').trim() || '-';
      var idSys = String(r[6] || '').trim();
      if (idSys === '' || idSys === '-') idSys = 'PIU-' + (new Date().getTime() + imported).toString().slice(-8);
      var waktu = Utilities.formatDate(tglObj, tz, "dd/MM/yyyy HH:mm:ss");
      sPi.appendRow([waktu, idSys, nota, toko, petugas, nilai, "'" + noAnggota, '']);
      imported++;
    }
    if (imported > 0) _invalidateKreditCaches();
    var msg = imported + ' tagihan piutang berhasil diimpor.';
    if (skipped.length) msg += ' ' + skipped.length + ' baris dilewati.';
    return { status: imported > 0 ? 'sukses' : 'error', msg: msg, skipped: skipped };
  } catch (e) { return { status: 'error', msg: "Gagal: " + e.toString() }; }
}

function getKreditTokoForMember(noAnggota) {
  try {
    var data = _getCachedPiutangData();
    var items = []; var total = 0;
    var tz = Session.getScriptTimeZone();
    var memberBayar = _getCachedMemberPaymentMap();
    var remaining = memberBayar[String(noAnggota).trim().toLowerCase()] || 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][6]).replace(/'/g,'').trim() === String(noAnggota).trim()) {
        var nilai = Number(data[i][5]) || 0;
        var idSys = String(data[i][1]);
        var used = Math.min(nilai, remaining);
        remaining -= used;
        var sisa = nilai - used;
        if (sisa < 0) sisa = 0;
        var fileId = String(data[i][7] || '');
        var tglRaw = data[i][0];
        var tglObj = _parseDate(tglRaw);
        var tglStr = tglObj ? Utilities.formatDate(tglObj, tz, "dd/MM/yyyy HH:mm") : String(tglRaw);

        items.push({ rowIdx: i+1, waktu: tglStr, waktuRaw: tglObj ? tglObj.getTime() : 0, id: idSys, nota: String(data[i][2]), toko: String(data[i][3]), petugas: String(data[i][4]), nilai: nilai, dibayar: used, sisa: sisa, noAnggota: String(data[i][6]), verifikasiFileId: fileId });
        total += sisa;
      }
    }
    var batasM = _getBatasKreditAnggota(noAnggota);
    // Biaya jasa 1,5%/bulan atas sisa belum lunas (masuk ke sisa tagihan)
    var feeRes = _simulasiBiayaJasaMember(String(noAnggota).trim().toLowerCase(), _buildPiutangEvents(data), _getCachedPaymentRows(), new Date());
    return { status: 'sukses', totalOutstanding: feeRes.balance, biayaJasa: feeRes.feeTotal, limitBulanIni: batasM.effective, limitGlobal: batasM.global, outstanding: batasM.outstanding, items: items };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function prosesVoucherPiutang(kode, userToko, notaToko, base64Photo, fileName, nominalOverride) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sVoucher = ss.getSheetByName(SHEET_VOUCHERS);
    var dVoucher = sVoucher.getDataRange().getValues();
    var sLaporan = ss.getSheetByName(SHEET_LAPORAN);
    var sPi = _ensureKreditTokoSheet();
    var vRow = -1; var vData = null;
    for (var i = 1; i < dVoucher.length; i++) { if (String(dVoucher[i][0]).trim() === String(kode).trim()) { vRow = i + 1; vData = dVoucher[i]; break; } }
    if (vRow === -1) return { status: 'error', msg: '❌ Voucher Tidak Ditemukan!' };

    var noAnggota = vData[1];
    // Cek blokir piutang
    if (_isPiutangBlocked(noAnggota)) {
      return { status: 'error', msg: '❌ Anggota ' + noAnggota + ' (' + vData[2] + ') sedang diblokir untuk transaksi piutang.' };
    }
    var valNum = Number(nominalOverride !== undefined && nominalOverride !== null ? nominalOverride : vData[5]) || 0;
    if (_isLimitKreditAktif()) {
      var batas = _getBatasKreditAnggota(noAnggota);
      var hutangBulanIni = _getKreditTokoBulanIni(noAnggota);
      if (batas.effective <= 0) return { status: 'error', msg: '❌ Kredit Toko diblokir: sisa tagihan bulan lalu (Rp ' + _formatRp(batas.outstanding) + ') sudah mencapai limit. Segera bayar tagihan agar dapat kredit baru.' };
      if (hutangBulanIni + valNum > batas.effective) return { status: 'error', msg: 'Limit piutang bulan ini tidak mencukupi! Batas (limit global ' + _formatRp(batas.global) + ' - sisa tagihan ' + _formatRp(batas.outstanding) + '): Rp ' + _formatRp(batas.effective) + '. Sisa: Rp ' + _formatRp(batas.effective - hutangBulanIni) };
    }

    var tz = Session.getScriptTimeZone();
    var st = String(vData[6]).trim();
    var ex = _parseDate(vData[7]); var startD = _parseDate(vData[8]); var today = new Date(); today.setHours(0,0,0,0);
    if (ex) ex.setHours(23,59,59,999); if (startD) startD.setHours(0,0,0,0);
    var isExp = (ex && today > ex); var isEarly = (startD && today < startD);
    if (st === 'Used') return { status: 'sudah_pakai_struk' };
    if (st !== 'Active' || isExp || isEarly) return { status: 'error', msg: '❌ Voucher tidak dapat digunakan (Belum Aktif / Kadaluarsa).' };

    var timestamp = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss");
    var idSys = "TRX-" + new Date().getTime().toString().slice(-6);
    // tandai voucher terpakai
    sVoucher.getRange(vRow, 7).setValue('Used');
    // simpan sebagai piutang
    var idPi = 'PIU-' + (new Date().getTime() + 2000).toString().slice(-8); // ID Transaksi Piutang
    // simpan di laporan mutasi dengan ID gabungan
    sLaporan.appendRow([timestamp, idSys + " / " + idPi, "'" + notaToko, "'" + vData[0], userToko.nama, "'" + userToko.user, valNum, vData[2], "'" + vData[1], vData[4]]);

    var fileId = '';
    if (base64Photo && base64Photo.length > 50) {
      var fn = fileName || ('piu_' + idPi + '.jpg');
      fileId = saveKreditTokoFileToDrive(base64Photo, fn);
    }
    // Append in order: Waktu, ID System, Nota Toko, Toko, Petugas, Nominal, No Anggota, Verifikasi
    sPi.appendRow([timestamp, idPi, notaToko || '-', userToko ? userToko.nama : '-', userToko ? userToko.user : '-', valNum, "'" + vData[1], fileId]);
    _invalidateKreditCaches();

    return { status: 'sukses', nota: { toko: userToko.nama, waktu: timestamp, notaToko: notaToko, idSystem: idSys + " / " + idPi, kasir: userToko.user, kode: vData[0], nama: vData[2], noAnggota: vData[1], kelompok: vData[4], masaBerlaku: (startD ? Utilities.formatDate(startD, tz, "dd/MM/yyyy") : "-") + " s/d " + (ex ? Utilities.formatDate(ex, tz, "dd/MM/yyyy") : "-"), nilai: "Rp " + _formatRp(valNum) } };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function getSemuaPiutangAdmin(bln, thn, startDate, endDate, q) {
  try {
    var dataPiutang = _getCachedPiutangData();
    if (dataPiutang.length < 2) return { status: 'sukses', data: [] };

    var dataUsers = _getCachedUsersData();
    var userMap = {};
    var hpMap = {};
    var emailMap = {};
    for (var j = 1; j < dataUsers.length; j++) {
      var key = String(dataUsers[j][0]).trim();
      userMap[key] = String(dataUsers[j][2] || "-");
      hpMap[key] = String(dataUsers[j][10] || "");
      emailMap[key] = String(dataUsers[j][9] || "");
    }

    var result = [];
    var tz = Session.getScriptTimeZone();
    var memberBayar = _getCachedMemberPaymentMap();
    var limitMap = _getBatasKreditAllMembers(dataPiutang);
    var globalLimit = _getGlobalLimitCached();

    // Normalisasi filter periode (diterapkan server-side untuk kurangi payload)
    var fBln = (bln === undefined || bln === null || bln === '') ? '' : String(bln).replace(/\D/g, '');
    var fThn = (thn === undefined || thn === null) ? '' : String(thn).trim();
    var fSd = (startDate === undefined || startDate === null) ? '' : String(startDate).trim();
    var fEd = (endDate === undefined || endDate === null) ? '' : String(endDate).trim();
    var fQ = (q === undefined || q === null) ? '' : String(q).trim().toLowerCase();

    function _matchPeriod(waktuStr) {
      var wp = String(waktuStr || '').split(' ')[0].split('/');
      if (wp.length === 3) {
        if (fBln !== '' && fBln !== 'ALL' && parseInt(wp[1], 10) !== parseInt(fBln, 10)) return false;
        if (fThn !== '' && wp[2] !== fThn) return false;
        var dNorm = wp[2] + '-' + wp[1] + '-' + wp[0];
        if (fSd !== '' && dNorm < fSd) return false;
        if (fEd !== '' && dNorm > fEd) return false;
        return true;
      }
      return (fBln === '' && fThn === '' && fSd === '' && fEd === '');
    }

    // Header sheet Piutang: 0:Waktu, 1:ID System, 2:Nota Toko, 3:Toko, 4:Petugas, 5:Nominal, 6:No Anggota, 7:Verifikasi, 8:Status Notif
    // Alokasikan pembayaran akumulatif per anggota ke baris piutang berurutan (FIFO)
    var allocRemain = {};
    for (var i = 1; i < dataPiutang.length; i++) {
      var row = dataPiutang[i];
      var noAnggota = String(row[6]).replace(/'/g, "").trim();
      var nKey = noAnggota.toLowerCase();
      var nilaiRow = Number(row[5]) || 0;
      if (allocRemain[nKey] === undefined) allocRemain[nKey] = memberBayar[nKey] || 0;

      var tglRaw = row[0];
      var tglObj = _parseDate(tglRaw);
      var waktuStr = tglObj ? Utilities.formatDate(tglObj, tz, "dd/MM/yyyy HH:mm") : String(tglRaw);

      var notif = String(row[8] || "").trim();
      if (notif === "") notif = "Belum";

      var used = Math.min(nilaiRow, allocRemain[nKey]);
      allocRemain[nKey] -= used;
      var sisa = nilaiRow - used;
      if (sisa < 0) sisa = 0;
      var sts = sisa <= 0 ? 'Lunas' : 'Belum Lunas';

      var limitInfo = limitMap[nKey];
      var effLimit = limitInfo ? limitInfo.effective : globalLimit;

      if (!_matchPeriod(waktuStr)) continue;

      if (fQ !== '') {
        var hayQ = (noAnggota + ' ' + (userMap[noAnggota] || '')).toLowerCase();
        if (hayQ.indexOf(fQ) === -1) continue;
      }

      var stNota = _storeFromNota(String(row[2]));

      result.push({
        waktu: waktuStr,
        waktuRaw: tglObj ? tglObj.getTime() : 0,
        id: String(row[1]).trim(),
        nota: String(row[2]),
        toko: String(row[3]),
        petugas: String(row[4]),
        nilai: nilaiRow,
        store: stNota.key,
        storeNama: stNota.nama,
        noAnggota: noAnggota,
        nama: userMap[noAnggota] || "-",
        noHp: hpMap[noAnggota] || "",
        email: emailMap[noAnggota] || "",
        status: sts,
        dibayar: used,
        sisa: sisa,
        notif: notif,
        verifikasiFileId: String(row[7] || ""),
        memberLimit: effLimit
      });
    }

    // Biaya jasa 1,5%/bulan atas sisa yang belum lunas (agregat admin)
    var piutangEventsF = _buildPiutangEvents(dataPiutang);
    var payRowsF = _getCachedPaymentRows();
    var asOfF = _asOfDateForRekap(fBln, fThn);
    var seenF = {};
    var totalBiayaJasaF = 0;
    for (var fi = 0; fi < result.length; fi++) {
      var fKey = String(result[fi].noAnggota || '').trim().toLowerCase();
      if (!fKey || seenF[fKey]) continue;
      seenF[fKey] = true;
      totalBiayaJasaF += _simulasiBiayaJasaMember(fKey, piutangEventsF, payRowsF, asOfF).feeTotal;
    }
    var totalPiutangF = 0;
    for (var fj = 0; fj < result.length; fj++) totalPiutangF += (Number(result[fj].sisa) || 0);

    return { status: 'sukses', data: result, totalBiayaJasa: totalBiayaJasaF, totalTagihan: totalPiutangF + totalBiayaJasaF };
  } catch (e) {
    return { status: 'error', msg: "Server Error: " + e.toString() };
  }
}

// =====================================================
// BIAYA JASA KREDIT TOKO (1,5% PER BULAN ATAS SISA BELUM LUNAS)
// =====================================================
function _getBiayaJasaRate() { return 0.015; }

// Toko asal nota berdasarkan awalan kode nota: prefix TKST = Toko Sutomo, selainnya = Toko INKA.
function _storeFromNota(notaRaw) {
  var n = String(notaRaw || '').replace(/^'/, '').trim().toUpperCase();
  if (n.indexOf('TKST') === 0) return { key: 'SUTOMO', nama: 'Toko Sutomo' };
  return { key: 'INKA', nama: 'Toko INKA' };
}

// Normalisasi seluruh baris piutang menjadi event kredit per anggota (urut waktu)
function _buildPiutangEvents(dataPiutang) {
  var events = [];
  for (var i = 1; i < dataPiutang.length; i++) {
    var row = dataPiutang[i];
    var noAng = String(row[6]).replace(/'/g, '').trim().toLowerCase();
    var tglObj = _parseDate(row[0]);
    var nilai = Number(row[5]) || 0;
    if (!noAng || !tglObj || nilai <= 0) continue;
    events.push({ noAng: noAng, t: tglObj.getTime(), val: nilai });
  }
  events.sort(function (a, b) { return a.t - b.t; });
  return events;
}

// Simulasi ledger kredit toko per anggota: kredit (+), pembayaran (-),
// di tiap pergantian bulan saldo sisa yang belum lunas dikenakan biaya jasa 1,5%
// (masuk ke sisa tagihan, berakumulasi/compounding) sampai bulan asOf (inclusive).
function _simulasiBiayaJasaMember(noAngLc, piutangEvents, payRows, asOfDate) {
  var events = [];
  for (var i = 0; i < piutangEvents.length; i++) {
    if (piutangEvents[i].noAng === noAngLc) events.push({ t: piutangEvents[i].t, val: piutangEvents[i].val });
  }
  for (var j = 0; j < payRows.length; j++) {
    if (payRows[j].noAng === noAngLc) events.push({ t: payRows[j].t, val: -payRows[j].jml });
  }
  if (events.length === 0) return { balance: 0, principal: 0, feeTotal: 0, feeEvents: [] };

  // Urutkan naik; pada tanggal yang sama kredit diproses dahulu
  events.sort(function (a, b) { if (a.t !== b.t) return a.t - b.t; return a.val > 0 ? -1 : 1; });

  var asOfY = asOfDate.getFullYear(), asOfM = asOfDate.getMonth();
  var balance = 0, feeTotal = 0, feeEvents = [];
  var evIdx = 0;
  var firstD = new Date(events[0].t);
  var year = firstD.getFullYear(), month = firstD.getMonth();

  while (year < asOfY || (year === asOfY && month <= asOfM)) {
    var mNext = new Date(year, month + 1, 1).getTime();
    while (evIdx < events.length && events[evIdx].t < mNext) {
      if (events[evIdx].val > 0) balance += events[evIdx].val;
      else balance = Math.max(0, balance + events[evIdx].val);
      evIdx++;
    }
    // Biaya jasa untuk sisa yang dibawa ke bulan berikutnya (hanya bulan yang sudah terlewati)
    if ((year < asOfY || (year === asOfY && month < asOfM)) && balance > 0) {
      var fee = Math.round(balance * _getBiayaJasaRate());
      if (fee > 0) {
        feeTotal += fee;
        balance += fee;
        var ny = year, nm = month + 1;
        if (nm === 12) { nm = 0; ny++; }
        feeEvents.push({ y: ny, m: nm, fee: fee });
      }
    }
    month++;
    if (month === 12) { month = 0; year++; }
    if (evIdx >= events.length && (year > asOfY || (year === asOfY && month > asOfM))) break;
  }

  var principal = balance - feeTotal;
  if (principal < 0) principal = 0;
  return { balance: balance, principal: principal, feeTotal: feeTotal, feeEvents: feeEvents };
}

// Tanggal as-of untuk snapshpot rekap: akhir bulan terpilih atau sekarang
function _asOfDateForRekap(fBulan, fTahun) {
  if (fTahun !== '' && /^\d+$/.test(String(fTahun))) {
    var y = parseInt(String(fTahun), 10);
    if (fBulan !== '' && fBulan !== 'ALL' && /^\d+$/.test(String(fBulan))) {
      var m = parseInt(String(fBulan), 10);
      if (m >= 1 && m <= 12) return new Date(y, m - 1, 15);
    }
    return new Date(y, 11, 31);
  }
  return new Date();
}

// =====================================================
// REKAP KREDIT PER ANGOTA (AGREGASI PER BULAN)
// =====================================================
function getRekapKreditPerAnggota(bulan, tahun) {
  try {
    var dataPiutang = _getCachedPiutangData();
    var dataUsers = _getCachedUsersData();

    // Daftar seluruh anggota dari sheet Users (role = Anggota) supaya yang
    // tidak punya piutang di periode terpilih tetap tampil dengan nominal 0.
    var userMap = {};
    for (var j = 1; j < dataUsers.length; j++) {
      var roleUser = String(dataUsers[j][3] || 'Anggota').trim();
      if (roleUser.toLowerCase() !== 'anggota') continue;
      var key = String(dataUsers[j][0]).trim();
      if (!key) continue;
      userMap[key] = {
        nama: String(dataUsers[j][2] || "-"),
        kelompok: String(dataUsers[j][4] || "-"),
        hp: String(dataUsers[j][10] || ""),
        email: String(dataUsers[j][9] || ""),
        statusAnggota: String(dataUsers[j][14] || 'Aktif')
      };
    }

    var memberBayar = _getCachedMemberPaymentMap();

    // Normalisasi filter
    var fBulan = (bulan === undefined || bulan === null || bulan === '') ? '' : String(bulan).replace(/\D/g, '');
    var fTahun = (tahun === undefined || tahun === null) ? '' : String(tahun).trim();

    // Seluruh baris piutang diurutkan per anggota berdasarkan tanggal (untuk alokasi FIFO)
    var rowsByMember = {};
    for (var i = 1; i < dataPiutang.length; i++) {
      var row = dataPiutang[i];
      var noAnggota = String(row[6]).replace(/'/g, "").trim();
      if (!noAnggota) continue;
      var tglObj = _parseDate(row[0]);
      if (!tglObj) continue;
      if (!rowsByMember[noAnggota]) rowsByMember[noAnggota] = [];
      rowsByMember[noAnggota].push({ tgl: tglObj, tglTime: tglObj.getTime(), nilai: Number(row[5]) || 0, nota: String(row[2] || '') });
    }

    // Gabungan seluruh anggota (userMap) + anggota yang ada piutang tapi tidak
    // terdaftar di sheet Users (data yatim piatu), supaya tidak ada yang hilang.
    var allMembers = {};
    for (var u in userMap) {
      if (userMap.hasOwnProperty(u)) allMembers[u] = true;
    }
    for (var p in rowsByMember) {
      if (rowsByMember.hasOwnProperty(p)) allMembers[p] = true;
    }

    // Persiapan simulasi biaya jasa (1,5%/bulan atas sisa yang belum lunas)
    var piutangEvents = _buildPiutangEvents(dataPiutang);
    var payRows = _getCachedPaymentRows();
    var asOf = _asOfDateForRekap(fBulan, fTahun);

    // Agregasi per anggota
    var aggMap = {};
    for (var noAng in allMembers) {
      if (!allMembers.hasOwnProperty(noAng)) continue;

      var rows = rowsByMember[noAng] || [];
      // Urutkan baris naik (tertua dulu) supaya alokasi FIFO benar
      rows.sort(function (a, b) { return a.tglTime - b.tglTime; });

      var allocRemain = memberBayar[noAng.toLowerCase()] || 0;
      var transByStore = {};
      var periodeTotal = 0, periodeTransaksi = 0, periodeDibayar = 0;

      for (var k = 0; k < rows.length; k++) {
        var r = rows[k];
        var rowMonth = r.tgl.getMonth() + 1;
        var rowYear = r.tgl.getFullYear();

        var used = Math.min(r.nilai, allocRemain);
        allocRemain -= used;
        var stKey = _storeFromNota(r.nota).key;

        // Hanya masukkan ke rekap bila baris ada di periode yang dipilih
        var inPeriod = true;
        if (fBulan !== '' && fBulan !== 'ALL' && rowMonth !== parseInt(fBulan, 10)) inPeriod = false;
        if (fTahun !== '' && String(rowYear) !== fTahun) inPeriod = false;

        if (inPeriod) {
          transByStore[stKey] = (transByStore[stKey] || 0) + r.nilai;
          periodeTotal += r.nilai;
          periodeTransaksi += 1;
          periodeDibayar += used;
        }
      }

      var periodeSisa = periodeTotal - periodeDibayar;
      if (periodeSisa < 0) periodeSisa = 0;

      // Biaya jasa 1,5%/bulan atas sisa tagihan yang belum lunas (snapshot s/d akhir periode)
      var feeRes = _simulasiBiayaJasaMember(noAng.toLowerCase(), piutangEvents, payRows, asOf);

      var transSutomo = transByStore['SUTOMO'] || 0;
      var transInka = transByStore['INKA'] || 0;
      // Biaya jasa per toko = 1,5% dari nilai transaksi toko tersebut
      var rateJasa = _getBiayaJasaRate();
      var jasaSutomo = Math.round(transSutomo * rateJasa);
      var jasaInka = Math.round(transInka * rateJasa);

      var status;
      if (periodeTransaksi <= 0 && feeRes.balance <= 0) status = 'Tidak Ada Transaksi';
      else status = feeRes.balance <= 0 ? 'Lunas' : 'Belum Lunas';

      var info = userMap[noAng] || { nama: '-', kelompok: '-', hp: '', email: '', statusAnggota: '' };
      aggMap[noAng] = {
        noAnggota: noAng,
        nama: info.nama,
        kelompok: info.kelompok,
        noHp: info.hp,
        email: info.email,
        jumlahTransaksi: periodeTransaksi,
        totalNilai: periodeTotal,
        totalDibayar: periodeDibayar,
        sisaPiutang: feeRes.principal,
        biayaJasa: jasaSutomo + jasaInka,
        sisa: feeRes.balance,
        transSutomo: transSutomo,
        jasaSutomo: jasaSutomo,
        transInka: transInka,
        jasaInka: jasaInka,
        total: transSutomo + jasaSutomo + transInka + jasaInka,
        status: status
      };
    }

    var result = [];
    for (var no in aggMap) {
      if (aggMap.hasOwnProperty(no)) result.push(aggMap[no]);
    }
    result.sort(function (a, b) { return b.totalNilai - a.totalNilai; });

    return { status: 'sukses', data: result };
  } catch (e) {
    return { status: 'error', msg: "Server Error: " + e.toString() };
  }
}

// Rekap pembayaran piutang toko (dari sheet PembayaranPiutang)
function getRekapPembayaran() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sBayar = _ensurePembayaranPiutangSheet();
    var dataBayar = sBayar.getDataRange().getValues();

    var sUs = ss.getSheetByName(SHEET_USERS);
    var dataUs = sUs.getDataRange().getValues();
    var userMap = {};
    for (var j = 1; j < dataUs.length; j++) {
      var key = String(dataUs[j][0]).trim();
      userMap[key] = { nama: String(dataUs[j][2] || "-"), kelompok: String(dataUs[j][4] || "-"), hp: String(dataUs[j][10] || "") };
    }

    // Total piutang per No Anggota
    var sPi = _ensureKreditTokoSheet();
    var dataPi = sPi.getDataRange().getValues();
    var piutangByMember = {};
    for (var i = 1; i < dataPi.length; i++) {
      var nA = String(dataPi[i][6]).replace(/'/g, '').trim().toLowerCase();
      if (!nA) continue;
      if (!piutangByMember[nA]) piutangByMember[nA] = 0;
      piutangByMember[nA] += Number(dataPi[i][5]) || 0;
    }

    // Total bayar per No Anggota
    var bayarByMember = {};
    for (var k = 1; k < dataBayar.length; k++) {
      var nB = String(dataBayar[k][1]).replace(/'/g, '').trim().toLowerCase();
      if (!nB) continue;
      if (!bayarByMember[nB]) bayarByMember[nB] = 0;
      bayarByMember[nB] += Number(dataBayar[k][2]) || 0;
    }

    var result = [];
    var tz = Session.getScriptTimeZone();
    for (var k = 1; k < dataBayar.length; k++) {
      var row = dataBayar[k];
      var tglObj = _parseDate(row[0]);
      var noAng = String(row[1]).replace(/'/g, '').trim();
      var noAngKey = noAng.toLowerCase();
      var totalPiutang = piutangByMember[noAngKey] || 0;
      var totalBayar = bayarByMember[noAngKey] || 0;
      var isLunas = totalBayar >= totalPiutang && totalPiutang > 0;
      result.push({
        waktu: tglObj ? Utilities.formatDate(tglObj, tz, "dd/MM/yyyy") : String(row[0] || ""),
        waktuRaw: tglObj ? tglObj.getTime() : 0,
        noAnggota: noAng,
        nama: userMap[noAng] ? userMap[noAng].nama : "-",
        kelompok: userMap[noAng] ? userMap[noAng].kelompok : "-",
        noHp: userMap[noAng] ? userMap[noAng].hp : "",
        nota: "-",
        nilaiPiutang: totalPiutang,
        jumlah: Number(row[2]) || 0,
        notaBayar: String(row[3] || "-"),
        metode: String(row[4] || "-"),
        status: isLunas ? 'Lunas' : String(row[5] || "Masuk")
      });
    }
    result.sort(function (a, b) { return String(a.waktuRaw) < String(b.waktuRaw) ? 1 : -1; });
    return { status: 'sukses', data: result };
  } catch (e) {
    return { status: 'error', msg: "Server Error: " + e.toString() };
  }
}

function kirimEmailNotifHutang(dataKirim) {
  try {
    if (!dataKirim || dataKirim.length === 0) return { status: 'error', msg: 'Data kirim kosong.' };
    var tz = Session.getScriptTimeZone();
    var tanggal = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
    var sukses = 0; var gagal = 0;
    
    for (var i = 0; i < dataKirim.length; i++) {
      var item = dataKirim[i];
      var email = item.email.trim();
      if (!email || email.indexOf('@') === -1) { gagal++; continue; }
      
      var nama = item.nama || 'Anggota';
      var total = Number(item.total) || 0;
      var count = item.count || 0;
      
      var subject = 'Pemberitahuan Kredit Toko - MyKopinka per ' + tanggal;
      var body = 'Yth. Bapak/Ibu ' + nama + ',\n\n';
      body += 'Berdasarkan data per ' + tanggal + ', tercatat informasi sebagai berikut:\n\n';
      body += '• No Anggota: ' + (item.noAnggota || '-') + '\n';
      
      // Jika ada detail per-transaksi (dari kirimEmailHutangSingle)
      if (item.waktu) {
        body += '• Waktu: ' + item.waktu + '\n';
        body += '• No Nota: ' + item.nota + '\n';
        body += '• Toko: ' + item.toko + '\n';
        body += '• Petugas: ' + item.petugas + '\n';
        body += '• ID Sistem: ' + item.idSistem + '\n';
        body += '• Nominal: Rp ' + _formatRp(total) + '\n\n';
      }
      
      body += '• Total Kredit: Rp ' + _formatRp(total) + '\n';
      body += '• Jumlah Transaksi: ' + count + '\n\n';
      body += 'Apabila transaksi tersebut dilakukan oleh Anda, silakan abaikan pemberitahuan ini.\n\n';
      body += 'Apabila Anda tidak melakukan transaksi tersebut, mohon segera menghubungi kami untuk verifikasi lebih lanjut.\n\n';
      body += 'Informasi detail dapat diakses melalui mykopinka.kopinka.com.\n\n';
      body += 'Terima kasih.\n\n';
      body += '_MyKopinka - Koperasi Pegawai INKA_';
      
      MailApp.sendEmail(email, subject, body, {name: 'MyKopinka'});
      sukses++;
    }
    
    return { status: 'sukses', msg: 'Email berhasil dikirim ke ' + sukses + ' anggota' + (gagal > 0 ? (', ' + gagal + ' gagal (email tidak valid).') : '.') };
  } catch (e) {
    return { status: 'error', msg: 'Gagal kirim email: ' + e.toString() };
  }
}

function markPiutangNotif(idPiutang, status) {
  try {
    var statusFinal = (status === 'Terkirim' || status === 'Sudah Terkirim') ? 'Terkirim' : 'Belum';
    var sPi = _ensureKreditTokoSheet();
    var data = sPi.getDataRange().getValues();
    var count = 0;
    var target = String(idPiutang).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][1] || "").trim().toLowerCase();
      var match = (rowId === target) || (rowId.indexOf(" / " + target) !== -1) || (rowId.indexOf("/" + target) !== -1);
      if (match) {
        sPi.getRange(i + 1, 9).setValue(statusFinal);
        count++;
      }
    }
    if (count > 0) _invalidateKreditCaches();
    return { status: count > 0 ? 'sukses' : 'error', msg: count > 0 ? 'Status notifikasi diperbarui.' : 'Transaksi tidak ditemukan.' };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function markPiutangNotifBulk(ids, status) {
  try {
    var statusFinal = (status === 'Terkirim' || status === 'Sudah Terkirim') ? 'Terkirim' : 'Belum';
    var sPi = _ensureKreditTokoSheet();
    var data = sPi.getDataRange().getValues();
    var set = {};
    for (var k = 0; k < ids.length; k++) {
      set[String(ids[k]).trim().toLowerCase()] = true;
    }
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][1] || "").trim().toLowerCase();
      if (set[rowId]) {
        sPi.getRange(i + 1, 9).setValue(statusFinal);
        count++;
      }
    }
    if (count > 0) _invalidateKreditCaches();
    return { status: count > 0 ? 'sukses' : 'error', msg: count > 0 ? 'Status notifikasi (' + count + ') diperbarui.' : 'Tidak ada transaksi yang cocok.' };
  } catch (e) { return { status: 'error', msg: e.toString() }; }
}

function kirimBackupHutangExcel(base64Excel, namaFile, emailAnggota) {
  try {
    if (!base64Excel || !namaFile || !emailAnggota) {
      return { status: 'error', msg: 'Data tidak lengkap.' };
    }
    
    // Decode base64 Excel
    var rawBase64 = base64Excel.replace(/[^A-Za-z0-9\+\/\=]/g, "");
    while (rawBase64.length % 4 !== 0) rawBase64 += "=";
    var blob = Utilities.newBlob(Utilities.base64Decode(rawBase64), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', namaFile);
    
    var tz = Session.getScriptTimeZone();
    var tanggal = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
    
    var subject = 'Backup Data Kredit Toko - MyKopinka per ' + tanggal;
    var body = 'Yth. Anggota,\n\n';
    body += 'Terlampir file backup data kredit toko Anda per ' + tanggal + '.\n\n';
    body += 'Silakan unduh dan simpan file untuk keperluan pencatatan pribadi.\n\n';
    body += 'Terima kasih.\n\n';
    body += '_MyKopinka - Koperasi Pegawai INKA_';
    
    MailApp.sendEmail({
      to: emailAnggota,
      subject: subject,
      body: body,
      name: 'MyKopinka',
      attachments: [blob]
    });
    
    return { status: 'sukses', msg: 'File Excel berhasil dikirim ke ' + emailAnggota };
  } catch (e) {
    return { status: 'error', msg: 'Gagal kirim: ' + e.toString() };
  }
}

// =====================================================
// FUNGSI KHUSUS ADMIN (Cek Member & Update Status)
// =====================================================
function getMemberVoucherAdmin(username) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sUsers = ss.getSheetByName(SHEET_USERS);
    var uData = sUsers.getDataRange().getValues();
    var userInfo = null;
    var uLower = String(username).trim().toLowerCase();
    
    // Mencari informasi profil member
    for (var i = 1; i < uData.length; i++) {
      if (String(uData[i][0]).trim().toLowerCase() === uLower) {
        userInfo = {
          username: String(uData[i][0]),
          nama: String(uData[i][2] || "-"),
          kelompok: String(uData[i][4] || "-"),
          foto: String(uData[i][5] || "")
        };
        break;
      }
    }
    
    if (!userInfo) return { status: 'error', msg: 'Data member tidak ditemukan di database user.' };
    
    var sVouchers = ss.getSheetByName(SHEET_VOUCHERS);
    var vData = sVouchers.getDataRange().getValues();
    var vouchers = [];
    var tz = Session.getScriptTimeZone();
    var today = new Date(); today.setHours(0,0,0,0);
    
    // Mencari daftar voucher milik member
    for (var j = 1; j < vData.length; j++) {
      if (String(vData[j][1]).trim().toLowerCase() === uLower) {
        var ex = _parseDate(vData[j][7]);
        var startD = _parseDate(vData[j][8]);
        if(ex) ex.setHours(23,59,59,999);
        if(startD) startD.setHours(0,0,0,0);

        var st = String(vData[j][6]).trim();
        if (st === 'Active') {
          if (ex && today > ex) st = 'Expired';
          else if (startD && today < startD) st = 'Belum Aktif';
        }

        vouchers.push({
          kode: String(vData[j][0]),
          nilai: Number(vData[j][5]) || 0,
          status: st,
          expDate: ex ? Utilities.formatDate(ex, tz, "dd/MM/yyyy") : "-"
        });
      }
    }
    
    return { status: 'sukses', info: userInfo, vouchers: vouchers };
  } catch (e) {
    return { status: 'error', msg: "Server Error: " + e.toString() };
  }
}

function updateStatusVoucherMember(username, status) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS);
    var data = sheet.getDataRange().getValues();
    var uLower = String(username).trim().toLowerCase();
    var count = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === uLower) {
        // Hanya mengubah voucher yang belum terpakai
        if (String(data[i][6]).trim() !== 'Used') {
          sheet.getRange(i + 1, 7).setValue(status);
          count++;
        }
      }
    }
    return { status: 'sukses', msg: count + ' voucher milik ' + username + ' berhasil diubah menjadi ' + status + '.' };
  } catch (e) {
    return { status: 'error', msg: "Server Error: " + e.toString() };
  }
}

function updateStatusVouchers(arrKode, status) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VOUCHERS);
    var data = sheet.getDataRange().getValues();
    var count = 0;
    var kodes = arrKode.map(function(k) { return String(k).trim().toUpperCase(); });
    
    for (var i = 1; i < data.length; i++) {
      var rowKode = String(data[i][0]).trim().toUpperCase();
      if (kodes.indexOf(rowKode) !== -1) {
        if (String(data[i][6]).trim() !== 'Used') {
          sheet.getRange(i + 1, 7).setValue(status);
          count++;
        }
      }
    }
    return { status: 'sukses', msg: count + ' voucher berhasil diubah menjadi ' + status + '.' };
  } catch (e) { return { status: 'error', msg: "Server Error: " + e.toString() }; }
}

// =====================================================
// FUNGSI API UNTUK MENERIMA FOTO DARI CLOUDFLARE WORKERS
// =====================================================
function doPost(e) {
  try {
    // Membaca data JSON yang dikirim dari Cloudflare (Frontend)
    var requestData = JSON.parse(e.postData.contents);
    var fotoBase64 = requestData.foto;
    
    // Pastikan ada data foto yang dikirim
    if (!fotoBase64 || fotoBase64.trim() === "") {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', msg: 'Data foto kosong' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    var namaFile = "Foto_Kamera_" + new Date().getTime() + ".jpg";
    
    // Kita panggil fungsi saveFileToDrive yang SUDAH ADA di dalam code.gs kamu
    // Fungsi ini akan mengubah base64 menjadi gambar dan menyimpannya ke FOLDER_ID
    var fileUrl = saveFileToDrive(fotoBase64, namaFile);

    if (fileUrl) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'sukses', url: fileUrl }))
                           .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', msg: 'Gagal menyimpan file ke Drive' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', msg: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================
// PENGATURAN GLOBAL
// =====================================================
function getGlobalLimit() {
  return _getSetting('limit_piutang') || 0;
}

function saveGlobalLimit(value) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getSheetByName(SHEET_SETTINGS);
    if (!s) {
      s = ss.insertSheet(SHEET_SETTINGS);
      s.appendRow(["Key", "Value"]);
    }
    var data = s.getDataRange().getValues();
    var key = 'limit_piutang';
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === key) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex !== -1) {
      s.getRange(rowIndex, 2).setValue(value);
    } else {
      s.appendRow([key, value]);
    }
    _clearSettingCache('limit_piutang');
    _invalidateKreditCaches();
    return { status: 'sukses', msg: 'Limit piutang global berhasil diperbarui.' };
  } catch (e) {
    return { status: 'error', msg: e.toString() };
  }
}

// Controller tampilan informasi limit di panel anggota
function getShowLimitMember() {
  var v = _getSetting('show_limit_member');
  if (v === null || v === undefined) return true;
  return String(v).toLowerCase() === 'true' || v === '1' || v === 'true';
}

function saveShowLimitMember(value) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getSheetByName(SHEET_SETTINGS);
    if (!s) {
      s = ss.insertSheet(SHEET_SETTINGS);
      s.appendRow(["Key", "Value"]);
    }
    var data = s.getDataRange().getValues();
    var key = 'show_limit_member';
    var boolVal = (value === true || value === 'true' || value === 1 || value === '1') ? 'true' : 'false';
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === key) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex !== -1) {
      s.getRange(rowIndex, 2).setValue(boolVal);
    } else {
      s.appendRow([key, boolVal]);
    }
    _clearSettingCache('show_limit_member');
    return { status: 'sukses', msg: 'Setelan tampilan limit anggota disimpan.' };
  } catch (e) {
    return { status: 'error', msg: e.toString() };
  }
}

// Controller fitur limit kredit (aktif/nonaktif)
function getLimitKreditAktif() {
  return _isLimitKreditAktif();
}

function saveLimitKreditAktif(value) {
  try {
    var val = (value === true || value === 'true' || value === 1 || value === '1') ? 'true' : 'false';
    _setSetting('limit_kredit_aktif', val);
    _clearSettingCache('limit_kredit_aktif');
    _invalidateKreditCaches();
    return { status: 'sukses', msg: val === 'true' ? 'Fitur limit kredit diaktifkan.' : 'Fitur limit kredit dinonaktifkan.' };
  } catch (e) {
    return { status: 'error', msg: e.toString() };
  }
}

// =====================================================
// SETTING TEMPLATE PESAN WA
// =====================================================
function _setSetting(key, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_SETTINGS);
  if (!s) {
    s = ss.insertSheet(SHEET_SETTINGS);
    s.appendRow(["Key", "Value"]);
  }
  var data = s.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex !== -1) {
    s.getRange(rowIndex, 2).setValue(value);
  } else {
    s.appendRow([key, value]);
  }
  return true;
}

var WA_TEMPLATE_KEYS = ['wa_tpl_struk', 'wa_tpl_admin', 'wa_tpl_kasir', 'wa_tpl_group'];

function getWaTemplates() {
  try {
    var result = {};
    for (var i = 0; i < WA_TEMPLATE_KEYS.length; i++) {
      var v = _getSetting(WA_TEMPLATE_KEYS[i]);
      if (v !== null && v !== undefined && String(v).trim() !== '') {
        result[WA_TEMPLATE_KEYS[i]] = String(v);
      }
    }
    return { status: 'sukses', data: result };
  } catch (e) {
    return { status: 'error', msg: e.toString() };
  }
}

function saveWaSettings(updates) {
  try {
    if (!updates || typeof updates !== 'object') throw new Error('Data tidak valid.');
    var saved = 0;
    for (var key in updates) {
      if (WA_TEMPLATE_KEYS.indexOf(key) === -1) continue;
      var val = String(updates[key]);
      if (val.trim() === '') continue;
      _setSetting(key, val);
      saved++;
    }
    return { status: 'sukses', msg: saved + ' template pesan WA berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', msg: e.toString() };
  }
}

/**
 * Handles user logout.
 * In Google Apps Script, there isn't a traditional server-side session to destroy.
 * This function primarily serves to acknowledge the client-side logout request.
 * @returns {boolean} True if logout process is acknowledged.
 */
function logoutUser() {
  return true;
}

function debugPiutang(bln, thn, sd, ed) {
  try {
    var raw = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Piutang').getDataRange().getValues();
    var cached = _getCachedPiutangData();

    var fBln = (bln === undefined || bln === null || bln === '') ? '' : String(bln).replace(/\D/g, '');
    var fThn = (thn === undefined || thn === null) ? '' : String(thn).trim();
    var fSd = (sd === undefined || sd === null) ? '' : String(sd).trim();
    var fEd = (ed === undefined || ed === null) ? '' : String(ed).trim();

    function testRow(tglRaw) {
      var tglObj = _parseDate(tglRaw);
      var waktuStr = tglObj ? Utilities.formatDate(tglObj, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(tglRaw);
      var wp = String(waktuStr || '').split(' ')[0].split('/');
      var result = true;
      var detail = '';
      if (wp.length === 3) {
        if (fBln !== '' && fBln !== 'ALL' && parseInt(wp[1], 10) !== parseInt(fBln, 10)) { result = false; detail = 'bulan'; }
        if (result && fThn !== '' && wp[2] !== fThn) { result = false; detail = 'tahun (' + wp[2] + ' vs ' + fThn + ')'; }
        if (result) {
          var dNorm = wp[2] + '-' + wp[1] + '-' + wp[0];
          if (fSd !== '' && dNorm < fSd) { result = false; detail = 'start'; }
          if (result && fEd !== '' && dNorm > fEd) { result = false; detail = 'end'; }
        }
      } else {
        if (fBln !== '' || fThn !== '' || fSd !== '' || fEd !== '') { result = false; detail = 'format-waktu (' + waktuStr + ')'; }
      }
      return { waktuStr: waktuStr, wp: wp.join('/'), result: result, detail: detail };
    }

    var samples = [];
    var contohLolos = 0;
    var cacheAsal = cached || raw;
    for (var i = 1; i < Math.min(6, cacheAsal.length); i++) {
      var t = testRow(cacheAsal[i][0]);
      t.sumber = cached ? 'CACHED' : 'RAW';
      t.rawVal = String(cacheAsal[i][0]);
      samples.push(t);
      if (t.result) contohLolos++;
    }

    var realCall = null;
    try {
      var r = getSemuaPiutangAdmin(bln, thn, sd, ed);
      realCall = { status: r.status, dataCount: r.data ? r.data.length : 0, msg: (r.msg || '').substring(0, 200) };
    } catch (e2) { realCall = { err: e2.toString() }; }

    return {
      fBln: fBln, fThn: fThn, fSd: fSd, fEd: fEd,
      rawRowCount: raw.length,
      cachedRowCount: cached ? cached.length : 0,
      typeofWaktuCache: cached && cached[1] ? typeof cached[1][0] : 'N/A',
      sampleRows: samples,
      realCall: realCall,
      tz: Session.getScriptTimeZone()
    };
  } catch (e) {
    return { error: e.toString() };
  }
}