/**
 * Code.gs — Google Apps Script untuk Web Peta PMCI
 * Spreadsheet: Form_Report_PMCI_Medan-Binjai_2026
 *
 * CARA DEPLOY:
 * 1. Buka script.google.com → project ini
 * 2. Ganti seluruh isi dengan kode ini
 * 3. Deploy → Manage Deployments → Edit (ikon pensil) → Version: "New version" → Deploy
 * 4. Pastikan: Execute as = Me, Who has access = Anyone
 * 5. Salin URL deployment → update SHEET_URL di index.html jika URL berubah
 */

// ══════════════════════════════════════════════
// KONFIGURASI — sesuaikan jika nama sheet berbeda
// ══════════════════════════════════════════════
var SHEET_NAME      = 'Form Report';   // nama tab spreadsheet utama
var PLAN_SHEET_NAME = 'PlanColors';    // nama tab untuk menyimpan warna plan

// Index kolom (0-based), sesuai struktur spreadsheet:
// Baris 1 (index 0): No | Jenis | Kode Titik | Kode Modul | Surveyor | Tanggal Sampling | Waktu | Koordinat Plan | | Koordinat Actual | | Elevasi | Jarak Offset | ...
// Baris 2 (index 1): sub-header:  | | | | | | | UTM_XP | UTM_YP | UTM_XA | UTM_YA | ...
// Data mulai baris 3 (index 2)
var COL_KODE_TITIK = 2;   // "Kode Titik"     → DL001, W1, dll
var COL_TANGGAL    = 5;   // "Tanggal Sampling"
var COL_UTM_XP     = 7;   // "UTM_XP"  (koordinat plan)
var COL_UTM_YP     = 8;   // "UTM_YP"
var COL_UTM_XA     = 9;   // "UTM_XA"  (koordinat actual/realisasi)
var COL_UTM_YA     = 10;  // "UTM_YA"
var DATA_START_ROW = 2;   // index baris pertama data (skip 2 baris header)

// ══════════════════════════════════════════════
// doGet — dipanggil saat peta fetch data
// ══════════════════════════════════════════════
function doGet(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return jsonResponse({ status: 'error', message: 'Sheet "' + SHEET_NAME + '" tidak ditemukan. Cek nama tab di spreadsheet.' });
    }

    var data = sheet.getDataRange().getValues();
    var rows = [];

    for (var i = DATA_START_ROW; i < data.length; i++) {
      var row  = data[i];
      var nama = String(row[COL_KODE_TITIK] || '').trim();
      if (!nama || nama === '') continue;

      // Konversi tanggal: Apps Script otomatis parse Date dari sel spreadsheet
      var tanggal    = row[COL_TANGGAL];
      var tanggalStr = '';
      if (tanggal instanceof Date && !isNaN(tanggal)) {
        tanggalStr = Utilities.formatDate(tanggal, 'Asia/Jakarta', 'yyyy-MM-dd');
      } else if (tanggal !== null && tanggal !== undefined && String(tanggal).trim() !== '') {
        // Fallback: angka serial Excel atau string
        var strVal = String(tanggal).trim();
        if (/^\d+(\.\d+)?$/.test(strVal)) {
          // Excel serial number → konversi manual
          var excelSerial = parseFloat(strVal);
          var d = new Date((excelSerial - 25569) * 86400 * 1000);
          if (!isNaN(d)) {
            tanggalStr = Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd');
          }
        } else {
          tanggalStr = strVal;
        }
      }

      // Koordinat — bisa jadi angka atau string
      var utmXa = _toNum(row[COL_UTM_XA]);
      var utmYa = _toNum(row[COL_UTM_YA]);
      var utmXp = _toNum(row[COL_UTM_XP]);
      var utmYp = _toNum(row[COL_UTM_YP]);

      rows.push({
        nama:    nama,
        tanggal: tanggalStr,
        utm_xp:  utmXp,
        utm_yp:  utmYp,
        utm_xa:  utmXa,
        utm_ya:  utmYa
      });
    }

    // Ambil plan colors dari sheet PlanColors
    var planData = _readPlanColors(ss);

    return jsonResponse({ status: 'ok', data: rows, plan: planData });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ══════════════════════════════════════════════
// doPost — dipanggil saat peta simpan plan colors
// ══════════════════════════════════════════════
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    if (payload.action === 'savePlan') {
      var ss        = SpreadsheetApp.getActiveSpreadsheet();
      var planSheet = ss.getSheetByName(PLAN_SHEET_NAME);

      // Buat sheet PlanColors jika belum ada
      if (!planSheet) {
        planSheet = ss.insertSheet(PLAN_SHEET_NAME);
      }

      // Tulis ulang seluruh isi (header + data)
      planSheet.clearContents();
      planSheet.getRange(1, 1, 1, 2).setValues([['nama', 'color']]);

      var entries = Object.entries(payload.data || {});
      if (entries.length > 0) {
        planSheet.getRange(2, 1, entries.length, 2).setValues(entries);
      }

      return jsonResponse({ status: 'ok', saved: entries.length });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action: ' + (payload.action || 'undefined') });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ══════════════════════════════════════════════
// Helper: baca sheet PlanColors
// ══════════════════════════════════════════════
function _readPlanColors(ss) {
  var result = {};
  try {
    var planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    if (!planSheet) return result;
    var planData  = planSheet.getDataRange().getValues();
    for (var p = 1; p < planData.length; p++) {
      var pNama  = String(planData[p][0] || '').trim();
      var pColor = String(planData[p][1] || '').trim();
      if (pNama && pColor) result[pNama] = pColor;
    }
  } catch (e) {}
  return result;
}

// ══════════════════════════════════════════════
// Helper: konversi nilai sel ke angka, kosong jika gagal
// ══════════════════════════════════════════════
function _toNum(val) {
  if (val === null || val === undefined || val === '') return '';
  var n = parseFloat(String(val).replace(',', '.'));
  return isNaN(n) ? '' : n;
}

// ══════════════════════════════════════════════
// Helper: buat ContentService JSON response dengan CORS
// ══════════════════════════════════════════════
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
