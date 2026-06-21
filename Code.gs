/**
 * =====================================================================
 * APPS SCRIPT — Progres PMCI
 * Menghubungkan web peta (index.html, Leaflet) ke Google Sheet:
 * https://docs.google.com/spreadsheets/d/1jkaoAE5h-gRsAdkZG2akKDGD_Frc-Tgy1vEsJ3aWWHc
 *
 * KONTRAK DATA dengan frontend (jangan diubah tanpa update index.html juga):
 *  GET  -> { status:'ok', data:[{nama, tanggal, utm_xa, utm_ya}, ...], plan:{nama:warna,...} }
 *  POST -> body: { action:'savePlan', data:{nama:warna,...} }  => { status:'ok' }
 *
 * CATATAN PENTING (baca sebelum deploy):
 * - Kode Titik di sheet ini formatnya 3 digit (DL001, DL082, ...),
 *   sedangkan array TITIK di index.html formatnya TANPA leading-zero ekstra
 *   (DL01, DL82, DL100, ...). Fungsi normalizeKode() di bawah menangani
 *   konversi ini. Kalau nanti ada kode titik dengan pola lain (bukan
 *   [huruf][angka]), cek ulang fungsi ini.
 * - Saya asumsikan nama tab data = "Sheet1", header di baris 1-2, data
 *   mulai baris 3, dan kolom C/F/J/K = Kode Titik/Tanggal Sampling/UTM_XA/UTM_YA.
 *   Ini sudah saya verifikasi silang dengan REALISASI_AWAL di index.html
 *   untuk baris-baris yang sempat saya lihat (s.d. DL098). Tolong cek lagi
 *   untuk baris-baris setelahnya kalau strukturnya beda.
 * =====================================================================
 */

const CONFIG = {
  SHEET_NAME: 'Sheet1',        // nama tab data utama — ganti kalau beda
  PLAN_SHEET_NAME: 'Plan',     // tab penyimpanan warna "Plan Mode" (auto-dibuat kalau belum ada)
  HEADER_ROWS: 2,               // baris 1 = judul kolom, baris 2 = sub-judul (UTM_XP/YP/XA/YA)
  COL_KODE_TITIK: 3,             // kolom C
  COL_TANGGAL_SAMPLING: 6,      // kolom F
  COL_UTM_XA: 10,                // kolom J
  COL_UTM_YA: 11                 // kolom K
};

// ─────────────────────────────────────────────
// ENTRY POINTS
// ─────────────────────────────────────────────

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet "' + CONFIG.SHEET_NAME + '" tidak ditemukan. Cek nama tab di CONFIG.SHEET_NAME.');
    }

    const lastRow = sheet.getLastRow();
    const data = [];

    if (lastRow > CONFIG.HEADER_ROWS) {
      const numRows = lastRow - CONFIG.HEADER_ROWS;
      const values = sheet.getRange(CONFIG.HEADER_ROWS + 1, 1, numRows, CONFIG.COL_UTM_YA).getValues();

      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const kodeRaw = row[CONFIG.COL_KODE_TITIK - 1];
        if (!kodeRaw || String(kodeRaw).trim() === '') continue; // skip baris kosong

        data.push({
          nama: normalizeKode(kodeRaw),
          tanggal: formatTanggal(row[CONFIG.COL_TANGGAL_SAMPLING - 1]),
          utm_xa: row[CONFIG.COL_UTM_XA - 1],
          utm_ya: row[CONFIG.COL_UTM_YA - 1]
        });
      }
    }

    const plan = readPlanSheet(ss);
    return jsonOutput({ status: 'ok', data: data, plan: plan });

  } catch (err) {
    return jsonOutput({ status: 'error', message: err.message });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.action === 'savePlan') {
      savePlanSheet(payload.data || {});
      return jsonOutput({ status: 'ok' });
    }

    return jsonOutput({ status: 'error', message: 'Action tidak dikenali: ' + payload.action });

  } catch (err) {
    return jsonOutput({ status: 'error', message: err.message });
  }
}

// ─────────────────────────────────────────────
// HELPER — output & format
// ─────────────────────────────────────────────

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Sheet bisa simpan tanggal sebagai objek Date (kalau kolom terformat Date)
// atau sebagai teks biasa. Dua kasus ini ditangani; hasil akhir selalu
// string 'yyyy-MM-dd' (atau '' kalau kosong) supaya aman dari masalah timezone
// saat di-parse ulang oleh `new Date(...)` di frontend.
function formatTanggal(val) {
  if (val === '' || val === null || val === undefined) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).trim();
}

// Konversi "DL001" -> "DL01", "DL082" -> "DL82", "DL100" -> "DL100"
// supaya match dengan format nama di TITIK/SUMUR pada index.html.
// Kalau formatnya bukan [huruf-huruf][angka] (mis. kode custom), nilai
// dikembalikan apa adanya (tidak diubah).
function normalizeKode(kode) {
  const s = String(kode).trim();
  const m = s.match(/^([A-Za-z]+)0*([0-9]+)$/);
  if (!m) return s;
  const prefix = m[1];
  let numStr = String(parseInt(m[2], 10));
  if (numStr.length < 2) numStr = '0' + numStr; // padding minimal 2 digit (DL01..DL09)
  return prefix + numStr;
}

// ─────────────────────────────────────────────
// HELPER — tab "Plan" (warna Plan Mode dari web)
// ─────────────────────────────────────────────

function getOrCreatePlanSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.PLAN_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['Kode Titik', 'Warna']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readPlanSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);
  const plan = {};
  if (!sheet) return plan;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return plan;

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  values.forEach(function (r) {
    const nama = r[0], warna = r[1];
    if (nama && warna) plan[normalizeKode(nama)] = String(warna).trim();
  });
  return plan;
}

function savePlanSheet(planData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreatePlanSheet(ss);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  }

  const rows = Object.keys(planData).map(function (nama) {
    return [nama, planData[nama]];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}
