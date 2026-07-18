/**
 * ABSENSI FOTO — BACKEND (Google Apps Script)
 * ---------------------------------------------------------
 * Cara pakai singkat (detail lengkap ada di PANDUAN.md):
 * 1. Buat Google Spreadsheet baru.
 * 2. Buka Extensions > Apps Script, hapus isi default,
 *    lalu tempel (paste) seluruh isi file ini.
 * 3. Jalankan sekali fungsi setup() dari editor Apps Script
 *    (pilih "setup" di dropdown, klik Run) untuk membuat
 *    header sheet & folder Drive secara otomatis.
 * 4. Deploy > New deployment > Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Salin URL Web App yang muncul, tempel ke variabel
 *    APPS_SCRIPT_URL di file index.html.
 */

const SHEET_NAME = 'Absensi';
const DRIVE_FOLDER_NAME = 'Foto Absensi';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Nama', 'ID', 'Tipe', 'Foto', 'Latitude', 'Longitude']);
    sheet.setFrozenRows(1);
  }
  getOrCreateFolder_();
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const nama = (body.nama || '').toString().trim();
    const idPeg = (body.id || '').toString().trim();
    const tipe = (body.tipe || '').toString().trim();
    const waktuIso = body.waktu || new Date().toISOString();
    const gambar = body.gambar || '';
    const lat = body.lat || '';
    const lng = body.lng || '';

    if (!nama) throw new Error('Nama wajib diisi');
    if (!gambar) throw new Error('Foto wajib disertakan');

    const fotoUrl = simpanFoto_(gambar, nama, waktuIso);

    const sheet = getSheet_();
    sheet.appendRow([new Date(waktuIso), nama, idPeg, tipe, fotoUrl, lat, lng]);

    return jsonResponse_({ ok: true, fotoUrl: fotoUrl });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}

function doGet(e) {
  return jsonResponse_({ ok: true, message: 'Endpoint absensi aktif. Gunakan metode POST untuk mengirim data.' });
}

/* ---------- Helper ---------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Nama', 'ID', 'Tipe', 'Foto', 'Latitude', 'Longitude']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateFolder_() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function simpanFoto_(dataUrl, nama, waktuIso) {
  const match = dataUrl.match(/^data:(image\/\w+);base64,(.*)$/);
  if (!match) throw new Error('Format gambar tidak valid');
  const mimeType = match[1];
  const base64 = match[2];
  const bytes = Utilities.base64Decode(base64);
  const ext = mimeType.split('/')[1] || 'jpg';
  const fileName = `${nama.replace(/[^a-zA-Z0-9]/g, '_')}_${waktuIso.replace(/[:.]/g, '-')}.${ext}`;
  const blob = Utilities.newBlob(bytes, mimeType, fileName);

  const folder = getOrCreateFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?id=${file.getId()}`;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
