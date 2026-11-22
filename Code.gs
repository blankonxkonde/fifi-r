/**
 * Apps Script untuk menerima hasil rating dari halaman GitHub Pages.
 *
 * Cara pasang:
 * 1. Buat Google Sheet baru → Extensions → Apps Script.
 * 2. Hapus isi Code.gs, tempel file ini, Save.
 * 3. Deploy → New deployment → pilih "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Salin URL Web App (berakhiran /exec), tempel ke SCRIPT_URL di index.html.
 *
 * Setiap kali tombol "Simpan ke Google Sheet" ditekan, baris baru ditambahkan
 * ke sheet "Hasil".
 */

const SHEET_NAME = "Hasil";
const HEADERS = ["Waktu", "Nama", "Kategori", "Menu", "Rating", "Mau Coba (1-5)", "Catatan"];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#f3f3f3");
  }
  return sh;
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const rows = Array.isArray(body.rows) ? body.rows : [];
    if (!rows.length) return json_({ ok: false, error: "no rows" });

    const now = new Date();
    const values = rows.map(r => [
      now,
      String(r.nama || "").slice(0, 100),
      String(r.kategori || "").slice(0, 50),
      String(r.menu || "").slice(0, 100),
      r.rating === "" || r.rating == null ? "" : String(r.rating),
      r.mauCoba === "" || r.mauCoba == null ? "" : Number(r.mauCoba),
      String(r.catatan || "").slice(0, 1000)
    ]);

    const sh = getSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, values.length, HEADERS.length).setValues(values);
    return json_({ ok: true, saved: values.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, message: "Endpoint rating aktif" });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
