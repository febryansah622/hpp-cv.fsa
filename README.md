/**
 * ============================================================
 *  REKAP HPP OTOMATIS - CV. FSA (v2)
 * ============================================================
 * Perubahan dari v1:
 * - Mingguan sekarang ditarik TERPISAH per metrik (Porsi, Omset,
 *   Pembelian, Pemakaian) masing-masing Pekan 1-5, bukan 1 angka
 *   ambigu per pekan seperti versi sebelumnya.
 * - Caranya: cari baris header yang memuat teks "Pekan 1".."Pekan 5"
 *   (jadi tahu kolom mana untuk pekan berapa), lalu di bawah header
 *   itu cari baris berlabel "Porsi"/"Omset"/"Pembelian"/"Pemakaian"
 *   dan ambil 5 angka di kolom pekan yang sesuai. Sheet biasanya
 *   punya beberapa tabel mingguan terpisah (1 utk porsi&omset,
 *   1 utk pembelian, 1 utk pemakaian) - script ini akan menjelajahi
 *   SEMUA tabel seperti itu di dalam sheet, bukan cuma yang pertama.
 *
 * Cara pakai: sama seperti v1 (lihat PANDUAN_SETUP.md).
 * ============================================================
 */

const SHEET_MASTER_INDEX = "Master Index";
const SHEET_DATA_GABUNGAN = "Data Gabungan";

const LABELS = {
  totalPorsi: "Total Porsi",
  totalOmset: "Total Omset",
  totalPembelian: "Total Pembelian",
  totalPemakaian: "Total Pemakaian",
  stokOnHand: "Stok On Hand",
};

function tarikSemuaData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = ss.getSheetByName(SHEET_MASTER_INDEX);
  if (!indexSheet) throw new Error('Tab "' + SHEET_MASTER_INDEX + '" tidak ditemukan.');

  const rows = indexSheet.getDataRange().getValues();
  const header = rows[0];
  const colBulan = header.indexOf("Bulan");
  const colUnit = header.indexOf("Unit");
  const colLink = header.indexOf("Link Sheet");
  const colStatus = header.indexOf("Status");

  const hasil = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const bulan = row[colBulan];
    const unit = row[colUnit];
    const link = row[colLink];
    if (!link) continue;

    try {
      const fileId = ekstrakFileId(link);
      const sourceSs = SpreadsheetApp.openById(fileId);
      const sourceSheet = sourceSs.getSheets()[0];
      const data = parseSheetDapur(sourceSheet);
      data.bulan = bulan;
      data.unit = unit;
      hasil.push(data);
      indexSheet.getRange(i + 1, colStatus + 1).setValue(
        "OK - ditarik " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM HH:mm")
      );
    } catch (err) {
      indexSheet.getRange(i + 1, colStatus + 1).setValue("ERROR: " + err.message);
    }
  }

  simpanKeDataGabungan(hasil);
  return hasil;
}

function ekstrakFileId(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) throw new Error("Link tidak valid: " + url);
  return match[1];
}

/**
 * Cari sel yang isinya mengandung `labelText`, ambil nilai di kanan atau bawahnya.
 * Dipakai untuk angka TOTAL (bukan mingguan).
 */
function cariNilaiByLabel(sheet, labelText, data) {
  data = data || sheet.getDataRange().getValues();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const cell = data[r][c];
      if (typeof cell === "string" && cell.toLowerCase().indexOf(labelText.toLowerCase()) !== -1) {
        if (c + 1 < data[r].length && typeof data[r][c + 1] === "number") {
          return data[r][c + 1];
        }
        if (r + 1 < data.length && typeof data[r + 1][c] === "number") {
          return data[r + 1][c];
        }
      }
    }
  }
  return null;
}

/**
 * Cari SEMUA tabel mingguan di dalam sheet (bisa lebih dari satu),
 * kembalikan objek { porsi:[5 angka], omset:[...], pembelian:[...], pemakaian:[...] }.
 */
function cariSemuaMingguan(data) {
  const result = {
    porsi: [0, 0, 0, 0, 0],
    omset: [0, 0, 0, 0, 0],
    pembelian: [0, 0, 0, 0, 0],
    pemakaian: [0, 0, 0, 0, 0],
  };

  for (let r = 0; r < data.length; r++) {
    // cari kolom "Pekan 1" di baris ini
    let col1 = -1;
    for (let c = 0; c < data[r].length; c++) {
      if (typeof data[r][c] === "string" && data[r][c].toLowerCase().replace(/\s+/g, " ").indexOf("pekan 1") !== -1) {
        col1 = c;
        break;
      }
    }
    if (col1 === -1) continue;

    // cari kolom Pekan 2..5 di baris yang sama (cari ke kanan dari col1)
    const pekanCols = [col1];
    for (let k = 2; k <= 5; k++) {
      let found = -1;
      for (let c = col1 + 1; c < data[r].length; c++) {
        if (typeof data[r][c] === "string" && data[r][c].toLowerCase().replace(/\s+/g, " ").indexOf("pekan " + k) !== -1) {
          found = c;
          break;
        }
      }
      if (found === -1) break;
      pekanCols.push(found);
    }
    if (pekanCols.length < 5) continue; // header tidak lengkap, bukan tabel mingguan yang valid

    // scan baris-baris di bawah header ini untuk label metrik, sampai ketemu baris kosong total
    for (let rr = r + 1; rr < Math.min(r + 12, data.length); rr++) {
      const rowIsEmpty = data[rr].every(function (v) { return v === "" || v === null; });
      if (rowIsEmpty) break;

      // cari label di kolom sebelum col1
      let labelText = "";
      for (let c = 0; c < col1; c++) {
        if (typeof data[rr][c] === "string" && data[rr][c].trim() !== "") {
          labelText = data[rr][c].toLowerCase();
        }
      }
      if (!labelText) continue;

      const values = pekanCols.map(function (pc) {
        return typeof data[rr][pc] === "number" ? data[rr][pc] : 0;
      });

      if (labelText.indexOf("porsi") !== -1) result.porsi = values;
      else if (labelText.indexOf("omset") !== -1) result.omset = values;
      else if (labelText.indexOf("pembelian") !== -1) result.pembelian = values;
      else if (labelText.indexOf("pemakaian") !== -1) result.pemakaian = values;
    }
  }

  return result;
}

/**
 * Parse satu sheet dapur jadi objek data ringkas.
 * CATATAN: Dapur Depok (omset subsidi terpisah) dan Dapur Tangsel
 * (gabungan Tangsel+Tangkot) mungkin perlu penyesuaian tambahan
 * kalau hasil tarikannya kurang pas - kabari saya kalau begitu.
 */
function parseSheetDapur(sheet) {
  const data = sheet.getDataRange().getValues();
  const mingguan = cariSemuaMingguan(data);
  return {
    totalPorsi: cariNilaiByLabel(sheet, LABELS.totalPorsi, data),
    totalOmset: cariNilaiByLabel(sheet, LABELS.totalOmset, data),
    totalPembelian: cariNilaiByLabel(sheet, LABELS.totalPembelian, data),
    totalPemakaian: cariNilaiByLabel(sheet, LABELS.totalPemakaian, data),
    stokOnHand: cariNilaiByLabel(sheet, LABELS.stokOnHand, data),
    mingguan: mingguan,
  };
}

function simpanKeDataGabungan(hasil) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DATA_GABUNGAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DATA_GABUNGAN);
  }

  const headerRow = [
    "Bulan", "Unit", "Total Porsi", "Total Omset", "Total Pembelian",
    "Total Pemakaian", "% Pemakaian/Omset", "Stok On Hand",
    "Porsi Pekan 1", "Porsi Pekan 2", "Porsi Pekan 3", "Porsi Pekan 4", "Porsi Pekan 5",
    "Omset Pekan 1", "Omset Pekan 2", "Omset Pekan 3", "Omset Pekan 4", "Omset Pekan 5",
    "Pembelian Pekan 1", "Pembelian Pekan 2", "Pembelian Pekan 3", "Pembelian Pekan 4", "Pembelian Pekan 5",
    "Pemakaian Pekan 1", "Pemakaian Pekan 2", "Pemakaian Pekan 3", "Pemakaian Pekan 4", "Pemakaian Pekan 5",
    "Terakhir Update"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headerRow);
  }

  const existing = sheet.getDataRange().getValues();
  const now = new Date();

  hasil.forEach(function (d) {
    const persen = d.totalOmset ? (d.totalPemakaian / d.totalOmset * 100) : 0;
    const m = d.mingguan;
    const rowValues = [
      d.bulan, d.unit, d.totalPorsi, d.totalOmset, d.totalPembelian,
      d.totalPemakaian, persen, d.stokOnHand,
      m.porsi[0], m.porsi[1], m.porsi[2], m.porsi[3], m.porsi[4],
      m.omset[0], m.omset[1], m.omset[2], m.omset[3], m.omset[4],
      m.pembelian[0], m.pembelian[1], m.pembelian[2], m.pembelian[3], m.pembelian[4],
      m.pemakaian[0], m.pemakaian[1], m.pemakaian[2], m.pemakaian[3], m.pemakaian[4],
      now
    ];

    let foundRow = -1;
    for (let i = 1; i < existing.length; i++) {
      if (existing[i][0] === d.bulan && existing[i][1] === d.unit) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
  });
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DATA_GABUNGAN);
  if (!sheet) {
    return jsonResponse({ error: "Data Gabungan belum ada. Jalankan tarikSemuaData() dulu." });
  }

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  let rows = data.slice(1).map(function (row) {
    const obj = {};
    header.forEach(function (h, idx) { obj[h] = row[idx]; });
    return obj;
  });

  const params = e && e.parameter ? e.parameter : {};
  if (params.bulan) {
    rows = rows.filter(function (r) { return r["Bulan"] === params.bulan; });
  }
  if (params.unit) {
    rows = rows.filter(function (r) { return r["Unit"] === params.unit; });
  }

  return jsonResponse({ data: rows, jumlah: rows.length });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
