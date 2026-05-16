const SHEET_ID = '1leDarZxebdSbxQiyVDUjKdftfUiIgBzvalVX6L1uy8c';
const SHEET_NAME = 'KHS TA - TIF 2026';

// Fungsi wajib untuk merender halaman HTML
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('RABIT - RAB Instant Tool');
}

// Fungsi untuk mengambil data KHS dari Spreadsheet
function getKatalogData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  // Mengambil data dari B9 sampai F732
  const data = sheet.getRange('B9:F732').getValues();
  
  const katalog = [];
  
  // Looping data dan memasukannya ke dalam array object
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== "") { // Memastikan baris designator tidak kosong
      katalog.push({
        designator: data[i][0],    // Kolom B
        uraian: data[i][1],        // Kolom C
        satuan: data[i][2],        // Kolom D
        hargaMaterial: data[i][3] || 0, // Kolom E
        hargaJasa: data[i][4] || 0      // Kolom F
      });
    }
  }
  return katalog;
}

function buatExcelRAB(itemRAB, namaPekerjaan) {
  if (!itemRAB || itemRAB.length === 0) {
    throw new Error("Tabel RAB masih kosong. Tambahkan minimal satu item pekerjaan.");
  }

  if (!namaPekerjaan || String(namaPekerjaan).trim() === "") {
    throw new Error("Nama pekerjaan belum diisi.");
  }

  const namaFileAman = String(namaPekerjaan)
    .trim()
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, "_")
    .substring(0, 80);

  const namaTemp = `TEMP_RAB_${namaFileAman}_${new Date().getTime()}`;
  const fileAsli = DriveApp.getFileById(SHEET_ID);
  const fileTemp = fileAsli.makeCopy(namaTemp);
  const tempId = fileTemp.getId();
  
  try {
    const ssTemp = SpreadsheetApp.openById(tempId);
    const sheet = ssTemp.getSheetByName(SHEET_NAME);
  
    if (!sheet) {
      throw new Error(`Sheet "${SHEET_NAME}" tidak ditemukan.`);
    }
  
    // HANYA SISAKAN SHEET KHS YANG DIGUNAKAN
    // Sheet lain dihapus dari file sementara, bukan dari file asli
    ssTemp.setActiveSheet(sheet);
  
    ssTemp.getSheets().forEach(function(sh) {
      if (sh.getSheetId() !== sheet.getSheetId()) {
        ssTemp.deleteSheet(sh);
      }
    });
  
    SpreadsheetApp.flush();
  
    // Sesuai range KHS yang Anda gunakan di getKatalogData()
    const START_ROW = 9;
    const END_ROW = 732;
    const TOTAL_ROWS = END_ROW - START_ROW + 1;
  
    const DESIGNATOR_COL = 2; // Kolom B
    const VOLUME_COL = 7;     // Kolom G

    // Buat map volume berdasarkan designator dari tabel web
    const volumeByDesignator = {};

    itemRAB.forEach(function(item) {
      const designator = normalisasiDesignator_(item.designator);
      const volume = parseFloat(item.volume) || 0;

      if (designator && volume > 0) {
        volumeByDesignator[designator] = volume;
      }
    });

    // Ambil semua designator dari kolom B sheet KHS
    const designatorSheet = sheet
      .getRange(START_ROW, DESIGNATOR_COL, TOTAL_ROWS, 1)
      .getValues();

    // Siapkan isi kolom G berdasarkan kecocokan designator
    const hasilVolume = designatorSheet.map(function(row) {
      const designatorKHS = normalisasiDesignator_(row[0]);

      if (volumeByDesignator.hasOwnProperty(designatorKHS)) {
        return [volumeByDesignator[designatorKHS]];
      }

      return [""];
    });

    // Kosongkan dulu kolom G, lalu isi volume hasil dari web
    sheet.getRange(START_ROW, VOLUME_COL, TOTAL_ROWS, 1).clearContent();
    sheet.getRange(START_ROW, VOLUME_COL, TOTAL_ROWS, 1).setValues(hasilVolume);

    SpreadsheetApp.flush();

    // Export salinan spreadsheet menjadi XLSX
    const exportUrl = `https://docs.google.com/spreadsheets/d/${tempId}/export?format=xlsx`;

    const response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error("Gagal export file Excel. Kode error: " + response.getResponseCode());
    }

    const fileName = `RAB_${namaFileAman}.xlsx`;
    const mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    const blob = response.getBlob().setName(fileName);

    return {
      fileName: fileName,
      mimeType: mimeType,
      base64: Utilities.base64Encode(blob.getBytes())
    };

  } finally {
    // Hapus file salinan sementara agar Drive tidak penuh
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}

function normalisasiDesignator_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}