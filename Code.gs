const SPREADSHEET_ID = "1leDarZxebdSbxQiyVDUjKdftfUiIgBzvalVX6L1uy8c";
const SHEET_KHS = "KHS TA - TIF 2026";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("RAB Generator");
}

function getMasterData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_KHS);

  const data = sheet.getRange("B9:F732").getValues();

  return data
    .filter(r => r[0] != "")
    .map((r, i) => ({
      urutan: i,
      designator: r[0],
      uraian: r[1],
      satuan: r[2],
      hargaMaterial: Number(r[3]) || 0,
      hargaJasa: Number(r[4]) || 0
    }));
}

const SHEET_RAB_HEADER = "RAB_HEADER";
const SHEET_RAB_DETAIL = "RAB_DETAIL";

function saveRab(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    if (!payload) {
      throw new Error("Data RAB tidak ditemukan.");
    }

    const namaPekerjaan = String(payload.namaPekerjaan || "").trim();
    const items = payload.items || [];

    if (!namaPekerjaan) {
      throw new Error("Nama pekerjaan wajib diisi.");
    }

    if (!items.length) {
      throw new Error("Item RAB masih kosong.");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const headerSheet = getOrCreateSheet_(
      ss,
      SHEET_RAB_HEADER,
      [
        "ID RAB",
        "Tanggal Simpan",
        "Nama Pekerjaan",
        "Jumlah Item",
        "Total Material",
        "Total Jasa",
        "Total RAB"
      ]
    );

    const detailSheet = getOrCreateSheet_(
      ss,
      SHEET_RAB_DETAIL,
      [
        "ID RAB",
        "No",
        "Designator",
        "Uraian",
        "Satuan",
        "Harga Material",
        "Harga Jasa",
        "Volume",
        "Total Material",
        "Total Jasa",
        "Total"
      ]
    );

    const rabId = generateRabId_(headerSheet);
    const now = new Date();

    let totalMaterial = 0;
    let totalJasa = 0;

    const detailRows = items.map((item, index) => {

      const hargaMaterial = Number(item.hargaMaterial) || 0;
      const hargaJasa = Number(item.hargaJasa) || 0;
      const volume = Number(item.volume) || 0;

      const itemTotalMaterial = hargaMaterial * volume;
      const itemTotalJasa = hargaJasa * volume;
      const itemTotal = itemTotalMaterial + itemTotalJasa;

      totalMaterial += itemTotalMaterial;
      totalJasa += itemTotalJasa;

      return [
        rabId,
        index + 1,
        item.designator || "",
        item.uraian || "",
        item.satuan || "",
        hargaMaterial,
        hargaJasa,
        volume,
        itemTotalMaterial,
        itemTotalJasa,
        itemTotal
      ];

    });

    const totalRab = totalMaterial + totalJasa;

    headerSheet.appendRow([
      rabId,
      now,
      namaPekerjaan,
      items.length,
      totalMaterial,
      totalJasa,
      totalRab
    ]);

    detailSheet
      .getRange(
        detailSheet.getLastRow() + 1,
        1,
        detailRows.length,
        detailRows[0].length
      )
      .setValues(detailRows);

    formatRabSheets_(headerSheet, detailSheet);

    return {
      success: true,
      rabId: rabId,
      namaPekerjaan: namaPekerjaan,
      jumlahItem: items.length,
      totalMaterial: totalMaterial,
      totalJasa: totalJasa,
      totalRab: totalRab
    };

  } finally {

    lock.releaseLock();

  }
}

function generateRabId_(headerSheet) {

  const timezone =
    Session.getScriptTimeZone() || "Asia/Jakarta";

  let rabId = "";

  while (true) {

    const timestamp =
      Utilities.formatDate(
        new Date(),
        timezone,
        "yyMMdd-HHmmss"
      );

    rabId = `RAB-${timestamp}`;

    const lastRow = headerSheet.getLastRow();

    if (lastRow <= 1) {
      break;
    }

    const existingIds =
      headerSheet
        .getRange(2, 1, lastRow - 1, 1)
        .getValues()
        .flat();

    if (!existingIds.includes(rabId)) {
      break;
    }

    Utilities.sleep(1000);

  }

  return rabId;

}

function getOrCreateSheet_(ss, sheetName, headers) {

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {

    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight("bold")
      .setBackground("#0f172a")
      .setFontColor("#ffffff");

    sheet.setFrozenRows(1);

  }

  return sheet;

}

function formatRabSheets_(headerSheet, detailSheet) {

  headerSheet.getRange("B:B")
    .setNumberFormat("dd/MM/yyyy HH:mm:ss");

  headerSheet.getRange("E:G")
    .setNumberFormat('"Rp" #,##0');

  detailSheet.getRange("F:G")
    .setNumberFormat('"Rp" #,##0');

  detailSheet.getRange("I:K")
    .setNumberFormat('"Rp" #,##0');

}