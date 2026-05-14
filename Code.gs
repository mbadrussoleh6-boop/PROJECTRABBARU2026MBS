const SHEET_ID = '1leDarZxebdSbxQiyVDUjKdftfUiIgBzvalVX6L1uy8c';
const SHEET_NAME = 'KHS TA - TIF 2026';

// Fungsi wajib untuk merender halaman HTML
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('Generator RAB Maintenance');
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