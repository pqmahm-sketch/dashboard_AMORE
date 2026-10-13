// ============================================================================
// OTOMASI BULANAN - DASHBOARD AMORE NASIONAL
// ============================================================================
// Script ini mengotomasi proses bulanan pada 2 spreadsheet:
// 1. Spreadsheet MD  - "[LIVE] Dashboard AMORE Nasional - For MD"
// 2. Spreadsheet AHM - "[LIVE] Dashboard AMORE Nasional - For AHM"
//
// Proses yang diotomasi:
// Task 1: Copy formula ke bulan berjalan & paste values bulan sebelumnya (H+7)
// Task 2: Rollover tahun baru (update header, clear data, pertahankan formula)
// Task 3: Update formula Data Mentah AHM ke sumber lokal
// Task 4: Update periode di Dashboard MD
// ============================================================================

var MD_SPREADSHEET_ID  = '1uU3mG1tu-d9aW1Z-eL-ocOQHhMzg2xPWAN4DvSYmo1o';
var AHM_SPREADSHEET_ID = '1JzPicTYV2iBnVpS-Lal1ZZ4IfaJARj9_fb6nIlEm42I';

var MONTH_NAMES_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun',
                      'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

var MONTH_NAMES_EN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// 6 tabel di "Rekap Result All MD"
var TABLE_NAMES = [
  'OVERALL PROCESS',
  'FLINKOTE',
  'REGULATOR',
  'TIMER',
  'MARKING SELANG',
  'MARKING FB'
];

// ============================================================================
// TRIGGER SETUP
// ============================================================================

/**
 * Pasang trigger harian jam 08:00 WIB (01:00 UTC).
 * Jalankan fungsi ini SEKALI secara manual untuk mengaktifkan otomasi.
 */
function setupDailyTrigger() {
  // Hapus trigger lama yang mungkin sudah ada
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runMonthlyAutomation') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Buat trigger harian jam 08:00 WIB (UTC+7 → 01:00 UTC)
  ScriptApp.newTrigger('runMonthlyAutomation')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();

  Logger.log('Trigger harian berhasil dipasang. Akan berjalan setiap hari jam 08:00 WIB.');
}

/**
 * Hapus semua trigger otomasi.
 */
function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runMonthlyAutomation') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  Logger.log('Semua trigger otomasi telah dihapus.');
}

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================

/**
 * Fungsi utama yang dipanggil oleh trigger harian.
 * Mengecek apakah hari ini adalah H+7 dari akhir bulan sebelumnya.
 * Jika ya, jalankan semua proses otomasi.
 */
function runMonthlyAutomation() {
  var today = new Date();
  var currentMonth = today.getMonth(); // 0-indexed (0=Jan)
  var currentYear  = today.getFullYear();

  // Hitung H+7: tanggal terakhir bulan sebelumnya + 7 hari
  var prevMonthLastDay = new Date(currentYear, currentMonth, 0); // hari terakhir bulan lalu
  var hPlus7 = new Date(prevMonthLastDay);
  hPlus7.setDate(hPlus7.getDate() + 7);

  // Cek apakah hari ini = H+7
  if (today.getDate() !== hPlus7.getDate() ||
      today.getMonth() !== hPlus7.getMonth() ||
      today.getFullYear() !== hPlus7.getFullYear()) {
    Logger.log('Hari ini bukan H+7. Tidak ada proses yang dijalankan.');
    return;
  }

  Logger.log('=== H+7 terdeteksi! Memulai proses otomasi bulanan ===');
  Logger.log('Tanggal: ' + today.toISOString());
  Logger.log('Bulan berjalan: ' + MONTH_NAMES_ID[currentMonth] + ' ' + currentYear);

  // Bulan sebelumnya (0-indexed)
  var prevMonth = currentMonth === 0 ? 11 : currentMonth - 1;
  var prevYear  = currentMonth === 0 ? currentYear - 1 : currentYear;

  // Cek apakah ini awal tahun baru (bulan berjalan = Januari)
  var isNewYear = (currentMonth === 0);

  try {
    if (isNewYear) {
      // TASK 2: Year rollover
      Logger.log('--- Tahun baru terdeteksi! Menjalankan year rollover ---');
      handleYearRollover(currentYear);
    } else {
      // TASK 1: Copy formula & paste values
      Logger.log('--- Task 1: Proses Rekap Result All MD ---');
      processRekapAllMD(currentMonth, prevMonth);
    }

    // TASK 3: Update formula Data Mentah AHM
    Logger.log('--- Task 3: Update Data Mentah AHM ---');
    updateAhmDataMentah(currentMonth);

    // TASK 4: Update Dashboard
    Logger.log('--- Task 4: Update Dashboard MD ---');
    updateDashboard(currentMonth, currentYear);

    // Update juga "Detail Problem" dan "Rekap Result Per MD" jika ada
    updateDetailProblem(currentMonth, currentYear);

    Logger.log('=== Semua proses otomasi selesai! ===');
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    throw e;
  }
}

// ============================================================================
// TASK 1: PROSES REKAP RESULT ALL MD
// ============================================================================

/**
 * Copy formula dari bulan sebelumnya ke bulan berjalan,
 * lalu paste values (freeze) pada kolom bulan sebelumnya.
 *
 * @param {number} currentMonth - Bulan berjalan (0-indexed)
 * @param {number} prevMonth - Bulan sebelumnya (0-indexed)
 */
function processRekapAllMD(currentMonth, prevMonth) {
  var ss = SpreadsheetApp.openById(MD_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Rekap Result All MD');
  if (!sheet) {
    Logger.log('Sheet "Rekap Result All MD" tidak ditemukan!');
    return;
  }

  // Kolom B=2 (Jan) sampai M=13 (Des), jadi bulan ke-i (0-indexed) → kolom i+2
  var prevCol    = prevMonth + 2;    // kolom bulan sebelumnya
  var currentCol = currentMonth + 2; // kolom bulan berjalan

  var sections = getTableSections(sheet);

  for (var s = 0; s < sections.length; s++) {
    var section = sections[s];
    var dataStartRow = section.dataStartRow;
    var dataEndRow   = section.dataEndRow;
    var numRows      = dataEndRow - dataStartRow + 1;

    Logger.log('Proses tabel: ' + section.name +
               ' (baris ' + dataStartRow + '-' + dataEndRow + ')');

    // 1. Copy formula dari bulan sebelumnya ke bulan berjalan
    var sourceRange = sheet.getRange(dataStartRow, prevCol, numRows, 1);
    var targetRange = sheet.getRange(dataStartRow, currentCol, numRows, 1);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

    // 2. Paste values (freeze) pada bulan sebelumnya
    var values = sourceRange.getValues();
    sourceRange.setValues(values);
  }

  Logger.log('Task 1 selesai: Formula dicopy ke ' + MONTH_NAMES_ID[currentMonth] +
             ', values di-freeze pada ' + MONTH_NAMES_ID[prevMonth]);
}

/**
 * Deteksi lokasi 6 tabel di sheet "Rekap Result All MD".
 * Mencari nama tabel di kolom A untuk menentukan baris header dan data.
 *
 * @param {Sheet} sheet - Sheet object
 * @return {Array} Array of {name, headerRow, dataStartRow, dataEndRow}
 */
function getTableSections(sheet) {
  var lastRow = sheet.getLastRow();
  var colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  var sections = [];

  for (var t = 0; t < TABLE_NAMES.length; t++) {
    var tableName = TABLE_NAMES[t];

    for (var r = 0; r < colA.length; r++) {
      var cellValue = String(colA[r][0]).trim().toUpperCase();

      if (cellValue === tableName.toUpperCase()) {
        var nameRow = r + 1; // 1-indexed
        var headerRow = nameRow + 1;
        var dataStartRow = headerRow + 1;

        // Cari akhir data: baris terakhir sebelum tabel berikutnya atau akhir sheet
        var dataEndRow = lastRow;
        if (t < TABLE_NAMES.length - 1) {
          // Cari tabel berikutnya
          for (var r2 = r + 1; r2 < colA.length; r2++) {
            var nextCell = String(colA[r2][0]).trim().toUpperCase();
            if (nextCell === TABLE_NAMES[t + 1].toUpperCase()) {
              dataEndRow = r2; // baris sebelum nama tabel berikutnya (1-indexed: r2+1-1 = r2)
              break;
            }
          }
        }

        // Mundur dari dataEndRow, skip baris kosong
        while (dataEndRow > dataStartRow) {
          var rowVals = sheet.getRange(dataEndRow, 1, 1, 14).getValues()[0];
          var hasContent = rowVals.some(function(v) { return v !== '' && v !== null; });
          if (hasContent) break;
          dataEndRow--;
        }

        sections.push({
          name: tableName,
          nameRow: nameRow,
          headerRow: headerRow,
          dataStartRow: dataStartRow,
          dataEndRow: dataEndRow
        });
        break;
      }
    }
  }

  Logger.log('Ditemukan ' + sections.length + ' tabel');
  return sections;
}

// ============================================================================
// TASK 2: YEAR ROLLOVER
// ============================================================================

/**
 * Proses rollover tahun baru:
 * 1. Copy formula Desember ke Januari (sebelum freeze)
 * 2. Freeze values Desember
 * 3. Update header bulan (misal "Jan-26" → "Jan-27")
 * 4. Clear data Feb-Des, pertahankan formula di Januari
 *
 * @param {number} newYear - Tahun baru (e.g. 2027)
 */
function handleYearRollover(newYear) {
  var ss = SpreadsheetApp.openById(MD_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Rekap Result All MD');
  if (!sheet) return;

  var sections = getTableSections(sheet);
  var yearSuffix = String(newYear).slice(-2); // "27"

  // STEP 1 & 2: Copy formula Des → Jan, lalu freeze Des
  for (var s = 0; s < sections.length; s++) {
    var section = sections[s];
    var startRow = section.dataStartRow;
    var endRow   = section.dataEndRow;
    var numRows  = endRow - startRow + 1;

    // Kolom Des=13 (bulan 11+2), kolom Jan=2 (bulan 0+2)
    var decRange = sheet.getRange(startRow, 13, numRows, 1);
    var janRange = sheet.getRange(startRow, 2, numRows, 1);

    // Copy formula Des ke Jan
    decRange.copyTo(janRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

    // Freeze values Des
    var decValues = decRange.getValues();
    decRange.setValues(decValues);
  }
  Logger.log('Formula Des dicopy ke Jan, Des di-freeze');

  // STEP 3: Update header bulan di semua tabel
  for (var s = 0; s < sections.length; s++) {
    var headerRow = sections[s].headerRow;
    for (var m = 0; m < 12; m++) {
      var col = m + 2; // B=2 untuk Jan
      var headerText = MONTH_NAMES_ID[m] + '-' + yearSuffix;
      sheet.getRange(headerRow, col).setValue(headerText);
    }
  }
  Logger.log('Header bulan diupdate ke tahun ' + newYear);

  // STEP 4: Clear data Feb-Des (kolom C=3 sampai M=13), set ke "-"
  for (var s = 0; s < sections.length; s++) {
    var section = sections[s];
    var startRow = section.dataStartRow;
    var endRow   = section.dataEndRow;
    var numRows  = endRow - startRow + 1;

    // Kolom C (Feb) sampai M (Des) = kolom 3 sampai 13 = 11 kolom
    var clearRange = sheet.getRange(startRow, 3, numRows, 11);
    var blanks = [];
    for (var r = 0; r < numRows; r++) {
      var row = [];
      for (var c = 0; c < 11; c++) {
        row.push('-');
      }
      blanks.push(row);
    }
    clearRange.setValues(blanks);
  }
  Logger.log('Data Feb-Des dikosongkan');

  // Update kolom N "Gr M vs M-1" header jika perlu (opsional, biasanya tidak berubah)

  // Update "Rekap Result Per MD" headers
  updateRekapPerMDHeaders(ss, newYear);

  Logger.log('Year rollover selesai untuk tahun ' + newYear);
}

/**
 * Update header bulan di sheet "Rekap Result Per MD".
 */
function updateRekapPerMDHeaders(ss, newYear) {
  var sheet = ss.getSheetByName('Rekap Result Per MD');
  if (!sheet) return;

  var yearSuffix = String(newYear).slice(-2);

  // Row 3 berisi header bulan ("Jan-26" dst), kolom B-M
  for (var m = 0; m < 12; m++) {
    var col = m + 2;
    sheet.getRange(3, col).setValue(MONTH_NAMES_ID[m] + '-' + yearSuffix);
  }
  Logger.log('Header "Rekap Result Per MD" diupdate');
}

// ============================================================================
// TASK 3: UPDATE DATA MENTAH AHM
// ============================================================================

/**
 * Update formula di cell A2 sheet "Data Mentah" (AHM) agar mengambil
 * data dari sheet lokal "Data Mentah All" dengan filter bulan berjalan.
 *
 * Formula: =IFERROR(SORT(FILTER('Data Mentah All'!A2:L, 'Data Mentah All'!O2:O=BULAN), 1, TRUE), "")
 *
 * @param {number} currentMonth - Bulan berjalan (0-indexed, 0=Jan)
 */
function updateAhmDataMentah(currentMonth) {
  var ss = SpreadsheetApp.openById(AHM_SPREADSHEET_ID);
  var sheetDataMentah = ss.getSheetByName('Data Mentah');
  if (!sheetDataMentah) {
    Logger.log('Sheet "Data Mentah" tidak ditemukan di AHM!');
    return;
  }

  // Verifikasi sheet sumber ada
  var sheetAll = ss.getSheetByName('Data Mentah All');
  if (!sheetAll) {
    Logger.log('Sheet "Data Mentah All" tidak ditemukan di AHM!');
    return;
  }

  // Cari kolom "Month" di "Data Mentah All" secara dinamis
  var monthCol = findMonthColumn(sheetAll);
  if (!monthCol) {
    Logger.log('Kolom "Month" tidak ditemukan di "Data Mentah All"! Menggunakan kolom O sebagai default.');
    monthCol = 'O';
  }

  var monthNumber = currentMonth + 1; // 1-indexed (1=Jan)

  // Simpan template formula kolom lokal (M-T) dari baris 2
  var localFormulas = [];
  var lastCol = sheetDataMentah.getLastColumn();
  if (lastCol > 12) {
    // Kolom M=13 sampai kolom terakhir
    var localRange = sheetDataMentah.getRange(2, 13, 1, lastCol - 12);
    localFormulas = localRange.getFormulas()[0];
  }

  // Hitung jumlah baris data lama
  var oldLastRow = sheetDataMentah.getLastRow();

  // Clear area data lama (baris 2 ke bawah, kolom A sampai L)
  if (oldLastRow >= 2) {
    sheetDataMentah.getRange(2, 1, oldLastRow - 1, lastCol).clearContent();
  }

  // Set formula FILTER di A2
  var formula = '=IFERROR(SORT(FILTER(\'Data Mentah All\'!A2:L,\'Data Mentah All\'!' +
                monthCol + '2:' + monthCol + '=' + monthNumber + '),1,TRUE),"")';
  sheetDataMentah.getRange('A2').setFormula(formula);

  // Tunggu sebentar agar formula ter-resolve
  SpreadsheetApp.flush();

  // Hitung jumlah baris data baru
  var newLastRow = sheetDataMentah.getLastRow();
  var newDataRows = newLastRow - 1; // minus header row

  Logger.log('Data Mentah AHM: ' + newDataRows + ' baris data untuk bulan ' + monthNumber);

  // Extend formula kolom lokal (M-T) ke semua baris data baru
  if (localFormulas.length > 0 && newDataRows > 0) {
    extendLocalFormulas(sheetDataMentah, localFormulas, newDataRows, 13);
  }

  Logger.log('Task 3 selesai: Formula Data Mentah AHM diupdate untuk bulan ' +
             MONTH_NAMES_ID[currentMonth]);
}

/**
 * Cari kolom "Month" di header row 1 secara dinamis.
 *
 * @param {Sheet} sheet - Sheet "Data Mentah All"
 * @return {string|null} Huruf kolom (e.g. "O") atau null jika tidak ditemukan
 */
function findMonthColumn(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === 'month') {
      return columnToLetter(i + 1);
    }
  }
  return null;
}

/**
 * Extend formula kolom lokal ke semua baris data.
 * Template formula dari baris 2 di-adjust reference-nya untuk setiap baris.
 *
 * @param {Sheet} sheet - Sheet target
 * @param {Array} templates - Array formula template dari baris 2
 * @param {number} numRows - Jumlah baris data
 * @param {number} startCol - Kolom awal (13 = M)
 */
function extendLocalFormulas(sheet, templates, numRows, startCol) {
  // Set formula di baris 2 dulu
  for (var c = 0; c < templates.length; c++) {
    if (templates[c] && templates[c] !== '') {
      sheet.getRange(2, startCol + c).setFormula(templates[c]);
    }
  }

  // Jika ada lebih dari 1 baris, copy formula ke baris 3 dan seterusnya
  if (numRows > 1) {
    var sourceRange = sheet.getRange(2, startCol, 1, templates.length);
    var targetRange = sheet.getRange(3, startCol, numRows - 1, templates.length);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  }

  Logger.log('Formula lokal kolom M-T di-extend ke ' + numRows + ' baris');
}

// ============================================================================
// TASK 4: UPDATE DASHBOARD
// ============================================================================

/**
 * Update cell B1:D1 (merged) di sheet "Dashboard" dengan periode bulan berjalan.
 * Format: "Okt 2026" atau "Oct 2026"
 *
 * @param {number} currentMonth - Bulan berjalan (0-indexed)
 * @param {number} currentYear - Tahun berjalan
 */
function updateDashboard(currentMonth, currentYear) {
  var ss = SpreadsheetApp.openById(MD_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Dashboard');
  if (!sheet) {
    Logger.log('Sheet "Dashboard" tidak ditemukan!');
    return;
  }

  // Cek format existing (EN atau ID)
  var existingValue = String(sheet.getRange('B1').getValue());
  var useEnglish = /Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec/.test(existingValue);
  var monthNames = useEnglish ? MONTH_NAMES_EN : MONTH_NAMES_ID;

  var periodText = monthNames[currentMonth] + ' ' + currentYear;
  sheet.getRange('B1').setValue(periodText);

  Logger.log('Dashboard diupdate ke: ' + periodText);
}

/**
 * Update cell B1:D1 di sheet "Detail Problem" jika ada.
 */
function updateDetailProblem(currentMonth, currentYear) {
  var ss = SpreadsheetApp.openById(MD_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Detail Problem');
  if (!sheet) return;

  var existingValue = String(sheet.getRange('B1').getValue());
  var useEnglish = /Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec/.test(existingValue);
  var monthNames = useEnglish ? MONTH_NAMES_EN : MONTH_NAMES_ID;

  var periodText = monthNames[currentMonth] + ' ' + currentYear;
  sheet.getRange('B1').setValue(periodText);

  Logger.log('Detail Problem diupdate ke: ' + periodText);
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Konversi nomor kolom ke huruf (1→A, 2→B, ..., 26→Z, 27→AA).
 *
 * @param {number} column - Nomor kolom (1-indexed)
 * @return {string} Huruf kolom
 */
function columnToLetter(column) {
  var letter = '';
  while (column > 0) {
    var temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp) / 26);
  }
  return letter;
}

/**
 * Konversi huruf kolom ke nomor (A→1, B→2, ..., Z→26, AA→27).
 *
 * @param {string} letter - Huruf kolom
 * @return {number} Nomor kolom (1-indexed)
 */
function letterToColumn(letter) {
  var column = 0;
  for (var i = 0; i < letter.length; i++) {
    column = column * 26 + (letter.charCodeAt(i) - 64);
  }
  return column;
}

// ============================================================================
// MANUAL RUN FUNCTIONS (untuk testing)
// ============================================================================

/**
 * Jalankan otomasi secara manual tanpa cek H+7.
 * Berguna untuk testing atau jika perlu menjalankan di luar jadwal.
 *
 * Ubah bulan dan tahun sesuai kebutuhan sebelum menjalankan.
 */
function manualRun() {
  var currentMonth = new Date().getMonth(); // 0-indexed
  var currentYear  = new Date().getFullYear();
  var prevMonth    = currentMonth === 0 ? 11 : currentMonth - 1;

  var isNewYear = (currentMonth === 0);

  Logger.log('=== MANUAL RUN ===');
  Logger.log('Bulan berjalan: ' + MONTH_NAMES_ID[currentMonth] + ' ' + currentYear);

  if (isNewYear) {
    handleYearRollover(currentYear);
  } else {
    processRekapAllMD(currentMonth, prevMonth);
  }

  updateAhmDataMentah(currentMonth);
  updateDashboard(currentMonth, currentYear);
  updateDetailProblem(currentMonth, currentYear);

  Logger.log('=== MANUAL RUN SELESAI ===');
}

/**
 * Test: hanya proses Rekap Result All MD.
 */
function testRekapAllMD() {
  var currentMonth = new Date().getMonth();
  var prevMonth = currentMonth === 0 ? 11 : currentMonth - 1;
  processRekapAllMD(currentMonth, prevMonth);
}

/**
 * Test: hanya update Data Mentah AHM.
 */
function testAhmDataMentah() {
  var currentMonth = new Date().getMonth();
  updateAhmDataMentah(currentMonth);
}

/**
 * Test: hanya update Dashboard.
 */
function testDashboard() {
  var currentMonth = new Date().getMonth();
  var currentYear = new Date().getFullYear();
  updateDashboard(currentMonth, currentYear);
}

/**
 * Test: year rollover.
 * HATI-HATI: akan mengubah data! Hanya jalankan jika memang ingin rollover.
 */
function testYearRollover() {
  var currentYear = new Date().getFullYear();
  handleYearRollover(currentYear);
}
