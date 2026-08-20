/**
 * ============================================================================
 *  DATABASE PROJECT – GOOGLE APPS SCRIPT BACKEND
 * ============================================================================
 *  File Excel "master.xlsx" adalah TEMPLATE database untuk satu project,
 *  berisi 8 sheet: PO_Tracking, Milestone_Tracking, FAT_Schedule,
 *  Shipment_Tracking, Action_Log, Risk_Register, S_Curve_Data, Master_Data
 *  — persis seperti struktur "Kutai.xlsx" yang sudah berjalan.
 *
 *  Cara memakai untuk PROJECT BARU:
 *  1. Duplikat master.xlsx (mis. jadi "ProjectBaru.xlsx") lalu upload/import
 *     ke Google Sheets (atau File > Save as Google Sheets kalau sudah di Drive).
 *  2. Buka Google Sheet hasil import > Extensions > Apps Script, tempel
 *     SELURUH isi file Code.gs ini (satu Code.gs dipakai untuk semua project,
 *     tidak perlu diubah).
 *  3. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 *  4. Salin URL /exec, tambahkan sebagai entry baru pada array PROJECTS di
 *     index.html (field apiUrl) untuk project tsb.
 *  Setiap project punya Google Sheet + deployment sendiri, tapi memakai
 *  Code.gs dan struktur kolom yang sama persis — jadi tinggal copy-paste.
 *
 *  Kolom "Cummulative Progress (Actual Progress)" di sheet PO_Tracking TIDAK
 *  diisi manual. Setiap kali data ditambah/diubah lewat menu PO Tracking,
 *  script ini:
 *    a. Menemukan SEMUA kolom yang mengandung teks "Weight Factor Percentage"
 *       beserta bobotnya (diambil otomatis dari angka % pada nama kolom).
 *    b. Mengisi kolom bobot tsb dengan 1 jika kolom "(Actual)" milestone yang
 *       bersangkutan sudah terisi tanggal, atau 0 jika belum.
 *    c. Menulis FORMULA (bukan angka mati) ke kolom Cummulative Progress:
 *       = SUM(setiap kolom Weight Factor * bobotnya)
 *       Sehingga nilainya tetap ter-update walau sel diedit langsung di Sheet.
 * ============================================================================
 */

var SHEET_NAME_PO = 'PO_Tracking';
var CUMULATIVE_PROGRESS_HEADER = 'Cummulative Progress (Actual Progress)';
var PO_ID_HEADER = 'PO Tracking ID';

// Sheet lain dalam database per-project (mengikuti struktur Kutai.xlsx / master.xlsx)
var OTHER_SHEETS = {
  Milestone_Tracking: { keyColumn: 'Milestone_ID' },
  FAT_Schedule: { keyColumn: 'FAT_ID' },
  Shipment_Tracking: { keyColumn: 'Shipment_ID' },
  Action_Log: { keyColumn: 'Action_ID' },
  Risk_Register: { keyColumn: 'Risk_ID' },
  S_Curve_Data: { keyColumn: 'Period' },
  Master_Data: { keyColumn: 'Key' }
};

// ----------------------------------------------------------------------------
// ENTRY POINTS
// ----------------------------------------------------------------------------

function doGet(e) {
  return handleRequest_(e, 'GET');
}

function doPost(e) {
  return handleRequest_(e, 'POST');
}

function handleRequest_(e, method) {
  var action, payload;
  try {
    if (method === 'GET') {
      action = e && e.parameter ? e.parameter.action : '';
      payload = e && e.parameter ? e.parameter : {};
    } else {
      var body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
      action = body.action;
      payload = body.payload || {};
    }
    var result = routeAction_(action, payload);
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ success: false, error: err.message });
  }
}

function routeAction_(action, payload) {
  switch (action) {
    case 'getAllPO':
      return { success: true, data: getAllPO_() };
    case 'getAllMilestones':
      return { success: true, data: getAllGeneric_('Milestone_Tracking') };
    case 'getAllFAT':
      return { success: true, data: getAllGeneric_('FAT_Schedule') };
    case 'getAllShipments':
      return { success: true, data: getAllGeneric_('Shipment_Tracking') };
    case 'getAllActions':
      return { success: true, data: getAllGeneric_('Action_Log') };
    case 'getAllRisks':
      return { success: true, data: getAllGeneric_('Risk_Register') };
    case 'getAllSCurve':
      return { success: true, data: getAllGeneric_('S_Curve_Data') };
    case 'getAllMasterData':
      return { success: true, data: getAllGeneric_('Master_Data') };
    case 'getDashboardData':
      return { success: true, data: getDashboardData_() };
    case 'getCurrencyRates':
      return { success: true, data: { USD: 1, IDR: 16300, JPY: 157 } };
    case 'closeAction':
      return updateRow_('Action_Log', 'Action_ID', payload.actionId || payload.id, {
        Status: 'CLOSED',
        Closed_Date: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd')
      });
    case 'addRow':
      return addRow_(payload.sheet, payload.rowData);
    case 'updateRow':
      return updateRow_(payload.sheet, payload.keyColumn, payload.keyValue, payload.newData);
    case 'deleteRow':
      return deleteRow_(payload.sheet, payload.keyColumn, payload.keyValue);
    case 'uploadFileToDrive':
      return uploadFileToDrive_(payload.base64Data, payload.fileName, payload.mimeType);
    default:
      return { success: false, error: 'Aksi tidak dikenal: ' + action };
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------------------
// SHEET / HEADER HELPERS
// ----------------------------------------------------------------------------

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" tidak ditemukan');
  return sh;
}

function getHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  var raw = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return raw.map(function (h) {
    return String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
  });
}

function formatCell_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd');
  }
  return v;
}

function rowToObject_(headers, rowArr) {
  var obj = {};
  headers.forEach(function (h, i) {
    obj[h] = formatCell_(rowArr[i]);
  });
  return obj;
}

function isRowEmpty_(rowArr) {
  return rowArr.every(function (v) {
    return v === '' || v === null || typeof v === 'undefined';
  });
}

function columnToLetter_(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ----------------------------------------------------------------------------
// WEIGHT FACTOR / CUMMULATIVE PROGRESS LOGIC
// (inti permintaan: kolom Cummulative Progress = akumulasi semua kolom yang
//  namanya mengandung "Weight Factor Percentage")
// ----------------------------------------------------------------------------

/**
 * Menemukan semua kolom header yang mengandung "Weight Factor Percentage",
 * lalu mengekstrak bobotnya dari angka % pada nama kolom itu sendiri.
 * Contoh: "Fabrication Completion (Weight Factor Percentage 30%)" -> weight 0.30
 */
function getWeightFactorColumns_(headers) {
  var cols = [];
  headers.forEach(function (h, idx) {
    var m = h.match(/Weight Factor Percentage\s*(\d+(?:\.\d+)?)\s*%/i);
    if (m) {
      cols.push({
        index: idx, // 0-based
        header: h,
        weight: parseFloat(m[1]) / 100,
        milestone: h.split('(')[0].trim()
      });
    }
  });
  return cols;
}

function getCumulativeProgressColIndex_(headers) {
  var idx = headers.indexOf(CUMULATIVE_PROGRESS_HEADER);
  if (idx === -1) {
    idx = headers.findIndex(function (h) {
      return /Cummulative Progress/i.test(h);
    });
  }
  return idx; // 0-based, -1 jika tidak ditemukan
}

/** Bangun formula "=U{row}*0.1+AG{row}*0.05+..." dari kolom weight yang ditemukan */
function buildCumulativeFormula_(weightCols, row) {
  var parts = weightCols.map(function (c) {
    return columnToLetter_(c.index + 1) + row + '*' + c.weight;
  });
  return '=' + parts.join('+');
}

/** Tulis formula Cummulative Progress ke baris tertentu di sheet */
function applyCumulativeFormula_(sheet, headers, row) {
  var weightCols = getWeightFactorColumns_(headers);
  var cumIdx = getCumulativeProgressColIndex_(headers);
  if (cumIdx === -1 || !weightCols.length) return;
  var cell = sheet.getRange(row, cumIdx + 1);
  cell.setFormula(buildCumulativeFormula_(weightCols, row));
  cell.setNumberFormat('0%');
}

/**
 * Isi otomatis nilai kolom Weight Factor (0/1) berdasarkan apakah kolom
 * "(Actual)" milik milestone yang sama sudah terisi tanggal atau belum.
 * rowArr dimodifikasi in-place.
 */
function syncWeightFlags_(headers, rowArr, weightCols) {
  weightCols.forEach(function (c) {
    var actualHeader = c.milestone + ' (Actual)';
    var actualIdx = headers.indexOf(actualHeader);
    if (actualIdx !== -1) {
      var actualVal = rowArr[actualIdx];
      rowArr[c.index] = actualVal !== '' && actualVal !== null && typeof actualVal !== 'undefined' ? 1 : 0;
    }
  });
}

/** Utility manual: jalankan sekali dari editor Apps Script untuk merapikan seluruh sheet */
function rebuildAllCumulativeProgress() {
  var sheet = getSheet_(SHEET_NAME_PO);
  var headers = getHeaders_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var weightCols = getWeightFactorColumns_(headers);
  var range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  var values = range.getValues();
  values.forEach(function (rowArr, i) {
    if (isRowEmpty_(rowArr)) return;
    syncWeightFlags_(headers, rowArr, weightCols);
  });
  range.setValues(values);
  for (var r = 2; r <= lastRow; r++) {
    applyCumulativeFormula_(sheet, headers, r);
  }
}

// ----------------------------------------------------------------------------
// PO TRACKING ID GENERATOR
// ----------------------------------------------------------------------------

function generateNextPOId_(sheet, headers, idIdx) {
  var lastRow = sheet.getLastRow();
  var max = 0;
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
    ids.forEach(function (r) {
      var m = String(r[0]).match(/(\d+)\s*$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  var next = max + 1;
  var padded = next < 1000 ? ('000' + next).slice(-3) : String(next);
  return 'PID' + padded;
}

// ----------------------------------------------------------------------------
// GET ALL PO
// ----------------------------------------------------------------------------

function getAllPO_() {
  var sheet = getSheet_(SHEET_NAME_PO);
  var headers = getHeaders_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var cumIdx = getCumulativeProgressColIndex_(headers);
  var data = [];
  values.forEach(function (row) {
    if (isRowEmpty_(row)) return;
    var obj = rowToObject_(headers, row);
    if (cumIdx !== -1) {
      var raw = Number(row[cumIdx]) || 0;
      obj['Cumulative Progress %'] = Math.round(raw * 100);
    }
    data.push(obj);
  });
  return data;
}

/** Baca seluruh baris dari sheet lain (Milestone_Tracking, FAT_Schedule, dst) apa adanya */
function getAllGeneric_(sheetName) {
  var sheet;
  try {
    sheet = getSheet_(sheetName);
  } catch (e) {
    return [];
  }
  var headers = getHeaders_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var data = [];
  values.forEach(function (row) {
    if (isRowEmpty_(row)) return;
    data.push(rowToObject_(headers, row));
  });
  return data;
}

// ----------------------------------------------------------------------------
// DASHBOARD (ringkasan sederhana khusus PO Tracking)
// ----------------------------------------------------------------------------

function getDashboardData_() {
  var sheet;
  try {
    sheet = getSheet_(SHEET_NAME_PO);
  } catch (e) {
    return {};
  }
  var headers = getHeaders_(sheet);
  var lastRow = sheet.getLastRow();
  var data = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var catIdx = headers.indexOf('Category');
  var cumIdx = getCumulativeProgressColIndex_(headers);
  var statusIdx = headers.indexOf('Order Status (Open / Close)');

  var total = 0, onSchedule = 0, atRisk = 0, delay = 0, open_ = 0, closed = 0, progressSum = 0;
  data.forEach(function (row) {
    if (isRowEmpty_(row)) return;
    total++;
    var cat = row[catIdx];
    if (cat === 'On Schedule') onSchedule++;
    else if (cat === 'At Risk') atRisk++;
    else if (cat === 'Delay') delay++;
    var st = row[statusIdx];
    if (st === 'Open') open_++;
    else if (st === 'Close') closed++;
    progressSum += Number(row[cumIdx]) || 0;
  });

  return {
    totalPO: total,
    onSchedule: onSchedule,
    atRisk: atRisk,
    delay: delay,
    openPO: open_,
    closedPO: closed,
    avgProgressPct: total ? Math.round((progressSum / total) * 100) : 0
  };
}

// ----------------------------------------------------------------------------
// CRUD GENERIK (dipakai oleh menu PO Tracking; sheet lain bisa memakai
// helper yang sama selama nama & header sheet-nya sudah dibuat)
// ----------------------------------------------------------------------------

function addRow_(sheetName, rowData) {
  if (!sheetName) return { success: false, error: 'Sheet tidak ditentukan' };
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  var rowArr = headers.map(function (h) {
    return rowData && Object.prototype.hasOwnProperty.call(rowData, h) ? rowData[h] : '';
  });

  if (sheetName === SHEET_NAME_PO) {
    var idIdx = headers.indexOf(PO_ID_HEADER);
    if (idIdx !== -1 && !rowArr[idIdx]) {
      rowArr[idIdx] = generateNextPOId_(sheet, headers, idIdx);
    }
    var weightCols = getWeightFactorColumns_(headers);
    syncWeightFlags_(headers, rowArr, weightCols);
  }

  sheet.appendRow(rowArr);
  var newRow = sheet.getLastRow();

  if (sheetName === SHEET_NAME_PO) {
    applyCumulativeFormula_(sheet, headers, newRow);
  }

  return { success: true, message: 'Data berhasil ditambahkan' };
}

function updateRow_(sheetName, keyColumn, keyValue, newData) {
  if (!sheetName || !keyColumn) return { success: false, error: 'Parameter tidak lengkap' };
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  var keyIdx = headers.indexOf(keyColumn);
  if (keyIdx === -1) return { success: false, error: 'Kolom kunci "' + keyColumn + '" tidak ditemukan' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: 'Data tidak ditemukan' };
  var range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    if (String(values[i][keyIdx]) === String(keyValue)) {
      var rowArr = values[i];
      headers.forEach(function (h, idx) {
        if (newData && Object.prototype.hasOwnProperty.call(newData, h)) {
          rowArr[idx] = newData[h];
        }
      });

      var sheetRow = i + 2;
      if (sheetName === SHEET_NAME_PO) {
        var weightCols = getWeightFactorColumns_(headers);
        syncWeightFlags_(headers, rowArr, weightCols);
      }

      sheet.getRange(sheetRow, 1, 1, headers.length).setValues([rowArr]);

      if (sheetName === SHEET_NAME_PO) {
        applyCumulativeFormula_(sheet, headers, sheetRow);
      }

      return { success: true, message: 'Data berhasil diupdate' };
    }
  }
  return { success: false, error: 'Data dengan ' + keyColumn + ' = ' + keyValue + ' tidak ditemukan' };
}

function deleteRow_(sheetName, keyColumn, keyValue) {
  if (!sheetName || !keyColumn) return { success: false, error: 'Parameter tidak lengkap' };
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  var keyIdx = headers.indexOf(keyColumn);
  if (keyIdx === -1) return { success: false, error: 'Kolom kunci "' + keyColumn + '" tidak ditemukan' };

  var lastRow = sheet.getLastRow();
  for (var r = 2; r <= lastRow; r++) {
    var val = sheet.getRange(r, keyIdx + 1).getValue();
    if (String(val) === String(keyValue)) {
      sheet.deleteRow(r);
      return { success: true, message: 'Data berhasil dihapus' };
    }
  }
  return { success: false, error: 'Data dengan ' + keyColumn + ' = ' + keyValue + ' tidak ditemukan' };
}

// ----------------------------------------------------------------------------
// UPLOAD FILE (Unpriced PO PDF, dsb) KE GOOGLE DRIVE
// ----------------------------------------------------------------------------

function uploadFileToDrive_(base64Data, fileName, mimeType) {
  try {
    if (!base64Data || !fileName) return { success: false, error: 'File tidak lengkap' };
    var folder = getUploadFolder_();
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl(), id: file.getId() };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getUploadFolder_() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('PO_UPLOAD_FOLDER_ID');
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // folder id tersimpan tapi sudah tidak valid, buat ulang di bawah
    }
  }
  var folders = DriveApp.getFoldersByName('PO_Tracking_Uploads');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('PO_Tracking_Uploads');
  props.setProperty('PO_UPLOAD_FOLDER_ID', folder.getId());
  return folder;
}
