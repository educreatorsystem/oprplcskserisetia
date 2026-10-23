/**
 * SISTEM OPR PLC — SK SERI SETIA
 * Backend Google Apps Script (Code.gs)
 *
 * Spreadsheet: 11F7GYjSjHcrVafEet_LLrzHqytpQ9f5lEI13P4D1XvM
 * Folder Drive: 1oCNCL11GL1ane_LQDmN_SshSUc3EWF3n
 *
 * PENTING:
 * 1. Gantikan SEMUA kod lama dalam Code.gs dengan kod ini.
 * 2. Jalankan fungsi setupSistem sekali dan benarkan akses.
 * 3. Deploy sebagai Web app: Execute as Me, Who has access Anyone.
 * 4. Kod ini sengaja TIDAK menggunakan file.setSharing(), kerana arahan itu
 *    sering disekat oleh akaun Google Workspace sekolah.
 */

var SPREADSHEET_ID = '11F7GYjSjHcrVafEet_LLrzHqytpQ9f5lEI13P4D1XvM';
var DRIVE_FOLDER_ID = '1oCNCL11GL1ane_LQDmN_SshSUc3EWF3n';
var SHEET_NAME = 'OPR_PLC';
var MAX_IMAGE_BYTES = 2 * 1024 * 1024;

var HEADERS = [
  'id', 'tajuk', 'bidang', 'alat', 'domain', 'tempat',
  'tarikhMula', 'tarikhTamat', 'durasi', 'matapelajaran',
  'pengurusan', 'fokus', 'improve', 'share', 'catatan',
  'penyedia', 'jawatanPenyedia', 'pengesah', 'jawatanPengesah',
  'img1', 'img2', 'img3', 'updatedAt', 'guruTerlibat'
];

function doGet() {
  return jsonOutput_({
    ok: true,
    system: 'Sistem OPR PLC SK Seri Setia',
    message: 'Apps Script aktif.',
    time: new Date().toISOString()
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Permintaan kosong atau tidak sah.');
    }

    var request = JSON.parse(e.postData.contents);
    var action = String(request.action || '').toLowerCase();

    if (action === 'create') {
      return jsonOutput_(saveRecord_(request.record, request.files || {}, false));
    }
    if (action === 'update') {
      return jsonOutput_(saveRecord_(request.record, request.files || {}, true));
    }
    if (action === 'delete') {
      return jsonOutput_(deleteRecord_(request.id));
    }

    throw new Error('Tindakan tidak dikenali: ' + action);
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return jsonOutput_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** Jalankan fungsi ini sekali dari editor Apps Script sebelum deploy. */
function setupSistem() {
  var sheet = getSheet_();
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  console.log('Google Sheet: ' + sheet.getParent().getName());
  console.log('Tab: ' + sheet.getName());
  console.log('Folder Drive: ' + folder.getName());
  console.log('Sistem sedia untuk digunakan.');

  return 'BERJAYA — Sheet dan folder Drive boleh diakses.';
}

function saveRecord_(inputRecord, files, isUpdate) {
  if (!inputRecord || typeof inputRecord !== 'object') {
    throw new Error('Maklumat laporan tidak diterima.');
  }

  var record = normalizeRecord_(inputRecord);
  validateRecord_(record);

  var sheet = getSheet_();
  var rowNumber = findRowById_(sheet, record.id);

  if (isUpdate && rowNumber < 2) {
    throw new Error('Rekod untuk dikemas kini tidak ditemui. Muat semula sistem.');
  }
  if (!isUpdate && rowNumber >= 2) {
    throw new Error('ID rekod telah wujud. Sila cuba simpan semula.');
  }

  var previous = rowNumber >= 2 ? rowToRecord_(sheet.getRange(rowNumber, 1, 1, HEADERS.length).getValues()[0]) : null;
  var uploadedIds = [];

  try {
    record.images = saveImages_(record.id, files, previous ? previous.images : record.images, uploadedIds);
    record.updatedAt = new Date().toISOString();

    var values = recordToRow_(record);
    if (rowNumber >= 2) {
      sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
    SpreadsheetApp.flush();

    return { ok: true, action: isUpdate ? 'update' : 'create', record: record };
  } catch (error) {
    // Bersihkan hanya fail baharu jika penulisan ke Sheet gagal.
    uploadedIds.forEach(function (fileId) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (ignore) {}
    });
    throw error;
  }
}

function deleteRecord_(id) {
  id = cleanText_(id, 150);
  if (!/^opr_[A-Za-z0-9_-]+$/.test(id)) {
    throw new Error('ID rekod tidak sah.');
  }

  var sheet = getSheet_();
  var rowNumber = findRowById_(sheet, id);
  if (rowNumber < 2) {
    throw new Error('Rekod tidak ditemui atau telah dipadam.');
  }

  sheet.deleteRow(rowNumber);
  SpreadsheetApp.flush();

  // Fail gambar tidak dipadam secara automatik supaya tiada gambar lain
  // terpadam jika folder atau fail dikongsi dengan rekod lain.
  return { ok: true, action: 'delete', id: id };
}

function saveImages_(recordId, files, existingImages, uploadedIds) {
  var result = {
    img1: cleanUrl_(existingImages && existingImages.img1),
    img2: cleanUrl_(existingImages && existingImages.img2),
    img3: cleanUrl_(existingImages && existingImages.img3)
  };

  var hasUpload = ['img1', 'img2', 'img3'].some(function (key) {
    return files && files[key] && files[key].base64;
  });
  if (!hasUpload) return result;

  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  ['img1', 'img2', 'img3'].forEach(function (key, index) {
    var item = files && files[key];
    if (!item || !item.base64) return;

    var mimeType = String(item.mimeType || '').toLowerCase();
    if (['image/jpeg', 'image/png', 'image/webp'].indexOf(mimeType) === -1) {
      throw new Error('Gambar ' + (index + 1) + ' mesti berformat JPG, PNG atau WEBP.');
    }

    var bytes;
    try {
      bytes = Utilities.base64Decode(String(item.base64));
    } catch (error) {
      throw new Error('Data Gambar ' + (index + 1) + ' tidak sah.');
    }
    if (!bytes.length || bytes.length > MAX_IMAGE_BYTES) {
      throw new Error('Saiz Gambar ' + (index + 1) + ' mesti tidak melebihi 2MB.');
    }

    var extension = mimeType === 'image/png' ? '.png' : mimeType === 'image/webp' ? '.webp' : '.jpg';
    var filename = safeFilename_(recordId + '_' + key + '_' + Date.now() + extension);
    var blob = Utilities.newBlob(bytes, mimeType, filename);
    var file = folder.createFile(blob);
    uploadedIds.push(file.getId());

    // Jangan gunakan file.setSharing() di sini. Fail mewarisi akses folder.
    result[key] = 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(file.getId()) + '&sz=w1600';
  });

  return result;
}

function getSheet_() {
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  } else {
    var currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
    var mismatch = HEADERS.some(function (header, index) {
      return String(currentHeaders[index] || '').trim() !== header;
    });
    if (mismatch) {
      throw new Error('Susunan tajuk lajur tab OPR_PLC tidak sepadan. Jangan ubah baris pertama Google Sheet.');
    }
  }

  return sheet;
}

function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  var match = sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(id)
    .matchEntireCell(true)
    .findNext();
  return match ? match.getRow() : -1;
}

function normalizeRecord_(source) {
  var images = source.images && typeof source.images === 'object' ? source.images : {};
  return {
    id: cleanText_(source.id, 150),
    tajuk: cleanText_(source.tajuk, 500),
    bidang: cleanText_(source.bidang, 150),
    alat: cleanText_(source.alat, 200),
    domain: cleanArray_(source.domain, 20, 200),
    tempat: cleanText_(source.tempat, 300),
    tarikhMula: cleanDate_(source.tarikhMula),
    tarikhTamat: cleanDate_(source.tarikhTamat),
    durasi: Math.max(1, Math.min(14400, Number(source.durasi) || 60)),
    matapelajaran: cleanArray_(source.matapelajaran, 50, 150),
    pengurusan: cleanArray_(source.pengurusan, 50, 200),
    fokus: cleanText_(source.fokus, 5000),
    improve: cleanText_(source.improve, 5000),
    share: cleanText_(source.share, 5000),
    catatan: cleanText_(source.catatan, 5000),
    penyedia: cleanText_(source.penyedia, 250),
    jawatanPenyedia: cleanText_(source.jawatanPenyedia, 250),
    pengesah: cleanText_(source.pengesah, 250),
    jawatanPengesah: cleanText_(source.jawatanPengesah, 250),
    images: {
      img1: cleanUrl_(images.img1),
      img2: cleanUrl_(images.img2),
      img3: cleanUrl_(images.img3)
    },
    updatedAt: '',
    guruTerlibat: cleanArray_(source.guruTerlibat, 500, 250)
  };
}

function validateRecord_(record) {
  if (!/^opr_[A-Za-z0-9_-]+$/.test(record.id)) throw new Error('ID laporan tidak sah.');
  if (!record.tajuk) throw new Error('Tajuk program diperlukan.');
  if (!record.bidang) throw new Error('Bidang diperlukan.');
  if (!record.alat) throw new Error('Alat kolaboratif diperlukan.');
  if (!record.domain.length) throw new Error('Pilih sekurang-kurangnya satu Domain SGM.');
  if (!record.tarikhMula || !record.tarikhTamat) throw new Error('Tarikh mula dan tarikh tamat diperlukan.');
  if (record.tarikhTamat < record.tarikhMula) throw new Error('Tarikh tamat tidak boleh sebelum tarikh mula.');
  if (!record.penyedia) throw new Error('Nama guru penyedia diperlukan.');
}

function recordToRow_(record) {
  return [
    record.id, record.tajuk, record.bidang, record.alat,
    JSON.stringify(record.domain), record.tempat, record.tarikhMula,
    record.tarikhTamat, record.durasi, JSON.stringify(record.matapelajaran),
    JSON.stringify(record.pengurusan), record.fokus, record.improve,
    record.share, record.catatan, record.penyedia, record.jawatanPenyedia,
    record.pengesah, record.jawatanPengesah, record.images.img1,
    record.images.img2, record.images.img3, record.updatedAt,
    JSON.stringify(record.guruTerlibat)
  ];
}

function rowToRecord_(row) {
  return {
    id: String(row[0] || ''),
    tajuk: String(row[1] || ''),
    bidang: String(row[2] || ''),
    alat: String(row[3] || ''),
    domain: parseArray_(row[4]),
    tempat: String(row[5] || ''),
    tarikhMula: cellDate_(row[6]),
    tarikhTamat: cellDate_(row[7]),
    durasi: Number(row[8]) || 0,
    matapelajaran: parseArray_(row[9]),
    pengurusan: parseArray_(row[10]),
    fokus: String(row[11] || ''),
    improve: String(row[12] || ''),
    share: String(row[13] || ''),
    catatan: String(row[14] || ''),
    penyedia: String(row[15] || ''),
    jawatanPenyedia: String(row[16] || ''),
    pengesah: String(row[17] || ''),
    jawatanPengesah: String(row[18] || ''),
    images: { img1: String(row[19] || ''), img2: String(row[20] || ''), img3: String(row[21] || '') },
    updatedAt: cellDateTime_(row[22]),
    guruTerlibat: parseArray_(row[23])
  };
}

function cleanText_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength || 5000);
}

function cleanArray_(value, maxItems, maxLength) {
  if (!Array.isArray(value)) return [];
  return value.slice(0, maxItems).map(function (item) {
    return cleanText_(item, maxLength);
  }).filter(function (item) { return item !== ''; });
}

function cleanDate_(value) {
  var text = String(value || '').slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : '';
}

function cleanUrl_(value) {
  var text = String(value || '').trim();
  return /^https:\/\//i.test(text) ? text.slice(0, 2000) : '';
}

function safeFilename_(value) {
  return String(value || 'gambar.jpg').replace(/[^A-Za-z0-9._-]+/g, '_').slice(0, 180);
}

function parseArray_(value) {
  if (Array.isArray(value)) return value;
  try {
    var parsed = JSON.parse(String(value || '[]'));
    return Array.isArray(parsed) ? parsed : [];
  } catch (ignore) {
    return [];
  }
}

function cellDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur', 'yyyy-MM-dd');
  }
  return String(value || '').slice(0, 10);
}

function cellDateTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.toISOString();
  return String(value || '');
}

function jsonOutput_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
