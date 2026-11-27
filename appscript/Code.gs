/**
 * Backend Google Apps Script untuk form Titip Jual Heru Property.
 *
 * Pasang sebagai script yang terikat (bound) pada Google Spreadsheet, atau isi
 * SPREADSHEET_ID lewat fungsi configure(). Jalankan setup() sekali sebelum
 * melakukan deploy sebagai Web App.
 */

const APP = Object.freeze({
  SHEET_NAME: 'Titip Jual',
  PHOTO_FOLDER_NAME: 'Foto Titip Jual - Heru Property',
  MAX_PHOTOS: 12,
  MAX_PHOTO_BYTES: 3 * 1024 * 1024,
  STATUS_NEW: 'Baru',
  STATUS_PUBLISHED: 'Tayang',
  HEADERS: [
    'ID', 'Waktu Masuk', 'Status', 'Nama', 'No HP', 'Jenis', 'Alamat',
    'Kecamatan', 'Kelurahan', 'Harga Jual (Juta)', 'Harga Net (Juta)', 'KPR',
    'Luas Tanah', 'Luas Bangunan', 'Lantai', 'Kamar Tidur', 'Kamar Mandi',
    'Akses Jalan', 'Listrik', 'Air', 'Sertifikat', 'Sertifikat Atas Nama',
    'PBB', 'Keterangan', 'Shareloc', 'Postingan FB', 'Latitude', 'Longitude',
    'Akurasi (m)', 'Sumber Lokasi', 'File Foto JSON', 'Folder Foto',
    'Terakhir Diubah'
  ]
});

/**
 * Opsional untuk standalone script. Jalankan dari editor satu kali:
 * configure('ID_SPREADSHEET', 'ID_FOLDER_DRIVE_OPSIONAL')
 */
function configure(spreadsheetId, folderId) {
  const props = PropertiesService.getScriptProperties();
  if (spreadsheetId) props.setProperty('SPREADSHEET_ID', String(spreadsheetId).trim());
  if (folderId) props.setProperty('PHOTO_FOLDER_ID', String(folderId).trim());
  return setup();
}

/** Membuat/memperbaiki sheet, header, validasi status, dan folder foto. */
function setup() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss);
  const folder = getPhotoRootFolder_();

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, APP.HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#d4521a')
    .setFontColor('#ffffff');
  sheet.autoResizeColumns(1, APP.HEADERS.length);
  sheet.setColumnWidth(24, 320);
  sheet.setColumnWidth(31, 260);

  const statusCol = APP.HEADERS.indexOf('Status') + 1;
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Baru', 'Diproses', 'Tayang', 'Terjual', 'Ditolak'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, statusCol, Math.max(sheet.getMaxRows() - 1, 1), 1)
    .setDataValidation(validation);

  return { spreadsheetId: ss.getId(), sheet: sheet.getName(), folderId: folder.getId() };
}

function doGet(e) {
  try {
    if (!e || !e.parameter || e.parameter.list !== '1') {
      return json_({ ok: true, service: 'Heru Property Titip Jual API' });
    }
    return json_({ ok: true, items: getPublishedListings_() });
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return json_({ ok: false, error: 'Server gagal membaca daftar.' });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const raw = e && e.postData ? e.postData.contents : '';
    if (!raw || raw.length > 45 * 1024 * 1024) throw new Error('Payload kosong atau terlalu besar.');
    const data = JSON.parse(raw);
    validateSubmission_(data);

    const id = Utilities.getUuid();
    const now = new Date();
    const photoInfo = savePhotos_(id, data.nama, data.photos);
    try {
      const sheet = getOrCreateSheet_(getSpreadsheet_());
      const record = {
        'ID': id,
        'Waktu Masuk': now,
        'Status': APP.STATUS_NEW,
        'Nama': safeCell_(data.nama),
        'No HP': safeCell_(data.wa),
        'Jenis': safeCell_(data.jenis || 'Rumah'),
        'Alamat': safeCell_(data.alamat),
        'Kecamatan': safeCell_(data.kecamatan),
        'Kelurahan': safeCell_(data.kelurahan),
        'Harga Jual (Juta)': numberOrBlank_(data.harga),
        'Harga Net (Juta)': numberOrBlank_(data.net),
        'KPR': safeCell_(data.kpr),
        'Luas Tanah': safeCell_(data.lt),
        'Luas Bangunan': safeCell_(data.lb),
        'Lantai': safeCell_(data.lantai),
        'Kamar Tidur': safeCell_(data.kt),
        'Kamar Mandi': safeCell_(data.km),
        'Akses Jalan': safeCell_(data.akses),
        'Listrik': safeCell_(data.listrik),
        'Air': safeCell_(data.air),
        'Sertifikat': safeCell_(data.sert),
        'Sertifikat Atas Nama': safeCell_(data.atasnama),
        'PBB': safeCell_(data.pbb),
        'Keterangan': safeCell_(data.ket),
        'Shareloc': safeCell_(data.shareloc),
        'Postingan FB': safeCell_(data.fbpost),
        'Latitude': coordinateOrBlank_(data.lat, -90, 90),
        'Longitude': coordinateOrBlank_(data.lon, -180, 180),
        'Akurasi (m)': numberOrBlank_(data.akurasi),
        'Sumber Lokasi': safeCell_(data.sumber),
        'File Foto JSON': JSON.stringify(photoInfo.files),
        'Folder Foto': photoInfo.folderUrl,
        'Terakhir Diubah': now
      };
      sheet.appendRow(APP.HEADERS.map(h => record[h] === undefined ? '' : record[h]));
    } catch (err) {
      photoInfo.folder.setTrashed(true);
      throw err;
    }

    return json_({ ok: true, id: id });
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return json_({ ok: false, error: publicError_(err) });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function validateSubmission_(data) {
  if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('Data tidak valid.');
  if (!String(data.nama || '').trim()) throw new Error('Nama wajib diisi.');
  if (String(data.nama).length > 150) throw new Error('Nama terlalu panjang.');
  if (!Array.isArray(data.photos) || data.photos.length < 1) throw new Error('Minimal satu foto.');
  if (data.photos.length > APP.MAX_PHOTOS) throw new Error('Maksimal 12 foto.');
  if (!data.lat && !data.lon && !String(data.shareloc || '').trim() && !String(data.alamat || '').trim()) {
    throw new Error('Lokasi wajib diisi.');
  }
  if (data.shareloc && !/^https:\/\//i.test(String(data.shareloc))) throw new Error('Link shareloc tidak valid.');
  if (data.fbpost && !/^https:\/\//i.test(String(data.fbpost))) throw new Error('Link Facebook tidak valid.');
  if (data.wa && !/^(?:\+?62|0)8\d{7,12}$/.test(String(data.wa).replace(/[\s.\-]/g, ''))) {
    throw new Error('Nomor HP tidak valid.');
  }
}

function savePhotos_(id, name, photos) {
  const root = getPhotoRootFolder_();
  const folderName = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyyMMdd-HHmmss') +
    ' - ' + cleanFileName_(name) + ' - ' + id.slice(0, 8);
  const folder = root.createFolder(folderName);
  const files = [];
  try {
    photos.forEach((photo, index) => {
      const base64 = String(photo && photo.data || '').replace(/^data:image\/\w+;base64,/, '');
      if (!base64) throw new Error('Ada foto yang kosong.');
      const bytes = Utilities.base64Decode(base64);
      if (bytes.length > APP.MAX_PHOTO_BYTES) throw new Error('Ukuran salah satu foto terlalu besar.');
      const type = /^image\/(?:jpeg|png|webp)$/i.test(String(photo.type || '')) ? photo.type : 'image/jpeg';
      const ext = type === 'image/png' ? '.png' : type === 'image/webp' ? '.webp' : '.jpg';
      const file = folder.createFile(Utilities.newBlob(bytes, type, String(index + 1).padStart(2, '0') + ext));
      files.push({ id: file.getId(), name: file.getName() });
    });
    return { folder: folder, folderUrl: folder.getUrl(), files: files };
  } catch (err) {
    folder.setTrashed(true);
    throw err;
  }
}

function getPublishedListings_() {
  const sheet = getOrCreateSheet_(getSpreadsheet_());
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, APP.HEADERS.length).getValues();
  const col = Object.fromEntries(APP.HEADERS.map((h, i) => [h, i]));

  return values.filter(row => String(row[col['Status']]).trim().toLowerCase() === APP.STATUS_PUBLISHED.toLowerCase())
    .map(row => {
      const lat = finiteNumber_(row[col['Latitude']]);
      const lon = finiteNumber_(row[col['Longitude']]);
      if (lat === null || lon === null) return null;
      const photoMeta = parsePhotoMeta_(row[col['File Foto JSON']]);
      const photos = photoMeta.map(meta => {
        makeFilePublic_(meta.id);
        return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(meta.id) + '&sz=w1200';
      });
      const kec = String(row[col['Kecamatan']] || '').trim();
      const kel = String(row[col['Kelurahan']] || '').trim();
      const address = String(row[col['Alamat']] || '').trim();
      const area = [kel, kec].filter(Boolean).join(', ') || address;
      const shareloc = String(row[col['Shareloc']] || '').trim();
      return {
        name: String(row[col['Nama']] || 'Properti Dijual'),
        area: area,
        lat: lat,
        lon: lon,
        maps: shareloc || 'https://www.google.com/maps?q=' + lat + ',' + lon,
        price: finiteNumber_(row[col['Harga Jual (Juta)']]),
        lt: valueOrNull_(row[col['Luas Tanah']]),
        lb: valueOrNull_(row[col['Luas Bangunan']]),
        kt: valueOrNull_(row[col['Kamar Tidur']]),
        km: valueOrNull_(row[col['Kamar Mandi']]),
        sert: valueOrNull_(row[col['Sertifikat']]),
        photos: photos
      };
    }).filter(Boolean);
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('SPREADSHEET_ID belum dikonfigurasi.');
  return active;
}

function getOrCreateSheet_(ss) {
  let sheet = ss.getSheetByName(APP.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(APP.SHEET_NAME);
  const current = sheet.getRange(1, 1, 1, APP.HEADERS.length).getDisplayValues()[0];
  if (current.every(v => !v)) sheet.getRange(1, 1, 1, APP.HEADERS.length).setValues([APP.HEADERS]);
  else if (APP.HEADERS.some((h, i) => current[i] !== h)) {
    throw new Error('Header sheet "' + APP.SHEET_NAME + '" tidak sesuai. Jalankan setup pada sheet kosong.');
  }
  return sheet;
}

function getPhotoRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (_) { props.deleteProperty('PHOTO_FOLDER_ID'); }
  }
  const folder = DriveApp.createFolder(APP.PHOTO_FOLDER_NAME);
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

function parsePhotoMeta_(value) {
  try {
    const parsed = JSON.parse(String(value || '[]'));
    return Array.isArray(parsed) ? parsed.filter(x => x && /^[\w-]+$/.test(String(x.id || ''))) : [];
  } catch (_) { return []; }
}

function makeFilePublic_(id) {
  try {
    const file = DriveApp.getFileById(id);
    if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
  } catch (err) {
    console.warn('Foto tidak dapat dipublikasikan: ' + id + ' - ' + err);
  }
}

function safeCell_(value) {
  let text = String(value == null ? '' : value).trim();
  if (text.length > 5000) text = text.slice(0, 5000);
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function numberOrBlank_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const n = Number(String(value).replace(',', '.'));
  return Number.isFinite(n) && n >= 0 ? n : '';
}

function coordinateOrBlank_(value, min, max) {
  const n = finiteNumber_(value);
  return n !== null && n >= min && n <= max ? n : '';
}

function finiteNumber_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function valueOrNull_(value) {
  return value === '' || value === null || value === undefined ? null : value;
}

function cleanFileName_(value) {
  return String(value || 'Properti').replace(/[\\/:*?"<>|#%{}~&]/g, '-').replace(/\s+/g, ' ').trim().slice(0, 80) || 'Properti';
}

function publicError_(err) {
  const allowed = /^(Data|Nama|Minimal|Maksimal|Lokasi|Link|Nomor|Ada foto|Ukuran)/;
  const message = err && err.message ? err.message : '';
  return allowed.test(message) ? message : 'Data gagal disimpan. Silakan coba lagi.';
}

function json_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
