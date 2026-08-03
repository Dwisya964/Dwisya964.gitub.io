/**
 * ==============================================================================
 * KODE.GS - BACKEND GOOGLE APPS SCRIPT (FULL & LENGKAP)
 * ==============================================================================
 * Hubungkan file ini ke Google Sheets Anda.
 * Cara Penggunaan:
 * 1. Buka Google Sheets baru -> Ekstensi -> Apps Script
 * 2. Salin seluruh isi file Kode.gs ini ke editor Apps Script
 * 3. Salin seluruh isi file Index.html ke file Index.html di Apps Script
 * 4. Klik 'Terapkan' (Deploy) -> 'Penataan baru' (New Deployment) -> Web App
 * 5. Jalankan authorizeDriveAccess() satu kali dari editor Apps Script
 *    dan izinkan akses Google Drive saat diminta.
 * 6. Deploy sebagai Web App: Execute as = Me, Who has access = Anyone.
 */

// ==============================================================================
// KONFIGURASI IDENTITAS APLIKASI
// ==============================================================================
var APP_NAME = 'dwisyafitriproject store';
var APP_SHORT_NAME = 'dwisyafitriproject store';
var RECEIPT_FOLDER_PROPERTY = 'RECEIPT_FOLDER_ID';
var RECEIPT_SHEET_NAME = 'StrukPDF';
var RECEIPT_CHUNK_SIZE = 45000;

// Ikon SVG disajikan oleh doGet agar manifest dan shortcut memakai ikon yang sama.
// Jika nama toko diganti, ubah juga teks di bagian <text> di bawah ini.
var APP_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">' +
  '<rect width="512" height="512" rx="96" fill="#1D4ED8"/>' +
  '<g fill="none" stroke="#fff" stroke-width="24" stroke-linecap="round" stroke-linejoin="round">' +
  '<rect x="140" y="210" width="232" height="196" rx="16"/>' +
  '<path d="M196 210c0-36 27-76 60-76s60 40 60 76"/>' +
  '<line x1="196" y1="290" x2="316" y2="290"/>' +
  '</g>' +
  '<text x="256" y="462" font-family="Arial,Helvetica,sans-serif" font-size="52" font-weight="700" fill="#BFDBFE" text-anchor="middle" letter-spacing="3">STORE</text>' +
  '</svg>';

// 1. FUNGSI UNTUK MENAMPILKAN HALAMAN WEB APP + MANIFEST/IKON SHORTCUT
function doGet(e) {
  var params = (e && e.parameter) || {};

  // Link struk publik. PDF disimpan di sheet StrukPDF sehingga pengiriman
  // struk tidak bergantung pada izin DriveApp.
  if (params.receipt) {
    return renderReceiptPage_(String(params.receipt));
  }

  // Manifest harus berupa endpoint nyata, bukan Blob URL sementara di browser.
  if (params.manifest === '1') {
    return ContentService
      .createTextOutput(JSON.stringify(getWebAppManifest_()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Endpoint ikon dipakai oleh browser ketika membuat shortcut.
  if (params.icon === '1') {
    return ContentService
      .createTextOutput(APP_ICON_SVG)
      .setMimeType(ContentService.MimeType.XML);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppManifest_() {
  var webAppUrl = ScriptApp.getService().getUrl() || '';
  var iconUrl = webAppUrl ? webAppUrl + '?icon=1' : '?icon=1';

  return {
    name: APP_NAME,
    short_name: APP_SHORT_NAME,
    description: 'Kasir POS & Arus Cash Keuangan',
    start_url: webAppUrl || './',
    scope: webAppUrl || './',
    display: 'standalone',
    orientation: 'portrait',
    background_color: '#EFF6FF',
    theme_color: '#1D4ED8',
    icons: [
      { src: iconUrl, sizes: '192x192', type: 'image/svg+xml', purpose: 'any maskable' },
      { src: iconUrl, sizes: '512x512', type: 'image/svg+xml', purpose: 'any maskable' }
    ]
  };
}

// 2. FUNGSI MENERIMA PERMINTAAN POST (API ROUTER)
function doPost(e) {
  try {
    var contents = e.postData ? e.postData.contents : '{}';
    var params = JSON.parse(contents || '{}');
    var action = params.action;
    var payload = params.payload;

    var result = handleAction(action, payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      message: 'Server Error: ' + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 3. FUNGSI ROUTER UTAMA DARI FRONTEND / GAS CALL
function handleAction(action, payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  switch (action) {
    // --- MANAJEMEN PRODUK ---
    case 'getProduk':
      return getSheetData(ss, 'Produk');
    case 'saveProduk':
      return saveSheetRow(ss, 'Produk', payload);
    case 'deleteProduk':
      return deleteSheetRow(ss, 'Produk', payload);

    // --- MANAJEMEN PENJUALAN (KASIR POS MULTI-ITEM) ---
    case 'getPenjualan':
      return getSheetData(ss, 'Penjualan');
    case 'savePenjualan':
      return saveSheetRow(ss, 'Penjualan', payload);
    case 'deletePenjualan':
      return deleteSheetRow(ss, 'Penjualan', payload);

    // --- MANAJEMEN PENGELUARAN (ARUS CASH KEUANGAN) ---
    case 'getPengeluaran':
      return getSheetData(ss, 'Pengeluaran');
    case 'savePengeluaran':
      return saveSheetRow(ss, 'Pengeluaran', payload);
    case 'deletePengeluaran':
      return deleteSheetRow(ss, 'Pengeluaran', payload);

    // --- MANAJEMEN ULASAN / REVIEW PELANGGAN ---
    case 'getReview':
      return getSheetData(ss, 'Review');
    case 'addReview':
    case 'updateReview':
      return saveSheetRow(ss, 'Review', payload);
    case 'deleteReview':
      return deleteSheetRow(ss, 'Review', payload);

    // --- AUTENTIKASI ADMIN & UPLOAD FOTO ---
    case 'loginAdmin':
      return checkAdminLogin(ss, payload);
    case 'uploadFoto':
      return uploadFotoDrive(payload);

    // =========================================================
    // Simpan PDF tanpa Drive dan kirim link struk
    // =========================================================
    case 'uploadReceiptPdfToDrive':
      return uploadReceiptPdfToDrive(payload);

    // Kirim pesan teks + link PDF otomatis melalui WhatsApp Cloud API.
    case 'sendReceiptWhatsApp':
      return sendReceiptWhatsApp(payload);

    // [LAMA - dipertahankan untuk kompatibilitas jika masih diperlukan]
    case 'createReceiptPdfAndSendWhatsApp':
      return createReceiptPdfAndSendWhatsApp(payload);

    default:
      return { ok: false, message: 'Aksi tidak dikenal: ' + action };
  }
}

// 4. MEMBACA DATA DARI TABEL GOOGLE SHEETS
function getSheetData(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initSheetHeaderAndDefaultData(sheet, sheetName);
  }

  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return { ok: true, data: [] };

  var headers = values[0];
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var headerKey = String(headers[j]).trim();
      obj[headerKey] = row[j];
    }
    list.push(obj);
  }
  return { ok: true, data: list };
}

// 5. MENYIMPAN / MENGUPDATE BARIS DATA SPREADSHEET
function saveSheetRow(ss, sheetName, payload) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initSheetHeaderAndDefaultData(sheet, sheetName);
  }

  var data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  var values = sheet.getDataRange().getValues();
  var headers = values[0];

  // Buat ID baru jika data baru
  if (!data.ID) {
    var prefix = sheetName.substring(0, 1).toUpperCase();
    data.ID = prefix + new Date().getTime();
  }

  // Khusus Penjualan: Lengkapi Otomatis Tanggal & No Transaksi jika kosong
  if (sheetName === 'Penjualan') {
    if (!data.TANGGAL) {
      data.TANGGAL = new Date().toISOString();
    }
    if (!data.NO_TRANSAKSI) {
      var d = new Date();
      var dateStr = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
      data.NO_TRANSAKSI = 'TRX-' + dateStr + '-' + Math.floor(1000 + Math.random() * 9000);
    }
  }

  // Khusus Pengeluaran: Lengkapi Otomatis Tanggal jika kosong
  if (sheetName === 'Pengeluaran' && !data.TANGGAL) {
    data.TANGGAL = new Date().toISOString();
  }

  // Cari index baris jika ID sudah ada (Update)
  var rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.ID)) {
      rowIndex = i + 1;
      break;
    }
  }

  var rowData = [];
  for (var h = 0; h < headers.length; h++) {
    var key = String(headers[h]).trim();
    rowData.push(data[key] !== undefined ? data[key] : '');
  }

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return getSheetData(ss, sheetName);
}

// 6. MENGHAPUS BARIS DATA SPREADSHEET
function deleteSheetRow(ss, sheetName, payload) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: false, message: 'Sheet tidak ditemukan.' };

  var values = sheet.getDataRange().getValues();
  var targetId = typeof payload === 'object' ? payload.ID : payload;

  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(targetId)) {
      sheet.deleteRow(i + 1);
      return getSheetData(ss, sheetName);
    }
  }
  return { ok: false, message: 'Data ID tidak ditemukan.' };
}

// 7. FUNGSI LOGIN ADMIN
function checkAdminLogin(ss, payload) {
  var d = typeof payload === 'string' ? JSON.parse(payload) : payload;
  var user = String(d.username || d.USERNAME || '').trim();
  var pass = String(d.password || d.PASSWORD || '').trim();

  // Validasi default admin
  if ((user === 'admin' && pass === 'admin123') || (user === 'admin' && pass === 'admin')) {
    return { ok: true, username: 'admin' };
  }

  // Cek sheet Admin jika ada
  var sheet = ss.getSheetByName('Admin');
  if (sheet) {
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === user && String(values[i][1]).trim() === pass) {
        return { ok: true, username: user };
      }
    }
  }

  return { ok: false, message: 'Username atau Password salah.' };
}

// 8. INITIALISASI HEADER & SAMPLE DATA PADA SHEETS BARU
function initSheetHeaderAndDefaultData(sheet, sheetName) {
  if (sheetName === 'Produk') {
    sheet.appendRow(['ID', 'NAMA', 'DESKRIPSI', 'HARGA', 'FOTO_URL', 'FOTO_URL2', 'FOTO_URL3', 'LINK', 'LINK_DESAIN', 'STATUS']);
    sheet.appendRow(['P001', 'Es Kopi Susu Aren', 'Kopi espresso dengan susu segar dan gula aren murni.', 18000, 'https://images.unsplash.com/photo-1517701604599-bb29b565090c?w=500', '', '', '', '', 'ditampilkan']);
    sheet.appendRow(['P002', 'Roti Bakar Cokelat Keju', 'Roti bakar renyah dengan isian cokelat lumer dan keju parut melimpah.', 22000, 'https://images.unsplash.com/photo-1584776296944-ab6fb57b0bdd?w=500', '', '', '', '', 'ditampilkan']);
  } else if (sheetName === 'Penjualan') {
    sheet.appendRow(['ID', 'NO_TRANSAKSI', 'TANGGAL', 'NAMA_PEMBELI', 'WHATSAPP', 'PRODUK', 'HARGA', 'BAYAR', 'METODE', 'KEMBALIAN']);
    sheet.appendRow(['S001', 'TRX0001', new Date().toISOString(), 'Pelanggan Umum', '628123456789', 'Es Kopi Susu Aren (2x), Roti Bakar (1x)', 58000, 60000, 'Cash', 2000]);
  } else if (sheetName === 'Pengeluaran') {
    sheet.appendRow(['ID', 'TANGGAL', 'KETERANGAN', 'JUMLAH']);
    sheet.appendRow(['E001', '2025-05-05T00:00:00.000Z', 'Pembelian Bahan Baku Kopi & Roti', 15500000]);
    sheet.appendRow(['E002', '2025-05-12T00:00:00.000Z', 'Peralatan & Perlengkapan Dapur', 10000000]);
    sheet.appendRow(['E003', '2025-05-20T00:00:00.000Z', 'Gaji Karyawan', 8000000]);
    sheet.appendRow(['E004', '2025-05-27T00:00:00.000Z', 'Biaya Operasional & Listrik', 7350000]);
  } else if (sheetName === 'Review') {
    sheet.appendRow(['ID', 'NAMA', 'WHATSAPP', 'RATING', 'REVIEW', 'TANGGAL', 'BALASAN']);
    sheet.appendRow(['R001', 'Andi', '628123456789', 5, 'Kopi susu aren-nya mantap sekali!', '2025-05-28T10:00:00.000Z', 'Terima kasih mas Andi!']);
  } else if (sheetName === 'Admin') {
    sheet.appendRow(['USERNAME', 'PASSWORD']);
    sheet.appendRow(['Tokodwi#1', 'KayaRaya@2026']);
  }
}

// 9. UPLOAD FOTO KE GOOGLE DRIVE
function uploadFotoDrive(payload) {
  try {
    var d = typeof payload === 'string' ? JSON.parse(payload) : payload;
    var base64Data = d.base64;
    var fileName = d.filename || ('foto_' + new Date().getTime() + '.jpg');

    if (!base64Data) return { ok: false, message: 'Data base64 tidak valid.' };

    var contentType = base64Data.substring(5, base64Data.indexOf(';'));
    var bytes = Utilities.base64Decode(base64Data.substring(base64Data.indexOf(',') + 1));
    var blob = Utilities.newBlob(bytes, contentType, fileName);

    var file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { ok: true, url: 'https://lh3.googleusercontent.com/d/' + file.getId() };
  } catch (err) {
    return { ok: false, message: 'Gagal upload foto: ' + err.toString() };
  }
}

// ============================================================
// Kompatibilitas aksi lama uploadReceiptPdfToDrive
// ============================================================
/**
 * Nama aksi lama tetap dipertahankan agar halaman yang masih tersimpan di
 * cache HP tidak rusak. Implementasinya sekarang tidak memakai DriveApp.
 */
function uploadReceiptPdfToDrive(payload) {
  return saveReceiptPdfToSheet_(payload);
}

/**
 * Mengambil folder struk dari Script Properties, atau membuat folder baru.
 *
 * Sengaja tidak memakai getFoldersByName(). Pembacaan daftar folder dapat
 * memunculkan error DriveApp.getFoldersByName sebelum otorisasi selesai.
 * Folder dibuat dan ID-nya disimpan sekali setelah Drive diotorisasi.
 */
function getReceiptFolder_(folderName) {
  var properties = PropertiesService.getScriptProperties();
  var savedId = properties.getProperty(RECEIPT_FOLDER_PROPERTY);

  if (savedId) {
    try {
      return DriveApp.getFolderById(savedId);
    } catch (ignored) {
      properties.deleteProperty(RECEIPT_FOLDER_PROPERTY);
    }
  }

  var folder = DriveApp.createFolder(folderName || 'Struk Transaksi dwisyafitriproject');
  properties.setProperty(RECEIPT_FOLDER_PROPERTY, folder.getId());
  return folder;
}

/**
 * Jalankan sekali secara manual dari editor Apps Script untuk meminta izin Drive
 * dan menyiapkan folder struk sebelum Web App dipakai pelanggan.
 */
function setupReceiptFolder() {
  var folder = getReceiptFolder_('Struk Transaksi dwisyafitriproject');
  Logger.log('Folder struk siap: ' + folder.getUrl());
  return { ok: true, folderId: folder.getId(), folderUrl: folder.getUrl() };
}

/**
 * Alias yang mudah ditemukan di menu Run.
 * Jalankan fungsi ini satu kali, lalu klik Review permissions / Allow.
 */
function authorizeDriveAccess() {
  var folder = getReceiptFolder_('Struk Transaksi dwisyafitriproject');
  return {
    ok: true,
    message: 'Izin Google Drive sudah aktif.',
    folderId: folder.getId(),
    folderUrl: folder.getUrl()
  };
}

/**
 * Mengunggah PDF ke Drive lalu mengirim teks berisi link PDF melalui
 * WhatsApp Cloud API. Jika API belum dikonfigurasi, hasilnya dikembalikan
 * sebagai fallback agar frontend dapat membuka chat wa.me.
 *
 * Script Properties yang diperlukan:
 * - WHATSAPP_ACCESS_TOKEN
 * - WHATSAPP_PHONE_NUMBER_ID
 * - WHATSAPP_API_VERSION (opsional, default v20.0)
 */
function sendReceiptWhatsApp(payload) {
  try {
    var d = typeof payload === 'string' ? JSON.parse(payload) : (payload || {});
    var phone = normalizeWhatsAppPhone_(d.phone);
    if (phone.length < 8) {
      return { ok: false, message: 'Nomor WhatsApp tujuan tidak valid.' };
    }

    // Simpan PDF di spreadsheet, bukan Google Drive. Ini menghilangkan
    // kebutuhan izin Drive yang menyebabkan error pada deployment lama.
    var receiptResult = saveReceiptPdfToSheet_(d);
    if (!receiptResult || !receiptResult.ok) {
      return receiptResult || { ok: false, message: 'Gagal membuat link PDF.' };
    }

    var properties = PropertiesService.getScriptProperties();
    var accessToken = String(properties.getProperty('WHATSAPP_ACCESS_TOKEN') || '').trim();
    var phoneNumberId = String(properties.getProperty('WHATSAPP_PHONE_NUMBER_ID') || '').trim();
    var apiVersion = String(properties.getProperty('WHATSAPP_API_VERSION') || 'v20.0').trim();
    var receiptUrl = receiptResult.fileUrl;

    // Tanpa Cloud API, frontend menggunakan link wa.me sebagai fallback.
    if (!accessToken || !phoneNumberId) {
      return {
        ok: true,
        sent: false,
        manualSend: true,
        fileUrl: receiptUrl,
        filename: receiptResult.filename,
        message: 'WhatsApp Cloud API belum dikonfigurasi.'
      };
    }

    // Ganti placeholder DI SERVER setelah upload selesai agar Cloud API
    // tidak pernah mengirim teks "{{PDF_LINK}}" ke pelanggan.
    var text = String(d.message || (
      'Halo, berikut link struk transaksi Anda: ' + receiptUrl
    )).replace(/\{\{PDF_LINK\}\}/g, receiptUrl);

    var textResult = callWhatsAppApi_(
      apiVersion,
      phoneNumberId,
      accessToken,
      {
        messaging_product: 'whatsapp',
        to: phone,
        type: 'text',
        text: {
          preview_url: true,
          body: text
        }
      }
    );

    if (!textResult.ok) {
      return {
        ok: false,
        fileUrl: receiptUrl,
        message: 'PDF sudah tersimpan dan link sudah dibuat, tetapi pesan WhatsApp gagal dikirim: ' +
          textResult.message
      };
    }

    return {
      ok: true,
      sent: true,
      fileUrl: receiptUrl,
      messageId: textResult.messageId,
      message: 'Pesan WhatsApp dan link PDF berhasil dikirim otomatis.'
    };
  } catch (err) {
    return {
      ok: false,
      message: 'Gagal mengirim WhatsApp otomatis: ' +
        String(err && err.message ? err.message : err)
    };
  }
}

/**
 * Menyimpan PDF sebagai beberapa bagian di sheet StrukPDF.
 * Batas setiap cell dibuat 45.000 karakter agar aman dari batas cell Sheets.
 */
function saveReceiptPdfToSheet_(payload) {
  try {
    var d = typeof payload === 'string' ? JSON.parse(payload) : (payload || {});
    var base64Data = String(d.base64 || '');
    var filename = String(d.filename || ('struk-' + new Date().getTime() + '.pdf'))
      .replace(/[^a-zA-Z0-9._-]/g, '-');

    if (!base64Data) {
      return { ok: false, message: 'Data PDF tidak valid.' };
    }

    var commaIndex = base64Data.indexOf(',');
    var encoded = commaIndex >= 0 ? base64Data.substring(commaIndex + 1) : base64Data;
    encoded = encoded.replace(/\s/g, '');
    if (!encoded) {
      return { ok: false, message: 'Isi PDF kosong.' };
    }

    var token = Utilities.getUuid().replace(/-/g, '');
    var sheet = getReceiptSheet_();
    var chunks = [];
    for (var i = 0; i < encoded.length; i += RECEIPT_CHUNK_SIZE) {
      chunks.push(encoded.substring(i, i + RECEIPT_CHUNK_SIZE));
    }

    var receiptData = d.receiptData || d.sale || {};
    var metadata = {
      NAMA_PEMBELI: String(receiptData.NAMA_PEMBELI || ''),
      WHATSAPP: String(receiptData.WHATSAPP || ''),
      NO_TRANSAKSI: String(receiptData.NO_TRANSAKSI || receiptData.ID || ''),
      TANGGAL: String(receiptData.TANGGAL || ''),
      PRODUK: String(receiptData.PRODUK || ''),
      HARGA: Number(receiptData.HARGA || 0),
      BAYAR: Number(receiptData.BAYAR || 0),
      METODE: String(receiptData.METODE || ''),
      KEMBALIAN: Number(receiptData.KEMBALIAN || 0)
    };
    var metadataJson = JSON.stringify(metadata);

    // Header baru: TOKEN | FILENAME | CREATED_AT | METADATA | PDF_PART_1 | ...
    // Baris lama tetap dapat dibaca karena renderReceiptPage_ mendeteksi formatnya.
    var minimumColumns = 4 + chunks.length;
    if (sheet.getMaxColumns() < minimumColumns) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), minimumColumns - sheet.getMaxColumns());
    }

    var row = [token, filename, new Date(), metadataJson];
    Array.prototype.push.apply(row, chunks);
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);

    var webAppUrl = ScriptApp.getService().getUrl();
    if (!webAppUrl) {
      return {
        ok: false,
        message: 'URL Web App belum tersedia. Deploy sebagai Web App terlebih dahulu.'
      };
    }

    return {
      ok: true,
      token: token,
      filename: filename,
      fileUrl: webAppUrl + '?receipt=' + encodeURIComponent(token)
    };
  } catch (err) {
    return {
      ok: false,
      message: 'Gagal menyimpan PDF: ' + String(err && err.message ? err.message : err)
    };
  }
}

function getReceiptSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RECEIPT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RECEIPT_SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([['TOKEN', 'FILENAME', 'CREATED_AT']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function renderReceiptPage_(token) {
  var safeToken = String(token || '').replace(/[^a-zA-Z0-9]/g, '');
  if (!safeToken) {
    return HtmlService.createHtmlOutput('<h2>Link struk tidak valid.</h2>');
  }

  try {
    var sheet = getReceiptSheet_();
    var record = null;
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    if (lastRow > 1) {
      var tokenValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
      for (var i = 0; i < tokenValues.length; i++) {
        if (String(tokenValues[i][0]) === safeToken) {
          var recordRow = i + 2;
          record = sheet.getRange(recordRow, 1, 1, lastColumn).getValues()[0];
          break;
        }
      }
    }

    if (!record) {
      return HtmlService.createHtmlOutput(
        '<!doctype html><html lang="id"><meta charset="utf-8">' +
        '<title>Struk tidak ditemukan</title>' +
        '<body style="font-family:Arial;padding:32px;text-align:center">' +
        '<h2>Struk tidak ditemukan</h2><p>Link mungkin sudah tidak valid.</p></body></html>'
      );
    }

    var metadata = {};
    var chunkStart = 3;
    if (record[3] && String(record[3]).charAt(0) === '{') {
      try {
        metadata = JSON.parse(String(record[3]));
        chunkStart = 4;
      } catch (e) {
        metadata = {};
      }
    }

    // Gabungkan semua chunk base64 menjadi satu string
    var encoded = '';
    for (var c = chunkStart; c < record.length; c++) {
      if (record[c]) encoded += String(record[c]);
    }

    var rawFilename = String(record[1] || 'struk.pdf');
    var safeFilename = rawFilename.replace(/['"\\]/g, '');
    var filename = escapeHtml_(rawFilename);

    var total    = formatReceiptMoney_(metadata.HARGA);
    var paid     = formatReceiptMoney_(metadata.BAYAR);
    var change   = formatReceiptMoney_(metadata.KEMBALIAN);
    var dateText = escapeHtml_(formatReceiptDate_(metadata.TANGGAL || record[2]));
    var buyer    = escapeHtml_(metadata.NAMA_PEMBELI || 'Pelanggan Umum');
    var trx      = escapeHtml_(metadata.NO_TRANSAKSI || '-');
    var products = escapeHtml_(metadata.PRODUK || 'Tidak ada detail produk').replace(/\n/g, '<br>');
    var method   = escapeHtml_(metadata.METODE || '-').toUpperCase();

    // Pecah base64 menjadi potongan 4000 karakter agar aman sebagai
    // literal JS di dalam tag <script> (menghindari batas string GAS).
    var JS_CHUNK = 4000;
    var b64Parts = [];
    for (var p = 0; p < encoded.length; p += JS_CHUNK) {
      b64Parts.push(encoded.substring(p, p + JS_CHUNK));
    }
    // Hasilkan: var _PDF_PARTS=["aaa","bbb",...];
    var pdfPartsJs = 'var _PDF_PARTS=' + JSON.stringify(b64Parts) + ';';

    var page =
      '<!doctype html><html lang="id"><head>' +
      '<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<meta name="theme-color" content="#1d4ed8"><title>Struk - ' + filename + '</title>' +
      '<style>' +
      'body{font-family:Arial,sans-serif;background:#eff6ff;margin:0;padding:16px;color:#111}' +
      '.paper{max-width:430px;margin:0 auto;background:#fff;border:2px solid #1d4ed8;' +
      'box-shadow:5px 5px 0 #1d4ed8;border-radius:10px;padding:20px}' +
      '.brand{text-align:center;border-bottom:2px dashed #1d4ed8;padding-bottom:14px;margin-bottom:14px}' +
      '.brand h1{font-size:24px;margin:0;font-weight:800;letter-spacing:1px}' +
      '.brand p{font-size:11px;margin:5px 0 0;color:#52627a}' +
      '.meta{font-size:12px;border-bottom:1px solid #dbeafe;padding-bottom:12px;margin-bottom:14px}' +
      '.line{display:flex;justify-content:space-between;gap:14px;margin:6px 0}' +
      '.label{color:#64748b}.value{text-align:right;font-weight:700}' +
      '.section{font-size:11px;color:#64748b;font-weight:800;text-transform:uppercase;margin-bottom:6px}' +
      '.products{font-size:14px;font-weight:700;line-height:1.45;' +
      'border-bottom:1px dashed #93c5fd;padding-bottom:14px;margin-bottom:14px}' +
      '.total{font-size:16px;font-weight:800}' +
      '.thanks{text-align:center;font-size:12px;color:#64748b;margin:20px 0 14px}' +
      '#btnUnduh{display:block;width:100%;box-sizing:border-box;text-align:center;cursor:pointer;' +
      'background:#1d4ed8;color:#fff;padding:13px;border:none;border-radius:7px;' +
      'font-weight:700;font-size:14px;margin-top:4px}' +
      '#btnUnduh:active{opacity:.85}' +
      '#msg{text-align:center;font-size:12px;color:#64748b;margin-top:8px;min-height:18px}' +
      '</style></head><body>' +
      '<main class="paper">' +
      '<div class="brand"><h1>dwisyafitriproject</h1><p>STRUK PEMBAYARAN KASIR</p></div>' +
      '<div class="meta">' +
      '<div class="line"><span class="label">No. Transaksi</span><span class="value">' + trx + '</span></div>' +
      '<div class="line"><span class="label">Tanggal</span><span class="value">' + dateText + '</span></div>' +
      '<div class="line"><span class="label">Pembeli</span><span class="value">' + buyer + '</span></div>' +
      '<div class="line"><span class="label">Metode</span><span class="value">' + method + '</span></div>' +
      '</div>' +
      '<div class="section">Detail Pesanan</div>' +
      '<div class="products">' + products + '</div>' +
      '<div class="line"><span class="label">Total Tagihan</span><span class="value total">' + total + '</span></div>' +
      '<div class="line"><span class="label">Bayar</span><span class="value">' + paid + '</span></div>' +
      '<div class="line"><span class="label">Kembalian</span><span class="value total">' + change + '</span></div>' +
      '<p class="thanks">Terima kasih telah berbelanja di dwisyafitriproject.</p>' +
      '<button id="btnUnduh" onclick="unduhPdf()">⬇ Unduh PDF Struk</button>' +
      '<p id="msg"></p>' +
      '</main>' +
      '<script>' +
      pdfPartsJs +
      'function unduhPdf(){' +
      '  var btn=document.getElementById("btnUnduh");' +
      '  var msg=document.getElementById("msg");' +
      '  btn.disabled=true;btn.textContent="Menyiapkan PDF...";msg.textContent="";' +
      '  try{' +
      '    var b64=_PDF_PARTS.join("");' +
      '    var bin=atob(b64);' +
      '    var arr=new Uint8Array(bin.length);' +
      '    for(var i=0;i<bin.length;i++){arr[i]=bin.charCodeAt(i);}' +
      '    var blob=new Blob([arr],{type:"application/pdf"});' +
      '    var a=document.createElement("a");' +
      '    a.href=URL.createObjectURL(blob);' +
      '    a.download="' + safeFilename + '";' +
      '    document.body.appendChild(a);a.click();' +
      '    setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(a.href);},1000);' +
      '    btn.textContent="✓ PDF Terunduh";msg.textContent="File tersimpan di folder Unduhan Anda.";' +
      '  }catch(e){' +
      '    btn.disabled=false;btn.textContent="⬇ Unduh PDF Struk";' +
      '    msg.textContent="Gagal: "+e.message;' +
      '  }' +
      '}' +
      '<\/script>' +
      '</body></html>';

    return HtmlService.createHtmlOutput(page).setTitle('Struk Pembayaran');

  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<h2>Gagal membuka struk</h2><p>' +
      escapeHtml_(String(err && err.message ? err.message : err)) +
      '</p>'
    );
  }
}

function formatReceiptMoney_(value) {
  var number = Number(value || 0);
  return 'Rp ' + number.toLocaleString('id-ID');
}

function formatReceiptDate_(value) {
  if (!value) return '-';
  var date = new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

function escapeHtml_(value) {
  return String(value || '').replace(/[&<>"']/g, function (character) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[character];
  });
}

function normalizeWhatsAppPhone_(value) {
  var phone = String(value || '').replace(/\D/g, '');
  if (phone.indexOf('0') === 0) {
    phone = '62' + phone.substring(1);
  }
  return phone;
}

function callWhatsAppApi_(apiVersion, phoneNumberId, accessToken, body) {
  var response = UrlFetchApp.fetch(
    'https://graph.facebook.com/' + apiVersion + '/' + phoneNumberId + '/messages',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + accessToken },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    }
  );

  var responseText = response.getContentText() || '{}';
  var result;
  try {
    result = JSON.parse(responseText);
  } catch (parseError) {
    result = {};
  }

  if (!result.messages || !result.messages.length) {
    return {
      ok: false,
      message: responseText
    };
  }

  return {
    ok: true,
    messageId: result.messages[0].id
  };
}

// ============================================================
// [LAMA - Dipertahankan] Kirim PDF via WhatsApp Cloud API
// ============================================================
/**
 * Fungsi lama yang mengirim PDF langsung sebagai dokumen melalui WhatsApp Cloud API.
 * Dipertahankan untuk kompatibilitas. Tidak lagi digunakan oleh tombol utama.
 *
 * Konfigurasi di Apps Script > Project Settings > Script Properties:
 * - WHATSAPP_ACCESS_TOKEN
 * - WHATSAPP_PHONE_NUMBER_ID
 * - WHATSAPP_API_VERSION (opsional, default v20.0)
 */
function createReceiptPdfAndSendWhatsApp(payload) {
  try {
    var d = typeof payload === 'string' ? JSON.parse(payload) : (payload || {});
    var base64Data = String(d.base64 || '');
    var phone = String(d.phone || '').replace(/\D/g, '');
    var filename = String(d.filename || ('struk-' + new Date().getTime() + '.pdf'))
      .replace(/[^a-zA-Z0-9._-]/g, '-');

    if (phone.indexOf('0') === 0) {
      phone = '62' + phone.substring(1);
    }
    if (phone.length < 8) {
      return { ok: false, message: 'Nomor WhatsApp tujuan tidak valid.' };
    }
    if (!base64Data || base64Data.indexOf(',') === -1) {
      return { ok: false, message: 'Data PDF tidak valid.' };
    }

    var properties = PropertiesService.getScriptProperties();
    var accessToken = properties.getProperty('WHATSAPP_ACCESS_TOKEN');
    var phoneNumberId = properties.getProperty('WHATSAPP_PHONE_NUMBER_ID');

    if (!accessToken || !phoneNumberId) {
      return {
        ok: true,
        sent: false,
        manualAttachment: true,
        message: 'WhatsApp Cloud API belum dikonfigurasi. PDF akan diunduh ke perangkat.',
      };
    }

    var contentType = 'application/pdf';
    var bytes = Utilities.base64Decode(base64Data.substring(base64Data.indexOf(',') + 1));
    var blob = Utilities.newBlob(bytes, contentType, filename);
    var apiVersion = properties.getProperty('WHATSAPP_API_VERSION') || 'v20.0';
    var mediaResponse = UrlFetchApp.fetch(
      'https://graph.facebook.com/' + apiVersion + '/' + phoneNumberId + '/media',
      {
        method: 'post',
        headers: { Authorization: 'Bearer ' + accessToken },
        payload: {
          messaging_product: 'whatsapp',
          type: contentType,
          file: blob,
        },
        muteHttpExceptions: true,
      }
    );
    var mediaResult = JSON.parse(mediaResponse.getContentText() || '{}');
    if (!mediaResult.id) {
      return {
        ok: false,
        message: 'Gagal mengunggah PDF ke WhatsApp: ' + mediaResponse.getContentText(),
      };
    }

    var sendResponse = UrlFetchApp.fetch(
      'https://graph.facebook.com/' + apiVersion + '/' + phoneNumberId + '/messages',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + accessToken },
        payload: JSON.stringify({
          messaging_product: 'whatsapp',
          to: phone,
          type: 'document',
          document: {
            id: mediaResult.id,
            filename: filename,
            caption: String(d.caption || 'Struk transaksi dari dwisyafitriproject'),
          },
        }),
        muteHttpExceptions: true,
      }
    );
    var sendResult = JSON.parse(sendResponse.getContentText() || '{}');
    if (!sendResult.messages || !sendResult.messages.length) {
      return {
        ok: false,
        message: 'Gagal mengirim PDF ke WhatsApp: ' + sendResponse.getContentText(),
      };
    }

    return {
      ok: true,
      sent: true,
      messageId: sendResult.messages[0].id,
    };
  } catch (err) {
    return { ok: false, message: 'Gagal membuat/mengirim PDF: ' + err.toString() };
  }
}
