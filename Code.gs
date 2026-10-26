/**
 * BACKEND GOOGLE APPS SCRIPT
 * Aplikasi: Peminjaman Tools SE
 */

// Konfigurasi Database & Script Terintegrasi
const SPREADSHEET_ID = "1OK7KVPXkcRtNcyvrjT7HB6Xy6fABnWDok0YDTL-6lsg";
const SCRIPT_ID = "1VgM6xzu3g0awATCfE11YhlYTkOWo4BxjD-7nxV-DrL6AQ-u4lvu1uMSn";
const GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwm6sHhmhQWzh5qRCd8kfEpnS0B2UkCYwSS9_kUTPbjyBjJI0DgBvQutmzXYBTdhngiNQ/exec";

function setupDatabase() {
  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  var sheet = ss.getSheetByName("Database");
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName("Database");
  }
  
  var headers = [
    "ID Transaksi", 
    "Timestamp Peminjaman", 
    "Nama Lengkap", 
    "NRP", 
    "No WhatsApp",
    "Perusahaan", 
    "Departemen", 
    "Section", 
    "Tanggal Peminjaman", 
    "Keterangan / Nama Barang", 
    "Foto Barang (Pinjam)", 
    "Status", 
    "Timestamp Pengembalian", 
    "Foto Barang (Kembali)"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#4F46E5");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else {
    // Pastikan jika kolom No WhatsApp belum ada di sheet lama, disisipkan tanpa merusak urutan
    var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var hasWa = false;
    for (var k = 0; k < currentHeaders.length; k++) {
      if (String(currentHeaders[k]).toLowerCase().indexOf("whatsapp") !== -1 || String(currentHeaders[k]).toLowerCase().indexOf("wa") !== -1) {
        hasWa = true;
        break;
      }
    }
    if (!hasWa && sheet.getLastColumn() >= 4) {
      sheet.insertColumnAfter(4);
      sheet.getRange(1, 5).setValue("No WhatsApp");
    }
  }
  
  var folders = DriveApp.getFoldersByName("Aplikasi_Peminjaman_Foto");
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Aplikasi_Peminjaman_Foto");
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  var props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', SPREADSHEET_ID);
  props.setProperty('FOLDER_ID', folder.getId());
  
  Logger.log("Setup Berhasil! Terhubung ke Spreadsheet: " + ss.getUrl());
  return "Setup Berhasil! Terhubung ke Spreadsheet ID: " + SPREADSHEET_ID;
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var result = { success: false, message: "Action tidak dikenal" };
    if (action === "getAllPinjamanAktif") {
      result = getAllPinjamanAktif();
    } else if (action === "getHistoriByNRP") {
      result = getHistoriByNRP(e.parameter.keyword || "");
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Peminjaman Tools SE')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var params = body.params || {};
    var result = { success: false, message: "Action tidak valid" };

    if (action === "submitPeminjaman") {
      result = submitPeminjaman(params.data);
    } else if (action === "getAllPinjamanAktif") {
      result = getAllPinjamanAktif();
    } else if (action === "submitPengembalian") {
      result = submitPengembalian(params.rowNumber, params.base64Foto, params.idTransaksi);
    } else if (action === "getHistoriByNRP") {
      result = getHistoriByNRP(params.keyword);
    } else if (action === "getAllData") {
      result = getAllData(params.password);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "API Error: " + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getActiveSheet() {
  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    var props = PropertiesService.getScriptProperties();
    var ssId = props.getProperty('SPREADSHEET_ID') || SPREADSHEET_ID;
    if (!ssId) throw new Error("Database belum disetup. Spreadsheet ID belum ditemukan.");
    ss = SpreadsheetApp.openById(ssId);
  }
  var sheet = ss.getSheetByName("Database");
  if (!sheet) {
    sheet = ss.getSheets()[0];
    if (sheet) {
      sheet.setName("Database");
    } else {
      sheet = ss.insertSheet("Database");
    }
  }
  return sheet;
}

// Cache folder Drive agar tidak dicari berulang kali
var _cachedFolder = null;
function getUploadFolder() {
  if (_cachedFolder) return _cachedFolder;
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('FOLDER_ID');
  if (folderId) {
    try {
      _cachedFolder = DriveApp.getFolderById(folderId);
      return _cachedFolder;
    } catch (e) {}
  }
  var folders = DriveApp.getFoldersByName("Aplikasi_Peminjaman_Foto");
  _cachedFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Aplikasi_Peminjaman_Foto");
  _cachedFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty('FOLDER_ID', _cachedFolder.getId());
  return _cachedFolder;
}

function saveImageToDrive(base64Data, filename) {
  if (!base64Data) return "";
  var folder = getUploadFolder();
  var splitBase = base64Data.split(',');
  var type = splitBase[0].split(';')[0].replace('data:', '') || 'image/jpeg';
  var byteCharacters = Utilities.base64Decode(splitBase[1]);
  var blob = Utilities.newBlob(byteCharacters, type, filename);
  
  // Folder sudah publik, pembuatan file langsung selesai tanpa menunggu setSharing individual (menghemat 4-6 detik)
  var file = folder.createFile(blob);
  return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
}

function formatSafeDate(val, pattern) {
  if (!val) return "-";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), pattern);
  }
  return String(val);
}

function submitPeminjaman(data) {
  try {
    var sheet = getActiveSheet();
    var idTransaksi = "TRX-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
    var timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var fotoPinjamUrl = saveImageToDrive(data.fotoPinjam, idTransaksi + "_pinjam.jpg");

    var rowData = [
      idTransaksi,
      timestampStr,
      data.nama,
      data.nrp,
      data.noWa || "-",
      data.perusahaan,
      data.departemen,
      data.section,
      data.tglPeminjaman,
      data.keterangan,
      fotoPinjamUrl,
      "DIPINJAM",
      "",
      ""
    ];
    
    sheet.appendRow(rowData);
    return { success: true, message: "Peminjaman berhasil dicatat dengan ID: " + idTransaksi };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}

function getAllPinjamanAktif() {
  try {
    var sheet = getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var results = [];

    // Deteksi letak index kolom secara dinamis
    var idxStatus = 11;
    var idxFoto = 10;
    var idxKeterangan = 9;
    var idxTglPinjam = 8;
    var idxDept = 6;
    var idxWa = 4;
    var idxNrp = 3;

    for (var h = 0; h < headerRow.length; h++) {
      var head = String(headerRow[h]).toLowerCase();
      if (head.indexOf("status") !== -1) idxStatus = h;
      else if (head.indexOf("pinjam") !== -1 && head.indexOf("foto") !== -1) idxFoto = h;
      else if (head.indexOf("keterangan") !== -1) idxKeterangan = h;
      else if (head.indexOf("tanggal peminjaman") !== -1) idxTglPinjam = h;
      else if (head.indexOf("departemen") !== -1) idxDept = h;
      else if (head.indexOf("whatsapp") !== -1 || head.indexOf("wa") !== -1) idxWa = h;
      else if (head.indexOf("nrp") !== -1) idxNrp = h;
    }
    
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      var status = (row[idxStatus] || "").trim().toUpperCase();
      
      if (status === "DIPINJAM") {
        results.push({
          rowNumber: i + 2,
          idTransaksi: row[0] || "",
          nama: row[2] || "",
          nrp: row[idxNrp] || "",
          noWa: row[idxWa] || "-",
          departemen: row[idxDept] || "",
          tglPeminjaman: row[idxTglPinjam] || "-",
          keterangan: row[idxKeterangan] || "",
          fotoUrl: row[idxFoto] || ""
        });
      }
    }
    return { success: true, data: results };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}

function submitPengembalian(rowNumber, base64Foto, idTransaksi) {
  try {
    var sheet = getActiveSheet();
    var r = parseInt(rowNumber, 10);
    var timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var fileName = (idTransaksi || ("TRX_KEMBALI_" + r)) + "_kembali.jpg";
    var fotoKembaliUrl = saveImageToDrive(base64Foto, fileName);

    // Update kolom 12 (Status), 13 (Timestamp Kembali), dan 14 (Foto Kembali) dalam 1 kali eksekusi atomic
    sheet.getRange(r, 12, 1, 3).setValues([["DONE", timestampStr, fotoKembaliUrl]]);
    
    return { success: true, message: "Barang berhasil dikembalikan!" };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}

function getHistoriByNRP(keyword) {
  try {
    var sheet = getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
    var results = [];
    var searchKey = String(keyword || "").trim().toLowerCase();
    
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      var idTrx = String(row[0] || "").trim();
      var nrp = String(row[3] || "").trim();
      var status = String(row[11] || row[10] || "").trim().toUpperCase();
      var isMatch = (nrp.toLowerCase() === searchKey) || (idTrx.toLowerCase().indexOf(searchKey) !== -1);
      
      if (searchKey !== "" && isMatch && status === "DONE") {
        results.push({
          idTransaksi: idTrx,
          nama: String(row[2] || ""),
          nrp: nrp,
          noWa: String(row[4] || "-"),
          departemen: String(row[6] || row[5] || ""),
          section: String(row[7] || row[6] || ""),
          perusahaan: String(row[5] || row[4] || ""),
          keterangan: String(row[9] || row[8] || ""),
          tglPinjam: row[8] || row[7] || row[1] || "-",
          tglKembali: row[12] || row[11] || "-",
          fotoPinjam: String(row[10] || row[9] || ""),
          fotoKembali: String(row[13] || row[12] || ""),
          status: status
        });
      }
    }
    return { success: true, data: results };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}

function getAllData(password) {
  try {
    if (password !== "admin123") {
      return { success: false, message: "Password admin salah!" };
    }
    
    var sheet = getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
    var results = [];
    
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      if (!row[0]) continue;

      results.push({
        idTransaksi: row[0] || "",
        timestampPinjam: row[1] || "",
        nama: row[2] || "",
        nrp: row[3] || "",
        noWa: row[4] || "-",
        perusahaan: row[5] || row[4] || "",
        departemen: row[6] || row[5] || "",
        section: row[7] || row[6] || "",
        tglPinjam: row[8] || row[7] || "",
        keterangan: row[9] || row[8] || "",
        fotoPinjam: row[10] || row[9] || "",
        status: (row[11] || row[10] || "DIPINJAM").trim().toUpperCase(),
        tglKembali: row[12] || row[11] || "-",
        fotoKembali: row[13] || row[12] || ""
      });
    }
    return { success: true, data: results };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}
