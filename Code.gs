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
      result = submitPengembalian(params.rowNumber, params.base64Foto);
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

function saveImageToDrive(base64Data, filename) {
  if (!base64Data) return "";
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('FOLDER_ID');
  var folder;
  if (folderId) {
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {}
  }
  if (!folder) {
    var folders = DriveApp.getFoldersByName("Aplikasi_Peminjaman_Foto");
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Aplikasi_Peminjaman_Foto");
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    props.setProperty('FOLDER_ID', folder.getId());
  }
  
  var splitBase = base64Data.split(',');
  var type = splitBase[0].split(';')[0].replace('data:', '');
  var byteCharacters = Utilities.base64Decode(splitBase[1]);
  var blob = Utilities.newBlob(byteCharacters, type, filename);
  
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
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
    var timestamp = new Date();
    var fotoPinjamUrl = saveImageToDrive(data.fotoPinjam, idTransaksi + "_pinjam.jpg");
    
    var rowData = [
      idTransaksi,
      timestamp,
      data.nama,
      data.nrp,
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
    
    var data = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
    var results = [];
    
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      var status = (row[10] || "").trim().toUpperCase();
      
      if (status === "DIPINJAM") {
        results.push({
          rowNumber: i + 2,
          idTransaksi: row[0] || "",
          nama: row[2] || "",
          nrp: row[3] || "",
          departemen: row[5] || "",
          tglPeminjaman: row[7] || "-",
          keterangan: row[8] || "",
          fotoUrl: row[9] || ""
        });
      }
    }
    return { success: true, data: results };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}

function submitPengembalian(rowNumber, base64Foto) {
  try {
    var sheet = getActiveSheet();
    var timestamp = new Date();
    var idTransaksi = sheet.getRange(rowNumber, 1).getValue();
    var fotoKembaliUrl = saveImageToDrive(base64Foto, idTransaksi + "_kembali.jpg");
    
    sheet.getRange(rowNumber, 11).setValue("DONE");
    sheet.getRange(rowNumber, 12).setValue(timestamp);
    sheet.getRange(rowNumber, 13).setValue(fotoKembaliUrl);
    
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

    // Menggunakan getDisplayValues() instan tanpa loop formatSafeDate
    var data = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var results = [];
    var searchKey = String(keyword || "").trim().toLowerCase();
    
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      var idTrx = String(row[0] || "").trim();
      var nrp = String(row[3] || "").trim();
      var status = String(row[10] || "").trim().toUpperCase();
      var isMatch = (nrp.toLowerCase() === searchKey) || (idTrx.toLowerCase().indexOf(searchKey) !== -1);
      
      if (searchKey !== "" && isMatch && status === "DONE") {
        results.push({
          idTransaksi: idTrx,
          nama: String(row[2] || ""),
          nrp: nrp,
          departemen: String(row[5] || ""),
          section: String(row[6] || ""),
          perusahaan: String(row[4] || ""),
          keterangan: String(row[8] || ""),
          tglPinjam: row[7] || row[1] || "-",
          tglKembali: row[11] || "-",
          fotoPinjam: String(row[9] || ""),
          fotoKembali: String(row[12] || ""),
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
    
    // Optimasi super cepat: Membaca teks terformat langsung dari spreadsheet dalam 1 batch
    var data = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    var results = [];
    
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      if (!row[0]) continue;

      results.push({
        idTransaksi: row[0] || "",
        timestampPinjam: row[1] || "",
        nama: row[2] || "",
        nrp: row[3] || "",
        perusahaan: row[4] || "",
        departemen: row[5] || "",
        section: row[6] || "",
        tglPinjam: row[7] || "",
        keterangan: row[8] || "",
        fotoPinjam: row[9] || "",
        status: row[10] || "DIPINJAM",
        tglKembali: row[11] || "-",
        fotoKembali: row[12] || ""
      });
    }
    return { success: true, data: results };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}
