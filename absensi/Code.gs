/**
 * ============================================================================
 * GOOGLE APPS SCRIPT (GAS) BACKEND - 3 FILES ARCHITECTURE
 * Gerakan Angkatan Muda Kristen Indonesia
 * ============================================================================
 */

var SPREADSHEET_ID = "1zXpTaQuTx8WqoDddkVXUxVJALhvLvBJMisp4Kap89jM";

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch(e) {
      return SpreadsheetApp.getActiveSpreadsheet();
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : null;
  if (action) {
    return handleApiGet(e);
  }
  
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase() : 'index';
  var fileName = (page === 'admin') ? 'admin' : 'index';
  var title = (page === 'admin') ? 'Admin Panel - Absensi GAMKI' : 'Absensi Kegiatan GAMKI';
  
  return HtmlService.createHtmlOutputFromFile(fileName)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function handleApiGet(e) {
  var action = e.parameter.action;
  var response = {};
  try {
    if (action === "getMembers") {
      response = { status: "success", data: getGAMKIMembers() };
    } else if (action === "getHistory") {
      response = { status: "success", data: getAttendanceHistory() };
    } else if (action === "getWebAppUrl") {
      response = { status: "success", data: getWebAppUrl() };
    } else {
      response = { status: "error", message: "Unknown GET action: " + action };
    }
  } catch (error) {
    response = { status: "error", message: error.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var response = {};
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    var data = payload.data;
    
    if (action === "submitAttendance") {
      response = submitAttendance(data);
    } else if (action === "addMember") {
      response = addMemberGAS(data.name);
    } else if (action === "updateMember") {
      response = updateMemberGAS(data.oldName, data.newName);
    } else if (action === "deleteMember") {
      response = deleteMemberGAS(data.name);
    } else if (action === "updateRecord") {
      response = updateAttendanceRecordGAS(data.id, data.memberName, data.eventDate, data.kegiatan);
    } else if (action === "deleteRecord") {
      response = deleteAttendanceRecordGAS(data.id);
    } else if (action === "clearAll") {
      response = clearAllAttendanceGAS();
    } else {
      response = { status: "error", message: "Unknown POST action: " + action };
    }
  } catch (error) {
    response = { status: "error", message: "Server error: " + error.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Kembalikan URL exec Web App ini untuk navigasi antar halaman di client side.
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// 1. Setup Spreadsheet Automatically
function setupSpreadsheet() {
  var ss = getSpreadsheet();
  
  // Tab 1: Presensi
  var sheetPresensi = ss.getSheetByName("Presensi") || ss.getActiveSheet();
  sheetPresensi.setName("Presensi");
  if (sheetPresensi.getLastRow() === 0) {
    var headers = [["ID Absensi", "Waktu Input", "Nama Anggota", "Tanggal Kegiatan", "Kegiatan", "Link Foto Bukti Drive"]];
    sheetPresensi.getRange(1, 1, 1, 6).setValues(headers).setFontWeight("bold").setBackground("#0f2b5c").setFontColor("#ffffff");
    sheetPresensi.setFrozenRows(1);
  }
  
  // Tab 2: Data Anggota
  var sheetAnggota = ss.getSheetByName("Data Anggota");
  if (!sheetAnggota) {
    sheetAnggota = ss.insertSheet("Data Anggota");
    var headersAnggota = [["No", "Nama Lengkap Anggota", "Status"]];
    sheetAnggota.getRange(1, 1, 1, 3).setValues(headersAnggota).setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
    sheetAnggota.setFrozenRows(1);
    
    var defaultMembers = getDefaultMembersArray();
    var rows = defaultMembers.map(function(name, idx) { return [idx + 1, name, "Aktif"]; });
    sheetAnggota.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}

// 2. Member Master CRUD (Sheet: Data Anggota)
function getGAMKIMembers() {
  setupSpreadsheet();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Data Anggota");
  var lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return getDefaultMembersArray();
  
  var values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var members = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().trim() !== "") {
      members.push(values[i][0].toString().trim());
    }
  }
  return members.length > 0 ? members : getDefaultMembersArray();
}

function addMemberGAS(name) {
  setupSpreadsheet();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Data Anggota");
  var lastRow = sheet.getLastRow();
  sheet.appendRow([lastRow, name, "Aktif"]);
  return { status: "success", message: "Anggota '" + name + "' berhasil ditambahkan ke Google Sheets!" };
}

function updateMemberGAS(oldName, newName) {
  setupSpreadsheet();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Data Anggota");
  var values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === oldName) {
      sheet.getRange(i + 2, 2).setValue(newName);
      return { status: "success", message: "Nama anggota berhasil diperbarui di Google Sheets!" };
    }
  }
  return { status: "error", message: "Anggota tidak ditemukan." };
}

function deleteMemberGAS(name) {
  setupSpreadsheet();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Data Anggota");
  var values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === name) {
      sheet.deleteRow(i + 2);
      return { status: "success", message: "Anggota '" + name + "' berhasil dihapus dari Google Sheets." };
    }
  }
  return { status: "error", message: "Anggota tidak ditemukan." };
}

// 3. Attendance CRUD & History (Sheet: Presensi)
function submitAttendance(data) {
  try {
    setupSpreadsheet();
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Presensi");
    
    var timestamp = new Date();
    var formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var idAbsensi = "ATT-" + timestamp.getTime();
    
    var photoUrl = "-";
    if (data.photoBase64) {
      photoUrl = savePhotoToDrive(data.photoBase64, data.memberName, data.eventDate);
    }
    
    sheet.appendRow([idAbsensi, formattedTimestamp, data.memberName, data.eventDate, data.kegiatan, photoUrl]);
    
    return {
      status: "success",
      message: "Presensi berhasil disimpan ke Google Sheets!",
      id: idAbsensi,
      photoUrl: photoUrl
    };
  } catch (error) {
    return { status: "error", message: "Gagal menyimpan absensi: " + error.toString() };
  }
}

function getAttendanceHistory() {
  setupSpreadsheet();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  var lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return [];
  
  var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var records = [];
  
  for (var i = values.length - 1; i >= 0; i--) {
    records.push({
      id: values[i][0],
      timestamp: values[i][1],
      memberName: values[i][2],
      eventDate: values[i][3],
      kegiatan: values[i][4],
      photoUrl: values[i][5]
    });
  }
  return records;
}

function updateAttendanceRecordGAS(id, memberName, eventDate, kegiatan) {
  setupSpreadsheet();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { status: "error", message: "Data tidak ditemukan." };
  
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === id) {
      var row = i + 2;
      sheet.getRange(row, 3).setValue(memberName);
      sheet.getRange(row, 4).setValue(eventDate);
      sheet.getRange(row, 5).setValue(kegiatan);
      return { status: "success", message: "Data absensi berhasil diperbarui di Google Sheets!" };
    }
  }
  return { status: "error", message: "ID Presensi tidak ditemukan." };
}

function deleteAttendanceRecordGAS(id) {
  setupSpreadsheet();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { status: "error", message: "Data tidak ditemukan." };
  
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === id) {
      sheet.deleteRow(i + 2);
      return { status: "success", message: "Data absensi berhasil dihapus dari Google Sheets." };
    }
  }
  return { status: "error", message: "ID Presensi tidak ditemukan." };
}

function clearAllAttendanceGAS() {
  setupSpreadsheet();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Presensi");
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return { status: "success", message: "Seluruh data absensi di Google Sheets berhasil dibersihkan!" };
}

// 4. Save Photo to Drive
function savePhotoToDrive(base64Data, memberName, eventDate) {
  try {
    var folderName = "Foto Bukti Absensi GAMKI";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    var parts = base64Data.split(",");
    var contentType = parts[0].split(";")[0].replace("data:", "");
    var decodedData = Utilities.base64Decode(parts[1]);
    
    var sanitizedName = memberName.replace(/[^a-zA-Z0-9 ]/g, "").replace(/\s+/g, "_");
    var fileName = "Absensi_" + sanitizedName + "_" + eventDate + ".jpg";
    
    var blob = Utilities.newBlob(decodedData, contentType, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (e) {
    return "Gagal Upload Foto: " + e.toString();
  }
}

function getDefaultMembersArray() {
  return [
    "BOYDO HARRIS K. PANJAITAN, S.H.", "Boy MF Tampubolon, S.E.", "John Martin Lumbangaol, S.Sos",
    "Gogo Satria Marbun, S.S.", "Tumpal Utrecht Napitupulu, S.H., M.H", "Herbert H Panjaitan, S.T., M.Si",
    "Gerald Siahaan, SE, SH, MM, MH", "Tulus Josep H Marpaung, S.Si., M.Si", "Leon Franciscus Lumban Batu, S.E., M.Si.",
    "Thomas F Hutahaean, SE., M.Si., M.Ak", "Elisabet Grace Damanik, S.Pd", "Hendrik Fernandes Naipospos, S.I.Kom., M.I.Kom",
    "Charli Sihombing, SH", "Metro Halomoan Hutabarat, S.T.", "Lied Apriani Pane, S.Kep.,Ners", "Pdt. Krisman Saragih, S.Th",
    "ARION PASARIBU, S.E.", "Febiola Panjaitan", "Marudut Simanjuntak, S.E.", "Jaya Berkat Perlindungan Sinaga",
    "Richard S.D. Hutapea, S.H.", "Belman Hasibuan, SE, S.H., M.H", "Ir. Antonius Simangunsong, S.T.",
    "Benni Simbolon, S.E.", "Asta Saragih", "Khant Asido Hutagaol", "Roy Ferdinan Tanjung, S.E.",
    "Luhut Pardamean Purba, SM. MPA", "Leo Sagala, S.Kom", "Immanuel Gerard Harianja", "Yohanes Abadi Simatupang",
    "Mellinda Manurung, S.Tr.Kep.,Ns", "Gomgom Nababan", "JHOLANT B AMELIA SINAGA, S.E., M.M. M.Ak",
    "Euodia Graceia Saragi", "Atri Sri R Zebua, SKM", "Marthin J Sinambela", "Sudarmo Elyanto Simangunsong",
    "Abed Nego Lumbangaol, S.Pd", "Yosua Tahyudi R. Panjaitan, S.H., M.H.", "Echo Parlaungan Siahaan, S.",
    "Ronald Y Simanjuntak, S.I.Kom", "Ricki A Hutabarat, S.Pd", "Christoffel Hasiholan Sihotang, S.H, M.H.",
    "Boston Erwin Simbolon, S.Pd", "Polomawaty Sinaga, S.E", "Fransisco Lumban Batu, S.H., M.H.",
    "Jonathan Liviera Marpaung, S.Si., M.Si", "Arianis Laia, S.Keb, Bdn", "Rido Adeward Sitompul, S.H.",
    "Ernes Sariah S Munthe", "Hotlan Listen Novianus Sitompul, S.E", "Hotbona Novandi Tambunan S.S., M.M.",
    "Abel Pramusti Sinaga, Amd.Kep", "Firman Sinaga"
  ];
}
