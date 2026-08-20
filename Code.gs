/**
 * ============================================================================
 * GOOGLE APPS SCRIPT (GAS) BACKEND
 * PORTAL PENDAFTARAN PELATIHAN UMKM GAMKI SUMATERA UTARA 2026
 * ============================================================================
 */

// 1. ENTRY POINT GET & POST

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
  var response = {};
  
  try {
    setupSpreadsheet();
    if (action === "getPendaftaran") {
      response = { success: true, data: getPendaftaranData() };
    } else if (action === "getPeserta") {
      response = { success: true, data: getPesertaData() };
    } else if (action === "getRegistrasi") {
      response = { success: true, data: getRegistrasiData() };
    } else {
      response = { success: false, message: "Action GET tidak dikenal: " + action };
    }
  } catch (err) {
    response = { success: false, message: "Error: " + err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var response = {};
  try {
    setupSpreadsheet();
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    
    if (action === "daftar") {
      response = handleDaftar(payload);
    } else if (action === "login") {
      response = handleLogin(payload);
    } else if (action === "terima") {
      response = handleTerima(payload);
    } else if (action === "tolak") {
      response = handleTolak(payload);
    } else if (action === "registrasi") {
      response = handleRegistrasiHariH(payload);
    } else {
      response = { success: false, message: "Action POST tidak dikenal: " + action };
    }
  } catch (err) {
    response = { success: false, message: "Error Server: " + err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}


// 2. SETUP SPREADSHEET & TAB OTOMATIS

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Tab 1: Pendaftaran
  var sheetPendaftaran = ss.getSheetByName("Pendaftaran") || ss.getActiveSheet();
  sheetPendaftaran.setName("Pendaftaran");
  if (sheetPendaftaran.getLastRow() === 0) {
    var headersPend = [[
      "ID Pendaftaran", "Timestamp", "Nama UMKM", "Bidang Usaha", "Nama Produk",
      "Legalitas Usaha", "Alamat", "Nama Pemilik", "Jabatan", "No HP", "Email",
      "Status NIB", "Nomor NIB", "Status NPWP", "Nomor NPWP", "Status PIRT", "Nomor PIRT",
      "Status Merk", "Nomor Merk", "Status Halal", "Nomor Halal", "Kegiatan Ekspor",
      "URL Foto Produk", "Status Verifikasi"
    ]];
    sheetPendaftaran.getRange(1, 1, 1, headersPend[0].length)
      .setValues(headersPend).setFontWeight("bold").setBackground("#0f2b5c").setFontColor("#ffffff");
    sheetPendaftaran.setFrozenRows(1);
  }
  
  // Tab 2: Peserta
  var sheetPeserta = ss.getSheetByName("Peserta");
  if (!sheetPeserta) {
    sheetPeserta = ss.insertSheet("Peserta");
    var headersPeserta = [["ID Peserta", "ID Pendaftaran", "Nama UMKM", "Nama Pemilik", "No HP", "Email", "Status Registrasi", "Tanggal ACC"]];
    sheetPeserta.getRange(1, 1, 1, headersPeserta[0].length)
      .setValues(headersPeserta).setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
    sheetPeserta.setFrozenRows(1);
  }
  
  // Tab 3: Registrasi (Hari H)
  var sheetRegistrasi = ss.getSheetByName("Registrasi");
  if (!sheetRegistrasi) {
    sheetRegistrasi = ss.insertSheet("Registrasi");
    var headersReg = [["ID Registrasi", "Waktu Registrasi", "ID Peserta", "Nama UMKM", "Nama Pemilik", "Petugas Admin"]];
    sheetRegistrasi.getRange(1, 1, 1, headersReg[0].length)
      .setValues(headersReg).setFontWeight("bold").setBackground("#10b981").setFontColor("#ffffff");
    sheetRegistrasi.setFrozenRows(1);
  }
  
  // Tab 4: Admin
  var sheetAdmin = ss.getSheetByName("Admin");
  if (!sheetAdmin) {
    sheetAdmin = ss.insertSheet("Admin");
    var headersAdmin = [["Username", "Password", "Nama Admin"]];
    sheetAdmin.getRange(1, 1, 1, 3).setValues(headersAdmin).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    sheetAdmin.appendRow(["admin", "gamki2026", "Panitia GAMKI SUMUT"]);
    sheetAdmin.setFrozenRows(1);
  }
}


// 3. HANDLER PENDAFTARAN (PUBLIK)

function formatTextNumber(val) {
  if (!val || val.toString().trim() === "" || val === "-") return "-";
  var str = val.toString().trim();
  if (str.startsWith("8")) {
    str = "0" + str;
  }
  return "'" + str;
}

function handleDaftar(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pendaftaran");
  
  var now = new Date();
  var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var idPendaftaran = "REG-2026-" + Math.floor(1000 + Math.random() * 9000);
  
  // Simpan foto ke Google Drive
  var fotoUrl = "-";
  if (data.foto_produk_base64) {
    fotoUrl = savePhotoToDrive(data.foto_produk_base64, data.nama_umkm);
  }
  
  sheet.appendRow([
    idPendaftaran,
    timestamp,
    data.nama_umkm || "-",
    data.bidang_usaha || "-",
    data.nama_produk || "-",
    data.legalitas || "-",
    data.alamat || "-",
    data.nama_pemilik || "-",
    data.jabatan || "-",
    formatTextNumber(data.no_hp),
    data.email || "-",
    data.status_nib || "Belum Memiliki",
    formatTextNumber(data.nomor_nib),
    data.status_npwp || "Belum Memiliki",
    formatTextNumber(data.nomor_npwp),
    data.status_pirt || "Belum Memiliki",
    formatTextNumber(data.nomor_pirt),
    data.status_merk || "Belum Memiliki",
    formatTextNumber(data.nomor_merk),
    data.status_halal || "Belum Memiliki",
    formatTextNumber(data.nomor_halal),
    data.kegiatan_export || "Belum Pernah",
    fotoUrl,
    "Menunggu Verifikasi"
  ]);
  
  // Kirim Email Konfirmasi Pendaftaran Diterima Sistem (Otomatis)
  sendEmailKonfirmasiPendaftaran(data.email, data.nama_pemilik, data.nama_umkm, idPendaftaran);
  
  return {
    success: true,
    message: "Pendaftaran berhasil dikirim! Email konfirmasi telah dikirim ke " + (data.email || "email Anda"),
    data: { id_pendaftaran: idPendaftaran }
  };
}

function savePhotoToDrive(base64Data, namaUmkm) {
  try {
    var folderName = "Foto Produk UMKM GAMKI 2026";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    
    var splitData = base64Data.split(",");
    var contentType = splitData[0].split(";")[0].replace("data:", "");
    var bytes = Utilities.base64Decode(splitData[1]);
    var fileName = "Foto_" + namaUmkm.replace(/[^a-zA-Z0-9]/g, "_") + "_" + new Date().getTime() + ".jpg";
    var blob = Utilities.newBlob(bytes, contentType, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w800";
  } catch (e) {
    return "Error Upload: " + e.toString();
  }
}


// 4. HANDLER LOGIN ADMIN

function handleLogin(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Admin");
  var rows = sheet.getDataRange().getValues();
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === data.username && rows[i][1].toString() === data.password) {
      return {
        success: true,
        message: "Login berhasil",
        data: { username: rows[i][0], name: rows[i][2] }
      };
    }
  }
  return { success: false, message: "Username atau Password salah!" };
}


// 5. READ DATA UNTUK DASHBOARD

function getPendaftaranData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pendaftaran");
  var rows = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < rows.length; i++) {
    result.push({
      id_pendaftaran: rows[i][0],
      timestamp: rows[i][1],
      nama_umkm: rows[i][2],
      bidang_usaha: rows[i][3],
      nama_produk: rows[i][4],
      legalitas: rows[i][5],
      alamat: rows[i][6],
      nama_pemilik: rows[i][7],
      jabatan: rows[i][8],
      no_hp: rows[i][9],
      email: rows[i][10],
      status_nib: rows[i][11],
      nomor_nib: rows[i][12],
      status_npwp: rows[i][13],
      nomor_npwp: rows[i][14],
      status_pirt: rows[i][15],
      nomor_pirt: rows[i][16],
      status_merk: rows[i][17],
      nomor_merk: rows[i][18],
      status_halal: rows[i][19],
      nomor_halal: rows[i][20],
      kegiatan_export: rows[i][21],
      foto_produk_url: rows[i][22],
      status: rows[i][23]
    });
  }
  return result;
}

function getPesertaData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Peserta");
  var rows = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < rows.length; i++) {
    result.push({
      id_peserta: rows[i][0],
      id_pendaftaran: rows[i][1],
      nama_umkm: rows[i][2],
      nama_pemilik: rows[i][3],
      no_hp: rows[i][4],
      email: rows[i][5],
      status_registrasi: rows[i][6],
      tanggal_acc: rows[i][7]
    });
  }
  return result;
}

function getRegistrasiData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Registrasi");
  var rows = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < rows.length; i++) {
    result.push({
      id_registrasi: rows[i][0],
      timestamp_registrasi: rows[i][1],
      id_peserta: rows[i][2],
      nama_umkm: rows[i][3],
      nama_pemilik: rows[i][4],
      admin_user: rows[i][5]
    });
  }
  return result;
}


// 6. VERIFIKASI ACC & EMAIL NOTIFIKASI DITERIMA

function handleTerima(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPend = ss.getSheetByName("Pendaftaran");
  var rowsPend = sheetPend.getDataRange().getValues();
  
  var targetRow = -1;
  var pendaftaranObj = null;
  
  for (var i = 1; i < rowsPend.length; i++) {
    if (rowsPend[i][0] === data.id_pendaftaran) {
      targetRow = i + 1;
      pendaftaranObj = rowsPend[i];
      break;
    }
  }
  
  if (targetRow === -1) {
    return { success: false, message: "ID Pendaftaran tidak ditemukan." };
  }
  
  // Update status di Sheet Pendaftaran
  sheetPend.getRange(targetRow, 24).setValue("Diterima");
  
  // Buat ID Peserta Unik
  var idPeserta = "PST-2026-" + Math.floor(1000 + Math.random() * 9000);
  var sheetPes = ss.getSheetByName("Peserta");
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  
  sheetPes.appendRow([
    idPeserta,
    pendaftaranObj[0],
    pendaftaranObj[2], // nama_umkm
    pendaftaranObj[7], // nama_pemilik
    formatTextNumber(pendaftaranObj[9]), // no_hp
    pendaftaranObj[10],// email
    "Belum Registrasi",
    dateStr
  ]);
  
  // Kirim Email DITERIMA
  sendEmailDiterima(pendaftaranObj[10], pendaftaranObj[7], pendaftaranObj[2], pendaftaranObj[0], idPeserta);
  
  return { success: true, message: "Pendaftaran DITERIMA! ID Peserta: " + idPeserta + " & Email Notifikasi telah terkirim." };
}


// 7. VERIFIKASI TOLAK & EMAIL NOTIFIKASI DITOLAK

function handleTolak(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPend = ss.getSheetByName("Pendaftaran");
  var rowsPend = sheetPend.getDataRange().getValues();
  
  var targetRow = -1;
  var pendaftaranObj = null;
  
  for (var i = 1; i < rowsPend.length; i++) {
    if (rowsPend[i][0] === data.id_pendaftaran) {
      targetRow = i + 1;
      pendaftaranObj = rowsPend[i];
      break;
    }
  }
  
  if (targetRow === -1) {
    return { success: false, message: "ID Pendaftaran tidak ditemukan." };
  }
  
  sheetPend.getRange(targetRow, 24).setValue("Ditolak");
  
  // Kirim Email DITOLAK
  sendEmailDitolak(pendaftaranObj[10], pendaftaranObj[7], pendaftaranObj[2], pendaftaranObj[0]);
  
  return { success: true, message: "Pendaftaran telah DITOLAK & Email Notifikasi telah terkirim." };
}


// 8. REGISTRASI HARI H (CHECK-IN LOCATION)

function handleRegistrasiHariH(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPes = ss.getSheetByName("Peserta");
  var rowsPes = sheetPes.getDataRange().getValues();
  
  var targetPeserta = null;
  var targetRow = -1;
  
  for (var i = 1; i < rowsPes.length; i++) {
    if (rowsPes[i][0] === data.id_peserta) {
      targetRow = i + 1;
      targetPeserta = rowsPes[i];
      break;
    }
  }
  
  if (!targetPeserta) {
    return { success: false, message: "ID Peserta '" + data.id_peserta + "' tidak ditemukan atau belum disetujui (ACC)." };
  }
  
  if (targetPeserta[6] === "Sudah Registrasi") {
    return { success: false, message: "Peserta ini SUDAH melakukan registrasi kehadiran sebelumnya!" };
  }
  
  // Update status peserta
  sheetPes.getRange(targetRow, 7).setValue("Sudah Registrasi");
  
  // Catat di Sheet Registrasi
  var sheetReg = ss.getSheetByName("Registrasi");
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var idReg = "REG-HADIR-" + new Date().getTime();
  
  sheetReg.appendRow([
    idReg,
    nowStr,
    targetPeserta[0], // id_peserta
    targetPeserta[2], // nama_umkm
    targetPeserta[3], // nama_pemilik
    data.admin_user || "Admin"
  ]);
  
  return { success: true, message: "Kehadiran atas nama " + targetPeserta[3] + " (" + targetPeserta[2] + ") berhasil dicatat!" };
}


// 9. FUNGSI KIRIM EMAIL NOTIFIKASI HTML

function sendEmailDiterima(email, namaPemilik, namaUmkm, idPendaftaran, idPeserta) {
  if (!email || email.indexOf("@") === -1) return;
  
  var subject = "[GAMKI SUMUT 2026] Selamat! Pendaftaran Pelatihan UMKM Anda DITERIMA";
  var htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #0f2b5c; color: #ffffff; padding: 20px; text-align: center;">
        <h2 style="margin: 0; font-size: 20px; font-weight: bold;">DPD GAMKI SUMATERA UTARA</h2>
        <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.9;">Pelatihan & Pameran UMKM 2026</p>
      </div>
      <div style="padding: 25px; color: #1e293b; line-height: 1.6;">
        <h3 style="color: #0f2b5c; margin-top:0;">Halo, ${namaPemilik}!</h3>
        <p>Selamat! Berkas pendaftaran untuk usaha <strong>${namaUmkm}</strong> telah diverifikasi dan <span style="color:#10b981; font-weight:bold;">DITERIMA</span> oleh Panitia Pelatihan UMKM GAMKI SUMUT 2026.</p>
        
        <div style="background-color: #f1f5f9; border-left: 4px solid #d97706; padding: 15px; margin: 20px 0; border-radius: 4px;">
          <p style="margin: 0 0 5px 0;"><strong>ID Pendaftaran:</strong> ${idPendaftaran}</p>
          <p style="margin: 0;"><strong>ID Peserta Resmi:</strong> <span style="font-size: 18px; color: #0f2b5c; font-weight: bold;">${idPeserta}</span></p>
        </div>
        
        <p><strong>Petunjuk Penting:</strong></p>
        <ul style="padding-left: 20px;">
          <li>Simpan Kode <strong>ID Peserta Resmi (${idPeserta})</strong> ini untuk dipakai saat Registrasi Kehadiran di lokasi acara pada Hari H.</li>
          <li>Panitia akan segera mengontak WhatsApp Anda untuk pembagian grup dan jadwal detail kegiatan.</li>
        </ul>
        
        <p style="margin-top: 30px;">Salam hangat,<br><strong>Panitia Pelatihan UMKM GAMKI SUMUT 2026</strong></p>
      </div>
      <div style="background-color: #f8fafc; border-top: 1px solid #e2e8f0; padding: 12px; text-align: center; font-size: 12px; color: #64748b;">
        Email ini dikirimkan secara otomatis oleh Sistem Portal GAMKI SUMUT 2026.
      </div>
    </div>
  `;
  
  try {
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
  } catch(e) {
    Logger.log("Gagal kirim email: " + e.toString());
  }
}

function sendEmailDitolak(email, namaPemilik, namaUmkm, idPendaftaran) {
  if (!email || email.indexOf("@") === -1) return;
  
  var subject = "[GAMKI SUMUT 2026] Informasi Pendaftaran Pelatihan UMKM";
  var htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #0f2b5c; color: #ffffff; padding: 20px; text-align: center;">
        <h2 style="margin: 0; font-size: 20px; font-weight: bold;">DPD GAMKI SUMATERA UTARA</h2>
        <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.9;">Pelatihan & Pameran UMKM 2026</p>
      </div>
      <div style="padding: 25px; color: #1e293b; line-height: 1.6;">
        <h3 style="color: #0f2b5c; margin-top:0;">Halo, ${namaPemilik}!</h3>
        <p>Terima kasih atas antusiasme Anda mendaftarkan usaha <strong>${namaUmkm}</strong> (ID: <code>${idPendaftaran}</code>) pada kegiatan Pelatihan UMKM GAMKI SUMUT 2026.</p>
        
        <p>Setelah dilakukan proses peninjauan berkas oleh panitia, kami menginformasikan bahwa pendaftaran Anda <strong>belum dapat disetujui</strong> pada gelombang kali ini karena keterbatasan kuota.</p>
        
        <p>Tetap semangat dan nantikan program pembinaan UMKM GAMKI selanjutnya!</p>
        
        <p style="margin-top: 30px;">Salam hangat,<br><strong>Panitia Pelatihan UMKM GAMKI SUMUT 2026</strong></p>
      </div>
      <div style="background-color: #f8fafc; border-top: 1px solid #e2e8f0; padding: 12px; text-align: center; font-size: 12px; color: #64748b;">
        Email ini dikirimkan secara otomatis oleh Sistem Portal GAMKI SUMUT 2026.
      </div>
    </div>
  `;
  
  try {
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
  } catch(e) {
    Logger.log("Gagal kirim email: " + e.toString());
  }
}

function sendEmailKonfirmasiPendaftaran(email, namaPemilik, namaUmkm, idPendaftaran) {
  if (!email || email.indexOf("@") === -1) return;
  
  var subject = "[GAMKI SUMUT 2026] Konfirmasi Pendaftaran Pelatihan UMKM (" + idPendaftaran + ")";
  var htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #0f2b5c; color: #ffffff; padding: 20px; text-align: center;">
        <h2 style="margin: 0; font-size: 20px; font-weight: bold;">DPD GAMKI SUMATERA UTARA</h2>
        <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.9;">Pelatihan & Pameran UMKM 2026</p>
      </div>
      <div style="padding: 25px; color: #1e293b; line-height: 1.6;">
        <h3 style="color: #0f2b5c; margin-top:0;">Halo, ${namaPemilik}!</h3>
        <p>Terima kasih telah mendaftarkan usaha <strong>${namaUmkm}</strong> pada kegiatan Pelatihan UMKM GAMKI SUMUT 2026.</p>
        
        <div style="background-color: #f1f5f9; border-left: 4px solid #0f2b5c; padding: 15px; margin: 20px 0; border-radius: 4px;">
          <p style="margin: 0 0 5px 0;"><strong>ID Pendaftaran Anda:</strong> <span style="font-size: 16px; color: #0f2b5c; font-weight: bold;">${idPendaftaran}</span></p>
          <p style="margin: 0; font-size: 13px; color: #64748b;">Status Berkas: <span style="color: #d97706; font-weight: bold;">Menunggu Verifikasi Panitia</span></p>
        </div>
        
        <p>Berkas pendaftaran dan foto produk Anda telah berhasil diterima oleh sistem database kami. Panitia akan segera melakukan verifikasi kelengkapan berkas Anda.</p>
        <p>Setelah proses verifikasi selesai, Anda akan menerima email notifikasi kelulusan resmi beserta ID Peserta.</p>
        
        <p style="margin-top: 30px;">Salam hangat,<br><strong>Panitia Pelatihan UMKM GAMKI SUMUT 2026</strong></p>
      </div>
      <div style="background-color: #f8fafc; border-top: 1px solid #e2e8f0; padding: 12px; text-align: center; font-size: 12px; color: #64748b;">
        Email ini dikirimkan secara otomatis oleh Sistem Portal GAMKI SUMUT 2026.
      </div>
    </div>
  `;
  
  try {
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: htmlBody });
  } catch(e) {
    Logger.log("Gagal kirim email konfirmasi: " + e.toString());
  }
}
