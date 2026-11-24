// ==========================================
// KONFIGURASI INFRASTRUKTUR CLOUD
// ==========================================
const CONFIG = {
  ROOT_FOLDER_ID: 'MASUKKAN_ID_FOLDER_GOOGLE_DRIVE_ANDA',
  ALLOWED_EXTENSIONS: ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.jpg', '.jpeg', '.png'],
  SHEETS: {
    USERS: 'Users',
    DOCUMENTS: 'Documents',
    LOGS: 'Logs'
  }
};

// ==========================================
// 1. AUTO-PROVISIONING DATABASE
// ==========================================
function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(CONFIG.SHEETS.USERS);
    userSheet.appendRow(['User_ID', 'Username', 'Password_Hash', 'Nama_Lengkap', 'Email', 'Ormawa', 'Role']);
    userSheet.getRange('A1:G1').setFontWeight('bold');
    
    userSheet.appendRow(['USR-001', 'admin', hashPassword('admin123'), 'Administrator Kemahasiswaan', 'admin@example.com', 'KEMAHASISWAAN', 'Admin']);
    userSheet.appendRow(['USR-002', 'user1', hashPassword('user123'), 'Pengurus BEM Fasilkom', 'ormawa@example.com', 'BEM FASILKOM', 'User']);
  }

  let docSheet = ss.getSheetByName(CONFIG.SHEETS.DOCUMENTS);
  if (!docSheet) {
    docSheet = ss.insertSheet(CONFIG.SHEETS.DOCUMENTS);
    docSheet.appendRow([
      'Doc_ID', 'Timestamp', 'User_ID', 'Nama_Pengirim', 'Ormawa', 
      'Jenis_Berkas', 'Nama_File_Asli', 'Nama_File_Cloud', 
      'File_ID_Drive', 'File_URL', 'Status', 'Catatan_Admin', 'Nama_Kegiatan'
    ]);
    docSheet.getRange('A1:M1').setFontWeight('bold');
  } else {
    ensureProkerHeader(docSheet);
  }

  let logSheet = ss.getSheetByName(CONFIG.SHEETS.LOGS);
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.SHEETS.LOGS);
    logSheet.appendRow(['Timestamp', 'Actor', 'Action', 'Detail']);
    logSheet.getRange('A1:D1').setFontWeight('bold');
  }

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  writeLog('SYSTEM', 'PROVISIONING', 'Inisialisasi struktur database selesai.');
}

function ensureProkerHeader(docSheet) {
  if (docSheet.getRange('M1').getValue() === '') {
    docSheet.getRange('M1').setValue('Nama_Kegiatan').setFontWeight('bold');
  }
}

// ==========================================
// 2. UTILITAS KEAMANAN & CLOUD LOGGING
// ==========================================
function hashPassword(password) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, 
    String(password), 
    Utilities.Charset.UTF_8
  );
  return rawHash.map(byte => {
    const v = (byte < 0) ? 256 + byte : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function writeLog(actor, action, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOGS);
  if (logSheet) {
    logSheet.appendRow([new Date(), actor, action, detail]);
  }
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

function sanitizeText(text) {
  return String(text || '').trim().replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '');
}

function validateFileExtension(fileName) {
  const name = String(fileName || '');
  const ext = name.includes('.') ? name.substring(name.lastIndexOf('.')).toLowerCase() : '';
  return {
    valid: CONFIG.ALLOWED_EXTENSIONS.includes(ext),
    ext: ext
  };
}

// ==========================================
// 3. AUTENTIKASI (LOGIN MULTI-ROLE)
// ==========================================
function loginUser(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    const data = userSheet.getDataRange().getValues();
    const inputHash = hashPassword(password);
    const cleanUsername = String(username).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowUsername = String(data[i][1]).trim().toLowerCase();
      const rowHash = String(data[i][2]);

      if (rowUsername === cleanUsername && rowHash === inputHash) {
        const userData = {
          userId: data[i][0],
          username: data[i][1],
          nama: data[i][3],
          email: data[i][4],
          divisi: String(data[i][5]).trim().toUpperCase(),
          role: data[i][6]
        };
        writeLog(userData.username, 'LOGIN', 'Login berhasil sebagai ' + userData.role);
        return { success: true, user: userData };
      }
    }

    writeLog(cleanUsername || 'UNKNOWN', 'LOGIN_FAILED', 'Username atau password salah');
    return { success: false, message: 'Username atau password salah.' };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// ==========================================
// 4. CLOUD FILE ENGINE (UPLOAD, REVISI, HAPUS)
// ==========================================
function uploadDocument(payload) {
  try {
    // Validasi Whitelist Format File di sisi Server
    const fileCheck = validateFileExtension(payload.fileName);
    if (!fileCheck.valid) {
      return { success: false, message: 'Format file tidak diizinkan. Gunakan PDF, Word, Excel, PPT, atau JPG/PNG.' };
    }

    const rootFolder = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    const ormawaNorm = String(payload.divisi).trim().toUpperCase();
    const namaKegiatan = String(payload.namaKegiatan || '').trim();

    if (!namaKegiatan) {
      return { success: false, message: 'Nama Kegiatan / Proker wajib diisi.' };
    }

    const ormawaClean = sanitizeText(ormawaNorm);
    const jenisClean = sanitizeText(payload.jenisBerkas);
    const prokerClean = sanitizeText(namaKegiatan);

    // Auto-Folder: Root -> ORMAWA -> JENIS_BERKAS
    const ormawaFolder = getOrCreateFolder(rootFolder, ormawaClean);
    const targetFolder = getOrCreateFolder(ormawaFolder, jenisClean);

    // Auto-Rename: [ORMAWA]_[JENIS]_[PROKER]_[YYYYMMDD_HHmm].[ext]
    const now = new Date();
    const dateStamp = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMdd_HHmm');
    const cloudFileName = `${ormawaClean}_${jenisClean}_${prokerClean}_${dateStamp}${fileCheck.ext}`;

    const decodedBytes = Utilities.base64Decode(payload.base64Data);
    const blob = Utilities.newBlob(decodedBytes, payload.mimeType, cloudFileName);

    const savedFile = targetFolder.createFile(blob);
    savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const docSheet = ss.getSheetByName(CONFIG.SHEETS.DOCUMENTS);
    ensureProkerHeader(docSheet);

    const docId = 'DOC-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMddHHmmss');

    docSheet.appendRow([
      docId,
      now,
      payload.userId,
      payload.namaPengirim,
      ormawaNorm,
      payload.jenisBerkas,
      payload.fileName,
      cloudFileName,
      savedFile.getId(),
      savedFile.getUrl(),
      'Pending',
      '-',
      namaKegiatan
    ]);

    writeLog(payload.namaPengirim, 'UPLOAD', `Upload ${docId} (${cloudFileName})`);

    return {
      success: true,
      message: 'Dokumen berhasil diunggah dan ditata otomatis di Cloud.',
      docId: docId
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengunggah dokumen: ' + err.message };
  }
}

function reviseDocument(payload) {
  try {
    // Validasi Whitelist WAJIB sebelum menyentuh file lama di Drive
    const fileCheck = validateFileExtension(payload.fileName);
    if (!fileCheck.valid) {
      return { success: false, message: 'Format file tidak diizinkan. Gunakan PDF, Word, Excel, PPT, atau JPG/PNG.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const docSheet = ss.getSheetByName(CONFIG.SHEETS.DOCUMENTS);
    const data = docSheet.getDataRange().getValues();

    let targetRowIndex = -1;
    let oldRow = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.docId)) {
        targetRowIndex = i + 1;
        oldRow = data[i];
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { success: false, message: 'Dokumen tidak ditemukan.' };
    }
    if (String(oldRow[2]) !== String(payload.userId)) {
      return { success: false, message: 'Anda tidak memiliki izin merevisi dokumen ini.' };
    }
    if (String(oldRow[10]) !== 'Revisi') {
      return { success: false, message: 'Hanya dokumen berstatus Revisi yang dapat diperbarui.' };
    }

    // 1. Buang file lama ke Trash Drive (Cloud Storage Lifecycle)
    const oldFileId = oldRow[8];
    if (oldFileId) {
      try {
        DriveApp.getFileById(oldFileId).setTrashed(true);
      } catch (e) {}
    }

    // 2. Simpan file revisi baru di folder yang sama
    const rootFolder = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    const ormawaClean = sanitizeText(oldRow[4]);
    const jenisClean = sanitizeText(oldRow[5]);
    const prokerClean = sanitizeText(oldRow[12] || oldRow[3]);

    const ormawaFolder = getOrCreateFolder(rootFolder, ormawaClean);
    const targetFolder = getOrCreateFolder(ormawaFolder, jenisClean);

    const now = new Date();
    const dateStamp = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMdd_HHmm');
    const cloudFileName = `${ormawaClean}_${jenisClean}_${prokerClean}_REV_${dateStamp}${fileCheck.ext}`;

    const decodedBytes = Utilities.base64Decode(payload.base64Data);
    const blob = Utilities.newBlob(decodedBytes, payload.mimeType, cloudFileName);
    const savedFile = targetFolder.createFile(blob);
    savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 3. Perbarui metadata di Sheets & kembalikan status ke Pending
    docSheet.getRange(targetRowIndex, 2).setValue(now);
    docSheet.getRange(targetRowIndex, 7).setValue(payload.fileName);
    docSheet.getRange(targetRowIndex, 8).setValue(cloudFileName);
    docSheet.getRange(targetRowIndex, 9).setValue(savedFile.getId());
    docSheet.getRange(targetRowIndex, 10).setValue(savedFile.getUrl());
    docSheet.getRange(targetRowIndex, 11).setValue('Pending');
    docSheet.getRange(targetRowIndex, 12).setValue('Revisi telah diunggah, menunggu verifikasi ulang.');

    writeLog(payload.namaPengirim, 'REVISE_UPLOAD', `Revisi ${payload.docId} -> ${cloudFileName}`);

    return {
      success: true,
      message: 'File revisi berhasil diunggah. File lama telah dibersihkan dari Cloud Drive.'
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengunggah revisi: ' + err.message };
  }
}

function deleteDocument(docId, userId, role, actorUsername) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const docSheet = ss.getSheetByName(CONFIG.SHEETS.DOCUMENTS);
    const data = docSheet.getDataRange().getValues();

    let targetRowIndex = -1;
    let targetRow = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(docId)) {
        targetRowIndex = i + 1;
        targetRow = data[i];
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { success: false, message: 'Dokumen tidak ditemukan.' };
    }

    // Validasi hak akses penghapusan
    if (role !== 'Admin') {
      if (String(targetRow[2]) !== String(userId) || String(targetRow[10]) !== 'Pending') {
        return { success: false, message: 'User hanya dapat menghapus dokumen sendiri yang masih Pending.' };
      }
    }

    // Hapus file fisik di Google Drive (Move to Trash)
    const fileId = targetRow[8];
    if (fileId) {
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch (e) {}
    }

    // Hapus metadata di Google Sheets
    docSheet.deleteRow(targetRowIndex);
    writeLog(actorUsername, 'DELETE_DOC', `Menghapus ${docId} (${targetRow[7]}) dari Sheets & Drive`);

    return { success: true, message: `Dokumen ${docId} berhasil dihapus dari Database & Cloud Drive.` };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus dokumen: ' + err.message };
  }
}

// ==========================================
// 5. AMBIL DATA DOKUMEN
// ==========================================
function getDocuments(userId, role) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const docSheet = ss.getSheetByName(CONFIG.SHEETS.DOCUMENTS);
    const data = docSheet.getDataRange().getValues();
    const results = [];

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (role === 'Admin' || String(row[2]) === String(userId)) {
        results.push({
          docId: row[0],
          timestamp: Utilities.formatDate(new Date(row[1]), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm'),
          userId: row[2],
          namaPengirim: row[3],
          divisi: row[4],
          jenisBerkas: row[5],
          namaFileAsli: row[6],
          namaFileCloud: row[7],
          fileId: row[8],
          fileUrl: row[9],
          status: row[10],
          catatanAdmin: row[11],
          namaKegiatan: row[12] || '-'
        });
      }
    }

    return {
      success: true,
      documents: results,
      emailQuota: role === 'Admin' ? MailApp.getRemainingDailyQuota() : null
    };
  } catch (err) {
    return { success: false, message: 'Gagal memuat dokumen: ' + err.message };
  }
}

// ==========================================
// 6. VERIFIKASI ADMIN & AUTO-EMAIL NOTIFIKASI
// ==========================================
function updateDocumentStatus(docId, newStatus, catatan, adminUsername) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const docSheet = ss.getSheetByName(CONFIG.SHEETS.DOCUMENTS);
    const docData = docSheet.getDataRange().getValues();

    let targetRowIndex = -1;
    let docInfo = null;

    for (let i = 1; i < docData.length; i++) {
      if (String(docData[i][0]) === String(docId)) {
        targetRowIndex = i + 1;
        docInfo = {
          userId: docData[i][2],
          namaPengirim: docData[i][3],
          ormawa: docData[i][4],
          namaFileCloud: docData[i][7],
          fileUrl: docData[i][9],
          namaKegiatan: docData[i][12] || '-'
        };
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { success: false, message: 'Dokumen tidak ditemukan.' };
    }

    const cleanCatatan = catatan && String(catatan).trim() !== '' ? String(catatan).trim() : '-';
    docSheet.getRange(targetRowIndex, 11).setValue(newStatus);
    docSheet.getRange(targetRowIndex, 12).setValue(cleanCatatan);

    const userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    const userData = userSheet.getDataRange().getValues();
    let userEmail = '';

    for (let j = 1; j < userData.length; j++) {
      if (String(userData[j][0]) === String(docInfo.userId)) {
        userEmail = userData[j][4];
        break;
      }
    }

    let emailStatus = 'Email tidak ditemukan';
    if (userEmail) {
      const subject = `[Cloud Document Portal] Status Dokumen ${docId}: ${newStatus}`;
      const body = 
        `Halo ${docInfo.namaPengirim} (${docInfo.ormawa}),\n\n` +
        `Dokumen kegiatan "${docInfo.namaKegiatan}" telah diverifikasi oleh Admin.\n\n` +
        `ID Dokumen    : ${docId}\n` +
        `Nama File     : ${docInfo.namaFileCloud}\n` +
        `Status Baru   : ${newStatus}\n` +
        `Catatan Admin : ${cleanCatatan}\n` +
        `Link Dokumen  : ${docInfo.fileUrl}\n\n` +
        (newStatus === 'Revisi' ? `Silakan login kembali ke portal dan gunakan tombol "Upload Revisi" untuk mengunggah berkas perbaikan.\n\n` : '') +
        `Terima kasih,\nSistem Cloud Document Portal`;

      MailApp.sendEmail(userEmail, subject, body);
      emailStatus = 'Email notifikasi terkirim ke ' + userEmail;
    }

    writeLog(adminUsername, 'UPDATE_STATUS', `${docId} -> ${newStatus} (${emailStatus})`);

    return {
      success: true,
      message: `Status diubah menjadi ${newStatus}. ${emailStatus}.`
    };
  } catch (err) {
    return { success: false, message: 'Gagal memperbarui status: ' + err.message };
  }
}

// ==========================================
// 7. WEB APP ENTRY POINT
// ==========================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Cloud Document Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// 8. MANAJEMEN AKUN ORMAWA (KHUSUS ADMIN)
// ==========================================
function createUserByAdmin(payload) {
  try {
    const cleanUsername = String(payload.username || '').trim().toLowerCase();
    const cleanNama = String(payload.nama || '').trim();
    const cleanEmail = String(payload.email || '').trim();
    const cleanOrmawa = String(payload.divisi || '').trim().toUpperCase();
    const password = String(payload.password || '');

    if (!cleanUsername || !cleanNama || !cleanEmail || !cleanOrmawa || password.length < 4) {
      return { success: false, message: 'Semua kolom wajib diisi (password minimal 4 karakter).' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    const data = userSheet.getDataRange().getValues();

    let maxNum = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === cleanUsername) {
        return { success: false, message: `Username "${cleanUsername}" sudah digunakan.` };
      }
      const numPart = parseInt(String(data[i][0]).replace('USR-', ''), 10);
      if (!isNaN(numPart) && numPart > maxNum) maxNum = numPart;
    }

    const newUserId = 'USR-' + String(maxNum + 1).padStart(3, '0');

    userSheet.appendRow([
      newUserId,
      cleanUsername,
      hashPassword(password),
      cleanNama,
      cleanEmail,
      cleanOrmawa,
      payload.role
    ]);

    writeLog(payload.adminUsername, 'CREATE_USER', `Membuat akun ${cleanUsername} (${cleanOrmawa} - ${payload.role})`);

    return {
      success: true,
      message: `Akun "${cleanUsername}" (${cleanOrmawa}) berhasil dibuat.`
    };
  } catch (err) {
    return { success: false, message: 'Gagal membuat akun: ' + err.message };
  }
}

function getAllUsers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    const data = userSheet.getDataRange().getValues();
    const users = [];

    for (let i = 1; i < data.length; i++) {
      users.push({
        userId: data[i][0],
        username: data[i][1],
        nama: data[i][3],
        email: data[i][4],
        ormawa: String(data[i][5]).trim().toUpperCase(),
        role: data[i][6]
      });
    }
    return { success: true, users: users };
  } catch (err) {
    return { success: false, message: 'Gagal memuat daftar akun: ' + err.message };
  }
}

function deleteUserByAdmin(targetUserId, currentAdminUserId, adminUsername) {
  try {
    if (String(targetUserId) === String(currentAdminUserId)) {
      return { success: false, message: 'Anda tidak dapat menghapus akun Admin yang sedang dipakai.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    const data = userSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(targetUserId)) {
        const deletedUsername = data[i][1];
        userSheet.deleteRow(i + 1);
        writeLog(adminUsername, 'DELETE_USER', `Menghapus akun ${deletedUsername} (${targetUserId})`);
        return { success: true, message: `Akun "${deletedUsername}" berhasil dihapus.` };
      }
    }

    return { success: false, message: 'Akun tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus akun: ' + err.message };
  }
}