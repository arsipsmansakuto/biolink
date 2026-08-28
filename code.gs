/**
 * BIO LINK PRO v7.0 TURBO - ENTERPRISE SUITE 2026
 * Backend Engine: Google Apps Script (Code.gs)
 * Otomatisasi Database Google Sheets & Manajemen Media Google Drive
 */

/**
 * Endpoint Utama GET:
 * Melayani antarmuka Web App atau mengembalikan data JSON jika diminta.
 */
function doGet(e) {
  setupDatabase();

  if (e && e.parameter && e.parameter.action === 'getAppData') {
    var data = getAppData();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: data }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var html = HtmlService.createHtmlOutputFromFile('index');
  html.setTitle('Bio Link Pro Dashboard');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

/**
 * Konfigurasi Otomatis Database (Zero-Config Setup):
 * Membuat Sheet 'Profile' dan 'Links' secara otomatis dengan header lengkap.
 */
function setupDatabase() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log('Lock timeout saat setupDatabase');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Inisialisasi Sheet 'Profile'
    var profileSheet = ss.getSheetByName('Profile');
    if (!profileSheet) {
      profileSheet = ss.insertSheet('Profile');
      profileSheet.appendRow(['Key', 'Value']);
      profileSheet.appendRow(['Name', 'Nama Anda']);
      profileSheet.appendRow(['Bio', 'Digital Creator & Innovator']);
      profileSheet.appendRow(['Avatar', 'data:image/svg+xml,<svg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 24 24\' fill=\'%2394A3B8\'><path d=\'M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 3c1.66 0 3 1.34 3 3s-1.34 3-3 3-3-1.34-3-3 1.34-3 3-3zm0 14.2c-2.5 0-4.71-1.28-6-3.22.03-1.99 4-3.08 6-3.08 1.99 0 5.97 1.09 6 3.08-1.29 1.94-3.5 3.22-6 3.22z\'/></svg>']);
      profileSheet.appendRow(['Background', 'linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%)']);
      profileSheet.appendRow(['Password', 'admin']);
      profileSheet.appendRow(['FontFamily', "'Inter', sans-serif"]);
      profileSheet.appendRow(['FontSize', '16px']);
      profileSheet.appendRow(['TextColor', '#000000']);
      profileSheet.appendRow(['ButtonTextColor', '#ffffff']);
      profileSheet.appendRow(['VerifiedBadge', 'true']);
      
      profileSheet.getRange("A1:B1").setFontWeight("bold");
      profileSheet.setFrozenRows(1);
    }

    // 2. Inisialisasi Sheet 'Links'
    var linksSheet = ss.getSheetByName('Links');
    if (!linksSheet) {
      linksSheet = ss.insertSheet('Links');
      linksSheet.appendRow(['ID', 'Type', 'Title', 'URL', 'Icon', 'Content', 'Color', 'Effect', 'Active', 'Shape', 'Password']);
      linksSheet.getRange("A1:K1").setFontWeight("bold");
      linksSheet.setFrozenRows(1);
    }

  } finally {
    lock.releaseLock();
  }
}

/**
 * Menyimpan data Base64 ke Google Drive secara otomatis
 * Menghindari limitasi sel 50.000 karakter pada Google Sheets
 */
function saveMediaToDriveIfBase64(base64Data, prefix) {
  if (!base64Data || typeof base64Data !== 'string' || !base64Data.startsWith('data:image')) {
    return base64Data;
  }
  try {
    var folderName = "BioLink_Media_Storage";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var parts = base64Data.split(',');
    var header = parts[0];
    var data = parts[1];
    var mime = header.split(':')[1].split(';')[0];
    var ext = mime.split('/')[1] || 'jpg';
    
    var decoded = Utilities.base64Decode(data);
    var blob = Utilities.newBlob(decoded, mime, (prefix || 'media') + '_' + Date.now() + '.' + ext);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch (err) {
    Logger.log("Drive storage error: " + err);
    return base64Data;
  }
}

/**
 * Membuat ID Unik Komponen
 */
function generateId() {
  return 'c_' + Math.random().toString(36).substring(2, 11);
}

/**
 * Mengambil Seluruh Data Profil & Daftar Komponen Terkini
 */
function getAppData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Membaca Konfigurasi Profil
  var profileSheet = ss.getSheetByName('Profile');
  var profileData = profileSheet ? profileSheet.getDataRange().getValues() : [];
  var profile = {};
  for (var i = 1; i < profileData.length; i++) {
    if (profileData[i][0]) {
      profile[profileData[i][0]] = profileData[i][1];
      // Menyediakan alias lowercase untuk kompatibilitas mutlak
      var k = String(profileData[i][0]).toLowerCase();
      profile[k] = profileData[i][1];
    }
  }

  // Membaca Daftar Komponen
  var linksSheet = ss.getSheetByName('Links');
  var linksData = linksSheet ? linksSheet.getDataRange().getValues() : [];
  var links = [];
  for (var j = 1; j < linksData.length; j++) {
    if (linksData[j][0]) {
      links.push({
        id: String(linksData[j][0]),
        type: String(linksData[j][1] || 'link'),
        title: String(linksData[j][2] || ''),
        url: String(linksData[j][3] || ''),
        icon: String(linksData[j][4] || ''),
        content: String(linksData[j][5] || ''),
        color: String(linksData[j][6] || '#4F46E5'),
        effect: String(linksData[j][7] || 'btn-effect-scale'),
        active: String(linksData[j][8]).toUpperCase() !== 'FALSE',
        shape: String(linksData[j][9] || 'rounded-2xl'),
        password: String(linksData[j][10] || '')
      });
    }
  }

  return {
    profile: profile,
    links: links
  };
}

/**
 * Verifikasi Kata Sandi Admin
 */
function verifyAuth(password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var profileSheet = ss.getSheetByName('Profile');
  if (!profileSheet) return false;
  var data = profileSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'Password') {
      return String(data[i][1]) === String(password);
    }
  }
  return false;
}

/**
 * Menyimpan & Memperbarui Pengaturan Profil Permanen
 */
function saveProfile(name, bio, avatar, password, background, fontFamily, fontSize, textColor, buttonTextColor, verifiedBadge) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server sibuk. Silakan coba kembali.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Profile');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Profile');
    }

    var savedAvatar = saveMediaToDriveIfBase64(avatar, 'avatar');
    var savedBg = saveMediaToDriveIfBase64(background, 'bg');

    var data = sheet.getDataRange().getValues();
    var keyRowMap = {};
    for (var i = 1; i < data.length; i++) {
      keyRowMap[data[i][0]] = i + 1;
    }

    function updateOrAdd(key, value) {
      if (value === undefined || value === null) return;
      if (keyRowMap[key]) {
        sheet.getRange(keyRowMap[key], 2).setValue(value);
      } else {
        sheet.appendRow([key, value]);
      }
    }

    updateOrAdd('Name', name);
    updateOrAdd('Bio', bio);
    updateOrAdd('Avatar', savedAvatar);
    updateOrAdd('Background', savedBg);
    updateOrAdd('FontFamily', fontFamily);
    updateOrAdd('FontSize', fontSize);
    updateOrAdd('TextColor', textColor);
    updateOrAdd('ButtonTextColor', buttonTextColor);
    updateOrAdd('VerifiedBadge', verifiedBadge || 'true');

    if (password && String(password).trim() !== '') {
      updateOrAdd('Password', String(password).trim());
    }

    return { 
      success: true, 
      message: 'Profil berhasil diperbarui secara permanen',
      data: {
        Avatar: savedAvatar,
        Background: savedBg
      }
    };

  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menyimpan Komponen Baru atau Mengubah Komponen yang Ada
 */
function saveComponent(id, type, title, url, icon, content, color, effect, active, shape, password) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server sibuk. Silakan coba kembali.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Links');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Links');
    }

    var savedIcon = saveMediaToDriveIfBase64(icon, 'icon');
    var savedUrl = (type === 'image') ? saveMediaToDriveIfBase64(url, 'banner') : (url || '');

    var activeStr = (active === false || active === 'false') ? 'FALSE' : 'TRUE';
    var targetId = id || generateId();
    var data = sheet.getDataRange().getValues();

    // Pembaruan baris data jika ID sudah ada
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(targetId)) {
        sheet.getRange(i + 1, 2, 1, 10).setValues([[
          type || 'link',
          title || '',
          savedUrl,
          savedIcon,
          content || '',
          color || '#4F46E5',
          effect || 'btn-effect-scale',
          activeStr,
          shape || 'rounded-2xl',
          password || ''
        ]]);
        return { success: true, id: targetId, icon: savedIcon, url: savedUrl };
      }
    }

    // Tambah baris baru jika ID belum ada
    sheet.appendRow([
      targetId,
      type || 'link',
      title || '',
      savedUrl,
      savedIcon,
      content || '',
      color || '#4F46E5',
      effect || 'btn-effect-scale',
      activeStr,
      shape || 'rounded-2xl',
      password || ''
    ]);
    return { success: true, id: targetId, icon: savedIcon, url: savedUrl };

  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Memperbarui Urutan Baris Komponen (Drag & Drop) Secara Permanen
 */
function updateComponentOrder(newIds) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server lock timeout.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Links');
    if (!sheet || !newIds || !newIds.length) return { success: false };

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true };

    var rows = data.slice(1);
    var rowMap = {};
    for (var i = 0; i < rows.length; i++) {
      rowMap[String(rows[i][0])] = rows[i];
    }

    var newRows = [];
    for (var j = 0; j < newIds.length; j++) {
      var currentId = String(newIds[j]);
      if (rowMap[currentId]) {
        newRows.push(rowMap[currentId]);
      }
    }

    // Pastikan jika ada row yang belum tercakup di newIds tetap dipertahankan di bawah
    for (var k = 0; k < rows.length; k++) {
      var existId = String(rows[k][0]);
      if (newIds.indexOf(existId) === -1) {
        newRows.push(rows[k]);
      }
    }

    if (newRows.length > 0) {
      var totalRows = sheet.getLastRow();
      if (totalRows > 1) {
        sheet.getRange(2, 1, totalRows - 1, sheet.getLastColumn()).clearContent();
      }
      sheet.getRange(2, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    return { success: true };

  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus Baris Komponen Berdasarkan ID
 */
function deleteLink(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server lock timeout.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Links');
    if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'ID komponen tidak ditemukan' };

  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Endpoint Utama POST:
 * Menerima payload JSON dari panggilan fetch() asinkron frontend.
 */
function doPost(e) {
  try {
    setupDatabase();
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }

    var action = payload.action;
    var result = { success: false, message: 'Aksi tidak valid' };

    if (action === 'getAppData') {
      result = { success: true, data: getAppData() };
    } else if (action === 'verifyAuth') {
      var isAuth = verifyAuth(payload.password);
      result = { success: isAuth, message: isAuth ? 'Autentikasi berhasil' : 'Kata sandi salah' };
    } else if (action === 'saveProfile') {
      result = saveProfile(
        payload.Name || payload.name, 
        payload.Bio || payload.bio, 
        payload.Avatar || payload.avatar, 
        payload.Password || payload.password,
        payload.Background || payload.background, 
        payload.FontFamily || payload.fontFamily, 
        payload.FontSize || payload.fontSize, 
        payload.TextColor || payload.textColor,
        payload.ButtonTextColor || payload.buttonTextColor,
        payload.VerifiedBadge || (payload.verified ? 'true' : 'false')
      );
    } else if (action === 'saveComponent') {
      result = saveComponent(
        payload.id, payload.type, payload.title, payload.url,
        payload.icon, payload.content, payload.color, payload.effect, payload.active,
        payload.shape, payload.password
      );
    } else if (action === 'updateOrder') {
      result = updateComponentOrder(payload.newIds);
    } else if (action === 'deleteLink') {
      result = deleteLink(payload.id);
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
