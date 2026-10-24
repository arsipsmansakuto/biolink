/**
 * ============================================================================
 * KSL BIO LINK - MULTI-PROJECT MANAGER (BACKEND GOOGLE APPS SCRIPT)
 * Versi: 5.1 (Full Module, Appearance Enhancements & Multi-Bio Support)
 * ============================================================================
 * Struktur Sheet yang otomatis dibuat:
 *  1. SETTINGS     : Konfigurasi global & deployment metadata
 *  2. USERS        : Akun administrator & editor
 *  3. PROJECTS     : Multi Bio profil (slug, nama, bio, avatar, tema, font, button style)
 *  4. LINKS        : Daftar tautan publik & modul smart campaign (URL kondisional)
 *  5. RECYCLE_BIN  : Tempat sampah tautan yang dihapus
 *  6. ANALYTICS    : Log kunjungan & klik tautan
 * ============================================================================
 */

// Nama-nama sheet standar
const SHEETS = {
  SETTINGS: 'SETTINGS',
  USERS: 'USERS',
  PROJECTS: 'PROJECTS',
  LINKS: 'LINKS',
  RECYCLE_BIN: 'RECYCLE_BIN',
  ANALYTICS: 'ANALYTICS'
};

/**
 * Endpoint Utama Web App (doGet)
 * Menampilkan antarmuka HTML atau melayani data publik jika diminta via URL
 */
function doGet(e) {
  // Pastikan struktur sheet tersedia sebelum melayani request
  initDatabaseTables();

  const template = HtmlService.createTemplateFromFile('index');
  
  // Masukkan metadata deployment agar client bisa mendeteksi URL resmi
  template.webAppUrl = getWebAppPublicUrl();
  template.requestedBio = (e && e.parameter && e.parameter.bio) ? e.parameter.bio : '';

  return template.evaluate()
    .setTitle('KSL Bio Link - Multi-Project Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Menyertakan file parsial (jika menggunakan pemisahan modular di Apps Script)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Mendapatkan URL deployment Web App resmi saat ini
 */
function getWebAppPublicUrl() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (err) {
    return '';
  }
}

/**
 * ============================================================================
 * 1. INISIALISASI DATABASE SHEET OTOMATIS
 * ============================================================================
 */
function initDatabaseTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Sheet USERS
  let userSheet = ss.getSheetByName(SHEETS.USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(SHEETS.USERS);
    userSheet.appendRow(['ID_USER', 'NAME', 'EMAIL', 'PASSWORD', 'ROLE', 'CREATED_AT']);
    userSheet.appendRow([
      'USR-' + Utilities.getUuid().slice(0, 8),
      'Super Administrator',
      'admin@ksl.local',
      'admin',
      'SUPER_ADMIN',
      new Date().toISOString()
    ]);
    formatHeaderRow(userSheet);
  }

  // 2. Sheet PROJECTS (Multi Bio)
  let projSheet = ss.getSheetByName(SHEETS.PROJECTS);
  if (!projSheet) {
    projSheet = ss.insertSheet(SHEETS.PROJECTS);
    projSheet.appendRow([
      'ID_PROJECT', 'SLUG', 'NAME', 'BIO', 'AVATAR_URL',
      'BG_COLOR', 'TEXT_COLOR', 'CUSTOM_BG_IMAGE', 'FOOTER_TEXT',
      'FONT_FAMILY', 'BTN_TYPE', 'BTN_EFFECT', 'UPDATED_AT'
    ]);
    projSheet.appendRow([
      'bio-utama',
      'utama',
      'Kucing Storia Labs',
      'Penyedia Solusi Web App, Google Apps Script, Dashboard Multi Bio Link & Automasi Bisnis.',
      'https://placehold.co/140x140/2563eb/ffffff?text=KSL',
      '#f8fafc',
      '#0f172a',
      '',
      'Powered by <strong>KSL Bio Link v5.0</strong>',
      "'Plus Jakarta Sans', sans-serif",
      'btn-shape-rounded',
      'btn-effect-lift',
      new Date().toISOString()
    ]);
    formatHeaderRow(projSheet);
  }

  // 3. Sheet LINKS (URL opsional / kondisional)
  let linkSheet = ss.getSheetByName(SHEETS.LINKS);
  if (!linkSheet) {
    linkSheet = ss.insertSheet(SHEETS.LINKS);
    linkSheet.appendRow([
      'ID_LINK', 'ID_PROJECT', 'ORDER_NUM', 'TITLE', 'URL', 'ICON',
      'STATUS', 'CTA_TYPE', 'CAMPAIGN', 'FEATURED', 'LOCKED', 'PASSWORD',
      'CLICKS', 'START_AT', 'END_AT', 'UPDATED_AT'
    ]);
    
    // Seed tautan awal demo
    const defaultLinks = [
      ['LNK-01', 'bio-utama', 1, 'Download Dashboard Generate Ver 2', 'https://drive.google.com/', '🔗', 'ACTIVE', 'DOWNLOAD', 'Source Code Dashboard', 'NO', 'NO', '', 2, '', '', new Date().toISOString()],
      ['LNK-02', 'bio-utama', 2, 'KSL Desaku Ver 1 ⭐ ⭐', 'https://youtube.com/', '▶️', 'ACTIVE', 'YOUTUBE', 'Launch Desaku Ver 1', 'YES', 'NO', '', 3, '', '', new Date().toISOString()],
      ['LNK-03', 'bio-utama', 3, 'KSL KasirKu Ver7 🔒', 'https://youtube.com/', '💻', 'ACTIVE', 'YOUTUBE', 'Launch KasirKu Ver 7', 'NO', 'YES', '123', 0, '', '', new Date().toISOString()],
      ['LNK-04', 'bio-utama', 4, 'Playlist Kas RT', 'https://youtube.com/', '🎬', 'ACTIVE', 'LINK', '', 'NO', 'NO', '', 13, '', '', new Date().toISOString()]
    ];
    defaultLinks.forEach(row => linkSheet.appendRow(row));
    formatHeaderRow(linkSheet);
  }

  // 4. Sheet RECYCLE_BIN
  let recycleSheet = ss.getSheetByName(SHEETS.RECYCLE_BIN);
  if (!recycleSheet) {
    recycleSheet = ss.insertSheet(SHEETS.RECYCLE_BIN);
    recycleSheet.appendRow([
      'ID_LINK', 'ID_PROJECT', 'ORDER_NUM', 'TITLE', 'URL', 'ICON',
      'STATUS', 'CTA_TYPE', 'CAMPAIGN', 'FEATURED', 'LOCKED', 'PASSWORD',
      'CLICKS', 'DELETED_AT'
    ]);
    formatHeaderRow(recycleSheet);
  }

  // 5. Sheet ANALYTICS
  let analyticsSheet = ss.getSheetByName(SHEETS.ANALYTICS);
  if (!analyticsSheet) {
    analyticsSheet = ss.insertSheet(SHEETS.ANALYTICS);
    analyticsSheet.appendRow(['TIMESTAMP', 'EVENT_TYPE', 'ID_PROJECT', 'ID_LINK', 'USER_AGENT']);
    formatHeaderRow(analyticsSheet);
  }

  // 6. Sheet SETTINGS
  let settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEETS.SETTINGS);
    settingsSheet.appendRow(['KEY', 'VALUE', 'DESCRIPTION']);
    settingsSheet.appendRow(['TOTAL_PAGE_VIEWS', '125', 'Total akumulasi kunjungan profil bio']);
    settingsSheet.appendRow(['APP_VERSION', '5.1.0', 'Versi build aplikasi']);
    formatHeaderRow(settingsSheet);
  }
}

function formatHeaderRow(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  header.setBackground('#1e293b')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setFontFamily('Arial');
  sheet.setFrozenRows(1);
}

/**
 * ============================================================================
 * 2. AUTENTIKASI ADMINISTRATOR
 * ============================================================================
 */
function apiLoginAdmin(email, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.USERS);
    if (!sheet) return { success: false, message: 'Database USERS tidak ditemukan.' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const userEmail = String(data[i][2]).toLowerCase().trim();
      const userPass = String(data[i][3]).trim();

      if (userEmail === String(email).toLowerCase().trim() && userPass === String(password).trim()) {
        const token = Utilities.base64Encode(Utilities.getUuid() + ':' + data[i][0]);
        return {
          success: true,
          token: token,
          user: {
            id: data[i][0],
            name: data[i][1],
            email: data[i][2],
            role: data[i][4]
          }
        };
      }
    }
    return { success: false, message: 'Email atau kata sandi tidak valid.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * ============================================================================
 * 3. PENGAMBILAN DATA KESELURUHAN (GET COMPLETE STATE)
 * ============================================================================
 */
function apiGetFullAppState() {
  try {
    initDatabaseTables();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Ambil Semua Project
    const projSheet = ss.getSheetByName(SHEETS.PROJECTS);
    const projData = projSheet.getDataRange().getValues();
    const projects = [];

    for (let i = 1; i < projData.length; i++) {
      projects.push({
        id: String(projData[i][0]),
        slug: String(projData[i][1]),
        name: String(projData[i][2]),
        bio: String(projData[i][3]),
        avatarUrl: String(projData[i][4]),
        bgColor: String(projData[i][5] || '#f8fafc'),
        textColor: String(projData[i][6] || '#0f172a'),
        customBgImage: String(projData[i][7] || ''),
        footerText: String(projData[i][8] || 'Powered by <strong>KSL Bio Link v5.0</strong>'),
        fontFamily: String(projData[i][9] || "'Plus Jakarta Sans', sans-serif"),
        btnType: String(projData[i][10] || 'btn-shape-rounded'),
        btnEffect: String(projData[i][11] || 'btn-effect-lift'),
        links: []
      });
    }

    // 2. Ambil Semua Links Aktif
    const linkSheet = ss.getSheetByName(SHEETS.LINKS);
    const linkData = linkSheet.getDataRange().getValues();

    for (let j = 1; j < linkData.length; j++) {
      const projId = String(linkData[j][1]);
      const parentProj = projects.find(p => p.id === projId);

      const linkObj = {
        id: String(linkData[j][0]),
        order: Number(linkData[j][2]) || 1,
        title: String(linkData[j][3]),
        url: String(linkData[j][4] || ''),
        icon: String(linkData[j][5] || '🔗'),
        status: String(linkData[j][6] || 'ACTIVE'),
        ctaType: String(linkData[j][7] || 'LINK'),
        campaign: String(linkData[j][8] || ''),
        featured: String(linkData[j][9]).toUpperCase() === 'YES',
        locked: String(linkData[j][10]).toUpperCase() === 'YES',
        password: String(linkData[j][11] || ''),
        clicks: Number(linkData[j][12]) || 0,
        startAt: String(linkData[j][13] || ''),
        endAt: String(linkData[j][14] || '')
      };

      if (parentProj) {
        parentProj.links.push(linkObj);
      }
    }

    // 3. Ambil Recycle Bin
    const recycleSheet = ss.getSheetByName(SHEETS.RECYCLE_BIN);
    const recycleData = recycleSheet.getDataRange().getValues();
    const recycleBin = [];

    for (let k = 1; k < recycleData.length; k++) {
      recycleBin.push({
        id: String(recycleData[k][0]),
        projectId: String(recycleData[k][1]),
        order: Number(recycleData[k][2]) || 1,
        title: String(recycleData[k][3]),
        url: String(recycleData[k][4] || ''),
        icon: String(recycleData[k][5]),
        status: String(recycleData[k][6]),
        ctaType: String(recycleData[k][7]),
        campaign: String(recycleData[k][8]),
        featured: String(recycleData[k][9]).toUpperCase() === 'YES',
        locked: String(recycleData[k][10]).toUpperCase() === 'YES',
        password: String(recycleData[k][11]),
        clicks: Number(recycleData[k][12]) || 0,
        deletedAt: String(recycleData[k][13])
      });
    }

    // 4. Ambil Daftar Pengguna
    const userSheet = ss.getSheetByName(SHEETS.USERS);
    const userData = userSheet.getDataRange().getValues();
    const adminUsers = [];

    for (let u = 1; u < userData.length; u++) {
      adminUsers.push({
        id: String(userData[u][0]),
        name: String(userData[u][1]),
        email: String(userData[u][2]),
        role: String(userData[u][4])
      });
    }

    // 5. Total Page Views
    const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
    const settingsData = settingsSheet.getDataRange().getValues();
    let totalViews = 125;
    for (let s = 1; s < settingsData.length; s++) {
      if (String(settingsData[s][0]) === 'TOTAL_PAGE_VIEWS') {
        totalViews = Number(settingsData[s][1]) || 0;
        break;
      }
    }

    return {
      success: true,
      bioProjects: projects,
      linksRecycleBin: recycleBin,
      adminUsers: adminUsers,
      analyticsViews: totalViews
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * ============================================================================
 * 4. PENGELOLAAN LINK (SAVE / REORDER / DELETE / RESTORE)
 * ============================================================================
 */
function apiSaveLink(linkData, projectId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.LINKS);
    const rows = sheet.getDataRange().getValues();

    const linkId = linkData.id || ('LNK-' + Utilities.getUuid().slice(0, 8).toUpperCase());
    const featured = linkData.featured ? 'YES' : 'NO';
    const locked = linkData.locked ? 'YES' : 'NO';
    const now = new Date().toISOString();

    let foundRowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(linkId)) {
        foundRowIndex = i + 1; // 1-indexed
        break;
      }
    }

    const rowPayload = [
      linkId,
      projectId,
      Number(linkData.order) || 1,
      linkData.title || '',
      linkData.url || '', // Kondisional (bisa kosong)
      linkData.icon || '🔗',
      linkData.status || 'ACTIVE',
      linkData.ctaType || 'LINK',
      linkData.campaign || '',
      featured,
      locked,
      linkData.password || '',
      Number(linkData.clicks) || 0,
      linkData.startAt || '',
      linkData.endAt || '',
      now
    ];

    if (foundRowIndex > 0) {
      sheet.getRange(foundRowIndex, 1, 1, rowPayload.length).setValues([rowPayload]);
    } else {
      sheet.appendRow(rowPayload);
    }

    return { success: true, id: linkId };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function apiReorderLinks(orderedList) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.LINKS);
    const data = sheet.getDataRange().getValues();

    orderedList.forEach(item => {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(item.id)) {
          sheet.getRange(i + 1, 3).setValue(Number(item.order));
          break;
        }
      }
    });

    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function apiMoveLinkToRecycleBin(linkId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const linkSheet = ss.getSheetByName(SHEETS.LINKS);
    const recycleSheet = ss.getSheetByName(SHEETS.RECYCLE_BIN);

    const data = linkSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(linkId)) {
        const deletedRow = [
          data[i][0], data[i][1], data[i][2], data[i][3], data[i][4],
          data[i][5], data[i][6], data[i][7], data[i][8], data[i][9],
          data[i][10], data[i][11], data[i][12], new Date().toLocaleString('id-ID')
        ];
        recycleSheet.appendRow(deletedRow);
        linkSheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Link tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function apiRestoreLink(linkId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const linkSheet = ss.getSheetByName(SHEETS.LINKS);
    const recycleSheet = ss.getSheetByName(SHEETS.RECYCLE_BIN);

    const data = recycleSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(linkId)) {
        const restoredRow = [
          data[i][0], data[i][1], data[i][2], data[i][3], data[i][4],
          data[i][5], data[i][6], data[i][7], data[i][8], data[i][9],
          data[i][10], data[i][11], data[i][12], '', '', new Date().toISOString()
        ];
        linkSheet.appendRow(restoredRow);
        recycleSheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Item di recycle bin tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function apiPermanentDelete(linkId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const recycleSheet = ss.getSheetByName(SHEETS.RECYCLE_BIN);

    const data = recycleSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(linkId)) {
        recycleSheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Item tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * ============================================================================
 * 5. PENGELOLAAN PROJECT MULTI-BIO & PENGATURAN TAMPILAN
 * ============================================================================
 */
function apiSaveProject(project) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.PROJECTS);
    const rows = sheet.getDataRange().getValues();

    const projId = project.id || ('bio-' + Utilities.getUuid().slice(0, 8));
    const now = new Date().toISOString();

    let foundIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(projId)) {
        foundIndex = i + 1;
        break;
      }
    }

    const rowPayload = [
      projId,
      project.slug || 'bio',
      project.name || 'Bio Name',
      project.bio || '',
      project.avatarUrl || 'https://placehold.co/140x140/2563eb/ffffff?text=KSL',
      project.bgColor || '#f8fafc',
      project.textColor || '#0f172a',
      project.customBgImage || '',
      project.footerText || 'Powered by <strong>KSL Bio Link v5.0</strong>',
      project.fontFamily || "'Plus Jakarta Sans', sans-serif",
      project.btnType || 'btn-shape-rounded',
      project.btnEffect || 'btn-effect-lift',
      now
    ];

    if (foundIndex > 0) {
      sheet.getRange(foundIndex, 1, 1, rowPayload.length).setValues([rowPayload]);
    } else {
      sheet.appendRow(rowPayload);
    }

    return { success: true, id: projId };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function apiDeleteProject(projectId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const projSheet = ss.getSheetByName(SHEETS.PROJECTS);
    const linkSheet = ss.getSheetByName(SHEETS.LINKS);

    // Hapus project
    const pData = projSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      if (String(pData[i][0]) === String(projectId)) {
        projSheet.deleteRow(i + 1);
        break;
      }
    }

    // Hapus tautan yang terafiliasi dengan project ini
    const lData = linkSheet.getDataRange().getValues();
    for (let j = lData.length - 1; j >= 1; j--) {
      if (String(lData[j][1]) === String(projectId)) {
        linkSheet.deleteRow(j + 1);
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function apiSaveAppearance(appearance, projectId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.PROJECTS);
    const rows = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(projectId)) {
        if (appearance.bgColor) sheet.getRange(i + 1, 6).setValue(appearance.bgColor);
        if (appearance.textColor) sheet.getRange(i + 1, 7).setValue(appearance.textColor);
        if (appearance.customBgImage !== undefined) sheet.getRange(i + 1, 8).setValue(appearance.customBgImage);
        if (appearance.footerText) sheet.getRange(i + 1, 9).setValue(appearance.footerText);
        if (appearance.fontFamily) sheet.getRange(i + 1, 10).setValue(appearance.fontFamily);
        if (appearance.btnType) sheet.getRange(i + 1, 11).setValue(appearance.btnType);
        if (appearance.btnEffect) sheet.getRange(i + 1, 12).setValue(appearance.btnEffect);
        if (appearance.avatarUrl) sheet.getRange(i + 1, 5).setValue(appearance.avatarUrl);
        sheet.getRange(i + 1, 13).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    return { success: false, message: 'Project tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * ============================================================================
 * 6. LOGGING KLIK & ANALYTICS
 * ============================================================================
 */
function apiIncrementClick(linkId, projectId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const linkSheet = ss.getSheetByName(SHEETS.LINKS);
    const data = linkSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(linkId)) {
        const currentClicks = Number(data[i][12]) || 0;
        linkSheet.getRange(i + 1, 13).setValue(currentClicks + 1);
        break;
      }
    }

    // Catat log klik di sheet ANALYTICS
    const analyticsSheet = ss.getSheetByName(SHEETS.ANALYTICS);
    if (analyticsSheet) {
      analyticsSheet.appendRow([
        new Date().toISOString(),
        'LINK_CLICK',
        projectId || '',
        linkId,
        ''
      ]);
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
