/**
 * Main entry point for Google Apps Script Web App.
 * Serves the index.html page or returns JSON data for API requests.
 */
function doGet(e) {
  setupDatabase();
  
  // If request contains API parameter, return JSON response
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
 * Automatically sets up required sheets ('Profile' and 'Links') if they do not exist.
 * Zero-config setup using SpreadsheetApp.getActiveSpreadsheet().
 */
function setupDatabase() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log('Lock timeout during setupDatabase');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Setup 'Profile' Sheet
    var profileSheet = ss.getSheetByName('Profile');
    if (!profileSheet) {
      profileSheet = ss.insertSheet('Profile');
      profileSheet.appendRow(['Key', 'Value']);
      profileSheet.appendRow(['Name', 'Nama Anda']);
      profileSheet.appendRow(['Bio', 'Digital Creator & Innovator']);
      profileSheet.appendRow(['Avatar', 'https://images.unsplash.com/photo-1534528741775-53994a69daeb?auto=format&fit=crop&w=300&q=80']);
      profileSheet.appendRow(['Background', 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)']);
      profileSheet.appendRow(['Password', 'admin123']);
      profileSheet.appendRow(['FontFamily', "'Inter', sans-serif"]);
      profileSheet.appendRow(['FontSize', '100%']);
      profileSheet.appendRow(['TextColor', 'rgb(255, 255, 255)']);
      profileSheet.appendRow(['ButtonTextColor', 'rgb(255, 255, 255)']);
      profileSheet.appendRow(['Subdomain', 'username']);
      profileSheet.appendRow(['CustomDomain', '']);
      profileSheet.appendRow(['VerifiedBadge', 'blue']);
      profileSheet.appendRow(['AnalyticsId', '']);
      profileSheet.appendRow(['MetaTitle', '']);
      profileSheet.appendRow(['MetaDesc', '']);
      profileSheet.appendRow(['CustomCss', '']);
      
      profileSheet.getRange("A1:B1").setFontWeight("bold");
      profileSheet.setFrozenRows(1);
    }

    // 2. Setup 'Links' Sheet
    var linksSheet = ss.getSheetByName('Links');
    if (!linksSheet) {
      linksSheet = ss.insertSheet('Links');
      linksSheet.appendRow(['ID', 'Type', 'Title', 'URL', 'Icon', 'Content', 'Color', 'Effect', 'Active']);
      linksSheet.appendRow([generateId(), 'link', 'Portofolio Utama', 'https://example.com', '💼', '', '#4F46E5', 'btn-effect-scale', 'TRUE']);
      linksSheet.appendRow([generateId(), 'whatsapp', 'Chat WhatsApp Kami', '628123456789', '💬', 'Halo, saya tertarik dengan jasa Anda!', '#25D366', 'btn-effect-scale', 'TRUE']);
      
      linksSheet.getRange("A1:I1").setFontWeight("bold");
      linksSheet.setFrozenRows(1);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Generates a random unique ID string for components.
 */
function generateId() {
  return 'id_' + Math.random().toString(36).substring(2, 11);
}

/**
 * Reads all data from 'Profile' and 'Links' sheets and formats it into a structured object.
 */
function getAppData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Read Profile Sheet
  var profileSheet = ss.getSheetByName('Profile');
  var profileData = profileSheet ? profileSheet.getDataRange().getValues() : [];
  var profile = {};
  for (var i = 1; i < profileData.length; i++) {
    if (profileData[i][0]) {
      profile[profileData[i][0]] = profileData[i][1];
    }
  }
  
  // Read Links Sheet
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
        color: String(linksData[j][6] || ''),
        effect: String(linksData[j][7] || 'btn-effect-scale'),
        active: String(linksData[j][8]).toUpperCase() !== 'FALSE'
      });
    }
  }
  
  return {
    profile: profile,
    links: links
  };
}

/**
 * Verifies admin password against stored password in Profile sheet.
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
 * Saves or updates profile key-value settings in the 'Profile' sheet.
 */
function saveProfile(name, bio, avatar, password, background, fontFamily, fontSize, textColor, buttonTextColor, subdomain, customDomain, verifiedBadge, analyticsId, metaTitle, metaDesc, customCss) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server busy. Lock timeout.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Profile');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Profile');
    }
    
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
    updateOrAdd('Avatar', avatar);
    updateOrAdd('Background', background);
    updateOrAdd('FontFamily', fontFamily);
    updateOrAdd('FontSize', fontSize);
    updateOrAdd('TextColor', textColor);
    updateOrAdd('ButtonTextColor', buttonTextColor);
    updateOrAdd('Subdomain', subdomain);
    updateOrAdd('CustomDomain', customDomain);
    updateOrAdd('VerifiedBadge', verifiedBadge);
    updateOrAdd('AnalyticsId', analyticsId);
    updateOrAdd('MetaTitle', metaTitle);
    updateOrAdd('MetaDesc', metaDesc);
    updateOrAdd('CustomCss', customCss);

    if (password && String(password).trim() !== '') {
      updateOrAdd('Password', String(password).trim());
    }
    
    return { success: true, message: 'Profil berhasil diperbarui' };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Saves a new component or updates an existing component in 'Links' sheet.
 */
function saveComponent(id, type, title, url, icon, content, color, effect, active) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server busy. Lock timeout.' };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Links');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Links');
    }
    
    var activeStr = (active === false || active === 'false') ? 'FALSE' : 'TRUE';
    var targetId = id || generateId();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(targetId)) {
        sheet.getRange(i + 1, 2, 1, 8).setValues([[
          type || 'link',
          title || '',
          url || '',
          icon || '',
          content || '',
          color || '#4F46E5',
          effect || 'btn-effect-scale',
          activeStr
        ]]);
        return { success: true, id: targetId };
      }
    }
    
    sheet.appendRow([
      targetId,
      type || 'link',
      title || '',
      url || '',
      icon || '',
      content || '',
      color || '#4F46E5',
      effect || 'btn-effect-scale',
      activeStr
    ]);
    return { success: true, id: targetId };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Updates row ordering in 'Links' sheet based on array of IDs.
 */
function updateComponentOrder(newIds) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Lock timeout.' };
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
 * Deletes a link component row from 'Links' sheet by ID.
 */
function deleteLink(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Lock timeout.' };
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
    return { success: false, message: 'ID not found' };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handles incoming POST requests sent via fetch() with stringified JSON payload.
 */
function doPost(e) {
  try {
    setupDatabase();
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }
    
    var action = payload.action;
    var result = { success: false, message: 'Invalid action' };
    
    if (action === 'getAppData') {
      result = { success: true, data: getAppData() };
    } else if (action === 'verifyAuth') {
      var isAuth = verifyAuth(payload.password);
      result = { success: isAuth, message: isAuth ? 'Auth success' : 'Invalid password' };
    } else if (action === 'saveProfile') {
      result = saveProfile(
        payload.Name, payload.Bio, payload.Avatar, payload.Password,
        payload.Background, payload.FontFamily, payload.FontSize, payload.TextColor,
        payload.ButtonTextColor, payload.Subdomain, payload.CustomDomain,
        payload.VerifiedBadge, payload.AnalyticsId, payload.MetaTitle,
        payload.MetaDesc, payload.CustomCss
      );
    } else if (action === 'saveComponent') {
      result = saveComponent(
        payload.id, payload.type, payload.title, payload.url,
        payload.icon, payload.content, payload.color, payload.effect, payload.active
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