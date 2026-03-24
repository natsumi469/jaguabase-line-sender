// ==========================================
// JAGUABASE LINE自動送信システム
// ==========================================

const SHEET_NAMES = {
  CUSTOMERS: '顧客管理',
  STAFF:     'スタッフマスタ',
  STORES:    '店舗マスタ',
  LOGS:      '送信ログ',
  QUEUE:     '処理キュー',
  CARTE:     'カルテ'
};

const DRIVE_FOLDER_ID  = '1k1YwBxS_Ep2OE7qIm0ve3N3IuK-YjK6r';
const SLACK_WEBHOOK_URL = 'https://hooks.slack.com/services/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';

// ==========================================
// エントリポイント
// ==========================================

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getCarte') {
    return getCarteForBookmarklet(e.parameter.storesId);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('JAGUABASE 施術写真送信')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let queueSheet = ss.getSheetByName(SHEET_NAMES.QUEUE);
    if (!queueSheet) {
      queueSheet = ss.insertSheet(SHEET_NAMES.QUEUE);
      queueSheet.appendRow(['受信日時', 'storeCode', 'payload', '処理済み']);
    }
    const storeCode = e.parameter.store ? e.parameter.store.toUpperCase() : '';
    queueSheet.appendRow([new Date(), storeCode, e.postData.contents, 'false']);
  } catch(err) {}
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// キュー処理
// ==========================================

function processQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName(SHEET_NAMES.QUEUE);
  if (!queueSheet) return;

  const data = queueSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][3]);
    if (status === 'true' || status === 'done' ||
        status.startsWith('error') || status.startsWith('skip')) continue;

    const storeCode = data[i][1];
    const payload   = data[i][2];
    try {
      const json = JSON.parse(payload);
      if (!json.events || json.events.length === 0) {
        queueSheet.getRange(i + 1, 4).setValue('skip:empty');
        continue;
      }
      const storeInfo = getStoreInfo(storeCode);
      if (!storeInfo) {
        queueSheet.getRange(i + 1, 4).setValue('error:store notfound:' + storeCode);
        continue;
      }

      let result = 'skip:no target event';
      for (const event of json.events) {
        if (event.type !== 'follow' && event.type !== 'message') continue;
        const userId  = event.source.userId;
        const profile = getLineProfile(storeInfo.token, userId);
        if (!profile) {
          result = 'error:profile null:' + userId;
          break;
        }
        upsertCustomer(profile.displayName, userId, storeCode);

        if (event.type === 'message' && event.message && event.message.type === 'text') {
          const cleanText     = event.message.text.trim().replace(/^text=/, '');
          const storesIdMatch = cleanText.match(/^storesId:(\d+)(?:\s+(同意|非同意))?$/);
          if (storesIdMatch) {
            const sid     = storesIdMatch[1];
            const consent = storesIdMatch[2] === '非同意' ? 'no' : 'yes';
            linkUserWithStores(userId, profile.displayName, sid, consent);
            result = 'done:linked:' + sid + ':' + consent;
            break;
          }
          const extractedName = extractNameFromMessage(cleanText);
          if (extractedName) updateCustomerName(userId, extractedName);
        }
        result = 'done:' + profile.displayName;
      }
      queueSheet.getRange(i + 1, 4).setValue(result);
    } catch(err) {
      queueSheet.getRange(i + 1, 4).setValue('error:' + err.message);
    }
  }
}

function setupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processQueue') ScriptApp.deleteTrigger(trigger);
  }
  ScriptApp.newTrigger('processQueue').timeBased().everyMinutes(1).create();
}

// ==========================================
// 顧客一覧（PWA用）
// ==========================================

function getTodayCustomers(dateOffset) {
  const offset     = dateOffset || 0;
  const base       = new Date();
  base.setDate(base.getDate() + offset);
  const today      = base;
  const startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
  const endOfDay   = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
  const sheet        = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CUSTOMERS);
  const customerData = sheet.getDataRange().getValues();
  const customers    = [];

  for (const calendar of CalendarApp.getAllCalendars()) {
    for (const event of calendar.getEvents(startOfDay, endOfDay)) {
      const title       = event.getTitle();
      const description = event.getDescription() || '';

      const nameMatch = title.match(/^(.+?)[\:：]/);
      if (!nameMatch) continue;
      const fullName = nameMatch[1].trim();
      const lastName = fullName.split(/[\s\u3000]/)[0];

      const staffMatch = description.match(/担当スタッフ[：:]\s*(.+)/);
      const staffName  = staffMatch ? staffMatch[1].trim() : '未設定';
      const menuMatch  = description.match(/メニュー[：:]\s*(.+)/);
      const menuText   = menuMatch ? menuMatch[1].trim() : '';

      const coubicMatch    = description.match(/coubic\.com\/dashboard\/customers\/(\d+)/);
      const storesAppMatch = description.match(/reserve\.stores\.app\/dashboard\/customers\/(\d+)/);
      const storesIdFromCal = storesAppMatch ? storesAppMatch[1] : (coubicMatch ? coubicMatch[1] : null);
      const isStoresAppId   = !!storesAppMatch;

      let matchedDisplayName = null;
      let matchedStoresId    = null;
      let matchedRow         = -1;

      // H列を照合（coubic・STORES.app問わず storesIdFromCal で検索）
      if (storesIdFromCal) {
        for (let i = 1; i < customerData.length; i++) {
          if (String(customerData[i][7]) === storesIdFromCal) {
            matchedDisplayName = customerData[i][0];
            matchedStoresId    = storesIdFromCal;
            matchedRow         = i + 1;
            break;
          }
        }
      }

      // 名前マッチング（H列で見つからない場合のみ）
      if (!matchedDisplayName) {
        const mmdd = String(today.getMonth()+1).padStart(2,'0') + String(today.getDate()).padStart(2,'0');
        for (let i = 1; i < customerData.length; i++) {
          const displayName = customerData[i][0];
          if (!displayName) continue;
          const parsed = parseLineName(displayName);
          if (parsed.date && parsed.date !== mmdd) continue;
          if (parsed.lastName && parsed.lastName === lastName) {
            matchedDisplayName = displayName;
            matchedStoresId    = customerData[i][7] ? String(customerData[i][7]) : null;
            matchedRow         = i + 1;
            break;
          }
        }
      }

      // STORES.appのIDのみH列に自動保存（coubic IDは混入させない）
      if (isStoresAppId && matchedRow > 0 && !matchedStoresId) {
        sheet.getRange(matchedRow, 8).setValue(storesIdFromCal);
        matchedStoresId = storesIdFromCal;
      } else if (isStoresAppId && !matchedStoresId) {
        matchedStoresId = storesIdFromCal;
      }

      // 店舗のbasicIdを取得
      let basicId = null;
      if (matchedRow > 0 && customerData[matchedRow - 1][4]) {
        const si = getStoreInfo(customerData[matchedRow - 1][4]);
        if (si) basicId = si.basicId;
      }

      const startTime = event.getStartTime();
      const timeStr   = startTime.getHours() + ':' + String(startTime.getMinutes()).padStart(2, '0');

      customers.push({
        displayName: matchedDisplayName || fullName,
        lastName:    fullName,
        time:        timeStr,
        staff:       staffName,
        menu:        menuText,
        hasLineId:   !!matchedDisplayName,
        storesId:    storesIdFromCal || matchedStoresId,
        basicId:     basicId
      });
    }
  }

  customers.sort((a, b) => a.time.localeCompare(b.time));
  return customers;
}

// ==========================================
// 写真送信
// ==========================================

function sendPhotoFromPWA(displayName, storesId, photosArray, carteData) {
  try {
    const folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const today   = new Date();
    const dateStr = (today.getMonth()+1) + String(today.getDate()).padStart(2,'0');

    let resolvedName = displayName;
    if (storesId) {
      const rows = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(SHEET_NAMES.CUSTOMERS).getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][7]) === String(storesId)) {
          resolvedName = rows[i][0];
          break;
        }
      }
    }

    const consent    = getCustomerConsent(resolvedName);
    const filePrefix = consent === 'no' ? 'NG_' : '';
    const design     = (carteData && carteData.design)
      ? '_' + carteData.design.replace(/[\/\\:*?"<>|]/g, '').trim()
      : '';
    const imageUrls  = [];

    for (let i = 0; i < photosArray.length; i++) {
      const photo    = photosArray[i];
      const mimeType = (photo.mimeType && photo.mimeType.startsWith('image/')) ? photo.mimeType : 'image/jpeg';
      const ext      = mimeType === 'image/png' ? 'png' : 'jpg';
      const fileName = filePrefix + dateStr + '_' + resolvedName + design + '_' + (i+1) + '.' + ext;
      const blob     = Utilities.newBlob(Utilities.base64Decode(photo.base64), mimeType, fileName);
      const file     = folder.createFile(blob);
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(shareErr) {
        console.log('setSharing失敗（共有制限の可能性）: ' + shareErr.message);
      }
      imageUrls.push('https://lh3.googleusercontent.com/d/' + file.getId());
    }

    if (carteData) {
      carteData.displayName = resolvedName;
      saveCarteData(carteData);
    }
    return sendPhotosToCustomer(resolvedName, storesId, imageUrls);
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function sendPhotosToCustomer(displayName, storesId, imageUrls) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CUSTOMERS);
  const data  = sheet.getDataRange().getValues();

  let row = -1;
  if (storesId) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][7]) === String(storesId)) { row = i; break; }
    }
  }
  if (row === -1) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === displayName) { row = i; break; }
    }
  }
  if (row === -1) {
    writeLog(displayName, '', 'error', 'customer not found');
    return { success: false, error: 'customer not found' };
  }

  const userId       = data[row][1];
  const storeCode    = data[row][4];
  const staffInitial = data[row][5];
  const resolvedName = data[row][0];

  if (!userId) {
    writeLog(resolvedName, storeCode, 'error', 'LINE ID not found');
    return { success: false, error: 'LINE ID not found' };
  }
  const storeInfo = getStoreInfo(storeCode);
  if (!storeInfo) {
    writeLog(resolvedName, storeCode, 'error', 'store not found');
    return { success: false, error: 'store not found' };
  }
  const staffName = staffInitial ? getStaffName(staffInitial) : null;
  try {
    sendLineMessage(storeInfo.token, userId, imageUrls, storeInfo.name, staffName);
    writeLog(resolvedName, storeCode, 'success', '');
    return { success: true };
  } catch(e) {
    writeLog(resolvedName, storeCode, 'error', e.message);
    return { success: false, error: e.message };
  }
}

function sendLineMessage(token, userId, imageUrls, storeName, staffName) {
  const storeLabel = storeName || 'JAGUABASE';
  let messageText = 'ご来店ありがとうございました\n施術のお写真をお送りします\n\nこちらの写真はSNSなどに自由にアップしていただいて構いません\nもしよければInstagramで\n@jaguabase_jaguatattoo\nをメンション・タグ付けしていただけるととっても嬉しいです\n\nまたのご来店をお待ちしております\n' + storeLabel;
  if (staffName) messageText += '\n担当：' + staffName;

  const messages = imageUrls.map(url => ({
    type: 'image', originalContentUrl: url, previewImageUrl: url
  }));
  messages.push({ type: 'text', text: messageText });

  const lineResp = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({ to: userId, messages: messages }),
    muteHttpExceptions: true
  });
  if (lineResp.getResponseCode() !== 200) {
    throw new Error('LINE API ' + lineResp.getResponseCode() + ': ' + lineResp.getContentText().substring(0, 300));
  }
}

// ==========================================
// カルテ
// ==========================================

function saveCarteData(carteData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (carteData.storesId) {
      const rows = ss.getSheetByName(SHEET_NAMES.CUSTOMERS).getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][7]) === String(carteData.storesId)) {
          carteData.displayName = rows[i][0];
          break;
        }
      }
    }
    let carteSheet = ss.getSheetByName(SHEET_NAMES.CARTE);
    if (!carteSheet) {
      carteSheet = ss.insertSheet(SHEET_NAMES.CARTE);
      carteSheet.appendRow(['LINE表示名', '施術日時', '担当スタッフ', 'サイズ・デザイン', '施術部位', 'メモ']);
    }
    carteSheet.appendRow([
      carteData.displayName,
      new Date(),
      carteData.staff  || '',
      carteData.design || '',
      carteData.area   || '',
      carteData.memo   || ''
    ]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getCarteForBookmarklet(storesId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const customerData = ss.getSheetByName(SHEET_NAMES.CUSTOMERS).getDataRange().getValues();
    let displayName = null;
    for (let i = 1; i < customerData.length; i++) {
      if (String(customerData[i][7]) === String(storesId)) {
        displayName = customerData[i][0];
        break;
      }
    }
    if (!displayName) return jsonResponse({ success: false, error: 'STORES ID not registered' });

    const carteSheet = ss.getSheetByName(SHEET_NAMES.CARTE);
    if (!carteSheet) return jsonResponse({ success: false, error: 'No carte sheet' });

    const carteData = carteSheet.getDataRange().getValues();
    const entries = [];
    for (let i = 1; i < carteData.length; i++) {
      if (carteData[i][0] === displayName) {
        entries.push({ date: carteData[i][1], staff: carteData[i][2], design: carteData[i][3], area: carteData[i][4], memo: carteData[i][5] });
      }
    }
    if (entries.length === 0) return jsonResponse({ success: false, error: 'No carte data' });

    const latest  = entries[entries.length - 1];
    const d       = new Date(latest.date);
    const dateStr = d.getFullYear() + '.' + String(d.getMonth()+1).padStart(2,'0') + '.' + String(d.getDate()).padStart(2,'0');
    let text = dateStr + (latest.staff ? '.' + latest.staff : '') + '\n';
    if (latest.design) text += latest.design;
    if (latest.area)   text += '(' + latest.area + ')';
    if (latest.design || latest.area) text += '\n';
    if (latest.memo)   text += latest.memo + '\n';
    return jsonResponse({ success: true, text: text.trim(), customerName: displayName });
  } catch(e) {
    return jsonResponse({ success: false, error: e.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 顧客管理
// ==========================================

function upsertCustomer(displayName, userId, storeCode) {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CUSTOMERS);
  const data   = sheet.getDataRange().getValues();
  const parsed = parseLineName(displayName);
  const store  = storeCode || parsed.store || '';

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userId) {
      sheet.getRange(i + 1, 1).setValue(displayName);
      if (store && !data[i][4]) sheet.getRange(i + 1, 5).setValue(store);
      return;
    }
  }
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === displayName) {
      sheet.getRange(i + 1, 2).setValue(userId);
      if (store && !data[i][4]) sheet.getRange(i + 1, 5).setValue(store);
      return;
    }
  }
  sheet.appendRow([displayName, userId, parsed.lastName || '', parsed.date || '', store, parsed.staffInitial || '', new Date(), '']);
}

function saveStoresId(displayName, storesId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CUSTOMERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === displayName) {
        sheet.getRange(i + 1, 8).setValue(storesId);
        return { success: true };
      }
    }
    return { success: false, error: 'customer not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function linkUserWithStores(userId, displayName, storesId, consent) {
  try {
    upsertCustomer(displayName, userId, '');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CUSTOMERS);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userId) {
        if (storesId) sheet.getRange(i + 1, 8).setValue(storesId);
        sheet.getRange(i + 1, 9).setValue(consent ? 'yes' : 'no');
        return { success: true };
      }
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function extractNameFromMessage(text) {
  const patterns = [
    /お名前[：:]\s*([^\s\n　]{2,10}(?:[\s　][^\s\n　]{1,6})?)/,
    /予約(?:した|の)?[\s　]?([一-龯]{1,4}[\s　][一-龯]{1,4}|[一-龯]{2,5})(?:です|と申します|といいます)/,
    /^([一-龯]{1,4}[\s　]?[一-龯]{1,4})(?:です|と申します|といいます)/m,
  ];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1].trim().length >= 2) return match[1].trim();
  }
  return null;
}

function updateCustomerName(userId, newName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CUSTOMERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] !== userId) continue;
    const hoursSince = (new Date() - new Date(data[i][6])) / (1000 * 60 * 60);
    if (hoursSince > 48) return;
    const parsed = parseLineName(newName);
    sheet.getRange(i + 1, 1).setValue(newName);
    if (parsed.lastName) sheet.getRange(i + 1, 3).setValue(parsed.lastName);
    return;
  }
}

// ==========================================
// マスタ参照
// ==========================================

function parseLineName(displayName) {
  if (!displayName) return { staffInitial: null, date: null, store: null, lastName: null };
  const staffMatch = displayName.match(/^([A-Za-z]{1,2})/);
  const dateMatch  = displayName.match(/(\d{4})/);
  const storeMatch = displayName.match(/([KODNkodn])$/);
  const nameMatch  = displayName.match(/[\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff]+/);
  return {
    staffInitial: staffMatch ? staffMatch[1].toUpperCase() : null,
    date:         dateMatch  ? dateMatch[1]                : null,
    store:        storeMatch ? storeMatch[1].toUpperCase() : null,
    lastName:     nameMatch  ? nameMatch[0]                : null
  };
}

function getStaffName(initial) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STAFF);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toUpperCase() === initial.toUpperCase()) return data[i][1];
  }
  return null;
}

// 店舗マスタ: A列=店舗コード, B列=店舗名, C列=チャネルトークン, D列=basicId
function getStoreInfo(storeCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STORES);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toUpperCase() === storeCode.toUpperCase()) {
      return { name: data[i][1], token: data[i][2], basicId: data[i][3] || null };
    }
  }
  return null;
}

function getLineProfile(token, userId) {
  try {
    const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + userId,
      { method: 'get', headers: { 'Authorization': 'Bearer ' + token } });
    return JSON.parse(response.getContentText());
  } catch(e) { return null; }
}

function getStores() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STORES);
  const data  = sheet.getDataRange().getValues();
  const stores = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][3]) {
      stores.push({ code: String(data[i][0]), name: String(data[i][1]), basicId: String(data[i][3]) });
    }
  }
  return stores;
}

function getQRDataUrl(lineUrl) {
  var url      = 'https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=' + encodeURIComponent(lineUrl);
  var response = UrlFetchApp.fetch(url);
  var base64   = Utilities.base64Encode(response.getBlob().getBytes());
  return 'data:image/png;base64,' + base64;
}

function getCustomerConsent(displayName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CUSTOMERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === displayName) return data[i][8] || '';
  }
  return '';
}

// ==========================================
// ログ・通知
// ==========================================

function writeLog(displayName, store, status, error) {
  try {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LOGS);
    if (!logSheet) return;
    logSheet.appendRow([new Date(), displayName, store, status, error || '']);
  } catch(e) {}
}

function checkDailyCarteAndNotify() {
  const today    = new Date();
  const todayStr = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');
  const customers = getTodayCustomers();
  if (!customers || customers.length === 0) return;

  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const carteSheet = ss.getSheetByName(SHEET_NAMES.CARTE);
  const logSheet   = ss.getSheetByName(SHEET_NAMES.LOGS);

  const carteToday = new Set();
  if (carteSheet) {
    for (const row of carteSheet.getDataRange().getValues().slice(1)) {
      const d = new Date(row[1]);
      const s = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
      if (s === todayStr) carteToday.add(row[0]);
    }
  }

  const photoToday = new Set();
  if (logSheet) {
    for (const row of logSheet.getDataRange().getValues().slice(1)) {
      const d = new Date(row[0]);
      const s = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
      if (s === todayStr && row[3] === 'success') photoToday.add(row[1]);
    }
  }

  const missing = [];
  for (const c of customers) {
    if (!c.hasLineId) continue;
    const noCarte = !carteToday.has(c.displayName);
    const noPhoto = !photoToday.has(c.displayName);
    if (noCarte || noPhoto) {
      missing.push(c.time + ' ' + c.displayName + (noCarte ? '[カルテ未入力]' : '') + (noPhoto ? '[写真未送信]' : ''));
    }
  }
  if (missing.length === 0) return;

  UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: '本日の未対応予約があります\n' + missing.join('\n') })
  });
}

function setupDailyNotifyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'checkDailyCarteAndNotify') ScriptApp.deleteTrigger(trigger);
  }
  ScriptApp.newTrigger('checkDailyCarteAndNotify').timeBased().atHour(21).nearMinute(0).everyDays(1).create();
}
