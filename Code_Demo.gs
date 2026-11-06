/**
 * =========================================================================
 * 🏫 智慧大屏班務系統 - 【公共 Demo 展示與 7 天試用專用版】
 * =========================================================================
 * ⚠️ 注意事項：
 * 1. 本檔案專門部署在展示用、體驗用的公共 GAS 專案上。
 * 2. 支援「使用者名稱 / 班級代碼 (Tenant ID)」多租戶隔離，多位老師同時試用互不干擾。
 * 3. 內建 7 天過期資料定時自動清理程序 (autoCleanupExpiredDemoData)。
 * 4. 若教師要建立正式獨立的班級系統，請使用專案根目錄的【純淨正式版 Code.gs】。
 * =========================================================================
 */

// 預備預設 Gemini API Key (可於 GAS「專案設定 ⚙️ ➔ 指令碼屬性」填入 GEMINI_API_KEY，或於試算表 Settings 分頁填入，以策安全不外洩)
const DEFAULT_FALLBACK_API_KEY = "";
const DEMO_EXPIRE_DAYS = 7; // 試用有效期 7 天

/**
 * 取得 Gemini API Key：
 * 1. 優先讀取 Google 試算表中的「Settings」或「設定」分頁
 * 2. 其次讀取 GAS 的指令碼屬性 (Script Properties)
 * 3. 若皆未填寫，自動使用系統後端預設 Key
 */
function getGeminiApiKey() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      const settingsSheet = ss.getSheetByName("Settings") || ss.getSheetByName("設定");
      if (settingsSheet && settingsSheet.getLastRow() >= 2) {
        const data = settingsSheet.getDataRange().getValues();
        for (let i = 0; i < data.length; i++) {
          const keyName = String(data[i][0] || "").trim().toUpperCase();
          if (keyName === "GEMINI_API_KEY" || keyName === "GEMINI API KEY" || keyName === "API_KEY" || keyName === "GEMINI_KEY" || keyName === "GEMINI金鑰") {
            const val = String(data[i][1] || "").trim();
            if (val && !val.includes("請填入") && !val.includes("你的API_KEY")) {
              return val;
            }
          }
        }
      }
    }
    const propKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (propKey && propKey.trim()) {
      return propKey.trim();
    }
  } catch (err) {
    Logger.log("讀取 Gemini API Key 失敗: " + err);
  }
  return DEFAULT_FALLBACK_API_KEY;
}

const GEMINI_API_KEY = getGeminiApiKey();
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// 計算當日與隔日字串 (以 GMT+8 台灣時間為準)
const now = new Date();
const todayDate = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");
const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
const tomorrowDate = Utilities.formatDate(tomorrow, "GMT+8", "yyyy-MM-dd");

// =========================================================================
// 🏢 多租戶資料庫初始化與管理模組
// =========================================================================

function getTenantId(data, e) {
  let tid = '';
  if (data && data.tenantId) tid = String(data.tenantId).trim();
  else if (e && e.parameter && e.parameter.tenantId) tid = String(e.parameter.tenantId).trim();
  return tid ? tid : 'demo';
}

function ensureDemoSheetsExist(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetDefs = [
    { name: "Demo_Tenants", headers: ["Tenant_ID", "建立時間", "最後活動時間", "班級名稱", "PIN碼"] },
    { name: "Students", headers: ["Tenant_ID", "座號", "姓名", "密碼", "未繳項目", "照片網址", "辨識狀態"] },
    { name: "All_Records", headers: ["Tenant_ID", "紀錄ID", "登記時間", "類別", "內容範圍", "執行日期", "來源"] },
    { name: "Missing_Items", headers: ["Tenant_ID", "紀錄ID", "登記日期", "學生座號", "項目名稱", "狀態", "銷案時間"] },
    { name: "Seating_Photos", headers: ["Tenant_ID", "座號", "姓名", "座位編號", "照片網址", "辨識狀態", "更新時間"] },
    { name: "Seating_State", headers: ["Tenant_ID", "座位佈局JSON", "更新時間"] }
  ];

  sheetDefs.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers])
           .setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }
  });
}

function initTenantIfNeeded(ss, tenantId) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureDemoSheetsExist(ss);

  const tSheet = ss.getSheetByName("Demo_Tenants");
  const data = tSheet.getDataRange().getValues();
  let foundRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === tenantId) {
      foundRow = i + 1;
      break;
    }
  }

  const nowIso = new Date().toISOString();

  if (foundRow > -1) {
    // 更新最後活躍時間
    tSheet.getRange(foundRow, 3).setValue(nowIso);
    return;
  }

  // 新租戶：註冊並播種 10 名預設學生
  tSheet.appendRow([tenantId, nowIso, nowIso, `${tenantId} 班`, "8888"]);

  const sSheet = ss.getSheetByName("Students");
  const defaultStudents = [
    [tenantId, "1", "王小明", "1234", "", "", "正常"],
    [tenantId, "2", "李大華", "1234", "", "", "正常"],
    [tenantId, "3", "張美美", "1234", "", "", "正常"],
    [tenantId, "4", "陳志豪", "1234", "", "", "正常"],
    [tenantId, "5", "林雅婷", "1234", "", "", "正常"],
    [tenantId, "6", "黃俊傑", "1234", "", "", "正常"],
    [tenantId, "7", "吳淑芬", "1234", "", "", "正常"],
    [tenantId, "8", "蔡宗憲", "1234", "", "", "正常"],
    [tenantId, "9", "許瑋甯", "1234", "", "", "正常"],
    [tenantId, "10", "鄭家豪", "1234", "", "", "正常"]
  ];
  defaultStudents.forEach(row => sSheet.appendRow(row));

  // 播種預設聯絡簿
  const rSheet = ss.getSheetByName("All_Records");
  rSheet.appendRow([tenantId, "REC-DEMO-1", Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss"), "作業", "國語第 1 課生字", tomorrowDate, "系統預設"]);
  rSheet.appendRow([tenantId, "REC-DEMO-2", Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss"), "提醒事項", "帶彩色筆、美勞剪刀", tomorrowDate, "系統預設"]);
}

// =========================================================================
// 🌐 GET 請求路由
// =========================================================================

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = (e && e.parameter) ? e.parameter.action : null;
  const tenantId = getTenantId(null, e);
  initTenantIfNeeded(ss, tenantId);

  if (action === 'getClassTitle' || action === 'getTeacherPin') {
    return ContentService.createTextOutput(JSON.stringify({ 
      success: true,
      className: getClassTitle(tenantId),
      teacherPin: getTeacherPin(tenantId)
    })).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'getStudentData') {
    return ContentService.createTextOutput(JSON.stringify(getStudentData(tenantId)))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'loadSeatingState') {
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: loadSeatingState(tenantId) }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'getStudentPhotos') {
    return ContentService.createTextOutput(JSON.stringify(getStudentPhotos(tenantId)))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'none') {
    return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.JSON);
  }

  // 座位表頁面
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase() : '';
  const role = (e && e.parameter && e.parameter.role) ? e.parameter.role : 'screen';
  const className = getClassTitle(tenantId);

  if (page === 'seating') {
    const template = HtmlService.createTemplateFromFile('Seating');
    template.origin = (e && e.parameter && e.parameter.origin) ? e.parameter.origin : '';
    template.role = role;
    template.className = className;
    template.geminiApiKey = '';
    template.mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : '';
    template.tenantId = tenantId;

    return template.evaluate()
      .setTitle(`${className} - 班級座位表`)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.userRole = role;
  template.className = className;
  template.tenantId = tenantId;

  return template.evaluate()
    .setTitle(`${className} - 大屏班務系統 (Demo試用版)`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =========================================================================
// 🌐 POST 請求路由
// =========================================================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'ping') {
      return ContentService.createTextOutput(JSON.stringify({ status: 'pong' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = data.action;
    const tenantId = getTenantId(data, e);
    initTenantIfNeeded(ss, tenantId);

    let result = {};

    if (action === 'getClassTitle' || action === 'getTeacherPin') {
      result = {
        success: true,
        className: getClassTitle(tenantId),
        teacherPin: getTeacherPin(tenantId)
      };
    } else if (action === 'parentLogin') {
      result = parentLogin(data.seatNo, data.password, tenantId);
    } else if (action === 'getParentViewData') {
      result = getParentViewData(data.studentNo, tenantId);
    } else if (action === 'submitParentReport') {
      result = submitParentReport(data.seatNo, data.note, data.absent, tenantId);
    } else if (action === 'clearParentNotification') {
      result = clearParentNotification(data.seatNo, tenantId);
    } else if (action === 'deleteParentRecord') {
      result = deleteParentRecord(data.recordId, tenantId);
    } else if (action === 'toggleCompleteRecord') {
      result = toggleCompleteRecord(data.recordId, data.completed, tenantId);
    } else if (action === 'getTodayTasks' || action === 'getTodayReminders') {
      result = getTodayTasks(data, tenantId);
    } else if (action === 'submitMissingText') {
      result = saveNewMissingRecords(data.targetItem, data.studentNos, data.todayStr, "打字輸入", tenantId);
    } else if (action === 'processHandwritingImage') {
      result = processHandwritingImage(data.base64Image, data.todayStr, tenantId);
    } else if (action === 'getNextWeekReminders') {
      result = getNextWeekReminders(data, tenantId);
    } else if (action === 'uploadStudentPhoto') {
      result = uploadAndRecognizeStudentPhoto(data, tenantId);
    } else if (action === 'getStudentPhotos') {
      result = getStudentPhotos(tenantId);
    } else if (action === 'searchLatestGeminiModel') {
      result = { success: true, model: getLatestFlashModelName() };
    } else if (action === 'getStudentData') {
      result = getStudentData(tenantId);
    } else if (action === 'loadSeatingState') {
      result = { success: true, data: loadSeatingState(tenantId) };
    } else if (action === 'saveSeatingState') {
      result = saveSeatingState(data.state || data, tenantId);
    } else if (action === 'exportSeatingToSheet') {
      result = exportSeatingToSheet(data.exportData || data, tenantId);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================================
// 🤖 GEMINI 模型探索與 API 調用
// =========================================================================

function getLatestFlashModelName(customKey) {
  const defaultModel = "gemini-2.5-flash";
  const apiKey = customKey || getGeminiApiKey();
  if (!apiKey) return defaultModel;
  const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  
  try {
    const response = UrlFetchApp.fetch(listUrl, { "muteHttpExceptions": true });
    const result = JSON.parse(response.getContentText());
    if (result.models && Array.isArray(result.models)) {
      const flashModels = result.models
        .filter(m => {
          const name = (m.name || '').toLowerCase();
          const hasGen = m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent");
          return hasGen && name.includes("flash") && !name.includes("omni") && !name.includes("exp");
        })
        .map(m => m.name.replace("models/", ""));

      if (flashModels.length > 0) {
        flashModels.sort((a, b) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }));
        return flashModels[0];
      }
    }
  } catch (err) {
    Logger.log("⚠️ 取得模型清單失敗: " + err.toString());
  }
  return defaultModel;
}

// =========================================================================
// 📊 多租戶核心業務邏輯
// =========================================================================

function getClassTitle(tenantId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tSheet = ss.getSheetByName("Demo_Tenants");
    if (!tSheet) return `${tenantId} 班`;
    const data = tSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === tenantId) {
        return data[i][3] || `${tenantId} 班`;
      }
    }
  } catch (err) {}
  return `${tenantId} 班`;
}

function getTeacherPin(tenantId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tSheet = ss.getSheetByName("Demo_Tenants");
    if (!tSheet) return "8888";
    const data = tSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === tenantId) {
        return String(data[i][4] || "8888").trim();
      }
    }
  } catch (err) {}
  return "8888";
}

function getStudentData(tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Students');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const students = [];

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === tenantId) {
      students.push({
        id: rows[i][1],
        name: rows[i][2],
        password: rows[i][3],
        missing: rows[i][4] || '',
        photoUrl: rows[i][5] || ''
      });
    }
  }
  return students;
}

function getTodayTasks(data, tenantId) {
  const targetDate = data.todayStr || todayDate;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recSheet = ss.getSheetByName("All_Records");
  const misSheet = ss.getSheetByName("Missing_Items");

  const regular = [];
  if (recSheet && recSheet.getLastRow() > 1) {
    const rData = recSheet.getDataRange().getValues();
    for (let i = 1; i < rData.length; i++) {
      if (String(rData[i][0]).trim() !== tenantId) continue;
      let exDate = rData[i][5];
      if (exDate instanceof Date) exDate = Utilities.formatDate(exDate, "GMT+8", "yyyy-MM-dd");
      else exDate = String(exDate).trim();
      if (exDate === targetDate) {
        regular.push({
          id: rData[i][1],
          type: rData[i][3],
          content: rData[i][4],
          executionDate: exDate
        });
      }
    }
  }

  const missing = [];
  if (misSheet && misSheet.getLastRow() > 1) {
    const mData = misSheet.getDataRange().getValues();
    for (let i = 1; i < mData.length; i++) {
      if (String(mData[i][0]).trim() !== tenantId) continue;
      if (String(mData[i][5]).trim() === "未繳") {
        missing.push({
          id: mData[i][1],
          date: mData[i][2],
          seatNo: mData[i][3],
          targetItem: mData[i][4],
          status: mData[i][5]
        });
      }
    }
  }

  return { success: true, regular: regular, missing: missing };
}

function saveRegularRecords(items, tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("All_Records");
  if (!sheet) return 0;

  let addedCount = 0;
  const nowTime = new Date();

  items.forEach(item => {
    const newId = "REC-" + Utilities.formatDate(nowTime, "GMT+8", "yyyyMMddHHmmss") + "-" + Math.floor(Math.random() * 1000);
    sheet.appendRow([
      tenantId,
      newId,
      Utilities.formatDate(nowTime, "GMT+8", "yyyy-MM-dd HH:mm:ss"),
      item.type,
      item.range || item.content,
      item.executionDate,
      "Demo 語音/手寫"
    ]);
    addedCount++;
  });
  return addedCount;
}

function saveNewMissingRecords(targetItem, studentNos, todayDateStr, source, tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Missing_Items");
  if (!sheet) return { success: false, count: 0 };

  let added = 0;
  studentNos.forEach(seatNo => {
    const newId = "MIS-" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMddHHmmss") + "-" + Math.floor(Math.random() * 1000);
    sheet.appendRow([
      tenantId,
      newId,
      todayDateStr,
      String(seatNo).trim(),
      targetItem,
      "未繳",
      ""
    ]);
    added++;
  });
  return { success: true, count: added };
}

function processHandwritingImage(base64Image, todayDateStr, tenantId) {
  const pureBase64 = base64Image.replace(/^data:image\/(png|jpeg);base64,/, "");
  const apiKey = getGeminiApiKey();
  const modelName = getLatestFlashModelName(apiKey);

  const prompt = `今天日期是 ${todayDateStr}，預設隔日日期為 ${tomorrowDate}。
請辨識圖片中的手寫文字內容，輸出 JSON：
{
  "regularItems": [{"type":"作業","range":"國語第1課","executionDate":"${tomorrowDate}"}],
  "missingItems": [{"targetItem":"理講","studentNos":["1","5","10"]}]
}`;

  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        { "inline_data": { "mime_type": "image/png", "data": pureBase64 } }
      ]
    }]
  };

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
  try {
    const response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
    let rawText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim().replace(/```json/g, "").replace(/```/g, "").trim();
    const parsedData = JSON.parse(rawText);

    let regAdded = 0, misAdded = 0;
    if (parsedData.regularItems && parsedData.regularItems.length > 0) {
      regAdded = saveRegularRecords(parsedData.regularItems, tenantId);
    }
    if (parsedData.missingItems && parsedData.missingItems.length > 0) {
      parsedData.missingItems.forEach(mi => {
        if (mi.targetItem && mi.studentNos) {
          const res = saveNewMissingRecords(mi.targetItem, mi.studentNos, todayDateStr, "Demo 手寫", tenantId);
          misAdded += res.count;
        }
      });
    }
    return { success: true, regularAdded: regAdded, missingAdded: misAdded };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function loadSeatingState(tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Seating_State");
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === tenantId) {
      try {
        return JSON.parse(data[i][1]);
      } catch (e) {
        return null;
      }
    }
  }
  return null;
}

function saveSeatingState(state, tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Seating_State");
  if (!sheet) return { success: false, error: "查無 Seating_State 分頁" };

  const data = sheet.getDataRange().getValues();
  const jsonStr = JSON.stringify(state);
  const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === tenantId) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[jsonStr, nowStr]]);
      return { success: true };
    }
  }

  sheet.appendRow([tenantId, jsonStr, nowStr]);
  return { success: true };
}

function getStudentPhotos(tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Seating_Photos");
  const map = {};
  if (!sheet || sheet.getLastRow() <= 1) return { success: true, photos: map };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === tenantId) {
      const seatNo = String(data[i][1]).trim();
      const url = String(data[i][4] || '').trim();
      if (seatNo && url) map[seatNo] = url;
    }
  }
  return { success: true, photos: map };
}

function uploadAndRecognizeStudentPhoto(data, tenantId) {
  try {
    const base64Image = data.base64Image || data.base64;
    if (!base64Image) return { success: false, error: "未接收到照片！" };
    const cleanBase64 = base64Image.replace(/^data:image\/(png|jpeg|jpg|webp);base64,/, "");
    const mimeType = data.mimeType || "image/jpeg";
    let seatNo = data.manualSeatNo ? String(data.manualSeatNo).trim() : null;
    let studentName = data.manualName ? String(data.manualName).trim() : "";

    if (!seatNo) {
      const apiKey = getGeminiApiKey();
      const modelUsed = getLatestFlashModelName(apiKey);
      const prompt = `辨識白紙上的【座號】與【姓名】，純JSON：{"seat_number": 數字或 null, "name": "姓名"}`;
      const payload = {
        "contents": [{
          "parts": [
            { "text": prompt },
            { "inline_data": { "mime_type": "image/jpeg", "data": cleanBase64 } }
          ]
        }]
      };
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelUsed}:generateContent?key=${apiKey}`;
      const response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
      const resObj = JSON.parse(response.getContentText());
      let rawText = resObj.candidates[0].content.parts[0].text.trim().replace(/```json/g, "").replace(/```/g, "").trim();
      const parsed = JSON.parse(rawText);
      if (parsed.seat_number) {
        seatNo = String(parseInt(parsed.seat_number));
        studentName = parsed.name || "";
      } else {
        return { success: false, error: "未能辨識座號數字，請手動指定。" };
      }
    }

    // 存入 Google Drive
    const folderName = "大屏班務_Demo照片";
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, `demo_${tenantId}_${seatNo}_${Date.now()}.jpg`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const photoUrl = `https://lh3.googleusercontent.com/d/${file.getId()}`;

    // 更新 Seating_Photos
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName("Seating_Photos");
    pSheet.appendRow([tenantId, seatNo, studentName, "", photoUrl, "已辨識", Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss")]);

    return { success: true, seatNumber: seatNo, name: studentName, photoUrl: photoUrl };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function parentLogin(seatNo, password, tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) return { success: false, error: "查無名冊" };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === tenantId && String(data[i][1]).trim() === String(seatNo).trim()) {
      if (String(data[i][3]).trim() === String(password).trim()) {
        return { success: true, studentName: data[i][2], seatNo: seatNo };
      } else {
        return { success: false, error: "密碼錯誤！" };
      }
    }
  }
  return { success: false, error: "查無此座號！" };
}

function getParentViewData(studentNo, tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet = ss.getSheetByName("Missing_Items");
  const rSheet = ss.getSheetByName("All_Records");

  const missingList = [];
  if (mSheet) {
    const mData = mSheet.getDataRange().getValues();
    for (let i = 1; i < mData.length; i++) {
      if (String(mData[i][0]).trim() === tenantId && String(mData[i][3]).trim() === String(studentNo).trim() && mData[i][5] === "未繳") {
        missingList.push({ id: mData[i][1], item: mData[i][4], date: mData[i][2] });
      }
    }
  }

  const tasks = [];
  if (rSheet) {
    const rData = rSheet.getDataRange().getValues();
    for (let i = 1; i < rData.length; i++) {
      if (String(rData[i][0]).trim() === tenantId) {
        tasks.push({ id: rData[i][1], type: rData[i][3], content: rData[i][4], executionDate: rData[i][5] });
      }
    }
  }

  return { success: true, missingList: missingList, tasks: tasks };
}

// =========================================================================
// ⏰ 7 天自動清理機制 (排程觸發器)
// =========================================================================

/**
 * 請在 GAS 專案觸發條件中新增：時間驅動 ➔ 每日執行 (例如凌晨 3:00~4:00)
 */
function autoCleanupExpiredDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tSheet = ss.getSheetByName("Demo_Tenants");
  if (!tSheet || tSheet.getLastRow() <= 1) return;

  const nowMs = Date.now();
  const expireThresholdMs = DEMO_EXPIRE_DAYS * 24 * 60 * 60 * 1000;
  const tData = tSheet.getDataRange().getValues();
  const expiredTenantSet = new Set();

  for (let i = 1; i < tData.length; i++) {
    const tenantId = String(tData[i][0]).trim();
    const createdTimeStr = tData[i][1];
    const createdMs = Date.parse(createdTimeStr);
    if (!isNaN(createdMs) && (nowMs - createdMs > expireThresholdMs)) {
      expiredTenantSet.add(tenantId);
    }
  }

  if (expiredTenantSet.size === 0) {
    Logger.log("✅ 無過期之 Demo 租戶資料。");
    return;
  }

  Logger.log(`🧹 開始清理過期租戶: ${Array.from(expiredTenantSet).join(", ")}`);

  const sheetsToClean = ["Students", "All_Records", "Missing_Items", "Seating_Photos", "Seating_State", "Demo_Tenants"];
  sheetsToClean.forEach(sName => {
    const sheet = ss.getSheetByName(sName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const sData = sheet.getDataRange().getValues();
    for (let r = sData.length - 1; r >= 1; r--) {
      const rowTenant = String(sData[r][0]).trim();
      if (expiredTenantSet.has(rowTenant)) {
        sheet.deleteRow(r + 1);
      }
    }
  });

  Logger.log("🎉 過期 Demo 租戶資料清理完畢！");
}
