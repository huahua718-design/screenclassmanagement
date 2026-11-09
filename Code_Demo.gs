/**
 * =========================================================================
 * 🏫 智慧大屏班務系統 - 【公共 Demo 展示與 7 天試用專用版】
 * =========================================================================
 * ⚠️ 注意事項：
 * 1. 本檔案專門部署在展示用、體驗用的公共 GAS 專案上 (試算表 B)。
 * 2. 完整支援大屏前端所有 API：今日任務、提醒事項、缺交銷案看板、座位表、家長登入。
 * 3. 支援「使用者名稱 / 班級代碼 (Tenant ID)」多租戶隔離，多位老師同時試用互不干擾。
 * 4. 首次連線自動播種 10 名學生與範例作業，開箱即有完整豐富體驗。
 * 5. 內建 7 天過期資料定時自動清理程序 (autoCleanupExpiredDemoData)。
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

function normalizeDateStr(dateVal) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, "GMT+8", "yyyy-MM-dd");
  }
  let s = String(dateVal).trim();
  let m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (m) {
    return `${m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3]).padStart(2, '0')}`;
  }
  return s;
}

// =========================================================================
// 🏢 多租戶資料庫初始化與管理模組
// =========================================================================

function getTenantId(data, e) {
  let tid = '';
  if (data && data.tenantId) tid = String(data.tenantId).trim();
  else if (e && e.parameter && e.parameter.tenantId) tid = String(e.parameter.tenantId).trim();
  else if (e && e.parameter && e.parameter.tenant) tid = String(e.parameter.tenant).trim();
  return tid ? tid : 'demo';
}

function ensureDemoSheetsExist(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetDefs = [
    { name: "Demo_Tenants", headers: ["Tenant_ID", "建立時間", "最後活動時間", "班級名稱", "PIN碼", "座位表JSON"] },
    { name: "Students", headers: ["Tenant_ID", "座號", "姓名", "密碼", "未繳項目", "照片網址", "辨識狀態"] },
    { name: "All_Records", headers: ["Tenant_ID", "紀錄ID", "登記時間", "類別", "內容範圍", "執行日期", "來源"] },
    { name: "Missing_Items", headers: ["Tenant_ID", "紀錄ID", "登記日期", "學生座號", "項目名稱", "狀態", "銷案時間"] },
    { name: "Seating_Photos", headers: ["Tenant_ID", "座號", "姓名", "座位編號", "照片網址", "辨識狀態", "更新時間"] },
    { name: "Settings", headers: ["設定項目", "設定值", "說明與備註"] }
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
    tSheet.getRange(foundRow, 3).setValue(nowIso);
    return;
  }

  // 新租戶：註冊租戶
  tSheet.appendRow([tenantId, nowIso, nowIso, `${tenantId} 班`, "8888", ""]);

  // 1. 播種 10 名預設學生
  const sSheet = ss.getSheetByName("Students");
  const defaultStudents = [
    [tenantId, "1", "王小明", "1234", "數學習作第 5 頁", "", "正常"],
    [tenantId, "2", "李大華", "1234", "", "", "正常"],
    [tenantId, "3", "張美美", "1234", "", "", "正常"],
    [tenantId, "4", "陳志豪", "1234", "", "", "正常"],
    [tenantId, "5", "林雅婷", "1234", "數學習作第 5 頁", "", "正常"],
    [tenantId, "6", "黃俊傑", "1234", "", "", "正常"],
    [tenantId, "7", "吳淑芬", "1234", "", "", "正常"],
    [tenantId, "8", "蔡宗憲", "1234", "", "", "正常"],
    [tenantId, "9", "許瑋甯", "1234", "", "", "正常"],
    [tenantId, "10", "鄭家豪", "1234", "", "", "正常"]
  ];
  defaultStudents.forEach(row => sSheet.appendRow(row));

  // 2. 播種今日任務與明日提醒
  const rSheet = ss.getSheetByName("All_Records");
  const timeNowStr = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss");
  // 今日任務
  rSheet.appendRow([tenantId, "REC-TODAY-1", timeNowStr, "作業", "國語第 1 課生字語詞", todayDate, "系統預設"]);
  rSheet.appendRow([tenantId, "REC-TODAY-2", timeNowStr, "考試", "數學 1-1 小考", todayDate, "系統預設"]);
  // 明日提醒
  rSheet.appendRow([tenantId, "REC-TOMO-1", timeNowStr, "作業", "數學習作第 6-7 頁", tomorrowDate, "系統預設"]);
  rSheet.appendRow([tenantId, "REC-TOMO-2", timeNowStr, "提醒事項", "帶彩色筆、美勞剪刀", tomorrowDate, "系統預設"]);

  // 3. 播種缺交作業
  const mSheet = ss.getSheetByName("Missing_Items");
  mSheet.appendRow([tenantId, "MIS-INIT-1", todayDate, "1", "數學習作第 5 頁", "未繳", ""]);
  mSheet.appendRow([tenantId, "MIS-INIT-2", todayDate, "5", "數學習作第 5 頁", "未繳", ""]);
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
    } else if (action === 'processAudioInput') {
      result = processAudioInput(data.base64Audio, data.mimeType, data.todayStr, tenantId);
    } else if (action === 'processMissingAudio') {
      result = processMissingAudio(data.base64Audio, data.mimeType, data.todayStr, data.targetItem, tenantId);
    } else if (action === 'getTomorrowReminders' || action === 'getThisWeekReminders') {
      result = getThisWeekReminders(data, tenantId);
    } else if (action === 'getNextWeekReminders') {
      result = getNextWeekReminders(data, tenantId);
    } else if (action === 'getUnsettledMissingItems') {
      result = getUnsettledMissingItems(tenantId);
    } else if (action === 'markSingleStudentCompleted') {
      result = markSingleStudentCompleted(data.studentNo, data.itemTitle, tenantId);
    } else if (action === 'getRecentAssignments') {
      result = getRecentAssignments(data.todayStr, tenantId);
    } else if (action === 'validateStudentLogin') {
      result = validateStudentLogin(data.seatNo, data.pwd, tenantId);
    } else if (action === 'getStudentMissingItems') {
      result = getStudentMissingItems(data.studentNo, tenantId);
    } else if (action === 'getTodayTasks' || action === 'getTodayReminders') {
      result = getTodayTasks(data, tenantId);
    } else if (action === 'submitMissingText') {
      result = saveNewMissingRecords(data.targetItem, data.studentNos, data.todayStr, "打字輸入", tenantId);
    } else if (action === 'processHandwritingImage') {
      result = processHandwritingImage(data.base64Image, data.todayStr, tenantId);
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
// 🤖 GEMINI 模型探索
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
  const className = getClassTitle(tenantId);
  if (!sheet) return { className: className, students: [] };

  const rows = sheet.getDataRange().getValues();
  const photoMap = getStudentPhotoMap(ss, tenantId);
  const students = [];

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === tenantId && rows[i][1] !== '') {
      const seatNo = String(rows[i][1]).trim();
      students.push({
        id: seatNo,
        name: String(rows[i][2] || '').trim(),
        password: String(rows[i][3] || '1234').trim(),
        missing: String(rows[i][4] || '').trim(),
        photoUrl: photoMap[seatNo] || String(rows[i][5] || '').trim()
      });
    }
  }
  return { className: className, students: students };
}

function getTodayTasks(payload, tenantId) {
  let targetDateStr = "";
  if (typeof payload === 'object' && payload !== null) {
    targetDateStr = payload.todayStr || payload.targetDateStr || "";
  } else {
    targetDateStr = String(payload || "");
  }
  targetDateStr = normalizeDateStr(targetDateStr);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("All_Records");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const tasks = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== tenantId) continue;
    let rowDate = normalizeDateStr(data[i][5]);
    if (rowDate === targetDateStr && rowDate !== "") {
      tasks.push({
        type: data[i][3] || "任務",
        range: data[i][4] || "",
        date: rowDate,
        source: data[i][6] || ""
      });
    }
  }
  return tasks;
}

function getThisWeekReminders(payload, tenantId) {
  const tomorrowStr = normalizeDateStr(payload.tomorrowStr);
  const saturdayStr = normalizeDateStr(payload.saturdayStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("All_Records");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== tenantId) continue;
    let exDate = normalizeDateStr(data[i][5]);
    if (exDate && exDate >= tomorrowStr && exDate <= saturdayStr) {
      list.push({
        id: data[i][1],
        type: data[i][3],
        range: data[i][4],
        date: exDate,
        source: data[i][6]
      });
    }
  }
  return list;
}

function getNextWeekReminders(payload, tenantId) {
  const startDateStr = normalizeDateStr(payload.startDateStr);
  const endDateStr = normalizeDateStr(payload.endDateStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("All_Records");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== tenantId) continue;
    let exDate = normalizeDateStr(data[i][5]);
    if (exDate && exDate >= startDateStr && exDate <= endDateStr) {
      list.push({
        id: data[i][1],
        type: data[i][3],
        range: data[i][4],
        date: exDate,
        source: data[i][6]
      });
    }
  }
  return list;
}

function getUnsettledMissingItems(tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Missing_Items");
  const grouped = {};
  if (!sheet) return grouped;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== tenantId) continue;
    const status = String(data[i][5]).trim();
    if (status === "未繳") {
      const itemTitle = String(data[i][4]).trim();
      const studentNo = String(data[i][3]).padStart(2, '0');
      if (!grouped[itemTitle]) grouped[itemTitle] = [];
      if (!grouped[itemTitle].includes(studentNo)) grouped[itemTitle].push(studentNo);
    }
  }

  // 排序座號
  for (let k in grouped) {
    grouped[k].sort((a, b) => parseInt(a, 10) - parseInt(b, 10));
  }
  return grouped;
}

function markSingleStudentCompleted(studentNo, itemTitle, tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Missing_Items");
  if (!sheet) return { success: false, error: "查無 Missing_Items" };

  const data = sheet.getDataRange().getValues();
  const targetNo = String(studentNo).replace(/^0+/, "");
  const targetItem = String(itemTitle).trim();
  const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
  let clearedCount = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== tenantId) continue;
    const rowNo = String(data[i][3]).replace(/^0+/, "");
    const rowItem = String(data[i][4]).trim();
    const rowStatus = String(data[i][5]).trim();

    if (rowNo === targetNo && rowItem === targetItem && rowStatus === "未繳") {
      sheet.getRange(i + 1, 6).setValue("已完成");
      sheet.getRange(i + 1, 7).setValue(nowStr);
      clearedCount++;
    }
  }
  updateStudentsMissingSummary(ss, tenantId);
  return { success: true, count: clearedCount };
}

function updateStudentsMissingSummary(ss, tenantId) {
  const sSheet = ss.getSheetByName("Students");
  const mSheet = ss.getSheetByName("Missing_Items");
  if (!sSheet || !mSheet) return;

  const mData = mSheet.getDataRange().getValues();
  const pendingByStudent = {};

  for (let i = 1; i < mData.length; i++) {
    if (String(mData[i][0]).trim() !== tenantId) continue;
    if (String(mData[i][5]).trim() === "未繳") {
      const seatNo = String(mData[i][3]).padStart(2, '0');
      const item = String(mData[i][4]).trim();
      if (!pendingByStudent[seatNo]) pendingByStudent[seatNo] = [];
      if (!pendingByStudent[seatNo].includes(item)) pendingByStudent[seatNo].push(item);
    }
  }

  const sData = sSheet.getDataRange().getValues();
  for (let i = 1; i < sData.length; i++) {
    if (String(sData[i][0]).trim() !== tenantId) continue;
    const seatNo = String(sData[i][1]).padStart(2, '0');
    const itemsList = pendingByStudent[seatNo] ? pendingByStudent[seatNo].join(", ") : "";
    sSheet.getRange(i + 1, 5).setValue(itemsList);
  }
}

function getRecentAssignments(todayDateStr, tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rSheet = ss.getSheetByName("All_Records");
  const assignments = [];
  if (!rSheet) return assignments;

  const data = rSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== tenantId) continue;
    const type = String(data[i][3]).trim();
    const content = String(data[i][4]).trim();
    if (type === "作業" && content && !assignments.includes(content)) {
      assignments.push(content);
    }
  }
  return assignments;
}

function validateStudentLogin(seatNo, pwd, tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) return { success: false, message: "查無名冊！" };

  const data = sheet.getDataRange().getValues();
  const targetSeat = String(seatNo).trim();
  const targetPwd = String(pwd).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== tenantId) continue;
    if (String(data[i][1]).trim() === targetSeat) {
      const realPwd = String(data[i][3]).trim();
      if (realPwd === targetPwd || realPwd === "") {
        return { success: true, studentName: String(data[i][2]).trim() };
      } else {
        return { success: false, message: "密碼錯誤！", error: "密碼錯誤！" };
      }
    }
  }
  return { success: false, message: "查無此座號！", error: "查無此座號！" };
}

function getStudentMissingItems(studentNo, tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Missing_Items");
  const result = [];
  if (!sheet) return result;

  const data = sheet.getDataRange().getValues();
  const targetNo = String(studentNo).trim().padStart(2, '0');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== tenantId) continue;
    const rowNo = String(data[i][3]).trim().padStart(2, '0');
    const status = String(data[i][5]).trim();
    if (rowNo === targetNo && status === "未繳") {
      result.push(`${data[i][4]} (${normalizeDateStr(data[i][2])})`);
    }
  }
  return result;
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
      normalizeDateStr(item.executionDate),
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
    const cleanNo = String(seatNo).trim().padStart(2, '0');
    const newId = "MIS-" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMddHHmmss") + "-" + Math.floor(Math.random() * 1000);
    sheet.appendRow([
      tenantId,
      newId,
      normalizeDateStr(todayDateStr),
      cleanNo,
      targetItem,
      "未繳",
      ""
    ]);
    added++;
  });
  updateStudentsMissingSummary(ss, tenantId);
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

function processAudioInput(base64Audio, mimeType, todayDateStr, tenantId) {
  const pureBase64 = base64Audio.includes(',') ? base64Audio.split(',')[1] : base64Audio;
  let cleanMime = mimeType ? mimeType.split(';')[0].trim() : 'audio/mp4';
  if (cleanMime === 'audio/m4a') cleanMime = 'audio/mp4';

  const prompt = `今天日期是 ${todayDateStr}，預設隔日日期為 ${tomorrowDate}。
請聽取語音內容並提取提醒事項，輸出 JSON 陣列：
[{"type":"考試","range":"數學 3-3","executionDate":"${tomorrowDate}"}]`;

  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        { "inline_data": { "mime_type": cleanMime, "data": pureBase64 } }
      ]
    }]
  };

  const apiKey = getGeminiApiKey();
  const modelName = getLatestFlashModelName(apiKey);
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

  try {
    const response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
    let rawText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim().replace(/```json/g, "").replace(/```/g, "").trim();
    const parsedData = JSON.parse(rawText);
    const addedCount = saveRegularRecords(parsedData, tenantId);
    return { success: true, added: addedCount };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function processMissingAudio(base64Audio, mimeType, todayDateStr, targetItem, tenantId) {
  const pureBase64 = base64Audio.includes(',') ? base64Audio.split(',')[1] : base64Audio;
  let cleanMime = mimeType ? mimeType.split(';')[0].trim() : 'audio/mp4';
  if (cleanMime === 'audio/m4a') cleanMime = 'audio/mp4';

  const prompt = `聽取語音中缺交學生的座號數字，回傳純整數數字陣列：["1","5","12"]`;
  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        { "inline_data": { "mime_type": cleanMime, "data": pureBase64 } }
      ]
    }]
  };

  const apiKey = getGeminiApiKey();
  const modelName = getLatestFlashModelName(apiKey);
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

  try {
    const response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
    let rawText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim().replace(/```json/g, "").replace(/```/g, "").trim();
    const studentNos = JSON.parse(rawText);
    const res = saveNewMissingRecords(targetItem, studentNos, todayDateStr, "Demo 語音", tenantId);
    return { success: true, added: res.count };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function loadSeatingState(tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tSheet = ss.getSheetByName("Demo_Tenants");
  if (!tSheet) return null;
  const data = tSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === tenantId) {
      const jsonStr = data[i][5];
      if (jsonStr) {
        try { return JSON.parse(jsonStr); } catch (e) {}
      }
      break;
    }
  }
  return null;
}

function saveSeatingState(state, tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tSheet = ss.getSheetByName("Demo_Tenants");
  if (!tSheet) return { success: false, error: "查無 Demo_Tenants" };

  const data = tSheet.getDataRange().getValues();
  const jsonStr = JSON.stringify(state);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === tenantId) {
      tSheet.getRange(i + 1, 6).setValue(jsonStr);
      return { success: true };
    }
  }
  return { success: false, error: "查無此租戶" };
}

function getStudentPhotoMap(ss, tenantId) {
  const map = {};
  const sheet = ss.getSheetByName("Seating_Photos");
  if (!sheet || sheet.getLastRow() <= 1) return map;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === tenantId) {
      const seatNo = String(data[i][1]).trim();
      const url = String(data[i][4] || '').trim();
      if (seatNo && url) map[seatNo] = url;
    }
  }
  return map;
}

function getStudentPhotos(tenantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return { success: true, photos: getStudentPhotoMap(ss, tenantId) };
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

function exportSeatingToSheet(exportData, tenantId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `${tenantId}_座位表`;
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) ss.deleteSheet(sheet);
    sheet = ss.insertSheet(sheetName);

    sheet.getRange(1, 1).setValue(`${getClassTitle(tenantId)} - 座位表 (導出時間: ${todayDate})`)
         .setFontWeight("bold").setFontSize(14);
    
    return { success: true, sheetName: sheetName };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =========================================================================
// ⏰ 7 天自動清理機制 (排程觸發器)
// =========================================================================

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
    if (tenantId === "demo") continue; // 保護預設 demo 不被清除
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

  const sheetsToClean = ["Students", "All_Records", "Missing_Items", "Seating_Photos", "Demo_Tenants"];
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
