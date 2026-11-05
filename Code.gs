// 請填入您的 Gemini API Key
const GEMINI_API_KEY = "請填入您的 Gemini API Key";
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// 計算當日與隔日字串 (以 GMT+8 台灣時間為準)
const now = new Date();
const todayDate = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd");

const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
const tomorrowDate = Utilities.formatDate(tomorrow, "GMT+8", "yyyy-MM-dd");

function doGet(e) {
  // 1. 處理前端 API 請求 (例如 ?action=getClassTitle)
  const action = (e && e.parameter) ? e.parameter.action : null;
  
  if (action === 'getClassTitle' || action === 'getTeacherPin') {
    const className = getClassTitle();
    const teacherPin = getTeacherPin();
    return ContentService.createTextOutput(JSON.stringify({ 
      success: true,
      className: className,
      teacherPin: teacherPin
    })).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'getStudentData') {
    return ContentService.createTextOutput(JSON.stringify(getStudentData()))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'loadSeatingState') {
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: loadSeatingState() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'getStudentPhotos') {
    return ContentService.createTextOutput(JSON.stringify(getStudentPhotos()))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'none') {
    return ContentService.createTextOutput("")
                         .setMimeType(ContentService.MimeType.JSON);
  }

  // 2. 讀取 Students 分頁 A1 的班級名稱
  let className = '智慧教室';
  try {
    if (typeof getClassTitle === 'function') {
      className = getClassTitle() || '智慧教室';
    }
  } catch (err) {
    console.error("讀取班級名稱失敗：", err);
  }

  // 3. 讀取頁面 (page) 與角色 (role) 網址參數
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase() : '';
  const role = (e && e.parameter && e.parameter.role) ? e.parameter.role : 'screen';

  // 4. 【座位表頁面】若網址帶有 ?page=seating，開啟 Seating.html 樣板
  if (page === 'seating') {
    const template = HtmlService.createTemplateFromFile('Seating');
    
    // 將後端接收到的 origin 與 role 注入樣板
    template.origin = (e && e.parameter && e.parameter.origin) ? e.parameter.origin : '';
    template.role = role;
    template.className = className;
    template.geminiApiKey = (typeof GEMINI_API_KEY !== 'undefined' && GEMINI_API_KEY !== "請填入您的 Gemini API Key") ? GEMINI_API_KEY : '';
    template.mode = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : '';

    return template.evaluate()
      .setTitle(`${className} - 班級座位表`)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 5. 【大屏主頁面】預設載入 Index.html 樣板
  const template = HtmlService.createTemplateFromFile('Index');
  
  // 將角色與班級名稱傳入 HTML 樣板 (前端可透過 <?= className ?> 使用)
  template.role = role;
  template.className = className;

  // 根據班級名稱與角色動態組裝網頁標題
  let pageTitle = `${className}的大屏看板`;
  if (role === 'teacher') pageTitle = `${className}教師控制台`;
  if (role === 'parent') pageTitle = `${className}線上聯絡簿`;

  return template.evaluate()
    .setTitle(pageTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // 🟢 已修正 F 大寫
}
// 💡 通用日期標準化工具 (解決斜線/橫線/時區偏差等問題)
function normalizeDateStr(val) {
  if (!val) return "";
  
  // 若為 Date 物件，強制轉為台灣時區的 yyyy-MM-dd
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Taipei", "yyyy-MM-dd");
  }
  
  var str = String(val).trim();
  if (str.includes("T")) str = str.split("T")[0]; // 處理 ISO 字串
  str = str.replace(/\//g, "-");                  // 斜線轉連字號
  
  // 處理 YYYY-M-D 補零
  var parts = str.split("-");
  if (parts.length === 3) {
    var y = parts[0];
    var m = parts[1].padStart(2, '0');
    var d = parts[2].padStart(2, '0');
    return y + "-" + m + "-" + d;
  }
  return str;
}
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'ping') {
      return ContentService.createTextOutput(JSON.stringify({ status: 'pong' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const action = data.action;
    let result = {};

    if (action === 'getClassTitle' || action === 'getTeacherPin') {
      // 取得 "Students" 分頁的 A1 (班級名稱) 與 C1 (教師管理 PIN 碼)
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Students");
      var title = "智慧教室";
      var pin = "8888";
      
      if (sheet) {
        title = sheet.getRange("A1").getValue() || "智慧教室";
        var valC1 = sheet.getRange("C1").getValue();
        if (valC1 && String(valC1).trim() !== "") {
          pin = String(valC1).trim();
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        className: String(title).trim(),
        teacherPin: String(pin).trim()
      })).setMimeType(ContentService.MimeType.JSON);
      
    } else if (action === 'processAudioInput') {
      result = processAudioInput(data.base64Audio, data.mimeType, data.todayStr);
    } else if (action === 'processMissingAudio') {
      result = processMissingAudio(data.base64Audio, data.mimeType, data.defaultItem, data.todayStr);
    } else if (action === 'getTomorrowReminders' || action === 'getThisWeekReminders') {
      result = getThisWeekReminders(data);
    } else if (action === 'getUnsettledMissingItems') {
      result = getUnsettledMissingItems();
    } else if (action === 'markSingleStudentCompleted') {
      result = markSingleStudentCompleted(data.studentNo, data.itemTitle);
    } else if (action === 'getRecentAssignments') {
      result = getRecentAssignments();
    } else if (action === 'validateStudentLogin') {
      result = validateStudentLogin(data);
    } else if (action === 'getStudentMissingItems') {
      result = getStudentMissingItems(data.studentNo);
    } else if (action === 'getTodayTasks' || action === 'getTodayReminders') {
      result = getTodayTasks(data);
    } else if (action === 'submitMissingText') {
      result = saveNewMissingRecords(data.targetItem, data.studentNos, data.todayStr, "打字輸入");
    } else if (action === 'processHandwritingImage') {
      result = processHandwritingImage(data.base64Image, data.todayStr);
    } else if (action === 'getNextWeekReminders') {
      result = getNextWeekReminders(data);
    } else if (action === 'uploadStudentPhoto') {
      result = uploadAndRecognizeStudentPhoto(data);
    } else if (action === 'getStudentPhotos') {
      result = getStudentPhotos();
    } else if (action === 'searchLatestGeminiModel') {
      result = { success: true, model: searchLatestGeminiVisionModel(data.apiKey) };
    } else if (action === 'getStudentData') {
      result = getStudentData();
    } else if (action === 'loadSeatingState') {
      result = { success: true, data: loadSeatingState() };
    } else if (action === 'saveSeatingState') {
      result = saveSeatingState(data.state || data);
    } else if (action === 'exportSeatingToSheet') {
      result = exportSeatingToSheet(data.exportData || data);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 🤖 動態取得最新 Flash 模型
function getLatestFlashModelName() {
  const defaultModel = "gemini-2.5-flash";
  if (!GEMINI_API_KEY) return defaultModel;
  const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${GEMINI_API_KEY}`;
  
  try {
    const response = UrlFetchApp.fetch(listUrl, { "muteHttpExceptions": true });
    const result = JSON.parse(response.getContentText());
    
    if (result.models && Array.isArray(result.models)) {
      const flashModels = result.models
        .filter(m => {
          const name = m.name.toLowerCase();
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
    Logger.log("⚠️ 取得模型清單失敗，使用預設值: " + err.toString());
  }
  return defaultModel;
}

// 🔄 自動同步 Students 頁面的「未繳項目」統計 (第 D 欄)
function updateStudentsSheetSummary() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const studentSheet = ss.getSheetByName("Students");
  const missingSheet = ss.getSheetByName("Missing_Items");
  
  if (!studentSheet || !missingSheet) return;

  const missingData = missingSheet.getDataRange().getValues();
  const pendingByStudent = {};

  for (let i = 1; i < missingData.length; i++) {
    if (missingData[i][4] === "未繳") {
      const studentNo = String(missingData[i][2]).padStart(2, '0');
      const item = missingData[i][3];
      if (!pendingByStudent[studentNo]) pendingByStudent[studentNo] = [];
      if (!pendingByStudent[studentNo].includes(item)) pendingByStudent[studentNo].push(item);
    }
  }

  const studentData = studentSheet.getDataRange().getValues();
  for (let i = 1; i < studentData.length; i++) {
    const seatNo = String(studentData[i][0]).padStart(2, '0');
    const itemsList = pendingByStudent[seatNo] ? pendingByStudent[seatNo].join(", ") : "";
    studentSheet.getRange(i + 1, 4).setValue(itemsList);
  }
}

/**
 * 家長登入驗證 (已修正欄位對應：A座號 / B姓名 / C密碼)
 */
function validateStudentLogin(data) {
  // 1. 取得前端傳入的座號與密碼
  var rawSeat = (typeof data === 'object' && data !== null) ? data.seatNo : data;
  var rawPwd = (typeof data === 'object' && data !== null) ? data.pwd : arguments[1];

  var inputSeatStr = String(rawSeat || "").trim();
  var inputPwd = normalizePassword(rawPwd);

  if (!inputSeatStr || !inputPwd) {
    return { success: false, message: "座號與密碼不能為空！", error: "座號與密碼不能為空！" };
  }

  // 將輸入座號轉為整數數字 (避免 3 與 03 比對不符的問題)
  var inputSeatNum = parseInt(inputSeatStr, 10);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Students") || ss.getSheetByName("學生資料");

  if (!sheet) {
    return { 
      success: false, 
      message: "找不到『Students』工作表，請檢查分頁名稱！",
      error: "找不到 Students 工作表" 
    };
  }

  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    // 欄位精準定位：A欄(0)座號 / B欄(1)姓名 / C欄(2)密碼
    var rowSeatNum = parseInt(rows[i][0], 10);           // A 欄：座號
    var studentName = String(rows[i][1] || "").trim();    // B 欄：姓名
    var sheetPwd = normalizePassword(rows[i][2]);        // C 欄：密碼 (自動補足4碼)

    // 1. 先比對座號 (轉數字比對，相容 3 與 03)
    if (!isNaN(rowSeatNum) && rowSeatNum === inputSeatNum) {
      // 2. 再比對密碼
      if (sheetPwd === inputPwd) {
        return {
          success: true,
          studentNo: String(rowSeatNum).padStart(2, '0'), // 回傳標準化兩位數座號 (如 "03")
          studentName: studentName
        };
      } else {
        return { success: false, message: "密碼錯誤！", error: "密碼錯誤！" };
      }
    }
  }

  return { success: false, message: "查無此座號！", error: "查無此座號！" };
}
// 2. 大屏手寫辨識 (僅保留一般提醒與缺交登記)
function processHandwritingImage(base64Image, todayDate) {
  const pureBase64 = base64Image.replace(/^data:image\/(png|jpeg);base64,/, "");
  
  const prompt = `
  今天日期是 ${todayDate}，預設隔日日期為 ${tomorrowDate}。
  請辨識圖片中的手寫文字內容，並分為兩類輸出 JSON：

  1. "regularItems": 一般課業提醒或考試項目。
     - "type": "考試"、"作業" 或 "提醒事項"。
     - "range": 內容描述。
     - "executionDate": 執行日期（格式 YYYY-MM-DD）。
       【關鍵日期規則】：若手寫內容中未明確提及具體日期，預設執行日期統一填入隔日 ${tomorrowDate}。

  2. "missingItems": 若文字含關鍵字 "缺" 或 "未" (例如 "理講缺1,5,10" 或 "數習未交 3,7")，代表缺交登記。
     - "targetItem": 作業名稱 (例如 "理講")。
     - "studentNos": 座號數字陣列 (例如 ["1", "5", "10"])。

  僅回傳純 JSON 物件：
  {
    "regularItems": [{"type":"作業","range":"國語第1課","executionDate":"${tomorrowDate}"}],
    "missingItems": [{"targetItem":"理講","studentNos":["1","5","10"]}]
  }
`;

  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        { "inline_data": { "mime_type": "image/png", "data": pureBase64 } }
      ]
    }]
  };

  const modelName = getLatestFlashModelName();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${GEMINI_API_KEY}`;

  try {
    const response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
    let rawText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.trim().replace(/```json/g, "").replace(/```/g, "").trim();
    const parsedData = JSON.parse(rawText);

    let regAdded = 0, misAdded = 0;

    if (parsedData.regularItems && parsedData.regularItems.length > 0) {
      regAdded = saveRegularRecords(parsedData.regularItems);
    }

    if (parsedData.missingItems && parsedData.missingItems.length > 0) {
      parsedData.missingItems.forEach(mi => {
        if (mi.targetItem && mi.studentNos) {
          const res = saveNewMissingRecords(mi.targetItem, mi.studentNos, todayDate);
          misAdded += res.count;
        }
      });
    }

    return { success: true, regularAdded: regAdded, missingAdded: misAdded };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// 內部：寫入一般紀錄
function saveRegularRecords(items) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("All_Records");
  
  if (!sheet) return 0;

  const existingData = sheet.getDataRange().getValues();
  let addedCount = 0;
  const now = new Date();

  items.forEach(item => {
    // 檢查 All_Records 中是否已有相同「類別 + 內容 + 執行日期」
    const isDuplicate = existingData.some(row => {
      const rowType = String(row[2]).trim();
      const rowRange = String(row[3]).trim();
      let exDate = row[4];
      if (exDate instanceof Date) {
        exDate = Utilities.formatDate(exDate, "GMT+8", "yyyy-MM-dd");
      } else {
        exDate = String(exDate).trim();
      }
      return rowType === item.type && rowRange === item.range.trim() && exDate === item.executionDate;
    });

    if (!isDuplicate) {
      const newId = "REC-" + Utilities.formatDate(now, "GMT+8", "yyyyMMddHHmmss") + "-" + Math.floor(Math.random() * 1000);
      sheet.appendRow([
        newId, 
        Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss"), 
        item.type, 
        item.range, 
        item.executionDate, 
        "語音/手寫"
      ]);
      addedCount++;
    }
  });
  return addedCount;
}

// 3. 教師一般語音記事 (修正版)
function processAudioInput(base64Audio, mimeType, todayDate) {
  // 1. 安全的 Base64 前綴清除
  const pureBase64 = base64Audio.includes(',') ? base64Audio.split(',')[1] : base64Audio;

  // 2. 清理 MIME Type
  let cleanMime = mimeType ? mimeType.split(';')[0].trim() : 'audio/mp4';
  if (cleanMime === 'audio/m4a') cleanMime = 'audio/mp4';

  // 3. 計算隔日日期 (用於未指定日期時的預設值)
  const today = todayDate ? new Date(todayDate) : new Date();
  const tomorrow = new Date(today.getTime() + 24 * 60 * 60 * 1000);
  const tomorrowDate = Utilities.formatDate(tomorrow, "GMT+8", "yyyy-MM-dd");

  // 4. 強化 Prompt：要求 range 必須包含科目名稱，且預設日期設為隔日
  const prompt = `今天日期是 ${todayDate}，預設隔日日期為 ${tomorrowDate}。
請聽取語音內容並提取提醒事項，輸出 JSON 陣列格式。

【欄位與規則】：
1. "type": "考試"、"作業" 或 "提醒事項"
2. "range": 科目與完整描述內容（極重要！必須同時包含「科目名稱」與「章節/範圍」，例如「考數學三之三」必須寫成「數學 3-3」或「數學三之三」，絕對不可遺漏科目「數學」）。
3. "executionDate": 執行日期（格式 YYYY-MM-DD）。若語音未明確指定具體日期，預設統一填入隔日 ${tomorrowDate}。

僅回傳純 JSON 陣列：
[{"type":"考試","range":"數學 3-3","executionDate":"${tomorrowDate}"}]`;

  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        { "inline_data": { "mime_type": cleanMime, "data": pureBase64 } }
      ]
    }]
  };

  const modelName = getLatestFlashModelName();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${GEMINI_API_KEY}`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });

    const resObj = JSON.parse(response.getContentText());

    if (resObj.error) {
      return { success: false, error: `Gemini API 拒絕請求 (${resObj.error.code}): ${resObj.error.message}` };
    }
    if (!resObj.candidates || resObj.candidates.length === 0 || !resObj.candidates[0].content) {
      return { success: false, error: "Gemini 無法解析此段語音（請確認是否有錄到聲音或時間太短）" };
    }

    let rawText = resObj.candidates[0].content.parts[0].text.trim().replace(/```json/g, "").replace(/```/g, "").trim();
    const parsedData = JSON.parse(rawText);
    const addedCount = saveRegularRecords(parsedData);
    return { success: true, added: addedCount };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
// 4. 查詢本周剩餘提醒事項 (明日 ~ 本周六)
function getThisWeekReminders(payload) {
  var tomorrowStr = "";
  var saturdayStr = "";

  if (typeof payload === 'object' && payload !== null) {
    tomorrowStr = payload.tomorrowStr || payload.targetDateStr || payload.todayStr || "";
    saturdayStr = payload.saturdayStr || "";
  }

  tomorrowStr = normalizeDateStr(tomorrowStr);
  saturdayStr = normalizeDateStr(saturdayStr);

  // 若前端未傳入區間，自動以「週日為一週第一天」計算明日與本週六
  if (!tomorrowStr || !saturdayStr) {
    var baseDate = new Date();
    var dayOfWeek = baseDate.getDay(); // 0(週日) ~ 6(週六)
    
    var tomorrow = new Date(baseDate.getTime() + 24 * 60 * 60 * 1000);
    tomorrowStr = normalizeDateStr(tomorrow);
    
    var saturday = new Date(baseDate.getTime() + (6 - dayOfWeek) * 24 * 60 * 60 * 1000);
    saturdayStr = normalizeDateStr(saturday);
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("All_Records");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const reminders = [];
  const weekDayNames = ["日", "一", "二", "三", "四", "五", "六"];

  for (let i = 1; i < data.length; i++) {
    let rowDate = normalizeDateStr(data[i][4]); // 第 5 欄：執行日期
    if (!rowDate) continue;

    // 篩選執行日期在 [tomorrowStr, saturdayStr] 區間內的項目
    if (rowDate >= tomorrowStr && rowDate <= saturdayStr) {
      const d = new Date(rowDate.replace(/-/g, '/'));
      const dayName = !isNaN(d.getDay()) ? weekDayNames[d.getDay()] : "";
      const monthDay = !isNaN(d.getTime()) ? `${d.getMonth() + 1}/${d.getDate()}` : rowDate;

      reminders.push({ 
        type: data[i][2] || "提醒",           // 第 3 欄：類別
        range: data[i][3] || "",            // 第 4 欄：內容範圍
        date: rowDate, 
        dateLabel: `${monthDay} (${dayName})`,// 前端顯示用的日期標籤 (例: 9/24 (四))
        source: data[i][5] || ""            // 第 6 欄：來源
      });
    }
  }

  // 依執行日期從小到大排序
  reminders.sort((a, b) => a.date.localeCompare(b.date));
  return reminders;
}

// 建立相容別名，確保舊版呼叫不中斷
function getTomorrowReminders(payload) {
  return getThisWeekReminders(payload);
}

// 📌 修改：取得「執行日期」在 3 天內的作業與提醒項目
function getRecentAssignments(payload) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("All_Records");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const results = [];

  var todayDateStr = "";
  if (typeof payload === 'object' && payload !== null) {
    todayDateStr = payload.todayStr || payload.today || "";
  } else {
    todayDateStr = String(payload || "");
  }
  todayDateStr = normalizeDateStr(todayDateStr);

  let baseDate = todayDateStr ? new Date(todayDateStr.replace(/-/g, '/')) : new Date();
  if (isNaN(baseDate.getTime())) baseDate = new Date();

  // 計算執行日期的前後 3 天範圍 (今天 -3 天 ~ 今天 +3 天)
  var minDateStr = normalizeDateStr(new Date(baseDate.getTime() - 3 * 24 * 60 * 60 * 1000));
  var maxDateStr = normalizeDateStr(new Date(baseDate.getTime() + 3 * 24 * 60 * 60 * 1000));

  for (let i = 1; i < data.length; i++) {
    let rowDate = normalizeDateStr(data[i][4]); // 💡 改為抓取 index 4 (執行日期)
    if (!rowDate) continue;

    // 比對執行日期是否在基準日期的 3 天之內
    if (rowDate >= minDateStr && rowDate <= maxDateStr) {
      const itemTitle = `[${data[i][2]}] ${data[i][3]}`;
      if (!results.includes(itemTitle)) results.push(itemTitle);
    }
  }
  return results;
}

// 6. 取得未完成缺交項目 (看板用)
function getUnsettledMissingItems() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Missing_Items");
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const grouped = {};

  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === "未繳") {
      const item = data[i][3];
      const studentNo = String(data[i][2]).padStart(2, '0');
      if (!grouped[item]) grouped[item] = [];
      if (!grouped[item].includes(studentNo)) grouped[item].push(studentNo);
    }
  }
  return grouped;
}

// 7. 教師缺交語音登記 (修正版)
function processMissingAudio(base64Audio, mimeType, defaultItem, todayDate) {
  const pureBase64 = base64Audio.includes(',') ? base64Audio.split(',')[1] : base64Audio;
  let cleanMime = mimeType ? mimeType.split(';')[0].trim() : 'audio/mp4';
  if (cleanMime === 'audio/m4a') cleanMime = 'audio/mp4';

  // 1. 日期正確計算 (避免字串解析時區偏差)
  const baseDate = todayDate ? new Date(todayDate.replace(/-/g, '/')) : new Date();
  
  const todayStr = Utilities.formatDate(baseDate, "GMT+8", "yyyy-MM-dd");
  
  const tomorrow = new Date(baseDate.getTime() + 24 * 60 * 60 * 1000);
  const tomorrowStr = Utilities.formatDate(tomorrow, "GMT+8", "yyyy-MM-dd");

  const afterTomorrow = new Date(baseDate.getTime() + 2 * 24 * 60 * 60 * 1000);
  const afterTomorrowStr = Utilities.formatDate(afterTomorrow, "GMT+8", "yyyy-MM-dd");

  // 2. 注入明確的日期對照表給 Gemini
  const prompt = `
你是一個課堂登記助理。請聽取語音內容，解析出「作業/科目名稱」、「目標日期」以及「座號數字清單」。

【日期參照基準】：
- 今天：${todayStr}
- 明天：${tomorrowStr}
- 後天：${afterTomorrowStr}

【解析規則】：
1. "targetItem"：提取作業名稱。優先使用語音中提到的名稱，若語音未提及則使用介面預選項目："${defaultItem || "無"}"。
2. "targetDate"：
   - 若語音提到「明天」或「明天的考試/作業」，必須回傳 "${tomorrowStr}"。
   - 若語音提到「後天」，必須回傳 "${afterTomorrowStr}"。
   - 若語音提到「今天」，回傳 "${todayStr}"。
   - 若語音「完全沒提及日期」，預設回傳明天："${tomorrowStr}"。
3. "studentNos"：提取座號數字（如 "5號10號" -> ["05", "10"]），請補齊為兩位數。

僅回傳純 JSON 物件格式：
{
  "targetItem": "國語講義",
  "targetDate": "${tomorrowStr}",
  "studentNos": ["05", "10"]
}
`;

  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        { "inline_data": { "mime_type": cleanMime, "data": pureBase64 } }
      ]
    }]
  };

  const modelName = getLatestFlashModelName();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${GEMINI_API_KEY}`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });

    const resObj = JSON.parse(response.getContentText());

    if (resObj.error) {
      return { success: false, error: `Gemini API 錯誤 (${resObj.error.code}): ${resObj.error.message}` };
    }
    if (!resObj.candidates || resObj.candidates.length === 0 || !resObj.candidates[0].content) {
      return { success: false, error: "無法解析此段語音，請確認聲音是否清晰。" };
    }

    let rawText = resObj.candidates[0].content.parts[0].text.trim();
    const jsonMatch = rawText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return { success: false, error: "無法解析出標準 JSON 資料結構。" };
    }
    
    const result = JSON.parse(jsonMatch[0]);
    const finalItem = (result.targetItem || defaultItem || "").trim();
    const finalDate = result.targetDate || tomorrowStr; // 預設帶入明天
    const studentNos = result.studentNos || [];

    if (!finalItem) {
      return { success: false, error: "語音中未聽出作業名稱，請先選擇作業項目或於語音中說出項目。" };
    }
    if (studentNos.length === 0) {
      return { success: false, error: "未辨識出任何缺交座號。" };
    }

    // 寫入試算表，使用 Gemini 解析出的指定日期 finalDate
    return saveNewMissingRecords(finalItem, studentNos, finalDate, "語音輸入");
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
// 8. 看板點擊座號快速銷案
function markSingleStudentCompleted(studentNo, item) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Missing_Items");
  const data = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  const targetNoStr = String(studentNo).padStart(2, '0');
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === item && String(data[i][2]).padStart(2, '0') === targetNoStr && data[i][4] === "未繳") {
      sheet.getRange(i + 1, 5).setValue("已補交");
      sheet.getRange(i + 1, 6).setValue(nowStr);
      updatedCount++;
    }
  }

  updateStudentsSheetSummary();
  return { success: true, count: updatedCount };
}

// 9. 家長端查詢個人缺交 (修正版)
function getStudentMissingItems(studentNo) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Missing_Items");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const results = [];
  
  // 將傳入的座號轉為整數數字（避免 "3" 與 "03" 格式不符的情況）
  const targetNo = parseInt(studentNo, 10);

  for (let i = 1; i < data.length; i++) {
    const rowSeatNo = parseInt(data[i][2], 10); // 第 3 欄 (Index 2)：座號
    const status = String(data[i][4]).trim();   // 第 5 欄 (Index 4)：狀態

    // 比對座號與狀態是否為「未繳」
    if (!isNaN(rowSeatNo) && rowSeatNo === targetNo && status === "未繳") {
      let regDate = data[i][1]; // 第 2 欄 (Index 1)：登記日期
      
      // 日期格式化處理
      if (regDate instanceof Date) {
        regDate = Utilities.formatDate(regDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        regDate = String(regDate).trim();
      }

      results.push({
        id: data[i][0],      // 第 1 欄 (Index 0)：流水號
        item: data[i][3],    // 第 4 欄 (Index 3)：缺交項目名稱
        date: regDate        // 登記日期
      });
    }
  }
  return results;
}

/**
 * 儲存缺交記錄至 Google 試算表
 * @param {string} targetItem 作業或項目名稱
 * @param {Array} studentNos 座號陣列 (例: ["2", "12", "30"])
 * @param {string} todayDate 登記日期
 */
/**
 * 寫入缺交記錄至『Missing_Items』分頁
 * 欄位順序：1.流水號 2.登記日期 3.座號 4.缺交項目名稱 5.狀態 6.銷案/補交時間 7.來源
 */
function saveNewMissingRecords(targetItem, studentNos, todayDate, source) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Missing_Items");
    
    if (!sheet) {
      return { success: false, error: "找不到『Missing_Items』工作表，請先檢查試算表分頁名稱！" };
    }

    if (!Array.isArray(studentNos)) {
      studentNos = [studentNos];
    }

    // 取得現有資料進行比對
    const existingData = sheet.getDataRange().getValues();
    let count = 0;
    const now = new Date();

    studentNos.forEach((no, index) => {
      const cleanNo = String(no).trim().padStart(2, '0');
      
      if (cleanNo !== "" && cleanNo !== "00") {
        // 防重檢查：座號 (第3欄 index 2) + 項目 (第4欄 index 3) + 狀態為未繳 (第5欄 index 4)
        const isDuplicate = existingData.some(row => {
          const rowSeat = String(row[2]).trim().padStart(2, '0');
          const rowItem = String(row[3]).trim();
          const rowStatus = String(row[4]).trim();
          return rowSeat === cleanNo && rowItem === targetItem && rowStatus === "未繳";
        });

        // 未重複才寫入
        if (!isDuplicate) {
          const serialNo = "M" + Utilities.formatDate(now, "GMT+8", "yyyyMMddHHmmss") + (index + 1);

          sheet.appendRow([
            serialNo,            // 1. 流水號
            todayDate,           // 2. 登記日期
            cleanNo,             // 3. 座號 (補齊為 03 格式)
            targetItem,          // 4. 缺交項目名稱
            "未繳",              // 5. 狀態
            "",                  // 6. 銷案/補交時間
            source || "教師端"   // 7. 來源
          ]);
          count++;
        }
      }
    });

    return { success: true, count: count };
  } catch (err) {
    return { success: false, error: "寫入『Missing_Items』失敗：" + err.toString() };
  }
}
// 5. 查詢下周全週提醒事項 (下周日 ~ 下周六，以週日為每週起始日)
function getNextWeekReminders(payload) {
  var startDateStr = "";
  var endDateStr = "";

  if (typeof payload === 'object' && payload !== null) {
    startDateStr = payload.startDateStr || "";
    endDateStr = payload.endDateStr || "";
  }

  startDateStr = normalizeDateStr(startDateStr);
  endDateStr = normalizeDateStr(endDateStr);

  // 若前端未傳入區間，自動計算下週日與下週六
  if (!startDateStr || !endDateStr) {
    var todayDateStr = (typeof payload === 'object' && payload !== null) ? (payload.todayStr || payload.today || "") : String(payload || "");
    todayDateStr = normalizeDateStr(todayDateStr);

    let baseDate = todayDateStr ? new Date(todayDateStr.replace(/-/g, '/')) : new Date();
    if (isNaN(baseDate.getTime())) baseDate = new Date();

    const dayOfWeek = baseDate.getDay(); // 0(週日) ~ 6(週六)
    
    // 下週日 = 當前日期 + (7 - 當前星期) 天
    const nextSun = new Date(baseDate.getTime() + (7 - dayOfWeek) * 24 * 60 * 60 * 1000);
    // 下週六 = 當前日期 + (13 - 當前星期) 天
    const nextSat = new Date(baseDate.getTime() + (13 - dayOfWeek) * 24 * 60 * 60 * 1000);
    
    startDateStr = normalizeDateStr(nextSun);
    endDateStr = normalizeDateStr(nextSat);
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("All_Records");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return [];

  const reminders = [];
  const weekDayNames = ["日", "一", "二", "三", "四", "五", "六"];
  
  for (let i = 1; i < data.length; i++) {
    let rowDate = normalizeDateStr(data[i][4]); // 第 5 欄：執行日期
    if (!rowDate) continue;
    
    // 篩選執行日期在下週日 ~ 下週六區間內的項目
    if (rowDate >= startDateStr && rowDate <= endDateStr) {
      const d = new Date(rowDate.replace(/-/g, '/'));
      const dayName = !isNaN(d.getDay()) ? weekDayNames[d.getDay()] : "";
      const monthDay = !isNaN(d.getTime()) ? `${d.getMonth() + 1}/${d.getDate()}` : rowDate;
      
      reminders.push({
        type: data[i][2] || "提醒",
        range: data[i][3] || "",
        date: rowDate,
        dateLabel: `${monthDay} (${dayName})`,
        source: data[i][5] || ""
      });
    }
  }
  
  reminders.sort((a, b) => a.date.localeCompare(b.date));
  return reminders;
}
// 📌 新增：查詢今日提醒事項
// 📌 取得「今日任務」項目（將原本的 getTodayReminders 改名，邏輯保持不變）
function getTodayTasks(payload) {
  var targetDateStr = "";
  if (typeof payload === 'object' && payload !== null) {
    targetDateStr = payload.todayStr || payload.targetDateStr || "";
  } else {
    targetDateStr = String(payload || "");
  }
  targetDateStr = normalizeDateStr(targetDateStr);

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("All_Records");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const tasks = [];

  for (let i = 1; i < data.length; i++) {
    let rowDate = normalizeDateStr(data[i][4]); // 第 5 欄：執行日期
    if (rowDate === targetDateStr && rowDate !== "") {
      tasks.push({ 
        type: data[i][2] || "任務",  // 若無分類，預設標示為任務
        range: data[i][3] || "", 
        date: rowDate, 
        source: data[i][5] || "" 
      });
    }
  }
  return tasks;
}

// 💡 建立相容別名：避免舊版前端呼叫時找不到函式
function getTodayReminders(payload) {
  return getTodayTasks(payload);
}
/**
 * 密碼標準化處理（自動補齊前端被試算表吃掉的 0）
 */
function normalizePassword(rawPwd) {
  if (rawPwd === null || rawPwd === undefined) return "";
  var str = String(rawPwd).trim();
  if (str === "") return "";
  
  // 自動補足至 4 碼 (例: "123" -> "0123", "5" -> "0005")
  return str.padStart(4, '0');
}
/**
 * 1. 從「Student」試算表分頁讀取學生名單 (欄位 A: 座號, 欄位 B: 姓名)
 */
function getStudentListFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Student') || ss.getSheetByName('學生名單');
  
  // 若沒有 Student 分頁則自動新建並注入預設資料
  if (!sheet) {
    sheet = ss.insertSheet('Student');
    sheet.appendRow(['座號', '姓名']);
    const defaultData = [
      [1, '李燐'], [2, '林瑜'], [3, '林廷'], [4, '林翰'],
      [5, '郇宸'], [6, '莊勳'], [7, '郭評'], [8, '陳澤'],
      [9, '陳澄'], [10, '陳安'], [11, '黃翔'], [12, '楊棟']
    ];
    sheet.getRange(2, 1, defaultData.length, 2).setValues(defaultData);
  }
  
  const values = sheet.getDataRange().getValues();
  const students = [];
  // 跳過首列標題
  for (let i = 1; i < values.length; i++) {
    const id = parseInt(values[i][0]);
    const name = String(values[i][1] || '').trim();
    if (!isNaN(id) && name) {
      students.push({ id: id, name: name });
    }
  }
  students.sort((a, b) => a.id - b.id);
  return students;
}
/**
 * 2. 讀取儲存在 GAS 專案中的座位與分組設定 (結合 Student 分頁最新名單)
 */
function getSeatingConfig() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('CLASS_SEATING_CONFIG');
  
  const sheetStudents = getStudentListFromSheet();
  
  let config = null;
  if (raw) {
    try {
      config = JSON.parse(raw);
    } catch (e) {
      console.error('解析儲存設定失敗', e);
    }
  }
  
  if (!config) {
    config = {
      activeMode: 'seats',
      cols: 7, 
      rows: 7, // 預設 7x7 矩陣
      emptySeats: [],
      studentSeats: {},
      drawnStudentIds: [],
      podiumPosition: 'top',
      currentGroupSlot: 0,
      groupSlots: Array.from({ length: 10 }, (_, i) => ({
        lessonName: i === 0 ? '課堂分組 A' : `課堂 ${i + 1}`,
        groupCount: 6,
        groups: []
      }))
    };
  }
  
  // 強制同步來自 Student 試算表最新學生名單
  config.customStudents = sheetStudents;
  return config;
}

/**
 * 3. 儲存座位與分組設定 JSON 至 ScriptProperties
 */
function saveSeatingConfig(configObj) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('CLASS_SEATING_CONFIG', JSON.stringify(configObj));
    return { success: true, message: '雲端座位與分組設定已成功儲存！' };
  } catch (e) {
    return { success: false, message: '儲存失敗：' + e.toString() };
  }
}

/**
 * 4. 將當前排好的座位表匯出寫入 Google 試算表（建立「座位表」分頁供列印）
 */
function exportSeatingToSheet(seatingData) {
  const { cols, rows, emptySeats, studentSeats, podiumPosition } = seatingData;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName('座位表');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('座位表');
  }
  
  sheet.setGridLines(true);
  
  // 設置大標題
  sheet.getRange(1, 1).setValue('【 班級座位表 】').setFontSize(16).setFontWeight('bold');
  
  let startRow = 3;
  if (podiumPosition === 'top') {
    sheet.getRange(startRow, 1, 1, cols).merge()
      .setValue('👨‍🏫 講 台 / 黑 板 ( 前 方 )')
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    startRow += 2;
  }
  
  // 繪製座位陣列
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const idx = r * cols + c;
      const targetCell = sheet.getRange(startRow + r, c + 1);
      
      if (emptySeats.includes(idx)) {
        targetCell.setValue('【走道】')
          .setBackground('#f1f5f9')
          .setFontColor('#94a3b8')
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      } else {
        const student = studentSeats[idx];
        if (student) {
          targetCell.setValue(`${student.id}. ${student.name}`)
            .setFontSize(12)
            .setFontWeight('bold')
            .setBackground('#ffffff')
            .setFontColor('#0f172a')
            .setBorder(true, true, true, true, null, null, '#64748b', SpreadsheetApp.BorderStyle.SOLID)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
        } else {
          targetCell.setValue('(空位)')
            .setFontSize(10)
            .setBackground('#fafafa')
            .setFontColor('#cbd5e1')
            .setBorder(true, true, true, true, null, null, '#cbd5e1', SpreadsheetApp.BorderStyle.DASHED)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
        }
      }
    }
    sheet.setRowHeight(startRow + r, 45);
  }
  
  if (podiumPosition === 'bottom') {
    const bottomPodiumRow = startRow + rows + 1;
    sheet.getRange(bottomPodiumRow, 1, 1, cols).merge()
      .setValue('👨‍🏫 講 台 / 黑 板 ( 後 方 )')
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }
  
  // 調整欄寬
  for (let c = 1; c <= cols; c++) {
    sheet.setColumnWidth(c, 110);
  }
  
  ss.setActiveSheet(sheet);
  return { success: true, sheetName: '座位表' };
}
/**
 * 讀取 Students 分頁 A1 儲存格的班級名稱
 */
function getClassTitle() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Students"); // 請確認試算表分頁名稱完全一致 (區分大小寫)
    
    if (!sheet) {
      Logger.log("⚠️ 找不到 Students 分頁");
      return "智慧教室";
    }
    
    const val = sheet.getRange("A1").getValue();
    return (val && val.toString().trim() !== "") ? val.toString().trim() : "智慧教室";
  } catch (err) {
    Logger.log("讀取 A1 失敗: " + err.toString());
    return "智慧教室";
  }
}

// 🟢 讀取 Students 分頁 C1 儲存格的教師大屏管理 PIN 碼 (預設: 8888)
function getTeacherPin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Students");
    if (!sheet) return "8888";
    const val = sheet.getRange("C1").getValue();
    return (val && String(val).trim() !== "") ? String(val).trim() : "8888";
  } catch (err) {
    return "8888";
  }
}
// 1. 讀取 Students 分頁名單
function getStudentData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Students');
  
  if (!sheet) {
    return { error: '找不到名為 "Students" 的分頁！' };
  }

  const className = sheet.getRange('A1').getValue() || '未命名班級';
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 3) {
    return { className: className, students: [] };
  }

  const lastCol = Math.max(sheet.getLastColumn(), 5);
  const values = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const photoMap = getStudentPhotoMap(ss);
  const students = values
    .filter(row => row[0] !== '' && row[1] !== '')
    .map(row => ({
      id: String(row[0]).trim(),
      photoUrl: photoMap[String(row[0]).trim()] || ((row.length >= 5 && row[4]) ? String(row[4]).trim() : ''),
      name: String(row[1]).trim()
    }));

  return { className: className, students: students };
}

// 2. 儲存座位狀態至 DocumentProperties
function saveSeatingState(stateJson) {
  try {
    PropertiesService.getDocumentProperties().setProperty('SEATING_LAYOUT_STATE', JSON.stringify(stateJson));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 3. 讀取儲存的座位狀態
function loadSeatingState() {
  try {
    const data = PropertiesService.getDocumentProperties().getProperty('SEATING_LAYOUT_STATE');
    return data ? JSON.parse(data) : null;
  } catch (e) {
    return null;
  }
}

// 4. 匯出座位表為新試算表分頁
function exportSeatingToSheet(exportData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
  const sheetName = `座位表_${timestamp}`;
  
  let sheet = ss.insertSheet(sheetName);
  const { rows, cols, podiumTop, grid, className } = exportData;

  sheet.getRange('A1').setValue(`${className} - 教室座位表`).setFontSize(16).setFontWeight('bold');
  
  let startRow = 3;
  if (podiumTop) {
    sheet.getRange(startRow, 1, 1, cols).merge().setValue('【 講 台 / 黑 板 】')
      .setHorizontalAlignment('center').setBackground('#e2e8f0').setFontWeight('bold');
    startRow += 2;
  }

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cellData = grid[r][c];
      const cell = sheet.getRange(startRow + r, c + 1);
      
      if (cellData.isAisle) {
        cell.setValue('【走道】').setBackground('#f1f5f9').setFontColor('#94a3b8');
      } else if (cellData.student) {
        cell.setValue(`(${cellData.student.id}) ${cellData.student.name}`).setBackground('#ffffff');
      } else {
        cell.setValue('[ 空位 ]').setBackground('#f8fafc').setFontColor('#cbd5e1');
      }
      cell.setBorder(true, true, true, true, false, false, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
      cell.setHorizontalAlignment('center').setVerticalAlignment('middle');
    }
  }

  if (!podiumTop) {
    startRow += rows + 1;
    sheet.getRange(startRow, 1, 1, cols).merge().setValue('【 講 台 / 黑 板 】')
      .setHorizontalAlignment('center').setBackground('#e2e8f0').setFontWeight('bold');
  }

  for (let c = 1; c <= cols; c++) {
    sheet.setColumnWidth(c, 120);
  }

  return { success: true, sheetName: sheetName };
}

// ==========================================
// 📸 學生照片座位表與 Gemini 視覺辨識專用模組
// ==========================================

/**
 * 上網動態搜尋 Google Gemini 最新 Vision / Flash 模型
 */
function searchLatestGeminiVisionModel(apiKey) {
  const key = apiKey || (typeof GEMINI_API_KEY !== 'undefined' ? GEMINI_API_KEY : '');
  const defaultModel = "gemini-2.5-flash";
  if (!key || key === "請填入您的 Gemini API Key") return defaultModel;
  
  const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${key}`;
  try {
    const response = UrlFetchApp.fetch(listUrl, { "muteHttpExceptions": true });
    const result = JSON.parse(response.getContentText());
    if (result.models && Array.isArray(result.models)) {
      const candidates = result.models
        .filter(m => {
          const name = (m.name || '').toLowerCase();
          const hasGen = m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent");
          return hasGen && name.includes("flash") && !name.includes("omni") && !name.includes("exp");
        })
        .map(m => m.name.replace("models/", ""));
      
      if (candidates.length > 0) {
        candidates.sort((a, b) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }));
        Logger.log("🔍 上網查詢到最新 Gemini 視覺模型：" + candidates[0]);
        return candidates[0];
      }
    }
  } catch (err) {
    Logger.log("⚠️ 搜尋最新 Gemini 模型失敗，使用預設值: " + err.toString());
  }
  return defaultModel;
}

/**
 * 取得最新學生照片對照表 { [seatNo]: photoUrl }
 */
function getStudentPhotos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const photoMap = getStudentPhotoMap(ss);
    return { success: true, photos: photoMap };
  } catch (err) {
    return { success: false, error: err.toString(), photos: {} };
  }
}

/**
 * 讀取 Seating_Photos 與 Students 分頁的照片對照表
 */
function getStudentPhotoMap(ss) {
  const photoMap = {};
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 先讀取 Seating_Photos
  const pSheet = ss.getSheetByName("Seating_Photos");
  if (pSheet && pSheet.getLastRow() > 1) {
    const pData = pSheet.getDataRange().getValues();
    for (let r = 1; r < pData.length; r++) {
      const seatNo = String(pData[r][0]).trim();
      const url = String(pData[r][3] || "").trim();
      if (seatNo && url) {
        photoMap[seatNo] = url;
      }
    }
  }

  // 2. 輔助讀取 Students 分頁第 E 欄 (照片網址)
  const sSheet = ss.getSheetByName("Students");
  if (sSheet && sSheet.getLastRow() >= 3) {
    const lastCol = Math.max(sSheet.getLastColumn(), 5);
    const sData = sSheet.getRange(3, 1, sSheet.getLastRow() - 2, lastCol).getValues();
    for (let r = 0; r < sData.length; r++) {
      const seatNo = String(sData[r][0]).trim();
      const url = sData[r].length >= 5 ? String(sData[r][4] || "").trim() : "";
      if (seatNo && url && !photoMap[seatNo]) {
        photoMap[seatNo] = url;
      }
    }
  }

  return photoMap;
}

/**
 * 接收手機/網頁傳送的圖片 Base64，呼叫 Gemini API，把圖片存入 Google Drive，並更新 Google Sheet
 */
function uploadAndRecognizeStudentPhoto(data) {
  try {
    const base64Image = data.base64Image || data.base64;
    if (!base64Image) {
      return { success: false, error: "未接收到照片資料！" };
    }
    const cleanBase64 = base64Image.replace(/^data:image\/(png|jpeg|jpg|webp);base64,/, "");
    const mimeType = data.mimeType || "image/jpeg";
    
    let seatNo = data.manualSeatNo ? String(data.manualSeatNo).trim() : null;
    let studentName = data.manualName ? String(data.manualName).trim() : "";
    let modelUsed = "手動指定";
    let recognitionSuccess = false;

    // 1. 若未手動指定座號，上網搜尋最新模型並呼叫 Gemini Vision API 辨識
    if (!seatNo) {
      modelUsed = searchLatestGeminiVisionModel(data.apiKey);
      const key = data.apiKey || (typeof GEMINI_API_KEY !== 'undefined' ? GEMINI_API_KEY : '');
      if (!key || key === "請填入您的 Gemini API Key") {
        return { success: false, error: "未設定 Gemini API Key，無法進行 AI 視覺辨識，請手動指定座號或填入 API Key！" };
      }

      const visionPrompt = `
你是一位精準的班級照片分析助手。
學生正手持或在胸前舉著一張白紙，紙上寫有學生的【座號】與【姓名】。
請仔細辨識照片中白紙上的文字，並輸出嚴格的純 JSON 格式：
{
  "seat_number": 數字或 null,
  "name": "學生姓名"
}
【重要規則】：
1. "seat_number" 必須是純整數數字（例如看到「05」、「5」、「5號」，請回傳 5；若看到「12」，請回傳 12）。若完全辨識不出座號數字，請回傳 null。
2. "name" 為白紙上書寫的學生繁體中文姓名（例如「張小明」）。若未寫姓名或字跡過度模糊無法辨認，請回傳空字串 ""。
3. 嚴禁輸出任何 markdown 區塊（例如 \`\`\`json ），僅回傳純 JSON 物件。
`;

      const payload = {
        "contents": [{
          "parts": [
            { "text": visionPrompt },
            { "inline_data": { "mime_type": mimeType, "data": cleanBase64 } }
          ]
        }],
        "generationConfig": {
          "response_mime_type": "application/json",
          "temperature": 0.1
        }
      };

      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelUsed}:generateContent?key=${key}`;
      const response = UrlFetchApp.fetch(url, {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      });

      const resObj = JSON.parse(response.getContentText());
      if (resObj.error) {
        return { success: false, error: `Gemini API 錯誤 (${resObj.error.code}): ${resObj.error.message}`, modelUsed: modelUsed };
      }

      if (!resObj.candidates || resObj.candidates.length === 0 || !resObj.candidates[0].content) {
        return { success: false, error: "Gemini 無法辨識此照片，請確認照片清晰度。", modelUsed: modelUsed };
      }

      let rawText = resObj.candidates[0].content.parts[0].text.trim().replace(/```json/g, "").replace(/```/g, "").trim();
      let parsed = {};
      try {
        parsed = JSON.parse(rawText);
      } catch (e) {
        const jsonMatch = rawText.match(/\{[\s\S]*\}/);
        if (jsonMatch) parsed = JSON.parse(jsonMatch[0]);
      }

      if (parsed.seat_number !== undefined && parsed.seat_number !== null && !isNaN(parseInt(parsed.seat_number))) {
        seatNo = String(parseInt(parsed.seat_number));
        studentName = parsed.name || "";
        recognitionSuccess = true;
      } else {
        return {
          success: false,
          error: "未能從照片中清晰辨識出座號數字（可能太模糊或反光），請重拍或使用手動指定座號補救！",
          recognizedName: parsed.name || "",
          modelUsed: modelUsed
        };
      }
    } else {
      recognitionSuccess = true;
    }

    // 2. 將照片存入 Google Drive
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let folder;
    const folderName = "大屏班務_座位表照片";
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    const decodedBytes = Utilities.base64Decode(cleanBase64);
    const fileName = `student_${seatNo.padStart(2, '0')}_${Date.now()}.jpg`;
    const blob = Utilities.newBlob(decodedBytes, mimeType, fileName);
    const file = folder.createFile(blob);
    
    // 設定共用權限為知道連結者可檢視
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 轉換為直連網址格式
    const photoUrl = `https://lh3.googleusercontent.com/d/${file.getId()}`;

    // 3. 取得學生姓名（若辨識未取得，比對 Students 分頁補齊）
    if (!studentName) {
      const studentSheet = ss.getSheetByName("Students");
      if (studentSheet) {
        const sData = studentSheet.getDataRange().getValues();
        for (let i = 2; i < sData.length; i++) {
          if (String(sData[i][0]).trim() === seatNo) {
            studentName = String(sData[i][1]).trim();
            break;
          }
        }
      }
    }

    // 4. 更新 Google Sheet: Seating_Photos 分頁
    let photoSheet = ss.getSheetByName("Seating_Photos");
    if (!photoSheet) {
      photoSheet = ss.insertSheet("Seating_Photos");
      photoSheet.getRange(1, 1, 1, 6).setValues([["座號", "姓名", "座位編號", "照片網址", "辨識狀態", "更新時間"]]);
      photoSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e2e8f0");
      photoSheet.setFrozenRows(1);
    }

    const pData = photoSheet.getDataRange().getValues();
    let foundRow = -1;
    for (let r = 1; r < pData.length; r++) {
      if (String(pData[r][0]).trim() === seatNo) {
        foundRow = r + 1;
        break;
      }
    }

    const timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    const seatId = data.seatId || `S${seatNo.padStart(2, '0')}`;
    const statusText = recognitionSuccess ? "Success" : "Manual";

    if (foundRow > 0) {
      // 重拍覆蓋
      photoSheet.getRange(foundRow, 1, 1, 6).setValues([[
        seatNo,
        studentName || pData[foundRow - 1][1] || `學生${seatNo}`,
        seatId,
        photoUrl,
        statusText,
        timestamp
      ]]);
    } else {
      // 新增紀錄
      photoSheet.appendRow([
        seatNo,
        studentName || `學生${seatNo}`,
        seatId,
        photoUrl,
        statusText,
        timestamp
      ]);
    }

    // 5. 同步更新 Students 分頁 (若有 E 欄 照片網址 與 F 欄 辨識狀態)
    const studentSheet = ss.getSheetByName("Students");
    if (studentSheet) {
      const sData = studentSheet.getDataRange().getValues();
      for (let i = 2; i < sData.length; i++) {
        if (String(sData[i][0]).trim() === seatNo) {
          studentSheet.getRange(i + 1, 5).setValue(photoUrl);
          studentSheet.getRange(i + 1, 6).setValue(statusText);
          break;
        }
      }
    }

    // 6. 更新儲存於 DocumentProperties 中的 SEATING_LAYOUT_STATE
    try {
      const saved = loadSeatingState();
      if (saved && saved.grid) {
        let updated = false;
        saved.grid.forEach(row => {
          row.forEach(cell => {
            if (cell.student && String(cell.student.id).trim() === seatNo) {
              cell.student.photoUrl = photoUrl;
              if (studentName) cell.student.name = studentName;
              updated = true;
            }
          });
        });
        if (updated) {
          saveSeatingState(saved);
        }
      }
    } catch (e) {
      Logger.log("更新 SEATING_LAYOUT_STATE 失敗: " + e.toString());
    }

    return {
      success: true,
      seatNumber: seatNo,
      name: studentName,
      photoUrl: photoUrl,
      fileId: file.getId(),
      modelUsed: modelUsed
    };
  } catch (err) {
    Logger.log("uploadAndRecognizeStudentPhoto 錯誤: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

