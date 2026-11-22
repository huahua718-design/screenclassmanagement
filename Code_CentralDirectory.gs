/**
 * =========================================================================
 * 📖 智慧班務系統 - 【中央名錄與 6 碼班級代碼發放中心】(Central Directory GAS)
 * =========================================================================
 * 
 * 🌟 系統定位與架構理念：
 * 1. 本 GAS 專門部署在開發者的【班級系統中央名錄】Google 試算表上（只需建立一次）。
 * 2. 核心功能為「電話簿（代碼 ➔ 網址）」：
 *    - 老師首次使用時，傳送個人 GAS 網址，本系統隨機核發一組不重複的 6 碼英數代碼（例如：K8M3P9）。
 *    - 教室大屏、手機端、家長端輸入或攜帶該 6 碼代碼，向本系統查詢對應的真實 GAS 網址。
 * 3. 「查詢一次，終身直連」：
 *    - 各端點查得真實 GAS 網址後，直接存入該裝置的 localStorage。
 *    - 後續所有班級名冊、座位表、缺交、聯絡簿與 API 呼叫，完全直連老師個人的 GAS。
 *    - 本中央試算表【完全不經手任何班級與學生隱私個資】，零負擔、極速、安全！
 * =========================================================================
 */

const SHEET_NAME = "Directory";
// 排除易混淆字元 (0/O, 1/I/L)，使用 31 個清晰英數字元，6 碼可提供近 9 億種不重複組合
const CODE_CHARACTERS = "23456789ABCDEFGHJKMNPQRSTUVWXYZ";
const CODE_LENGTH = 6;

/**
 * 試算表結構初始化
 */
function ensureDirectorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Code", "GasUrl", "CreatedAt", "LastQueriedAt", "Note"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f1f5f9");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 產生不重複的 6 碼隨機代碼
 */
function generateUniqueCode(sheet) {
  const existingCodes = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      const c = String(data[i][0]).trim().toUpperCase();
      if (c) existingCodes.add(c);
    }
  }

  let newCode = "";
  let attempts = 0;
  do {
    newCode = "";
    for (let i = 0; i < CODE_LENGTH; i++) {
      const idx = Math.floor(Math.random() * CODE_CHARACTERS.length);
      newCode += CODE_CHARACTERS.charAt(idx);
    }
    attempts++;
    if (attempts > 1000) {
      throw new Error("代碼池生成碰撞過多，請檢查代碼長度設定。");
    }
  } while (existingCodes.has(newCode));

  return newCode;
}

/**
 * 格式化 JSON 回傳
 */
function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🔍 查詢 API (doGet)：透過 6 碼代碼換取真實 GasUrl
 * 呼叫範例：
 * GET https://script.google.com/macros/s/.../exec?code=K8M3P9
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const rawCode = params.code || params.classCode || "";
    const code = String(rawCode).trim().toUpperCase();

    if (!code) {
      return jsonOutput({
        success: false,
        error: "請提供欲查詢的班級代碼 (例如: ?code=K8M3P9)"
      });
    }

    const sheet = ensureDirectorySheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return jsonOutput({
        success: false,
        error: "查無此班級代碼，請確認代碼是否輸入正確或向班級導師索取。"
      });
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      const currentCode = String(data[i][0]).trim().toUpperCase();
      if (currentCode === code) {
        const gasUrl = String(data[i][1]).trim();
        // 更新最後查詢時間戳記
        try {
          sheet.getRange(i + 2, 4).setValue(new Date().toISOString());
        } catch (err) {}

        return jsonOutput({
          success: true,
          code: code,
          gasUrl: gasUrl
        });
      }
    }

    return jsonOutput({
      success: false,
      error: `查無此班級代碼【${code}】，請確認是否輸入正確或重新註冊。`
    });

  } catch (err) {
    return jsonOutput({
      success: false,
      error: `中央名錄查詢異常: ${err.toString()}`
    });
  }
}

/**
 * 📝 註冊 API (doPost)：接收老師的 GasUrl，核發 6 碼專屬班級代碼
 * 呼叫範例：
 * POST https://script.google.com/macros/s/.../exec
 * Body (JSON): { "action": "register", "gasUrl": "https://script.google.com/macros/s/.../exec" }
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const rawUrl = payload.gasUrl || payload.url || "";
    const gasUrl = String(rawUrl).trim();

    // 格式防呆檢查
    if (!gasUrl || !gasUrl.startsWith("http") || !gasUrl.includes("script.google.com")) {
      return jsonOutput({
        success: false,
        error: "無效的 Google Apps Script 網頁應用程式網址，網址必須以 https:// 開頭並包含 script.google.com"
      });
    }

    const sheet = ensureDirectorySheet();
    const lastRow = sheet.getLastRow();

    // 1. 防重複註冊：若此 GAS 網址過去曾註冊過，直接回傳既有的 6 碼代碼
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        const rowUrl = String(data[i][1]).trim();
        if (rowUrl === gasUrl) {
          const existingCode = String(data[i][0]).trim().toUpperCase();
          // 更新最後使用時間
          sheet.getRange(i + 2, 4).setValue(new Date().toISOString());
          return jsonOutput({
            success: true,
            code: existingCode,
            gasUrl: gasUrl,
            isExisting: true,
            message: `此網址已註冊過，已為您取回專屬班級代碼：${existingCode}`
          });
        }
      }
    }

    // 2. 新註冊：產生全新的 6 碼隨機不重複代碼
    const newCode = generateUniqueCode(sheet);
    const nowIso = new Date().toISOString();
    sheet.appendRow([newCode, gasUrl, nowIso, nowIso, payload.note || ""]);

    return jsonOutput({
      success: true,
      code: newCode,
      gasUrl: gasUrl,
      isNew: true,
      message: `註冊成功！您的專屬班級代碼為：${newCode}`
    });

  } catch (err) {
    return jsonOutput({
      success: false,
      error: `中央名錄註冊異常: ${err.toString()}`
    });
  }
}
