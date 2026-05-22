function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Create a custom menu
  ui.createMenu("額外選單").addItem("顯示報到側邊欄", "showSidebar").addToUi();
}

function showSidebar() {
  const template = HtmlService.createTemplateFromFile("Sidebar.html");
  const id = PropertiesService.getScriptProperties().getProperty("id");
  template.url = id ? "https://kuomartin.github.io/reg_dev/v3/?id=" + id : "";
  const htmlOutput = template
    .evaluate()
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("學員報到系統");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function doGet(e) {
  const isBridgeMode = e && e.parameter && e.parameter.bridge === "1";
  const serviceUrl = ScriptApp.getService().getUrl();
  const match =
    serviceUrl && serviceUrl.match(/\/macros\/s\/([^/]+)\/(?:exec|dev)/);
  const deploymentId = match ? match[1] : null;
  const shouldSave = e && e.parameter && e.parameter.save === "1";
  if (shouldSave && deploymentId) {
    PropertiesService.getScriptProperties().setProperty("id", deploymentId);
  }
  const template = HtmlService.createTemplateFromFile("Bridge");
  template.deploymentId = deploymentId || "";
  template.redirectUrl = deploymentId
    ? "https://kuomartin.github.io/reg_dev/v3/?id=" + deploymentId
    : "https://kuomartin.github.io/reg_dev/v3/";
  template.saveUrl = serviceUrl
    ? serviceUrl + (serviceUrl.indexOf("?") === -1 ? "?" : "&") + "save=1"
    : "";
  template.saved = shouldSave;
  template.isBridgeMode = isBridgeMode;
  return template
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 優化後的學員搜尋：僅搜尋關鍵欄位，不載入全表
 * @param {string} query
 * @returns {{data: Map<string, any>, range: GoogleAppsScript.Spreadsheet.Range} | null}
 */
function findStudentRow(query) {
  if (!query) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("學員名單");
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow <= 1) return null; // 只有標題或空表

  // 1. 取得標題行（用於後續對應欄位名稱）
  const title = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const idIndex = title.indexOf("學員代號") + 1; // getRange 是 1-based
  const phoneIndex = title.indexOf("手機") + 1;

  // 2. 使用 TextFinder 鎖定特定欄位進行全文匹配
  // 我們先找「學員代號」，沒找到再找「手機」
  const searchIndices = [idIndex, phoneIndex].filter((idx) => idx > 0);
  let foundRange = null;

  for (let colIdx of searchIndices) {
    const searchRange = sheet.getRange(2, colIdx, lastRow - 1, 1);
    foundRange = searchRange
      .createTextFinder(query)
      .matchEntireCell(true) // 精確匹配
      .findNext();

    if (foundRange) break;
  }

  // 3. 如果找到了，只讀取那一行的資料
  if (foundRange) {
    const rowIndex = foundRange.getRow();
    const rowRange = sheet.getRange(rowIndex, 1, 1, lastColumn);
    const rowValues = rowRange.getValues()[0];

    return {
      data: new Map(rowValues.map((value, index) => [title[index], value])),
      range: rowRange,
    };
  }

  return null;
}

/**
 * replace string template with given values
 * @param {string} template
 * @param {Map<string,any>} data
 * @returns {string}
 */
function msgFormat(template, data) {
  let s = template;
  for (let [key, value] of data) {
    s = s.replaceAll(`{{${key}}}`, value);
  }
  return s;
}

/**
 * @param {string} studentId
 * @param {string} sheetName
 * @param {boolean} [autoJump=true]
 * @returns {{success: boolean, message: string}}
 */
function checkInStudent(studentId, sheetName, autoJump) {
  if (autoJump === undefined) autoJump = true;
  if (!sheetName) sheetName = "學員報到";
  const lock = LockService.getDocumentLock();
  try {
    let studentResult = findStudentRow(studentId);
    if (!studentResult) {
      return { success: false, message: "查無此學員" };
    }
    let student = studentResult.data;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["學員代號", "報到時間"]);
    }

    // Acquire lock to avoid race conditions
    lock.waitLock(10000);

    // Fast and safe approach: appendRow automatically adds data right after the last row
    // and handles sheet expansion without having to fetch the entire sheet.
    sheet.appendRow([`'${studentId}`, new Date()]);

    lock.releaseLock();

    // 報到成功後，若設定自動跳轉，則嘗試將試算表畫面跳轉到學員名單中對應的行
    if (autoJump) {
      try {
        studentResult.range.activate();
      } catch (activateErr) {
        // 若是從外部 Web App 呼叫，activate 可能會失敗，這裡直接忽略錯誤
      }
    }

    let msgTemplate = getMsgTemplate();
    return { success: true, message: msgFormat(msgTemplate, student) };
  } catch (e) {
    return { success: false, message: "報到失敗: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}
function int2Col(col) {
  let result = "";
  while (col > 0) {
    let remainder = col % 26;
    if (remainder === 0) {
      remainder = 26;
      col = col - 1;
    }
    result = String.fromCharCode(64 + remainder) + result;
    col = Math.floor(col / 26);
  }
  return result;
}

let CACHED_TEMPLATE = null;

function getMsgTemplate() {
  if (CACHED_TEMPLATE !== null) return CACHED_TEMPLATE;

  const cache = CacheService.getScriptCache();
  let jsonTemplate = cache.get("msgTemplate");

  if (!jsonTemplate) {
    jsonTemplate =
      PropertiesService.getScriptProperties().getProperty("msgTemplate");
    if (jsonTemplate) {
      cache.put("msgTemplate", jsonTemplate, 21600); // 快取 6 小時
    }
  }

  if (!jsonTemplate) {
    CACHED_TEMPLATE = "報到成功；姓名：{{姓名}}，序號：{{序號}}";
  } else {
    try {
      CACHED_TEMPLATE = JSON.parse(jsonTemplate);
    } catch (e) {
      CACHED_TEMPLATE = jsonTemplate;
    }
  }
  return CACHED_TEMPLATE;
}

function updateMsgTemplate(template) {
  const jsonTemplate = JSON.stringify(template);
  PropertiesService.getScriptProperties().setProperty(
    "msgTemplate",
    jsonTemplate,
  );
  CacheService.getScriptCache().put("msgTemplate", jsonTemplate, 21600);
  CACHED_TEMPLATE = template;
  return { success: true, message: "訊息模板已更新" };
}

function getCheckInSheets() {
  const props = PropertiesService.getScriptProperties();
  const sheetsJson = props.getProperty("checkInSheets");
  if (!sheetsJson) {
    return ["學員報到"];
  }
  try {
    return JSON.parse(sheetsJson);
  } catch (e) {
    return ["學員報到"];
  }
}

function updateCheckInSheets(sheets) {
  if (!Array.isArray(sheets))
    return { success: false, message: "無效的資料格式" };
  const jsonSheets = JSON.stringify(sheets);
  PropertiesService.getScriptProperties().setProperty(
    "checkInSheets",
    jsonSheets,
  );
  return { success: true, message: "報到分頁已更新" };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
