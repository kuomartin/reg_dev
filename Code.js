function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Create a custom menu
  ui.createMenu("額外選單")
    .addItem("顯示報到側邊欄", "showSidebar")
    .addItem("備份公式", "backupFomula")
    .addItem("還原公式", "restoreFomula")
    .addToUi();
}

function showSidebar() {
  const template = HtmlService.createTemplateFromFile('index.html');
  template.url = PropertiesService.getScriptProperties().getProperty('url');
  const htmlOutput = template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('學員報到系統');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function doGet(e) {
  // 必須設定 ALLOWALL，GitHub Pages 才能用 iframe 嵌入它
  return HtmlService.createHtmlOutputFromFile('Bridge')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * 優化後的學員搜尋：僅搜尋關鍵欄位，不載入全表
 * @param {string} query
 * @returns {Map<string, any> | null}
 */
function findStudentRow(query) {
  if (!query) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('學員名單');
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow <= 1) return null; // 只有標題或空表

  // 1. 取得標題行（用於後續對應欄位名稱）
  const title = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const idIndex = title.indexOf('學員代號') + 1; // getRange 是 1-based
  const phoneIndex = title.indexOf('手機') + 1;

  // 2. 使用 TextFinder 鎖定特定欄位進行全文匹配
  // 我們先找「學員代號」，沒找到再找「手機」
  const searchIndices = [idIndex, phoneIndex].filter(idx => idx > 0);
  let foundRange = null;

  for (let colIdx of searchIndices) {
    const searchRange = sheet.getRange(2, colIdx, lastRow - 1, 1);
    foundRange = searchRange.createTextFinder(query)
      .matchEntireCell(true) // 精確匹配
      .findNext();

    if (foundRange) break;
  }

  // 3. 如果找到了，只讀取那一行的資料
  if (foundRange) {
    const rowIndex = foundRange.getRow();
    const rowValues = sheet.getRange(rowIndex, 1, 1, lastColumn).getValues()[0];

    return new Map(rowValues.map((value, index) => [title[index], value]));
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
 * @returns {{success: boolean, message: string}}
 */
function checkInStudent(studentId) {
  const lock = LockService.getDocumentLock();
  try {
    let student = findStudentRow(studentId);
    if (!student) {
      return { success: false, message: '查無此學員' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('學員報到');

    // Acquire lock to avoid race conditions
    lock.waitLock(10000);

    // Fast and safe approach: appendRow automatically adds data right after the last row
    // and handles sheet expansion without having to fetch the entire sheet.
    sheet.appendRow([`'${studentId}`, new Date()]);

    lock.releaseLock();
    let msgTemplate = getMsgTemplate();
    return { success: true, message: msgFormat(msgTemplate, student) };
  } catch (e) {
    return { success: false, message: '報到失敗: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function int2Col(col) {
  let result = '';
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

function backupFomula() {
  const backup = [];
  const cells = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sheet => {
    let formulaArray = sheet.getDataRange().getFormulas();
    let formulas = [];
    for (let i = 0; i < formulaArray.length; i++) {
      for (let j = 0; j < formulaArray[i].length; j++) {
        if (formulaArray[i][j] !== '') {
          formulas.push({ row: i + 1, col: j + 1, formula: formulaArray[i][j] });
          cells.push(`${sheet.getName()}!${int2Col(j + 1)}${i + 1}`);
        }
      }
    }
    backup.push({ sheetName: sheet.getName(), formulas: formulas });
  })
  PropertiesService.getScriptProperties().setProperty('backup', JSON.stringify(backup));
  const ui = SpreadsheetApp.getUi();
  ui.alert('備份成功:\n' + cells.join('\n'));
}

function restoreFomula() {
  const ui = SpreadsheetApp.getUi();
  const backup = PropertiesService.getScriptProperties().getProperty('backup');
  const cells = [];
  if (!backup) {
    ui.alert('沒有備份資料');
    return;
  }
  const backupData = JSON.parse(backup);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  backupData.forEach(item => {
    const sheet = ss.getSheetByName(item.sheetName);
    if (sheet) {
      item.formulas.forEach(formula => {
        sheet.getRange(formula.row, formula.col).setFormula(formula.formula);
        cells.push(`${sheet.getName()}!${int2Col(formula.col)}${formula.row}`);
      })
    }
  })
  ui.alert('還原成功:\n' + cells.join('\n'));
}

let CACHED_TEMPLATE = null;

function getMsgTemplate() {
  if (CACHED_TEMPLATE !== null) return CACHED_TEMPLATE;

  const cache = CacheService.getScriptCache();
  let jsonTemplate = cache.get('msgTemplate');

  if (!jsonTemplate) {
    jsonTemplate = PropertiesService.getScriptProperties().getProperty('msgTemplate');
    if (jsonTemplate) {
      cache.put('msgTemplate', jsonTemplate, 21600); // 快取 6 小時
    }
  }

  if (!jsonTemplate) {
    CACHED_TEMPLATE = '報到成功；姓名：{{姓名}}，序號：{{序號}}';
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
  PropertiesService.getScriptProperties().setProperty('msgTemplate', jsonTemplate);
  CacheService.getScriptCache().put('msgTemplate', jsonTemplate, 21600);
  CACHED_TEMPLATE = template;
  return { success: true, message: '訊息模板已更新' };
}