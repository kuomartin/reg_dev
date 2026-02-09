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
  const url = ScriptApp.getService().getUrl();
  PropertiesService.getScriptProperties().setProperty('url', url);
  const template = HtmlService.createTemplateFromFile('index.html');
  template.url = url;
  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('學員報到系統');
}
/**
 * @param {string} query
 * @returns {Map<string, any> | null}
 */
function findStudentRow(query) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('學員名單');
  let data = sheet.getDataRange().getValues();
  let title = data.shift();
  let idIndex = title.indexOf('學員代號');
  let phoneIndex = title.indexOf('手機');

  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let id = row[idIndex];
    let phone = row[phoneIndex];
    if (query === id || query === phone) {
      return new Map(row.map((value, index) => [title[index], value]));
    }
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

    // 準備寫入的資料 (二維陣列形式)
    let rowData = [[`'${studentId}`, new Date()]];

    // 取得鎖定，確保在高併發時不會覆蓋資料
    lock.waitLock(10000);

    // 1. 找到最後一列資料的位置
    let maxRows = sheet.getMaxRows();
    let emptyRows = sheet.getRange(1, 1, maxRows).getValues().reverse().findIndex(c => c != '');
    const targetRow = maxRows - emptyRows + 1;

    // 2. 判斷空間是否足夠，不夠則新增列
    if (targetRow > maxRows) {
      sheet.insertRowsAfter(maxRows, 1);
    }

    // 3. 在資料範圍的下一列寫入
    // getRange(row, column, numRows, numColumns)
    sheet.getRange(targetRow, 1, 1, rowData[0].length).setValues(rowData);

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

function getMsgTemplate() {
  const jsonTemplate = PropertiesService.getScriptProperties().getProperty('msgTemplate');
  if (!jsonTemplate) return '報到成功；姓名：{{姓名}}，序號：{{序號}}';
  try {
    return JSON.parse(jsonTemplate);
  } catch (e) {
    return jsonTemplate;
  }
}

function updateMsgTemplate(template) {
  const jsonTemplate = JSON.stringify(template);
  PropertiesService.getScriptProperties().setProperty('msgTemplate', jsonTemplate);
  return { success: true, message: '訊息模板已更新' };
}