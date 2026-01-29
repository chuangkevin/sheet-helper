/**
 * 車輛異動表單處理 - Apps Script
 * 綁定到 Google Form 的 onFormSubmit 觸發器
 *
 * 使用方式：
 * 1. 在 Google Sheet 中點選「擴充功能 → Apps Script」
 * 2. 貼上此程式碼
 * 3. 設定觸發器：onFormSubmit, 從試算表, 提交表單時
 */

// 設定 - 欄位索引（0-indexed）
const CONFIG = {
  SHEET_NAME: '整合庫存',
  ITEM_COL: 0,      // A欄 - item
  SOURCE_COL: 1,    // B欄 - 來源
  BRAND_COL: 2,     // C欄 - Brand
  YEAR_COL: 3,      // D欄 - 年式
  MFG_DATE_COL: 4,  // E欄 - 出廠年月
  MILEAGE_COL: 5,   // F欄 - 里程
  MODEL_COL: 6,     // G欄 - Model
  VIN_COL: 7,       // H欄 - 引擎碼
  CONDITION_COL: 8, // I欄 - 車況
  STATUS_COL: 9,    // J欄 - 狀態
  EXT_COLOR_COL: 10,// K欄 - 外觀色
  INT_COLOR_COL: 11,// L欄 - 內裝色
  MOD_COL: 12,      // M欄 - 改裝
  PO_STATUS_COL: 13,// N欄 - PO狀態
  OWNER_COL: 14,    // O欄 - 負責人
  PRICE_COL: 15,    // P欄 - 開價
  NOTES_COL: 16,    // Q欄 - 備註
};

/**
 * 表單提交時觸發
 * @param {Object} e - 表單提交事件
 */
function onFormSubmit(e) {
  const responses = e.namedValues;
  const actionType = responses['異動類型'][0];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    Logger.log('找不到工作表: ' + CONFIG.SHEET_NAME);
    return;
  }

  switch (actionType) {
    case '新車入庫':
      handleNewCar(sheet, responses);
      break;
    case '車輛售出':
      handleCarSold(sheet, responses);
      break;
    case '車輛退車':
      handleCarReturn(sheet, responses);
      break;
    default:
      Logger.log('未知的異動類型: ' + actionType);
  }
}

/**
 * 處理新車入庫
 * @param {Sheet} sheet - 目標工作表
 * @param {Object} responses - 表單回應
 */
function handleNewCar(sheet, responses) {
  // 取得下一個 item 編號
  const source = responses['來源'] ? responses['來源'][0] : '國外進口';
  const nextItem = getNextItemNumber(sheet, source);

  // 準備新行資料（17 欄）
  const newRow = new Array(17).fill('');
  newRow[CONFIG.ITEM_COL] = nextItem;
  newRow[CONFIG.SOURCE_COL] = source;
  newRow[CONFIG.BRAND_COL] = responses['Brand'] ? responses['Brand'][0] : '';
  newRow[CONFIG.YEAR_COL] = responses['年式'] ? responses['年式'][0] : '';
  newRow[CONFIG.MFG_DATE_COL] = responses['出廠年月'] ? responses['出廠年月'][0] : '';
  newRow[CONFIG.MILEAGE_COL] = responses['里程'] ? responses['里程'][0] : '';
  newRow[CONFIG.MODEL_COL] = responses['Model'] ? responses['Model'][0] : '';
  newRow[CONFIG.VIN_COL] = responses['引擎碼'] ? responses['引擎碼'][0] : '';
  newRow[CONFIG.CONDITION_COL] = responses['車況'] ? responses['車況'][0] : '';
  newRow[CONFIG.STATUS_COL] = '新到貨';
  newRow[CONFIG.EXT_COLOR_COL] = responses['外觀色'] ? responses['外觀色'][0] : '';
  newRow[CONFIG.INT_COLOR_COL] = responses['內裝色'] ? responses['內裝色'][0] : '';
  newRow[CONFIG.MOD_COL] = responses['改裝'] ? responses['改裝'][0] : '';
  newRow[CONFIG.PO_STATUS_COL] = '未PO';
  newRow[CONFIG.OWNER_COL] = responses['負責人'] ? responses['負責人'][0] : '';
  newRow[CONFIG.PRICE_COL] = responses['開價'] ? responses['開價'][0] : '';
  newRow[CONFIG.NOTES_COL] = responses['備註'] ? responses['備註'][0] : '';

  // 新增到表格最後
  sheet.appendRow(newRow);

  // 設定背景色為淺黃色（新到貨）
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 17).setBackground('#fff2cc');

  Logger.log('新增車輛: ' + nextItem);
}

/**
 * 處理車輛售出
 * @param {Sheet} sheet - 目標工作表
 * @param {Object} responses - 表單回應
 */
function handleCarSold(sheet, responses) {
  const itemNumber = responses['Item編號'] ? responses['Item編號'][0] : '';
  if (!itemNumber) {
    Logger.log('缺少 Item 編號');
    return;
  }

  const row = findRowByItem(sheet, itemNumber);
  if (row === -1) {
    Logger.log('找不到車輛: ' + itemNumber);
    return;
  }

  // 更新狀態為已售出
  sheet.getRange(row, CONFIG.STATUS_COL + 1).setValue('已售出');

  // 組合備註
  const currentNotes = sheet.getRange(row, CONFIG.NOTES_COL + 1).getValue();
  const soldDate = responses['售出日期'] ? responses['售出日期'][0] : new Date().toLocaleDateString('zh-TW');
  const soldPrice = responses['實際售價'] ? responses['實際售價'][0] : '';
  const soldNote = responses['售出備註'] ? responses['售出備註'][0] : '';

  let newNote = `[${soldDate}售出`;
  if (soldPrice) newNote += ` $${soldPrice}`;
  newNote += ']';
  if (soldNote) newNote += ` ${soldNote}`;

  sheet.getRange(row, CONFIG.NOTES_COL + 1).setValue(
    currentNotes ? currentNotes + '; ' + newNote : newNote
  );

  // 設定背景色為紅色（已售出）
  sheet.getRange(row, 1, 1, 17).setBackground('#ffcdd2');

  Logger.log('車輛售出: ' + itemNumber);
}

/**
 * 處理車輛退車
 * @param {Sheet} sheet - 目標工作表
 * @param {Object} responses - 表單回應
 */
function handleCarReturn(sheet, responses) {
  const itemNumber = responses['Item編號'] ? responses['Item編號'][0] : '';
  if (!itemNumber) {
    Logger.log('缺少 Item 編號');
    return;
  }

  const row = findRowByItem(sheet, itemNumber);
  if (row === -1) {
    Logger.log('找不到車輛: ' + itemNumber);
    return;
  }

  // 更新狀態為特殊（退車）
  sheet.getRange(row, CONFIG.STATUS_COL + 1).setValue('特殊');

  // 附加退車原因到備註
  const currentNotes = sheet.getRange(row, CONFIG.NOTES_COL + 1).getValue();
  const returnReason = responses['退車原因'] ? responses['退車原因'][0] : '未說明';
  const returnNote = `[退車] ${returnReason}`;

  sheet.getRange(row, CONFIG.NOTES_COL + 1).setValue(
    currentNotes ? currentNotes + '; ' + returnNote : returnNote
  );

  // 設定背景色為紫色（特殊）
  sheet.getRange(row, 1, 1, 17).setBackground('#e1bee7');

  Logger.log('車輛退車: ' + itemNumber);
}

/**
 * 根據 item 編號找到對應的行號
 * @param {Sheet} sheet - 工作表
 * @param {string} itemNumber - 車輛編號
 * @returns {number} 行號（從1開始），找不到則回傳 -1
 */
function findRowByItem(sheet, itemNumber) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.ITEM_COL]).trim() === String(itemNumber).trim()) {
      return i + 1; // Sheet 行號從 1 開始
    }
  }
  return -1;
}

/**
 * 根據來源取得下一個 item 編號
 * @param {Sheet} sheet - 工作表
 * @param {string} source - 來源類型
 * @returns {string} 新的 item 編號
 */
function getNextItemNumber(sheet, source) {
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  let prefix = '';

  // 根據來源決定前綴
  switch (source) {
    case '國外進口':
      prefix = ''; // 純數字
      break;
    case '台灣車':
      prefix = 'A';
      break;
    case '託售':
      prefix = 'T';
      break;
    case '寄賣':
      prefix = 'B';
      break;
    default:
      prefix = '';
  }

  // 找出該前綴的最大編號
  for (let i = 1; i < data.length; i++) {
    const item = String(data[i][CONFIG.ITEM_COL]).trim();
    if (prefix === '') {
      // 純數字（排除 P 開頭）
      if (/^\d+$/.test(item)) {
        const num = parseInt(item);
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    } else if (item.toUpperCase().startsWith(prefix)) {
      const num = parseInt(item.substring(prefix.length));
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }

  return prefix + (maxNum + 1);
}

/**
 * 手動測試 - 新車入庫
 */
function testNewCar() {
  const mockEvent = {
    namedValues: {
      '異動類型': ['新車入庫'],
      '來源': ['國外進口'],
      'Brand': ['Ferrari'],
      '年式': ['2024'],
      '出廠年月': ['2024/01'],
      '里程': ['500'],
      'Model': ['296 GTB'],
      '引擎碼': ['V6 3.0L Hybrid'],
      '車況': ['6/A'],
      '外觀色': ['紅'],
      '內裝色': ['黑'],
      '改裝': [''],
      '負責人': ['Kevin'],
      '開價': ['1680'],
      '備註': ['測試資料']
    }
  };
  onFormSubmit(mockEvent);
}

/**
 * 手動測試 - 車輛售出
 */
function testCarSold() {
  const mockEvent = {
    namedValues: {
      '異動類型': ['車輛售出'],
      'Item編號': ['123'],
      '售出日期': ['2024/01/15'],
      '實際售價': ['1580'],
      '售出備註': ['客戶滿意']
    }
  };
  onFormSubmit(mockEvent);
}

/**
 * 手動測試 - 車輛退車
 */
function testCarReturn() {
  const mockEvent = {
    namedValues: {
      '異動類型': ['車輛退車'],
      'Item編號': ['456'],
      '退車原因': ['車況不符預期']
    }
  };
  onFormSubmit(mockEvent);
}
