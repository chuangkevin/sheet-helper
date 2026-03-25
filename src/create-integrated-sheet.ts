import { google } from 'googleapis';
import * as fs from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';
import { preflight } from './lib/preflight';

dotenv.config();

// 顏色判斷
function getColorType(bgColor: any): string {
  if (!bgColor) return '白色';
  const r = (bgColor.red || 0);
  const g = (bgColor.green || 0);
  const b = (bgColor.blue || 0);
  if (r > 0.95 && g > 0.95 && b > 0.95) return '白色';
  if (r > 0.99 && g < 0.01 && b > 0.99) return '紫紅';
  if (r > 0.99 && g > 0.9 && b > 0.7 && b < 0.85) return '淺黃';
  if (r > 0.7 && g < 0.5 && b < 0.5) return '紅色';
  if (r > 0.8 && g < 0.8 && b < 0.8 && r > g && r > b) return '淺紅';
  return '其他';
}

// 根據顏色判斷狀態
function getStatusFromColor(color: string, textStatus: string): string {
  if (color === '紅色' || color === '淺紅') return '已售出';
  if (color === '淺黃') return '新到貨';
  if (color === '紫紅') return '特殊';
  if (textStatus === 'Sold') return '已售出';
  if (textStatus === '海運') return '海運中';
  if (textStatus === '驗車' || textStatus === '驗車完成') return '驗車中';
  return '在庫';
}

// 判斷來源
function getSourceType(item: string): string {
  if (/^B\d+/.test(item)) return '寄賣';
  if (/^T\d+/.test(item)) return '託售';
  if (/^A\d+/.test(item)) return '台灣車';
  if (/^P\d+/.test(item)) return '國外進口';
  if (/^\d+$/.test(item)) return '國外進口';
  return '其他';
}

// 解析內裝色，分離顏色和改裝
function parseInteriorColor(raw: string): { color: string; modification: string } {
  if (!raw) return { color: '', modification: '' };

  // 格式1: 顏色(改裝) - 例如 "黑(星空頂)"、"米(ACC/米勒)"
  const bracketMatch = raw.match(/^([^(（]+)[(（](.+)[)）]$/);
  if (bracketMatch) {
    return {
      color: bracketMatch[1].trim(),
      modification: bracketMatch[2].trim()
    };
  }

  // 格式2: 顏色 改裝 - 例如 "紅 Mansory"（英文改裝品牌前有空格）
  const spaceMatch = raw.match(/^([^\s]+)\s+(Mansory|V-Specification|BB版|BB|Bespok)(.*)$/i);
  if (spaceMatch) {
    return {
      color: spaceMatch[1].trim(),
      modification: (spaceMatch[2] + (spaceMatch[3] || '')).trim()
    };
  }

  // 格式3: 純顏色或雙色 - 例如 "黑"、"黑/白"、"米黑"
  return { color: raw.trim(), modification: '' };
}

interface CarRecord {
  item: string;
  source: string;
  brand: string;
  year: string;
  manufactureDate: string;
  mileage: string;
  model: string;
  vin: string;
  condition: string;
  status: string;
  exteriorColor: string;
  interiorColor: string;
  modification: string;
  note: string;
  poStatus: string;
  owner: string;
  price: string;
  bgColor: string;
}

async function main() {
  const auth = await preflight({ needCarPrompts: false });
  const sheets = google.sheets({ version: 'v4', auth });
  const spreadsheetId = process.env.SPREADSHEET_ID!;

  console.log('📊 開始讀取表格資料...\n');

  // 完整讀取車源表
  console.log('讀取車源表...');
  const sourceResp = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: ['車源!A1:Z1000'],
    includeGridData: true,
  });

  // 完整讀取庫存表
  console.log('讀取庫存表...');
  const invResp = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: ['庫存!A1:Z500'],
    includeGridData: true,
  });

  const sourceRows = sourceResp.data.sheets?.[0]?.data?.[0]?.rowData || [];
  const invRows = invResp.data.sheets?.[0]?.data?.[0]?.rowData || [];

  console.log(`車源表讀取: ${sourceRows.length} 行`);
  console.log(`庫存表讀取: ${invRows.length} 行`);

  // 解析所有車輛資料
  const allCars: CarRecord[] = [];
  const seenItems = new Set<string>();

  // 解析車源表
  console.log('\n解析車源表...');
  let sourceCount = 0;
  let modificationCount = 0;

  sourceRows.forEach((row: any, idx: number) => {
    if (idx < 2) return;
    const cells = row.values || [];
    const item = cells[0]?.formattedValue?.trim() || '';

    if (!item || item === 'item' || item === '台灣' || item === '寄賣') return;
    if (seenItems.has(item)) return;

    const bgColor = getColorType(cells[0]?.effectiveFormat?.backgroundColor);
    const textStatus = cells[9]?.formattedValue || '';

    // 解析內裝色和改裝
    const rawInterior = cells[12]?.formattedValue || '';
    const { color: interiorColor, modification } = parseInteriorColor(rawInterior);

    if (modification) modificationCount++;

    seenItems.add(item);
    sourceCount++;

    allCars.push({
      item,
      source: getSourceType(item),
      brand: cells[2]?.formattedValue || '',
      year: cells[3]?.formattedValue || '',
      manufactureDate: cells[4]?.formattedValue || '',
      mileage: cells[5]?.formattedValue || '',
      model: cells[6]?.formattedValue || '',
      vin: cells[7]?.formattedValue || '',
      condition: cells[8]?.formattedValue || '',
      status: getStatusFromColor(bgColor, textStatus),
      exteriorColor: cells[11]?.formattedValue || '',
      interiorColor,
      modification,
      note: cells[10]?.formattedValue || '',
      poStatus: '未PO',
      owner: '',
      price: '',
      bgColor,
    });
  });
  console.log(`車源表解析: ${sourceCount} 筆有效資料`);
  console.log(`有改裝資訊: ${modificationCount} 筆`);

  // 解析庫存表（補充/更新資訊）
  console.log('\n解析庫存表...');

  // 先找出「分配」欄位的索引
  const invHeader = invRows[0]?.values || [];
  let assignIdx = -1;
  invHeader.forEach((cell: any, i: number) => {
    const val = cell?.formattedValue?.trim() || '';
    if (val === '分配') {
      assignIdx = i;
      console.log(`找到「分配」欄位，索引: ${i}`);
    }
  });

  let invCount = 0;
  let updatedCount = 0;
  let ownerCount = 0;
  invRows.forEach((row: any, idx: number) => {
    if (idx < 1) return;
    const cells = row.values || [];
    const item = cells[0]?.formattedValue?.trim() || '';

    if (!item || item === 'item' || item === '台灣' || item === '寄賣') return;

    invCount++;

    const existing = allCars.find(c => c.item === item);
    if (existing) {
      // 更新車況
      const invCondition = cells[8]?.formattedValue || '';
      if (invCondition && invCondition !== '車況') {
        existing.condition = invCondition;
        updatedCount++;
      }
      // 更新外觀色
      const invExtColor = cells[11]?.formattedValue || '';
      if (invExtColor && !existing.exteriorColor) {
        existing.exteriorColor = invExtColor;
      }
      // 更新內裝色和改裝
      const rawInterior = cells[12]?.formattedValue || '';
      if (rawInterior) {
        const { color, modification } = parseInteriorColor(rawInterior);
        if (color && !existing.interiorColor) {
          existing.interiorColor = color;
        }
        if (modification && !existing.modification) {
          existing.modification = modification;
        }
      }
      // 更新備註
      const invNote = cells[10]?.formattedValue || '';
      if (invNote) {
        if (existing.note && existing.note !== invNote) {
          existing.note = existing.note + ' | ' + invNote;
        } else {
          existing.note = invNote;
        }
      }
      // 更新開價
      const invPrice = cells[17]?.formattedValue || '';
      if (invPrice) {
        existing.price = invPrice;
      }
      // 更新負責人（從「分配」欄位）
      if (assignIdx >= 0) {
        const owner = cells[assignIdx]?.formattedValue?.trim() || '';
        if (owner) {
          existing.owner = owner;
          ownerCount++;
        }
      }
    }
  });
  console.log(`庫存表解析: ${invCount} 筆，更新了 ${updatedCount} 筆車況`);
  if (assignIdx >= 0) {
    console.log(`更新了 ${ownerCount} 筆負責人`);
  } else {
    console.log('⚠️ 未找到「分配」欄位');
  }
  console.log(`庫存表解析: ${invCount} 筆，更新了 ${updatedCount} 筆車況`);

  // 統計
  const inStock = allCars.filter(c => c.status === '在庫' || c.status === '新到貨' || c.status === '驗車中' || c.status === '海運中');
  const sold = allCars.filter(c => c.status === '已售出');
  const special = allCars.filter(c => c.status === '特殊');
  const withMod = allCars.filter(c => c.modification);

  console.log(`\n📊 整合統計:`);
  console.log(`總車輛數: ${allCars.length}`);
  console.log(`在庫/新到貨/驗車中/海運中: ${inStock.length}`);
  console.log(`已售出: ${sold.length}`);
  console.log(`特殊: ${special.length}`);
  console.log(`有改裝資訊: ${withMod.length}`);

  // 創建新工作表
  console.log('\n📝 創建「整合庫存」工作表...');

  // 安全名單：只有這些是程式產生的工作表，可以刪除重建
  const SAFE_TO_DELETE = ['整合庫存'];
  // 保護名單：這些是使用者的原始資料，絕對不能刪
  const PROTECTED_SHEETS = ['車源', '庫存'];

  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const existingSheet = spreadsheet.data.sheets?.find(s => s.properties?.title === '整合庫存');

  if (existingSheet) {
    const sheetTitle = existingSheet.properties?.title || '';
    if (PROTECTED_SHEETS.includes(sheetTitle)) {
      console.error(`❌ 安全保護：不允許刪除「${sheetTitle}」工作表`);
      process.exit(1);
    }
    if (!SAFE_TO_DELETE.includes(sheetTitle)) {
      console.error(`❌ 安全保護：「${sheetTitle}」不在允許刪除的名單中`);
      process.exit(1);
    }
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          deleteSheet: { sheetId: existingSheet.properties?.sheetId }
        }]
      }
    });
    console.log('已刪除舊的「整合庫存」工作表');
  }

  const addSheetResp = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        addSheet: {
          properties: {
            title: '整合庫存',
            gridProperties: { rowCount: 1000, columnCount: 20 }
          }
        }
      }]
    }
  });

  const newSheetId = addSheetResp.data.replies?.[0]?.addSheet?.properties?.sheetId;
  console.log(`新工作表已創建，ID: ${newSheetId}`);

  // 準備資料 - 改裝獨立欄位
  const headers = [
    'item', '來源', 'Brand', '年式', '出廠年月', '里程', 'Model', '引擎碼(VIN)',
    '車況', '狀態', '外觀色', '內裝色', '改裝', 'PO狀態', '負責人', '開價', '備註'
  ];

  // 寫入全部車輛
  const dataToWrite = allCars.map(car => [
    car.item,
    car.source,
    car.brand,
    car.year,
    car.manufactureDate,
    car.mileage,
    car.model,
    car.vin,
    car.condition,
    car.status,
    car.exteriorColor,
    car.interiorColor,
    car.modification,
    car.poStatus,
    car.owner,
    car.price,
    car.note,
  ]);

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: '整合庫存!A1',
    valueInputOption: 'RAW',
    requestBody: {
      values: [headers, ...dataToWrite]
    }
  });

  console.log(`已寫入 ${dataToWrite.length} 筆車輛資料（全部）`);

  // 設定下拉選單
  console.log('\n🔽 設定下拉選單...');

  const validationRequests = [
    // 來源下拉 (B欄, index 1)
    {
      setDataValidation: {
        range: { sheetId: newSheetId, startRowIndex: 1, endRowIndex: 1000, startColumnIndex: 1, endColumnIndex: 2 },
        rule: {
          condition: { type: 'ONE_OF_LIST', values: [
            { userEnteredValue: '國外進口' },
            { userEnteredValue: '台灣車' },
            { userEnteredValue: '託售' },
            { userEnteredValue: '寄賣' },
          ]},
          showCustomUi: true
        }
      }
    },
    // 車況下拉 (I欄, index 8)
    {
      setDataValidation: {
        range: { sheetId: newSheetId, startRowIndex: 1, endRowIndex: 1000, startColumnIndex: 8, endColumnIndex: 9 },
        rule: {
          condition: { type: 'ONE_OF_LIST', values: [
            { userEnteredValue: '6/A' },
            { userEnteredValue: '5/A' },
            { userEnteredValue: '5/B' },
            { userEnteredValue: '4.5/A' },
            { userEnteredValue: '4.5/B' },
            { userEnteredValue: '4/A' },
            { userEnteredValue: '4/B' },
            { userEnteredValue: '4/C' },
          ]},
          showCustomUi: true
        }
      }
    },
    // 狀態下拉 (J欄, index 9)
    {
      setDataValidation: {
        range: { sheetId: newSheetId, startRowIndex: 1, endRowIndex: 1000, startColumnIndex: 9, endColumnIndex: 10 },
        rule: {
          condition: { type: 'ONE_OF_LIST', values: [
            { userEnteredValue: '在庫' },
            { userEnteredValue: '新到貨' },
            { userEnteredValue: '海運中' },
            { userEnteredValue: '驗車中' },
            { userEnteredValue: '已售出' },
            { userEnteredValue: '特殊' },
          ]},
          showCustomUi: true
        }
      }
    },
    // PO狀態下拉 (N欄, index 13)
    {
      setDataValidation: {
        range: { sheetId: newSheetId, startRowIndex: 1, endRowIndex: 1000, startColumnIndex: 13, endColumnIndex: 14 },
        rule: {
          condition: { type: 'ONE_OF_LIST', values: [
            { userEnteredValue: '未PO' },
            { userEnteredValue: '部分PO' },
            { userEnteredValue: '已PO' },
            { userEnteredValue: '不需PO' },
          ]},
          showCustomUi: true
        }
      }
    },
  ];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: validationRequests }
  });

  console.log('下拉選單設定完成');

  // 設定格式
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        // 凍結首行
        {
          updateSheetProperties: {
            properties: { sheetId: newSheetId, gridProperties: { frozenRowCount: 1 } },
            fields: 'gridProperties.frozenRowCount'
          }
        },
        // 標題列格式
        {
          repeatCell: {
            range: { sheetId: newSheetId, startRowIndex: 0, endRowIndex: 1 },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.2, green: 0.2, blue: 0.2 },
                textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } }
              }
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat)'
          }
        },
        // 自動調整欄寬
        {
          autoResizeDimensions: {
            dimensions: { sheetId: newSheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 17 }
          }
        }
      ]
    }
  });

  console.log('\n✅ 完成！請查看 Google Sheets 的「整合庫存」工作表');
  console.log(`📊 共 ${dataToWrite.length} 筆車輛（全部）`);
  console.log(`   - 在庫/新到貨/驗車中/海運中: ${inStock.length} 筆`);
  console.log(`   - 已售出: ${sold.length} 筆`);
  console.log(`   - 有改裝資訊: ${withMod.length} 筆`);
}

main().catch(console.error);
