import { google } from 'googleapis';
import * as fs from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';
import { preflight, findCarDataFolder } from './lib/preflight';

dotenv.config();

// 整合庫存表欄位索引
const COL = {
  ITEM: 0,
  SOURCE: 1,
  BRAND: 2,
  YEAR: 3,
  MFG_DATE: 4,
  MILEAGE: 5,
  MODEL: 6,
  VIN: 7,
  CONDITION: 8,
  STATUS: 9,
  EXT_COLOR: 10,
  INT_COLOR: 11,
  MOD: 12,
  PO_STATUS: 13,
  OWNER: 14,
  PRICE: 15,
  NOTES: 16,
};

interface CarData {
  item: string;
  source: string;
  brand: string;
  year: string;
  mfgDate: string;
  mileage: string;
  model: string;
  vin: string;
  condition: string;
  status: string;
  extColor: string;
  intColor: string;
  modification: string;
  poStatus: string;
  owner: string;
  price: string;
  notes: string;
}

/**
 * 清理資料夾名稱中的非法字元
 */
function sanitizeFolderName(name: string): string {
  return name
    .replace(/[<>:"/\\|?*]/g, '') // 移除 Windows 非法字元
    .replace(/\s+/g, ' ')          // 多個空格變單一空格
    .trim();
}

/**
 * 產生資料夾名稱
 * 格式：【流水編號】.【年份】"【品牌 車型】 車色-內裝 業務
 * 範例：P11.2023" Bentley GT V8 S 灰-灰白 Mita
 */
function generateFolderName(car: CarData): string {
  // 流水編號.年份"
  let folderName = `${car.item}.${car.year}" ${car.brand} ${car.model}`;

  // 車色-內裝
  const colors: string[] = [];
  if (car.extColor && car.extColor.trim()) {
    colors.push(car.extColor.trim());
  }
  if (car.intColor && car.intColor.trim()) {
    colors.push(car.intColor.trim());
  }
  if (colors.length > 0) {
    folderName += ` ${colors.join('-')}`;
  }

  // 業務
  if (car.owner && car.owner.trim()) {
    folderName += ` ${car.owner.trim()}`;
  }

  return sanitizeFolderName(folderName);
}

/**
 * 產生車輛基本資料 Markdown
 */
function generateCarMarkdown(car: CarData): string {
  const lines = [
    `# ${car.year} ${car.brand} ${car.model}`,
    '',
    '## 基本資料',
    '',
    '| 欄位 | 資料 |',
    '|------|------|',
    `| 編號 | ${car.item} |`,
    `| 來源 | ${car.source} |`,
    `| 品牌 | ${car.brand} |`,
    `| 年式 | ${car.year} |`,
    `| 出廠年月 | ${car.mfgDate || '-'} |`,
    `| 里程 | ${car.mileage || '-'} |`,
    `| 車型 | ${car.model} |`,
    `| 引擎碼 | ${car.vin || '-'} |`,
    `| 車況 | ${car.condition || '-'} |`,
    `| 狀態 | ${car.status || '-'} |`,
    `| 外觀色 | ${car.extColor || '-'} |`,
    `| 內裝色 | ${car.intColor || '-'} |`,
    `| 改裝 | ${car.modification || '-'} |`,
    `| 負責人 | ${car.owner || '-'} |`,
    '',
  ];

  // 如果有備註，加入備註區塊
  if (car.notes && car.notes.trim()) {
    lines.push('## 備註');
    lines.push('');
    lines.push(car.notes);
    lines.push('');
  }

  // 加入 JSON 格式（供 post-helper 使用）
  lines.push('## JSON 資料');
  lines.push('');
  lines.push('```json');
  lines.push(JSON.stringify({
    item: car.item,
    brand: car.brand,
    year: car.year,
    model: car.model,
    vin: car.vin,
    mileage: car.mileage,
    condition: car.condition,
    extColor: car.extColor,
    intColor: car.intColor,
    modification: car.modification,
    owner: car.owner,
  }, null, 2));
  lines.push('```');
  lines.push('');

  return lines.join('\n');
}

async function main() {
  const auth = await preflight();
  const carPromptsBase = process.env.CAR_PROMPTS_PATH || '';
  const carDataResult = findCarDataFolder(carPromptsBase);
  const CAR_PROMPTS_PATH = carDataResult.path;

  const sheets = google.sheets({ version: 'v4', auth });
  const spreadsheetId = process.env.SPREADSHEET_ID!;

  console.log('讀取整合庫存表...');

  // 讀取整合庫存
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: '整合庫存!A1:Q600',
  });

  const rows = res.data.values || [];
  const headers = rows[0];
  const data = rows.slice(1);

  console.log(`總共 ${data.length} 筆資料`);

  // 篩選未 PO 且在庫的車輛
  const unpostedCars: CarData[] = [];

  data.forEach((row, index) => {
    const status = row[COL.STATUS] || '';
    const poStatus = row[COL.PO_STATUS] || '';
    const brand = row[COL.BRAND] || '';

    // 跳過已售出、特殊、沒有品牌的
    if (status === '已售出' || status === '特殊' || !brand) {
      return;
    }

    // 只處理未 PO 的
    if (poStatus !== '未PO' && poStatus !== '') {
      return;
    }

    const car: CarData = {
      item: row[COL.ITEM] || '',
      source: row[COL.SOURCE] || '',
      brand: row[COL.BRAND] || '',
      year: row[COL.YEAR] || '',
      mfgDate: row[COL.MFG_DATE] || '',
      mileage: row[COL.MILEAGE] || '',
      model: row[COL.MODEL] || '',
      vin: row[COL.VIN] || '',
      condition: row[COL.CONDITION] || '',
      status: row[COL.STATUS] || '',
      extColor: row[COL.EXT_COLOR] || '',
      intColor: row[COL.INT_COLOR] || '',
      modification: row[COL.MOD] || '',
      poStatus: row[COL.PO_STATUS] || '',
      owner: row[COL.OWNER] || '',
      price: row[COL.PRICE] || '',
      notes: row[COL.NOTES] || '',
    };

    unpostedCars.push(car);
  });

  console.log(`\n找到 ${unpostedCars.length} 台未 PO 車輛\n`);

  // 確保目標資料夾存在
  if (!fs.existsSync(CAR_PROMPTS_PATH)) {
    fs.mkdirSync(CAR_PROMPTS_PATH, { recursive: true });
  }

  // 為每台車建立資料夾
  let created = 0;
  let skipped = 0;

  for (const car of unpostedCars) {
    const folderName = generateFolderName(car);
    const folderPath = path.join(CAR_PROMPTS_PATH, folderName);

    // 檢查資料夾是否已存在
    if (fs.existsSync(folderPath)) {
      console.log(`⏭️  已存在: ${folderName}`);
      skipped++;
      continue;
    }

    // 建立資料夾
    fs.mkdirSync(folderPath, { recursive: true });

    // 建立車輛資料 MD 檔
    const mdContent = generateCarMarkdown(car);
    const mdPath = path.join(folderPath, 'car-info.md');
    fs.writeFileSync(mdPath, mdContent, 'utf-8');

    console.log(`✅ 建立: ${folderName}`);
    created++;
  }

  console.log(`\n========== 完成 ==========`);
  console.log(`新建立: ${created} 個資料夾`);
  console.log(`已存在: ${skipped} 個資料夾`);
  console.log(`總共: ${unpostedCars.length} 台未 PO 車輛`);
}

main().catch(console.error);
