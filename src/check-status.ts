import { google } from 'googleapis';
import * as fs from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';
import { preflight, findCarDataFolder } from './lib/preflight';

dotenv.config();

async function main() {
  const auth = await preflight();
  const sheets = google.sheets({ version: 'v4', auth });
  const spreadsheetId = process.env.SPREADSHEET_ID!;
  const carPromptsBase = process.env.CAR_PROMPTS_PATH || '';
  const carDataResult = findCarDataFolder(carPromptsBase);
  const CAR_PROMPTS_PATH = carDataResult.path;

  // 讀取整合庫存
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: '整合庫存!A1:Q600',
  });

  const rows = res.data.values || [];
  const headers = rows[0];
  const data = rows.slice(1);

  // 欄位索引
  const statusIdx = headers.indexOf('狀態');
  const brandIdx = headers.indexOf('Brand');
  const yearIdx = headers.indexOf('年式');
  const modelIdx = headers.indexOf('Model');

  // 統計狀態
  const statusCount: { [key: string]: number } = {};
  data.forEach(row => {
    const status = row[statusIdx] || '(空)';
    statusCount[status] = (statusCount[status] || 0) + 1;
  });

  console.log('=== 狀態統計 ===\n');
  Object.entries(statusCount).sort((a, b) => b[1] - a[1]).forEach(([status, count]) => {
    console.log(`${status}: ${count} 筆`);
  });

  // 庫存 = 非已售出
  const inStock = data.filter(row => {
    const status = row[statusIdx] || '';
    return status !== '已售出' && status !== '特殊' && row[brandIdx];
  });

  console.log(`\n=== 庫存統計 ===`);
  console.log(`庫存車輛: ${inStock.length} 台`);

  // 檢查本地資料夾
  const localFolders = fs.existsSync(CAR_PROMPTS_PATH)
    ? fs.readdirSync(CAR_PROMPTS_PATH).filter(f =>
        fs.statSync(path.join(CAR_PROMPTS_PATH, f)).isDirectory()
      )
    : [];

  console.log(`本地 PO 資料夾: ${localFolders.length} 個`);
  console.log('\n本地資料夾:');
  localFolders.forEach(f => console.log(`  - ${f}`));

  // 嘗試匹配
  let poCount = 0;
  const notPoList: string[] = [];

  inStock.forEach(row => {
    const brand = row[brandIdx] || '';
    const year = row[yearIdx] || '';
    const model = row[modelIdx] || '';

    // 嘗試匹配本地資料夾
    const matched = localFolders.find(f => {
      const fLower = f.toLowerCase();
      const brandLower = brand.toLowerCase();
      const modelLower = model.toLowerCase();
      return fLower.includes(year) &&
             (fLower.includes(brandLower) || brandLower.includes(fLower.split(' ')[1]?.toLowerCase() || '')) &&
             (fLower.includes(modelLower.split(' ')[0]?.toLowerCase() || ''));
    });

    if (matched) {
      poCount++;
    } else {
      notPoList.push(`${year} ${brand} ${model}`);
    }
  });

  console.log(`\n=== PO 統計 ===`);
  console.log(`已 PO: ${poCount} 台`);
  console.log(`未 PO: ${inStock.length - poCount} 台`);

  console.log(`\n=== 未 PO 車輛清單（庫存中）===\n`);
  notPoList.slice(0, 30).forEach((car, i) => {
    console.log(`${i + 1}. ${car}`);
  });
  if (notPoList.length > 30) {
    console.log(`... 還有 ${notPoList.length - 30} 台`);
  }
}

main().catch(console.error);
