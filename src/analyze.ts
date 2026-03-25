import { google } from 'googleapis';
import * as fs from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';
import { preflight } from './lib/preflight';

dotenv.config();

async function main() {
  const auth = await preflight({ needCarPrompts: false });
  const sheets = google.sheets({ version: 'v4', auth });
  const spreadsheetId = process.env.SPREADSHEET_ID!;

  // 讀取車源表
  const sourceResp = await sheets.spreadsheets.get({
    spreadsheetId,
    ranges: ['車源!A1:Z600'],
    includeGridData: true,
  });

  const sourceRows = sourceResp.data.sheets?.[0]?.data?.[0]?.rowData || [];

  // 收集所有內裝色欄位的值
  console.log('=== 內裝色欄位分析 ===\n');

  const interiorColors: string[] = [];

  sourceRows.forEach((row: any, idx: number) => {
    if (idx < 2) return;
    const cells = row.values || [];
    const item = cells[0]?.formattedValue?.trim() || '';
    if (!item || item === 'item') return;

    const interiorColor = cells[12]?.formattedValue || '';
    if (interiorColor) {
      interiorColors.push(interiorColor);
    }
  });

  // 找出不重複的值
  const uniqueColors = [...new Set(interiorColors)];
  console.log(`共 ${uniqueColors.length} 種不同的內裝色值：\n`);

  uniqueColors.forEach((color, i) => {
    console.log(`${i + 1}. "${color}"`);
  });
}

main().catch(console.error);
