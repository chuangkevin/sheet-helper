import { google } from 'googleapis';
import * as fs from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';

dotenv.config();

const TOKEN_PATH = path.join(__dirname, '..', 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, '..', 'credentials.json');

async function main() {
  const content = fs.readFileSync(CREDENTIALS_PATH);
  const credentials = JSON.parse(content.toString());
  const { client_secret, client_id } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, 'urn:ietf:wg:oauth:2.0:oob');
  const token = fs.readFileSync(TOKEN_PATH);
  oAuth2Client.setCredentials(JSON.parse(token.toString()));

  const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
  const spreadsheetId = process.env.SPREADSHEET_ID!;

  // 取得整合庫存的 sheetId
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = spreadsheet.data.sheets?.find(s => s.properties?.title === '整合庫存');

  if (!sheet) {
    console.log('找不到「整合庫存」工作表');
    return;
  }

  const sheetId = sheet.properties?.sheetId;
  console.log('調整欄寬中...');

  // 手動設定每個欄位的寬度（像素）
  const columnWidths = [
    { col: 0, width: 60 },   // item
    { col: 1, width: 80 },   // 來源
    { col: 2, width: 110 },  // Brand
    { col: 3, width: 50 },   // 年式
    { col: 4, width: 85 },   // 出廠年月
    { col: 5, width: 100 },  // 里程
    { col: 6, width: 220 },  // Model
    { col: 7, width: 190 },  // 引擎碼(VIN)
    { col: 8, width: 55 },   // 車況
    { col: 9, width: 70 },   // 狀態
    { col: 10, width: 70 },  // 外觀色
    { col: 11, width: 70 },  // 內裝色
    { col: 12, width: 280 }, // 改裝
    { col: 13, width: 75 },  // PO狀態
    { col: 14, width: 70 },  // 負責人
    { col: 15, width: 70 },  // 開價
    { col: 16, width: 350 }, // 備註
  ];

  const requests = columnWidths.map(({ col, width }) => ({
    updateDimensionProperties: {
      range: {
        sheetId,
        dimension: 'COLUMNS',
        startIndex: col,
        endIndex: col + 1
      },
      properties: { pixelSize: width },
      fields: 'pixelSize'
    }
  }));

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests }
  });

  console.log('✅ 欄寬調整完成！所有資料應該都能完整顯示了。');
}

main().catch(console.error);
