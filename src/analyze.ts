import { google } from 'googleapis';
import * as fs from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';

dotenv.config();

const TOKEN_PATH = path.join(__dirname, '..', 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, '..', 'credentials.json');

async function authorize() {
  const content = fs.readFileSync(CREDENTIALS_PATH);
  const credentials = JSON.parse(content.toString());
  const { client_secret, client_id } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, 'urn:ietf:wg:oauth:2.0:oob');
  const token = fs.readFileSync(TOKEN_PATH);
  oAuth2Client.setCredentials(JSON.parse(token.toString()));
  return oAuth2Client;
}

async function main() {
  const auth = await authorize();
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
