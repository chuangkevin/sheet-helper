import { google } from 'googleapis';
import * as fs from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';
import { parseCarPrompts, CarData } from './carParser';

// Load environment variables
dotenv.config();

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = path.join(__dirname, '..', 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, '..', 'credentials.json');

// Load client secrets from a local file.
function loadCredentials(): Promise<any> {
  return new Promise((resolve, reject) => {
    fs.readFile(CREDENTIALS_PATH, (err, content) => {
      if (err) return reject('Error loading client secret file: ' + err);
      resolve(JSON.parse(content.toString()));
    });
  });
}

// Create an OAuth2 client with the given credentials
async function authorize() {
  const credentials = await loadCredentials();
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, 'urn:ietf:wg:oauth:2.0:oob');

  // Check if we have previously stored a token.
  try {
    const token = fs.readFileSync(TOKEN_PATH);
    oAuth2Client.setCredentials(JSON.parse(token.toString()));
    return oAuth2Client;
  } catch (err) {
    return await getNewToken(oAuth2Client);
  }
}

// Get and store new token after prompting for user authorization
async function getNewToken(oAuth2Client: any): Promise<any> {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });

  console.log('請在瀏覽器中開啟此 URL 進行授權:', authUrl);
  console.log('授權後，複製授權碼並貼上到這裡:');

  // For simplicity, let's use a manual approach
  const readline = require('readline');
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  return new Promise((resolve, reject) => {
    rl.question('請輸入授權碼: ', async (code: string) => {
      rl.close();
      try {
        const { tokens } = await oAuth2Client.getToken(code.trim());
        oAuth2Client.setCredentials(tokens);

        // Store the token to disk for later program executions
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
        console.log('Token 已儲存到', TOKEN_PATH);
        resolve(oAuth2Client);
      } catch (err) {
        reject(err);
      }
    });
  });
}

// Example function to read from sheet
async function readSheet(auth: any, spreadsheetId: string, range: string) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });
  return res.data.values;
}

// Example function to write to sheet
async function writeSheet(auth: any, spreadsheetId: string, range: string, values: any[][]) {
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: 'RAW',
    requestBody: { values },
  });
}

// Sync car data to sheet
async function syncCarInventory(auth: any, spreadsheetId: string, cars: CarData[]) {
  const headers = [['車輛名稱', '資料夾路徑', '有 Post Helper', '有 Facebook', '有 Yahoo', '有 Official', '有 8891']];
  const data = cars.map(car => [
    car.name,
    car.folder,
    car.postHelper ? '是' : '否',
    car.facebook ? '是' : '否',
    car.yahoo ? '是' : '否',
    car.official ? '是' : '否',
    car['8891'] ? '是' : '否',
  ]);

  const values = headers.concat(data);
  await writeSheet(auth, spreadsheetId, 'Inventory!A1:G', values);
  console.log('庫存已同步到 Google Sheets');
}

// Main function
async function main() {
  try {
    const auth = await authorize();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    if (!spreadsheetId) {
      console.error('請在 .env 中設定 SPREADSHEET_ID');
      return;
    }

    const carPromptsPath = process.env.CAR_PROMPTS_PATH || 'D:\\Projects\\car-prompts';
    const cars = parseCarPrompts(carPromptsPath);
    console.log(`找到 ${cars.length} 輛車`);

    // Sync to sheet
    await syncCarInventory(auth, spreadsheetId, cars);

  } catch (error) {
    console.error('錯誤:', error);
  }
}

main();