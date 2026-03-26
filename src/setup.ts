import * as fs from 'fs';
import * as path from 'path';
import {
  checkStatus,
  authorize,
  parseSpreadsheetUrl,
  validateCredentialsFile,
  findCarDataFolder,
  sanitizePath,
  writeEnvValue,
  openBrowser,
  CREDENTIALS_PATH,
  TOKEN_PATH,
  PROJECT_ROOT,
} from './lib/preflight';
import { google } from 'googleapis';

// ──────────────────────────────────────────────────────────────────
// CLI-driven setup — 所有操作都是子命令，不需要互動式輸入
//
// 用法：
//   npm run setup                              顯示目前設定狀態與下一步指引
//   npm run setup -- status                    同上（JSON 格式輸出）
//   npm run setup -- credentials <path>        設定 Google OAuth 憑證
//   npm run setup -- spreadsheet <url-or-id>   設定 Google Sheets
//   npm run setup -- car-prompts <path>        設定 car-prompts 路徑
//   npm run setup -- auth                      執行 OAuth 授權（開啟瀏覽器，自動接收回調）
//   npm run setup -- guide                     顯示完整設定教學
//   npm run setup -- reset-auth                清除授權 token 並重新授權
// ──────────────────────────────────────────────────────────────────

const args = process.argv.slice(2);
const command = args[0] || '';
const commandArg = args.slice(1).join(' ');

async function main() {
  switch (command) {
    case 'credentials':
      await cmdCredentials(commandArg);
      break;
    case 'spreadsheet':
      await cmdSpreadsheet(commandArg);
      break;
    case 'car-prompts':
      await cmdCarPrompts(commandArg);
      break;
    case 'auth':
      await cmdAuth();
      break;
    case 'reset-auth':
      await cmdResetAuth();
      break;
    case 'guide':
      cmdGuide();
      break;
    case 'status':
      cmdStatusJson();
      break;
    default:
      cmdStatus();
      break;
  }
}

// ── credentials <path> ──────────────────────────────────────────
async function cmdCredentials(filePath: string) {
  if (!filePath) {
    console.error('❌ 請提供憑證 JSON 檔案路徑');
    console.error('用法：npm run setup -- credentials <path-to-json>');
    console.error('');
    console.error('取得方式：');
    console.error('  1. 前往 https://console.cloud.google.com/apis/credentials');
    console.error('  2. 建立 OAuth 2.0 用戶端 ID（類型選「桌面應用程式」）');
    console.error('  3. 下載 JSON 檔案');
    console.error('  4. 執行此命令並帶入檔案路徑');
    process.exit(1);
  }

  const cleaned = sanitizePath(filePath);

  if (!fs.existsSync(cleaned)) {
    console.error(`❌ 找不到檔案：${cleaned}`);
    process.exit(1);
  }

  if (!validateCredentialsFile(cleaned)) {
    console.error('❌ 檔案格式不正確');
    console.error('請確認下載的是「桌面應用程式」類型的 OAuth 憑證');
    console.error('（不是 API Key，也不是 Service Account）');
    process.exit(1);
  }

  fs.copyFileSync(cleaned, CREDENTIALS_PATH);
  console.log('✅ Google API 憑證已設定');

  // Show next step
  const status = checkStatus();
  printNextStep(status);
}

// ── spreadsheet <url-or-id> ─────────────────────────────────────
async function cmdSpreadsheet(input: string) {
  if (!input) {
    console.error('❌ 請提供 Google Sheets 網址或 ID');
    console.error('用法：npm run setup -- spreadsheet <url-or-id>');
    console.error('');
    console.error('範例：');
    console.error('  npm run setup -- spreadsheet "https://docs.google.com/spreadsheets/d/1ABC.../edit"');
    console.error('  npm run setup -- spreadsheet 1ABCxyz...');
    process.exit(1);
  }

  const id = parseSpreadsheetUrl(input.trim());
  if (!id) {
    console.error('❌ 無法解析 Spreadsheet ID');
    console.error('請提供完整的 Google Sheets 網址或 Spreadsheet ID');
    process.exit(1);
  }

  // Try to verify if we have credentials
  const status = checkStatus();
  if (status.credentialsValid && status.hasToken) {
    try {
      const auth = await authorize();
      const sheets = google.sheets({ version: 'v4', auth });
      const res = await sheets.spreadsheets.get({ spreadsheetId: id });

      const title = res.data.properties?.title || '(未命名)';
      const sheetNames = res.data.sheets?.map(s => s.properties?.title).filter(Boolean) || [];

      writeEnvValue('SPREADSHEET_ID', id);
      console.log(`✅ 已連結 Google Sheets：「${title}」`);
      console.log(`   工作表：${sheetNames.join('、')}`);
    } catch (err: any) {
      if (err?.code === 403 || err?.code === 401) {
        console.error('❌ 沒辦法打開這個表格');
        console.error('');
        console.error('可能的原因：');
        console.error('  1. 這個表格沒有分享給你 — 請打開表格，點右上角「共用」，把你自己的 Gmail 加進去');
        console.error('  2. 你剛剛登入的 Google 帳號不是表格的擁有者 — 請確認是同一個帳號');
        console.error('');
        console.error('修好之後，再執行一次同樣的指令就可以了。');
        process.exit(1);
      } else if (err?.code === 404) {
        console.error('❌ 找不到這個表格');
        console.error('請確認你貼的網址是正確的，打開那個網址應該要能看到表格。');
        process.exit(1);
      } else {
        // Save anyway, might be a network issue
        writeEnvValue('SPREADSHEET_ID', id);
        console.log(`⚠️ 無法驗證表格（網路問題？），已儲存 ID：${id}`);
      }
    }
  } else {
    // No auth yet, just save the ID
    writeEnvValue('SPREADSHEET_ID', id);
    console.log(`✅ 已儲存 Spreadsheet ID：${id}`);
    console.log('   （尚未授權，無法驗證表格是否存在）');
  }

  const newStatus = checkStatus();
  printNextStep(newStatus);
}

// ── car-prompts <path> ──────────────────────────────────────────
async function cmdCarPrompts(inputPath: string) {
  if (!inputPath) {
    console.error('❌ 請提供 car-prompts 專案路徑');
    console.error('用法：npm run setup -- car-prompts <path>');
    console.error('');
    console.error('範例：');
    console.error('  npm run setup -- car-prompts "D:\\Projects\\car-prompts"');
    console.error('');
    console.error('路徑下應包含「汽車資料」子資料夾');
    process.exit(1);
  }

  const cleaned = sanitizePath(inputPath);

  if (!fs.existsSync(cleaned)) {
    console.error(`❌ 找不到路徑：${cleaned}`);
    process.exit(1);
  }

  const result = findCarDataFolder(cleaned);

  if (!result.found) {
    console.error('❌ 在這個路徑下找不到「汽車資料」資料夾');
    console.error('car-prompts 專案裡面應該要有一個叫「汽車資料」的子資料夾');
    process.exit(1);
  }

  writeEnvValue('CAR_PROMPTS_PATH', cleaned);
  console.log(`✅ 已設定 car-prompts 路徑`);
  console.log(`   汽車資料位置：${result.path}`);
  console.log(`   找到 ${result.count} 台車的資料`);

  const newStatus = checkStatus();
  printNextStep(newStatus);
}

// ── auth ────────────────────────────────────────────────────────
async function cmdAuth() {
  const status = checkStatus();

  if (!status.credentialsValid) {
    console.error('❌ 請先設定 Google API 憑證');
    console.error('執行：npm run setup -- credentials <path-to-json>');
    process.exit(1);
  }

  if (status.hasToken) {
    console.log('ℹ️ 已經授權過了。如需重新授權，請執行：');
    console.log('  npm run setup -- reset-auth');
    return;
  }

  console.log('正在開啟瀏覽器，請在瀏覽器登入你的 Google 帳號並按「允許」...\n');

  await authorize();

  const newStatus = checkStatus();
  printNextStep(newStatus);

  // Force exit — OAuth server may keep sockets open
  process.exit(0);
}

// ── reset-auth ──────────────────────────────────────────────────
async function cmdResetAuth() {
  if (fs.existsSync(TOKEN_PATH)) {
    fs.unlinkSync(TOKEN_PATH);
    console.log('已清除舊的授權 token');
  }

  const status = checkStatus();
  if (!status.credentialsValid) {
    console.error('❌ 請先設定 Google API 憑證');
    console.error('執行：npm run setup -- credentials <path-to-json>');
    process.exit(1);
  }

  console.log('正在重新授權...\n');
  await authorize();

  const newStatus = checkStatus();
  printNextStep(newStatus);
  process.exit(0);
}

// ── guide ───────────────────────────────────────────────────────
function cmdGuide() {
  console.log(`
╔══════════════════════════════════════════════════════════════╗
║         Sheet Helper — 完整設定教學                         ║
╚══════════════════════════════════════════════════════════════╝

本工具需要以下 4 個步驟完成設定：

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
步驟 1：取得 Google API 憑證
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  1. 開啟 Google Cloud Console：
     https://console.cloud.google.com/projectcreate

  2. 建立一個新專案（名字隨便取，例如「車輛管理」）

  3. 啟用 Google Sheets API：
     https://console.cloud.google.com/apis/library/sheets.googleapis.com

  4. 設定 OAuth 同意畫面：
     https://console.cloud.google.com/apis/credentials/consent
     - User Type 選「外部」
     - 應用程式名稱隨便填
     - 填入你的 email
     - 在「測試使用者」加入你自己的 Google 帳號 email

  5. 建立 OAuth 用戶端 ID：
     https://console.cloud.google.com/apis/credentials/oauthclient
     - 類型選「桌面應用程式」
     - 建立後下載 JSON 檔案

  6. ⚠️ 重要：在 OAuth 用戶端設定中，新增授權的重新導向 URI：
     http://localhost:3456/oauth2callback

  7. 執行以下命令，帶入下載的 JSON 檔案路徑：
     npm run setup -- credentials "C:\\path\\to\\downloaded.json"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
步驟 2：Google 帳號授權
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  執行：npm run setup -- auth
  瀏覽器會自動開啟，登入 Google 後授權即完成。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
步驟 3：連結 Google Sheets
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  打開你的車輛庫存 Google Sheets，複製網址，然後執行：
  npm run setup -- spreadsheet "https://docs.google.com/spreadsheets/d/xxxxx/edit"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
步驟 4：連結 car-prompts 資料夾
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  執行：npm run setup -- car-prompts "D:\\Projects\\car-prompts"
  路徑下應包含「汽車資料」子資料夾。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
設定完成後，可使用以下命令：
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  npm run integrate    建立整合庫存表
  npm run status       檢查庫存和 PO 狀態
  npm run folders      為未 PO 車輛建立資料夾
  npm run fix-width    調整表格欄寬
`);
}

// ── status (human-friendly) ─────────────────────────────────────
function cmdStatus() {
  const status = checkStatus();

  console.log(`\n🚗 Sheet Helper 設定狀態\n`);
  console.log(`  ${status.credentialsValid ? '✅' : '❌'} Google API 憑證    ${status.credentialsValid ? '— 已設定' : ''}`);
  console.log(`  ${status.hasToken ? '✅' : '❌'} Google 帳號授權    ${status.hasToken ? '— 已授權' : ''}`);
  console.log(`  ${status.hasSpreadsheetId ? '✅' : '❌'} Google Sheets      ${status.hasSpreadsheetId ? `— ${status.spreadsheetId!.slice(0, 12)}...` : ''}`);
  console.log(`  ${status.carPromptsValid ? '✅' : '❌'} car-prompts 資料夾 ${status.carPromptsValid ? `— ${status.carPromptsPath}（${status.carCount} 台車）` : ''}`);

  const allDone = status.credentialsValid && status.hasToken && status.hasSpreadsheetId && status.carPromptsValid;

  if (allDone) {
    console.log(`\n🎉 設定完成！可以開始使用：`);
    console.log(`  npm run integrate    建立整合庫存表`);
    console.log(`  npm run status       檢查庫存和 PO 狀態`);
    console.log(`  npm run folders      為未 PO 車輛建立資料夾`);
    console.log(`  npm run fix-width    調整表格欄寬\n`);
  } else {
    printNextStep(status);
  }
}

// ── status (JSON for CLI tools like Gemini) ─────────────────────
function cmdStatusJson() {
  const status = checkStatus();
  const allDone = status.credentialsValid && status.hasToken && status.hasSpreadsheetId && status.carPromptsValid;

  const output = {
    ready: allDone,
    credentials: { ok: status.credentialsValid },
    auth: { ok: status.hasToken },
    spreadsheet: { ok: status.hasSpreadsheetId, id: status.spreadsheetId },
    carPrompts: {
      ok: status.carPromptsValid,
      path: status.carPromptsPath,
      carCount: status.carCount,
    },
    nextStep: allDone ? null : getNextStepCommand(status),
  };

  console.log(JSON.stringify(output, null, 2));
}

// ── Helpers ─────────────────────────────────────────────────────
function getNextStepCommand(status: ReturnType<typeof checkStatus>): string {
  if (!status.credentialsValid) return 'npm run setup -- credentials <path-to-json>';
  if (!status.hasToken) return 'npm run setup -- auth';
  if (!status.hasSpreadsheetId) return 'npm run setup -- spreadsheet <url>';
  if (!status.carPromptsValid) return 'npm run setup -- car-prompts <path>';
  return '';
}

function printNextStep(status: ReturnType<typeof checkStatus>) {
  const next = getNextStepCommand(status);
  if (next) {
    console.log(`\n📋 下一步：${next}`);
    console.log(`   （執行 npm run setup -- guide 查看完整教學）\n`);
  }
}

// ── Entry point ─────────────────────────────────────────────────
main().catch(err => {
  if (err?.code === 'ENOTFOUND' || err?.code === 'ETIMEDOUT') {
    console.error('❌ 無法連上 Google，請確認網路連線');
  } else {
    console.error('❌ 錯誤：', err?.message || err);
  }
  process.exit(1);
});
