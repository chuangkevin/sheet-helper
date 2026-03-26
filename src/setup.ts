import * as fs from 'fs';
import * as path from 'path';
import { execSync } from 'child_process';
import {
  checkStatus,
  authorize,
  getAuthorizedEmail,
  hasAdc,
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
//   npm run setup -- init-gcloud               全自動設定 Google API（推薦）
//   npm run setup -- credentials <path>        手動設定 Google OAuth 憑證
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
    case 'init-gcloud':
      await cmdInitGcloud();
      break;
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

// ── init-gcloud (全自動 Google API 設定) ─────────────────────────
function runCmd(cmd: string, label: string): boolean {
  try {
    execSync(cmd, { stdio: 'inherit' });
    return true;
  } catch {
    console.error(`❌ ${label} 失敗`);
    return false;
  }
}

function runCmdSilent(cmd: string): string {
  try {
    return execSync(cmd, { encoding: 'utf-8' }).trim();
  } catch {
    return '';
  }
}

async function cmdInitGcloud() {
  console.log('');
  console.log('🚀 全自動 Google API 設定');
  console.log('   只需要在瀏覽器點兩次「允許」就搞定！');
  console.log('');

  // Step 0: Check gcloud is installed
  const gcloudVersion = runCmdSilent('gcloud --version');
  if (!gcloudVersion) {
    console.error('❌ 找不到 gcloud CLI');
    console.error('');
    if (process.platform === 'darwin') {
      console.error('請先安裝：brew install google-cloud-sdk');
    } else {
      console.error('請先安裝：https://cloud.google.com/sdk/docs/install');
    }
    process.exit(1);
  }
  console.log('✅ 找到 gcloud CLI');

  // Step 1: gcloud auth login
  console.log('');
  console.log('━━━ 第 1 步：登入 Google 帳號 ━━━');
  console.log('（瀏覽器會自動打開，請登入並按「允許」）');
  console.log('');
  if (!runCmd('gcloud auth login --brief', '登入 Google 帳號')) {
    process.exit(1);
  }
  console.log('✅ Google 帳號已登入');

  // Get account email
  const email = runCmdSilent('gcloud config get-value account 2>/dev/null');
  if (email) {
    console.log(`   帳號：${email}`);
  }

  // Step 2: Create project (or use existing)
  console.log('');
  console.log('━━━ 第 2 步：建立 Google Cloud 專案 ━━━');
  const projectId = `sheet-helper-${Date.now().toString(36)}`;
  const createResult = runCmdSilent(`gcloud projects create ${projectId} --name="車輛管理" 2>&1`);

  if (createResult.includes('already exists') || createResult.includes('ALREADY_EXISTS')) {
    console.log(`ℹ️ 使用既有專案：${projectId}`);
  } else if (createResult.includes('ERROR')) {
    // Try to use existing project
    const existingProject = runCmdSilent('gcloud config get-value project 2>/dev/null');
    if (existingProject) {
      console.log(`ℹ️ 無法建立新專案，使用目前專案：${existingProject}`);
    } else {
      console.error('❌ 無法建立 Google Cloud 專案');
      console.error(createResult);
      process.exit(1);
    }
  } else {
    console.log(`✅ 已建立專案：${projectId}`);
  }

  // Set project
  const activeProject = runCmdSilent(`gcloud config set project ${projectId} 2>&1`) ? projectId
    : runCmdSilent('gcloud config get-value project 2>/dev/null');
  console.log(`   使用專案：${activeProject}`);

  // Step 3: Enable Sheets API
  console.log('');
  console.log('━━━ 第 3 步：啟用 Google Sheets API ━━━');
  runCmd(`gcloud services enable sheets.googleapis.com --project=${activeProject}`, '啟用 Sheets API');
  console.log('✅ Google Sheets API 已啟用');

  // Step 4: ADC login with scopes
  console.log('');
  console.log('━━━ 第 4 步：取得應用程式憑證 ━━━');
  console.log('（瀏覽器會再打開一次，請按「允許」）');
  console.log('');

  const scopeStr = 'https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/userinfo.email';
  if (!runCmd(
    `gcloud auth application-default login --scopes=${scopeStr}`,
    '取得應用程式憑證'
  )) {
    process.exit(1);
  }

  // Set quota project
  runCmdSilent(`gcloud auth application-default set-quota-project ${activeProject} 2>&1`);

  console.log('✅ 應用程式憑證已取得');

  // Verify
  console.log('');
  console.log('━━━ 驗證設定 ━━━');

  if (hasAdc()) {
    console.log('✅ Google API 設定完成！');
  } else {
    console.error('❌ 找不到應用程式憑證，請重新執行');
    process.exit(1);
  }

  if (email) {
    console.log(`   登入帳號：${email}`);
    console.log(`   記得把你的 Google 表格分享給 ${email}（如果你是表格擁有者就不用）`);
  }

  // Save gcloud project to env
  writeEnvValue('GCLOUD_PROJECT', activeProject);

  const status = checkStatus();
  printNextStep(status);

  console.log('');
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
  if (status.authReady) {
    const auth = await authorize();
    try {
      const sheets = google.sheets({ version: 'v4', auth });
      const res = await sheets.spreadsheets.get({ spreadsheetId: id });

      const title = res.data.properties?.title || '(未命名)';
      const sheetNames = res.data.sheets?.map(s => s.properties?.title).filter(Boolean) || [];

      writeEnvValue('SPREADSHEET_ID', id);
      console.log(`✅ 已連結 Google Sheets：「${title}」`);
      console.log(`   工作表：${sheetNames.join('、')}`);
    } catch (err: any) {
      if (err?.code === 403 || err?.code === 401) {
        const email = await getAuthorizedEmail(auth);
        console.error('❌ 沒辦法打開這個表格');
        console.error('');
        if (email) {
          console.error(`你目前登入的帳號是：${email}`);
          console.error('');
          console.error('請這樣做：');
          console.error('  1. 在瀏覽器打開你的 Google 表格');
          console.error('  2. 點右上角綠色的「共用」按鈕');
          console.error(`  3. 把 ${email} 加進去，權限選「編輯者」`);
          console.error('  4. 按「完成」');
        } else {
          console.error('請打開你的 Google 表格，點右上角「共用」，把你剛剛登入的 Gmail 加進去。');
        }
        console.error('');
        console.error('弄好之後，再執行一次同樣的指令就可以了。');
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

  const auth = await authorize();
  const email = await getAuthorizedEmail(auth);
  if (email) {
    console.log(`已登入帳號：${email}`);
  }

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

  const authLabel = status.hasAdc ? '— gcloud（自動）' : status.credentialsValid ? '— credentials.json' : '';
  const authOk = status.authReady;

  console.log(`\n🚗 Sheet Helper 設定狀態\n`);
  console.log(`  ${authOk ? '✅' : '❌'} Google API 認證    ${authLabel}`);
  console.log(`  ${status.hasSpreadsheetId ? '✅' : '❌'} Google Sheets      ${status.hasSpreadsheetId ? `— ${status.spreadsheetId!.slice(0, 12)}...` : ''}`);
  console.log(`  ${status.carPromptsValid ? '✅' : '❌'} car-prompts 資料夾 ${status.carPromptsValid ? `— ${status.carPromptsPath}（${status.carCount} 台車）` : ''}`);

  const allDone = authOk && status.hasSpreadsheetId && status.carPromptsValid;

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
  const allDone = status.authReady && status.hasSpreadsheetId && status.carPromptsValid;

  const output = {
    ready: allDone,
    auth: { ok: status.authReady, method: status.hasAdc ? 'gcloud-adc' : status.credentialsValid ? 'credentials.json' : 'none' },
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
  if (!status.authReady) return 'npm run setup -- init-gcloud';
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
