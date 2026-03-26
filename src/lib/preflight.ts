import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import * as fs from 'fs';
import * as path from 'path';
import * as http from 'http';
import * as dotenv from 'dotenv';
import { execSync } from 'child_process';

dotenv.config();

// ── Path constants ──────────────────────────────────────────────
export const PROJECT_ROOT = path.join(__dirname, '..', '..');
export const CREDENTIALS_PATH = path.join(PROJECT_ROOT, 'credentials.json');
export const SERVICE_ACCOUNT_PATH = path.join(PROJECT_ROOT, 'service-account.json');
export const TOKEN_PATH = path.join(PROJECT_ROOT, 'token.json');
export const ENV_PATH = path.join(PROJECT_ROOT, '.env');
export const ENV_EXAMPLE_PATH = path.join(PROJECT_ROOT, '.env.example');

const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/userinfo.email',
];

// ADC path (gcloud auth application-default login)
function getAdcPath(): string {
  if (process.platform === 'win32') {
    return path.join(process.env.APPDATA || '', 'gcloud', 'application_default_credentials.json');
  }
  return path.join(process.env.HOME || '', '.config', 'gcloud', 'application_default_credentials.json');
}

export function hasAdc(): boolean {
  return fs.existsSync(getAdcPath());
}

// ── Status types ────────────────────────────────────────────────
export interface PreflightStatus {
  hasNodeModules: boolean;
  hasCredentials: boolean;
  credentialsValid: boolean;
  hasServiceAccount: boolean;
  serviceAccountValid: boolean;
  hasToken: boolean;
  hasAdc: boolean;            // gcloud ADC available
  authReady: boolean;         // (credentials+token) or service-account or ADC
  hasEnv: boolean;
  hasSpreadsheetId: boolean;
  spreadsheetId: string | null;
  hasCarPromptsPath: boolean;
  carPromptsPath: string | null;
  carPromptsValid: boolean;
  carCount: number;
}

// ── Status check (synchronous, no side effects) ─────────────────
export function checkStatus(): PreflightStatus {
  const hasNodeModules = fs.existsSync(path.join(PROJECT_ROOT, 'node_modules'));
  const hasCredentials = fs.existsSync(CREDENTIALS_PATH);

  let credentialsValid = false;
  if (hasCredentials) {
    try {
      const content = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf-8'));
      credentialsValid = !!(content?.installed?.client_id && content?.installed?.client_secret);
    } catch {}
  }

  const hasServiceAccount = fs.existsSync(SERVICE_ACCOUNT_PATH);
  let serviceAccountValid = false;
  if (hasServiceAccount) {
    try {
      const content = JSON.parse(fs.readFileSync(SERVICE_ACCOUNT_PATH, 'utf-8'));
      serviceAccountValid = content?.type === 'service_account' && !!content?.client_email && !!content?.private_key;
    } catch {}
  }

  const hasToken = fs.existsSync(TOKEN_PATH);
  const adcAvailable = hasAdc();
  const authReady = (credentialsValid && hasToken) || serviceAccountValid || adcAvailable;

  const hasEnv = fs.existsSync(ENV_PATH);
  let spreadsheetId: string | null = null;
  let carPromptsPath: string | null = null;

  if (hasEnv) {
    const envConfig = dotenv.parse(fs.readFileSync(ENV_PATH, 'utf-8'));
    spreadsheetId = envConfig.SPREADSHEET_ID || process.env.SPREADSHEET_ID || null;
    carPromptsPath = envConfig.CAR_PROMPTS_PATH || process.env.CAR_PROMPTS_PATH || null;
  }

  const hasSpreadsheetId = !!spreadsheetId && spreadsheetId !== 'your_spreadsheet_id_here';
  const hasCarPromptsPath = !!carPromptsPath && carPromptsPath.length > 0;

  let carPromptsValid = false;
  let carCount = 0;
  if (hasCarPromptsPath && carPromptsPath) {
    const result = findCarDataFolder(carPromptsPath);
    carPromptsValid = result.found;
    carCount = result.count;
  }

  return {
    hasNodeModules,
    hasCredentials,
    credentialsValid,
    hasServiceAccount,
    serviceAccountValid,
    hasToken,
    hasAdc: adcAvailable,
    authReady,
    hasEnv,
    hasSpreadsheetId,
    spreadsheetId: hasSpreadsheetId ? spreadsheetId : null,
    hasCarPromptsPath,
    carPromptsPath,
    carPromptsValid,
    carCount,
  };
}

// ── Get authorized email ─────────────────────────────────────────
export async function getAuthorizedEmail(auth: any): Promise<string | null> {
  // If it's a ServiceAccount object or has credentials with client_email
  if (auth?.credentials?.client_email) {
    return auth.credentials.client_email;
  }
  if (auth?.client_email) {
    return auth.client_email;
  }

  try {
    const oauth2 = google.oauth2({ version: 'v2', auth });
    const res = await oauth2.userinfo.get();
    return res.data.email || null;
  } catch {
    return null;
  }
}

// ── Authorize (try Service Account, then ADC, then credentials.json) ────────────
export async function authorize(): Promise<any> {
  // Strategy 1: Service Account (Recommended for servers/CLI without local browser)
  if (fs.existsSync(SERVICE_ACCOUNT_PATH)) {
    try {
      const auth = new google.auth.GoogleAuth({
        keyFile: SERVICE_ACCOUNT_PATH,
        scopes: SCOPES,
      });
      const client = await auth.getClient();
      return client;
    } catch (err) {
      // Service account failed, try next
    }
  }

  // Strategy 2: ADC (gcloud auth application-default login)
  if (hasAdc()) {
    try {
      const auth = new GoogleAuth({ scopes: SCOPES });
      const client = await auth.getClient();
      return client;
    } catch {
      // ADC failed, try credentials.json
    }
  }

  // Strategy 3: Traditional credentials.json + token.json
  if (!fs.existsSync(CREDENTIALS_PATH)) {
    throw new Error('找不到 Google 憑證。請執行 npm run setup -- init-gcloud 自動設定，或 npm run setup -- credentials <path> 手動設定');
  }

  const content = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf-8'));
  const { client_secret, client_id } = content.installed;
  const redirectUri = 'http://localhost:3456/oauth2callback';
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirectUri);

  if (fs.existsSync(TOKEN_PATH)) {
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf-8'));
    oAuth2Client.setCredentials(token);
    return oAuth2Client;
  }

  // No token — need OAuth via local server
  return await getNewTokenViaLocalServer(oAuth2Client);
}

async function getNewTokenViaLocalServer(oAuth2Client: any): Promise<any> {
  return new Promise((resolve, reject) => {
    const server = http.createServer(async (req, res) => {
      try {
        const parsedUrl = new URL(req.url || '', 'http://localhost:3456');
        if (parsedUrl.pathname !== '/oauth2callback') {
          res.writeHead(404);
          res.end();
          return;
        }

        const code = parsedUrl.searchParams.get('code');
        if (!code) {
          res.writeHead(400);
          res.end('缺少授權碼');
          return;
        }

        const { tokens } = await oAuth2Client.getToken(code);
        oAuth2Client.setCredentials(tokens);
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));

        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end('<html><body><h1>✅ 授權成功！</h1><p>你可以關閉這個頁面了。</p></body></html>');

        server.close(() => {
          // Ensure all connections are destroyed so process can exit
          server.unref();
        });
        console.log('✅ Google 帳號已連結成功！');
        resolve(oAuth2Client);
      } catch (err) {
        res.writeHead(500);
        res.end('授權失敗');
        server.close();
        reject(err);
      }
    });

    server.listen(3456, () => {
      const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
      });

      console.log('正在開啟瀏覽器進行 Google 授權...');
      console.log(`\n授權網址：${authUrl}\n`);
      console.log('等待瀏覽器授權回調...');

      openBrowser(authUrl);
    });

    // Timeout after 5 minutes
    setTimeout(() => {
      server.close();
      reject(new Error('授權逾時（5 分鐘），請重新執行'));
    }, 5 * 60 * 1000);
  });
}

// ── Preflight gate (call at top of each command) ────────────────
export async function preflight(options?: { needCarPrompts?: boolean }): Promise<any> {
  const status = checkStatus();
  const needCar = options?.needCarPrompts ?? true;

  if (!status.authReady) {
    console.error('❌ 還沒設定好，請先執行：npm run setup -- init-gcloud');
    console.error('   （缺少 Google 憑證，init-gcloud 可以全自動搞定）');
    process.exit(1);
  }

  if (!status.hasSpreadsheetId) {
    console.error('❌ 還沒設定好，請先執行：npm run setup');
    console.error('   （還沒連結你的 Google 表格）');
    process.exit(1);
  }

  if (needCar && !status.hasCarPromptsPath) {
    console.error('❌ 還沒設定好，請先執行：npm run setup');
    console.error('   （還沒指定汽車資料的位置）');
    process.exit(1);
  }

  const auth = await authorize();
  return auth;
}

// ── Utility: parse Spreadsheet URL ──────────────────────────────
export function parseSpreadsheetUrl(input: string): string | null {
  const match = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(input.trim())) return input.trim();
  return null;
}

// ── Utility: validate credentials JSON ──────────────────────────
export function validateCredentialsFile(filePath: string): boolean {
  try {
    const content = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    return !!(content?.installed?.client_id && content?.installed?.client_secret);
  } catch {
    return false;
  }
}

// ── Utility: validate service account JSON ──────────────────────
export function validateServiceAccountFile(filePath: string): boolean {
  try {
    const content = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    return content?.type === 'service_account' && !!content?.client_email && !!content?.private_key;
  } catch {
    return false;
  }
}

// ── Utility: find 汽車資料 folder ───────────────────────────────
export function findCarDataFolder(basePath: string): { found: boolean; path: string; count: number } {
  const carDataPath = path.join(basePath, '汽車資料');

  if (fs.existsSync(carDataPath) && fs.statSync(carDataPath).isDirectory()) {
    const count = countCarFolders(carDataPath);
    return { found: true, path: carDataPath, count };
  }

  if (fs.existsSync(basePath) && fs.statSync(basePath).isDirectory()) {
    const entries = fs.readdirSync(basePath);
    const carLike = entries.filter(e => {
      try {
        return fs.statSync(path.join(basePath, e)).isDirectory() && /^\d{4}\s/.test(e);
      } catch { return false; }
    });
    if (carLike.length > 3) {
      return { found: true, path: basePath, count: carLike.length };
    }
  }

  return { found: false, path: '', count: 0 };
}

function countCarFolders(dirPath: string): number {
  try {
    return fs.readdirSync(dirPath).filter(f => {
      const full = path.join(dirPath, f);
      return fs.statSync(full).isDirectory();
    }).length;
  } catch {
    return 0;
  }
}

// ── Utility: sanitize path ──────────────────────────────────────
export function sanitizePath(input: string): string {
  let p = input.trim();
  if ((p.startsWith('"') && p.endsWith('"')) || (p.startsWith("'") && p.endsWith("'"))) {
    p = p.slice(1, -1);
  }
  return p.trim();
}

// ── Utility: read/write .env ────────────────────────────────────
export function writeEnvValue(key: string, value: string): void {
  let lines: string[] = [];

  if (fs.existsSync(ENV_PATH)) {
    lines = fs.readFileSync(ENV_PATH, 'utf-8').split('\n');
  } else if (fs.existsSync(ENV_EXAMPLE_PATH)) {
    lines = fs.readFileSync(ENV_EXAMPLE_PATH, 'utf-8').split('\n');
  }

  let found = false;
  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    if (trimmed.startsWith(`${key}=`) || trimmed === key) {
      lines[i] = `${key}=${value}`;
      found = true;
      break;
    }
  }

  if (!found) {
    lines.push(`${key}=${value}`);
  }

  fs.writeFileSync(ENV_PATH, lines.join('\n'));
  dotenv.config({ override: true });
}

// ── Utility: open browser ───────────────────────────────────────
export async function openBrowser(url: string): Promise<void> {
  try {
    const open = (await import('open')).default;
    await open(url);
  } catch {
    try {
      // Fallback: platform-specific command
      const { execSync } = require('child_process');
      const cmd = process.platform === 'darwin'
        ? `open "${url}"`
        : process.platform === 'win32'
          ? `start "" "${url}"`
          : `xdg-open "${url}"`;
      execSync(cmd, { stdio: 'ignore' });
    } catch {
      // All methods failed, show URL for manual opening
      console.log(`\n⚠️ 無法自動開啟瀏覽器，請手動開啟以下網址：`);
      console.log(`\n  ${url}\n`);
    }
  }
}
