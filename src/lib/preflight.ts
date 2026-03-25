import { google } from 'googleapis';
import * as fs from 'fs';
import * as path from 'path';
import * as http from 'http';
import * as url from 'url';
import * as dotenv from 'dotenv';

dotenv.config();

// ── Path constants ──────────────────────────────────────────────
export const PROJECT_ROOT = path.join(__dirname, '..', '..');
export const CREDENTIALS_PATH = path.join(PROJECT_ROOT, 'credentials.json');
export const TOKEN_PATH = path.join(PROJECT_ROOT, 'token.json');
export const ENV_PATH = path.join(PROJECT_ROOT, '.env');
export const ENV_EXAMPLE_PATH = path.join(PROJECT_ROOT, '.env.example');

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

// ── Status types ────────────────────────────────────────────────
export interface PreflightStatus {
  hasNodeModules: boolean;
  hasCredentials: boolean;
  credentialsValid: boolean;
  hasToken: boolean;
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

  const hasToken = fs.existsSync(TOKEN_PATH);

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
    hasToken,
    hasEnv,
    hasSpreadsheetId,
    spreadsheetId: hasSpreadsheetId ? spreadsheetId : null,
    hasCarPromptsPath,
    carPromptsPath,
    carPromptsValid,
    carCount,
  };
}

// ── Authorize (consolidated OAuth flow) ─────────────────────────
export async function authorize(): Promise<any> {
  if (!fs.existsSync(CREDENTIALS_PATH)) {
    throw new Error('credentials.json 不存在，請先執行 npm run setup credentials <path>');
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
        const parsedUrl = url.parse(req.url || '', true);
        if (parsedUrl.pathname !== '/oauth2callback') {
          res.writeHead(404);
          res.end();
          return;
        }

        const code = parsedUrl.query.code as string;
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

        server.close();
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

  if (!status.credentialsValid) {
    console.error('❌ 還沒設定 Google API 憑證');
    console.error('請先執行：npm run setup credentials <path-to-json>');
    process.exit(1);
  }

  if (!status.hasSpreadsheetId) {
    console.error('❌ 還沒連結 Google Sheets');
    console.error('請先執行：npm run setup spreadsheet <url>');
    process.exit(1);
  }

  if (needCar && !status.hasCarPromptsPath) {
    console.error('❌ 還沒設定 car-prompts 資料夾路徑');
    console.error('請先執行：npm run setup car-prompts <path>');
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
    const { exec } = require('child_process');
    exec(`start "" "${url}"`);
  }
}
