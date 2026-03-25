# Spec: 首次使用引導精靈（Onboarding Setup Wizard）

## 概述

使用者 clone 這個專案後，透過 CLI 子命令逐步完成設定。所有操作都是非互動式的（不需要 stdin 輸入），可由 AI CLI 工具（如 Gemini CLI）自動驅動。使用者**不需要知道任何檔案要放哪裡、環境變數叫什麼名字**——工具自動處理一切。

## 動機

- 使用者打開專案，不知道要幹嘛
- 設定步驟散落在不同地方，容易漏掉
- 缺少任何一個設定，程式就會報錯
- 需要支援 AI CLI 工具（如 Gemini）自動執行，不能依賴互動式 stdin

## CLI 介面設計

所有設定操作都是獨立的子命令，不需要互動式輸入：

```bash
npm run setup                              # 顯示目前設定狀態與下一步指引
npm run setup -- status                    # 輸出 JSON 格式狀態（供 AI CLI 解析）
npm run setup -- credentials <path>        # 設定 Google OAuth 憑證
npm run setup -- auth                      # 執行 OAuth 授權（本地 HTTP server 接收回調）
npm run setup -- spreadsheet <url-or-id>   # 連結 Google Sheets
npm run setup -- car-prompts <path>        # 設定 car-prompts 路徑
npm run setup -- guide                     # 顯示完整設定教學
npm run setup -- reset-auth                # 清除 token 並重新授權
```

---

## 詳細流程

### AI CLI 工具（如 Gemini）的自動化流程

AI CLI 工具拿到這個專案後，應該：

1. 執行 `npm run setup -- status` 取得 JSON 狀態
2. 根據 `nextStep` 欄位，依序執行所需的子命令
3. 每個子命令都會明確回報成功或失敗，以及下一步

```json
{
  "ready": false,
  "credentials": { "ok": false },
  "auth": { "ok": false },
  "spreadsheet": { "ok": false, "id": null },
  "carPrompts": { "ok": false, "path": null, "carCount": 0 },
  "nextStep": "npm run setup -- credentials <path-to-json>"
}
```

### 步驟 1：設定 Google API 憑證

```bash
npm run setup -- credentials "C:\Users\xxx\Downloads\client_secret_xxxxx.json"
```

- 驗證檔案存在
- 驗證 JSON 格式（必須包含 `installed.client_id` 和 `installed.client_secret`）
- 自動複製到專案正確位置（使用者不需知道放哪）
- 格式不對時顯示友善提示（不是 API Key、不是 Service Account）

### 步驟 2：Google OAuth 授權

```bash
npm run setup -- auth
```

- 在 `localhost:3456` 啟動臨時 HTTP server
- 自動開啟瀏覽器到 Google 授權頁面
- 使用者在瀏覽器登入並允許存取
- 授權碼透過 HTTP 回調自動接收（無需手動複製貼上）
- Token 自動儲存
- 5 分鐘逾時自動結束

**前提條件**：
- 已完成步驟 1（credentials）
- 使用者需在 Google Cloud Console 的 OAuth 用戶端設定中，新增重新導向 URI：`http://localhost:3456/oauth2callback`

### 步驟 3：連結 Google Sheets

```bash
npm run setup -- spreadsheet "https://docs.google.com/spreadsheets/d/1lT0X2rx.../edit"
```

- 接受完整 URL 或單純的 Spreadsheet ID
- 自動從 URL 解析出 ID
- 如果已有授權，會驗證表格是否可存取，並顯示表格名稱和工作表列表
- 自動寫入設定

### 步驟 4：連結 car-prompts 資料夾

```bash
npm run setup -- car-prompts "D:\Projects\car-prompts"
```

- 自動在路徑下尋找「汽車資料」子資料夾
- 如果路徑本身就是汽車資料（裡面有 `年份 品牌...` 格式的子資料夾），也能辨識
- 顯示找到幾台車的資料
- 自動寫入設定

---

## 自動偵測機制（Preflight Check）

每個指令（integrate、status、folders、fix-width）執行時，會先做 preflight check：

```
if (!credentials) → 印出「請先執行 npm run setup -- credentials <path>」並結束
if (!spreadsheetId) → 印出「請先執行 npm run setup -- spreadsheet <url>」並結束
if (!carPromptsPath) → 印出「請先執行 npm run setup -- car-prompts <path>」並結束（僅需要 car-prompts 的指令）
if (!token) → 自動啟動 OAuth 流程
```

### preflight 選項

| 指令 | needCarPrompts |
|------|---------------|
| integrate | false |
| fix-width | false |
| analyze | false |
| status | true |
| folders | true |

---

## 設定狀態顯示

### 人類友善格式（`npm run setup`）

```
🚗 Sheet Helper 設定狀態

  ✅ Google API 憑證    — 已設定
  ✅ Google 帳號授權    — 已授權
  ✅ Google Sheets      — 1lT0X2rxALWO...
  ✅ car-prompts 資料夾 — D:\Projects\car-prompts（75 台車）

🎉 設定完成！可以開始使用：
  npm run integrate    建立整合庫存表
  npm run status       檢查庫存和 PO 狀態
  npm run folders      為未 PO 車輛建立資料夾
  npm run fix-width    調整表格欄寬
```

### JSON 格式（`npm run setup -- status`）

供 AI CLI 工具解析：
```json
{
  "ready": true,
  "credentials": { "ok": true },
  "auth": { "ok": true },
  "spreadsheet": { "ok": true, "id": "1lT0X2rx..." },
  "carPrompts": { "ok": true, "path": "D:\\Projects\\car-prompts", "carCount": 75 },
  "nextStep": null
}
```

---

## 錯誤處理

| 情境 | 處理方式 |
|------|---------|
| JSON 檔案格式錯誤 | 「檔案格式不正確，請確認下載的是『桌面應用程式』類型的 OAuth 憑證」 |
| 網路連線失敗 | 「無法連上 Google，請確認網路連線」 |
| Sheets 無權限 | 「沒有權限存取這個表格，請確認表格已分享給你授權的 Google 帳號」 |
| car-prompts 路徑不存在 | 「找不到路徑：xxx」 |
| car-prompts 內無汽車資料 | 「在這個路徑下找不到『汽車資料』資料夾」 |
| OAuth 逾時 | 「授權逾時（5 分鐘），請重新執行」 |
| 未帶參數 | 顯示該子命令的用法和範例 |

所有錯誤都以 exit code 1 結束，方便 CI/AI 工具判斷。

---

## 技術實作

### 檔案結構

- `src/setup.ts` — CLI 子命令路由與各步驟實作
- `src/lib/preflight.ts` — Preflight check + 共用 utilities（auth、env 讀寫、路徑驗證）

### OAuth 回調機制

使用 Node.js `http` 模組在 `localhost:3456` 建立臨時 HTTP server：
- 接收 `/oauth2callback?code=xxx` 回調
- 自動交換 token
- 回傳成功頁面給瀏覽器
- 關閉 server

### package.json scripts

```json
{
  "setup": "ts-node src/setup.ts",
  "integrate": "ts-node src/create-integrated-sheet.ts",
  "status": "ts-node src/check-status.ts",
  "folders": "ts-node src/create-car-folders.ts",
  "fix-width": "ts-node src/fix-width.ts"
}
```

### 依賴

全部使用 Node.js 內建模組 + 已安裝的套件，不需新增依賴：
- `http`（內建）— OAuth 回調 server
- `url`（內建）— URL parsing
- `fs` / `path`（內建）— 檔案操作
- `open`（已安裝）— 開啟瀏覽器

---

## 不做什麼

- 不做互動式 stdin 輸入（readline）— 所有操作都透過 CLI 參數
- 不做 GUI 介面
- 不自動建立 Google Cloud 專案（需要使用者手動登入 Google）
- 不處理 Service Account 認證方式（只支援 OAuth Desktop App）
- 不在設定過程中教使用者怎麼用 Google Sheets（只設定連線）
