# Spec: 首次使用引導精靈（Onboarding Setup Wizard）

## 概述

使用者 clone 這個專案後，透過 CLI 子命令逐步完成設定。所有操作都是非互動式的（不需要 stdin 輸入），可由 AI CLI 工具（如 Gemini CLI）自動驅動。使用者**不需要知道任何檔案要放哪裡、環境變數叫什麼名字**——工具自動處理一切。

## 動機

- 使用者打開專案，不知道要幹嘛
- 設定步驟散落在不同地方，容易漏掉
- 缺少任何一個設定，程式就會報錯
- 需要支援 AI CLI 工具（如 Gemini）自動執行，不能依賴互動式 stdin
- 手動建立 Google Cloud OAuth 憑證對一般使用者太困難（需要建專案、同意畫面、憑證、測試使用者）

## 認證方式

### 推薦方式：gcloud ADC（Application Default Credentials）

使用 `gcloud` CLI 全自動設定，**完全不需要手動建立 OAuth 同意畫面、憑證、或加測試使用者**。

使用者全程只需在瀏覽器按兩次「允許」：

1. `gcloud auth login` — 登入 Google 帳號
2. `gcloud auth application-default login` — 授權應用程式存取

### 備用方式：手動 credentials.json

如果 gcloud CLI 不可用，可改用傳統方式：手動在 Google Cloud Console 建立 OAuth 桌面應用程式憑證，下載 JSON 檔案後設定。

### 認證優先順序

程式 `authorize()` 會依序嘗試：

1. **ADC**：檢查 `~/.config/gcloud/application_default_credentials.json`（macOS/Linux）或 `%APPDATA%\gcloud\application_default_credentials.json`（Windows）
2. **credentials.json + token.json**：傳統 OAuth Desktop App 方式

## CLI 介面設計

所有設定操作都是獨立的子命令，不需要互動式輸入：

```bash
npm run setup                              # 顯示目前設定狀態與下一步指引
npm run setup -- status                    # 輸出 JSON 格式狀態（供 AI CLI 解析）
npm run setup -- init-gcloud               # 全自動 Google API 設定（推薦）
npm run setup -- credentials <path>        # 手動設定 Google OAuth 憑證
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

1. 確認 `gcloud` CLI 已安裝（沒有的話 `brew install google-cloud-sdk`）
2. 執行 `npm run setup -- init-gcloud`（自動完成 Google API 設定）
3. 執行 `npm run setup -- spreadsheet <url>`（使用者提供表格網址）
4. 執行 `npm run setup -- car-prompts ../car-prompts`（自動偵測隔壁資料夾）
5. 執行 `npm run setup -- status` 確認 `"ready": true`

### init-gcloud 子命令

```bash
npm run setup -- init-gcloud
```

自動執行以下步驟：

1. **檢查 gcloud CLI** — 找不到則提示安裝方式
2. **`gcloud auth login --brief`** — 開啟瀏覽器，使用者按「允許」
3. **`gcloud projects create`** — 自動建立 Google Cloud 專案（ID: `sheet-helper-{timestamp}`）
4. **`gcloud services enable sheets.googleapis.com`** — 啟用 Sheets API
5. **`gcloud auth application-default login --scopes=...`** — 開啟瀏覽器取得 ADC，使用者按「允許」
6. **`gcloud auth application-default set-quota-project`** — 設定配額專案
7. **驗證 ADC 檔案存在** — 確認設定成功
8. **顯示登入帳號** — 提醒使用者確認表格有分享給該帳號

### 連結 Google Sheets

```bash
npm run setup -- spreadsheet "https://docs.google.com/spreadsheets/d/1lT0X2rx.../edit"
```

- 接受完整 URL 或單純的 Spreadsheet ID
- 自動從 URL 解析出 ID
- 如果已有授權，會驗證表格是否可存取，並顯示表格名稱和工作表列表
- **權限錯誤時**：自動查詢已授權帳號的 email，告訴使用者「把 xxx@gmail.com 加進表格的共用設定」
- 自動寫入 `.env`

### 連結 car-prompts 資料夾

```bash
npm run setup -- car-prompts "../car-prompts"
```

- 自動在路徑下尋找「汽車資料」子資料夾
- 如果路徑本身就是汽車資料（裡面有 `年份 品牌...` 格式的子資料夾），也能辨識
- 顯示找到幾台車的資料
- 自動寫入 `.env`

---

## 自動偵測機制（Preflight Check）

每個指令（integrate、status、folders、fix-width）執行時，會先做 preflight check：

```
if (!authReady) → 印出「請先執行 npm run setup -- init-gcloud」並結束
if (!spreadsheetId) → 印出「請先執行 npm run setup」並結束
if (!carPromptsPath) → 印出「請先執行 npm run setup」並結束（僅需要 car-prompts 的指令）
```

`authReady` 為 `true` 的條件：ADC 存在 **或** (credentials.json 有效 + token.json 存在)

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

  ✅ Google API 認證    — gcloud（自動）
  ✅ Google Sheets      — 1lT0X2rxALWO...
  ✅ car-prompts 資料夾 — ../car-prompts（75 台車）

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
  "auth": { "ok": true, "method": "gcloud-adc" },
  "spreadsheet": { "ok": true, "id": "1lT0X2rx..." },
  "carPrompts": { "ok": true, "path": "../car-prompts", "carCount": 75 },
  "nextStep": null
}
```

`auth.method` 可能的值：`"gcloud-adc"` | `"credentials.json"` | `"none"`

---

## 錯誤處理

| 情境 | 處理方式 |
|------|---------|
| gcloud CLI 未安裝 | macOS: 「請先安裝：brew install google-cloud-sdk」 |
| gcloud 登入失敗 | 「登入 Google 帳號失敗」+ exit 1 |
| JSON 檔案格式錯誤 | 「檔案格式不正確，請確認下載的是『桌面應用程式』類型的 OAuth 憑證」 |
| 網路連線失敗 | 「無法連上 Google，請確認網路連線」 |
| Sheets 無權限（403/401） | 顯示已登入 email + 分享步驟（「點右上角共用，把 xxx@gmail.com 加進去」）|
| Sheets 找不到（404） | 「找不到這個表格，請確認網址是正確的」 |
| car-prompts 路徑不存在 | 「找不到路徑：xxx」 |
| car-prompts 內無汽車資料 | 「在這個路徑下找不到『汽車資料』資料夾」 |
| OAuth 逾時 | 「授權逾時（5 分鐘），請重新執行」 |
| 未帶參數 | 顯示該子命令的用法和範例 |

所有錯誤都以 exit code 1 結束，方便 CI/AI 工具判斷。
所有面向使用者的訊息使用繁體中文白話文，不使用技術術語。

---

## 技術實作

### 檔案結構

- `src/setup.ts` — CLI 子命令路由與各步驟實作
- `src/lib/preflight.ts` — Preflight check + 共用 utilities（auth、env 讀寫、路徑驗證）

### 認證模組

`preflight.ts` 的 `authorize()` 函數：

1. 檢查 ADC 檔案是否存在（`getAdcPath()`）
2. 如果有，使用 `GoogleAuth` 建立 client
3. 如果沒有，fallback 到 `credentials.json` + `token.json` 流程
4. `token.json` 不存在時，啟動本地 HTTP server（`localhost:3456`）接收 OAuth 回調

### OAuth 回調機制（備用方式）

使用 Node.js `http` 模組在 `localhost:3456` 建立臨時 HTTP server：

- 接收 `/oauth2callback?code=xxx` 回調
- 自動交換 token
- 回傳成功頁面給瀏覽器
- 關閉 server + `process.exit(0)` 確保程式結束

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

- `googleapis` — Google Sheets API client
- `google-auth-library`（googleapis 的依賴）— ADC 支援
- `http`（內建）— OAuth 回調 server
- `child_process`（內建）— 執行 gcloud 指令
- `open`（已安裝）— 開啟瀏覽器

---

## 不做什麼

- 不做互動式 stdin 輸入（readline）— 所有操作都透過 CLI 參數
- 不做 GUI 介面
- 不在設定過程中教使用者怎麼用 Google Sheets（只設定連線）
- 不處理 Service Account 認證方式
- 不自動修改或刪除 Google Sheets 的「車源」和「庫存」工作表
