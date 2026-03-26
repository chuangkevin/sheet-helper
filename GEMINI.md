# Sheet Helper — AI 助手指引

## 專案簡介

這是「車輛庫存管理整合工具」，用來同步 Google Sheets 車輛庫存資料，並與 `car-prompts` 專案整合產生 PO 文。

## 工作區結構

使用者的工作區通常包含三個相關專案：

```
汽車維護/                  ← VS Code workspace 根目錄
├── car-prompts/           ← 車輛 PO 文資料（裡面有「汽車資料」子資料夾）
├── post-helper/           ← PO 文生成工具
└── sheet-helper/          ← 本專案：Google Sheets 庫存管理
```

## 使用者初次使用時的引導流程

當使用者不知道要做什麼，或專案尚未設定完成時，請按以下步驟引導：

### 0. 設定 VS Code Workspace

如果母資料夾（例如 `~/Documents/汽車維護/`）還沒有 `.code-workspace` 檔案，請把 repo 內的 `汽車維護.code-workspace` 複製到母資料夾：

```bash
cp sheet-helper/汽車維護.code-workspace ../汽車維護.code-workspace
```

之後使用者只要雙擊這個檔案，就能在 VS Code 同時開啟 car-prompts、post-helper、sheet-helper 三個專案。

### 1. 確認 gcloud CLI 已安裝

```bash
gcloud --version
```

如果沒有安裝：
```bash
brew install google-cloud-sdk
```

### 2. 全自動設定 Google API（推薦方式）

```bash
cd sheet-helper
npm install
npm run setup -- init-gcloud
```

這個指令會自動完成以下所有事情：
1. 登入 Google 帳號（瀏覽器會打開，使用者只需要按「允許」）
2. 建立 Google Cloud 專案
3. 啟用 Google Sheets API
4. 取得應用程式憑證（瀏覽器會再打開一次，使用者按「允許」）

**使用者全程只需要在瀏覽器按兩次「允許」，其他全自動。**

### 3. 連結 Google Sheets

請使用者提供他的車輛庫存 Google Sheets 網址，然後執行：
```bash
npm run setup -- spreadsheet "<Google Sheets 網址>"
```

如果出現權限錯誤，程式會顯示目前登入的 email，請使用者去表格的「共用」設定中加入該 email。

### 4. 連結 car-prompts 資料夾

car-prompts 專案通常就在隔壁資料夾：
```bash
npm run setup -- car-prompts "../car-prompts"
```

### 5. 確認設定完成

```bash
npm run setup -- status
```
當 `"ready": true` 時，所有設定完成。

### 備用方式：手動設定 Google API

如果 gcloud CLI 無法使用，可以改用手動方式：

1. 開啟 https://console.cloud.google.com/projectcreate 建立專案
2. 開啟 https://console.cloud.google.com/apis/library/sheets.googleapis.com 啟用 Sheets API
3. 開啟 https://console.cloud.google.com/apis/credentials/consent 設定 OAuth 同意畫面
   - User Type 選「外部」，填入 email，在「測試使用者」加入自己的 email
4. 開啟 https://console.cloud.google.com/apis/credentials/oauthclient 建立憑證
   - 類型選「桌面應用程式」
   - **重要：在「授權的重新導向 URI」加入 `http://localhost:3456/oauth2callback`**
5. 下載 JSON 檔案後執行：
```bash
npm run setup -- credentials "<下載的 JSON 檔案路徑>"
npm run setup -- auth
```

## 日常操作指令

設定完成後，使用者可以使用以下指令：

| 指令 | 用途 |
|------|------|
| `npm run integrate` | 讀取車源+庫存表，建立整合庫存表 |
| `npm run status` | 檢查庫存狀態和 PO 進度 |
| `npm run folders` | 為未 PO 車輛在 car-prompts 建立資料夾 |
| `npm run fix-width` | 調整整合庫存表的欄寬 |
| `npm run setup` | 查看/修改設定 |
| `npm run setup -- init-gcloud` | 重新設定 Google API |

## 常見任務

- **「幫我更新庫存」** → 執行 `npm run integrate`
- **「有哪些車還沒 PO？」** → 執行 `npm run status`
- **「幫沒 PO 的車建資料夾」** → 執行 `npm run folders`
- **「重新設定 Google 帳號」** → 執行 `npm run setup -- init-gcloud`

## 重要規則

### 絕對不能做的事

- **絕對不能刪除既有的工作表（sheet）**。「車源」和「庫存」是使用者手動維護的原始資料，刪掉就沒了。
- 如果需要產出新資料，**只能新增工作表**（例如「整合庫存」），不能覆蓋或刪除原有的工作表。
- `npm run integrate` 只會刪除並重建「整合庫存」這個由程式產生的工作表，不會動到「車源」和「庫存」。
- 如果使用者要你讀取表格資料或做分析，**讀取就好，不要修改原表**。

### 其他注意事項

- 敏感檔案（credentials.json、token.json、.env）不可提交到 Git
- Google Sheets API 有請求頻率限制，避免短時間內大量執行
- 整合庫存表是合併產生，修改後不會回寫原表
