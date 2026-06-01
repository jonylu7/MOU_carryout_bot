# MOU 攜出申請自動化

自動化台灣財政部財資中心資料攜出審核流程。每週掃描 Google Drive 申請表、填寫 Word 攜出單（含範例截圖）、等待人工審核（HITL），核准後轉換 PDF 並寄送給財資中心承辦人。

## 流程

```
Google Drive 申請表（每週掃描）
  ↓ 讀取「送出申請 = FALSE」的新列
  ↓ 依（攜出日期 × 子題）分組
  ↓ 填入 .docx 攜出單（含範例 workbook 截圖）
  ↓ 發 HITL 審核信給協調人（含 Word 路徑 + 財資中心信件預覽）
  ↓ 協調人修改 Word → 點擊「確認送出」
  ↓ 轉 PDF → 寄給財資中心承辦人
  ↓ 在 Google Drive 標記已送出
```

## 安裝

```bash
pip install -r requirements.txt
```

## 設定

### 1. 環境變數

複製範本並填入：

```bash
cp .env.example .env           # 個人 / 聯絡資訊
cp .env.gemini.example .env.gemini   # Gemini API 金鑰（選用）
```

**`.env`** — 個人資訊：

| 變數 | 說明 |
|---|---|
| `DEFAULT_APPLICANT` | 申請人姓名 |
| `BASE_DIR` | 資料根目錄絕對路徑（例：Dropbox 同步資料夾） |
| `DRIVE_FILE_ID` | Google Drive 申請表的 File ID |
| `GOOGLE_SERVICE_ACCOUNT_KEY` | 服務帳戶金鑰 JSON 的檔名（放在本專案資料夾） |
| `EMAIL_SENDER` | Gmail 寄件地址 |
| `EMAIL_PASSWORD` | Gmail App Password（非登入密碼） |
| `EMAIL_COORDINATOR` | HITL 審核通知收件人 |
| `EMAIL_MOF_TO` | 財資中心信件收件人（測試時填自己，正式改為承辦人 email） |
| `EMAIL_MOF_CC` | 財資中心信件 CC，逗號分隔（可留空） |

**`.env.gemini`** — Gemini API 金鑰（選用，無則退化為純文字整理）：

| 變數 | 說明 |
|---|---|
| `GEMINI_API_KEY` | 至 [Google AI Studio](https://aistudio.google.com/app/apikey) 取得 |

### 2. Google 服務帳戶

1. 至 [Google Cloud Console](https://console.cloud.google.com/) 建立專案並啟用 **Google Drive API** 與 **Google Sheets API**
2. 建立服務帳戶 → 下載 JSON 金鑰 → 放至專案資料夾
3. 在 Google Drive 對服務帳戶 email 共用申請表（**編輯者**權限）

### 3. 電子簽章（選用）

將簽名圖片（PNG，透明背景，約 300×100px）存為 `signature.png` 放在專案資料夾。

## 執行

```bash
# 測試模式（只讀，不寫入、不寄信）
python main.py --dry-run

# 正式執行
python main.py
```

## 每週排程（macOS cron）

```
0 9 * * 1 cd "/path/to/carry_out_automation" && python main.py >> automation.log 2>&1
```

## 財資中心承辦人

設定於 `config.py` 的 `MOU_COORDINATORS`：

| 子題 | 承辦人 | Email |
|---|---|---|
| 子題一 | 薛秀英 | n107217@fia.gov.tw |
| 子題二 | 許巧君 | n108266@fia.gov.tw |
| 子題一子議題 | 陳靜 | n108362@fia.gov.tw |
| 稅收估計 | 許宏韜 | N108237@fia.gov.tw |

## 檔案說明

| 檔案 | 說明 |
|---|---|
| `config.py` | 所有設定參數（個人資訊從 `.env` 載入） |
| `sheets_reader.py` | 讀取 / 寫回 Google Drive 申請表 |
| `form_filler.py` | 填寫 .docx + 截圖範例 + 電子簽章 + PDF 轉換 |
| `hitl_server.py` | 本機 HITL 審核 Web 伺服器 |
| `email_sender.py` | 發送審核 / 財資中心通知 / 退回 Email |
| `main.py` | 主流程 |
| `.env` | 個人資訊（不提交） |
| `.env.gemini` | Gemini API 金鑰（不提交） |
| `*.json` | Google 服務帳戶金鑰（不提交） |
| `signature.png` | 電子簽章圖片（不提交） |
