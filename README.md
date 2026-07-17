# 合約管理系統

> 個人工作產出 — 為公司設計並獨立實現的合約管理平台

一套基於 Google Apps Script 的全端合約管理系統，整合 Google Drive 雲端存檔、Google Sheets 資料庫與 Gemini AI，提供合約上傳建檔、多維搜索及 AI 智能問答三大功能，並以密碼保護的網頁作為操作入口。

---

## 功能概覽

### 🔍 搜尋 / 同步
- 支援 Contract ID（新舊編號）、公司名稱、合約全文關鍵字三維搜索
- 一鍵快速同步，自動掃描 Google Drive 新增及已刪除文件

### 📤 上傳 / 建檔
- 支援 PDF、Word（.doc / .docx），上限 8MB
- 系統自動分配 Contract ID（`公司代碼-0001` 格式），可手動覆寫
- 上傳後自動：提取全文 → AI 識別舊合約編號 → 寫入試算表

### 🤖 AI 助理（RAG 雙層搜索）
- 自然語言提問，例如「找所有跟贊助有關的合約」
- **第一層**：搜索合約 ID、公司名稱、日期、舊 ID（結構化欄位）
- **第二層**：搜索合約原文全文（`_FullTextData` 快取）
- 回覆自動附上 Google Drive 可點擊連結

### 📖 左側說明書側欄
- 點擊展開，不遮蓋主系統畫面
- 涵蓋 Drive 結構、Sheets 欄位說明、三大功能操作指引

---

## 技術架構

| 層級 | 技術 |
|---|---|
| 前端 | HTML / CSS / JavaScript（內嵌於 Apps Script） |
| 後端 | Google Apps Script（無伺服器） |
| 資料庫 | Google Sheets（主表 + 全文索引表 + 公司代碼表） |
| 文件存儲 | Google Drive |
| AI | Gemini API（`gemini-flash-latest`） |
| 文字提取 | Google Drive OCR（PDF / Word 轉 Google Docs） |

---

## Google Sheets 結構

**主表 `contract`（8 欄）**

| 欄位 | 說明 |
|---|---|
| Contract ID | 唯一編號，格式 `XXX-0001` |
| Contract Name | 原始文件名稱 |
| Company | 所屬公司全稱 |
| Submit Time | 上傳日期 |
| Contract Link | Google Drive 直連 |
| Contract Excerpt | 合約前 300 字摘要 |
| File ID | Drive 文件唯一識別碼 |
| Old Contract ID | AI 偵測的舊合約編號 |

**隱藏表 `_FullTextData`**：儲存每份合約完整提取文字，供 AI 全文搜索使用

**隱藏表 `_CompanyMap`**：公司全稱 ↔ 內部代碼對照（如 `FM Event Limited → FME`）

---

## 部署方式

1. 前往 [Google Apps Script](https://script.google.com)，建立新專案
2. 將 `ContractSystem_v2.gs` 的內容貼入編輯器
3. 填入以下四項設定：

```javascript
var ROOT_FOLDER_ID = 'YOUR_GOOGLE_DRIVE_FOLDER_ID';  // Google Drive 根資料夾 ID
var SHEET_ID       = 'YOUR_GOOGLE_SHEET_ID';          // Google Sheets 試算表 ID
var SITE_PASSWORD  = 'YOUR_SITE_PASSWORD';             // 自訂網頁密碼
var GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY';            // 前往 Google AI Studio 申請
```

4. 在 Apps Script 左側啟用兩個服務：**Drive API v3**、**Sheets API v4**
5. 點擊「部署 → 新增部署項目 → 網頁應用程式」，存取權設為「所有人」
6. 複製部署 URL 即可使用

> **API Key 申請**：前往 [Google AI Studio](https://aistudio.google.com/apikey) 免費申請 Gemini API Key

---

## 試算表選單工具

開啟綁定的 Google Sheets 後，頂部選單會出現「合同工具」：

- **立即同步**：新增 + 刪除一次完成
- **補全舊 ID**：AI 逐份偵測空白欄位，每次最多處理 15 份（可重複執行直到完成）
- **清理已刪除**：只移除已移至垃圾桶的記錄

---

## 輔助腳本（本地預處理）

| 檔案 | 說明 |
|---|---|
| `process_contracts.py` | 本地批量掃描 PDF / Word，匯出 Excel 清單 |
| `classify_by_content.py` | 依合約內容關鍵字自動分類未歸檔文件 |

```bash
pip install -r requirements.txt
python process_contracts.py
```
