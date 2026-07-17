## Project Overview

This is an **internal contract management system** developed independently as a personal work output for my employer.

The system provides a serverless, password-protected web interface for uploading, indexing, and searching company contracts, with an AI assistant for natural-language queries. All data lives in Google Workspace — no external server required.

Built end to end as a solo project: system architecture, AI integration, RAG pipeline, frontend UI, and deployment.

---

## Approach & Methods

- **Storage & Database:**
  - Contracts stored in Google Drive, organized by company folder
  - Google Sheets used as the database (main table, full-text index, company code map)
  - Text extracted from PDF and Word files via Google Drive OCR (converting to Google Docs and reading the body text)

- **AI Integration (RAG):**
  - Implemented a two-layer retrieval pipeline before passing context to Gemini
  - Layer 1: keyword match against structured fields (Contract ID, company name, date, old ID)
  - Layer 2: keyword match against full extracted text cached in a hidden sheet (`_FullTextData`)
  - Matched contracts (up to 5, 1500 chars each) sent as context to Gemini; responses include clickable Drive links

- **Web Interface:**
  - Full HTML/CSS/JS UI embedded in a single Apps Script `doGet()` function
  - Three-tab layout: Search/Sync, Upload/File, AI Assistant
  - Left-side slide-out documentation panel that does not overlay the main content
  - Password gate before the main interface

- **Auto Contract ID:**
  - Company codes derived from name initials (with Chinese pinyin initial support and collision resolution)
  - Sequential numbering per company (`FME-0001`, `FME-0002`, etc.)
  - Live preview shown before upload

---

## Technical Challenges

- **Gemini rate limits on a corporate Google Workspace account:** Free-tier quotas were near zero for most models (`gemini-2.0-flash`, `gemini-1.5-flash`). Diagnosed by testing four models individually; only `gemini-flash-latest` returned HTTP 200. RAG pre-filtering also reduced the number of API calls needed per query.

- **Deployed web app showing 404 while direct function test showed 200:** The web app was serving a cached old deployment version. Identified that Apps Script requires explicitly selecting "New Version" in Manage Deployments — redeploying with the same version does not update the live URL.

- **Backfill function with no visible progress:** Original implementation used an unconditional 5-second sleep and batched all writes at the end, making it appear to hang. Refactored to write each cell immediately after the AI call, added strict skip logic for non-empty cells, and set a 4.3-minute safety timeout (under Apps Script's 6-minute execution limit) so the function can be re-run incrementally.

- **AI responses containing raw Unicode separator characters:** Gemini occasionally returned `─────────` separator lines in its output. Added a regex filter to strip lines consisting entirely of Unicode box-drawing or dash characters before rendering the HTML response.

---

## Key Features

- Multi-condition contract search: Contract ID (new or old), company, and full-text keyword
- One-click Drive sync: scans for newly added and trashed files, updates the index
- Upload with automatic ID assignment and AI-extracted legacy contract number
- AI assistant with two-layer RAG search and clickable contract links in responses
- Spreadsheet-side menu tools: sync, backfill old IDs, clean up deleted records
- Local Python scripts for pre-processing and classifying contracts in bulk before system import

---

## Tools & Technologies

- **Backend:** Google Apps Script (serverless)
- **Frontend:** HTML / CSS / JavaScript (embedded)
- **Database:** Google Sheets
- **File Storage:** Google Drive
- **AI:** Gemini API (`gemini-flash-latest`)
- **Text Extraction:** Google Drive OCR (PDF / Word to Google Docs)
- **Caching:** Apps Script CacheService (full-text map, 5-minute TTL)
- **Local Preprocessing:** Python, pdfplumber, python-docx, openpyxl

---

## Project Structure

```
contract-management-system/
├── ContractSystem_v2.gs        Main Apps Script file — UI, backend logic, AI integration
├── classify_by_content.py      Local script: classify unindexed contracts by content keyword
├── process_contracts.py        Local script: batch scan PDF/Word files and export an Excel list
├── requirements.txt
└── README.md
```

---

## Deployment

1. Open [Google Apps Script](https://script.google.com) and create a new project
2. Paste the contents of `ContractSystem_v2.gs` into the editor
3. Fill in the four config variables at the top of the file:

```javascript
var ROOT_FOLDER_ID = 'YOUR_GOOGLE_DRIVE_FOLDER_ID';
var SHEET_ID       = 'YOUR_GOOGLE_SHEET_ID';
var SITE_PASSWORD  = 'YOUR_SITE_PASSWORD';
var GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY';  // https://aistudio.google.com/apikey
```

4. Enable two services under **Services**: Drive API v3, Sheets API v4
5. Click **Deploy > New deployment > Web app**, set access to "Anyone"
6. Copy the deployment URL

To bind the spreadsheet menu tools (`onOpen`, `syncAll`, etc.), open the target Google Sheet, go to **Extensions > Apps Script**, and paste the same file there.

---

## Local Preprocessing Scripts

```bash
pip install -r requirements.txt
python process_contracts.py       # Scan local folder, export Excel list
python classify_by_content.py     # Classify unindexed files by content keyword
```
