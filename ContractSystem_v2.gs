/****************************************
 * 合同管理系統 - Apps Script 終極全端版（防呆排序版）
 * 含：網頁端上傳、動態預覽 ID、手動快速同步、AI智能提取
 ****************************************/

var ROOT_FOLDER_ID = 'YOUR_GOOGLE_DRIVE_FOLDER_ID';   // Google Drive 根資料夾 ID
var SHEET_ID = 'YOUR_GOOGLE_SHEET_ID';               // Google Sheets 試算表 ID
var SHEET_NAME = 'contract';
var FULLTEXT_SHEET_NAME = '_FullTextData';
var COMPANY_MAP_SHEET_NAME = '_CompanyMap';
var HEADER = ['Contract ID','Contract Name','Company','Submit Time','Contract Link','Contract Excerpt','File ID','Old Contract ID'];

var SITE_PASSWORD = 'YOUR_SITE_PASSWORD';            // 自訂網頁入口密碼
var GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY';          // 前往 Google AI Studio 申請：https://aistudio.google.com/apikey

// ==========================================
// 核心輔助工具 (移至最上方確保優先載入)
// ==========================================
function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function checkPassword(pwd) {
  return pwd === SITE_PASSWORD;
}

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function ensureFullTextSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(FULLTEXT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FULLTEXT_SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([['File ID', 'Contract ID', 'Full Text']]);
    sheet.hideSheet();
  }
  return sheet;
}

function ensureHeader(sheet) {
  var firstRow = sheet.getRange(1,1,1,HEADER.length).getValues()[0];
  if (firstRow.join('') === '') { sheet.getRange(1,1,1,HEADER.length).setValues([HEADER]); sheet.setFrozenRows(1); }
}

function getCodeMap() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(COMPANY_MAP_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(COMPANY_MAP_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['Company Name', 'Code']]);
    var initialMap = {
      '2P Entertainment (Macau) Ltd':'2PE', '2P Workshop':'2PW', 'Delia Limited':'DEL',
      'FM Event Limited':'FME', 'FM Group (Holdings) Ltd':'FMG', 'FM Investment Limited':'FMI',
      'FM Projects (HK) Limited':'FMPHK', 'FM Projects Limited':'FMP', 'FM Telemedia Limited':'FMT',
      'FM Travel Agent Ltd':'FMTA', 'FUN FUN CHANGE CO':'FFC', 'Film Mall Entertainment Limited':'FMEN',
      'Film Mall Limited':'FML', 'Film Mall Producao Limitada':'FMPA', 'Film Mall Production Limited':'FMPL',
      'Love Smart':'LSM', 'TIEN RIVER - BVI':'TRB', 'Vanuatu':'VAN', '深圳影市堂':'FMSZ', '花蜜项目策划(深圳)有限公司':'HMSZ'
    };
    var initialRows = [];
    for (var key in initialMap) { initialRows.push([key, initialMap[key]]); }
    if (initialRows.length > 0) sheet.getRange(2, 1, initialRows.length, 2).setValues(initialRows);
    sheet.setFrozenRows(1);
  }
  var lastRow = sheet.getLastRow();
  var map = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 2).getValues().forEach(function(row) {
      if (row[0]) map[row[0].toString().trim()] = row[1].toString().trim();
    });
  }
  return map;
}

function generateUniqueCode(companyName, currentMap) {
  var code = "";
  var hasChinese = /[一-龥]/.test(companyName);
  var existingCodes = [];
  for (var key in currentMap) existingCodes.push(currentMap[key].toUpperCase());
  if (hasChinese) {
    for (var i = 0; i < companyName.length; i++) {
      var char = companyName.charAt(i);
      if (/[a-zA-Z0-9]/.test(char)) code += char.toUpperCase();
      else if (/[一-龥]/.test(char)) code += getChineseInitial(char);
    }
    return resolveCollision(code, existingCodes);
  } else {
    var words = companyName.split(/[\s_\-\(\)\.,]+/);
    var validWords = [];
    for (var j = 0; j < words.length; j++) {
      var w = words[j].toLowerCase();
      if (w.length > 0 && w !== "ltd" && w !== "limited") validWords.push(words[j]);
    }
    if (validWords.length === 0) validWords = words.filter(function(w){ return w.length > 0; });
    var baseCode = "";
    validWords.forEach(function(w) { baseCode += w.charAt(0).toUpperCase(); });
    if (existingCodes.indexOf(baseCode) === -1) return baseCode;
    var lastWord = validWords[validWords.length - 1];
    var lastChar = lastWord.charAt(lastWord.length - 1).toUpperCase();
    return resolveCollision(baseCode + lastChar, existingCodes);
  }
}

function resolveCollision(baseCode, existingCodes) {
  if (existingCodes.indexOf(baseCode) === -1) return baseCode;
  var counter = 1;
  while (existingCodes.indexOf(baseCode + counter) !== -1) counter++;
  return baseCode + counter;
}

function getChineseInitial(char) {
  var boundaries = "阿八嚓搭蛾发噶哈击喀垃妈拿哦啪期然撒塌挖昔压匝";
  var letters = "ABCDEFGHJKLMNOPQRSTWXYZ";
  for (var i = boundaries.length - 1; i >= 0; i--) {
    if (char.localeCompare(boundaries.charAt(i), 'zh-CN') >= 0) return letters.charAt(i);
  }
  return "";
}

function testGeminiApi() {
  var models = ['gemini-flash-latest', 'gemini-2.0-flash', 'gemini-2.0-flash-lite', 'gemini-1.5-flash'];
  var payload = { 'contents': [{ 'parts': [{ 'text': '用一句話說你好' }] }] };
  var base = 'https://generativelanguage.googleapis.com/v1beta/models/';
  models.forEach(function(model) {
    var options = {
      'method': 'post', 'contentType': 'application/json',
      'headers': { 'X-goog-api-key': GEMINI_API_KEY },
      'payload': JSON.stringify(payload), 'muteHttpExceptions': true
    };
    var res = UrlFetchApp.fetch(base + model + ':generateContent', options);
    var code = res.getResponseCode();
    Logger.log('[' + model + '] HTTP ' + code);
    if (code === 200) {
      var json = JSON.parse(res.getContentText());
      Logger.log('  ✅ 成功回應: ' + json.candidates[0].content.parts[0].text.trim().substring(0, 80));
    } else {
      var err = JSON.parse(res.getContentText());
      var msg = err.error ? err.error.message.substring(0, 120) : res.getContentText().substring(0, 120);
      Logger.log('  ❌ ' + msg);
    }
  });
}

function aiExtractOldId(text) {
  if (!GEMINI_API_KEY || GEMINI_API_KEY === '請貼上你的_Gemini_API_Key') return "";
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent";
  var prompt = "找出合約『過往的舊合約編號 / 舊ID』。可能標註為 'contract number'、'合同编号'等。只回傳ID字串(如 'FME-001')。無則回 'None'。勿加解釋。\n\n內容：\n" + text.substring(0, 1200);
  var payload = { "contents": [{"parts": [{"text": prompt}]}] };
  var options = { "method": "post", "contentType": "application/json", "headers": { "X-goog-api-key": GEMINI_API_KEY }, "payload": JSON.stringify(payload), "muteHttpExceptions": true };
  var maxRetries = 1;
  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();
      var rawText = response.getContentText();
      if (code === 429 || code === 503) { Utilities.sleep((attempt + 1) * 8000); continue; }
      if (code !== 200) return null;
      var json = JSON.parse(rawText);
      if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) return null;
      var result = json.candidates[0].content.parts[0].text.trim();
      if (result.toLowerCase() === 'none' || result.indexOf(' ') !== -1 || result.length > 30) return "";
      return result;
    } catch (e) { return null; }
  }
  return null;
}

function getPreviewContractId(companyName) {
  if (!companyName) return "";
  var currentMap = getCodeMap();
  var code = currentMap[companyName];
  if (!code) code = generateUniqueCode(companyName, currentMap);
  var seqMap = buildSeqMap(getSheet());
  var seq = (seqMap[code] || 0) + 1;
  return code + '-' + ('0000' + seq).slice(-4);
}

// ==========================================
// 網頁版介面 (Wix 嵌入)
// ==========================================
function doGet(e) {
  var currentCodeMap = getCodeMap();
  var companies = Object.keys(currentCodeMap);
  var options = '<option value="">（請選擇公司或在下方輸入新公司）</option>';
  companies.forEach(function(c) { options += '<option value="' + escapeHtml(c) + '">' + escapeHtml(c) + '</option>'; });

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
    + '<style>body{font-family:Arial,sans-serif;margin:0;background:#fafafa;color:#222;}#pwdScreen{display:flex;align-items:center;justify-content:center;height:100vh;background:#f4f4f4;}.box{background:#fff;padding:32px 28px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,0.1);width:280px;text-align:center;}.box h2{font-size:16px;margin-bottom:18px;color:#333;}.box input{width:100%;padding:10px;box-sizing:border-box;font-size:14px;border:1px solid #ccc;border-radius:6px;margin-bottom:10px;}.box button{width:100%;padding:10px;font-size:14px;background:#1a73e8;color:#fff;border:none;border-radius:6px;cursor:pointer;}#pwdError{color:#d33;font-size:12px;margin-top:6px;min-height:16px;}#mainScreen{display:none;padding:20px;max-width:800px;margin:0 auto;}.tabs{display:flex;margin-bottom:20px;border-bottom:2px solid #ddd;}.tab-btn{flex:1;padding:12px;cursor:pointer;background:none;border:none;font-size:15px;font-weight:bold;color:#666;outline:none;}.tab-btn.active{color:#1a73e8;border-bottom:3px solid #1a73e8;}.tab-content{display:none;background:#fff;padding:20px;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,0.1);}.tab-content.active{display:block;}h2.title{font-size:18px;margin-bottom:16px;margin-top:0;}label{display:block;margin-top:12px;margin-bottom:4px;font-weight:bold;font-size:13px;color:#444;}input[type=text], input[type=file], select{width:100%;padding:10px;box-sizing:border-box;font-size:14px;border:1px solid #ccc;border-radius:6px;background:#fff;}button.actionBtn{padding:12px;font-size:14px;background:#1a73e8;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:bold;transition:0.2s;}button.actionBtn:disabled{background:#aaa;cursor:not-allowed;}table{width:100%;border-collapse:collapse;margin-top:20px;font-size:13px;}th,td{border:1px solid #ddd;padding:10px;text-align:left;}th{background:#f0f0f0;}a{color:#1a73e8;text-decoration:none;font-weight:bold;}#status, #uploadStatus{margin-top:15px;font-size:14px;font-weight:bold;color:#555;}.btn-group{display:flex; gap:10px; margin-top:20px;}.chat-container{height:400px; overflow-y:auto; border:1px solid #ddd; padding:15px; border-radius:8px; background:#f4f7f6; margin-bottom:15px;}.msg-row{margin-bottom:15px; overflow:hidden;}.ai-msg{background:#fff; color:#333; padding:12px 16px; border-radius:12px 12px 12px 0; border:1px solid #e0e0e0; float:left; clear:both; max-width:85%; line-height:1.6; font-size:14px; box-shadow:0 1px 2px rgba(0,0,0,0.05);}.user-msg{background:#1a73e8; color:#fff; padding:12px 16px; border-radius:12px 12px 0 12px; float:right; clear:both; max-width:85%; line-height:1.6; font-size:14px; box-shadow:0 1px 2px rgba(0,0,0,0.1);}.ai-msg ul{margin:8px 0; padding-left:20px;} .ai-msg li{margin-bottom:6px;}.tab-btn{flex:1;padding:12px;cursor:pointer;background:none;border:none;font-size:14px;font-weight:bold;color:#666;outline:none;}'
    + '#helpToggle{position:fixed;left:0;top:50%;transform:translateY(-50%);z-index:1001;background:#1a73e8;color:#fff;border:none;border-radius:0 8px 8px 0;padding:14px 7px;cursor:pointer;writing-mode:vertical-rl;font-size:12px;font-weight:bold;letter-spacing:2px;box-shadow:2px 0 6px rgba(0,0,0,0.25);}'
    + '#helpSidebar{position:fixed;left:0;top:0;width:max(240px,calc((100vw - 840px) / 2));max-width:400px;height:100vh;background:#fff;z-index:1000;box-shadow:3px 0 15px rgba(0,0,0,0.15);overflow-y:auto;transform:translateX(-100%);transition:transform 0.3s ease;}'
    + '#helpSidebar.open{transform:translateX(0);}'
    + '.sb-hdr{background:#1a73e8;color:#fff;padding:14px 16px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;z-index:1;}'
    + '.sb-hdr h3{margin:0;font-size:14px;}'
    + '.sb-close{background:none;border:none;color:#fff;font-size:18px;cursor:pointer;line-height:1;padding:0;}'
    + '.sb-body{padding:14px 16px;}'
    + '.sb-sec{margin-bottom:18px;}'
    + '.sb-sec h4{color:#1a73e8;font-size:11px;font-weight:bold;margin:0 0 8px;padding-bottom:5px;border-bottom:2px solid #e8f0fe;letter-spacing:.5px;}'
    + '.sb-sec p,.sb-sec li{font-size:12px;line-height:1.75;color:#444;margin:3px 0;}'
    + '.sb-sec ul{padding-left:14px;margin:5px 0;}'
    + '.sb-col{background:#f8f9fa;border-radius:5px;padding:5px 8px;margin-bottom:4px;font-size:11.5px;line-height:1.6;}'
    + '.sb-col b{color:#1a73e8;}'
    + '</style></head><body>'

    + '<button id="helpToggle" onclick="toggleSidebar()">📖 說明書</button>'
    + '<div id="helpSidebar">'
    + '  <div class="sb-hdr"><h3>📖 系統使用說明書</h3><button class="sb-close" onclick="toggleSidebar()">✕</button></div>'
    + '  <div class="sb-body">'

    + '  <div class="sb-sec"><h4>🔰 系統概覽</h4>'
    + '  <p>本系統整合 <b>Google Drive</b> 雲端存檔、<b>Google Sheets</b> 資料庫與 <b>Gemini AI</b>，提供合約的上傳、搜索及智能問答。</p>'
    + '  <p><b>系統密碼：</b>27302666</p></div>'

    + '  <div class="sb-sec"><h4>📁 Google Drive 結構</h4><ul>'
    + '  <li>根目錄下按<b>公司名稱</b>分設子資料夾</li>'
    + '  <li>新增公司時系統<b>自動建立</b>對應資料夾</li>'
    + '  <li>支援格式：<b>PDF</b>、<b>Word</b>（.doc / .docx）</li>'
    + '  <li>單份文件上限：<b>8MB</b></li>'
    + '  <li>子資料夾內文件亦會被掃描建檔</li>'
    + '  </ul></div>'

    + '  <div class="sb-sec"><h4>📊 Google Sheets 說明</h4>'
    + '  <p><b>主表 contract（8 欄）：</b></p>'
    + '  <div class="sb-col"><b>① Contract ID</b> — 唯一編號，格式 XXX-0001</div>'
    + '  <div class="sb-col"><b>② Contract Name</b> — 原始文件名稱</div>'
    + '  <div class="sb-col"><b>③ Company</b> — 所屬公司全稱</div>'
    + '  <div class="sb-col"><b>④ Submit Time</b> — 文件上傳日期</div>'
    + '  <div class="sb-col"><b>⑤ Contract Link</b> — Google Drive 直連</div>'
    + '  <div class="sb-col"><b>⑥ Contract Excerpt</b> — 合約前 300 字摘要</div>'
    + '  <div class="sb-col"><b>⑦ File ID</b> — Drive 文件唯一識別碼</div>'
    + '  <div class="sb-col"><b>⑧ Old Contract ID</b> — AI 偵測的舊合約編號</div>'
    + '  <p style="margin-top:8px;"><b>隱藏表 _FullTextData：</b>儲存每份合約完整全文，供 AI 全文搜索使用</p>'
    + '  <p><b>隱藏表 _CompanyMap：</b>公司全稱 ↔ 內部代碼對照（如 FM Event Limited → FME）</p>'
    + '  </div>'

    + '  <div class="sb-sec"><h4>🔍 搜尋 / 同步</h4><ul>'
    + '  <li><b>Contract ID：</b>支援新舊 ID，部分輸入即可（輸入 "FME" 找出全部 FME 合約）</li>'
    + '  <li><b>公司篩選：</b>下拉選擇，精確匹配公司全稱</li>'
    + '  <li><b>內容關鍵字：</b>搜索合約原文全文（如「贊助」「罰款」「獨家」）</li>'
    + '  <li><b>快速同步：</b>掃描 Drive 所有新增及已刪除文件，建議每次上傳後手動同步一次</li>'
    + '  </ul></div>'

    + '  <div class="sb-sec"><h4>📤 上傳 / 建檔</h4><ul>'
    + '  <li>選擇已有公司，或在下方輸入<b>新公司名稱</b>（系統自動建立 Drive 資料夾）</li>'
    + '  <li>Contract ID 由系統自動計算，可手動覆寫</li>'
    + '  <li>上傳後自動執行：提取全文 → AI 識別舊 ID → 寫入試算表</li>'
    + '  <li>預計耗時約 <b>15～20 秒</b></li>'
    + '  </ul></div>'

    + '  <div class="sb-sec"><h4>🤖 AI 助理</h4><ul>'
    + '  <li>用自然語言提問，找到的合約附<b>可點擊連結</b></li>'
    + '  <li><b>第一層搜索：</b>Contract ID、公司名稱、日期、舊 ID（結構化）</li>'
    + '  <li><b>第二層搜索：</b>合約全文內容（關鍵字比對）</li>'
    + '  </ul>'
    + '  <p><b>查詢範例：</b></p><ul>'
    + '  <li>「找所有 FME 的合約」</li>'
    + '  <li>「有沒有提到贊助的合約？」</li>'
    + '  <li>「2024 年之後的合約有哪些？」</li>'
    + '  <li>「找包含罰款條款的合約」</li>'
    + '  <li>「Film Mall 跟誰簽了獨家授權？」</li>'
    + '  </ul>'
    + '  <p>⚠️ 免費版每天 20 次 AI 問答，超出請翌日再試或開通付費方案。</p>'
    + '  </div>'

    + '  <div class="sb-sec"><h4>🔧 試算表選單工具</h4><ul>'
    + '  <li><b>立即同步：</b>新增 + 刪除一次完成</li>'
    + '  <li><b>補全舊 ID：</b>AI 逐份偵測空白欄位，每次最多 15 份，可多次執行</li>'
    + '  <li><b>清理已刪除：</b>只移除已移至垃圾桶的記錄</li>'
    + '  </ul></div>'

    + '  </div>'
    + '</div>'

    + '<div id="pwdScreen"><div class="box"><h2>🔒 合同系統入口</h2><input type="password" id="pwdInput" placeholder="請輸入密碼"><button onclick="submitPwd()">進入</button><div id="pwdError"></div></div></div>'

    + '<div id="mainScreen">'
    + '  <div class="tabs">'
    + '    <button class="tab-btn active" onclick="switchTab(event, \'searchTab\')">🔍 搜尋/同步</button>'
    + '    <button class="tab-btn" onclick="switchTab(event, \'uploadTab\')">📤 上傳/建檔</button>'
    + '    <button class="tab-btn" onclick="switchTab(event, \'aiTab\')">🤖 AI 助理</button>'
    + '  </div>'

    + '  <div id="searchTab" class="tab-content active">'
    + '    <h2 class="title">合同資料庫</h2>'
    + '    <label>Contract ID / 舊 ID</label>'
    + '    <input type="text" id="contractId" placeholder="輸入新編號或舊編號皆可">'
    + '    <label>Company (公司)</label><select id="company">' + options + '</select>'
    + '    <label>內容關鍵字</label><input type="text" id="excerpt" placeholder="搜尋合同全文內容">'
    + '    <div class="btn-group">'
    + '      <button class="actionBtn" onclick="doSearch()" style="flex:2;">🔍 執行搜尋</button>'
    + '      <button id="quickSyncBtn" class="actionBtn" onclick="doQuickSync()" style="flex:1; background:#34a853;">🔄 快速同步</button>'
    + '    </div>'
    + '    <div id="status"></div>'
    + '    <div id="resultArea"></div>'
    + '  </div>'

    + '  <div id="uploadTab" class="tab-content">'
    + '    <h2 class="title">雲端上傳與 AI 建檔</h2>'
    + '    <label>1. 公司名稱</label>'
    + '    <select id="uploadCompany" onchange="updatePreviewId()">' + options + '</select>'
    + '    <input type="text" id="newCompany" placeholder="或在此輸入全新公司名稱 (系統將自動開資料夾)" oninput="updatePreviewId()" style="margin-top:8px;">'
    + '    <label>2. Contract ID (系統已為您計算，可手動覆寫)</label>'
    + '    <input type="text" id="manualContractId" placeholder="選擇公司後，系統會自動算出下一個 ID..." style="background:#e8f0fe; color:#1a73e8; font-weight:bold;">'
    + '    <label>3. 選擇合同檔案 (支援 PDF, Word，上限 8MB)</label>'
    + '    <input type="file" id="fileInput" accept=".pdf,.doc,.docx">'
    + '    <button id="uploadBtn" class="actionBtn" onclick="doUpload()" style="width:100%; margin-top:20px;">開始上傳並建檔</button>'
    + '    <div id="uploadStatus"></div>'
    + '  </div>'

    + '  <div id="aiTab" class="tab-content">'
    + '    <h2 class="title">🤖 老闆專屬 AI 助理</h2>'
    + '    <div id="chatBox" class="chat-container">'
    + '      <div class="msg-row"><div class="ai-msg">老闆您好！我是您的專屬 AI 合約助理。<br>您可以直接用語音口吻吩咐我，例如：<br>👉 「幫我找最新的一份合約」<br>👉 「幫我找這個月 FME 的所有合約」<br>👉 「有沒有跟贊助有關的合約？」</div></div>'
    + '    </div>'
    + '    <div style="display:flex; gap:10px;">'
    + '      <input type="text" id="aiInput" placeholder="請輸入您的指令..." style="flex:1; margin:0;" onkeypress="if(event.key===\'Enter\') sendAiMsg();">'
    + '      <button id="aiSendBtn" class="actionBtn" onclick="sendAiMsg()" style="width:80px; margin:0;">發送</button>'
    + '    </div>'
    + '  </div>'
    + '</div>'

    + '<script>'
    + 'function switchTab(evt, tabId) {'
    + '  document.querySelectorAll(".tab-content").forEach(function(el){ el.classList.remove("active"); });'
    + '  document.querySelectorAll(".tab-btn").forEach(function(el){ el.classList.remove("active"); });'
    + '  document.getElementById(tabId).classList.add("active");'
    + '  evt.currentTarget.classList.add("active");'
    + '}'
    + 'function submitPwd(){'
    + '  var pwd = document.getElementById("pwdInput").value;'
    + '  document.getElementById("pwdError").innerText = "驗證中...";'
    + '  google.script.run.withSuccessHandler(function(ok){'
    + '    if(ok){ document.getElementById("pwdScreen").style.display = "none"; document.getElementById("mainScreen").style.display = "block"; }'
    + '    else{ document.getElementById("pwdError").innerText = "密碼錯誤，請重試"; }'
    + '  }).checkPassword(pwd);'
    + '}'
    + 'document.getElementById("pwdInput").addEventListener("keydown", function(e){ if(e.key === "Enter") submitPwd(); });'

    + 'function doSearch(){'
    + '  document.getElementById("status").innerText = "搜尋中...";'
    + '  document.getElementById("resultArea").innerHTML = "";'
    + '  var cid = document.getElementById("contractId").value, comp = document.getElementById("company").value, exc = document.getElementById("excerpt").value;'
    + '  google.script.run.withSuccessHandler(renderResults).searchContractsWeb(cid, comp, exc);'
    + '}'
    + 'function renderResults(results){'
    + '  document.getElementById("status").innerText = "找到 " + results.length + " 筆結果";'
    + '  if(results.length === 0) return;'
    + '  var html = "<table><tr><th>Contract ID</th><th>Name</th><th>Company</th><th>Submit Time</th><th>連結</th></tr>";'
    + '  results.forEach(function(r){ html += "<tr><td>"+r.id+"</td><td>"+r.name+"</td><td>"+r.company+"</td><td>"+r.submitTime+"</td><td><a href=\\""+r.link+"\\" target=\\"_blank\\">打開</a></td></tr>"; });'
    + '  html += "</table>"; document.getElementById("resultArea").innerHTML = html;'
    + '}'

    + 'function doQuickSync(){'
    + '  var btn = document.getElementById("quickSyncBtn");'
    + '  var stat = document.getElementById("status");'
    + '  btn.disabled = true; btn.innerText = "掃描中...";'
    + '  stat.innerText = "⏳ 正在對接雲端掃描，請稍候...";'
    + '  google.script.run.withSuccessHandler(function(res){'
    + '    stat.innerText = "✅ " + res; btn.disabled = false; btn.innerText = "🔄 快速同步";'
    + '  }).withFailureHandler(function(err){'
    + '    stat.innerText = "❌ 錯誤: " + (err.message || err); btn.disabled = false; btn.innerText = "🔄 快速同步";'
    + '  }).syncAllWeb();'
    + '}'

    + 'var previewTimer;'
    + 'function updatePreviewId(){'
    + '  clearTimeout(previewTimer);'
    + '  var comp = document.getElementById("newCompany").value.trim() || document.getElementById("uploadCompany").value;'
    + '  if(!comp) { document.getElementById("manualContractId").value = ""; return; }'
    + '  document.getElementById("manualContractId").value = "計算專屬 ID 中...";'
    + '  previewTimer = setTimeout(function(){'
    + '    google.script.run.withSuccessHandler(function(id){'
    + '      document.getElementById("manualContractId").value = id;'
    + '    }).getPreviewContractId(comp);'
    + '  }, 500);'
    + '}'

    + 'function doUpload(){'
    + '  var fileInput = document.getElementById("fileInput");'
    + '  var companySel = document.getElementById("uploadCompany").value;'
    + '  var newCompany = document.getElementById("newCompany").value.trim();'
    + '  var manualId = document.getElementById("manualContractId").value.trim();'
    + '  if(manualId.indexOf("計算") !== -1) manualId = "";'

    + '  var finalCompany = newCompany ? newCompany : companySel;'
    + '  if(!finalCompany){ alert("請選擇或填寫公司名稱！"); return; }'
    + '  if(fileInput.files.length === 0){ alert("請選擇要上傳的文件！"); return; }'
    + '  var file = fileInput.files[0];'
    + '  if(file.size > 8*1024*1024){ alert("文件過大！請上傳小於 8MB 的檔案。"); return; }'

    + '  document.getElementById("uploadStatus").innerText = "⏳ 上傳並 AI 智慧建檔中，請稍候 (約 15~20 秒)...";'
    + '  document.getElementById("uploadBtn").disabled = true;'

    + '  var reader = new FileReader();'
    + '  reader.onload = function(e){'
    + '    var data = e.target.result;'
    + '    var base64 = data.split(",")[1];'
    + '    google.script.run'
    + '      .withSuccessHandler(function(res){'
    + '         document.getElementById("uploadStatus").innerText = res;'
    + '         document.getElementById("uploadBtn").disabled = false;'
    + '         if(res.indexOf("✅") !== -1){'
    + '            fileInput.value = ""; document.getElementById("newCompany").value = ""; updatePreviewId();'
    + '         }'
    + '      })'
    + '      .withFailureHandler(function(err){'
    + '         var errMsg = err ? (err.message || err.toString() || err) : "網路中斷或檔案過大";'
    + '         document.getElementById("uploadStatus").innerText = "❌ 系統錯誤: " + errMsg;'
    + '         document.getElementById("uploadBtn").disabled = false;'
    + '      })'
    + '      .uploadFileWeb(base64, file.type, file.name, finalCompany, manualId);'
    + '  };'
    + '  reader.onerror = function(){'
    + '    document.getElementById("uploadStatus").innerText = "❌ 讀取檔案失敗";'
    + '    document.getElementById("uploadBtn").disabled = false;'
    + '  };'
    + '  reader.readAsDataURL(file);'
    + '}'

    + 'function sendAiMsg() {'
    + '  var input = document.getElementById("aiInput");'
    + '  var msg = input.value.trim();'
    + '  if(!msg) return;'
    + '  var chatBox = document.getElementById("chatBox");'
    + '  chatBox.innerHTML += \'<div class="msg-row"><div class="user-msg">\' + msg + \'</div></div>\';'
    + '  var loadingId = "load_" + new Date().getTime();'
    + '  chatBox.innerHTML += \'<div class="msg-row" id="\' + loadingId + \'"><div class="ai-msg">🔍 正在為您翻閱資料庫，請稍候...</div></div>\';'
    + '  chatBox.scrollTop = chatBox.scrollHeight;'
    + '  input.value = "";'
    + '  document.getElementById("aiSendBtn").disabled = true;'
    + '  google.script.run'
    + '    .withSuccessHandler(function(res){'
    + '      document.getElementById(loadingId).remove();'
    + '      chatBox.innerHTML += \'<div class="msg-row"><div class="ai-msg">\' + res + \'</div></div>\';'
    + '      chatBox.scrollTop = chatBox.scrollHeight;'
    + '      document.getElementById("aiSendBtn").disabled = false;'
    + '    })'
    + '    .withFailureHandler(function(err){'
    + '      document.getElementById(loadingId).remove();'
    + '      chatBox.innerHTML += \'<div class="msg-row"><div class="ai-msg" style="color:red;">❌ 連線異常，請再試一次。</div></div>\';'
    + '      chatBox.scrollTop = chatBox.scrollHeight;'
    + '      document.getElementById("aiSendBtn").disabled = false;'
    + '    })'
    + '    .chatWithAI(msg);'
    + '}'

    + 'function toggleSidebar(){var sb=document.getElementById("helpSidebar");var btn=document.getElementById("helpToggle");var isOpen=sb.classList.toggle("open");btn.style.display=isOpen?"none":"";}'
    + '</script></body></html>';

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).setTitle('合同搜索與上傳系統');
}

// ==========================================
// AI 助理
// ==========================================
function chatWithAI(userMessage) {
  // 提取關鍵字（長度 > 1，過濾標點）
  var keywords = userMessage.toLowerCase()
    .split(/[\s？?！!。，,、「」【】（）()\.\-\/]+/)
    .filter(function(w) { return w.length > 1; });

  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  var matchedFileIds = {};   // { fileId: true }
  var contractMetaMap = {};  // { fileId: rowArray }
  var mainData = [];

  // ── 第一層：搜主表結構化欄位 ──────────────────────────
  if (lastRow >= 2) {
    mainData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    mainData.forEach(function(row) {
      var fileId = row[6];
      if (!fileId) return;
      contractMetaMap[fileId] = row;
      if (keywords.length === 0) return;
      var year = row[3] ? new Date(row[3]).getFullYear().toString() : '';
      var searchStr = [row[0], row[1], row[2], row[7], year].join('|').toLowerCase();
      if (keywords.some(function(k) { return searchStr.indexOf(k) !== -1; })) {
        matchedFileIds[fileId] = true;
      }
    });
  }

  // ── 第二層：搜 _FullTextData 合約全文 ────────────────
  var ftSheet = ensureFullTextSheet();
  var ftLastRow = ftSheet.getLastRow();
  var ftData = [];
  if (ftLastRow >= 2) {
    ftData = ftSheet.getRange(2, 1, ftLastRow - 1, 3).getValues();
    if (keywords.length > 0) {
      ftData.forEach(function(row) {
        var fileId = row[0];
        if (!fileId || !row[2]) return;
        var fullTextLower = row[2].toString().toLowerCase();
        if (keywords.some(function(k) { return fullTextLower.indexOf(k) !== -1; })) {
          matchedFileIds[fileId] = true;
        }
      });
    }
  }

  // ── 組裝 AI 上下文 ────────────────────────────────────
  var matchedIds = Object.keys(matchedFileIds);
  var contextParts = [];
  var noMatch = matchedIds.length === 0;

  if (noMatch) {
    // 無命中 → 通用模式，只傳最近 10 筆摘要
    var recentRows = mainData.slice(-10);
    var recentLines = recentRows.map(function(row) {
      var date = row[3] ? new Date(row[3]).toLocaleDateString('zh-TW') : '';
      return row[0] + ' | ' + row[1] + ' | ' + row[2] + ' | ' + date + ' | 連結: ' + (row[4] || '');
    });
    contextParts.push('最近 ' + recentLines.length + ' 份合約（格式：ID|名稱|公司|日期|連結）：\n' + recentLines.join('\n'));
  } else {
    // 有命中 → 取全文，最多 5 份、每份 1500 字
    var ftMap = {};
    ftData.forEach(function(row) { if (matchedFileIds[row[0]]) ftMap[row[0]] = row[2]; });

    var shown = 0;
    matchedIds.forEach(function(fileId) {
      if (shown >= 5) return;
      var row = contractMetaMap[fileId];
      var meta = row
        ? '合約ID: ' + row[0] + ' | 公司: ' + row[2]
          + ' | 日期: ' + (row[3] ? new Date(row[3]).toLocaleDateString('zh-TW') : '未知')
          + (row[7] ? ' | 舊ID: ' + row[7] : '')
          + ' | 連結: ' + (row[4] || '')
        : 'File ID: ' + fileId;
      var fullText = ftMap[fileId] ? ftMap[fileId].toString().substring(0, 1500) : '（無全文快取）';
      contextParts.push('【合約 ' + (shown + 1) + '】\n' + meta + '\n內容節錄：\n' + fullText);
      shown++;
    });
    if (matchedIds.length > 5) {
      contextParts.push('（另有 ' + (matchedIds.length - 5) + ' 份相關合約因篇幅省略）');
    }
  }

  var systemPrompt = '你是公司合約管理 AI 助理，服務對象是老闆。\n'
    + (noMatch
        ? '以下是最近合約清單供參考：\n'
        : '以下是根據問題找到的相關合約（共 ' + matchedIds.length + ' 份，顯示前 ' + Math.min(matchedIds.length, 5) + ' 份全文節錄）：\n')
    + contextParts.join('\n\n') + '\n\n'
    + '請用繁體中文回答。列出合約時，每份合約必須用 <a href="連結網址" target="_blank">合約ID</a> 格式輸出可點擊連結。列表用 <ul><li> 格式；換行用 <br>；重點用 <strong>。\n'
    + '老闆問：' + userMessage;

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent';
  var payload = { 'contents': [{ 'parts': [{ 'text': systemPrompt }] }] };
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': { 'X-goog-api-key': GEMINI_API_KEY },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  for (var attempt = 0; attempt < 3; attempt++) {
    try {
      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();
      if (code === 429 || code === 503) { Utilities.sleep((attempt + 1) * 10000); continue; }
      if (code !== 200) return '抱歉，AI 服務暫時無法使用（錯誤碼：' + code + '）';
      var json = JSON.parse(response.getContentText());
      if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) return '抱歉，AI 未能生成回覆，請再試一次。';
      var text = json.candidates[0].content.parts[0].text.trim();
      // Markdown → HTML
      text = text.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
      text = text.replace(/^#{1,3}\s+(.*)/gm, '<strong>$1</strong>');
      // 移除 Gemini 常用的 Unicode 分隔線與 markdown 水平線
      text = text.replace(/^[─━═\-_]{3,}\s*$/gm, '');
      var textLines = text.split('\n');
      var out = '', inList = false;
      for (var i = 0; i < textLines.length; i++) {
        var line = textLines[i].trim();
        if (!line) continue;
        if (/^[-*•]\s/.test(line)) {
          if (!inList) { out += '<ul>'; inList = true; }
          out += '<li>' + line.replace(/^[-*•]\s+/, '') + '</li>';
        } else {
          if (inList) { out += '</ul>'; inList = false; }
          out += (out ? '<br>' : '') + line;
        }
      }
      if (inList) out += '</ul>';
      return out || '（AI 未回傳有效內容，請再試一次）';
    } catch (e) {
      return '❌ 連線異常：' + e.message;
    }
  }
  return '⏳ AI 目前仍在限流中，請稍後再試。';
}

// ==========================================
// 系統核心處理邏輯
// ==========================================
function uploadFileWeb(base64Data, mimeType, fileName, companyName, manualContractId) {
  try {
    var root = DriveApp.getFolderById(ROOT_FOLDER_ID);
    var folders = root.getFoldersByName(companyName);
    var targetFolder = folders.hasNext() ? folders.next() : root.createFolder(companyName);

    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    var newFile = targetFolder.createFile(blob);

    var currentMap = getCodeMap();
    var code = currentMap[companyName];
    if (!code) {
       code = generateUniqueCode(companyName, currentMap);
       SpreadsheetApp.openById(SHEET_ID).getSheetByName(COMPANY_MAP_SHEET_NAME).appendRow([companyName, code]);
       currentMap[companyName] = code;
    }

    var contractId = manualContractId ? manualContractId.trim() : "";
    if (!contractId || contractId.indexOf("計算") !== -1) {
       var seqMap = buildSeqMap(getSheet());
       var seq = (seqMap[code] || 0) + 1;
       contractId = code + '-' + ('0000' + seq).slice(-4);
    }

    var fullText = extractFullText(newFile);
    var excerpt = fullText.substring(0, 300);
    var oldContractId = aiExtractOldId(fullText);

    var sheet = getSheet();
    var fullTextSheet = ensureFullTextSheet();
    ensureHeader(sheet);

    sheet.appendRow([contractId, newFile.getName(), companyName, newFile.getDateCreated(), newFile.getUrl(), excerpt, newFile.getId(), oldContractId || "（未偵測到舊ID）"]);
    fullTextSheet.appendRow([newFile.getId(), contractId, fullText]);

    return "✅ 成功！已建立。\n合約 ID：" + contractId + "\n舊 ID：" + (oldContractId || "無");
  } catch (e) {
    return "❌ 系統錯誤：" + e.message;
  }
}

function syncAllWeb() {
  syncAll();
  return "同步完成！所有新增/刪除皆已更新。";
}

function syncAll() { removeDeletedFiles(); scanAndFillContracts(); }

function scanAndFillContracts() {
  var startTime = new Date().getTime();
  var sheet = getSheet();
  var fullTextSheet = ensureFullTextSheet();
  ensureHeader(sheet);
  var existingIds = getExistingFileIds(sheet);
  var seqMap = buildSeqMap(sheet);
  var rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
  var companyFolders = rootFolder.getFolders();
  var newRows = [];
  var newFullTextRows = [];
  var isTimeout = false;
  var currentCodeMap = getCodeMap();
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var mapSheet = ss.getSheetByName(COMPANY_MAP_SHEET_NAME);

  while (companyFolders.hasNext() && !isTimeout) {
    var companyFolder = companyFolders.next();
    var companyName = companyFolder.getName().trim();
    var code = currentCodeMap[companyName];
    if (!code) {
      code = generateUniqueCode(companyName, currentCodeMap);
      mapSheet.appendRow([companyName, code]);
      currentCodeMap[companyName] = code;
    }
    isTimeout = scanFolderRecursive(companyFolder, companyName, code, existingIds, seqMap, newRows, newFullTextRows, startTime, 260000);
  }

  if (newRows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, HEADER.length).setValues(newRows);
  if (newFullTextRows.length > 0) fullTextSheet.getRange(fullTextSheet.getLastRow() + 1, 1, newFullTextRows.length, 3).setValues(newFullTextRows);
}

function scanFolderRecursive(folder, companyName, code, existingIds, seqMap, newRows, newFullTextRows, startTime, maxTimeMs) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    if (new Date().getTime() - startTime > maxTimeMs) return true;
    var file = files.next();
    var fileId = file.getId();
    if (existingIds.has(fileId)) continue;
    var mime = file.getMimeType();
    if (mime !== MimeType.PDF && mime !== MimeType.MICROSOFT_WORD && mime !== MimeType.GOOGLE_DOCS) continue;

    var seq = (seqMap[code] || 0) + 1;
    seqMap[code] = seq;
    var contractId = code + '-' + ('0000' + seq).slice(-4);
    var fullText = extractFullText(file);
    var excerpt = fullText.substring(0, 300);
    var oldContractId = aiExtractOldId(fullText);
    if (oldContractId === null) oldContractId = "";

    newRows.push([contractId, file.getName(), companyName, file.getDateCreated(), file.getUrl(), excerpt, fileId, oldContractId]);
    newFullTextRows.push([fileId, contractId, fullText]);
    existingIds.add(fileId);
  }
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    if (scanFolderRecursive(subFolders.next(), companyName, code, existingIds, seqMap, newRows, newFullTextRows, startTime, maxTimeMs)) return true;
  }
  return false;
}

function backfillOldContractIds() {
  Logger.log('🚀 開始執行舊 ID【極速快進、嚴格跳過版】任務...');
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('沒有資料需要補全！');
    return;
  }

  if (sheet.getMaxColumns() < HEADER.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADER.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 8).setValue('Old Contract ID');

  var fullTextMap = buildFullTextMap();
  var mainData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var updatedCount = 0, quotaSkippedCount = 0;
  var startTime = new Date().getTime();

  for (var i = 0; i < mainData.length; i++) {
    if (new Date().getTime() - startTime > 260000) {
      Logger.log('⏰ 接近執行時間上限 (4.3 分鐘)，自動安全退出。');
      Logger.log('✅ 本次進度：即時成功寫入了 ' + updatedCount + ' 筆。請再點擊一次繼續往後推進！');
      return;
    }

    var contractId = mainData[i][0];
    var fileId = mainData[i][6];
    var currentOldId = mainData[i][7] ? mainData[i][7].toString().trim() : "";
    var rowNum = i + 2;

    if (currentOldId !== "") {
      continue;
    }

    if (fullTextMap[fileId]) {
      Logger.log('🔍 [' + contractId + '] 位於第 ' + rowNum + ' 行 -> 偵測到純白儲存格，觸發 AI 識別...');
      var detectedOldId = aiExtractOldId(fullTextMap[fileId]);

      if (detectedOldId === null) {
        quotaSkippedCount++;
        Logger.log('⏭️ 遭遇限制或異常，暫時跳過第 ' + rowNum + ' 行 (下次執行會再試)');
      } else {
        var finalValue = detectedOldId || "（未偵測到舊ID）";
        sheet.getRange(rowNum, 8).setValue(finalValue);
        mainData[i][7] = finalValue;
        updatedCount++;
        Logger.log('✍️ [即時寫入成功] 第 ' + rowNum + ' 行已寫入 -> ' + finalValue);
      }

      Utilities.sleep(8000);
    } else {
      Logger.log('ℹ️ 第 ' + rowNum + ' 行在隱藏分頁無全文快取，自動跳過');
    }
  }

  Logger.log('🎉 恭喜！整個工作表已全部掃描、推進補全完畢！');
}

function removeDeletedFiles() {
  var sheet = getSheet();
  var fullTextSheet = ensureFullTextSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var fileIds = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
  var deletedIds = {};
  var deletedCount = 0;
  for (var i = fileIds.length - 1; i >= 0; i--) {
    if (!fileIds[i][0]) continue;
    if (!isFileStillValid(fileIds[i][0])) {
      sheet.deleteRow(i + 2);
      deletedIds[fileIds[i][0]] = true;
      deletedCount++;
    }
  }
  if (deletedCount > 0) {
    var ftLastRow = fullTextSheet.getLastRow();
    if (ftLastRow >= 2) {
      var ftIds = fullTextSheet.getRange(2, 1, ftLastRow - 1, 1).getValues();
      for (var j = ftIds.length - 1; j >= 0; j--) {
        if (ftIds[j][0] && deletedIds[ftIds[j][0]]) fullTextSheet.deleteRow(j + 2);
      }
    }
  }
}

function isFileStillValid(fileId) { try { var f = DriveApp.getFileById(fileId); return !f.isTrashed(); } catch (e) { return false; } }

function buildFullTextIdSet(fullTextSheet) {
  var set = new Set();
  var lastRow = fullTextSheet.getLastRow();
  if (lastRow >= 2) fullTextSheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function(r){ if(r[0]) set.add(r[0]); });
  return set;
}

function getExistingFileIds(sheet) {
  var set = new Set(), lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.getRange(2, 7, lastRow - 1, 1).getValues().forEach(function(r){ if (r[0]) set.add(r[0]); });
  return set;
}

function buildSeqMap(sheet) {
  var seqMap = {}, lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function(r){
      if (!r[0]) return;
      var parts = r[0].toString().split('-');
      if (parts.length < 2) return;
      var num = parseInt(parts[1], 10);
      if (!seqMap[parts[0]] || num > seqMap[parts[0]]) seqMap[parts[0]] = num;
    });
  }
  return seqMap;
}

function extractFullText(file) {
  try {
    var blob = file.getBlob();
    var tempDoc = Drive.Files.create({name: 'temp_ocr_' + file.getId(), mimeType: MimeType.GOOGLE_DOCS}, blob);
    var text = DocumentApp.openById(tempDoc.id).getBody().getText();
    DriveApp.getFileById(tempDoc.id).setTrashed(true);
    return text;
  } catch (e) { return '（無法提取文字）'; }
}

function createAutoTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'syncAll' || t.getHandlerFunction() === 'scanAndFillContracts') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncAll').timeBased().everyMinutes(5).create();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('合同工具')
    .addItem('🔍 搜索合同（多條件）', 'showSearchDialog')
    .addItem('♻️ 清除搜索篩選', 'clearSearchFilter')
    .addItem('🔄 立即同步（新增+刪除）', 'syncAll')
    .addItem('🧠 一鍵補全新創立合同的舊ID', 'backfillOldContractIds')
    .addItem('🗑️ 僅清理已刪除文件', 'removeDeletedFiles')
    .addToUi();
}

function searchContractsWeb(contractId, company, excerptKeyword) {
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
  var fullTextMap = buildFullTextMap();
  contractId = (contractId || '').trim().toLowerCase();
  company = (company || '').trim();
  excerptKeyword = (excerptKeyword || '').trim().toLowerCase();
  var results = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var matchId = !contractId || row[0].toString().toLowerCase().indexOf(contractId) !== -1 || (row[7] && row[7].toString().toLowerCase().indexOf(contractId) !== -1);
    var matchCompany = !company || row[2] === company;
    var matchExcerpt = true;
    if (excerptKeyword) {
      var fileId = row[6];
      var fullText = (fullTextMap[fileId] || row[5] || '').toString().toLowerCase();
      matchExcerpt = fullText.indexOf(excerptKeyword) !== -1;
    }
    if (matchId && matchCompany && matchExcerpt) {
      results.push({ id: row[0] + (row[7] ? " (舊: " + row[7] + ")" : ""), name: row[1], company: row[2], submitTime: row[3] ? new Date(row[3]).toLocaleDateString() : '', link: row[4] });
    }
  }
  return results;
}

function buildFullTextMap() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('fullTextMap');
  if (cached) { return JSON.parse(cached); }
  var sheet = ensureFullTextSheet();
  var lastRow = sheet.getLastRow();
  var map = {};
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 3).getValues().forEach(function(r) { if (r[0]) map[r[0]] = r[2]; });
  }
  try { cache.put('fullTextMap', JSON.stringify(map), 300); } catch (e) {}
  return map;
}
