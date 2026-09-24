/**
 * ============================================================================
 * 特搜場控即時看板 — Google Apps Script 後端 (Code.gs)
 * ============================================================================
 *
 * 【這份程式碼要怎麼用】
 * 1. 建立一份新的「看板控制中心」Google Sheet（永久保留、不是每次演習重建的那份）。
 * 2. 打開該 Sheet 的「擴充功能 > Apps Script」，把這份 Code.gs 的內容整個貼進去。
 * 3. 在「看板控制中心」這份 Sheet 裡建立兩個分頁（名稱要完全一樣，含全形/半形）：
 *
 *    A. 分頁名稱：「設定」
 *       第 1 列放標題（隨意），從第 2 列開始，每一列是一組「項目 / 值」：
 *         A 欄                          B 欄
 *         場控時序檔案網址或ID           <貼上「10月場控時序.xlsx」上傳到雲端硬碟後、
 *                                        或另存成 Google Sheet 後的網址或檔案ID>
 *         演習開始日期                   2026-10-15   ← 對應場控時序表 columns 的第一欄（09:00）
 *
 *       程式是用「A欄文字」去找對應的 B 欄值，所以 A 欄文字要跟上面列的完全一致，
 *       但列的先後順序、中間有沒有空白列都不影響（見 readConfigValue 函式）。
 *
 *    B. 分頁名稱：「處置記錄」
 *       第 1 列放標題：ID | 時間戳記 | 事件分類 | 狀態 | 內容備註
 *       從第 2 列開始，每新增一筆記錄就是新增一列（由 addLog 動作自動寫入，
 *       使用者也可以自己在 Google Sheet 手動補登一列）。
 *
 * 4. 部署：Apps Script 編輯器右上角「部署 > 新增部署作業」，類型選「網頁應用程式」，
 *    「具有存取權的使用者」選「所有人」（因為場控組多人要用不同裝置存取，且沒有登入機制），
 *    存取權執行身分選「以我的身分執行」。部署後會得到一個
 *    https://script.google.com/macros/s/xxxxx/exec 網址，把這個網址貼到
 *    看板前端（index.html）右上角「設定」面板的「Apps Script 網址」欄位。
 * 5. 第一次部署或改過程式碼重新部署後，Google 會要求你「授權」（因為要讀寫 Google Sheet）。
 *    照畫面指示選擇自己的帳號、允許存取即可。
 *
 * 【重要假設，之後如果範本欄列順序改變，改這裡就好，不用重寫整支程式】
 * - 這支程式假設「場控時序」原始檔案裡，時序主表所在分頁固定叫做「場控時序」
 *   （對應 場控時序_結構分析.md 的分析結果）。
 * - 下面的 ROW_LABELS / GROUP_NAMES 是「看板要顯示」的標籤（給前端 API 用的固定命名），
 *   ROW_LABEL_SHEET_ALIASES 則是「這些標籤在實際 Excel/Sheet 裡 A 欄真正寫的文字」，
 *   因為分析報告發現兩者用字有些微差異（例如 Excel 裡寫「OCC隊長會議」，看板要顯示「OCC會議」），
 *   如果你拿到的場控時序檔案欄位文字不一樣，只要改這個對照表即可，不用改其他程式邏輯。
 * - 這支程式碼的 doGet / doPost 邏輯是依照「場控時序_結構分析.md」與使用者確認過的
 *   API contract 純邏輯推演寫成，並未實際連線到真正的 Google Sheet 執行測試
 *   （Apps Script 需要在瀏覽器登入 Google 部署後才能真正測試），
 *   請部署後務必先用瀏覽器打開 exec 網址 + ?action=timeline 實際看一次回傳的 JSON 內容再上線使用。
 */

// ===== 設定分頁 / 處置記錄分頁的名稱 =====
var SETTINGS_SHEET_NAME = '設定';
var LOG_SHEET_NAME = '處置記錄';

// ===== 場控時序原始檔案裡，時序主表所在分頁的名稱 =====
var SOURCE_TIMELINE_SHEET_NAME = '場控時序';

// ===== 看板要抓的「固定類別列」標籤（API 回傳給前端用的標籤，跟 app.js 的 ROW_FIELD_MAP 要一致）=====
var ROW_LABELS = ['階段', '區域', '傷患編號', 'INJECT', 'LEMA會議', 'OCC會議', '拍照攝影', '救護車司機', '後送醫院'];

// 上面 ROW_LABELS 對應到「場控時序」分頁 A 欄實際寫的文字。
// 如果 Excel A 欄文字跟 ROW_LABELS 一模一樣，這裡可以不用列出來（會直接 fallback 用同一個字串比對）。
var ROW_LABEL_SHEET_ALIASES = {
  'OCC會議': 'OCC隊長會議',
  '救護車司機': 'LEMA救護車司機'
};

// ===== 六個小組名稱（API 回傳給前端用的標籤）=====
var GROUP_NAMES = ['總控A', '總控B', '現場A', '現場B', '傷控A', '傷控B'];

// 場控時序分頁裡，組別標題列實際文字是「總控A組」這種「名稱+組」的格式
var GROUP_NAME_SHEET_SUFFIX = '組';

// 每個組別區塊固定 4 列：組名（標題，不使用）/ 任務 / 負責人 / 位置
var GROUP_BLOCK_OFFSET_TASK = 1;
var GROUP_BLOCK_OFFSET_OWNER = 2;
var GROUP_BLOCK_OFFSET_LOCATION = 3;

// ===== 休息時段的判斷依據 =====
// 場控時序_結構分析.md 記載：休息時段固定用淺橘色 FFFAE2D5（ARGB），
// Apps Script 的 Range.getBackgrounds() 回傳格式是不含 alpha 的 "#rrggbb" 小寫，
// 所以這裡設成 "#fae2d5"。如果之後範本改了休息底色，改這個常數即可。
var REST_BACKGROUND_COLOR = '#fae2d5';

// 時間軸最多掃描到第幾欄（B欄=第2欄開始），設寬一點確保涵蓋完整 36 小時演習，
// 有需要的話可以調大這個數字。
var MAX_TIME_AXIS_COLUMN = 60;

// A 欄（標籤欄）最多掃描到第幾列，用來找各列標籤與組別區塊的位置
var MAX_LABEL_ROW = 60;


/**
 * ============================================================================
 * doGet — 處理所有「讀取」類型的請求
 * 前端呼叫方式：
 *   GET {APPS_SCRIPT_URL}?action=timeline  → 回傳目前時段資料
 *   GET {APPS_SCRIPT_URL}?action=log       → 回傳處置記錄列表
 * ============================================================================
 */
function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;
  var result;
  try {
    if (action === 'timeline') {
      result = handleGetTimeline();
    } else if (action === 'log') {
      result = handleGetLog();
    } else {
      result = { error: '未知的 action 參數：' + action };
    }
  } catch (err) {
    result = { error: '伺服器處理發生錯誤：' + err };
  }
  return jsonResponse(result);
}


/**
 * ============================================================================
 * doPost — 處理所有「寫入」類型的請求
 * 前端呼叫方式（body 是 JSON 字串）：
 *   {"action":"addLog", "timestamp":"...", "category":"...", "status":"...", "note":"..."}
 *   {"action":"setConfig", "spreadsheetUrl":"...", "exerciseDate":"2026-11-XX"}
 * ============================================================================
 */
function doPost(e) {
  var result;
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.action === 'addLog') {
      result = handleAddLog(payload);
    } else if (payload.action === 'setConfig') {
      result = handleSetConfig(payload);
    } else {
      result = { success: false, error: '未知的 action：' + payload.action };
    }
  } catch (err) {
    result = { success: false, error: '伺服器處理發生錯誤：' + err };
  }
  return jsonResponse(result);
}


/**
 * 把 JavaScript 物件轉成 JSON 回應。
 * 注意：Apps Script Web App 的 doGet 對一般瀏覽器 fetch(GET) 已足夠支援跨網域讀取，
 * doPost 那邊前端故意用 text/plain 送出 body 是為了閃避瀏覽器 CORS 的 preflight 檢查
 * （Apps Script 網頁應用程式不支援自訂 CORS 標頭，這是社群公認的繞過方式）。
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


/**
 * ============================================================================
 * 讀取「設定」分頁裡的設定值
 * ============================================================================
 * 用「A欄文字 = label」去找對應的 B 欄值，不依賴固定的儲存格位置或列的先後順序。
 */
function readConfigValue(settingsSheet, label) {
  var lastRow = settingsSheet.getLastRow();
  if (lastRow < 1) return '';
  var values = settingsSheet.getRange(1, 1, lastRow, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    var rowLabel = String(values[i][0]).trim();
    if (rowLabel === label) {
      return String(values[i][1]).trim();
    }
  }
  return '';
}

/**
 * 把使用者貼的 Google Sheet 網址轉成純 ID；如果本來就是 ID（不含網址格式），原樣回傳。
 */
function extractSpreadsheetId(urlOrId) {
  var match = String(urlOrId).match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];
  return String(urlOrId).trim();
}

/**
 * 開啟「看板控制中心」自己這份 Sheet 的「設定」分頁，回傳 { spreadsheetId, exerciseDate }
 */
function getActiveConfig() {
  var controlSheet = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = controlSheet.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    throw new Error('找不到「' + SETTINGS_SHEET_NAME + '」分頁，請檢查看板控制中心 Sheet 的分頁名稱設定');
  }
  var spreadsheetUrlOrId = readConfigValue(settingsSheet, '場控時序檔案網址或ID');
  var exerciseDate = readConfigValue(settingsSheet, '演習開始日期');
  if (!spreadsheetUrlOrId) {
    throw new Error('「設定」分頁裡尚未填寫「場控時序檔案網址或ID」');
  }
  return {
    spreadsheetId: extractSpreadsheetId(spreadsheetUrlOrId),
    exerciseDate: exerciseDate
  };
}


/**
 * ============================================================================
 * handleGetTimeline — 讀取場控時序主表，組成前端要的 JSON 格式
 * ============================================================================
 */
function handleGetTimeline() {
  var config = getActiveConfig();
  var sourceSpreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  var sheet = sourceSpreadsheet.getSheetByName(SOURCE_TIMELINE_SHEET_NAME);
  if (!sheet) {
    throw new Error('場控時序檔案裡找不到「' + SOURCE_TIMELINE_SHEET_NAME + '」分頁');
  }

  // ---- 第一步：找出所有「有值」的時間軸欄位（B欄=第2欄起），並記錄它們在原表格的實際欄位編號 ----
  // 依照 場控時序_結構分析.md 的提醒：時間軸欄位可能有跳號（合併儲存格造成的排版空欄），
  // 所以「以有值的時間欄為準」而非假設每一欄都固定代表一小時往後推。
  var headerRowValues = sheet.getRange(1, 1, 1, MAX_TIME_AXIS_COLUMN).getValues()[0];
  var columns = []; // 要回傳給前端的時間文字，例如 "09:00"
  var sourceColumnIndexes = []; // 每個 columns[i] 對應到原始表格的第幾欄（1-based）

  for (var c = 1; c < headerRowValues.length; c++) { // c=1 代表第2欄（B欄），因為陣列是 0-based
    var cellValue = headerRowValues[c];
    if (cellValue === '' || cellValue === null) continue; // 跳過沒有時間值的空欄
    var timeText;
    if (Object.prototype.toString.call(cellValue) === '[object Date]') {
      timeText = Utilities.formatDate(cellValue, sourceSpreadsheet.getSpreadsheetTimeZone(), 'HH:mm');
    } else {
      timeText = String(cellValue);
    }
    columns.push(timeText);
    sourceColumnIndexes.push(c + 1); // 轉回 1-based 的實際欄位編號給 getRange 用
  }

  // ---- 第二步：用 A 欄文字建立「標籤 → 列號」的對照表 ----
  var labelColumnValues = sheet.getRange(1, 1, MAX_LABEL_ROW, 1).getValues();
  var rowNumberByLabel = {};
  for (var r = 0; r < labelColumnValues.length; r++) {
    var label = String(labelColumnValues[r][0]).trim();
    if (label) {
      rowNumberByLabel[label] = r + 1; // 轉成 1-based 列號
    }
  }

  // ---- 第三步：依 ROW_LABELS 逐列，把每個有效時間欄的值抓出來 ----
  var rows = [];
  for (var li = 0; li < ROW_LABELS.length; li++) {
    var outputLabel = ROW_LABELS[li];
    var sheetLabel = ROW_LABEL_SHEET_ALIASES[outputLabel] || outputLabel;
    var rowNumber = rowNumberByLabel[sheetLabel];
    var values = [];
    if (rowNumber) {
      for (var ci = 0; ci < sourceColumnIndexes.length; ci++) {
        var val = sheet.getRange(rowNumber, sourceColumnIndexes[ci]).getValue();
        values.push(val === null ? '' : String(val));
      }
    } else {
      // 找不到這個標籤對應的列，回傳全空陣列，避免整支程式因為單一列缺漏而掛掉
      for (var ce = 0; ce < sourceColumnIndexes.length; ce++) values.push('');
    }
    rows.push({ label: outputLabel, values: values });
  }

  // ---- 第四步：六個小組區塊（每組固定 4 列：組名/任務/負責人/位置）----
  var groups = [];
  var groupTaskRowNumbers = []; // 記錄每組「任務」列的列號，等一下要拿來判斷休息底色
  for (var gi = 0; gi < GROUP_NAMES.length; gi++) {
    var groupName = GROUP_NAMES[gi];
    var groupSheetLabel = groupName + GROUP_NAME_SHEET_SUFFIX; // 例如 "總控A" + "組" = "總控A組"
    var groupHeaderRow = rowNumberByLabel[groupSheetLabel];
    var task = [], owner = [], location = [];
    var taskRowNumber = groupHeaderRow ? groupHeaderRow + GROUP_BLOCK_OFFSET_TASK : null;
    groupTaskRowNumbers.push(taskRowNumber);

    if (groupHeaderRow) {
      var ownerRowNumber = groupHeaderRow + GROUP_BLOCK_OFFSET_OWNER;
      var locationRowNumber = groupHeaderRow + GROUP_BLOCK_OFFSET_LOCATION;
      for (var gc = 0; gc < sourceColumnIndexes.length; gc++) {
        var colIndex = sourceColumnIndexes[gc];
        task.push(String(sheet.getRange(taskRowNumber, colIndex).getValue() || ''));
        owner.push(String(sheet.getRange(ownerRowNumber, colIndex).getValue() || ''));
        location.push(String(sheet.getRange(locationRowNumber, colIndex).getValue() || ''));
      }
    } else {
      for (var gce = 0; gce < sourceColumnIndexes.length; gce++) {
        task.push(''); owner.push(''); location.push('');
      }
    }
    groups.push({ name: groupName, task: task, owner: owner, location: location });
  }

  // ---- 第五步：判斷每個時段是否為「休息」 ----
  // 判斷依據（雙重保險）：
  //   (a) 任一小組「任務」列的儲存格背景色 === REST_BACKGROUND_COLOR
  //   (b) 任一小組「任務」列的文字內容剛好就是「休息」二字（防止之後改版拿掉底色但留文字）
  var restColumns = [];
  for (var rc = 0; rc < sourceColumnIndexes.length; rc++) {
    var colIndex = sourceColumnIndexes[rc];
    var isResting = false;
    for (var tr = 0; tr < groupTaskRowNumbers.length; tr++) {
      var taskRow = groupTaskRowNumbers[tr];
      if (!taskRow) continue;
      var cell = sheet.getRange(taskRow, colIndex);
      var bg = String(cell.getBackground()).toLowerCase();
      var text = String(cell.getValue() || '').trim();
      if (bg === REST_BACKGROUND_COLOR || text === '休息') {
        isResting = true;
        break;
      }
    }
    restColumns.push(isResting);
  }

  return {
    exerciseDate: config.exerciseDate,
    columns: columns,
    rows: rows,
    groups: groups,
    restColumns: restColumns
  };
}


/**
 * ============================================================================
 * handleGetLog — 讀取「處置記錄」分頁所有列
 * ============================================================================
 * 分頁欄位順序固定：ID | 時間戳記 | 事件分類 | 狀態 | 內容備註（第1列為標題，不讀取）
 * 排序交給前端處理，這裡照表格原始順序回傳即可。
 */
function handleGetLog() {
  var controlSheet = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = controlSheet.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    throw new Error('找不到「' + LOG_SHEET_NAME + '」分頁，請檢查看板控制中心 Sheet 的分頁名稱設定');
  }
  var lastRow = logSheet.getLastRow();
  var logs = [];
  if (lastRow >= 2) {
    var values = logSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (!row[0] && !row[1]) continue; // 跳過完全空白的列
      logs.push({
        id: String(row[0]),
        timestamp: formatTimestampForOutput(row[1]),
        category: String(row[2] || ''),
        status: String(row[3] || ''),
        note: String(row[4] || '')
      });
    }
  }
  return { logs: logs };
}

/**
 * 處置記錄的時間戳記欄位可能是 Apps Script 自動轉換的 Date 物件，也可能是使用者手動輸入的文字，
 * 統一轉成含台灣時區 +08:00 的 ISO 8601 字串，跟前端 API contract 保持一致。
 */
function formatTimestampForOutput(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ssXXX");
  }
  return String(value || '');
}


/**
 * ============================================================================
 * handleAddLog — 新增一筆處置記錄到「處置記錄」分頁最後一列
 * ============================================================================
 */
function handleAddLog(payload) {
  var controlSheet = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = controlSheet.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    throw new Error('找不到「' + LOG_SHEET_NAME + '」分頁，請檢查看板控制中心 Sheet 的分頁名稱設定');
  }
  var id = Utilities.getUuid();
  logSheet.appendRow([
    id,
    payload.timestamp || '',
    payload.category || '',
    payload.status || '',
    payload.note || ''
  ]);
  return { success: true, id: id };
}


/**
 * ============================================================================
 * handleSetConfig — 更新「設定」分頁的場控時序檔案網址／演習日期
 * ============================================================================
 * 用途：場控組如果要切換到另一場演習的場控時序檔案，不用重新部署 Apps Script，
 * 直接透過前端（未來若加上對應設定介面）呼叫這個 action 更新設定分頁即可。
 */
function handleSetConfig(payload) {
  var controlSheet = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = controlSheet.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    throw new Error('找不到「' + SETTINGS_SHEET_NAME + '」分頁，請檢查看板控制中心 Sheet 的分頁名稱設定');
  }
  if (payload.spreadsheetUrl) {
    upsertConfigValue(settingsSheet, '場控時序檔案網址或ID', payload.spreadsheetUrl);
  }
  if (payload.exerciseDate) {
    upsertConfigValue(settingsSheet, '演習開始日期', payload.exerciseDate);
  }
  return { success: true };
}

/**
 * 在「設定」分頁裡，找到 A 欄等於 label 的列就更新 B 欄；找不到就在最後新增一列。
 */
function upsertConfigValue(settingsSheet, label, value) {
  var lastRow = settingsSheet.getLastRow();
  if (lastRow >= 1) {
    var values = settingsSheet.getRange(1, 1, lastRow, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === label) {
        settingsSheet.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
  }
  settingsSheet.appendRow([label, value]);
}
