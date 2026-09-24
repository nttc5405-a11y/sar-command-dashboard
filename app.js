/*
 * 特搜場控即時看板 — 前端邏輯 (app.js)
 * ------------------------------------------------------------
 * 這支檔案負責：
 *   1. 從「Apps Script 後端」或「本地 mock-data.json」抓資料
 *   2. 把資料畫到畫面上（時段總覽卡片、六組任務卡片、處置記錄列表）
 *   3. 處理使用者互動（切換時段、新增記錄、設定面板）
 *
 * 重要原則：不管資料是從 Apps Script 來、還是從 mock-data.json 來，
 * 一律先整理成同一種「內部資料格式」，再交給同一套 render 函式畫面，
 * 避免真實模式跟示範模式要各寫一套畫面邏輯（容易兩邊長得不一樣）。
 */

(function () {
  'use strict';

  // ===== 常數設定 =====
  var STORAGE_KEY_URL = 'ctrlboard_apps_script_url';
  var STORAGE_KEY_MODE = 'ctrlboard_mode'; // 'mock' 或 'live'
  var MOCK_DATA_URL = 'mock-data.json';
  var REFRESH_INTERVAL_MS = 17000; // 每 17 秒自動重新抓一次資料（15~20 秒區間內）
  var CATEGORY_OPTIONS = ['通報', '派遣', '到場', '處理中', '後送', '完成', '其他'];
  var STATUS_OPTIONS = ['待命', '出動', '處理中', '完成'];

  // ===== 全域狀態 =====
  var state = {
    mode: 'mock', // 目前模式：mock（示範資料）或 live（真實 Apps Script）
    appsScriptUrl: '',
    timeline: null, // 目前已載入的時段資料（格式見 README/程式內註解）
    logs: [], // 目前已載入的處置記錄陣列
    groupCardsBuilt: false, // 六組卡片的骨架是否已建立（只需建立一次）
    currentColumnIndex: 0, // 系統時間對應到的「真正目前時段」索引
    viewingColumnIndex: 0, // 使用者目前正在檢視的時段索引（可能不等於 currentColumnIndex）
    sortOrder: 'desc', // 處置記錄排序：'desc' = 最新在前，'asc' = 最舊在前
    lastLogsSignature: '', // 用來判斷 logs 有沒有變動，變動才重繪列表，減少閃爍
    refreshTimer: null,
    mockDataCache: null // 快取讀過的 mock-data.json 內容，避免每次都重新 fetch
  };

  // ===== 小工具函式 =====

  // 把數字補成兩位數字串，例如 9 -> "09"
  function pad2(n) {
    return n < 10 ? '0' + n : String(n);
  }

  // 把 Date 物件格式化成「MM/DD HH:mm」，用在下拉選單跟提示文字
  function formatDateShort(d) {
    return pad2(d.getMonth() + 1) + '/' + pad2(d.getDate()) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
  }

  // 依 exerciseDate（演習第一個時段的日期，對應 columns[0]，固定為當天 09:00）
  // 加上「索引 * 1 小時」，算出每個時段欄位實際對應的日期時間。
  // 依照需求規格：不要去解析 columns 裡的文字時間字串猜日期，只用索引位移計算，
  // 這樣就算欄位文字剛好跨過午夜（例如第二天的 00:00、01:00），也不會算錯日期。
  function getColumnDateTime(exerciseDate, columnIndex) {
    var base = new Date(exerciseDate + 'T09:00:00+08:00');
    return new Date(base.getTime() + columnIndex * 60 * 60 * 1000);
  }

  // 找出「目前時段」的索引：從第一欄開始往後找，找到最後一個「時間 <= 現在」的欄位
  function computeCurrentColumnIndex(timeline) {
    if (!timeline || !timeline.columns || timeline.columns.length === 0) return 0;
    var now = new Date();
    var idx = 0;
    for (var i = 0; i < timeline.columns.length; i++) {
      var colTime = getColumnDateTime(timeline.exerciseDate, i);
      if (colTime.getTime() <= now.getTime()) {
        idx = i;
      } else {
        break;
      }
    }
    return idx;
  }

  // 用 JSON 字串比較兩份 log 陣列是否相同（簡單但夠用的「有沒有變動」判斷法）
  function computeLogsSignature(logs) {
    try {
      return JSON.stringify(logs);
    } catch (e) {
      return String(logs.length);
    }
  }

  // ===== localStorage 存取 =====

  function loadSettingsFromStorage() {
    try {
      state.appsScriptUrl = localStorage.getItem(STORAGE_KEY_URL) || '';
      var savedMode = localStorage.getItem(STORAGE_KEY_MODE);
      // 沒有 Apps Script 網址、或使用者沒存過模式時，預設用示範資料，
      // 這樣使用者第一次直接雙擊打開 index.html 就看得到畫面。
      if (savedMode === 'live' && state.appsScriptUrl) {
        state.mode = 'live';
      } else {
        state.mode = 'mock';
      }
    } catch (e) {
      // 某些環境（例如直接用 file:// 開啟）可能會擋 localStorage，失敗就用預設值
      state.appsScriptUrl = '';
      state.mode = 'mock';
    }
  }

  function saveUrlToStorage(url) {
    try {
      localStorage.setItem(STORAGE_KEY_URL, url);
    } catch (e) {
      /* 忽略儲存失敗（例如無痕模式），畫面仍可正常運作，只是重新整理後要重填 */
    }
  }

  function saveModeToStorage(mode) {
    try {
      localStorage.setItem(STORAGE_KEY_MODE, mode);
    } catch (e) {
      /* 同上，忽略 */
    }
  }

  // ===== 資料存取層：mock 模式 / 真實模式共用同一組函式介面 =====

  // 讀取 mock-data.json（只讀一次，之後用快取），回傳 Promise
  function loadMockData() {
    if (state.mockDataCache) {
      return Promise.resolve(state.mockDataCache);
    }
    return fetch(MOCK_DATA_URL)
      .then(function (res) {
        if (!res.ok) throw new Error('讀取 mock-data.json 失敗：HTTP ' + res.status);
        return res.json();
      })
      .then(function (data) {
        state.mockDataCache = data;
        return data;
      });
  }

  // 取得「時段資料」(timeline)：對應 API contract 的 GET ?action=timeline
  function fetchTimeline() {
    if (state.mode === 'mock') {
      return loadMockData().then(function (data) {
        return {
          exerciseDate: data.exerciseDate,
          columns: data.columns,
          rows: data.rows,
          groups: data.groups,
          restColumns: data.restColumns
        };
      });
    }
    var url = state.appsScriptUrl + '?action=timeline';
    return fetch(url).then(function (res) {
      if (!res.ok) throw new Error('讀取時段資料失敗：HTTP ' + res.status);
      return res.json();
    });
  }

  // 取得「處置記錄」(log)：對應 API contract 的 GET ?action=log
  function fetchLog() {
    if (state.mode === 'mock') {
      return loadMockData().then(function (data) {
        return { logs: data.logs || [] };
      });
    }
    var url = state.appsScriptUrl + '?action=log';
    return fetch(url).then(function (res) {
      if (!res.ok) throw new Error('讀取處置記錄失敗：HTTP ' + res.status);
      return res.json();
    });
  }

  // 新增一筆處置記錄。
  // mock 模式下沒有真正的後端可以寫入，所以只在瀏覽器記憶體中暫時新增一筆，
  // 重新整理頁面後就會消失——這是示範模式的已知限制，已寫在 README 裡。
  function postAddLog(payload) {
    if (state.mode === 'mock') {
      return loadMockData().then(function (data) {
        var newLog = {
          id: 'local-' + Date.now(),
          timestamp: payload.timestamp,
          category: payload.category,
          status: payload.status,
          note: payload.note
        };
        data.logs = data.logs || [];
        data.logs.push(newLog);
        return { success: true, id: newLog.id };
      });
    }
    var body = Object.assign({ action: 'addLog' }, payload);
    return fetch(state.appsScriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    }).then(function (res) {
      if (!res.ok) throw new Error('新增記錄失敗：HTTP ' + res.status);
      return res.json();
    });
  }

  // 切換演習檔案設定（設定分頁用的 setConfig action）。
  // 目前設定面板 UI 沒有對應輸入欄位（規格書的設定面板只要求網址跟模式切換），
  // 這個函式先寫好、符合 API contract，未來若要加「切換演習場次」的畫面可以直接呼叫它。
  function postSetConfig(spreadsheetUrl, exerciseDate) {
    if (state.mode === 'mock') {
      return Promise.resolve({ success: true });
    }
    var body = { action: 'setConfig', spreadsheetUrl: spreadsheetUrl, exerciseDate: exerciseDate };
    return fetch(state.appsScriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    }).then(function (res) {
      if (!res.ok) throw new Error('更新設定失敗：HTTP ' + res.status);
      return res.json();
    });
  }

  // ===== 畫面渲染：時鐘、頂部時段選單 =====

  function renderClock() {
    var el = document.getElementById('nowClock');
    if (!el) return;
    var now = new Date();
    el.textContent = now.getFullYear() + '/' + pad2(now.getMonth() + 1) + '/' + pad2(now.getDate()) +
      ' ' + pad2(now.getHours()) + ':' + pad2(now.getMinutes()) + ':' + pad2(now.getSeconds());
  }

  // 重建時段下拉選單的選項（時段資料剛載入、或時段數量改變時才需要重建）
  function renderSlotSelect() {
    var select = document.getElementById('slotSelect');
    if (!select || !state.timeline) return;
    var columns = state.timeline.columns;
    var html = '';
    for (var i = 0; i < columns.length; i++) {
      var dt = getColumnDateTime(state.timeline.exerciseDate, i);
      var label = formatDateShort(dt) + (state.restColumnsSafe()[i] ? '（休息）' : '');
      html += '<option value="' + i + '">' + label + '</option>';
    }
    select.innerHTML = html;
    select.value = String(state.viewingColumnIndex);
  }

  // 更新「非目前時段」提示橫幅的顯示/隱藏
  function renderNonCurrentBanner() {
    var banner = document.getElementById('nonCurrentBanner');
    var text = document.getElementById('nonCurrentText');
    if (!banner || !text || !state.timeline) return;
    if (state.viewingColumnIndex === state.currentColumnIndex) {
      banner.classList.add('hidden');
      return;
    }
    var dt = getColumnDateTime(state.timeline.exerciseDate, state.viewingColumnIndex);
    text.textContent = '正在查看：' + formatDateShort(dt) + '（非目前時段）';
    banner.classList.remove('hidden');
  }

  // ===== 畫面渲染：目前時段總覽卡片 =====

  // rows 陣列裡的 label 跟畫面欄位 id 的對照表（同一套設定，前後端都要對齊 API contract）
  var ROW_FIELD_MAP = {
    '階段': 'slotPhase',
    '區域': 'slotArea',
    '傷患編號': 'slotVictim',
    'INJECT': 'slotInject',
    'LEMA會議': 'slotLema',
    'OCC會議': 'slotOcc',
    '拍照攝影': 'slotPhoto',
    '救護車司機': 'slotAmbulance',
    '後送醫院': 'slotHospital'
  };

  function safeValue(v) {
    return v === undefined || v === null || v === '' ? '—' : v;
  }

  function renderCurrentSlotCard() {
    if (!state.timeline) return;
    var idx = state.viewingColumnIndex;
    var card = document.getElementById('currentSlotCard');
    var restBadge = document.getElementById('restBadge');
    var isResting = !!state.restColumnsSafe()[idx];

    if (card) {
      card.classList.toggle('is-resting', isResting);
    }
    if (restBadge) {
      restBadge.classList.toggle('hidden', !isResting);
    }

    (state.timeline.rows || []).forEach(function (row) {
      var fieldId = ROW_FIELD_MAP[row.label];
      if (!fieldId) return; // 遇到 API contract 沒定義的列標籤，先忽略，不讓畫面壞掉
      var el = document.getElementById(fieldId);
      if (el) {
        el.textContent = safeValue(row.values[idx]);
      }
    });
  }

  // ===== 畫面渲染：六組任務卡片 =====

  function buildGroupCardsSkeleton() {
    var container = document.getElementById('groupCards');
    if (!container || !state.timeline || !state.timeline.groups) return;
    var html = '';
    state.timeline.groups.forEach(function (group, i) {
      html += '' +
        '<div class="card group-card" id="group-card-' + i + '">' +
        '  <h3 class="group-card-name">' + group.name + '</h3>' +
        '  <dl class="group-card-body">' +
        '    <dt>任務</dt><dd id="group-' + i + '-task">—</dd>' +
        '    <dt>負責人</dt><dd id="group-' + i + '-owner">—</dd>' +
        '    <dt>位置</dt><dd id="group-' + i + '-location">—</dd>' +
        '  </dl>' +
        '</div>';
    });
    container.innerHTML = html;
    state.groupCardsBuilt = true;
  }

  function renderGroupCards() {
    if (!state.timeline || !state.timeline.groups) return;
    if (!state.groupCardsBuilt) {
      buildGroupCardsSkeleton();
    }
    var idx = state.viewingColumnIndex;
    state.timeline.groups.forEach(function (group, i) {
      var taskEl = document.getElementById('group-' + i + '-task');
      var ownerEl = document.getElementById('group-' + i + '-owner');
      var locEl = document.getElementById('group-' + i + '-location');
      if (taskEl) taskEl.textContent = safeValue(group.task[idx]);
      if (ownerEl) ownerEl.textContent = safeValue(group.owner[idx]);
      if (locEl) locEl.textContent = safeValue(group.location[idx]);
    });
  }

  // ===== 畫面渲染：處置記錄列表 =====

  function sortLogs(logs) {
    var copy = logs.slice();
    copy.sort(function (a, b) {
      var ta = new Date(a.timestamp).getTime();
      var tb = new Date(b.timestamp).getTime();
      return state.sortOrder === 'desc' ? tb - ta : ta - tb;
    });
    return copy;
  }

  function formatLogTimestamp(iso) {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return iso;
    return pad2(d.getMonth() + 1) + '/' + pad2(d.getDate()) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
  }

  function renderLogList() {
    var listEl = document.getElementById('logList');
    if (!listEl) return;
    var signature = computeLogsSignature(state.logs) + '|' + state.sortOrder;
    if (signature === state.lastLogsSignature) {
      return; // 資料跟排序都沒變，不重繪，避免不必要的畫面閃動
    }
    state.lastLogsSignature = signature;

    var sorted = sortLogs(state.logs);
    if (sorted.length === 0) {
      listEl.innerHTML = '<li class="log-empty">目前尚無處置記錄</li>';
      return;
    }
    var html = sorted.map(function (log) {
      return '' +
        '<li class="log-item">' +
        '  <span class="log-time">' + formatLogTimestamp(log.timestamp) + '</span>' +
        '  <span class="log-category">' + escapeHtml(log.category) + '</span>' +
        '  <span class="log-status log-status-' + escapeHtml(log.status) + '">' + escapeHtml(log.status) + '</span>' +
        '  <span class="log-note">' + escapeHtml(log.note || '') + '</span>' +
        '</li>';
    }).join('');
    listEl.innerHTML = html;
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ===== 設定面板 =====

  function renderSettingsPanel() {
    var urlInput = document.getElementById('appsScriptUrlInput');
    var modeMock = document.getElementById('modeMock');
    var modeLive = document.getElementById('modeLive');
    var statusEl = document.getElementById('connectionStatus');
    if (urlInput) urlInput.value = state.appsScriptUrl;
    if (modeMock) modeMock.checked = state.mode === 'mock';
    if (modeLive) modeLive.checked = state.mode === 'live';
    if (statusEl) {
      if (state.mode === 'mock') {
        statusEl.textContent = '目前狀態：示範模式（讀取本地 mock-data.json）';
        statusEl.className = 'connection-status status-mock';
      } else {
        statusEl.textContent = state.appsScriptUrl
          ? '目前狀態：已連接真實資料（' + state.appsScriptUrl + '）'
          : '目前狀態：真實模式，但尚未設定 Apps Script 網址';
        statusEl.className = 'connection-status status-live';
      }
    }
  }

  function openSettingsPanel() {
    document.getElementById('settingsPanel').classList.remove('hidden');
    document.getElementById('settingsOverlay').classList.remove('hidden');
    renderSettingsPanel();
  }

  function closeSettingsPanel() {
    document.getElementById('settingsPanel').classList.add('hidden');
    document.getElementById('settingsOverlay').classList.add('hidden');
  }

  // ===== 主要流程：載入資料、渲染畫面 =====

  // state.restColumnsSafe()：restColumns 若缺漏（例如舊資料或後端漏回傳），一律回傳全 false 陣列，避免畫面壞掉
  state.restColumnsSafe = function () {
    if (state.timeline && Array.isArray(state.timeline.restColumns)) {
      return state.timeline.restColumns;
    }
    return (state.timeline && state.timeline.columns ? state.timeline.columns : []).map(function () {
      return false;
    });
  };

  function refreshAll(isFirstLoad) {
    return Promise.all([fetchTimeline(), fetchLog()])
      .then(function (results) {
        var timeline = results[0];
        var logResult = results[1];
        state.timeline = timeline;
        state.logs = (logResult && logResult.logs) || [];

        state.currentColumnIndex = computeCurrentColumnIndex(timeline);
        if (isFirstLoad) {
          // 第一次載入：預設顯示「目前時段」
          state.viewingColumnIndex = state.currentColumnIndex;
          state.groupCardsBuilt = false; // 資料結構可能第一次才確定，重建一次卡片骨架
        }
        // 防呆：如果使用者正在看的時段超出新資料的範圍，拉回最後一欄
        if (state.viewingColumnIndex >= timeline.columns.length) {
          state.viewingColumnIndex = timeline.columns.length - 1;
        }

        renderSlotSelect();
        renderNonCurrentBanner();
        renderCurrentSlotCard();
        renderGroupCards();
        renderLogList();
      })
      .catch(function (err) {
        console.error('[特搜場控看板] 資料載入失敗：', err);
        showLoadError(err);
      });
  }

  function showLoadError(err) {
    var statusEl = document.getElementById('connectionStatus');
    if (statusEl) {
      statusEl.textContent = '資料載入失敗：' + err.message + '（請確認設定面板中的 Apps Script 網址是否正確、或切換回示範模式）';
      statusEl.className = 'connection-status status-error';
    }
  }

  function startAutoRefresh() {
    if (state.refreshTimer) {
      clearInterval(state.refreshTimer);
    }
    state.refreshTimer = setInterval(function () {
      refreshAll(false);
    }, REFRESH_INTERVAL_MS);
  }

  // ===== 事件綁定 =====

  function bindEvents() {
    // 時鐘每秒更新
    setInterval(renderClock, 1000);
    renderClock();

    // 上一時段 / 下一時段
    document.getElementById('prevSlotBtn').addEventListener('click', function () {
      if (!state.timeline) return;
      state.viewingColumnIndex = Math.max(0, state.viewingColumnIndex - 1);
      onViewingIndexChanged();
    });
    document.getElementById('nextSlotBtn').addEventListener('click', function () {
      if (!state.timeline) return;
      state.viewingColumnIndex = Math.min(state.timeline.columns.length - 1, state.viewingColumnIndex + 1);
      onViewingIndexChanged();
    });
    // 下拉選單直接選時段
    document.getElementById('slotSelect').addEventListener('change', function (e) {
      state.viewingColumnIndex = parseInt(e.target.value, 10) || 0;
      onViewingIndexChanged();
    });
    // 回到目前時段
    document.getElementById('backToCurrentBtn').addEventListener('click', function () {
      state.viewingColumnIndex = state.currentColumnIndex;
      onViewingIndexChanged();
    });

    // 設定面板開關
    document.getElementById('settingsBtn').addEventListener('click', openSettingsPanel);
    document.getElementById('closeSettingsBtn').addEventListener('click', closeSettingsPanel);
    document.getElementById('settingsOverlay').addEventListener('click', closeSettingsPanel);

    // 儲存 Apps Script 網址
    document.getElementById('saveUrlBtn').addEventListener('click', function () {
      var input = document.getElementById('appsScriptUrlInput');
      var url = (input.value || '').trim();
      state.appsScriptUrl = url;
      saveUrlToStorage(url);
      renderSettingsPanel();
    });

    // 套用新演習場次設定（切換到新的場控時序檔案）
    document.getElementById('applyConfigBtn').addEventListener('click', function () {
      var urlInput = document.getElementById('newSpreadsheetUrlInput');
      var dateInput = document.getElementById('newExerciseDateInput');
      var configStatusEl = document.getElementById('configStatus');
      var spreadsheetUrl = (urlInput.value || '').trim();
      var exerciseDate = dateInput.value;

      if (!spreadsheetUrl || !exerciseDate) {
        configStatusEl.textContent = '請把「Google Sheets 網址」跟「演習開始日期」都填寫完整';
        configStatusEl.className = 'connection-status status-error';
        return;
      }
      if (state.mode === 'mock') {
        configStatusEl.textContent = '目前是示範模式，請先切換到「使用真實資料」再套用';
        configStatusEl.className = 'connection-status status-error';
        return;
      }
      if (!state.appsScriptUrl) {
        configStatusEl.textContent = '請先填寫並儲存 Apps Script 網址';
        configStatusEl.className = 'connection-status status-error';
        return;
      }

      configStatusEl.textContent = '套用中...';
      configStatusEl.className = 'connection-status';
      postSetConfig(spreadsheetUrl, exerciseDate).then(function (result) {
        if (result && result.success) {
          configStatusEl.textContent = '已套用新場次設定，正在重新載入資料...';
          configStatusEl.className = 'connection-status status-live';
          urlInput.value = '';
          dateInput.value = '';
          state.lastLogsSignature = '';
          state.groupCardsBuilt = false; // 換了新檔案，欄位結構可能不同，強制重建卡片骨架
          refreshAll(true);
        } else {
          configStatusEl.textContent = '套用失敗，請確認網址跟日期是否正確';
          configStatusEl.className = 'connection-status status-error';
        }
      }).catch(function (err) {
        configStatusEl.textContent = '套用失敗：' + err.message;
        configStatusEl.className = 'connection-status status-error';
      });
    });

    // 切換 mock / live 模式
    document.getElementById('modeMock').addEventListener('change', function () {
      switchMode('mock');
    });
    document.getElementById('modeLive').addEventListener('change', function () {
      switchMode('live');
    });

    // 排序切換
    document.getElementById('sortToggleBtn').addEventListener('click', function () {
      state.sortOrder = state.sortOrder === 'desc' ? 'asc' : 'desc';
      document.getElementById('sortToggleBtn').textContent = state.sortOrder === 'desc' ? '最新在前 ▾' : '最舊在前 ▴';
      state.lastLogsSignature = ''; // 強制重繪一次
      renderLogList();
    });

    // 新增處置記錄表單
    document.getElementById('logForm').addEventListener('submit', function (e) {
      e.preventDefault();
      var timestampInput = document.getElementById('logTimestamp').value; // "YYYY-MM-DDTHH:mm"
      var category = document.getElementById('logCategory').value;
      var status = document.getElementById('logStatus').value;
      var note = document.getElementById('logNote').value.trim();

      if (!timestampInput) {
        alert('請填寫時間戳記');
        return;
      }
      // datetime-local 輸入沒有時區資訊，這裡當作台灣時間 +08:00 補上
      var isoTimestamp = timestampInput.length === 16 ? timestampInput + ':00+08:00' : timestampInput + '+08:00';

      var submitBtn = e.target.querySelector('button[type="submit"]');
      if (submitBtn) submitBtn.disabled = true;

      postAddLog({ timestamp: isoTimestamp, category: category, status: status, note: note })
        .then(function () {
          document.getElementById('logNote').value = '';
          setDefaultLogTimestamp();
          return refreshAll(false);
        })
        .catch(function (err) {
          alert('新增記錄失敗：' + err.message);
        })
        .then(function () {
          if (submitBtn) submitBtn.disabled = false;
        });
    });
  }

  function onViewingIndexChanged() {
    renderSlotSelect();
    renderNonCurrentBanner();
    renderCurrentSlotCard();
    renderGroupCards();
  }

  function switchMode(mode) {
    state.mode = mode;
    saveModeToStorage(mode);
    state.lastLogsSignature = ''; // 換模式後強制重繪列表
    renderSettingsPanel();
    refreshAll(true);
  }

  // 表單「時間戳記」欄位預設帶入現在時間（可手動修改，供事後補登用）
  function setDefaultLogTimestamp() {
    var input = document.getElementById('logTimestamp');
    if (!input) return;
    var now = new Date();
    var value = now.getFullYear() + '-' + pad2(now.getMonth() + 1) + '-' + pad2(now.getDate()) +
      'T' + pad2(now.getHours()) + ':' + pad2(now.getMinutes());
    input.value = value;
  }

  // ===== 啟動 =====

  function init() {
    loadSettingsFromStorage();
    bindEvents();
    setDefaultLogTimestamp();
    renderSettingsPanel();
    refreshAll(true).then(function () {
      startAutoRefresh();
    });
  }

  document.addEventListener('DOMContentLoaded', init);
})();
