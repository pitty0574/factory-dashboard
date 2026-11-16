/* ══════════════════════════════════════════════════════════════════
   C 棟點檢電子表單 — 設定檔
   頁面（c-inspection.html）完全依這份設定產生頁籤與欄位，
   新增點檢項目只要在 ITEMS 加一個物件，不需要改 HTML。

   ┌────────────────────────────────────────────────────────────────┐
   │ 新增點檢項目範本（複製到 ITEMS 陣列後填寫）                     │
   └────────────────────────────────────────────────────────────────┘
   {
     id: '',              // 英文代號，小寫加底線，全系統唯一；網址 ?item=<id> 直接開啟
     nameZh: '',          // 頁籤 / 標題中文名稱
     nameEn: '',          // 頁籤 / 標題英文名稱
     tab: '',             // 寫入的 Google Sheet 分頁名稱（第一次提交時由 Apps Script 建立，第 1 列為表頭）
     source: 'script',    // 'script' = 由本頁面填寫（會出現在頁籤）
                          // 'form'   = 由 Google 表單填寫，本頁只登記、不產生表單
     dailyCount: 1,       // 每日應填次數；同一日期已達此筆數時，提交前會跳確認
     deadlineParams: [],  // 截止時間參數名稱（對應系統設定分頁），依序對應第 1、2… 次
     fields: [            // 依序即為寫入 Sheet 的欄位順序（headers 取 nameZh）
       {
         key: '',         // 欄位代號（同一項目內唯一）
         nameZh: '',      // 中文標籤（也是 Sheet 表頭）
         nameEn: '',      // 英文標籤
         type: '',        // 'date'         日期（寫入 YYYY-MM-DD，文字格式）
                          // 'weekday'      星期，由 from 指定的 date 欄自動計算，不可編輯
                          // 'time'         時間（寫入 HH:mm，文字格式）
                          // 'number'       數字，可搭配 inputmode、specParam
                          // 'select'       下拉選單，需 options
                          // 'toggle'       大按鈕單選，需 options
                          // 'employee'     4 位數工號（文字格式），localStorage 記住上次輸入
                          // 'employeeName' 姓名，由 from 指定的 employee 欄查員工表自動帶出
                          // 'result'       判定，依 specParam 欄位自動填 OK / NG / 空字串
                          // 'text'         一般文字
         required: false, // 必填
         default: '',     // 預設值：date 可用 'today'、time 可用 'now'，其他填固定值
         from: '',        // weekday / employeeName 的來源欄位 key
         options: [],     // select / toggle 選項：字串，或 { value: '寫入值', en: '英文顯示' }
         inputmode: '',   // number 欄位的手機鍵盤：'decimal'（可輸入小數點）或 'numeric'
         specParam: '',   // 規格參數前綴，例如 'xxx' → 讀取系統設定的 xxx_min / xxx_max
         requiredWhenNG: false  // 判定為 NG 時改為必填（用於備註）
       }
     ]
   }

   【系統設定分頁（SETTINGS_TAB）要新增的參數】A 欄參數名稱、B 欄參數值，第 1 列是表頭
     · 有 specParam: 'xxx' 的欄位 → 新增 xxx_min、xxx_max 兩列
         值留空 = 該側不判定；兩側都留空 = 該欄位不判定
     · deadlineParams 裡的每個名稱 → 新增一列，值填時間（例：12:00）
     · 若要在 Sheet 未設定時也有預設值，請一併加到下方 DEFAULTS
   ══════════════════════════════════════════════════════════════════ */

window.C_INSPECTION = (function () {
  const SHEET_ID = '1PMGwahs2541WqFGqaLMBezJxOX7oA1NW_RFvGyCkwBo';
  // 與射出日報共用的 Sheets API Key（唯讀）
  const API_KEY = 'AIzaSyBMgXHsua64mciOFJwqH4NreGI11gMXR-s';
  // 補漆房共用 Sheet（讀員工名單用）：員工表 A 工號 / B 姓名 / C 崗位 / D 狀態
  const WORKER_SHEET_ID = '1xQoIxoLlM4LkKjDrYZ_87uIahtebPgPcG6Hitck1T5c';
  const WORKER_RANGE = '員工表!A2:D1000';

  const DEFAULTS = {
    primer_ohm_min: 3,
    primer_ohm_max: 10,
    primer_deadline: '12:00',
    patrol_deadline_1: '11:00',
    patrol_deadline_2: '15:30'
  };

  const ITEMS = [
    {
      id: 'primer_resistance',
      nameZh: '底漆阻抗',
      nameEn: 'Primer Resistance',
      tab: '底漆阻抗',
      source: 'script',
      dailyCount: 1,
      deadlineParams: ['primer_deadline'],
      fields: [
        { key: 'date',         nameZh: '日期',     nameEn: 'Date',              type: 'date', default: 'today', required: true },
        { key: 'weekday',      nameZh: '星期',     nameEn: 'Day',               type: 'weekday', from: 'date' },
        { key: 'coatingTime',  nameZh: '塗裝時間', nameEn: 'Coating Time',      type: 'time' },
        { key: 'measureTime',  nameZh: '量測時間', nameEn: 'Measure Time',      type: 'time', default: 'now' },
        { key: 'panel1',       nameZh: '試片1',    nameEn: 'Panel 1',           type: 'number', inputmode: 'decimal', required: true, specParam: 'primer_ohm' },
        { key: 'panel2',       nameZh: '試片2',    nameEn: 'Panel 2',           type: 'number', inputmode: 'decimal', required: true, specParam: 'primer_ohm' },
        { key: 'supplier',     nameZh: '廠商',     nameEn: 'Supplier',          type: 'select', options: ['華山', '健裕', { value: '其他', en: 'Other' }] },
        { key: 'filterBag',    nameZh: '濾袋更換', nameEn: 'Filter Bag Changed', type: 'toggle', required: true,
          options: [{ value: '有', en: 'Yes' }, { value: '無', en: 'No' }] },
        { key: 'employeeId',   nameZh: '工號',     nameEn: 'Employee ID',       type: 'employee', required: true },
        { key: 'employeeName', nameZh: '姓名',     nameEn: 'Name',              type: 'employeeName', from: 'employeeId' },
        { key: 'result',       nameZh: '判定',     nameEn: 'Result',            type: 'result' },
        { key: 'remark',       nameZh: '備註',     nameEn: 'Remark',            type: 'text', requiredWhenNG: true }
      ]
    },
    {
      // 由 Google 表單填寫，這裡只登記（給之後的看板 / 逾期提醒使用）
      id: 'booth_patrol',
      nameZh: '噴房巡檢',
      nameEn: 'Booth Patrol',
      tab: '第 1 张表单回复',
      source: 'form',
      dailyCount: 6,
      deadlineParams: ['patrol_deadline_1', 'patrol_deadline_2'],
      areas: ['底漆靜置室', '面漆噴房', '面漆靜置室'],
      fields: []
    }
  ];

  // ── Google Sheets 讀取 ──────────────────────────────────────────
  // 分頁名稱含空白、數字開頭或「-」時，A1 表示法必須加單引號
  function q(sheet) { return `'${String(sheet).replace(/'/g, "''")}'`; }

  async function readRange(range, sheetId) {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId || SHEET_ID}/values/${encodeURIComponent(range)}?key=${API_KEY}`;
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) {
      const errJson = await res.json().catch(() => null);
      const err = new Error((errJson && errJson.error && errJson.error.message) || `HTTP ${res.status}`);
      err.status = res.status;
      throw err;
    }
    const json = await res.json();
    return json.values || [];
  }
  // 讀整個分頁；分頁不存在時回傳 null（不丟錯）
  async function readTab(tab) {
    try {
      return await readRange(q(tab) + '!A:Z');
    } catch (e) {
      if (isMissingTab(e)) return null;
      throw e;
    }
  }
  function isMissingTab(e) {
    return e && e.status === 400 && /Unable to parse range|not found/i.test(e.message);
  }

  async function loadWorkers() {
    const rows = await readRange(WORKER_RANGE, WORKER_SHEET_ID);
    const map = {};
    rows.forEach(r => {
      const [id, name, role, status] = r;
      if (!id || !String(id).trim()) return;
      map[String(id).trim()] = { name: (name || '').trim(), role: (role || '').trim(), active: (status || '').trim() !== '離職' };
    });
    return map;
  }

  // ── 值的容錯解析 ──
  function pad(n) { return String(n).padStart(2, '0'); }
  // '12:00'、'12:00:00'、'下午 12:00:00'、'3:30 PM'、'1899/12/30 下午 3:30:00' → 'HH:mm'；無法解析回傳 null
  function toHHmm(v) {
    const s = String(v == null ? '' : v).trim();
    const m = s.match(/(上午|下午|AM|PM)?\s*(\d{1,2})[:：](\d{2})(?:[:：]\d{2})?\s*(上午|下午|AM|PM)?/i);
    if (!m) return null;
    let h = parseInt(m[2], 10);
    const min = parseInt(m[3], 10);
    const ap = (m[1] || m[4] || '').toUpperCase();
    if ((ap === '下午' || ap === 'PM') && h < 12) h += 12;
    if ((ap === '上午' || ap === 'AM') && h === 12) h = 0;
    if (h > 23 || min > 59) return null;
    return `${pad(h)}:${pad(min)}`;
  }
  // '2026-11-16'、'2026/11/16'、'2026/11/16 上午 08:00:00'、'11/16/2026' → 'YYYY-MM-DD'；無法解析回傳 ''
  function toYMD(v) {
    const s = String(v == null ? '' : v).trim().split(/[\sT]/)[0].replace(/-/g, '/');
    const m1 = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
    if (m1) return `${m1[1]}-${pad(m1[2])}-${pad(m1[3])}`;
    const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m2) return `${m2[3]}-${pad(m2[1])}-${pad(m2[2])}`;
    return '';
  }
  function parseSettingValue(raw, fallback) {
    const s = String(raw == null ? '' : raw).trim();
    if (s === '') return null;                                   // 空字串 = 不判定
    const n = s.replace(/,/g, '');
    if (/^-?\d*\.?\d+$/.test(n)) return parseFloat(n);
    const looksTime = /[:：]/.test(s) || (typeof fallback === 'string' && /^\d{1,2}:\d{2}$/.test(fallback));
    if (looksTime) {
      const t = toHHmm(s);
      if (t) return t;
      console.warn('[c-inspection] 時間參數無法解析，使用預設值', raw);
      return fallback === undefined ? s : fallback;
    }
    return s;
  }

  // ── 系統設定 ──
  // 回傳 { 參數名稱: 參數值 }：Sheet 有列就以 Sheet 為準（值空白 = null），沒有列才用 DEFAULTS
  async function loadSettings() {
    const settings = Object.assign({}, DEFAULTS);
    try {
      const rows = await readRange(q(api.SETTINGS_TAB) + '!A2:B');
      rows.forEach(r => {
        const name = String(r[0] == null ? '' : r[0]).trim();
        if (!name) return;
        settings[name] = parseSettingValue(r[1], DEFAULTS[name]);
      });
    } catch (e) {
      console.warn('[c-inspection] 系統設定讀取失敗，全部使用預設值', e);
    }
    api.settings = settings;
    return settings;
  }

  // ── 規格 / 判定 ──
  function num(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
    return isFinite(n) ? n : null;
  }
  // 欄位規格 { min, max }；某一側為 null 就不檢查該側。沒有 specParam 回傳 null
  function getSpec(field, settings) {
    if (!field || !field.specParam) return null;
    const s = settings || api.settings || DEFAULTS;
    return { min: num(s[field.specParam + '_min']), max: num(s[field.specParam + '_max']) };
  }
  // 單一值是否超規（空值或無規格時回傳 false）
  function isOutOfSpec(value, spec) {
    const n = num(value);
    if (n === null || !spec) return false;
    return (spec.min !== null && n < spec.min) || (spec.max !== null && n > spec.max);
  }
  // values = { 欄位 key: 值 } → 'NG' / 'OK' / ''（所有 specParam 欄位上下限皆為 null）
  function judge(item, values, settings) {
    let hasLimit = false, ng = false;
    item.fields.forEach(f => {
      const spec = getSpec(f, settings);
      if (!spec || (spec.min === null && spec.max === null)) return;
      hasLimit = true;
      if (isOutOfSpec(values[f.key], spec)) ng = true;
    });
    if (!hasLimit) return '';
    return ng ? 'NG' : 'OK';
  }
  // 規格顯示文字：'3～10'、'≥ 3'、'≤ 10'；兩側皆 null 回傳 ''
  function specText(spec) {
    if (!spec) return '';
    if (spec.min !== null && spec.max !== null) return `${spec.min}～${spec.max}`;
    if (spec.min !== null) return `≥ ${spec.min}`;
    if (spec.max !== null) return `≤ ${spec.max}`;
    return '';
  }

  const api = {
    SHEET_ID,
    SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbxQ4wxhk1liwfRsP0b9FaFoJwp76Jog0D4zDx7fX8dpUwpFz_MW7Z2RPxcqoifMmtwm/exec',
    ADMIN_TAB: '管理員設定',
    SETTINGS_TAB: '系統設定',
    API_KEY,
    WORKER_SHEET_ID,
    DEFAULTS,
    ITEMS,
    settings: null,
    loadSettings,
    loadWorkers,
    readRange,
    readTab,
    isMissingTab,
    toHHmm,
    toYMD,
    getSpec,
    isOutOfSpec,
    judge,
    specText
  };
  return api;
})();
