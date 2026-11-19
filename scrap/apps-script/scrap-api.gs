/**
 * 部件報廢申請單 API（Google Apps Script，container-bound）
 *
 * ── 部署方式 ──────────────────────────────────────────────
 * 1. 建立一個新的 Google Sheet → 擴充功能 → Apps Script，將本檔內容整份貼上並儲存。
 * 2. 首次使用：在編輯器上方函式選單選擇 setupSheets，按「執行」一次並完成授權。
 *    會建立「報廢明細」「下拉選單資料」「料號主檔」三個分頁（已存在則跳過，不覆蓋資料）。
 * 3. 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *      執行身分：我
 *      存取權：所有人
 *    部署後取得 /exec 網址，填入前端設定。
 * 4. 之後修改程式碼：部署 → 管理部署作業 → 編輯 → 版本選「新版本」，網址不變。
 *
 * ── API ───────────────────────────────────────────────────
 * GET  ?action=options
 * GET  ?action=list&from=yyyy-MM-dd&to=yyyy-MM-dd[&unit=xxx]
 * POST body（JSON 字串，Content-Type: text/plain）：
 *      {action:'add', date, unit, shift, car, machine, partNo, name, qty,
 *       defectType, reason, respUnit, operator, note}
 *      {action:'delete', id}
 *      {action:'markPrinted', date, unit}
 * 錯誤一律回傳 {ok:false, error:訊息}。
 *
 * 新增報廢時，若料號主檔沒有相同的 車型+件號+車色，會自動在最後補一列（來源＝現場新增）。
 */

var SHEET_RECORDS = '報廢明細';
var SHEET_OPTIONS = '下拉選單資料';
var SHEET_PARTS = '料號主檔';

var STATUS_UNPRINTED = '未列印';
var STATUS_PRINTED = '已列印';

// 報廢明細欄位：表頭 ↔ 英文 key（順序即欄位順序）
var RECORD_FIELDS = [
  ['ID', 'id'],
  ['建立時間', 'createdAt'],
  ['日期', 'date'],
  ['單位', 'unit'],
  ['班別', 'shift'],
  ['車型', 'car'],
  ['機台', 'machine'],
  ['件號', 'partNo'],
  ['車色', 'name'],
  ['報廢數量', 'qty'],
  ['不良類別', 'defectType'],
  ['不良原因', 'reason'],
  ['責任單位', 'respUnit'],
  ['填寫人', 'operator'],
  ['狀態', 'status'],
  ['列印時間', 'printedAt'],
  ['備註', 'note']
];

// 以文字格式儲存的欄位（避免日期時區偏移、件號被轉成數字）
var TEXT_KEYS = ['id', 'date', 'car', 'machine', 'partNo', 'name'];

var OPTION_COLUMNS = [
  ['單位', ['塗裝']],
  ['班別', ['日班', '夜班']],
  ['不良類別', ['異常', '換模', '開機']],
  ['不良原因', ['嚴重撞傷', '素材凹', '孔洞偏移', '掛架上掉落', '打磨破', '擦傷', '碰凸', '其他']],
  ['責任單位', ['塗裝', '射出', '前處理', '素材廠商', '其他']]
];

var PART_HEADERS = ['車型', '件號', '車色', '來源', '新增日期'];
var PART_SOURCE_FIELD = '現場新增';
var PART_DEFAULTS = [
  ['BH', '85022', 'GKT/GBL'],
  ['BH', '85022', '235'],
  ['ZS', '62022', 'CH'],
  ['ZS', '85022', 'JQ'],
  ['ZS', '65022', 'JQ'],
  ['CS', '75059', '素材'],
  ['PU', '85022', '素材'],
  ['JJ', 'H5920', 'PZ'],
  ['JS', '85022', 'KS']
];

var DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

/* ============================================================
 * 初始化（只需手動執行一次）
 * ============================================================ */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // (a) 報廢明細
  if (!ss.getSheetByName(SHEET_RECORDS)) {
    var sh = ss.insertSheet(SHEET_RECORDS);
    var headers = RECORD_FIELDS.map(function (f) { return f[0]; });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    TEXT_KEYS.forEach(function (key) {
      sh.getRange(1, colOf_(key), sh.getMaxRows(), 1).setNumberFormat('@');
    });
    sh.setFrozenRows(1);
  }

  // (b) 下拉選單資料
  if (!ss.getSheetByName(SHEET_OPTIONS)) {
    var so = ss.insertSheet(SHEET_OPTIONS);
    var maxLen = 0;
    OPTION_COLUMNS.forEach(function (c) { maxLen = Math.max(maxLen, c[1].length); });
    var grid = [];
    for (var r = 0; r <= maxLen; r++) {
      grid.push(OPTION_COLUMNS.map(function (c) {
        return r === 0 ? c[0] : (c[1][r - 1] || '');
      }));
    }
    so.getRange(1, 1, grid.length, OPTION_COLUMNS.length).setNumberFormat('@').setValues(grid);
    so.getRange(1, 1, 1, OPTION_COLUMNS.length).setFontWeight('bold');
    so.setFrozenRows(1);
  }

  // (c) 料號主檔
  var sp = ss.getSheetByName(SHEET_PARTS);
  if (!sp) {
    sp = ss.insertSheet(SHEET_PARTS);
    sp.getRange(1, 1, sp.getMaxRows(), PART_HEADERS.length).setNumberFormat('@');
    sp.getRange(1, 1, 1, PART_HEADERS.length).setValues([PART_HEADERS]).setFontWeight('bold');
    sp.getRange(2, 1, PART_DEFAULTS.length, PART_DEFAULTS[0].length).setValues(PART_DEFAULTS);
    sp.setFrozenRows(1);
  } else {
    // 既有分頁：只補 D1「來源」、E1「新增日期」（已有內容則不覆蓋），不動資料
    for (var c = 4; c <= PART_HEADERS.length; c++) {
      var cell = sp.getRange(1, c);
      if (String(cell.getDisplayValue()).trim() === '') {
        cell.setValue(PART_HEADERS[c - 1]).setFontWeight('bold');
      }
    }
  }
}

/* ============================================================
 * GET
 * ============================================================ */
function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    switch (p.action) {
      case 'options':
        return json_(getOptions_());
      case 'list':
        return json_(listRecords_(p.from, p.to, p.unit));
      default:
        return json_({ ok: false, error: '未知的 action：' + (p.action || '(空白)') });
    }
  } catch (err) {
    return json_({ ok: false, error: errMsg_(err) });
  }
}

function getOptions_() {
  var so = mustSheet_(SHEET_OPTIONS);
  var lastRow = so.getLastRow();
  var cols = [[], [], [], [], []];
  if (lastRow > 1) {
    var values = so.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
    values.forEach(function (row) {
      for (var c = 0; c < 5; c++) {
        var v = String(row[c]).trim();
        if (v) cols[c].push(v);
      }
    });
  }

  var sp = mustSheet_(SHEET_PARTS);
  var parts = [];
  var pLast = sp.getLastRow();
  if (pLast > 1) {
    sp.getRange(2, 1, pLast - 1, 3).getDisplayValues().forEach(function (row) {
      var car = String(row[0]).trim();
      var partNo = String(row[1]).trim();
      var name = String(row[2]).trim();
      if (car || partNo || name) parts.push({ car: car, partNo: partNo, name: name });
    });
  }

  return {
    ok: true,
    units: cols[0],
    shifts: cols[1],
    defectTypes: cols[2],
    reasons: cols[3],
    respUnits: cols[4],
    parts: parts
  };
}

function listRecords_(from, to, unit) {
  if (!DATE_RE.test(from || '') || !DATE_RE.test(to || '')) {
    return { ok: false, error: 'from / to 需為 yyyy-MM-dd' };
  }
  unit = unit ? String(unit).trim() : '';
  var rows = readRecords_().filter(function (r) {
    return r.date >= from && r.date <= to && (!unit || r.unit === unit);
  });
  return { ok: true, rows: rows };
}

/* ============================================================
 * POST
 * ============================================================ */
function doPost(e) {
  try {
    var body;
    try {
      body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    } catch (parseErr) {
      return json_({ ok: false, error: 'body 不是有效的 JSON' });
    }

    var handler;
    switch (body.action) {
      case 'add': handler = addRecord_; break;
      case 'delete': handler = deleteRecord_; break;
      case 'markPrinted': handler = markPrinted_; break;
      default:
        return json_({ ok: false, error: '未知的 action：' + (body.action || '(空白)') });
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) {
      return json_({ ok: false, error: '系統忙碌中，請稍後再試' });
    }
    try {
      return json_(handler(body));
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: errMsg_(err) });
  }
}

function addRecord_(b) {
  var rec = {};
  ['date', 'unit', 'shift', 'car', 'machine', 'partNo', 'name',
    'defectType', 'reason', 'respUnit', 'operator', 'note'].forEach(function (k) {
    rec[k] = b[k] == null ? '' : String(b[k]).trim();
  });

  if (!DATE_RE.test(rec.date)) return { ok: false, error: '日期格式需為 yyyy-MM-dd' };
  var required = [['unit', '單位'], ['shift', '班別'], ['car', '車型'], ['partNo', '件號'],
    ['name', '車色'], ['defectType', '不良類別'], ['reason', '不良原因'],
    ['respUnit', '責任單位'], ['operator', '填寫人']];
  for (var i = 0; i < required.length; i++) {
    if (!rec[required[i][0]]) return { ok: false, error: '缺少欄位：' + required[i][1] };
  }
  var qty = Number(b.qty);
  if (!isFinite(qty) || qty <= 0 || Math.floor(qty) !== qty) {
    return { ok: false, error: '報廢數量需為正整數' };
  }

  rec.id = Utilities.getUuid();
  rec.createdAt = nowStr_();
  rec.qty = qty;
  rec.status = STATUS_UNPRINTED;
  rec.printedAt = '';

  var sh = mustSheet_(SHEET_RECORDS);
  var rowIdx = sh.getLastRow() + 1;
  var range = sh.getRange(rowIdx, 1, 1, RECORD_FIELDS.length);
  range.setNumberFormats([RECORD_FIELDS.map(function (f) {
    return TEXT_KEYS.indexOf(f[1]) >= 0 ? '@' : 'General';
  })]);
  range.setValues([RECORD_FIELDS.map(function (f) { return rec[f[1]]; })]);

  // 料號主檔回寫：失敗只記 log，不影響報廢明細的寫入結果（呼叫端已持有 ScriptLock）
  try {
    appendPartIfMissing_(rec.car, rec.partNo, rec.name);
  } catch (partErr) {
    console.log('料號主檔回寫失敗：' + errMsg_(partErr) + '｜' + rec.car + ' / ' + rec.partNo + ' / ' + rec.name);
  }

  return { ok: true, id: rec.id };
}

// 料號主檔沒有相同 車型+件號+車色 時在最後新增一列（車型、件號以大寫比對，三欄皆 trim）
function appendPartIfMissing_(car, partNo, name) {
  car = String(car || '').trim();
  partNo = String(partNo || '').trim();
  name = String(name || '').trim();
  if (!car || !partNo || !name) return;

  var sp = mustSheet_(SHEET_PARTS);
  var lastRow = sp.getLastRow();
  if (lastRow > 1) {
    var keyCar = car.toUpperCase(), keyPart = partNo.toUpperCase();
    var rows = sp.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toUpperCase() === keyCar &&
          String(rows[i][1]).trim().toUpperCase() === keyPart &&
          String(rows[i][2]).trim() === name) {
        return;
      }
    }
  }

  var today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  var target = sp.getRange(Math.max(lastRow, 1) + 1, 1, 1, PART_HEADERS.length);
  target.setNumberFormat('@');
  target.setValues([[car, partNo, name, PART_SOURCE_FIELD, today]]);
}

function deleteRecord_(b) {
  var id = String(b.id || '').trim();
  if (!id) return { ok: false, error: '缺少 id' };

  var sh = mustSheet_(SHEET_RECORDS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, error: '找不到此筆資料' };

  var idCol = colOf_('id');
  var statusCol = colOf_('status');
  var ids = sh.getRange(2, idCol, lastRow - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) {
      var row = i + 2;
      var status = String(sh.getRange(row, statusCol).getDisplayValue()).trim();
      if (status !== STATUS_UNPRINTED) return { ok: false, error: '已列印，無法刪除' };
      sh.deleteRow(row);
      return { ok: true };
    }
  }
  return { ok: false, error: '找不到此筆資料' };
}

function markPrinted_(b) {
  var date = String(b.date || '').trim();
  var unit = String(b.unit || '').trim();
  if (!DATE_RE.test(date)) return { ok: false, error: '日期格式需為 yyyy-MM-dd' };
  if (!unit) return { ok: false, error: '缺少單位' };

  var sh = mustSheet_(SHEET_RECORDS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, updated: 0 };

  var n = lastRow - 1;
  var tz = tz_();
  var dateVals = sh.getRange(2, colOf_('date'), n, 1).getValues();
  var unitVals = sh.getRange(2, colOf_('unit'), n, 1).getDisplayValues();
  var statusRange = sh.getRange(2, colOf_('status'), n, 1);
  var printedRange = sh.getRange(2, colOf_('printedAt'), n, 1);
  var statusVals = statusRange.getValues();
  var printedVals = printedRange.getValues();

  var now = nowStr_();
  var updated = 0;
  for (var i = 0; i < n; i++) {
    if (toDateStr_(dateVals[i][0], tz) !== date) continue;
    if (String(unitVals[i][0]).trim() !== unit) continue;
    if (String(statusVals[i][0]).trim() !== STATUS_PRINTED) updated++;
    statusVals[i][0] = STATUS_PRINTED;
    if (printedVals[i][0] === '' || printedVals[i][0] == null) printedVals[i][0] = now;
  }
  statusRange.setValues(statusVals);
  printedRange.setValues(printedVals);

  return { ok: true, updated: updated };
}

/* ============================================================
 * 共用工具
 * ============================================================ */
function readRecords_() {
  var sh = mustSheet_(SHEET_RECORDS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var tz = tz_();
  var values = sh.getRange(2, 1, lastRow - 1, RECORD_FIELDS.length).getValues();
  var out = [];
  values.forEach(function (row) {
    if (!String(row[0]).trim()) return; // 無 ID 的空列略過
    var o = {};
    RECORD_FIELDS.forEach(function (f, c) {
      var v = row[c];
      switch (f[1]) {
        case 'date': o.date = toDateStr_(v, tz); break;
        case 'createdAt':
        case 'printedAt': o[f[1]] = toDateTimeStr_(v, tz); break;
        case 'qty': o.qty = v === '' ? 0 : Number(v); break;
        default: o[f[1]] = v == null ? '' : String(v).trim();
      }
    });
    out.push(o);
  });
  return out;
}

function colOf_(key) {
  for (var i = 0; i < RECORD_FIELDS.length; i++) {
    if (RECORD_FIELDS[i][1] === key) return i + 1;
  }
  throw new Error('未知欄位：' + key);
}

function mustSheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('找不到分頁「' + name + '」，請先執行 setupSheets()');
  return sh;
}

function tz_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Asia/Taipei';
}

function nowStr_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss');
}

// 日期欄應為文字；若被手動改成日期值，仍以試算表時區轉回 yyyy-MM-dd
function toDateStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return v == null ? '' : String(v).trim();
}

function toDateTimeStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss');
  return v == null ? '' : String(v).trim();
}

function errMsg_(err) {
  return (err && err.message) ? err.message : String(err);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
