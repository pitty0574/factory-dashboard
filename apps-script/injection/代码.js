// ============================================================
// A 棟射出生產 - Apps Script 完整版
// ============================================================

// === 1. 建立所有分頁（一次性，之前執行過就不用再跑）===
// ⚠️ 會清空所有分頁資料，執行前會跳確認視窗
function buildSheets() {
  const ui = SpreadsheetApp.getUi();
  const ans = ui.alert(
    '⚠️ 重建分頁',
    '會清空 5 台機台分頁、射出機狀態、下拉選單資料的「所有資料」，確定要執行？',
    ui.ButtonSet.YES_NO
  );
  if (ans !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  const machines = ['500T', '750T', '2500T', '3000T-1', '3000T-2'];
  const machineHeaders = ['時間戳記', '記錄日期', '工號', '車型', '件號', '料號', '模具編號', '本次模次', '產出數量', '班別', '狀態', '備註'];
  const headerBg = '#1c3a6e';
  const headerFg = '#ffffff';

  machines.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clear();
    sheet.appendRow(machineHeaders);
    sheet.getRange(1, 1, 1, machineHeaders.length)
      .setFontWeight('bold').setBackground(headerBg).setFontColor(headerFg)
      .setFontSize(11).setHorizontalAlignment('center');
    sheet.getRange('A:A').setNumberFormat('yyyy/MM/dd HH:mm:ss');
    sheet.getRange('B:B').setNumberFormat('yyyy-MM-dd');
    sheet.getRange('C:C').setNumberFormat('@');
    sheet.getRange('H:I').setNumberFormat('0');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, machineHeaders.length);
  });

  let statusSheet = ss.getSheetByName('射出機狀態');
  if (!statusSheet) statusSheet = ss.insertSheet('射出機狀態');
  statusSheet.clear();
  statusSheet.appendRow(['機台', '狀態', '更新時間', '更新者', '備註', '模式']);
  statusSheet.getRange(1, 1, 1, 6)
    .setFontWeight('bold').setBackground(headerBg).setFontColor(headerFg)
    .setFontSize(11).setHorizontalAlignment('center');
  const now = new Date();
  machines.forEach(m => statusSheet.appendRow([m, '待機', now, '系統', '初始化', '自動']));
  statusSheet.getRange('C:C').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['生產中', '試模中', '待機'], true).build();
  statusSheet.getRange('B2:B100').setDataValidation(statusRule);
  const modeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['手動', '自動'], true).build();
  statusSheet.getRange('F2:F100').setDataValidation(modeRule);
  statusSheet.setFrozenRows(1);
  statusSheet.autoResizeColumns(1, 6);

  let ddSheet = ss.getSheetByName('下拉選單資料');
  if (!ddSheet) ddSheet = ss.insertSheet('下拉選單資料');
  ddSheet.clear();
  ddSheet.appendRow(['車型', '件號', '料號']);
  ddSheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold').setBackground(headerBg).setFontColor(headerFg)
    .setFontSize(11).setHorizontalAlignment('center');
  const models = ['RV', 'PU', 'JJ', 'JD', 'ZS', 'AQ', 'TF', 'BH', 'JS', 'JL', 'JE', 'JH', 'CS'];
  const parts = ['前保', '後保', '側裙', '前下巴', '後下巴', '其他'];
  models.forEach((m, i) => ddSheet.getRange(i + 2, 1).setValue(m));
  parts.forEach((p, i) => ddSheet.getRange(i + 2, 2).setValue(p));
  ddSheet.setFrozenRows(1);
  ddSheet.autoResizeColumns(1, 3);

  ['工作表1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  const order = ['500T', '750T', '2500T', '3000T-1', '3000T-2', '射出機狀態', '下拉選單資料'];
  order.forEach((name, i) => {
    const s = ss.getSheetByName(name);
    if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); }
  });

  ui.alert('✅ 建立完成！\n\n7 個分頁：' + order.join(' / '));
}


// === 1b. 批次功能欄位（一次性，可重複執行；已存在的分頁／欄位會跳過，不清資料）===
const MACHINES = ['500T', '750T', '2500T', '3000T-1', '3000T-2'];
const BATCH_SHEET = '批次紀錄';
const BATCH_HEADERS = ['批次ID', '機台', '模穴數', '件號1', '件號2', '料號', '模具編號', '批次數量', '開始時間', '開始人工號', '結束時間', '狀態'];
const MACHINE_EXTRA_HEADERS = ['批次ID', '模穴數'];
const BATCH_COL = BATCH_HEADERS.reduce((o, h, i) => (o[h] = i, o), {});  // 欄名 → 0-based index

function setupBatchSchema() {
  const ss = SpreadsheetApp.getActive();
  const headerBg = '#1c3a6e';
  const headerFg = '#ffffff';
  const log = [];
  const styleHeader = (sheet, col, n) => sheet.getRange(1, col, 1, n)
    .setFontWeight('bold').setBackground(headerBg).setFontColor(headerFg)
    .setFontSize(11).setHorizontalAlignment('center');

  // 「批次紀錄」分頁
  let bs = ss.getSheetByName(BATCH_SHEET);
  if (!bs) {
    bs = ss.insertSheet(BATCH_SHEET);
    bs.getRange(1, 1, 1, BATCH_HEADERS.length).setValues([BATCH_HEADERS]);
    styleHeader(bs, 1, BATCH_HEADERS.length);
    // 批次ID / 機台 / 件號 / 料號 / 模具 / 工號存成純文字，避免被轉成數字或日期
    ['A:B', 'D:G', 'J:J'].forEach(a => bs.getRange(a).setNumberFormat('@'));
    bs.getRange('I:I').setNumberFormat('yyyy/MM/dd HH:mm:ss');
    bs.getRange('K:K').setNumberFormat('yyyy/MM/dd HH:mm:ss');
    bs.getRange('C:C').setNumberFormat('0');
    bs.getRange('H:H').setNumberFormat('0');
    bs.getRange('L2:L').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['進行中', '完成', '中止'], true).build());
    bs.setFrozenRows(1);
    bs.autoResizeColumns(1, BATCH_HEADERS.length);
    log.push('新增分頁「' + BATCH_SHEET + '」');
  } else {
    const added = appendMissingHeaders_(bs, BATCH_HEADERS, styleHeader);
    log.push('「' + BATCH_SHEET + '」已存在' + (added.length ? '，補上欄位：' + added.join('、') : '，跳過'));
  }

  // 5 個機台分頁：在最後面補「批次ID | 模穴數」，既有欄位不動
  MACHINES.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { log.push('⚠️ 找不到分頁 ' + name); return; }
    const added = appendMissingHeaders_(sh, MACHINE_EXTRA_HEADERS, styleHeader);
    if (added.length) {
      const map = headerMap_(sh);
      sh.getRange(1, map['批次ID'] + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    }
    log.push(name + '：' + (added.length ? '補上 ' + added.join('、') : '欄位已存在，跳過'));
  });

  Logger.log(log.join('\n'));
  try { SpreadsheetApp.getUi().alert('✅ setupBatchSchema 完成\n\n' + log.join('\n')); } catch (e) {}
}

// 標題列缺少的欄位依序加在最後一欄之後，回傳實際新增的欄名
function appendMissingHeaders_(sheet, headers, styleHeader) {
  const lastCol = sheet.getLastColumn();
  const existing = lastCol ? sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h).trim()) : [];
  const missing = headers.filter(h => existing.indexOf(h) < 0);
  if (!missing.length) return [];
  if (sheet.getMaxColumns() < lastCol + missing.length) {
    sheet.insertColumnsAfter(Math.max(lastCol, 1), lastCol + missing.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  styleHeader(sheet, lastCol + 1, missing.length);
  return missing;
}

// 標題名稱 → 0-based 欄位 index
function headerMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  const out = {};
  if (!lastCol) return out;
  sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].forEach((h, i) => {
    h = String(h).trim();
    if (h && !(h in out)) out[h] = i;
  });
  return out;
}

function fmtTime_(d) {
  return d instanceof Date ? Utilities.formatDate(d, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss') : String(d || '');
}

function batchRowToObj_(r) {
  const g = h => r[BATCH_COL[h]];
  return {
    batchId:   String(g('批次ID')),
    machine:   String(g('機台')),
    cavity:    Number(g('模穴數')) === 2 ? 2 : 1,
    partNo:    String(g('件號1')),
    partNo2:   String(g('件號2') || ''),
    materialNo: String(g('料號') || ''),
    moldNo:    String(g('模具編號') || ''),
    batchQty:  Number(g('批次數量')) || 0,
    startTime: fmtTime_(g('開始時間')),
    startPid:  String(g('開始人工號') || ''),
    endTime:   fmtTime_(g('結束時間')),
    status:    String(g('狀態'))
  };
}

function getBatchSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(BATCH_SHEET);
  if (!sh) throw new Error('找不到「' + BATCH_SHEET + '」分頁，請先執行 setupBatchSchema');
  return sh;
}

// 讀「批次紀錄」全部資料列（不含標題），回傳 { sheet, rows }
function readBatchRows_() {
  const sh = getBatchSheet_();
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, BATCH_HEADERS.length).getValues() : [];
  return { sheet: sh, rows: rows };
}

const isPosInt_ = v => {
  const n = Number(v);
  return v !== '' && v !== null && v !== undefined && Number.isInteger(n) && n > 0;
};


// === 2. 表單 API（給 injection-report.html 與 3D 戰情室用）===

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);

    // 3D 戰情室：管理員改機台狀態（沒有 pid，必須放在基本驗證之前）
    if (params.action === 'setMachineStatus') return json(setMachineStatus_(params));
    if (params.action === 'startBatch') return json(startBatch_(params));
    if (params.action === 'endBatch') return json(endBatch_(params));

    // 支援 Claude Code 自訂命名
    const machine    = params.machine;
    const pid        = params.pid;
    const model      = params.model      || '';
    const partNo     = params.partNo     || params.part || '';
    const partNo2    = String(params.partNo2 || '').trim();
    const materialNo = params.materialNo || params.partNumber || '';
    const moldNo     = params.moldNo     || params.moldNumber || '';
    const shots      = params.shots      || params.cycleCount || 1;
    const qtyRaw     = params.qty != null ? params.qty : params.totalCount;
    const shift      = params.shift      || '';
    const status     = params.status     || '生產中';
    const note       = params.note       || params.remark || '';
    const batchId    = String(params.batchId || '').trim();
    // 舊版前端沒帶 cavity → 留白；新版帶 1 或 2
    const cavityRaw  = params.cavity;
    const cavity     = Number(cavityRaw) === 2 ? 2 : 1;

    // 基本驗證
    if (!machine || !pid) return json({ ok: false, error: '缺少機台或工號' });
    if (!MACHINES.includes(machine)) return json({ ok: false, error: '無效機台: ' + machine });
    if (!isPosInt_(qtyRaw)) return json({ ok: false, error: '產出數量必須是正整數：' + qtyRaw });
    const qty = Number(qtyRaw);
    if (cavity === 2) {
      if (!partNo2) return json({ ok: false, error: '一模兩穴缺少件號2' });
      if (partNo2 === String(partNo).trim()) return json({ ok: false, error: '件號2 不能與件號1 相同' });
    }

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(machine);
    if (!sheet) return json({ ok: false, error: '找不到分頁: ' + machine });

    // 寫入機台分頁（A~L 固定 12 欄；批次ID／模穴數依標題名稱找欄，沒跑過 setupBatchSchema 就不寫）
    const now = new Date();
    const today = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');
    const hmap = headerMap_(sheet);
    const cBatch = hmap['批次ID'], cCavity = hmap['模穴數'];
    const width = Math.max(12, (cBatch == null ? 0 : cBatch + 1), (cCavity == null ? 0 : cCavity + 1));
    const makeRow = (part, h) => {
      const row = [
        now,                          // A 時間戳記
        today,                        // B 記錄日期
        String(pid).trim(),           // C 工號
        model,                        // D 車型
        part,                         // E 件號
        materialNo,                   // F 料號
        moldNo,                       // G 模具編號
        h,                            // H 本次模次
        qty,                          // I 產出數量（兩穴時＝模次數，兩列各記一次）
        shift,                        // J 班別
        status,                       // K 狀態
        note                          // L 備註
      ];
      while (row.length < width) row.push('');
      if (cBatch != null) row[cBatch] = batchId;
      if (cCavity != null) row[cCavity] = cavityRaw == null || cavityRaw === '' ? '' : cavity;
      return row;
    };
    // 一模兩穴：件號1 那列 H=1、件號2 那列 H=0，加總 H 欄＝實際模次數
    const rows = cavity === 2
      ? [makeRow(partNo, 1), makeRow(partNo2, 0)]
      : [makeRow(partNo, Number(shots) || 1)];

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, width).setValues(rows);
    } finally {
      lock.releaseLock();
    }

    // 同步更新「射出機狀態」
    // 手動模式時不覆蓋 B 欄狀態（保留管理員設定），C/D 照常更新
    const statusSheet = ss.getSheetByName('射出機狀態');
    if (statusSheet) {
      const col = statusSheet.getRange('A:A').getValues();
      for (let i = 0; i < col.length; i++) {
        if (col[i][0] === machine) {
          const row = i + 1;
          const mode = String(statusSheet.getRange(row, 6).getValue()).trim();
          if (mode !== '手動') statusSheet.getRange(row, 2).setValue(status);
          statusSheet.getRange(row, 3).setValue(now);
          statusSheet.getRange(row, 4).setValue(String(pid).trim());
          break;
        }
      }
    }

    return json({
      ok: true,
      machine,
      pid,
      batchId,
      rows: rows.length,
      time: Utilities.formatDate(now, 'Asia/Taipei', 'HH:mm:ss')
    });
  } catch (err) {
    return json({ ok: false, error: err.toString() });
  }
}


// === 3. 3D 戰情室：管理員設定機台狀態 ===
// 寫入「射出機狀態」：B 狀態、C 更新時間、D 更新者、E 備註、F 模式
function setMachineStatus_(params) {
  const STATUS_OK = ['生產中', '試模中', '待機'];
  const MODE_OK = ['手動', '自動'];

  const machine = String(params.machine || params['機台'] || '').trim();
  const status  = String(params.status || '').trim();
  const mode    = String(params.mode || '').trim();
  const user    = String(params.user || '').trim() || '管理員';
  const note    = String(params.note || '').trim() || '手動設定';

  if (!machine) return { ok: false, message: '缺少 machine' };
  if (STATUS_OK.indexOf(status) < 0) return { ok: false, message: 'status 不合法：' + status };
  if (MODE_OK.indexOf(mode) < 0) return { ok: false, message: 'mode 不合法：' + mode };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('射出機狀態');
    if (!sh) return { ok: false, message: '找不到「射出機狀態」分頁' };

    const names = sh.getRange(1, 1, sh.getLastRow(), 1).getDisplayValues();
    let row = -1;
    for (let i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim() === machine) { row = i + 1; break; }
    }
    if (row < 0) return { ok: false, message: '找不到機台：' + machine };

    const now = new Date();
    sh.getRange(row, 2, 1, 5).setValues([[status, now, user, note, mode]]); // B~F
    return {
      ok: true,
      machine: machine,
      status: status,
      mode: mode,
      time: Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss')
    };
  } finally {
    lock.releaseLock();
  }
}


// === 4. 批次：開始 / 結束（給 injection-report.html 幹部區塊用）===
// 開始新批次：同機台舊的「進行中」先改成「中止」，再寫一列「進行中」
function startBatch_(params) {
  const machine    = String(params.machine || '').trim();
  const cavity     = Number(params.cavity) === 2 ? 2 : 1;
  const partNo     = String(params.partNo || '').trim();
  const partNo2    = cavity === 2 ? String(params.partNo2 || '').trim() : '';
  const materialNo = String(params.materialNo || '').trim();
  const moldNo     = String(params.moldNo || '').trim();
  const pid        = String(params.pid || '').trim();
  const batchQty   = params.batchQty;

  if (!MACHINES.includes(machine)) return { ok: false, error: '無效機台: ' + machine };
  if (!partNo) return { ok: false, error: '缺少件號1' };
  if (cavity === 2 && !partNo2) return { ok: false, error: '一模兩穴缺少件號2' };
  if (cavity === 2 && partNo2 === partNo) return { ok: false, error: '件號2 不能與件號1 相同' };
  if (!materialNo) return { ok: false, error: '缺少料號' };
  if (!isPosInt_(batchQty)) return { ok: false, error: '批次數量必須是正整數：' + batchQty };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const { sheet, rows } = readBatchRows_();
    const now = new Date();
    let batchId = machine + '-' + Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMdd-HHmmss');
    // 同一秒重按 → 加尾碼避免 ID 重複
    const ids = new Set(rows.map(r => String(r[BATCH_COL['批次ID']])));
    for (let k = 2; ids.has(batchId); k++) batchId = batchId.replace(/(_\d+)?$/, '_' + k);
    const aborted = [];
    rows.forEach((r, i) => {
      if (String(r[BATCH_COL['機台']]) === machine && String(r[BATCH_COL['狀態']]) === '進行中') {
        sheet.getRange(i + 2, BATCH_COL['結束時間'] + 1).setValue(now);
        sheet.getRange(i + 2, BATCH_COL['狀態'] + 1).setValue('中止');
        aborted.push(String(r[BATCH_COL['批次ID']]));
      }
    });
    const row = [batchId, machine, cavity, partNo, partNo2, materialNo, moldNo, Number(batchQty), now, pid, '', '進行中'];
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    return { ok: true, batchId: batchId, batch: batchRowToObj_(row), aborted: aborted };
  } finally {
    lock.releaseLock();
  }
}

// 結束批次：指定批次 → 完成 + 結束時間（只接受「進行中」的批次）
function endBatch_(params) {
  const batchId = String(params.batchId || '').trim();
  if (!batchId) return { ok: false, error: '缺少 batchId' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const { sheet, rows } = readBatchRows_();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][BATCH_COL['批次ID']]) !== batchId) continue;
      const st = String(rows[i][BATCH_COL['狀態']]);
      if (st !== '進行中') return { ok: false, error: '批次 ' + batchId + ' 狀態為「' + st + '」，無法結束' };
      const now = new Date();
      sheet.getRange(i + 2, BATCH_COL['結束時間'] + 1).setValue(now);
      sheet.getRange(i + 2, BATCH_COL['狀態'] + 1).setValue('完成');
      return { ok: true, batchId: batchId, endTime: fmtTime_(now) };
    }
    return { ok: false, error: '找不到批次 ' + batchId };
  } finally {
    lock.releaseLock();
  }
}

// === 5. 批次：查詢 ===
// 該機台「進行中」的批次（有多筆時取最後一筆），沒有回傳 null
function findActiveBatch_(machine) {
  const { rows } = readBatchRows_();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][BATCH_COL['機台']]) === machine && String(rows[i][BATCH_COL['狀態']]) === '進行中') {
      return batchRowToObj_(rows[i]);
    }
  }
  return null;
}

// 機台分頁中同一批次ID 的產出數量，依件號加總 → { 件號: qty }
function batchProgress_(machine, batchId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(machine);
  if (!sh) throw new Error('找不到分頁: ' + machine);
  const hmap = headerMap_(sh);
  const cB = hmap['批次ID'], cP = hmap['件號'], cQ = hmap['產出數量'];
  const out = {};
  if (cB == null || cP == null || cQ == null) return out;
  const n = sh.getLastRow() - 1;
  if (n <= 0) return out;
  const width = Math.max(cB, cP, cQ) + 1;
  sh.getRange(2, 1, n, width).getValues().forEach(r => {
    if (String(r[cB]).trim() !== batchId) return;
    const part = String(r[cP]).trim();
    out[part] = (out[part] || 0) + (Number(r[cQ]) || 0);
  });
  return out;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'getActiveBatch') {
      const machine = String(p.machine || '').trim();
      if (!MACHINES.includes(machine)) return json({ ok: false, error: '無效機台: ' + machine });
      const batch = findActiveBatch_(machine);
      // 順便帶 progress，前端載入時少打一次
      return json({ ok: true, batch: batch, progress: batch ? batchProgress_(machine, batch.batchId) : null });
    }
    if (p.action === 'getBatchProgress') {
      const batchId = String(p.batchId || '').trim();
      if (!batchId) return json({ ok: false, error: '缺少 batchId' });
      // 機台名稱含「-」（3000T-1），從批次紀錄查機台，不從 ID 拆
      const { rows } = readBatchRows_();
      const r = rows.find(x => String(x[BATCH_COL['批次ID']]) === batchId);
      if (!r) return json({ ok: false, error: '找不到批次 ' + batchId });
      const batch = batchRowToObj_(r);
      return json({ ok: true, batchId: batchId, status: batch.status, progress: batchProgress_(batch.machine, batchId) });
    }
  } catch (err) {
    return json({ ok: false, error: err.toString() });
  }
  return json({ ok: true, message: 'A 棟射出生產 API 已就緒，請用 POST 送資料' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}