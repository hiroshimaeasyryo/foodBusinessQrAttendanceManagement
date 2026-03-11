/**
 * 飲食系出勤管理（AccessControlFoodBusiness）
 * スプレッドシート ID 指定で開く
 */
const SPREADSHEET_ID = '1BzM547ikvZIjXLEQBktuRxeU0UaiPH7Td1m80tBoSGE';
const STAFF_SHEET_NAME = 'スタッフDB';
const TIMESTAMP_SHEET_NAME = '打刻記録';
const TRANSFER_SHEET_NAME = '転記用';

/** 打刻記録シートの想定カラム名（生ログ。区分・勤務時間・備考は転記用で賄う） */
const TIMESTAMP_SHEET_COLUMNS = [
  'タイムスタンプ',
  'スタッフID',
  '氏名',
  '日付',
  '出勤時刻',
  '退勤時刻',
  '休憩時間(時間)'
];

/**
 * 転記用シートのカラム名（別Spreadsheetと同じ構成。確定フラグは転記先で使うため本シートには設けない）。
 * 転記バッチで「打刻記録」および各スタッフシートの「休み」列を観察し、重複なく追記する。
 */
const TRANSFER_SHEET_COLUMNS = [
  '年月',
  'スタッフID',
  '氏名',
  '日付',
  '区分',
  '出勤開始',
  '出勤終了',
  '休憩時間(時間)',
  '勤務時間(時間)',
  '備考'
];

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * シートの指定行（1始まり）をヘッダーとして読み、カラム名の並びと名前→列インデックス(0始まり)を返す。
 * 列の移動・挿入に耐えるため、書き込み時はこの並びで行データを組み立てる。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} [headerRowOneBased] ヘッダー行（1始まり）。省略時は 1。
 * @returns {{ headerNames: string[], nameToIndex: Object<string, number> }}
 */
function getSheetHeaderOrder(sheet, headerRowOneBased) {
  var row = headerRowOneBased == null ? 1 : Math.max(1, headerRowOneBased);
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    return { headerNames: [], nameToIndex: {} };
  }
  const headerRow = sheet.getRange(row, 1, row, lastCol).getValues()[0];
  const headerNames = headerRow.map(function (cell) { return String(cell || '').trim(); });
  const nameToIndex = {};
  headerNames.forEach(function (name, idx) {
    if (name) nameToIndex[name] = idx;
  });
  return { headerNames: headerNames, nameToIndex: nameToIndex };
}

/**
 * "HH:mm" または "H:mm" を分に変換。24:00 → 1440、25:00 → 1500 等の24時超え表記に対応。
 */
function parseTimeToMinutes(timeStr) {
  if (!timeStr || typeof timeStr !== 'string') return 0;
  const s = String(timeStr).trim();
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return 0;
  const hours = parseInt(m[1], 10) || 0;
  const minutes = parseInt(m[2], 10) || 0;
  return hours * 60 + minutes;
}

/**
 * 出勤・退勤時刻（"HH:mm"）と休憩時間(時間)から、勤務時間(時間)を算出。
 * 休憩を差し引いた実労働時間を返す。
 */
function calcWorkHours(startTimeStr, endTimeStr, breakHours) {
  const startM = parseTimeToMinutes(startTimeStr);
  let endM = parseTimeToMinutes(endTimeStr);
  if (endM <= startM && startM >= 0) endM += 24 * 60; // 日跨ぎ
  const workMinutes = Math.max(0, endM - startM - (Number(breakHours) || 0) * 60);
  return Math.round(workMinutes * 100) / 100 / 60;
}

/**
 * 転記用シートを取得または作成し、ヘッダーが無ければ設定する。
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureTransferSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(TRANSFER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TRANSFER_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, TRANSFER_SHEET_COLUMNS.length).setValues([TRANSFER_SHEET_COLUMNS]);
    sheet.getRange(1, 1, 1, TRANSFER_SHEET_COLUMNS.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * 転記用シートに既に存在する (日付, スタッフID) のセットを返す。重複排除に使う。
 * @returns {Set<string>} 各要素は "日付\tスタッフID"
 */
function getTransferSheetExistingKeys(transferSheet) {
  const keys = new Set();
  const map = getTransferSheetExistingRows(transferSheet);
  Object.keys(map).forEach(function (k) { keys.add(k); });
  return keys;
}

/**
 * 転記用シートの既存行を行番号(2始まり)と行データで返す。更新反映で使用。
 * @returns {Object<string, { rowIndex: number, rowData: Array }} キーは "日付\tスタッフID"
 */
function getTransferSheetExistingRows(transferSheet) {
  const result = {};
  const lastRow = transferSheet.getLastRow();
  if (lastRow < 2) return result;
  const header = getSheetHeaderOrder(transferSheet);
  const dateIdx = header.nameToIndex['日付'];
  const uuidIdx = header.nameToIndex['スタッフID'];
  if (dateIdx === undefined || uuidIdx === undefined) return result;
  const numCols = transferSheet.getLastColumn();
  const rows = transferSheet.getRange(2, 1, lastRow, numCols).getValues();
  rows.forEach(function (row, i) {
    const dateVal = row[dateIdx];
    const uuidVal = row[uuidIdx];
    let dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (dateVal) {
      dateStr = String(dateVal).trim();
    }
    if (dateStr && uuidVal) {
      const key = dateStr + '\t' + String(uuidVal);
      result[key] = { rowIndex: i + 2, rowData: row };
    }
  });
  return result;
}

/**
 * 日付文字列 (yyyy-MM-dd) から年月 (YYYYMM) を返す。
 */
function dateToYearMonth(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return '';
  const s = String(dateStr).trim();
  const m = s.match(/^(\d{4})-(\d{2})/);
  return m ? m[1] + m[2] : '';
}

/**
 * セル値（Date または "2026/03/08" / "2026年3月8日(日)" 等）を yyyy-MM-dd に正規化。dateSet 照合用。
 */
function normalizeDateStr(dateVal, tz) {
  if (!dateVal && dateVal !== 0) return '';
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
  }
  const s = String(dateVal).trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (!m) m = s.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (m) {
    const y = m[1];
    const mon = ('0' + m[2]).slice(-2);
    const d = ('0' + m[3]).slice(-2);
    return y + '-' + mon + '-' + d;
  }
  return s || '';
}

/**
 * 転記用の「出勤開始」「出勤終了」用に時刻を "HH:mm" 文字列で返す。Date の場合は時刻のみ書式化。
 */
function formatTimeForTransfer(val, tz) {
  if (val == null || val === '') return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz, 'HH:mm');
  }
  const s = String(val).trim();
  if (/^\d{1,2}:\d{2}$/.test(s)) return s;
  return s;
}

/**
 * 開始日〜終了日（yyyy-MM-dd）の日付文字列配列を返す（両端含む）。
 */
function getDateRangeInclusive(startDateStr, endDateStr) {
  const tz = Session.getScriptTimeZone();
  const start = new Date(startDateStr + 'T00:00:00');
  const end = new Date(endDateStr + 'T00:00:00');
  if (start.getTime() > end.getTime()) return [];
  const list = [];
  const d = new Date(start);
  while (d.getTime() <= end.getTime()) {
    list.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
    d.setDate(d.getDate() + 1);
  }
  return list;
}

/**
 * 日付範囲内の「正規1行」を 打刻記録 と 各スタッフシート から構築する。
 * キーは "日付\tスタッフID"。出勤があれば出勤、なければ休みがあれば休み。
 * @returns {Object<string, { yearMonth, uuid, name, date, kubun, start, end, breakHours, workHours, biko }>}
 */
function buildCanonicalRowsForDateRange(ss, startDateStr, endDateStr, staffList, familyNameToStaff, tz) {
  const dateSet = {};
  getDateRangeInclusive(startDateStr, endDateStr).forEach(function (d) { dateSet[d] = true; });

  const canonical = {};
  const tsSheet = ss.getSheetByName(TIMESTAMP_SHEET_NAME);
  if (tsSheet && tsSheet.getLastRow() >= 2) {
    const tsHeader = getSheetHeaderOrder(tsSheet);
    const tsDateIdx = tsHeader.nameToIndex['日付'];
    const tsUuidIdx = tsHeader.nameToIndex['スタッフID'];
    const tsNameIdx = tsHeader.nameToIndex['氏名'];
    const tsStartIdx = tsHeader.nameToIndex['出勤時刻'];
    const tsEndIdx = tsHeader.nameToIndex['退勤時刻'];
    const tsBreakIdx = tsHeader.nameToIndex['休憩時間(時間)'];
    if (tsDateIdx !== undefined && tsUuidIdx !== undefined) {
      const tsRows = tsSheet.getRange(2, 1, tsSheet.getLastRow(), tsSheet.getLastColumn()).getValues();
      tsRows.forEach(function (row) {
        const dateStr = normalizeDateStr(row[tsDateIdx], tz);
        if (!dateStr || !dateSet[dateStr]) return;
        const key = dateStr + '\t' + String(row[tsUuidIdx] || '');
        const start = formatTimeForTransfer(row[tsStartIdx], tz);
        const end = formatTimeForTransfer(row[tsEndIdx], tz);
        const breakH = row[tsBreakIdx] != null ? Number(row[tsBreakIdx]) : 0;
        canonical[key] = {
          yearMonth: dateToYearMonth(dateStr),
          uuid: row[tsUuidIdx],
          name: row[tsNameIdx] != null ? String(row[tsNameIdx]) : '',
          date: dateStr,
          kubun: '出勤',
          start: start,
          end: end,
          breakHours: breakH,
          workHours: calcWorkHours(start, end, breakH),
          biko: ''
        };
      });
    }
  }

  // スタッフ別シート: 1-2行目がヘッダー、3行目以降がデータ（recordTimestamp と同様）。2行目に「日付」「休み」がなければ1行目をヘッダーとする。
  const sheetNames = [STAFF_SHEET_NAME, TIMESTAMP_SHEET_NAME, TRANSFER_SHEET_NAME];
  ss.getSheets().forEach(function (s) {
    const name = s.getName();
    if (sheetNames.indexOf(name) >= 0) return;
    const staff = familyNameToStaff[name];
    if (!staff) return;
    var shHeader = getSheetHeaderOrder(s, 2);
    var dateIdx = shHeader.nameToIndex['日付'];
    var yasumiIdx = shHeader.nameToIndex['休み'];
    var dataFirstRow = 3;
    if (dateIdx === undefined || yasumiIdx === undefined) {
      shHeader = getSheetHeaderOrder(s, 1);
      dateIdx = shHeader.nameToIndex['日付'];
      yasumiIdx = shHeader.nameToIndex['休み'];
      dataFirstRow = 2;
    }
    if (dateIdx === undefined || yasumiIdx === undefined) return;
    const lastRow = s.getLastRow();
    if (lastRow < dataFirstRow) return;
    const bikoIdx = shHeader.nameToIndex['備考'];
    const rows = s.getRange(dataFirstRow, 1, lastRow, s.getLastColumn()).getValues();
    rows.forEach(function (row) {
      const dateStr = normalizeDateStr(row[dateIdx], tz);
      if (!dateStr || !dateSet[dateStr]) return;
      const yasumi = row[yasumiIdx];
      if (yasumi === undefined || yasumi === null || String(yasumi).trim() === '') return;
      const key = dateStr + '\t' + staff.uuid;
      if (canonical[key]) return; // 出勤優先
      const biko = (bikoIdx !== undefined && row[bikoIdx] != null) ? String(row[bikoIdx]).trim() : '';
      canonical[key] = {
        yearMonth: dateToYearMonth(dateStr),
        uuid: staff.uuid,
        name: staff.name,
        date: dateStr,
        kubun: String(yasumi).trim(),
        start: '',
        end: '',
        breakHours: '',
        workHours: '',
        biko: biko
      };
    });
  });

  return canonical;
}

/** 正規行オブジェクトを転記用ヘッダー順の配列に変換する */
function canonicalRowToValues(r, headerNames) {
  const map = {
    '年月': r.yearMonth,
    'スタッフID': r.uuid,
    '氏名': r.name,
    '日付': r.date,
    '区分': r.kubun,
    '出勤開始': r.start,
    '出勤終了': r.end,
    '休憩時間(時間)': r.breakHours,
    '勤務時間(時間)': r.workHours,
    '備考': r.biko
  };
  return headerNames.map(function (colName) {
    const v = map[colName];
    return v !== undefined && v !== null ? v : '';
  });
}

/** 転記用の既存1行と正規1行が「同じ内容」か比較する（日付は文字列に正規化） */
function transferRowEquals(existingRow, canonicalValues, header) {
  if (existingRow.length !== canonicalValues.length) return false;
  const dateIdx = header.nameToIndex['日付'];
  for (var i = 0; i < existingRow.length; i++) {
    var ex = existingRow[i];
    var canon = canonicalValues[i];
    if (dateIdx === i && ex instanceof Date) {
      ex = Utilities.formatDate(ex, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    if (typeof ex === 'number' && typeof canon === 'number') {
      if (Math.abs(ex - canon) > 1e-6) return false;
    } else if (String(ex).trim() !== String(canon).trim()) {
      return false;
    }
  }
  return true;
}

/**
 * 日次転記バッチ: 打刻記録と各スタッフシートの「休み」を転記用シートに重複なく追記し、
 * 過去範囲に指定がある場合は既存行の内容を正規データと照合して更新する。
 * @param {string} [targetDateStr] 対象日 (yyyy-MM-dd)。省略時は「昨日」。
 * @param {{ backwardDays?: number }} [options] backwardDays: 対象日の何日前までを「更新対象」にするか。0のときは追記のみ。例: 60 で過去60日分の変更を反映。
 * @returns {{ ok: boolean, message: string, appended: number, updated: number }}
 */
function runDailyTransfer(targetDateStr, options) {
  const tz = Session.getScriptTimeZone();
  if (!targetDateStr) {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    targetDateStr = Utilities.formatDate(yesterday, tz, 'yyyy-MM-dd');
  }
  options = options || {};
  const backwardDays = Math.max(0, parseInt(options.backwardDays, 10) || 0);

  const ss = getSpreadsheet();
  const tsSheet = ss.getSheetByName(TIMESTAMP_SHEET_NAME);
  if (!tsSheet) {
    return { ok: false, message: '打刻記録 シートが見つかりません', appended: 0, updated: 0 };
  }

  const transferSheet = ensureTransferSheet();
  const header = getSheetHeaderOrder(transferSheet);
  const staffList = getStaffList();
  const familyNameToStaff = {};
  staffList.forEach(function (s) {
    const family = (s.name || '').trim().split(/\s+/)[0] || '';
    if (family && !familyNameToStaff[family]) familyNameToStaff[family] = s;
  });

  var startDateStr = targetDateStr;
  if (backwardDays > 0) {
    var startDate = new Date(targetDateStr + 'T00:00:00');
    startDate.setDate(startDate.getDate() - backwardDays);
    startDateStr = Utilities.formatDate(startDate, tz, 'yyyy-MM-dd');
  }
  const canonical = buildCanonicalRowsForDateRange(ss, startDateStr, targetDateStr, staffList, familyNameToStaff, tz);
  const existingRows = getTransferSheetExistingRows(transferSheet);
  var lastRowBeforeChanges = transferSheet.getLastRow();

  var appended = 0;
  var updated = 0;
  const toAppend = [];
  const toUpdate = [];

  Object.keys(canonical).forEach(function (key) {
    const r = canonical[key];
    const canonicalValues = canonicalRowToValues(r, header.headerNames);
    const existing = existingRows[key];
    if (!existing) {
      toAppend.push(canonicalValues);
      appended++;
    } else {
      if (!transferRowEquals(existing.rowData, canonicalValues, header)) {
        toUpdate.push({ rowIndex: existing.rowIndex, values: canonicalValues });
        updated++;
      }
    }
  });

  if (toUpdate.length > 0) {
    var numCols = header.headerNames.length;
    toUpdate.forEach(function (item) {
      // getRange(row, column, numRows, numColumns) で 1 行だけ指定（終了行指定だと numRows と解釈され範囲がずれる）
      transferSheet.getRange(item.rowIndex, 1, 1, numCols).setValues([item.values]);
    });
  }
  if (toAppend.length > 0) {
    var numAppendRows = toAppend.length;
    var appendStartRow = lastRowBeforeChanges + 1;
    var numCols = header.headerNames.length;
    for (var i = 0; i < toAppend.length; i++) {
      if (toAppend[i].length !== numCols) {
        throw new Error('転記用: 行' + (i + 1) + 'の列数がヘッダーと一致しません（データ=' + toAppend[i].length + ', ヘッダー=' + numCols + '）。転記用シート1行目の列数を確認してください。');
      }
    }
    // getRange(row, column, numRows, numColumns) で行数・列数を明示（終了行計算のずれを防ぐ）
    transferSheet.getRange(appendStartRow, 1, numAppendRows, numCols).setValues(toAppend);
  }

  var msg = '';
  if (appended > 0) msg += '転記用に ' + appended + ' 件追記';
  if (updated > 0) msg += (msg ? '。' : '') + '既存 ' + updated + ' 件を更新（過去の変更を反映）';
  if (appended === 0 && updated === 0) msg = '転記対象なし（既に転記済みかつ変更なし）';
  return { ok: true, message: msg, appended: appended, updated: updated };
}

/**
 * 毎日定時実行用。トリガーにはこの関数を指定する。
 * 昨日分を転記し、過去60日分の「休み」変更・行挿入などを転記用に反映する。
 */
function runDailyTransferScheduled() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const targetDateStr = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return runDailyTransfer(targetDateStr, { backwardDays: 60 });
}

/**
 * メイン画面表示 / API エンドポイント
 */
function doGet(e) {
  const action = e.parameter.action;

  if (action) {
    try {
      let result;
      switch (action) {
        case 'getStaffList':
          result = getStaffList();
          break;
        case 'getStaffListByBirthKey':
          result = getStaffListByBirthKey(e.parameter.birthKey);
          break;
        case 'verifyStaff':
          result = verifyStaff(e.parameter.uuid, e.parameter.birthdate);
          break;
        case 'recordTimestamp': {
          // case 内で const/let を使う場合は必ず {} でブロック（未ブロックだと SyntaxError になる）
          const payload = JSON.parse(e.parameter.payload);
          result = recordTimestamp(payload);
          break;
        }
        case 'clearStaffCache':
          result = clearStaffCache();
          break;
        default:
          throw new Error('Unknown action: ' + action);
      }
      const output = JSON.stringify(result);
      const callback = e.parameter.callback;
      if (callback) {
        return ContentService.createTextOutput(callback + '(' + output + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(output)
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      const errorOutput = JSON.stringify({ ok: false, message: err.message });
      const callback = e.parameter.callback;
      if (callback) {
        return ContentService.createTextOutput(callback + '(' + errorOutput + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(errorOutput)
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  const tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.placeId = e.parameter.place || '';
  try {
    tmpl.gasUrl = ScriptApp.getService().getUrl();
  } catch (err) {
    tmpl.gasUrl = '';
  }
  // ロゴURL（GitHub Pages 等の絶対URLを指定すると hero と favicon に表示）
  tmpl.logoUrl = '';
  return tmpl.evaluate()
    .setTitle('出勤記録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * スタッフ一覧を取得（uuid, name, birthdate, property）
 * フロントで誕生日4桁→候補表示するため、birthKey (MMdd) も付与
 * return: [{ uuid, name, birthdate, property, birthKey }, ...]
 */
function getStaffList() {
  const cacheKey = 'staff_list_cache';
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);

  if (cachedData) {
    return JSON.parse(cachedData);
  }

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) {
    throw new Error('スタッフDB シートが見つかりません');
  }

  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const uuidIndex = header.indexOf('uuid');
  const nameIndex = header.indexOf('name');
  const birthIndex = header.indexOf('birthdate');
  const propertyIndex = header.indexOf('property');

  if (uuidIndex === -1 || nameIndex === -1 || birthIndex === -1) {
    throw new Error('スタッフDB に uuid / name / birthdate 列がありません');
  }

  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[uuidIndex] || !row[nameIndex]) continue;

    let birthStr = '';
    const cell = row[birthIndex];
    if (cell instanceof Date) {
      birthStr = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (cell) {
      birthStr = String(cell).trim();
    }

    // birthKey = MMdd（4桁）でフロントの即時マッチ用
    let birthKey = '';
    if (birthStr.length >= 10) {
      birthKey = birthStr.slice(5, 7) + birthStr.slice(8, 10);
    }

    result.push({
      uuid: String(row[uuidIndex]),
      name: String(row[nameIndex]).trim(),
      birthdate: birthStr,
      property: propertyIndex >= 0 && row[propertyIndex] ? String(row[propertyIndex]).trim() : '',
      birthKey: birthKey
    });
  }

  try {
    cache.put(cacheKey, JSON.stringify(result), 21600);
  } catch (e) {
    // キャッシュ書き込み失敗時は無視（次回取得で再取得）
  }

  return result;
}

/**
 * 誕生日4桁 (MMdd) でスタッフ候補を取得（キャッシュ経由で一覧からフィルタ）
 */
function getStaffListByBirthKey(birthKey) {
  const list = getStaffList();
  const key = String(birthKey).replace(/\D/g, '');
  if (key.length !== 4) return [];
  return list.filter(function (s) { return s.birthKey === key; });
}

function clearStaffCache() {
  CacheService.getScriptCache().remove('staff_list_cache');
  return { ok: true, message: 'キャッシュをクリアしました' };
}

/**
 * 生年月日とスタッフUUIDを検証
 */
function verifyStaff(uuid, birthdateStr) {
  const list = getStaffList();
  const normalized = String(birthdateStr).trim();

  for (let i = 0; i < list.length; i++) {
    if (list[i].uuid !== uuid) continue;
    if (list[i].birthdate === normalized) {
      return { ok: true, name: list[i].name };
    }
    return { ok: false, message: '生年月日が一致しません' };
  }
  return { ok: false, message: '該当のスタッフが見つかりません' };
}

/**
 * 属性に応じたデフォルト時刻を返す
 * 社員: shiftType 'default_11' => 11:00-24:00 2h休憩, 'default_18' => 18:00-24:00 休憩なし
 * 契約社員: 18:00-22:00 休憩なし
 */
function getDefaultTimes(property, shiftType) {
  if (property === '契約社員') {
    return { start: '18:00', end: '22:00', breakMinutes: 0 };
  }
  if (property === '社員') {
    if (shiftType === 'default_18') {
      return { start: '18:00', end: '24:00', breakMinutes: 0 };
    }
    return { start: '11:00', end: '24:00', breakMinutes: 120 };
  }
  return { start: '', end: '', breakMinutes: 0 };
}

/**
 * 打刻を記録
 * payload: {
 *   uuid, name, property,
 *   startTime?, endTime?, breakMinutes? (イレギュラー用・分。シートには時間に換算して記録),
 *   shiftType? ('default_11' | 'default_18') 社員用
 * }
 * 業務委託の場合は何も記録しない。
 */
function recordTimestamp(payload) {
  const property = payload.property || '';

  if (property === '業務委託') {
    return {
      ok: true,
      message: '記録なし（業務委託）',
      recorded: false
    };
  }

  const ss = getSpreadsheet();
  const tsSheet = ss.getSheetByName(TIMESTAMP_SHEET_NAME);
  if (!tsSheet) {
    throw new Error('打刻記録 シートが見つかりません');
  }

  let startTime = payload.startTime;
  let endTime = payload.endTime;
  let breakMinutes = payload.breakMinutes;

  if (startTime == null || startTime === '' || endTime == null || endTime === '') {
    const def = getDefaultTimes(property, payload.shiftType);
    startTime = startTime != null && startTime !== '' ? startTime : def.start;
    endTime = endTime != null && endTime !== '' ? endTime : def.end;
    if (breakMinutes == null || breakMinutes === '') breakMinutes = def.breakMinutes;
  }

  if (startTime == null) startTime = '';
  if (endTime == null) endTime = '';
  if (breakMinutes == null || breakMinutes === '') breakMinutes = 0;
  breakMinutes = Number(breakMinutes) || 0;
  // シートへは時間単位で記録（入力は分のまま: 90分 → 1.5）
  const breakHours = breakMinutes / 60;

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // ヘッダーが無い場合は1行目に想定カラムを設定（列の移動・挿入に耐えるため、以降はカラム名で位置解決）
  if (tsSheet.getLastRow() === 0) {
    tsSheet.getRange(1, 1, 1, TIMESTAMP_SHEET_COLUMNS.length).setValues([TIMESTAMP_SHEET_COLUMNS]);
    tsSheet.getRange(1, 1, 1, TIMESTAMP_SHEET_COLUMNS.length).setFontWeight('bold');
  }

  const header = getSheetHeaderOrder(tsSheet);
  const rowData = {
    'タイムスタンプ': now,
    'スタッフID': payload.uuid || '',
    '氏名': payload.name || '',
    '日付': dateStr,
    '出勤時刻': startTime,
    '退勤時刻': endTime,
    '休憩時間(時間)': breakHours
  };
  const rowValues = header.headerNames.map(function (colName) {
    return rowData[colName] !== undefined ? rowData[colName] : '';
  });

  var rowBefore = tsSheet.getLastRow();
  try {
    tsSheet.appendRow(rowValues);
  } catch (appendErr) {
    return {
      ok: false,
      message: 'シートへの書き込みに失敗しました: ' + appendErr.message,
      recorded: false
    };
  }
  var rowAfter = tsSheet.getLastRow();
  if (rowAfter !== rowBefore + 1) {
    return {
      ok: false,
      message: '打刻記録シートに行が追加されていません（権限・シート名・共有設定を確認してください）',
      recorded: false
    };
  }

  // スタッフ別シート: 氏名は "姓 名" → シート名は姓（先頭の単語）
  const namePart = (payload.name || '').trim().split(/\s+/)[0] || '';
  if (namePart) {
    const staffSheet = ss.getSheetByName(namePart);
    if (staffSheet) {
      // 1-2行目がヘッダーのため、3行目以降に追記。既存データの末尾に追加
      const lastRow = staffSheet.getLastRow();
      const headerRows = 2;
      const nextRow = Math.max(lastRow + 1, headerRows + 1);
      // getRange(row, column, numRows, numColumns) — 第3・4引数は「行数・列数」。1行4列に誤って nextRow を渡すと
      // 「データは1行だが範囲は nextRow 行」エラーになるため、numRows=1, numColumns=4 で指定する
      staffSheet.getRange(nextRow, 1, 1, 4).setValues([[
        dateStr,
        startTime,
        endTime,
        breakHours
      ]]);
    }
  }

  return {
    ok: true,
    message: '出勤記録を保存しました',
    recorded: true,
    row: rowAfter,
    timestamp: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    startTime: startTime,
    endTime: endTime,
    breakMinutes: breakMinutes
  };
}
