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
 * シートの1行目（ヘッダー）を読み、カラム名の並びと名前→列インデックス(0始まり)を返す。
 * 列の移動・挿入に耐えるため、書き込み時はこの並びで行データを組み立てる。
 * @returns {{ headerNames: string[], nameToIndex: Object<string, number> }}
 */
function getSheetHeaderOrder(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    return { headerNames: [], nameToIndex: {} };
  }
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
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
  const lastRow = transferSheet.getLastRow();
  if (lastRow < 2) return keys;
  const header = getSheetHeaderOrder(transferSheet);
  const dateIdx = header.nameToIndex['日付'];
  const uuidIdx = header.nameToIndex['スタッフID'];
  if (dateIdx === undefined || uuidIdx === undefined) return keys;
  const numCols = transferSheet.getLastColumn();
  const rows = transferSheet.getRange(2, 1, lastRow, numCols).getValues();
  rows.forEach(function (row) {
    const dateVal = row[dateIdx];
    const uuidVal = row[uuidIdx];
    let dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (dateVal) {
      dateStr = String(dateVal).trim();
    }
    if (dateStr && uuidVal) keys.add(dateStr + '\t' + String(uuidVal));
  });
  return keys;
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
 * 日次転記バッチ: 打刻記録と各スタッフシートの「休み」を転記用シートに重複なく追記する。
 * @param {string} [targetDateStr] 対象日 (yyyy-MM-dd)。省略時は「昨日」。
 * @returns {{ ok: boolean, message: string, appended: number }}
 */
function runDailyTransfer(targetDateStr) {
  const tz = Session.getScriptTimeZone();
  if (!targetDateStr) {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    targetDateStr = Utilities.formatDate(yesterday, tz, 'yyyy-MM-dd');
  }

  const ss = getSpreadsheet();
  const tsSheet = ss.getSheetByName(TIMESTAMP_SHEET_NAME);
  if (!tsSheet) {
    return { ok: false, message: '打刻記録 シートが見つかりません', appended: 0 };
  }

  const transferSheet = ensureTransferSheet();
  const existingKeys = getTransferSheetExistingKeys(transferSheet);
  const header = getSheetHeaderOrder(transferSheet);
  const staffList = getStaffList();
  const familyNameToStaff = {};
  staffList.forEach(function (s) {
    const family = (s.name || '').trim().split(/\s+/)[0] || '';
    if (family && !familyNameToStaff[family]) familyNameToStaff[family] = s;
  });

  const toAppend = [];
  const tsHeader = getSheetHeaderOrder(tsSheet);
  const tsDateIdx = tsHeader.nameToIndex['日付'];
  const tsUuidIdx = tsHeader.nameToIndex['スタッフID'];
  const tsNameIdx = tsHeader.nameToIndex['氏名'];
  const tsStartIdx = tsHeader.nameToIndex['出勤時刻'];
  const tsEndIdx = tsHeader.nameToIndex['退勤時刻'];
  const tsBreakIdx = tsHeader.nameToIndex['休憩時間(時間)'];

  if (tsSheet.getLastRow() >= 2 && tsDateIdx !== undefined && tsUuidIdx !== undefined) {
    const tsRows = tsSheet.getRange(2, 1, tsSheet.getLastRow(), tsSheet.getLastColumn()).getValues();
    tsRows.forEach(function (row) {
      let dateStr = '';
      const dateVal = row[tsDateIdx];
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
      } else if (dateVal) {
        dateStr = String(dateVal).trim();
      }
      if (dateStr !== targetDateStr) return;
      const key = dateStr + '\t' + String(row[tsUuidIdx] || '');
      if (existingKeys.has(key)) return;
      existingKeys.add(key);
      const start = row[tsStartIdx] != null ? String(row[tsStartIdx]) : '';
      const end = row[tsEndIdx] != null ? String(row[tsEndIdx]) : '';
      const breakH = row[tsBreakIdx] != null ? Number(row[tsBreakIdx]) : 0;
      toAppend.push({
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
      });
    });
  }

  const sheetNames = [STAFF_SHEET_NAME, TIMESTAMP_SHEET_NAME, TRANSFER_SHEET_NAME];
  ss.getSheets().forEach(function (s) {
    const name = s.getName();
    if (sheetNames.indexOf(name) >= 0) return;
    const staff = familyNameToStaff[name];
    if (!staff) return;
    const shHeader = getSheetHeaderOrder(s);
    const dateIdx = shHeader.nameToIndex['日付'];
    const yasumiIdx = shHeader.nameToIndex['休み'];
    if (dateIdx === undefined || yasumiIdx === undefined) return;
    if (s.getLastRow() < 2) return;
    const rows = s.getRange(2, 1, s.getLastRow(), s.getLastColumn()).getValues();
    rows.forEach(function (row) {
      let dateStr = '';
      const dateVal = row[dateIdx];
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
      } else if (dateVal) {
        dateStr = String(dateVal).trim();
      }
      if (dateStr !== targetDateStr) return;
      const yasumi = row[yasumiIdx];
      if (yasumi === undefined || yasumi === null || String(yasumi).trim() === '') return;
      const key = dateStr + '\t' + staff.uuid;
      if (existingKeys.has(key)) return;
      existingKeys.add(key);
      toAppend.push({
        yearMonth: dateToYearMonth(dateStr),
        uuid: staff.uuid,
        name: staff.name,
        date: dateStr,
        kubun: String(yasumi).trim(),
        start: '',
        end: '',
        breakHours: '',
        workHours: '',
        biko: ''
      });
    });
  });

  if (toAppend.length === 0) {
    return { ok: true, message: '転記対象なし（既に転記済みまたは該当データなし）', appended: 0 };
  }

  const rowValues = toAppend.map(function (r) {
    return header.headerNames.map(function (colName) {
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
      return map[colName] !== undefined ? map[colName] : '';
    });
  });
  transferSheet.getRange(transferSheet.getLastRow() + 1, 1, transferSheet.getLastRow() + rowValues.length, header.headerNames.length).setValues(rowValues);
  return { ok: true, message: '転記用シートに ' + toAppend.length + ' 件追記しました', appended: toAppend.length };
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
    console.error('Failed to put cache:', e);
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
