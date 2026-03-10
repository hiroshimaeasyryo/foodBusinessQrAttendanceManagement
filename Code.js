/**
 * 飲食系出勤管理（AccessControlFoodBusiness）
 * スプレッドシート ID 指定で開く
 */
const SPREADSHEET_ID = '1BzM547ikvZIjXLEQBktuRxeU0UaiPH7Td1m80tBoSGE';
const STAFF_SHEET_NAME = 'スタッフDB';
const TIMESTAMP_SHEET_NAME = '打刻記録';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
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
 * 契約社員: 18:00-24:00 休憩なし
 */
function getDefaultTimes(property, shiftType) {
  if (property === '契約社員') {
    return { start: '18:00', end: '24:00', breakMinutes: 0 };
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

  // 打刻記録シート: タイムスタンプ, スタッフID, 氏名, 出勤時刻, 退勤時刻, 休憩時間(時間)
  var rowBefore = tsSheet.getLastRow();
  try {
    tsSheet.appendRow([
      now,
      payload.uuid || '',
      payload.name || '',
      startTime,
      endTime,
      breakHours
    ]);
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
