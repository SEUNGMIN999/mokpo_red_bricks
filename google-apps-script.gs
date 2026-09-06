/* 목포 : 붉은 벽돌의 시간 — Google Sheets 기록관 API
 * Google Sheets에서 확장 프로그램 > Apps Script를 열고 이 파일을 붙여 넣으세요.
 */

const SPREADSHEET_ID = '여기에_구글_시트_ID를_입력하세요';
const SHEET_NAME = '기억의 기록관';
const MAX_MESSAGE_LENGTH = 100;

function doGet(e) {
  try {
    if (String((e && e.parameter && e.parameter.action) || '') !== 'list') {
      return output_({ ok: false, error: '지원하지 않는 요청입니다.' }, e);
    }
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return output_({ ok: true, records: [] }, e);

    const rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const records = rows
      .map(row => ({
        id: String(row[0]),
        createdAt: row[1] instanceof Date ? row[1].toISOString() : String(row[1]),
        alias: String(row[2]),
        message: String(row[3])
      }))
      .slice(-24);
    return output_({ ok: true, records }, e);
  } catch (error) {
    return output_({ ok: false, error: '기록을 불러오지 못했습니다.' }, e);
  }
}

function doPost(e) {
  try {
    const params = (e && e.parameter) || {};
    if (String(params.action || '') !== 'submit') {
      return json_({ ok: false, error: '지원하지 않는 요청입니다.' });
    }

    const recordId = String(params.recordId || '').trim();
    const deviceToken = String(params.deviceToken || '').trim();
    const message = cleanMessage_(params.message);
    if (!/^[A-Za-z0-9_-]{16,80}$/.test(recordId) ||
        !/^[A-Za-z0-9_-]{16,80}$/.test(deviceToken)) {
      return json_({ ok: false, error: '잘못된 기록 식별자입니다.' });
    }
    if (!message || Array.from(message).length > MAX_MESSAGE_LENGTH || containsContact_(message)) {
      return json_({ ok: false, error: '기록 내용을 확인해 주세요.' });
    }

    const fingerprint = sha256_(deviceToken);
    const cache = CacheService.getScriptCache();
    const rateKey = `write-${fingerprint.slice(0, 40)}`;
    if (cache.get(rateKey)) {
      return json_({ ok: false, error: '잠시 후 다시 시도해 주세요.' });
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) return json_({ ok: false, error: '기록관이 혼잡합니다.' });
    try {
      const sheet = getSheet_();
      const row = sheet.getLastRow() + 1;
      const alias = `기록자 ${String(row - 1).padStart(3, '0')}`;
      sheet.getRange(row, 1, 1, 5).setValues([
        [recordId, new Date(), alias, '', fingerprint]
      ]);
      // 수식으로 해석되지 않도록 메시지 셀을 일반 텍스트로 고정합니다.
      sheet.getRange(row, 4).setNumberFormat('@').setValue(message);
      cache.put(rateKey, '1', 60);
      return json_({ ok: true, id: recordId, alias });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return json_({ ok: false, error: '기록을 저장하지 못했습니다.' });
  }
}

function getSheet_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf('여기에_') === 0) {
    throw new Error('SPREADSHEET_ID를 설정하세요.');
  }
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['기록 ID', '작성 시각', '기록자', '메시지', '기기 해시']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#7f3929').setFontColor('#ffffff');
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 110);
    sheet.setColumnWidth(4, 440);
  }
  return sheet;
}

function cleanMessage_(value) {
  return String(value || '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .trim();
}

function containsContact_(message) {
  return /(https?:\/\/|www\.|[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}|01[016789][ -]?\d{3,4}[ -]?\d{4})/i.test(message);
}

function sha256_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8)
    .map(byte => (byte + 256) % 256)
    .map(byte => byte.toString(16).padStart(2, '0'))
    .join('');
}

function output_(payload, e) {
  const prefix = String((e && e.parameter && e.parameter.prefix) || '');
  if (/^[A-Za-z_$][0-9A-Za-z_$]{0,80}$/.test(prefix)) {
    return ContentService
      .createTextOutput(`${prefix}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
