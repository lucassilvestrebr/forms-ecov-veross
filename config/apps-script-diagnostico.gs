/**
 * ECOV / Veross - Recebimento de diagnóstico via HTML estático
 * 1) Crie uma planilha no Google Sheets.
 * 2) Cole o ID da planilha abaixo.
 * 3) Publique este script como Web App.
 */

const SPREADSHEET_ID = '1SFGuSgEn9mNR1r0zXN4TvzX5DoUESertGU7_ESg2IQo';
const SHEET_NAME = 'Diagnosticos';
const EXPECTED_TOKEN = 'ecov_diagnostico_v1';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const rawPayload = e && e.parameter && e.parameter.payload
      ? e.parameter.payload
      : e && e.postData && e.postData.contents
        ? e.postData.contents
        : '{}';

    const payload = JSON.parse(rawPayload);

    if (payload.form_token !== EXPECTED_TOKEN) {
      throw new Error('Token inválido.');
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    const data = flattenObject(payload);
    const headers = syncHeaders_(sheet, Object.keys(data));
    const row = headers.map(function(header) {
      const value = data[header];
      if (Array.isArray(value)) return value.join(' | ');
      if (value === null || value === undefined) return '';
      return value;
    });

    sheet.appendRow(row);

    return json_({ ok: true, message: 'Diagnóstico recebido com sucesso.' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

function syncHeaders_(sheet, newHeaders) {
  const lastColumn = sheet.getLastColumn();
  let headers = [];

  if (lastColumn > 0) {
    headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(String);
  }

  if (headers.length === 0) {
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    sheet.setFrozenRows(1);
    return newHeaders;
  }

  const missingHeaders = newHeaders.filter(function(header) {
    return headers.indexOf(header) === -1;
  });

  if (missingHeaders.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    headers = headers.concat(missingHeaders);
  }

  return headers;
}

function flattenObject(obj, prefix, out) {
  out = out || {};
  prefix = prefix || '';

  Object.keys(obj || {}).forEach(function(key) {
    const value = obj[key];
    const newKey = prefix ? prefix + '.' + key : key;

    if (
      value &&
      typeof value === 'object' &&
      !Array.isArray(value) &&
      Object.prototype.toString.call(value) !== '[object Date]'
    ) {
      flattenObject(value, newKey, out);
    } else {
      out[newKey] = value;
    }
  });

  return out;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
