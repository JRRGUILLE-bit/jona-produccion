const SPREADSHEET_ID = '1ptU07Hvm5bqUSykgblUIWUckILW4eJX2uLmFzsN14jo';
const PAYMENTS_SHEET_NAMES = ['Pagos', 'Presupuesto'];
const NEW_PAYMENTS_SHEET_NAME = 'Pagos';
const TOTALS_SHEET_NAME = 'Totales';
const DEFAULT_CURRENCY = 'UYU';

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const ss = getSpreadsheet_();
    const params = e && e.parameter ? e.parameter : {};
    const nombre = String(params.nombre || '').trim();
    const moneda = normalizeCurrency_(params.moneda);
    const monto = parseAmount_(params.monto);

    if (!nombre) throw new Error('Falta elegir quién pagó.');
    if (!(monto > 0)) throw new Error('El monto debe ser mayor que cero.');

    const sheet = ensurePaymentsSheet_(ss);
    appendPayment_(sheet, {
      timestamp: new Date(),
      nombre: nombre,
      monto: monto,
      moneda: moneda
    });

    rebuildTotals_(ss, sheet);
    SpreadsheetApp.flush();

    return iframeResponse_({
      ok: true,
      moneda: moneda,
      spreadsheetId: ss.getId(),
      sheetName: sheet.getName()
    });
  } catch (error) {
    return iframeResponse_({
      ok: false,
      message: error && error.message ? error.message : 'No se pudo guardar el pago.'
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {}
  }
}

function doGet() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ensurePaymentsSheet_(ss);
    rebuildTotals_(ss, sheet);
    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        spreadsheet: ss.getName(),
        spreadsheetId: ss.getId(),
        paymentsSheet: sheet.getName()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        message: error && error.message ? error.message : 'Error de diagnóstico.'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setupNow() {
  const ss = getSpreadsheet_();
  const sheet = ensurePaymentsSheet_(ss);
  rebuildTotals_(ss, sheet);
  SpreadsheetApp.flush();
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function ensurePaymentsSheet_(ss) {
  let sheet = null;

  for (let i = 0; i < PAYMENTS_SHEET_NAMES.length; i++) {
    sheet = ss.getSheetByName(PAYMENTS_SHEET_NAMES[i]);
    if (sheet) break;
  }

  if (!sheet) sheet = ss.insertSheet(NEW_PAYMENTS_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([
      ['Timestamp', 'Nombre', 'Monto', 'Moneda']
    ]);
  }

  ensureRequiredColumns_(sheet);
  fillLegacyCurrency_(sheet);
  formatPaymentsSheet_(sheet);
  return sheet;
}

function ensureRequiredColumns_(sheet) {
  const required = [
    { key: 'timestamp', title: 'Timestamp' },
    { key: 'nombre', title: 'Nombre' },
    { key: 'monto', title: 'Monto' },
    { key: 'moneda', title: 'Moneda' }
  ];

  const map = getHeaderMap_(sheet);

  required.forEach(function(item) {
    if (!map[item.key]) {
      const column = sheet.getLastColumn() + 1;
      sheet.getRange(1, column).setValue(item.title);
      map[item.key] = column;
    }
  });
}

function fillLegacyCurrency_(sheet) {
  if (sheet.getLastRow() <= 1) return;

  const map = getHeaderMap_(sheet);
  const range = sheet.getRange(2, map.moneda, sheet.getLastRow() - 1, 1);
  const values = range.getValues();
  let changed = false;

  values.forEach(function(row) {
    const original = String(row[0] || '').trim();
    const normalized = original ? normalizeCurrency_(original) : DEFAULT_CURRENCY;

    if (original !== normalized) {
      row[0] = normalized;
      changed = true;
    }
  });

  if (changed) range.setValues(values);
}

function appendPayment_(sheet, payment) {
  const map = getHeaderMap_(sheet);
  const rowNumber = sheet.getLastRow() + 1;
  const values = new Array(sheet.getLastColumn()).fill('');

  values[map.timestamp - 1] = payment.timestamp;
  values[map.nombre - 1] = payment.nombre;
  values[map.monto - 1] = payment.monto;
  values[map.moneda - 1] = payment.moneda;

  sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
  sheet.getRange(rowNumber, map.timestamp).setNumberFormat('dd/MM/yyyy HH:mm');
  sheet.getRange(rowNumber, map.monto).setNumberFormat('#,##0.00');
}

function rebuildTotals_(ss, paymentsSheet) {
  let totals = ss.getSheetByName(TOTALS_SHEET_NAME);
  if (!totals) totals = ss.insertSheet(TOTALS_SHEET_NAME);

  const map = getHeaderMap_(paymentsSheet);
  const sums = { UYU: 0, USD: 0 };
  const counts = { UYU: 0, USD: 0 };

  if (paymentsSheet.getLastRow() > 1) {
    const data = paymentsSheet
      .getRange(2, 1, paymentsSheet.getLastRow() - 1, paymentsSheet.getLastColumn())
      .getValues();

    data.forEach(function(row) {
      const amount = Number(row[map.monto - 1]);
      const currency = normalizeCurrency_(row[map.moneda - 1]);

      if (!isFinite(amount) || amount <= 0) return;
      sums[currency] += amount;
      counts[currency] += 1;
    });
  }

  totals.getRange(1, 1, totals.getMaxRows(), totals.getMaxColumns()).breakApart();
  totals.clear();
  totals.getRange('A1:C1').merge().setValue('Totales por moneda');
  totals.getRange('A2:C2').merge().setValue(
    'Actualizado: ' + Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone(),
      'dd/MM/yyyy HH:mm'
    )
  );
  totals.getRange(4, 1, 3, 3).setValues([
    ['Moneda', 'Total', 'Pagos'],
    ['UYU', sums.UYU, counts.UYU],
    ['USD', sums.USD, counts.USD]
  ]);

  totals.getRange('A1:C1')
    .setFontWeight('bold')
    .setFontSize(16)
    .setBackground('#151515')
    .setFontColor('#f4f1e8');
  totals.getRange('A4:C4')
    .setFontWeight('bold')
    .setBackground('#333333')
    .setFontColor('#ffffff');
  totals.getRange('B5').setNumberFormat('"$" #,##0.00');
  totals.getRange('B6').setNumberFormat('"US$" #,##0.00');
  totals.setFrozenRows(4);
  totals.autoResizeColumns(1, 3);

  ss.setActiveSheet(totals);
  ss.moveActiveSheet(1);
  ss.setActiveSheet(paymentsSheet);
  ss.moveActiveSheet(2);
  ss.setActiveSheet(totals);
}

function formatPaymentsSheet_(sheet) {
  const map = getHeaderMap_(sheet);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setFontWeight('bold')
    .setBackground('#151515')
    .setFontColor('#f4f1e8');

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, map.timestamp, sheet.getLastRow() - 1, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.getRange(2, map.monto, sheet.getLastRow() - 1, 1)
      .setNumberFormat('#,##0.00');

    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['UYU', 'USD'], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, map.moneda, sheet.getLastRow() - 1, 1)
      .setDataValidation(validation);
  }

  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function getHeaderMap_(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getDisplayValues()[0];

  return headerMapFromValues_(headers);
}

function headerMapFromValues_(headers) {
  const map = {};

  headers.forEach(function(value, index) {
    const normalized = normalizeText_(value);
    const column = index + 1;

    if (!map.timestamp && ['timestamp', 'fecha', 'fecha y hora', 'hora'].indexOf(normalized) !== -1) map.timestamp = column;
    if (!map.nombre && ['nombre', 'persona', 'quien pago'].indexOf(normalized) !== -1) map.nombre = column;
    if (!map.monto && ['monto', 'importe', 'cuanto pago'].indexOf(normalized) !== -1) map.monto = column;
    if (!map.moneda && ['moneda', 'currency'].indexOf(normalized) !== -1) map.moneda = column;
  });

  return map;
}

function normalizeText_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function normalizeCurrency_(value) {
  const normalized = normalizeText_(value).toUpperCase();
  if (normalized === 'USD' || normalized === 'DOLAR' || normalized === 'DOLARES') return 'USD';
  return 'UYU';
}

function parseAmount_(value) {
  const normalized = String(value == null ? '' : value)
    .trim()
    .replace(',', '.');
  const amount = Number(normalized);
  return isFinite(amount) ? amount : NaN;
}

function iframeResponse_(data) {
  const payload = JSON.stringify(
    Object.assign({ source: 'dineros-tumberos' }, data)
  ).replace(/</g, '\\u003c');

  return HtmlService
    .createHtmlOutput(
      '<!doctype html><meta charset="utf-8">' +
      '<script>window.parent.postMessage(' + payload + ', "*");<\/script>'
    )
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
