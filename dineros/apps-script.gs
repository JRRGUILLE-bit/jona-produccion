const SPREADSHEET_ID = '1ptU07Hvm5bqUSykgbLUIWUckILW4eJX2uLmFzsN14jo';
const PAYMENTS_SHEET_NAMES = ['Pagos', 'Presupuesto'];
const NEW_PAYMENTS_SHEET_NAME = 'Pagos';
const TOTALS_SHEET_NAME = 'Totales';
const DEFAULT_CURRENCY = 'UYU';
const PEOPLE_ORDER = ['Pedro', 'Manu', 'Male', 'Maite', 'Mateo'];

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const ss = getSpreadsheet_();
    const params = e && e.parameter ? e.parameter : {};
    const nombre = canonicalName_(params.nombre);
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
  sheet.getRange(rowNumber, map.monto).setNumberFormat(
    payment.moneda === 'USD' ? '"US$" #,##0.00' : '"$" #,##0.00'
  );
}

function rebuildTotals_(ss, paymentsSheet) {
  let totals = ss.getSheetByName(TOTALS_SHEET_NAME);
  if (!totals) totals = ss.insertSheet(TOTALS_SHEET_NAME);

  const map = getHeaderMap_(paymentsSheet);
  const perPerson = {};
  const extraNames = [];
  const general = { UYU: 0, USD: 0, count: 0 };

  PEOPLE_ORDER.forEach(function(name) {
    perPerson[name] = { UYU: 0, USD: 0, count: 0 };
  });

  if (paymentsSheet.getLastRow() > 1) {
    const data = paymentsSheet
      .getRange(2, 1, paymentsSheet.getLastRow() - 1, paymentsSheet.getLastColumn())
      .getValues();

    data.forEach(function(row) {
      const name = canonicalName_(row[map.nombre - 1]);
      const amount = Number(row[map.monto - 1]);
      const currency = normalizeCurrency_(row[map.moneda - 1]);

      if (!name || !isFinite(amount) || amount <= 0) return;

      if (!perPerson[name]) {
        perPerson[name] = { UYU: 0, USD: 0, count: 0 };
        extraNames.push(name);
      }

      perPerson[name][currency] += amount;
      perPerson[name].count += 1;
      general[currency] += amount;
      general.count += 1;
    });
  }

  extraNames.sort(function(a, b) {
    return a.localeCompare(b, 'es', { sensitivity: 'base' });
  });

  const names = PEOPLE_ORDER.concat(extraNames);
  const bodyRows = names.map(function(name) {
    return [
      name,
      perPerson[name].UYU,
      perPerson[name].USD,
      perPerson[name].count
    ];
  });

  const bodyStartRow = 5;
  const bodyEndRow = bodyStartRow + bodyRows.length - 1;
  const totalRow = bodyEndRow + 2;

  ensureGridSize_(totals, totalRow, 4);
  totals.getRange(1, 1, totals.getMaxRows(), totals.getMaxColumns()).breakApart();
  totals.getBandings().forEach(function(banding) { banding.remove(); });
  totals.getCharts().forEach(function(chart) { totals.removeChart(chart); });
  totals.clear();
  totals.setConditionalFormatRules([]);
  totals.setHiddenGridlines(true);
  totals.setTabColor('#151515');

  totals.getRange('A1:D1').merge().setValue('Resumen de pagos');
  totals.getRange('A2:D2').merge().setValue(
    'Actualizado: ' + Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone(),
      'dd/MM/yyyy HH:mm'
    )
  );
  totals.getRange('A4:D4').setValues([
    ['Nombre', 'Total UYU', 'Total USD', 'Pagos']
  ]);
  totals.getRange(bodyStartRow, 1, bodyRows.length, 4).setValues(bodyRows);
  totals.getRange(totalRow, 1, 1, 4).setValues([
    ['TOTAL GENERAL', general.UYU, general.USD, general.count]
  ]);

  totals.getRange('A1:D1')
    .setFontWeight('bold')
    .setFontSize(17)
    .setBackground('#151515')
    .setFontColor('#f4f1e8')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  totals.getRange('A2:D2')
    .setFontSize(10)
    .setFontColor('#6b6256')
    .setHorizontalAlignment('center');

  totals.getRange('A4:D4')
    .setFontWeight('bold')
    .setBackground('#0d2f4f')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  const bodyRange = totals.getRange(bodyStartRow, 1, bodyRows.length, 4);
  bodyRange
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#d8d2c7', SpreadsheetApp.BorderStyle.SOLID);

  for (let row = bodyStartRow; row <= bodyEndRow; row++) {
    if ((row - bodyStartRow) % 2 === 1) {
      totals.getRange(row, 1, 1, 4).setBackground('#f5f3ee');
    }
  }

  totals.getRange(bodyStartRow, 1, bodyRows.length, 1).setFontWeight('bold');
  totals.getRange(bodyStartRow, 2, bodyRows.length, 1).setNumberFormat('"$" #,##0.00');
  totals.getRange(bodyStartRow, 3, bodyRows.length, 1).setNumberFormat('"US$" #,##0.00');
  totals.getRange(bodyStartRow, 4, bodyRows.length, 1).setNumberFormat('0');

  totals.getRange(totalRow, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#dfe7ee')
    .setBorder(true, true, true, true, true, true, '#0d2f4f', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  totals.getRange(totalRow, 2).setNumberFormat('"$" #,##0.00');
  totals.getRange(totalRow, 3).setNumberFormat('"US$" #,##0.00');
  totals.getRange(totalRow, 4).setNumberFormat('0');

  totals.getRange(4, 2, totalRow - 3, 3).setHorizontalAlignment('right');
  totals.setFrozenRows(4);
  totals.setColumnWidth(1, 150);
  totals.setColumnWidth(2, 130);
  totals.setColumnWidth(3, 130);
  totals.setColumnWidth(4, 80);
  totals.setRowHeight(1, 34);
  totals.setRowHeight(2, 22);
  totals.setRowHeight(4, 28);

  ss.setActiveSheet(totals);
  ss.moveActiveSheet(1);
  ss.setActiveSheet(paymentsSheet);
  ss.moveActiveSheet(2);
  ss.setActiveSheet(totals);
}

function formatPaymentsSheet_(sheet) {
  const map = getHeaderMap_(sheet);
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(false);
  sheet.setTabColor('#8a6a3f');
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setFontWeight('bold')
    .setBackground('#151515')
    .setFontColor('#f4f1e8')
    .setHorizontalAlignment('center');

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['UYU', 'USD'], true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, map.moneda, Math.max(sheet.getMaxRows() - 1, 1), 1)
    .setDataValidation(validation);

  if (sheet.getLastRow() > 1) {
    const rowCount = sheet.getLastRow() - 1;
    sheet.getRange(2, map.timestamp, rowCount, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm');

    const currencyValues = sheet.getRange(2, map.moneda, rowCount, 1).getValues();
    const numberFormats = currencyValues.map(function(row) {
      return [normalizeCurrency_(row[0]) === 'USD'
        ? '"US$" #,##0.00'
        : '"$" #,##0.00'];
    });

    sheet.getRange(2, map.monto, rowCount, 1).setNumberFormats(numberFormats);
  }

  sheet.setColumnWidth(map.timestamp, 150);
  sheet.setColumnWidth(map.nombre, 120);
  sheet.setColumnWidth(map.monto, 120);
  sheet.setColumnWidth(map.moneda, 90);
}

function ensureGridSize_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
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

function canonicalName_(value) {
  const raw = String(value || '').trim();
  const normalized = normalizeText_(raw);

  for (let i = 0; i < PEOPLE_ORDER.length; i++) {
    if (normalizeText_(PEOPLE_ORDER[i]) === normalized) return PEOPLE_ORDER[i];
  }

  return raw;
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
