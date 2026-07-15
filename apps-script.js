const SHEET_ID = '17pgg3mwruTWm_Zj5cmyXXEfjKuihu8RIcD1JviSaKf0';
const APP_VERSION = 'finanzas-familia-write-v7';
const API_TOKEN = '05b58aae476517f01223cb670ad12d0e26c9a99ae2fef405';

const SHEETS = {
  gastos: {
    name: 'Gastos',
    headers: ['ID', 'Fecha', 'Día', 'Mes', 'Año', 'Descripción', 'Monto', 'Tipo', 'Categoría', 'Origen']
  },
  ingresos: {
    name: 'Ingresos',
    headers: ['ID', 'Fecha', 'Día', 'Mes', 'Año', 'Descripción', 'Quién', 'Monto', 'Origen', 'RecibidoMes']
  },
  recordatorios: {
    name: 'Recordatorios',
    headers: ['ID', 'Fecha', 'Descripción', 'Monto', 'Día', 'Frecuencia', 'Notas', 'Origen', 'ApartadoMes', 'PagadoMes', 'GastoPagadoId']
  },
  abonosIngresos: {
    name: 'AbonosIngresos',
    headers: ['ID', 'IngresoID', 'Fecha', 'Día', 'Mes', 'Año', 'Monto', 'Nota', 'Origen']
  }
};

function setup() {
  const ss = getSpreadsheet();
  ensureSheets(ss);
  return jsonOutput({ ok: true, version: APP_VERSION, message: 'Hojas y columnas listas' });
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'health') {
    return jsonOutput({ ok: true, version: APP_VERSION, writes: true });
  }

  if (action === 'getAll' || action === 'add' || action === 'update' || action === 'delete') {
    if (!isAuthorized(e.parameter || {})) {
      return jsonOutput({ ok: false, version: APP_VERSION, error: 'No autorizado' });
    }
  }

  if (action === 'getAll') {
    const ss = getSpreadsheet();
    ensureSheets(ss);

    return jsonOutput({
      ok: true,
      version: APP_VERSION,
      gastos: leerGastos(ss.getSheetByName(SHEETS.gastos.name)),
      ingresos: leerIngresos(ss.getSheetByName(SHEETS.ingresos.name)),
      recordatorios: leerRecordatorios(ss.getSheetByName(SHEETS.recordatorios.name)),
      abonosIngresos: leerAbonosIngresos(ss.getSheetByName(SHEETS.abonosIngresos.name))
    });
  }

  if (action === 'add' || action === 'update' || action === 'delete') {
    return ejecutarAccion(e.parameter || {});
  }

  return jsonOutput({ ok: false, version: APP_VERSION, error: 'Accion no reconocida' });
}

function doPost(e) {
  let data = {};
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (error) {
    return jsonOutput({ ok: false, version: APP_VERSION, error: 'JSON invalido' });
  }
  if (!isAuthorized(data)) {
    return jsonOutput({ ok: false, version: APP_VERSION, error: 'No autorizado' });
  }
  return ejecutarAccion(data);
}

function isAuthorized(data) {
  return String(data.token || '') === API_TOKEN;
}

function ejecutarAccion(data) {
  let result = false;

  if (data.action === 'add') {
    if (data.tipo === 'gasto') {
      agregarGasto(data);
      result = true;
    }
    if (data.tipo === 'ingreso') {
      agregarIngreso(data);
      result = true;
    }
    if (data.tipo === 'recordatorio') {
      agregarRecordatorio(data);
      result = true;
    }
    if (data.tipo === 'abonoIngreso') {
      agregarAbonoIngreso(data);
      result = true;
    }
  }

  if (data.action === 'delete') {
    result = borrarRegistro(data.tipo, data.id);
  }

  if (data.action === 'update') {
    result = actualizarRegistro(data.tipo, data.id, data);
  }

  return jsonOutput({
    ok: result !== false,
    version: APP_VERSION,
    action: data.action || '',
    tipo: data.tipo || '',
    id: data.id || '',
    result: result
  });
}

function agregarGasto(data) {
  const ss = getSpreadsheet();
  ensureSheets(ss);

  const fecha = new Date();
  ss.getSheetByName(SHEETS.gastos.name).appendRow([
    data.id || fecha.toISOString(),
    data.fecha || fecha,
    numberOr(data.dia, fecha.getDate()),
    numberOr(data.mes, fecha.getMonth()),
    numberOr(data.anio || data.ano, fecha.getFullYear()),
    safeSheetText(data.desc || data.descripcion || ''),
    numberOr(data.monto, 0),
    safeSheetText(data.tipogasto || data.tipoGasto || data.clase || ''),
    safeSheetText(data.cat || data.categoria || ''),
    safeSheetText(data.origen || 'telegram')
  ]);
}

function agregarIngreso(data) {
  const ss = getSpreadsheet();
  ensureSheets(ss);

  const fecha = new Date();
  ss.getSheetByName(SHEETS.ingresos.name).appendRow([
    data.id || fecha.toISOString(),
    data.fecha || fecha,
    numberOr(data.dia, fecha.getDate()),
    numberOr(data.mes, fecha.getMonth()),
    numberOr(data.anio || data.ano, fecha.getFullYear()),
    safeSheetText(data.desc || data.descripcion || ''),
    safeSheetText(data.quien || data['quién'] || ''),
    numberOr(data.monto, 0),
    safeSheetText(data.origen || 'telegram'),
    normalizeMonthMarker(data.recibidoMes)
  ]);
}

function agregarRecordatorio(data) {
  const ss = getSpreadsheet();
  ensureSheets(ss);

  const fecha = new Date();
  ss.getSheetByName(SHEETS.recordatorios.name).appendRow([
    data.id || fecha.toISOString(),
    data.fecha || fecha,
    safeSheetText(data.desc || data.descripcion || ''),
    numberOr(data.monto, 0),
    numberOr(data.dia, fecha.getDate()),
    safeSheetText(data.frecuencia || data.tipoRec || data.tipo_recordatorio || 'mensual'),
    safeSheetText(data.notas || ''),
    safeSheetText(data.origen || 'telegram'),
    data.apartadoMes || '',
    data.pagadoMes || '',
    data.gastoPagadoId || ''
  ]);
}

function agregarAbonoIngreso(data) {
  const ss = getSpreadsheet();
  ensureSheets(ss);

  const fecha = new Date();
  ss.getSheetByName(SHEETS.abonosIngresos.name).appendRow([
    data.id || fecha.toISOString(),
    data.ingresoId || data.ingresoID || '',
    data.fecha || fecha,
    numberOr(data.dia, fecha.getDate()),
    numberOr(data.mes, fecha.getMonth()),
    numberOr(data.anio || data.ano, fecha.getFullYear()),
    numberOr(data.monto, 0),
    safeSheetText(data.nota || data.notas || ''),
    safeSheetText(data.origen || 'manual')
  ]);
}

function actualizarRegistro(tipo, id, data) {
  if (!id) return false;

  const ss = getSpreadsheet();
  ensureSheets(ss);

  const sheetName = getSheetNameByTipo(tipo);
  if (!sheetName) return false;

  const sheet = ss.getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return false;

  const headers = values[0].map(String);
  const idText = String(id);

  for (let row = values.length - 1; row >= 1; row--) {
    if (String(values[row][0]) === idText) {
      actualizarCamposPorTipo(sheet, headers, row + 1, tipo, data);
      return true;
    }
  }

  return false;
}

function actualizarCamposPorTipo(sheet, headers, row, tipo, data) {
  if (tipo === 'gasto') {
    updateCellIfPresent(sheet, headers, row, 'Fecha', firstDefined(data.fecha));
    updateCellIfPresent(sheet, headers, row, 'Día', firstDefined(data.dia));
    updateCellIfPresent(sheet, headers, row, 'Mes', firstDefined(data.mes));
    updateCellIfPresent(sheet, headers, row, 'Año', firstDefined(data.anio, data.ano));
    updateCellIfPresent(sheet, headers, row, 'Descripción', firstDefined(data.desc, data.descripcion));
    updateCellIfPresent(sheet, headers, row, 'Monto', firstDefined(data.monto));
    updateCellIfPresent(sheet, headers, row, 'Tipo', firstDefined(data.tipogasto, data.tipoGasto, data.clase));
    updateCellIfPresent(sheet, headers, row, 'Categoría', firstDefined(data.cat, data.categoria));
    updateCellIfPresent(sheet, headers, row, 'Origen', firstDefined(data.origen));
    return;
  }

  if (tipo === 'ingreso') {
    updateCellIfPresent(sheet, headers, row, 'Fecha', firstDefined(data.fecha));
    updateCellIfPresent(sheet, headers, row, 'Día', firstDefined(data.dia));
    updateCellIfPresent(sheet, headers, row, 'Mes', firstDefined(data.mes));
    updateCellIfPresent(sheet, headers, row, 'Año', firstDefined(data.anio, data.ano));
    updateCellIfPresent(sheet, headers, row, 'Descripción', firstDefined(data.desc, data.descripcion));
    updateCellIfPresent(sheet, headers, row, 'Quién', firstDefined(data.quien, data['quién']));
    updateCellIfPresent(sheet, headers, row, 'Monto', firstDefined(data.monto));
    updateCellIfPresent(sheet, headers, row, 'Origen', firstDefined(data.origen));
    updateCellIfPresent(sheet, headers, row, 'RecibidoMes', firstDefined(data.recibidoMes));
    return;
  }

  if (tipo === 'recordatorio') {
    updateCellIfPresent(sheet, headers, row, 'Fecha', firstDefined(data.fecha));
    updateCellIfPresent(sheet, headers, row, 'Descripción', firstDefined(data.desc, data.descripcion));
    updateCellIfPresent(sheet, headers, row, 'Monto', firstDefined(data.monto));
    updateCellIfPresent(sheet, headers, row, 'Día', firstDefined(data.dia));
    updateCellIfPresent(sheet, headers, row, 'Frecuencia', firstDefined(data.frecuencia, data.tipoRec, data.tipo_recordatorio));
    updateCellIfPresent(sheet, headers, row, 'Notas', firstDefined(data.notas, data.nota));
    updateCellIfPresent(sheet, headers, row, 'Origen', firstDefined(data.origen));
    updateCellIfPresent(sheet, headers, row, 'ApartadoMes', firstDefined(data.apartadoMes));
    updateCellIfPresent(sheet, headers, row, 'PagadoMes', firstDefined(data.pagadoMes));
    updateCellIfPresent(sheet, headers, row, 'GastoPagadoId', firstDefined(data.gastoPagadoId));
    return;
  }

  if (tipo === 'abonoIngreso') {
    updateCellIfPresent(sheet, headers, row, 'IngresoID', firstDefined(data.ingresoId, data.ingresoID));
    updateCellIfPresent(sheet, headers, row, 'Fecha', firstDefined(data.fecha));
    updateCellIfPresent(sheet, headers, row, 'Día', firstDefined(data.dia));
    updateCellIfPresent(sheet, headers, row, 'Mes', firstDefined(data.mes));
    updateCellIfPresent(sheet, headers, row, 'Año', firstDefined(data.anio, data.ano));
    updateCellIfPresent(sheet, headers, row, 'Monto', firstDefined(data.monto));
    updateCellIfPresent(sheet, headers, row, 'Nota', firstDefined(data.nota, data.notas));
    updateCellIfPresent(sheet, headers, row, 'Origen', firstDefined(data.origen));
  }
}

function updateCellIfPresent(sheet, headers, row, header, value) {
  const index = headers.indexOf(header);
  if (index >= 0 && value !== undefined) {
    const cell = sheet.getRange(row, index + 1);
    if (header === 'ApartadoMes' || header === 'PagadoMes' || header === 'GastoPagadoId' || header === 'RecibidoMes') {
      cell.setNumberFormat('@');
      cell.setValue(String(value || ''));
      return;
    }
    cell.setValue(safeSheetValue(header, value));
  }
}

function safeSheetValue(header, value) {
  const textHeaders = ['Descripción', 'Tipo', 'Categoría', 'Origen', 'Quién', 'Frecuencia', 'Notas', 'Nota', 'IngresoID'];
  return textHeaders.indexOf(header) >= 0 ? safeSheetText(value) : value;
}

function safeSheetText(value) {
  let text = String(value || '').slice(0, 300);
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function firstDefined() {
  for (let i = 0; i < arguments.length; i++) {
    if (arguments[i] !== undefined) return arguments[i];
  }
  return undefined;
}

function borrarRegistro(tipo, id) {
  if (!id) return false;

  const ss = getSpreadsheet();
  ensureSheets(ss);

  const sheetName = getSheetNameByTipo(tipo);
  if (!sheetName) return false;

  const sheet = ss.getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const idText = String(id);

  for (let row = values.length - 1; row >= 1; row--) {
    if (String(values[row][0]) === idText) {
      sheet.deleteRow(row + 1);
      return true;
    }
  }

  return false;
}

function getSheetNameByTipo(tipo) {
  if (tipo === 'gasto') return SHEETS.gastos.name;
  if (tipo === 'ingreso') return SHEETS.ingresos.name;
  if (tipo === 'recordatorio') return SHEETS.recordatorios.name;
  if (tipo === 'abonoIngreso') return SHEETS.abonosIngresos.name;
  return '';
}

function leerGastos(sheet) {
  return readRows(sheet).map(row => ({
    id: row.ID || row.Id,
    fecha: row.Fecha,
    dia: numberOr(row.Dia || row['Día'], null),
    mes: numberOr(row.Mes, null),
    anio: numberOr(row.Ano || row['Año'], null),
    desc: row.Descripcion || row['Descripción'] || '',
    monto: numberOr(row.Monto, 0),
    tipo: row.Tipo || '',
    cat: row.Categoria || row['Categoría'] || '',
    origen: row.Origen || ''
  }));
}

function leerIngresos(sheet) {
  return readRows(sheet).map(row => ({
    id: row.ID || row.Id,
    fecha: row.Fecha,
    dia: numberOr(row.Dia || row['Día'], null),
    mes: numberOr(row.Mes, null),
    anio: numberOr(row.Ano || row['Año'], null),
    desc: row.Descripcion || row['Descripción'] || '',
    quien: row.Quien || row['Quién'] || '',
    monto: numberOr(row.Monto, 0),
    origen: row.Origen || '',
    recibidoMes: normalizeMonthMarker(row.RecibidoMes)
  }));
}

function leerRecordatorios(sheet) {
  return readRows(sheet).map(row => ({
    id: row.ID || row.Id,
    fecha: row.Fecha,
    desc: row.Descripcion || row['Descripción'] || '',
    monto: numberOr(row.Monto, 0),
    dia: numberOr(row.Dia || row['Día'], null),
    frecuencia: row.Frecuencia || 'mensual',
    tipo: row.Frecuencia || 'mensual',
    notas: row.Notas || '',
    origen: row.Origen || '',
    apartadoMes: normalizeMonthMarker(row.ApartadoMes),
    pagadoMes: normalizeMonthMarker(row.PagadoMes),
    gastoPagadoId: row.GastoPagadoId || ''
  }));
}

function leerAbonosIngresos(sheet) {
  return readRows(sheet).map(row => ({
    id: row.ID || row.Id,
    ingresoId: row.IngresoID || row.IngresoId || row.ingresoId || '',
    fecha: row.Fecha,
    dia: numberOr(row.Dia || row['Día'], null),
    mes: numberOr(row.Mes, null),
    anio: numberOr(row.Ano || row['Año'], null),
    monto: numberOr(row.Monto, 0),
    nota: row.Nota || row.Notas || '',
    origen: row.Origen || ''
  }));
}

function normalizeMonthMarker(value) {
  if (!value) return '';
  const text = String(value);
  const match = text.match(/^(\d{4})-(\d{2})/);
  return match ? match[1] + '-' + match[2] : text;
}

function readRows(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(cell => cell !== '')).map(row => {
    const item = {};
    headers.forEach((header, index) => {
      const value = row[index];
      item[header] = value instanceof Date ? value.toISOString() : value;
    });
    return item;
  });
}

function ensureSheets(ss) {
  Object.keys(SHEETS).forEach(key => {
    const config = SHEETS[key];
    let sheet = ss.getSheetByName(config.name);

    if (!sheet) sheet = ss.insertSheet(config.name);
    if (sheet.getLastRow() === 0) sheet.appendRow(config.headers);
    ensureHeaders(sheet, config.headers);
  });
}

function ensureHeaders(sheet, expectedHeaders) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);

  expectedHeaders.forEach(header => {
    if (!currentHeaders.includes(header)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      currentHeaders.push(header);
    }
  });
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function numberOr(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback;
  const normalized = String(value).replace(/[$,\s]/g, '');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : fallback;
}
