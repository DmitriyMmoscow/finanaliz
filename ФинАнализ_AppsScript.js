// ============================================================
//  ФинАнализ — Google Apps Script
//  Вставить в: script.google.com → Новый проект
//  Опубликовать: Развернуть → Новое развёртывание →
//    Тип: Веб-приложение
//    Выполнять как: Я
//    Доступ: Все (в т.ч. анонимные)
//  Скопировать URL и вставить в Настройки → ФинАнализ
// ============================================================

const SHEET_NAME_TX   = 'Транзакции';
const SHEET_NAME_DIR  = 'Справочник';
const SHEET_NAME_CATS = 'СтатьиДДС';
const SHEET_NAME_META = 'Мета';

// ── Заголовки листов ──────────────────────────────────────
const TX_HEADERS  = ['id','date','counterparty','purpose','amount','type','category','project','account','uploadName','description'];
const DIR_HEADERS = ['counterparty','inn','category'];
const CAT_HEADERS = ['id','name','type','keywords'];

// ============================================================
//  CORS — нужен для fetch из браузера
// ============================================================
function setCORS(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// OPTIONS preflight
function doOptions() {
  return setCORS(ContentService.createTextOutput(''));
}

// ============================================================
//  GET — чтение данных
//  ?action=read  → возвращает все данные
//  ?action=ping  → проверка связи
// ============================================================
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'read';

    if (action === 'ping') {
      return setCORS(json({ ok: true, message: 'ФинАнализ AppsScript работает' }));
    }

    if (action === 'read') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      return setCORS(json({
        ok: true,
        transactions: readSheet(ss, SHEET_NAME_TX,  TX_HEADERS),
        directory:    readSheet(ss, SHEET_NAME_DIR,  DIR_HEADERS),
        categories:   readSheet(ss, SHEET_NAME_CATS, CAT_HEADERS),
      }));
    }

    return setCORS(json({ ok: false, error: 'Unknown action: ' + action }));
  } catch(err) {
    return setCORS(json({ ok: false, error: err.toString() }));
  }
}

// ============================================================
//  POST — запись данных
//  body: { action, payload }
//
//  action = 'write'  — полная перезапись (payload = { transactions, directory, categories })
//  action = 'append' — добавить транзакции (payload = { transactions: [...] })
//  action = 'updateTx' — обновить одну транзакцию (payload = { id, fields: {...} })
// ============================================================
function doPost(e) {
  try {
    const body    = JSON.parse(e.postData.contents);
    const action  = body.action || 'write';
    const payload = body.payload || {};
    const ss      = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'write') {
      // Полная перезапись всех данных
      if (payload.transactions !== undefined)
        writeSheet(ss, SHEET_NAME_TX,  TX_HEADERS,  payload.transactions);
      if (payload.directory !== undefined)
        writeSheet(ss, SHEET_NAME_DIR,  DIR_HEADERS, payload.directory);
      if (payload.categories !== undefined)
        writeSheet(ss, SHEET_NAME_CATS, CAT_HEADERS, payload.categories);
      writeMeta(ss, payload.transactions ? payload.transactions.length : null);
      return setCORS(json({ ok: true, action: 'write' }));
    }

    if (action === 'append') {
      // Добавить новые транзакции (без дублей по id)
      const newTxs = payload.transactions || [];
      const sheet  = getOrCreateSheet(ss, SHEET_NAME_TX, TX_HEADERS);
      const existingIds = getExistingIds(sheet);
      const toAdd = newTxs.filter(tx => !existingIds.has(String(tx.id)));
      if (toAdd.length) {
        const rows = toAdd.map(tx => TX_HEADERS.map(h => tx[h] !== undefined ? tx[h] : ''));
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, TX_HEADERS.length).setValues(rows);
      }
      return setCORS(json({ ok: true, action: 'append', added: toAdd.length, skipped: newTxs.length - toAdd.length }));
    }

    if (action === 'updateTx') {
      // Обновить поля одной транзакции по id
      const txId  = String(payload.id);
      const sheet = getOrCreateSheet(ss, SHEET_NAME_TX, TX_HEADERS);
      const data  = sheet.getDataRange().getValues();
      const idCol = TX_HEADERS.indexOf('id');
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][idCol]) === txId) {
          Object.entries(payload.fields || {}).forEach(([key, val]) => {
            const col = TX_HEADERS.indexOf(key);
            if (col >= 0) sheet.getRange(r + 1, col + 1).setValue(val);
          });
          return setCORS(json({ ok: true, action: 'updateTx', updated: txId }));
        }
      }
      return setCORS(json({ ok: false, error: 'Transaction not found: ' + txId }));
    }

    return setCORS(json({ ok: false, error: 'Unknown action: ' + action }));
  } catch(err) {
    return setCORS(json({ ok: false, error: err.toString() }));
  }
}

// ============================================================
//  Вспомогательные функции
// ============================================================

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readSheet(ss, name, headers) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const hdr = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    hdr.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function writeSheet(ss, name, headers, rows) {
  const sheet = getOrCreateSheet(ss, name, headers);
  // Очищаем данные (не заголовок)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  if (!rows || !rows.length) return;
  const values = rows.map(row => headers.map(h => {
    const v = row[h];
    if (v === null || v === undefined) return '';
    if (Array.isArray(v)) return v.join(', ');
    return v;
  }));
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function getExistingIds(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return new Set();
  const idCol = data[0].indexOf('id');
  if (idCol < 0) return new Set();
  return new Set(data.slice(1).map(r => String(r[idCol])));
}

function writeMeta(ss, txCount) {
  const sheet = getOrCreateSheet(ss, SHEET_NAME_META, ['key','value']);
  sheet.clearContents();
  sheet.getRange(1,1,1,2).setValues([['key','value']]);
  sheet.getRange(2,1,3,2).setValues([
    ['last_sync', new Date().toISOString()],
    ['tx_count',  txCount || 0],
    ['version',   '1.0'],
  ]);
}
