// ═══════════════════════════════════════════════════════════════════
// СС Шахматка — Google Apps Script
// Таблица: https://docs.google.com/spreadsheets/d/1n0oy9wt8sLOXY6lJNkeHaChQJJiWAyBKE1872ALeBbQ
// ═══════════════════════════════════════════════════════════════════

const FACT_SHEET      = 'Факт';
const REF_SHEET       = 'Справочник работ';
const FLOORS_SHEET    = 'Реестр этажей';
const ADMIN_PASSWORD  = 'adminCC';

// Справочник работ — столбцы (0-индекс для массива)
// A(0)=ID_работы  B(1)=Раздел  C(2)=Система  D(3)=Работа  E(4)=Локация
// F(5)=Захватка   G(6)=Группировка  H(7)=Комментарий  I(8)=Название для шахматки

// Реестр этажей — столбцы (0-индекс)
// A(0)=ID_этажа  B(1)=Очередь  C(2)=Корпус_числом  D(3)=Корпус_текстом  E(4)=Этаж

// Факт — столбцы (1-индекс для GAS range; 0-индекс для массива)
// A(1/0)=ID_работы  B(2/1)=ID_этажа  C(3/2)=Система  D(4/3)=Корпус_текстом  E(5/4)=Этаж
// F(6/5)=Статус  G(7/6)=Дата готовности  H(8/7)=Процент готовности
// ⚠️ ПИСАТЬ ТОЛЬКО В F–H (столбцы 6–8). A–E никогда не трогать.

var FACT_EDIT_START = 6; // столбец F
var FACT_EDIT_COLS  = 3; // F, G, H

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var action = p.action || '';

    if (action === 'getData')       return jsonOut(getData());
    if (action === 'ping')          return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    if (action === 'checkPassword') return jsonOut({ ok: p.pwd === ADMIN_PASSWORD });
    if (action === 'clearCache')    { clearCache(); return jsonOut({ ok: true }); }
    if (action === 'saveCell') {
      saveCellFact(p.workId, p.floorId, p.undo === 'true');
      return jsonOut({ ok: true });
    }


    return HtmlService.createHtmlOutputFromFile('index.html')
      .setTitle('СС · Шахматка')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return jsonOut({ error: err.toString() });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'saveCell') {
      saveCellFact(body.workId, body.floorId, !!body.undo);
      return jsonOut({ ok: true });
    }
    return jsonOut({ error: 'Unknown action' });
  } catch (err) {
    return jsonOut({ error: err.toString() });
  }
}

// ─── Основные данные ────────────────────────────────────────────────────────

function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    works  : readWorks(ss),
    floors : readFloors(ss),
    facts  : readFacts(ss),
  };
}

function readWorks(ss) {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get('cc_works');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  var sheet = ss.getSheetByName(REF_SHEET);
  var works = [];
  if (sheet) {
    var last = sheet.getLastRow();
    if (last >= 2) {
      var vals = sheet.getRange(2, 1, last - 1, 10).getValues();
      vals.forEach(function (row) {
        var id = String(row[0]).trim();
        if (!id) return;
        works.push({
          workId   : id,
          razdel   : String(row[1]).trim(),
          sistema  : String(row[2]).trim(),
          work     : String(row[3]).trim(),
          lokacia  : String(row[4]).trim(),
          zahvatka : String(row[5]).trim(),
          group    : String(row[6]).trim(),
          nameGrid : String(row[8]).trim() || String(row[3]).trim(),
        });
      });
    }
  }
  try { cache.put('cc_works', JSON.stringify(works), 3600); } catch (e) {}
  return works;
}

function readFloors(ss) {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get('cc_floors');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  var sheet  = ss.getSheetByName(FLOORS_SHEET);
  var floors = [];
  if (sheet) {
    var last = sheet.getLastRow();
    if (last >= 2) {
      var vals = sheet.getRange(2, 1, last - 1, 5).getValues();
      vals.forEach(function (row) {
        var id = String(row[0]).trim();
        if (!id) return;
        var floorNum = parseFloat(String(row[4]).replace(',', '.'));
        if (isNaN(floorNum)) return;
        floors.push({
          floorId : id,
          corpus  : String(row[3]).trim(),
          floor   : floorNum,
        });
      });
    }
  }
  try { cache.put('cc_floors', JSON.stringify(floors), 3600); } catch (e) {}
  return floors;
}

function readFacts(ss) {
  var sheet = ss.getSheetByName(FACT_SHEET);
  var facts = [];
  if (!sheet) return facts;

  var last = sheet.getLastRow();
  if (last < 2) return facts;

  var vals = sheet.getRange(2, 1, last - 1, 8).getValues();
  vals.forEach(function (row) {
    var workId  = String(row[0]).trim();
    var floorId = String(row[1]).trim();
    if (!workId || !floorId) return;
    var status = String(row[5]).trim();
    var date   = formatDateOut(row[6]);
    var pct    = String(row[7]).trim();
    if (!status && !date && !pct) return; // пустые строки не передаём
    facts.push({ workId: workId, floorId: floorId, status: status, date: date, pct: pct });
  });
  return facts;
}

// ─── Сохранение ─────────────────────────────────────────────────────────────

function saveCellFact(workId, floorId, undo) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(FACT_SHEET);
  if (!sheet) throw new Error('Лист «Факт» не найден');

  var last = sheet.getLastRow();
  if (last < 2) throw new Error('Лист «Факт» пуст');

  // Ищем строку по A=workId и B=floorId — никогда не трогаем A–E
  var keys = sheet.getRange(2, 1, last - 1, 2).getValues();
  var targetRow = -1;
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === String(workId) &&
        String(keys[i][1]).trim() === String(floorId)) {
      targetRow = i + 2; // +1 заголовок, +1 смещение 0→1
      break;
    }
  }
  if (targetRow < 0) throw new Error('Строка не найдена: ' + workId + ' / ' + floorId);

  if (undo) {
    sheet.getRange(targetRow, FACT_EDIT_START, 1, FACT_EDIT_COLS).setValues([['', '', '']]);
  } else {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    sheet.getRange(targetRow, FACT_EDIT_START, 1, FACT_EDIT_COLS)
         .setValues([['СМР окончены', today, true]]);
  }
  SpreadsheetApp.flush();
}

// ─── Утилиты ────────────────────────────────────────────────────────────────

function clearCache() {
  try {
    var c = CacheService.getScriptCache();
    c.remove('cc_works');
    c.remove('cc_floors');
  } catch (e) {}
}

function formatDateOut(v) {
  if (!v) return '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return pad(v.getDate()) + '.' + pad(v.getMonth() + 1) + '.' + v.getFullYear();
  }
  return String(v).trim();
}

function pad(n) { return n < 10 ? '0' + n : String(n); }
