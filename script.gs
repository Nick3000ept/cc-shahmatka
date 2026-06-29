// ═══════════════════════════════════════════════════════════════════
// СС Шахматка — Google Apps Script
// Таблица: https://docs.google.com/spreadsheets/d/1n0oy9wt8sLOXY6lJNkeHaChQJJiWAyBKE1872ALeBbQ
// ═══════════════════════════════════════════════════════════════════

const FACT_SHEET      = 'Факт';
const REF_SHEET       = 'Справочник работ';
const FLOORS_SHEET    = 'Реестр этажей';
const CHK_HIER_SHEET  = 'Чек листы иерархия';
const CHK_SHEET       = 'Чек листы';
const ADMIN_PASSWORD  = 'adminCC';

// Справочник работ — столбцы (0-индекс для массива)
// A(0)=ID_работы  B(1)=Раздел  C(2)=Система  D(3)=Работа  E(4)=Локация
// F(5)=Захватка   G(6)=Группировка  H(7)=Комментарий  I(8)=Название для шахматки
// J(9)=Система форма КП  K(10)=ИД чек лист (код иерархии L2, связь с чек-листами)
//
// Чек листы иерархия — столбцы (0-индекс): C(2)=L1 код  D(3)=L2 код (= Справочник.K)
//   E(4)=Раздел  F(5)=Подраздел  G(6)=Работа  (имена для матчинга с листом «Чек листы»)
//
// Чек листы — столбцы (0-индекс): A(0)=ID  C(2)=Корпус  D(3)=Этажи
//   E(4)=Раздел  F(5)=Подраздел  G(6)=Работа  H(7)=Номер акта  I(8)=Дата  J(9)=Статус
//   O(14)=Файлы (S3) URL  P(15)=Google Drive URL

// Реестр этажей — столбцы (0-индекс)
// A(0)=ID_этажа  B(1)=Очередь  C(2)=Корпус_числом  D(3)=Корпус_текстом  E(4)=Этаж

// Факт — столбцы (1-индекс для GAS range; 0-индекс для массива)
// A(1/0)=ID_работы  B(2/1)=ID_этажа  C(3/2)=Система  D(4/3)=Корпус_текстом  E(5/4)=Этаж
// F(6/5)=Статус  G(7/6)=Дата готовности  H(8/7)=Процент готовности  I(9/8)=Доп признак (подъезд)
// ⚠️ ПИСАТЬ ТОЛЬКО В F–H (столбцы 6–8). A–E и I никогда не трогать.

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
  var factData = readFacts(ss);
  var chk = readChecklists(ss);
  return {
    works      : readWorks(ss),
    floors     : readFloors(ss),
    facts      : factData.facts,
    floorPod   : factData.floorPod,   // floorId → подъезд (Доп признак)
    checkByCell: chk.checkByCell,      // `workId\x00corpus\x00floor` → [{link,status,akt,date,id}]
    checkStats : chk.stats,            // диагностика совпадений (для проверки связей)
  };
}

function readWorks(ss) {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get('cc_works_v2');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  var sheet = ss.getSheetByName(REF_SHEET);
  var works = [];
  if (sheet) {
    var last = sheet.getLastRow();
    if (last >= 2) {
      var vals = sheet.getRange(2, 1, last - 1, 11).getValues();
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
          checkId  : String(row[10]).trim(), // K — ИД чек лист (код иерархии L2)
        });
      });
    }
  }
  try { cache.put('cc_works_v2', JSON.stringify(works), 3600); } catch (e) {}
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
  var floorPod = {};             // floorId → подъезд (Доп признак), свойство этажа
  if (!sheet) return { facts: facts, floorPod: floorPod };

  var last = sheet.getLastRow();
  if (last < 2) return { facts: facts, floorPod: floorPod };

  var vals = sheet.getRange(2, 1, last - 1, 9).getValues();
  vals.forEach(function (row) {
    var workId  = String(row[0]).trim();
    var floorId = String(row[1]).trim();
    if (!workId || !floorId) return;

    // Доп признак (подъезд) — столбец I (index 8). Один и тот же для всех строк этажа.
    var pod = String(row[8]).trim();
    if (pod && !floorPod[floorId]) floorPod[floorId] = pod;

    var status = String(row[5]).trim();
    var date   = formatDateOut(row[6]);
    var pct    = String(row[7]).trim();
    if (!status && !date && !pct) return; // пустые строки факта не передаём
    facts.push({ workId: workId, floorId: floorId, status: status, date: date, pct: pct });
  });
  return { facts: facts, floorPod: floorPod };
}

// ─── Связь работ с чек-листами ───────────────────────────────────────────────
// Цепочка: Справочник.K (ИД чек лист) == иерархия.D (код L2)
//   код → тройка имён (Раздел, Подраздел, Работа) из иерархии
//   «Чек листы» матчатся по той же тройке + Корпус + Этаж → ссылка Файлы (S3)
// Результат: checkByCell[`workId\x00corpus\x00floor`] = [{link,status,akt,date,id}]
function readChecklists(ss) {
  // 1) Иерархия: код L2 (D=3) → ключ тройки имён (E=4 Раздел, F=5 Подраздел, G=6 Работа)
  var codeToTriple = {};
  var h = ss.getSheetByName(CHK_HIER_SHEET);
  if (h && h.getLastRow() >= 2) {
    h.getRange(2, 1, h.getLastRow() - 1, 7).getValues().forEach(function (r) {
      var code = String(r[3]).trim();
      if (code) codeToTriple[code] = tripleKey(r[4], r[5], r[6]);
    });
  }

  // 2) Справочник: тройка имён → [workId] (через присвоенный код K=10)
  var tripleToWorks = {};
  var codeMissing = 0, worksWithCode = 0;
  var s = ss.getSheetByName(REF_SHEET);
  if (s && s.getLastRow() >= 2) {
    s.getRange(2, 1, s.getLastRow() - 1, 11).getValues().forEach(function (r) {
      var workId = String(r[0]).trim();
      var code   = String(r[10]).trim();
      if (!workId || !code) return;
      worksWithCode++;
      var tk = codeToTriple[code];
      if (!tk) { codeMissing++; return; } // код из K не найден в иерархии
      (tripleToWorks[tk] = tripleToWorks[tk] || []).push(workId);
    });
  }

  // 3) Чек листы: матч по тройке → раскидываем по ячейкам (workId, корпус, этаж)
  var checkByCell = {};
  var rows = 0, matchedRows = 0, cellLinks = 0;
  var unmatched = {};
  var c = ss.getSheetByName(CHK_SHEET);
  if (c && c.getLastRow() >= 2) {
    c.getRange(2, 1, c.getLastRow() - 1, 16).getValues().forEach(function (r) {
      rows++;
      var tk = tripleKey(r[4], r[5], r[6]);
      var works = tripleToWorks[tk];
      if (!works || !works.length) {
        var lbl = String(r[4]).trim() + ' || ' + String(r[5]).trim() + ' || ' + String(r[6]).trim();
        unmatched[lbl] = (unmatched[lbl] || 0) + 1;
        return;
      }
      matchedRows++;
      var corpus = String(r[2]).trim();
      var floor  = floorNorm(r[3]);
      var rec = {
        id    : r[0],
        status: String(r[9]).trim(),
        akt   : String(r[7]).trim(),
        date  : formatDateOut(r[8]),
        link  : String(r[14]).trim() || String(r[15]).trim(),
      };
      works.forEach(function (workId) {
        var key = workId + '\x00' + corpus + '\x00' + floor;
        (checkByCell[key] = checkByCell[key] || []).push(rec);
        cellLinks++;
      });
    });
  }

  // Топ-10 несовпавших троек — для диагностики словаря
  var unmatchedTop = Object.keys(unmatched)
    .sort(function (a, b) { return unmatched[b] - unmatched[a]; })
    .slice(0, 10)
    .map(function (k) { return k + '  (' + unmatched[k] + ')'; });

  return {
    checkByCell: checkByCell,
    stats: {
      hierCodes        : Object.keys(codeToTriple).length,
      worksWithCode    : worksWithCode,
      codeMissingInHier: codeMissing,
      tripleKeys       : Object.keys(tripleToWorks).length,
      cheklistRows     : rows,
      cheklistMatched  : matchedRows,
      cellsLinked      : Object.keys(checkByCell).length,
      cellLinkPairs    : cellLinks,
      unmatchedTop     : unmatchedTop,
    },
  };
}

// Ключ тройки имён (нормализованный): раздел \x00 подраздел \x00 работа
function tripleKey(razdel, podrazdel, rabota) {
  return normName(razdel) + '\x00' + normName(podrazdel) + '\x00' + normName(rabota);
}
function normName(s) {
  return String(s).trim().toLowerCase().replace(/\s+/g, ' ');
}
// Нормализация этажа: "2"→"2", "1,1"→"1.1", "-1"→"-1"; нечисловое — как есть
function floorNorm(v) {
  var s = String(v).trim().replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? s : String(n);
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
    c.remove('cc_works_v2');
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
