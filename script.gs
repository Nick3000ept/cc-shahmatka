// ═══════════════════════════════════════════════════════════════════
// СС Шахматка — Google Apps Script
// Таблица: https://docs.google.com/spreadsheets/d/1n0oy9wt8sLOXY6lJNkeHaChQJJiWAyBKE1872ALeBbQ
// ═══════════════════════════════════════════════════════════════════

const FACT_SHEET      = 'Факт';
const REF_SHEET       = 'Справочник работ';
const FLOORS_SHEET    = 'Реестр этажей';
const CHK_HIER_SHEET  = 'Чек листы иерархия';
const CHK_SHEET       = 'Чек листы';
const CHK_INT_SHEET   = 'Чек листы_внутренние';
const CHK_FOLDER_NAME = 'Чек листы СС';
const INCLUDE_EXTERNAL = false; // внешние чек-листы (лист «Чек листы» + иерархия) отключены; true — вернуть
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
    if (action === 'getChecklists') return jsonOut(getChecklists());
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
    if (body.action === 'addChecklist') {
      return jsonOut(saveInternalChecklist(body));
    }
    if (body.action === 'setChecklistStatus') {
      return jsonOut(setInternalChecklistStatus(body));
    }
    if (body.action === 'deleteChecklist') {
      return jsonOut(deleteInternalChecklist(body));
    }
    if (body.action === 'addChecklistFile') {
      return jsonOut(addInternalChecklistFile(body));
    }
    return jsonOut({ error: 'Unknown action' });
  } catch (err) {
    return jsonOut({ error: err.toString() });
  }
}

// ─── Основные данные ────────────────────────────────────────────────────────

// Быстрый эндпоинт: сетка (работы/этажи/факт). Чек-листы грузятся отдельно (getChecklists),
// т.к. их матчинг и JSON тяжёлые — иначе загрузка сетки блокируется на ~12с.
function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var factData = readFacts(ss);
  return {
    works   : readWorks(ss),
    floors  : readFloors(ss),
    facts   : factData.facts,
    floorPod: factData.floorPod,   // floorId → подъезд (Доп признак)
  };
}

// Отдельный (тяжёлый) эндпоинт: связь ячеек с чек-листами. Грузится фоном.
function getChecklists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var checkByCell, stats;
  if (INCLUDE_EXTERNAL) {
    var chk = readChecklists(ss);
    checkByCell = chk.checkByCell;
    stats = chk.stats;
  } else {
    checkByCell = {};             // внешние отключены
    stats = { external: 'off' };
  }
  addInternalChecklists(ss, checkByCell); // внутренние чек-листы (загруженные через сайт)
  return {
    checkByCell: checkByCell, // `workId\x00corpus\x00floor` → [{...,id,source}]
    checkStats : stats,
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
        source: 'external',
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

// Внутренние чек-листы (лист «Чек листы_внутренние»): прямой ID_работы + корпус + этаж.
// Дописываем в тот же checkByCell, что и внешние.
// Колонки: A ID  B Дата  C Корпус  D ID_работы  E Этаж  F Статус  G Ссылка  H Файл  I Комментарий  J Номер
//          K Доп файлы (JSON-массив [{n:имя, u:ссылка}] — догруженные фотографии)
//          L Загрузил (имя из формы)  M Изменено (кто · когда · что — последнее изменение)
function addInternalChecklists(ss, checkByCell) {
  var sheet = ss.getSheetByName(CHK_INT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues().forEach(function (r) {
    var workId = String(r[3]).trim();
    var corpus = String(r[2]).trim();
    if (!workId || !corpus) return;
    if (String(r[5]).trim() === 'Удалён') return; // «мягко» удалённые не отдаём
    var extra = [];
    if (r[10]) {
      try { var ex = JSON.parse(String(r[10])); if (ex && ex.length) extra = ex; } catch (e) {}
    }
    var key = workId + '\x00' + corpus + '\x00' + floorNorm(r[4]);
    (checkByCell[key] = checkByCell[key] || []).push({
      id     : r[0],
      status : String(r[5]).trim(),
      akt    : '',
      number : String(r[9]).trim(),
      date   : formatDateOut(r[1]),
      link   : String(r[6]).trim(),
      file   : String(r[7]).trim(),
      comment: String(r[8]).trim(),
      extra  : extra,
      by     : String(r[11]).trim(),
      source : 'internal',
    });
  });
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

// ─── Загрузка внутреннего чек-листа ──────────────────────────────────────────
// body: { corpus, cells:[{workId, floor}], status, comment, fileName, mimeType, fileData(base64) }
// Файл → папка «Чек листы СС» на Диске; по строке на каждую (работа+этаж) в лист.
function saveInternalChecklist(body) {
  var corpus  = String(body.corpus  || '').trim();
  var cells   = body.cells || [];
  var status  = String(body.status  || '').trim();
  var comment = String(body.comment || '').trim();
  var number  = String(body.number  || '').trim();       // номер чек-листа
  var user    = String(body.user    || '').trim();       // кто загрузил (имя из формы)
  var dateStr = normDateIn(body.date);                    // дата чек-листа (dd.MM.yyyy)
  if (!corpus)        throw new Error('Не указан корпус');
  if (!cells.length)  throw new Error('Не выбраны работы');
  if (!body.fileData) throw new Error('Файл не передан');

  // 1) Файл на Google Диск (доступ «по ссылке»), с понятным именем:
  //    СС_К4_эт.2,3_№7_Есть замечания_10.12.2025.pdf
  var floorsSet = {};
  cells.forEach(function (c) { var f = floorNorm(c.floor); if (f) floorsSet[f] = true; });
  var floorsArr = Object.keys(floorsSet).sort(function (a, b) { return parseFloat(a) - parseFloat(b); });
  var ext  = fileExt_(body.fileName, body.mimeType);
  var nice = safeFileName_(
    'СС_' + corpus + '_эт.' + floorsArr.join(',') +
    (number ? '_№' + number : '') +
    '_' + (status || 'Без статуса') + '_' + dateStr
  ) + '.' + ext;

  var folder = getOrCreateChecklistFolder();
  var blob   = Utilities.newBlob(Utilities.base64Decode(body.fileData),
    body.mimeType || 'application/octet-stream', nice);
  var file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var url   = file.getUrl();
  var fname = file.getName();

  // 2) Строки в лист «Чек листы_внутренние» — по одной на (работа+этаж)
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHK_INT_SHEET) || createInternalSheet(ss);
  ensureIntHeaders_(sheet);
  var group = 'IC-' + Date.now();
  var rows  = cells.map(function (c) {
    return [group, dateStr, corpus, String(c.workId).trim(), c.floor, status, url, fname, comment, number, '', user];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();

  return { ok: true, id: group, url: url, rows: rows.length };
}

// Шапки столбцов J–M (лист мог быть создан старой версией с меньшим числом столбцов)
function ensureIntHeaders_(sheet) {
  var want = { 10: 'Номер', 11: 'Доп файлы', 12: 'Загрузил', 13: 'Изменено' };
  for (var col in want) {
    if (String(sheet.getRange(1, Number(col)).getValue()).trim() !== want[col]) {
      sheet.getRange(1, Number(col)).setValue(want[col]);
    }
  }
}

// Имя файла без запрещённых символов, схлопнутые пробелы
function safeFileName_(s) {
  return String(s).replace(/[\\\/:*?"<>|#]/g, '-').replace(/\s+/g, ' ').trim();
}

// Расширение файла: из имени, иначе из mime-типа
function fileExt_(name, mime) {
  var m = String(name || '').match(/\.([A-Za-z0-9]{1,5})$/);
  if (m) return m[1].toLowerCase();
  if (/pdf/i.test(mime))  return 'pdf';
  if (/png/i.test(mime))  return 'png';
  if (/jpe?g/i.test(mime)) return 'jpg';
  if (/webp/i.test(mime)) return 'webp';
  if (/heic/i.test(mime)) return 'heic';
  return 'bin';
}

// Отметка «кто · когда · что» для столбца M «Изменено»
function changeMark_(user, what) {
  var d = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  return (user ? user + ' · ' : '') + d + ' · ' + what;
}

// Сменить статус внутреннего чек-листа (всей группы) — обновляет столбец F во всех строках с этим ID
function setInternalChecklistStatus(body) {
  var id     = String(body.id     || '').trim();
  var status = String(body.status || '').trim();
  var user   = String(body.user   || '').trim();
  if (!id)     throw new Error('Не указан ID чек-листа');
  if (!status) throw new Error('Не указан статус');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHK_INT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Лист «Чек листы_внутренние» пуст');
  ensureIntHeaders_(sheet);

  var n    = sheet.getLastRow() - 1;
  var idA  = sheet.getRange(2, 1, n, 1).getValues();   // столбец A (ID группы)
  var fCol = sheet.getRange(2, 6, n, 1).getValues();   // столбец F (Статус)
  var mCol = sheet.getRange(2, 13, n, 1).getValues();  // столбец M (Изменено)
  var mark = changeMark_(user, 'статус: ' + status);
  var updated = 0;
  for (var i = 0; i < n; i++) {
    if (String(idA[i][0]).trim() === id) { fCol[i][0] = status; mCol[i][0] = mark; updated++; }
  }
  if (!updated) throw new Error('Чек-лист не найден: ' + id);
  sheet.getRange(2, 6, n, 1).setValues(fCol);
  sheet.getRange(2, 13, n, 1).setValues(mCol);
  SpreadsheetApp.flush();
  return { ok: true, updated: updated };
}

// Догрузка файла (фото) к существующему внутреннему чек-листу.
// body: { id, user, fileName, mimeType, fileData(base64) }
// Файл → папка «Чек листы СС» с понятным именем (СС_К4_эт.2,3_№7_фото2.jpg);
// запись [{n:имя, u:ссылка}] дописывается в JSON-массив столбца K «Доп файлы»
// во всех строках группы, столбец M получает отметку «кто · когда · что».
function addInternalChecklistFile(body) {
  var id   = String(body.id   || '').trim();
  var user = String(body.user || '').trim();
  if (!id)            throw new Error('Не указан ID чек-листа');
  if (!body.fileData) throw new Error('Файл не передан');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHK_INT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Лист «Чек листы_внутренние» пуст');

  var n    = sheet.getLastRow() - 1;
  var vals = sheet.getRange(2, 1, n, 11).getValues(); // A–K
  var rowsIdx = [], floorsSet = {}, corpus = '', number = '';
  for (var i = 0; i < n; i++) {
    if (String(vals[i][0]).trim() !== id) continue;
    rowsIdx.push(i + 2);
    if (!corpus) corpus = String(vals[i][2]).trim();
    if (!number) number = String(vals[i][9]).trim();
    var fl = floorNorm(vals[i][4]);
    if (fl) floorsSet[fl] = true;
  }
  if (!rowsIdx.length) throw new Error('Чек-лист не найден: ' + id);

  ensureIntHeaders_(sheet);

  // Уже приложенные файлы (K одинаков во всей группе)
  var list = [];
  try {
    var cur = JSON.parse(String(vals[rowsIdx[0] - 2][10]));
    if (cur && cur.length) list = cur;
  } catch (e2) {}

  var floors = Object.keys(floorsSet).sort(function (a, b) { return parseFloat(a) - parseFloat(b); });
  var ext  = fileExt_(body.fileName, body.mimeType);
  var nice = safeFileName_(
    'СС_' + corpus + '_эт.' + floors.join(',') +
    (number ? '_№' + number : '') + '_фото' + (list.length + 1)
  ) + '.' + ext;

  var folder = getOrCreateChecklistFolder();
  var blob   = Utilities.newBlob(Utilities.base64Decode(body.fileData),
    body.mimeType || 'application/octet-stream', nice);
  var file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

  list.push({ n: file.getName(), u: file.getUrl() });
  var json = JSON.stringify(list);
  var mark = changeMark_(user, '+фото');
  rowsIdx.forEach(function (row) {
    sheet.getRange(row, 11).setValue(json);
    sheet.getRange(row, 13).setValue(mark);
  });
  SpreadsheetApp.flush();
  return { ok: true, name: file.getName(), url: file.getUrl(), files: list.length };
}

// «Удаление» внутреннего чек-листа. Строки из листа НЕ удаляются (правило безопасности) —
// вся группа помечается статусом «Удалён» (столбец F) и перестаёт отдаваться фронту.
// Файл на Диске остаётся. Восстановление — вручную вернуть прежний статус в листе.
function deleteInternalChecklist(body) {
  var id   = String(body.id   || '').trim();
  var user = String(body.user || '').trim();
  if (!id) throw new Error('Не указан ID чек-листа');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHK_INT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Лист «Чек листы_внутренние» пуст');
  ensureIntHeaders_(sheet);

  var n    = sheet.getLastRow() - 1;
  var idA  = sheet.getRange(2, 1, n, 1).getValues();   // столбец A (ID группы)
  var fCol = sheet.getRange(2, 6, n, 1).getValues();   // столбец F (Статус)
  var mCol = sheet.getRange(2, 13, n, 1).getValues();  // столбец M (Изменено)
  var mark = changeMark_(user, 'удалил');
  var updated = 0;
  for (var i = 0; i < n; i++) {
    if (String(idA[i][0]).trim() === id) { fCol[i][0] = 'Удалён'; mCol[i][0] = mark; updated++; }
  }
  if (!updated) throw new Error('Чек-лист не найден: ' + id);
  sheet.getRange(2, 6, n, 1).setValues(fCol);
  sheet.getRange(2, 13, n, 1).setValues(mCol);
  SpreadsheetApp.flush();
  return { ok: true, updated: updated };
}

function createInternalSheet(ss) {
  var sh = ss.insertSheet(CHK_INT_SHEET);
  sh.getRange(1, 1, 1, 13).setValues([[
    'ID', 'Дата', 'Корпус', 'ID_работы', 'Этаж', 'Статус', 'Ссылка', 'Файл', 'Комментарий', 'Номер',
    'Доп файлы', 'Загрузил', 'Изменено'
  ]]);
  sh.setFrozenRows(1);
  return sh;
}

// Дата из формы (YYYY-MM-DD) → dd.MM.yyyy; пусто → сегодня
function normDateIn(v) {
  var s = String(v || '').trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return m[3] + '.' + m[2] + '.' + m[1];
  if (s) return s;
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
}

function getOrCreateChecklistFolder() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('cc_chk_folder');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var it = DriveApp.getFoldersByName(CHK_FOLDER_NAME);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(CHK_FOLDER_NAME);
  props.setProperty('cc_chk_folder', folder.getId());
  return folder;
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
