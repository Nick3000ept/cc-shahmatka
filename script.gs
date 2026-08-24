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
const CHK_ARC_SHEET   = 'Архив';
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
    if (action === 'migrateInt')    return jsonOut(migrateInternalSheet());
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
    if (body.action === 'editChecklistCells') {
      return jsonOut(editInternalChecklistCells(body));
    }
    if (body.action === 'deleteArchiveChecklist') {
      return jsonOut(deleteArchiveChecklist(body));
    }
    if (body.action === 'removeArchiveWork') {
      return jsonOut(removeArchiveWork(body));
    }
    if (body.action === 'editArchiveCells') {
      return jsonOut(editArchiveCells(body));
    }
    if (body.action === 'importArchive') {
      return jsonOut(importArchiveRows(body));
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
  stats.archive = addArchiveChecklists(ss, checkByCell); // архив acons-app (лист «Архив»)
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
// Колонки: A ID  B Дата  C Корпус  D Система (название)  E Этаж  F Статус  G Ссылка  H Файл
//          I Комментарий  J Номер  K Доп файлы (JSON [{n,u}] — догруженные фото)
//          L Загрузил  M Изменено (кто · когда · что)  N ID_работы (ключ)  O Работа (название)
// Ключ привязки — N; в старых строках (до миграции) ID лежал в D — читаем с запасным вариантом.
function addInternalChecklists(ss, checkByCell) {
  var sheet = ss.getSheetByName(CHK_INT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues().forEach(function (r) {
    var workId = String(r[13]).trim() || String(r[3]).trim();
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

// ─── Архив чек-листов acons-app (лист «Архив») ──────────────────────────────
// Одна строка = чек-лист × этаж × работа шахматки (привязка выполнена заранее).
// Колонки: A ID_архива  B Номер акта  C Дата  D Статус  E Корпус  F Этаж  G ID_работы
//          H Система  I Работа  J Локация  K Захватка  L Метод привязки  M Файл
//          N Ссылка  O Подрядчик  P Комментарий  Q Пометка («Удалён · кто · когда» — скрыта)
// Архив read-only: правится только пометка Q (мягкое удаление всей группы по ID_архива).
var ARC_HEADERS = ['ID_архива', 'Номер акта', 'Дата', 'Статус', 'Корпус', 'Этаж', 'ID_работы',
  'Система', 'Работа', 'Локация', 'Захватка', 'Метод привязки', 'Файл', 'Ссылка',
  'Подрядчик', 'Комментарий', 'Пометка'];

function addArchiveChecklists(ss, checkByCell) {
  var sheet = ss.getSheetByName(CHK_ARC_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var added = 0;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues().forEach(function (r) {
    var workId = String(r[6]).trim();
    var corpus = String(r[4]).trim();
    if (!workId || !corpus) return;
    if (String(r[16]).trim().indexOf('Удалён') === 0) return; // «мягко» удалённые не отдаём
    var key = workId + '\x00' + corpus + '\x00' + floorNorm(r[5]);
    (checkByCell[key] = checkByCell[key] || []).push({
      id     : String(r[0]).trim(),
      status : String(r[3]).trim(),
      akt    : String(r[1]).trim(),
      date   : formatDateOut(r[2]),
      link   : String(r[13]).trim(),
      file   : String(r[12]).trim(),
      comment: String(r[15]).trim(),
      by     : String(r[14]).trim(), // подрядчик
      source : 'archive',
    });
    added++;
  });
  return added;
}

// «Удаление» архивного чек-листа: строки НЕ удаляются, вся группа (по ID_архива)
// получает пометку в столбце Q и перестаёт отдаваться фронту.
// Восстановление — очистить Q у строк группы в листе «Архив».
function deleteArchiveChecklist(body) {
  var id   = String(body.id   || '').trim();
  var user = String(body.user || '').trim();
  if (!id) throw new Error('Не указан ID чек-листа');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHK_ARC_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Лист «Архив» пуст');

  var n    = sheet.getLastRow() - 1;
  var idA  = sheet.getRange(2, 1, n, 1).getValues();   // столбец A (ID_архива)
  var qCol = sheet.getRange(2, 17, n, 1).getValues();  // столбец Q (Пометка)
  var mark = changeMark_(user, 'удалил'); // «кто · когда · удалил»
  var updated = 0;
  for (var i = 0; i < n; i++) {
    if (String(idA[i][0]).trim() === id) { qCol[i][0] = 'Удалён · ' + mark; updated++; }
  }
  if (!updated) throw new Error('Чек-лист не найден: ' + id);
  sheet.getRange(2, 17, n, 1).setValues(qCol);
  SpreadsheetApp.flush();
  return { ok: true, updated: updated };
}

// Убрать одну работу из архивного чек-листа (правка привязки): все строки группы
// (ID_архива) с этим ID_работы получают пометку в Q и перестают отдаваться фронту.
// Строки не удаляются; восстановление — очистить Q в листе «Архив».
function removeArchiveWork(body) {
  var id     = String(body.id     || '').trim();
  var workId = String(body.workId || '').trim();
  var user   = String(body.user   || '').trim();
  if (!id)     throw new Error('Не указан ID чек-листа');
  if (!workId) throw new Error('Не указан ID работы');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHK_ARC_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Лист «Архив» пуст');

  var n    = sheet.getLastRow() - 1;
  var idA  = sheet.getRange(2, 1, n, 1).getValues();  // A ID_архива
  var wG   = sheet.getRange(2, 7, n, 1).getValues();  // G ID_работы
  var qCol = sheet.getRange(2, 17, n, 1).getValues(); // Q Пометка
  var mark = 'Удалён · ' + changeMark_(user, 'убрал работу ' + workId);
  var updated = 0;
  for (var i = 0; i < n; i++) {
    if (String(idA[i][0]).trim() === id && String(wG[i][0]).trim() === workId &&
        String(qCol[i][0]).trim().indexOf('Удалён') !== 0) {
      qCol[i][0] = mark; updated++;
    }
  }
  if (!updated) throw new Error('Строки не найдены: ' + id + ' / ' + workId);
  sheet.getRange(2, 17, n, 1).setValues(qCol);
  SpreadsheetApp.flush();
  return { ok: true, updated: updated };
}

// Правка состава архивного чек-листа (режим выделения ячеек на шахматке).
// body: { id, cells:[{workId,floor}], user } — полный новый набор ячеек группы.
// Строки не удаляются: убранная ячейка — пометка в Q, добавленная — новая строка
// с копией общих полей группы и данными работы из справочника (метод — «правка вручную»).
// Мета (акт/дата/статус/корпус/файл) не меняется.
function editArchiveCells(body) {
  var id    = String(body.id   || '').trim();
  var user  = String(body.user || '').trim();
  var cells = body.cells || [];
  if (!id)           throw new Error('Не указан ID чек-листа');
  if (!cells.length) throw new Error('Нужна хотя бы одна ячейка');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHK_ARC_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Лист «Архив» пуст');

  var n    = sheet.getLastRow() - 1;
  var vals = sheet.getRange(2, 1, n, 17).getValues(); // A–Q

  // Новый набор: ключ «работа + этаж»
  var want = {};
  cells.forEach(function (c) {
    want[String(c.workId).trim() + '\x00' + floorNorm(c.floor)] = c;
  });

  // Активные строки группы (без помеченных)
  var active = [], have = {}, tpl = null;
  for (var i = 0; i < n; i++) {
    if (String(vals[i][0]).trim() !== id) continue;
    if (String(vals[i][16]).trim().indexOf('Удалён') === 0) continue;
    active.push(i);
    if (!tpl) tpl = vals[i];
    have[String(vals[i][6]).trim() + '\x00' + floorNorm(vals[i][5])] = true;
  }
  if (!active.length) throw new Error('Чек-лист не найден: ' + id);

  var mark = 'Удалён · ' + changeMark_(user, 'правка состава');

  // Убранные ячейки — пометить
  var removed = 0;
  active.forEach(function (i) {
    var key = String(vals[i][6]).trim() + '\x00' + floorNorm(vals[i][5]);
    if (!want[key]) { sheet.getRange(i + 2, 17).setValue(mark); removed++; }
  });

  // Добавленные — новые строки: общие поля из шаблона группы, работа — из справочника
  var wmap = {};
  readWorks(ss).forEach(function (w) { wmap[w.workId] = w; });
  var newRows = [];
  for (var key2 in want) {
    if (have[key2]) continue;
    var p = key2.split('\x00');
    var w = wmap[p[0]] || {};
    newRows.push([
      id, tpl[1], tpl[2], tpl[3], tpl[4], p[1], p[0],
      w.sistema || '', w.work || '', w.lokacia || '', w.zahvatka || '',
      'правка вручную (' + user + ')', tpl[12], tpl[13], tpl[14], tpl[15], '',
    ].map(function (c) { return safeCellArc_(c); }));
  }
  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 17).setValues(newRows);
  }
  SpreadsheetApp.flush();
  return { ok: true, added: newRows.length, removed: removed };
}

// Пакетная загрузка строк в лист «Архив» (разовый перенос из выгрузки acons-app).
// body: { pwd, rows: [[A..P] × N] } — pwd обязателен; строки только дописываются.
function importArchiveRows(body) {
  if (String(body.pwd) !== ADMIN_PASSWORD) throw new Error('Нет доступа');
  var rows = body.rows || [];
  if (!rows.length) throw new Error('Нет строк');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHK_ARC_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CHK_ARC_SHEET);
    sheet.getRange(1, 1, 1, ARC_HEADERS.length).setValues([ARC_HEADERS]);
    sheet.setFrozenRows(1);
  }
  var vals = rows.map(function (r) {
    var row = r.slice(0, 16);
    while (row.length < 16) row.push('');
    row.push(''); // Q Пометка — пустая
    return row.map(function (c) { return safeCellArc_(c); });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, vals.length, 17).setValues(vals);
  SpreadsheetApp.flush();
  return { ok: true, added: vals.length, total: sheet.getLastRow() - 1 };
}

// Экранирование значения для листа «Архив»: строки, начинающиеся с =,+,@ — апостроф
// спереди (защита от formula injection); ограничение длины.
function safeCellArc_(v) {
  var s = String(v == null ? '' : v);
  if (s.length > 5000) s = s.slice(0, 5000);
  if (/^[=+@]/.test(s)) s = "'" + s;
  return s;
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
  var works = worksById_(ss);
  var group = 'IC-' + Date.now();
  var rows  = cells.map(function (c) {
    var wid = String(c.workId).trim();
    var w   = works[wid] || {};
    // A id · B дата · C корпус · D система · E этаж · F статус · G ссылка · H файл ·
    // I комментарий · J номер · K доп файлы · L загрузил · M изменено · N ID_работы · O работа
    return [group, dateStr, corpus, w.sistema || '', c.floor, status, url, fname, comment, number,
            '', user, '', wid, w.nameGrid || w.work || wid];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();

  return { ok: true, id: group, url: url, rows: rows.length };
}

// Карта справочника: workId → работа (для названий системы и работы)
function worksById_(ss) {
  var map = {};
  readWorks(ss).forEach(function (w) { map[w.workId] = w; });
  return map;
}

// Шапки столбцов (лист мог быть создан старой версией с меньшим числом столбцов).
// Заполняем только ПУСТЫЕ ячейки шапки — переименованные пользователем не трогаем.
function ensureIntHeaders_(sheet) {
  var want = {
    4: 'Система', 10: 'Номер', 11: 'Доп файлы', 12: 'Загрузил',
    13: 'Изменено', 14: 'ID_работы', 15: 'Работа',
  };
  for (var col in want) {
    var cell = sheet.getRange(1, Number(col));
    if (!String(cell.getValue()).trim()) cell.setValue(want[col]);
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
    if (String(vals[i][5]).trim() === 'Удалён') continue; // убранные этажи не учитываем
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

// Правка состава чек-листа.
// body: { id, cells:[{workId,floor}], number?, date?(YYYY-MM-DD), comment?, user }
// cells — полный новый набор ячеек (работа+этаж) группы, минимум одна.
// number/date/comment, если переданы, обновляют № (J), дату (B) и комментарий (I) всей группы.
// Строки НЕ удаляются (правило безопасности): убранная ячейка — её строка помечается
// статусом «Удалён»; добавленная — новая строка с копией полей из существующей.
function editInternalChecklistCells(body) {
  var id      = String(body.id   || '').trim();
  var user    = String(body.user || '').trim();
  var cells   = body.cells || [];
  var number  = (body.number  !== undefined) ? String(body.number).trim()  : null;
  var comment = (body.comment !== undefined) ? String(body.comment).trim() : null;
  var dateStr = body.date ? normDateIn(body.date) : null;
  if (!id)           throw new Error('Не указан ID чек-листа');
  if (!cells.length) throw new Error('Нужна хотя бы одна ячейка');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHK_INT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Лист «Чек листы_внутренние» пуст');
  ensureIntHeaders_(sheet);

  var n    = sheet.getLastRow() - 1;
  var vals = sheet.getRange(2, 1, n, 15).getValues(); // A–O

  // Новый набор: ключ «работа + этаж»
  var want = {};
  cells.forEach(function (c) {
    want[String(c.workId).trim() + '\x00' + floorNorm(c.floor)] = c;
  });

  // Активные строки группы (без «мягко» удалённых); ID работы — N, в старых строках — D
  var active = [], have = {}, tpl = null;
  for (var i = 0; i < n; i++) {
    if (String(vals[i][0]).trim() !== id) continue;
    if (String(vals[i][5]).trim() === 'Удалён') continue;
    active.push(i);
    if (!tpl) tpl = vals[i];
    var rowWid = String(vals[i][13]).trim() || String(vals[i][3]).trim();
    have[rowWid + '\x00' + floorNorm(vals[i][4])] = true;
  }
  if (!active.length) throw new Error('Чек-лист не найден: ' + id);

  var mark = changeMark_(user, 'изменён');

  // Итоговые дата/№/комментарий группы (переданы — новые, нет — прежние)
  var newDate = (dateStr !== null) ? dateStr : tpl[1];
  var newNum  = (number  !== null) ? number  : tpl[9];
  var newComm = (comment !== null) ? comment : tpl[8];
  var metaChanged =
    (dateStr !== null && dateStr !== formatDateOut(tpl[1])) ||
    (number  !== null && number  !== String(tpl[9]).trim()) ||
    (comment !== null && comment !== String(tpl[8]).trim());

  // Убранные ячейки — пометить «Удалён»; остающиеся — обновить дату/№/комментарий
  var removed = 0;
  active.forEach(function (i) {
    var rowWid = String(vals[i][13]).trim() || String(vals[i][3]).trim();
    var key = rowWid + '\x00' + floorNorm(vals[i][4]);
    if (!want[key]) {
      sheet.getRange(i + 2, 6).setValue('Удалён');
      sheet.getRange(i + 2, 13).setValue(mark);
      removed++;
    } else if (metaChanged) {
      sheet.getRange(i + 2, 2).setValue(newDate);   // B дата
      sheet.getRange(i + 2, 9).setValue(newComm);   // I комментарий
      sheet.getRange(i + 2, 10).setValue(newNum);   // J номер
      sheet.getRange(i + 2, 13).setValue(mark);     // M изменено
    }
  });

  // Добавленные ячейки — новые строки с копией полей группы (и новыми датой/№/комментарием)
  var works   = worksById_(ss);
  var newRows = [];
  for (var key in want) {
    if (have[key]) continue;
    var c = want[key];
    var flVal = parseFloat(c.floor);
    if (isNaN(flVal)) flVal = c.floor;
    var wid = String(c.workId).trim();
    var w   = works[wid] || {};
    // A id · B дата · C корпус · D система · E этаж · F статус · G ссылка · H файл ·
    // I комментарий · J номер · K доп файлы · L загрузил · M изменено · N ID_работы · O работа
    newRows.push([id, newDate, tpl[2], w.sistema || '', flVal,
                  tpl[5], tpl[6], tpl[7], newComm, newNum, tpl[10], tpl[11], mark,
                  wid, w.nameGrid || w.work || wid]);
  }
  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 15).setValues(newRows);
  }
  SpreadsheetApp.flush();
  return { ok: true, added: newRows.length, removed: removed };
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
  sh.getRange(1, 1, 1, 15).setValues([[
    'ID', 'Дата', 'Корпус', 'Система', 'Этаж', 'Статус', 'Ссылка', 'Файл', 'Комментарий', 'Номер',
    'Доп файлы', 'Загрузил', 'Изменено', 'ID_работы', 'Работа'
  ]]);
  sh.setFrozenRows(1);
  return sh;
}

// Разовая миграция листа «Чек листы_внутренние» (GET ?action=migrateInt, идемпотентна):
// в старых строках ID работы лежал в D — переносим его в N, в D пишем название системы,
// в O — название работы (по справочнику). Строки, где N уже заполнен, не трогаем.
function migrateInternalSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CHK_INT_SHEET);
  if (!sheet) return { ok: true, migrated: 0 };
  ensureIntHeaders_(sheet);
  var n = sheet.getLastRow() - 1;
  if (n < 1) return { ok: true, migrated: 0 };

  var works = worksById_(ss);
  var dCol = sheet.getRange(2, 4,  n, 1).getValues();
  var nCol = sheet.getRange(2, 14, n, 1).getValues();
  var oCol = sheet.getRange(2, 15, n, 1).getValues();
  var migrated = 0;
  for (var i = 0; i < n; i++) {
    if (String(nCol[i][0]).trim()) continue;      // уже мигрирована
    var d = String(dCol[i][0]).trim();
    var w = works[d];
    if (!w) continue;                             // в D не ID работы — не трогаем
    nCol[i][0] = d;
    dCol[i][0] = w.sistema || '';
    oCol[i][0] = w.nameGrid || w.work || d;
    migrated++;
  }
  sheet.getRange(2, 4,  n, 1).setValues(dCol);
  sheet.getRange(2, 14, n, 1).setValues(nCol);
  sheet.getRange(2, 15, n, 1).setValues(oCol);
  SpreadsheetApp.flush();
  return { ok: true, migrated: migrated };
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
