/**
 * ============================================================
 * GOOGLE APPS SCRIPT — Эко-урбанистика (ПОЛНАЯ ВЕРСИЯ)
 * ============================================================
 * Листы:
 *   Лист1  — Блок 3: Актуализация (actuly.html)
 *   Лист2  — Блок 6: Итог (itog.html)
 *   Лист3  — Суммарные баллы (Лист1 + Лист2)
 *   Лист4  — Рефлексия студентов (reflection.html)
 *
 * КАК ПОДКЛЮЧИТЬ:
 *  1. Google Sheets → Расширения → Apps Script → вставь этот код
 *  2. Deploy → New deployment → Web App
 *     Execute as: Me | Who has access: Anyone
 *  3. Скопируй URL → вставь в GAS_URL во всех HTML-файлах
 * ============================================================
 */

const SHEETS = {
  block3:     'Лист1',
  block6:     'Лист2',
  totals:     'Лист3',
  reflection: 'Лист4'
};

// ---------- CORS headers ----------
function setCORS(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET,POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function doOptions() {
  return setCORS(ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT));
}

// ---------- POST handler ----------
function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch(err) { return jsonOut({status:'error',msg:'bad json'}); }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- Блок 3: актуализация ----
  if (data.action === 'submit') {
    const s = getOrCreateSheet(ss, SHEETS.block3, ['Дата/Время','ФИО','Баллы (Блок3)','Ответы']);
    s.appendRow([data.timestamp, data.fio, Number(data.score)||0, data.answers||'']);
    updateTotals(ss, data.fio, Number(data.score)||0, 0);
    return jsonOut({status:'ok'});
  }

  // ---- Блок 6: итог ----
  if (data.action === 'submitItog') {
    const s = getOrCreateSheet(ss, SHEETS.block6, ['Дата/Время','ФИО','Баллы (Блок6)','Ответы']);
    s.appendRow([data.timestamp, data.fio, Number(data.score6)||0, data.answers||'']);
    updateTotals(ss, data.fio, 0, Number(data.score6)||0);
    return jsonOut({status:'ok'});
  }

  // ---- Рефлексия студентов (reflection.html) → Лист4 ----
  if (data.action === 'submitStudentReflection') {
    const s = getOrCreateSheet(ss, SHEETS.reflection, ['Дата/Время','ФИО','Эмоция','Звёзды','Понимание %','Заметка']);
    s.appendRow([data.timestamp, data.fio||'', data.mood||'', Number(data.stars)||0, Number(data.score)||0, data.note||'']);
    return jsonOut({status:'ok'});
  }

  // ---- Рефлексия преподавателя (index.html, Блок 7) → Лист4 ----
  if (data.action === 'submitTeacherReflection') {
    const s = getOrCreateSheet(ss, SHEETS.reflection, ['Дата/Время','ФИО','Эмоция','Звёзды','Понимание %','Заметка']);
    s.appendRow([data.timestamp, '(Преподаватель)', data.mood||'', Number(data.stars)||0, Number(data.score)||0, data.note||'']);
    return jsonOut({status:'ok'});
  }

  return jsonOut({status:'ignored'});
}

// ---------- GET handler ----------
function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Топ-3 для подиума Блока 3 (из Лист3)
  if (action === 'getTop3') {
    const s = ss.getSheetByName(SHEETS.totals);
    if (!s || s.getLastRow() < 2) return jsonOut([]);
    const rows = s.getRange(2, 1, s.getLastRow()-1, 4).getValues();
    const top3 = rows
      .filter(r => r[0])
      .sort((a,b) => (b[3]||0) - (a[3]||0))
      .slice(0,3)
      .map(r => ({name: r[0], score: r[3]}));
    return jsonOut(top3);
  }

  // Статистика рефлексии для графика на index.html (из Лист4)
  if (action === 'getReflectionStats') {
    const s = ss.getSheetByName(SHEETS.reflection);
    if (!s || s.getLastRow() < 2) return jsonOut([]);
    const rows = s.getRange(2, 1, s.getLastRow()-1, 6).getValues();
    const moodMap = {};
    rows.forEach(r => {
      const m = r[2] || '❓';
      moodMap[m] = (moodMap[m]||0) + 1;
    });
    const result = Object.entries(moodMap).map(([mood, count]) => ({mood, count}));
    return jsonOut(result);
  }

  return jsonOut({status:'ok'});
}

// ---------- HELPERS ----------
function getOrCreateSheet(ss, name, headers) {
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(headers);
    s.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a4a1a')
      .setFontColor('#76FF03');
  }
  return s;
}

function updateTotals(ss, fio, delta3, delta6) {
  const s = getOrCreateSheet(ss, SHEETS.totals, ['ФИО','Блок 3','Блок 6','ИТОГО']);
  const lr = s.getLastRow();
  let found = false;
  if (lr > 1) {
    const data = s.getRange(2, 1, lr-1, 4).getValues();
    for (let i=0; i<data.length; i++) {
      if ((data[i][0]||'').toLowerCase() === fio.toLowerCase()) {
        const row = i + 2;
        const prev3 = Number(data[i][1])||0;
        const prev6 = Number(data[i][2])||0;
        const n3 = delta3 > 0 ? delta3 : prev3;
        const n6 = delta6 > 0 ? delta6 : prev6;
        s.getRange(row, 2, 1, 3).setValues([[n3, n6, n3+n6]]);
        found = true;
        break;
      }
    }
  }
  if (!found) s.appendRow([fio, delta3, delta6, delta3+delta6]);
  // Сортировка по убыванию ИТОГО
  if (s.getLastRow() > 2) {
    s.getRange(2, 1, s.getLastRow()-1, 4).sort({column:4, ascending:false});
  }
}

function jsonOut(data) {
  return setCORS(
    ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON)
  );
}
