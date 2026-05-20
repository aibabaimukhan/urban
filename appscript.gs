/**
 * GOOGLE APPS SCRIPT — Эко-урбанистика
 * 
 * Как использовать:
 * 1. Открой Google Sheets → Extensions → Apps Script
 * 2. Вставь этот код
 * 3. Нажми Deploy → New deployment → Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Скопируй URL и вставь в GAS_URL в index.html, actuly.html, itog.html
 *
 * Структура листов:
 *   Лист 1 (Sheet1): Данные из actuly.html (Блок 3 — Актуализация)
 *   Лист 2 (Sheet2): Данные из itog.html (Блок 6 — Итог)
 *   Лист 3 (Sheet3): Суммарные баллы по студентам
 */

const SHEET_NAMES = {
  block3: 'Лист1',
  block6: 'Лист2',
  totals: 'Лист3'
};

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',msg:'bad json'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (data.action === 'submit') {
    // Block 3 — actuly.html
    const sheet1 = ss.getSheetByName(SHEET_NAMES.block3) || ss.insertSheet(SHEET_NAMES.block3);
    if (sheet1.getLastRow() === 0) {
      sheet1.appendRow(['Дата/Время', 'ФИО', 'Баллы (Блок 3)', 'Ответы JSON']);
    }
    sheet1.appendRow([data.timestamp, data.fio, data.score, data.answers]);
    updateTotals(ss, data.fio, parseInt(data.score) || 0, 0);
    return jsonOk();
  }

  if (data.action === 'submitItog') {
    // Block 6 — itog.html
    const sheet2 = ss.getSheetByName(SHEET_NAMES.block6) || ss.insertSheet(SHEET_NAMES.block6);
    if (sheet2.getLastRow() === 0) {
      sheet2.appendRow(['Дата/Время', 'ФИО', 'Баллы (Блок 6)', 'Ответы JSON']);
    }
    sheet2.appendRow([data.timestamp, data.fio, data.score6, data.answers]);
    updateTotals(ss, data.fio, 0, parseInt(data.score6) || 0);
    return jsonOk();
  }

  return jsonOk();
}

function updateTotals(ss, fio, delta3, delta6) {
  const sheet3 = ss.getSheetByName(SHEET_NAMES.totals) || ss.insertSheet(SHEET_NAMES.totals);
  
  if (sheet3.getLastRow() === 0) {
    sheet3.appendRow(['ФИО', 'Блок 3 (актуализация)', 'Блок 6 (итог)', 'СУММАРНЫЙ БАЛЛ']);
    sheet3.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1a4a1a').setFontColor('#76FF03');
  }

  // Find existing row for this student
  const data = sheet3.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === fio.toLowerCase()) {
      const row = i + 1;
      const prev3 = parseInt(data[i][1]) || 0;
      const prev6 = parseInt(data[i][2]) || 0;
      const new3 = delta3 > 0 ? delta3 : prev3;
      const new6 = delta6 > 0 ? delta6 : prev6;
      sheet3.getRange(row, 2).setValue(new3);
      sheet3.getRange(row, 3).setValue(new6);
      sheet3.getRange(row, 4).setValue(new3 + new6);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet3.appendRow([fio, delta3, delta6, delta3 + delta6]);
  }
  
  // Sort by total descending
  const lastRow = sheet3.getLastRow();
  if (lastRow > 2) {
    sheet3.getRange(2, 1, lastRow-1, 4).sort({column: 4, ascending: false});
  }
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = e.parameter.action;
  const sheetNum = e.parameter.sheet;

  if (action === 'getTop3') {
    const sheetName = sheetNum === '1' ? SHEET_NAMES.totals : SHEET_NAMES.totals;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
      return jsonOk([]);
    }
    const data = sheet.getRange(2, 1, Math.min(sheet.getLastRow()-1, 3), 4).getValues();
    const top3 = data.map(row => ({ name: row[0], score: row[3] }));
    return ContentService.createTextOutput(JSON.stringify(top3))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return jsonOk([]);
}

function jsonOk(data) {
  return ContentService.createTextOutput(JSON.stringify({status:'ok', data: data || null}))
    .setMimeType(ContentService.MimeType.JSON);
}
