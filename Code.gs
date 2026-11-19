/**
 * Сервер опроса графика работы (Google Apps Script).
 *
 * Как подключить:
 * 1. Откройте таблицу «Опрос графика работы: результаты».
 * 2. Расширения → Apps Script. Удалите всё, что там есть, и вставьте этот код.
 * 3. Замените ADMIN_KEY ниже на свой длинный ключ (его вводит только администратор).
 * 4. Нажмите «Сохранить», затем «Развернуть» → «Новое развёртывание» → тип «Веб-приложение».
 *    «Выполнять как»: я. «У кого есть доступ»: все. Нажмите «Развернуть» и разрешите доступ.
 * 5. Скопируйте адрес веб-приложения (заканчивается на /exec) и вставьте его
 *    в index.html в строку var API_URL = '';
 *
 * Голоса записываются на первый лист: столбцы A-E (дата, ФИО, должность, город, график).
 * Столбцы K и L служебные (защита от повторного голосования).
 * Формулы подсчёта в столбцах G-I продолжают работать.
 */
var ADMIN_KEY = 'ЗАМЕНИТЕ_НА_ДЛИННЫЙ_КЛЮЧ';
var FIRST_ROW = 5;
var OPTIONS = {
  a: 'Вариант 1: 9:00 – 18:00',
  b: 'Вариант 2: 8:30 – 17:30',
  c: 'Вариант 3: 8:00 – 17:00'
};

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function clean_(s, max) {
  s = String(s || '').replace(/\s+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function fioKey_(s) {
  return String(s).toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ').trim();
}

function vote_(p) {
  var fio = String(p.fio || '').replace(/\s+/g, ' ').trim().slice(0, 120);
  var position = clean_(p.position, 120);
  var city = clean_(p.city, 80);
  var option = String(p.option || '');
  var uid = String(p.uid || '').slice(0, 80);
  if (fio.split(' ').length < 2 || fio.length < 5 || position.length < 2 || city.length < 2 || !OPTIONS[option] || !uid) {
    return { ok: false, error: 'invalid' };
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sh = sheet_();
    var last = sh.getLastRow();
    var key = fioKey_(fio);
    if (last >= FIRST_ROW) {
      var keys = sh.getRange(FIRST_ROW, 11, last - FIRST_ROW + 1, 2).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (keys[i][0] === uid || keys[i][1] === key) return { ok: false, error: 'duplicate' };
      }
    }
    var row = Math.max(last + 1, FIRST_ROW);
    sh.getRange(row, 1, 1, 5).setValues([[new Date(), clean_(fio, 120), position, city, OPTIONS[option]]]);
    sh.getRange(row, 11, 1, 2).setValues([[uid, key]]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function results_(p) {
  if (!ADMIN_KEY || ADMIN_KEY.indexOf('ЗАМЕНИТЕ') === 0 || p.key !== ADMIN_KEY) {
    return { ok: false, error: 'forbidden' };
  }
  var sh = sheet_();
  var last = sh.getLastRow();
  var voters = [];
  if (last >= FIRST_ROW) {
    var n = last - FIRST_ROW + 1;
    var rows = sh.getRange(FIRST_ROW, 1, n, 5).getValues();
    var ids = sh.getRange(FIRST_ROW, 11, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      var r = rows[i];
      if (!r[1]) continue;
      var text = String(r[4]);
      var opt = text.indexOf('Вариант 1') === 0 ? 'a' : text.indexOf('Вариант 2') === 0 ? 'b' : text.indexOf('Вариант 3') === 0 ? 'c' : '';
      if (!opt) continue;
      voters.push({
        id: ids[i][0] || ('row' + (FIRST_ROW + i)),
        t: r[0] instanceof Date ? r[0].getTime() : 0,
        fio: r[1],
        position: r[2],
        city: r[3],
        option: opt
      });
    }
  }
  return { ok: true, voters: voters };
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.action === 'vote') out = vote_(p);
    else if (p.action === 'results') out = results_(p);
    else if (p.action === 'ping') out = { ok: true };
    else out = { ok: false, error: 'bad_action' };
  } catch (err) {
    out = { ok: false, error: 'server' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
