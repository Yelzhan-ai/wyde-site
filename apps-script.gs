/**
 * WYDE — приём анкет «Первая тысяча» в Google Таблицу.
 *
 * Этот код вставляется в Google Таблицу: Расширения → Apps Script.
 * Пошаговая инструкция — в README.md, раздел «Как подключить Google Таблицу».
 */

// Название листа, куда падают анкеты (создаётся автоматически).
const SHEET_NAME = 'WYDE';

// Ключ из анкеты → заголовок столбца.
// Порядок строк здесь = порядок столбцов в таблице. Первая строка — дата отправки.
const COLUMNS = [
  ['date',      'Дата'],
  ['name',      'Имя'],
  ['whatsapp',  'WhatsApp'],
  ['instagram', 'Instagram'],
  ['city',      'Город'],
  ['gender',    'Пол'],
  ['lenL',      'Длина левой стопы, см'],
  ['lenR',      'Длина правой стопы, см'],
  ['ball',      'Ширина в пучках, см'],
  ['heel',      'Ширина пятки, см'],
  ['girth',     'Обхват в пучках, см'],
  ['instep',    'Обхват в подъёме, см'],
  ['insole',    'Ширина стельки, см'],
  ['widerMm',   'Стопа шире стельки, мм'],
  ['jis',       'Полнота JIS'],
  ['type',      'Тип стопы'],
  ['pains',     'Где обувь проигрывает'],
  ['brands',    'Бренды и ширины'],
  ['model',     'Голос: модель'],
  ['color',     'Голос: цвет'],
  ['lang',      'Язык сайта'],
  ['consent',   'Согласие на обработку данных'],
  ['utm',       'Откуда пришёл']
];

// Сайт присылает анкету сюда.
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // чтобы две анкеты, пришедшие одновременно, не перепутались
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) { // первая анкета — рисуем шапку
      sheet.appendRow(COLUMNS.map(function (c) { return c[1]; }));
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    }

    sheet.appendRow(COLUMNS.map(function (c) {
      return c[0] === 'date' ? new Date() : clean(data[c[0]]);
    }));
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// Проверка: открой URL веб-приложения в браузере — должно показать {"ok":true,"service":"WYDE"}.
function doGet() {
  return json({ ok: true, service: 'WYDE' });
}

// Числа оставляем числами. Текст, который начинается с = + - @,
// записываем как обычный текст, чтобы таблица не приняла его за формулу.
function clean(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return v;
  v = String(v).slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
