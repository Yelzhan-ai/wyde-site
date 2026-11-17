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
// Новые столбцы всегда добавляются В КОНЕЦ списка, чтобы старые строки таблицы не съехали.
const COLUMNS = [
  ['date',      'Дата'],
  ['name',      'Имя'],
  ['whatsapp',  'WhatsApp'],
  ['instagram', 'Instagram'],
  ['city',      'Город'],
  ['gender',    'Пол'],
  ['lenL',      'Длина левой стопы, см'],
  ['lenR',      'Длина правой стопы, см'],
  ['ball',      'Ширина стопы, см'],
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
  ['utm',       'Откуда пришёл'],
  // способ замера: "snapfeet" (приложение) или "paper" (бумага A4)
  ['method',                'Способ замера'],
  // цифры из приложения Snapfeet, в миллиметрах
  ['snapfeet_left_length',  'Snapfeet: длина левой, мм'],
  ['snapfeet_right_length', 'Snapfeet: длина правой, мм'],
  ['snapfeet_left_width',   'Snapfeet: ширина левой, мм'],
  ['snapfeet_right_width',  'Snapfeet: ширина правой, мм'],
  ['snapfeet_left_instep',  'Snapfeet: подъём левой, мм'],
  ['snapfeet_right_instep', 'Snapfeet: подъём правой, мм']
];

// Сайт присылает анкету сюда.
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // чтобы две анкеты, пришедшие одновременно, не перепутались
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // Шапка: на пустом листе рисуем её, на старом — дописываем новые столбцы справа.
    const headers = COLUMNS.map(function (c) { return c[1]; });
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }
    const head = sheet.getRange(1, 1, 1, headers.length);
    if (head.getValues()[0].join('|') !== headers.join('|')) head.setValues([headers]);
    head.setFontWeight('bold');

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
