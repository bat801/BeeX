// ============================================================
// BeeX — обработчик заявок (Google Apps Script)
// Таблица Google -> Расширения -> Apps Script
//
// Script properties (Project Settings -> Script properties):
//   TELEGRAM_TOKEN, TELEGRAM_CHAT_ID, SECRET_KEY,
//   NOTIFY_EMAIL (необяз., по умолч. beex805@gmail.com),
//   VK_TOKEN, VK_USER_ID,
//   MAX_TOKEN, MAX_CHAT_ID, MAX_API_URL (пока заглушка),
//   ADMIN_PASSWORD (для admin.html)
//
// Deploy -> Web app: Execute as: Me | Who has access: Anyone
// После любых правок кода — переиздать деплой (New deployment / Update).
// ============================================================

function getProps() {
  const p = PropertiesService.getScriptProperties();
  return {
    TELEGRAM_TOKEN: p.getProperty('TELEGRAM_TOKEN'),
    TELEGRAM_CHAT_ID: p.getProperty('TELEGRAM_CHAT_ID'),
    SECRET: p.getProperty('SECRET_KEY'),
    EMAIL: p.getProperty('NOTIFY_EMAIL') || 'beex805@gmail.com',
    VK_TOKEN: p.getProperty('VK_TOKEN'),
    VK_USER_ID: p.getProperty('VK_USER_ID'),
    MAX_TOKEN: p.getProperty('MAX_TOKEN'),
    MAX_CHAT_ID: p.getProperty('MAX_CHAT_ID'),
    MAX_API_URL: p.getProperty('MAX_API_URL'),
    ADMIN_PASSWORD: p.getProperty('ADMIN_PASSWORD') || 'beex-admin'
  };
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) return json({ ok: false, error: 'no data' });
    const p = getProps();
    const data = JSON.parse(e.postData.contents);
    if (data.key !== p.SECRET) return json({ ok: false, error: 'bad key' });

    const name = String(data.name || '—');
    const phone = String(data.phone || '—');
    const message = String(data.message || '');
    const date = new Date().toLocaleString('ru-RU');

    const leadText =
      'Новая заявка BeeX\n\n' +
      'Имя: ' + name + '\n' +
      'Телефон: ' + phone + '\n' +
      'Сообщение: ' + message + '\n\n' +
      date;

    // 1) Google Таблица
    try {
      SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
        .appendRow([date, name, phone, message]);
    } catch (err) { log('sheet', err); }

    // 2) Email (Gmail SMTP)
    try {
      GmailApp.sendEmail(p.EMAIL, 'Новая заявка с сайта BeeX', leadText);
    } catch (err) { log('email', err); }

    // 3) Telegram
    try { sendToTelegram(leadText, p); } catch (err) { log('telegram', err); }

    // 4) VK
    try { sendToVk(leadText, p); } catch (err) { log('vk', err); }

    // 5) MAX (заглушка — активируется при наличии MAX_API_URL/MAX_TOKEN)
    try { sendToMax(leadText, p); } catch (err) { log('max', err); }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function sendToTelegram(text, p) {
  if (!p.TELEGRAM_TOKEN || !p.TELEGRAM_CHAT_ID) return;
  const url = 'https://api.telegram.org/bot' + p.TELEGRAM_TOKEN +
    '/sendMessage?chat_id=' + p.TELEGRAM_CHAT_ID +
    '&text=' + encodeURIComponent(text) +
    '&parse_mode=Markdown';
  UrlFetchApp.fetch(url, { muteHttpExceptions: true });
}

function sendToVk(text, p) {
  if (!p.VK_TOKEN || !p.VK_USER_ID) return;
  const url = 'https://api.vk.ru/method/messages.send' +
    '?access_token=' + encodeURIComponent(p.VK_TOKEN) +
    '&user_id=' + encodeURIComponent(p.VK_USER_ID) +
    '&random_id=' + Math.floor(Math.random() * 1e9) +
    '&v=5.199' +
    '&message=' + encodeURIComponent(text);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  try {
    const r = JSON.parse(res.getContentText());
    if (r && r.error) log('vk-api', r.error.error_msg || JSON.stringify(r.error));
  } catch (e) {}
}

function sendToMax(text, p) {
  // ЗАГЛУШКА: Bot API для MAX пока не подтверждён.
  // Когда появится документация — укажите MAX_API_URL (и при необходимости
  // поправьте формат параметров под реальный эндпоинт). Пока не настроено — пропускаем.
  if (!p.MAX_API_URL || !p.MAX_TOKEN) return;
  const url = p.MAX_API_URL +
    '?token=' + encodeURIComponent(p.MAX_TOKEN) +
    '&chat_id=' + encodeURIComponent(p.MAX_CHAT_ID || '') +
    '&text=' + encodeURIComponent(text);
  UrlFetchApp.fetch(url, { muteHttpExceptions: true });
}

function doGet(e) {
  const p = getProps();
  const params = e.parameter || {};
  if (params.mode === 'list') {
    if (params.key !== p.SECRET || params.pass !== p.ADMIN_PASSWORD) {
      return json({ ok: false, error: 'unauthorized' });
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const values = sheet.getDataRange().getValues();
    const rows = values.slice(1).map(function (r) {
      return r.map(function (c) { return String(c); });
    });
    return json({ ok: true, rows: rows });
  }
  return json({ ok: true, msg: 'BeeX lead handler' });
}

function log(channel, err) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('log');
    if (!sh) sh = ss.insertSheet('log');
    sh.appendRow([new Date().toLocaleString('ru-RU'), channel, String(err)]);
  } catch (e) {}
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
