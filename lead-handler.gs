// ============================================================
// BeeX — обработчик заявок (Google Apps Script)
// Разместить: Таблица Google -> Расширения -> Apps Script
// Данные (токен/чат/секрет) задать в PropertiesService:
//   Script properties -> TELEGRAM_TOKEN, TELEGRAM_CHAT_ID,
//   SECRET_KEY, NOTIFY_EMAIL (необязательно, по умолчанию beex805@gmail.com)
// Опубликовать: Deploy -> Web app
//   Execute as: Me  |  Who has access: Anyone
// ============================================================

function getProps() {
  const p = PropertiesService.getScriptProperties();
  return {
    TOKEN:   p.getProperty('TELEGRAM_TOKEN'),
    CHAT_ID: p.getProperty('TELEGRAM_CHAT_ID'),
    SECRET:  p.getProperty('SECRET_KEY'),
    EMAIL:   p.getProperty('NOTIFY_EMAIL') || 'beex805@gmail.com'
  };
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.key !== getProps().SECRET) {
      return json({ ok: false, error: 'bad key' });
    }

    const name    = String(data.name    || '—');
    const phone   = String(data.phone   || '—');
    const message = String(data.message || '');
    const date    = new Date().toLocaleString('ru-RU');

    // 1) Google Таблица (первый лист активной таблицы)
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheets()[0]
      .appendRow([date, name, phone, message]);

    // 2) Email через Gmail (Gmail SMTP от имени вашего аккаунта)
    const emailBody =
      'Новая заявка с сайта BeeX\n\n' +
      'Имя: '       + name    + '\n' +
      'Телефон: '   + phone   + '\n' +
      'Сообщение: ' + message + '\n\n' +
      'Время: '     + date;
    GmailApp.sendEmail(getProps().EMAIL, 'Новая заявка с сайта BeeX', emailBody);

    // 3) Telegram
    const tText =
      'Новая заявка BeeX\n\n' +
      'Имя: '       + name    + '\n' +
      'Телефон: '   + phone   + '\n' +
      'Сообщение: ' + message + '\n\n' +
      date;
    const tgUrl =
      'https://api.telegram.org/bot' + getProps().TOKEN +
      '/sendMessage?chat_id=' + getProps().CHAT_ID +
      '&text=' + encodeURIComponent(tText) +
      '&parse_mode=Markdown';
    UrlFetchApp.fetch(tgUrl, { muteHttpExceptions: true });

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, msg: 'BeeX lead handler' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
