// ═══════════════════════════════════════════════════════════════
//  SALON SHE — Web App + PDF Generator
//  Деплой: Extensions → Apps Script → Deploy → New deployment
//           Type: Web app | Execute as: Me | Access: Anyone
// ═══════════════════════════════════════════════════════════════

// ── Замініть ці значення на свої ────────────────────────────────
var TEMPLATE_ID  = '1QcFYuWKuBGCavz3BgdeeCh0jZkIQPdLIyXK25VK-M8Y'; // Google Docs шаблон
var FOLDER_ID    = '1fojOa6FFAN7yepbfBVAk4JwC4fpdob_T';              // Папка для PDF
var SHEET_ID     = '1_jAqCKaG4T1MzlRKhi8y_sYSa5glQ00tm7FniGWunhA';   // !! ОБОВ'ЯЗКОВО: вставте ID вашої Google Таблиці (з URL: /spreadsheets/d/XXXXXX/edit)
var SHEET_NAME   = 'HourlyAccess';                                    // Назва аркушу

// ── Платіжний режим ──────────────────────────────────────────────
var PAYMENT_REQUIRED = false; // true → генерувати PDF тільки після PAID

// ── Колонки Sheet ────────────────────────────────────────────────
var SHEET_HEADERS = [
  'timestamp', 'fullName', 'phone', 'email', 'address',
  'visitDateTime', 'hours', 'amountPLN', 'revolutLink',
  'invoiceRequested', 'nip', 'idDoc',
  'payment_status', 'payment_reference',
  'pdfFileId', 'statusEmailSent', 'errorMessage'
];

// ════════════════════════════════════════════════════════════════
//  doGet — повертає HTML сторінку (залишено для сумісності)
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Salon SHE — Реєстрація')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ════════════════════════════════════════════════════════════════
//  doPost — точка входу для fetch() з GitHub Pages
// ════════════════════════════════════════════════════════════════
function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ ok: false, error: 'Invalid JSON: ' + err.toString() });
  }

  try {
    var result = submitForm(payload);
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
//  submitForm — головна точка входу з фронтенду
// ════════════════════════════════════════════════════════════════
function submitForm(payload) {
  // 1. Валідація обов'язкових полів
  var errors = validatePayload(payload);
  if (errors.length > 0) {
    throw new Error('Validation: ' + errors.join('; '));
  }

  var sheet     = getOrCreateSheet();
  var rowData   = buildRowData(payload);
  var rowIndex  = appendRow(sheet, rowData);

  // 2. Платіжний gate
  if (PAYMENT_REQUIRED && payload.payment_status !== 'PAID') {
    updateCell(sheet, rowIndex, 'payment_status', 'PENDING');
    return { ok: false, paymentPending: true, message: 'Очікуємо підтвердження оплати.' };
  }

  // 3. Генерація PDF і відправка email
  var pdfFileId      = '';
  var emailSent      = false;
  var errorMessage   = '';

  try {
    var result   = generateAndSendPdf(payload);
    pdfFileId    = result.pdfFileId;
    emailSent    = result.emailSent;
  } catch (err) {
    errorMessage = err.toString();
    Logger.log('submitForm error: ' + errorMessage);
    updateCell(sheet, rowIndex, 'errorMessage', errorMessage);
    throw new Error('Помилка генерації PDF: ' + errorMessage);
  }

  // 4. Оновлюємо рядок у Sheet
  updateCell(sheet, rowIndex, 'pdfFileId',       pdfFileId);
  updateCell(sheet, rowIndex, 'statusEmailSent', emailSent ? 'TRUE' : 'FALSE');
  updateCell(sheet, rowIndex, 'errorMessage',    errorMessage);

  return {
    ok:      true,
    message: 'PDF згенеровано та відправлено.',
    pdfFileId: pdfFileId
  };
}

// ════════════════════════════════════════════════════════════════
//  Валідація payload
// ════════════════════════════════════════════════════════════════
function validatePayload(p) {
  var errs = [];
  if (!p.fullName || !p.fullName.trim())      errs.push('fullName required');
  if (!p.phone    || !p.phone.trim())         errs.push('phone required');
  if (!p.email    || !p.email.trim())         errs.push('email required');
  if (!p.visitDateTime)                       errs.push('visitDateTime required');
  if (!p.hours || p.hours < 1 || p.hours > 12) errs.push('hours must be 1-12');
  if (!p.acceptedRegulamin)                   errs.push('regulamin not accepted');
  if (!p.acceptedMonitoring)                  errs.push('monitoring not accepted');
  if (!p.signatureBase64 || p.signatureBase64.length < 100) errs.push('signature required');
  if (p.invoiceRequested && !p.nip)           errs.push('nip required when invoice requested');
  return errs;
}

// ════════════════════════════════════════════════════════════════
//  PDF: генерація + вставка підпису + відправка email
// ════════════════════════════════════════════════════════════════
function generateAndSendPdf(payload) {
  var templateFile = DriveApp.getFileById(TEMPLATE_ID);
  var folder       = DriveApp.getFolderById(FOLDER_ID);
  var docName      = 'Umowa_' + payload.fullName.trim();

  // Копіюємо шаблон
  var copy = templateFile.makeCopy(docName, folder);
  var doc  = DocumentApp.openById(copy.getId());
  var body = doc.getBody();

  // Форматуємо дату візиту
  var visitFormatted = formatVisitDate(payload.visitDateTime);

  // Замінюємо плейсхолдери текстом
  var textMap = {
    '{{nazwa}}':    payload.fullName   || '',
    '{{adres}}':    payload.address    || '',
    '{{dokument}}': payload.idDoc      || '',
    '{{Email}}':    payload.email      || '',
    '{{nip}}':      payload.nip        || '',
    '{{telefon}}':  payload.phone      || '',
    '{{data}}':     Utilities.formatDate(new Date(), 'GMT+2', 'dd.MM.yyyy'),
    '{{dataWizyty}}': visitFormatted,
    '{{godziny}}':  String(payload.hours || ''),
    '{{kwota}}':    String(payload.calculatedAmountPLN || '') + ' PLN',
    '{{pesel}}':    ''
  };

  for (var tag in textMap) {
    body.replaceText(tag, textMap[tag]);
  }

  // Вставляємо підпис замість {{SIGNATURE}}
  insertSignatureImage(body, payload.signatureBase64);

  doc.saveAndClose();

  // Генеруємо PDF
  var pdfBlob = DriveApp.getFileById(copy.getId()).getAs(MimeType.PDF);
  pdfBlob.setName(docName + '.pdf');
  var pdfFile = folder.createFile(pdfBlob);

  // Видаляємо тимчасовий Google Doc
  copy.setTrashed(true);

  // Відправляємо email
  var emailSent = false;
  if (payload.email) {
    MailApp.sendEmail({
      to:          payload.email,
      subject:     'Ваш договір — Salon SHE',
      body:        'Доброго дня, ' + payload.fullName + '!\n\n' +
                   'Дякуємо за реєстрацію. Ваш договір (PDF) додано до цього листа.\n\n' +
                   'Деталі візиту:\n' +
                   '  Дата і час: ' + visitFormatted + '\n' +
                   '  Кількість годин: ' + payload.hours + '\n' +
                   '  Сума: ' + payload.calculatedAmountPLN + ' PLN\n\n' +
                   'З повагою,\nAlbina Boichuk / Salon SHE',
      attachments: [pdfBlob]
    });
    emailSent = true;
  }

  return { pdfFileId: pdfFile.getId(), emailSent: emailSent };
}

// ════════════════════════════════════════════════════════════════
//  Вставка підпису в документ
// ════════════════════════════════════════════════════════════════
function insertSignatureImage(body, signatureBase64) {
  // Конвертуємо dataURL → blob
  var base64Data = signatureBase64;
  if (base64Data.indexOf(',') !== -1) {
    base64Data = base64Data.split(',')[1];
  }
  var sigBlob = Utilities.newBlob(
    Utilities.base64Decode(base64Data),
    'image/png',
    'signature.png'
  );

  // Шукаємо {{SIGNATURE}} в тексті
  var found = body.findText('\\{\\{SIGNATURE\\}\\}');
  if (!found) {
    // Якщо плейсхолдер не знайдено — додаємо підпис в кінець
    body.appendImage(sigBlob);
    return;
  }

  // Отримуємо елемент, що містить плейсхолдер
  var element   = found.getElement();
  var paragraph = element.getParent();

  // Визначаємо індекс параграфу в body
  var paraIndex = body.getChildIndex(paragraph);

  // Вставляємо зображення після параграфу з плейсхолдером
  var imgPara = body.insertImage(paraIndex + 1, sigBlob);

  // Задаємо розмір підпису (ширина ~200px, висота пропорційна)
  imgPara.setWidth(200);
  imgPara.setHeight(80);

  // Видаляємо параграф з плейсхолдером
  paragraph.removeFromParent();
}

// ════════════════════════════════════════════════════════════════
//  Google Sheet: отримати або створити
// ════════════════════════════════════════════════════════════════
function getOrCreateSheet() {
  var ss;
  if (!SHEET_ID || SHEET_ID.length === 0) {
    // Шукаємо існуючу таблицю за назвою в папці (без створення — потребує окремих прав)
    var files = DriveApp.getFolderById(FOLDER_ID).getFilesByName('SalonSHE_Logs');
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      throw new Error(
        'SHEET_ID не вказано і таблицю SalonSHE_Logs не знайдено в папці. ' +
        'Будь ласка: 1) Створіть Google Таблицю вручну, ' +
        '2) Скопіюйте її ID з URL (частина між /d/ і /edit), ' +
        '3) Вставте в константу SHEET_ID у code.gs, ' +
        '4) Зробіть новий Deploy.'
      );
    }
  } else {
    ss = SpreadsheetApp.openById(SHEET_ID);
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(SHEET_HEADERS);
    sheet.setFrozenRows(1);
    // Форматуємо заголовок
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
         .setFontWeight('bold')
         .setBackground('#b5845a')
         .setFontColor('#ffffff');
  }
  return sheet;
}

// ════════════════════════════════════════════════════════════════
//  Sheet: додати рядок
// ════════════════════════════════════════════════════════════════
function buildRowData(payload) {
  return [
    new Date(),                                  // timestamp
    payload.fullName        || '',
    payload.phone           || '',
    payload.email           || '',
    payload.address         || '',
    payload.visitDateTime   || '',
    payload.hours           || '',
    payload.calculatedAmountPLN || '',
    payload.revolutLink     || '',
    payload.invoiceRequested ? 'TRUE' : 'FALSE',
    payload.nip             || '',
    payload.idDoc           || '',
    payload.payment_status  || 'SKIPPED',
    payload.payment_reference || '',
    '',   // pdfFileId — заповнюється після генерації
    '',   // statusEmailSent
    ''    // errorMessage
  ];
}

function appendRow(sheet, rowData) {
  sheet.appendRow(rowData);
  return sheet.getLastRow();
}

function updateCell(sheet, rowIndex, colName, value) {
  var colIndex = SHEET_HEADERS.indexOf(colName) + 1;
  if (colIndex > 0) {
    sheet.getRange(rowIndex, colIndex).setValue(value);
  }
}

// ════════════════════════════════════════════════════════════════
//  Утиліти
// ════════════════════════════════════════════════════════════════
function formatVisitDate(dtString) {
  if (!dtString) return '';
  try {
    var d = new Date(dtString);
    return Utilities.formatDate(d, 'GMT+2', 'dd.MM.yyyy HH:mm');
  } catch (e) {
    return dtString;
  }
}

// ════════════════════════════════════════════════════════════════
//  Збережено для зворотної сумісності з Google Form trigger
// ════════════════════════════════════════════════════════════════
function onFormSubmit(e) {
  // 1. ID ваших файлів (замініть на свої з посилань в браузері)
  const templateId = TEMPLATE_ID;
  const folderId   = FOLDER_ID;

  // 2. Отримуємо дані з форми
  const responses = e.namedValues;
  const userEmail = responses['Email'] ? responses['Email'][0] : null;

  const dataMap = {
    "{{nazwa}}":    responses['Imię i nazwisko/Nazwa firmy'] ? responses['Imię i nazwisko/Nazwa firmy'][0] : "",
    "{{adres}}":    responses['twój adres zamieszkania (lub rozliczenie podatkowe)'] ? responses['twój adres zamieszkania (lub rozliczenie podatkowe)'][0] : "",
    "{{pesel}}":    responses['PESEL'] ? responses['PESEL'][0] : "",
    "{{dokument}}": responses['Nr dokumentu tożsamości'] ? responses['Nr dokumentu tożsamości'][0] : "",
    "{{Email}}":    responses['Email'] ? responses['Email'][0] : "",
    "{{nip}}":      responses['NIP (jeśli potrzebujesz fakturę)'] && responses['NIP (jeśli potrzebujesz fakturę)'][0] ? responses['NIP (jeśli potrzebujesz fakturę)'][0] : "",
    "{{data}}":     Utilities.formatDate(new Date(), "GMT+2", "dd.MM.yyyy")
  };

  try {
    const templateFile = DriveApp.getFileById(templateId);
    const folder       = DriveApp.getFolderById(folderId);
    const copy         = templateFile.makeCopy('Умова - ' + dataMap["{{nazwa}}"], folder);
    const doc          = DocumentApp.openById(copy.getId());
    const body         = doc.getBody();

    for (let tag in dataMap) {
      body.replaceText(tag, dataMap[tag]);
    }
    doc.saveAndClose();

    const pdfBlob = copy.getAs(MimeType.PDF);
    pdfBlob.setName('Umowa_' + dataMap["{{nazwa}}"] + '.pdf');
    folder.createFile(pdfBlob);
    copy.setTrashed(true);

    if (userEmail) {
      MailApp.sendEmail({
        to:          userEmail,
        subject:     'Ваш договір про користування інфраструктурою',
        body:        'Доброго дня! \n\nДякуємо за заповнення форми. Ваш договір (PDF) додано до цього листа. \n\nЗ повагою, Albina Boichuk',
        attachments: [pdfBlob]
      });
    }
  } catch (error) {
    Logger.log('Помилка: ' + error.toString());
  }
}