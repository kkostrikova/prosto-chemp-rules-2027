/**
 * PROSTO CHEMP — agreements + payment receipt collector
 * Google Apps Script Web App
 */

const SHEET_ID = '13ncm5NW4SmAYWL6UALzEgugty4yxIH_IlWK4QKRZoYQ';
const SIGNATURE_FOLDER_ID = '1D53NN34oJcOBi7SMm_2GPvV27H8hFQeM';
const RECEIPT_FOLDER_ID = '16-PgfALTv96xqApa8uG3FJzwCEhHZnk0';
const COMPULSORY_FOLDER_ID = '1Wki6yogM-b4H3fCGAcAhPtvRYqDzLTW2';

const SIGNATURE_SHEET = 'Signatures';
const PAYMENT_SHEET = 'Payments';
const COMPULSORY_SHEET = 'Compulsory';
const HOST_SHEET = 'Для ведучої';
const LOG_SHEET = 'SystemLog';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'artRoutineDescription') {
    try {
      const result = saveArtRoutineDescription_({
        athleteName:p.athleteName || '',
        category:p.category || '',
        ageCategory:p.ageCategory || '',
        apparatus:p.apparatus || '',
        routineDescription:p.routineDescription || '',
        rulesVersion:p.rulesVersion || '2026/27'
      });
      return jsonp_(result, p.callback);
    } catch (err) {
      return jsonp_({ok:false,error:String(err)}, p.callback);
    }
  }
  return json_({ok:true,service:'PROSTO CHEMP backend',version:'14'});
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (data.submissionType === 'paymentReceipt') {
      return savePaymentReceipt_(data);
    }

    if (data.submissionType === 'compulsoryForm') {
      return saveCompulsoryForm_(data);
    }

    if (data.submissionType === 'artRoutineDescription') {
      return json_(saveArtRoutineDescription_(data));
    }

    return saveAgreement_(data);
  } catch (err) {
    try {
      const failed = JSON.parse((e && e.postData && e.postData.contents) || '{}');
      log_(
        failed.submissionType || 'unknown',
        'doPost',
        failed,
        'ERROR',
        String(err && err.stack ? err.stack : err)
      );
    } catch (logErr) {}
    return json_({ok:false,error:String(err)});
  }
}

function saveAgreement_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SIGNATURE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SIGNATURE_SHEET);
    sh.appendRow([
      'Timestamp','Agreement type','Athlete','Athlete DOB',
      'Representative','Representative role','Phone','Email',
      'Rules version','Signature file'
    ]);
  }

  let signatureUrl = '';
  if (data.signatureDataUrl) {
    const parts = data.signatureDataUrl.split(',');
    const raw = Utilities.base64Decode(parts[1] || '');
    const safeName = safe_(data.athleteName || data.representativeName || 'signature');
    const blob = Utilities.newBlob(raw, 'image/png', safeName + '_' + Date.now() + '.png');
    const file = DriveApp.getFolderById(SIGNATURE_FOLDER_ID).createFile(blob);
    signatureUrl = file.getUrl();
  }

  sh.appendRow([
    new Date(),
    data.agreementType || '',
    data.athleteName || '',
    data.athleteDob || '',
    data.representativeName || '',
    data.representativeRole || '',
    data.phone || '',
    data.email || '',
    data.rulesVersion || '2026/27',
    signatureUrl
  ]);

  return json_({ok:true,type:'agreement'});
}

function savePaymentReceipt_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(PAYMENT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PAYMENT_SHEET);
    sh.appendRow([
      'Дата/час','ID квитанції','ПІБ спортсмена','Сума',
      'Назва файлу','MIME','Посилання на квитанцію','Статус','Коментар'
    ]);
  }

  if (!data.athleteName || !data.receiptDataUrl) {
    throw new Error('Athlete name and receipt file are required.');
  }

  const now = new Date();
  const receiptId = 'RCPT-2027-' +
    Utilities.formatDate(now, Session.getScriptTimeZone() || 'Etc/GMT', 'yyyyMMdd-HHmmss') +
    '-' + Math.floor(1000 + Math.random() * 9000);

  const parts = String(data.receiptDataUrl).split(',');
  const raw = Utilities.base64Decode(parts[1] || '');
  const mime = data.mimeType || 'application/octet-stream';
  const originalName = data.fileName || 'receipt';
  const ext = extension_(originalName, mime);
  const safeName = safe_(data.athleteName);
  const blob = Utilities.newBlob(raw, mime, receiptId + '_' + safeName + ext);
  const file = DriveApp.getFolderById(RECEIPT_FOLDER_ID).createFile(blob);

  sh.appendRow([
    now,
    receiptId,
    data.athleteName || '',
    data.amount || '',
    originalName,
    mime,
    file.getUrl(),
    'Отримано',
    data.comment || ''
  ]);

  return json_({ok:true,type:'paymentReceipt',receiptId:receiptId});
}


function saveCompulsoryForm_(data) {
  log_('compulsoryForm', 'START', data, 'OK', 'Request received');

  const sportDeadline = new Date('2027-01-01T23:59:59+02:00');
  if (new Date() > sportDeadline) {
    throw new Error('SPORT compulsory form deadline has passed.');
  }

  if (!data.athleteName || !data.ageCategory || !data.category || !data.apparatus) {
    throw new Error('Required athlete fields are missing.');
  }

  const elements = Array.isArray(data.elements) ? data.elements : [];
  if (!elements.length) throw new Error('No compulsory elements selected.');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(COMPULSORY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(COMPULSORY_SHEET);
    sh.appendRow(['ПІБ спортсмена','Снаряд','Категорія','PDF файл']);
  }

  const now = new Date();
  const formId = 'COMP-2027-' +
    Utilities.formatDate(now, Session.getScriptTimeZone() || 'Etc/GMT', 'yyyyMMdd-HHmmss') +
    '-' + Math.floor(1000 + Math.random() * 9000);

  log_('compulsoryForm', 'HTML_PDF_CREATE', data, 'OK', formId);

  let availableImages = 0;
  const elementRows = elements.map(function(el, index) {
    let imageHtml = '<div class="no-image">—</div>';

    if (el.driveId) {
      try {
        const blob = DriveApp.getFileById(el.driveId).getBlob();
        const mime = blob.getContentType() || 'image/png';
        const b64 = Utilities.base64Encode(blob.getBytes());
        imageHtml = '<img src="data:' + escapeHtml_(mime) + ';base64,' + b64 + '" alt="' +
          escapeHtml_(el.code || '') + '">';
        availableImages++;
      } catch (imageErr) {
        imageHtml = '<div class="no-image">Зображення недоступне</div>';
        log_('compulsoryForm', 'IMAGE_' + (el.code || index), data, 'WARN', String(imageErr));
      }
    }

    const title = el.title ? '<b>' + escapeHtml_(el.title) + '</b>' : '';
    const description = el.description ? '<div class="desc">' + escapeHtml_(el.description) + '</div>' : '';

    return '<tr>' +
      '<td class="num">' + (index + 1) + '</td>' +
      '<td class="code">' + escapeHtml_(el.code || '') + '</td>' +
      '<td class="visual">' + imageHtml + '</td>' +
      '<td class="copy">' + title + description + '</td>' +
      '</tr>';
  }).join('');

  const html =
    '<!doctype html><html><head><meta charset="utf-8"><style>' +
    '@page{size:A4 portrait;margin:9mm}' +
    '*{box-sizing:border-box}' +
    'body{font-family:Arial,sans-serif;color:#111;margin:0;font-size:10pt}' +
    'h1{text-align:center;font-size:18pt;margin:0 0 6mm}' +
    '.meta{display:grid;grid-template-columns:42mm 1fr;border:1px solid #777;margin-bottom:5mm}' +
    '.meta div{padding:2.4mm 3mm;border-bottom:1px solid #aaa}' +
    '.meta div:nth-last-child(-n+2){border-bottom:0}' +
    '.label{font-weight:700;background:#f2f2f2}' +
    'table{width:100%;border-collapse:collapse;table-layout:fixed}' +
    'th,td{border:1px solid #777;padding:2mm;vertical-align:top}' +
    'th{background:#e8e8e8;text-align:center;font-size:8.5pt}' +
    'tr{page-break-inside:avoid}' +
    '.num{width:8mm;text-align:center}' +
    '.code{width:17mm;text-align:center;font-weight:700}' +
    '.visual{width:38mm;text-align:center;vertical-align:middle}' +
    '.visual img{display:block;max-width:34mm;max-height:31mm;width:auto;height:auto;margin:auto}' +
    '.no-image{font-size:7pt;color:#777}' +
    '.copy{font-size:8.5pt;line-height:1.25}' +
    '.copy b{font-size:9pt}' +
    '.desc{white-space:pre-line;margin-top:1mm}' +
    '.foot{margin-top:3mm;text-align:right;color:#777;font-size:7pt}' +
    '</style></head><body>' +
    '<h1>PROSTO CHEMP — ОБОВ\'ЯЗКОВІ ЕЛЕМЕНТИ</h1>' +
    '<div class="meta">' +
      '<div class="label">СПОРТСМЕН</div><div>' + escapeHtml_(data.athleteName) + '</div>' +
      '<div class="label">ВІКОВА КАТЕГОРІЯ</div><div>' + escapeHtml_(data.ageCategory) + '</div>' +
      '<div class="label">РОЗРЯД</div><div>' + escapeHtml_(data.category) + '</div>' +
      '<div class="label">СНАРЯД</div><div>' + escapeHtml_(data.apparatus) + '</div>' +
    '</div>' +
    '<table><thead><tr><th>№</th><th>Код</th><th>Візуальний приклад</th><th>Назва та опис</th></tr></thead>' +
    '<tbody>' + elementRows + '</tbody></table>' +
    '<div class="foot">PROSTO CHEMP · форма сформована автоматично · ' + escapeHtml_(formId) + '</div>' +
    '</body></html>';

  const pdfName = formId + '_' + safe_(data.athleteName) + '.pdf';
  const pdfBlob = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF).setName(pdfName);
  const pdfFile = DriveApp.getFolderById(COMPULSORY_FOLDER_ID).createFile(pdfBlob);

  log_(
    'compulsoryForm',
    'PDF_CREATED',
    data,
    'OK',
    pdfFile.getUrl() + ' | Images: ' + availableImages + '/' + elements.length
  );

  sh.appendRow([
    data.athleteName || '',
    data.apparatus || '',
    data.category || '',
    pdfFile.getUrl()
  ]);
  log_('compulsoryForm', 'SHEET_SAVED', data, 'OK', 'Row appended');

  return json_({ok:true,type:'compulsoryForm',formId:formId,pdfUrl:pdfFile.getUrl()});
}

function saveArtRoutineDescription_(data) {
  const artDeadline = new Date('2027-01-01T23:59:59+02:00');
  if (new Date() > artDeadline) {
    throw new Error('ART routine description deadline has passed.');
  }

  if (!data.athleteName || !data.category || !data.ageCategory || !data.apparatus || !data.routineDescription) {
    throw new Error('All ART description fields are required.');
  }

  data.athleteName = String(data.athleteName).trim();
  data.category = String(data.category).trim();
  data.ageCategory = String(data.ageCategory).trim();
  data.apparatus = String(data.apparatus).trim();
  data.routineDescription = String(data.routineDescription).trim();

  if (data.routineDescription.length > 100) {
    throw new Error('ART routine description must be 100 characters or fewer.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(HOST_SHEET);
    if (!sh) {
      sh = ss.insertSheet(HOST_SHEET);
      sh.appendRow(['ПІБ','Категорія','Вікова категорія','Снаряд','Опис номера для ведучої']);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,5).setFontWeight('bold');
    }

    const norm = function(v) {
      return String(v || '').trim().toLocaleLowerCase('uk-UA').replace(/\s+/g,' ');
    };
    const key = [
      norm(data.athleteName),
      norm(data.category),
      norm(data.ageCategory),
      norm(data.apparatus)
    ].join('|');

    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const rows = sh.getRange(2,1,lastRow-1,5).getValues();
      for (let i = 0; i < rows.length; i++) {
        const rowKey = [norm(rows[i][0]),norm(rows[i][1]),norm(rows[i][2]),norm(rows[i][3])].join('|');
        if (rowKey === key) {
          const targetRow = i + 2;
          sh.getRange(targetRow,1,1,5).setValues([[
            data.athleteName,
            data.category,
            data.ageCategory,
            data.apparatus,
            data.routineDescription
          ]]);
          return {ok:true,type:'artRoutineDescription',action:'updated',row:targetRow};
        }
      }
    }

    sh.appendRow([
      data.athleteName,
      data.category,
      data.ageCategory,
      data.apparatus,
      data.routineDescription
    ]);
    return {ok:true,type:'artRoutineDescription',action:'created',row:sh.getLastRow()};
  } finally {
    lock.releaseLock();
  }
}

function log_(type, stage, data, status, details) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(LOG_SHEET);
      sh.appendRow(['Дата/час','Тип','Етап','ПІБ','Снаряд','Категорія','Статус','Помилка/деталі']);
    }
    sh.appendRow([
      new Date(),
      type || '',
      stage || '',
      (data && data.athleteName) || '',
      (data && data.apparatus) || '',
      (data && data.category) || '',
      status || '',
      details || ''
    ]);
  } catch (ignore) {}
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function safe_(value) {
  return String(value || '')
    .replace(/[^a-zA-Z0-9а-яА-ЯіїєґІЇЄҐ _-]/g,'_')
    .trim()
    .slice(0,80) || 'file';
}

function extension_(name, mime) {
  const match = String(name || '').match(/(\.[a-zA-Z0-9]{1,8})$/);
  if (match) return match[1];
  if (mime === 'application/pdf') return '.pdf';
  if (mime === 'image/png') return '.png';
  if (mime === 'image/webp') return '.webp';
  return '.jpg';
}

function jsonp_(obj, callback) {
  const cb = String(callback || '');
  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(cb)) {
    return json_({ok:false,error:'Invalid callback'});
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(obj) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
