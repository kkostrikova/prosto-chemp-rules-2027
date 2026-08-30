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
  return json_({ok:true,service:'PROSTO CHEMP backend',version:'13'});
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

  log_('compulsoryForm', 'SHEET_PDF_CREATE', data, 'OK', formId);

  // Build a temporary Google Sheet instead of Google Docs.
  // This avoids the DocumentApp permission issue while keeping images private in Drive.
  const temp = SpreadsheetApp.create(formId + ' — ' + safe_(data.athleteName));
  const out = temp.getSheets()[0];
  out.setName('Форма');

  out.setColumnWidth(1, 42);
  out.setColumnWidth(2, 75);
  out.setColumnWidth(3, 170);
  out.setColumnWidth(4, 430);

  out.getRange('A1:D1').merge()
    .setValue('PROSTO CHEMP — ОБОВ\'ЯЗКОВІ ЕЛЕМЕНТИ')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center');

  out.getRange('A3:B3').merge().setValue('СПОРТСМЕН').setFontWeight('bold');
  out.getRange('C3:D3').merge().setValue(data.athleteName);
  out.getRange('A4:B4').merge().setValue('ВІКОВА КАТЕГОРІЯ').setFontWeight('bold');
  out.getRange('C4:D4').merge().setValue(data.ageCategory);
  out.getRange('A5:B5').merge().setValue('РОЗРЯД').setFontWeight('bold');
  out.getRange('C5:D5').merge().setValue(data.category);
  out.getRange('A6:B6').merge().setValue('СНАРЯД').setFontWeight('bold');
  out.getRange('C6:D6').merge().setValue(data.apparatus);

  out.getRange('A8:D8').setValues([['№','Код','Візуальний приклад','Назва та опис']])
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#e5e5e5');

  let row = 9;
  let insertedImages = 0;

  elements.forEach(function(el, index) {
    const title = (el.title ? el.title + '\n' : '') + (el.description || '');
    out.getRange(row,1).setValue(index + 1).setHorizontalAlignment('center');
    out.getRange(row,2).setValue(el.code || '').setFontWeight('bold').setHorizontalAlignment('center');
    out.getRange(row,4).setValue(title).setWrap(true).setVerticalAlignment('top');

    if (el.driveId) {
      try {
        const blob = DriveApp.getFileById(el.driveId).getBlob();
        const image = out.insertImage(blob,3,row);
        const w = image.getWidth();
        const h = image.getHeight();
        if (w > 150) {
          const ratio = 150 / w;
          image.setWidth(150).setHeight(Math.max(80, Math.round(h * ratio)));
        }
        insertedImages++;
      } catch (imageErr) {
        out.getRange(row,3).setValue('[Зображення недоступне]').setWrap(true);
        log_('compulsoryForm', 'IMAGE_' + (el.code || index), data, 'WARN', String(imageErr));
      }
    }

    out.setRowHeight(row, 170);
    row++;
  });

  const lastRow = Math.max(8,row-1);
  out.getRange(3,1,lastRow-2,4).setBorder(true,true,true,true,true,true);
  out.getRange(1,1,lastRow,4).setVerticalAlignment('top');
  out.setFrozenRows(0);

  SpreadsheetApp.flush();
  log_('compulsoryForm', 'TEMP_SHEET_READY', data, 'OK', 'Images inserted: ' + insertedImages + '/' + elements.length);

  const sourceFile = DriveApp.getFileById(temp.getId());
  const pdfName = formId + '_' + safe_(data.athleteName) + '.pdf';
  const pdfBlob = sourceFile.getAs(MimeType.PDF).setName(pdfName);
  const pdfFile = DriveApp.getFolderById(COMPULSORY_FOLDER_ID).createFile(pdfBlob);
  log_('compulsoryForm', 'PDF_CREATED', data, 'OK', pdfFile.getUrl());

  sh.appendRow([
    data.athleteName || '',
    data.apparatus || '',
    data.category || '',
    pdfFile.getUrl()
  ]);
  log_('compulsoryForm', 'SHEET_SAVED', data, 'OK', 'Row appended');

  try { sourceFile.setTrashed(true); } catch (trashErr) {}

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
