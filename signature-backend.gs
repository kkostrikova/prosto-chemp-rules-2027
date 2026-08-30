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

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (data.submissionType === 'paymentReceipt') {
      return savePaymentReceipt_(data);
    }

    if (data.submissionType === 'compulsoryForm') {
      return saveCompulsoryForm_(data);
    }

    return saveAgreement_(data);
  } catch (err) {
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
  if (!data.athleteName || !data.ageCategory || !data.category || !data.apparatus) {
    throw new Error('Required athlete fields are missing.');
  }
  const elements = Array.isArray(data.elements) ? data.elements : [];
  if (!elements.length) throw new Error('No compulsory elements selected.');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(COMPULSORY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(COMPULSORY_SHEET);
    sh.appendRow([
      'Дата/час','ID форми','ПІБ спортсмена','Вікова категорія','Розряд','Снаряд',
      'Порядок 1','Порядок 2','Порядок 3','Порядок 4','Порядок 5','Порядок 6',
      'Кількість груп','Версія правил','Статус','Посилання на PDF'
    ]);
  }

  const now = new Date();
  const formId = 'COMP-2027-' +
    Utilities.formatDate(now, Session.getScriptTimeZone() || 'Etc/GMT', 'yyyyMMdd-HHmmss') +
    '-' + Math.floor(1000 + Math.random() * 9000);

  const doc = DocumentApp.create(formId + ' — ' + safe_(data.athleteName));
  const body = doc.getBody();
  body.clear();
  body.appendParagraph('PROSTO CHEMP').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('ОБОВ\'ЯЗКОВІ ЕЛЕМЕНТИ').setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const info = body.appendTable([
    ['СПОРТСМЕН', data.athleteName],
    ['ВІКОВА КАТЕГОРІЯ', data.ageCategory],
    ['РОЗРЯД', data.category],
    ['НАПРЯМОК', data.apparatus]
  ]);
  info.setBorderWidth(1);

  body.appendParagraph('');
  body.appendParagraph('ПОРЯДОК ВИКОНАННЯ').setHeading(DocumentApp.ParagraphHeading.HEADING2);

  elements.forEach(function(el, index) {
    const rowTitle = (index + 1) + '. ' + (el.code || '') + (el.title ? ' — ' + el.title : '');
    body.appendParagraph(rowTitle).setHeading(DocumentApp.ParagraphHeading.HEADING3);

    if (el.driveId) {
      try {
        const blob = DriveApp.getFileById(el.driveId).getBlob();
        const image = body.appendImage(blob);
        const w = image.getWidth(), h = image.getHeight();
        if (w > 220) {
          const ratio = 220 / w;
          image.setWidth(220).setHeight(Math.round(h * ratio));
        }
      } catch (imageErr) {}
    }

    if (el.description) body.appendParagraph(el.description);
    body.appendParagraph('');
  });

  body.appendParagraph('Форма сформована автоматично на платформі PROSTO CHEMP.');
  doc.saveAndClose();

  const pdfName = formId + '_' + safe_(data.athleteName) + '.pdf';
  const pdfBlob = DriveApp.getFileById(doc.getId()).getBlob().getAs(MimeType.PDF).setName(pdfName);
  const pdfFile = DriveApp.getFolderById(COMPULSORY_FOLDER_ID).createFile(pdfBlob);
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  const ordered = elements.map(function(el){ return (el.code || '') + (el.title ? ' — ' + el.title : ''); });
  while (ordered.length < 6) ordered.push('');

  sh.appendRow([
    now,
    formId,
    data.athleteName || '',
    data.ageCategory || '',
    data.category || '',
    data.apparatus || '',
    ordered[0], ordered[1], ordered[2], ordered[3], ordered[4], ordered[5],
    elements.length,
    data.rulesVersion || '2026/27',
    'Отримано',
    pdfFile.getUrl()
  ]);

  return json_({ok:true,type:'compulsoryForm',formId:formId,pdfUrl:pdfFile.getUrl()});
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

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
