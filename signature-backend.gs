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
const HUNT_SHEET = 'Гра P4';
const NG_SHEET = 'New Generation';
const OSCAR_SHEET = 'OSCAR номінанти';

// Папки для фото створюються поряд із папкою квитанцій, щоб не заводити ID руками.
const NG_FOLDER_NAME = 'PROSTO CHEMP — New Generation';
const OSCAR_FOLDER_NAME = 'PROSTO CHEMP — OSCAR';

const OSCAR_APPARATUS = ['Пілон','Кільце','Полотна','Оригінальний жанр'];


// Пошук семи знаків P4. Час рахує сервер: браузер не може надіслати вигаданий
// результат, бо не він вирішує, скільки часу минуло.
const HUNT_MARKS = 7;
const HUNT_MIN_SECONDS = 20;    // швидше фізично не проскролити сторінку і не натиснути 7 знаків
const HUNT_MAX_SECONDS = 3600;  // довші забіги вважаємо покинутими
const HUNT_TOP = 10;


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
  if (p.action === 'huntStart') {
    try {
      return jsonp_(huntStart_(), p.callback);
    } catch (err) {
      return jsonp_({ok:false,error:String(err)}, p.callback);
    }
  }

  if (p.action === 'huntFinish') {
    try {
      return jsonp_(huntFinish_(p.token, p.nickname), p.callback);
    } catch (err) {
      return jsonp_({ok:false,error:String(err)}, p.callback);
    }
  }

  if (p.action === 'huntTop') {
    try {
      return jsonp_(huntTop_(), p.callback);
    } catch (err) {
      return jsonp_({ok:false,error:String(err)}, p.callback);
    }
  }

  return json_({ok:true,service:'PROSTO CHEMP backend',version:'15'});
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

    if (data.submissionType === 'newGeneration') {
      return json_(saveNewGeneration_(data));
    }

    if (data.submissionType === 'oscarNominee') {
      return json_(saveOscarNominee_(data));
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
    '.num{width:6mm;text-align:center;font-size:8pt}' +
    '.code{width:12mm;text-align:center;font-weight:700;font-size:8.5pt}' +
    '.visual{width:46mm;text-align:center;vertical-align:middle}' +
    '.visual img{display:block;max-width:43mm;max-height:48mm;width:auto;height:auto;margin:auto}' +
    '.no-image{font-size:7pt;color:#777}' +
    '.copy{font-size:8.8pt;line-height:1.28}' +
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

// Знаходить папку за назвою поряд із папкою квитанцій або створює її.
function mediaFolder_(name) {
  const parents = DriveApp.getFolderById(RECEIPT_FOLDER_ID).getParents();
  const root = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const found = root.getFoldersByName(name);
  return found.hasNext() ? found.next() : root.createFolder(name);
}

function savePhoto_(dataUrl, folderName, baseName, mime) {
  if (!dataUrl) return '';
  const parts = String(dataUrl).split(',');
  const raw = Utilities.base64Decode(parts[1] || '');
  const type = mime || 'image/jpeg';
  const blob = Utilities.newBlob(raw, type, safe_(baseName) + '_' + Date.now() + extension_('', type));
  return mediaFolder_(folderName).createFile(blob).getUrl();
}

function studio_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().slice(0, 60);
}

// Реєстрація в дитячу суддівську панель New Generation.
function saveNewGeneration_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(NG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(NG_SHEET);
    sh.appendRow(['Дата/час','Ім’я дитини','Вік','Студія','Контакт батьків','Згода батьків','Фото','Статус']);
  }

  const name = String(data.childName || '').trim().slice(0, 80);
  const age = Number(data.childAge) || 0;
  const studio = studio_(data.studio);
  if (!name || !studio) throw new Error('Вкажіть ім’я та студію.');
  if (age < 6 || age > 17) throw new Error('Вік має бути від 6 до 17 років.');
  if (!data.parentConsent) throw new Error('Потрібна згода батьків.');
  if (!data.photoDataUrl) throw new Error('Додайте фото.');

  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase() === name.toLowerCase() && studio_(rows[i][3]).toLowerCase() === studio.toLowerCase()) {
      throw new Error('Ця дитина вже зареєстрована.');
    }
  }

  const photoUrl = savePhoto_(data.photoDataUrl, NG_FOLDER_NAME, name, data.mimeType);
  sh.appendRow([new Date(), name, age, studio, String(data.parentContact || '').slice(0,120), 'так', photoUrl, 'На модерації']);
  return {ok:true, type:'newGeneration'};
}

// Номінація на OSCAR: студія подає спортсмена в снаряді або тренера.
function saveOscarNominee_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(OSCAR_SHEET);
  if (!sh) {
    sh = ss.insertSheet(OSCAR_SHEET);
    sh.appendRow(['Дата/час','Тип','Снаряд','Ім’я номінанта','Студія','Контакт студії','Фото','Статус']);
  }

  const kind = data.kind === 'coach' ? 'coach' : 'athlete';
  const name = String(data.nomineeName || '').trim().slice(0, 80);
  const studio = studio_(data.studio);
  const apparatus = kind === 'athlete' ? String(data.apparatus || '') : '';
  if (!name || !studio) throw new Error('Вкажіть ім’я номінанта та студію.');
  if (kind === 'athlete' && OSCAR_APPARATUS.indexOf(apparatus) === -1) throw new Error('Оберіть снаряд.');

  // одна номінація на студію в кожній категорії
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const sameStudio = studio_(rows[i][4]).toLowerCase() === studio.toLowerCase();
    const sameKind = String(rows[i][1]) === kind;
    const sameApparatus = String(rows[i][2]) === apparatus;
    if (sameStudio && sameKind && sameApparatus) {
      throw new Error(kind === 'coach'
        ? 'Від цієї студії тренера вже подано.'
        : 'Від цієї студії вже подано спортсмена в цьому снаряді.');
    }
  }

  const photoUrl = savePhoto_(data.photoDataUrl, OSCAR_FOLDER_NAME, name, data.mimeType);
  sh.appendRow([new Date(), kind, apparatus, name, studio, String(data.studioContact || '').slice(0,120), photoUrl, 'На модерації']);
  return {ok:true, type:'oscarNominee'};
}

function huntSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(HUNT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(HUNT_SHEET);
    sh.appendRow(['Дата/час','Instagram','Час, сек','Час','Спроб']);
  }
  return sh;
}

// Нік інстаграму, а не довільний текст: safe_ вирізав би крапки, дозволені в нікнеймах.
function huntNick_(value) {
  const raw = String(value || '').trim().replace(/^https?:\/\/(www\.)?instagram\.com\//i, '').replace(/^@/, '').replace(/\/.*$/, '');
  return /^[A-Za-z0-9._]{1,30}$/.test(raw) ? raw : '';
}

function huntFormat_(seconds) {
  const m = Math.floor(seconds / 60), s = seconds % 60;
  return m + ':' + ('0' + s).slice(-2);
}

// Видає токен забігу. Момент старту зберігається на сервері, браузер його не бачить
// і не може змінити.
function huntStart_() {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('hunt_' + token, String(Date.now()), HUNT_MAX_SECONDS);
  return {ok:true, token:token};
}

function huntFinish_(token, nickname) {
  const cache = CacheService.getScriptCache();
  const key = 'hunt_' + String(token || '');
  const startedAt = cache.get(key);
  if (!startedAt) {
    return {ok:false, error:'expired', message:'Забіг не знайдено або він застарів. Почни пошук спочатку.'};
  }
  cache.remove(key); // токен одноразовий

  const seconds = Math.round((Date.now() - Number(startedAt)) / 1000);
  if (seconds < HUNT_MIN_SECONDS) {
    log_('huntP4','finish',{athleteName:nickname},'REJECTED','seconds=' + seconds);
    return {ok:false, error:'too_fast', message:'Такий час неможливий. Спробуй ще раз.'};
  }
  if (seconds > HUNT_MAX_SECONDS) {
    return {ok:false, error:'too_slow', message:'Забіг тривав надто довго і не зарахований.'};
  }

  const name = huntNick_(nickname);
  if (!name) {
    return {ok:false, error:'no_nickname', message:'Вкажи свій Instagram — латиниця, крапки й підкреслення.'};
  }

  const sh = huntSheet_();
  const rows = sh.getDataRange().getValues();
  let targetRow = 0, previous = 0, attempts = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase() === name.toLowerCase()) {
      targetRow = i + 1;
      previous = Number(rows[i][2]) || 0;
      attempts = Number(rows[i][4]) || 1;
      break;
    }
  }

  if (targetRow) {
    // один рядок на нікнейм — лишаємо найкращий час
    if (previous && seconds >= previous) {
      sh.getRange(targetRow, 5).setValue(attempts + 1);
      return {ok:true, seconds:seconds, best:previous, improved:false, top:huntTop_().top};
    }
    sh.getRange(targetRow, 1, 1, 5).setValues([[
      new Date(), name, seconds, huntFormat_(seconds), attempts + 1
    ]]);
    return {ok:true, seconds:seconds, best:seconds, improved:true, top:huntTop_().top};
  }

  sh.appendRow([new Date(), name, seconds, huntFormat_(seconds), 1]);
  return {ok:true, seconds:seconds, best:seconds, improved:true, top:huntTop_().top};
}

// Публічна таблиця: тільки нік і час.
function huntTop_() {
  const rows = huntSheet_().getDataRange().getValues().slice(1);
  const top = rows
    .filter(function (r) { return r[1] && Number(r[2]) > 0; })
    .map(function (r) { return {nickname:String(r[1]), seconds:Number(r[2]), time:huntFormat_(Number(r[2]))}; })
    .sort(function (a, b) { return a.seconds - b.seconds; })
    .slice(0, HUNT_TOP);
  return {ok:true, top:top, marks:HUNT_MARKS};
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
