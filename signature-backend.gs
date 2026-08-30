/**
 * PROSTO CHEMP - signature collector backend (Google Apps Script)
 *
 * Setup:
 * 1) Create a Google Sheet and a Drive folder for signature images.
 * 2) Replace SHEET_ID and FOLDER_ID below.
 * 3) Deploy as Web App: Execute as "Me"; access "Anyone".
 * 4) Put the deployment URL into the website's SIGNATURE_ENDPOINT constant.
 */

const SHEET_ID = '13ncm5NW4SmAYWL6UALzEgugty4yxIH_IlWK4QKRZoYQ';
const FOLDER_ID = '1D53NN34oJcOBi7SMm_2GPvV27H8hFQeM';
const SHEET_NAME = 'Signatures';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
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
      const safeName = String(data.athleteName || data.representativeName || 'signature')
        .replace(/[^a-zA-Z0-9а-яА-ЯіїєґІЇЄҐ _-]/g,'_');
      const blob = Utilities.newBlob(raw, 'image/png', safeName + '_' + Date.now() + '.png');
      const file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
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

    return ContentService
      .createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
