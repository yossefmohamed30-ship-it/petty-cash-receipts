/***** إعدادات *****/
const CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',
  MODEL: 'gemini-2.5-flash',          // غيّره لو عايز موديل تاني
  FOLDER_NAME: 'Petty Cash Receipts',  // فولدر الإيصالات على الدرايف
  DATE_FORMAT: 'yyyy-mm-dd',
};

/***** المنيو *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧾 المصروفات')
    .addItem('➕ إضافة إيصال', 'openSidebar')
    .addSeparator()
    .addItem('🔑 إعداد مفتاح Gemini', 'setApiKey')
    .addToUi();
}

function openSidebar() {
  const t = HtmlService.createTemplateFromFile('ReceiptForm');
  t.defaultSheet = SpreadsheetApp.getActiveSheet().getName();
  t.accessKey = getKey_();
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('إضافة إيصال'));
}

// نفس الصفحة كـ Web App (تفتحها في تاب أو من الموبايل)
function doGet(e) {
  if (!e || !e.parameter || e.parameter.k !== getKey_()) {
    return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;text-align:center;margin-top:40px">⛔ اللينك ده مش صح</h2>');
  }
  const t = HtmlService.createTemplateFromFile('ReceiptForm');
  t.defaultSheet = '';
  t.accessKey = e.parameter.k;
  return t.evaluate()
    .setTitle('إضافة إيصال')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('مفتاح Gemini', 'الصق الـ API key من aistudio.google.com/apikey', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  ui.alert('✅ اتحفظ المفتاح');
}

/***** Helpers *****/
// مفتاح سري في اللينك — محدش يقدر يستخدم الصفحة من غيره
function getKey_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('ACCESS_KEY');
  if (!k) { k = Utilities.getUuid().replace(/-/g, ''); p.setProperty('ACCESS_KEY', k); }
  return k;
}

function checkKey_(k) {
  if (k !== getKey_()) throw new Error('⛔ مش مسموح');
}

// شغّلها من المحرر عشان تشوف المفتاح في الـ Execution log
function showAccessKey() {
  Logger.log('KEY=' + getKey_());
}

function ss_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  const it = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function lastRowInCol_(sheet, col) {
  const last = sheet.getLastRow();
  if (last === 0) return 0;
  const vals = sheet.getRange(1, col, last, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '' && vals[i][0] !== null) return i + 1;
  }
  return 0;
}

/***** بيستدعيها الـ HTML *****/

// التابات اللي فيها نفس الهيدر (Date / Amount / Description / link)
function getSheetsList(k) {
  checkKey_(k);
  const sheets = ss_().getSheets();
  const matching = sheets.filter(sh => {
    if (sh.getLastColumn() < 4) return false;
    const h = sh.getRange(1, 1, 1, 4).getValues()[0].map(v => String(v).trim().toLowerCase());
    return h[1] === 'amount' && h[3] === 'link';
  }).map(sh => sh.getName());
  return matching.length ? matching : sheets.map(sh => sh.getName());
}

// يرفع الصورة على الدرايف ويحللها
function uploadAndAnalyze(data) {
  checkKey_(data.k);
  const bytes = Utilities.base64Decode(data.base64);
  const blob = Utilities.newBlob(bytes, data.mimeType, data.fileName || ('receipt_' + Date.now()));
  const file = getFolder_().createFile(blob);

  const result = { fileId: file.getId(), url: file.getUrl() };
  try {
    Object.assign(result, analyze_(data.base64, data.mimeType));
  } catch (e) {
    result.aiError = e.message;
  }
  return result;
}

function analyze_(base64, mimeType) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return ocrAnalyze_(base64, mimeType); // من غير مفتاح: OCR بتاع جوجل درايف (ببلاش)

  const prompt =
    'This is an Egyptian receipt, invoice, or money-transfer screenshot (InstaPay, OPay, Vodafone Cash, bank app, shop receipt, etc.). ' +
    'Extract the TOTAL amount actually paid/transferred in EGP (the final total, not subtotal or individual items; ignore fees unless included in the total). ' +
    'Convert Arabic-Indic digits (٠١٢٣٤٥٦٧٨٩) to normal digits. ' +
    'Return ONLY JSON: {"amount": number|null, "date": "YYYY-MM-DD"|null, "merchant": string|null, ' +
    '"type": "out"|"in", "summary": string}. ' +
    '"type" is "in" only if the image clearly shows money RECEIVED; otherwise "out". ' +
    '"summary" is a very short Arabic (Egyptian) description of what it is, max 6 words.';

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: mimeType, data: base64 } },
      ],
    }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0 },
  };

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG.MODEL + ':generateContent';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('Gemini error ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
  }
  const json = JSON.parse(res.getContentText());
  const text = json.candidates[0].content.parts.map(p => p.text || '').join('');
  const out = JSON.parse(text.replace(/```json|```/g, '').trim());

  return {
    amount: typeof out.amount === 'number' ? out.amount : (out.amount ? Number(String(out.amount).replace(/[^\d.]/g, '')) : null),
    date: out.date || null,
    merchant: out.merchant || '',
    type: out.type === 'in' ? 'in' : 'out',
    summary: out.summary || '',
  };
}

// يسجل الصف في الشيت
function saveEntry(e) {
  checkKey_(e.k);
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sheet = ss_().getSheetByName(e.sheetName);
    if (!sheet) throw new Error('التاب "' + e.sheetName + '" مش موجودة');

    const [y, m, d] = String(e.date).split('-').map(Number);
    const date = new Date(y, m - 1, d);
    const amount = Math.abs(Number(e.amount)) * (e.type === 'in' ? 1 : -1);

    const row = lastRowInCol_(sheet, 1) + 1;
    sheet.getRange(row, 1, 1, 3).setValues([[date, amount, e.description]]);
    sheet.getRange(row, 1).setNumberFormat(CONFIG.DATE_FORMAT);
    sheet.getRange(row, 4).setFormula('=HYPERLINK("' + e.url + '","🧾 الإيصال")');

    // اسم الملف على الدرايف يبقى واضح
    try {
      const f = DriveApp.getFileById(e.fileId);
      const ext = (f.getName().match(/\.[a-z0-9]+$/i) || [''])[0];
      f.setName(e.date + ' | ' + amount + ' | ' + e.description + ext);
    } catch (err) {}

    return { row: row, sheet: e.sheetName };
  } finally {
    lock.releaseLock();
  }
}

// لو لغيت قبل ما تسجل — الصورة تروح الـ Trash
function discardUpload(fileId, k) {
  checkKey_(k);
  if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
}

/***** OCR من غير مفتاح — جوجل درايف بيقرا الصورة *****/
function ocrAnalyze_(base64, mimeType) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, 'ocr_tmp_' + Date.now());
  const doc = Drive.Files.create({ name: 'ocr_tmp_' + Date.now(), mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'ar' });
  let text = '';
  try {
    text = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + doc.id + '/export?mimeType=text/plain',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }).getContentText();
  } finally {
    try { DriveApp.getFileById(doc.id).setTrashed(true); } catch (err) {}
  }
  return parseReceiptText_(text);
}

function parseReceiptText_(raw) {
  const t = String(raw)
    .replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
    .replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d))
    .replace(/٫/g, '.').replace(/٬/g, ',');

  // التاريخ
  let date = null, m;
  if ((m = t.match(/(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/))) date = m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]);
  else if ((m = t.match(/(\d{1,2})[-\/.](\d{1,2})[-\/.](20\d{2})/))) date = m[3] + '-' + pad_(m[2]) + '-' + pad_(m[1]);

  // نشيل التواريخ والساعات والأرقام الطويلة (موبايلات / أرقام عمليات)
  const clean = t
    .replace(/\d{1,4}[-\/.]\d{1,2}[-\/.]\d{1,4}/g, ' ')
    .replace(/\d{1,2}:\d{2}(:\d{2})?/g, ' ')
    .replace(/\d{9,}/g, ' ');
  const lines = clean.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const numRe = /\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?/g;
  const nums = s => (s.match(numRe) || []).map(x => Number(x.replace(/,/g, ''))).filter(n => n > 0 && n < 10000000);

  const strong = /(الإجمالي|الاجمالي|إجمالي|اجمالي|الصافي|المطلوب|total|grand)/i;
  const medium = /(المبلغ|مبلغ|amount|paid|المدفوع|EGP|جنيه|ج\.م|L\.E|LE)/i;

  let amount = null;
  for (const re of [strong, medium]) {
    const cands = [];
    lines.forEach((l, i) => {
      if (re.test(l)) {
        cands.push(...nums(l));
        if (!nums(l).length && lines[i + 1]) cands.push(...nums(lines[i + 1]));
      }
    });
    if (cands.length) { amount = Math.max(...cands); break; }
  }
  if (amount === null) {
    const all = nums(clean);
    const dec = all.filter(n => n % 1 !== 0);
    amount = dec.length ? Math.max(...dec) : (all.length ? Math.max(...all) : null);
  }

  const type = /(استلام|استلمت|received|وارد|إيداع|ايداع)/i.test(t) ? 'in' : 'out';
  return { amount: amount, date: date, merchant: (lines.find(l => /[A-Za-z\u0600-\u06FF]{2,}/.test(l)) || '').slice(0, 40), type: type, summary: '' };
}

function pad_(n) { return String(n).padStart(2, '0'); }

/***** API لصفحة GitHub Pages (بيتنادى من doPost في Code.gs) *****/
function receiptApi_(data) {
  let out;
  try {
    let result;
    if (data.action === 'rcpt_sheets') result = getSheetsList(data.k);
    else if (data.action === 'rcpt_upload') result = uploadAndAnalyze(data);
    else if (data.action === 'rcpt_save') result = saveEntry(data);
    else if (data.action === 'rcpt_discard') result = discardUpload(data.fileId, data.k) || true;
    else throw new Error('unknown action');
    out = { ok: true, result: result };
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
