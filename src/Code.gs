/**
 * נטו הזמנות – צד שרת (Google Apps Script)
 * הקוד מחובר לגיליון "נטו הזמנות – מערכת". כל הנתונים נשמרים בגיליון.
 * התקנה: מדביקים את הקובץ הזה ב-Apps Script, מריצים setup פעם אחת, ואז פורסים כאפליקציית אינטרנט.
 */

const FOLDER_NAME = 'נטו הזמנות';
const SH = { settings: 'הגדרות', products: 'מוצרים', customers: 'לקוחות', orders: 'הזמנות', lines: 'שורות הזמנה', promos: 'מבצעים' };
const ZONES = ['קפוא', 'מצונן', 'יבש', 'חד-פעמי וניקיון'];

/* ================= התקנה ================= */

function setup() {
  const ss = SpreadsheetApp.getActive();
  const folder = appFolder_();
  // תיקיות משנה
  const imgFolder = subFolder_(folder, 'תמונות');
  subFolder_(folder, 'הזמנות PDF');
  subFolder_(folder, 'הקלטות');
  imgFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // פריסת תמונות מקובץ ה-ZIP (פעם אחת)
  const have = {};
  const ex = imgFolder.getFiles();
  while (ex.hasNext()) have[ex.next().getName()] = true;
  const all = folder.getFiles();
  while (all.hasNext()) {
    const z = all.next();
    if (!/^נטו-תמונות.*\.zip$/.test(z.getName())) continue;
    Utilities.unzip(z.getBlob().setContentType('application/zip')).forEach(b => { if (!have[b.getName()]) { imgFolder.createFile(b); have[b.getName()] = true; } });
  }
  const ids = {};
  const it = imgFolder.getFiles();
  while (it.hasNext()) { const f = it.next(); ids[f.getName()] = f.getId(); }
  PropertiesService.getScriptProperties().setProperty('LOGO_ID', ids['logo.png'] || '');

  // מזהי תמונות בגיליון המוצרים
  const ps = ss.getSheetByName(SH.products);
  const pv = ps.getDataRange().getValues();
  const h = pv[0], cImg = h.indexOf('תמונה'), cId = h.indexOf('מזהה תמונה');
  for (let r = 1; r < pv.length; r++) {
    const fn = pv[r][cImg];
    pv[r][cId] = fn && ids[fn] ? ids[fn] : '';
  }
  ps.getRange(1, 1, pv.length, h.length).setValues(pv);

  // סיסמאות ומפתחות אישיים ללקוחות
  const cs = ss.getSheetByName(SH.customers);
  const cv = cs.getDataRange().getValues();
  const ch = cv[0], cPass = ch.indexOf('סיסמה'), cKey = ch.indexOf('מפתח אישי'), cPhone = ch.indexOf('טלפון');
  for (let r = 1; r < cv.length; r++) {
    if (!cv[r][cPass]) {
      const digits = String(cv[r][cPhone] || '').replace(/\D/g, '');
      cv[r][cPass] = digits.length >= 4 ? digits.slice(-4) : String(1000 + Math.floor(Math.random() * 9000));
    }
    if (!cv[r][cKey]) cv[r][cKey] = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  }
  cs.getRange(1, 1, cv.length, ch.length).setNumberFormat('@').setValues(cv);

  // סיסמת מנהל
  if (!setting_('שעת סיכום')) setSetting_('שעת סיכום', '13');
  if (!setting_('סיכום יומי')) setSetting_('סיכום יומי', 'פעיל');
  if (!setting_('כתובת האפליקציה')) setSetting_('כתובת האפליקציה', 'https://shahar1234566-source.github.io/neto/');
  ensureCols_(cs, ['ימי תזכורת', 'שעת תזכורת', 'מגוון']);
  ensureCols_(ps, ['מגוונים', 'מזהה פנימי']);
  const imp = importCatalog_();
  fillInternalIds_();
  if (!setting_('סיסמת מנהל')) setSetting_('סיסמת מנהל', String(100000 + Math.floor(Math.random() * 900000)));

  // תזכורת יומית
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'sendReminders') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendReminders').timeBased().everyHours(1).create();

  SpreadsheetApp.getUi && toast_('ההתקנה הסתיימה' + (imp ? ' · ' + imp : '') + '. עכשיו: פריסה ← ניהול פריסות ← גרסה חדשה.');
  return imp;
}

function toast_(m) { try { SpreadsheetApp.getActive().toast(m, 'נטו הזמנות', 10); } catch (e) {} }

function appFolder_() {
  const f = DriveApp.getFoldersByName(FOLDER_NAME);
  return f.hasNext() ? f.next() : DriveApp.createFolder(FOLDER_NAME);
}
function subFolder_(parent, name) {
  const f = parent.getFoldersByName(name);
  return f.hasNext() ? f.next() : parent.createFolder(name);
}

/* ================= הגשת האפליקציה ================= */

function doGet(e) {
  const files = appFolder_().getFilesByName('app-index.html');
  if (!files.hasNext()) return HtmlService.createHtmlOutput('חסר הקובץ app-index.html בתיקייה "נטו הזמנות"');
  let html = files.next().getBlob().getDataAsString('UTF-8');
  const boot = {
    params: e && e.parameter ? { c: e.parameter.c || '', k: e.parameter.k || '', admin: e.parameter.admin || '' } : {},
    business: setting_('שם העסק') || 'נטו ליאור לב',
    logo: logoUrl_(),
    url: ScriptApp.getService().getUrl()
  };
  html = html.replace('__BOOT__', JSON.stringify(boot).replace(/</g, '\\u003c'));
  return HtmlService.createHtmlOutput(html)
    .setTitle('נטו הזמנות')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function logoUrl_(sz) {
  const id = PropertiesService.getScriptProperties().getProperty('LOGO_ID');
  return id ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w' + (sz || 400) : '';
}
// מזהה מוצר באפליקציה: מק"ט, ואם אין – ברקוד, ואם אין – מזהה פנימי
function pid_(r) {
  const sku = String(r['מק"ט'] || '').trim();
  if (sku) return sku;
  const bc = String(r['ברקוד'] || '').trim();
  if (bc) return 'b' + bc;
  return 'n' + String(r['מזהה פנימי'] || '').trim();
}
function listOf_(s) { return String(s || '').split(',').map(x => x.trim()).filter(Boolean); }
function imgUrl_(id) { return id ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400' : ''; }

/* ================= API (נקרא מהדפדפן) ================= */

function api(action, p) {
  try {
    p = p || {};
    switch (action) {
      case 'login': return login_(p);
      case 'catalog': return catalog_(auth_(p.token, 'c'));
      case 'submit': return submit_(auth_(p.token, 'c'), p);
      case 'adminLogin': return adminLogin_(p);
      case 'adminData': auth_(p.token, 'a'); return adminData_();
      case 'setStatus': auth_(p.token, 'a'); return setStatus_(p.no, p.status);
      case 'setAccess': auth_(p.token, 'a'); return setAccess_(p.code, p.access);
      case 'pdf': { const s = auth_(p.token); return { url: orderPdf_(p.no, s) }; }
      case 'saveRow': auth_(p.token, 'a'); saveRow_(p.kind, p.row, p.data); return adminData_();
      case 'deleteRow': auth_(p.token, 'a'); deleteRow_(p.kind, p.row, p.check); return adminData_();
      case 'saveSettings': auth_(p.token, 'a'); Object.keys(p.data || {}).forEach(k => setSetting_(k, p.data[k])); return adminData_();
      case 'uploadImage': auth_(p.token, 'a'); return uploadImage_(p.b64, p.name);
      case 'bulkProducts': auth_(p.token, 'a'); bulkProducts_(p); return adminData_();
      case 'saveAssorts': auth_(p.token, 'a'); saveAssorts_(p.names || [], p.renames || {}); return adminData_();
    }
    throw new Error('פעולה לא מוכרת');
  } catch (err) {
    return { error: String(err && err.message || err) };
  }
}

function auth_(token, type) {
  const raw = token && CacheService.getScriptCache().get('t_' + token);
  if (!raw) throw new Error('SESSION');
  const s = JSON.parse(raw);
  if (type && s.t !== type) throw new Error('אין הרשאה');
  CacheService.getScriptCache().put('t_' + token, raw, 21600);
  return s;
}
function newToken_(obj) {
  const t = Utilities.getUuid();
  CacheService.getScriptCache().put('t_' + t, JSON.stringify(obj), 21600);
  return t;
}

function login_(p) {
  const code = String(p.code || '').trim().replace(/^0+(?=\d{4})/, '');
  const cust = customers_().find(c =>
    (p.k ? c['מפתח אישי'] === String(p.k) && c['קוד'] === String(p.c) : c['קוד'] === code.padStart(4, '0') && String(c['סיסמה']) === String(p.pass || '').trim()));
  if (!cust) throw new Error(p.k ? 'הקישור האישי לא תקין. בקש קישור חדש.' : 'קוד לקוח או סיסמה שגויים');
  if (!truthy_(cust['גישה'])) throw new Error('עדיין אין לך גישה לאפליקציה. פנה לנטו ליאור לב.');
  return { token: newToken_({ t: 'c', code: cust['קוד'] }), name: cust['שם'] };
}

function adminLogin_(p) {
  if (String(p.pass || '').trim() !== String(setting_('סיסמת מנהל'))) throw new Error('סיסמה שגויה');
  return { token: newToken_({ t: 'a' }) };
}

function catalog_(s) {
  const c = customers_().find(x => x['קוד'] === s.code);
  const asg = String(c['מגוון'] || '').trim();
  const products = rows_(SH.products).filter(r => truthy_(r['פעיל']) && (!asg || listOf_(r['מגוונים']).indexOf(asg) >= 0)).map(r => ({
    id: pid_(r), item: String(r['מק"ט'] || ''), name: r['שם מוצר'], cat: r['קטגוריה'],
    sub: r['תת-קטגוריה'] || 'כללי', carton: Number(r["יח' בקרטון"]) || null, barcode: String(r['ברקוד'] || ''),
    img: imgUrl_(r['מזהה תמונה']), tag: r['תגית'] || ''
  }));
  const promos = rows_(SH.promos).filter(r => truthy_(r['פעיל'])).map(r => ({ title: r['כותרת'], text: r['טקסט'], id: String(r['מק"ט'] || '') }));
  const mine = ordersWithLines_().filter(o => o.cust === s.code).slice(0, 5);
  return {
    customer: { code: c['קוד'], name: c['שם'], addr: c['כתובת'] || '', days: c['ימי אספקה'] || '', cutoff: c['שעת סגירה'] || setting_('שעת סגירה') || '' },
    products, promos, orders: mine,
    owner: { wa: normPhone_(setting_('וואטסאפ')), waShown: setting_('וואטסאפ'), email: setting_('מייל הזמנות') }
  };
}

function submit_(s, p) {
  const lines = (p.lines || []).filter(l => Number(l[1]) > 0);
  if (!lines.length) throw new Error('ההזמנה ריקה');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const c = customers_().find(x => x['קוד'] === s.code);
    const prod = {}; rows_(SH.products).forEach(r => prod[pid_(r)] = r);
    const no = Number(setting_('מספר הזמנה הבא')) || 1001;
    setSetting_('מספר הזמנה הבא', no + 1);
    const now = new Date();
    let voiceUrl = '';
    if (p.voice && p.voice.b64) {
      const blob = Utilities.newBlob(Utilities.base64Decode(p.voice.b64), p.voice.mime || 'audio/mpeg', 'הזמנה-' + no + '-' + (p.voice.name || 'הקלטה'));
      const f = subFolder_(appFolder_(), 'הקלטות').createFile(blob);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      voiceUrl = f.getUrl();
    }
    let cartons = 0, units = 0;
    const lr = lines.map(([id, q, u]) => {
      const r = prod[id] || {};
      q = Number(q);
      if (u === 'u') units += q; else cartons += q;
      return [no, String(r['מק"ט'] || ''), r['שם מוצר'] || id, q, u === 'u' ? 'יחידה' : 'קרטון', r['קטגוריה'] || '', String(r['ברקוד'] || ''), Number(r["יח' בקרטון"]) || ''];
    });
    ss.getSheetByName(SH.lines).getRange(ss.getSheetByName(SH.lines).getLastRow() + 1, 1, lr.length, lr[0].length).setValues(lr);
    ss.getSheetByName(SH.orders).appendRow([no, now, c['קוד'], c['שם'], 'חדשה', lines.length, cartons, units, p.note || '', voiceUrl, '']);
    SpreadsheetApp.flush();
    lock.releaseLock();

    const order = ordersWithLines_().find(o => o.no === no);
    const pdf = makePdf_(order);
    setOrderField_(no, 'PDF', pdf.url);
    // מייל לבעל העסק
    const owner = setting_('מייל הזמנות');
    const subj = 'הזמנה חדשה ' + no + ' – ' + c['שם'];
    if (owner) MailApp.sendEmail({ to: owner, subject: subj, htmlBody: orderEmailHtml_(order, true), attachments: [pdf.blob], name: 'נטו הזמנות' });
    // אישור ללקוח
    if (c['מייל']) MailApp.sendEmail({ to: c['מייל'], subject: 'ההזמנה שלך התקבלה – ' + no, htmlBody: orderEmailHtml_(order, false), attachments: [pdf.blob], name: setting_('שם העסק') || 'נטו ליאור לב' });
    return { no, at: now.getTime(), pdfUrl: pdf.url, cartons, units, count: lines.length };
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

function adminData_() {
  const ss = SpreadsheetApp.getActive();
  const settings = {};
  ss.getSheetByName(SH.settings).getDataRange().getValues().slice(1).forEach(r => { if (String(r[0]).trim()) settings[String(r[0]).trim()] = clean_(r[1]); });
  const pub = settings['כתובת האפליקציה'] || ScriptApp.getService().getUrl();
  const custs = sheetRows_(SH.customers).map(x => { const c = x.data; c['קוד'] = String(c['קוד']).padStart(4, '0'); return Object.assign({ _row: x.row }, c); });
  const orders = ordersWithLines_().slice(0, 300);
  return {
    orders,
    customers: custs.map(c => Object.assign(c, { _wa: normPhone_(c['וואטסאפ'] || c['טלפון']), _access: truthy_(c['גישה']) })),
    products: sheetRows_(SH.products).map(x => Object.assign({ _row: x.row, _img: imgUrl_(x.data['מזהה תמונה']), _active: truthy_(x.data['פעיל']) }, x.data)),
    promos: sheetRows_(SH.promos).map(x => Object.assign({ _row: x.row, _active: truthy_(x.data['פעיל']) }, x.data)),
    due: dueToday_(custs, orders),
    assorts: listOf_(settings['מגוונים']),
    settings, url: pub, appUrl: ScriptApp.getService().getUrl(),
    business: settings['שם העסק'] || 'נטו ליאור לב',
    sheetUrl: ss.getUrl(), folderUrl: appFolder_().getUrl()
  };
}

/* ================= עריכה מתוך האפליקציה ================= */

const EDITABLE = { products: SH.products, customers: SH.customers, promos: SH.promos };
const TEXT_COLS = ['מק"ט', 'ברקוד', 'מזהה פנימי', 'קוד', 'טלפון', 'וואטסאפ', 'סיסמה', 'מפתח אישי', 'ימי אספקה', 'ימי תזכורת', 'שעת סגירה'];

function sheetRows_(name) {
  const v = SpreadsheetApp.getActive().getSheetByName(name).getDataRange().getValues();
  const h = v[0].map(x => String(x).trim());
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i].some(x => x !== '')) continue;
    const d = {}; h.forEach((k, j) => { if (k) d[k] = clean_(v[i][j]); });
    out.push({ row: i + 1, data: d });
  }
  return out;
}
function clean_(x) {
  if (x instanceof Date) return Utilities.formatDate(x, 'Asia/Jerusalem', x.getFullYear() < 1901 ? 'HH:mm' : 'dd/MM/yyyy HH:mm');
  return x;
}
function ensureCols_(sh, keys) {
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  keys.filter(k => k && k[0] !== '_' && h.indexOf(k) < 0).forEach(k => { sh.getRange(1, h.length + 1).setValue(k).setFontWeight('bold'); h.push(k); });
  return h;
}
function saveRow_(kind, row, data) {
  const name = EDITABLE[kind]; if (!name) throw new Error('לא ניתן לערוך');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(name);
    const h = ensureCols_(sh, Object.keys(data));
    if (kind === 'customers') {
      data['קוד'] = String(data['קוד'] || '').trim().padStart(4, '0');
      if (!/^\d{4,}$/.test(data['קוד'])) throw new Error('קוד לקוח צריך להיות מספר (לפחות 4 ספרות)');
      const dup = sheetRows_(name).find(x => String(x.data['קוד']).padStart(4, '0') === data['קוד'] && x.row !== Number(row));
      if (dup) throw new Error('כבר יש לקוח עם הקוד ' + data['קוד']);
      if (!data['סיסמה']) data['סיסמה'] = String(1000 + Math.floor(Math.random() * 9000));
      if (!data['מפתח אישי'] && !row) data['מפתח אישי'] = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    }
    if (kind === 'products' && data['מק"ט']) {
      const dup = sheetRows_(name).find(x => String(x.data['מק"ט']) === String(data['מק"ט']) && x.row !== Number(row));
      if (dup) throw new Error('כבר יש מוצר עם המק"ט ' + data['מק"ט']);
    }
    if (kind === 'products' && !row && !data['מק"ט'] && !data['ברקוד'] && !data['מזהה פנימי']) data['מזהה פנימי'] = nextInternalId_();
    const r = row ? Number(row) : sh.getLastRow() + 1;
    const cur = row ? sh.getRange(r, 1, 1, h.length).getValues()[0] : h.map(() => '');
    h.forEach((k, i) => { if (Object.prototype.hasOwnProperty.call(data, k)) cur[i] = data[k]; });
    h.forEach((k, i) => { if (TEXT_COLS.indexOf(k) >= 0) sh.getRange(r, i + 1).setNumberFormat('@'); });
    sh.getRange(r, 1, 1, h.length).setValues([cur]);
  } finally { lock.releaseLock(); }
}
function deleteRow_(kind, row, check) {
  const name = EDITABLE[kind]; if (!name) throw new Error('לא ניתן למחוק');
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const first = String(sh.getRange(Number(row), 1).getValue());
  if (String(check) !== first) throw new Error('הנתונים השתנו בינתיים. רענן ונסה שוב.');
  sh.deleteRow(Number(row));
}
/* ---------- מגוונים ופעולות על הרבה מוצרים ---------- */
function bulkProducts_(p) {
  const rows = (p.rows || []).map(Number).filter(Boolean);
  if (!rows.length) throw new Error('לא נבחרו מוצרים');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SH.products);
    const h = ensureCols_(sh, ['מגוונים']);
    const v = sh.getRange(1, 1, sh.getLastRow(), h.length).getValues();
    const cAct = h.indexOf('פעיל'), cAs = h.indexOf('מגוונים');
    const actCol = v.map(r => [r[cAct]]), asCol = v.map(r => [r[cAs]]);
    rows.forEach(r => {
      const i = r - 1; if (i < 1 || i >= v.length) return;
      if (typeof p.active === 'boolean') actCol[i][0] = p.active;
      let l = listOf_(asCol[i][0]);
      if (p.addTo && l.indexOf(p.addTo) < 0) l.push(p.addTo);
      if (p.removeFrom) l = l.filter(x => x !== p.removeFrom);
      asCol[i][0] = l.join(', ');
    });
    sh.getRange(1, cAct + 1, v.length, 1).setValues(actCol);
    sh.getRange(1, cAs + 1, v.length, 1).setValues(asCol);
  } finally { lock.releaseLock(); }
}
function saveAssorts_(names, renames) {
  names = names.map(x => String(x).replace(/,/g, ' ').trim()).filter(Boolean);
  const keep = {}; names.forEach(n => keep[n] = true);
  const fix = l => listOf_(l).map(x => renames[x] || x).filter(x => keep[x]);
  const ss = SpreadsheetApp.getActive();
  [[SH.products, 'מגוונים', true], [SH.customers, 'מגוון', false]].forEach(([name, col, multi]) => {
    const sh = ss.getSheetByName(name);
    const h = ensureCols_(sh, [col]);
    const n = sh.getLastRow(); if (n < 2) return;
    const rng = sh.getRange(2, h.indexOf(col) + 1, n - 1, 1);
    rng.setValues(rng.getValues().map(r => [multi ? fix(r[0]).join(', ') : (fix(r[0])[0] || '')]));
  });
  setSetting_('מגוונים', names.join(', '));
}

/* ---------- מזהה פנימי למוצר בלי מק"ט וברקוד ---------- */
function nextInternalId_() {
  let max = 0;
  rows_(SH.products).forEach(r => { const m = String(r['מזהה פנימי'] || '').match(/(\d+)/); if (m) max = Math.max(max, Number(m[1])); });
  return 'P' + String(max + 1).padStart(4, '0');
}
function fillInternalIds_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.products);
  const h = ensureCols_(sh, ['מזהה פנימי']);
  const v = sh.getDataRange().getValues();
  const cS = h.indexOf('מק"ט'), cB = h.indexOf('ברקוד'), cI = h.indexOf('מזהה פנימי');
  let max = 0;
  v.forEach((r, i) => { if (i) { const m = String(r[cI] || '').match(/(\d+)/); if (m) max = Math.max(max, Number(m[1])); } });
  const col = v.map((r, i) => {
    if (!i) return [r[cI]];
    if (!r.some(x => x !== '')) return [r[cI]];
    if (!String(r[cS] || '').trim() && !String(r[cB] || '').trim() && !String(r[cI] || '').trim()) return ['P' + String(++max).padStart(4, '0')];
    return [r[cI]];
  });
  sh.getRange(1, cI + 1, col.length, 1).setNumberFormat('@').setValues(col);
}

/* ---------- ייבוא מוצרים מקובץ ייבוא-מוצרים.json (פעם אחת) ---------- */
function importCatalog_() {
  const folder = appFolder_();
  const it = folder.getFilesByName('ייבוא-מוצרים.json');
  if (!it.hasNext()) return '';
  const file = it.next();
  const job = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.products);
  const h = ensureCols_(sh, ['מגוונים', 'מזהה פנימי', 'עמוד בקטלוג', 'מה לבדוק']);
  const v = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), h.length).getValues();
  const c = k => h.indexOf(k);
  const bySku = {}; v.forEach((r, i) => { if (i && String(r[c('מק"ט')]).trim()) bySku[String(r[c('מק"ט')]).trim()] = i; });
  const clearName = s => String(s || '').split(',').map(x => x.trim()).filter(x => x && x !== 'שם לבדיקה').join(', ');
  let fixed = 0, skus = 0;
  (job.setSku || []).forEach(f => {
    const i = v.findIndex((r, j) => j && !String(r[c('מק"ט')]).trim() && String(r[c('שם מוצר')]).trim() === f.find);
    if (i < 1) return;
    if (f.sku && bySku[f.sku] === undefined) { v[i][c('מק"ט')] = f.sku; bySku[f.sku] = i; skus++; }
    if (f.name) v[i][c('שם מוצר')] = f.name;
    if (f.barcode && !String(v[i][c('ברקוד')]).trim()) v[i][c('ברקוד')] = f.barcode;
    if (f.carton && !v[i][c("יח' בקרטון")]) v[i][c("יח' בקרטון")] = f.carton;
    v[i][c('מה לבדוק')] = clearName(String(v[i][c('מה לבדוק')]).replace('חסר מק"ט', ''));
  });
  (job.fixName || []).forEach(f => {
    const i = bySku[String(f.sku)]; if (i === undefined) return;
    if (f.name) v[i][c('שם מוצר')] = f.name;
    if (f.newSku && bySku[f.newSku] === undefined) { v[i][c('מק"ט')] = f.newSku; bySku[f.newSku] = i; }
    v[i][c('מה לבדוק')] = clearName(v[i][c('מה לבדוק')]);
    fixed++;
  });
  const add = (job.add || []).filter(a => !a['מק"ט'] || bySku[a['מק"ט']] === undefined);
  const newRows = add.map(a => h.map(k => k === 'פעיל' ? false : (a[k] === undefined ? '' : a[k])));
  sh.getRange(1, 1, v.length, h.length).setValues(v);
  if (newRows.length) {
    const start = sh.getLastRow() + 1;
    h.forEach((k, i) => { if (TEXT_COLS.indexOf(k) >= 0) sh.getRange(start, i + 1, newRows.length, 1).setNumberFormat('@'); });
    sh.getRange(start, 1, newRows.length, h.length).setValues(newRows);
  }
  file.setName('ייבוא-מוצרים-בוצע-' + Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd') + '.json');
  return 'נוספו ' + newRows.length + ' מוצרים, תוקנו ' + fixed + ' שמות והושלמו ' + skus + ' מק"טים';
}

function uploadImage_(b64, name) {
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', (name || ('img-' + Date.now())) + '.jpg');
  const f = subFolder_(appFolder_(), 'תמונות').createFile(blob);
  return { id: f.getId(), name: f.getName(), url: imgUrl_(f.getId()) };
}

/* ================= מי צריך להזמין היום ================= */

const DAYS = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ש'];
function todayIdx_(d) { return Number(Utilities.formatDate(d || new Date(), 'Asia/Jerusalem', 'u')) % 7; }
function parseDays_(s) { return String(s || '').replace(/['׳"]/g, '').split(/[,\s]+/).map(x => DAYS.indexOf(x.trim().charAt(0))).filter(i => i >= 0); }
function remindDays_(c) {
  const rd = parseDays_(c['ימי תזכורת']);
  if (rd.length) return rd;
  return parseDays_(c['ימי אספקה']).map(i => (i + 6) % 7);
}
function remindHour_(c) { return Number(c['שעת תזכורת']) || Number(setting_('שעת תזכורת')) || 10; }
function dueToday_(custs, orders) {
  const now = new Date(), t = todayIdx_(now);
  return custs.filter(c => truthy_(c['גישה']) && remindDays_(c).indexOf(t) >= 0).map(c => {
    const last = orders.find(o => o.cust === c['קוד']);
    return { code: c['קוד'], name: c['שם'], hour: remindHour_(c), ordered: !!(last && now - last.at < 24 * 3600 * 1000), wa: normPhone_(c['וואטסאפ'] || c['טלפון']), key: c['מפתח אישי'] || '', email: c['מייל'] || '' };
  });
}

function setStatus_(no, status) { setOrderField_(Number(no), 'סטטוס', status); return { ok: true }; }

function setAccess_(code, access) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.customers);
  const v = sh.getDataRange().getValues(), h = v[0];
  const r = v.findIndex((row, i) => i > 0 && String(row[h.indexOf('קוד')]) === String(code));
  if (r < 1) throw new Error('לקוח לא נמצא');
  sh.getRange(r + 1, h.indexOf('גישה') + 1).setValue(!!access);
  return { ok: true };
}

/* ================= PDF ומייל ================= */

function orderPdf_(no, s) {
  const o = ordersWithLines_().find(x => x.no === Number(no));
  if (!o) throw new Error('ההזמנה לא נמצאה');
  if (s.t === 'c' && o.cust !== s.code) throw new Error('אין הרשאה');
  if (o.pdf) return o.pdf;
  const pdf = makePdf_(o); setOrderField_(o.no, 'PDF', pdf.url); return pdf.url;
}

function makePdf_(o) {
  const html = orderFormHtml_(o);
  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName('הזמנה-' + o.no + '.pdf');
  const f = subFolder_(appFolder_(), 'הזמנות PDF').createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { blob, url: f.getUrl() };
}

function logoDataUri_() {
  const id = PropertiesService.getScriptProperties().getProperty('LOGO_ID');
  if (!id) return '';
  try { const b = DriveApp.getFileById(id).getBlob(); return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); } catch (e) { return ''; }
}

function orderFormHtml_(o) {
  const c = customers_().find(x => x['קוד'] === o.cust) || {};
  const biz = setting_('שם העסק') || 'נטו ליאור לב';
  let i = 0, totU = 0;
  const body = ZONES.map(z => {
    const ls = o.lines.filter(l => l.cat === z);
    if (!ls.length) return '';
    return '<tr><td colspan="7" style="background:#f3f3f3;font-weight:bold">' + z + '</td></tr>' + ls.map(l => {
      i++; const u = l.unit === 'יחידה' ? l.q : (l.carton ? l.q * l.carton : '');
      if (u) totU += u;
      return '<tr><td>' + i + '</td><td>' + esc_(l.item) + '</td><td>' + esc_(l.barcode) + '</td><td>' + esc_(l.name) + '</td><td style="text-align:center;font-weight:bold">' + (l.unit === 'קרטון' ? l.q : '') + '</td><td style="text-align:center;font-weight:bold">' + u + '</td><td></td></tr>';
    }).join('');
  }).join('') + o.lines.filter(l => ZONES.indexOf(l.cat) < 0).map(l => { i++; return '<tr><td>' + i + '</td><td>' + esc_(l.item) + '</td><td>' + esc_(l.barcode) + '</td><td>' + esc_(l.name) + '</td><td>' + (l.unit === 'קרטון' ? l.q : '') + '</td><td>' + (l.unit === 'יחידה' ? l.q : '') + '</td><td></td></tr>'; }).join('');
  const logo = logoDataUri_();
  const d = new Date(o.at);
  return '<html dir="rtl"><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;font-size:12px;direction:rtl}table{width:100%;border-collapse:collapse}th,td{border:1px solid #999;padding:4px;text-align:right}th{background:#e8e8e8}.sig td{border:0;border-top:1px solid #000;padding-top:4px;color:#555}</style></head><body>' +
    '<table style="border:0"><tr><td style="border:0;width:90px">' + (logo ? '<img src="' + logo + '" style="height:56px">' : '') + '</td><td style="border:0"><div style="font-size:20px;font-weight:bold">טופס הזמנה ' + o.no + '</div><div>' + esc_(biz) + '</div></td>' +
    '<td style="border:0;text-align:left">תאריך: ' + Utilities.formatDate(d, 'Asia/Jerusalem', 'dd/MM/yyyy') + '<br>שעה: ' + Utilities.formatDate(d, 'Asia/Jerusalem', 'HH:mm') + '<br>סטטוס: ' + esc_(o.status) + '</td></tr></table><hr>' +
    '<p><b>לקוח:</b> ' + esc_(c['שם'] || o.custName) + ' (' + esc_(o.cust) + ') &nbsp; <b>טלפון:</b> ' + esc_(c['טלפון'] || '—') + ' &nbsp; <b>כתובת:</b> ' + esc_(c['כתובת'] || '—') + ' &nbsp; <b>ימי אספקה:</b> ' + esc_(c['ימי אספקה'] || '—') + '</p>' +
    '<table><tr><th>#</th><th>מק"ט</th><th>ברקוד</th><th>תאור הסחורה</th><th>קרטונים</th><th>סה"כ יח\'</th><th>סופק ✓</th></tr>' + body +
    '<tr><td colspan="4" style="font-weight:bold;background:#e8e8e8">סה"כ ' + o.lines.length + ' מוצרים</td><td style="text-align:center;font-weight:bold;background:#e8e8e8">' + o.cartons + '</td><td style="text-align:center;font-weight:bold;background:#e8e8e8">' + (totU || '') + '</td><td style="background:#e8e8e8"></td></tr></table>' +
    '<p style="border:1px dashed #999;padding:6px"><b>הערות הלקוח:</b> ' + esc_(o.note || '') + (o.voice ? ' · צורפה הערה קולית' : '') + '</p>' +
    '<br><br><table class="sig"><tr><td>ליקט</td><td style="width:20px;border:0"></td><td>נהג</td><td style="width:20px;border:0"></td><td>חתימת הלקוח</td></tr></table></body></html>';
}

function orderEmailHtml_(o, forOwner) {
  const rows = o.lines.map(l => '<tr><td style="padding:3px 6px;border-bottom:1px solid #eee">' + esc_(l.item) + '</td><td style="padding:3px 6px;border-bottom:1px solid #eee">' + esc_(l.name) + '</td><td style="padding:3px 6px;border-bottom:1px solid #eee;font-weight:bold">' + l.q + ' ' + (l.unit === 'קרטון' ? 'קרט׳' : 'יח׳') + '</td></tr>').join('');
  const head = forOwner ? '<h2 style="margin:0">הזמנה חדשה ' + o.no + '</h2><p>מאת <b>' + esc_(o.custName) + '</b> (' + esc_(o.cust) + ')</p>'
    : '<h2 style="margin:0">ההזמנה שלך התקבלה ✓</h2><p>מספר הזמנה ' + o.no + '. נעדכן כשהיא תאושר.</p>';
  return '<div dir="rtl" style="font-family:Arial,sans-serif">' + head +
    '<p>' + o.lines.length + ' מוצרים · ' + o.cartons + ' קרטונים' + (o.units ? ' · ' + o.units + ' יחידות' : '') + '</p>' +
    '<table style="border-collapse:collapse;font-size:14px">' + rows + '</table>' +
    (o.note ? '<p><b>הערות:</b> ' + esc_(o.note) + '</p>' : '') + (o.voice ? '<p><a href="' + o.voice + '">🎙 הערה קולית</a></p>' : '') +
    '<p style="color:#777">טופס ההזמנה המלא מצורף כ-PDF.</p></div>';
}

/* ================= תזכורות ================= */

// רץ כל שעה: תזכורת במייל ללקוחות בשעה וביום שלהם, וסיכום יומי לבעל העסק.
function sendReminders() {
  const now = new Date();
  const H = Number(Utilities.formatDate(now, 'Asia/Jerusalem', 'H'));
  const biz = setting_('שם העסק') || 'נטו ליאור לב';
  const pub = setting_('כתובת האפליקציה') || ScriptApp.getService().getUrl();
  const custs = customers_();
  const orders = ordersWithLines_();
  const due = dueToday_(custs, orders);
  const text = setting_('נוסח תזכורת') || 'שלום {שם}, תזכורת להזמנה מ' + biz + '.';
  due.filter(d => d.hour === H && !d.ordered && d.email).forEach(d => {
    const c = custs.find(x => x['קוד'] === d.code);
    const msg = text.replace('{שם}', d.name).replace('{שעת סגירה}', c['שעת סגירה'] || setting_('שעת סגירה') || '');
    const link = pub + '?u=' + d.code + '&k=' + d.key;
    MailApp.sendEmail({ to: d.email, subject: 'תזכורת להזמנה – ' + biz, htmlBody: '<div dir="rtl" style="font-family:Arial">' + esc_(msg) + '<p><a href="' + link + '" style="background:#e63a2f;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none">להזמנה</a></p></div>', name: biz });
  });
  // סיכום יומי
  const sumHour = Number(setting_('שעת סיכום')) || 13;
  if (H !== sumHour || String(setting_('סיכום יומי')).trim() === 'כבוי' || !due.length) return;
  const owner = setting_('מייל הזמנות'); if (!owner) return;
  const missing = due.filter(d => !d.ordered), done = due.filter(d => d.ordered);
  const rows = missing.map(d => {
    const link = pub + '?u=' + d.code + '&k=' + d.key;
    const wmsg = 'שלום ' + d.name + ', תזכורת מ' + biz + ': עוד לא קיבלנו ממך הזמנה היום 🙂 להזמנה: ' + link;
    return '<tr><td style="padding:8px;border-bottom:1px solid #eee"><b>' + esc_(d.name) + '</b> (' + d.code + ')</td><td style="padding:8px;border-bottom:1px solid #eee">' +
      (d.wa ? '<a href="https://wa.me/' + d.wa + '?text=' + encodeURIComponent(wmsg) + '" style="background:#1fa855;color:#fff;padding:6px 12px;border-radius:8px;text-decoration:none">שלח תזכורת בוואטסאפ</a>' : 'אין טלפון') + '</td></tr>';
  }).join('');
  const html = '<div dir="rtl" style="font-family:Arial;font-size:15px">' +
    (missing.length ? '<h2 style="margin:0 0 8px">' + missing.length + ' לקוחות עוד לא הזמינו היום</h2><table style="border-collapse:collapse">' + rows + '</table>'
      : '<h2>כל הלקוחות הזמינו היום ✓</h2>') +
    (done.length ? '<p style="color:#1f8a4c">הזמינו: ' + done.map(d => esc_(d.name)).join(', ') + '</p>' : '') +
    '<p style="color:#777">הסיכום נשלח כל יום בשעה ' + sumHour + ':00. אפשר לשנות או לכבות במסך ההגדרות באפליקציה.</p></div>';
  MailApp.sendEmail({ to: owner, subject: missing.length ? 'נטו הזמנות: ' + missing.length + ' לקוחות עוד לא הזמינו' : 'נטו הזמנות: כולם הזמינו היום ✓', htmlBody: html, name: 'נטו הזמנות' });
}

/* ================= עזר לגיליון ================= */

function rows_(name) {
  const v = SpreadsheetApp.getActive().getSheetByName(name).getDataRange().getValues();
  const h = v.shift().map(x => String(x).trim());
  return v.filter(r => r.some(x => x !== '')).map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
}
function customers_() { return rows_(SH.customers).map(c => { c['קוד'] = String(c['קוד']).padStart(4, '0'); return c; }); }

function ordersWithLines_() {
  const lines = {};
  rows_(SH.lines).forEach(l => {
    const no = Number(l['מספר הזמנה']);
    (lines[no] = lines[no] || []).push({ item: String(l['מק"ט'] || ''), name: l['שם מוצר'], q: Number(l['כמות']), unit: l['יחידה'], cat: l['קטגוריה'], barcode: String(l['ברקוד'] || ''), carton: Number(l["יח' בקרטון"]) || null });
  });
  return rows_(SH.orders).map(o => ({
    no: Number(o['מספר']), at: new Date(o['תאריך']).getTime(), cust: String(o['קוד לקוח']).padStart(4, '0'), custName: o['שם לקוח'],
    status: o['סטטוס'], cartons: Number(o['קרטונים']) || 0, units: Number(o['יחידות']) || 0, note: o['הערה'] || '', voice: o['הקלטה'] || '', pdf: o['PDF'] || '',
    lines: lines[Number(o['מספר'])] || []
  })).sort((a, b) => b.no - a.no);
}

function setOrderField_(no, field, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.orders);
  const v = sh.getDataRange().getValues(), h = v[0];
  const r = v.findIndex((row, i) => i > 0 && Number(row[0]) === Number(no));
  if (r > 0) sh.getRange(r + 1, h.indexOf(field) + 1).setValue(value);
}

function setting_(key) {
  const v = SpreadsheetApp.getActive().getSheetByName(SH.settings).getDataRange().getValues();
  const r = v.find(x => String(x[0]).trim() === key);
  return r ? r[1] : '';
}
function setSetting_(key, val) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.settings);
  const v = sh.getDataRange().getValues();
  const i = v.findIndex(x => String(x[0]).trim() === key);
  if (i >= 0) sh.getRange(i + 1, 2).setNumberFormat('@').setValue(String(val)); else { sh.appendRow([key, '']); sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(String(val)); }
}

function truthy_(x) { return x === true || /^(true|כן|1|v|✓)$/i.test(String(x).trim()); }
function normPhone_(p) { const d = String(p || '').replace(/\D/g, ''); return d.startsWith('0') ? '972' + d.slice(1) : d; }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
