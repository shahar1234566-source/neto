/**
 * נטו הזמנות – טוען (Google Apps Script)
 * זה הקוד היחיד שמודבק ב-Apps Script, והוא לא צריך להשתנות יותר.
 * כל הלוגיקה נמצאת בקובץ server.js בתיקיית Drive "נטו הזמנות", ונטענת ממנו בכל פעם.
 * עדכונים עתידיים = החלפת server.js ב-Drive בלבד (בלי הדבקה, בלי setup, בלי פריסה).
 */
var LOADER_FOLDER = 'נטו הזמנות';
var LOADER_FILE = 'server.js';
var __srv = null;

function __server() {
  if (__srv) return __srv;
  var folders = DriveApp.getFoldersByName(LOADER_FOLDER);
  if (!folders.hasNext()) throw new Error('לא נמצאה התיקייה "' + LOADER_FOLDER + '" ב-Drive');
  var files = folders.next().getFilesByName(LOADER_FILE);
  if (!files.hasNext()) throw new Error('חסר הקובץ ' + LOADER_FILE + ' בתיקייה "' + LOADER_FOLDER + '"');
  var src = files.next().getBlob().getDataAsString('UTF-8');
  __srv = (0, eval)('(function(){\n' + src + '\n})()');
  return __srv;
}

function doGet(e) {
  try { return __server().doGet(e); }
  catch (err) { return HtmlService.createHtmlOutput('<div dir="rtl" style="font-family:Arial;padding:20px">תקלה בטעינת האפליקציה: ' + String(err && err.message || err) + '</div>'); }
}
function api(action, p) {
  try { return __server().api(action, p); }
  catch (err) { return { error: String(err && err.message || err) }; }
}
function setup() { return __server().setup(); }
function sendReminders() { return __server().sendReminders(); }

// לא נקרא אף פעם. נמצא כאן רק כדי שגוגל יבקש את כל ההרשאות (קריאה וכתיבה) שהקוד ב-Drive צריך.
function __permissions_() {
  var ss = SpreadsheetApp.getActive(); ss.getSheets()[0].getRange(1, 1).setValue(ss.getSheets()[0].getRange(1, 1).getValue()); ss.toast('');
  var f = DriveApp.getRootFolder(); f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var nf = f.createFolder('x'); var file = nf.createFile(Utilities.newBlob('x')); file.setName('y'); file.setTrashed(true);
  DriveApp.createFile(Utilities.newBlob('x')); Utilities.unzip(Utilities.newBlob('x'));
  MailApp.sendEmail({ to: 'x', subject: 'x', htmlBody: 'x' }); MailApp.getRemainingDailyQuota();
  ScriptApp.newTrigger('sendReminders').timeBased().everyHours(1).create(); ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  HtmlService.createHtmlOutput('').getBlob().getAs('application/pdf');
  CacheService.getScriptCache(); PropertiesService.getScriptProperties().setProperty('x', 'x'); LockService.getScriptLock(); Utilities.getUuid();
  UrlFetchApp.fetch('https://www.google.com');
}
