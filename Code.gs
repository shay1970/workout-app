/************************************************************
 *  תוכנית אימונים - שי  |  Cloud Sync backend (Apps Script)
 *  ---------------------------------------------------------
 *  קובץ זה הוא ה-"שרת" של הסנכרון לענן. הוא תואם בדיוק
 *  לפרוטוקול ש-index.html משתמש בו:
 *
 *    העלאה (שמירה לענן):  GET  ?action=chunk&i=<idx>&n=<total>&d=<chunk>&callback=<cb>
 *                         כל בקשה מחזירה JSONP: cb({"status":"chunk_ok"})
 *                         בצ'אנק האחרון הנתונים המורכבים נשמרים בגיליון.
 *
 *    הורדה (טעינה מהענן): GET  ?callback=<cb>   (בלי action)
 *                         מחזיר JSONP: cb(<אובייקט המצב שנשמר>)
 *
 *  ---------- הוראות פרסום (חד-פעמי) ----------
 *  1. פתח גיליון Google Sheets חדש (הוא ישמש כמאגר).
 *  2. תפריט: Extensions → Apps Script.
 *  3. מחק את קוד ברירת המחדל והדבק את כל הקובץ הזה.
 *  4. שמור (💾).
 *  5. Deploy → New deployment → סוג: Web app.
 *       - Execute as:      Me
 *       - Who has access:  Anyone
 *     Deploy → אשר הרשאות → העתק את כתובת ה-/exec.
 *  6. הדבק את הכתובת באפליקציה: כפתור "☁️ סנכרן" → שדה הכתובת.
 *
 *  שים לב: בכל שינוי בקוד צריך Deploy → Manage deployments →
 *  עריכה → Version: New version, אחרת האפליקציה תראה קוד ישן.
 ************************************************************/

var SHEET_NAME = 'WorkoutData';   // שם גיליון המאגר (נוצר אוטומטית)
var DATA_CELL  = 'A1';            // התא שבו נשמר ה-JSON המלא
var STAGE_KEY  = 'wa_stage_buf';  // מפתח לאחסון זמני של הצ'אנקים

/** מחזיר (ויוצר אם צריך) את גיליון המאגר */
function _sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(DATA_CELL).setValue('');
  }
  return sh;
}

/** עוטף אובייקט כתשובת JSONP (JavaScript) שהדפדפן יריץ */
function _jsonp(callback, obj) {
  var cb = callback || 'callback';
  var body = cb + '(' + JSON.stringify(obj) + ');';
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/** נקודת הכניסה היחידה — האפליקציה תמיד קוראת ב-GET (JSONP) */
function doGet(e) {
  var p  = (e && e.parameter) || {};
  var cb = p.callback || 'callback';
  try {
    if (p.action === 'chunk') {
      return _handleChunk(p, cb);
    }
    // אין action → בקשת טעינה: החזר את המצב השמור כאובייקט
    var raw = _sheet().getRange(DATA_CELL).getValue();
    var obj = {};
    if (raw) {
      try { obj = JSON.parse(raw); } catch (err) { obj = {}; }
    }
    return _jsonp(cb, obj);
  } catch (err) {
    return _jsonp(cb, { status: 'error', msg: String(err) });
  }
}

/** תמיכה גם ב-POST (לא בשימוש כרגע, ליתר בטחון) */
function doPost(e) { return doGet(e); }

/** צובר צ'אנק אחד; בצ'אנק האחרון מאמת ושומר את ה-JSON המלא */
function _handleChunk(p, cb) {
  var i = parseInt(p.i, 10);
  var n = parseInt(p.n, 10);
  var d = p.d || '';
  if (isNaN(i) || isNaN(n) || n < 1) {
    return _jsonp(cb, { status: 'error', msg: 'bad chunk params' });
  }

  var cache = CacheService.getScriptCache();
  var lock  = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // בצ'אנק הראשון מתחילים חוצץ נקי; אחרת ממשיכים מהקיים
    var buf = (i === 0) ? '' : (cache.get(STAGE_KEY) || '');
    buf += d;

    if (i < n - 1) {
      cache.put(STAGE_KEY, buf, 1800); // אחסון זמני ל-30 דקות
      return _jsonp(cb, { status: 'chunk_ok' });
    }

    // הצ'אנק האחרון — אמת JSON ושמור לצמיתות
    var parsed;
    try {
      parsed = JSON.parse(buf);
    } catch (err) {
      cache.remove(STAGE_KEY);
      return _jsonp(cb, { status: 'error', msg: 'invalid JSON assembled (retry sync)' });
    }
    _sheet().getRange(DATA_CELL).setValue(JSON.stringify(parsed));
    cache.remove(STAGE_KEY);
    return _jsonp(cb, { status: 'chunk_ok' });
  } finally {
    lock.releaseLock();
  }
}
