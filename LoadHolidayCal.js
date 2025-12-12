// ==========================
//  טעינת חגים לקאש לשנה
//  Cache key: "holidays_<year>"
// ==========================
function preloadHolidayCache(yearOpt) {
  var year = yearOpt || new Date().getFullYear();
  var cacheKey = "holidays_" + year;

  var url =
    "https://www.hebcal.com/hebcal" +
    "?v=1" +
    "&cfg=json" +
    "&maj=on" +        // כמו בקוד המקורי שלך – חגים גדולים
    "&year=" + year +
    "&c=on" +
    "&geo=IL";

  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());

  var holidays = (data.items || [])
    .filter(function (item) {
      return (
        ["holiday", "festival"].includes(item.category) ||
        item.title.indexOf("ערב ") === 0 ||
        item.title.indexOf("Erev ") === 0
      );
    })
    .map(function (item) {
      // לחתוך ל־yyyy-MM-dd (אם יש T00:00:00 וכו')
      var d = (item.date || "").substring(0, 10);
      return {
        date: d,
        title: item.title
      };
    });

  Logger.log("==== רשימת חגים לשנה " + year + " (מה-API) ====");
  holidays.forEach(function (h) {
    Logger.log("API חג: " + h.date + " → " + h.title);
  });

  CacheService.getScriptCache().put(cacheKey, JSON.stringify(holidays), 21600); // 6 שעות
  Logger.log("✅ נשמרו " + holidays.length + " חגים בקאש לשנה " + year + " (key=" + cacheKey + ")");
}

// ==========================
//  המרת ערך לתאריך אמיתי
//  (גיבוי אם זה לא Date אלא מספר/טקסט)
// ==========================
function normalizeDateCell(raw) {
  if (!raw) return null;

  if (raw instanceof Date) {
    if (!isNaN(raw.getTime())) return raw;
    return null;
  }

  if (typeof raw === "number") {
    // serial של Sheets/Excel
    var millis = (raw - 25569) * 24 * 60 * 60 * 1000;
    var dFromNumber = new Date(millis);
    if (!isNaN(dFromNumber.getTime())) return dFromNumber;
  }

  if (typeof raw === "string") {
    var t = raw.trim();
    if (!t) return null;

    // פורמט dd/MM/yyyy או דומים
    var m = t.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
    if (m) {
      var day   = parseInt(m[1], 10);
      var month = parseInt(m[2], 10);
      var year  = parseInt(m[3], 10);
      if (year < 100) year += 2000; // 25 → 2025

      var d = new Date(year, month - 1, day);
      if (!isNaN(d.getTime())) return d;
    }

    // fallback – ISO-8601
    var d2 = new Date(t);
    if (!isNaN(d2.getTime())) return d2;
  }

  return null;
}

// ==========================
//  מחזיר שם חג (או "") עבור תאריך
// ==========================
function getHolidayInfo(date) {
  if (!date) return "";
  if (!(date instanceof Date)) date = new Date(date);

  var year = date.getFullYear();
  var cacheKey = "holidays_" + year;
  var cached = CacheService.getScriptCache().get(cacheKey);
  if (!cached) return "";

  var holidays = JSON.parse(cached);
  var formatted = Utilities.formatDate(date, "GMT+3", "yyyy-MM-dd");
  var match = holidays.find(function (h) {
    return h.date === formatted;
  });
  return match ? match.title : "";
}

// ==========================
//  עוזר: שם גיליון בסגנון MM/yyyy
// ==========================
function _sheetNameForMonthYear_(monthNumber, year) {
  var mm = ("0" + monthNumber).slice(-2);
  return mm + "/" + year;
}

/**
 * מילוי עמודת "חג" עבור גיליון חודש ספציפי
 * לדוגמה:
 *   fillHolidayColumnForMonth(10);       // אוקטובר השנה
 *   fillHolidayColumnForMonth(12, 2025); // דצמבר 2025
 */
function fillHolidayColumnForMonth(monthNumber, yearOpt) {
  if (monthNumber < 1 || monthNumber > 12) {
    throw new Error("חודש לא תקין: " + monthNumber + " (צריך 1–12)");
  }

  var now = new Date();
  var year = yearOpt || now.getFullYear();
  var sheetName = _sheetNameForMonthYear_(monthNumber, year);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("⚠️ לא נמצא גיליון לחודש: " + sheetName);
    return;
  }

  var headerRow = 4; 
  var headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // ⚠️ פה השינוי: לגיליונות החודש התאריך נמצא בעמודת "subject"
  var subjectIndex    = headers.indexOf("subject") + 1;     // עמודת תאריך אמיתית (01/12/2025)
  var startDateIndex  = headers.indexOf("start date") + 1;  // אצלך זה שם היום (יום שני וכו')
  var holidayColIndex = headers.indexOf("חג") + 1;

  if (holidayColIndex === 0) {
    throw new Error('לא נמצאה העמודה "חג" בגיליון "' + sheetName + '"');
  }

  var dateColIndex = 0;

  if (subjectIndex > 0) {
    // בגיליונות החודש – זה העמודה החשובה
    dateColIndex = subjectIndex;
    Logger.log('📅 בגיליון "' + sheetName + '" משתמשים בעמודת "subject" כתאריך.');
  } else if (startDateIndex > 0) {
    // fallback (למקרה של גיליון אחר שבו "start date" כן מכיל תאריך)
    dateColIndex = startDateIndex;
    Logger.log('📅 בגיליון "' + sheetName + '" משתמשים בעמודת "start date" כתאריך.');
  } else {
    throw new Error('לא נמצאו העמודות "subject" או "start date" בגיליון "' + sheetName + '"');
  }

  // טוען חגים לשנה (כמו בקוד המקורי)
  preloadHolidayCache(year);

  var numRows = sheet.getLastRow() - 1;
  if (numRows <= 0) {
    Logger.log("ℹ️ אין שורות נתונים בגיליון " + sheetName);
    return;
  }

  var tz = Session.getScriptTimeZone() || "Asia/Jerusalem";

  // כאן אנחנו קוראים את עמודת התאריך **הנכונה** (subject / start date)
  var rawDates = sheet.getRange(2, dateColIndex, numRows, 1).getValues();
  var existingHolidays = sheet.getRange(2, holidayColIndex, numRows, 1).getValues();

  var results = [];
  Logger.log("==== התחלת מילוי חגים בגיליון " + sheetName + " ====");

  for (var i = 0; i < rawDates.length; i++) {
    var rowIndex = i + 2; // שורה בפועל
    var raw = rawDates[i][0];

    var d = normalizeDateCell(raw);
    if (!d) {
      Logger.log(
        "ℹ️ שורה " +
          rowIndex +
          " בגיליון " +
          sheetName +
          " – אין תאריך בעמודת תאריך (raw=" +
          raw +
          ") → skip, לא משנים את 'חג'"
      );
      results.push(existingHolidays[i]); // משאיר את הערך הקודם
      continue;
    }

    var yyyyMMdd = Utilities.formatDate(d, "GMT+3", "yyyy-MM-dd");
    var holidayTitle = getHolidayInfo(d);
    var dateStr = Utilities.formatDate(d, tz, "dd/MM/yyyy");

    Logger.log(
      "שורה " +
        rowIndex +
        " בגיליון " +
        sheetName +
        " – rawType=" +
        (raw instanceof Date ? "Date" : typeof raw) +
        ", parsed=" +
        yyyyMMdd +
        ", חג=" +
        (holidayTitle || "<אין>")
    );

    if (holidayTitle) {
      Logger.log(
        "🎉 חג '" +
          holidayTitle +
          "' נוסף לתאריך " +
          dateStr +
          " (שורה " +
          rowIndex +
          ", גיליון " +
          sheetName +
          ")"
      );
    }

    results.push([holidayTitle]);
  }

  sheet.getRange(2, holidayColIndex, results.length, 1).setValues(results);
  Logger.log("✅ עמודת 'חג' עודכנה בגיליון: " + sheetName);
}

/**
 * קיצור דרך: מילוי לחודש הנוכחי
 */
function fillHolidayColumnCurrentMonth() {
  var now = new Date();
  var month = now.getMonth() + 1; // 1–12
  var year  = now.getFullYear();
  fillHolidayColumnForMonth(month, year);
}

/**
 * קיצור דרך: מילוי לחודש הבא
 */
function fillHolidayColumnNextMonth() {
  var now = new Date();
  var next = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  var month = next.getMonth() + 1;   // 1–12
  var year  = next.getFullYear();    // כולל מעבר שנה (דצמבר→ינואר)
  fillHolidayColumnForMonth(month, year);
}

// נשאר כמו שהיה – בדיקת חג לוגית לפי תאריך
function isIsraelHoliday(date) {
  return getHolidayInfo(date) !== "";
}

/**
 * (אופציונלי) שמירת רשימת חגים גם לגיליון "חגים"
 */
function updateHolidaySheet() {
  var sheetName = "חגים";
  var year = new Date().getFullYear();
  preloadHolidayCache(year);

  var cacheKey = "holidays_" + year;
  var cached = CacheService.getScriptCache().get(cacheKey);
  if (!cached) {
    Logger.log("⚠️ אין חגים בקאש לשנה " + year);
    return;
  }

  var holidays = JSON.parse(cached).map(function (item) {
    return [item.date, item.title];
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }

  sheet.getRange(1, 1, 1, 2).setValues([["תאריך", "שם החג"]]);
  if (holidays.length) {
    sheet.getRange(2, 1, holidays.length, 2).setValues(holidays);
  }

  Logger.log("✅ עודכן גיליון החגים '" + sheetName + "' עם " + holidays.length + " שורות");
}

// ===== WRAPPER FOR TRIGGER / MENU =====
function addNextMonthHolidays() {
  Logger.log("=== מילוי חגים לחודש הבא התחיל ===");
  fillHolidayColumnNextMonth();
  Logger.log("=== מילוי חגים לחודש הבא הסתיים ===");
}
