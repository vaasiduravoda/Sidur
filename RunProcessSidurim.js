/*********************************************************
 * On-edit runner for "סידור כללי" A1:B1
 * - When user picks "שלח שיבוצים כעת", it calls callOtherAccount()
 * - Shows toast on success/failure
 * - Resets A1:B1 to "שיבוץ" and avoids recursive re-trigger
 *********************************************************/
var SHEET_NAME_SEND = 'סידור כללי';
var WATCH_RANGE_A1  = 'A1:B1';
// Hebrew labels for each remote function (used in the dropdown)
var TRIGGERS = {
  processSidurim_previewOnly:               'תצוגה בלבד - כל היומנים',
  processSidurim_real:                      'שיבוץ - כל היומנים',
  processSidurim_previewOnly_currentMonth:  'תצוגה בלבד - כל היומנים חודש נוכחי',
  processSidurim_previewOnly_nextMonth:     'תצוגה בלבד -כל היומנים חודש הבא',
  processSidurim_real_currentMonth:         'שיבוץ - כל היומנים חודש נוכחי',
  processSidurim_real_nextMonth:            'שיבוץ - כל היומנים חודש הבא',
  processSidurim_houseCalendars:            'שיבוץ - יומני בית',
  processSidurim_houseCalendars_previewOnly:'תצוגה בלבד - יומני בית',
};
var RESET_VALUE     = 'פעולות';
var SILENCE_FLAG    = 'ONEDIT_SILENCE_UNTIL_MS'; // Script Property key
var SILENCE_MS      = 3000;                      // ms

/**** Installable on-edit handler ****/
function onEditSendNow(e) {
  var ss = (e && e.source) ? e.source : SpreadsheetApp.getActive();
  if (!ss) return;

  // Debounce: ignore while we're resetting cells
  if (isSilenced_()) return;

  var r = e && e.range ? e.range : null;
  if (!r || !r.getSheet) return;
  var sh = r.getSheet();
  if (sh.getName() !== SHEET_NAME_SEND) return;

  // Only A1:B1
  var row = r.getRow();
  var col = r.getColumn();
  var colsEnd = col + r.getNumColumns() - 1;
  var intersects = (row === 1) && (col <= 2) && (colsEnd >= 1);
  if (!intersects) return;

  // // Map the selected dropdown value (Hebrew label) to action key(s)
  // var newVal = (typeof e.value !== 'undefined') ? e.value : r.getDisplayValue();
  // var actions = [];
  // for (var key in TRIGGERS) {
  //   if (Object.prototype.hasOwnProperty.call(TRIGGERS, key) && TRIGGERS[key] === newVal) {
  //     actions.push(key);
  //   }
  // }
  // if (!actions.length) return; // not one of our trigger labels

  // Map the selected dropdown value (Hebrew label) to action key(s)
var newVal = (typeof e.value !== 'undefined') ? e.value : r.getDisplayValue();
Logger.log('Selected value: "' + newVal + '"');

var actions = [];
for (var key in TRIGGERS) {
  if (Object.prototype.hasOwnProperty.call(TRIGGERS, key)) {
    var triggerVal = TRIGGERS[key];
    Logger.log('Comparing key=' + key + ', trigger="' + triggerVal + '", match=' + (triggerVal === newVal));
    if (triggerVal === newVal) {
      actions.push(key);
    }
  }
}

Logger.log('Found actions: ' + JSON.stringify(actions));


  // Silence re-entries while we run & reset
  silenceFor_(SILENCE_MS);

  try {
    ss.toast('⏳ שולח שיבוצים כעת...', 'סטטוס', 5);
    callOtherAccount(actions); // pass selected remote actions
    ss.toast('✅ השיבוצים נשלחו (בדוק Executions → Time-driven)', 'סטטוס', 8);
  } catch (err) {
    ss.toast('❌ שליחה נכשלה: ' + String(err), 'שגיאה', 10);
  } finally {
    try {
      sh.getRange(WATCH_RANGE_A1).setValues([[RESET_VALUE, RESET_VALUE]]);
    } catch (resetErr) {
      ss.toast('⚠️ שגיאה בשחזור הערכים: ' + String(resetErr), 'אזהרה', 8);
    } finally {
      silenceFor_(SILENCE_MS);
    }
  }
}

/**** Helpers (keep once in the project) ****/
function isSilenced_() {
  var now = Date.now();
  var until = Number(PropertiesService.getScriptProperties().getProperty(SILENCE_FLAG) || '0');
  return now < until;
}
function silenceFor_(ms) {
  var until = Date.now() + Math.max(0, ms | 0);
  PropertiesService.getScriptProperties().setProperty(SILENCE_FLAG, String(until));
}

/*********************************************************
 * Calls the target Web App.
 * - First PING (auth only), then RUN (schedules background worker)
 * - Defensive JSON checks with helpful error messages
 *********************************************************/
function callOtherAccount(actions) {
  const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzJlLfieTizknyY2ASS7XE6ZBoFAdIb3HpXk1mXfRDifYKlO-SKpkUBWiDV4g-phxZF/exec';
  const SECRET = 'testsec';

  const ss = SpreadsheetApp.getActive();
  const statusRangeA1 = "'סידור כללי'!C1";
  const jobId = Utilities.getUuid();

  ss.getRange(statusRangeA1).setValue('⏳ התחיל… (' + jobId.slice(0, 8) + ')');

  // Tell the target Web App which functions to run in the other account.
  // Make sure these names match the actionMap/doPost implementation there.
  const payload = {
    secret: SECRET,
    jobId: jobId,
    originSpreadsheetId: ss.getId(),
    statusRangeA1: statusRangeA1,
    actions: Array.isArray(actions) ? actions : []
  };

  const res = UrlFetchApp.fetch(WEB_APP_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { 'X-Webhook-Secret': SECRET },
    muteHttpExceptions: true,
    followRedirects: true   // ← allow the googleusercontent echo hop
  });

  const code = res.getResponseCode();
  const ct  = (res.getHeaders()['Content-Type'] || '').toLowerCase();
  const txt = res.getContentText() || '';

  if (code !== 200) throw new Error('HTTP ' + code + ' from target. Body: ' + txt.slice(0, 300));
  if (!ct.includes('application/json') || txt.trim().startsWith('<')) {
    throw new Error('Non-JSON response. Snippet: ' + txt.slice(0, 300));
  }
  const body = JSON.parse(txt);
  if (!body.ok) throw new Error('Remote error: ' + (body.error || 'unknown') + ' — reqId=' + (body.reqId || '-'));
}
