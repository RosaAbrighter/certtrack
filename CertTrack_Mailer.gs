/*************************************************************************
 *  CertTrack Mailer  —  Google Apps Script email engine
 *  Built for Rosa's certificate renewal tracker
 *
 *  WHAT IT DOES
 *  Once a day it reads a Google Sheet of employee certificates, works out
 *  how many days until each one expires, and emails a renewal reminder at
 *  the 90 / 60 / 30 / 20-day marks (and once more the day it expires).
 *  The employee gets the notice; HR gets one daily digest of everyone due.
 *  Each reminder is sent only once per window — it remembers what it sent.
 *
 *  ── SETUP (about 5 minutes, one time) ─────────────────────────────────
 *  1. Make a Google Sheet. Put these headers in row 1 (order doesn't
 *     matter, capitalisation doesn't matter):
 *        Employee Name | Employee Email | Certificate | Issue Date | Expiry Date
 *     This is EXACTLY what CertTrack's "Export" button produces, so you can
 *     paste an export straight in. (Extra columns are fine and ignored.)
 *  2. In the Sheet menu:  Extensions ▸ Apps Script.
 *  3. Delete whatever is there, paste this whole file, click Save.
 *  4. Edit the CONFIG block below — set HR_EMAIL and ORG_NAME at minimum.
 *  5. Run the function `installDailyTrigger` once (pick it in the toolbar
 *     dropdown, press Run). Approve the permission prompt the first time.
 *  6. Done. It now runs every morning on its own. To test immediately,
 *     run `sendTest` — it emails YOU a preview of today's alerts only.
 *
 *  Free Gmail sends up to ~100 recipients/day — plenty for under 50 staff.
 *************************************************************************/

var CONFIG = {
  HR_EMAIL:   "hr@company.com",        // <-- Rosa / HR admin. Gets the daily digest.
  HR_NAME:    "Rosa",                  // signature name on employee emails
  ORG_NAME:   "Your Organization",     // appears in the email body + subject

  WINDOWS:    [90, 60, 30, 20],        // days-before-expiry to alert. Edit freely.
  ALERT_ON_EXPIRY: true,               // also send the day it expires
  SEND_HR_DIGEST: true,                // one summary email to HR per day
  CC_HR_ON_EACH:  false,               // OR cc HR on every individual email instead

  SHEET_NAME: "",                      // leave "" to use the first/active sheet
  RUN_HOUR:   7                        // hour of day the daily check runs (0-23)
};

/* Header aliases — so it works whether columns say "Expiry Date", "Expires", etc. */
var COLS = {
  name:   ["employee name", "name", "employee"],
  email:  ["employee email", "email"],
  cert:   ["certificate", "certificate name", "cert", "license"],
  issue:  ["issue date", "issued", "issue"],
  expiry: ["expiry date", "expiry", "expires", "expiration", "expiration date"],
  sent:   ["alerts sent"]   // managed automatically; created if missing
};

/*======================================================================*/
/*  MAIN — this is what the daily trigger calls                          */
/*======================================================================*/
function checkAndSend() { return _run(false); }

/* Preview: emails only YOU (the script owner) today's alerts, sends nothing
   to employees, and does NOT mark anything as sent. */
function sendTest() { return _run(true); }

function _run(testMode) {
  var sheet = CONFIG.SHEET_NAME
    ? SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME)
    : SpreadsheetApp.getActive().getSheets()[0];
  if (!sheet) throw new Error("Sheet not found. Check CONFIG.SHEET_NAME.");

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) { Logger.log("No data rows."); return; }

  var header = data[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var ix = {};
  for (var key in COLS) ix[key] = _findCol(header, COLS[key]);

  if (ix.name < 0 || ix.cert < 0 || ix.expiry < 0)
    throw new Error("Missing a required column (need Employee Name, Certificate, Expiry Date).");

  // Ensure an "Alerts Sent" column exists for de-duplication.
  if (ix.sent < 0) {
    ix.sent = header.length;
    sheet.getRange(1, ix.sent + 1).setValue("Alerts Sent");
  }

  var windows = CONFIG.WINDOWS.slice().sort(function (a, b) { return a - b; }); // ascending
  var maxWin = windows[windows.length - 1];
  var today = _midnight(new Date());
  var digest = [];          // items to summarise to HR
  var sentUpdates = [];      // {row, value} writes for the Alerts Sent column
  var sentCount = 0;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var name  = String(row[ix.name]  || "").trim();
    var email = ix.email  >= 0 ? String(row[ix.email]  || "").trim() : "";
    var cert  = String(row[ix.cert]  || "").trim();
    var expDate = _parseDate(row[ix.expiry]);
    if (!name || !cert || !expDate) continue;

    var daysLeft = Math.round((_midnight(expDate) - today) / 86400000);

    var alreadySent = String(row[ix.sent] || "").split(",")
      .map(function (s) { return s.trim(); }).filter(Boolean);

    // Renewed? If it's comfortably valid again, clear history so future cycles fire.
    if (daysLeft > maxWin) {
      if (alreadySent.length && !testMode)
        sentUpdates.push({ row: r + 1, value: "" });
      continue;
    }

    // Which alert tokens apply right now?
    var triggered = [];
    if (daysLeft < 0 && CONFIG.ALERT_ON_EXPIRY) triggered.push("EXPIRED");
    if (daysLeft >= 0) {
      for (var w = 0; w < windows.length; w++)
        if (daysLeft <= windows[w]) triggered.push(String(windows[w]));
    }

    // Only the windows we haven't emailed yet.
    var fresh = triggered.filter(function (t) { return alreadySent.indexOf(t) < 0; });
    if (!fresh.length) continue;

    // Most urgent fresh window = the one we actually email about today.
    var token = fresh.indexOf("EXPIRED") > -1
      ? "EXPIRED"
      : fresh.map(Number).sort(function (a, b) { return a - b; })[0].toString();

    var item = { name: name, email: email, cert: cert,
                 daysLeft: daysLeft, token: token, expiry: expDate };
    digest.push(item);

    if (!testMode) {
      if (email) {
        var msg = _composeEmail(item);
        var opts = { name: CONFIG.ORG_NAME + " — Certifications" };
        if (CONFIG.CC_HR_ON_EACH && CONFIG.HR_EMAIL) opts.cc = CONFIG.HR_EMAIL;
        MailApp.sendEmail(email, msg.subject, msg.body, opts);
        sentCount++;
      }
      // Record every fresh token so nothing repeats, even if email was blank.
      var merged = alreadySent.concat(fresh)
        .filter(function (v, i, a) { return a.indexOf(v) === i; });
      sentUpdates.push({ row: r + 1, value: merged.join(",") });
    }
  }

  // Write de-dup column back in one pass.
  sentUpdates.forEach(function (u) {
    sheet.getRange(u.row, ix.sent + 1).setValue(u.value);
  });

  // HR digest / test summary.
  if (digest.length) {
    if (testMode) {
      MailApp.sendEmail(Session.getActiveUser().getEmail(),
        "[TEST] CertTrack — " + digest.length + " alert(s) would send today",
        _digestBody(digest, true));
    } else if (CONFIG.SEND_HR_DIGEST && CONFIG.HR_EMAIL) {
      MailApp.sendEmail(CONFIG.HR_EMAIL,
        "Certificate renewals due — " + _todayStr() + " (" + digest.length + ")",
        _digestBody(digest, false));
    }
  }

  Logger.log((testMode ? "[TEST] " : "") + "Alerts queued: " + digest.length +
             ", employee emails sent: " + (testMode ? 0 : sentCount));
}

/*======================================================================*/
/*  EMAIL TEMPLATES                                                      */
/*======================================================================*/
function _composeEmail(it) {
  var expired = it.token === "EXPIRED";
  var when = _fmt(it.expiry);
  var subject = expired
    ? "Action needed: " + it.cert + " has expired"
    : it.cert + " expires in " + it.daysLeft + " days";

  var body = "Hello " + it.name + ",\n\n";
  body += expired
    ? "Our records show your " + it.cert + " expired on " + when + " ("
        + Math.abs(it.daysLeft) + " days ago). Please renew it as soon as "
        + "possible and send confirmation to " + CONFIG.ORG_NAME + "."
    : "This is a reminder that your " + it.cert + " is due to expire on "
        + when + " — " + it.daysLeft + " days from today. Please begin the "
        + "renewal process so your certification stays current.";
  body += "\n\nThank you,\n" + CONFIG.HR_NAME + "\n" + CONFIG.ORG_NAME;
  return { subject: subject, body: body };
}

function _digestBody(items, isTest) {
  items.sort(function (a, b) { return a.daysLeft - b.daysLeft; });
  var lines = [];
  lines.push(isTest
    ? "Preview only — no employee emails were sent.\n"
    : "Certificate renewals due as of " + _todayStr() + ":\n");
  items.forEach(function (it) {
    var tag = it.token === "EXPIRED" ? "EXPIRED" : it.token + "-day";
    lines.push("• [" + tag + "]  " + it.name + " — " + it.cert +
      "  (expires " + _fmt(it.expiry) +
      (it.token === "EXPIRED" ? ", " + Math.abs(it.daysLeft) + "d ago"
                              : ", " + it.daysLeft + "d left") + ")" +
      (it.email ? "" : "   ⚠ no email on file"));
  });
  lines.push("\n— CertTrack Mailer");
  return lines.join("\n");
}

/*======================================================================*/
/*  TRIGGER MANAGEMENT                                                   */
/*======================================================================*/
function installDailyTrigger() {
  // Remove any old copies so you don't stack duplicates.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "checkAndSend") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkAndSend")
    .timeBased().everyDays(1).atHour(CONFIG.RUN_HOUR).create();
  Logger.log("Daily trigger installed for ~" + CONFIG.RUN_HOUR + ":00.");
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "checkAndSend") ScriptApp.deleteTrigger(t);
  });
  Logger.log("Triggers removed.");
}

/*======================================================================*/
/*  UTILITIES                                                            */
/*======================================================================*/
function _findCol(header, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var k = header.indexOf(aliases[i]);
    if (k > -1) return k;
  }
  return -1;
}
function _parseDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  var s = String(v || "").trim(); if (!s) return null;
  var m;
  if (m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/)) return new Date(+m[1], +m[2]-1, +m[3]);
  if (m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)) return new Date(+m[3], +m[1]-1, +m[2]);
  var d = new Date(s); return isNaN(d) ? null : d;
}
function _midnight(d) { var x = new Date(d); x.setHours(0,0,0,0); return x; }
function _fmt(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM d, yyyy");
}
function _todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy");
}
