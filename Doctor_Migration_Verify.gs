// ============================================================================
// Doctor_Migration_Verify.gs  —  Crescentia HealthTech
// TEMPORARY. Delete this file once Phase 3 is signed off.
// Compares the NEW generated slot grid against the OLD hardcoded CLINIC_SLOTS
// grid minus legacy 'Blocked' rows, for a given doctor and date.
// ============================================================================

var VERIFY_LEGACY_SLOTS = [
  "10:00 AM","10:15 AM","10:30 AM","10:45 AM","11:00 AM","11:15 AM","11:30 AM","11:45 AM",
  "12:00 PM","12:15 PM","12:30 PM","12:45 PM","05:00 PM","05:15 PM","05:30 PM","05:45 PM",
  "06:00 PM","06:15 PM","06:30 PM","06:45 PM","07:00 PM","07:15 PM","07:30 PM","07:45 PM",
  "08:00 PM","08:15 PM","08:30 PM","08:45 PM"
];

/**
 * RUN THIS FROM THE APPS SCRIPT EDITOR. Read the log.
 * @param {string} dateStr  "yyyy-MM-dd"
 * @param {string} doctorId defaults to DC_DEFAULT_DOCTOR
 */
function verifySlotGrid(dateStr, doctorId) {
  var docId = dc_str_(doctorId) || DC_DEFAULT_DOCTOR;
  var dateKey = dc_dateKey_(dateStr);
  if (!dateKey) return "Pass a date as 'yyyy-MM-dd'.";

  var dayName = DS_WEEKDAYS[ds_parseDateKey_(dateKey).getDay()];

  // --- OLD MODEL: hardcoded slots minus this date's legacy Blocked rows ----
  var blocked = {};
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appointments");
  if (sh && sh.getLastRow() > 1) {
    var data = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (dc_dateKey_(data[i][3]) !== dateKey) continue;
      if (dc_str_(data[i][6]) !== "Blocked") continue;
      blocked[dc_to24_(data[i][4])] = true;
    }
  }
  var oldSet = {}, oldList = [];
  VERIFY_LEGACY_SLOTS.forEach(function (s) {
    var k = dc_to24_(s);
    if (blocked[k]) return;
    oldSet[k] = true; oldList.push(k);
  });

  // --- NEW MODEL: generated slots ----------------------------------------
  var gen = ds_generateSlots_(docId, dateKey);
  var newSet = {}, newList = [];
  gen.forEach(function (s) { newSet[s.time24] = true; newList.push(s.time24); });

  // --- diff ---------------------------------------------------------------
  var missing = oldList.filter(function (t) { return !newSet[t]; });
  var extra   = newList.filter(function (t) { return !oldSet[t]; });

  var out = [];
  out.push("=== SLOT GRID VERIFICATION ===");
  out.push("Doctor : " + docId);
  out.push("Date   : " + dateKey + " (" + dayName + ")");
  out.push("Legacy blocked rows found on this date: " + Object.keys(blocked).length);
  out.push("");
  out.push("OLD grid slot count : " + oldList.length);
  out.push("NEW grid slot count : " + newList.length);
  out.push("");

  if (missing.length === 0 && extra.length === 0) {
    out.push("RESULT: MATCH. Every slot is identical.");
  } else {
    out.push("RESULT: MISMATCH.");
    if (missing.length) {
      out.push("  Present in OLD, absent in NEW (" + missing.length + "):");
      out.push("    " + missing.map(dc_to12_).join(", "));
    }
    if (extra.length) {
      out.push("  Present in NEW, absent in OLD (" + extra.length + "):");
      out.push("    " + extra.map(dc_to12_).join(", "));
    }
    out.push("");
    out.push("  Expected causes: Sunday (legacy array had no weekday awareness);");
    out.push("  a LEAVE/BLOCK exception you added; or the seeded template window");
    out.push("  differing from the legacy 10:00-12:45 / 17:00-20:45 pattern.");
  }

  out.push("");
  out.push("--- Booking safety check ---");
  var booked = ds_bookingsByTime_(docId, dateKey);
  var orphaned = Object.keys(booked).filter(function (t) { return !newSet[t]; });
  if (orphaned.length === 0) {
    out.push("Every existing booking on this date falls inside the new grid.");
  } else {
    out.push("WARNING: " + orphaned.length + " booking time(s) fall OUTSIDE the new grid:");
    orphaned.forEach(function (t) {
      booked[t].forEach(function (b) {
        out.push("    " + dc_to12_(t) + "  " + b.patientId + "  " + b.patientName + "  [" + b.status + "]");
      });
    });
    out.push("These are NOT lost - the UI shows them under 'Off-schedule bookings'.");
  }

  var text = out.join("\n");
  Logger.log(text);
  return text;
}

/** Sweeps the next N days and reports only the dates that mismatch. */
function verifySlotGridRange(startDateStr, days, doctorId) {
  var start = ds_parseDateKey_(dc_dateKey_(startDateStr));
  if (!start) return "Pass a start date as 'yyyy-MM-dd'.";
  var n = dc_int_(days) || 14;
  if (n > 60) n = 60;

  var lines = ["=== RANGE SWEEP: " + n + " day(s) from " + dc_dateKey_(startDateStr) + " ==="];
  var cursor = new Date(start.getTime());
  for (var i = 0; i < n; i++) {
    var key = Utilities.formatDate(cursor, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var res = verifySlotGrid(key, doctorId);
    var verdict = (res.indexOf("RESULT: MATCH") !== -1) ? "MATCH" : "MISMATCH";
    var warn = (res.indexOf("WARNING:") !== -1) ? "  <-- has off-schedule bookings" : "";
    lines.push(key + "  " + DS_WEEKDAYS[cursor.getDay()].substring(0, 3) + "  " + verdict + warn);
    cursor.setDate(cursor.getDate() + 1);
  }
  var text = lines.join("\n");
  Logger.log(text);
  return text;
}