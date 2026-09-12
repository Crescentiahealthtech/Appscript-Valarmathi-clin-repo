// =========================================================================
// 📅 DATE UTILITIES — one parser for every date the app reads
// Crescentia HealthTech / CresRx
// -------------------------------------------------------------------------
// Every module used to carry its own two-line date helper, and all of them
// were the same two lines:
//
//     var d = new Date(String(cellValue));
//     return isNaN(d.getTime()) ? null : d;
//
// That is wrong for the values this clinic's sheets actually hold, and it is
// wrong SILENTLY, which is why "Date of Admission" kept printing rubbish on
// the discharge summary long after the time-of-admission cell was fixed:
//
//   * "10/09/2026" is the 10th of September everywhere in this clinic, and
//     the 9th of October to `new Date`. No error — just the wrong day, and a
//     length of stay short by a month.
//   * "10-09-2026" is not a format `new Date` accepts at all. It returns
//     Invalid Date, the helper returns null, and the field prints empty or
//     falls back to echoing the raw cell text.
//   * A time-only cell ("Sat Dec 30 1899 23:31:00 GMT+0521") parses fine and
//     yields a date in 1899, which then formats as a real-looking admission
//     date.
//   * A numeric cell is a Sheets serial, not epoch milliseconds. Read as
//     epoch it lands in January 1970.
//
// Every one of those produces a plausible-looking wrong answer rather than a
// visible failure, so nothing downstream can detect it. cresc_toDate_ handles
// all of them explicitly and returns null when it genuinely cannot tell, so a
// caller can print "—" instead of a fabricated date.
//
// The per-module helpers (dsx_toDate_, hb_toDate_, acc_toDate_, ipr_ms_, …)
// are kept as thin wrappers around this file so their call sites do not move
// and each module keeps its own null/fallback convention.
// =========================================================================

/** Sheets' own epoch. A serial number counts days from here. */
var CRESC_SHEETS_EPOCH_MS = Date.UTC(1899, 11, 30);

/**
 * A value read out of a spreadsheet cell, as a Date — or null.
 *
 * Day-first is the house rule: this is an Indian clinic and every date typed
 * into these sheets, printed on these forms and read back by this staff is
 * dd/mm/yyyy. An unambiguous value still wins over the rule (13/07 can only
 * be the 13th of July; 2026-09-10 is ISO and reads as ISO).
 *
 * @param {*} v            cell value: Date, string, or number
 * @param {boolean=} allowTimeOnly  keep an 1899 time-only Date instead of
 *                                  rejecting it. Only dsx_time_() wants this.
 * @return {Date|null}
 */
function cresc_toDate_(v, allowTimeOnly) {
  if (v === null || v === undefined || v === '') return null;

  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    // A time-only cell is a Date pinned to the Sheets epoch. As a DATE it is
    // meaningless, and printing it is how "30-Dec-1899" reaches a document.
    if (!allowTimeOnly && v.getFullYear() < 1900) return null;
    return v;
  }

  if (typeof v === 'number' && isFinite(v)) {
    // Below ~100000 it is a Sheets serial (year 2173 at the top of that
    // range); above it, epoch milliseconds.
    if (Math.abs(v) < 100000) {
      // Read the serial in UTC, then rebuild it as LOCAL midnight. Adding the
      // serial to the epoch alone gives midnight UTC, which is the previous
      // calendar day in any negative-offset timezone.
      var u = new Date(CRESC_SHEETS_EPOCH_MS + Math.round(v * 86400000));
      return cresc_mk_(u.getUTCFullYear(), u.getUTCMonth() + 1, u.getUTCDate(),
                       u.getUTCHours(), u.getUTCMinutes(), u.getUTCSeconds());
    }
    var dn = new Date(v);
    return isNaN(dn.getTime()) ? null : dn;
  }

  var s = String(v).trim();
  if (!s) return null;

  // ISO first: yyyy-mm-dd, with or without a time part. Unambiguous, and the
  // shape every timestamp this app WRITES uses.
  var iso = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?/.exec(s);
  if (iso) {
    return cresc_mk_(+iso[1], +iso[2], +iso[3], +(iso[4] || 0), +(iso[5] || 0), +(iso[6] || 0));
  }

  // dd/mm/yyyy, dd-mm-yyyy, dd.mm.yyyy — with an optional trailing time.
  var dmy = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:[,\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?$/.exec(s);
  if (dmy) {
    var a = +dmy[1], b = +dmy[2], year = +dmy[3];
    if (year < 100) year += (year < 70) ? 2000 : 1900;

    var day, month;
    if (a > 12 && b <= 12)      { day = a; month = b; }   // 13/07 — day-first, certain
    else if (b > 12 && a <= 12) { day = b; month = a; }   // 07/13 — month-first, certain
    else                        { day = a; month = b; }   // ambiguous — house rule
    if (day < 1 || day > 31 || month < 1 || month > 12) return null;

    return cresc_mk_(year, month, day,
                     cresc_hour24_(+(dmy[4] || 0), dmy[7]), +(dmy[5] || 0), +(dmy[6] || 0));
  }

  // "10-Sep-2026", "10 Sep 2026", "10 September 2026" (+ optional time).
  var dMon = /^(\d{1,2})[\-\s]([A-Za-z]{3,})[\-\s](\d{4})(?:[,\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?)?$/.exec(s);
  if (dMon) {
    var mi = cresc_monthIndex_(dMon[2]);
    if (mi > 0) {
      return cresc_mk_(+dMon[3], mi, +dMon[1],
                       cresc_hour24_(+(dMon[4] || 0), dMon[7]), +(dMon[5] || 0), +(dMon[6] || 0));
    }
  }

  // Anything else — RFC strings, "Sep 10 2026", whatever a paste produced.
  var native = new Date(s);
  if (isNaN(native.getTime())) return null;
  if (!allowTimeOnly && native.getFullYear() < 1900) return null;
  return native;
}

/** Local-time Date from calendar parts, rejecting overflow (31 February). */
function cresc_mk_(y, mo, d, h, mi, se) {
  var dt = new Date(y, mo - 1, d, h || 0, mi || 0, se || 0, 0);
  if (isNaN(dt.getTime())) return null;
  // new Date(2026, 1, 31) silently rolls into March. A rolled date is a typo
  // in the sheet, not a date, and guessing at it is how bad data spreads.
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return null;
  return dt;
}

function cresc_hour24_(h, meridiem) {
  if (!meridiem) return h;
  var pm = /^[Pp]/.test(meridiem);
  if (pm && h < 12) return h + 12;
  if (!pm && h === 12) return 0;
  return h;
}

var CRESC_MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
                    'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

/** 1-12, or 0 when the word is not a month. */
function cresc_monthIndex_(word) {
  var w = String(word || '').toLowerCase().substring(0, 3);
  return CRESC_MONTHS.indexOf(w) + 1;
}

/** Formatted, or '' — never the raw cell text, and never a 1899 date. */
function cresc_fmt_(v, pattern, tz) {
  var d = cresc_toDate_(v);
  if (!d) return '';
  try {
    return Utilities.formatDate(d, tz || cresc_tz_(), pattern || 'dd-MMM-yyyy');
  } catch (e) { return ''; }
}

/**
 * A CLOCK TIME, formatted — the one place an 1899-epoch value is legitimate.
 *
 * Sheets stores a time-only cell as a Date pinned to its epoch, so
 * cresc_toDate_ rejects it: as a calendar date it is meaningless and printing
 * it is how "30-Dec-1899" reaches a document. Formatted as a time it is
 * exactly right, so this is the formatter a time column uses.
 */
function cresc_fmtTime_(v, pattern, tz) {
  var d = cresc_toDate_(v, true);
  if (!d) return '';
  try {
    return Utilities.formatDate(d, tz || cresc_tz_(), pattern || 'hh:mm a');
  } catch (e) { return ''; }
}

/** Epoch milliseconds, or 0 when the value is not a date. */
function cresc_ms_(v) {
  var d = cresc_toDate_(v);
  return d ? d.getTime() : 0;
}

function cresc_tz_() {
  try { return Session.getScriptTimeZone() || 'Asia/Kolkata'; }
  catch (e) { return 'Asia/Kolkata'; }
}

/**
 * Whole days from start to end, counting both ends — the way a ward counts a
 * length of stay. Midnight-anchored, so a 23:00 admission discharged at 01:00
 * the next morning is 2 days, not 1; and null when the start is unreadable,
 * so a caller shows "—" rather than a confident wrong number.
 */
function cresc_stayDays_(startVal, endVal) {
  var s = cresc_toDate_(startVal);
  if (!s) return null;
  var e = cresc_toDate_(endVal) || new Date();
  var d0 = Date.UTC(s.getFullYear(), s.getMonth(), s.getDate());
  var d1 = Date.UTC(e.getFullYear(), e.getMonth(), e.getDate());
  return Math.max(1, Math.round((d1 - d0) / 86400000) + 1);
}

/**
 * Self-check for the parser. Run from the script editor after touching this
 * file: every case here is a shape one of the clinic's sheets has actually
 * produced.
 */
function crescVerifyDates() {
  var f = function (v) { var d = cresc_toDate_(v); return d ? cresc_fmt_(d, 'yyyy-MM-dd HH:mm') : 'NULL'; };
  var cases = [
    ['10/09/2026',            '2026-09-10 00:00'],   // day-first house rule
    ['13/07/2026',            '2026-07-13 00:00'],   // day-first, unambiguous
    ['07/13/2026',            '2026-07-13 00:00'],   // month-first, unambiguous
    ['10-09-2026',            '2026-09-10 00:00'],   // new Date() cannot read this
    ['2026-09-10',            '2026-09-10 00:00'],
    ['2026-09-10T17:34:00',   '2026-09-10 17:34'],
    ['10-Sep-2026',           '2026-09-10 00:00'],
    ['10 September 2026',     '2026-09-10 00:00'],
    ['10/09/2026 05:34 PM',   '2026-09-10 17:34'],
    ['10/09/2026 12:15 AM',   '2026-09-10 00:15'],
    ['31/02/2026',            'NULL'],               // not a date, do not roll to March
    [new Date(1899, 11, 30, 23, 31), 'NULL'],        // time-only cell
    [46275,                   '2026-09-10 00:00'],   // Sheets serial
    ['',                      'NULL'],
    ['not a date',            'NULL']
  ];
  var fails = [];
  cases.forEach(function (c) {
    var got = f(c[0]);
    if (got !== c[1]) fails.push(String(c[0]) + ' -> ' + got + ' (expected ' + c[1] + ')');
  });
  var report = fails.length
    ? 'crescVerifyDates: ' + fails.length + ' failure(s)\n' + fails.join('\n')
    : 'crescVerifyDates: all ' + cases.length + ' cases pass.';
  Logger.log(report);
  return { success: fails.length === 0, message: report };
}
