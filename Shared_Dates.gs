// ============================================================================
// Shared_Dates.gs  —  Crescentia HealthTech / CresRx
// ONE date parser for the whole project.
// ----------------------------------------------------------------------------
// WHY THIS FILE EXISTS
//
// Every module had grown its own three-line parser, and all of them were the
// same three lines:
//
//     var d = new Date(String(v));
//     return isNaN(d.getTime()) ? null : d;
//
// That is correct only for values Sheets hands back as a real Date object.
// The moment a cell is TEXT — which happens constantly, because staff paste
// dates, imports arrive as strings, and a column formatted as Plain Text
// keeps everything typed into it as a string — `new Date()` is wrong in two
// different ways, and one of them is silent:
//
//   "13/09/2026"  -> Invalid Date.   Date of admission renders blank, length
//                                    of stay renders blank, the bed map shows
//                                    its "DOA?" badge, and the discharge
//                                    summary reports no admission date.
//   "01/02/2026"  -> 2 January 2026. NOT 1 February, which is what whoever
//                                    typed it meant. No error, no blank — a
//                                    wrong date on a discharge summary and a
//                                    wrong length of stay on a bill.
//
// The second is the reason this is a shared file rather than a patch in
// DS_Assembly.gs. India writes dates day-first; the V8 parser reads slashed
// dates month-first; there is no warning in between.
//
// WHAT IS PARSED
//
//   - Date objects (returned as-is once validated)
//   - epoch milliseconds, and Sheets' own day-serial numbers
//   - ISO 8601, with or without a time and a zone
//   - yyyy-MM-dd / yyyy/MM/dd
//   - dd/MM/yyyy, dd-MM-yyyy, dd.MM.yyyy  — DAY FIRST (see below)
//   - dd-MMM-yyyy, dd MMM yyyy, MMM dd yyyy, and the same with a time
//   - "dd-MMM-yyyy hh:mm a" — the format this project's own printers emit,
//     so a value that has been round-tripped through a print still parses
//
// THE DAY-FIRST RULE
//
// For an all-numeric slashed or dashed date the first field is read as the
// DAY, because that is what this clinic writes. Two guards keep that from
// making things worse than the engine default:
//
//   - a first field above 12 can only be a day, so "13/09/2026" is settled
//     by arithmetic rather than by convention;
//   - a SECOND field above 12 can only be a day, so an American "09/13/2026"
//     that reaches the sheet from an import is still read correctly.
//
// Only a genuinely ambiguous date — both fields 12 or below — falls to the
// day-first convention, and that is the convention the people typing it use.
// A four-digit first field (2026-09-13) is always ISO, never day-first.
//
// USAGE
//
//   cresc_parseDate_(v)            -> Date or null. Never throws.
//   cresc_formatDate_(v, pattern)  -> formatted string, or '' if unparseable
//   cresc_dateOnly_(v)             -> midnight local, for day comparisons
//   cresc_isSheetEpoch_(d)         -> true for a Sheets time-only cell
//   cresc_timeText_(v, pattern)    -> clock time from a time-only cell
//   cresc_daysBetween_(a, b)       -> whole days, or null
//
// Every module helper (dsx_toDate_, hb_toDate_, acc_toDate_, ipa_*, ipr_*,
// _dashToDate_, ipp_when_) now delegates here, so a date fixed in this file
// is fixed on every screen at once. Those helpers keep their names and their
// return contracts, so nothing that calls them had to change.
// ============================================================================

/** Month names, indexed for the abbreviated-month branch. */
var CRESC_MONTHS = {
  JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5,
  JUL: 6, AUG: 7, SEP: 8, SEPT: 8, OCT: 9, NOV: 10, DEC: 11
};

/**
 * Sheets stores a time-only cell as a Date pinned to 30 December 1899. Any
 * date at or before this is a clock time that lost its date, not a date.
 */
var CRESC_SHEET_EPOCH_MS = Date.UTC(1900, 0, 1);

function cresc_tz_() {
  try { return Session.getScriptTimeZone() || 'Asia/Kolkata'; }
  catch (e) { return 'Asia/Kolkata'; }
}

/** True when `d` is Sheets' 1899/1900 stand-in for a time with no date. */
function cresc_isSheetEpoch_(d) {
  return (d instanceof Date) && !isNaN(d.getTime()) && d.getTime() < CRESC_SHEET_EPOCH_MS;
}

/**
 * The project's one date parser.
 *
 * @param {*} v  anything a sheet cell, a form field or a JSON payload holds
 * @return {Date|null}  null for blank and for anything genuinely unreadable
 */
function cresc_parseDate_(v) {
  if (v === null || v === undefined || v === '') return null;

  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;

  // Numbers arrive two ways: epoch milliseconds from a client, and Sheets'
  // day serial from a cell read with getValues() on a numeric column.
  if (typeof v === 'number' && isFinite(v)) return cresc_fromNumber_(v);

  var s = String(v).trim();
  if (!s) return null;

  // A bare number that arrived as text is still a number.
  if (/^\d+(\.\d+)?$/.test(s)) return cresc_fromNumber_(parseFloat(s));

  var m;

  // ---- ISO 8601, and yyyy-MM-dd / yyyy/MM/dd ------------------------------
  // Handed straight to the engine, which reads these unambiguously. A
  // date-only ISO string is built by hand instead: the engine reads
  // "2026-09-13" as UTC midnight, which is the PREVIOUS day in any zone west
  // of Greenwich and shifts the date on screen for half the world.
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) return cresc_build_(+m[1], +m[2], +m[3], 0, 0, 0);

  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})[T ](.+)$/);
  if (m) {
    var isoTime = cresc_parseClock_(m[4]);
    if (isoTime) return cresc_build_(+m[1], +m[2], +m[3], isoTime.h, isoTime.min, isoTime.sec);
    var iso = new Date(s);
    return isNaN(iso.getTime()) ? null : iso;
  }

  // ---- dd-MMM-yyyy / dd MMM yyyy / MMM dd, yyyy (with optional time) ------
  var withMonthName = cresc_parseMonthName_(s);
  if (withMonthName) return withMonthName;

  // ---- all-numeric dd/MM/yyyy, dd-MM-yyyy, dd.MM.yyyy --------------------
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})(?:[T ,]+(.+))?$/);
  if (m) {
    var a = +m[1], b = +m[2], year = cresc_year_(+m[3]);
    var day, month;
    if (a > 12)      { day = a; month = b; }   // settled: only a day can exceed 12
    else if (b > 12) { day = b; month = a; }   // settled the other way: an imported MM/dd
    else             { day = a; month = b; }   // ambiguous -> day-first, as written here
    var clock = m[4] ? cresc_parseClock_(m[4]) : null;
    return cresc_build_(year, month, day,
      clock ? clock.h : 0, clock ? clock.min : 0, clock ? clock.sec : 0);
  }

  // ---- a clock time on its own -------------------------------------------
  // Kept as a time on the Sheets epoch, so cresc_isSheetEpoch_() can tell a
  // caller that this value carries no date.
  var only = cresc_parseClock_(s);
  if (only && /^\d{1,2}\s*:/.test(s)) {
    return new Date(1899, 11, 30, only.h, only.min, only.sec);
  }

  // ---- last resort: let the engine try ------------------------------------
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Epoch milliseconds, or a Sheets day serial, to a Date. */
function cresc_fromNumber_(n) {
  if (!isFinite(n)) return null;
  // A 0 in a date column is an EMPTY date column, never 30 December 1899.
  // Returning the Sheets epoch here would put "30-Dec-1899" on screen
  // wherever a date had simply not been filled in.
  if (n === 0) return null;
  // A Sheets day serial counts days from 30 December 1899; anything under
  // ~100,000 is a serial rather than a millisecond count (100,000 ms is two
  // minutes past the 1970 epoch, which is never a real clinical timestamp).
  if (Math.abs(n) < 100000) {
    // Built as a LOCAL date. Going via epoch milliseconds would land the
    // serial at UTC midnight, which formats as the previous day in any zone
    // west of Greenwich.
    var whole = Math.floor(n);
    var dayMs = Math.round((n - whole) * 86400000);
    var base = new Date(1899, 11, 30, 0, 0, 0, 0);
    base.setDate(base.getDate() + whole);
    if (dayMs) base = new Date(base.getTime() + dayMs);
    return isNaN(base.getTime()) ? null : base;
  }
  var d = new Date(n);
  return isNaN(d.getTime()) ? null : d;
}

/** Two-digit years: 70-99 are 1900s, everything else 2000s. */
function cresc_year_(y) {
  if (y >= 1000) return y;
  if (y >= 70 && y <= 99) return 1900 + y;
  return 2000 + y;
}

/** A local Date, guarding against month/day overflow silently rolling over. */
function cresc_build_(year, month, day, h, min, sec) {
  if (!year || month < 1 || month > 12 || day < 1 || day > 31) return null;
  var d = new Date(year, month - 1, day, h || 0, min || 0, sec || 0, 0);
  if (isNaN(d.getTime())) return null;
  // "31/02/2026" would otherwise become 3 March. A date nobody can have been
  // admitted on is an error to report, not a date to invent.
  if (d.getMonth() !== month - 1 || d.getDate() !== day) return null;
  return d;
}

/** "14:05", "2:05 PM", "02:05:30 pm" -> {h, min, sec}, or null. */
function cresc_parseClock_(text) {
  var s = String(text || '').trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?\s*([AaPp][Mm])?/);
  if (!m) return null;
  var h = +m[1], min = +m[2], sec = m[3] ? +m[3] : 0;
  var mer = m[4] ? m[4].toUpperCase() : '';
  if (mer === 'PM' && h < 12) h += 12;
  if (mer === 'AM' && h === 12) h = 0;
  if (h > 23 || min > 59 || sec > 59) return null;
  return { h: h, min: min, sec: sec };
}

/** "13-Sep-2026 05:34 PM", "13 September 2026", "Sep 13, 2026". */
function cresc_parseMonthName_(s) {
  var m = s.match(/^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s,]*(\d{2,4})(?:[T\s,]+(.+))?$/);
  if (m) {
    var mon = CRESC_MONTHS[m[2].substring(0, 3).toUpperCase()];
    if (mon === undefined) return null;
    var c = m[4] ? cresc_parseClock_(m[4]) : null;
    return cresc_build_(cresc_year_(+m[3]), mon + 1, +m[1],
      c ? c.h : 0, c ? c.min : 0, c ? c.sec : 0);
  }
  m = s.match(/^([A-Za-z]{3,9})[-\s]+(\d{1,2})[-\s,]+(\d{2,4})(?:[T\s,]+(.+))?$/);
  if (m) {
    var mon2 = CRESC_MONTHS[m[1].substring(0, 3).toUpperCase()];
    if (mon2 === undefined) return null;
    var c2 = m[4] ? cresc_parseClock_(m[4]) : null;
    return cresc_build_(cresc_year_(+m[3]), mon2 + 1, +m[2],
      c2 ? c2.h : 0, c2 ? c2.min : 0, c2 ? c2.sec : 0);
  }
  return null;
}

/**
 * A formatted date, or '' when the value is not a date.
 *
 * Returning '' rather than the raw text is deliberate: a cell holding
 * "asdf" should leave the field blank on a discharge summary, not print
 * "asdf" where a date belongs.
 */
function cresc_formatDate_(v, pattern) {
  var d = cresc_parseDate_(v);
  if (!d) return '';
  try { return Utilities.formatDate(d, cresc_tz_(), pattern || 'dd-MMM-yyyy'); }
  catch (e) { return ''; }
}

/** Midnight on the same day, for comparing dates without their times. */
function cresc_dateOnly_(v) {
  var d = cresc_parseDate_(v);
  if (!d) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
}

/** 'yyyy-MM-dd', the key every grouping in this project uses. */
function cresc_dayKey_(v) { return cresc_formatDate_(v, 'yyyy-MM-dd'); }

/**
 * A clock time for printing beside a date.
 *
 * A Sheets time-only cell is FORMATTED, never stringified: String() on it
 * yields "Sat Dec 30 1899 23:31:00 GMT+0521", which is what used to appear
 * next to the date of admission. Text that is already a time passes through.
 */
function cresc_timeText_(v, pattern) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return isNaN(v.getTime()) ? '' : cresc_formatDate_(v, pattern || 'hh:mm a');
  }
  var s = String(v).trim();
  var d = cresc_parseDate_(s);
  if (d && cresc_isSheetEpoch_(d)) return cresc_formatDate_(d, pattern || 'hh:mm a');
  return s;
}

/**
 * A DATE, AS TEXT, WITH NO TIME ON THE END OF IT.
 *
 * `String(cell)` on a Sheets date cell yields
 *
 *     Thu Jun 01 2028 00:00:00 GMT+0530 (India Standard Time)
 *
 * and that string was reaching printed documents — the expiry on a pharmacy
 * invoice was the reported case, but any column Sheets decided was a date
 * behaves the same way. Sheets coerces on entry, so a batch expiry typed as
 * "2028-06" becomes a real Date without anybody choosing that, and the same
 * column can hold text on one row and a Date on the next.
 *
 * This formats a Date and passes text through unchanged, so both rows print
 * the same. Use it for any value that is a date and NOT a moment: an expiry,
 * a date of birth, a review date, a due date.
 *
 * @param {*} v
 * @param {string} [pattern]  default 'dd-MMM-yyyy'
 * @return {string}
 */
function cresc_dateText_(v, pattern) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return isNaN(v.getTime()) ? '' : cresc_formatDate_(v, pattern || 'dd-MMM-yyyy');
  }
  return String(v).trim();
}

/**
 * The same, for a batch expiry, which is a MONTH and not a day.
 *
 * A medicine expires at the end of its printed month, so "Jun 2028" is the
 * honest rendering and "01-Jun-2028" invents a precision the pack does not
 * carry. Text already in yyyy-MM or MM/yyyy form passes through as typed.
 */
function cresc_expiryText_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return isNaN(v.getTime()) ? '' : cresc_formatDate_(v, 'MMM yyyy');
  }
  var s = String(v).trim();
  // "2028-06" and "2028-06-01" both read better as "Jun 2028".
  var m = s.match(/^(\d{4})-(\d{1,2})(?:-\d{1,2})?$/);
  if (m) {
    var d = cresc_build_(parseInt(m[1], 10), parseInt(m[2], 10), 1, 0, 0, 0);
    if (d) return cresc_formatDate_(d, 'MMM yyyy');
  }
  return s;
}

/** "13-Sep-2026 05:34 PM" from a date cell and a separate time cell. */
function cresc_dateTimeText_(dateVal, timeVal) {
  var d = cresc_formatDate_(dateVal, 'dd-MMM-yyyy');
  var t = cresc_timeText_(timeVal);
  if (!d) return t;
  return t ? (d + ' ' + t) : d;
}

/**
 * Whole days from `a` to `b` (default: now), or null if either is unreadable.
 * Compared at midnight, so an admission at 23:00 and a discharge at 01:00 the
 * next morning is one day, not zero.
 */
function cresc_daysBetween_(a, b) {
  var from = cresc_dateOnly_(a);
  if (!from) return null;
  var to = b ? cresc_dateOnly_(b) : cresc_dateOnly_(new Date());
  if (!to) return null;
  return Math.round((to.getTime() - from.getTime()) / 86400000);
}

/** Length of stay in days, counting the day of admission. Minimum 1. */
function cresc_los_(doa, dod) {
  var days = cresc_daysBetween_(doa, dod);
  if (days === null) return null;
  return Math.max(1, days + 1);
}

/** Epoch milliseconds, or 0 — for sorting rows by a date column. */
function cresc_ms_(v) {
  var d = cresc_parseDate_(v);
  return d ? d.getTime() : 0;
}

// ---------------------------------------------------------------------------
// SELF-TEST
// ---------------------------------------------------------------------------
/**
 * Run from the Apps Script editor after changing anything above.
 * Logs one line per failure and returns the number of failures.
 */
function cresc_testDates() {
  crescEditorOnly_('cresc_testDates');
  var fails = [];
  var check = function (label, got, want) {
    if (got !== want) fails.push(label + ': got ' + got + ', wanted ' + want);
  };
  var fmt = function (v) { return cresc_formatDate_(v, 'yyyy-MM-dd'); };

  check('dd/MM/yyyy unambiguous', fmt('13/09/2026'), '2026-09-13');
  check('dd/MM/yyyy ambiguous is day-first', fmt('01/02/2026'), '2026-02-01');
  check('imported MM/dd/yyyy still lands', fmt('09/13/2026'), '2026-09-13');
  check('dd-MM-yyyy', fmt('13-09-2026'), '2026-09-13');
  check('dd.MM.yyyy', fmt('13.09.2026'), '2026-09-13');
  check('ISO date keeps its day', fmt('2026-09-13'), '2026-09-13');
  check('ISO with time', fmt('2026-09-13T17:34:00'), '2026-09-13');
  check('dd-MMM-yyyy', fmt('13-Sep-2026'), '2026-09-13');
  check('dd MMM yyyy with time', fmt('13 Sep 2026 05:34 PM'), '2026-09-13');
  check('MMM dd, yyyy', fmt('Sep 13, 2026'), '2026-09-13');
  check('two-digit year', fmt('13/09/26'), '2026-09-13');
  check('impossible date is rejected', fmt('31/02/2026'), '');
  check('blank is blank', fmt(''), '');
  check('rubbish is blank', fmt('not a date'), '');
  check('Date object passes through', fmt(new Date(2026, 8, 13)), '2026-09-13');

  var t = cresc_parseDate_('13/09/2026 05:34 PM');
  check('time of day is kept', t ? t.getHours() : -1, 17);

  check('length of stay counts the admission day',
    cresc_los_('10/09/2026', '13/09/2026'), 4);
  check('same-day stay is one day',
    cresc_los_('13/09/2026', '13/09/2026'), 1);

  check('a time-only cell is recognised',
    cresc_isSheetEpoch_(cresc_parseDate_('05:34 PM')), true);
  check('a time-only cell prints as a time',
    cresc_timeText_(new Date(1899, 11, 30, 17, 34)), '05:34 PM');

  if (fails.length) {
    fails.forEach(function (f) { Logger.log('FAIL  ' + f); });
  } else {
    Logger.log('cresc_testDates: all ' + 20 + ' checks passed.');
  }
  return fails.length;
}
