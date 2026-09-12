// ============================================================================
// DS_Data.gs  —  Crescentia HealthTech / CresRx
// IP Discharge Summary Engine · Phase 2 · data access, payload, config, hashing
// ----------------------------------------------------------------------------
// Prefix rules: see the header of DS_Setup.gs. Everything here is `dsx_` /
// `DSX_` because `ds_` belongs to the doctor session store and scheduler.
//
// Two rules govern every read in this file:
//   1. HEADER MAPS ONLY. Discovery found the live header order differing from
//      the declared constants on IP_Timeline_DB and IP_CaseSheets_DB, and
//      found settleDischarge writing IP_Admissions by hard-coded index. No
//      positional access appears anywhere in this module.
//   2. NO getDataRange() ON CLINICAL SHEETS. Row lookups go through TextFinder
//      restricted to one column, then re-read that row to verify. A 30-day ICU
//      stay must not pull the whole timeline into memory.
// ============================================================================

var DSX_MAX_CHUNK = 45000;   // Sheets caps a cell at 50,000 characters.
var DSX_CHUNKS    = 8;
var DSX_SCHEMA_VERSION = 1;

var DSX_STATUS = {
  GENERATED:             'GENERATED',
  IN_PREPARATION:        'IN_PREPARATION',
  PENDING_SIGNATURE:     'PENDING_SIGNATURE',
  RETURNED:              'RETURNED',
  SIGNED:                'SIGNED',
  AMENDMENT_IN_PROGRESS: 'AMENDMENT_IN_PROGRESS',
  CANCELLED:             'CANCELLED'
};

var DSX_DISCHARGE_TYPES = ['NORMAL', 'LAMA', 'DAMA', 'REFERRED', 'DEATH', 'ABSCONDED'];

/** Sections no role may edit; they are rendered from workflow data. */
var DSX_READONLY_SECTIONS = ['PATIENT_BANNER', 'SIGNATURES'];

// ---------------------------------------------------------------------------
// SECTION A — scalars, ids, time
// ---------------------------------------------------------------------------

function dsx_str_(v)   { return (v === null || v === undefined) ? '' : String(v).trim(); }
function dsx_upper_(v) { return dsx_str_(v).toUpperCase(); }
function dsx_int_(v)   { var n = parseInt(v, 10); return isNaN(n) ? 0 : n; }
function dsx_bool_(v)  {
  var s = dsx_upper_(v);
  return s === 'TRUE' || s === 'YES' || s === 'Y' || s === '1' || v === true;
}

/** Canonical patient-ID form. Lower-case rows exist historically (Discovery §11.9). */
function dsx_pid_(v) { return dsx_upper_(v); }

/** Canonical IP-number form. */
function dsx_ip_(v) { return dsx_upper_(v).replace(/\s+/g, ''); }

function dsx_tenant_() {
  try { return getTenantId_(); } catch (e) { return 'VALARMATHI'; }
}

function dsx_tz_() {
  try { return Session.getScriptTimeZone() || 'Asia/Kolkata'; } catch (e) { return 'Asia/Kolkata'; }
}

function dsx_nowIso_() { return new Date().toISOString(); }

function dsx_toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = dsx_str_(v);
  if (!s) return null;
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function dsx_fmt_(v, pattern) {
  var d = dsx_toDate_(v);
  if (!d) return '';
  try { return Utilities.formatDate(d, dsx_tz_(), pattern || 'dd-MMM-yyyy hh:mm a'); }
  catch (e) { return ''; }
}

/**
 * A clock time, for printing beside a date.
 *
 * Sheets stores a time-only cell as a Date pinned to the spreadsheet epoch,
 * so String() on it yields "Sat Dec 30 1899 23:31:00 GMT+0521" — which is
 * exactly what appeared next to the date of admission on the summary. A Date
 * is therefore FORMATTED, never stringified; text that is already a time
 * ("05:34 PM") passes through untouched.
 */
function dsx_time_(v, pattern) {
  if (v instanceof Date) {
    return isNaN(v.getTime()) ? '' : dsx_fmt_(v, pattern || 'hh:mm a');
  }
  return dsx_str_(v);
}

/** "10-Sep-2026 05:34 PM" from a date cell and a separate time cell. */
function dsx_dateTime_(dateVal, timeVal) {
  var d = dsx_fmt_(dateVal, 'dd-MMM-yyyy');
  var t = dsx_time_(timeVal);
  if (!d) return t;
  return t ? (d + ' ' + t) : d;
}

function dsx_newEventId_() {
  return 'DSE-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toUpperCase();
}

/** One summary per admission, and the ID is derivable from the IP number. */
function dsx_summaryIdFor_(ipNumber) {
  var ip = dsx_ip_(ipNumber);
  if (!ip) throw new Error('VALIDATION_FAILED: an IP number is required.');
  return 'DS-' + ip;
}

function dsx_ipFromSummaryId_(summaryId) {
  var s = dsx_upper_(summaryId);
  return s.indexOf('DS-') === 0 ? s.substring(3) : '';
}

// ---------------------------------------------------------------------------
// SECTION B — sheet access
// ---------------------------------------------------------------------------

var DSX_HEADER_CACHE = {};   // per-execution only; Apps Script discards it after the call

function dsx_ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/**
 * A DS_* sheet, or null. This module never creates a sheet on a read path —
 * that is setupDischargeSummaryModule()'s job, so a missing sheet surfaces as
 * a clear setup error instead of an empty screen.
 */
function dsx_sheet_(name) {
  return dsx_ss_().getSheetByName(name);
}

function dsx_requireSheet_(name) {
  var sh = dsx_sheet_(name);
  if (!sh) {
    throw new Error('VALIDATION_FAILED: sheet "' + name +
                    '" is missing. Run setupDischargeSummaryModule() once.');
  }
  return sh;
}

/** {HeaderName: zeroBasedIndex}, cached for the life of this execution. */
function dsx_headerMap_(sheet) {
  var key = sheet.getSheetId();
  if (DSX_HEADER_CACHE[key]) return DSX_HEADER_CACHE[key];

  var lastCol = Math.max(1, sheet.getLastColumn());
  var head = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < head.length; i++) {
    var h = dsx_str_(head[i]);
    if (h && map[h] === undefined) map[h] = i;
  }
  DSX_HEADER_CACHE[key] = map;
  return map;
}

function dsx_resetHeaderCache_() { DSX_HEADER_CACHE = {}; }

/**
 * Row numbers whose `headerName` column equals `key`, via TextFinder scoped to
 * that one column, then verified by re-reading the cell. TextFinder matches on
 * the displayed string, so the verification pass is what makes this safe for a
 * date- or number-formatted column.
 *
 * @return {Array<number>} 1-based row numbers, ascending
 */
function dsx_findRowsByKey_(sheet, headerName, key) {
  var out = [];
  var needle = dsx_upper_(key);
  if (!needle) return out;

  var map = dsx_headerMap_(sheet);
  var col = map[headerName];
  if (col === undefined) return out;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return out;

  var range = sheet.getRange(2, col + 1, lastRow - 1, 1);
  var found;
  try {
    found = range.createTextFinder(key)
                 .matchEntireCell(true)
                 .matchCase(false)
                 .ignoreDiacritics(false)
                 .findAll();
  } catch (e) {
    found = [];
  }

  // Verification re-read: TextFinder works on display text, so confirm the
  // stored value really normalises to the key before acting on the row.
  var values = null;
  if (!found.length) {
    // Fall back to a single bounded column read. A TextFinder miss on a
    // formatted column would otherwise look like "no such admission".
    values = range.getDisplayValues();
    for (var i = 0; i < values.length; i++) {
      if (dsx_upper_(values[i][0]) === needle) out.push(i + 2);
    }
    return out;
  }

  // One bounded read spanning the candidate rows, not one read per candidate.
  // A summary with fifty workflow events used to cost fifty round trips here
  // alone, and those seconds are the difference between a reply and a
  // request that dies on the way back.
  var rows = [];
  for (var j = 0; j < found.length; j++) rows.push(found[j].getRow());
  rows.sort(function (a, b) { return a - b; });

  var first = rows[0], last = rows[rows.length - 1];
  var block = sheet.getRange(first, col + 1, last - first + 1, 1).getDisplayValues();
  for (var k = 0; k < rows.length; k++) {
    var r = rows[k];
    if (dsx_upper_(block[r - first][0]) === needle && out.indexOf(r) === -1) out.push(r);
  }
  out.sort(function (a, b) { return a - b; });
  return out;
}

/**
 * The named columns of every data row, in as few round trips as the layout
 * allows.
 *
 * Reading a row at a time is what makes the discharge screens slow, but
 * getDataRange() on DS_Snapshots would drag eight payload chunks per row into
 * memory. So: take the columns actually wanted, group them into contiguous
 * runs, and read one range per run. For DS_Snapshots that is two reads
 * (the metadata before the payload, and the two columns after it) instead of
 * one per snapshot, and the payload columns are never touched.
 *
 * @param {Sheet} sheet
 * @param {Array<string>} headerNames
 * @param {Array<number>=} rowNumbers  1-based; omit for every data row
 * @return {Array<Object>} row objects keyed by header name, plus _row
 */
function dsx_readColumns_(sheet, headerNames, rowNumbers) {
  var map = dsx_headerMap_(sheet);
  var wanted = [];
  headerNames.forEach(function (h) {
    if (map[h] !== undefined && wanted.indexOf(map[h]) === -1) wanted.push(map[h]);
  });
  if (!wanted.length) return [];
  wanted.sort(function (a, b) { return a - b; });

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var rows;
  if (rowNumbers && rowNumbers.length) {
    rows = rowNumbers.slice().sort(function (a, b) { return a - b; });
  } else {
    rows = [];
    for (var r = 2; r <= lastRow; r++) rows.push(r);
  }
  if (!rows.length) return [];

  var top = rows[0], bottom = rows[rows.length - 1];
  var height = bottom - top + 1;

  // Contiguous runs of columns, tolerating a gap of one so two adjacent-ish
  // fields do not cost two round trips.
  var runs = [], cur = [wanted[0], wanted[0]];
  for (var i = 1; i < wanted.length; i++) {
    if (wanted[i] - cur[1] <= 2) cur[1] = wanted[i];
    else { runs.push(cur); cur = [wanted[i], wanted[i]]; }
  }
  runs.push(cur);

  var blocks = runs.map(function (run) {
    return sheet.getRange(top, run[0] + 1, height, run[1] - run[0] + 1).getValues();
  });

  return rows.map(function (rowNum) {
    var obj = { _row: rowNum };
    headerNames.forEach(function (h) {
      var col = map[h];
      if (col === undefined) return;
      for (var b = 0; b < runs.length; b++) {
        if (col >= runs[b][0] && col <= runs[b][1]) {
          obj[h] = blocks[b][rowNum - top][col - runs[b][0]];
          return;
        }
      }
    });
    return obj;
  });
}

/** The first matching row number, or 0. */
function dsx_findRowByKey_(sheet, headerName, key) {
  var rows = dsx_findRowsByKey_(sheet, headerName, key);
  return rows.length ? rows[0] : 0;
}

/** One row as an object keyed by header name. */
function dsx_readRow_(sheet, rowNumber) {
  var map = dsx_headerMap_(sheet);
  var lastCol = Math.max(1, sheet.getLastColumn());
  var vals = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  var obj = { _row: rowNumber };
  Object.keys(map).forEach(function (h) { obj[h] = vals[map[h]]; });
  return obj;
}

/** Writes ONLY the keys present in `obj`, each cast for a sheet cell. */
function dsx_writeRow_(sheet, rowNumber, obj) {
  var map = dsx_headerMap_(sheet);
  Object.keys(obj).forEach(function (h) {
    if (h.charAt(0) === '_') return;
    var col = map[h];
    if (col === undefined) return;
    sheet.getRange(rowNumber, col + 1).setValue(dsx_cast_(obj[h]));
  });
}

/** Appends in header order, casting every value. */
function dsx_appendRow_(sheet, obj) {
  var map = dsx_headerMap_(sheet);
  var width = Math.max(1, sheet.getLastColumn());
  var row = new Array(width);
  for (var i = 0; i < width; i++) row[i] = '';
  Object.keys(obj).forEach(function (h) {
    if (h.charAt(0) === '_') return;
    var col = map[h];
    if (col === undefined) return;
    row[col] = dsx_cast_(obj[h]);
  });
  sheet.appendRow(row);
  return sheet.getLastRow();
}

/**
 * Sheets stores dates, numbers, booleans and strings. Anything else — an
 * object slipped in by a caller — becomes JSON rather than "[object Object]".
 */
function dsx_cast_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v;
  var t = typeof v;
  if (t === 'number') return isFinite(v) ? v : '';
  if (t === 'boolean') return v;
  if (t === 'string') return v;
  try { return JSON.stringify(v); } catch (e) { return String(v); }
}

// ---------------------------------------------------------------------------
// SECTION C — payload packing
// ---------------------------------------------------------------------------

/**
 * A JSON string as up to 8 chunks of at most 45,000 characters.
 * @throws PAYLOAD_TOO_LARGE when it does not fit.
 */
function dsx_packPayload_(jsonString) {
  var s = String(jsonString === null || jsonString === undefined ? '' : jsonString);
  var limit = DSX_MAX_CHUNK * DSX_CHUNKS;
  if (s.length > limit) {
    throw new Error('PAYLOAD_TOO_LARGE: the summary is ' + s.length +
                    ' characters; the limit is ' + limit +
                    '. Shorten the hospital course or the investigations table.');
  }
  var out = [];
  for (var i = 0; i < DSX_CHUNKS; i++) {
    out.push(s.substr(i * DSX_MAX_CHUNK, DSX_MAX_CHUNK) || '');
  }
  return out;
}

/** The inverse. Accepts any row object carrying Payload_1..8. */
function dsx_unpackPayload_(rowObj) {
  var s = '';
  for (var i = 1; i <= DSX_CHUNKS; i++) {
    var part = rowObj['Payload_' + i];
    if (part === null || part === undefined) continue;
    s += String(part);
  }
  if (!s) return null;

  var declared = dsx_int_(rowObj.Payload_Chars);
  if (declared && declared !== s.length) {
    throw new Error('VALIDATION_FAILED: stored payload is ' + s.length +
                    ' characters but the row declares ' + declared +
                    '. The draft is corrupt; restore it from the last snapshot.');
  }
  try {
    return JSON.parse(s);
  } catch (e) {
    throw new Error('VALIDATION_FAILED: stored payload is not valid JSON (' + e.message + ').');
  }
}

/**
 * Forces the Payload_* cells of one row to plain text BEFORE anything is
 * written into them.
 *
 * setValue() parses what it is given the way the UI would: a chunk that
 * happens to begin with "=", "+", "-" or "@" becomes a formula, and one that
 * looks numeric becomes a number. Chunk 1 always starts with "{" and is safe,
 * but chunk boundaries fall mid-JSON, so any payload over 45,000 characters
 * can start a chunk on a character Sheets rewrites. The value then reads back
 * a different length, Payload_Chars no longer matches, and the draft is
 * declared corrupt — for a document that was stored perfectly well.
 */
function dsx_prepPayloadCells_(sheet, rowNumber) {
  var map = dsx_headerMap_(sheet);
  for (var i = 1; i <= DSX_CHUNKS; i++) {
    var col = map['Payload_' + i];
    if (col === undefined) continue;
    try { sheet.getRange(rowNumber, col + 1).setNumberFormat('@'); } catch (e) { /* format is advisory */ }
  }
}

/** Spreads a packed payload across the row fields a DS_* sheet expects. */
function dsx_payloadFields_(jsonString) {
  var chunks = dsx_packPayload_(jsonString);
  var o = { Payload_Chars: String(jsonString).length };
  for (var i = 0; i < DSX_CHUNKS; i++) o['Payload_' + (i + 1)] = chunks[i];
  return o;
}

// ---------------------------------------------------------------------------
// SECTION D — canonical JSON and hashing
// ---------------------------------------------------------------------------

/**
 * Deterministic JSON: object keys sorted recursively, "\n" line endings, no
 * incidental whitespace. Two runs over equal content MUST produce byte-equal
 * output or the hash chain is worthless.
 */
function dsx_canonicalJson_(obj) {
  return dsx_canon_(obj).replace(/\r\n/g, '\n');
}

function dsx_canon_(v) {
  if (v === null || v === undefined) return 'null';
  if (v instanceof Date) return JSON.stringify(isNaN(v.getTime()) ? null : v.toISOString());

  var t = typeof v;
  if (t === 'number') return isFinite(v) ? JSON.stringify(v) : 'null';
  if (t === 'boolean' || t === 'string') return JSON.stringify(v);

  if (Object.prototype.toString.call(v) === '[object Array]') {
    var items = [];
    for (var i = 0; i < v.length; i++) items.push(dsx_canon_(v[i]));
    return '[' + items.join(',') + ']';
  }

  if (t === 'object') {
    var keys = Object.keys(v).filter(function (k) { return v[k] !== undefined; });
    keys.sort();
    var pairs = keys.map(function (k) { return JSON.stringify(k) + ':' + dsx_canon_(v[k]); });
    return '{' + pairs.join(',') + '}';
  }
  return 'null';
}

function dsx_sha256Hex_(str) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, String(str), Utilities.Charset.UTF_8);
  return dsx_bytesToHex_(bytes);
}

function dsx_bytesToHex_(bytes) {
  var out = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] & 0xFF;
    out += (b < 16 ? '0' : '') + b.toString(16);
  }
  return out;
}

function dsx_bytesToB64Url_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

/**
 * The module's own HMAC secret. Discovery found no SecurityTokens.js and no
 * secret accessor anywhere in the project, so this is it. Created on first use
 * and never logged.
 */
function dsx_secret_() {
  var props = PropertiesService.getScriptProperties();
  var key = 'DS_VERIFY_SECRET';
  var v = props.getProperty(key);
  if (!v) {
    v = Utilities.base64EncodeWebSafe(
          Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
                                  Utilities.getUuid() + ':' + Date.now() + ':' + Utilities.getUuid(),
                                  Utilities.Charset.UTF_8));
    props.setProperty(key, v);
  }
  return v;
}

function dsx_hmacSha256B64Url_(message) {
  var sig = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256, String(message), dsx_secret_(), Utilities.Charset.UTF_8);
  return dsx_bytesToB64Url_(sig);
}

/** Length-independent, content-constant-time comparison. */
function dsx_timingSafeEqual_(a, b) {
  var x = String(a === null || a === undefined ? '' : a);
  var y = String(b === null || b === undefined ? '' : b);
  // Compare digests so the loop length never depends on the secret's length.
  var hx = dsx_sha256Hex_(x), hy = dsx_sha256Hex_(y);
  var diff = hx.length ^ hy.length;
  for (var i = 0; i < hx.length && i < hy.length; i++) {
    diff |= hx.charCodeAt(i) ^ hy.charCodeAt(i);
  }
  return diff === 0;
}

/** The 12-character form printed on the document and shown in the UI. */
function dsx_shortHash_(hash) { return dsx_str_(hash).substring(0, 12); }

// ---------------------------------------------------------------------------
// SECTION E — configuration
// ---------------------------------------------------------------------------

var DSX_CONFIG_DEFAULTS = {
  // Who may prepare. Confirmed with the CEO: nurse and duty doctor both.
  DS_PREPARER_ROLES: 'nurse,doctor',
  // Solo-clinic fast path: OFF by default, per-tenant switch.
  DS_REQUIRE_PREPARER_REVIEW: 'true',
  // Consultant of record, or any doctor who logs an on-behalf reason.
  DS_SIGNER_POLICY: 'CONSULTANT_OR_ANY_DOCTOR_WITH_REASON',
  // Billing gate at settlement: OFF | WARN | BLOCK. Confirmed: WARN.
  DS_BILLING_GATE: 'WARN',
  DS_TAT_AMBER_MIN: '120',
  DS_TAT_RED_MIN: '240',
  DS_AI_ENABLED: 'false',
  DS_AI_REQUIRE_CONSENT: 'true',
  DS_PUSH_DISCHARGE_RX: 'false',
  DS_BOOK_FOLLOWUP: 'false',
  // Signing credential. Passwords are clear text today (Discovery §5), so the
  // password fallback is a deliberate, switchable compromise.
  DS_SIGN_ALLOW_PASSWORD: 'true',
  // Facility identity for the printed document, confirmed by the CEO. The
  // phone already prints from IPP_CLINIC, so it is not repeated here. Both are
  // overridable per tenant through Script Properties.
  DS_FACILITY_ADDRESS: 'Crescentia Healthtech, Tamil Nadu, India',
  // No facility registration or licence number has been supplied yet. Empty
  // prints nothing rather than a placeholder.
  DS_FACILITY_REG_LINE: ''
};

/**
 * Config from Script Properties — the project's only config mechanism
 * (there is no config sheet). Unset keys take the defaults above.
 */
function dsx_config_() {
  var props = {};
  try { props = PropertiesService.getScriptProperties().getProperties() || {}; } catch (e) {}

  var cfg = {};
  Object.keys(DSX_CONFIG_DEFAULTS).forEach(function (k) {
    var raw = dsx_str_(props[k]);
    cfg[k] = raw === '' ? DSX_CONFIG_DEFAULTS[k] : raw;
  });

  return {
    raw: cfg,
    preparerRoles: cfg.DS_PREPARER_ROLES.toLowerCase().split(',')
                     .map(function (s) { return s.trim(); }).filter(String),
    requirePreparerReview: dsx_bool_(cfg.DS_REQUIRE_PREPARER_REVIEW),
    signerPolicy: dsx_upper_(cfg.DS_SIGNER_POLICY),
    billingGate: dsx_upper_(cfg.DS_BILLING_GATE) || 'WARN',
    tatAmberMin: dsx_int_(cfg.DS_TAT_AMBER_MIN) || 120,
    tatRedMin: dsx_int_(cfg.DS_TAT_RED_MIN) || 240,
    aiEnabled: dsx_bool_(cfg.DS_AI_ENABLED),
    aiRequireConsent: dsx_bool_(cfg.DS_AI_REQUIRE_CONSENT),
    pushDischargeRx: dsx_bool_(cfg.DS_PUSH_DISCHARGE_RX),
    bookFollowUp: dsx_bool_(cfg.DS_BOOK_FOLLOWUP),
    signAllowPassword: dsx_bool_(cfg.DS_SIGN_ALLOW_PASSWORD),
    facilityAddress: cfg.DS_FACILITY_ADDRESS,
    facilityRegLine: cfg.DS_FACILITY_REG_LINE
  };
}

// ---------------------------------------------------------------------------
// SECTION F — the payload schema
// ---------------------------------------------------------------------------

/**
 * The one shape that travels between assembly, the editor, the snapshots and
 * the renderer.
 *
 * {
 *   schemaVersion: 1,
 *   dischargeType: "NORMAL" | "LAMA" | "DAMA" | "REFERRED" | "DEATH" | "ABSCONDED",
 *   meta: {
 *     ipNumber:         string,
 *     patientId:        string,
 *     generatedAt:      ISO-8601 string,
 *     sourceBundleHash: sha256 hex of the canonical source bundle
 *   },
 *   sections: {
 *     <SECTION_KEY>: {
 *       title:    string,
 *       format:   "TEXT" | "TABLE" | "LIST" | "FIELDS",
 *       content:  string            when format is TEXT
 *                 Array<string>     when format is LIST
 *                 {columns:[…], rows:[[…]]}  when format is TABLE
 *                 {fieldKey: value} when format is FIELDS
 *       origin:   "AUTO" | "AI" | "MANUAL",
 *       edited:   boolean,
 *       editedBy: username or "",
 *       editedAt: ISO-8601 or "",
 *       sourceRefs: [{type:"NOTE"|"LAB"|"CASESHEET"|"ADMISSION"|"PATIENT"|"QUEUE",
 *                     id:string, at:ISO-8601 or ""}],
 *       sourceChangedSinceEdit: boolean,
 *       reviewed:    boolean,
 *       aiAcceptedBy: username or "",
 *       aiAcceptedAt: ISO-8601 or ""
 *     }
 *   }
 * }
 *
 * `content` is the ONLY field a human edit may change, plus `reviewed` and the
 * AI acceptance fields. Everything else is engine-owned.
 */
function dsx_newSection_(title, format, content, origin, sourceRefs) {
  return {
    title: dsx_str_(title),
    format: dsx_upper_(format) || 'TEXT',
    content: (content === undefined || content === null) ? dsx_emptyContent_(format) : content,
    origin: dsx_upper_(origin) || 'AUTO',
    edited: false,
    editedBy: '',
    editedAt: '',
    sourceRefs: sourceRefs || [],
    sourceChangedSinceEdit: false,
    reviewed: false,
    aiAcceptedBy: '',
    aiAcceptedAt: ''
  };
}

function dsx_emptyContent_(format) {
  switch (dsx_upper_(format)) {
    case 'LIST':   return [];
    case 'TABLE':  return { columns: [], rows: [] };
    case 'FIELDS': return {};
    default:       return '';
  }
}

/** True when a section carries nothing a reader would see. */
function dsx_sectionIsEmpty_(sec) {
  if (!sec) return true;
  var c = sec.content;
  switch (dsx_upper_(sec.format)) {
    case 'LIST':   return !c || !c.length;
    case 'TABLE':  return !c || !c.rows || !c.rows.length;
    case 'FIELDS':
      if (!c) return true;
      return Object.keys(c).every(function (k) { return dsx_str_(c[k]) === ''; });
    default:       return dsx_str_(c) === '';
  }
}

/**
 * Structural validation. Clinical completeness is the readiness engine's job
 * (DS_Assembly.gs); this only rejects a payload that cannot be stored,
 * rendered or hashed.
 * @return {{ok:boolean, errors:Array<string>}}
 */
function dsx_validatePayload_(payload) {
  var errors = [];

  if (!payload || typeof payload !== 'object') {
    return { ok: false, errors: ['Payload is not an object.'] };
  }
  if (dsx_int_(payload.schemaVersion) !== DSX_SCHEMA_VERSION) {
    errors.push('Unsupported schemaVersion: ' + payload.schemaVersion +
                ' (this build writes ' + DSX_SCHEMA_VERSION + ').');
  }
  if (DSX_DISCHARGE_TYPES.indexOf(dsx_upper_(payload.dischargeType)) === -1) {
    errors.push('Unknown dischargeType: ' + payload.dischargeType + '.');
  }

  var meta = payload.meta || {};
  if (!dsx_str_(meta.ipNumber))  errors.push('meta.ipNumber is missing.');
  if (!dsx_str_(meta.patientId)) errors.push('meta.patientId is missing.');

  var sections = payload.sections;
  if (!sections || typeof sections !== 'object') {
    errors.push('sections is missing.');
    return { ok: false, errors: errors };
  }

  var allowedFormats = ['TEXT', 'TABLE', 'LIST', 'FIELDS'];
  var allowedOrigins = ['AUTO', 'AI', 'MANUAL'];

  Object.keys(sections).forEach(function (key) {
    var s = sections[key];
    if (!s || typeof s !== 'object') { errors.push(key + ': not an object.'); return; }
    if (allowedFormats.indexOf(dsx_upper_(s.format)) === -1) {
      errors.push(key + ': unknown format "' + s.format + '".');
    }
    if (allowedOrigins.indexOf(dsx_upper_(s.origin)) === -1) {
      errors.push(key + ': unknown origin "' + s.origin + '".');
    }
    switch (dsx_upper_(s.format)) {
      case 'LIST':
        if (Object.prototype.toString.call(s.content) !== '[object Array]') {
          errors.push(key + ': LIST content must be an array.');
        }
        break;
      case 'TABLE':
        if (!s.content || Object.prototype.toString.call(s.content.rows) !== '[object Array]' ||
            Object.prototype.toString.call(s.content.columns) !== '[object Array]') {
          errors.push(key + ': TABLE content must be {columns:[], rows:[]}.');
        }
        break;
      case 'FIELDS':
        if (!s.content || typeof s.content !== 'object' ||
            Object.prototype.toString.call(s.content) === '[object Array]') {
          errors.push(key + ': FIELDS content must be an object.');
        }
        break;
      default:
        if (typeof s.content !== 'string') errors.push(key + ': TEXT content must be a string.');
    }
    if (Object.prototype.toString.call(s.sourceRefs) !== '[object Array]') {
      errors.push(key + ': sourceRefs must be an array.');
    }
  });

  // The stored form must fit. Catching it here beats discovering it at save.
  try {
    dsx_packPayload_(JSON.stringify(payload));
  } catch (e) {
    errors.push(e.message);
  }

  return { ok: errors.length === 0, errors: errors };
}

// ---------------------------------------------------------------------------
// SECTION G — DS_* row access
// ---------------------------------------------------------------------------

function dsx_summariesSheet_() { return dsx_requireSheet_(DSX_SHEETS.SUMMARIES); }
function dsx_workingSheet_()   { return dsx_requireSheet_(DSX_SHEETS.WORKING); }
function dsx_snapshotsSheet_() { return dsx_requireSheet_(DSX_SHEETS.SNAPSHOTS); }
function dsx_logSheet_()       { return dsx_requireSheet_(DSX_SHEETS.LOG); }
function dsx_phrasesSheet_()   { return dsx_requireSheet_(DSX_SHEETS.PHRASES); }

/** The DS_Summaries header row, or null. Also returns its row number. */
function dsx_getHeader_(summaryId) {
  var sh = dsx_summariesSheet_();
  var row = dsx_findRowByKey_(sh, 'Summary_ID', dsx_upper_(summaryId));
  if (!row) return null;
  var obj = dsx_readRow_(sh, row);
  if (dsx_str_(obj.Tenant_ID) && dsx_upper_(obj.Tenant_ID) !== dsx_upper_(dsx_tenant_())) return null;
  return obj;
}

/** The live draft payload, or null when the working row is absent. */
function dsx_getWorking_(summaryId) {
  var sh = dsx_workingSheet_();
  var row = dsx_findRowByKey_(sh, 'Summary_ID', dsx_upper_(summaryId));
  if (!row) return null;
  var obj = dsx_readRow_(sh, row);
  return { row: row, baseSnapshotNo: dsx_int_(obj.Base_Snapshot_No), payload: dsx_unpackPayload_(obj) };
}

/** Creates or replaces the single working row for a summary. Caller holds the lock. */
function dsx_putWorking_(summaryId, payload, baseSnapshotNo, username) {
  var sh = dsx_workingSheet_();
  var json = JSON.stringify(payload);
  var fields = dsx_payloadFields_(json);

  fields.Summary_ID = dsx_upper_(summaryId);
  fields.Tenant_ID = dsx_tenant_();
  fields.Base_Snapshot_No = dsx_int_(baseSnapshotNo);
  fields.Updated_At = new Date();
  fields.Updated_By = dsx_str_(username);

  var row = dsx_findRowByKey_(sh, 'Summary_ID', dsx_upper_(summaryId));
  if (!row) {
    row = dsx_appendRow_(sh, { Summary_ID: dsx_upper_(summaryId), Tenant_ID: dsx_tenant_() });
  }
  dsx_prepPayloadCells_(sh, row);
  dsx_writeRow_(sh, row, fields);
  return row;
}

/**
 * Appends a frozen snapshot. Append-only: this never updates an existing row.
 * Caller holds the lock.
 */
function dsx_appendSnapshot_(summaryId, snapshotNo, type, payload, contentHash, prevSignedHash, username, verifyTokenHash) {
  var sh = dsx_snapshotsSheet_();
  var json = JSON.stringify(payload);
  var fields = dsx_payloadFields_(json);

  fields.Snapshot_Key = dsx_upper_(summaryId) + ':' + dsx_int_(snapshotNo);
  fields.Tenant_ID = dsx_tenant_();
  fields.Summary_ID = dsx_upper_(summaryId);
  fields.Snapshot_No = dsx_int_(snapshotNo);
  fields.Snapshot_Type = dsx_upper_(type);
  fields.Content_Hash = dsx_str_(contentHash);
  fields.Prev_Signed_Hash = dsx_str_(prevSignedHash);
  fields.Created_At = new Date();
  fields.Created_By = dsx_str_(username);
  fields.Verify_Token_Hash = dsx_str_(verifyTokenHash);

  var row = dsx_appendRow_(sh, { Snapshot_Key: fields.Snapshot_Key, Tenant_ID: fields.Tenant_ID });
  dsx_prepPayloadCells_(sh, row);
  dsx_writeRow_(sh, row, fields);
  return row;
}

/** Every snapshot row for a summary, payloads excluded, newest first. */
function dsx_listSnapshots_(summaryId) {
  var sh = dsx_snapshotsSheet_();
  var rows = dsx_findRowsByKey_(sh, 'Summary_ID', dsx_upper_(summaryId));
  if (!rows.length) return [];

  var out = dsx_readColumns_(sh,
    ['Snapshot_No', 'Snapshot_Type', 'Content_Hash', 'Prev_Signed_Hash',
     'Created_At', 'Created_By'], rows
  ).map(function (v) {
    return {
      row: v._row,
      snapshotNo: dsx_int_(v['Snapshot_No']),
      type: dsx_upper_(v['Snapshot_Type']),
      contentHash: dsx_str_(v['Content_Hash']),
      prevSignedHash: dsx_str_(v['Prev_Signed_Hash']),
      createdAt: dsx_toDate_(v['Created_At']),
      createdBy: dsx_str_(v['Created_By'])
    };
  });
  out.sort(function (a, b) { return b.snapshotNo - a.snapshotNo; });
  return out;
}

/** One snapshot's payload, by number. */
function dsx_getSnapshotPayload_(summaryId, snapshotNo) {
  var sh = dsx_snapshotsSheet_();
  var key = dsx_upper_(summaryId) + ':' + dsx_int_(snapshotNo);
  var row = dsx_findRowByKey_(sh, 'Snapshot_Key', key);
  if (!row) return null;
  var obj = dsx_readRow_(sh, row);
  return { meta: obj, payload: dsx_unpackPayload_(obj) };
}

/** The most recent snapshot of a given type, or null. */
function dsx_latestSnapshotOfType_(summaryId, type) {
  var list = dsx_listSnapshots_(summaryId).filter(function (s) {
    return s.type === dsx_upper_(type);
  });
  return list.length ? list[0] : null;
}

/** Append-only workflow event. Never throws upward — logging must not lose a commit. */
function dsx_logEvent_(summaryId, actor, action, fromStatus, toStatus, snapshotNo, contentHash, comment, meta) {
  try {
    dsx_appendRow_(dsx_logSheet_(), {
      Event_ID: dsx_newEventId_(),
      Tenant_ID: dsx_tenant_(),
      Summary_ID: dsx_upper_(summaryId),
      Timestamp: new Date(),
      Actor_Username: dsx_str_(actor && actor.username),
      Actor_Role: dsx_str_(actor && actor.role),
      Action: dsx_upper_(action),
      From_Status: dsx_upper_(fromStatus),
      To_Status: dsx_upper_(toStatus),
      Snapshot_No: dsx_int_(snapshotNo),
      Content_Hash: dsx_str_(contentHash),
      Comment: dsx_str_(comment),
      Meta_JSON: meta ? JSON.stringify(meta) : ''
    });
  } catch (e) {
    Logger.log('DS workflow log failed: ' + e.message);
  }
}

/** The latest `limit` workflow events for a summary, newest first. */
function dsx_recentEvents_(summaryId, limit) {
  var sh = dsx_logSheet_();
  var rows = dsx_findRowsByKey_(sh, 'Summary_ID', dsx_upper_(summaryId));
  rows.reverse();
  if (limit > 0) rows = rows.slice(0, limit);
  if (!rows.length) return [];

  // dsx_readColumns_ hands rows back in ascending row order; an append-only
  // log means descending row order is newest first, which is what callers want.
  var raw = dsx_readColumns_(sh,
    ['Event_ID', 'Timestamp', 'Actor_Username', 'Actor_Role', 'Action',
     'From_Status', 'To_Status', 'Snapshot_No', 'Comment', 'Meta_JSON'], rows);
  raw.sort(function (a, b) { return b._row - a._row; });

  return raw.map(function (v) {
    var meta = {};
    try { meta = JSON.parse(dsx_str_(v['Meta_JSON']) || '{}'); } catch (e) {}
    return {
      eventId: dsx_str_(v['Event_ID']),
      at: dsx_toDate_(v['Timestamp']),
      atText: dsx_fmt_(v['Timestamp'], 'dd-MMM-yyyy hh:mm a'),
      actor: dsx_str_(v['Actor_Username']),
      role: dsx_str_(v['Actor_Role']),
      action: dsx_str_(v['Action']),
      fromStatus: dsx_str_(v['From_Status']),
      toStatus: dsx_str_(v['To_Status']),
      snapshotNo: dsx_int_(v['Snapshot_No']),
      comment: dsx_str_(v['Comment']),
      meta: meta
    };
  });
}

/** Mirrors a mutation into the project's clinical audit sheet. Best effort. */
function dsx_audit_(actor, actionType, summaryId, details) {
  try {
    if (typeof logAudit_ === 'function') {
      logAudit_({ username: actor && actor.username, role: actor && actor.role,
                  doctorId: actor && actor.doctorId },
                dsx_upper_(actionType), 'DISCHARGE_SUMMARY', dsx_upper_(summaryId), details || {});
    }
  } catch (e) {
    Logger.log('DS audit failed: ' + e.message);
  }
}

// ---------------------------------------------------------------------------
// SECTION H — the standard envelope
// ---------------------------------------------------------------------------

function dsx_ok_(message, data) {
  return { success: true, message: dsx_str_(message), data: data === undefined ? null : data };
}

// ---------------------------------------------------------------------------
// SECTION H2 — the transport boundary
// ---------------------------------------------------------------------------
//
// WHY THIS EXISTS
//   google.script.run serialises whatever a server function returns. When the
//   graph contains something it cannot carry — an Invalid Date, a NaN or an
//   Infinity from a division, a function, a cycle — the client's SUCCESS
//   handler is called with null. Not an error, not a message: null. That is
//   exactly the "The server returned nothing for DS-IP...." the discharge
//   editor reports, and it is indistinguishable at the browser from a dropped
//   connection, which is why it has been so hard to pin down.
//
//   Every read endpoint therefore hands its envelope through dsx_wire_ before
//   returning it. Dates survive as Dates (Apps Script carries those natively)
//   so no client contract changes; everything the transport cannot represent
//   is turned into something it can.

var DSX_WIRE_MAX_DEPTH = 32;

/** A deep copy of `v` containing only values google.script.run can carry. */
function dsx_wire_(v) { return dsx_wireValue_(v, 0, []); }

function dsx_wireValue_(v, depth, stack) {
  if (v === null || v === undefined) return null;

  // A Date is carried natively — but an Invalid Date is not, and one
  // unparsable cell anywhere in the graph would sink the whole reply.
  // Tested by tag rather than instanceof so a Date built in another context
  // (or a subclass) is still recognised as one.
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? null : v;
  }

  var t = typeof v;
  if (t === 'string' || t === 'boolean') return v;
  if (t === 'number') return isFinite(v) ? v : null;   // NaN / ±Infinity
  if (t === 'function' || t === 'undefined') return null;
  if (t !== 'object') return String(v);                // symbol and friends

  if (depth >= DSX_WIRE_MAX_DEPTH) return null;
  if (stack.indexOf(v) !== -1) return null;            // cycle
  stack.push(v);

  var out;
  if (Object.prototype.toString.call(v) === '[object Array]') {
    out = [];
    for (var i = 0; i < v.length; i++) out.push(dsx_wireValue_(v[i], depth + 1, stack));
  } else {
    out = {};
    var keys = Object.keys(v);
    for (var k = 0; k < keys.length; k++) {
      if (!keys[k]) continue;                          // an unaddressable key
      out[keys[k]] = dsx_wireValue_(v[keys[k]], depth + 1, stack);
    }
  }
  stack.pop();
  return out;
}

/**
 * The JSON form of a value, as a value.
 *
 * The store holds every payload as JSON, so anything rebuilt in memory is
 * put through the same round trip before it is used as a payload — otherwise
 * a freshly assembled draft carries Date objects and non-finite numbers that
 * a stored one never has, and the two disagree about what the document says.
 */
function dsx_jsonNormalize_(v) {
  try { return JSON.parse(JSON.stringify(v)); } catch (e) { return v; }
}

/** Serialised length of a reply, or -1 when it cannot be measured. */
function dsx_wireSize_(v) {
  try { return JSON.stringify(v).length; } catch (e) { return -1; }
}

function dsx_err_(code, message, data) {
  var out = { success: false, code: dsx_upper_(code), message: dsx_str_(message) };
  if (data !== undefined) out.data = data;
  return out;
}

/**
 * Turns a thrown error into the envelope, recovering the code when the message
 * was thrown as "CODE: text" by a helper in this module.
 */
function dsx_fromError_(e) {
  var msg = (e && e.message) ? String(e.message) : String(e);
  var m = msg.match(/^([A-Z_]{4,}):\s*(.*)$/);
  if (m) return dsx_err_(m[1], m[2]);
  return dsx_err_('VALIDATION_FAILED', msg);
}