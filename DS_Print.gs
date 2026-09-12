// ============================================================================
// DS_Print.gs  —  Crescentia HealthTech / CresRx
// IP Discharge Summary Engine · Phase 7 · one renderer, stored PDF, QR verify
// ----------------------------------------------------------------------------
// ONE RENDERER, TWO OUTPUTS
//   dsx_renderDocumentHtml_() produces the HTML for browser printing AND for
//   the stored PDF. They cannot drift, because they are the same string.
//
// IT COMPOSES INTO IP_Print_Kit.gs
//   ipp_doc_() already emits the A4 shell this document needs: one outer
//   <table class="page"> with a repeating <thead> patient strip and a pinned
//   <tfoot>, table-layout:fixed everywhere, no flexbox, @page A4 with 12-14mm
//   margins, print-color-adjust:exact. That file documents, with reasons, the
//   CSS that survives Apps Script's HTML-to-PDF converter. Writing a second
//   stylesheet here would mean re-learning all of it.
//
// THE SPIKE IS NOT OPTIONAL
//   Three things are unproven in this project: a data-URI image inside the
//   PDF converter, Tamil glyphs, and page breaks inside a long table. Run
//   ds_printSpike() on the deployment and read its report BEFORE relying on
//   the stored PDF. The renderer degrades honestly: if the QR image does not
//   survive, the printed verification CODE and URL still do, because they are
//   plain text.
// ============================================================================

var DSX_PDF_FOLDER_ROOT = 'CresRx';
var DSX_PDF_FOLDER_SUB  = 'Discharge_Summaries';

// ---------------------------------------------------------------------------
// SECTION A — FACILITY IDENTITY
// ---------------------------------------------------------------------------

/**
 * Name, phone, address and registration line for the letterhead.
 *
 * Discovery §6: the project has no stored facility address and no registration
 * number anywhere. IPP_CLINIC is the only identity that exists. Script
 * Properties override it, so the clinic can supply the missing lines without
 * a code change, and nothing looks different until they do.
 */
function dsx_facility_() {
  var cfg = dsx_config_();
  var base = (typeof IPP_CLINIC === 'object' && IPP_CLINIC) ? IPP_CLINIC : {};
  var props = {};
  try { props = PropertiesService.getScriptProperties().getProperties() || {}; } catch (e) {}

  return {
    name: dsx_str_(props.CLINIC_NAME) || dsx_str_(base.name) || 'Clinic',
    tagline: dsx_str_(base.tagline),
    phone: dsx_str_(base.phone),
    address: cfg.facilityAddress,
    regLine: cfg.facilityRegLine
  };
}

// ---------------------------------------------------------------------------
// SECTION B — THE RENDERER
// ---------------------------------------------------------------------------

/** Sections in the order a clinician reads them on paper. */
var DSX_PRINT_ORDER = [
  'ADMISSION_DETAILS', 'DIAGNOSIS', 'PRESENTING_COMPLAINTS', 'HISTORY',
  'ALLERGIES', 'EXAM_ON_ADMISSION', 'INVESTIGATIONS', 'PENDING_RESULTS',
  'PROCEDURES', 'HOSPITAL_COURSE', 'TREATMENT_GIVEN', 'CONDITION_AT_DISCHARGE',
  'DISCHARGE_MEDICATIONS', 'ADVICE', 'FOLLOW_UP',
  'RED_FLAGS', 'LAMA_DETAILS', 'REFERRAL_DETAILS', 'DEATH_DETAILS', 'ABSCOND_DETAILS'
];

/**
 * @param {string} summaryId
 * @param {string} snapshotRef  "WORKING" | "SIGNED" | "SIGNED:n"
 * @param {string} mode         DRAFT | FINAL | REPRINT
 * @param {Object} [opts]       {printedBy}
 * @return {string} a complete HTML document
 */
function dsx_renderDocumentHtml_(summaryId, snapshotRef, mode, opts) {
  opts = opts || {};
  var M = dsx_upper_(mode) || 'FINAL';

  var header = dsx_getHeader_(summaryId);
  if (!header) throw new Error('VALIDATION_FAILED: no discharge summary found for ' + summaryId + '.');

  var ref = dsx_resolveRef_(summaryId, snapshotRef);
  if (!ref.payload) throw new Error('VALIDATION_FAILED: version "' + snapshotRef + '" could not be read.');
  var payload = ref.payload;

  var snapshotNo = (function () {
    var m = String(ref.label).split(':');
    return m.length > 1 ? dsx_int_(m[1]) : dsx_int_(header.Current_Snapshot_No);
  })();

  var fac = dsx_facility_();
  var banner = (payload.sections.PATIENT_BANNER && payload.sections.PATIENT_BANNER.content) || {};
  var body = [];

  // ---- letterhead extras the shared kit does not carry --------------------
  if (fac.address || fac.regLine) {
    body.push('<div style="text-align:center;font-size:9pt;color:#334155;margin:-8px 0 10px;">' +
      (fac.address ? ipp_esc_(fac.address) : '') +
      (fac.address && fac.regLine ? ' &middot; ' : '') +
      (fac.regLine ? ipp_esc_(fac.regLine) : '') + '</div>');
  }

  // ---- mode banner ---------------------------------------------------------
  if (M === 'DRAFT') {
    body.push('<div style="border:1.5pt dashed #b91c1c;color:#b91c1c;text-align:center;' +
      'font-weight:700;letter-spacing:.12em;padding:6px;margin-bottom:10px;' +
      'font-size:12pt;-webkit-print-color-adjust:exact;print-color-adjust:exact;">' +
      'DRAFT &ndash; NOT VALID FOR DISCHARGE</div>');
  } else if (M === 'REPRINT') {
    body.push('<div style="border:.8pt solid #64748b;color:#334155;text-align:center;' +
      'font-size:9pt;padding:4px;margin-bottom:10px;">Duplicate copy &ndash; printed ' +
      ipp_esc_(ipp_when_(new Date())) + ' by ' + ipp_esc_(opts.printedBy || '') + '</div>');
  }

  // ---- superseded notice ---------------------------------------------------
  var lastSigned = dsx_int_(header.Last_Signed_Snapshot_No);
  if (snapshotNo && lastSigned && snapshotNo < lastSigned) {
    var amendReason = dsx_amendmentReasonFor_(summaryId, lastSigned);
    body.push('<div style="border:1pt solid #b45309;background:#fef3c7;color:#78350f;' +
      'padding:6px 8px;margin-bottom:10px;font-size:9.5pt;' +
      '-webkit-print-color-adjust:exact;print-color-adjust:exact;">' +
      '<strong>Superseded</strong> by version ' + lastSigned + ' on ' +
      ipp_esc_(dsx_fmt_(header.Signed_At, 'dd-MMM-yyyy')) +
      (amendReason ? ': ' + ipp_esc_(amendReason) : '') + '</div>');
  }

  // ---- sections ------------------------------------------------------------
  DSX_PRINT_ORDER.forEach(function (key) {
    var sec = payload.sections[key];
    if (!sec) return;

    // FINAL omits empty sections rather than printing "N/A" filler. Allergies
    // is the exception: a discharge summary that is silent about allergies is
    // a dangerous document, so it always prints its stated status.
    var empty = dsx_sectionIsEmpty_(sec);
    if (empty && M === 'FINAL' && key !== 'ALLERGIES') return;

    body.push(dsx_printSection_(key, sec, empty));
  });

  // ---- signature block -----------------------------------------------------
  body.push(dsx_printSignatureBlock_(header, payload, M, snapshotNo));

  // ---- patient acknowledgement --------------------------------------------
  body.push(ipp_sec_('Patient / attendant acknowledgement',
    '<div style="font-size:9.5pt;margin-bottom:14px;">' +
    'Received this discharge summary; medicines and warning signs were explained to me ' +
    'in a language I understand.</div>' +
    ipp_cols_(
      ipp_kv_([['Name', '&nbsp;'], ['Relation', '&nbsp;']], { narrow: true, keepEmpty: true }),
      ipp_sig_('', 'Signature / thumb impression'))));

  var wb = [dsx_str_(banner.ward), dsx_str_(banner.bed)].filter(String).join(' / ');
  var footNote = 'Summary ' + dsx_str_(header.Summary_ID) + ' v' + snapshotNo +
                 (dsx_str_(header.Signed_Hash) ? ' · ' + dsx_shortHash_(header.Signed_Hash) : '');

  return ipp_doc_({
    docTitle: dsx_str_(payload.documentTitle) || dsx_documentTitle_(header.Discharge_Type),
    patient: {
      name: dsx_str_(banner.name),
      pid: dsx_str_(banner.patientId),
      ipNumber: dsx_str_(banner.ipNumber) || dsx_str_(header.IP_Number),
      ageSex: dsx_str_(banner.ageSex),
      wardBed: wb,
      consultant: dsx_consultantNameFor_(header),
      diagnosis: dsx_finalDiagnosisOf_(payload)
    },
    bodyHtml: body.join(''),
    footNote: footNote
  });
}

/** One section, rendered by its format through the shared print kit. */
function dsx_printSection_(key, sec, empty) {
  var title = dsx_str_(sec.title) || key;

  if (key === 'ALLERGIES') {
    // High-visibility box, always printed, whatever it says.
    var text = empty ? 'NOT RECORDED' : dsx_str_(sec.content);
    return '<div style="border:1.2pt solid #b91c1c;background:#fef2f2;color:#7f1d1d;' +
      'padding:7px 9px;margin-bottom:12px;page-break-inside:avoid;break-inside:avoid;' +
      '-webkit-print-color-adjust:exact;print-color-adjust:exact;">' +
      '<strong style="letter-spacing:.05em;">ALLERGIES:</strong> ' +
      ipp_escMultiline_(text) + '</div>';
  }

  if (empty) return ipp_sec_(title, '<div style="color:#64748b;font-size:9.5pt;">&mdash;</div>');

  switch (dsx_upper_(sec.format)) {
    case 'LIST':
      return ipp_sec_(title,
        '<ul style="margin:0 0 4px 16px;padding:0;font-size:10pt;">' +
        sec.content.map(function (i) {
          return '<li style="margin-bottom:2px;">' + dsx_bilingual_(i) + '</li>';
        }).join('') + '</ul>');

    case 'TABLE':
      // ipp_table_ escapes its own headers; escaping here would double it.
      return ipp_sec_(title, ipp_table_(
        sec.content.columns.slice(),
        sec.content.rows.map(function (r) {
          return sec.content.columns.map(function (c, i) {
            var v = r[i] === undefined ? '' : r[i];
            // Generic names print in capitals: that is what a pharmacist reads.
            if (/^generic$/i.test(c)) return '<strong>' + ipp_esc_(String(v).toUpperCase()) + '</strong>';
            return ipp_esc_(v);
          });
        })));

    case 'FIELDS':
      var pairs = Object.keys(sec.content)
        .filter(function (f) { return dsx_str_(sec.content[f]); })
        .map(function (f) {
          return [dsx_humanise_(f), ipp_escMultiline_(sec.content[f])];
        });
      return ipp_sec_(title, ipp_kv_(pairs));

    default:
      return ipp_sec_(title,
        '<div style="font-size:10pt;">' + ipp_escMultiline_(sec.content) + '</div>');
  }
}

/**
 * A phrase-library line prints English with Tamil beneath it when the clinic
 * has authored a Tamil version. Nothing is machine-translated: the Tamil text
 * is only ever what a human at the clinic wrote and reviewed.
 */
function dsx_bilingual_(text) {
  var en = dsx_str_(text);
  var ta = dsx_phraseTamil_(en);
  if (!ta) return ipp_esc_(en);
  return ipp_esc_(en) +
         '<div style="font-size:9.5pt;color:#334155;">' + ipp_esc_(ta) + '</div>';
}

var DSX_PHRASE_CACHE = null;

function dsx_phraseTamil_(englishText) {
  if (DSX_PHRASE_CACHE === null) {
    DSX_PHRASE_CACHE = {};
    try {
      var sh = dsx_sheet_(DSX_SHEETS.PHRASES);
      if (sh && sh.getLastRow() >= 2) {
        var map = dsx_headerMap_(sh);
        var n = sh.getLastRow() - 1;
        var width = Math.max(1, sh.getLastColumn());
        var vals = sh.getRange(2, 1, n, width).getValues();
        for (var i = 0; i < n; i++) {
          if (!dsx_bool_(vals[i][map['Active']])) continue;
          var en = dsx_str_(vals[i][map['Text_EN']]);
          var ta = dsx_str_(vals[i][map['Text_TA']]);
          if (en && ta) DSX_PHRASE_CACHE[en.toUpperCase()] = ta;
        }
      }
    } catch (e) { /* the document prints in English if the library is unreadable */ }
  }
  return DSX_PHRASE_CACHE[dsx_upper_(englishText)] || '';
}

function dsx_printSignatureBlock_(header, payload, mode, snapshotNo) {
  var sig = payload.signature || {};
  var signed = dsx_upper_(header.Status) === 'SIGNED' && dsx_str_(sig.signerName);

  var preparedBy = dsx_str_(sig.preparedBy) || dsx_str_(header.Prepared_By);
  var left = ipp_kv_([
    ['Prepared by', ipp_esc_(preparedBy || '—')],
    ['Submitted', ipp_esc_(dsx_fmt_(header.Submitted_At, 'dd-MMM-yyyy hh:mm a') || '—')]
  ], { narrow: true, keepEmpty: true });

  var right;
  if (signed) {
    right = ipp_kv_([
      ['Verified &amp; signed by', '<strong>' + ipp_esc_(sig.signerName) + '</strong>'],
      ['Qualification', ipp_esc_(sig.signerQualification)],
      ['Reg. No', ipp_esc_(sig.signerRegNo)],
      ['Signed at', ipp_esc_(dsx_fmt_(header.Signed_At, 'dd-MMM-yyyy hh:mm a'))]
    ], { narrow: true, keepEmpty: true }) +
    (dsx_str_(sig.onBehalfReason)
      ? '<div style="font-size:9pt;color:#7f1d1d;margin-top:3px;">Signed on behalf of the ' +
        'consultant of record: ' + ipp_esc_(sig.onBehalfReason) + '</div>'
      : '');
  } else {
    right = '<div style="color:#b91c1c;font-size:10pt;font-weight:700;">Not yet signed</div>';
  }

  var verify = signed ? dsx_verifyBlock_(header, snapshotNo) : '';

  return ipp_sec_('Signatures',
    ipp_cols_(left, right) +
    '<div style="height:8px"></div>' +
    ipp_cols_(
      ipp_sig_(signed ? dsx_str_(sig.signerName) : '', 'Signature and seal'),
      verify),
    { keepEmpty: true });
}

/**
 * The verification block. The CODE and the URL are plain text and always
 * print; the QR is an image and is therefore the part that depends on the
 * spike. Printing only a QR would be a document nobody can verify from a
 * photocopy.
 */
function dsx_verifyBlock_(header, snapshotNo) {
  var summaryId = dsx_str_(header.Summary_ID);
  var token = dsx_hmacSha256B64Url_(dsx_upper_(summaryId) + ':' + dsx_int_(snapshotNo));
  var url = dsx_webAppUrl_() + '?verifyDS=' + encodeURIComponent(token);
  var code = dsx_shortHash_(header.Signed_Hash);

  var img = '';
  try {
    var qr = DSX_QR(0, 'M');
    qr.addData(url);
    qr.make();
    img = '<img src="' + qr.createDataURL(3, 2) + '" width="86" height="86" alt="">';
  } catch (e) {
    img = '';   // the code below still verifies the document
  }

  return '<table style="width:100%;border-collapse:collapse;"><tr>' +
    (img ? '<td style="width:92px;vertical-align:top;">' + img + '</td>' : '') +
    '<td style="vertical-align:top;font-size:8.5pt;color:#334155;">' +
    '<div><strong>Verify this document</strong></div>' +
    '<div>Code <strong style="letter-spacing:.08em;">' + ipp_esc_(code) + '</strong></div>' +
    '<div style="word-break:break-all;">' + ipp_esc_(url) + '</div>' +
    '</td></tr></table>';
}

function dsx_webAppUrl_() {
  try {
    var u = ScriptApp.getService().getUrl();
    if (u) return u;
  } catch (e) {}
  try {
    return dsx_str_(PropertiesService.getScriptProperties().getProperty('DS_WEBAPP_URL'));
  } catch (e) { return ''; }
}

function dsx_consultantNameFor_(header) {
  var adm = dsx_admissionRow_(dsx_str_(header.IP_Number));
  return adm ? dsx_str_(adm.Consultant) : '';
}

function dsx_finalDiagnosisOf_(payload) {
  var d = payload.sections.DIAGNOSIS;
  if (!d || !d.content) return '';
  return dsx_str_(d.content.final) || dsx_str_(d.content.provisional);
}

/** The reason logged when the amendment that produced version n was started. */
function dsx_amendmentReasonFor_(summaryId, snapshotNo) {
  var events = dsx_recentEvents_(summaryId, 200);
  for (var i = 0; i < events.length; i++) {
    if (events[i].action === 'DS_AMEND_START') return events[i].comment;
  }
  return '';
}

// ---------------------------------------------------------------------------
// SECTION C — BROWSER PRINTING
// ---------------------------------------------------------------------------

/**
 * HTML for the client's hidden print iframe — the same pattern the lab reports
 * use. A draft print is watermarked and never available once signed.
 */
function ds_getPrintHtml(token, summaryId, snapshotRef, mode) {
  var lock = LockService.getScriptLock();
  try {
    var M = dsx_upper_(mode) || 'FINAL';
    var actor = dsx_requireRole_(token, (M === 'DRAFT') ? 'printDraft' : 'printFinal');

    var header = dsx_getHeader_(summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');

    var status = dsx_upper_(header.Status);
    if (M !== 'DRAFT' && status !== DSX_STATUS.SIGNED) {
      return dsx_err_('INVALID_STATE',
        'This summary is ' + status + '. Only a signed summary can be printed as a final copy; ' +
        'use Print draft instead.');
    }
    if (M === 'DRAFT' && status === DSX_STATUS.SIGNED) {
      M = 'REPRINT';
    }

    // A signed document always prints from the signed snapshot, never from the
    // working row — otherwise a later edit could reach paper unsigned.
    var ref = (status === DSX_STATUS.SIGNED)
      ? ('SIGNED:' + dsx_int_(header.Last_Signed_Snapshot_No))
      : (snapshotRef || 'WORKING');

    if (M !== 'DRAFT' && dsx_int_(header.Print_Count) > 0) M = 'REPRINT';

    // Rendering happens outside the lock.
    var html = dsx_renderDocumentHtml_(summaryId, ref, M, { printedBy: actor.displayName });

    // Only the print counter takes the lock, and only briefly.
    if (M !== 'DRAFT') {
      try {
        lock.waitLock(5000);
        dsx_resetHeaderCache_();
        var fresh = dsx_getHeader_(summaryId);
        if (fresh) {
          dsx_writeRow_(dsx_summariesSheet_(), fresh._row, {
            Print_Count: dsx_int_(fresh.Print_Count) + 1,
            Last_Printed_At: new Date()
          });
        }
      } finally {
        try { lock.releaseLock(); } catch (e) {}
      }
    }

    dsx_logEvent_(summaryId, actor, 'DS_PRINT', status, status,
                  dsx_int_(header.Current_Snapshot_No), '', '', { mode: M });

    return dsx_ok_('', { html: html, mode: M, ref: ref });

  } catch (e) {
    return dsx_fromError_(e);
  }
}

// ---------------------------------------------------------------------------
// SECTION D — STORED PDF
// ---------------------------------------------------------------------------

/**
 * Generates and stores the PDF for a signed summary. Runs OUTSIDE the signing
 * lock, by design: a Drive hiccup must never roll back a signature.
 *
 * @return {{success, message, data?}}
 */
function dsx_generatePdf_(summaryId) {
  var lock = LockService.getScriptLock();
  try {
    var header = dsx_getHeader_(summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No summary ' + summaryId + '.');
    if (dsx_upper_(header.Status) !== DSX_STATUS.SIGNED) {
      return dsx_err_('INVALID_STATE', 'Only a signed summary has a stored PDF.');
    }

    var snapshotNo = dsx_int_(header.Last_Signed_Snapshot_No);
    var html = dsx_renderDocumentHtml_(summaryId, 'SIGNED:' + snapshotNo, 'FINAL', {});

    var name = dsx_str_(header.Summary_ID) + '_v' + snapshotNo + '.pdf';
    var blob = Utilities.newBlob(html, 'text/html', name.replace(/\.pdf$/, '.html'))
                        .getAs('application/pdf').setName(name);

    var signedAt = dsx_toDate_(header.Signed_At) || new Date();
    var folder = dsx_pdfFolder_(
      Utilities.formatDate(signedAt, dsx_tz_(), 'yyyy'),
      Utilities.formatDate(signedAt, dsx_tz_(), 'MM'));

    // Replace rather than duplicate when a retry lands after a partial run.
    var existing = folder.getFilesByName(name);
    while (existing.hasNext()) existing.next().setTrashed(true);

    var file = folder.createFile(blob);
    // Never link-share. The file is served back through ds_getPdfBase64 only.

    lock.waitLock(10000);
    dsx_resetHeaderCache_();
    var fresh = dsx_getHeader_(summaryId);
    if (fresh) {
      dsx_writeRow_(dsx_summariesSheet_(), fresh._row, {
        Pdf_File_ID: file.getId(),
        Pdf_Status: 'READY'
      });
    }
    SpreadsheetApp.flush();

    return dsx_ok_('PDF stored.', { fileId: file.getId(), name: name });

  } catch (e) {
    try {
      var h = dsx_getHeader_(summaryId);
      if (h) dsx_writeRow_(dsx_summariesSheet_(), h._row, { Pdf_Status: 'FAILED' });
    } catch (e2) {}
    Logger.log('DS PDF generation failed for ' + summaryId + ': ' + e.message);
    return dsx_err_('VALIDATION_FAILED', 'PDF generation failed: ' + e.message);
  } finally {
    try { lock.releaseLock(); } catch (e3) {}
  }
}

function dsx_pdfFolder_(year, month) {
  var root = dsx_folderNamed_(DriveApp.getRootFolder(), DSX_PDF_FOLDER_ROOT);
  var sub = dsx_folderNamed_(root, DSX_PDF_FOLDER_SUB);
  return dsx_folderNamed_(dsx_folderNamed_(sub, year), month);
}

function dsx_folderNamed_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/**
 * Time-driven handler. Parameterless and idempotent, so it is safe to run by
 * hand from the editor. Retries PENDING and FAILED rows.
 */
function dsCron_processPdfQueue() {
  var done = 0, failed = 0;
  try {
    var sh = dsx_sheet_(DSX_SHEETS.SUMMARIES);
    if (!sh || sh.getLastRow() < 2) return 'No summaries.';

    var map = dsx_headerMap_(sh);
    var width = Math.max(1, sh.getLastColumn());
    var values = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();

    for (var i = 0; i < values.length; i++) {
      var status = dsx_upper_(values[i][map['Status']]);
      var pdf = dsx_upper_(values[i][map['Pdf_Status']]);
      if (status !== DSX_STATUS.SIGNED) continue;
      if (pdf !== 'PENDING' && pdf !== 'FAILED') continue;

      var res = dsx_generatePdf_(dsx_str_(values[i][map['Summary_ID']]));
      if (res.success) done++; else failed++;
      if (done + failed >= 20) break;   // stay inside the execution budget
    }
  } catch (e) {
    Logger.log('dsCron_processPdfQueue: ' + e.message);
    return 'Queue run failed: ' + e.message;
  }
  var msg = 'Discharge PDFs: ' + done + ' stored, ' + failed + ' failed.';
  Logger.log(msg);
  return msg;
}

/** Admin-run once. Installs the ten-minute trigger, without duplicating it. */
function installDischargePdfTrigger() {
  var existing = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'dsCron_processPdfQueue';
  });
  if (existing.length) return 'The discharge PDF trigger is already installed.';

  ScriptApp.newTrigger('dsCron_processPdfQueue').timeBased().everyMinutes(10).create();
  return 'Discharge PDF trigger installed (every 10 minutes).';
}

/** Serves the stored PDF to an authorised role. No Drive link is ever shared. */
function ds_getPdfBase64(token, summaryId, snapshotNo) {
  try {
    var actor = dsx_requireRole_(token, 'printFinal');
    var header = dsx_getHeader_(summaryId);
    if (!header) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found.');
    if (dsx_upper_(header.Status) !== DSX_STATUS.SIGNED) {
      return dsx_err_('INVALID_STATE', 'Only a signed summary has a stored PDF.');
    }

    var fileId = dsx_str_(header.Pdf_File_ID);
    if (!fileId) {
      // Generate on demand rather than telling the user to come back later.
      var gen = dsx_generatePdf_(summaryId);
      if (!gen.success) return gen;
      fileId = gen.data.fileId;
    }

    var file = DriveApp.getFileById(fileId);
    dsx_logEvent_(summaryId, actor, 'DS_PDF_DOWNLOAD', DSX_STATUS.SIGNED, DSX_STATUS.SIGNED,
                  dsx_int_(header.Last_Signed_Snapshot_No), '', '', {});

    return dsx_ok_('', {
      name: file.getName(),
      base64: Utilities.base64Encode(file.getBlob().getBytes())
    });
  } catch (e) {
    return dsx_fromError_(e);
  }
}

// ---------------------------------------------------------------------------
// SECTION E — PUBLIC VERIFICATION PAGE
// ---------------------------------------------------------------------------

/**
 * Answers ?verifyDS=<token>. Anonymous, and it carries NO clinical content:
 * facility, document type, summary ID, version, signer and registration
 * number, signed-at, validity, patient initials and a masked IP number. That
 * is everything needed to confirm a document is genuine and nothing more.
 *
 * Called from doGet in CodeMV.gs.
 */
function dsx_verifyPage_(rawToken) {
  var esc = function (v) {
    return String(v === null || v === undefined ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  };

  var state = { status: 'NOT FOUND', tone: '#b91c1c', rows: [] };

  try {
    var token = dsx_str_(rawToken);
    var hash = token ? dsx_sha256Hex_(token) : '';
    var found = null;

    if (hash) {
      var sh = dsx_sheet_(DSX_SHEETS.SNAPSHOTS);
      if (sh && sh.getLastRow() >= 2) {
        var rows = dsx_findRowsByKey_(sh, 'Verify_Token_Hash', hash);
        // Constant-time confirmation on the matched row, so a near-miss in the
        // sheet lookup cannot be used to probe for valid tokens.
        for (var i = 0; i < rows.length; i++) {
          var o = dsx_readRow_(sh, rows[i]);
          if (dsx_timingSafeEqual_(dsx_str_(o.Verify_Token_Hash), hash)) { found = o; break; }
        }
      }
    }

    if (found) {
      var header = dsx_getHeader_(dsx_str_(found.Summary_ID));
      var snapNo = dsx_int_(found.Snapshot_No);
      var lastSigned = header ? dsx_int_(header.Last_Signed_Snapshot_No) : snapNo;

      if (header && dsx_upper_(header.Status) === DSX_STATUS.SIGNED && snapNo === lastSigned) {
        state.status = 'VALID'; state.tone = '#047857';
      } else if (snapNo < lastSigned) {
        state.status = 'SUPERSEDED by version ' + lastSigned; state.tone = '#b45309';
      } else {
        state.status = 'NOT CURRENT'; state.tone = '#b45309';
      }

      var initials = dsx_initialsOf_(header);
      state.rows = [
        ['Document', dsx_documentTitle_(header && header.Discharge_Type)],
        ['Summary ID', dsx_str_(found.Summary_ID)],
        ['Version', String(snapNo)],
        ['Signed by', dsx_str_(header && header.Signed_By)],
        ['Registration No', dsx_str_(header && header.Signer_Reg_No)],
        ['Signed at', dsx_fmt_(header && header.Signed_At, 'dd-MMM-yyyy hh:mm a')],
        ['Patient', initials],
        ['IP number', dsx_maskIp_(header && header.IP_Number)]
      ];
    }
  } catch (e) {
    state = { status: 'NOT FOUND', tone: '#b91c1c', rows: [] };
  }

  var fac = dsx_facility_();
  var html =
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>Verify discharge summary</title><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;' +
    'background:#f8fafc;margin:0;padding:28px 16px;color:#0f172a;}' +
    '.card{max-width:420px;margin:0 auto;background:#fff;border:1px solid #e2e8f0;' +
    'border-radius:14px;padding:22px;box-shadow:0 4px 18px rgba(15,23,42,.06);}' +
    'h1{font-size:16px;margin:0 0 2px;}h2{font-size:12px;font-weight:400;color:#64748b;margin:0 0 16px;}' +
    '.state{font-size:18px;font-weight:700;padding:10px;border-radius:10px;text-align:center;' +
    'margin-bottom:16px;color:#fff;}' +
    'table{width:100%;border-collapse:collapse;font-size:13px;}' +
    'td{padding:6px 0;border-bottom:1px solid #f1f5f9;vertical-align:top;}' +
    'td.k{color:#64748b;width:42%;}' +
    '.note{font-size:11px;color:#94a3b8;margin-top:14px;line-height:1.5;}' +
    '</style></head><body><div class="card">' +
    '<h1>' + esc(fac.name) + '</h1>' +
    '<h2>Discharge summary verification</h2>' +
    '<div class="state" style="background:' + state.tone + '">' + esc(state.status) + '</div>' +
    (state.rows.length
      ? '<table>' + state.rows.map(function (r) {
          return '<tr><td class="k">' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td></tr>';
        }).join('') + '</table>'
      : '<div style="font-size:13px;color:#64748b;">This verification code does not match any ' +
        'signed document issued by this facility.</div>') +
    '<div class="note">This page confirms only that a document was issued and signed. ' +
    'It contains no clinical information. To obtain a copy of the summary, contact the facility.</div>' +
    '</div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Verify discharge summary')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function dsx_initialsOf_(header) {
  if (!header) return '';
  var adm = dsx_admissionRow_(dsx_str_(header.IP_Number));
  var name = adm ? dsx_str_(adm.Patient_Name) : '';
  return name.split(/\s+/).filter(String).map(function (w) {
    return w.charAt(0).toUpperCase() + '.';
  }).join(' ');
}

function dsx_maskIp_(ip) {
  var s = dsx_str_(ip);
  if (s.length <= 4) return s;
  return s.substring(0, s.length - 4).replace(/[^-]/g, '*') + s.slice(-4);
}

// ---------------------------------------------------------------------------
// SECTION F — THE SPIKE
// ---------------------------------------------------------------------------

/**
 * READ-ONLY probe of Apps Script's HTML-to-PDF converter. Writes nothing to
 * any sheet and stores nothing in Drive; it converts in memory and reports
 * what survived.
 *
 * Run this from the Apps Script editor on the real deployment and read the
 * log BEFORE relying on the stored PDF. Nothing in this project has previously
 * put an image or Tamil text through the converter, so nothing here is assumed.
 */
function ds_printSpike() {
  var lines = [];
  var probe = function (label, html) {
    try {
      var t = Date.now();
      var blob = Utilities.newBlob(html, 'text/html', 'spike.html').getAs('application/pdf');
      var bytes = blob.getBytes().length;
      lines.push('pass  ' + label + ' — ' + bytes + ' byte PDF in ' + (Date.now() - t) + ' ms');
      return bytes;
    } catch (e) {
      lines.push('FAIL  ' + label + ' — ' + e.message);
      return 0;
    }
  };

  var shell = function (body) {
    return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
           ipp_css_() + '</style></head><body>' + body + '</body></html>';
  };

  // (a) table-based A4 layout with a header repeated on every page
  var longRows = '';
  for (var i = 1; i <= 120; i++) {
    longRows += '<tr><td>Row ' + i + '</td><td>Paracetamol 500 mg</td><td>Oral</td>' +
                '<td>1-0-1</td><td>5 days</td></tr>';
  }
  var baseline = probe('(a) A4 table layout, repeating header, 120-row table',
    shell('<table class="page"><thead><tr><td>Repeating patient strip</td></tr></thead>' +
          '<tbody><tr><td><table style="width:100%"><tbody>' + longRows +
          '</tbody></table></td></tr></tbody></table>'));

  // (b) a base64 image via data URI — the QR and any logo depend on this
  var qrHtml = '';
  try {
    var qr = DSX_QR(0, 'M');
    qr.addData('https://example.invalid/?verifyDS=SPIKE');
    qr.make();
    qrHtml = '<img src="' + qr.createDataURL(3, 2) + '" width="86" height="86" alt="">';
    lines.push('pass  QR encoder produced a ' + qr.getModuleCount() + '-module data URI');
  } catch (e) {
    lines.push('FAIL  QR encoder threw: ' + e.message);
  }
  var withImg = probe('(b) base64 data-URI image', shell('<div>' + qrHtml + '</div>'));
  if (withImg && baseline) {
    lines.push('      note: the converter accepted the page. Whether the IMAGE is actually ' +
               'drawn can only be confirmed by opening the PDF — compare the byte size with ' +
               'and without the image below.');
  }
  var withoutImg = probe('(b2) the same page with the image removed', shell('<div></div>'));
  if (withImg && withoutImg) {
    lines.push('      image adds ' + (withImg - withoutImg) + ' bytes — a figure near zero ' +
               'means the image was dropped.');
  }

  // (c) Tamil text
  var tamil = 'மருந்து உணவு ' +
              'நேரம்';
  var withTa = probe('(c) Tamil text', shell('<div style="font-size:12pt">' + tamil + '</div>'));
  var withoutTa = probe('(c2) the same page in English', shell('<div style="font-size:12pt">Food and timing</div>'));
  if (withTa && withoutTa) {
    lines.push('      Tamil differs by ' + (withTa - withoutTa) + ' bytes — open the PDF and ' +
               'confirm the glyphs render rather than showing boxes.');
  }

  // (d) page breaks inside a long table
  probe('(d) page breaks inside a long table', shell(
    '<table class="page"><thead><tr><td>Header</td></tr></thead><tbody><tr><td>' +
    '<table style="width:100%"><tbody>' + longRows + longRows + '</tbody></table>' +
    '</td></tr></tbody></table>'));

  lines.push('');
  lines.push('WHAT TO DO WITH THIS');
  lines.push('  · If (b) shows the image adding real bytes and the PDF displays the QR, nothing');
  lines.push('    needs to change — the renderer already embeds it.');
  lines.push('  · If the image is dropped, the document still verifies: the code and URL are');
  lines.push('    plain text and always print. Decide whether to keep the QR for browser print');
  lines.push('    only, or move the stored PDF to a Google Docs template.');
  lines.push('  · If Tamil renders as boxes, set DS_PHRASE_LIBRARY entries to English only for');
  lines.push('    the stored PDF, or route the stored PDF through Google Docs. Browser printing');
  lines.push('    is unaffected — the browser has the fonts.');

  var report = 'ds_printSpike\n' + lines.join('\n');
  Logger.log(report);
  return { success: true, message: report, data: { lines: lines } };
}
