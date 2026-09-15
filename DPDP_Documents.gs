// ============================================================================
// DPDP_Documents.gs — Crescentia HealthTech
// Handing a patient their report without publishing it to the internet.
// ----------------------------------------------------------------------------
// WHAT WAS WRONG (DPDP_READINESS.md finding H1, HIGH)
//
// Five places in this project did the same thing:
//
//     file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
//
// for a lab report, a prescription, a pharmacy invoice or a lab invoice, and
// then put the resulting Drive URL into a WhatsApp message. Each of those
// files carries the patient's name, their patient ID, and — for a report —
// their results and the diagnosis they were ordered for.
//
// Four separate problems, from one line:
//
//   * The link never expires. A report shared in 2023 is still readable.
//   * The link is a bearer token with no owner. WhatsApp messages are
//     forwarded, quoted, backed up to someone else's cloud and restored onto
//     someone else's phone; every copy works.
//   * Google indexes what is linked to. "Anyone with the link" means anyone
//     who ever ends up with the link, by any route.
//   * Nothing recorded that the file existed, so nothing could revoke it.
//
// s.8(5) (reasonable security safeguards) and s.8(7) (keep no longer than
// necessary) at once.
//
// WHAT THIS FILE DOES INSTEAD
//
// Every document the clinic hands a patient is REGISTERED before it is sent.
// The register knows its patient, its expiry, who sent it and how, and it can
// be closed in one call — which is the part the bare setSharing() line could
// not do at all.
//
// WHY THE LINK IS A DRIVE LINK
//
// The first version of this file kept the file private and sent the patient a
// link back into this web app:
//
//     https://script.google.com/macros/s/AKfycb…/exec?doc=DOC-…&k=…
//
// That is the strongest link technically and the weakest link in practice. A
// patient receiving it on WhatsApp sees a script.google.com URL with a
// 32-character key on the end, arriving from a clinic, asking them to open
// their medical report. It looks exactly like the thing they are told every
// week not to click, it renders in WhatsApp with no preview and no file name,
// and on a phone that has never signed into the deploying Google account it
// could not be opened at all. A report a patient will not open is not a
// report delivered.
//
// So the document is published to Drive and the patient gets the link Drive
// itself generates:
//
//     https://drive.google.com/file/d/<id>/view
//
// which WhatsApp previews as a PDF, which opens in the Drive viewer the
// patient already has, and which downloads without an account.
//
// WHAT KEEPS THAT SAFE
//
// The four problems in the original setSharing() line were: the link never
// expired, nothing recorded that it existed, nothing could revoke it, and it
// was a permanent window rather than one document. Publishing to Drive
// re-opens only the first of those, and only until the sweep runs:
//
//   * Every published file is a row in Document_Grants with an expiry.
//   * dpdpExpireDocumentGrants() runs daily off dpdpInstallTriggers() and
//     sets every expired file back to PRIVATE. The link then stops working
//     for everyone who ever received it, forwarded copies included.
//   * dpdpRevokeDocumentLink() does the same immediately, on demand.
//   * The grant is one file. There is no folder link, and the folders
//     themselves are never shared.
//
// The honest trade this makes, written down rather than glossed: between
// sending and expiry, anyone holding the link can open that one document,
// and Drive does not tell us who did. The portal route below counted opens;
// Drive cannot. A patient who signs in to the portal gets an authenticated
// view instead, and that is still where a patient with a password should be
// sent. What a WhatsApp message gets is a document that expires.
//
// If link sharing is refused — a Workspace policy, a shared drive — the
// grant falls back to the portal-served route automatically rather than
// failing, and says so in the row. dpdpServeDocument_() below is that route,
// and it also keeps every link issued before this change working.
// ============================================================================

var DPDP_DOC_CFG = {
  SHEET: 'Document_Grants',

  /** How long a link lives. Long enough to reach a patient who is at work
   *  when the message arrives; short enough that a forwarded message is
   *  useless within the month. */
  DEFAULT_DAYS: 14,

  /** A document opened more often than this is not being read by one patient.
   *  The grant closes itself and the count is left in the register, because
   *  the number is evidence. Portal-served grants only: Drive does not report
   *  opens back to us, which is the one thing that route gives up. */
  MAX_OPENS: 25,

  KEY_LENGTH: 32,

  /** DRIVE hands out the link Drive generates; PORTAL keeps the file private
   *  and serves it through ?doc=&k= below. DRIVE is what a patient can
   *  actually open from a WhatsApp message. */
  DELIVERY_DEFAULT: 'DRIVE'
};

function dpdp_grantSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DPDP_DOC_CFG.SHEET, [
    'Grant_ID', 'Key_Digest', 'File_ID', 'File_Name', 'Doc_Type', 'Patient_ID',
    'Issued_At', 'Issued_By', 'Expires_At', 'Status', 'Opens', 'Last_Opened_At',
    'Revoked_At', 'Revoked_By',
    // Appended, so rows written before this change keep their meaning: a
    // blank Delivery is a PORTAL grant, which is what they all were.
    'Delivery', 'Link'
  ]);
}

/** A random key for one document. Platform randomness, not Math.random(). */
function dpdp_newKey_() {
  var s = '';
  while (s.length < DPDP_DOC_CFG.KEY_LENGTH) {
    s += Utilities.getUuid().replace(/-/g, '');
  }
  return s.substring(0, DPDP_DOC_CFG.KEY_LENGTH);
}

/** What the register stores in place of the key itself. */
function dpdp_keyDigest_(key) {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(key)));
}

/** The published web app URL, or '' when the script has never been deployed. */
function dpdp_webAppUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/**
 * Registers a file against a patient and returns the link to send them.
 *
 * Call this INSTEAD of a bare setSharing(ANYONE_WITH_LINK). The difference
 * is not the sharing — it is that the file is now in a register that knows
 * when it expires and can close it.
 *
 * @param {DriveApp.File} file
 * @param {string} docType      LAB_REPORT, OP_PRESCRIPTION, …
 * @param {string} patientId
 * @param {string} issuedBy     username of the person sending it
 * @param {number} [days]       lifetime; defaults to DPDP_DOC_CFG.DEFAULT_DAYS
 * @return {{success:boolean, url:string, grantId:string, expiresAt:string,
 *           delivery:string, message:string}}
 */
function dpdpIssueDocumentLink_(file, docType, patientId, issuedBy, days) {
  try {
    if (!file) return { success: false, url: '', message: 'No file to share.' };

    var now = new Date();
    var life = (days && days > 0) ? days : DPDP_DOC_CFG.DEFAULT_DAYS;
    var expires = new Date(now.getTime() + life * 86400000);
    var key = dpdp_newKey_();
    var id = 'DOC-' + Utilities.formatDate(now, DPDP_CFG.TZ, 'yyMMdd-HHmmss') + '-' +
             Utilities.getUuid().substring(0, 4).toUpperCase();

    // ---- publish to Drive, which is the link the patient can open --------
    var delivery = '', url = '';
    if (DPDP_DOC_CFG.DELIVERY_DEFAULT === 'DRIVE') {
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        delivery = 'DRIVE';
        // Built rather than read from getUrl(): getUrl() returns the /edit or
        // /view form Drive happens to choose for the type, and a PDF handed to
        // a patient should always open in the viewer.
        url = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';
      } catch (e) {
        // A Workspace policy or a shared drive can refuse link sharing. That
        // is a reason to fall back, not to fail: the patient still needs the
        // document, and the portal route below still works.
        delivery = '';
      }
    }

    // ---- fall back to the portal-served route ----------------------------
    if (!delivery) {
      try {
        file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      } catch (e) { /* the grant still binds */ }
      var base = dpdp_webAppUrl_();
      if (!base) {
        return { success: false, url: '',
                 message: 'Drive would not publish this file and the script has ' +
                          'no deployed web app URL to serve it from either, so ' +
                          'there is no link to send. Deploy the web app, or check ' +
                          'whether link sharing is blocked for this account.' };
      }
      delivery = 'PORTAL';
      url = base + '?doc=' + encodeURIComponent(id) + '&k=' + encodeURIComponent(key);
    }

    var sh = dpdp_grantSheet_();
    var m = dc_headerMap_(sh);
    var row = new Array(sh.getLastColumn()).fill('');
    row[m['Grant_ID']]      = id;
    row[m['Key_Digest']]    = dpdp_keyDigest_(key);
    row[m['File_ID']]       = file.getId();
    row[m['File_Name']]     = file.getName();
    row[m['Doc_Type']]      = dpdp_str_(docType);
    row[m['Patient_ID']]    = dpdp_str_(patientId).toUpperCase();
    row[m['Issued_At']]     = now;
    row[m['Issued_By']]     = dpdp_str_(issuedBy) || 'SYSTEM';
    row[m['Expires_At']]    = expires;
    row[m['Status']]        = 'ACTIVE';
    row[m['Opens']]         = 0;
    row[m['Delivery']]      = delivery;
    // The link is stored so the register can answer "what was this patient
    // sent" (s.11(1)(b)) without rebuilding it. The KEY is not stored — only
    // its digest — so a PORTAL row in the register is not itself usable.
    row[m['Link']]          = (delivery === 'DRIVE') ? url : '';
    sh.appendRow(row);
    dc_invalidate_(DPDP_DOC_CFG.SHEET);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: dpdp_str_(issuedBy) || 'SYSTEM', role: '' },
                'DOCUMENT_LINK_ISSUED', 'Document', id,
                { docType: dpdp_str_(docType), patientId: dpdp_str_(patientId),
                  delivery: delivery, expiresAt: dpdp_fmt_(expires) });
    } catch (e) {}

    return { success: true, url: url, grantId: id, delivery: delivery,
             expiresAt: dpdp_fmt_(expires),
             message: 'Link valid until ' +
                      Utilities.formatDate(expires, DPDP_CFG.TZ, 'dd-MMM-yyyy') + '.' };
  } catch (err) {
    return { success: false, url: '', message: 'Could not issue a document link: ' +
             err.message };
  }
}

/**
 * Takes a published file back to private. Used by the expiry sweep and by a
 * revocation, which want the identical effect: the link stops working for
 * everyone who ever received it.
 *
 * @return {string} '' on success, or why it could not be done
 */
function dpdp_unpublish_(fileId) {
  if (!dpdp_str_(fileId)) return 'no file id on the grant';
  try {
    DriveApp.getFileById(dpdp_str_(fileId))
            .setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    return '';
  } catch (e) {
    // A file the clinic has since deleted is the same outcome as a file made
    // private, so it is reported rather than treated as a failure.
    return e.message;
  }
}

/**
 * Resolves a grant from the URL. Fails closed, and says as little as it can:
 * a wrong key and an expired grant get the same answer, because the
 * difference is only useful to somebody guessing.
 *
 * @return {{ok:boolean, row:number, fileId:string, name:string,
 *           patientId:string, docType:string, reason:string}}
 */
function dpdp_resolveGrant_(grantId, key) {
  var sh = dpdp_grantSheet_();
  var m = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  var want = dpdp_str_(grantId).toUpperCase();
  if (!want || !dpdp_str_(key)) return { ok: false, reason: 'MISSING' };

  for (var i = 1; i < (data ? data.length : 0); i++) {
    if (dpdp_str_(data[i][m['Grant_ID']]).toUpperCase() !== want) continue;

    var status = dpdp_str_(data[i][m['Status']]).toUpperCase();
    var expires = (typeof cresc_parseDate_ === 'function')
      ? cresc_parseDate_(data[i][m['Expires_At']])
      : new Date(data[i][m['Expires_At']]);
    var opens = parseInt(data[i][m['Opens']], 10) || 0;

    // Constant-time, like a password: the digest is compared, never the key.
    var match = (typeof crescPwdEquals_ === 'function')
      ? crescPwdEquals_(dpdp_keyDigest_(key), dpdp_str_(data[i][m['Key_Digest']]))
      : (dpdp_keyDigest_(key) === dpdp_str_(data[i][m['Key_Digest']]));

    if (!match) return { ok: false, reason: 'DENIED' };
    if (status !== 'ACTIVE') return { ok: false, reason: status };
    if (!expires || isNaN(expires.getTime()) || expires.getTime() < Date.now()) {
      sh.getRange(i + 1, m['Status'] + 1).setValue('EXPIRED');
      return { ok: false, reason: 'EXPIRED' };
    }
    if (opens >= DPDP_DOC_CFG.MAX_OPENS) {
      sh.getRange(i + 1, m['Status'] + 1).setValue('EXHAUSTED');
      return { ok: false, reason: 'EXHAUSTED' };
    }

    return { ok: true, row: i + 1,
             fileId: dpdp_str_(data[i][m['File_ID']]),
             name: dpdp_str_(data[i][m['File_Name']]),
             patientId: dpdp_str_(data[i][m['Patient_ID']]),
             docType: dpdp_str_(data[i][m['Doc_Type']]),
             opens: opens, reason: '' };
  }
  return { ok: false, reason: 'DENIED' };
}

/** One page, one message, no clinical content. Used for every refusal. */
function dpdp_docMessage_(title, body) {
  var html =
    '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>' + title + '</title><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;' +
    'background:#f8fafc;margin:0;padding:40px 20px;display:flex;justify-content:center;' +
    'align-items:center;min-height:100vh}' +
    '.card{background:#fff;padding:36px 26px;border-radius:16px;max-width:380px;width:100%;' +
    'border:1px solid #e5e7eb;box-shadow:0 4px 20px rgba(0,0,0,.08);text-align:center}' +
    'h2{color:#0f172a;margin:0 0 10px;font-size:1.15rem}p{color:#64748b;font-size:.9rem;line-height:1.5;margin:0}' +
    '</style></head><body><div class="card"><h2>' + title + '</h2><p>' + body + '</p></div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle(title);
}

/**
 * doGet route for ?doc=&k=. Serves one registered document, once counted.
 *
 * The file is fetched with the deploying account's Drive access — the same
 * access that created it — and handed to the browser as a download. The file
 * itself stays private, so this route is the ONLY way to it, and this route
 * has an expiry, a counter and an audit row.
 */
function dpdpServeDocument_(grantId, key) {
  var g;
  try { g = dpdp_resolveGrant_(grantId, key); }
  catch (e) { g = { ok: false, reason: 'ERROR' }; }

  if (!g.ok) {
    if (g.reason === 'EXPIRED' || g.reason === 'EXHAUSTED') {
      return dpdp_docMessage_('This link has expired',
        'Documents are shared for a limited time so that a message forwarded ' +
        'months later cannot still open your records. Please contact the clinic ' +
        'for a fresh copy.');
    }
    if (g.reason === 'REVOKED') {
      return dpdp_docMessage_('This link has been withdrawn',
        'The clinic has withdrawn access to this document. Please contact them ' +
        'if you still need a copy.');
    }
    // Everything else — wrong key, unknown id, missing parameter — gets one
    // answer, so the page cannot be used to test which ids exist.
    return dpdp_docMessage_('This link is not valid',
      'Check that you have the whole link from the message, including the part ' +
      'after the last "&". If it still does not open, ask the clinic to send it ' +
      'again.');
  }

  try {
    var file = DriveApp.getFileById(g.fileId);
    var blob = file.getBlob();
    var b64 = Utilities.base64Encode(blob.getBytes());
    var name = (g.name || 'document.pdf').replace(/[^\w.\- ]/g, '_');

    // Counted and stamped BEFORE the page is built: if the render fails the
    // open still happened, and an undercount is the wrong way to be wrong.
    try {
      var sh = dpdp_grantSheet_();
      var m = dc_headerMap_(sh);
      sh.getRange(g.row, m['Opens'] + 1).setValue((g.opens || 0) + 1);
      sh.getRange(g.row, m['Last_Opened_At'] + 1).setValue(new Date());
      dc_invalidate_(DPDP_DOC_CFG.SHEET);
    } catch (e) {}

    try {
      // s.11(1)(b) and finding M2: who read what, and when. The reader is not
      // identified — a link is a capability, not an identity — and the row
      // says so rather than implying a name it does not have.
      logAudit_({ username: 'LINK:' + grantId, role: 'document-link' },
                'DOCUMENT_OPENED', 'Document', String(grantId),
                { docType: g.docType, patientId: g.patientId,
                  open: (g.opens || 0) + 1, of: DPDP_DOC_CFG.MAX_OPENS });
    } catch (e) {}

    var page =
      '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<title>Your document</title><style>' +
      'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;' +
      'background:#f8fafc;margin:0;padding:40px 20px;display:flex;justify-content:center;' +
      'align-items:center;min-height:100vh}' +
      '.card{background:#fff;padding:36px 24px;border-radius:16px;max-width:360px;width:100%;' +
      'border:1px solid #e5e7eb;box-shadow:0 4px 20px rgba(0,0,0,.08);text-align:center}' +
      '.icon{font-size:44px}h2{color:#0f172a;margin:12px 0 6px;font-size:1.15rem}' +
      'p{color:#64748b;font-size:.85rem;margin:0 0 6px}' +
      '.btn{background:#10b981;color:#fff;border:0;padding:16px;border-radius:12px;' +
      'font-size:15px;font-weight:700;width:100%;margin-top:22px;cursor:pointer}' +
      '.note{color:#94a3b8;font-size:.72rem;margin-top:16px;line-height:1.5}' +
      '</style></head><body><div class="card">' +
      '<div class="icon">📄</div><h2>Your document is ready</h2>' +
      '<p>' + name.replace(/[<>&]/g, '') + '</p>' +
      '<button id="dl" class="btn">Download</button>' +
      '<div class="note">This link is personal to you and stops working after a ' +
      'short time. Please do not forward it — anyone who receives it can open ' +
      'this document.</div></div>' +
      '<script>document.getElementById("dl").addEventListener("click",function(){' +
      'this.innerText="Downloading…";var a=document.createElement("a");' +
      'a.href="data:' + blob.getContentType() + ';base64,' + b64 + '";' +
      'a.download="' + name + '";document.body.appendChild(a);a.click();' +
      'document.body.removeChild(a);var b=this;setTimeout(function(){b.innerText="Downloaded";},2000);' +
      '});<\/script></body></html>';

    return HtmlService.createHtmlOutput(page).setTitle('Your document')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (err) {
    return dpdp_docMessage_('This document is no longer available',
      'The file could not be opened. Please contact the clinic for a copy.');
  }
}

/**
 * FRONTEND ENTRY. Withdraws one document link immediately.
 *
 * s.6(6) is about consent, but the same principle applies to a document
 * already sent: the clinic has to be able to take it back, and "ask Google to
 * un-share a file nobody wrote down" is not a procedure.
 */
function dpdpRevokeDocumentLink(grantId, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['dpdp.manage', 'admin.config', 'patient.write']);
    var sh = dpdp_grantSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var want = dpdp_str_(grantId).toUpperCase();

    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Grant_ID']]).toUpperCase() !== want) continue;

      // A DRIVE grant is only withdrawn when the FILE stops being readable.
      // Marking the row REVOKED and leaving the file published would be a
      // register that lies, which is worse than no register.
      var delivery = dpdp_str_(data[i][m['Delivery']]).toUpperCase() || 'PORTAL';
      var problem = '';
      if (delivery === 'DRIVE') {
        problem = dpdp_unpublish_(dpdp_str_(data[i][m['File_ID']]));
      }

      sh.getRange(i + 1, m['Status'] + 1).setValue(problem ? 'REVOKE_FAILED' : 'REVOKED');
      sh.getRange(i + 1, m['Revoked_At'] + 1).setValue(new Date());
      sh.getRange(i + 1, m['Revoked_By'] + 1).setValue(actor.username);
      dc_invalidate_(DPDP_DOC_CFG.SHEET);
      SpreadsheetApp.flush();
      try {
        logAudit_({ username: actor.username, role: actor.role },
                  'DOCUMENT_LINK_REVOKED', 'Document', want,
                  { delivery: delivery, problem: problem });
      } catch (e) {}

      if (problem) {
        return { success: false,
                 message: 'The register row for ' + want + ' is marked, but the ' +
                          'file itself could NOT be made private (' + problem +
                          '). Anyone holding the link can still open it — ' +
                          'un-share or delete the file in Drive.' };
      }
      return { success: true,
               message: 'Link ' + want + ' withdrawn. The file is private again, ' +
                        'so forwarded copies of the link no longer open it.' };
    }
    return { success: false, message: 'No document link with the id ' + grantId + '.' };
  } catch (err) {
    return { success: false, message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * FRONTEND ENTRY. Every document link issued for one patient, newest first —
 * which is also the s.11(1)(b) answer to "who has my data been given to".
 */
function dpdpListDocumentLinks(patientId, sessionToken) {
  try {
    crescRequire_(sessionToken, ['dpdp.manage', 'patient.read']);
    var sh = dpdp_grantSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var want = dpdp_str_(patientId).toUpperCase();
    var rows = [];
    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (want && dpdp_str_(data[i][m['Patient_ID']]).toUpperCase() !== want) continue;
      rows.push({
        grantId: dpdp_str_(data[i][m['Grant_ID']]),
        document: dpdp_str_(data[i][m['File_Name']]),
        docType: dpdp_str_(data[i][m['Doc_Type']]),
        patientId: dpdp_str_(data[i][m['Patient_ID']]),
        issuedAt: dpdp_str_(data[i][m['Issued_At']]),
        issuedBy: dpdp_str_(data[i][m['Issued_By']]),
        expiresAt: dpdp_str_(data[i][m['Expires_At']]),
        status: dpdp_str_(data[i][m['Status']]),
        delivery: dpdp_str_(data[i][m['Delivery']]) || 'PORTAL',
        link: dpdp_str_(data[i][m['Link']]),
        // Only a PORTAL grant has a meaningful open count. Drive does not
        // report reads back to us, so 0 here means "not counted", not "not
        // opened", and the screen says so rather than implying nobody looked.
        opens: parseInt(data[i][m['Opens']], 10) || 0,
        opensCounted: (dpdp_str_(data[i][m['Delivery']]).toUpperCase() !== 'DRIVE'),
        lastOpened: dpdp_str_(data[i][m['Last_Opened_At']])
      });
    }
    rows.reverse();
    return { success: true, rows: rows, message: '' };
  } catch (err) {
    return { success: false, rows: [],
             message: String(err.message || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * Closes every grant past its expiry. Attach to the daily trigger installed by
 * dpdpInstallTriggers(); it costs nothing when there is nothing to close.
 *
 * THIS IS THE JOB THAT MAKES A DRIVE LINK SAFE TO SEND. Until it runs, an
 * expired grant's file is still published; after it runs, the link is dead
 * for every copy of the message it was ever sent in. If it is not installed,
 * the clinic is back to permanent Drive links — dpdpReadinessCheck() reports
 * a missing trigger for exactly this reason.
 *
 * A file that cannot be made private is counted and NAMED. It is the one
 * outcome here that leaves a patient's document readable, so it must not
 * disappear into a success message.
 */
function dpdpExpireDocumentGrants() {
  var sh = dpdp_grantSheet_();
  var m = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  if (!data || data.length < 2) return 'No document grants.';

  var now = Date.now(), closed = 0, live = 0, unpublished = 0, stuck = [];
  for (var i = 1; i < data.length; i++) {
    if (dpdp_str_(data[i][m['Status']]).toUpperCase() !== 'ACTIVE') continue;
    var exp = (typeof cresc_parseDate_ === 'function')
      ? cresc_parseDate_(data[i][m['Expires_At']]) : new Date(data[i][m['Expires_At']]);
    if (exp && !isNaN(exp.getTime()) && exp.getTime() >= now) { live++; continue; }

    var delivery = dpdp_str_(data[i][m['Delivery']]).toUpperCase() || 'PORTAL';
    var problem = '';
    if (delivery === 'DRIVE') {
      problem = dpdp_unpublish_(dpdp_str_(data[i][m['File_ID']]));
      if (!problem) unpublished++;
    }

    if (problem) {
      sh.getRange(i + 1, m['Status'] + 1).setValue('EXPIRY_FAILED');
      stuck.push(dpdp_str_(data[i][m['Grant_ID']]) + ' (' +
                 dpdp_str_(data[i][m['File_Name']]) + '): ' + problem);
    } else {
      sh.getRange(i + 1, m['Status'] + 1).setValue('EXPIRED');
      closed++;
    }
  }
  dc_invalidate_(DPDP_DOC_CFG.SHEET);
  var msg = closed + ' document link(s) expired (' + unpublished +
            ' Drive file(s) made private again), ' + live + ' still live.' +
            (stuck.length
              ? '\n\nSTILL READABLE — un-share or delete these in Drive by hand:\n' +
                stuck.join('\n')
              : '');
  Logger.log(msg);
  return msg;
}
