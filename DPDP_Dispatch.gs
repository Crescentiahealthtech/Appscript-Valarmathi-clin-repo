// ============================================================================
// DPDP_Dispatch.gs — Crescentia HealthTech
// Nothing leaves the building for a patient who did not agree to be sent it.
// ----------------------------------------------------------------------------
// WHAT WAS MISSING
//
// DPDP_PURPOSES (DPDP_Compliance.gs) has carried a COMMUNICATION purpose from
// the day the consent register was built:
//
//     'Reports and reminders by WhatsApp or email —
//      Sending your prescription, lab report or invoice to your mobile
//      number or email address.'
//
// It was asked at the desk, recorded, withdrawable in the portal, and read by
// NOTHING. Eight dispatch endpoints — the pharmacy invoice, the lab invoice,
// the lab report, the archived lab invoice and the OP prescription, each with
// a WhatsApp and an email path — generated a PDF and sent it without once
// looking at the register. A patient who had explicitly REFUSED reports by
// WhatsApp, or withdrawn that consent in the portal an hour earlier, got the
// message anyway.
//
// A consent that no code consults is not a consent record. It is a note about
// a conversation, and s.6(10) puts the burden of proving consent on the
// clinic.
//
// WHAT THIS FILE DOES
//
//   1. dpdpRequireDispatchConsent_() — the gate. One line at the top of each
//      dispatch endpoint. Refuses unless the register says GIVEN, and says
//      what to do about it when it refuses.
//
//   2. dpdp_resolvePatientFor_() — the patient this document belongs to.
//      Three of the five callers of dpdpIssueDocumentLink_ were passing an
//      INVOICE NUMBER or an ORDER ID into its patientId argument, and a
//      fourth passed a variable that did not exist in its scope, so the
//      Document_Grants register could not answer the one question s.11(1)(b)
//      gives a patient the right to ask: what have you sent about me, and to
//      whom. The gate needs the real id, so the real id is resolved here and
//      the register gets it too.
//
//   3. THE NOTICE, AND WHY IT NAMES META.
//      s.5(1) says the notice must describe what is actually done with the
//      data. What is actually done, when a desk clicks "WhatsApp", is that
//      the patient's name, their document and a link to it are handed to
//      WhatsApp — which is Meta Platforms, a company outside the clinic,
//      outside this application and outside India. "We will send it to your
//      mobile" does not describe that. The text below does, in the words a
//      patient reads before they agree, and the same text is shown at the
//      desk before the send so the person clicking knows what they are
//      telling the patient.
//
//      The clinic has no Data Processing Agreement with Meta and cannot get
//      one through consumer WhatsApp — see docs/PROCESSORS.md. That makes
//      naming them a requirement rather than a courtesy: the patient's
//      agreement is the only lawful basis this route has.
// ============================================================================

var DPDP_DISPATCH = {

  /**
   * Every channel a patient document can leave by, and who carries it.
   *
   * `processor` is the third party that sees the message. It is written out
   * in full because a notice that says "a messaging service" has told the
   * patient nothing they did not already assume.
   */
  CHANNELS: {
    WHATSAPP: {
      key: 'WHATSAPP',
      label: 'WhatsApp',
      processor: 'Meta Platforms (WhatsApp)',
      notice:
        'Your document is sent as a WhatsApp message. WhatsApp is run by ' +
        'Meta, a company outside this clinic and outside India, and the ' +
        'message passes through their service to reach your phone. The ' +
        'message carries your name and a link that opens the document. ' +
        'Anyone who gets hold of that message — if your phone is shared, if ' +
        'the message is forwarded, or if your chats are backed up somewhere ' +
        'else — can open it, so please do not forward it. The link stops ' +
        'working after a short time. You can tell us to stop sending ' +
        'documents this way at any time, in the app under My privacy, or by ' +
        'telling us at the clinic.'
    },
    EMAIL: {
      key: 'EMAIL',
      label: 'Email',
      processor: 'Google (Gmail)',
      notice:
        'Your document is sent as an email with the file attached, from the ' +
        'clinic\'s Gmail account through Google\'s mail service. Email is not ' +
        'encrypted end to end: your email provider, and anyone with access to ' +
        'your inbox, can read it. You can tell us to stop sending documents ' +
        'this way at any time, in the app under My privacy, or by telling us ' +
        'at the clinic.'
    }
  },

  /** The consent purpose all of this is gated on. */
  PURPOSE: 'COMMUNICATION'
};

function dpdp_channel_(channel) {
  var k = String(channel || '').trim().toUpperCase();
  return DPDP_DISPATCH.CHANNELS[k] || null;
}

// ---------------------------------------------------------------------------
// WHERE THE DOCUMENT'S PATIENT COMES FROM
// ---------------------------------------------------------------------------

/**
 * The patient id behind a document reference, whatever kind of reference it is.
 *
 * Every dispatch endpoint is handed an invoice number, a bill id, a lab order
 * id or an encounter id — never a patient id. Rather than four half-written
 * lookups scattered across four files (which is what produced the bug this
 * fixes), one function knows them all.
 *
 * @param {string} kind  PHARMACY_INVOICE | LAB_INVOICE | LAB_ORDER | OP_ENCOUNTER
 * @param {string} ref   the id of that thing
 * @return {string} the patient id, or '' when it genuinely cannot be resolved
 */
function dpdp_resolvePatientFor_(kind, ref) {
  var want = String(ref || '').trim().toUpperCase();
  if (!want) return '';
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var lookup = function (sheetName, refHeader, patientHeader) {
    try {
      var sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return '';
      var m = dc_headerMap_(sh);
      if (m[refHeader] === undefined || m[patientHeader] === undefined) return '';
      var data = sh.getDataRange().getDisplayValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][m[refHeader]] || '').trim().toUpperCase() !== want) continue;
        return String(data[i][m[patientHeader]] || '').trim().toUpperCase();
      }
    } catch (e) {}
    return '';
  };

  switch (String(kind || '').toUpperCase()) {
    case 'PHARMACY_INVOICE':
      // Pharmacy_Invoices is read positionally elsewhere in Pharmacy.gs
      // (column D is the patient), but by header here so a column added to
      // that sheet later does not silently move the answer.
      return lookup('Pharmacy_Invoices', 'Invoice_No', 'Patient_ID') ||
             dpdp_pharmacyInvoicePatientByIndex_(want);

    case 'LAB_INVOICE':
      // A lab bill id, or the order id the archive path passes instead.
      return lookup('LAB_BILLING', 'BillID', 'PatientID') ||
             lookup('LAB_BILLING', 'OrderID', 'PatientID') ||
             lookup('LAB_ORDERS', 'OrderID', 'PatientID');

    case 'LAB_ORDER':
      return lookup('LAB_ORDERS', 'OrderID', 'PatientID');

    case 'OP_ENCOUNTER':
      // OP_Encounters is written as a positional row (A encounter, B patient)
      // by saveOPEncounter, and its header row has been spelled differently
      // across deployments, so the name lookup is tried and then the columns
      // the writer actually uses.
      return lookup('OP_Encounters', 'Encounter_ID', 'Patient_ID') ||
             lookup('OP_Encounters', 'Appt_ID', 'Patient_ID') ||
             dpdp_firstColMatch_('OP_Encounters', want, 0, 1);
  }
  return '';
}

/**
 * Positional fallback: find the row whose column `refCol` equals `want` and
 * return column `patCol`. For the two sheets in this project that are written
 * by index rather than by header.
 */
function dpdp_firstColMatch_(sheetName, want, refCol, patCol) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return '';
    var data = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][refCol] || '').trim().toUpperCase() !== want) continue;
      var pid = String(data[i][patCol] || '').trim().toUpperCase();
      return (pid === 'WALK-IN' || pid === 'DIRECT') ? '' : pid;
    }
  } catch (e) {}
  return '';
}

/**
 * Fallback for a Pharmacy_Invoices sheet whose header row does not match the
 * names above. Column D has been the patient id since the sheet was created
 * and getInvoiceForPrint() already reads it that way.
 */
function dpdp_pharmacyInvoicePatientByIndex_(invoiceNoUpper) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pharmacy_Invoices');
    if (!sh || sh.getLastRow() < 2) return '';
    var data = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toUpperCase() !== invoiceNoUpper) continue;
      var pid = String(data[i][3] || '').trim().toUpperCase();
      return (pid === 'WALK-IN' || pid === 'DIRECT') ? '' : pid;
    }
  } catch (e) {}
  return '';
}

// ---------------------------------------------------------------------------
// WHERE THE CONSENT COMES FROM
// ---------------------------------------------------------------------------

/**
 * The current state of one consent purpose for one patient.
 *
 * Deliberately NOT getConsentStatus(): that one is a frontend entry with its
 * own permission check and it builds the whole purpose list. This is the
 * internal read, used by a gate that already knows who is calling.
 *
 * @return {{state:string, at:string, by:string, noticeVersion:string, stale:boolean}}
 *         state is GIVEN | REFUSED | WITHDRAWN | NOT_ASKED
 */
function dpdp_consentState_(patientId, purposeKey) {
  var out = { state: 'NOT_ASKED', at: '', by: '', noticeVersion: '', stale: false };
  var want = String(patientId || '').trim().toUpperCase();
  var purpose = String(purposeKey || '').trim().toUpperCase();
  if (!want || !purpose) return out;

  try {
    var sh = dpdp_consentSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (dpdp_str_(data[i][m['Patient_ID']]).toUpperCase() !== want) continue;
      if (dpdp_str_(data[i][m['Purpose']]).toUpperCase() !== purpose) continue;
      // The register is append-only, so the LAST matching row is the answer.
      out.state = dpdp_str_(data[i][m['Withdrawn_At']])
        ? 'WITHDRAWN'
        : dpdp_str_(data[i][m['Decision']]).toUpperCase();
      out.at = dpdp_fmt_(data[i][m['Recorded_At']]);
      out.by = dpdp_str_(data[i][m['Recorded_By']]);
      out.noticeVersion = dpdp_str_(data[i][m['Notice_Version']]);
    }
    if (out.noticeVersion && out.noticeVersion !== dpdp_noticeVersion_()) {
      out.stale = true;
    }
  } catch (e) { /* an unreadable register is NOT_ASKED, which fails closed */ }
  return out;
}

// ---------------------------------------------------------------------------
// THE GATE
// ---------------------------------------------------------------------------

/**
 * May this document be sent to this patient on this channel.
 *
 * FAILS CLOSED, including on a patient id that could not be resolved: a
 * dispatch whose patient is unknown is a dispatch whose consent cannot be
 * checked, and "we could not tell" is not permission.
 *
 * The one deliberate exception is a walk-in with no patient record at all
 * (`allowAnonymous`), used by the pharmacy counter: an over-the-counter sale
 * to somebody who never gave a name has no patient to consent, and no
 * clinical content on the invoice either. The register still gets a row.
 *
 * @param {string} patientId
 * @param {string} channel        WHATSAPP | EMAIL
 * @param {Object} actor          from crescActor_/crescRequire_
 * @param {string} docType        for the audit row
 * @param {Object} [opts]         { allowAnonymous:boolean }
 * @return {{ok:boolean, state:string, message:string, notice:string}}
 */
function dpdpRequireDispatchConsent_(patientId, channel, actor, docType, opts) {
  opts = opts || {};
  var ch = dpdp_channel_(channel);
  if (!ch) {
    return { ok: false, state: 'NOT_ASKED', notice: '',
             message: 'Unknown dispatch channel "' + channel + '".' };
  }

  var pid = String(patientId || '').trim().toUpperCase();
  if (!pid) {
    if (opts.allowAnonymous) {
      return { ok: true, state: 'NO_PATIENT', notice: ch.notice, message: '' };
    }
    return { ok: false, state: 'UNKNOWN_PATIENT', notice: ch.notice,
             message: 'This document is not linked to a patient record, so there ' +
                      'is no consent to check. Open the record and send it from ' +
                      'there, or record the patient on the bill first.' };
  }

  var c = dpdp_consentState_(pid, DPDP_DISPATCH.PURPOSE);

  if (c.state !== 'GIVEN') {
    var why = {
      REFUSED:   pid + ' was asked and said no to receiving documents this way.',
      WITHDRAWN: pid + ' has withdrawn consent to receive documents this way.',
      NOT_ASKED: pid + ' has never been asked whether we may send documents ' +
                 'this way, so we may not.'
    }[c.state] || (pid + ' has no consent on record for this.');

    return {
      ok: false, state: c.state, notice: ch.notice,
      message: why + '\n\nAsk them, read them what the message involves, and ' +
               'record their answer — Privacy → Consent, or the consent block ' +
               'on the registration screen. If they agree, this will send. ' +
               'Nothing was sent.'
    };
  }

  // Consent given against an OLDER notice is not consent to the current one
  // (s.5), but refusing a send over it would strand every patient whenever
  // the clinic reworded its notice. It is allowed and flagged, and the
  // re-ask shows up in the readiness report rather than at the counter.
  try {
    logAudit_({ username: (actor && actor.username) || 'SYSTEM',
                role: (actor && actor.role) || '' },
              'DPDP_DISCLOSURE', 'Patient', pid,
              { channel: ch.key, processor: ch.processor,
                docType: String(docType || ''),
                consentAt: c.at, noticeVersion: c.noticeVersion,
                staleNotice: c.stale });
  } catch (e) {}

  return { ok: true, state: 'GIVEN', notice: ch.notice, message: '' };
}

// ---------------------------------------------------------------------------
// WHAT THE DESK AND THE PATIENT SEE
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. The notice for one channel, with no patient in it.
 *
 * Unauthenticated on purpose, exactly as getDPDPNotice() is: this is the text
 * a patient is entitled to read BEFORE they are asked, and it contains
 * nothing but the clinic's own description of what it does.
 */
function getDispatchChannelNotice(channel) {
  crescEditorOnly_('getDispatchChannelNotice');
  var ch = dpdp_channel_(channel);
  if (!ch) {
    return { success: false, message: 'Unknown channel.', channels: dpdpDispatchChannels_() };
  }
  return { success: true, channel: ch.key, label: ch.label,
           processor: ch.processor, notice: ch.notice, message: '' };
}

/** FRONTEND ENTRY. Every channel and its notice, for a settings screen. */
function dpdpDispatchChannels_() {
  return Object.keys(DPDP_DISPATCH.CHANNELS).map(function (k) {
    var c = DPDP_DISPATCH.CHANNELS[k];
    return { key: c.key, label: c.label, processor: c.processor, notice: c.notice };
  });
}

/**
 * FRONTEND ENTRY. Where this patient stands on being sent documents, and the
 * notice to read them if they have not been asked.
 *
 * Called by a dispatch button BEFORE it sends, so the desk can put the
 * question rather than discovering the refusal after generating a PDF.
 *
 * @param {string} patientId
 * @param {string} channel
 * @param {string} sessionToken
 */
function getDispatchConsent(patientId, channel, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken,
                              ['billing.read', 'patient.read', 'lab.read',
                               'emr.read', 'portal.self']);
    var pid = String(patientId || '').trim().toUpperCase();
    if (actor.role === 'patient' &&
        String(actor.username || '').trim().toUpperCase() !== pid) {
      return { success: false, message: 'You can only see your own consent record.' };
    }

    var ch = dpdp_channel_(channel) || DPDP_DISPATCH.CHANNELS.WHATSAPP;
    var c = pid ? dpdp_consentState_(pid, DPDP_DISPATCH.PURPOSE)
                : { state: 'UNKNOWN_PATIENT', at: '', by: '', stale: false };

    return {
      success: true,
      patientId: pid,
      channel: ch.key,
      label: ch.label,
      processor: ch.processor,
      notice: ch.notice,
      state: c.state,
      maySend: c.state === 'GIVEN',
      recordedAt: c.at,
      recordedBy: c.by,
      staleNotice: !!c.stale,
      noticeVersion: dpdp_noticeVersion_(),
      message: ''
    };
  } catch (err) {
    return { success: false,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * FRONTEND ENTRY. Records the patient's answer at the moment of sending, and
 * says whether the send may now go ahead.
 *
 * This exists so that "the patient has not been asked" is answerable at the
 * counter in one click instead of being a dead end that teaches the desk to
 * route around consent. It writes through recordConsent(), so it lands in the
 * same append-only register with the same notice version as any other
 * consent — the only difference is the `method`, which records that it was
 * taken verbally at the desk rather than on a form.
 *
 * @param {{patientId, channel, agreed:boolean}} payload
 * @param {string} sessionToken
 */
function recordDispatchConsent(payload, sessionToken) {
  try {
    payload = payload || {};
    var pid = String(payload.patientId || '').trim().toUpperCase();
    if (!pid) return { success: false, message: 'No patient was named.' };

    var decisions = {};
    decisions[DPDP_DISPATCH.PURPOSE] = (payload.agreed === true);

    var res = recordConsent({
      patientId: pid,
      decisions: decisions,
      method: 'VERBAL_AT_DESK',
      notes: 'Asked at the point of sending a document by ' +
             ((dpdp_channel_(payload.channel) || {}).label || 'message') +
             '. The channel notice was read to the patient.'
    }, sessionToken);

    if (!res || !res.success) return res || { success: false, message: 'Not recorded.' };
    return { success: true, maySend: (payload.agreed === true),
             message: payload.agreed
               ? 'Consent recorded. Sending now.'
               : 'Recorded that they said no. Nothing will be sent this way.' };
  } catch (err) {
    return { success: false,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * ADMIN / REPORT. Which patients the clinic is currently unable to send
 * anything to, and why.
 *
 * The gate above turns a missing consent into a refusal at the counter, which
 * is correct and also invisible until somebody tries. This is the list to
 * work through beforehand.
 */
function dpdpDispatchReadiness(sessionToken) {
  try {
    crescRequire_(sessionToken, ['dpdp.manage', 'admin.config']);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
    if (!sh || sh.getLastRow() < 2) {
      return { success: true, rows: [], message: 'No patients on file.' };
    }
    var data = sh.getDataRange().getDisplayValues();
    var rows = [], counts = { GIVEN: 0, REFUSED: 0, WITHDRAWN: 0, NOT_ASKED: 0 };

    for (var i = 1; i < data.length; i++) {
      var pid = String(data[i][0] || '').trim().toUpperCase();
      if (!pid) continue;
      var c = dpdp_consentState_(pid, DPDP_DISPATCH.PURPOSE);
      counts[c.state] = (counts[c.state] || 0) + 1;
      if (c.state === 'GIVEN' && !c.stale) continue;
      rows.push({
        patientId: pid,
        name: String(data[i][2] || ''),
        mobile: String(data[i][6] || ''),
        state: c.state,
        staleNotice: !!c.stale,
        recordedAt: c.at
      });
    }
    return {
      success: true, rows: rows, counts: counts,
      message: rows.length
        ? rows.length + ' patient(s) cannot be sent documents until they are asked.'
        : 'Every patient on file has been asked.'
    };
  } catch (err) {
    return { success: false, rows: [],
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}
