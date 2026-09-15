// ============================================================================
// DS_Record_Bridge.gs — Crescentia HealthTech
// A signed discharge summary, in the two places clinicians look for it.
// ----------------------------------------------------------------------------
// WHERE A SIGNED SUMMARY WENT, AND WHERE IT DID NOT
//
// Signing wrote a snapshot to DS_Snapshots, stamped DS_Summaries and queued a
// PDF. Nothing else. Which means:
//
//   WARD NOTES (IP_Notes_UI, IP_Timeline_DB)
//     The notes timeline is the continuous record of a stay: admission,
//     nursing rounds, doctor's rounds, procedures — and then it simply ends.
//     The discharge, the document that concludes the admission and is the
//     only place the final diagnosis, the course, the discharge script and
//     the follow-up plan are written down together, left NO ENTRY AT ALL. A
//     nurse scrolling the stay could not see that the patient had been
//     discharged, let alone on what.
//
//   EMR MASTER TIMELINE (EMR_MasterTimeline_UI)
//     mt_pushAdmissions_ enriched each stay with six fields from
//     ipr_dischargeSummaryBlock_: the summary id, its status, who signed it,
//     the hash, the date and the final diagnosis. Useful, and not the
//     summary. A doctor seeing this patient for the first time, looking at
//     their last admission, got "Discharged · Type 2 DM with cellulitis ·
//     signed by Dr X · a1b2c3d4e5f6" — and no course, no medications, no
//     advice, no follow-up. The information exists, signed and hashed, one
//     sheet away.
//
// WHAT THIS FILE DOES
//
//   1. dsx_writeSummaryToTimeline_() — called at the end of signing. Writes
//      ONE note of type DISCHARGE_SUMMARY into IP_Timeline_DB carrying every
//      section of the signed payload, so the ward record concludes where the
//      stay concludes.
//
//      Idempotent per version: signing version 1 writes one note; signing an
//      amendment as version 2 writes a second, marked as an amendment, and
//      leaves the first standing. That is deliberate — the ward record must
//      show that the summary changed and when, which a mutated note cannot.
//
//   2. getDischargeSummaryFull() — every section of the signed summary for
//      one admission, as data, for the master timeline to expand inline.
//
//   3. mtPrintEncounter() — print HTML for ONE encounter: a single OPD visit,
//      or a single admission with its whole discharge summary. Every event on
//      the master timeline is individually printable, which is what a patient
//      asking for "the paper from my visit in March" actually needs.
//
// WHY THE TIMELINE STILL LOADS LIGHT
//
// getDischargeSummaryFull is a SEPARATE call, made when a stay is expanded,
// rather than being folded into buildLongitudinalTimeline. A patient with six
// admissions would otherwise carry six complete discharge summaries through
// google.script.run on every timeline open, to display one of them.
// ============================================================================

/** The note type the ward timeline stores a discharge summary under. */
var DSB_NOTE_TYPE = 'DISCHARGE_SUMMARY';

/**
 * Sections in the order a reader wants them, with the labels the printed
 * document uses. Drawn from the payload rather than assumed, so a section the
 * assembler adds later appears without touching this list — DSX_PRINT_ORDER
 * governs the order and anything outside it is appended.
 */
function dsb_orderedSections_(payload) {
  var out = [];
  var seen = {};
  var order = (typeof DSX_PRINT_ORDER !== 'undefined' && DSX_PRINT_ORDER)
    ? DSX_PRINT_ORDER : [];

  var push = function (key) {
    if (seen[key]) return;
    var sec = payload.sections[key];
    if (!sec) return;
    seen[key] = true;
    out.push({
      key: key,
      title: dsx_str_(sec.title) || key,
      format: dsx_str_(sec.format) || 'TEXT',
      content: sec.content,
      empty: (typeof dsx_sectionIsEmpty_ === 'function') ? dsx_sectionIsEmpty_(sec) : false
    });
  };

  order.forEach(push);
  Object.keys(payload.sections || {}).forEach(push);
  return out;
}

/**
 * Renders one section's content as PLAIN TEXT, for the ward note.
 *
 * Deliberately not the print HTML. A note in IP_Timeline_DB is read on a ward
 * screen, in a search, and in the complete-file print, none of which want a
 * nested document's markup; and storing markup in a JSON note column is how a
 * record stops being greppable.
 */
function dsb_sectionText_(sec) {
  var c = sec.content;
  if (c === null || c === undefined) return '';

  if (sec.format === 'LIST' && Object.prototype.toString.call(c) === '[object Array]') {
    return c.filter(String).map(function (l) { return '• ' + dsx_str_(l); }).join('\n');
  }

  if (sec.format === 'TABLE' && c && c.columns) {
    var cols = c.columns || [];
    var rows = c.rows || [];
    if (!rows.length) return '';
    return rows.map(function (r) {
      // "Medicine Name: Tab Amoxicillin 500mg · Dosage / Sig: 1-0-1 · Days: 5"
      // reads on one line in a note column and survives a search for the
      // drug name, which a rendered table does not.
      return cols.map(function (h, i) {
        var v = dsx_str_(r[i]);
        return v ? (h + ': ' + v) : '';
      }).filter(String).join(' · ');
    }).join('\n');
  }

  if (sec.format === 'FIELDS' && typeof c === 'object') {
    return Object.keys(c).map(function (f) {
      var v = dsx_str_(c[f]);
      if (!v) return '';
      var label = (typeof dsx_humanise_ === 'function') ? dsx_humanise_(f) : f;
      return label + ': ' + v;
    }).filter(String).join('\n');
  }

  return dsx_str_(c);
}

/**
 * Writes the signed summary into the ward notes timeline.
 *
 * NEVER THROWS. Its caller is the signing commit, and a signature that has
 * already been written must not be rolled back because a note could not be
 * appended — the summary is safe in DS_Snapshots either way. A failure is
 * logged and reported on the sign reply so it is visible rather than silent.
 *
 * @param {string} summaryId
 * @param {Object} header      the DS_Summaries row
 * @param {Object} payload     the SIGNED payload
 * @param {number} snapshotNo
 * @param {Object} actor       the signer
 * @return {{ok:boolean, noteId:string, message:string}}
 */
function dsx_writeSummaryToTimeline_(summaryId, header, payload, snapshotNo, actor) {
  try {
    var ipNumber = dsx_upper_(header.IP_Number);
    if (!ipNumber) return { ok: false, noteId: '', message: 'the summary has no IP number' };

    var banner = (payload.sections && payload.sections.PATIENT_BANNER &&
                  payload.sections.PATIENT_BANNER.content) || {};
    var patientId = dsx_upper_(banner.patientId);

    var sheet = (typeof ipc_timelineSheet_ === 'function')
      ? ipc_timelineSheet_()
      : SpreadsheetApp.getActiveSpreadsheet().getSheetByName('IP_Timeline_DB');
    if (!sheet) return { ok: false, noteId: '', message: 'IP_Timeline_DB is missing' };

    var m = dc_headerMap_(sheet);

    // One note per signed VERSION. Re-running the signing commit — a retry, a
    // resumed execution — must not produce two identical entries; an
    // amendment signed as the next version must.
    var marker = summaryId + '#v' + snapshotNo;
    try {
      var existing = dc_sheetValues_(sheet);
      for (var i = 1; i < (existing ? existing.length : 0); i++) {
        if (dsx_upper_(existing[i][m['Role_Type']]) !== DSB_NOTE_TYPE) continue;
        if (String(existing[i][m['Note_Data_JSON']] || '').indexOf(marker) === -1) continue;
        return { ok: true, noteId: dsx_str_(existing[i][m['Note_ID']]),
                 message: 'already recorded' };
      }
    } catch (e) { /* a failed dedupe check must not stop the write */ }

    var sections = dsb_orderedSections_(payload).filter(function (s) { return !s.empty; });

    var noteData = {
      // The marker the dedupe above looks for, and the link back to the
      // signed artefact a reader can verify.
      summaryRef: marker,
      summaryId: summaryId,
      version: snapshotNo,
      isAmendment: snapshotNo > 1,
      documentTitle: dsx_str_(payload.documentTitle) ||
                     (typeof dsx_documentTitle_ === 'function'
                       ? dsx_documentTitle_(header.Discharge_Type) : 'Discharge summary'),
      dischargeType: dsx_upper_(header.Discharge_Type) || 'NORMAL',
      signedAt: dsx_fmt_(header.Signed_At || new Date(), 'dd-MMM-yyyy hh:mm a'),
      signedBy: dsx_str_(actor && actor.displayName) || dsx_str_(header.Signed_By),
      signerRegNo: dsx_str_(header.Signer_Reg_No),
      shortHash: (typeof dsx_shortHash_ === 'function')
                   ? dsx_shortHash_(header.Signed_Hash) : '',
      finalDiagnosis: (typeof dsx_finalDiagnosisOf_ === 'function')
                        ? dsx_finalDiagnosisOf_(payload) : '',
      // The WHOLE summary, section by section, as text.
      sections: sections.map(function (s) {
        return { key: s.key, title: s.title, text: dsb_sectionText_(s) };
      }),
      // A one-line summary for a collapsed timeline row.
      summary: 'Discharge summary v' + snapshotNo + ' signed' +
               (snapshotNo > 1 ? ' (amendment)' : '') + '.'
    };

    var noteId = (typeof _generateNoteId_ === 'function')
      ? _generateNoteId_('IPN')
      : 'IPN-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    var row = new Array(sheet.getLastColumn()).fill('');
    var put = function (h, v) { if (m[h] !== undefined) row[m[h]] = v; };
    put('Timestamp', new Date());
    put('IP_Number', ipNumber);
    put('Patient_ID', patientId);
    put('Role_Type', DSB_NOTE_TYPE);
    put('Note_Data_JSON', JSON.stringify(noteData));
    put('Author', dsx_str_(actor && actor.displayName) || dsx_str_(header.Signed_By));
    put('Shift', (typeof _getCurrentShift_ === 'function') ? _getCurrentShift_() : '');
    put('Flags', snapshotNo > 1 ? 'AMENDMENT' : '');
    put('Note_ID', noteId);
    put('Author_Doctor_ID', dsx_str_(actor && actor.doctorId));
    put('Author_Username', dsx_str_(actor && actor.username));
    sheet.appendRow(row);
    dc_invalidate_('IP_Timeline_DB');

    return { ok: true, noteId: noteId, message: '' };

  } catch (err) {
    try { Logger.log('dsx_writeSummaryToTimeline_: ' + err.message); } catch (e) {}
    return { ok: false, noteId: '', message: String(err && err.message || err) };
  }
}

/**
 * ADMIN / REPAIR. Writes the ward-note entry for summaries signed BEFORE this
 * bridge existed.
 *
 * Every stay discharged up to now has no discharge entry in its ward record.
 * Idempotent, so it is safe to re-run, and it reports per summary.
 */
function dsxBackfillSummaryNotes(sessionToken) {
  try {
    var actor = (typeof dsx_requireRole_ === 'function')
      ? dsx_requireRole_(sessionToken, 'printFinal')
      : crescRequire_(sessionToken, ['ward.discharge', 'admin.config']);

    var sh = dsx_summariesSheet_();
    var data = dc_sheetValues_(sh);
    var m = dc_headerMap_(sh);
    if (!data || data.length < 2) return { success: true, written: 0, message: 'No summaries.' };

    var written = 0, already = 0, failed = [];
    for (var i = 1; i < data.length; i++) {
      if (dsx_upper_(data[i][m['Status']]) !== DSX_STATUS.SIGNED) continue;
      var summaryId = dsx_str_(data[i][m['Summary_ID']]);
      var header = dsx_getHeader_(summaryId);
      if (!header) continue;

      var snapNo = dsx_int_(header.Last_Signed_Snapshot_No);
      var ref = dsx_resolveRef_(summaryId, 'SIGNED:' + snapNo);
      if (!ref || !ref.payload) { failed.push(summaryId + ' (no readable snapshot)'); continue; }

      var res = dsx_writeSummaryToTimeline_(summaryId, header, ref.payload, snapNo, actor);
      if (!res.ok) failed.push(summaryId + ' (' + res.message + ')');
      else if (res.message === 'already recorded') already++;
      else written++;
    }
    SpreadsheetApp.flush();

    return {
      success: true, written: written,
      message: written + ' discharge summary(ies) added to the ward record, ' +
               already + ' already there.' +
               (failed.length ? '\n\nCould not do: ' + failed.join('; ') : '')
    };
  } catch (err) {
    return { success: false, written: 0,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}

// ---------------------------------------------------------------------------
// THE WHOLE SUMMARY, FOR THE MASTER TIMELINE
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. Every section of the signed summary for one admission.
 *
 * Reads the SIGNED snapshot, never the working row: a summary that has been
 * signed and then edited must not show its edits here, for the same reason
 * ds_getPrintHtml refuses to print them.
 *
 * @param {string} ipNumber
 * @param {string} sessionToken
 * @return {{success, ipNumber, summaryId, status, version, signedAt, signedBy,
 *           signerRegNo, shortHash, sections:Array, message}}
 */
function getDischargeSummaryFull(ipNumber, sessionToken) {
  try {
    crescRequire_(sessionToken, ['emr.read', 'ward.read']);

    var summaryId = (typeof dsx_summaryIdFor_ === 'function')
      ? dsx_summaryIdFor_(dsx_upper_(ipNumber)) : '';
    var header = summaryId ? dsx_getHeader_(summaryId) : null;
    if (!header) {
      return { success: false, sections: [],
               message: 'No discharge summary has been started for ' + ipNumber + '.' };
    }

    var status = dsx_upper_(header.Status);
    var signed = (status === DSX_STATUS.SIGNED);
    var snapNo = signed ? dsx_int_(header.Last_Signed_Snapshot_No)
                        : dsx_int_(header.Current_Snapshot_No);

    var ref = dsx_resolveRef_(summaryId, signed ? ('SIGNED:' + snapNo) : 'WORKING');
    if (!ref || !ref.payload) {
      return { success: false, sections: [],
               message: 'The summary for ' + ipNumber + ' could not be read.' };
    }

    try {
      dpdpLogRead_(crescActor_(sessionToken), 'Patient',
                   dsx_upper_((ref.payload.sections.PATIENT_BANNER || {}).content &&
                              ref.payload.sections.PATIENT_BANNER.content.patientId),
                   { endpoint: 'getDischargeSummaryFull', summaryId: summaryId });
    } catch (e) {}

    var sections = dsb_orderedSections_(ref.payload)
      .filter(function (s) { return !s.empty || s.key === 'ALLERGIES'; })
      .map(function (s) {
        return { key: s.key, title: s.title, format: s.format,
                 text: dsb_sectionText_(s) };
      });

    return {
      success: true,
      ipNumber: dsx_upper_(ipNumber),
      summaryId: summaryId,
      status: status,
      // A draft is returned, clearly labelled, rather than withheld: a
      // clinician looking at an unfinished summary needs to see that it is
      // unfinished, not an empty panel that looks like a fault.
      signed: signed,
      version: snapNo,
      signedAt: dsx_fmt_(header.Signed_At, 'dd-MMM-yyyy hh:mm a'),
      signedBy: dsx_str_(header.Signed_By),
      signerRegNo: dsx_str_(header.Signer_Reg_No),
      shortHash: (typeof dsx_shortHash_ === 'function')
                   ? dsx_shortHash_(header.Signed_Hash) : '',
      dischargeType: dsx_upper_(header.Discharge_Type) || 'NORMAL',
      sections: sections,
      message: signed ? '' :
        'This summary is ' + status.replace(/_/g, ' ').toLowerCase() +
        ' — it has NOT been signed, and what you are reading may still change.'
    };
  } catch (err) {
    return { success: false, sections: [],
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}

// ---------------------------------------------------------------------------
// ONE ENCOUNTER, ON PAPER
// ---------------------------------------------------------------------------
//
// Nothing on the master timeline was printable on its own. The whole record
// could be printed as a file, and a prescription could be printed from the OP
// consult while it was still open — but a patient at the counter asking for
// "the paper from my visit in March", or a doctor wanting one admission's
// summary to send with a referral, had no route to it. The information is on
// screen; it just could not leave the screen.
//
// mtPrintOPEncounter() renders ONE consultation as the document it is:
// vitals, complaints, history, diagnosis, the prescription, the
// investigations, the advice and the review date, on the clinic's letterhead
// through the same print kit the discharge summary and the ward notes use —
// so the three do not arrive as paper from three different organisations.
//
// An admission does not need its own renderer: the signed discharge summary
// IS that document, and ds_getPrintHtml already produces it from the signed
// snapshot. The timeline calls that.

/**
 * FRONTEND ENTRY. One OPD consultation, as a printable document.
 *
 * @param {string} encounterId
 * @param {string} sessionToken
 * @return {{success:boolean, html:string, message:string}}
 */
function mtPrintOPEncounter(encounterId, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['emr.read', 'patient.read']);

    var res = getEncounterForPrint(String(encounterId || '').trim());
    if (!res || !res.success) {
      return { success: false, html: '',
               message: (res && res.message) || 'That consultation could not be read.' };
    }
    var d = res.data;

    try {
      dpdpLogRead_(actor, 'Patient', String(d.patientId || ''),
                   { endpoint: 'mtPrintOPEncounter', encounterId: d.encounterId });
    } catch (e) {}

    var body = [];

    // ---- vitals ----------------------------------------------------------
    var v = d.vitals || {};
    var bp = [String(v.sysBp || '').trim(), String(v.diaBp || '').trim()]
               .filter(String).join('/');
    var vitalRows = [
      ['BP (mmHg)', bp],
      ['Pulse (/min)', v.hr],
      ['SpO2 (%)', v.spo2],
      ['Temperature', v.temp],
      ['Weight (kg)', v.weight],
      ['BMI', v.bmi]
    ].filter(function (r) { return String(r[1] || '').trim(); });
    if (vitalRows.length) body.push(ipp_sec_('Vitals', ipp_kv_(vitalRows)));

    // ---- clinical --------------------------------------------------------
    var c = d.clinical || {};
    if (String(c.complaints || '').trim()) {
      body.push(ipp_sec_('Presenting complaints',
        '<div style="font-size:10pt;">' + ipp_escMultiline_(c.complaints) + '</div>'));
    }
    if (String(c.history || '').trim()) {
      body.push(ipp_sec_('History',
        '<div style="font-size:10pt;">' + ipp_escMultiline_(c.history) + '</div>'));
    }
    if (String(c.diagnosis || '').trim()) {
      body.push(ipp_sec_('Diagnosis',
        '<div style="font-size:11pt;font-weight:700;">' +
        ipp_escMultiline_(c.diagnosis) + '</div>'));
    }

    // ---- prescription ----------------------------------------------------
    // The same five columns the OP consult, the IP case sheet and the
    // discharge script prescribe in, so a patient holding two of them is not
    // reading two different layouts of the same instruction.
    var meds = Array.isArray(d.meds) ? d.meds : [];
    if (meds.length) {
      body.push(ipp_sec_('Prescription', ipp_table_(
        ['#', 'Medicine', 'Dosage / Sig', 'Days', 'Notes / Timing'],
        meds.map(function (m, i) {
          var sig = String(m.sig || '').trim();
          // An IV order's rate is half the instruction; without it the sheet
          // shows a volume apparently infused at no particular speed.
          if (String(m.type || '').toUpperCase() === 'IV' && String(m.rate || '').trim()) {
            sig = (sig + ' @ ' + String(m.rate).trim()).trim();
          }
          return [
            String(i + 1),
            '<strong>' + ipp_esc_(m.drugName || m.brand || '') + '</strong>' +
              (m.generic ? '<br><span style="font-size:8pt;color:#64748b;">' +
                           ipp_esc_(m.generic) + '</span>' : ''),
            ipp_esc_(sig),
            ipp_esc_(m.duration || m.days || ''),
            ipp_esc_(m.instructions || m.notes || '')
          ];
        }),
        ['6%', '32%', '24%', '10%', '28%'])));
    }

    // ---- investigations --------------------------------------------------
    var labs = Array.isArray(d.labs) ? d.labs : [];
    if (labs.length) {
      body.push(ipp_sec_('Investigations', ipp_table_(
        ['Test', 'Result', 'Reference'],
        labs.map(function (l) {
          return [ipp_esc_(l.testName || ''),
                  ipp_esc_((l.result || '') + (l.unit ? ' ' + l.unit : '')),
                  ipp_esc_(l.referenceRange || '')];
        }), ['50%', '25%', '25%'])));
    }

    // ---- advice and review ----------------------------------------------
    var tail = [];
    if (String(d.advice || '').trim()) tail.push(['Advice', ipp_escMultiline_(d.advice)]);
    if (String(d.reviewDate || '').trim()) tail.push(['Review on', ipp_esc_(d.reviewDate)]);
    if (tail.length) body.push(ipp_sec_('Advice and follow-up', ipp_kv_(tail)));

    // ---- signature -------------------------------------------------------
    body.push(ipp_sec_('', ipp_cols_('',
      ipp_sig_(d.doctorName || '',
               [d.doctorSpecialty, d.doctorRegNo ? 'Reg. ' + d.doctorRegNo : '']
                 .filter(String).join(' · ')))));

    var html = ipp_doc_({
      docTitle: 'OUTPATIENT CONSULTATION',
      patient: {
        name: d.patientName,
        pid: d.patientId,
        ageSex: d.patientAgeSex,
        consultant: d.doctorName,
        diagnosis: c.diagnosis
      },
      bodyHtml: body.join(''),
      footNote: 'Consultation ' + String(d.encounterId || '') + ' · ' +
                String(d.date || '')
    });

    return { success: true, html: html, message: '' };

  } catch (err) {
    return { success: false, html: '',
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}
