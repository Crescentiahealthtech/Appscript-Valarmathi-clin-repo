// ============================================================================
// Patient_Portal.gs — Crescentia HealthTech
// What the patient's own app is allowed to know, and where it comes from.
// ----------------------------------------------------------------------------
// WHAT THE PORTAL WAS
//
// PatientApp.html was written once and left. Three of its four panels were
// HARDCODED MARKUP — not stale data, not a failed fetch, literally typed into
// the file:
//
//     Metformin 500mg  ·  Telmisartan 40mg  ·  Atorvastatin 10mg
//     Height 165 cm  Weight 72 kg  BMI 26.4  HbA1c 8.2%  BP 130/85
//     hba1cChart: data: [7.2, 7.5, 7.8, 8.0, 8.2]
//
// Every patient who signed in saw the same three drugs, the same weight and
// the same rising HbA1c. A patient who has never been prescribed metformin
// was being shown metformin. "Investigations & Prescriptions" and "Order
// Medications" were single paragraphs reading "mapped here."
//
// The booking screen worked, but booked against getAvailableTimeSlots — one
// global grid of quarter-hours with no doctor in it at all, so a patient
// could not choose whom they were seeing, and a slot showed as free when the
// doctor they had in mind was on leave.
//
// WHAT IT IS NOW
//
// One call, portalHome(), returns the patient's actual record: their next
// appointment and the review date their doctor set, what they are actually
// taking, their real last vitals, their released lab results, and whichever
// condition panels apply to them — diabetes, hypertension, pregnancy,
// childhood immunisation. Nothing is invented; a panel with no data says so
// rather than showing a plausible number.
//
// TWO RULES THIS FILE KEEPS THROUGHOUT
//
// 1. THE PATIENT ID COMES FROM THE SESSION, NEVER FROM THE BROWSER.
//    Every other portal endpoint in this project takes a patientId argument
//    and then checks it against the session (crescRequireOwnRecord_), which
//    works but means the check can be forgotten in the next endpoint somebody
//    adds. Here the id is READ OFF THE SESSION and the argument does not
//    exist, so there is nothing to forge and nothing to remember to check.
//    Patient ids are sequential and printed on every barcode label, so this
//    matters more here than it would elsewhere.
//
// 2. NOTHING UNVERIFIED REACHES THE PATIENT.
//    A lab result that has not been verified, a discharge summary that has
//    not been signed, a draft prescription — none of them appear. A patient
//    acting on a number no clinician has released is the specific harm a
//    portal can cause that a paper report cannot.
// ============================================================================

var PP = {
  /** How many past appointments and lab orders the portal carries. */
  VISIT_LIMIT: 12,
  LAB_LIMIT: 20,

  /**
   * Parameter-name aliases, because a lab catalogue is typed by humans.
   *
   * LAB_TEST_CATALOG has no code a trend could key on — the same analyte is
   * "HbA1c" in one clinic's catalogue and "Glycated Haemoglobin" in the next,
   * and this deployment's has been edited by hand. Matching on the name with
   * a list of spellings is how the chart finds the values instead of showing
   * an empty panel to a diabetic with three years of results.
   *
   * Only the analytes something on screen actually plots are listed. Adding a
   * key here does nothing on its own; it needs a series and a panel.
   */
  ANALYTES: {
    HBA1C: ['HBA1C', 'HB A1C', 'GLYCATED HAEMOGLOBIN', 'GLYCOSYLATED HEMOGLOBIN',
            'GLYCATED HEMOGLOBIN', 'A1C'],
    FBS:   ['FBS', 'FASTING BLOOD SUGAR', 'FASTING GLUCOSE', 'GLUCOSE FASTING',
            'BLOOD SUGAR FASTING', 'FASTING PLASMA GLUCOSE'],
    PPBS:  ['PPBS', 'PBS', 'POST PRANDIAL BLOOD SUGAR', 'POSTPRANDIAL GLUCOSE',
            'POST PRANDIAL GLUCOSE', 'GLUCOSE POST PRANDIAL', 'PP BLOOD SUGAR',
            'POST PRANDIAL PLASMA GLUCOSE'],
    RBS:   ['RBS', 'RANDOM BLOOD SUGAR', 'RANDOM GLUCOSE']
  }
};

function pp_str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function pp_up_(v) { return pp_str_(v).toUpperCase(); }
function pp_num_(v) {
  // "8.2 %" and "110 mg/dL" both arrive as result values.
  var n = parseFloat(String(v === null || v === undefined ? '' : v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? null : n;
}
function pp_date_(v) {
  if (!v) return null;
  if (typeof cresc_parseDate_ === 'function') return cresc_parseDate_(v);
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var d = new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}
function pp_fmt_(v, pattern) {
  var d = pp_date_(v);
  if (!d) return '';
  try { return Utilities.formatDate(d, 'Asia/Kolkata', pattern || 'dd MMM yyyy'); }
  catch (e) { return ''; }
}
function pp_iso_(v) {
  var d = pp_date_(v);
  return d ? Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd') : '';
}

/**
 * THE ONE AUTHORISATION IN THIS FILE.
 *
 * Returns the signed-in patient's own id, or throws. Staff are refused
 * outright rather than shown the portal for whoever they name: a receptionist
 * wanting a patient's record has the patient screens, which audit the read.
 *
 * @return {{patientId:string, username:string}}
 */
function pp_me_(sessionToken) {
  var sess = null;
  try { sess = dc_validateSession_(sessionToken); } catch (e) { sess = null; }
  if (!sess) throw new Error('Your session has expired. Please sign in again.');

  var role = pp_str_(sess.role).toLowerCase();
  if (role !== 'patient') {
    throw new Error('These screens are the patient\'s own. Staff should open the ' +
                    'patient record from the clinical screens, where the read is ' +
                    'logged against the patient.');
  }
  var pid = pp_up_(sess.username);
  if (!pid) throw new Error('This login is not linked to a patient record.');
  return { patientId: pid, username: pp_str_(sess.username) };
}

/** One error shape for every endpoint here. */
function pp_fail_(err, extra) {
  var out = { success: false,
              message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  if (extra) Object.keys(extra).forEach(function (k) { out[k] = extra[k]; });
  return out;
}

/** Does a lab parameter name mean this analyte? */
function pp_isAnalyte_(parameterName, analyte) {
  var n = pp_up_(parameterName).replace(/[^A-Z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  var list = PP.ANALYTES[analyte] || [];
  for (var i = 0; i < list.length; i++) {
    if (n === list[i]) return true;
    // A catalogue entry is often "HbA1c (Glycated Haemoglobin)".
    if (n.indexOf(list[i]) === 0) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------
// APPOINTMENTS
// ---------------------------------------------------------------------------

/**
 * This patient's appointments, split into what is still to come and what has
 * been.
 *
 * Reads the Appointments sheet by header where the columns were appended, and
 * by index for the original nine — which is how every other reader in this
 * project treats that sheet.
 */
function pp_appointments_(pid) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
  var out = { upcoming: [], past: [], lastVisit: '' };
  if (!sh || sh.getLastRow() < 2) return out;

  var m = dc_headerMap_(sh);
  var docCol = (m['Doctor_ID'] === undefined) ? -1 : m['Doctor_ID'];
  var nameCol = (m['Doctor_Name_Snapshot'] === undefined) ? -1 : m['Doctor_Name_Snapshot'];
  var data = sh.getDataRange().getDisplayValues();

  var todayIso = pp_iso_(new Date());
  var LIVE = ['Booked', 'Arrived', 'In-Progress'];

  for (var i = 1; i < data.length; i++) {
    if (pp_up_(data[i][1]) !== pid) continue;
    var status = pp_str_(data[i][6]);
    if (status === 'Blocked' || status === 'DELETE') continue;

    var dateIso = (typeof dc_dateKey_ === 'function')
      ? dc_dateKey_(data[i][3]) : pp_iso_(data[i][3]);
    var docId = (docCol === -1) ? '' : pp_str_(data[i][docCol]);
    var docName = (nameCol === -1) ? '' : pp_str_(data[i][nameCol]);
    if (!docName && docId) {
      try {
        var prof = dc_getDoctorById_(docId);
        if (prof) docName = prof.name;
      } catch (e) {}
    }

    var rec = {
      apptId: pp_str_(data[i][0]),
      date: dateIso,
      dateText: pp_fmt_(dateIso, 'EEE, dd MMM yyyy'),
      time: pp_str_(data[i][4]),
      purpose: pp_str_(data[i][5]),
      status: status,
      doctorId: docId,
      doctorName: docName || 'the clinic'
    };

    if (dateIso >= todayIso && LIVE.indexOf(status) !== -1) out.upcoming.push(rec);
    else out.past.push(rec);
  }

  out.upcoming.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
  out.past.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });

  // "Last visit" is a visit that HAPPENED. A cancellation is not a visit, and
  // showing one as the last visit tells the patient they were seen when they
  // were not.
  for (var p = 0; p < out.past.length; p++) {
    if (out.past[p].status === 'Completed' || out.past[p].status === 'Arrived' ||
        out.past[p].status === 'In-Progress') {
      out.lastVisit = out.past[p].dateText;
      break;
    }
  }
  out.past = out.past.slice(0, PP.VISIT_LIMIT);
  return out;
}

// ---------------------------------------------------------------------------
// CONSULTATIONS: vitals, prescriptions, the review date
// ---------------------------------------------------------------------------

/**
 * Every OP encounter for this patient, newest first, with the fields the
 * portal shows.
 *
 * OP_Encounters is written positionally by saveOPEncounter (A encounter,
 * B patient, C timestamp, D–J vitals, K complaints, L history, V diagnosis,
 * W meds JSON, X labs JSON, Y advice, Z review date), so it is read the same
 * way, with the appended Doctor_ID read by header name.
 */
function pp_encounters_(pid) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OP_Encounters');
  if (!sh || sh.getLastRow() < 2) return [];

  var data = sh.getDataRange().getDisplayValues();
  var hdr = {};
  (data[0] || []).forEach(function (h, c) { hdr[pp_str_(h)] = c; });
  var out = [];

  for (var i = 1; i < data.length; i++) {
    if (pp_up_(data[i][1]) !== pid) continue;
    var when = pp_date_(data[i][2]);
    var meds = [], labs = [];
    try { meds = data[i][22] ? JSON.parse(data[i][22]) : []; } catch (e) { meds = []; }
    try { labs = data[i][23] ? JSON.parse(data[i][23]) : []; } catch (e) { labs = []; }

    var docId = (hdr['Doctor_ID'] !== undefined) ? pp_str_(data[i][hdr['Doctor_ID']]) : '';
    var docName = (hdr['Doctor_Signature_Snapshot'] !== undefined)
      ? pp_str_(data[i][hdr['Doctor_Signature_Snapshot']]) : '';
    if (!docName && docId) {
      try { var pr = dc_getDoctorById_(docId); if (pr) docName = pr.name; } catch (e) {}
    }

    out.push({
      encounterId: pp_str_(data[i][0]),
      at: when,
      date: pp_iso_(when),
      dateText: pp_fmt_(when, 'dd MMM yyyy'),
      doctorName: docName,
      vitals: {
        sysBp: pp_num_(data[i][3]), diaBp: pp_num_(data[i][4]),
        pulse: pp_num_(data[i][5]), spo2: pp_num_(data[i][6]),
        temp: pp_str_(data[i][7]), weight: pp_num_(data[i][8]),
        bmi: pp_num_(data[i][9])
      },
      complaints: pp_str_(data[i][10]),
      diagnosis: pp_str_(data[i][21]),
      meds: Array.isArray(meds) ? meds : [],
      labs: Array.isArray(labs) ? labs : [],
      advice: pp_str_(data[i][24]),
      reviewDate: pp_str_(data[i][25])
    });
  }
  out.sort(function (a, b) { return (b.at || 0) - (a.at || 0); });
  return out;
}

/**
 * What the patient is currently taking.
 *
 * The LATEST prescription wins outright — the one from the most recent
 * consultation or the most recent signed discharge summary, whichever is
 * newer. It is NOT merged with older ones: a drug stopped at the last visit
 * would reappear from the visit before, and a portal that tells a patient to
 * keep taking something they were told to stop is worse than one that shows
 * nothing.
 *
 * `asOf` and `source` ride along so the screen can say where the list came
 * from and how old it is, which is the honest way to show a list that may
 * have been superseded by a conversation.
 */
function pp_currentMeds_(pid, encounters) {
  var latest = { asOf: null, source: '', prescriber: '', items: [] };

  var enc = (encounters && encounters.length) ? encounters[0] : null;
  if (enc && enc.meds.length) {
    latest = {
      asOf: enc.at,
      source: 'Consultation on ' + enc.dateText,
      prescriber: enc.doctorName,
      items: enc.meds.map(function (m) {
        var sig = pp_str_(m.sig);
        if (pp_up_(m.type) === 'IV' && pp_str_(m.rate)) {
          sig = (sig + ' @ ' + pp_str_(m.rate)).trim();
        }
        return {
          name: pp_str_(m.drugName || m.brand),
          generic: pp_str_(m.generic),
          type: pp_str_(m.type),
          sig: sig,
          days: pp_str_(m.duration || m.days),
          notes: pp_str_(m.instructions || m.notes)
        };
      })
    };
  }

  // A discharge script is the more recent instruction when the stay ended
  // after the last outpatient visit.
  try {
    var ds = pp_latestDischargeScript_(pid);
    if (ds && (!latest.asOf || (ds.asOf && ds.asOf > latest.asOf))) latest = ds;
  } catch (e) { /* the discharge module may not be on this deployment */ }

  return latest;
}

/**
 * The discharge medications from this patient's most recent SIGNED summary.
 *
 * Signed only. An unsigned summary is a draft that a doctor is still writing,
 * and its medication table is the section the desk is told never to prefill
 * — showing it to the patient would hand them a prescription nobody has
 * approved.
 */
function pp_latestDischargeScript_(pid) {
  if (typeof dsx_summariesSheet_ !== 'function') return null;
  var sh = dsx_summariesSheet_();
  var m = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  if (!data || data.length < 2) return null;

  var best = null;
  for (var i = 1; i < data.length; i++) {
    if (dsx_upper_(data[i][m['Status']]) !== DSX_STATUS.SIGNED) continue;
    var ip = dsx_str_(data[i][m['IP_Number']]);
    if (!ip) continue;
    var signedAt = pp_date_(data[i][m['Signed_At']]);
    if (!signedAt) continue;
    if (best && best.signedAt >= signedAt) continue;
    best = { summaryId: dsx_str_(data[i][m['Summary_ID']]), ipNumber: ip,
             signedAt: signedAt,
             snapNo: dsx_int_(data[i][m['Last_Signed_Snapshot_No']]) };
  }
  if (!best) return null;

  // Cheapest possible confirmation that this stay is THIS patient's: the
  // banner in the signed payload carries the patient id.
  var ref = dsx_resolveRef_(best.summaryId, 'SIGNED:' + best.snapNo);
  if (!ref || !ref.payload) return null;
  var banner = (ref.payload.sections.PATIENT_BANNER || {}).content || {};
  if (dsx_upper_(banner.patientId) !== pid) return null;

  var sec = ref.payload.sections.DISCHARGE_MEDICATIONS;
  if (!sec || !sec.content || !sec.content.rows || !sec.content.rows.length) return null;

  var cols = sec.content.columns || [];
  var idx = function (name) { return cols.indexOf(name); };
  var iName = idx('Medicine Name'), iSig = idx('Dosage / Sig'),
      iDays = idx('Days'), iType = idx('Type'), iNotes = idx('Notes / Timing');

  return {
    asOf: best.signedAt,
    source: 'Discharge on ' + pp_fmt_(best.signedAt) + ' (' + best.ipNumber + ')',
    prescriber: dsx_str_(ref.payload.signedBy) || '',
    items: sec.content.rows.map(function (r) {
      return {
        name: (iName > -1) ? pp_str_(r[iName]) : '',
        generic: '',
        type: (iType > -1) ? pp_str_(r[iType]) : '',
        sig: (iSig > -1) ? pp_str_(r[iSig]) : '',
        days: (iDays > -1) ? pp_str_(r[iDays]) : '',
        notes: (iNotes > -1) ? pp_str_(r[iNotes]) : ''
      };
    }).filter(function (x) { return x.name; })
  };
}

/**
 * The review date the doctor actually set, from the most recent consultation
 * that set one.
 *
 * The old portal parsed this out of the free-text PLAN field of EMR_Records
 * with a line-by-line keyword search for "follow-up" or "next visit", and
 * showed whatever was left of the line. OP_Encounters column Z holds the date
 * the doctor picked in the consult, which is the answer to the question.
 */
function pp_reviewDate_(encounters) {
  for (var i = 0; i < encounters.length; i++) {
    var raw = pp_str_(encounters[i].reviewDate);
    if (!raw) continue;
    var iso = pp_iso_(raw);
    return {
      // An unparseable value is shown as written rather than dropped: "after
      // 2 weeks" is still what the doctor said.
      date: iso,
      text: iso ? pp_fmt_(iso, 'EEEE, dd MMM yyyy') : raw,
      setOn: encounters[i].dateText,
      setBy: encounters[i].doctorName,
      overdue: !!(iso && iso < pp_iso_(new Date()))
    };
  }
  return null;
}

/**
 * The most recent vitals, from wherever they were most recently taken.
 *
 * An outpatient consultation and an antenatal visit both record them, and for
 * a patient who was recently an inpatient the ward notes are newer than
 * either. All three are considered and the newest wins, because "last vitals"
 * that silently mean "last OUTPATIENT vitals" are wrong for exactly the
 * patient whose vitals matter most.
 */
function pp_lastVitals_(pid, encounters, antenatal) {
  var best = null;
  var consider = function (at, source, v) {
    if (!at) return;
    if (best && best.at >= at) return;
    best = { at: at, source: source, v: v };
  };

  if (encounters.length) {
    var e = encounters[0];
    consider(e.at, 'Consultation on ' + e.dateText, {
      bp: (e.vitals.sysBp && e.vitals.diaBp) ? (e.vitals.sysBp + '/' + e.vitals.diaBp) : '',
      pulse: e.vitals.pulse, spo2: e.vitals.spo2, temp: e.vitals.temp,
      weight: e.vitals.weight, bmi: e.vitals.bmi
    });
  }

  if (antenatal && antenatal.lastVisit) {
    var lv = antenatal.lastVisit;
    var lvAt = pp_date_((antenatal.trend.length
      ? antenatal.trend[antenatal.trend.length - 1].date : ''));
    consider(lvAt, 'Antenatal visit on ' + lv.date, {
      bp: lv.bp, pulse: null, spo2: null, temp: '',
      weight: lv.weightKg, bmi: null
    });
  }

  var ip = pp_lastWardVitals_(pid);
  if (ip) consider(ip.at, ip.source, ip.v);

  if (!best) return null;
  return { source: best.source, takenOn: pp_fmt_(best.at, 'dd MMM yyyy'),
           vitals: best.v };
}

/** The newest nursing-note vitals for this patient, or null. */
function pp_lastWardVitals_(pid) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('IP_Timeline_DB');
    if (!sh || sh.getLastRow() < 2) return null;
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var best = null;
    for (var i = 1; i < data.length; i++) {
      if (pp_up_(data[i][m['Patient_ID']]) !== pid) continue;
      var at = pp_date_(data[i][m['Timestamp']]);
      if (!at || (best && best.at >= at)) continue;
      var nd = null;
      try { nd = JSON.parse(pp_str_(data[i][m['Note_Data_JSON']]) || '{}'); } catch (e) { nd = null; }
      var v = nd && (nd.vitals || nd.vitalSigns);
      if (!v) continue;
      best = { at: at, source: 'Ward round on ' + pp_fmt_(at, 'dd MMM yyyy'),
               v: { bp: pp_str_(v.bp || ((v.sysBp && v.diaBp) ? v.sysBp + '/' + v.diaBp : '')),
                    pulse: pp_num_(v.pr || v.pulse || v.hr),
                    spo2: pp_num_(v.spo2),
                    temp: pp_str_(v.temp),
                    weight: pp_num_(v.weight),
                    bmi: null } };
    }
    return best;
  } catch (e) { return null; }
}

// ---------------------------------------------------------------------------
// TRENDS
// ---------------------------------------------------------------------------

/**
 * A named analyte's values over time, oldest first, from verified results only.
 *
 * @param {Array} orders  as returned by lpv_resultsFor_
 */
function pp_analyteSeries_(orders, analyte) {
  var pts = [];
  (orders || []).forEach(function (o) {
    var at = pp_date_(o.date);
    if (!at) return;
    (o.results || []).forEach(function (r) {
      if (!pp_isAnalyte_(r.parameterName, analyte)) return;
      var n = pp_num_(r.value);
      if (n === null) return;
      pts.push({ date: pp_iso_(at), dateText: pp_fmt_(at, 'dd MMM yy'),
                 value: n, unit: pp_str_(r.unit), flag: pp_str_(r.flag),
                 ref: pp_str_(r.refRangeText) });
    });
  });
  pts.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
  return pts;
}

/** Blood pressure over time from the consultations, oldest first. */
function pp_bpSeries_(encounters, antenatal) {
  var pts = [];
  (encounters || []).forEach(function (e) {
    if (!e.vitals.sysBp || !e.vitals.diaBp) return;
    pts.push({ date: e.date, dateText: pp_fmt_(e.date, 'dd MMM yy'),
               sys: e.vitals.sysBp, dia: e.vitals.diaBp,
               pulse: e.vitals.pulse });
  });
  // An antenatal patient's blood pressure is taken at every ANC visit and
  // those readings are the ones that matter; leaving them out would show a
  // pregnant woman a BP chart with most of her readings missing.
  if (antenatal && antenatal.trend) {
    antenatal.trend.forEach(function (v) {
      if (!v.sysBp || !v.diaBp) return;
      pts.push({ date: v.date, dateText: pp_fmt_(v.date, 'dd MMM yy'),
                 sys: v.sysBp, dia: v.diaBp, pulse: null, antenatal: true });
    });
  }
  pts.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
  return pts;
}

/** Weight over time, from consultations and antenatal visits. */
function pp_weightSeries_(encounters, antenatal) {
  var pts = [];
  (encounters || []).forEach(function (e) {
    if (!e.vitals.weight) return;
    pts.push({ date: e.date, dateText: pp_fmt_(e.date, 'dd MMM yy'),
               value: e.vitals.weight });
  });
  if (antenatal && antenatal.trend) {
    antenatal.trend.forEach(function (v) {
      if (!v.weightKg) return;
      pts.push({ date: v.date, dateText: pp_fmt_(v.date, 'dd MMM yy'),
                 value: v.weightKg, antenatal: true });
    });
  }
  pts.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
  return pts;
}

/**
 * Which condition panels this patient should see.
 *
 * Driven by the RECORD, not by a flag somebody has to remember to set:
 *
 *   diabetes      the Conditions field says so, OR there is an HbA1c or a
 *                 fasting sugar on file. A patient with three years of HbA1c
 *                 results and a blank Conditions field is a diabetic whose
 *                 registration form was filled in quickly.
 *   hypertension  the Conditions field says so, OR there are at least three
 *                 recorded blood pressures. Three readings is the point at
 *                 which a line is a trend rather than two dots.
 *   pregnancy     an OPEN antenatal record. Nothing else — a Conditions field
 *                 saying "pregnant" from two years ago is not a pregnancy.
 *   immunisation  a child, by date of birth, or any dose already recorded.
 */
function pp_conditions_(profile, orders, bpSeries, antenatal, immunisation) {
  var text = pp_up_(profile.comorb);
  var saysDm = /\bDM\b|DIABET|T2DM|T1DM|NIDDM|IDDM/.test(text);
  var saysHtn = /\bHTN\b|HYPERTEN|\bHT\b/.test(text);

  var hba1c = pp_analyteSeries_(orders, 'HBA1C');
  var fbs = pp_analyteSeries_(orders, 'FBS');
  var ppbs = pp_analyteSeries_(orders, 'PPBS');
  var rbs = pp_analyteSeries_(orders, 'RBS');

  return {
    diabetes: saysDm || hba1c.length > 0 || fbs.length > 0,
    diabetesStatedOnFile: saysDm,
    hypertension: saysHtn || bpSeries.length >= 3,
    hypertensionStatedOnFile: saysHtn,
    pregnancy: !!antenatal,
    child: !!immunisation,
    series: { hba1c: hba1c, fbs: fbs, ppbs: ppbs, rbs: rbs }
  };
}

/**
 * Should the immunisation panel be shown to this person at all?
 *
 * ONE RULE NOW: is this a child. The card is a child's document and its
 * purpose is telling a mother when her child's next vaccination falls due, so
 * an adult opening the portal should not be offered it — there is nothing on
 * the schedule after MC.IMMUN_MAX_AGE_YEARS for them to be told about.
 *
 * The previous test also asked whether any dose had been recorded, which was
 * the wrong question in both directions: it showed the panel to an adult who
 * happened to have one row in the register, and it hid it from a child whose
 * doses had simply never been typed in — which, now that doses are presumed
 * given rather than entered, is most children.
 */
function pp_wantsImmunisation_(view) {
  return !!(view && view.eligible);
}

// ---------------------------------------------------------------------------
// THE HOME CALL
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. Everything the patient's home screen shows, in one call.
 *
 * One call rather than eight: the portal is the screen most likely to be
 * opened on a phone on mobile data, and eight round trips each carrying a
 * session check and a sheet open is most of a second before anything paints.
 *
 * @param {string} sessionToken
 */
function portalHome(sessionToken) {
  try {
    var me = pp_me_(sessionToken);
    var pid = me.patientId;

    var profile = pt_readProfile_(pid);
    if (!profile) {
      return { success: false, message: 'Your record could not be read. Please ' +
                                        'contact the clinic.' };
    }

    // s.11(1)(b) / finding M2: a patient reading their own record is still a
    // read of that record, and the log should show it was them.
    try {
      dpdpLogRead_({ username: me.username, role: 'patient' }, 'Patient', pid,
                   { endpoint: 'portalHome', self: true });
    } catch (e) {}

    var appts = pp_appointments_(pid);
    var encounters = pp_encounters_(pid);
    var labs = lpv_resultsFor_(pid, PP.LAB_LIMIT);
    var orders = (labs && labs.success) ? labs.orders : [];

    var antenatal = null;
    try {
      if (typeof mc_antenatalView_ === 'function') antenatal = mc_antenatalView_(pid);
    } catch (e) { antenatal = null; }

    var immunisation = null;
    try {
      if (typeof mc_immunisationView_ === 'function') {
        var view = mc_immunisationView_(pid, profile.dob);
        if (pp_wantsImmunisation_(view)) immunisation = view;
      }
    } catch (e) { immunisation = null; }

    var bp = pp_bpSeries_(encounters, antenatal);
    var conditions = pp_conditions_(profile, orders, bp, antenatal, immunisation);

    return {
      success: true,
      patient: {
        id: profile.id,
        name: profile.name,
        firstName: pp_str_(profile.name).split(/\s+/)[0] || pp_str_(profile.name),
        age: profile.age,
        gender: profile.gender,
        bloodGroup: profile.bloodGroup,
        conditions: profile.comorb
      },

      appointments: {
        upcoming: appts.upcoming,
        past: appts.past,
        lastVisit: appts.lastVisit || ''
      },
      review: pp_reviewDate_(encounters),

      medications: pp_currentMeds_(pid, encounters),

      vitals: pp_lastVitals_(pid, encounters, antenatal),
      weightTrend: pp_weightSeries_(encounters, antenatal),
      bpTrend: bp,

      // Released results only — lpv_resultsFor_ drops drafts and anything not
      // verified, which is the whole reason the portal reuses it rather than
      // reading LAB_RESULTS itself.
      labs: {
        orders: orders.slice(0, 6),
        totalOrders: orders.length,
        abnormalRecent: orders.slice(0, 6).reduce(function (n, o) {
          return n + (o.abnormal || 0);
        }, 0)
      },

      conditions: conditions,
      antenatal: antenatal,
      immunisation: immunisation,

      lastVisitNote: encounters.length ? {
        dateText: encounters[0].dateText,
        doctorName: encounters[0].doctorName,
        diagnosis: encounters[0].diagnosis,
        advice: encounters[0].advice
      } : null,

      message: ''
    };
  } catch (err) {
    return pp_fail_(err);
  }
}

// ---------------------------------------------------------------------------
// BOOKING, WITH A CHOICE OF DOCTOR
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. The doctors a patient may book, and the next day each of
 * them has a free slot.
 *
 * "Next available" is worth the extra work: a list of names with no dates
 * makes the patient click through every doctor to find out who can see them
 * this week, which is the one thing they came to find out.
 */
function portalBookableDoctors(sessionToken) {
  try {
    pp_me_(sessionToken);
    var docs = getActiveDoctors() || [];
    var out = docs.map(function (d) {
      var next = pp_nextFreeDay_(d.doctorId, 30);
      return {
        doctorId: d.doctorId,
        name: d.name,
        specialty: d.specialty,
        regNo: d.regNo,
        nextFree: next ? next.dateIso : '',
        nextFreeText: next ? pp_fmt_(next.dateIso, 'EEE, dd MMM') : '',
        nextFreeTime: next ? next.time12 : '',
        freeCount: next ? next.count : 0
      };
    });
    // Whoever can see them soonest, first. A doctor with nothing free in the
    // next month goes to the bottom rather than being hidden — the patient
    // may want to wait for a particular person.
    out.sort(function (a, b) {
      if (!a.nextFree && !b.nextFree) return a.name.localeCompare(b.name);
      if (!a.nextFree) return 1;
      if (!b.nextFree) return -1;
      return String(a.nextFree).localeCompare(String(b.nextFree));
    });
    return { success: true, doctors: out,
             message: out.length ? '' :
               'No doctor is taking bookings at the moment. Please telephone the clinic.' };
  } catch (err) {
    return pp_fail_(err, { doctors: [] });
  }
}

/** The first day within `days` on which this doctor has a bookable slot. */
function pp_nextFreeDay_(doctorId, days) {
  var today = new Date();
  for (var i = 0; i < days; i++) {
    var probe = new Date(today.getTime() + i * 86400000);
    var key = dc_fmtDate_(probe);
    var free = pp_freeSlots_(doctorId, key, i === 0);
    if (free.length) {
      return { dateIso: key, time12: free[0].time12, count: free.length };
    }
  }
  return null;
}

/**
 * The slots a PATIENT may book with one doctor on one date.
 *
 * Deliberately not getDoctorSlotsForDate: that one returns every slot with
 * `bookings` attached — the names, ids and purposes of the other patients in
 * the clinic that day. It is a staff endpoint and it refuses patient sessions
 * for that reason. This returns free times and nothing else.
 *
 * @param {boolean} isToday  drop slots that have already passed
 */
function pp_freeSlots_(doctorId, dateKey, isToday) {
  var grid = [];
  try { grid = ds_generateSlots_(doctorId, dateKey) || []; } catch (e) { return []; }
  if (!grid.length) return [];

  var booked = {};
  try { booked = ds_bookingsByTime_(doctorId, dateKey) || {}; } catch (e) { booked = {}; }

  var now = new Date();
  var nowMins = now.getHours() * 60 + now.getMinutes();

  return grid.filter(function (s) {
    if ((booked[s.time24] || []).length >= (s.maxPerSlot || 1)) return false;
    // A slot in the past is not a slot. Offering one produces a booking the
    // patient cannot attend and the desk has to unpick.
    if (isToday && dc_minutes_(s.time24) <= nowMins + 15) return false;
    return true;
  }).map(function (s) {
    return { time24: s.time24, time12: s.time12, session: s.sessionLabel };
  });
}

/**
 * FRONTEND ENTRY. Free slots for one doctor on one date.
 *
 * @param {string} doctorId
 * @param {string} dateStr
 * @param {string} sessionToken
 */
function portalDoctorSlots(doctorId, dateStr, sessionToken) {
  try {
    pp_me_(sessionToken);
    var did = pp_str_(doctorId);
    if (!did) return { success: false, slots: [], message: 'Choose a doctor first.' };

    var doc = dc_getDoctorById_(did);
    if (!doc) return { success: false, slots: [], message: 'That doctor is not on file.' };
    if (doc.status !== 'ACTIVE') {
      return { success: false, slots: [],
               message: doc.name + ' is not taking bookings at the moment.' };
    }

    var key = dc_dateKey_(dateStr);
    if (!key) return { success: false, slots: [], message: 'Choose a date.' };

    var todayKey = dc_fmtDate_(new Date());
    if (key < todayKey) {
      return { success: false, slots: [], message: 'That date has passed.' };
    }

    var free = pp_freeSlots_(did, key, key === todayKey);
    var grid = [];
    try { grid = ds_generateSlots_(did, key) || []; } catch (e) { grid = []; }

    // Three different empties, three different messages. "No slots" for a
    // doctor who does not work Sundays, for a day that is fully booked, and
    // for a day whose clinic has already finished are three different things
    // the patient should do three different things about.
    var msg = '';
    if (!free.length) {
      if (!grid.length) msg = doc.name + ' does not hold a clinic on this day.';
      else if (key === todayKey) msg = 'Today\'s clinic is over or fully booked. Try tomorrow.';
      else msg = doc.name + ' is fully booked on this day.';
    }

    return { success: true, doctorId: doc.doctorId, doctorName: doc.name,
             date: key, slots: free, message: msg };
  } catch (err) {
    return pp_fail_(err, { slots: [] });
  }
}

/**
 * FRONTEND ENTRY. Books the signed-in patient with the doctor they chose.
 *
 * Every figure the browser could lie about is re-derived here: the patient is
 * the session's, the fee is the doctor's own consult fee from the Doctors
 * master, and the slot is checked against the doctor's generated grid rather
 * than trusted because the browser offered it.
 *
 * @param {{doctorId, date, time, purpose}} payload
 */
function portalBookAppointment(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var me = pp_me_(sessionToken);
    var pid = me.patientId;

    var doc = dc_getDoctorById_(pp_str_(payload.doctorId));
    if (!doc) return { success: false, message: 'Choose a doctor.' };
    if (doc.status !== 'ACTIVE') {
      return { success: false, message: doc.name + ' is not taking bookings.' };
    }

    var key = dc_dateKey_(payload.date);
    var time24 = dc_to24_(payload.time);
    if (!key) return { success: false, message: 'Choose a date.' };
    if (!time24) return { success: false, message: 'Choose a time.' };

    var todayKey = dc_fmtDate_(new Date());
    if (key < todayKey) return { success: false, message: 'That date has passed.' };

    // The slot must be free in the doctor's OWN grid at the moment of
    // booking. The browser's list is a snapshot from when the screen painted.
    var free = pp_freeSlots_(doc.doctorId, key, key === todayKey);
    var ok = false;
    for (var i = 0; i < free.length; i++) if (free[i].time24 === time24) { ok = true; break; }
    if (!ok) {
      return { success: false, code: 'SLOT_GONE',
               message: 'That time has just been taken. Please pick another.' };
    }

    // One live appointment per patient per doctor per day. A patient who
    // double-books is a patient the desk has to telephone.
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
    if (!sh) return { success: false, message: 'Appointments are unavailable. Please telephone the clinic.' };
    var m = dc_headerMap_(sh);
    var docCol = (m['Doctor_ID'] === undefined) ? -1 : m['Doctor_ID'];
    if (docCol === -1) {
      return { success: false,
               message: 'Online booking is not configured yet. Please telephone the clinic.' };
    }

    var data = sh.getDataRange().getDisplayValues();
    var LIVE = ['Booked', 'Arrived', 'In-Progress'];
    for (var r = 1; r < data.length; r++) {
      if (pp_up_(data[r][1]) !== pid) continue;
      if (LIVE.indexOf(pp_str_(data[r][6])) === -1) continue;
      if (dc_dateKey_(data[r][3]) !== key) continue;
      var rowDoc = pp_up_(pp_str_(data[r][docCol]) || DC_DEFAULT_DOCTOR);
      if (rowDoc !== pp_up_(doc.doctorId)) continue;
      return { success: false, code: 'ALREADY_BOOKED',
               message: 'You already have an appointment with ' + doc.name +
                        ' on ' + pp_fmt_(key, 'EEE, dd MMM') + ' at ' +
                        pp_str_(data[r][4]) + '.' };
    }

    var PURPOSES = ['Consultation', 'Follow-up', 'Routine Checkup',
                    'Antenatal check-up', 'Immunisation'];
    var purpose = pp_str_(payload.purpose);
    if (PURPOSES.indexOf(purpose) === -1) purpose = 'Consultation';

    var profile = pt_readProfile_(pid);
    var apptId = apt_newId_();
    var row = new Array(sh.getLastColumn()).fill('');
    row[0] = apptId;
    row[1] = pid;
    row[2] = pp_str_(profile && profile.name);
    row[3] = key;
    row[4] = dc_to12_(time24);
    row[5] = purpose;
    row[6] = 'Booked';
    // The doctor's own tariff, read here — never the figure the browser sent.
    row[7] = dc_money_(doc.consultFee);
    row[8] = new Date();
    row[docCol] = doc.doctorId;
    if (m['Doctor_Name_Snapshot'] !== undefined) row[m['Doctor_Name_Snapshot']] = doc.name;
    if (m['Booked_By'] !== undefined) row[m['Booked_By']] = me.username;
    if (m['Attribution_Source'] !== undefined) row[m['Attribution_Source']] = 'PORTAL';
    sh.appendRow(row);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: me.username, role: 'patient' },
                'APPOINTMENT_BOOK', 'Appointment', apptId,
                { doctorId: doc.doctorId, patientId: pid, date: key,
                  time: time24, via: 'portal' });
    } catch (e) {}

    return { success: true, apptId: apptId,
             doctorName: doc.name, date: key,
             dateText: pp_fmt_(key, 'EEEE, dd MMM yyyy'),
             time: dc_to12_(time24),
             message: 'Booked with ' + doc.name + ' on ' +
                      pp_fmt_(key, 'EEEE, dd MMM') + ' at ' + dc_to12_(time24) + '.' };
  } catch (err) {
    return pp_fail_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * FRONTEND ENTRY. The patient cancels their own appointment.
 *
 * A booking a patient can make and cannot cancel is a booking they simply do
 * not attend, and a clinic that cannot tell those apart cannot reuse the
 * slot. Only their own, only a future one, and the row is marked rather than
 * deleted so the slot's history survives.
 */
function portalCancelAppointment(apptId, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var me = pp_me_(sessionToken);
    var want = pp_str_(apptId);
    if (!want) return { success: false, message: 'Which appointment?' };

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Appointments');
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'Not found.' };

    var cell = sh.getRange(2, 1, sh.getLastRow() - 1, 1)
                 .createTextFinder(want).matchEntireCell(true).findNext();
    if (!cell) return { success: false, message: 'That appointment is not on file.' };

    var r = cell.getRow();
    var row = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];
    if (pp_up_(row[1]) !== me.patientId) {
      // Deliberately the same answer as a missing appointment: whether an
      // appointment id exists is not something this endpoint will confirm.
      return { success: false, message: 'That appointment is not on file.' };
    }

    var key = dc_dateKey_(row[3]);
    if (key && key < dc_fmtDate_(new Date())) {
      return { success: false,
               message: 'That appointment has already passed, so there is nothing ' +
                        'to cancel.' };
    }
    var status = pp_str_(row[6]);
    if (status === 'Cancelled') {
      return { success: true, message: 'That appointment was already cancelled.' };
    }
    if (status !== 'Booked') {
      return { success: false,
               message: 'You have already been checked in for this appointment. ' +
                        'Please speak to the front desk.' };
    }

    sh.getRange(r, 7).setValue('Cancelled');
    SpreadsheetApp.flush();
    try {
      logAudit_({ username: me.username, role: 'patient' },
                'APPOINTMENT_STATUS', 'Appointment', want,
                { from: status, to: 'Cancelled', patientId: me.patientId,
                  via: 'portal' });
    } catch (e) {}

    return { success: true, message: 'Cancelled. The slot is free for someone else.' };
  } catch (err) {
    return pp_fail_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ---------------------------------------------------------------------------
// RECORDS
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. This patient's own lab results — verified ones only.
 *
 * @param {string} sessionToken
 */
function portalLabResults(sessionToken) {
  try {
    var me = pp_me_(sessionToken);
    try {
      dpdpLogRead_({ username: me.username, role: 'patient' }, 'Patient',
                   me.patientId, { endpoint: 'portalLabResults', self: true });
    } catch (e) {}
    var res = lpv_resultsFor_(me.patientId, PP.LAB_LIMIT);
    return { success: !!res.success, orders: res.orders || [],
             message: res.success ? '' : res.message };
  } catch (err) {
    return pp_fail_(err, { orders: [] });
  }
}

/**
 * FRONTEND ENTRY. The patient's visit history: every consultation with what
 * was found, prescribed and advised, and every admission with its signed
 * discharge summary.
 */
function portalRecords(sessionToken) {
  try {
    var me = pp_me_(sessionToken);
    var pid = me.patientId;
    try {
      dpdpLogRead_({ username: me.username, role: 'patient' }, 'Patient', pid,
                   { endpoint: 'portalRecords', self: true });
    } catch (e) {}

    var encounters = pp_encounters_(pid).slice(0, PP.VISIT_LIMIT).map(function (e) {
      return {
        encounterId: e.encounterId,
        dateText: e.dateText,
        doctorName: e.doctorName,
        complaints: e.complaints,
        diagnosis: e.diagnosis,
        advice: e.advice,
        reviewDate: e.reviewDate,
        vitals: e.vitals,
        meds: e.meds.map(function (m) {
          return { name: pp_str_(m.drugName || m.brand),
                   sig: pp_str_(m.sig),
                   days: pp_str_(m.duration || m.days),
                   notes: pp_str_(m.instructions || m.notes) };
        })
      };
    });

    var stays = [];
    try { stays = pp_admissions_(pid); } catch (e) { stays = []; }

    return { success: true, consultations: encounters, admissions: stays,
             message: (encounters.length || stays.length) ? '' :
               'Nothing recorded yet. Your visits will appear here.' };
  } catch (err) {
    return pp_fail_(err, { consultations: [], admissions: [] });
  }
}

/**
 * This patient's admissions, each with its SIGNED discharge summary if there
 * is one.
 *
 * An unsigned summary is not shown at all — not even as sections. The patient
 * is told the summary is being prepared, which is true and is the only thing
 * they can act on.
 */
function pp_admissions_(pid) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('IP_Admissions');
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getDataRange().getValues();
  var out = [];

  for (var i = 1; i < rows.length; i++) {
    if (pp_up_(rows[i][1]) !== pid) continue;
    var doa = pp_date_(rows[i][4]);
    if (!doa) continue;
    var ip = pp_str_(rows[i][0]);
    var dod = pp_date_(rows[i][12]);

    var rec = {
      ipNumber: ip,
      admittedText: pp_fmt_(doa),
      dischargedText: dod ? pp_fmt_(dod) : '',
      stillAdmitted: !dod,
      ward: pp_str_(rows[i][7]) || pp_str_(rows[i][8]),
      diagnosis: pp_str_(rows[i][10]),
      consultant: pp_str_(rows[i][9]),
      summaryReady: false,
      summaryNote: 'Your discharge summary is being prepared.',
      sections: []
    };

    try {
      if (typeof dsx_summaryIdFor_ === 'function') {
        var sid = dsx_summaryIdFor_(ip);
        var header = sid ? dsx_getHeader_(sid) : null;
        if (header && dsx_upper_(header.Status) === DSX_STATUS.SIGNED) {
          var ref = dsx_resolveRef_(sid, 'SIGNED:' + dsx_int_(header.Last_Signed_Snapshot_No));
          if (ref && ref.payload) {
            rec.summaryReady = true;
            rec.summaryNote = 'Signed ' + pp_fmt_(header.Signed_At) +
                              (dsx_str_(header.Signed_By) ? ' by ' + dsx_str_(header.Signed_By) : '');
            rec.sections = dsb_orderedSections_(ref.payload)
              .filter(function (s) { return !s.empty; })
              .map(function (s) {
                return { key: s.key, title: s.title, text: dsb_sectionText_(s) };
              });
          }
        }
      }
    } catch (e) { /* the discharge module may not be on this deployment */ }

    out.push({ sortKey: pp_iso_(doa), rec: rec });
  }
  out.sort(function (a, b) { return String(b.sortKey).localeCompare(String(a.sortKey)); });
  return out.map(function (x) { return x.rec; });
}
