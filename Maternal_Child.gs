// ============================================================================
// Maternal_Child.gs — Crescentia HealthTech
// Antenatal care and childhood immunisation: two registers the clinic keeps
// and the patient can see.
// ----------------------------------------------------------------------------
// WHY THESE ARE NEW
//
// The patient portal is asked to show a pregnant patient her current trend,
// her next checkup and the baby's growth, and to show a parent how far
// through the immunisation schedule their child is and what is next. Neither
// could be done, because neither existed anywhere in this project: a grep
// for LMP, EDD, antenatal, immunisation or vaccine found two rows of the OP
// complaint phrase library and nothing else. Antenatal visits were being
// written as free-text progress notes, and vaccinations were being recorded
// on the card the parent carries and nowhere the clinic could read.
//
// Free text cannot be trended and a paper card cannot be reminded from. So
// the two things the portal needs to show are the two things the clinic needs
// to record, and this file is the register for both.
//
// WHAT AN ANTENATAL VISIT RECORDS
//
// The minimum set a trend needs and an ANC card carries: weight, blood
// pressure, haemoglobin, fundal height, fetal heart rate, presentation,
// urine albumin and sugar, and the date of the next visit. Gestational age is
// NOT stored per visit — it is computed from the LMP, because a stored week
// number is a number that goes stale and disagrees with itself between visits.
//
// WHAT THE IMMUNISATION SCHEDULE IS
//
// India's Universal Immunisation Programme, as the National Immunization
// Schedule publishes it, with each dose's due age in DAYS from birth. Due
// dates come from the child's date of birth, so the schedule is reference
// data and the register only records what was GIVEN. That way a schedule
// revision does not rewrite history, and a child registered today gets their
// whole schedule without anybody entering thirty rows.
//
// WHAT IT DOES NOT DO
//
// It does not decide anything clinical. Nothing here flags a pregnancy as
// high risk, computes a Bishop score, or says a vaccine may be given. It
// records what a clinician wrote and computes dates from dates.
// ============================================================================

var MC = {
  ANC:      'ANC_Register',
  ANC_VISITS: 'ANC_Visits',
  IMMUN:    'Immunisation_Given',
  TZ:       'Asia/Kolkata',

  /** A pregnancy is closed once it is this far past the EDD with no outcome. */
  ANC_STALE_DAYS: 60
};

function mc_str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function mc_up_(v) { return mc_str_(v).toUpperCase(); }
function mc_num_(v) { var n = parseFloat(v); return isNaN(n) ? null : n; }

function mc_date_(v) {
  if (!v) return null;
  if (typeof cresc_parseDate_ === 'function') return cresc_parseDate_(v);
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var d = new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}

function mc_fmt_(v, pattern) {
  var d = mc_date_(v);
  if (!d) return '';
  try { return Utilities.formatDate(d, MC.TZ, pattern || 'dd-MMM-yyyy'); }
  catch (e) { return ''; }
}

/** An ISO date, which is what a date input and a chart axis both want. */
function mc_iso_(v) {
  var d = mc_date_(v);
  return d ? Utilities.formatDate(d, MC.TZ, 'yyyy-MM-dd') : '';
}

/** Whole days between two dates, positive when `to` is later. */
function mc_days_(from, to) {
  var a = mc_date_(from), b = mc_date_(to);
  if (!a || !b) return null;
  // Midnight-to-midnight, so a visit at 09:00 and one at 18:00 on the same
  // day are zero days apart rather than "0.375".
  var a0 = new Date(a.getFullYear(), a.getMonth(), a.getDate());
  var b0 = new Date(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.round((b0 - a0) / 86400000);
}

// ---------------------------------------------------------------------------
// SHEETS
// ---------------------------------------------------------------------------

function mc_ancSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), MC.ANC, [
    'ANC_ID', 'Patient_ID', 'Patient_Name', 'LMP', 'EDD',
    'Gravida', 'Para', 'Living', 'Abortions',
    'Blood_Group', 'Height_Cm', 'Booking_Weight_Kg',
    'Risk_Notes', 'Status', 'Outcome', 'Outcome_Date', 'Outcome_Notes',
    'Opened_By', 'Opened_At', 'Closed_By', 'Closed_At'
  ]);
}

function mc_ancVisitSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), MC.ANC_VISITS, [
    'Visit_ID', 'ANC_ID', 'Patient_ID', 'Visit_Date',
    'Weight_Kg', 'Sys_BP', 'Dia_BP', 'Pallor',
    'Haemoglobin', 'Urine_Albumin', 'Urine_Sugar',
    'Fundal_Height_Cm', 'Fetal_Heart_Rate', 'Presentation', 'Fetal_Movements',
    'Oedema', 'Complaints', 'Advice', 'Next_Visit_Date',
    'Recorded_By', 'Recorded_At', 'Doctor_ID'
  ]);
}

function mc_immunSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), MC.IMMUN, [
    'Entry_ID', 'Patient_ID', 'Vaccine_Code', 'Dose_Label',
    'Given_On', 'Batch_No', 'Site', 'Route',
    'Given_By', 'Recorded_By', 'Recorded_At', 'Notes'
  ]);
}

/** ONE-OFF. Creates both registers. Safe to re-run. */
function mcSetup() {
  mc_ancSheet_();
  mc_ancVisitSheet_();
  mc_immunSheet_();
  return 'Maternal and child registers ready: ' +
         [MC.ANC, MC.ANC_VISITS, MC.IMMUN].join(', ') + '.';
}

// ---------------------------------------------------------------------------
// THE IMMUNISATION SCHEDULE — reference data, not a register
// ---------------------------------------------------------------------------

/**
 * India's Universal Immunisation Programme, as the National Immunization
 * Schedule publishes it for infants and children.
 *
 * `dueDays` is days from DATE OF BIRTH, so every due date is computed rather
 * than stored. `window` is how long after the due date the dose is still
 * ordinarily given, which is what separates "due" from "overdue" — a dose one
 * day past its date is due, not missed, and telling a parent otherwise is how
 * a schedule stops being believed.
 *
 * This is the national schedule and NOT a clinical decision: a clinician may
 * give a dose early, late, or not at all for reasons this table cannot know.
 * The register records what was given; this table only says what was expected
 * and when.
 */
var MC_SCHEDULE = [
  { code: 'BCG',      label: 'BCG',                    dueDays: 0,    window: 365, at: 'At birth' },
  { code: 'HEPB0',    label: 'Hepatitis B — birth dose', dueDays: 0,  window: 1,   at: 'At birth (within 24 hours)' },
  { code: 'OPV0',     label: 'OPV — zero dose',        dueDays: 0,    window: 15,  at: 'At birth' },

  { code: 'OPV1',     label: 'OPV 1',                  dueDays: 42,   window: 30,  at: '6 weeks' },
  { code: 'PENTA1',   label: 'Pentavalent 1',          dueDays: 42,   window: 30,  at: '6 weeks' },
  { code: 'ROTA1',    label: 'Rotavirus 1',            dueDays: 42,   window: 30,  at: '6 weeks' },
  { code: 'PCV1',     label: 'PCV 1',                  dueDays: 42,   window: 30,  at: '6 weeks' },
  { code: 'IPV1',     label: 'fIPV 1',                 dueDays: 42,   window: 30,  at: '6 weeks' },

  { code: 'OPV2',     label: 'OPV 2',                  dueDays: 70,   window: 30,  at: '10 weeks' },
  { code: 'PENTA2',   label: 'Pentavalent 2',          dueDays: 70,   window: 30,  at: '10 weeks' },
  { code: 'ROTA2',    label: 'Rotavirus 2',            dueDays: 70,   window: 30,  at: '10 weeks' },

  { code: 'OPV3',     label: 'OPV 3',                  dueDays: 98,   window: 30,  at: '14 weeks' },
  { code: 'PENTA3',   label: 'Pentavalent 3',          dueDays: 98,   window: 30,  at: '14 weeks' },
  { code: 'ROTA3',    label: 'Rotavirus 3',            dueDays: 98,   window: 30,  at: '14 weeks' },
  { code: 'PCV2',     label: 'PCV 2',                  dueDays: 98,   window: 30,  at: '14 weeks' },
  { code: 'IPV2',     label: 'fIPV 2',                 dueDays: 98,   window: 30,  at: '14 weeks' },

  { code: 'MR1',      label: 'Measles–Rubella 1',      dueDays: 270,  window: 90,  at: '9–12 months' },
  { code: 'JE1',      label: 'JE 1 (endemic areas)',   dueDays: 270,  window: 90,  at: '9–12 months' },
  { code: 'PCVB',     label: 'PCV booster',            dueDays: 270,  window: 90,  at: '9–12 months' },
  { code: 'VITA1',    label: 'Vitamin A — 1st dose',   dueDays: 270,  window: 90,  at: '9 months' },

  { code: 'DPTB1',    label: 'DPT booster 1',          dueDays: 480,  window: 120, at: '16–24 months' },
  { code: 'MR2',      label: 'Measles–Rubella 2',      dueDays: 480,  window: 120, at: '16–24 months' },
  { code: 'OPVB',     label: 'OPV booster',            dueDays: 480,  window: 120, at: '16–24 months' },
  { code: 'JE2',      label: 'JE 2 (endemic areas)',   dueDays: 480,  window: 120, at: '16–24 months' },

  { code: 'DPTB2',    label: 'DPT booster 2',          dueDays: 1825, window: 365, at: '5–6 years' },
  { code: 'TD10',     label: 'Td — 10 years',          dueDays: 3650, window: 365, at: '10 years' },
  { code: 'TD16',     label: 'Td — 16 years',          dueDays: 5840, window: 365, at: '16 years' }
];

/** FRONTEND ENTRY. The schedule itself, for a picker. No patient data. */
function mcGetSchedule() {
  return { success: true, schedule: MC_SCHEDULE.map(function (v) {
    return { code: v.code, label: v.label, at: v.at, dueDays: v.dueDays };
  }) };
}

function mc_scheduleEntry_(code) {
  var c = mc_up_(code);
  for (var i = 0; i < MC_SCHEDULE.length; i++) {
    if (MC_SCHEDULE[i].code === c) return MC_SCHEDULE[i];
  }
  return null;
}

// ---------------------------------------------------------------------------
// ANTENATAL — internal readers
// ---------------------------------------------------------------------------

/**
 * The open pregnancy for a patient, or null.
 *
 * A patient may have several rows over the years; the OPEN one is the current
 * pregnancy. If more than one is open — a data-entry slip — the most recently
 * opened wins and the others are left alone rather than guessed at.
 */
function mc_openAnc_(patientId) {
  var pid = mc_up_(patientId);
  if (!pid) return null;
  var sh = mc_ancSheet_();
  var m = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  var best = null;
  for (var i = 1; i < (data ? data.length : 0); i++) {
    if (mc_up_(data[i][m['Patient_ID']]) !== pid) continue;
    if (mc_up_(data[i][m['Status']]) !== 'OPEN') continue;
    var rec = {
      row: i + 1,
      ancId: mc_str_(data[i][m['ANC_ID']]),
      lmp: mc_date_(data[i][m['LMP']]),
      edd: mc_date_(data[i][m['EDD']]),
      gravida: mc_str_(data[i][m['Gravida']]),
      para: mc_str_(data[i][m['Para']]),
      living: mc_str_(data[i][m['Living']]),
      abortions: mc_str_(data[i][m['Abortions']]),
      bloodGroup: mc_str_(data[i][m['Blood_Group']]),
      heightCm: mc_num_(data[i][m['Height_Cm']]),
      bookingWeight: mc_num_(data[i][m['Booking_Weight_Kg']]),
      riskNotes: mc_str_(data[i][m['Risk_Notes']]),
      openedAt: mc_date_(data[i][m['Opened_At']])
    };
    if (!best || (rec.openedAt && best.openedAt && rec.openedAt > best.openedAt)) best = rec;
    else if (!best) best = rec;
  }
  return best;
}

/** Every visit on one ANC record, oldest first. */
function mc_ancVisits_(ancId) {
  var want = mc_up_(ancId);
  if (!want) return [];
  var sh = mc_ancVisitSheet_();
  var m = dc_headerMap_(sh);
  var data = dc_sheetValues_(sh);
  var out = [];
  for (var i = 1; i < (data ? data.length : 0); i++) {
    if (mc_up_(data[i][m['ANC_ID']]) !== want) continue;
    out.push({
      visitId: mc_str_(data[i][m['Visit_ID']]),
      date: mc_date_(data[i][m['Visit_Date']]),
      weightKg: mc_num_(data[i][m['Weight_Kg']]),
      sysBp: mc_num_(data[i][m['Sys_BP']]),
      diaBp: mc_num_(data[i][m['Dia_BP']]),
      haemoglobin: mc_num_(data[i][m['Haemoglobin']]),
      urineAlbumin: mc_str_(data[i][m['Urine_Albumin']]),
      urineSugar: mc_str_(data[i][m['Urine_Sugar']]),
      fundalHeightCm: mc_num_(data[i][m['Fundal_Height_Cm']]),
      fetalHeartRate: mc_num_(data[i][m['Fetal_Heart_Rate']]),
      presentation: mc_str_(data[i][m['Presentation']]),
      fetalMovements: mc_str_(data[i][m['Fetal_Movements']]),
      oedema: mc_str_(data[i][m['Oedema']]),
      complaints: mc_str_(data[i][m['Complaints']]),
      advice: mc_str_(data[i][m['Advice']]),
      nextVisit: mc_date_(data[i][m['Next_Visit_Date']]),
      recordedBy: mc_str_(data[i][m['Recorded_By']])
    });
  }
  out.sort(function (a, b) { return (a.date || 0) - (b.date || 0); });
  return out;
}

/**
 * Gestational age at a date, from the LMP.
 *
 * Computed, never stored. A week number written into a visit row is correct
 * for one day and wrong afterwards, and two visit rows then disagree about
 * the same pregnancy.
 *
 * @return {{days:number, weeks:number, remainder:number, text:string}|null}
 */
function mc_gestation_(lmp, at) {
  var days = mc_days_(lmp, at || new Date());
  if (days === null || days < 0) return null;
  var w = Math.floor(days / 7);
  var d = days % 7;
  return { days: days, weeks: w, remainder: d,
           text: w + ' weeks' + (d ? ' ' + d + ' day' + (d === 1 ? '' : 's') : '') };
}

/** EDD by Naegele's rule: LMP + 280 days. */
function mc_edd_(lmp) {
  var d = mc_date_(lmp);
  if (!d) return null;
  return new Date(d.getTime() + 280 * 86400000);
}

// ---------------------------------------------------------------------------
// ANTENATAL — write endpoints (the clinic side)
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. Opens an antenatal record.
 *
 * The EDD is DERIVED from the LMP unless one is given explicitly — a scan EDD
 * legitimately differs from the menstrual one, and a clinician who has a scan
 * date should be able to say so rather than have the arithmetic overrule them.
 *
 * @param {{patientId, lmp, edd, gravida, para, living, abortions,
 *          bloodGroup, heightCm, bookingWeightKg, riskNotes}} payload
 */
function mcOpenAntenatal(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var actor = crescRequire_(sessionToken, ['emr.write', 'ward.write', 'patient.write']);

    var pid = mc_up_(payload.patientId);
    if (!pid) return { success: false, message: 'Name the patient.' };

    var lmp = mc_date_(payload.lmp);
    var edd = mc_date_(payload.edd) || mc_edd_(lmp);
    if (!lmp && !edd) {
      return { success: false,
               message: 'Give the last menstrual period, or an EDD from a scan. ' +
                        'Without one of them nothing on the antenatal card can be ' +
                        'dated — not the gestational age, not the next visit, not ' +
                        'the growth chart.' };
    }
    if (lmp && mc_days_(lmp, new Date()) > 320) {
      return { success: false,
               message: 'That LMP is more than 45 weeks ago. Check the date — if ' +
                        'this is a past pregnancy, record its outcome instead of ' +
                        'opening it as current.' };
    }

    var existing = mc_openAnc_(pid);
    if (existing) {
      return { success: false, code: 'ALREADY_OPEN', ancId: existing.ancId,
               message: 'This patient already has an open antenatal record (' +
                        existing.ancId + ', EDD ' + mc_fmt_(existing.edd) +
                        '). Close it with its outcome before opening another.' };
    }

    var profile = null;
    try { profile = pt_readProfile_(pid); } catch (e) { profile = null; }
    if (!profile) return { success: false, message: 'No patient with the ID ' + pid + '.' };

    var now = new Date();
    var ancId = 'ANC-' + Utilities.formatDate(now, MC.TZ, 'yyMMdd') + '-' +
                Utilities.getUuid().substring(0, 4).toUpperCase();

    var sh = mc_ancSheet_();
    var m = dc_headerMap_(sh);
    var row = new Array(sh.getLastColumn()).fill('');
    var put = function (h, v) { if (m[h] !== undefined) row[m[h]] = v; };
    put('ANC_ID', ancId);
    put('Patient_ID', pid);
    put('Patient_Name', mc_str_(profile.name));
    put('LMP', lmp || '');
    put('EDD', edd || '');
    put('Gravida', mc_str_(payload.gravida));
    put('Para', mc_str_(payload.para));
    put('Living', mc_str_(payload.living));
    put('Abortions', mc_str_(payload.abortions));
    put('Blood_Group', mc_str_(payload.bloodGroup));
    put('Height_Cm', mc_num_(payload.heightCm) || '');
    put('Booking_Weight_Kg', mc_num_(payload.bookingWeightKg) || '');
    put('Risk_Notes', mc_str_(payload.riskNotes));
    put('Status', 'OPEN');
    put('Opened_By', actor.username);
    put('Opened_At', now);
    sh.appendRow(row);
    dc_invalidate_(MC.ANC);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'ANC_OPENED', 'Patient', pid,
                { ancId: ancId, lmp: mc_iso_(lmp), edd: mc_iso_(edd) });
    } catch (e) {}

    var g = mc_gestation_(lmp, now);
    return { success: true, ancId: ancId, edd: mc_fmt_(edd),
             gestation: g ? g.text : '',
             message: 'Antenatal record opened. EDD ' + mc_fmt_(edd) +
                      (g ? ', currently ' + g.text + '.' : '.') };
  } catch (err) {
    return { success: false,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * FRONTEND ENTRY. Records one antenatal visit.
 *
 * @param {{patientId, ancId, visitDate, weightKg, sysBp, diaBp, haemoglobin,
 *          urineAlbumin, urineSugar, fundalHeightCm, fetalHeartRate,
 *          presentation, fetalMovements, oedema, complaints, advice,
 *          nextVisitDate}} payload
 */
function mcRecordAntenatalVisit(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var actor = crescRequire_(sessionToken, ['emr.write', 'ward.write']);

    var pid = mc_up_(payload.patientId);
    var anc = payload.ancId ? { ancId: mc_str_(payload.ancId) } : mc_openAnc_(pid);
    if (!anc || !anc.ancId) {
      return { success: false, code: 'NO_ANC',
               message: 'There is no open antenatal record for ' + pid +
                        '. Open one first — the visit has to hang off an LMP, or ' +
                        'nothing about it can be dated.' };
    }

    var visitDate = mc_date_(payload.visitDate) || new Date();
    var next = mc_date_(payload.nextVisitDate);
    if (next && mc_days_(visitDate, next) < 0) {
      return { success: false,
               message: 'The next visit date is before this visit. Check it.' };
    }

    var now = new Date();
    var visitId = 'ANCV-' + Utilities.formatDate(now, MC.TZ, 'yyMMdd-HHmmss') + '-' +
                  Utilities.getUuid().substring(0, 4).toUpperCase();

    var sh = mc_ancVisitSheet_();
    var m = dc_headerMap_(sh);
    var row = new Array(sh.getLastColumn()).fill('');
    var put = function (h, v) { if (m[h] !== undefined) row[m[h]] = v; };
    put('Visit_ID', visitId);
    put('ANC_ID', anc.ancId);
    put('Patient_ID', pid);
    put('Visit_Date', visitDate);
    put('Weight_Kg', mc_num_(payload.weightKg) || '');
    put('Sys_BP', mc_num_(payload.sysBp) || '');
    put('Dia_BP', mc_num_(payload.diaBp) || '');
    put('Pallor', mc_str_(payload.pallor));
    put('Haemoglobin', mc_num_(payload.haemoglobin) || '');
    put('Urine_Albumin', mc_str_(payload.urineAlbumin));
    put('Urine_Sugar', mc_str_(payload.urineSugar));
    put('Fundal_Height_Cm', mc_num_(payload.fundalHeightCm) || '');
    put('Fetal_Heart_Rate', mc_num_(payload.fetalHeartRate) || '');
    put('Presentation', mc_str_(payload.presentation));
    put('Fetal_Movements', mc_str_(payload.fetalMovements));
    put('Oedema', mc_str_(payload.oedema));
    put('Complaints', mc_str_(payload.complaints));
    put('Advice', mc_str_(payload.advice));
    put('Next_Visit_Date', next || '');
    put('Recorded_By', actor.username);
    put('Recorded_At', now);
    put('Doctor_ID', mc_str_(actor.doctorId));
    sh.appendRow(row);
    dc_invalidate_(MC.ANC_VISITS);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'ANC_VISIT', 'Patient', pid,
                { ancId: anc.ancId, visitId: visitId, date: mc_iso_(visitDate) });
    } catch (e) {}

    return { success: true, visitId: visitId,
             message: 'Antenatal visit recorded' +
                      (next ? '. Next visit ' + mc_fmt_(next) + '.' :
                       '. No next visit date was set — she has nothing to come back on.') };
  } catch (err) {
    return { success: false,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** FRONTEND ENTRY. Closes an antenatal record with its outcome. */
function mcCloseAntenatal(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var actor = crescRequire_(sessionToken, ['emr.write', 'ward.write']);

    var want = mc_up_(payload.ancId);
    var outcome = mc_up_(payload.outcome);
    var ALLOWED = ['LIVE_BIRTH', 'STILLBIRTH', 'ABORTION', 'REFERRED_OUT',
                   'LOST_TO_FOLLOW_UP'];
    if (!want) return { success: false, message: 'Name the antenatal record.' };
    if (ALLOWED.indexOf(outcome) === -1) {
      return { success: false,
               message: 'Say how the pregnancy ended: ' + ALLOWED.join(', ') + '.' };
    }

    var sh = mc_ancSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (mc_up_(data[i][m['ANC_ID']]) !== want) continue;
      if (mc_up_(data[i][m['Status']]) !== 'OPEN') {
        return { success: false, message: 'That record is already closed.' };
      }
      var r = i + 1;
      sh.getRange(r, m['Status'] + 1).setValue('CLOSED');
      sh.getRange(r, m['Outcome'] + 1).setValue(outcome);
      sh.getRange(r, m['Outcome_Date'] + 1)
        .setValue(mc_date_(payload.outcomeDate) || new Date());
      sh.getRange(r, m['Outcome_Notes'] + 1).setValue(mc_str_(payload.notes));
      sh.getRange(r, m['Closed_By'] + 1).setValue(actor.username);
      sh.getRange(r, m['Closed_At'] + 1).setValue(new Date());
      dc_invalidate_(MC.ANC);
      SpreadsheetApp.flush();
      try {
        logAudit_({ username: actor.username, role: actor.role },
                  'ANC_CLOSED', 'Patient', mc_up_(data[i][m['Patient_ID']]),
                  { ancId: want, outcome: outcome });
      } catch (e) {}
      return { success: true, message: 'Antenatal record closed — ' +
               outcome.replace(/_/g, ' ').toLowerCase() + '.' };
    }
    return { success: false, message: 'No antenatal record with the id ' + want + '.' };
  } catch (err) {
    return { success: false,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ---------------------------------------------------------------------------
// IMMUNISATION — write endpoint
// ---------------------------------------------------------------------------

/**
 * FRONTEND ENTRY. Records a dose that was given.
 *
 * The register holds GIVEN doses only. Due dates come from the schedule and
 * the child's date of birth, so nothing has to be pre-created and a schedule
 * revision does not rewrite what already happened.
 *
 * @param {{patientId, vaccineCode, givenOn, batchNo, site, route, givenBy,
 *          notes}} payload
 */
function mcRecordImmunisation(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};
    var actor = crescRequire_(sessionToken, ['emr.write', 'ward.write']);

    var pid = mc_up_(payload.patientId);
    if (!pid) return { success: false, message: 'Name the child.' };

    var entry = mc_scheduleEntry_(payload.vaccineCode);
    // A dose outside the national schedule (a private vaccine, a catch-up
    // course) is allowed, with its own label, rather than refused — refusing
    // it would send the record back to paper for exactly the doses most worth
    // having written down.
    var code = entry ? entry.code : mc_up_(payload.vaccineCode);
    var label = entry ? entry.label : mc_str_(payload.doseLabel) || code;
    if (!code) return { success: false, message: 'Name the vaccine.' };

    var givenOn = mc_date_(payload.givenOn) || new Date();
    if (mc_days_(givenOn, new Date()) < 0) {
      return { success: false,
               message: 'That date is in the future. Record a dose when it has ' +
                        'been given, not before.' };
    }

    // Already recorded? A second identical row would make the schedule show a
    // dose twice and the count wrong.
    var sh = mc_immunSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (mc_up_(data[i][m['Patient_ID']]) !== pid) continue;
      if (mc_up_(data[i][m['Vaccine_Code']]) !== code) continue;
      return { success: false, code: 'ALREADY_GIVEN',
               message: label + ' is already recorded for ' + pid + ', given ' +
                        mc_fmt_(data[i][m['Given_On']]) + '.' };
    }

    var now = new Date();
    var row = new Array(sh.getLastColumn()).fill('');
    var put = function (h, v) { if (m[h] !== undefined) row[m[h]] = v; };
    put('Entry_ID', 'IMM-' + Utilities.formatDate(now, MC.TZ, 'yyMMdd-HHmmss') + '-' +
                    Utilities.getUuid().substring(0, 4).toUpperCase());
    put('Patient_ID', pid);
    put('Vaccine_Code', code);
    put('Dose_Label', label);
    put('Given_On', givenOn);
    put('Batch_No', mc_str_(payload.batchNo));
    put('Site', mc_str_(payload.site));
    put('Route', mc_str_(payload.route));
    put('Given_By', mc_str_(payload.givenBy) || actor.displayName || actor.username);
    put('Recorded_By', actor.username);
    put('Recorded_At', now);
    put('Notes', mc_str_(payload.notes));
    sh.appendRow(row);
    dc_invalidate_(MC.IMMUN);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'IMMUNISATION_GIVEN', 'Patient', pid,
                { vaccine: code, givenOn: mc_iso_(givenOn),
                  offSchedule: !entry });
    } catch (e) {}

    return { success: true, message: label + ' recorded, given ' + mc_fmt_(givenOn) + '.' };
  } catch (err) {
    return { success: false,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ---------------------------------------------------------------------------
// READERS shared by the clinic screens and the portal
// ---------------------------------------------------------------------------

/**
 * The antenatal picture for one patient. NO AUTHORISATION OF ITS OWN —
 * callers establish that first.
 *
 * @return {Object|null} null when there is no open pregnancy
 */
function mc_antenatalView_(patientId) {
  var anc = mc_openAnc_(patientId);
  if (!anc) return null;

  var visits = mc_ancVisits_(anc.ancId);
  var now = new Date();
  var g = mc_gestation_(anc.lmp, now);

  // The next appointment is whatever the LAST visit said to come back on. A
  // date that has passed is reported as overdue rather than hidden — it is
  // the single most useful thing on this panel.
  var last = visits.length ? visits[visits.length - 1] : null;
  var nextVisit = last ? last.nextVisit : null;
  var nextVisitIn = nextVisit ? mc_days_(now, nextVisit) : null;

  return {
    ancId: anc.ancId,
    lmp: mc_iso_(anc.lmp),
    lmpText: mc_fmt_(anc.lmp),
    edd: mc_iso_(anc.edd),
    eddText: mc_fmt_(anc.edd),
    gestationText: g ? g.text : '',
    gestationWeeks: g ? g.weeks : null,
    trimester: g ? (g.weeks < 14 ? 1 : (g.weeks < 28 ? 2 : 3)) : null,
    gravida: anc.gravida, para: anc.para,
    living: anc.living, abortions: anc.abortions,
    bloodGroup: anc.bloodGroup,
    riskNotes: anc.riskNotes,
    visitCount: visits.length,

    nextVisit: mc_iso_(nextVisit),
    nextVisitText: mc_fmt_(nextVisit),
    nextVisitInDays: nextVisitIn,
    nextVisitOverdue: (nextVisitIn !== null && nextVisitIn < 0),

    // ---- the trends ----------------------------------------------------
    // Gestational age on the x-axis, not the visit number: two visits three
    // weeks apart and two a fortnight apart are not the same interval, and a
    // growth curve plotted against visit number hides exactly the gap a
    // growth curve exists to show.
    trend: visits.map(function (v) {
      var vg = mc_gestation_(anc.lmp, v.date);
      return {
        date: mc_iso_(v.date),
        dateText: mc_fmt_(v.date, 'dd-MMM'),
        weeks: vg ? vg.weeks : null,
        weightKg: v.weightKg,
        sysBp: v.sysBp,
        diaBp: v.diaBp,
        haemoglobin: v.haemoglobin,
        fundalHeightCm: v.fundalHeightCm,
        fetalHeartRate: v.fetalHeartRate,
        presentation: v.presentation,
        urineAlbumin: v.urineAlbumin,
        urineSugar: v.urineSugar
      };
    }),

    lastVisit: last ? {
      date: mc_fmt_(last.date),
      weightKg: last.weightKg,
      bp: (last.sysBp && last.diaBp) ? (last.sysBp + '/' + last.diaBp) : '',
      haemoglobin: last.haemoglobin,
      fundalHeightCm: last.fundalHeightCm,
      fetalHeartRate: last.fetalHeartRate,
      presentation: last.presentation,
      advice: last.advice
    } : null
  };
}

/**
 * The immunisation picture for one child. NO AUTHORISATION OF ITS OWN.
 *
 * Returns null when there is no usable date of birth: every due date on this
 * panel is computed from it, and a schedule with no anchor would show a child
 * as overdue for everything from BCG onwards.
 */
function mc_immunisationView_(patientId, dob) {
  var birth = mc_date_(dob);
  if (!birth) return null;

  var ageDays = mc_days_(birth, new Date());
  if (ageDays === null || ageDays < 0) return null;

  var pid = mc_up_(patientId);
  var given = {};
  try {
    var sh = mc_immunSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (mc_up_(data[i][m['Patient_ID']]) !== pid) continue;
      given[mc_up_(data[i][m['Vaccine_Code']])] = {
        on: mc_date_(data[i][m['Given_On']]),
        label: mc_str_(data[i][m['Dose_Label']]),
        batch: mc_str_(data[i][m['Batch_No']]),
        by: mc_str_(data[i][m['Given_By']])
      };
    }
  } catch (e) { /* an unreadable register shows the schedule with nothing given */ }

  var doses = MC_SCHEDULE.map(function (v) {
    var due = new Date(birth.getTime() + v.dueDays * 86400000);
    var g = given[v.code];
    var inDays = mc_days_(new Date(), due);
    var state;
    if (g) state = 'GIVEN';
    else if (inDays > 0) state = 'UPCOMING';
    else if (-inDays <= v.window) state = 'DUE';
    else state = 'OVERDUE';

    return {
      code: v.code, label: v.label, at: v.at,
      due: mc_iso_(due), dueText: mc_fmt_(due),
      dueInDays: inDays,
      state: state,
      givenOn: g ? mc_iso_(g.on) : '',
      givenOnText: g ? mc_fmt_(g.on) : '',
      givenBy: g ? g.by : ''
    };
  });

  // Doses recorded that the national schedule does not contain — a private
  // vaccine, a catch-up course. Shown, because a parent looking at this list
  // should see everything their child has had, not only the government ones.
  var extra = [];
  Object.keys(given).forEach(function (code) {
    if (mc_scheduleEntry_(code)) return;
    extra.push({
      code: code, label: given[code].label || code, at: '',
      due: '', dueText: '', dueInDays: null, state: 'GIVEN',
      givenOn: mc_iso_(given[code].on), givenOnText: mc_fmt_(given[code].on),
      givenBy: given[code].by, offSchedule: true
    });
  });

  // "Due so far" is the honest denominator. Counting a newborn as 3 of 27 is
  // a progress bar that says a healthy baby is 11% vaccinated.
  var dueSoFar = doses.filter(function (d) { return d.state !== 'UPCOMING'; });
  var doneSoFar = dueSoFar.filter(function (d) { return d.state === 'GIVEN'; });
  var next = doses.filter(function (d) { return d.state !== 'GIVEN'; })
                  .sort(function (a, b) { return (a.dueInDays || 0) - (b.dueInDays || 0); });
  var overdue = doses.filter(function (d) { return d.state === 'OVERDUE'; });

  return {
    dob: mc_iso_(birth),
    dobText: mc_fmt_(birth),
    ageDays: ageDays,
    ageText: mc_ageText_(ageDays),
    doses: doses.concat(extra),
    // Ordered so the panel can show "next due" without sorting again.
    nextDue: next.length ? next[0] : null,
    nextDueGroup: next.filter(function (d) {
      return next.length && d.due === next[0].due;
    }),
    overdue: overdue,
    completedOfDue: doneSoFar.length,
    totalDue: dueSoFar.length,
    percentOfDue: dueSoFar.length
      ? Math.round((doneSoFar.length / dueSoFar.length) * 100) : 100,
    totalInSchedule: doses.length
  };
}

/** "8 months", "2 years 3 months", "11 days" — the unit a parent uses. */
function mc_ageText_(days) {
  if (days === null || days === undefined) return '';
  if (days < 31) return days + ' day' + (days === 1 ? '' : 's');
  var months = Math.floor(days / 30.44);
  if (months < 24) return months + ' month' + (months === 1 ? '' : 's');
  var years = Math.floor(months / 12);
  var rem = months % 12;
  return years + ' year' + (years === 1 ? '' : 's') +
         (rem ? ' ' + rem + ' month' + (rem === 1 ? '' : 's') : '');
}

// ---------------------------------------------------------------------------
// CLINIC-SIDE READ ENDPOINTS
// ---------------------------------------------------------------------------

/** FRONTEND ENTRY. One patient's antenatal record, for a clinician. */
function mcGetAntenatal(patientId, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['emr.read', 'ward.read']);
    var pid = mc_up_(patientId);
    if (!pid) return { success: false, message: 'Name the patient.' };
    try { dpdpLogRead_(actor, 'Patient', pid, { endpoint: 'mcGetAntenatal' }); } catch (e) {}
    var view = mc_antenatalView_(pid);
    return view
      ? { success: true, antenatal: view, message: '' }
      : { success: true, antenatal: null,
          message: 'No open antenatal record for ' + pid + '.' };
  } catch (err) {
    return { success: false,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}

/** FRONTEND ENTRY. One child's immunisation schedule and what is given. */
function mcGetImmunisation(patientId, sessionToken) {
  try {
    var actor = crescRequire_(sessionToken, ['emr.read', 'ward.read']);
    var pid = mc_up_(patientId);
    if (!pid) return { success: false, message: 'Name the child.' };
    try { dpdpLogRead_(actor, 'Patient', pid, { endpoint: 'mcGetImmunisation' }); } catch (e) {}

    var profile = null;
    try { profile = pt_readProfile_(pid); } catch (e) { profile = null; }
    if (!profile) return { success: false, message: 'No patient with the ID ' + pid + '.' };

    var view = mc_immunisationView_(pid, profile.dob);
    return view
      ? { success: true, immunisation: view, message: '' }
      : { success: true, immunisation: null,
          message: 'No usable date of birth on record for ' + pid +
                   ', so no due date can be worked out. Record the date of ' +
                   'birth on the patient file first.' };
  } catch (err) {
    return { success: false,
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}

/**
 * FRONTEND ENTRY. Every pregnancy currently open, soonest EDD first, with
 * whose next visit has passed.
 *
 * The worklist an ANC clinic runs on. Without it the register is only
 * readable one patient at a time, which means nobody notices the woman who
 * stopped coming.
 */
function mcListOpenAntenatal(sessionToken) {
  try {
    crescRequire_(sessionToken, ['emr.read', 'ward.read', 'appointment.read']);
    var sh = mc_ancSheet_();
    var m = dc_headerMap_(sh);
    var data = dc_sheetValues_(sh);
    var now = new Date();
    var rows = [];

    for (var i = 1; i < (data ? data.length : 0); i++) {
      if (mc_up_(data[i][m['Status']]) !== 'OPEN') continue;
      var pid = mc_up_(data[i][m['Patient_ID']]);
      var view = mc_antenatalView_(pid);
      if (!view) continue;
      rows.push({
        patientId: pid,
        name: mc_str_(data[i][m['Patient_Name']]),
        ancId: view.ancId,
        edd: view.eddText,
        eddIso: view.edd,
        gestationText: view.gestationText,
        trimester: view.trimester,
        visitCount: view.visitCount,
        nextVisitText: view.nextVisitText,
        nextVisitOverdue: view.nextVisitOverdue,
        nextVisitInDays: view.nextVisitInDays,
        // A pregnancy well past its EDD with no outcome is either a delivery
        // nobody recorded or a patient nobody followed up. Both need chasing,
        // and neither shows up anywhere else.
        pastEdd: (view.edd && mc_days_(view.edd, now) > 0),
        stale: (view.edd && mc_days_(view.edd, now) > MC.ANC_STALE_DAYS)
      });
    }
    rows.sort(function (a, b) { return String(a.eddIso).localeCompare(String(b.eddIso)); });

    return { success: true, rows: rows,
             message: rows.length ? '' : 'No antenatal record is open.' };
  } catch (err) {
    return { success: false, rows: [],
             message: String((err && err.message) || err).replace('FORBIDDEN: ', '') };
  }
}
