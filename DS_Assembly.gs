// ============================================================================
// DS_Assembly.gs  —  Crescentia HealthTech / CresRx
// IP Discharge Summary Engine · Phase 4 · deterministic clinical assembly
// ----------------------------------------------------------------------------
// NO AI IN THIS FILE. Every generated line traces to a source row, and every
// section carries the refs that produced it so the doctor can open the original
// note or result from a chip in the editor.
//
// WHERE THERE IS NO DATA, THE SECTION STAYS EMPTY.
// Never "N/A", never "Nil", never a placeholder sentence. A blank section is an
// honest statement that the ward did not record something; filler text is a
// lie that a doctor will sign.
//
// PERFORMANCE
//   Reads are bounded — TextFinder on one key column, then only the matched
//   rows. There is no getDataRange() on a clinical sheet anywhere in this file,
//   and nothing here takes the script lock.
// ============================================================================

var DSX_SRC = {
  ADMISSIONS: 'IP_Admissions',
  PATIENTS:   'Patients',
  CASESHEETS: 'IP_CaseSheets_DB',
  NOTES:      'IP_Timeline_DB',
  PHARMQ:     'IP_Pharmacy_Queue',
  INVENTORY:  'Pharmacy_Inventory',
  LAB_ORDERS: 'LAB_ORDERS',
  LAB_TESTS:  'LAB_ORDER_TESTS',
  LAB_RESULTS:'LAB_RESULTS'
};

// ---------------------------------------------------------------------------
// SECTION A — THE SOURCE BUNDLE
// ---------------------------------------------------------------------------

/** One admission row as an object, header-mapped. Used by DS_Workflow too. */
function dsx_admissionRow_(ipNumber) {
  var sh = dsx_ss_().getSheetByName(DSX_SRC.ADMISSIONS);
  if (!sh) return null;
  var row = dsx_findRowByKey_(sh, 'IP_Number', dsx_ip_(ipNumber));
  if (!row) return null;
  return dsx_readRow_(sh, row);
}

/**
 * Everything this admission has, read once. No writes, no lock.
 *
 * Malformed JSON never throws: it adds a SOURCE_PARSE_WARNING and the row is
 * skipped. A single corrupt note must not make a patient undischargeable.
 *
 * @return {{bundle:Object, warnings:Array<string>, sourceBundleHash:string, timings:Object}}
 */
function dsx_buildSourceBundle_(ipNumber) {
  var t0 = Date.now(), timings = {}, warnings = [];
  var ip = dsx_ip_(ipNumber);

  var mark = function (label, since) { timings[label] = Date.now() - since; };

  // -- admission ------------------------------------------------------------
  var s = Date.now();
  var admission = dsx_admissionRow_(ip);
  if (!admission) throw new Error('VALIDATION_FAILED: admission ' + ip + ' was not found.');
  var wb = ipa_resolveWardBed_(admission.Ward_Bed, admission.Bed);
  admission._ward = wb.ward;
  admission._bed = wb.bed;
  mark('admission', s);

  // -- patient --------------------------------------------------------------
  s = Date.now();
  var patient = null;
  var psh = dsx_ss_().getSheetByName(DSX_SRC.PATIENTS);
  if (psh) {
    var prow = dsx_findRowByKey_(psh, 'Patient_ID', dsx_pid_(admission.Patient_ID));
    if (prow) patient = dsx_readRow_(psh, prow);
  }
  if (!patient) warnings.push('Patient record ' + dsx_str_(admission.Patient_ID) + ' was not found.');
  mark('patient', s);

  // -- casesheets -----------------------------------------------------------
  s = Date.now();
  var casesheets = [];
  var csh = dsx_ss_().getSheetByName(DSX_SRC.CASESHEETS);
  if (csh) {
    dsx_findRowsByKey_(csh, 'IP_Number', ip).forEach(function (r) {
      var o = dsx_readRow_(csh, r);
      // Phase 7 amendment chain: a superseded casesheet is history, not truth.
      if (dsx_str_(o.Superseded_By)) return;
      casesheets.push(o);
    });
    casesheets.sort(function (a, b) {
      var da = dsx_toDate_(a.Timestamp), db = dsx_toDate_(b.Timestamp);
      return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
    });
  }
  if (!casesheets.length) warnings.push('No casesheet was found for this admission.');
  mark('casesheets', s);

  // -- notes ----------------------------------------------------------------
  s = Date.now();
  var notes = [];
  var nsh = dsx_ss_().getSheetByName(DSX_SRC.NOTES);
  if (nsh) {
    dsx_findRowsByKey_(nsh, 'IP_Number', ip).forEach(function (r) {
      var o = dsx_readRow_(nsh, r);
      var data = {};
      var raw = dsx_str_(o.Note_Data_JSON);
      if (raw) {
        try {
          data = JSON.parse(raw) || {};
        } catch (e) {
          warnings.push('SOURCE_PARSE_WARNING: note on row ' + r + ' (' +
                        dsx_fmt_(o.Timestamp, 'dd-MMM hh:mm a') + ') holds invalid JSON and was skipped.');
          return;
        }
      }
      notes.push({
        row: r,
        noteId: dsx_str_(o.Note_ID) || ('ROW-' + r),
        at: dsx_toDate_(o.Timestamp),
        roleType: dsx_upper_(o.Role_Type),
        // "Loading..." is a UI placeholder that was persisted on historic rows
        // (Discovery §11.5). Showing it on a discharge summary would be absurd.
        author: dsx_cleanAuthor_(o.Author),
        authorDoctorId: dsx_str_(o.Author_Doctor_ID),
        shift: dsx_str_(o.Shift),
        data: data
      });
    });
    notes.sort(function (a, b) {
      return (a.at ? a.at.getTime() : 0) - (b.at ? b.at.getTime() : 0);
    });
  }
  mark('notes', s);

  // -- pharmacy queue (cross-check only) ------------------------------------
  s = Date.now();
  var queue = [];
  var qsh = dsx_ss_().getSheetByName(DSX_SRC.PHARMQ);
  if (qsh) {
    dsx_findRowsByKey_(qsh, 'IP_Number', ip).forEach(function (r) {
      var o = dsx_readRow_(qsh, r);
      queue.push({
        queueId: dsx_str_(o.Queue_ID),
        drugName: dsx_str_(o.Drug_Name),
        dose: dsx_str_(o.Dose),
        freq: dsx_str_(o.Frequency),
        route: dsx_str_(o.Route),
        instructions: dsx_str_(o.Instructions),
        status: dsx_str_(o.Status),
        actionFlag: dsx_upper_(o.Action_Flag),
        orderedAt: dsx_toDate_(o.Ordered_At),
        modifiedAt: dsx_toDate_(o.Modified_At)
      });
    });
  }
  mark('pharmacyQueue', s);

  // -- labs ------------------------------------------------------------------
  s = Date.now();
  var labs = dsx_readLabs_(ip, warnings);
  mark('labs', s);

  var bundle = {
    ipNumber: ip,
    admission: admission,
    patient: patient,
    casesheets: casesheets,
    notes: notes,
    queue: queue,
    labs: labs
  };

  // The hash covers only clinical content, not row numbers or read order, so
  // regenerating an unchanged admission produces the same hash.
  var hashBasis = dsx_canonicalJson_({
    admission: dsx_pickForHash_(admission),
    casesheets: casesheets.map(dsx_pickForHash_),
    notes: notes.map(function (n) {
      return { at: n.at ? n.at.toISOString() : '', roleType: n.roleType, data: n.data };
    }),
    labs: labs
  });
  var sourceBundleHash = dsx_sha256Hex_(hashBasis);

  timings.total = Date.now() - t0;
  return { bundle: bundle, warnings: warnings, sourceBundleHash: sourceBundleHash, timings: timings };
}

function dsx_cleanAuthor_(v) {
  var a = dsx_str_(v);
  if (!a || /^loading\.*$/i.test(a)) return '';
  return a;
}

/** Row objects carry a _row key; hashing must ignore it. */
function dsx_pickForHash_(obj) {
  var out = {};
  Object.keys(obj || {}).forEach(function (k) {
    if (k.charAt(0) === '_') return;
    var v = obj[k];
    out[k] = (v instanceof Date) ? (isNaN(v.getTime()) ? '' : v.toISOString()) : v;
  });
  return out;
}

/**
 * Verified lab results for this admission.
 *
 * Join: LAB_ORDERS.AdmissionID -> OrderID -> LAB_ORDER_TESTS / LAB_RESULTS.
 * LAB_ORDER_TESTS and LAB_RESULTS carry no AdmissionID of their own.
 *
 * This deliberately does NOT reuse getIPLabResults() — that function reads the
 * dead Lab_Queue_DB sheet positionally and returns nothing (Discovery §4).
 */
function dsx_readLabs_(ipNumber, warnings) {
  var out = { orders: [], tests: [], results: [] };

  var osh = dsx_ss_().getSheetByName(DSX_SRC.LAB_ORDERS);
  if (!osh) return out;

  var orderIds = {};
  dsx_findRowsByKey_(osh, 'AdmissionID', dsx_ip_(ipNumber)).forEach(function (r) {
    var o = dsx_readRow_(osh, r);
    var id = dsx_str_(o.OrderID);
    if (!id) return;
    orderIds[dsx_upper_(id)] = true;
    out.orders.push({
      orderId: id,
      testNames: dsx_str_(o.TestNames),
      priority: dsx_str_(o.Priority),
      status: dsx_upper_(o.OrderStatus),
      orderedBy: dsx_str_(o.OrderingDoctorName),
      createdAt: dsx_toDate_(o.CreatedAt)
    });
  });
  if (!out.orders.length) return out;

  var tsh = dsx_ss_().getSheetByName(DSX_SRC.LAB_TESTS);
  if (tsh) {
    Object.keys(orderIds).forEach(function (oid) {
      dsx_findRowsByKey_(tsh, 'OrderID', oid).forEach(function (r) {
        var o = dsx_readRow_(tsh, r);
        out.tests.push({
          orderId: dsx_str_(o.OrderID),
          testId: dsx_str_(o.TestID),
          testName: dsx_str_(o.TestName),
          status: dsx_upper_(o.TestStatus),
          verifiedAt: dsx_toDate_(o.VerifiedAt),
          verifiedBy: dsx_str_(o.VerifiedBy)
        });
      });
    });
  }

  var rsh = dsx_ss_().getSheetByName(DSX_SRC.LAB_RESULTS);
  if (rsh) {
    Object.keys(orderIds).forEach(function (oid) {
      dsx_findRowsByKey_(rsh, 'OrderID', oid).forEach(function (r) {
        var o = dsx_readRow_(rsh, r);
        // Only the latest, verified, non-draft version may reach a signed
        // document. An amended result supersedes the one it replaced.
        if (dsx_bool_(o.IsDraft)) return;
        if (!dsx_toDate_(o.VerifiedAt)) return;
        var latest = dsx_str_(o.IsLatest);
        if (latest !== '' && !dsx_bool_(latest)) return;

        out.results.push({
          resultId: dsx_str_(o.ResultID),
          orderId: dsx_str_(o.OrderID),
          testId: dsx_str_(o.TestID),
          parameterName: dsx_str_(o.ParameterName) || dsx_str_(o.TestID),
          value: dsx_str_(o.ResultValue),
          unit: dsx_str_(o.Unit),
          refRange: dsx_str_(o.RefRangeText),
          flag: dsx_upper_(o.Flag),
          interpretation: dsx_str_(o.Interpretation),
          verifiedAt: dsx_toDate_(o.VerifiedAt),
          verifiedBy: dsx_str_(o.VerifiedBy)
        });
      });
    });
    out.results.sort(function (a, b) {
      return (a.verifiedAt ? a.verifiedAt.getTime() : 0) - (b.verifiedAt ? b.verifiedAt.getTime() : 0);
    });
  }
  return out;
}

// ---------------------------------------------------------------------------
// SECTION B — MEDICATION REPLAY
// ---------------------------------------------------------------------------

/**
 * Replays every medication order in chronological order and reports the state
 * of each drug at discharge.
 *
 * KEYING. Discovery §3 found no drug ID and no inventory ID anywhere in the IP
 * note path: a medication's only identity is the free-text brand in
 * `drugName`. The key is therefore `_normDrug_(drugName)` — the project's own
 * normaliser, which already handles "azithral tab 500" vs "500 Azithral Tab".
 * The generic name is resolved by an OPTIONAL lookup into Pharmacy_Inventory
 * and left EMPTY when the brand is not stocked. It is never guessed: printing
 * a wrong generic name on a discharge summary is a dispensing error waiting to
 * happen.
 *
 * Two different drugs are never silently merged — an ambiguous key raises a
 * conflict that becomes a SOFT readiness warning.
 *
 * @return {{drugs:Array, conflicts:Array, keying:string}}
 */
function dsx_replayMedications_(notes, admissionMeds) {
  var byKey = {};
  var conflicts = [];
  var order = [];

  var admitted = {};
  (admissionMeds || []).forEach(function (m) {
    var k = _normDrug_(m.drugName);
    if (k) admitted[k] = m;
  });

  notes.forEach(function (note) {
    var orders = (note.data && note.data.medOrders) || [];
    if (!orders.length) return;

    orders.forEach(function (o) {
      var name = dsx_str_(o.drugName || o.newDrugName);
      if (!name) return;
      var key = _normDrug_(name);
      if (!key) return;

      var action = dsx_upper_(o.action) || 'NEW';
      var at = note.at;

      if (!byKey[key]) {
        byKey[key] = {
          key: key,
          brand: name,
          names: [name],
          dose: dsx_str_(o.dose),
          route: dsx_str_(o.route),
          freq: dsx_str_(o.freq),
          instructions: dsx_str_(o.instructions),
          state: 'ACTIVE',
          startedAt: at,
          stoppedAt: null,
          changed: false,
          continuedFromAdmission: !!admitted[key],
          events: [],
          sourceRefs: []
        };
        order.push(key);
      }
      var d = byKey[key];

      // The same normalised key reached by two visibly different brand strings
      // is exactly the "P 500" class of bug. Surface it; never merge silently.
      if (d.names.indexOf(name) === -1) {
        d.names.push(name);
        conflicts.push('"' + d.names[0] + '" and "' + name + '" resolve to the same drug key — confirm they are the same medicine.');
      }

      d.events.push({ at: at, action: action, noteId: note.noteId });
      d.sourceRefs.push({ type: 'NOTE', id: note.noteId, at: at ? at.toISOString() : '' });

      switch (action) {
        case 'NEW':
          d.state = 'ACTIVE';
          if (!d.startedAt || (at && at < d.startedAt)) d.startedAt = at;
          d.stoppedAt = null;
          if (dsx_str_(o.dose)) d.dose = dsx_str_(o.dose);
          if (dsx_str_(o.route)) d.route = dsx_str_(o.route);
          if (dsx_str_(o.freq)) d.freq = dsx_str_(o.freq);
          if (dsx_str_(o.instructions)) d.instructions = dsx_str_(o.instructions);
          break;
        case 'MODIFY':
          d.state = 'ACTIVE';
          d.changed = true;
          d.stoppedAt = null;
          if (dsx_str_(o.dose)) d.dose = dsx_str_(o.dose);
          if (dsx_str_(o.route)) d.route = dsx_str_(o.route);
          if (dsx_str_(o.freq)) d.freq = dsx_str_(o.freq);
          if (dsx_str_(o.instructions)) d.instructions = dsx_str_(o.instructions);
          break;
        case 'HOLD':
          d.state = 'HELD';
          break;
        case 'STOP':
          d.state = 'STOPPED';
          d.stoppedAt = at;
          break;
        case 'CONT':
        default:
          if (d.state === 'STOPPED') {
            // A CONT after a STOP is contradictory. Take the later instruction
            // but say so, rather than quietly resurrecting a stopped drug.
            d.state = 'ACTIVE';
            d.stoppedAt = null;
            conflicts.push('"' + d.brand + '" was stopped and then continued — confirm the intended instruction.');
          }
          break;
      }
    });
  });

  var generics = dsx_genericIndex_();
  var drugs = order.map(function (k) {
    var d = byKey[k];
    d.generic = generics[k] || dsx_genericFuzzy_(generics, d.brand) || '';
    d.injectable = dsx_isInjectable_(d.route, d.brand);
    d.status = d.continuedFromAdmission ? (d.changed ? 'CHANGED' : 'CONTINUED')
             : (d.changed ? 'CHANGED' : 'NEW');
    return d;
  });

  return {
    drugs: drugs,
    conflicts: conflicts,
    keying: 'normalised brand text (_normDrug_) — no drug or inventory ID exists in IP notes'
  };
}

/** {normalisedBrandKey: genericName} from Pharmacy_Inventory. Best effort. */
function dsx_genericIndex_() {
  var idx = {};
  try {
    var sh = dsx_ss_().getSheetByName(DSX_SRC.INVENTORY);
    if (!sh || sh.getLastRow() < 2) return idx;
    var map = dsx_headerMap_(sh);
    var bCol = map['Brand Name'], gCol = map['Generic Name'];
    if (bCol === undefined || gCol === undefined) return idx;

    var n = sh.getLastRow() - 1;
    var brands = sh.getRange(2, bCol + 1, n, 1).getDisplayValues();
    var gens = sh.getRange(2, gCol + 1, n, 1).getDisplayValues();
    for (var i = 0; i < n; i++) {
      var b = dsx_str_(brands[i][0]);
      var g = dsx_str_(gens[i][0]);
      if (!b || !g) continue;
      var k = _normDrug_(b);
      if (k && !idx[k]) idx[k] = g.toUpperCase();
    }
  } catch (e) { /* generic resolution is optional; the brand always prints */ }
  return idx;
}

/** A single unambiguous inventory brand contained in the ordered name. */
function dsx_genericFuzzy_(index, brand) {
  var b = _normDrug_(brand);
  if (!b) return '';
  var words = b.split(' ');
  var hits = [];
  Object.keys(index).forEach(function (k) {
    if (!k) return;
    var kw = k.split(' ');
    var all = kw.every(function (w) { return words.indexOf(w) !== -1; });
    if (all && hits.indexOf(index[k]) === -1) hits.push(index[k]);
  });
  return hits.length === 1 ? hits[0] : '';
}

function dsx_isInjectable_(route, brand) {
  var r = dsx_upper_(route);
  if (/^(IV|IM|SC|S\/C|I\/V|I\/M|INJ|INTRAVENOUS|INTRAMUSCULAR|SUBCUTANEOUS)\b/.test(r)) return true;
  if (/INFUSION|INJECT/.test(r)) return true;
  return /\bINJ\b|INJECTION|\bIV\b/i.test(dsx_str_(brand));
}

// ---------------------------------------------------------------------------
// SECTION C — SECTION BUILDERS
// ---------------------------------------------------------------------------

function dsx_refsFrom_(type, ids, at) {
  return (ids || []).filter(String).map(function (id) {
    return { type: type, id: String(id), at: at || '' };
  });
}

/** Parses a casesheet JSON column into an array, tolerating plain text. */
function dsx_csArray_(v, warnings, label) {
  var raw = dsx_str_(v);
  if (!raw) return [];
  if (raw.charAt(0) !== '[' && raw.charAt(0) !== '{') return [{ condition: raw, duration: '' }];
  try {
    var parsed = JSON.parse(raw);
    return (Object.prototype.toString.call(parsed) === '[object Array]') ? parsed : [parsed];
  } catch (e) {
    if (warnings) warnings.push('SOURCE_PARSE_WARNING: casesheet ' + label + ' is not valid JSON; it was printed as text.');
    return [{ condition: raw, duration: '' }];
  }
}

/** "Cold — 2 Days" lines from a complaint/history array. */
function dsx_conditionLines_(arr) {
  return (arr || []).map(function (c) {
    var prefix = dsx_str_(c.prefix);
    var cond = dsx_str_(c.condition || c.text || c.name);
    var dur = dsx_str_(c.duration);
    if (!cond) return '';
    return (prefix ? prefix + ' ' : '') + cond + (dur ? ' — ' + dur : '');
  }).filter(String);
}

function dsx_buildSections_(bundle, dischargeType, warnings) {
  var sections = {};
  var adm = bundle.admission;
  var pat = bundle.patient || {};
  var cs = bundle.casesheets.length ? bundle.casesheets[bundle.casesheets.length - 1] : null;
  var csRefs = bundle.casesheets.map(function (c) { return dsx_str_(c.Encounter_ID); });

  var doa = dsx_toDate_(adm.DOA);
  var dod = dsx_toDate_(adm.DOD);

  // -- PATIENT_BANNER (read-only) -------------------------------------------
  sections.PATIENT_BANNER = dsx_newSection_('Patient', 'FIELDS', {
    name: dsx_str_(adm.Patient_Name) || dsx_str_(pat.Name),
    ageSex: dsx_str_(adm.Age_Sex) ||
            [dsx_str_(pat.Age), dsx_str_(pat.Gender)].filter(String).join(' / '),
    patientId: dsx_pid_(adm.Patient_ID),
    ipNumber: dsx_str_(adm.IP_Number),
    address: dsx_str_(pat.Address),
    mobile: dsx_str_(pat.Mobile),
    bloodGroup: dsx_str_(pat.Blood_Group)
  }, 'AUTO', dsx_refsFrom_('ADMISSION', [dsx_str_(adm.IP_Number)])
       .concat(dsx_refsFrom_('PATIENT', [dsx_pid_(adm.Patient_ID)])));

  // -- ADMISSION_DETAILS -----------------------------------------------------
  var los = (doa && (dod || new Date()))
    ? Math.max(1, Math.round(((dod || new Date()).getTime() - doa.getTime()) / 86400000))
    : '';
  var wardHistory = dsx_wardHistory_(bundle);
  sections.ADMISSION_DETAILS = dsx_newSection_('Admission details', 'FIELDS', {
    dateOfAdmission: dsx_dateTime_(doa, adm.TOA),
    dateOfDischarge: dsx_fmt_(dod, 'dd-MMM-yyyy'),
    lengthOfStay: los === '' ? '' : (los + ' day' + (los === 1 ? '' : 's')),
    admissionType: dsx_str_(adm.Admission_Type),
    wardBed: [adm._ward, adm._bed].filter(String).join(' / '),
    wardHistory: wardHistory,
    consultant: dsx_str_(adm.Consultant),
    referredBy: dsx_str_(pat.Referred_By)
  }, 'AUTO', dsx_refsFrom_('ADMISSION', [dsx_str_(adm.IP_Number)]));

  // -- DIAGNOSIS -------------------------------------------------------------
  var provisional = dsx_str_(adm.Diagnosis);
  var csDx = cs ? dsx_str_(cs['Primary Diagnosis']) : '';
  // The latest structured doctor assessment, if any, is the best candidate for
  // a final diagnosis. When there is none the field stays EMPTY for the doctor.
  var latestDx = dsx_latestDiagnosisFromNotes_(bundle.notes);
  sections.DIAGNOSIS = dsx_newSection_('Diagnosis', 'FIELDS', {
    provisional: provisional || csDx,
    final: latestDx.text,
    secondary: '',
    icd10: ''
  }, 'AUTO', dsx_refsFrom_('ADMISSION', [dsx_str_(adm.IP_Number)])
       .concat(dsx_refsFrom_('CASESHEET', csRefs))
       .concat(latestDx.refs));

  // -- PRESENTING_COMPLAINTS / HISTORY --------------------------------------
  sections.PRESENTING_COMPLAINTS = dsx_newSection_('Presenting complaints', 'LIST',
    cs ? dsx_conditionLines_(dsx_csArray_(cs.Chief_Complaints, warnings, 'Chief_Complaints')) : [],
    'AUTO', dsx_refsFrom_('CASESHEET', csRefs));

  sections.HISTORY = dsx_newSection_('History', 'LIST',
    cs ? dsx_conditionLines_(dsx_csArray_(cs.History, warnings, 'History')) : [],
    'AUTO', dsx_refsFrom_('CASESHEET', csRefs));

  // -- ALLERGIES -------------------------------------------------------------
  // The ONLY recorded source is Patients.Allergies (Discovery §2). When it is
  // blank the section stays blank and a HARD readiness item forces a clinician
  // to state NKDA or list the allergies.
  sections.ALLERGIES = dsx_newSection_('Allergies', 'TEXT',
    dsx_str_(pat.Allergies), 'AUTO',
    dsx_refsFrom_('PATIENT', [dsx_pid_(adm.Patient_ID)]));

  // -- EXAM_ON_ADMISSION -----------------------------------------------------
  // Vitals are one field PER READING, not a single "BP 120/80 · PR 88 · …"
  // string. A clinician correcting the admission pulse had to retype the
  // whole line and keep the separators right; and on screen the row now
  // lays out as a strip of small labelled boxes.
  sections.EXAM_ON_ADMISSION = dsx_newSection_('Examination on admission', 'FIELDS',
    cs ? dsx_merge_(dsx_vitalFields_(cs), {
      generalExamination: dsx_generalExam_(cs),
      cvs: dsx_str_(cs.CVS),
      rs: dsx_str_(cs.RS),
      pa: dsx_str_(cs.PA),
      cns: dsx_str_(cs.CNS)
    }) : {},
    'AUTO', dsx_refsFrom_('CASESHEET', csRefs));

  // -- INVESTIGATIONS / PENDING_RESULTS -------------------------------------
  var labView = dsx_significantLabs_(bundle.labs);
  sections.INVESTIGATIONS = dsx_newSection_('Significant investigations', 'TABLE',
    { columns: ['Test', 'First value', 'First on', 'Last value', 'Last on', 'Reference', 'Flag'],
      rows: labView.rows },
    'AUTO', labView.refs);

  sections.PENDING_RESULTS = dsx_newSection_('Results pending at discharge', 'LIST',
    labView.pending, 'AUTO', labView.pendingRefs);

  // -- PROCEDURES ------------------------------------------------------------
  var proc = dsx_procedures_(bundle.notes);
  sections.PROCEDURES = dsx_newSection_('Procedures', 'TABLE',
    { columns: ['Date', 'Procedure', 'Operator', 'Findings', 'Complications'], rows: proc.rows },
    'AUTO', proc.refs);

  // -- medication replay feeds three sections -------------------------------
  var admissionMeds = cs ? dsx_admissionMeds_(cs, warnings) : [];
  var replay = dsx_replayMedications_(bundle.notes, admissionMeds);

  var given = replay.drugs.map(function (d) {
    return [
      d.generic || '', d.brand, d.dose, d.route, d.freq,
      dsx_fmt_(d.startedAt, 'dd-MMM'),
      d.stoppedAt ? dsx_fmt_(d.stoppedAt, 'dd-MMM') : '',
      d.state
    ];
  });
  sections.TREATMENT_GIVEN = dsx_newSection_('Treatment given during the stay', 'TABLE',
    { columns: ['Generic', 'Brand', 'Dose', 'Route', 'Frequency', 'Started', 'Stopped', 'Status'],
      rows: given },
    'AUTO', dsx_refsFrom_('NOTE', bundle.notes.filter(function (n) {
      return n.data && n.data.medOrders && n.data.medOrders.length;
    }).map(function (n) { return n.noteId; })));

  // WRITTEN BY HAND, never prefilled.
  //
  // This table used to be seeded from the inpatient medication replay: every
  // active non-injectable drug arrived as a row. What a patient goes home on
  // is a fresh prescribing decision, not the ward chart minus the drips —
  // doses change, courses finish, and a row that is already on the page is a
  // row nobody re-reads. It starts empty and the drug column autocompletes
  // from pharmacy stock, exactly as prescribing does in OP and IP.
  //
  // Treatment given during the stay (above) still carries the full replay, so
  // nothing is lost from the record — it simply is not the discharge script.
  sections.DISCHARGE_MEDICATIONS = dsx_newSection_('Discharge medications', 'TABLE',
    { columns: ['Generic', 'Brand', 'Strength / Dose', 'Route', 'Frequency', 'Timing',
                'Food', 'Duration', 'Instructions', 'Status'],
      rows: [] },
    'MANUAL', []);
  sections.DISCHARGE_MEDICATIONS.reviewed = false;

  // "Medications stopped or held" is gone. It restated, in its own table,
  // what the Status column of Treatment given during the stay already says
  // for the same drugs, and it printed as a second list of drug names on a
  // document whose whole job is to be unambiguous about what the patient
  // takes home.

  // -- HOSPITAL_COURSE -------------------------------------------------------
  var course = dsx_hospitalCourse_(bundle, replay, doa);
  sections.HOSPITAL_COURSE = dsx_newSection_('Course in hospital', 'TEXT',
    course.text, 'AUTO', course.refs);

  // -- CONDITION_AT_DISCHARGE ------------------------------------------------
  var lastVitals = dsx_lastVitals_(bundle.notes, cs);
  sections.CONDITION_AT_DISCHARGE = dsx_newSection_('Condition at discharge', 'FIELDS',
    dsx_merge_({ generalCondition: '' },
      dsx_splitVitalText_(lastVitals.text),
      { vitalsRecordedAt: lastVitals.at ? dsx_fmt_(lastVitals.at, 'dd-MMM-yyyy hh:mm a') : '' }),
    'AUTO', lastVitals.refs);

  // -- advice / follow-up / red flags — empty, phrase library fills them ------
  // Manual, like the follow-up and the discharge script below it. It used to
  // inherit the ADMISSION casesheet's Advice field, which is advice given on
  // the way IN — carrying it to the summary put week-old instructions under a
  // "what to do at home" heading. The editor offers the clinic's advice
  // phrase library instead (ds_getPickers -> ADVICE), the same list OP and IP
  // prescribe from.
  sections.ADVICE = dsx_newSection_('Advice on discharge', 'FIELDS',
    { diet: '', activity: '', woundCare: '', other: '' }, 'MANUAL', []);

  sections.FOLLOW_UP = dsx_newSection_('Follow-up', 'FIELDS',
    { date: '', doctorOrDepartment: '', investigations: '', notRequired: '' }, 'MANUAL', []);

  sections.RED_FLAGS = dsx_newSection_('When to seek urgent care', 'LIST', [], 'AUTO', []);

  // -- type-specific ---------------------------------------------------------
  var typeSection = dsx_typeSpecificSection_(dischargeType);
  if (typeSection) sections[typeSection.key] = typeSection.section;

  // -- SIGNATURES (read-only; rendered from workflow data) -------------------
  sections.SIGNATURES = dsx_newSection_('Signatures', 'FIELDS', {}, 'AUTO', []);

  return { sections: sections, replay: replay, labView: labView };
}

function dsx_typeSpecificSection_(dischargeType) {
  switch (dsx_upper_(dischargeType)) {
    case 'LAMA':
    case 'DAMA':
      return { key: 'LAMA_DETAILS', section: dsx_newSection_(
        'Discharge against medical advice', 'FIELDS',
        { risksExplained: '', explainedBy: '', attendantName: '',
          attendantRelation: '', declarationSigned: '' }, 'AUTO', []) };
    case 'REFERRED':
      return { key: 'REFERRAL_DETAILS', section: dsx_newSection_(
        'Referral details', 'FIELDS',
        { referredTo: '', reason: '', conditionAtTransfer: '', modeOfTransport: '' }, 'AUTO', []) };
    case 'DEATH':
      return { key: 'DEATH_DETAILS', section: dsx_newSection_(
        'Death details', 'FIELDS',
        { dateTimeOfDeath: '', immediateCause: '', antecedentCause: '',
          underlyingCause: '', resuscitationSummary: '' }, 'AUTO', []) };
    case 'ABSCONDED':
      return { key: 'ABSCOND_DETAILS', section: dsx_newSection_(
        'Absconding details', 'FIELDS',
        { lastSeenAt: '', informedTo: '', policeIntimation: '' }, 'AUTO', []) };
    default:
      return null;
  }
}

/** The printed document title changes with the discharge type. */
function dsx_documentTitle_(dischargeType) {
  switch (dsx_upper_(dischargeType)) {
    case 'LAMA': return 'LAMA Summary';
    case 'DAMA': return 'Discharge Against Medical Advice Summary';
    case 'REFERRED': return 'Referral Summary';
    case 'DEATH': return 'Death Summary';
    case 'ABSCONDED': return 'Absconding Summary';
    default: return 'Discharge Summary';
  }
}

// ---- small builders --------------------------------------------------------

function dsx_wardHistory_(bundle) {
  var seen = [], out = [];
  var push = function (ward, bed, at) {
    var label = [ward, bed].filter(String).join(' / ');
    if (!label || seen.indexOf(label) !== -1) return;
    seen.push(label);
    out.push(label + (at ? ' (' + dsx_fmt_(at, 'dd-MMM') + ')' : ''));
  };
  bundle.casesheets.forEach(function (c) {
    push(dsx_str_(c.Ward), dsx_str_(c.Bed), dsx_toDate_(c.Timestamp));
  });
  push(bundle.admission._ward, bundle.admission._bed, null);
  return out.join(' → ');
}

/** Shallow-merges its arguments left to right into a new object. */
function dsx_merge_() {
  var out = {};
  for (var i = 0; i < arguments.length; i++) {
    var o = arguments[i] || {};
    Object.keys(o).forEach(function (k) { out[k] = o[k]; });
  }
  return out;
}

/**
 * The seven vitals as SEPARATE fields, in the order a chart records them.
 *
 * Keys carry their unit so the printed label reads "BP (mmHg)" without the
 * renderer knowing anything about vitals, and so an empty box still says
 * what belongs in it.
 */
function dsx_vitalFields_(cs) {
  var bp = [dsx_str_(cs.Sys_BP), dsx_str_(cs.Dia_BP)].filter(String).join('/');
  return {
    bpMmHg:       bp,
    pulsePerMin:  dsx_str_(cs.PR),
    spo2Percent:  dsx_str_(cs.SpO2),
    temperature:  dsx_str_(cs.Temp),
    respRatePerMin: dsx_str_(cs.RR),
    weightKg:     dsx_str_(cs.Weight),
    heightCm:     dsx_str_(cs.Height)
  };
}

/**
 * Splits a "BP 120/80 · PR 88 · SpO2 97" line back into the same field names
 * dsx_vitalFields_ produces.
 *
 * The discharge vitals come from a nursing note, whose stored shape is a
 * single formatted line (dsx_lastVitals_). Rather than reach back into the
 * note's JSON from here — it has several historic key spellings — the line is
 * parsed once, so the two vitals blocks on the document present the identical
 * set of boxes. Anything unrecognised is kept in `otherVitals` rather than
 * dropped, because losing a recorded observation is worse than an odd label.
 */
function dsx_splitVitalText_(text) {
  var out = { bpMmHg: '', pulsePerMin: '', spo2Percent: '', temperature: '',
              respRatePerMin: '', otherVitals: '' };
  var rest = [];
  dsx_str_(text).split(/\s*[·|,]\s*/).forEach(function (bit) {
    var s = dsx_str_(bit);
    if (!s) return;
    var m;
    if ((m = /^BP\s+(.+?)(?:\s*mmHg)?$/i.exec(s)))        { out.bpMmHg = dsx_str_(m[1]); return; }
    if ((m = /^(?:PR|Pulse|HR)\s+(.+?)(?:\s*\/min)?$/i.exec(s))) { out.pulsePerMin = dsx_str_(m[1]); return; }
    if ((m = /^SpO2\s+(.+?)%?$/i.exec(s)))                  { out.spo2Percent = dsx_str_(m[1]); return; }
    if ((m = /^Temp(?:erature)?\s+(.+)$/i.exec(s)))         { out.temperature = dsx_str_(m[1]); return; }
    if ((m = /^RR\s+(.+?)(?:\s*\/min)?$/i.exec(s)))         { out.respRatePerMin = dsx_str_(m[1]); return; }
    rest.push(s);
  });
  out.otherVitals = rest.join(' · ');
  return out;
}

function dsx_admissionVitals_(cs) {
  var bits = [];
  var bp = [dsx_str_(cs.Sys_BP), dsx_str_(cs.Dia_BP)].filter(String).join('/');
  if (bp) bits.push('BP ' + bp + ' mmHg');
  if (dsx_str_(cs.PR)) bits.push('PR ' + dsx_str_(cs.PR) + '/min');
  if (dsx_str_(cs.SpO2)) bits.push('SpO2 ' + dsx_str_(cs.SpO2) + '%');
  if (dsx_str_(cs.Temp)) bits.push('Temp ' + dsx_str_(cs.Temp));
  if (dsx_str_(cs.Weight)) bits.push('Weight ' + dsx_str_(cs.Weight) + ' kg');
  return bits.join(' · ');
}

function dsx_generalExam_(cs) {
  var signs = ['Pallor', 'Icterus', 'Cyanosis', 'Clubbing', 'Edema'];
  var present = signs.filter(function (k) { return /^(YES|Y|PRESENT|\+)$/.test(dsx_upper_(cs[k])); });
  var other = dsx_str_(cs['Other GE findings']);
  var parts = [];
  if (present.length) parts.push(present.join(', ') + ' present');
  else if (signs.some(function (k) { return dsx_str_(cs[k]); })) parts.push('No pallor, icterus, cyanosis, clubbing or oedema');
  if (other) parts.push(other);
  return parts.join('. ');
}

function dsx_admissionMeds_(cs, warnings) {
  var arr = dsx_csArray_(cs.Prescription_JSON, warnings, 'Prescription_JSON');
  return arr.map(function (m) {
    return {
      drugName: dsx_str_(m.drugName || m.brand || m.name),
      dose: dsx_str_(m.strength || m.dose),
      freq: dsx_str_(m.sig || m.freq),
      duration: dsx_str_(m.duration)
    };
  }).filter(function (m) { return !!m.drugName; });
}

/** The most recent structured assessment that reads like a diagnosis. */
function dsx_latestDiagnosisFromNotes_(notes) {
  for (var i = notes.length - 1; i >= 0; i--) {
    var n = notes[i];
    if (n.roleType !== 'DOCTOR' && n.roleType !== 'CONSULTANT') continue;
    var d = n.data || {};
    var text = dsx_str_(d.diagnosis || d.finalDiagnosis || d.assessment);
    if (text) {
      return { text: text, refs: [{ type: 'NOTE', id: n.noteId, at: n.at ? n.at.toISOString() : '' }] };
    }
  }
  return { text: '', refs: [] };
}

/**
 * Significant results: every abnormal value, plus the first and last of any
 * repeated test. A discharge summary that reprints forty normal CBCs is one
 * nobody reads.
 */
function dsx_significantLabs_(labs) {
  var byParam = {};
  (labs.results || []).forEach(function (r) {
    var k = dsx_upper_(r.parameterName);
    if (!k) return;
    (byParam[k] = byParam[k] || []).push(r);
  });

  var rows = [], refs = [];
  Object.keys(byParam).sort().forEach(function (k) {
    var series = byParam[k];
    var first = series[0], last = series[series.length - 1];
    var abnormal = series.some(function (r) { return r.flag && r.flag !== 'N' && r.flag !== 'NORMAL'; });
    var repeated = series.length > 1;
    if (!abnormal && !repeated) {
      // A single normal result is still worth one line — it is evidence the
      // test was done. It just does not get a first/last pair.
      rows.push([first.parameterName, dsx_valueWithUnit_(first), dsx_fmt_(first.verifiedAt, 'dd-MMM'),
                 '', '', first.refRange, first.flag]);
    } else {
      rows.push([first.parameterName,
                 dsx_valueWithUnit_(first), dsx_fmt_(first.verifiedAt, 'dd-MMM'),
                 repeated ? dsx_valueWithUnit_(last) : '',
                 repeated ? dsx_fmt_(last.verifiedAt, 'dd-MMM') : '',
                 last.refRange || first.refRange,
                 last.flag || first.flag]);
    }
    series.forEach(function (r) {
      refs.push({ type: 'LAB', id: r.resultId, at: r.verifiedAt ? r.verifiedAt.toISOString() : '' });
    });
  });

  // Pending = ordered but with no verified result.
  var verifiedTests = {};
  (labs.results || []).forEach(function (r) { verifiedTests[dsx_upper_(r.orderId) + '|' + dsx_upper_(r.testId)] = true; });
  var pending = [], pendingRefs = [];
  (labs.tests || []).forEach(function (t) {
    if (verifiedTests[dsx_upper_(t.orderId) + '|' + dsx_upper_(t.testId)]) return;
    if (t.verifiedAt) return;
    pending.push(t.testName + ' (ordered on ' + (t.orderId || '') + ')');
    pendingRefs.push({ type: 'LABORDER', id: t.orderId, at: '' });
  });

  return { rows: rows, refs: refs, pending: pending, pendingRefs: pendingRefs };
}

function dsx_valueWithUnit_(r) {
  return [dsx_str_(r.value), dsx_str_(r.unit)].filter(String).join(' ');
}

function dsx_procedures_(notes) {
  var rows = [], refs = [];
  notes.forEach(function (n) {
    if (n.roleType !== 'PROCEDURE') return;
    var d = n.data || {};
    var name = dsx_str_(d.procedureName || d.procedure);
    if (!name) return;
    rows.push([
      dsx_fmt_(n.at, 'dd-MMM-yyyy'),
      name,
      dsx_str_(d.operator) || n.author,
      dsx_str_(d.findings),
      dsx_str_(d.complications)
    ]);
    refs.push({ type: 'NOTE', id: n.noteId, at: n.at ? n.at.toISOString() : '' });
  });
  return { rows: rows, refs: refs };
}

/**
 * Day-wise digest. One paragraph per day the ward wrote something, built from
 * the doctor's own words plus the events of that day. Plain editable text —
 * the preparer is expected to tidy it, and the doctor to read it.
 */
function dsx_hospitalCourse_(bundle, replay, doa) {
  var days = {}, refs = [];
  var dayKey = function (d) { return d ? dsx_fmt_(d, 'yyyy-MM-dd') : ''; };
  var dayNo = function (d) {
    if (!doa || !d) return 0;
    return Math.max(1, Math.floor((d.getTime() - doa.getTime()) / 86400000) + 1);
  };
  var bucket = function (d) {
    var k = dayKey(d);
    if (!k) return null;
    if (!days[k]) days[k] = { at: d, no: dayNo(d), notes: [], procedures: [], meds: [], labs: [] };
    return days[k];
  };

  bundle.notes.forEach(function (n) {
    var b = bucket(n.at);
    if (!b) return;
    var d = n.data || {};

    if (n.roleType === 'DOCTOR' || n.roleType === 'CONSULTANT') {
      var text = dsx_str_(d.subjectiveObjective || d.assessment || d.plan || d.notes);
      if (text) {
        b.notes.push(text);
        refs.push({ type: 'NOTE', id: n.noteId, at: n.at ? n.at.toISOString() : '' });
      }
    }
    if (n.roleType === 'PROCEDURE') {
      var pname = dsx_str_(d.procedureName || d.procedure);
      if (pname) {
        b.procedures.push(pname);
        refs.push({ type: 'NOTE', id: n.noteId, at: n.at ? n.at.toISOString() : '' });
      }
    }
    (d.medOrders || []).forEach(function (o) {
      var name = dsx_str_(o.drugName || o.newDrugName);
      var action = dsx_upper_(o.action);
      if (!name || action === 'CONT') return;
      var verb = action === 'NEW' ? 'started' : action === 'STOP' ? 'stopped'
               : action === 'HOLD' ? 'held' : 'changed';
      b.meds.push(name + ' ' + verb);
    });
  });

  (bundle.labs.results || []).forEach(function (r) {
    if (!r.flag || r.flag === 'N' || r.flag === 'NORMAL') return;
    var b = bucket(r.verifiedAt);
    if (!b) return;
    b.labs.push(r.parameterName + ' ' + dsx_valueWithUnit_(r) + (r.flag ? ' (' + r.flag + ')' : ''));
  });

  var keys = Object.keys(days).sort();
  var paragraphs = keys.map(function (k) {
    var b = days[k];
    var bits = [];
    if (b.notes.length) bits.push(b.notes.join(' '));
    if (b.procedures.length) bits.push('Procedure: ' + b.procedures.join(', ') + '.');
    if (b.meds.length) bits.push('Medication: ' + dsx_unique_(b.meds).join(', ') + '.');
    if (b.labs.length) bits.push('Abnormal results: ' + dsx_unique_(b.labs).join('; ') + '.');
    if (!bits.length) return '';
    return 'Day ' + b.no + ' (' + dsx_fmt_(b.at, 'dd-MMM') + '): ' + bits.join(' ');
  }).filter(String);

  return { text: paragraphs.join('\n\n'), refs: refs };
}

function dsx_unique_(arr) {
  var seen = {}, out = [];
  (arr || []).forEach(function (v) { if (!seen[v]) { seen[v] = true; out.push(v); } });
  return out;
}

/** The most recent recorded vitals, from nursing notes first, casesheet last. */
function dsx_lastVitals_(notes, cs) {
  for (var i = notes.length - 1; i >= 0; i--) {
    var n = notes[i];
    var v = (n.data && (n.data.vitals || n.data.vitalSigns)) || null;
    if (!v) continue;
    var bits = [];
    if (dsx_str_(v.bp)) bits.push('BP ' + dsx_str_(v.bp));
    if (dsx_str_(v.pulse || v.pr)) bits.push('PR ' + dsx_str_(v.pulse || v.pr));
    if (dsx_str_(v.spo2)) bits.push('SpO2 ' + dsx_str_(v.spo2));
    if (dsx_str_(v.temp)) bits.push('Temp ' + dsx_str_(v.temp));
    if (dsx_str_(v.rr)) bits.push('RR ' + dsx_str_(v.rr));
    if (!bits.length) continue;
    return { text: bits.join(' · '), at: n.at,
             refs: [{ type: 'NOTE', id: n.noteId, at: n.at ? n.at.toISOString() : '' }] };
  }
  if (cs) {
    var text = dsx_admissionVitals_(cs);
    if (text) return { text: text, at: dsx_toDate_(cs.Timestamp),
                       refs: [{ type: 'CASESHEET', id: dsx_str_(cs.Encounter_ID), at: '' }] };
  }
  return { text: '', at: null, refs: [] };
}

/** "1-0-1" from a frequency string when it already is one; otherwise blank. */
function dsx_timingGrid_(freq) {
  var f = dsx_str_(freq);
  return /^\d+(\s*-\s*\d+){1,3}$/.test(f) ? f.replace(/\s+/g, '') : '';
}

// ---------------------------------------------------------------------------
// SECTION D — READINESS ENGINE
// ---------------------------------------------------------------------------

/**
 * HARD items block submission and signing. SOFT items are warnings the
 * preparer acknowledges. The server recomputes this at submit and at sign —
 * whatever the client believed is never trusted.
 *
 * @return {{hard:Array<{id,label,sectionKey}>, soft:Array<{id,label,sectionKey}>}}
 */
function dsx_readiness_(payload, bundle, header) {
  var hard = [], soft = [];
  var S = (payload && payload.sections) || {};
  var type = dsx_upper_(payload && payload.dischargeType);

  var need = function (list, id, label, sectionKey) {
    list.push({ id: id, label: label, sectionKey: sectionKey });
  };
  var field = function (key, name) {
    var sec = S[key];
    return (sec && sec.content) ? dsx_str_(sec.content[name]) : '';
  };

  // ---- HARD ---------------------------------------------------------------
  if (!field('DIAGNOSIS', 'final')) {
    need(hard, 'FINAL_DIAGNOSIS', 'A final diagnosis is required.', 'DIAGNOSIS');
  }
  if (dsx_sectionIsEmpty_(S.ALLERGIES)) {
    need(hard, 'ALLERGY_STATUS',
         'Allergy status must be stated — enter the allergies or "No known drug allergies".',
         'ALLERGIES');
  }
  var meds = S.DISCHARGE_MEDICATIONS;
  if (meds && !meds.reviewed) {
    need(hard, 'MEDS_REVIEWED',
         'Discharge medications must be reviewed and marked Reviewed (tick it even when there are none).',
         'DISCHARGE_MEDICATIONS');
  }
  if (!field('CONDITION_AT_DISCHARGE', 'generalCondition')) {
    need(hard, 'CONDITION', 'The general condition at discharge is required.', 'CONDITION_AT_DISCHARGE');
  }
  var fuNot = dsx_upper_(field('FOLLOW_UP', 'notRequired'));
  var fuGiven = field('FOLLOW_UP', 'date') || field('FOLLOW_UP', 'doctorOrDepartment');
  if (!fuGiven && fuNot !== 'TRUE' && fuNot !== 'YES') {
    need(hard, 'FOLLOW_UP', 'Give a follow-up plan, or tick "follow-up not required".', 'FOLLOW_UP');
  }

  // type-specific required fields
  var typeRules = {
    LAMA:  { key: 'LAMA_DETAILS', fields: ['risksExplained', 'explainedBy', 'attendantName'] },
    DAMA:  { key: 'LAMA_DETAILS', fields: ['risksExplained', 'explainedBy', 'attendantName'] },
    REFERRED: { key: 'REFERRAL_DETAILS', fields: ['referredTo', 'reason', 'conditionAtTransfer'] },
    DEATH: { key: 'DEATH_DETAILS', fields: ['dateTimeOfDeath', 'immediateCause'] },
    ABSCONDED: { key: 'ABSCOND_DETAILS', fields: ['lastSeenAt', 'informedTo'] }
  };
  var rule = typeRules[type];
  if (rule) {
    rule.fields.forEach(function (f) {
      if (!field(rule.key, f)) {
        need(hard, 'TYPE_' + f.toUpperCase(),
             dsx_humanise_(f) + ' is required for a ' + type + ' discharge.', rule.key);
      }
    });
  }

  // AI sections must be accepted before signing.
  Object.keys(S).forEach(function (k) {
    if (dsx_upper_(S[k].origin) === 'AI' && !dsx_str_(S[k].aiAcceptedBy)) {
      need(hard, 'AI_UNACCEPTED_' + k, 'AI-drafted text in ' + (S[k].title || k) +
           ' must be accepted or rejected.', k);
    }
  });

  // Size.
  try { dsx_packPayload_(JSON.stringify(payload)); }
  catch (e) { need(hard, 'PAYLOAD_SIZE', e.message, ''); }

  // ---- SOFT ---------------------------------------------------------------
  var vitalsAt = field('CONDITION_AT_DISCHARGE', 'vitalsRecordedAt');
  if (!vitalsAt) {
    need(soft, 'NO_VITALS', 'No vitals are recorded for this admission.', 'CONDITION_AT_DISCHARGE');
  } else {
    var d = dsx_toDate_(vitalsAt);
    if (d && (Date.now() - d.getTime()) > 24 * 3600 * 1000) {
      need(soft, 'STALE_VITALS', 'The last recorded vitals are more than 24 hours old.', 'CONDITION_AT_DISCHARGE');
    }
  }
  if (!dsx_sectionIsEmpty_(S.PENDING_RESULTS)) {
    need(soft, 'PENDING_RESULTS',
         'Some ordered investigations have no verified result yet.', 'PENDING_RESULTS');
  }
  if (payload && payload.assembly && payload.assembly.activeInjectables &&
      payload.assembly.activeInjectables.length) {
    need(soft, 'ACTIVE_INJECTABLES',
         'Injectables were still active at discharge and are not on the discharge list: ' +
         payload.assembly.activeInjectables.join(', ') + '.', 'DISCHARGE_MEDICATIONS');
  }
  Object.keys(S).forEach(function (k) {
    if (S[k].sourceChangedSinceEdit) {
      need(soft, 'SOURCE_CHANGED_' + k,
           (S[k].title || k) + ' was edited, and the source record has changed since.', k);
    }
  });
  var dup = dsx_duplicateGenerics_(S.DISCHARGE_MEDICATIONS);
  if (dup.length) {
    need(soft, 'DUPLICATE_GENERIC',
         'The same generic appears more than once in the discharge medications: ' + dup.join(', ') + '.',
         'DISCHARGE_MEDICATIONS');
  }
  if (payload && payload.assembly) {
    (payload.assembly.warnings || []).forEach(function (w, i) {
      need(soft, 'SOURCE_WARN_' + i, w, '');
    });
    (payload.assembly.medConflicts || []).forEach(function (w, i) {
      need(soft, 'MED_CONFLICT_' + i, w, 'DISCHARGE_MEDICATIONS');
    });
  }

  return { hard: hard, soft: soft };
}

/**
 * Field keys whose camel-case spelling does not humanise into something a
 * clinician would accept on a printed document. "bpMmHg" becomes
 * "Bp mm hg" otherwise, and the unit belongs in the label so an empty box
 * still says what goes in it.
 *
 * ds_getPickers() serves this same table to the editor, so the label over a
 * box on screen and the label beside it on paper cannot drift apart.
 */
var DSX_FIELD_LABELS = {
  bpMmHg:          'BP (mmHg)',
  pulsePerMin:     'Pulse (/min)',
  spo2Percent:     'SpO\u2082 (%)',
  temperature:     'Temperature',
  respRatePerMin:  'Resp. rate (/min)',
  weightKg:        'Weight (kg)',
  heightCm:        'Height (cm)',
  otherVitals:     'Other vitals',
  vitalsRecordedAt:'Vitals recorded at',
  generalCondition:'General condition',
  generalExamination: 'General examination',
  cvs: 'CVS', rs: 'RS', pa: 'P/A', cns: 'CNS',
  icd10: 'ICD-10',
  dateOfAdmission: 'Date of admission',
  dateOfDischarge: 'Date of discharge',
  lengthOfStay:    'Length of stay',
  admissionType:   'Admission type',
  wardBed:         'Ward / bed',
  wardHistory:     'Ward history',
  referredBy:      'Referred by',
  doctorOrDepartment: 'Doctor or department',
  notRequired:     'Not required',
  woundCare:       'Wound care',
  ageSex:          'Age / sex',
  patientId:       'Patient ID',
  ipNumber:        'IP number',
  bloodGroup:      'Blood group'
};

function dsx_humanise_(camel) {
  var mapped = DSX_FIELD_LABELS[camel];
  if (mapped) return mapped;
  var s = String(camel).replace(/([A-Z])/g, ' $1').toLowerCase().trim();
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function dsx_duplicateGenerics_(section) {
  if (!section || !section.content || !section.content.rows) return [];
  var seen = {}, dup = [];
  section.content.rows.forEach(function (r) {
    var g = dsx_upper_(r[0]);
    if (!g) return;
    if (seen[g] && dup.indexOf(g) === -1) dup.push(g);
    seen[g] = true;
  });
  return dup;
}

// ---------------------------------------------------------------------------
// SECTION E — ASSEMBLE AND REGENERATE
// ---------------------------------------------------------------------------

/**
 * Builds a complete payload for an admission. Runs outside the lock.
 * @return {{payload:Object, warnings:Array, timings:Object}}
 */
function dsx_assemble_(ipNumber, dischargeType, actor) {
  var t0 = Date.now();
  var built = dsx_buildSourceBundle_(ipNumber);
  var warnings = built.warnings;
  var tBundle = Date.now() - t0;

  var tSec = Date.now();
  var out = dsx_buildSections_(built.bundle, dischargeType, warnings);
  tSec = Date.now() - tSec;

  var activeInjectables = out.replay.drugs
    .filter(function (d) { return d.state === 'ACTIVE' && d.injectable; })
    .map(function (d) { return d.brand; });

  var payload = {
    schemaVersion: DSX_SCHEMA_VERSION,
    dischargeType: dsx_upper_(dischargeType) || 'NORMAL',
    documentTitle: dsx_documentTitle_(dischargeType),
    meta: {
      ipNumber: built.bundle.ipNumber,
      patientId: dsx_pid_(built.bundle.admission.Patient_ID),
      generatedAt: dsx_nowIso_(),
      sourceBundleHash: built.sourceBundleHash
    },
    // Engine-owned diagnostics. Not a section, so it never prints, but the
    // readiness engine and the UI both read it.
    assembly: {
      warnings: warnings,
      medConflicts: out.replay.conflicts,
      medKeying: out.replay.keying,
      activeInjectables: activeInjectables,
      noteCount: built.bundle.notes.length,
      resultCount: (built.bundle.labs.results || []).length
    },
    sections: out.sections
  };

  var check = dsx_validatePayload_(payload);
  if (!check.ok) throw new Error('VALIDATION_FAILED: ' + check.errors.join(' '));

  var timings = {
    bundleMs: tBundle,
    sectionsMs: tSec,
    totalMs: Date.now() - t0,
    bundleDetail: built.timings
  };
  Logger.log('DS assembly ' + ipNumber + ': ' + JSON.stringify(timings));

  return { payload: payload, warnings: warnings, timings: timings };
}

/**
 * Rebuilds the machine draft and merges it into the working copy.
 *
 * AUTO sections nobody has edited are replaced outright. Edited sections keep
 * the human's content and, when the fresh machine output differs from what the
 * machine produced last time, are flagged sourceChangedSinceEdit so the doctor
 * sees that the record moved under a line somebody had already corrected.
 * Human edits are never overwritten.
 */
function ds_regenerate(token, summaryId, expectedRowVersion) {
  var lock = LockService.getScriptLock();
  try {
    var actor = dsx_requireRole_(token, 'generate');

    var header0 = dsx_getHeader_(summaryId);
    if (!header0) return dsx_err_('VALIDATION_FAILED', 'No discharge summary found for ' + summaryId + '.');
    var t0 = dsx_transition_('REGENERATE', dsx_upper_(header0.Status));
    if (!t0.ok) return dsx_err_('INVALID_STATE', t0.message);

    var working0 = dsx_getWorking_(summaryId);
    if (!working0 || !working0.payload) {
      return dsx_err_('VALIDATION_FAILED', 'The working draft is missing.');
    }

    // Assembly outside the lock.
    var fresh = dsx_assemble_(dsx_str_(header0.IP_Number),
                              dsx_upper_(header0.Discharge_Type), actor);

    var merged = dsx_mergeRegenerated_(working0.payload, fresh.payload);

    lock.waitLock(DSX_LOCK_MS);
    dsx_resetHeaderCache_();

    var header = dsx_getHeader_(summaryId);
    dsx_checkVersion_(header, expectedRowVersion);
    var status = dsx_upper_(header.Status);
    var t = dsx_transition_('REGENERATE', status);
    if (!t.ok) return dsx_err_('INVALID_STATE', t.message);

    var snapshotNo = dsx_int_(header.Current_Snapshot_No) + 1;
    var hash = dsx_sha256Hex_(dsx_canonicalJson_(fresh.payload));

    dsx_appendSnapshot_(summaryId, snapshotNo, 'BASELINE', fresh.payload, hash, '', actor.username, '');
    dsx_putWorking_(summaryId, merged.payload, snapshotNo, actor.username);

    var rowVersion = dsx_bumpVersion_(dsx_summariesSheet_(), header, actor, {
      Current_Snapshot_No: snapshotNo
    });

    dsx_logEvent_(summaryId, actor, 'DS_REGENERATE', status, status, snapshotNo, hash, '',
                  { refreshed: merged.refreshed, flagged: merged.flagged, kept: merged.kept });
    SpreadsheetApp.flush();

    return dsx_ok_(
      'Regenerated. ' + merged.refreshed.length + ' section(s) refreshed, ' +
      merged.kept.length + ' edited section(s) kept' +
      (merged.flagged.length ? ', ' + merged.flagged.length + ' flagged as source-changed' : '') + '.',
      { rowVersion: rowVersion, snapshotNo: snapshotNo,
        refreshed: merged.refreshed, kept: merged.kept, flagged: merged.flagged,
        warnings: fresh.warnings, timings: fresh.timings });

  } catch (e) {
    return dsx_conflictOr_(e);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function dsx_mergeRegenerated_(current, fresh) {
  var refreshed = [], kept = [], flagged = [];
  var out = {
    schemaVersion: fresh.schemaVersion,
    dischargeType: fresh.dischargeType,
    documentTitle: fresh.documentTitle,
    meta: fresh.meta,
    assembly: fresh.assembly,
    sections: {}
  };
  if (current.signature) out.signature = current.signature;

  Object.keys(fresh.sections).forEach(function (k) {
    var f = fresh.sections[k];
    var c = current.sections[k];

    if (!c) { out.sections[k] = f; refreshed.push(k); return; }

    // A human edit is never overwritten, whatever the machine now says.
    if (c.edited || dsx_upper_(c.origin) === 'MANUAL' || dsx_upper_(c.origin) === 'AI') {
      c.sourceRefs = f.sourceRefs;
      if (dsx_canonicalJson_(f.content) !== dsx_canonicalJson_(c.content)) {
        c.sourceChangedSinceEdit = true;
        flagged.push(k);
      }
      out.sections[k] = c;
      kept.push(k);
      return;
    }

    // Untouched AUTO section: take the fresh one, but keep the review tick so a
    // regeneration does not silently un-review the medications.
    f.reviewed = c.reviewed;
    out.sections[k] = f;
    refreshed.push(k);
  });

  // Sections that exist only in the current draft (e.g. a type-specific block
  // from a type that has since changed) are kept rather than dropped.
  Object.keys(current.sections).forEach(function (k) {
    if (!out.sections[k]) { out.sections[k] = current.sections[k]; kept.push(k); }
  });

  return { payload: out, refreshed: refreshed, kept: kept, flagged: flagged };
}

// ---------------------------------------------------------------------------
// SECTION F — SOURCE LOOK-THROUGH
// ---------------------------------------------------------------------------

/**
 * The original record behind a source chip, scoped to this admission so a
 * crafted ref can never read another patient's note.
 */
function dsx_readSourceItem_(ipNumber, type, id) {
  var ip = dsx_ip_(ipNumber);

  if (type === 'NOTE') {
    var nsh = dsx_ss_().getSheetByName(DSX_SRC.NOTES);
    if (!nsh) return null;
    var rows = dsx_findRowsByKey_(nsh, 'Note_ID', id);
    if (!rows.length && /^ROW-\d+$/.test(id)) rows = [dsx_int_(id.substring(4))];
    for (var i = 0; i < rows.length; i++) {
      var o = dsx_readRow_(nsh, rows[i]);
      if (dsx_ip_(o.IP_Number) !== ip) continue;
      var data = {};
      try { data = JSON.parse(dsx_str_(o.Note_Data_JSON) || '{}'); } catch (e) {}
      return {
        type: 'NOTE',
        id: id,
        at: dsx_fmt_(o.Timestamp, 'dd-MMM-yyyy hh:mm a'),
        roleType: dsx_upper_(o.Role_Type),
        author: dsx_cleanAuthor_(o.Author),
        shift: dsx_str_(o.Shift),
        data: data
      };
    }
    return null;
  }

  if (type === 'LAB') {
    var rsh = dsx_ss_().getSheetByName(DSX_SRC.LAB_RESULTS);
    if (!rsh) return null;
    var rrow = dsx_findRowByKey_(rsh, 'ResultID', id);
    if (!rrow) return null;
    var r = dsx_readRow_(rsh, rrow);
    // Confirm the order belongs to this admission before returning anything.
    if (!dsx_orderBelongsToAdmission_(dsx_str_(r.OrderID), ip)) return null;
    return {
      type: 'LAB', id: id,
      testName: dsx_str_(r.ParameterName),
      value: dsx_str_(r.ResultValue), unit: dsx_str_(r.Unit),
      refRange: dsx_str_(r.RefRangeText), flag: dsx_str_(r.Flag),
      interpretation: dsx_str_(r.Interpretation),
      verifiedBy: dsx_str_(r.VerifiedBy),
      at: dsx_fmt_(r.VerifiedAt, 'dd-MMM-yyyy hh:mm a')
    };
  }

  if (type === 'CASESHEET') {
    var csh = dsx_ss_().getSheetByName(DSX_SRC.CASESHEETS);
    if (!csh) return null;
    var crow = dsx_findRowByKey_(csh, 'Encounter_ID', id);
    if (!crow) return null;
    var c = dsx_readRow_(csh, crow);
    if (dsx_ip_(c.IP_Number) !== ip) return null;
    return {
      type: 'CASESHEET', id: id,
      at: dsx_fmt_(c.Timestamp, 'dd-MMM-yyyy hh:mm a'),
      author: dsx_str_(c["Doctor's Name"]),
      diagnosis: dsx_str_(c['Primary Diagnosis']),
      complaints: dsx_str_(c.Chief_Complaints),
      history: dsx_str_(c.History),
      advice: dsx_str_(c.Advice)
    };
  }

  if (type === 'ADMISSION') {
    var a = dsx_admissionRow_(ip);
    if (!a) return null;
    return {
      type: 'ADMISSION', id: ip,
      doa: dsx_fmt_(a.DOA, 'dd-MMM-yyyy'),
      admissionType: dsx_str_(a.Admission_Type),
      consultant: dsx_str_(a.Consultant),
      diagnosis: dsx_str_(a.Diagnosis),
      status: dsx_upper_(a.Status)
    };
  }

  if (type === 'PATIENT') {
    var psh = dsx_ss_().getSheetByName(DSX_SRC.PATIENTS);
    if (!psh) return null;
    var adm = dsx_admissionRow_(ip);
    if (!adm || dsx_pid_(adm.Patient_ID) !== dsx_pid_(id)) return null;
    var prow = dsx_findRowByKey_(psh, 'Patient_ID', dsx_pid_(id));
    if (!prow) return null;
    var p = dsx_readRow_(psh, prow);
    return {
      type: 'PATIENT', id: dsx_pid_(id),
      name: dsx_str_(p.Name), age: dsx_str_(p.Age), gender: dsx_str_(p.Gender),
      address: dsx_str_(p.Address), mobile: dsx_str_(p.Mobile),
      allergies: dsx_str_(p.Allergies), conditions: dsx_str_(p.Conditions)
    };
  }

  return null;
}

function dsx_orderBelongsToAdmission_(orderId, ip) {
  var osh = dsx_ss_().getSheetByName(DSX_SRC.LAB_ORDERS);
  if (!osh) return false;
  var row = dsx_findRowByKey_(osh, 'OrderID', orderId);
  if (!row) return false;
  return dsx_ip_(dsx_readRow_(osh, row).AdmissionID) === ip;
}

// ---------------------------------------------------------------------------
// SECTION G — DRY RUN
// ---------------------------------------------------------------------------

/**
 * READ-ONLY. Assembles a real admission and logs what it produced without
 * writing a single cell. Run it from the Apps Script editor with a real IP
 * number before trusting the engine on live patients.
 */
function ds_dryRunAssembly(ipNumber) {
  var ip = dsx_ip_(ipNumber);
  if (!ip) {
    var msg = 'ds_dryRunAssembly: pass an IP number, e.g. ds_dryRunAssembly("IP2609-0002").';
    Logger.log(msg);
    return { success: false, message: msg };
  }

  try {
    var built = dsx_assemble_(ip, 'NORMAL', { username: 'DRY_RUN', role: 'doctor' });
    var readiness = dsx_readiness_(built.payload, null, null);
    var lines = [];

    lines.push('IP ' + ip + ' — ' + built.payload.documentTitle);
    lines.push('Source bundle hash: ' + built.payload.meta.sourceBundleHash.substring(0, 16) + '…');
    lines.push('Notes read: ' + built.payload.assembly.noteCount +
               ' · verified results: ' + built.payload.assembly.resultCount);
    lines.push('Medication keying: ' + built.payload.assembly.medKeying);
    lines.push('');
    lines.push('SECTIONS');

    Object.keys(built.payload.sections).forEach(function (k) {
      var s = built.payload.sections[k];
      var size;
      switch (s.format) {
        case 'TABLE':  size = s.content.rows.length + ' row(s)'; break;
        case 'LIST':   size = s.content.length + ' item(s)'; break;
        case 'FIELDS': size = Object.keys(s.content).filter(function (f) {
                         return dsx_str_(s.content[f]); }).length + ' field(s) filled'; break;
        default:       size = dsx_str_(s.content).length + ' char(s)'; break;
      }
      lines.push('  ' + (dsx_sectionIsEmpty_(s) ? '[empty] ' : '        ') +
                 k + ' (' + s.format + ') — ' + size +
                 ' · ' + s.sourceRefs.length + ' source ref(s)');
    });

    lines.push('');
    lines.push('READINESS — HARD (' + readiness.hard.length + ')');
    readiness.hard.forEach(function (h) { lines.push('  ✗ ' + h.label); });
    lines.push('READINESS — SOFT (' + readiness.soft.length + ')');
    readiness.soft.forEach(function (h) { lines.push('  ! ' + h.label); });

    lines.push('');
    lines.push('WARNINGS (' + built.warnings.length + ')');
    built.warnings.forEach(function (w) { lines.push('  · ' + w); });

    lines.push('');
    lines.push('TIMINGS ' + JSON.stringify(built.timings));
    lines.push('PAYLOAD ' + JSON.stringify(built.payload).length + ' characters of ' +
               (DSX_MAX_CHUNK * DSX_CHUNKS));

    var report = lines.join('\n');
    Logger.log(report);
    return { success: true, message: report,
             data: { readiness: readiness, timings: built.timings, warnings: built.warnings } };

  } catch (e) {
    var err = 'ds_dryRunAssembly failed for ' + ip + ': ' + e.message;
    Logger.log(err);
    return { success: false, message: err };
  }
}