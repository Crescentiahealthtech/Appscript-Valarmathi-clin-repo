// ============================================================================
// Dose_Reference.gs  —  Crescentia HealthTech
// Per-drug dose reference ranges, on a sheet the clinic owns.
// ----------------------------------------------------------------------------
// WHY THIS EXISTS
//
// Dose_Calculator.html does correct arithmetic on whatever mg/kg figure is
// typed into it. That is the whole of its safety: it checks the ANSWER
// against a crude weight band, and it has nothing at all to say about
// whether the mg/kg figure ITSELF was right for the drug. 15 mg/kg of
// paracetamol and 15 mg/kg of digoxin are the same arithmetic and very
// different events.
//
// So the range belongs beside the drug, and the range has to be the
// clinic's own — a table baked into the script is a table nobody can
// correct on the morning they find it wrong. `Drug_Dose_Reference` is that
// table: one row per drug (optionally per indication), edited in the
// spreadsheet, read by the calculator.
//
// WHAT THIS IS NOT
//
// It is not a pharmacopoeia and it does not prescribe. It reports what the
// clinic wrote down, and the calculator says plainly when a drug is not in
// the table rather than implying a silent "within range". Nothing here ever
// blocks a prescription.
//
// SETUP
//   Run setupDoseReference() once from the Apps Script editor. It creates
//   the sheet, seeds a starter set, and is safe to re-run — existing rows
//   are never overwritten.
// ============================================================================

var DREF_SHEET = "Drug_Dose_Reference";
var DREF_HEADERS = [
  "Drug",                // the name a prescriber types (brand or generic)
  "Generic",             // what it actually is
  "Class",               // therapeutic class — also used by duplicate checking
  "Indication",          // blank = the general case
  "Route",
  "Basis",               // DAY or DOSE — what Min/Max per kg are expressed in
  "Min_Per_Kg",
  "Max_Per_Kg",
  "Usual_Per_Kg",        // what the calculator prefills
  "Max_Single_Dose_mg",
  "Max_Daily_Dose_mg",
  "Doses_Per_Day",
  "Adult_Dose",
  "Min_Age_Months",      // below this the drug is flagged, not blocked
  "Formulations",        // free text, e.g. "125 mg/5 ml; 250 mg tab"
  "Notes",
  "Source",
  "Status"               // ACTIVE / anything else = ignored
];

function dref_sheet_() {
  var sh = dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DREF_SHEET, DREF_HEADERS);
  sh.setFrozenRows(1);
  return sh;
}

/** '' for anything unusable. */
function dref_str_(v) { return (v === null || v === undefined) ? "" : String(v).trim(); }

/** A number, or null — so "no ceiling recorded" is distinguishable from 0. */
function dref_num_(v) {
  var s = dref_str_(v);
  if (!s) return null;
  var n = parseFloat(s.replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
}

/**
 * The lookup key for a drug name.
 *
 * Prescribers type brands, strengths and pack sizes into the same box:
 * "Augmentin 625", "T. PARACETAMOL 500mg", "Syp. Amoxyclav 228/5ml". The key
 * is the alphabetic stem, lower-cased, so all three reach the same row.
 */
function dref_key_(name) {
  return dref_str_(name)
    .toLowerCase()
    .replace(/\b(tab|tabs|tablet|cap|caps|capsule|syp|syrup|susp|suspension|inj|injection|drops?|oint|cream|neb|nebuliser|nebulizer|resp(?:ule|ules)?)\b\.?/g, " ")
    .replace(/[0-9]+(\.[0-9]+)?\s*(mg|mcg|g|ml|iu|%)?/g, " ")
    .replace(/[^a-z ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Every ACTIVE reference row, keyed for lookup. Read once per execution.
 * @return {{rows:Array<Object>, byKey:Object}}
 */
function dref_all_() {
  var out = { rows: [], byKey: {} };
  var sh;
  try { sh = dref_sheet_(); } catch (e) { return out; }

  var values = dc_sheetValues_(sh);
  if (!values || values.length < 2) return out;

  var hdr = {};
  for (var c = 0; c < values[0].length; c++) hdr[dref_str_(values[0][c])] = c;
  function g(row, name) { return hdr[name] === undefined ? "" : row[hdr[name]]; }

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var drug = dref_str_(g(row, "Drug"));
    if (!drug) continue;
    if (dref_str_(g(row, "Status")).toUpperCase() !== "ACTIVE") continue;

    var rec = {
      drug:        drug,
      generic:     dref_str_(g(row, "Generic")),
      klass:       dref_str_(g(row, "Class")),
      indication:  dref_str_(g(row, "Indication")),
      route:       dref_str_(g(row, "Route")),
      basis:       (dref_str_(g(row, "Basis")).toUpperCase() === "DOSE") ? "dose" : "day",
      minPerKg:    dref_num_(g(row, "Min_Per_Kg")),
      maxPerKg:    dref_num_(g(row, "Max_Per_Kg")),
      usualPerKg:  dref_num_(g(row, "Usual_Per_Kg")),
      maxSingle:   dref_num_(g(row, "Max_Single_Dose_mg")),
      maxDaily:    dref_num_(g(row, "Max_Daily_Dose_mg")),
      dosesPerDay: dref_num_(g(row, "Doses_Per_Day")),
      adultDose:   dref_str_(g(row, "Adult_Dose")),
      minAgeMonths: dref_num_(g(row, "Min_Age_Months")),
      formulations: dref_str_(g(row, "Formulations")),
      notes:       dref_str_(g(row, "Notes")),
      source:      dref_str_(g(row, "Source"))
    };
    out.rows.push(rec);

    // Both the written name and the generic reach the row, so "Crocin" and
    // "Paracetamol" answer the same. First writer wins: a row keyed on its
    // own Drug column is never displaced by another row's Generic.
    [rec.drug, rec.generic].forEach(function (n) {
      var k = dref_key_(n);
      if (k && !out.byKey[k]) out.byKey[k] = rec;
    });
  }
  return out;
}

// ---------------------------------------------------------------------------
// FRONTEND ENTRY POINTS
// ---------------------------------------------------------------------------

/**
 * The reference row for one drug, or a miss that says so.
 *
 * A MISS IS AN ANSWER. The calculator shows "no reference on file for this
 * drug" rather than nothing, because a blank strip is indistinguishable
 * from "checked, and fine".
 *
 * @param {string} drugName  whatever is in the prescription's name box
 * @return {{success:boolean, found:boolean, reference:Object, message:string}}
 */
function getDoseReference(drugName, sessionToken) {
  try {
    crescRequire_(sessionToken, 'reference.read');
    var key = dref_key_(drugName);
    if (!key) {
      return { success: true, found: false, reference: null,
               message: "Type a drug name to see its reference range." };
    }
    var all = dref_all_();
    var hit = all.byKey[key];

    // A near miss is worth catching: "amoxycillin" vs "amoxicillin" is a
    // spelling difference, not a different drug. Only a containment match
    // either way, so unrelated drugs never collide.
    if (!hit) {
      var keys = Object.keys(all.byKey);
      for (var i = 0; i < keys.length; i++) {
        if (keys[i].length < 4) continue;
        if (key.indexOf(keys[i]) !== -1 || keys[i].indexOf(key) !== -1) {
          hit = all.byKey[keys[i]];
          break;
        }
      }
    }

    if (!hit) {
      return { success: true, found: false, reference: null,
               message: "No dose reference on file for \"" + dref_str_(drugName) +
                        "\". Add it to the " + DREF_SHEET + " sheet to have it " +
                        "checked here in future." };
    }
    return { success: true, found: true, reference: hit, message: "" };
  } catch (err) {
    return { success: false, found: false, reference: null,
             message: "Dose reference lookup failed: " + err.message };
  }
}

/**
 * The whole table, for the picker and for review.
 * @return {{success:boolean, count:number, rows:Array<Object>, message:string}}
 */
function listDoseReference() {
  try {
    var all = dref_all_();
    return { success: true, count: all.rows.length, rows: all.rows, message: "" };
  } catch (err) {
    return { success: false, count: 0, rows: [], message: err.message };
  }
}

// ---------------------------------------------------------------------------
// SETUP
// ---------------------------------------------------------------------------

/**
 * ONE-OFF. Creates the sheet and seeds a starter set. Safe to re-run: a drug
 * already in the sheet is left exactly as the clinic has edited it.
 *
 * THE SEED IS A STARTING POINT, NOT AN AUTHORITY. Every row carries
 * Source = "SEED — verify" so that a row nobody has reviewed is visibly
 * different from one the clinic has signed off. Review it against your own
 * formulary and your own product labels before relying on it.
 *
 * Columns are ordered to match DREF_HEADERS.
 */
function setupDoseReference() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sh = dref_sheet_();
    var values = dc_sheetValues_(sh);

    var have = {};
    for (var i = 1; i < values.length; i++) {
      var k = dref_key_(values[i][0]);
      if (k) have[k] = true;
    }

    // Drug, Generic, Class, Indication, Route, Basis, Min/kg, Max/kg,
    // Usual/kg, MaxSingle, MaxDaily, Doses/day, Adult, MinAgeMonths,
    // Formulations, Notes
    var seed = [
      ["Paracetamol", "Paracetamol", "Analgesic / Antipyretic", "", "Oral", "DOSE",
       10, 15, 15, 1000, 4000, 4, "500-1000 mg 4-6 hourly", 0,
       "125 mg/5 ml; 250 mg/5 ml; 500 mg tab",
       "Maximum 75 mg/kg/day in children. Reduce in hepatic impairment."],

      ["Ibuprofen", "Ibuprofen", "NSAID", "", "Oral", "DOSE",
       5, 10, 10, 400, 1200, 3, "400 mg 8 hourly", 6,
       "100 mg/5 ml; 200/400 mg tab",
       "With food. Avoid in dehydration, renal impairment, active bleeding, dengue."],

      ["Amoxicillin", "Amoxicillin", "Antibiotic - Penicillin", "", "Oral", "DAY",
       25, 50, 40, 1000, 3000, 3, "500 mg 8 hourly", 0,
       "125 mg/5 ml; 250 mg/5 ml; 250/500 mg cap",
       "Higher end (80-90 mg/kg/day) for otitis media and pneumonia."],

      ["Amoxicillin-Clavulanate", "Amoxicillin + Clavulanic acid", "Antibiotic - Penicillin", "", "Oral", "DAY",
       25, 45, 40, 875, 2625, 2, "625 mg 8 hourly or 1 g 12 hourly", 0,
       "228.5 mg/5 ml; 457 mg/5 ml; 625 mg tab",
       "Dose by the amoxicillin component. Diarrhoea is common."],

      ["Azithromycin", "Azithromycin", "Antibiotic - Macrolide", "", "Oral", "DAY",
       10, 12, 10, 500, 500, 1, "500 mg once daily", 6,
       "200 mg/5 ml; 250/500 mg tab",
       "3-5 day course. QT prolongation; check interactions."],

      ["Cefixime", "Cefixime", "Antibiotic - Cephalosporin", "", "Oral", "DAY",
       8, 10, 8, 400, 400, 2, "200 mg 12 hourly", 6,
       "50 mg/5 ml; 100 mg/5 ml; 200 mg tab", "Renal dose adjustment required."],

      ["Ceftriaxone", "Ceftriaxone", "Antibiotic - Cephalosporin", "", "IV/IM", "DAY",
       50, 75, 50, 2000, 4000, 1, "1-2 g once daily", 0,
       "250/500/1000 mg vial",
       "100 mg/kg/day for meningitis. Do NOT mix with calcium-containing fluids in neonates."],

      ["Ondansetron", "Ondansetron", "Antiemetic", "", "Oral/IV", "DOSE",
       0.1, 0.15, 0.15, 8, 24, 3, "4-8 mg 8 hourly", 1,
       "2 mg/5 ml; 4 mg tab / MD tab",
       "QT prolongation. Avoid with other QT-prolonging drugs."],

      ["Salbutamol", "Salbutamol", "Bronchodilator - SABA", "Nebulisation", "Nebulisation", "DOSE",
       0.1, 0.15, 0.15, 5, 20, 4, "2.5-5 mg per nebulisation", 0,
       "2.5 mg/2.5 ml respule; 5 mg/ml solution",
       "Under 5 years 2.5 mg; over 5 years 2.5-5 mg per nebulisation. Tachycardia, tremor."],

      ["Ipratropium", "Ipratropium bromide", "Bronchodilator - Anticholinergic", "Nebulisation", "Nebulisation", "DOSE",
       null, null, null, 0.5, 2, 4, "500 mcg per nebulisation", 0,
       "250 mcg/ml; 500 mcg/2 ml respule",
       "Under 12 years 250 mcg; 12 and over 500 mcg. Dosed per nebulisation, not per kg."],

      ["Budesonide", "Budesonide", "Corticosteroid - Inhaled", "Nebulisation", "Nebulisation", "DOSE",
       null, null, null, 1, 2, 2, "0.5-1 mg twice daily", 0,
       "0.5 mg/2 ml; 1 mg/2 ml respule",
       "Rinse the mouth after. Dosed per nebulisation, not per kg."],

      ["Prednisolone", "Prednisolone", "Corticosteroid - Oral", "", "Oral", "DAY",
       1, 2, 1, 60, 60, 1, "40-60 mg once daily", 0,
       "5 mg tab; 15 mg/5 ml",
       "Short courses. Taper if given beyond two weeks."],

      ["Hydrocortisone", "Hydrocortisone", "Corticosteroid - IV", "", "IV", "DOSE",
       2, 4, 4, 200, 800, 4, "100-200 mg 6 hourly", 0,
       "100 mg vial", "Acute asthma, anaphylaxis, adrenal crisis."],

      ["Adrenaline", "Adrenaline (Epinephrine)", "Vasopressor", "Anaphylaxis", "IM", "DOSE",
       0.01, 0.01, 0.01, 0.5, null, 1, "0.5 mg IM (0.5 ml of 1:1000)", 0,
       "1 mg/ml (1:1000) ampoule",
       "0.01 mg/kg IM, anterolateral thigh, repeat at 5-15 minutes. THIS IS mg, NOT ml."],

      ["Furosemide", "Furosemide", "Diuretic - Loop", "", "Oral/IV", "DOSE",
       0.5, 1, 1, 40, 160, 2, "20-40 mg 12 hourly", 0,
       "10 mg/ml injection; 40 mg tab",
       "Monitor potassium, sodium and renal function."],

      ["Metformin", "Metformin", "Antidiabetic - Biguanide", "", "Oral", "DAY",
       null, null, null, 1000, 2000, 2, "500-1000 mg twice daily", 120,
       "500/850/1000 mg tab",
       "With food. Withhold around contrast imaging. Not dosed per kg in adults."],

      ["Enalapril", "Enalapril", "Antihypertensive - ACE inhibitor", "", "Oral", "DAY",
       0.1, 0.5, 0.1, 20, 40, 2, "5-20 mg daily", 0,
       "2.5/5/10 mg tab",
       "Monitor potassium and creatinine. Contraindicated in pregnancy."],

      ["Amlodipine", "Amlodipine", "Antihypertensive - CCB", "", "Oral", "DAY",
       0.1, 0.3, 0.1, 10, 10, 1, "5-10 mg once daily", 72,
       "2.5/5/10 mg tab", "Ankle oedema is the common dose-limiting effect."],

      ["Pantoprazole", "Pantoprazole", "Proton pump inhibitor", "", "Oral/IV", "DAY",
       0.5, 1, 1, 40, 80, 1, "40 mg once daily", 60,
       "40 mg tab; 40 mg vial", "Before food."],

      ["Ranitidine", "Ranitidine", "H2 receptor antagonist", "", "Oral/IV", "DAY",
       2, 4, 4, 150, 300, 2, "150 mg twice daily", 0,
       "150 mg tab; 25 mg/ml injection",
       "Withdrawn in many markets — check local availability before prescribing."],

      ["Domperidone", "Domperidone", "Prokinetic", "", "Oral", "DOSE",
       0.2, 0.4, 0.25, 10, 30, 3, "10 mg 8 hourly", 0,
       "1 mg/ml suspension; 10 mg tab",
       "QT prolongation. Shortest effective course."],

      ["Cetirizine", "Cetirizine", "Antihistamine", "", "Oral", "DAY",
       0.2, 0.25, 0.25, 10, 10, 1, "10 mg once daily", 6,
       "5 mg/5 ml; 10 mg tab", "2-6 years 2.5-5 mg daily; 6 and over 5-10 mg daily."],

      ["Metronidazole", "Metronidazole", "Antibiotic - Nitroimidazole", "", "Oral/IV", "DAY",
       20, 30, 22.5, 750, 2250, 3, "400 mg 8 hourly", 0,
       "200 mg/5 ml; 200/400 mg tab", "Avoid alcohol during and for 48 hours after."],

      ["Ranferon", "Ferrous ascorbate + Folic acid", "Haematinic", "Iron deficiency", "Oral", "DAY",
       3, 6, 3, null, 200, 1, "100 mg elemental iron daily", 0,
       "Elemental iron 100 mg", "Dose by ELEMENTAL iron. On an empty stomach with vitamin C."],

      ["Vitamin D3", "Cholecalciferol", "Vitamin", "Deficiency", "Oral", "DAY",
       null, null, null, null, null, 1, "60,000 IU weekly for 8 weeks", 0,
       "60,000 IU sachet", "Dosed in IU, not mg/kg. Recheck the level after the course."]
    ];

    var toAppend = [];
    seed.forEach(function (r) {
      if (have[dref_key_(r[0])]) return;
      toAppend.push([
        r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10],
        r[11], r[12], r[13], r[14], r[15], "SEED — verify", "ACTIVE"
      ]);
    });

    if (toAppend.length) {
      sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, DREF_HEADERS.length)
        .setValues(toAppend);
      dc_invalidate_(DREF_SHEET);
    }
    SpreadsheetApp.flush();

    var msg = DREF_SHEET + ": " + toAppend.length + " drug(s) seeded, " +
              (seed.length - toAppend.length) + " already present. " +
              "Every seeded row is marked \"SEED — verify\" — review it against " +
              "your own formulary and change Source once you have.";
    Logger.log(msg);
    return msg;

  } catch (e) {
    return "setupDoseReference failed: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Which drugs the clinic actually prescribes but has no reference row for.
 * Run occasionally: it turns "the calculator never checks anything" into a
 * finite list of rows to add.
 */
function auditDoseReferenceCoverage() {
  try {
    var known = dref_all_().byKey;
    var seen = {}, missing = {};

    // The prescribing sheets, and the column each one keeps the drug name in.
    [["OP_Prescriptions", "Drug_Name"],
     ["IP_Medication_Queue", "Drug_Name"]].forEach(function (spec) {
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spec[0]);
      if (!sh) return;
      var values = dc_sheetValues_(sh);
      if (!values || values.length < 2) return;
      var col = -1;
      for (var c = 0; c < values[0].length; c++) {
        if (dref_str_(values[0][c]) === spec[1]) { col = c; break; }
      }
      if (col === -1) return;
      for (var i = 1; i < values.length; i++) {
        var name = dref_str_(values[i][col]);
        var k = dref_key_(name);
        if (!k) continue;
        seen[k] = true;
        if (!known[k]) missing[k] = (missing[k] || 0) + 1;
      }
    });

    var list = Object.keys(missing).sort(function (a, b) { return missing[b] - missing[a]; });
    var report = "Dose reference coverage: " +
      (Object.keys(seen).length - list.length) + " of " + Object.keys(seen).length +
      " prescribed drugs have a reference row.\n" +
      (list.length
        ? "Missing, most prescribed first:\n  " +
          list.slice(0, 40).map(function (k) { return k + "  (" + missing[k] + ")"; }).join("\n  ")
        : "Nothing missing.");
    Logger.log(report);
    return report;
  } catch (e) {
    return "auditDoseReferenceCoverage failed: " + e.message;
  }
}
