// ============================================================================
// OP_Rx_Engine.gs  —  Crescentia HealthTech
// Prescription engine: seed library, Rx bundles, taper mode, safety guards,
// last-visit recall, complaint-driven workup, paediatric dose suggestions.
// ----------------------------------------------------------------------------
// REQUIRES: Doctor_Core.gs, Doctor_Session_Store.gs, OP_Doctor_Engine.gs,
//           OP_Templates_Engine.gs
//
// INTERACTION RULE (applies to every suggestion in this file)
//   Nothing here writes into a prescription. Bundles, learned suggestions,
//   last-visit meds and paediatric doses all return data for the UI to render
//   as a card with an explicit "+". The doctor's click is the act of
//   prescribing. One behaviour, learnt once.
// ============================================================================

var RX_MAX_TAPER_STEPS = 24;

// ============================================================================
// SECTION A — MIGRATION
// ============================================================================

/** Run once from the editor. Idempotent. */
function runRxEngineMigration() {
  var lock = LockService.getScriptLock();
  var out = [];
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var pat = ss.getSheetByName("Patients");
    if (pat) { dc_ensureColumn_(pat, "Allergies"); out.push("Patients: Allergies added."); }
    else out.push("Patients sheet absent — skipped.");

    var rx = ss.getSheetByName("Pharmacy_Queue_DB");
    if (rx) {
      dc_ensureColumn_(rx, "Computed_Qty");
      out.push("Pharmacy_Queue_DB: Computed_Qty added (tapers bill wrong without it).");
    }

    dc_ensureSheet_(ss, "Clinical_Seed_Library", [
      "Seed_ID", "Category", "Department", "Text", "Aliases",
      "Suggested_Labs", "Suggested_Dx", "Sort_Order", "Status"
    ]);
    dc_ensureSheet_(ss, "Rx_Bundles", [
      "Bundle_ID", "Tenant_ID", "Doctor_ID", "Scope", "Name", "Department",
      "Indication", "Items_JSON", "Use_Count", "Created_By",
      "Created_At", "Updated_At", "Status"
    ]);
    out.push("Clinical_Seed_Library / Rx_Bundles ready.");

    SpreadsheetApp.flush();
    out.push("--- Now run seedClinicalLibrary() ---");
    return out.join("\n");
  } catch (e) {
    return "RX MIGRATION FAILED: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

function rx_seedSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Clinical_Seed_Library", [
    "Seed_ID", "Category", "Department", "Text", "Aliases",
    "Suggested_Labs", "Suggested_Dx", "Sort_Order", "Status"
  ]);
}

function rx_bundleSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Rx_Bundles", [
    "Bundle_ID", "Tenant_ID", "Doctor_ID", "Scope", "Name", "Department",
    "Indication", "Items_JSON", "Use_Count", "Created_By",
    "Created_At", "Updated_At", "Status"
  ]);
}

/**
 * Seeds the starter phrase library. Skips anything already present, so it is
 * safe to re-run after you add rows to the array below.
 * Columns: Category | Department | Text | Aliases | Labs | Dx
 */
function seedClinicalLibrary() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var SEED = [
      // --- GENERAL MEDICINE : chief complaints -------------------------------
      ["CC","General Medicine","Fever","temperature, pyrexia, jvaram, thermal","CBC, Peripheral Smear, Dengue NS1","Acute Febrile Illness"],
      ["CC","General Medicine","Cough with expectoration","productive cough, sputum, phlegm, irumal","CBC, Chest X-Ray","Lower Respiratory Tract Infection"],
      ["CC","General Medicine","Dry cough","non-productive cough, tickling throat","",""],
      ["CC","General Medicine","Sore throat","throat pain, painful swallowing, odynophagia, tondai vali","Throat Swab","Acute Pharyngitis"],
      ["CC","General Medicine","Breathlessness","dyspnoea, SOB, short of breath, mookku adaippu","CBC, Chest X-Ray, ECG",""],
      ["CC","General Medicine","Chest pain","chest discomfort, angina, tightness","ECG, Troponin",""],
      ["CC","General Medicine","Headache","head pain, cephalgia, thalai vali","",""],
      ["CC","General Medicine","Giddiness","dizziness, vertigo, light headed","CBC, RBS, BP charting",""],
      ["CC","General Medicine","Vomiting","emesis, throwing up, thookkam","Serum Electrolytes",""],
      ["CC","General Medicine","Loose stools","diarrhoea, watery motion, loose motion","Stool Routine, Serum Electrolytes","Acute Gastroenteritis"],
      ["CC","General Medicine","Abdominal pain","stomach pain, belly pain, vayiru vali","CBC, USG Abdomen",""],
      ["CC","General Medicine","Generalised weakness","tiredness, fatigue, lethargy, weakness","CBC, RBS, TSH, Vitamin B12",""],
      ["CC","General Medicine","Swelling of legs","pedal edema, leg swelling, foot swelling","RFT, Urine Routine, Serum Albumin",""],
      ["CC","General Medicine","Joint pain","arthralgia, body pain, joint ache","CBC, ESR, RA Factor, Uric Acid",""],
      ["CC","General Medicine","Burning micturition","pain during micturition, burning urine, dysuria, painful urination","Urine Routine, Urine Culture","Urinary Tract Infection"],
      ["CC","General Medicine","Increased frequency of urination","frequent urination, polyuria","Urine Routine, FBS, PPBS",""],
      ["CC","General Medicine","Weight loss","losing weight, reduced weight","CBC, RBS, TSH, HIV",""],
      ["CC","General Medicine","Itching all over body","pruritus, generalised itching","CBC, LFT, RFT",""],

      // --- OBSTETRICS & GYNAECOLOGY ------------------------------------------
      ["CC","Obstetrics & Gynaecology","Abnormal heavy menstrual flow","menorrhagia, heavy periods, excessive bleeding","CBC, TSH, USG Pelvis","Abnormal Uterine Bleeding"],
      ["CC","Obstetrics & Gynaecology","Irregular menstrual cycles","irregular periods, oligomenorrhoea","TSH, Prolactin, USG Pelvis","PCOS for evaluation"],
      ["CC","Obstetrics & Gynaecology","White discharge per vaginum","leucorrhoea, vaginal discharge, white discharge","High Vaginal Swab","Vaginitis"],
      ["CC","Obstetrics & Gynaecology","Lower abdominal pain","pelvic pain, lower belly pain","USG Pelvis, Urine Routine",""],
      ["CC","Obstetrics & Gynaecology","Amenorrhoea","missed periods, no periods","Urine Pregnancy Test, TSH, Prolactin",""],
      ["CC","Obstetrics & Gynaecology","Dysmenorrhoea","painful periods, menstrual cramps","USG Pelvis",""],
      ["CC","Obstetrics & Gynaecology","Antenatal check up","ANC visit, pregnancy check, routine antenatal","CBC, Blood Grouping, Urine Routine, OGTT",""],
      ["CC","Obstetrics & Gynaecology","Infertility","unable to conceive, not conceiving","TSH, Prolactin, Semen Analysis, USG Pelvis",""],

      // --- PAEDIATRICS -------------------------------------------------------
      ["CC","Paediatrics","Fever with cold","febrile with coryza, running nose with fever","CBC","Viral Upper Respiratory Infection"],
      ["CC","Paediatrics","Refusal of feeds","not feeding, poor feeding, decreased intake","CBC, RBS",""],
      ["CC","Paediatrics","Loose stools in child","diarrhoea child, watery stools infant","Stool Routine, Serum Electrolytes","Acute Gastroenteritis with dehydration"],
      ["CC","Paediatrics","Not gaining weight","failure to thrive, poor weight gain","CBC, Thyroid Profile",""],
      ["CC","Paediatrics","Immunisation visit","vaccination, vaccine due, routine immunisation","",""],
      ["CC","Paediatrics","Wheezing","noisy breathing, whistling chest","Chest X-Ray","Reactive Airway Disease"],

      // --- DERMATOLOGY -------------------------------------------------------
      ["CC","Dermatology","Itchy skin lesions","rash with itching, pruritic rash","",""],
      ["CC","Dermatology","Hair fall","alopecia, hair loss, hair thinning","CBC, Ferritin, TSH",""],
      ["CC","Dermatology","Acne over face","pimples, face eruption","",""],

      // --- ENT ---------------------------------------------------------------
      ["CC","ENT","Ear pain","otalgia, earache, kadhu vali","","Acute Otitis Media"],
      ["CC","ENT","Reduced hearing","hearing loss, hard of hearing","Pure Tone Audiometry",""],
      ["CC","ENT","Nasal block","blocked nose, stuffy nose, nasal obstruction","","Allergic Rhinitis"],

      // --- HISTORY / COMORBIDITIES -------------------------------------------
      ["HX","General Medicine","Type 2 Diabetes Mellitus","T2DM, diabetes, sugar, DM","HbA1c, FBS, PPBS",""],
      ["HX","General Medicine","Systemic Hypertension","HTN, high BP, blood pressure","RFT, ECG",""],
      ["HX","General Medicine","Bronchial Asthma","asthma, reactive airway","",""],
      ["HX","General Medicine","Hypothyroidism","thyroid, low thyroid","TSH",""],
      ["HX","General Medicine","Coronary Artery Disease","CAD, IHD, heart disease","ECG, Echo",""],
      ["HX","General Medicine","Chronic Kidney Disease","CKD, renal failure, kidney disease","RFT, Urine Routine",""],
      ["HX","General Medicine","Seizure disorder","epilepsy, fits, convulsions","",""],
      ["HX","General Medicine","Pulmonary Tuberculosis","TB, koch's, tuberculosis","Chest X-Ray, Sputum AFB",""],
      ["HX","General Medicine","Dyslipidaemia","high cholesterol, lipid disorder","Lipid Profile",""],
      ["HX","Obstetrics & Gynaecology","Polycystic Ovarian Syndrome","PCOS, PCOD","TSH, USG Pelvis",""],

      // --- ADVICE -------------------------------------------------------------
      ["ADVICE","General Medicine","Take plenty of oral fluids","hydration, drink water, fluids","",""],
      ["ADVICE","General Medicine","Complete the full course of antibiotics","finish antibiotics, full course","",""],
      ["ADVICE","General Medicine","Review immediately if fever persists beyond 3 days","review if fever continues","",""],
      ["ADVICE","General Medicine","Salt restricted diet","low salt, reduce salt","",""],
      ["ADVICE","General Medicine","Diabetic diet, avoid sugar and sweets","diabetic diet, sugar free diet","",""],
      ["ADVICE","General Medicine","Regular blood pressure monitoring at home","BP charting, monitor BP","",""],
      ["ADVICE","General Medicine","Steam inhalation twice daily","steam, inhalation","",""],
      ["ADVICE","General Medicine","Adequate rest, avoid strenuous activity","rest, avoid exertion","",""],
      ["ADVICE","Paediatrics","Continue breastfeeding on demand","breastfeed, feeding advice","",""],
      ["ADVICE","Paediatrics","ORS after each loose stool","ORS, oral rehydration","",""],
      ["ADVICE","Obstetrics & Gynaecology","Iron and calcium supplements daily","iron calcium, supplements","",""],
      ["ADVICE","Obstetrics & Gynaecology","Maintain a menstrual calendar","period diary, menstrual chart","",""]
    ];

    var sh = rx_seedSheet_();
    var existing = {};
    var data = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      existing[dc_upper_(data[i][1]) + "|" + dc_str_(data[i][3]).toLowerCase()] = true;
    }

    var rows = [], order = data.length;
    SEED.forEach(function (s) {
      var key = dc_upper_(s[0]) + "|" + String(s[2]).toLowerCase();
      if (existing[key]) return;
      existing[key] = true;
      rows.push([
        "SEED-" + Utilities.getUuid().substring(0, 8).toUpperCase(),
        String(s[0]), String(s[1]), String(s[2]), String(s[3]),
        String(s[4]), String(s[5]), order++, "ACTIVE"
      ]);
    });

    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
    }
    SpreadsheetApp.flush();
    return "Seeded " + rows.length + " new phrase(s). " +
           (SEED.length - rows.length) + " already present.";
  } catch (e) {
    return "SEED FAILED: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// SECTION B — COMPOSER CONTEXT (one call, not six)
// ============================================================================

/**
 * FRONTEND ENTRY. Everything the consultation form needs to open, in one
 * round trip: three phrase tiers, Rx bundles, exam defaults, slash snippets.
 * @return {{success, phrases:{CC,HX,ADVICE}, bundles, examDefaults, slash}}
 */
function getComposerContext(doctorId, sessionToken) {
  var blank = { CC: [], HX: [], ADVICE: [] };
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) {
      return { success: false, message: scope.message,
               phrases: blank, bundles: [], examDefaults: OPT_FALLBACK_EXAM, slash: [] };
    }

    var me = dc_str_(doctorId) || scope.selfDoctorId;
    var prof = me ? dc_getDoctorById_(me) : null;
    var dept = prof ? dc_str_(prof.specialty) : "";

    // --- tier 3: learned + clinic (from OP_Templates_Engine) ---------------
    var learned = getScopedTemplates(me, sessionToken);

    // --- tier 1: seed library, department-filtered -------------------------
    var seed = { CC: [], HX: [], ADVICE: [] };
    var workup = {};
    var sh = rx_seedSheet_();
    var data = sh.getDataRange().getDisplayValues();

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][8]) === "INACTIVE") continue;
      var cat = dc_upper_(data[i][1]);
      if (!seed[cat]) continue;

      var rowDept = dc_str_(data[i][2]);
      // Show the doctor's own department first, but never hide the rest —
      // a general physician sees everything.
      var isMine = (!dept || !rowDept ||
                    rowDept.toLowerCase() === dept.toLowerCase());

      var text = dc_str_(data[i][3]);
      seed[cat].push({
        text: text,
        aliases: dc_str_(data[i][4]),
        department: rowDept,
        scope: "SEED",
        inDepartment: isMine,
        count: 0
      });

      if (cat === "CC") {
        var labs = dc_str_(data[i][5]);
        var dx   = dc_str_(data[i][6]);
        if (labs || dx) {
          workup[text.toLowerCase()] = {
            labs: labs ? labs.split(",").map(function (x) { return x.trim(); }).filter(Boolean) : [],
            dx: dx
          };
        }
      }
    }

    // Department matches float to the top of the seed tier.
    ["CC", "HX", "ADVICE"].forEach(function (c) {
      seed[c].sort(function (a, b) {
        if (a.inDepartment !== b.inDepartment) return a.inDepartment ? -1 : 1;
        return a.text < b.text ? -1 : 1;
      });
    });

    // --- merge: PERSONAL > CLINIC > SEED ----------------------------------
    var phrases = { CC: [], HX: [], ADVICE: [] };
    ["CC", "HX", "ADVICE"].forEach(function (c) {
      var seen = {};
      var merged = [];
      (learned[c] || []).forEach(function (p) {
        var k = p.text.toLowerCase();
        if (seen[k]) return;
        seen[k] = true;
        merged.push(p);
      });
      seed[c].forEach(function (p) {
        var k = p.text.toLowerCase();
        if (seen[k]) return;      // the doctor's own version wins
        seen[k] = true;
        merged.push(p);
      });
      phrases[c] = merged;
    });

    var bundles = listRxBundles(me, sessionToken);
    var exam = getExamDefaults(me, sessionToken);

    return {
      success: true,
      doctorId: me,
      department: dept,
      phrases: phrases,
      workup: workup,
      bundles: bundles.success ? bundles.bundles : [],
      examDefaults: exam.defaults,
      examIsCustom: exam.isCustom
    };
  } catch (e) {
    return { success: false, message: "Could not load templates: " + e.message,
             phrases: blank, bundles: [], examDefaults: OPT_FALLBACK_EXAM };
  }
}

// ============================================================================
// SECTION C — Rx BUNDLES  (suggested, never auto-inserted)
// ============================================================================

/** FRONTEND ENTRY. Preset bundles + this doctor's own. */
function listRxBundles(doctorId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) return { success: false, message: scope.message, bundles: [] };

    var me = dc_upper_(dc_str_(doctorId) || scope.selfDoctorId);
    var tenant = getTenantId_();
    var sh = rx_bundleSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var personal = [], preset = [];

    for (var i = 1; i < data.length; i++) {
      if (dc_str_(data[i][1]) !== tenant) continue;
      if (dc_upper_(data[i][12]) === "DELETED") continue;

      var sc = dc_upper_(data[i][3]) || "PRESET";
      var owner = dc_upper_(data[i][2]);
      if (sc === "PERSONAL" && owner !== me) continue;

      var items = [];
      try { items = data[i][7] ? JSON.parse(data[i][7]) : []; } catch (e) { continue; }

      var entry = {
        bundleId: dc_str_(data[i][0]),
        scope: sc,
        name: dc_str_(data[i][4]),
        department: dc_str_(data[i][5]),
        indication: dc_str_(data[i][6]),
        items: items,
        itemCount: items.length,
        useCount: dc_int_(data[i][8]),
        editable: (sc === "PERSONAL" && owner === me)
      };
      (sc === "PERSONAL" ? personal : preset).push(entry);
    }

    var byUse = function (a, b) { return b.useCount - a.useCount; };
    personal.sort(byUse);
    preset.sort(byUse);

    return { success: true, bundles: personal.concat(preset) };
  } catch (e) {
    return { success: false, message: "Could not load bundles: " + e.message, bundles: [] };
  }
}

/**
 * FRONTEND ENTRY.
 * payload = { bundleId?, doctorId?, scope, name, department, indication,
 *             items:[{type,drugName,sig,duration,comments,source}] }
 */
function saveRxBundle(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var w = resolveWriteDoctor_(sessionToken, payload && payload.doctorId);
    if (!w.ok) return { success: false, message: w.message };

    var name = dc_str_(payload.name);
    if (!name) return { success: false, message: "Give the bundle a name." };

    var items = (payload.items || []).filter(function (it) {
      return dc_str_(it.drugName);
    });
    if (!items.length) return { success: false, message: "A bundle needs at least one medicine." };

    var requested = dc_upper_(payload.scope) || "PERSONAL";
    var role = dc_str_(w.sess.role).toLowerCase();
    if (requested === "PRESET" && role !== "admin") {
      return { success: false, message: "Only an administrator can save a shared preset." };
    }

    var clean = items.map(function (it) {
      return {
        type: dc_str_(it.type) || "Tab",
        drugName: dc_str_(it.drugName),
        sig: dc_str_(it.sig),
        duration: dc_str_(it.duration),
        comments: dc_str_(it.comments),
        source: dc_upper_(it.source) || "EXTERNAL"
      };
    });

    var sh = rx_bundleSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var now = new Date();
    var existingId = dc_str_(payload.bundleId);

    if (existingId) {
      for (var i = 1; i < data.length; i++) {
        if (dc_upper_(data[i][0]) !== dc_upper_(existingId)) continue;
        var sc = dc_upper_(data[i][3]) || "PRESET";
        if (sc === "PERSONAL" && dc_upper_(data[i][2]) !== dc_upper_(w.doctorId)) {
          return { success: false, message: "You can only edit your own bundles." };
        }
        if (sc === "PRESET" && role !== "admin") {
          return { success: false, message: "Only an administrator can edit a shared preset." };
        }
        var r = i + 1;
        sh.getRange(r, 5).setValue(String(name));
        sh.getRange(r, 6).setValue(String(dc_str_(payload.department)));
        sh.getRange(r, 7).setValue(String(dc_str_(payload.indication)));
        sh.getRange(r, 8).setValue(String(JSON.stringify(clean)));
        sh.getRange(r, 12).setValue(now);
        SpreadsheetApp.flush();
        logAudit_(w.sess, "RX_BUNDLE_UPDATE", "Rx_Bundle", existingId, { name: name });
        return { success: true, bundleId: existingId, message: "Bundle updated." };
      }
      return { success: false, message: "Bundle not found." };
    }

    for (var j = 1; j < data.length; j++) {
      if (dc_upper_(data[j][12]) === "DELETED") continue;
      if (dc_str_(data[j][4]).toLowerCase() !== name.toLowerCase()) continue;
      if (dc_upper_(data[j][3]) === requested &&
          (requested === "PRESET" || dc_upper_(data[j][2]) === dc_upper_(w.doctorId))) {
        return { success: false, message: "You already have a bundle called \"" + name + "\"." };
      }
    }

    var newId = "RXB-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    sh.appendRow([
      String(newId), String(getTenantId_()),
      String(requested === "PRESET" ? "" : w.doctorId),
      String(requested), String(name),
      String(dc_str_(payload.department)), String(dc_str_(payload.indication)),
      String(JSON.stringify(clean)), 0,
      String(w.sess.username), now, now, "ACTIVE"
    ]);

    SpreadsheetApp.flush();
    logAudit_(w.sess, "RX_BUNDLE_CREATE", "Rx_Bundle", newId,
              { name: name, scope: requested, items: clean.length });
    return { success: true, bundleId: newId, message: "Bundle \"" + name + "\" saved." };

  } catch (e) {
    return { success: false, message: "Could not save the bundle: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteRxBundle(bundleId, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired." };

    var sh = rx_bundleSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var role = dc_str_(sess.role).toLowerCase();

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== dc_upper_(bundleId)) continue;
      var sc = dc_upper_(data[i][3]) || "PRESET";
      if (sc === "PERSONAL" && dc_upper_(data[i][2]) !== dc_upper_(sess.doctorId)) {
        return { success: false, message: "You can only delete your own bundles." };
      }
      if (sc === "PRESET" && role !== "admin") {
        return { success: false, message: "Only an administrator can delete a shared preset." };
      }
      sh.getRange(i + 1, 13).setValue("DELETED");
      SpreadsheetApp.flush();
      return { success: true, message: "Bundle removed." };
    }
    return { success: false, message: "Bundle not found." };
  } catch (e) {
    return { success: false, message: "Could not delete: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** Records that a bundle was actually used, for ranking. Fire-and-forget. */
function noteRxBundleUsed(bundleId, sessionToken) {
  try {
    var sh = rx_bundleSheet_();
    var data = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) === dc_upper_(bundleId)) {
        sh.getRange(i + 1, 9).setValue(dc_int_(data[i][8]) + 1);
        return { success: true };
      }
    }
    return { success: false };
  } catch (e) { return { success: false }; }
}

// ============================================================================
// SECTION D — TAPER BUILDER
// ============================================================================

/**
 * FRONTEND ENTRY. Expands a taper spec into printable steps and a total
 * tablet count. Pure computation — safe to call on every keystroke.
 *
 * spec = { drugName, unitStrength, startDose, stepSize, stepDays,
 *          pattern: 'LINEAR'|'HALVING'|'CUSTOM', minDose,
 *          customSteps:[{dose,days}], frequency }
 *
 * Returns steps, a human-readable sig, and computedQty — which pharmacy MUST
 * use, because a taper's quantity cannot be inferred from duration alone.
 */
function buildTaperPlan(spec) {
  try {
    spec = spec || {};
    var unit  = parseFloat(spec.unitStrength) || 0;
    var freq  = dc_int_(spec.frequency) || 1;
    var pattern = dc_upper_(spec.pattern) || "LINEAR";
    var steps = [];

    if (pattern === "CUSTOM") {
      (spec.customSteps || []).forEach(function (s) {
        var d = parseFloat(s.dose), days = dc_int_(s.days);
        if (d > 0 && days > 0) steps.push({ dose: d, days: days });
      });
    } else {
      var dose = parseFloat(spec.startDose) || 0;
      var stepDays = dc_int_(spec.stepDays) || 0;
      var minDose = parseFloat(spec.minDose) || 0;
      var stepSize = parseFloat(spec.stepSize) || 0;

      if (dose <= 0)     return { success: false, message: "Enter a starting dose." };
      if (stepDays <= 0) return { success: false, message: "Enter how many days each step lasts." };
      if (pattern === "LINEAR" && stepSize <= 0) {
        return { success: false, message: "Enter the reduction per step." };
      }

      var guard = 0;
      while (dose > 0 && guard < RX_MAX_TAPER_STEPS) {
        steps.push({ dose: Math.round(dose * 100) / 100, days: stepDays });
        guard++;
        var next = (pattern === "HALVING") ? dose / 2 : dose - stepSize;
        if (next < minDose || next <= 0) break;
        dose = next;
      }
      if (guard >= RX_MAX_TAPER_STEPS) {
        return { success: false, message: "That taper runs past " + RX_MAX_TAPER_STEPS + " steps — check the values." };
      }
    }

    if (!steps.length) return { success: false, message: "No taper steps produced." };

    // --- expand into printable rows + quantity ----------------------------
    var dayCursor = 1, totalDays = 0, totalUnits = 0;
    var rows = steps.map(function (s) {
      var from = dayCursor;
      var to = dayCursor + s.days - 1;
      dayCursor = to + 1;
      totalDays += s.days;

      var unitsPerDose = (unit > 0) ? (s.dose / unit) : 0;
      totalUnits += unitsPerDose * freq * s.days;

      return {
        label: "Day " + from + (s.days > 1 ? "\u2013" + to : ""),
        dose: s.dose,
        days: s.days,
        unitsPerDose: Math.round(unitsPerDose * 100) / 100,
        text: "Day " + from + (s.days > 1 ? "\u2013" + to : "") + ": " +
              s.dose + " mg" + (freq > 1 ? " \u00d7 " + freq + "/day" : " OD")
      };
    });

    var qty = Math.ceil(totalUnits);

    return {
      success: true,
      steps: rows,
      totalDays: totalDays,
      computedQty: qty,
      // Compact form for the sig field; the full table prints separately.
      sig: "TAPER: " + rows.map(function (r) { return r.dose + "mg \u00d7 " + r.days + "d"; }).join(" \u2192 "),
      summary: totalDays + " days, " +
               (unit > 0 ? qty + " unit(s) of " + unit + " mg" : "quantity needs a unit strength"),
      warning: (unit > 0) ? "" :
        "Enter the tablet strength or pharmacy cannot compute the quantity."
    };
  } catch (e) {
    return { success: false, message: "Could not build the taper: " + e.message };
  }
}

// ============================================================================
// SECTION E — SAFETY GUARDS
// ----------------------------------------------------------------------------
// These reduce risk. They do not remove it. Wording is deliberately a prompt
// to check, never an all-clear: a text match cannot know that Augmentin is a
// penicillin. Never render "no allergy detected".
// ============================================================================

/** Brand -> generic map from the pharmacy master, plus coverage stats. */
function rx_genericMap_() {
  var map = {}, total = 0, withGeneric = 0;
  try {
    (fetchOPDrugMaster() || []).forEach(function (d) {
      total++;
      var g = dc_str_(d.generic);
      if (g) withGeneric++;
      map[String(d.brand).toLowerCase()] = {
        generic: g.toLowerCase(), status: d.status
      };
    });
  } catch (e) { /* non-fatal */ }
  return { map: map, total: total, withGeneric: withGeneric };
}

/** Patient allergies as a normalised token list. */
function getPatientAllergies(patientId, sessionToken) {
  try {
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired.", allergies: [] };

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Patients");
    if (!sh) return { success: true, allergies: [], recorded: false };

    var m = dc_headerMap_(sh);
    var col = (m["Allergies"] === undefined) ? -1 : m["Allergies"];
    if (col === -1) return { success: true, allergies: [], recorded: false };

    var data = sh.getDataRange().getDisplayValues();
    var pid = dc_upper_(patientId);
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== pid) continue;
      var raw = dc_str_(data[i][col]);
      var list = raw ? raw.split(/[,;/]/).map(function (x) { return x.trim(); })
                          .filter(Boolean) : [];
      return { success: true, allergies: list, recorded: raw !== "", raw: raw };
    }
    return { success: true, allergies: [], recorded: false };
  } catch (e) {
    return { success: false, message: "Could not read allergies: " + e.message, allergies: [] };
  }
}

/** FRONTEND ENTRY. Saves/updates the allergy field for a patient. */
function savePatientAllergies(patientId, allergiesText, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired." };

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Patients");
    if (!sh) return { success: false, message: "Patients sheet missing." };

    var col = dc_ensureColumn_(sh, "Allergies");
    var data = sh.getDataRange().getDisplayValues();
    var pid = dc_upper_(patientId);

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== pid) continue;
      sh.getRange(i + 1, col + 1).setValue(String(dc_str_(allergiesText)));
      SpreadsheetApp.flush();
      logAudit_(sess, "ALLERGY_UPDATE", "Patient", pid, {});
      return { success: true, message: "Allergies updated." };
    }
    return { success: false, message: "Patient not found." };
  } catch (e) {
    return { success: false, message: "Could not save allergies: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * FRONTEND ENTRY. Checks a prescription list for allergy matches and
 * duplicate therapy. Call before saving, and on adding each drug.
 *
 * @return {{success, alerts:[{level,type,message}], coverage:{...}}}
 *   level: 'DANGER' (allergy) | 'WARN' (duplicate) | 'INFO'
 */
function checkRxSafety(patientId, meds, sessionToken) {
  try {
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired.", alerts: [] };

    meds = meds || [];
    var gm = rx_genericMap_();
    var alerts = [];

    // ---- allergy check ---------------------------------------------------
    var allergyRes = getPatientAllergies(patientId, sessionToken);
    var allergies = (allergyRes.allergies || []).map(function (a) { return a.toLowerCase(); });

    if (!allergyRes.recorded) {
      alerts.push({
        level: "INFO", type: "NO_ALLERGY_RECORD",
        message: "No allergy history on file for this patient. Ask and record it."
      });
    }

    meds.forEach(function (md) {
      var name = dc_str_(md.drugName);
      if (!name) return;
      var lower = name.toLowerCase();
      var entry = gm.map[lower];
      var generic = entry ? entry.generic : "";

      allergies.forEach(function (al) {
        if (!al || al.length < 3) return;
        var hit = (lower.indexOf(al) !== -1) ||
                  (generic && generic.indexOf(al) !== -1) ||
                  (al.indexOf(lower) !== -1);
        if (hit) {
          alerts.push({
            level: "DANGER", type: "ALLERGY",
            message: name + " may match the recorded allergy \u201c" + al + "\u201d. Confirm before prescribing."
          });
        }
      });
    });

    // ---- duplicate therapy ----------------------------------------------
    var byGeneric = {};
    meds.forEach(function (md) {
      var name = dc_str_(md.drugName);
      if (!name) return;
      var entry = gm.map[name.toLowerCase()];
      if (!entry || !entry.generic) return;
      if (!byGeneric[entry.generic]) byGeneric[entry.generic] = [];
      byGeneric[entry.generic].push(name);
    });

    Object.keys(byGeneric).forEach(function (g) {
      if (byGeneric[g].length < 2) return;
      alerts.push({
        level: "WARN", type: "DUPLICATE",
        message: byGeneric[g].join(" and ") + " both contain " + g + " \u2014 duplicate therapy."
      });
    });

    // Honest about how much of the formulary this can actually check.
    var coveragePct = gm.total ? Math.round((gm.withGeneric / gm.total) * 100) : 0;

    return {
      success: true,
      alerts: alerts,
      allergiesRecorded: allergyRes.recorded,
      coverage: {
        drugsInMaster: gm.total,
        withGenericName: gm.withGeneric,
        percent: coveragePct,
        note: coveragePct < 90
          ? "Duplicate-therapy checking only covers drugs with a Generic name in Pharmacy_Inventory (" +
            coveragePct + "% of your formulary)."
          : ""
      }
    };
  } catch (e) {
    return { success: false, message: "Safety check failed: " + e.message, alerts: [] };
  }
}

// ============================================================================
// SECTION F — SAME AS LAST VISIT
// ============================================================================

/**
 * FRONTEND ENTRY. The patient's previous encounter, for one-click recall.
 * Returns data only — the UI renders it as a card the doctor must accept.
 */
function getLastVisitRecall(patientId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, null);
    if (!scope.ok) return { success: false, message: scope.message };

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OP_Encounters");
    if (!sh || sh.getLastRow() < 2) {
      return { success: true, found: false, message: "No previous visit on record." };
    }

    var m = dc_headerMap_(sh);
    var docCol  = (m["Doctor_ID"] === undefined) ? -1 : m["Doctor_ID"];
    var sigCol  = (m["Doctor_Signature_Snapshot"] === undefined) ? -1 : m["Doctor_Signature_Snapshot"];
    var radOCol = (m["Radiology_Orders"] === undefined) ? -1 : m["Radiology_Orders"];
    var data = sh.getDataRange().getDisplayValues();
    var pid = dc_upper_(patientId);

    for (var i = data.length - 1; i >= 1; i--) {
      if (dc_upper_(data[i][1]) !== pid) continue;

      var rowDoc = (docCol === -1) ? DC_DEFAULT_DOCTOR
                                   : (dc_str_(data[i][docCol]) || DC_DEFAULT_DOCTOR);
      if (!dc_inScope_(scope, rowDoc)) continue;   // respects doctor scoping

      var meds = [], labs = [];
      try { meds = data[i][22] ? JSON.parse(data[i][22]) : []; } catch (e) {}
      try { labs = data[i][23] ? JSON.parse(data[i][23]) : []; } catch (e) {}

      var docRec = dc_getDoctorById_(rowDoc);

      return {
        success: true,
        found: true,
        encounterId: dc_str_(data[i][0]),
        date: dc_str_(data[i][2]),
        doctorId: rowDoc,
        doctorName: docRec ? docRec.name : rowDoc,
        complaints: dc_str_(data[i][10]),
        history: dc_str_(data[i][11]),
        exam: {
          cvs: dc_str_(data[i][17]), rs: dc_str_(data[i][18]),
          pa: dc_str_(data[i][19]), cns: dc_str_(data[i][20])
        },
        diagnosis: dc_str_(data[i][21]),
        meds: meds,
        labs: labs.filter(function (l) { return l.type === "Order"; }),
        advice: dc_str_(data[i][24]),
        radiology: (radOCol === -1) ? "" : dc_str_(data[i][radOCol]),
        message: ""
      };
    }
    return { success: true, found: false, message: "No previous visit on record." };
  } catch (e) {
    return { success: false, message: "Could not load the last visit: " + e.message };
  }
}

// ============================================================================
// SECTION G — PAEDIATRIC DOSE SUGGESTION  (suggested, never applied)
// ============================================================================

/**
 * FRONTEND ENTRY. Computes a weight-based dose and returns it as a
 * suggestion. Deliberately does NOT write to the sig field — the doctor
 * inserts it with a click, like every other suggestion in this engine.
 */
function suggestPaediatricDose(drugName, weightKg, ageYears, frequency) {
  try {
    var w = parseFloat(weightKg);
    var age = parseFloat(ageYears);
    var freq = dc_int_(frequency) || 2;

    if (!w || w <= 0) {
      return { success: true, applicable: false, message: "Record the weight to calculate a dose." };
    }
    if (age && age >= 12) {
      return { success: true, applicable: false, message: "" };
    }

    var name = dc_str_(drugName).toLowerCase();
    if (!name) return { success: true, applicable: false, message: "" };

    var ref = null, unit = 0, generic = "";
    (fetchOPDrugMaster() || []).forEach(function (d) {
      if (String(d.brand).toLowerCase() !== name) return;
      var r = String(d.refDose || "").match(/[\d.]+/);
      if (r) ref = parseFloat(r[0]);
      generic = dc_str_(d.generic);
      var u = String(d.unit || "").match(/[\d.]+/);
      if (u) unit = parseFloat(u[0]);
    });

    if (!ref) {
      return {
        success: true, applicable: false,
        message: "No reference dose on file for " + drugName +
                 ". Add mg/kg/day to Pharmacy_Inventory to enable this."
      };
    }

    var perDay = w * ref;
    var perDose = perDay / freq;

    return {
      success: true,
      applicable: true,
      drugName: drugName,
      generic: generic,
      refDose: ref,
      weightKg: w,
      perDay: Math.round(perDay * 10) / 10,
      perDose: Math.round(perDose * 10) / 10,
      frequency: freq,
      unitStrength: unit,
      units: (unit > 0) ? Math.round((perDose / unit) * 100) / 100 : null,
      sig: (unit > 0)
        ? (Math.round((perDose / unit) * 100) / 100) + " unit " + freqWordServer_(freq) +
          " (" + Math.round(perDose * 10) / 10 + " mg)"
        : Math.round(perDose * 10) / 10 + " mg " + freqWordServer_(freq),
      note: "Based on " + ref + " mg/kg/day at " + w + " kg. Verify against the child's condition."
    };
  } catch (e) {
    return { success: false, message: "Dose calculation failed: " + e.message };
  }
}

function freqWordServer_(f) {
  return ({ 1: "OD", 2: "BD", 3: "TDS", 4: "QID" })[f] || (f + "x/day");
}