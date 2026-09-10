// ============================================================================
// Drug_Interactions.gs  —  Crescentia HealthTech
// Drug–drug interaction checking, driven by a sheet the clinic maintains.
// ----------------------------------------------------------------------------
// WHY A SHEET, NOT A HARD-CODED TABLE
// Interaction knowledge changes, and the clinic's formulary is its own. A
// sheet lets the prescribing doctors curate it without a deployment. The
// checker degrades honestly: it reports how many pairs it knows, so nobody
// mistakes an empty table for a clean prescription.
//
// THIS IS A SAFETY NET, NOT A PHARMACOPOEIA. It only knows the pairs entered
// in the sheet. It never blocks a prescription — it surfaces a warning the
// prescriber acknowledges.
// ============================================================================

var DI_SHEET = "Drug_Interactions";
var DI_HEADERS = [
  "Interaction_ID", "Drug_A", "Drug_B", "Severity", "Mechanism",
  "Clinical_Effect", "Management", "Source", "Status"
];

/** Severity ordering, worst first. Anything unrecognised is treated as MODERATE. */
var DI_SEVERITY = ["CONTRAINDICATED", "MAJOR", "MODERATE", "MINOR"];

function di_sheet_() {
  var sh = dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), DI_SHEET, DI_HEADERS);
  sh.setFrozenRows(1);
  return sh;
}

/**
 * ONE-OFF SETUP. Creates the sheet and seeds a starter set of interactions
 * that matter in general medicine. Safe to re-run: existing pairs are kept.
 * Curate and extend this in the sheet — it is meant to be the clinic's own.
 */
function setupDrugInteractions() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sh = di_sheet_();
    var data = dc_sheetValues_(sh);

    var have = {};
    for (var i = 1; i < data.length; i++) {
      have[di_pairKey_(data[i][1], data[i][2])] = true;
    }

    var seed = [
      ["Warfarin", "Aspirin", "MAJOR", "Additive antiplatelet + anticoagulant",
       "Markedly increased bleeding risk", "Avoid together; if unavoidable, monitor INR and watch for bleeding"],
      ["Warfarin", "Ciprofloxacin", "MAJOR", "CYP inhibition raises warfarin levels",
       "INR rises, bleeding risk", "Recheck INR within 3-5 days of starting"],
      ["Warfarin", "Fluconazole", "MAJOR", "CYP2C9 inhibition",
       "INR rises sharply", "Avoid, or reduce warfarin dose and monitor INR"],
      ["Clopidogrel", "Omeprazole", "MODERATE", "CYP2C19 inhibition reduces clopidogrel activation",
       "Reduced antiplatelet effect", "Prefer pantoprazole"],
      ["Metformin", "Contrast Media", "MAJOR", "Risk of lactic acidosis with renal impairment",
       "Lactic acidosis", "Withhold metformin around contrast imaging; recheck renal function"],
      ["ACE Inhibitor", "Spironolactone", "MAJOR", "Additive potassium retention",
       "Hyperkalaemia", "Monitor serum potassium and renal function"],
      ["Enalapril", "Spironolactone", "MAJOR", "Additive potassium retention",
       "Hyperkalaemia", "Monitor serum potassium"],
      ["Ramipril", "Spironolactone", "MAJOR", "Additive potassium retention",
       "Hyperkalaemia", "Monitor serum potassium"],
      ["NSAID", "ACE Inhibitor", "MODERATE", "Reduced renal perfusion",
       "Acute kidney injury, blunted BP control", "Avoid prolonged use; monitor renal function"],
      ["Diclofenac", "Enalapril", "MODERATE", "Reduced renal perfusion",
       "Acute kidney injury", "Avoid prolonged combination; monitor creatinine"],
      ["Ibuprofen", "Aspirin", "MODERATE", "Competes at COX-1",
       "Reduced cardioprotective effect of aspirin", "Give aspirin 2 hours before ibuprofen"],
      ["Statin", "Clarithromycin", "MAJOR", "CYP3A4 inhibition",
       "Myopathy, rhabdomyolysis", "Suspend the statin during the course"],
      ["Atorvastatin", "Clarithromycin", "MAJOR", "CYP3A4 inhibition",
       "Myopathy, rhabdomyolysis", "Suspend atorvastatin during the course"],
      ["Simvastatin", "Amlodipine", "MODERATE", "CYP3A4 interaction",
       "Increased statin exposure, myopathy", "Cap simvastatin at 20 mg daily"],
      ["Digoxin", "Furosemide", "MODERATE", "Hypokalaemia potentiates digoxin",
       "Digoxin toxicity", "Monitor potassium and digoxin level"],
      ["Digoxin", "Amiodarone", "MAJOR", "Reduced digoxin clearance",
       "Digoxin toxicity", "Halve the digoxin dose and monitor levels"],
      ["Methotrexate", "Trimethoprim", "CONTRAINDICATED", "Additive antifolate effect",
       "Severe myelosuppression", "Do not combine"],
      ["Methotrexate", "NSAID", "MAJOR", "Reduced methotrexate clearance",
       "Methotrexate toxicity", "Avoid; monitor counts if unavoidable"],
      ["Tramadol", "SSRI", "MAJOR", "Additive serotonergic effect",
       "Serotonin syndrome, seizures", "Avoid; use an alternative analgesic"],
      ["Sertraline", "Tramadol", "MAJOR", "Additive serotonergic effect",
       "Serotonin syndrome", "Avoid combination"],
      ["Amiodarone", "Levofloxacin", "MAJOR", "Additive QT prolongation",
       "Torsades de pointes", "Avoid; ECG if unavoidable"],
      ["Ondansetron", "Amiodarone", "MODERATE", "Additive QT prolongation",
       "QT prolongation", "Use the lowest dose; consider ECG"],
      ["Azithromycin", "Amiodarone", "MAJOR", "Additive QT prolongation",
       "Torsades de pointes", "Avoid; choose another antibiotic"],
      ["Phenytoin", "Fluconazole", "MAJOR", "CYP2C9 inhibition",
       "Phenytoin toxicity", "Monitor phenytoin level"],
      ["Levothyroxine", "Calcium Carbonate", "MODERATE", "Reduced absorption",
       "Under-treated hypothyroidism", "Separate doses by at least 4 hours"],
      ["Levothyroxine", "Iron", "MODERATE", "Reduced absorption",
       "Under-treated hypothyroidism", "Separate doses by at least 4 hours"],
      ["Ciprofloxacin", "Calcium Carbonate", "MODERATE", "Chelation reduces absorption",
       "Treatment failure", "Separate doses by at least 2 hours"],
      ["Spironolactone", "Potassium Chloride", "MAJOR", "Additive potassium load",
       "Hyperkalaemia", "Avoid unless potassium is being monitored"],
      ["Insulin", "Propranolol", "MODERATE", "Beta blockade masks hypoglycaemia",
       "Unrecognised hypoglycaemia", "Warn the patient; prefer a cardioselective agent"],
      ["Prednisolone", "NSAID", "MODERATE", "Additive GI mucosal injury",
       "GI bleeding, ulceration", "Add gastroprotection"]
    ];

    var added = [];
    seed.forEach(function (r) {
      var k = di_pairKey_(r[0], r[1]);
      if (have[k]) return;
      have[k] = true;
      added.push([
        "DI-" + Utilities.getUuid().substring(0, 8).toUpperCase(),
        r[0], r[1], r[2], r[3], r[4], r[5], "Clinic starter set", "ACTIVE"
      ]);
    });

    if (added.length) {
      sh.getRange(sh.getLastRow() + 1, 1, added.length, DI_HEADERS.length).setValues(added);
    }
    dc_invalidate_(DI_SHEET);
    SpreadsheetApp.flush();

    return "Drug_Interactions ready. Added " + added.length + " pair(s); " +
           (Object.keys(have).length) + " total.\n" +
           "Curate this sheet directly — Drug_A and Drug_B match against both the " +
           "brand typed and its Generic name in Pharmacy_Inventory, so an entry " +
           "like 'NSAID' will not match unless a drug is literally named that. " +
           "Prefer generic names.";
  } catch (e) {
    return "Setup failed: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

/** Order-independent key so A+B and B+A are one pair. */
function di_pairKey_(a, b) {
  var x = dc_upper_(a), y = dc_upper_(b);
  return (x < y) ? (x + "||" + y) : (y + "||" + x);
}

/** Loaded once per execution: [{a, b, severity, mechanism, effect, management}] */
function di_rules_() {
  var out = [];
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DI_SHEET);
    if (!sh) return out;
    var data = dc_sheetValues_(sh);
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][8]) === "INACTIVE") continue;
      var a = dc_str_(data[i][1]), b = dc_str_(data[i][2]);
      if (!a || !b) continue;
      out.push({
        a: a.toLowerCase(), b: b.toLowerCase(),
        severity: dc_upper_(data[i][3]) || "MODERATE",
        mechanism: dc_str_(data[i][4]),
        effect: dc_str_(data[i][5]),
        management: dc_str_(data[i][6])
      });
    }
  } catch (e) { /* no sheet yet — the checker reports zero coverage */ }
  return out;
}

/**
 * Every name a prescribed drug could be known by: what was typed, and its
 * Generic from Pharmacy_Inventory. "Tab Azithral 500" has to match a rule
 * written against "Azithromycin".
 */
function di_aliases_(drugName, genericMap) {
  var typed = dc_str_(drugName).toLowerCase();
  var out = [typed];
  var entry = genericMap && genericMap.map ? genericMap.map[typed] : null;
  if (entry && entry.generic) out.push(String(entry.generic).toLowerCase());
  return out;
}

/** True if any alias contains the rule term, or vice versa. */
function di_matches_(aliases, term) {
  if (!term || term.length < 3) return false;
  for (var i = 0; i < aliases.length; i++) {
    var a = aliases[i];
    if (!a) continue;
    if (a.indexOf(term) !== -1 || term.indexOf(a) !== -1) return true;
  }
  return false;
}

/**
 * Interaction alerts for a medication list.
 * @param {Array} meds  [{ drugName }]
 * @return {{alerts:[], rulesLoaded:number}}
 */
function checkDrugInteractions(meds, genericMap) {
  var result = { alerts: [], rulesLoaded: 0 };
  try {
    meds = (meds || []).filter(function (m) { return dc_str_(m.drugName); });
    if (meds.length < 2) return result;

    var rules = di_rules_();
    result.rulesLoaded = rules.length;
    if (!rules.length) return result;

    var gm = genericMap || rx_genericMap_();
    var aliases = meds.map(function (m) { return di_aliases_(m.drugName, gm); });
    var seen = {};

    for (var i = 0; i < meds.length; i++) {
      for (var j = i + 1; j < meds.length; j++) {
        for (var r = 0; r < rules.length; r++) {
          var rule = rules[r];
          var hit = (di_matches_(aliases[i], rule.a) && di_matches_(aliases[j], rule.b)) ||
                    (di_matches_(aliases[i], rule.b) && di_matches_(aliases[j], rule.a));
          if (!hit) continue;

          var key = di_pairKey_(meds[i].drugName, meds[j].drugName) + "||" + rule.severity;
          if (seen[key]) continue;
          seen[key] = true;

          result.alerts.push({
            level: (rule.severity === "CONTRAINDICATED" || rule.severity === "MAJOR")
                     ? "DANGER" : "WARN",
            type: "INTERACTION",
            severity: rule.severity,
            drugs: [dc_str_(meds[i].drugName), dc_str_(meds[j].drugName)],
            message: dc_str_(meds[i].drugName) + " + " + dc_str_(meds[j].drugName) +
                     " — " + rule.severity.toLowerCase() + " interaction" +
                     (rule.effect ? ": " + rule.effect : "") +
                     (rule.management ? ". " + rule.management + "." : ".")
          });
        }
      }
    }

    // Worst first, so the prescriber reads the dangerous pair before the minor one.
    result.alerts.sort(function (x, y) {
      return DI_SEVERITY.indexOf(x.severity) - DI_SEVERITY.indexOf(y.severity);
    });
  } catch (e) {
    Logger.log("Interaction check failed: " + e.message);
  }
  return result;
}
