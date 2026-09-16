// ============================================================================
// Drug_Interaction.gs  —  Crescentia HealthTech
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
       "GI bleeding, ulceration", "Add gastroprotection"],

      // --- pairs that turn up in Indian general practice and were missing ---
      // Written in GENERIC terms, because that is what di_aliases_() resolves
      // a typed brand to. A pair named after a class ("NSAID") only fires on
      // a drug literally called that, so the common members are listed too.
      ["Metformin", "Alcohol", "MODERATE", "Additive lactate production",
       "Lactic acidosis", "Advise abstinence during treatment"],
      ["Glimepiride", "Fluconazole", "MODERATE", "CYP2C9 inhibition raises sulfonylurea levels",
       "Hypoglycaemia", "Warn the patient; monitor sugars during the course"],
      ["Glimepiride", "Ciprofloxacin", "MODERATE", "Altered glucose handling",
       "Hypo- or hyperglycaemia", "Monitor blood sugar through the course"],
      ["Amlodipine", "Clarithromycin", "MAJOR", "CYP3A4 inhibition",
       "Profound hypotension, acute kidney injury", "Suspend amlodipine during the course"],
      ["Atenolol", "Verapamil", "MAJOR", "Additive AV nodal blockade",
       "Bradycardia, heart block", "Avoid; use a dihydropyridine instead"],
      ["Diltiazem", "Atenolol", "MAJOR", "Additive AV nodal blockade",
       "Bradycardia, heart block", "Avoid the combination"],
      ["Salbutamol", "Propranolol", "MAJOR", "Non-selective beta blockade opposes the bronchodilator",
       "Bronchospasm; the reliever stops working", "Avoid propranolol in asthma"],
      ["Theophylline", "Ciprofloxacin", "MAJOR", "CYP1A2 inhibition",
       "Theophylline toxicity, seizures", "Avoid; choose another antibiotic"],
      ["Carbamazepine", "Clarithromycin", "MAJOR", "CYP3A4 inhibition",
       "Carbamazepine toxicity", "Avoid; use azithromycin instead"],
      ["Warfarin", "Metronidazole", "MAJOR", "CYP2C9 inhibition",
       "INR rises sharply, bleeding", "Recheck INR within 3 days"],
      ["Warfarin", "Diclofenac", "MAJOR", "Antiplatelet effect plus GI injury",
       "Serious GI bleeding", "Avoid; use paracetamol for analgesia"],
      ["Aspirin", "Enalapril", "MODERATE", "Reduced prostaglandin-mediated vasodilation",
       "Blunted BP control, renal impairment", "Acceptable at cardioprotective doses; monitor"],
      ["Clopidogrel", "Pantoprazole", "MINOR", "Weaker CYP2C19 inhibition than omeprazole",
       "Small reduction in antiplatelet effect", "Preferred PPI if one is needed"],
      ["Furosemide", "Gentamicin", "MAJOR", "Additive ototoxicity and nephrotoxicity",
       "Hearing loss, acute kidney injury", "Avoid; monitor levels and creatinine if unavoidable"],
      ["Lithium", "Enalapril", "MAJOR", "Reduced lithium clearance",
       "Lithium toxicity", "Monitor lithium levels closely"],
      ["Lithium", "Diclofenac", "MAJOR", "Reduced lithium clearance",
       "Lithium toxicity", "Avoid NSAIDs; use paracetamol"],
      ["Allopurinol", "Azathioprine", "CONTRAINDICATED", "Blocked azathioprine metabolism",
       "Profound myelosuppression", "Do not combine"],
      ["Rifampicin", "Oral Contraceptive", "MAJOR", "Enzyme induction",
       "Contraceptive failure", "Advise an additional barrier method"],
      ["Rifampicin", "Warfarin", "MAJOR", "Enzyme induction",
       "Loss of anticoagulation", "Expect a large dose increase; monitor INR"],
      ["Domperidone", "Fluconazole", "MAJOR", "CYP3A4 inhibition plus additive QT effect",
       "QT prolongation, arrhythmia", "Avoid the combination"],
      ["Ondansetron", "Domperidone", "MODERATE", "Additive QT prolongation",
       "QT prolongation", "Use one antiemetic, not two"],
      ["Tramadol", "Amitriptyline", "MAJOR", "Additive serotonergic effect, lowered seizure threshold",
       "Serotonin syndrome, seizures", "Avoid; choose a different analgesic"],
      ["Amitriptyline", "Fluoxetine", "MAJOR", "CYP2D6 inhibition raises tricyclic levels",
       "Tricyclic toxicity, arrhythmia", "Avoid; reduce the tricyclic dose substantially if unavoidable"],
      ["Atorvastatin", "Fenofibrate", "MODERATE", "Additive muscle toxicity",
       "Myopathy, rhabdomyolysis", "Monitor CK; warn about muscle pain"],
      ["Pantoprazole", "Iron", "MODERATE", "Raised gastric pH reduces iron absorption",
       "Treatment failure in anaemia", "Give iron with vitamin C, separated from the PPI"],
      ["Calcium Carbonate", "Doxycycline", "MODERATE", "Chelation reduces absorption",
       "Treatment failure", "Separate doses by at least 2 hours"],
      ["Metoclopramide", "Haloperidol", "MAJOR", "Additive dopamine blockade",
       "Extrapyramidal reactions", "Avoid the combination"]
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
 * THERAPEUTIC CLASSES, SO A RULE CAN BE WRITTEN ONCE.
 *
 * A clinic curating the interaction sheet writes what the reference books
 * write: "NSAID + ACE Inhibitor", "Statin + Clarithromycin", "Tramadol +
 * SSRI". Those rules could never fire, because matching was done against the
 * drug's own names only — "aceclofenac" is not the word "NSAID", so a rule
 * naming the class matched nothing, and the pair that matters was missed
 * silently. Thirty curated rules were loaded; a third of them were
 * unreachable.
 *
 * Each class here is therefore added to a drug's alias list, so a rule may
 * name either the class or the molecule and both resolve. It also gives
 * ds_identify_() a class for the duplicate-therapy check when
 * Drug_Dose_Reference has no row for the drug — which, on a new deployment,
 * is every drug.
 *
 * Members are matched whole-word against the resolved generic, never against
 * the brand: "Telma" must not become an ACE inhibitor because a brand name
 * happens to contain letters.
 */
var DI_CLASSES = [
  { klass: 'nsaid',
    members: /^(ibuprofen|diclofenac|aceclofenac|naproxen|indomethacin|ketorolac|mefenamic acid|piroxicam|etoricoxib|nimesulide|flurbiprofen|ketoprofen)$/ },
  { klass: 'ace inhibitor',
    members: /^(enalapril|ramipril|lisinopril|perindopril|captopril|benazepril|fosinopril|quinapril|trandolapril|imidapril)$/ },
  { klass: 'arb',
    members: /^(telmisartan|losartan|valsartan|olmesartan|irbesartan|candesartan|azilsartan)$/ },
  { klass: 'statin',
    members: /^(atorvastatin|rosuvastatin|simvastatin|pravastatin|lovastatin|fluvastatin|pitavastatin)$/ },
  { klass: 'ssri',
    members: /^(fluoxetine|sertraline|paroxetine|citalopram|escitalopram|fluvoxamine)$/ },
  { klass: 'snri',
    members: /^(venlafaxine|desvenlafaxine|duloxetine|milnacipran)$/ },
  { klass: 'proton pump inhibitor',
    members: /^(omeprazole|esomeprazole|pantoprazole|pantaprazole|rabeprazole|lansoprazole|dexlansoprazole|dexrabeprazole)$/ },
  { klass: 'h2 receptor blocker',
    members: /^(ranitidine|famotidine|cimetidine|nizatidine|roxatidine)$/ },
  { klass: 'benzodiazepine',
    members: /^(diazepam|lorazepam|alprazolam|clonazepam|midazolam|nitrazepam|chlordiazepoxide|etizolam)$/ },
  { klass: 'opioid',
    members: /^(morphine|tramadol|fentanyl|codeine|oxycodone|buprenorphine|pethidine|tapentadol)$/ },
  { klass: 'antihistamine',
    members: /^(cetirizine|levocetirizine|loratadine|fexofenadine|chlorpheniramine|hydroxyzine|diphenhydramine|promethazine|bilastine)$/ },
  { klass: 'macrolide',
    members: /^(azithromycin|erythromycin|clarithromycin|roxithromycin)$/ },
  { klass: 'fluoroquinolone',
    members: /^(ciprofloxacin|levofloxacin|ofloxacin|moxifloxacin|norfloxacin)$/ },
  { klass: 'sulfonylurea',
    members: /^(glimepiride|glipizide|gliclazide|glibenclamide|glyburide)$/ },
  { klass: 'anticoagulant',
    members: /^(warfarin|acenocoumarol|dabigatran|rivaroxaban|apixaban|edoxaban|heparin|enoxaparin)$/ },
  { klass: 'antiplatelet',
    members: /^(aspirin|clopidogrel|ticagrelor|prasugrel|dipyridamole|cilostazol)$/ },
  { klass: 'oral corticosteroid',
    members: /^(prednisolone|prednisone|methylprednisolone|dexamethasone|deflazacort|hydrocortisone|betamethasone)$/ },
  { klass: 'thiazide diuretic',
    members: /^(hydrochlorothiazide|chlorthalidone|indapamide|metolazone)$/ },
  { klass: 'loop diuretic',
    members: /^(furosemide|torsemide|bumetanide)$/ },
  { klass: 'potassium sparing diuretic',
    members: /^(spironolactone|eplerenone|amiloride|triamterene)$/ },
  { klass: 'beta blocker',
    members: /^(propranolol|metoprolol|atenolol|bisoprolol|carvedilol|nebivolol|labetalol|sotalol)$/ },
  { klass: 'calcium channel blocker',
    members: /^(amlodipine|nifedipine|felodipine|diltiazem|verapamil|cilnidipine)$/ },
  { klass: 'analgesic antipyretic',
    members: /^(paracetamol|para|acetaminophen)$/ }
];

/**
 * The classes a set of resolved generic names belongs to.
 * @param {Array<string>} generics  lower-cased, already split out of a
 *                                  combination product
 * @return {Array<string>}
 */
function di_classesFor_(generics) {
  var out = [];
  (generics || []).forEach(function (g) {
    var t = String(g || '').toLowerCase().trim();
    if (!t) return;
    DI_CLASSES.forEach(function (c) {
      if (c.members.test(t) && out.indexOf(c.klass) === -1) out.push(c.klass);
    });
  });
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

  // Strip the dose form and strength a prescription carries but a rule does
  // not: "Tab Azithral 500" has to match a rule written as "Azithromycin",
  // and its bare-name form is what the generic map is keyed on.
  var bare = typed
    .replace(/\b(tab|tabs|tablet|cap|caps|capsule|syp|syrup|inj|injection|susp|suspension|oint|ointment|cream|gel|drop|drops|soln|solution|sachet|powder|spray|lotion|patch|supp)\b/g, " ")
    .replace(/\b\d+(\.\d+)?\s*(mg|mcg|g|gm|ml|iu|units?|%)?\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (bare && bare !== typed) out.push(bare);

  var entry = (genericMap && genericMap.map)
    ? (genericMap.map[typed] || genericMap.map[bare])
    : null;
  if (entry && entry.generic) {
    var g = String(entry.generic).toLowerCase();
    out.push(g);
    // A combination product lists several generics; each can interact.
    g.split(/[+,/]/).forEach(function (part) {
      var t = part.trim();
      if (t.length >= 4 && out.indexOf(t) === -1) out.push(t);
    });
  }
  // The therapeutic classes this drug belongs to, so a rule written as
  // "NSAID" or "Statin" reaches it. Derived from the RESOLVED generics only,
  // never from the typed brand.
  var generics = out.slice(1);
  di_classesFor_(generics).forEach(function (k) {
    if (out.indexOf(k) === -1) out.push(k);
    // "proton pump inhibitor" is often written "PPI", and "ACE Inhibitor"
    // without the space. Add the short form so either spelling matches.
    var short = { 'proton pump inhibitor': 'ppi',
                  'h2 receptor blocker': 'h2 blocker',
                  'oral corticosteroid': 'corticosteroid',
                  'calcium channel blocker': 'ccb',
                  'analgesic antipyretic': 'paracetamol' }[k];
    if (short && out.indexOf(short) === -1) out.push(short);
  });

  return out.filter(function (x, i) { return x && out.indexOf(x) === i; });
}

/**
 * True if a rule term matches one of a drug's names.
 *
 * Word-boundary anchored, NOT plain substring. "Iron" inside "Ironsucrose" is
 * a real match; "iron" inside "Environ" is not, and a bare indexOf would
 * accept both. A false interaction warning that fires on every prescription
 * gets ignored, which is worse than no warning at all.
 */
function di_matches_(aliases, term) {
  if (!term || term.length < 3) return false;
  var t = String(term).toLowerCase().trim();
  var re;
  try {
    re = new RegExp("(^|[^a-z])" + t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "([^a-z]|$)", "i");
  } catch (e) { return false; }

  for (var i = 0; i < aliases.length; i++) {
    var a = aliases[i];
    if (!a) continue;
    if (a === t) return true;
    if (re.test(a)) return true;
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

// ---------------------------------------------------------------------------
// DIAGNOSTICS — answer "is my interaction data actually being used?"
// ---------------------------------------------------------------------------

/**
 * Run from the editor. Takes two drug names as they would be TYPED on a
 * prescription and explains, step by step, whether a rule fires and why.
 *
 *   testDrugInteraction("Tab Azithral 500", "Tab Amiodarone 100")
 *
 * Use it after curating the sheet: it shows the names each drug resolves to,
 * so a rule written in generic terms can be checked against the brand the
 * clinic actually stocks.
 */
function testDrugInteraction(drugA, drugB) {
  var out = [];
  var gm = rx_genericMap_();
  var rules = di_rules_();

  out.push("Rules loaded from " + DI_SHEET + ": " + rules.length);
  if (!rules.length) {
    out.push("");
    out.push("NOTHING IS BEING CHECKED. Either the sheet is empty, every row is");
    out.push("marked INACTIVE, or Drug_A / Drug_B is blank on every row.");
    Logger.log(out.join("\n"));
    return out.join("\n");
  }

  var aliasA = di_aliases_(drugA, gm);
  var aliasB = di_aliases_(drugB, gm);
  out.push("");
  out.push("\"" + drugA + "\" resolves to: " + aliasA.join(" | "));
  out.push("\"" + drugB + "\" resolves to: " + aliasB.join(" | "));

  if (aliasA.length < 2) {
    out.push("  NOTE: no generic found for A. Add its Generic name in the");
    out.push("        Pharmacy_Inventory sheet, or write the rule against the brand.");
  }
  if (aliasB.length < 2) {
    out.push("  NOTE: no generic found for B. Same fix.");
  }

  var res = checkDrugInteractions([{ drugName: drugA }, { drugName: drugB }], gm);
  out.push("");
  if (res.alerts.length) {
    out.push("MATCHED " + res.alerts.length + " rule(s):");
    res.alerts.forEach(function (a) { out.push("  [" + a.severity + "] " + a.message); });
  } else {
    out.push("NO RULE MATCHED.");
    out.push("");
    out.push("Rules mentioning either drug, for comparison:");
    var near = 0;
    rules.forEach(function (r) {
      if (di_matches_(aliasA, r.a) || di_matches_(aliasA, r.b) ||
          di_matches_(aliasB, r.a) || di_matches_(aliasB, r.b)) {
        out.push("  " + r.a + "  +  " + r.b + "   (" + r.severity + ")");
        near++;
      }
    });
    if (!near) {
      out.push("  none — no rule names either drug under any of the forms above.");
      out.push("");
      out.push("Write Drug_A and Drug_B using the names shown in the 'resolves to'");
      out.push("lines. Generic names work only when the brand carries that Generic");
      out.push("in Pharmacy_Inventory.");
    }
  }

  Logger.log(out.join("\n"));
  return out.join("\n");
}

/**
 * Audits the whole interaction sheet against the live formulary and reports
 * which rules can never fire, because neither side matches any stocked drug
 * or its generic. A rule that cannot fire is worse than a missing one: it
 * looks like coverage.
 */
function auditDrugInteractionCoverage() {
  var gm = rx_genericMap_();
  var rules = di_rules_();
  if (!rules.length) return "No interaction rules loaded. Run setupDrugInteractions() first.";

  // Every name the formulary can present, brand and generic.
  var stocked = [];
  try {
    var inv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pharmacy_Inventory");
    var data = dc_sheetValues_(inv);
    for (var i = 1; i < data.length; i++) {
      var brand = dc_str_(data[i][1]);
      if (brand) stocked.push(di_aliases_(brand, gm));
    }
  } catch (e) { return "Could not read Pharmacy_Inventory: " + e.message; }

  var termHits = function (term) {
    for (var i = 0; i < stocked.length; i++) {
      if (di_matches_(stocked[i], term)) return true;
    }
    return false;
  };

  var dead = [], live = 0;
  rules.forEach(function (r) {
    var a = termHits(r.a), b = termHits(r.b);
    if (a && b) { live++; return; }
    dead.push("  " + r.a + " + " + r.b +
              "   (" + (!a ? "'" + r.a + "' not stocked" : "") +
              (!a && !b ? "; " : "") +
              (!b ? "'" + r.b + "' not stocked" : "") + ")");
  });

  var out = [
    "Interaction rules: " + rules.length,
    "Can fire against the current formulary: " + live,
    "Cannot fire: " + dead.length
  ];
  if (dead.length) {
    out.push("");
    out.push("These rules will never trigger — neither the brand nor its Generic");
    out.push("in Pharmacy_Inventory matches the term written in the rule:");
    dead.forEach(function (d) { out.push(d); });
    out.push("");
    out.push("Either stock the drug, or rewrite the rule using a name the");
    out.push("formulary actually carries. Check a specific pair with");
    out.push("testDrugInteraction(\"brand A\", \"brand B\").");
  }
  var report = out.join("\n");
  Logger.log(report);
  return report;
}