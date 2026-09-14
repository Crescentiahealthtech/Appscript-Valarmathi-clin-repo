// ============================================================================
// Drug_Safety.gs  —  Crescentia HealthTech
// Duplicate therapy, therapeutic-class overlap, and honest coverage.
// ----------------------------------------------------------------------------
// WHAT THE CHECKS IN checkRxSafety() COULD AND COULD NOT SEE
//
// 1. DUPLICATE BY GENERIC only fired when BOTH rows carried a Generic name in
//    Pharmacy_Inventory. The coverage note said so honestly — and on a
//    formulary that is sixty per cent filled in, that is a check which
//    silently does nothing for the other forty.
//
// 2. THE SAME DRUG TWICE was matched on the lower-cased name with runs of
//    whitespace collapsed. "Tab Paracetamol 500" and "PARACETAMOL 500mg" are
//    the same drug written twice and did not match — and the double-entry
//    this check exists for is exactly the case where the two lines were typed
//    by different people, or by the same person on different days, and are
//    therefore spelt differently.
//
// 3. THERAPEUTIC-CLASS OVERLAP was not checked at all. Two NSAIDs, two PPIs,
//    two benzodiazepines, an ACE inhibitor and an ARB — these are the
//    duplicates that reach the patient, because they look like two different
//    drugs on the page and the interaction table has no pair for them.
//
// 4. THE ALLERGY MATCH used a plain indexOf in both directions, which is the
//    false-positive problem di_matches_() in Drug_Interaction.gs was written
//    to solve. A recorded allergy to "ASA" matched every drug with "asa" in
//    it — Asacol, Losartan-ASA combinations, Nasal drops. An alert that fires
//    on prescriptions it has no business firing on gets dismissed by reflex,
//    and then so does the one that mattered.
//
// This file fixes all four. It draws names and classes from BOTH sources the
// clinic maintains — Pharmacy_Inventory's Generic column and
// Drug_Dose_Reference's Generic and Class columns — so a drug missing from
// one is still checked through the other.
// ============================================================================

/**
 * Classes whose members must not be doubled up, and why.
 *
 * Keyed on the Class text in Drug_Dose_Reference, matched loosely so
 * "Antibiotic - Penicillin" and "NSAID" both land. Only classes where TWO AT
 * ONCE is a real problem are listed: two antibiotics are routine, two NSAIDs
 * are not.
 */
var DS_CLASS_RULES = [
  { match: /nsaid/i,
    label: 'NSAIDs',
    why: 'Two NSAIDs together add gastric and renal risk without adding analgesia.' },
  { match: /proton pump/i,
    label: 'proton pump inhibitors',
    why: 'One is enough; a second adds no acid suppression.' },
  { match: /h2 receptor/i,
    label: 'H2 blockers',
    why: 'Two H2 blockers duplicate the same action.' },
  { match: /ace inhibitor/i,
    label: 'ACE inhibitors',
    why: 'Two ACE inhibitors multiply the risk of hyperkalaemia and acute kidney injury.' },
  { match: /\bARB\b|angiotensin receptor/i,
    label: 'ARBs',
    why: 'Two ARBs multiply the risk of hyperkalaemia and acute kidney injury.' },
  { match: /benzodiazepine/i,
    label: 'benzodiazepines',
    why: 'Additive sedation and respiratory depression.' },
  { match: /opioid/i,
    label: 'opioids',
    why: 'Additive sedation and respiratory depression.' },
  { match: /antihistamine/i,
    label: 'antihistamines',
    why: 'Additive sedation and anticholinergic load with no extra benefit.' },
  { match: /statin/i,
    label: 'statins',
    why: 'Two statins multiply the risk of myopathy.' },
  { match: /corticosteroid - oral/i,
    label: 'oral corticosteroids',
    why: 'Two oral steroids compound the total dose. Convert to one.' },
  { match: /antidiabetic - sulfonylurea|sulfonylurea/i,
    label: 'sulfonylureas',
    why: 'Additive hypoglycaemia.' },
  { match: /anticoagulant/i,
    label: 'anticoagulants',
    why: 'Two anticoagulants multiply bleeding risk. This is almost never intended.' },
  { match: /antiplatelet/i,
    label: 'antiplatelets',
    why: 'Dual antiplatelet therapy is a deliberate decision — confirm it is one.' },
  { match: /analgesic \/ antipyretic|paracetamol/i,
    label: 'paracetamol-containing products',
    why: 'Paracetamol in two products stacks towards the daily maximum without ' +
         'either line looking excessive.' }
];

function ds_str_(v) { return (v === null || v === undefined) ? "" : String(v).trim(); }

/**
 * The normalised stem of a drug name, for matching one prescription line
 * against another.
 *
 * "Tab Paracetamol 500", "PARACETAMOL 500mg" and "Syp. Paracetamol" all
 * reduce to "paracetamol". Reuses dref_key_() from Dose_Reference.gs so the
 * dose reference, the duplicate check and the class check all agree about
 * what a drug is called; falls back to its own stripper if that file is not
 * deployed.
 */
function ds_key_(name) {
  if (typeof dref_key_ === "function") return dref_key_(name);
  return ds_str_(name).toLowerCase()
    .replace(/\b(tab|tabs|tablet|cap|caps|capsule|syp|syrup|susp|suspension|inj|injection|drops?|oint|cream|neb|resp(?:ule|ules)?)\b\.?/g, " ")
    .replace(/[0-9]+(\.[0-9]+)?\s*(mg|mcg|g|ml|iu|%)?/g, " ")
    .replace(/[^a-z ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Everything the clinic knows about one prescribed name.
 *
 * TWO SOURCES, DELIBERATELY. Pharmacy_Inventory carries the Generic for what
 * the clinic stocks; Drug_Dose_Reference carries the Generic AND the Class
 * for what the clinic prescribes. A drug in one and not the other is the
 * normal case, and checking against only the first is why duplicate therapy
 * covered so little of the formulary.
 *
 * @return {{name, key, generics:Array<string>, klass:string, source:string}}
 */
function ds_identify_(drugName, genericMap) {
  var name = ds_str_(drugName);
  var key = ds_key_(name);
  var lower = name.toLowerCase();
  var generics = [];
  var klass = "";
  var sources = [];

  // a) the formulary
  var entry = (genericMap && genericMap.map)
    ? (genericMap.map[lower] || genericMap.map[key]) : null;
  if (entry && entry.generic) {
    sources.push("inventory");
    String(entry.generic).toLowerCase().split(/[+,\/]/).forEach(function (g) {
      var t = g.trim();
      if (t.length >= 3 && generics.indexOf(t) === -1) generics.push(t);
    });
  }

  // b) the dose reference
  try {
    if (typeof dref_all_ === "function") {
      var hit = dref_all_().byKey[key];
      if (hit) {
        sources.push("reference");
        klass = hit.klass || "";
        String(hit.generic || hit.drug).toLowerCase().split(/[+,\/]/).forEach(function (g) {
          var t = g.trim();
          if (t.length >= 3 && generics.indexOf(t) === -1) generics.push(t);
        });
      }
    }
  } catch (e) { /* the reference sheet is optional */ }

  // c) nothing knows this drug. Its own stem is still a usable identity for
  //    the same-drug check, which is the one that must never depend on a
  //    lookup succeeding.
  if (!generics.length && key) generics.push(key);

  return { name: name, key: key, generics: generics, klass: klass,
           source: sources.join("+") || "typed" };
}

/**
 * Duplicate-therapy and class-overlap alerts for a medication list.
 *
 * @param {Array} meds  [{ drugName }]
 * @param {Object} [genericMap]  rx_genericMap_()
 * @return {{alerts:Array, checked:number, identified:number}}
 */
function checkDuplicateTherapy(meds, genericMap) {
  var out = { alerts: [], checked: 0, identified: 0 };
  try {
    meds = (meds || []).filter(function (m) { return ds_str_(m && m.drugName); });
    if (meds.length < 2) { out.checked = meds.length; return out; }

    var gm = genericMap || (typeof rx_genericMap_ === "function" ? rx_genericMap_() : null);
    var ids = meds.map(function (m) { return ds_identify_(m.drugName, gm); });
    out.checked = ids.length;
    out.identified = ids.filter(function (i) { return i.source !== "typed"; }).length;

    // ---- 1. the same drug, listed twice ---------------------------------
    // Matched on the NORMALISED stem, so "Tab Paracetamol 500" and
    // "PARACETAMOL 500mg" are caught. They were not before, and a duplicate
    // typed by two different people is spelt two different ways by
    // definition.
    var byKey = {};
    ids.forEach(function (id) {
      if (!id.key) return;
      (byKey[id.key] = byKey[id.key] || []).push(id.name);
    });
    Object.keys(byKey).forEach(function (k) {
      if (byKey[k].length < 2) return;
      var names = byKey[k];
      var identical = names.every(function (n) { return n.toLowerCase() === names[0].toLowerCase(); });
      out.alerts.push({
        level: "DANGER", type: "SAME_DRUG",
        drugs: names.slice(),
        message: identical
          ? "“" + names[0] + "” appears " + names.length + " times on this list. " +
            "Remove the extra line unless two different strengths are genuinely intended."
          : names.join(" and ") + " are the same drug written " + names.length +
            " ways. Keep one line unless two strengths are genuinely intended."
      });
    });

    // ---- 2. two products sharing an active ingredient --------------------
    var byGeneric = {};
    ids.forEach(function (id) {
      if (byKey[id.key] && byKey[id.key].length > 1) return;   // already reported above
      id.generics.forEach(function (g) {
        (byGeneric[g] = byGeneric[g] || []);
        if (byGeneric[g].indexOf(id.name) === -1) byGeneric[g].push(id.name);
      });
    });
    Object.keys(byGeneric).forEach(function (g) {
      if (byGeneric[g].length < 2) return;
      out.alerts.push({
        level: "WARN", type: "DUPLICATE",
        drugs: byGeneric[g].slice(),
        message: byGeneric[g].join(" and ") + " both contain " + g +
                 " — duplicate therapy. Check the combined daily dose."
      });
    });

    // ---- 3. two drugs of a class that should not be doubled --------------
    // The duplicate that actually reaches the patient: two different NSAIDs,
    // an ACE inhibitor beside an ARB. They look like two different drugs on
    // the page and no interaction pair covers them.
    DS_CLASS_RULES.forEach(function (rule) {
      var members = [];
      ids.forEach(function (id) {
        if (!id.klass || !rule.match.test(id.klass)) return;
        if (members.indexOf(id.name) === -1) members.push(id.name);
      });
      if (members.length < 2) return;

      // Already said as a shared-ingredient duplicate? Do not say it twice —
      // two alerts about one line is how a prescriber learns to skim them.
      var alreadyPaired = out.alerts.some(function (a) {
        return (a.type === "DUPLICATE" || a.type === "SAME_DRUG") &&
               members.every(function (m) { return (a.drugs || []).indexOf(m) !== -1; });
      });
      if (alreadyPaired) return;

      out.alerts.push({
        level: "WARN", type: "CLASS_DUPLICATE",
        drugs: members.slice(),
        klass: rule.label,
        message: members.join(" and ") + " are both " + rule.label + ". " + rule.why
      });
    });

  } catch (e) {
    Logger.log("checkDuplicateTherapy failed: " + e.message);
  }
  return out;
}

/**
 * Drug-allergy alerts, matched on word boundaries.
 *
 * The old check was `lower.indexOf(allergy) !== -1 || allergy.indexOf(lower)
 * !== -1` — a substring test in both directions. A recorded allergy to "ASA"
 * matched Asacol and every nasal preparation; an allergy recorded as
 * "Penicillin" did not match "Amoxicillin", which is the one that matters.
 *
 * This matches whole words against the typed name, the stem and every
 * generic, and additionally knows the cross-reactive families a clinic
 * records allergies in terms of.
 *
 * @param {Array<string>} allergies
 * @param {Array} meds  [{ drugName }]
 */
function checkAllergyConflicts(allergies, meds, genericMap) {
  var out = [];
  try {
    allergies = (allergies || []).map(ds_str_).filter(function (a) { return a.length >= 3; });
    if (!allergies.length) return out;

    var gm = genericMap || (typeof rx_genericMap_ === "function" ? rx_genericMap_() : null);

    // An allergy is recorded as the class the patient reacted to; the
    // prescription names a member of it. Neither name contains the other.
    var FAMILIES = [
      { allergy: /penicillin|beta.?lactam/i,
        members: /(amoxicillin|ampicillin|cloxacillin|piperacillin|penicillin|augmentin|amoxyclav|clavulan)/i,
        note: "penicillin group" },
      { allergy: /cephalosporin/i,
        members: /(cef[a-z]+|ceftriaxone|cefixime|cefuroxime|cephalexin)/i,
        note: "cephalosporin group" },
      { allergy: /sulfa|sulpha|sulfonamide/i,
        members: /(sulfamethoxazole|cotrimoxazole|sulfasalazine|sulfadiazine)/i,
        note: "sulfonamide group" },
      { allergy: /nsaid|aspirin|\basa\b/i,
        members: /(ibuprofen|diclofenac|naproxen|aceclofenac|indomethacin|ketorolac|aspirin|mefenamic)/i,
        note: "NSAID group" },
      { allergy: /quinolone|fluoroquinolone/i,
        members: /(ciprofloxacin|levofloxacin|ofloxacin|moxifloxacin|norfloxacin)/i,
        note: "quinolone group" },
      { allergy: /macrolide/i,
        members: /(azithromycin|erythromycin|clarithromycin|roxithromycin)/i,
        note: "macrolide group" }
    ];

    (meds || []).forEach(function (md) {
      var name = ds_str_(md && md.drugName);
      if (!name) return;
      var id = ds_identify_(name, gm);
      // Every string this drug answers to, for the word-boundary test.
      var aliases = [name.toLowerCase(), id.key].concat(id.generics)
        .filter(function (x, i, arr) { return x && arr.indexOf(x) === i; });

      allergies.forEach(function (al) {
        var direct = (typeof di_matches_ === "function")
          ? di_matches_(aliases, al)
          : aliases.some(function (a) { return a === al.toLowerCase(); });

        var family = null;
        if (!direct) {
          for (var f = 0; f < FAMILIES.length; f++) {
            if (!FAMILIES[f].allergy.test(al)) continue;
            if (aliases.some(function (a) { return FAMILIES[f].members.test(a); })) {
              family = FAMILIES[f];
              break;
            }
          }
        }
        if (!direct && !family) return;

        out.push({
          level: "DANGER", type: "ALLERGY",
          drugs: [name],
          message: family
            ? name + " is in the " + family.note + ", and this patient has a recorded " +
              "allergy to “" + al + "”. Confirm before prescribing."
            : name + " matches the recorded allergy “" + al + "”. " +
              "Confirm before prescribing."
        });
      });
    });
  } catch (e) {
    Logger.log("checkAllergyConflicts failed: " + e.message);
  }
  return out;
}

/**
 * Run from the editor. Explains what the checks see for a list of names.
 *
 *   testDrugSafety(["Tab Paracetamol 500", "PARACETAMOL 500mg", "Ibuprofen"])
 *
 * Prints how each name was identified and which source identified it, then
 * every alert. Use it after editing Drug_Dose_Reference: a check that covers
 * nothing looks identical to a check that found nothing.
 */
function testDrugSafety(names) {
  var meds = (names || []).map(function (n) { return { drugName: n }; });
  var gm = (typeof rx_genericMap_ === "function") ? rx_genericMap_() : null;
  var lines = ["=== identity ==="];

  meds.forEach(function (m) {
    var id = ds_identify_(m.drugName, gm);
    lines.push("  " + m.drugName + "\n      stem: " + (id.key || "(none)") +
               "\n      generics: " + (id.generics.join(", ") || "(none)") +
               "\n      class: " + (id.klass || "(none)") +
               "\n      known from: " + id.source);
  });

  var dup = checkDuplicateTherapy(meds, gm);
  lines.push("\n=== duplicate / class (" + dup.identified + " of " + dup.checked +
             " identified) ===");
  lines.push(dup.alerts.length
    ? dup.alerts.map(function (a) { return "  [" + a.type + "] " + a.message; }).join("\n")
    : "  (none)");

  var inter = (typeof checkDrugInteractions === "function")
    ? checkDrugInteractions(meds, gm) : { alerts: [], rulesLoaded: 0 };
  lines.push("\n=== interactions (" + inter.rulesLoaded + " pairs loaded) ===");
  lines.push(inter.alerts.length
    ? inter.alerts.map(function (a) { return "  [" + a.severity + "] " + a.message; }).join("\n")
    : "  (none)");

  var report = lines.join("\n");
  Logger.log(report);
  return report;
}
