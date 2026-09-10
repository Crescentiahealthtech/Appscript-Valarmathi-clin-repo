// ============================================================================
// OP_Templates_Engine.gs  —  Crescentia HealthTech
// PHASE 4a : Doctor-specific clinical templates
// ----------------------------------------------------------------------------
// REQUIRES: Doctor_Core.gs, Doctor_Session_Store.gs, OP_Doctor_Engine.gs
//
// FOUR LAYERS
//   1. Phrase memory      — Clinical_Templates, now scoped PERSONAL / CLINIC
//   2. Consult templates  — OP_Consult_Templates (named order sets, NO meds)
//   3. Exam defaults      — Doctor_Exam_Defaults (replaces hardcoded JS object)
//   4. Dx -> drug hints   — derived live from the doctor's OWN encounters
//
// SAFETY BOUNDARY
//   Consult templates deliberately carry no medications. Drug suggestions are
//   a separate, explicit path: they are shown with the evidence count behind
//   them, capped, drawn only from the requesting doctor's own history, and
//   never written into a prescription row without the doctor clicking.
// ============================================================================

var OPT_MAX_SUGGESTIONS   = 5;
var OPT_SUGGESTION_CACHE  = 1800;   // 30 min
var OPT_PROMOTE_MIN_USES  = 5;
var OPT_PROMOTE_MIN_DOCS  = 2;

// ============================================================================
// SECTION A — SCHEMA
// ============================================================================

/** Run once from the editor after adding this file. Idempotent. */
function runTemplateMigration() {
  var lock = LockService.getScriptLock();
  var out = [];
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- Clinical_Templates: scope the existing global library ------------
    var ct = ss.getSheetByName("Clinical_Templates");
    if (!ct) {
      ct = ss.insertSheet("Clinical_Templates");
      ct.appendRow(["Category", "Text", "UseCount", "Doctor_ID", "Scope"]);
      ct.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9ead3");
      out.push("Clinical_Templates created.");
    } else {
      dc_ensureColumn_(ct, "Doctor_ID");
      dc_ensureColumn_(ct, "Scope");
      // Everything that already exists was clinic-wide. Say so explicitly.
      var scopeIdx = dc_col_(ct, "Scope");
      var last = ct.getLastRow();
      if (last > 1) {
        var rng = ct.getRange(2, scopeIdx + 1, last - 1, 1);
        var vals = rng.getValues();
        var filled = 0;
        for (var i = 0; i < vals.length; i++) {
          if (dc_str_(vals[i][0]) === "") { vals[i][0] = "CLINIC"; filled++; }
        }
        if (filled) rng.setValues(vals);
        out.push("Clinical_Templates: " + filled + " legacy row(s) marked CLINIC.");
      }
    }

    // --- New sheets --------------------------------------------------------
    dc_ensureSheet_(ss, "OP_Consult_Templates", [
      "Template_ID", "Tenant_ID", "Doctor_ID", "Scope", "Name", "Specialty",
      "Complaints", "History", "Exam_JSON", "Diagnosis", "Diagnosis_Type",
      "Labs_JSON", "Advice", "Review_Days", "Use_Count",
      "Created_By", "Created_At", "Updated_At", "Status"
    ]);
    dc_ensureSheet_(ss, "Doctor_Exam_Defaults", [
      "Doctor_ID", "Tenant_ID", "CVS", "RS", "PA", "CNS", "Updated_At"
    ]);
    out.push("OP_Consult_Templates / Doctor_Exam_Defaults ready.");

    // --- OP_Encounters: stop discarding radiology and diagnosis type ------
    var op = ss.getSheetByName("OP_Encounters");
    if (op) {
      ["Radiology_Orders", "Radiology_Findings", "Diagnosis_Type"].forEach(function (h) {
        dc_ensureColumn_(op, h);
      });
      out.push("OP_Encounters: radiology + diagnosis-type columns added.");
    } else {
      out.push("OP_Encounters absent — skipped.");
    }

    SpreadsheetApp.flush();
    return out.join("\n");
  } catch (e) {
    return "TEMPLATE MIGRATION FAILED: " + e.message;
  } finally {
    lock.releaseLock();
  }
}

function opt_phraseSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Clinical_Templates",
    ["Category", "Text", "UseCount", "Doctor_ID", "Scope"]);
}

function opt_consultSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "OP_Consult_Templates", [
    "Template_ID", "Tenant_ID", "Doctor_ID", "Scope", "Name", "Specialty",
    "Complaints", "History", "Exam_JSON", "Diagnosis", "Diagnosis_Type",
    "Labs_JSON", "Advice", "Review_Days", "Use_Count",
    "Created_By", "Created_At", "Updated_At", "Status"
  ]);
}

function opt_examSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Doctor_Exam_Defaults",
    ["Doctor_ID", "Tenant_ID", "CVS", "RS", "PA", "CNS", "Updated_At"]);
}

// ============================================================================
// SECTION B — LAYER 1 : SCOPED PHRASE MEMORY
// ============================================================================

/**
 * FRONTEND ENTRY. Replaces fetchClinicalTemplates() in the OP module.
 * Personal phrases rank above clinic phrases; each tier sorted by use count.
 * @return {{success, CC:[], HX:[], ADVICE:[]}}  items = {text, count, scope, id}
 */
function getScopedTemplates(doctorId, sessionToken) {
  var empty = { CC: [], HX: [], ADVICE: [] };
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) {
      return { success: false, message: scope.message, CC: [], HX: [], ADVICE: [] };
    }
    var me = dc_upper_(dc_str_(doctorId) || scope.selfDoctorId);

    var sh = opt_phraseSheet_();
    var m = dc_headerMap_(sh);
    var data = sh.getDataRange().getDisplayValues();

    var personal = { CC: [], HX: [], ADVICE: [] };
    var clinic   = { CC: [], HX: [], ADVICE: [] };

    for (var i = 1; i < data.length; i++) {
      var cat = dc_upper_(data[i][0]);
      var text = dc_str_(data[i][1]);
      if (!text || !empty[cat]) continue;

      var owner = (m["Doctor_ID"] === undefined) ? "" : dc_upper_(data[i][m["Doctor_ID"]]);
      var sc    = (m["Scope"] === undefined) ? "CLINIC" : (dc_upper_(data[i][m["Scope"]]) || "CLINIC");

      var item = {
        rowId: i + 1,
        text: text,
        count: dc_int_(data[i][2]),
        scope: sc
      };

      if (sc === "PERSONAL") {
        if (me && owner === me) personal[cat].push(item);
      } else {
        clinic[cat].push(item);
      }
    }

    var byCount = function (a, b) { return b.count - a.count; };
    var out = { success: true, CC: [], HX: [], ADVICE: [] };

    ["CC", "HX", "ADVICE"].forEach(function (cat) {
      personal[cat].sort(byCount);
      clinic[cat].sort(byCount);
      // Personal first — a doctor's own wording should always outrank the house set.
      out[cat] = personal[cat].concat(clinic[cat]);
    });

    return out;
  } catch (e) {
    return { success: false, message: "Templates unavailable: " + e.message,
             CC: [], HX: [], ADVICE: [] };
  }
}

/**
 * Records phrases against the doctor who actually used them.
 * Called from saveOPEncounterScoped — replaces the global learnTemplates()
 * for the OP path. Existing callers of learnTemplates() are unaffected.
 */
function learnTemplatesScoped_(items, doctorId) {
  if (!items || !items.length) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    var sh = opt_phraseSheet_();
    var m = dc_headerMap_(sh);
    var docIdx   = m["Doctor_ID"];
    var scopeIdx = m["Scope"];
    var me = dc_upper_(doctorId);
    if (!me) return;

    var data = sh.getDataRange().getDisplayValues();

    // Index existing rows by category + text + owner.
    var index = {};
    for (var i = 1; i < data.length; i++) {
      var owner = (docIdx === undefined) ? "" : dc_upper_(data[i][docIdx]);
      var key = dc_upper_(data[i][0]) + "|" +
                dc_str_(data[i][1]).toLowerCase() + "|" + owner;
      index[key] = i + 1;
    }

    var appends = [];
    items.forEach(function (it) {
      var cat = dc_upper_(it.category);
      var text = dc_str_(it.text);
      if (!text || ["CC", "HX", "ADVICE"].indexOf(cat) === -1) return;

      var personalKey = cat + "|" + text.toLowerCase() + "|" + me;
      var clinicKey   = cat + "|" + text.toLowerCase() + "|";

      if (index[personalKey]) {
        var r = index[personalKey];
        sh.getRange(r, 3).setValue(dc_int_(sh.getRange(r, 3).getValue()) + 1);
      } else if (index[clinicKey]) {
        // Already in the house library — count it there, don't duplicate.
        var r2 = index[clinicKey];
        sh.getRange(r2, 3).setValue(dc_int_(sh.getRange(r2, 3).getValue()) + 1);
      } else {
        var row = [cat, text, 1];
        while (row.length <= Math.max(docIdx || 0, scopeIdx || 0)) row.push("");
        if (docIdx !== undefined)   row[docIdx] = me;
        if (scopeIdx !== undefined) row[scopeIdx] = "PERSONAL";
        appends.push(row);
        index[personalKey] = -1;   // guard against duplicates in one batch
      }
    });

    if (appends.length) {
      var width = sh.getLastColumn();
      appends.forEach(function (r) { while (r.length < width) r.push(""); });
      sh.getRange(sh.getLastRow() + 1, 1, appends.length, width).setValues(appends);
    }

    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log("learnTemplatesScoped_ error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * FRONTEND ENTRY. Lets a doctor prune their own phrase list.
 * This is the mitigation for learned typos — pruning, not a use threshold.
 */
function deletePhraseTemplate(rowId, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired." };

    var sh = opt_phraseSheet_();
    var m = dc_headerMap_(sh);
    var r = dc_int_(rowId);
    if (r < 2 || r > sh.getLastRow()) return { success: false, message: "Phrase not found." };

    var row = sh.getRange(r, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    var sc = (m["Scope"] === undefined) ? "CLINIC" : dc_upper_(row[m["Scope"]]);
    var owner = (m["Doctor_ID"] === undefined) ? "" : dc_upper_(row[m["Doctor_ID"]]);

    if (sc === "CLINIC") {
      var role = dc_str_(sess.role).toLowerCase();
      if (role !== "admin") {
        return { success: false, message: "Only an administrator can remove a shared phrase." };
      }
    } else if (owner !== dc_upper_(sess.doctorId)) {
      return { success: false, message: "You can only remove your own phrases." };
    }

    sh.deleteRow(r);
    SpreadsheetApp.flush();
    return { success: true, message: "Phrase removed." };
  } catch (e) {
    return { success: false, message: "Could not remove phrase: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * FRONTEND ENTRY (admin). Promotes a personal phrase into the shared library.
 * Gated so the house set doesn't fill with one doctor's idiosyncrasies.
 */
function promotePhraseToClinic(rowId, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired." };
    if (dc_str_(sess.role).toLowerCase() !== "admin") {
      return { success: false, message: "Only an administrator can promote a phrase." };
    }

    var sh = opt_phraseSheet_();
    var m = dc_headerMap_(sh);
    var data = sh.getDataRange().getDisplayValues();
    var r = dc_int_(rowId);
    if (r < 2 || r > data.length) return { success: false, message: "Phrase not found." };

    var row = data[r - 1];
    var cat = dc_upper_(row[0]);
    var text = dc_str_(row[1]).toLowerCase();

    // How many distinct doctors use this phrase, and how often in total?
    var totalUses = 0, doctors = {};
    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== cat) continue;
      if (dc_str_(data[i][1]).toLowerCase() !== text) continue;
      totalUses += dc_int_(data[i][2]);
      var d = (m["Doctor_ID"] === undefined) ? "" : dc_upper_(data[i][m["Doctor_ID"]]);
      if (d) doctors[d] = true;
    }
    var docCount = Object.keys(doctors).length;

    if (totalUses < OPT_PROMOTE_MIN_USES || docCount < OPT_PROMOTE_MIN_DOCS) {
      return { success: false,
               message: "Needs at least " + OPT_PROMOTE_MIN_USES + " uses across " +
                        OPT_PROMOTE_MIN_DOCS + " doctors (currently " + totalUses +
                        " uses, " + docCount + " doctor(s))." };
    }

    if (m["Scope"] !== undefined)     sh.getRange(r, m["Scope"] + 1).setValue("CLINIC");
    if (m["Doctor_ID"] !== undefined) sh.getRange(r, m["Doctor_ID"] + 1).setValue("");

    SpreadsheetApp.flush();
    logAudit_(sess, "PHRASE_PROMOTE", "Clinical_Template", dc_str_(row[1]), { category: cat });
    return { success: true, message: "Phrase added to the clinic library." };
  } catch (e) {
    return { success: false, message: "Could not promote phrase: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// SECTION C — LAYER 2 : CONSULT TEMPLATES (ORDER SETS, NO MEDICATIONS)
// ============================================================================

/** FRONTEND ENTRY. Templates this doctor may use: their own + the clinic set. */
function listConsultTemplates(doctorId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) return { success: false, message: scope.message, templates: [] };

    var me = dc_upper_(dc_str_(doctorId) || scope.selfDoctorId);
    var tenant = getTenantId_();

    var sh = opt_consultSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var personal = [], clinic = [];

    for (var i = 1; i < data.length; i++) {
      if (dc_str_(data[i][1]) !== tenant) continue;
      if (dc_upper_(data[i][18]) === "DELETED") continue;

      var sc = dc_upper_(data[i][3]) || "CLINIC";
      var owner = dc_upper_(data[i][2]);
      if (sc === "PERSONAL" && owner !== me) continue;

      var item = {
        templateId: dc_str_(data[i][0]),
        scope: sc,
        ownerDoctorId: owner,
        name: dc_str_(data[i][4]),
        specialty: dc_str_(data[i][5]),
        diagnosis: dc_str_(data[i][9]),
        useCount: dc_int_(data[i][14]),
        editable: (sc === "PERSONAL" && owner === me)
      };
      (sc === "PERSONAL" ? personal : clinic).push(item);
    }

    var byUse = function (a, b) { return b.useCount - a.useCount; };
    personal.sort(byUse);
    clinic.sort(byUse);

    return {
      success: true,
      templates: personal.concat(clinic),
      message: (personal.length + clinic.length) ? "" :
        "No consultation templates yet. Save one from a completed consult."
    };
  } catch (e) {
    return { success: false, message: "Could not load templates: " + e.message, templates: [] };
  }
}

/**
 * FRONTEND ENTRY. Full template body, and increments its use count.
 * Returns NO medications by design — drugs never travel inside a template.
 */
function applyConsultTemplate(templateId, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var scope = resolveScope_(sessionToken, null);
    if (!scope.ok) return { success: false, message: scope.message };

    var sh = opt_consultSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var target = dc_upper_(templateId);
    var me = dc_upper_(scope.selfDoctorId);

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== target) continue;
      if (dc_upper_(data[i][18]) === "DELETED") {
        return { success: false, message: "This template has been removed." };
      }
      var sc = dc_upper_(data[i][3]) || "CLINIC";
      if (sc === "PERSONAL" && dc_upper_(data[i][2]) !== me) {
        return { success: false, message: "That template belongs to another doctor." };
      }

      var exam = {}, labs = [];
      try { exam = data[i][8] ? JSON.parse(data[i][8]) : {}; } catch (e) {}
      try { labs = data[i][11] ? JSON.parse(data[i][11]) : []; } catch (e) {}

      sh.getRange(i + 1, 15).setValue(dc_int_(data[i][14]) + 1);
      SpreadsheetApp.flush();

      return {
        success: true,
        template: {
          templateId: dc_str_(data[i][0]),
          name: dc_str_(data[i][4]),
          complaints: dc_str_(data[i][6]),
          history: dc_str_(data[i][7]),
          exam: exam,
          diagnosis: dc_str_(data[i][9]),
          diagnosisType: dc_str_(data[i][10]),
          labs: labs,
          advice: dc_str_(data[i][12]),
          reviewDays: dc_int_(data[i][13])
        },
        message: "Template applied. Review every field before saving."
      };
    }
    return { success: false, message: "Template not found." };
  } catch (e) {
    return { success: false, message: "Could not apply template: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * FRONTEND ENTRY. Create or update a consult template.
 * payload = { templateId?, doctorId?, scope, name, specialty, complaints,
 *             history, exam:{cvs,rs,pa,cns}, diagnosis, diagnosisType,
 *             labs:[{testName,source}], advice, reviewDays }
 * Any `meds` key on the payload is IGNORED — templates never carry drugs.
 */
function saveConsultTemplate(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var w = resolveWriteDoctor_(sessionToken, payload && payload.doctorId);
    if (!w.ok) return { success: false, message: w.message };

    var name = dc_str_(payload.name);
    if (!name) return { success: false, message: "Give the template a name." };

    var requested = dc_upper_(payload.scope) || "PERSONAL";
    var role = dc_str_(w.sess.role).toLowerCase();
    if (requested === "CLINIC" && role !== "admin") {
      return { success: false,
               message: "Only an administrator can add to the shared clinic library." };
    }

    var sh = opt_consultSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var now = new Date();
    var tenant = getTenantId_();

    var examJson = JSON.stringify({
      cvs: dc_str_((payload.exam || {}).cvs),
      rs:  dc_str_((payload.exam || {}).rs),
      pa:  dc_str_((payload.exam || {}).pa),
      cns: dc_str_((payload.exam || {}).cns)
    });
    var labsJson = JSON.stringify(payload.labs || []);

    // ---- update -----------------------------------------------------------
    var existingId = dc_str_(payload.templateId);
    if (existingId) {
      for (var i = 1; i < data.length; i++) {
        if (dc_upper_(data[i][0]) !== dc_upper_(existingId)) continue;

        var sc = dc_upper_(data[i][3]) || "CLINIC";
        if (sc === "PERSONAL" && dc_upper_(data[i][2]) !== dc_upper_(w.doctorId)) {
          return { success: false, message: "You can only edit your own templates." };
        }
        if (sc === "CLINIC" && role !== "admin") {
          return { success: false, message: "Only an administrator can edit a shared template." };
        }

        var r = i + 1;
        sh.getRange(r, 5).setValue(String(name));
        sh.getRange(r, 6).setValue(String(dc_str_(payload.specialty)));
        sh.getRange(r, 7).setValue(String(dc_str_(payload.complaints)));
        sh.getRange(r, 8).setValue(String(dc_str_(payload.history)));
        sh.getRange(r, 9).setValue(String(examJson));
        sh.getRange(r, 10).setValue(String(dc_str_(payload.diagnosis)));
        sh.getRange(r, 11).setValue(String(dc_str_(payload.diagnosisType)));
        sh.getRange(r, 12).setValue(String(labsJson));
        sh.getRange(r, 13).setValue(String(dc_str_(payload.advice)));
        sh.getRange(r, 14).setValue(dc_int_(payload.reviewDays));
        sh.getRange(r, 18).setValue(now);

        SpreadsheetApp.flush();
        logAudit_(w.sess, "TEMPLATE_UPDATE", "OP_Consult_Template", existingId, { name: name });
        return { success: true, templateId: existingId, message: "Template updated." };
      }
      return { success: false, message: "Template not found." };
    }

    // ---- create -----------------------------------------------------------
    for (var j = 1; j < data.length; j++) {
      if (dc_upper_(data[j][18]) === "DELETED") continue;
      if (dc_str_(data[j][4]).toLowerCase() !== name.toLowerCase()) continue;
      var jsc = dc_upper_(data[j][3]) || "CLINIC";
      if (jsc === requested &&
          (jsc === "CLINIC" || dc_upper_(data[j][2]) === dc_upper_(w.doctorId))) {
        return { success: false, message: "You already have a template called \"" + name + "\"." };
      }
    }

    var newId = "TPL-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    sh.appendRow([
      String(newId), String(tenant),
      String(requested === "CLINIC" ? "" : w.doctorId),
      String(requested), String(name), String(dc_str_(payload.specialty)),
      String(dc_str_(payload.complaints)), String(dc_str_(payload.history)),
      String(examJson), String(dc_str_(payload.diagnosis)),
      String(dc_str_(payload.diagnosisType)), String(labsJson),
      String(dc_str_(payload.advice)), dc_int_(payload.reviewDays), 0,
      String(w.sess.username), now, now, "ACTIVE"
    ]);

    SpreadsheetApp.flush();
    logAudit_(w.sess, "TEMPLATE_CREATE", "OP_Consult_Template", newId,
              { name: name, scope: requested });
    return { success: true, templateId: newId, message: "Template \"" + name + "\" saved." };

  } catch (e) {
    return { success: false, message: "Could not save the template: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** FRONTEND ENTRY. Soft-delete — the row stays for audit. */
function deleteConsultTemplate(templateId, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired." };

    var sh = opt_consultSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var target = dc_upper_(templateId);
    var role = dc_str_(sess.role).toLowerCase();

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== target) continue;
      var sc = dc_upper_(data[i][3]) || "CLINIC";
      if (sc === "PERSONAL" && dc_upper_(data[i][2]) !== dc_upper_(sess.doctorId)) {
        return { success: false, message: "You can only delete your own templates." };
      }
      if (sc === "CLINIC" && role !== "admin") {
        return { success: false, message: "Only an administrator can delete a shared template." };
      }
      sh.getRange(i + 1, 19).setValue("DELETED");
      sh.getRange(i + 1, 18).setValue(new Date());
      SpreadsheetApp.flush();
      logAudit_(sess, "TEMPLATE_DELETE", "OP_Consult_Template", templateId, {});
      return { success: true, message: "Template removed." };
    }
    return { success: false, message: "Template not found." };
  } catch (e) {
    return { success: false, message: "Could not delete the template: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// SECTION D — LAYER 3 : PER-DOCTOR EXAM DEFAULTS
// ============================================================================

var OPT_FALLBACK_EXAM = {
  cvs: "S1S2+, No Murmur",
  rs:  "B/L NVBS, No added sounds",
  pa:  "Soft, Non-tender, No organomegaly",
  cns: "Conscious, Oriented, NFND"
};

/** FRONTEND ENTRY. This doctor's normal-findings text, or the house default. */
function getExamDefaults(doctorId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) {
      return { success: true, defaults: OPT_FALLBACK_EXAM, isCustom: false };
    }
    var me = dc_upper_(dc_str_(doctorId) || scope.selfDoctorId);
    if (!me) return { success: true, defaults: OPT_FALLBACK_EXAM, isCustom: false };

    var sh = opt_examSheet_();
    var data = sh.getDataRange().getDisplayValues();

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== me) continue;
      return {
        success: true,
        isCustom: true,
        defaults: {
          cvs: dc_str_(data[i][2]) || OPT_FALLBACK_EXAM.cvs,
          rs:  dc_str_(data[i][3]) || OPT_FALLBACK_EXAM.rs,
          pa:  dc_str_(data[i][4]) || OPT_FALLBACK_EXAM.pa,
          cns: dc_str_(data[i][5]) || OPT_FALLBACK_EXAM.cns
        }
      };
    }
    return { success: true, defaults: OPT_FALLBACK_EXAM, isCustom: false };
  } catch (e) {
    return { success: true, defaults: OPT_FALLBACK_EXAM, isCustom: false };
  }
}

/** FRONTEND ENTRY. payload = { doctorId?, cvs, rs, pa, cns } */
function saveExamDefaults(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var w = resolveWriteDoctor_(sessionToken, payload && payload.doctorId);
    if (!w.ok) return { success: false, message: w.message };

    var sh = opt_examSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var now = new Date();
    var me = dc_upper_(w.doctorId);

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== me) continue;
      var r = i + 1;
      sh.getRange(r, 3, 1, 5).setValues([[
        String(dc_str_(payload.cvs)), String(dc_str_(payload.rs)),
        String(dc_str_(payload.pa)),  String(dc_str_(payload.cns)), now
      ]]);
      SpreadsheetApp.flush();
      return { success: true, message: "Examination defaults saved." };
    }

    sh.appendRow([
      String(w.doctorId), String(getTenantId_()),
      String(dc_str_(payload.cvs)), String(dc_str_(payload.rs)),
      String(dc_str_(payload.pa)),  String(dc_str_(payload.cns)), now
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: "Examination defaults saved." };
  } catch (e) {
    return { success: false, message: "Could not save defaults: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// SECTION E — LAYER 4 : DIAGNOSIS -> DRUG SUGGESTIONS
// ----------------------------------------------------------------------------
// Derived live from the requesting doctor's OWN encounters. No lookup table,
// so a suggestion can never drift from what is actually in the record.
//
// CLINICAL SAFETY — these constraints are deliberate, not incidental:
//   * suggestions come only from the requesting doctor's own history
//   * every item carries the evidence count it was derived from
//   * capped at OPT_MAX_SUGGESTIONS
//   * the caller must never pre-fill a prescription row from this
// ============================================================================

var OPT_STOPWORDS = {
  "the": 1, "and": 1, "with": 1, "for": 1, "acute": 1, "chronic": 1,
  "mild": 1, "moderate": 1, "severe": 1, "suspected": 1, "probable": 1,
  "type": 1, "left": 1, "right": 1, "under": 1, "evaluation": 1
};

function opt_dxTokens_(text) {
  var out = [];
  String(text || "").toLowerCase().split(/[^a-z0-9]+/).forEach(function (t) {
    if (t.length < 4) return;
    if (OPT_STOPWORDS[t]) return;
    out.push(t);
  });
  return out;
}

/**
 * FRONTEND ENTRY. Drugs this doctor has previously prescribed for a similar
 * diagnosis. Returns [] rather than an error when there is no history.
 */
function getDiagnosisDrugSuggestions(diagnosis, doctorId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) return { success: false, message: scope.message, suggestions: [] };

    var me = dc_str_(doctorId) || scope.selfDoctorId;
    var dx = dc_str_(diagnosis);
    if (!me || !dx) return { success: true, suggestions: [], matchedEncounters: 0 };

    var tokens = opt_dxTokens_(dx);
    if (!tokens.length) return { success: true, suggestions: [], matchedEncounters: 0 };

    var cacheKey = "DXDRUG_" + dc_upper_(me) + "_" +
                   Utilities.base64EncodeWebSafe(tokens.join("_")).substring(0, 60);
    var cache = CacheService.getScriptCache();
    try {
      var hit = cache.get(cacheKey);
      if (hit) return JSON.parse(hit);
    } catch (e) { /* cache is optional */ }

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OP_Encounters");
    if (!sh || sh.getLastRow() < 2) return { success: true, suggestions: [], matchedEncounters: 0 };

    var m = dc_headerMap_(sh);
    var docCol = (m["Doctor_ID"] === undefined) ? -1 : m["Doctor_ID"];
    var data = sh.getDataRange().getDisplayValues();
    var meU = dc_upper_(me);

    var tally = {}, matched = 0;

    for (var i = data.length - 1; i >= 1; i--) {
      // Cheapest filters first — this sheet grows without bound.
      if (docCol !== -1) {
        var rowDoc = dc_str_(data[i][docCol]) || DC_DEFAULT_DOCTOR;
        if (dc_upper_(rowDoc) !== meU) continue;
      }
      var rowDx = data[i][21];
      if (!rowDx) continue;

      var rowTokens = opt_dxTokens_(rowDx);
      var overlap = false;
      for (var t = 0; t < tokens.length && !overlap; t++) {
        if (rowTokens.indexOf(tokens[t]) !== -1) overlap = true;
      }
      if (!overlap) continue;

      var meds = [];
      try { meds = data[i][22] ? JSON.parse(data[i][22]) : []; } catch (e) { continue; }
      if (!meds.length) continue;

      matched++;
      var seenThisEncounter = {};
      meds.forEach(function (md) {
        var nm = dc_str_(md.drugName);
        if (!nm) return;
        var key = nm.toLowerCase();
        if (seenThisEncounter[key]) return;      // count each drug once per visit
        seenThisEncounter[key] = true;

        if (!tally[key]) {
          tally[key] = { drugName: nm, strength: dc_str_(md.strength),
                         sig: dc_str_(md.sig), duration: dc_str_(md.duration),
                         count: 0 };
        }
        tally[key].count++;
        // Keep the most recent sig, since prescribing habits evolve.
        if (tally[key].count === 1) {
          tally[key].sig = dc_str_(md.sig);
          tally[key].duration = dc_str_(md.duration);
        }
      });

      if (matched >= 60) break;   // recent history is what matters
    }

    // ---- stock status, so a delisted drug is flagged not hidden ----------
    var stock = {};
    try {
      (fetchOPDrugMaster() || []).forEach(function (d) {
        stock[String(d.brand).toLowerCase()] = d.status;
      });
    } catch (e) { /* non-fatal */ }

    var suggestions = Object.keys(tally).map(function (k) { return tally[k]; })
      .sort(function (a, b) { return b.count - a.count; })
      .slice(0, OPT_MAX_SUGGESTIONS)
      .map(function (s) {
        var st = stock[s.drugName.toLowerCase()];
        return {
          drugName: s.drugName, strength: s.strength,
          sig: s.sig, duration: s.duration,
          count: s.count, outOf: matched,
          evidence: "Prescribed in " + s.count + " of " + matched + " similar consults",
          stockStatus: st || "unknown",
          inPharmacy: !!st
        };
      });

    var result = {
      success: true,
      suggestions: suggestions,
      matchedEncounters: matched,
      message: suggestions.length ? "" : "No prior prescribing history for this diagnosis."
    };

    try { cache.put(cacheKey, JSON.stringify(result), OPT_SUGGESTION_CACHE); } catch (e) {}
    return result;

  } catch (e) {
    return { success: false, message: "Could not load suggestions: " + e.message,
             suggestions: [] };
  }
}

// ============================================================================
// SECTION F — FRONTEND ADAPTER: single assist call for the diagnosis box
// Wraps getDiagnosisDrugSuggestions and adds this doctor's own lab + advice
// habits for the same diagnosis. Returns drug CLASSES as [] by design —
// there is no class table, and inventing one is not safe.
// ============================================================================
function getAssistForDiagnosis(diagnosis, doctorId, sessionToken) {
  var blank = { success: true, matched: "", yours: [], classes: [],
                labs: [], advice: "", reviewDays: 0, message: "" };
  try {
    var dx = dc_str_(diagnosis);
    if (!dx) return blank;

    var drugs = getDiagnosisDrugSuggestions(dx, doctorId, sessionToken);
    if (!drugs || !drugs.success) {
      return { success: false, message: (drugs && drugs.message) || "Assist unavailable.",
               yours: [], classes: [], labs: [], advice: "", reviewDays: 0 };
    }

    var labs = [], advice = "";
    var scope = resolveScope_(sessionToken, doctorId);
    var me = dc_upper_(dc_str_(doctorId) || (scope.ok ? scope.selfDoctorId : ""));

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OP_Encounters");
    if (sh && sh.getLastRow() > 1) {
      var m = dc_headerMap_(sh);
      var docCol = (m["Doctor_ID"] === undefined) ? -1 : m["Doctor_ID"];
      var data = sh.getDataRange().getDisplayValues();
      var tokens = opt_dxTokens_(dx);
      var labTally = {}, adviceTally = {}, scanned = 0;

      for (var i = data.length - 1; i >= 1 && scanned < 60; i--) {
        if (me && docCol !== -1) {
          var rowDoc = dc_str_(data[i][docCol]) || DC_DEFAULT_DOCTOR;
          if (dc_upper_(rowDoc) !== me) continue;
        }
        var rowDx = data[i][21];              // V — Diagnosis
        if (!rowDx) continue;

        var rowTokens = opt_dxTokens_(rowDx), hit = false;
        for (var t = 0; t < tokens.length && !hit; t++) {
          if (rowTokens.indexOf(tokens[t]) !== -1) hit = true;
        }
        if (!hit) continue;
        scanned++;

        try {                                 // X — Labs JSON
          (JSON.parse(data[i][23] || "[]") || []).forEach(function (l) {
            if (dc_upper_(l.type) !== "ORDER") return;
            var n = dc_str_(l.testName);
            if (n) labTally[n] = (labTally[n] || 0) + 1;
          });
        } catch (e) { /* malformed row — skip, never fail the call */ }

        // Y — Advice. saveOPEncounter appends "| Future Labs: ..." — drop that.
        var adv = dc_str_(data[i][24]).split("|")[0].trim();
        if (adv) adviceTally[adv] = (adviceTally[adv] || 0) + 1;
      }

      labs = Object.keys(labTally)
        .sort(function (a, b) { return labTally[b] - labTally[a]; })
        .slice(0, 6);

      var advKeys = Object.keys(adviceTally)
        .sort(function (a, b) { return adviceTally[b] - adviceTally[a]; });
      // One prior use is a coincidence. Two is a habit worth surfacing.
      if (advKeys.length && adviceTally[advKeys[0]] >= 2) advice = advKeys[0];
    }

    var n = drugs.matchedEncounters || 0;
    var sugg = drugs.suggestions || [];
    return {
      success: true,
      matched: n ? (n + " similar consult" + (n === 1 ? "" : "s")) : "",
      yours: sugg,
      classes: [],
      labs: labs,
      advice: advice,
      reviewDays: 0,
      message: (sugg.length || labs.length || advice) ? "" :
               "Nothing on file for this diagnosis yet."
    };
  } catch (e) {
    return { success: false, message: "Assist unavailable: " + e.message,
             yours: [], classes: [], labs: [], advice: "", reviewDays: 0 };
  }
}