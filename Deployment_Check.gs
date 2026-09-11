// ============================================================================
// Deployment_Check.gs  —  Crescentia HealthTech
// Detects a PARTIAL deployment before it turns into mystery bugs.
// ----------------------------------------------------------------------------
// Apps Script has no module system: every .gs file shares one global scope, and
// a file that was not copied across simply leaves its functions undefined. The
// caller then throws a ReferenceError which some try/catch swallows, and the
// symptom surfaces somewhere unrelated —
//
//   Doctor_Core.gs stale  ->  "Admission IP2609-0002 was not found"
//                         ->  "Your role cannot post notes on this chart"
//                         ->  "Repair failed: dc_invalidate_ is not defined"
//
// all from one missing file. Run verifyDeployment() FIRST whenever the app
// behaves strangely after an update. It names the file to copy.
// ============================================================================

/**
 * Every function this codebase depends on across file boundaries, mapped to
 * the file that defines it. Extend this when you add a cross-file dependency.
 */
var DEP_MAP = {
  "Doctor_Core.gs": [
    "dc_headerMap_", "dc_col_", "dc_ensureColumn_", "dc_ensureSheet_",
    "dc_invalidate_", "dc_resetCache_", "dc_sheetValues_",
    "dc_str_", "dc_upper_", "dc_int_", "dc_money_", "dc_dateKey_",
    "dc_getDoctorById_", "resolveScope_", "resolveWriteDoctor_",
    "dc_inScope_", "getDoctorPickerContext",
    "dc_careTeamSheet_", "getIPCareTeam", "addIPCareTeamMember",
    "removeIPCareTeamMember", "dc_isOnCareTeam_",
    "dc_normaliseTeamRole_", "getIPTeamRoles"
  ],
  "Doctor_Session_Store.gs": ["dc_validateSession_", "dc_sessionName_", "revokeSession"],
  "Doctors_Engine.gs":       ["getTenantId_", "validateSession_", "issueSession_", "logAudit_", "getActiveDoctors"],
  "IP_Clinical_Access.gs": [
    "ipc_timelineSheet_", "ipc_casesheetSheet_", "resolveIPWrite_", "resolveIPRead_",
    "ipc_admissionRow_", "ipc_primaryDoctorId_", "ipc_ensurePrimaryOnCareTeam_",
    "ipc_mayWriteOnAdmission_", "ipc_mayReadAdmission_", "ipc_wardVisibilityFilter_",
    "ipc_authorLabel_", "getIPCareTeamPanel", "getIPNotePermissions",
    "getIPNotesBundle", "updateIPDiagnosis", "getIPDiagnosis",
    "verifyIPClinicalSchema", "verifyIPCareTeamCoverage", "repairDuplicateCareTeamRows"
  ],
  "IP_Casesheet_Engine.gs": [
    "getIPCasesheetWard", "getIPCasesheetContext", "saveIPCasesheet",
    "ipc_writeCasesheetRow_", "ipc_currentCasesheet_", "ipc_admissionsWithCasesheet_",
    "getIPCasesheetPrintHtml", "amendIPCasesheet", "getIPCasesheetForEdit",
    "getIPCasesheetVersions", "ipc_safe_"
  ],
  "IP_Notes_Logic.gs": [
    "saveIPNote", "getIPTimeline", "getClinicalContext", "getActiveIPAdmissionsForNotes",
    "_normDrug_", "_ensureIPPharmacyQueueSheet_", "fetchPharmacyMasterForIP",
    "getIPNotesPrintHtml", "getIPNotesLabCatalog", "generateIPHandoverSummary"
  ],
  "IP_Schema_Repair.gs": [
    "repairCasesheetHeaderDrift", "runIPHealthCheck", "repairDuplicatePharmacyQueueRows"
  ],
  "Drug_Interactions.gs": [
    "di_sheet_", "setupDrugInteractions", "checkDrugInteractions", "di_rules_"
  ],
  "OP_Rx_Engine.gs":        ["checkRxSafety", "getPatientAllergies", "savePatientAllergies",
                             "rx_genericMap_", "rx_seedSheet_", "getComposerContext",
                             "listRxBundles", "buildTaperPlan", "suggestPaediatricDose"],
  "OP_Templates_Engine.gs": ["getScopedTemplates", "learnTemplatesScoped_", "listConsultTemplates",
                             "applyConsultTemplate", "getExamDefaults", "getAssistForDiagnosis",
                             "opt_seedAssistFor_", "opt_parseList_", "opt_dxTokens_"],
  "OP_Database_Engine.gs":  ["fetchOPDrugMaster", "fetchUniversalDrugs", "saveOPEncounter"],
  "Lab_OPD_Bridge.gs":      ["getOPDOrderableTests", "createOPDLabOrder"],
  "LabIntegrationEngine.gs":["createLabRequest", "getOrderableTests"],
  "IP_Admissions_Logic.gs": ["getActiveIPWard", "saveNewAdmissionLedger"],
  "Barcode_Engine.gs":      ["bc_nextPatientId_", "bc_nextDailyId_", "getScanRouterConfig",
                             "resolveScan", "barcodeUpdateAppointment", "bc_auditScan_",
                             "auditDuplicatePatientIds", "getPatientLabelData",
                             "getSampleLabelData", "verifyCollectionIdentity",
                             "receiveLabSampleByBarcode"]
};

/**
 * ENTRY POINT. Reports any file whose functions are missing from this project.
 * Returns a human-readable report; also logs it.
 */
function verifyDeployment() {
  var missingByFile = {};
  var totalMissing = 0, totalChecked = 0;

  Object.keys(DEP_MAP).forEach(function (file) {
    DEP_MAP[file].forEach(function (fn) {
      totalChecked++;
      var present = false;
      try {
        // Reading an undefined global throws; a defined one yields a function.
        present = (typeof this[fn] === "function") ||
                  (eval("typeof " + fn) === "function");
      } catch (e) { present = false; }
      if (!present) {
        if (!missingByFile[file]) missingByFile[file] = [];
        missingByFile[file].push(fn);
        totalMissing++;
      }
    });
  });

  var files = Object.keys(missingByFile);
  var out = [];
  out.push("DEPLOYMENT CHECK  —  " + (totalChecked - totalMissing) + "/" + totalChecked +
           " expected functions present.");
  out.push("");

  if (!files.length) {
    out.push("All files are deployed and up to date.");
  } else {
    out.push("INCOMPLETE DEPLOYMENT. Copy these files into the Apps Script");
    out.push("project again, then re-run this check:");
    out.push("");
    files.forEach(function (f) {
      out.push("  " + f);
      out.push("      missing: " + missingByFile[f].join(", "));
    });
    out.push("");
    out.push("Until every file above is updated, expect symptoms that look");
    out.push("unrelated to the missing file — 'admission not found', 'your role");
    out.push("cannot post notes', and repairs that fail on an undefined helper");
    out.push("are all the same cause.");
  }

  var report = out.join("\n");
  Logger.log(report);
  return report;
}

/**
 * Run this after any update. Checks the deployment FIRST, and only then the
 * data — a schema repair against a half-deployed project reports nonsense.
 */
function runFullHealthCheck() {
  var dep = verifyDeployment();
  if (dep.indexOf("INCOMPLETE DEPLOYMENT") !== -1) {
    var msg = dep + "\n\nSTOPPING: data repairs were not run. Fix the deployment first.";
    Logger.log(msg);
    return msg;
  }
  return dep + "\n\n" + runIPHealthCheck();
}

// ---------------------------------------------------------------------------
// DUPLICATE DEFINITIONS
//
// Apps Script has one global scope across every .gs file. When two files define
// the same function, the one loaded LAST silently wins and the other is dead
// code — with no warning anywhere. This project has several, including doGet,
// the web app's entry point.
//
// This does not guess: it prints the first line of the body that is actually
// live, so you can see which file won.
// ---------------------------------------------------------------------------

var DUP_SUSPECTS = [
  "doGet", "include",
  "getAppointmentsByDate", "submitNewAppointment", "updateAppointmentStatus",
  "getAvailableTimeSlots", "getPatientDemographics", "fetchDailyLedger",
  "formatTimeSafely", "getPatientDashboardStats", "saveAdminAvailability",
  "getFinanceDashboard", "recordBankTransfer",
  "accd_day_", "accd_month_", "accd_income_", "accd_isCash_", "accd_ipSettlements_",
  "getPendingCreditBills", "searchLabRecords", "_esc"
];

/**
 * Reports which definition of each duplicated function is live.
 * Run it, then delete the losing copy from the file that is NOT shown.
 */
function whichDefinitionWins() {
  var out = ["LIVE DEFINITIONS OF DUPLICATED FUNCTIONS", ""];
  out.push("Each of these names is defined in more than one .gs file. Only the");
  out.push("body shown below is running; the other copy is dead code. Delete the");
  out.push("copy that does NOT match, so future edits land where they take effect.");
  out.push("");

  DUP_SUSPECTS.forEach(function (name) {
    var src;
    try { src = eval(name + ".toString()"); }
    catch (e) { out.push(name + "  ->  NOT DEFINED"); return; }

    // First two non-empty lines of the body identify which copy this is.
    var lines = String(src).split("\n")
      .map(function (l) { return l.trim(); })
      .filter(function (l) { return l && l.indexOf("//") !== 0; })
      .slice(0, 3);
    out.push(name + "()");
    lines.forEach(function (l) {
      out.push("    " + (l.length > 110 ? l.substring(0, 110) + "..." : l));
    });
    out.push("");
  });

  out.push("KNOWN PAIRS IN THIS PROJECT:");
  out.push("  doGet, include, getPatientDashboardStats, saveAdminAvailability");
  out.push("      -> CodeMV.gs  vs  PatientDetails.gs");
  out.push("  getAppointmentsByDate, submitNewAppointment, updateAppointmentStatus,");
  out.push("  getAvailableTimeSlots, getPatientDemographics, fetchDailyLedger,");
  out.push("  formatTimeSafely");
  out.push("      -> Appointment.gs  vs  CodeMV.gs");
  out.push("  getFinanceDashboard, recordBankTransfer, accd_*");
  out.push("      -> AccountsDashboard.gs  vs  Accounts_Dashboard.gs   (whole file duplicated)");
  out.push("  getPendingCreditBills  -> Pharmacy.gs vs PharmacyReturnsBackend.gs");
  out.push("  searchLabRecords       -> LabIntegrationEngine.gs vs LabRecordsLogic.gs");
  out.push("  _esc                   -> LabBillingLogic.gs vs LabIntegrationEngine.gs");
  out.push("");
  out.push("doGet matters most: CodeMV.gs takes (e) and serves WhatsApp lab-report");
  out.push("links; PatientDetails.gs takes no argument and cannot. If the second is");
  out.push("live, those links are broken.");

  var report = out.join("\n");
  Logger.log(report);
  return report;
}