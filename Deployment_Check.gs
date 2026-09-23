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
  // Shared_Dates.gs is listed FIRST because everything else depends on it:
  // every module's own date helper now delegates here, so a project missing
  // this one file loses dates on every screen at once.
  "Shared_Dates.gs": [
    "cresc_parseDate_", "cresc_formatDate_", "cresc_dateOnly_", "cresc_dayKey_",
    "cresc_timeText_", "cresc_dateTimeText_", "cresc_daysBetween_", "cresc_los_",
    "cresc_ms_", "cresc_isSheetEpoch_"
  ],
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
  "Doctor_Session_Store.gs": ["dc_validateSession_", "dc_sessionName_", "revokeSession_"],
  // Pharmacy.gs reads the signed discharge script through these, so the
  // take-home prescription reaches the counter. Both are optional at runtime
  // (a project without the discharge module just contributes nothing), but
  // listing them means verifyDeployment() names the file when they are gone.
  "DS_Data.gs": ["dsx_latestSnapshotOfType_", "dsx_summaryIdFor_", "dsx_unpackPayload_",
                 "dsx_upgradePayload_", "dsx_toDate_", "dsx_getHeader_",
                 "dsx_getWorking_", "dsx_sectionIsEmpty_"],
  "Appointment.gs":          ["apt_newId_", "submitNewAppointment", "getAppointmentsByDate"],
  "Doctors_Engine.gs":       ["getTenantId_", "validateSession_", "issueSession_", "logAudit_", "getActiveDoctors_"],
  // The permission matrix and the guard. Every module that carries a
  // crescRequire_() call depends on this file being present, so a deployment
  // missing it must be named rather than discovered one refusal at a time.
  "RBAC.gs": ["crescRequire_", "crescActor_", "crescCan_", "crescPermsFor_",
              "crescRequireOwnRecord_", "crescGetMyPermissions", "cresc_reason_",
              "crescRbacSelfTest", "crescRbacCoverage",
              // The system-owner flag. Auth_Credentials.gs reads it to decide
              // who may create, disable or reset a doctor or an administrator;
              // with this file stale, nobody can and nothing says why.
              "crescIsElevated_", "crescEnsureSuperAdminColumn_"],
  // Sign-in audit and lockout. AuthLogin.gs calls into these on every
  // attempt, so without this file nobody can sign in at all.
  "Auth_Audit.gs": ["crescAuthAudit_", "crescAuthGuard_", "crescAuthFailed_",
                    "crescAuthPassed_", "crescLogSignOut", "crescUnlockAccount",
                    "crescGetLoginAudit"],
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
    "getIPNotesPrintHtml", "getIPNotesLabCatalog"
  ],
  "IP_Schema_Repair.gs": [
    "repairCasesheetHeaderDrift", "runIPHealthCheck", "repairDuplicatePharmacyQueueRows"
  ],
  "Drug_Interaction.gs": [
    "di_sheet_", "setupDrugInteractions", "checkDrugInteractions_", "di_rules_"
  ],
  "Accounts_Shifts.gs": [
    "acc_cashSince_", "acc_counterKey_", "getShiftState", "openShift",
    "closeShift", "getAllDrawerStates"
  ],
  "DPDP_Compliance.gs": [
    "dpdpSetup", "dpdpSetGrievanceOfficer", "getDPDPNotice", "recordConsent",
    "getConsentStatus", "withdrawConsent", "raiseDPDPRequest",
    "listDPDPRequests", "closeDPDPRequest", "exportPatientData",
    "dpdpRegisterSharedFile_", "dpdpExpireSharedLinks", "dpdpRetentionReport",
    "dpdpReadinessCheck", "dpdpVerifyRequester", "dpdpSubmitPublicRequest",
    "recordNomination", "getNomination", "revokeNomination", "dpdpLogRead_",
    "dpdpSaveGrievanceOfficer", "dpdpConsoleSnapshot", "dpdpConsoleHousekeeping",
    "dpdpRetentionReportUI"
  ],
  "DPDP_Documents.gs": [
    "dpdpIssueDocumentLink_", "dpdpServeDocument_", "dpdpRevokeDocumentLink",
    "dpdpListDocumentLinks", "dpdpExpireDocumentGrants_", "dpdp_unpublish_"
  ],
  "Lab_Patient_View.gs": ["getPatientLabResults", "lpv_resultsFor_"],
  // IP Records reads the discharge summary's own sections through the
  // discharge engine's version resolver rather than parsing DS_Snapshots
  // itself. Optional at runtime — a deployment without the discharge module
  // shows the admission fact and nothing else — but listing it means
  // verifyDeployment() names the file when it is gone.
  "DS_Workflow.gs": ["dsx_resolveRef_"],
  "Deployment_Probe.gs": [
    "depProbeRecord_", "dep_probeSummary_", "depDeploymentFinding",
    "dep_attestation_", "dpdpConfirmDeploymentAccess", "RUN_deploymentEvidence"
  ],
  "Patient_Portal.gs": [
    "portalHome", "portalBookableDoctors", "portalDoctorSlots",
    "portalBookAppointment", "portalCancelAppointment", "portalLabResults",
    "portalRecords", "pp_me_", "pp_freeSlots_", "pp_serialisable_"
  ],
  "Maternal_Child.gs": [
    "mcSetup", "mcGetSchedule", "mcOpenAntenatal", "mcRecordAntenatalVisit",
    "mcCloseAntenatal", "mcRecordImmunisation", "mcGetAntenatal",
    "mcGetImmunisation", "mcListOpenAntenatal",
    "mc_antenatalView_", "mc_immunisationView_"
  ],
  "DS_Record_Bridge.gs": [
    "dsx_writeSummaryToTimeline_", "dsxBackfillSummaryNotes",
    "getDischargeSummaryFull", "mtPrintOPEncounter",
    "dsb_orderedSections_", "dsb_sectionText_"
  ],
  "DPDP_Consent_Backfill.gs": [
    "dpdpBackfillConsent", "dpdpConsentQueue", "dpdp_consentMatrix_",
    "dpdp_allPatients_", "RUN_consentBackfill_DRYRUN",
    "RUN_consentBackfill_FOR_REAL", "RUN_consentQueue"
  ],
  "DPDP_Dispatch.gs": [
    "dpdpRequireDispatchConsent_", "dpdp_resolvePatientFor_", "dpdp_consentState_",
    "getDispatchChannelNotice", "dpdpDispatchChannels_", "getDispatchConsent",
    "recordDispatchConsent", "dpdpDispatchReadiness"
  ],
  "DPDP_Breach.gs": [
    "dpdp_anomalyScan_", "dpdpAnomalyScan", "dpdpRaiseBreach",
    "dpdpAssessBreach", "dpdpRecordBreachNotification", "dpdpCloseBreach",
    "dpdpListBreaches", "dpdpBreachNotice"
  ],
  "RUN_Setup.gs": [
    "RUN_01_checkFilesArrived", "RUN_02_checkRoleMatrix", "RUN_03_checkPasswordStorage",
    "RUN_04_timePasswordHashing", "RUN_04b_makeSignInFaster", "RUN_05_createRegisters",
    "RUN_06_nameGrievanceOfficer", "RUN_07a_migrateCredentials_DRYRUN",
    "RUN_07b_migrateCredentials_FOR_REAL", "RUN_08_installScheduledJobs",
    "RUN_09_checkScheduledJobs", "RUN_10a_revokeOldPublicLinks_DRYRUN",
    "RUN_10b_revokeOldPublicLinks_FOR_REAL", "RUN_11a_eraseUnusedFields_DRYRUN",
    "RUN_11b_eraseUnusedFields_FOR_REAL", "RUN_12_readinessCheck",
    "RUN_13_retentionReport", "RUN_14_auditReviewThisWeek",
    "RUN_90_resetOnePassword", "RUN_91_setVoicePolicy", "RUN_99_removeScheduledJobs"
  ],
  "DPDP_Triggers.gs": [
    "dpdpInstallTriggers", "dpdpRemoveTriggers", "dpdpTriggerStatus",
    "dpdpDailyMaintenance", "dpdpWeeklyReview", "dpdpMonthlyRetentionReport"
  ],
  "Auth_Credentials.gs": [
    "crescPwdEncode_", "crescPwdVerify_", "crescPwdIsHashed_", "crescPwdEquals_",
    "crescRandomPassword_", "crescChangePassword", "crescAdminResetPassword",
    "crescCredentialStatus", "crescMigrateCredentials", "crescPwdBenchmark",
    "crescCreateStaffAccount", "crescSetAccountActive", "crescListStaffAccounts",
    "crescUserAdminScope"
  ],
  "Auth_Reset.gs": [
    "crescRequestPasswordReset", "cresc_sendResetEmail_", "cresc_resetThrottle_",
    "crescResetReadiness"
  ],
  "Clinic_Profile.gs": [
    "cresc_clinic_", "getClinicProfile", "setClinicProfile",
    "cresc_clinicWebsite_", "cresc_publicBaseUrl_", "cresc_publicLinkBase_"
  ],
  "Patient_Profile_Edit.gs": [
    "updatePatientProfile", "getPatientProfileHistory", "ppe_validate_"
  ],
  "Lab_Cancellation.gs": [
    "labEnsureCancellationColumns", "labCancelPendingOrder", "labCancelBill",
    "labListCancelled", "labx_billFor_", "labx_orderStatus_"
  ],
  "Doctor_Visiting.gs": [
    "dv_sheet_", "dv_isVisitingSlot_", "dv_identityFor_", "dv_signature_",
    "getVisitingConsultantState", "declareVisitingConsultant",
    "endVisitingConsultant", "listVisitingConsultants", "setupVisitingConsultant"
  ],
  "OP_Doctor_Engine.gs": [
    "op_referralSheet_", "createOPReferral", "respondToOPReferral",
    "listMyOPReferrals", "getOPReferralLetter"
  ],
  "Drug_Safety.gs": [
    "ds_identify_", "checkDuplicateTherapy_", "checkAllergyConflicts_",
    "testDrugSafety"
  ],
  "Dose_Reference.gs": [
    "dref_sheet_", "dref_all_", "getDoseReference", "listDoseReference",
    "setupDoseReference"
  ],
  "OP_Rx_Engine.gs":        ["checkRxSafety", "getPatientAllergies", "savePatientAllergies",
                             "rx_genericMap_", "rx_seedSheet_", "getComposerContext",
                             "listRxBundles", "buildTaperPlan", "suggestPaediatricDose"],
  "OP_Templates_Engine.gs": ["getScopedTemplates", "learnTemplatesScoped_", "listConsultTemplates",
                             "applyConsultTemplate", "getExamDefaults", "getAssistForDiagnosis",
                             "opt_seedAssistFor_", "opt_parseList_", "opt_dxTokens_"],
  "OP_Database_Engine.gs":  ["fetchOPDrugMaster", "fetchUniversalDrugs", "saveOPEncounter_"],
  "Lab_OPD_Bridge.gs":      ["getOPDOrderableTests", "createOPDLabOrder"],
  "LabIntegrationEngine.gs":["createLabRequest", "getOrderableTests"],
  "IP_Admissions_Logic.gs": ["getActiveIPWard", "saveNewAdmissionLedger"],
  "Barcode_Engine.gs":      ["bc_nextPatientId_", "bc_nextDailyId_", "getScanRouterConfig",
                             "resolveScan", "barcodeUpdateAppointment", "bc_auditScan_",
                             "auditDuplicatePatientIds", "getPatientLabelData",
                             "getSampleLabelData", "verifyCollectionIdentity",
                             "receiveLabSampleByBarcode"],
  // Every printed document asks this file for the patient-ID barcode. It is
  // called defensively, so a missing file costs the barcode and not the
  // document — which is precisely why it needs naming here instead.
  "Barcode_Print.gs":       ["bcp_code128_", "bcp_code128Svg_", "bcp_patientBarcodeBlock_"],
  // The discharge summary engine. Its absence used to surface as
  // "ds_getStatusMap is not a function" on the IP Admissions screen and as a
  // Discharge Desk that spun for ever, neither of which names a file.
  "DS_Data.gs":             ["dsx_summariesSheet_", "dsx_headerMap_", "dsx_newSection_",
                             "dsx_wire_", "dsx_readColumns_", "dsx_jsonNormalize_"],
  "DS_Workflow.gs":         ["ds_getQueue", "ds_getStatusMap", "ds_getSummary",
                             "ds_getSummaryLite", "ds_getSnapshots",
                             "ds_initiateDischarge", "ds_saveWorking", "ds_sign",
                             "ds_getDiff", "ds_getSourceItem", "ds_heartbeatEditing",
                             "dsx_requireRole_", "dsx_permissions_", "dsx_userRow_",
                             "dsx_fitForWire_"],
  "DS_Assembly.gs":         ["dsx_assemble_", "dsx_admissionRow_"],
  // Hospital billing. Without this file the billing desk boots into
  // "hb_getBootstrap is not a function", which names nothing useful.
  "Hospital_Billing.gs":    ["setupHospitalBilling", "hb_getBootstrap", "hb_getPatientContext",
                             "hb_saveInvoice", "hb_getInvoices", "hb_getInvoice",
                             "hb_recordPayment", "acc_hospitalRows_", "hb_billedApptIds_"],
  "DS_Print.gs":            ["ds_getPrintHtml"],
  "DS_Gate.gs":             ["dsx_gateCheck_", "ds_getGateStatus"]
};

/**
 * ENTRY POINT. Reports any file whose functions are missing from this project.
 * Returns a human-readable report; also logs it.
 */
function verifyDeployment() {
  crescEditorOnly_('verifyDeployment');
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
  crescEditorOnly_('runFullHealthCheck');
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
// Apps Script has one global scope across every .gs file: when two files
// define the same function, the one loaded LAST silently wins. This file used
// to carry whichDefinitionWins(), a hand-kept list of suspected pairs that
// had gone stale — none of them are duplicated any more. The check now runs
// against the source itself, before anything is pasted in:
//
//     node tools/dupes.js        (part of ./tools/check.sh)
// ---------------------------------------------------------------------------