// =========================================================================
// 📦 MODULE LOADER — serves the application in bundles instead of one page
// Crescentia HealthTech / CresRx
// -------------------------------------------------------------------------
// Index.html used to inline all 47 partials into a single document: 1.56 MB
// of markup and script that every user downloaded, parsed and executed on
// every load, whatever their role. A pharmacist paid for the 186 KB of ward
// notes; a doctor paid for the whole accounts suite. Nothing could paint
// until all of it had been parsed.
//
// Index.html now carries a ~165 KB shell — the theme, sign-in, the rail, the
// modals, the router and the landing screens — and the rest is fetched from
// here, one bundle per area of work, on first navigation into it. After
// sign-in the client also prewarms the bundles the signed-in role actually
// uses, in the background, so the first click into a module is usually
// already paid for.
//
// The bundles hold UI only: markup, styles and client script, identical to
// what was inlined before. No patient data passes through this file, and
// nothing here reads a session — exactly as when Index.html served the same
// markup to an anonymous page load.
// =========================================================================

/**
 * name -> the partials that make it up, in the order they must be injected.
 *
 * Two rules govern which partial belongs where:
 *
 *   1. Everything a rail click lands on is either in the shell or in a
 *      bundle keyed to that click, so no navigation can reach a screen
 *      whose markup is absent.
 *   2. A partial goes in the SAME bundle as the code its inline handlers
 *      call. Admin_Accounts_Workspace is the accounts hub and every tile on
 *      it calls AccX/AccDash, so it ships with the accounts scripts rather
 *      than in the shell.
 *
 * The very large clinical screens are split one per bundle. A doctor opening
 * the case sheet should not wait on the notes module, and a single reply
 * carrying both would be 360 KB of string through google.script.run.
 */
var CRESC_BUNDLES = {
  'patient':       ['PatientApp'],
  'appointments':  ['Admin_Appointments', 'Admin_Availability'],
  'patients':      ['Admin_Patient_Register', 'Admin_Patient_Search'],
  'billing':       ['Admin_Hospital_Billing'],

  'emr-ward':      ['IP_Admissions_UI', 'IP_Ledger_UI', 'IP_BedMap_UI'],
  'emr-casesheet': ['IP_Casesheet_UI'],
  'emr-notes':     ['IP_Notes_UI'],
  'emr-op':        ['OP_Module'],
  'emr-records':   ['EMR_MasterTimeline_UI', 'IP_Records_UI'],
  'emr-discharge': ['DS_Desk_UI', 'DS_Desk_Scripts', 'DS_Discharge_Flow'],

  'pharmacy':      ['PharmacyStockAddUI', 'PharmacyInventoryLedger',
                    'Pharmacy_Billing_UI', 'PharmacyDeskReturnsUI',
                    'PharmacyDashboardUI'],
  'lab':           ['Admin_Lab_Workspace', 'LabModuleUI', 'LabBillingUI',
                    'LabRecordsUI'],
  'accounts':      ['Admin_Accounts_Workspace', 'Accounts_Scripts',
                    'Accounts_Shifts_Scripts', 'Accounts_Dashboard_Scripts',
                    'Accounts_Discharge_Script', 'AccountsPayablesScripts',
                    'AccountsTaxAuditsScripts', 'AccountsInsuranceScripts',
                    'AccountsIPChargesUI'],
  'barcode':       ['Barcode_QR_Lib', 'Barcode_Labels', 'Barcode_Router']
};

/**
 * FRONTEND ENTRY. One bundle's markup, ready to inject.
 *
 * @param {string} name  a key of CRESC_BUNDLES
 * @return {{success:boolean, name:string, html:string, message:string}}
 */
function crescGetBundle(name) {
  try {
    var key = String(name || '').trim();
    var files = CRESC_BUNDLES[key];
    if (!files) {
      return { success: false, name: key, html: '',
               message: 'Unknown module bundle "' + key + '".' };
    }
    var parts = [];
    for (var i = 0; i < files.length; i++) {
      // A missing file is named rather than silently dropped: a bundle that
      // arrives half-built produces a screen with no explanation, which is
      // the hardest kind of fault to trace back to a deployment.
      try {
        parts.push(HtmlService.createHtmlOutputFromFile(files[i]).getContent());
      } catch (e) {
        return { success: false, name: key, html: '',
                 message: 'Module file "' + files[i] + '" is missing from the ' +
                          'Apps Script project, so the ' + key + ' screens cannot load.' };
      }
    }
    return { success: true, name: key, html: parts.join('\n'), message: '' };
  } catch (err) {
    return { success: false, name: String(name || ''), html: '',
             message: 'Bundle load failed: ' + err.message };
  }
}

/**
 * Every partial referenced by a bundle exists, and none is claimed twice.
 * Run from the script editor after adding or renaming a UI file.
 */
function crescVerifyBundles() {
  var seen = {}, problems = [], count = 0;
  Object.keys(CRESC_BUNDLES).forEach(function (key) {
    CRESC_BUNDLES[key].forEach(function (f) {
      count++;
      if (seen[f]) problems.push(f + ' is in both "' + seen[f] + '" and "' + key + '".');
      seen[f] = key;
      try {
        HtmlService.createHtmlOutputFromFile(f);
      } catch (e) {
        problems.push(f + ' (bundle "' + key + '") does not exist.');
      }
    });
  });
  var report = problems.length
    ? 'crescVerifyBundles: ' + problems.length + ' problem(s)\n' + problems.join('\n')
    : 'crescVerifyBundles: ' + count + ' file(s) across ' +
      Object.keys(CRESC_BUNDLES).length + ' bundle(s), all present, none duplicated.';
  Logger.log(report);
  return { success: problems.length === 0, message: report };
}
