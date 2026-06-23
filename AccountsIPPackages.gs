// =========================================================================
// 🏥 CRESCENTIA — IP PACKAGE MASTER + INSURERS
// Run setupPackages() ONCE to seed Package_Master. Edit rates in the sheet
// afterwards — no code change needed. getPackages()/getInsurers() feed the UI.
// Package cap = the bundled price (default = lower bound, editable at discharge).
// Stay_Days pre-fills the ward multipliers; Exclusions are billed on top.
// =========================================================================

function ipk_ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function ipk_sheet_(name, headers) {
  var ss = ipk_ss_(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}

// Run once from the editor.
function setupPackages() {
  var sh = ipk_sheet_('Package_Master', ['Code', 'Name', 'Amount_Min', 'Amount_Max', 'Default_Cap', 'Stay_Days', 'Room_Type', 'Inclusions', 'Exclusions']);
  // [Code, Name, min, max, defaultCap, stayDays, roomType, inclusions, exclusions]
  var P = [
    ['NORMDEL', 'Normal Delivery', 15000, 30000, 15000, 2, 'Standard', 'Obs-Gyn, 48h nursing, CBC+Blood Group, baby assessment', 'NICU, ICU, Blood'],
    ['LSCS', 'LSCS (C-Section)', 40000, 65000, 40000, 3, 'Standard', 'Obs-Gyn, assistant, anaesthetist, Major OT, 72h nursing, newborn assessment', 'NICU, ICU, Blood'],
    ['APPEND', 'Appendicectomy', 35000, 60000, 35000, 3, 'Standard', 'Surgeon, assistant, anaesthetist, Major OT, 72h nursing, dressing', 'ICU, Blood'],
    ['LAPCHOLE', 'Lap Cholecystectomy', 55000, 90000, 55000, 2, 'Standard', 'Surgeon, assistant, anaesthetist, Lap OT, 48h nursing, lap instruments', 'ICU'],
    ['HERNIA', 'Inguinal Hernia Repair', 35000, 70000, 35000, 2, 'Standard', 'Surgeon, assistant, anaesthetist, OT, 48h nursing, standard mesh', 'Premium mesh'],
    ['VHYST', 'Vaginal Hysterectomy', 60000, 100000, 60000, 4, 'Standard', 'Surgeon, assistant, anaesthetist, Major OT, 96h nursing, foley/dressings', 'ICU'],
    ['TAHBSO', 'TAH + BSO', 80000, 150000, 80000, 5, 'Standard', 'Surgeon, assistant, anaesthetist, Major OT, 120h nursing, post-op care', 'ICU, Blood'],
    ['TURP', 'TURP', 60000, 120000, 60000, 3, 'Standard', 'Urologist, anaesthetist, Endoscopic OT, 72h nursing, foley', 'ICU'],
    ['DJSTENT', 'DJ Stenting', 20000, 40000, 20000, 1, 'Standard', 'Urologist, anaesthetist, Endoscopic OT, 24h nursing, DJ stent', 'ICU'],
    ['URSL', 'URSL', 60000, 120000, 60000, 2, 'Standard', 'Urologist, anaesthetist, Endoscopic OT, 48h nursing, stone extraction', 'ICU'],
    ['PCNL', 'PCNL', 90000, 180000, 90000, 3, 'Standard', 'Urologist, assistant, anaesthetist, Major OT, 72h nursing, nephrostomy care', 'Blood'],
    ['PHACO', 'Cataract (Phaco)', 12000, 50000, 12000, 0, 'Day Care', 'Ophthalmologist, Ophthal OT, standard IOL', 'Premium IOL'],
    ['TKR', 'Total Knee Replacement', 200000, 450000, 200000, 5, 'Deluxe', 'Ortho, assistant, anaesthetist, Major OT, 1d ICU, 120h nursing, implant+physio', 'Revision implant'],
    ['THR', 'Total Hip Replacement', 250000, 500000, 250000, 5, 'Deluxe', 'Ortho, assistant, anaesthetist, Major OT, 1d ICU, 120h nursing, implant+physio', 'Revision implant'],
    ['ANGIO', 'Angioplasty (1 DES)', 150000, 300000, 150000, 2, 'ICU + Room', 'Cardiologist, Cath Lab, 24h ICU, cardiac panel, one DES stent', 'Additional stents'],
    ['DENGUE', 'Dengue Package', 15000, 35000, 15000, 3, 'Standard', 'Physician, 72h nursing, CBC monitoring, IV fluids', 'SDP/RDP'],
    ['MICU', 'Medical ICU (per day)', 8000, 20000, 8000, 1, 'ICU', 'Intensivist, 1:1 nursing, routine labs, basic ICU drugs, monitoring', 'Ventilator, Dialysis']
  ];
  // clear existing data rows, rewrite
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  sh.getRange(2, 1, P.length, P[0].length).setValues(P);
  SpreadsheetApp.flush();
  return { success: true, message: P.length + " packages seeded." };
}

function getPackages() {
  try {
    var sh = ipk_ss_().getSheetByName('Package_Master');
    if (!sh) return { success: true, packages: [], message: "Run setupPackages() first." };
    var d = sh.getDataRange().getValues(), out = [];
    for (var i = 1; i < d.length; i++) {
      if (!String(d[i][0]).trim()) continue;
      out.push({ code: String(d[i][0]), name: String(d[i][1]), min: Number(d[i][2]) || 0, max: Number(d[i][3]) || 0, cap: Number(d[i][4]) || 0, stayDays: Number(d[i][5]) || 0, roomType: String(d[i][6]), inclusions: String(d[i][7]), exclusions: String(d[i][8]) });
    }
    return { success: true, packages: out };
  } catch (e) { return { success: false, message: e.message }; }
}

// Insurers/TPAs. Reads Insurers_Master if present, else returns a default list.
function getInsurers() {
  try {
    var sh = ipk_ss_().getSheetByName('Insurers_Master'), out = [];
    if (sh) { var d = sh.getDataRange().getValues(); for (var i = 1; i < d.length; i++) if (String(d[i][0]).trim()) out.push(String(d[i][0])); }
    if (!out.length) out = ['Star Health', 'Care Health', 'Niva Bupa', 'HDFC ERGO', 'ICICI Lombard', 'Bajaj Allianz', 'Tata AIG', 'New India Assurance', 'United India', 'Oriental Insurance', 'National Insurance', 'SBI General', 'Aditya Birla Health', 'ManipalCigna', 'Medi Assist (TPA)', 'Paramount (TPA)', 'MDIndia (TPA)', 'Govt — PMJAY/CMCHIS', 'Other'];
    return { success: true, insurers: out };
  } catch (e) { return { success: false, message: e.message }; }
}