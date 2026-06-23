// =========================================================================
// 💰 CRESCENTIA HEALTHTECH — ACCOUNTS CORE (Phase 1 & 2: Virtual Merge)
// Model: Finance_Master_Ledger holds EXPENSES/MANUAL entries only. 
// Income is read live from source billing sheets (Pharmacy, Lab, Appointments).
// The "Master Ledger" view is a VIRTUAL merge — no duplicate posting, no double-counting.
// =========================================================================

var ACC_CFG = {
  LEDGER: 'Finance_Master_Ledger',
  TAX: 'Tax_Ledger',
  AUDIT: 'Audit_Event_Ledger',
  AR: 'Accounts_Receivable',
  CLAIMS: 'Insurance_Claims_Ledger',
  SHIFTS: 'Shift_Registers',
  PH_INVOICES: 'Pharmacy_Invoices',
  LAB_BILLING: 'LAB_BILLING',
  TZ: 'Asia/Kolkata',
  VARIANCE_FLAG: 500,        // cash mismatch threshold ₹
  TPA_DELAY_DAYS: 90,        // delayed claim threshold
  LOCK_PROP: 'ACC_LOCKED_PERIODS' // ScriptProperties key -> CSV of YYYY-MM
};

// ---------- Small Helpers ----------
function acc_sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error("Sheet missing: " + name + ". Run Accounts DB setup first.");
  return sh;
}
function acc_now_() { return Utilities.formatDate(new Date(), ACC_CFG.TZ, "yyyy-MM-dd HH:mm:ss"); }
function acc_money_(v) { var n = parseFloat(v); return isNaN(n) ? 0 : Math.round(n * 100) / 100; }
function acc_str_(v) { return (v === null || v === undefined) ? "" : String(v); }
function acc_period_(d) {
  var dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) dt = new Date();
  return Utilities.formatDate(dt, ACC_CFG.TZ, "yyyy-MM");
}
function acc_toDate_(v) {
  if (v instanceof Date) return v;
  var d = new Date(v);
  return isNaN(d.getTime()) ? new Date() : d;
}
function acc_lockedSet_() {
  var raw = PropertiesService.getScriptProperties().getProperty(ACC_CFG.LOCK_PROP) || "";
  return raw ? raw.split(",").filter(String) : [];
}
function acc_isLocked_(period) { return acc_lockedSet_().indexOf(period) !== -1; }

// Append an immutable audit line. Never throws upward (best-effort logging).
function acc_audit_(user, action, module, refId, oldVal, newVal, reason) {
  try {
    var sh = acc_sheet_(ACC_CFG.AUDIT);
    var id = "AUD-" + Date.now().toString().slice(-8);
    sh.appendRow([
      id, acc_now_(), acc_str_(user) || "SYSTEM", acc_str_(action),
      acc_str_(module), acc_str_(refId), acc_str_(oldVal), acc_str_(newVal), acc_str_(reason)
    ]);
  } catch (e) { /* swallow: audit must not break the txn */ }
}

// =========================================================================
// VIRTUAL MERGE NORMALIZERS (Source of Truth Engine)
// =========================================================================

// Helper to convert sheet into JSON objects safely
function acc_readObjects_(sheetName) {
  try {
    var sh = acc_sheet_(sheetName);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return [];
    var headers = data[0];
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var obj = { _row: i + 1 };
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j]; // mapped by exact header name
      }
      out.push(obj);
    }
    return out;
  } catch(e) { return []; }
}

// 1. Pharmacy Income Normalizer 
function acc_pharmaRows_() {
  return acc_readObjects_(ACC_CFG.PH_INVOICES).map(function(r) {
    var stat = acc_str_(r['Payment_Status'] || r['PaymentStatus'] || r['Pay_Status']).toUpperCase();
    var net = acc_money_(r['Net'] || r['Net_Amount'] || r['NetAmount'] || r['Total_Amount']);
    var d = acc_toDate_(r['Timestamp'] || r['Date']);
    return {
      source: 'Pharmacy',
      billId: acc_str_(r['Invoice_No'] || r['Invoice_ID'] || r['InvoiceID'] || r['Invoice ID'] || r[0]),
      patientId: acc_str_(r['Patient_ID'] || r['PatientID'] || r['Patient ID']),
      name: acc_str_(r['Patient_Name'] || r['PatientName'] || r['Patient Name']),
      admissionId: '',
      net: net,
      balance: (stat === 'PENDING' || stat === 'CREDIT') ? net : 0,
      mode: acc_str_(r['Payment_Mode'] || r['PaymentMode'] || r['Pay_Mode']) || 'Cash',
      billDate: d, realizedDate: d,
      realized: (stat === 'PAID'),
      open: (stat === 'PENDING' || stat === 'CREDIT'),
      _row: r._row
    };
  });
}

// 2. Lab Income Normalizer
function acc_labRows_() {
  return acc_readObjects_(ACC_CFG.LAB_BILLING).map(function(r) {
    var stat = acc_str_(r['PaymentStatus']).toUpperCase();
    var net = acc_money_(r['NetAmount']);
    var bal = acc_money_(r['BalanceAmount']);
    var d = acc_toDate_(r['BilledAt'] || r['Timestamp'] || r['Date']);
    return {
      source: 'Lab',
      billId: acc_str_(r['BillID']),
      patientId: acc_str_(r['PatientID']),
      name: acc_str_(r['PatientName']),
      admissionId: acc_str_(r['AdmissionID']),
      net: net,
      balance: bal || (stat === 'ON_ACCOUNT' ? net : 0),
      mode: acc_str_(r['PaymentMode']) || 'Cash',
      billDate: d, realizedDate: d,
      realized: (stat === 'PAID'),
      open: (stat === 'ON_ACCOUNT'),
      _row: r._row
    };
  });
}

// 3. OP Consultation Income Normalizer 
function acc_opRows_() {
  return acc_readObjects_('Appointments').map(function (r) {
    var status = acc_str_(r['Status']).toUpperCase();
    var fee = acc_money_(r['Fee']);
    var d = acc_toDate_(r['Timestamp'] || r['Date']);
    return {
      source: 'OP_Consultation', 
      billId: acc_str_(r['Appt_ID'] || r['Appt ID'] || r['ApptId']),
      patientId: acc_str_(r['Patient_ID'] || r['Patient ID']), 
      name: acc_str_(r['Patient_Name'] || r['Patient Name']),
      admissionId: '', 
      net: fee, 
      balance: 0, 
      mode: 'Cash', // Default OP collections to Cash
      billDate: d, realizedDate: d,
      realized: (status === 'COMPLETED' && fee > 0),
      open: false,
      _row: r._row
    };
  });
}

// =========================================================================
// READ: Dashboard Payload (Virtual Merge of Ledger + Clinical Sources)
// =========================================================================
function getAccountsDashboard() {
  try {
    var thisPeriod = acc_period_(new Date());
    var todayStr = Utilities.formatDate(new Date(), ACC_CFG.TZ, "yyyy-MM-dd");
    var masterList = [];

    // 1. Pull Expenses & Manual entries from Finance_Master_Ledger
    try {
      var lData = acc_sheet_(ACC_CFG.LEDGER).getDataRange().getValues();
      for (var i = 1; i < lData.length; i++) {
        var r = lData[i];
        if (!r[0] || r[0] === "Txn_ID") continue; 
        var ts = acc_toDate_(r[1]);
        masterList.push({
          txnId: acc_str_(r[0]),
          ts: ts, 
          period: acc_period_(ts),
          category: acc_str_(r[3]),
          entity: acc_str_(r[6]),
          mode: acc_str_(r[7]),
          amtIn: acc_money_(r[8]),
          amtOut: acc_money_(r[9]),
          locked: (acc_str_(r[12]).toUpperCase() === 'TRUE') || acc_isLocked_(acc_period_(ts)),
          isToday: Utilities.formatDate(ts, ACC_CFG.TZ, "yyyy-MM-dd") === todayStr
        });
      }
    } catch(e) {}

    // 2. Pull Clinical Income (Virtual Merge)
    var virtualIncome = acc_pharmaRows_().concat(acc_labRows_()).concat(acc_opRows_());
    virtualIncome.forEach(function(inc) {
      if (inc.realized) {
        var ts = inc.realizedDate;
        masterList.push({
          txnId: inc.billId,
          ts: ts,
          period: acc_period_(ts),
          category: inc.source,
          entity: inc.name || inc.patientId || "Walk-In",
          mode: inc.mode,
          amtIn: inc.net,
          amtOut: 0,
          locked: acc_isLocked_(acc_period_(ts)),
          isToday: Utilities.formatDate(ts, ACC_CFG.TZ, "yyyy-MM-dd") === todayStr
        });
      }
    });

    // 3. Sort chronologically (Newest first)
    masterList.sort(function(a, b) { return b.ts.getTime() - a.ts.getTime(); });

    // --- OPTION A: Live Drawer Cash (Tied to OPEN Shifts) ---
    var openDrawerCash = 0;
    var earliestOpen = Infinity;
    var hasOpenShift = false;

    try {
      // UPGRADE: We now use acc_readObjects_ to find data by Header Name, 
      // rendering it immune to column shuffling or index mismatches!
      var shiftRows = acc_readObjects_(ACC_CFG.SHIFTS);
      
      shiftRows.forEach(function(s) {
        if (acc_str_(s['Status']).toUpperCase() === 'OPEN') { 
          hasOpenShift = true;
          openDrawerCash += acc_money_(s['Opening_Cash']); // Add Opening Float
          
          var openTime = acc_toDate_(s['Timestamp']); 
          if (openTime.getTime() < earliestOpen) {
            earliestOpen = openTime.getTime();
          }
        }
      });
    } catch(e) {
      // Silently catch missing sheet errors
    }

    // 4. Calculate KPIs & Format for UI
    var mtdIn = 0, mtdOut = 0;
    var displayRows = [];

    for (var k = 0; k < masterList.length; k++) {
      var row = masterList[k];
      
      if (row.period === thisPeriod) {
        mtdIn += row.amtIn;
        mtdOut += row.amtOut;
      }
      
      // OPTION A MATH: Add cash transactions that happened AFTER the shift opened
      if (hasOpenShift && row.mode.toLowerCase() === 'cash' && row.ts.getTime() >= earliestOpen) {
        openDrawerCash += (row.amtIn - row.amtOut); 
      }

      if (displayRows.length < 100) {
        displayRows.push({
          txnId: row.txnId,
          ts: Utilities.formatDate(row.ts, ACC_CFG.TZ, "yyyy-MM-dd HH:mm"),
          category: row.category,
          entity: row.entity,
          mode: row.mode,
          amtIn: row.amtIn,
          amtOut: row.amtOut,
          locked: row.locked
        });
      }
    }

    // ---- Red Flags ----
    var flags = { cashMismatch: 0, pendingWriteOffs: 0, delayedClaims: 0, pendingDischarges: 0 };

    try {
      var shData = acc_sheet_(ACC_CFG.SHIFTS).getDataRange().getValues();
      for (var s2 = 1; s2 < shData.length; s2++) {
        if (!shData[s2][0]) continue;
        if (Math.abs(acc_money_(shData[s2][10])) > ACC_CFG.VARIANCE_FLAG &&
            acc_str_(shData[s2][11]).toLowerCase() !== 'reconciled') flags.cashMismatch++;
      }
    } catch (e) {}

    var openCount = virtualIncome.filter(function(r) { return r.open && r.balance > 0; }).length;
    flags.pendingDischarges = openCount; 

    try {
      var cData = acc_sheet_(ACC_CFG.CLAIMS).getDataRange().getValues();
      var now = new Date();
      for (var c = 1; c < cData.length; c++) {
        if (!cData[c][0]) continue;
        var settle = acc_str_(cData[c][12]).toLowerCase();
        if (settle === 'settled' || settle === 'rejected') continue;
        var cd = acc_toDate_(cData[c][1]);
        if (!isNaN(cd.getTime()) && Math.floor((now - cd) / 86400000) > ACC_CFG.TPA_DELAY_DAYS) {
          flags.delayedClaims++;
        }
      }
    } catch (e) {}

    return {
      success: true,
      kpis: { collectedMTD: acc_money_(mtdIn), paidMTD: acc_money_(mtdOut), drawerCash: acc_money_(openDrawerCash) },
      flags: flags,
      currentPeriod: thisPeriod,
      periodLocked: acc_isLocked_(thisPeriod),
      lockedPeriods: acc_lockedSet_(),
      ledger: displayRows
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// =========================================================================
// WRITE: Record a realized cash-flow entry (Expenses & Manual Incomes)
// =========================================================================
function recordLedgerEntry(obj) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    if (!obj) return { success: false, message: "No data received." };
    var dir = acc_str_(obj.direction).toUpperCase();
    if (dir !== 'IN' && dir !== 'OUT') return { success: false, message: "direction must be IN or OUT." };
    var amount = acc_money_(obj.amount);
    if (amount <= 0) return { success: false, message: "Amount must be greater than 0." };

    var period = acc_period_(new Date());
    if (acc_isLocked_(period)) return { success: false, message: "Period " + period + " is locked. Entries are frozen." };

    var sh = acc_sheet_(ACC_CFG.LEDGER);
    var txnId = "TXN-" + Date.now().toString().slice(-9);
    var ts = new Date(); 

    var amountIn = (dir === 'IN') ? amount : 0;
    var amountOut = (dir === 'OUT') ? amount : 0;

    sh.appendRow([
      acc_str_(txnId),                                   // A Txn_ID
      ts,                                                // B Timestamp (Date)
      acc_str_(dir === 'IN' ? 'Receipt' : 'Payment'),    // C Voucher_Type
      acc_str_(obj.category) || 'General',               // D Category
      acc_str_(obj.department),                          // E Department
      acc_str_(obj.refId),                               // F Reference_ID
      acc_str_(obj.entity),                              // G Entity_Name
      acc_str_(obj.mode) || 'Cash',                      // H Payment_Mode
      amountIn,                                          // I Amount_In
      amountOut,                                         // J Amount_Out
      acc_str_(obj.loggedBy) || 'UNKNOWN',               // K Logged_By
      acc_str_(obj.attachmentUrl),                       // L Attachment_URL
      "FALSE",                                           // M Is_Locked
      acc_str_(obj.notes)                                // N Notes
    ]);

    // Optional GST split -> Tax_Ledger
    if (obj.gst && (acc_money_(obj.gst.base) > 0)) {
      var g = obj.gst;
      var cgst = acc_money_(g.cgst), sgst = acc_money_(g.sgst), igst = acc_money_(g.igst);
      acc_sheet_(ACC_CFG.TAX).appendRow([
        "TAX-" + Date.now().toString().slice(-9),        
        ts,                                              
        acc_str_(txnId),                                 
        acc_str_(dir === 'IN' ? 'Output' : 'Input'),     
        acc_str_(g.category),                            
        acc_money_(g.base),                              
        cgst, sgst, igst,                                
        acc_money_(cgst + sgst + igst),                  
        "Unfiled"                                        
      ]);
    }

    acc_audit_(obj.loggedBy, dir === 'IN' ? 'RECORD_INCOME' : 'RECORD_EXPENSE',
               'Master_Ledger', txnId, '', amount, acc_str_(obj.notes));

    SpreadsheetApp.flush();
    return { success: true, message: "Entry recorded.", txnId: txnId };
  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// WRITE: Lock a financial period (YYYY-MM). Stamps Is_Locked on its rows.
// =========================================================================
function lockFinancialPeriod(period, loggedBy) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    period = acc_str_(period) || acc_period_(new Date());
    if (!/^\d{4}-\d{2}$/.test(period)) return { success: false, message: "Invalid period format (YYYY-MM)." };

    var set = acc_lockedSet_();
    if (set.indexOf(period) !== -1) return { success: false, message: "Period " + period + " is already locked." };

    var sh = acc_sheet_(ACC_CFG.LEDGER);
    var data = sh.getDataRange().getValues();
    var stamped = 0;
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (acc_period_(data[i][1]) === period && acc_str_(data[i][12]).toUpperCase() !== 'TRUE') {
        sh.getRange(i + 1, 13).setValue("TRUE"); // col M
        stamped++;
      }
    }

    set.push(period);
    PropertiesService.getScriptProperties().setProperty(ACC_CFG.LOCK_PROP, set.join(","));
    acc_audit_(loggedBy, 'LOCK_PERIOD', 'Accounts', period, 'OPEN', 'LOCKED', stamped + ' rows frozen');

    SpreadsheetApp.flush();
    return { success: true, message: "Period " + period + " locked. " + stamped + " entries frozen." };
  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function getLockedPeriods() {
  try { return { success: true, periods: acc_lockedSet_() }; }
  catch (e) { return { success: false, message: e.message }; }
}