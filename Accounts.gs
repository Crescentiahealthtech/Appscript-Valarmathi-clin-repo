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
/**
 * The accounting period ("yyyy-MM") a value falls in, or '' when it holds no
 * readable date.
 *
 * The old body fell back to `new Date()` for anything unparseable, which put
 * a row with a broken timestamp into THE CURRENT MONTH's figures. Worse, it
 * tested with `isNaN(new Date(d).getTime())`, and `new Date(null)` is not
 * NaN - it is 1 January 1970 - so a null date came back as the period
 * "1970-01" with no warning at all.
 *
 * An empty period is the honest answer: acc_isLocked_('') is false, and
 * row.period === thisPeriod is false, so such a row is counted nowhere
 * rather than counted in the wrong place.
 */
function acc_period_(d) {
  var dt = (typeof cresc_parseDate_ === 'function') ? cresc_parseDate_(d)
         : ((d instanceof Date && !isNaN(d.getTime())) ? d : null);
  if (!dt) return '';
  return Utilities.formatDate(dt, ACC_CFG.TZ, "yyyy-MM");
}

/**
 * A bill's date, and a day on its own.
 *
 * Both are called by the receivables screen and neither existed, so every
 * call to getReceivablesAgeing() died on "acc_fmtTs_ is not defined" and the
 * outer catch reported that as the reason there were no receivables. A bad
 * or missing date returns "" rather than "Invalid Date": an unknown bill date
 * should read as unknown, not as a value someone might chase.
 */
function acc_fmtTs_(d) {
  if (!d) return "";
  var dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return "";
  return Utilities.formatDate(dt, ACC_CFG.TZ, "dd-MMM-yyyy hh:mm a");
}

function acc_dayStr_(d) {
  if (!d) return "";
  var dt = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dt.getTime())) return "";
  return Utilities.formatDate(dt, ACC_CFG.TZ, "dd-MMM-yyyy");
}

/**
 * "yyyy-MM-dd" for a date, or "" when there is no readable date.
 *
 * THIS IS WHY THE FINANCE HUB WOULD NOT OPEN.
 *
 * getAccountsDashboard() stamped every row with
 *
 *     isToday: Utilities.formatDate(ts, ACC_CFG.TZ, "yyyy-MM-dd") === todayStr
 *
 * and `ts` is acc_toDate_(...), which returns NULL for a cell whose date
 * cannot be read - a blank Timestamp on a bill, or a day-first string the
 * parser rejects. Utilities.formatDate(null, ...) does not return "": it
 * throws "The parameters (null,String,String) don't match the method
 * signature".
 *
 * In the ledger loop that throw was swallowed by a bare `catch (e) {}`, so
 * every manual expense simply vanished from the hub. In the clinical-income
 * loop it was not caught at all, so ONE paid bill with an unreadable date
 * took down getAccountsDashboard() entirely - it returned
 * { success:false } and the Finance Hub, the Master Ledger, the Cash Drawer
 * and every spoke that loads behind it showed nothing, for every role
 * including admin.
 *
 * A row with no readable date is not today. Saying so costs nothing and
 * cannot throw.
 */
function acc_ymd_(d) {
  if (!d) return "";
  var dt = (d instanceof Date) ? d : acc_toDate_(d);
  if (!dt || isNaN(dt.getTime())) return "";
  return Utilities.formatDate(dt, ACC_CFG.TZ, "yyyy-MM-dd");
}
/**
 * A date from a ledger, bill or shift row - or null.
 *
 * Two bugs lived in the four lines this replaces. It parsed with
 * `new Date(v)`, so a day-first "13/09/2026" in a text-formatted column was
 * Invalid Date and an ambiguous "01/02/2026" silently became 2 January. And
 * it then returned `new Date()` for anything it could not read, so a row
 * with a broken timestamp was BOOKED TO TODAY: it appeared in today's
 * collection figure, this month's period, and the open shift's cash
 * reconciliation, with nothing on screen to say the date was never readable.
 *
 * It now returns null, which is what every call site already tests for.
 * Shared_Dates.gs does the parsing.
 */
function acc_toDate_(v) {
  if (typeof cresc_parseDate_ === 'function') return cresc_parseDate_(v);
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
/**
 * A unique id for a financial record.
 *
 * Every id in this module was built as
 *
 *     PREFIX + Date.now().toString().slice(-9)
 *
 * which is not unique in either of the two ways that matter for a ledger.
 * Two entries recorded in the SAME MILLISECOND get the same id - and they
 * do: settling a discharge posts a receipt and its tax split together, and
 * a busy counter has two people pressing Save at once. And slice(-9) keeps
 * only the last nine digits of a thirteen-digit clock, which wraps roughly
 * every eleven and a half days, so two entries that far apart collide too.
 *
 * A duplicate id in a ledger is not a cosmetic problem. Reversing,
 * reconciling and auditing all find a transaction by its id, and a lookup
 * that matches two rows either picks one arbitrarily or reverses the wrong
 * entry.
 *
 * The replacement is sortable, readable and unique: a timestamp anyone can
 * read at a glance, plus six characters from a UUID.
 *
 *     TXN-260913-174233-A19F4C
 *
 * Nothing parses these ids - they are matched whole - so the change is safe
 * for rows already written, which keep the ids they have.
 *
 * @param {string} prefix  e.g. 'TXN', 'TAX', 'AUD'
 * @return {string}
 */
function acc_newId_(prefix) {
  var stamp = Utilities.formatDate(new Date(), ACC_CFG.TZ, 'yyMMdd-HHmmss');
  var rand;
  try {
    rand = Utilities.getUuid().replace(/-/g, '').substring(0, 6).toUpperCase();
  } catch (e) {
    // getUuid() is not available in every execution context. Six random
    // base-36 characters are still far better than none.
    rand = ('000000' + Math.floor(Math.random() * 2176782336).toString(36).toUpperCase()).slice(-6);
  }
  return String(prefix || 'ID') + '-' + stamp + '-' + rand;
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
    var id = acc_newId_("AUD");
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
//
// An appointment's Fee column is now the consultation TARIFF, not a receipt:
// the booking modal no longer asks for a figure, and the money is taken on the
// Hospital Billing desk, which writes a real invoice (Hospital_Billing.gs).
// Counting both would bill the hospital twice for one consultation.
//
// So a completed appointment contributes its fee ONLY while no hospital
// invoice names it. That keeps every historic visit — which has a fee and no
// invoice, because there was nowhere to raise one — counted exactly as it was,
// while every visit billed the new way is counted once, from its invoice.
function acc_opRows_() {
  var billed = {};
  try { if (typeof hb_billedApptIds_ === 'function') billed = hb_billedApptIds_() || {}; }
  catch (e) { billed = {}; }

  return acc_readObjects_('Appointments').map(function (r) {
    var status = acc_str_(r['Status']).toUpperCase();
    var fee = acc_money_(r['Fee']);
    var d = acc_toDate_(r['Timestamp'] || r['Date']);
    var billId = acc_str_(r['Appt_ID'] || r['Appt ID'] || r['ApptId']);
    var invoiced = !!billed[billId.toUpperCase()];
    return {
      source: 'OP_Consultation', 
      billId: billId,
      patientId: acc_str_(r['Patient_ID'] || r['Patient ID']), 
      name: acc_str_(r['Patient_Name'] || r['Patient Name']),
      admissionId: '', 
      net: invoiced ? 0 : fee,
      balance: 0, 
      mode: 'Cash', // Default OP collections to Cash
      billDate: d, realizedDate: d,
      realized: (!invoiced && status === 'COMPLETED' && fee > 0),
      open: false,
      _row: r._row
    };
  });
}

/**
 * 4. Hospital Billing Income Normalizer.
 *
 * The reader itself lives beside its schema in Hospital_Billing.gs. This
 * wrapper exists so a deployment that has not copied that file across still
 * loads the Finance Hub — with no hospital income in it — rather than failing
 * on "acc_hospitalRows_ is not defined".
 */
function acc_hospitalRowsSafe_() {
  try {
    if (typeof acc_hospitalRows_ !== 'function') return [];
    return acc_hospitalRows_() || [];
  } catch (e) { return []; }
}

// =========================================================================
// READ: Dashboard Payload (Virtual Merge of Ledger + Clinical Sources)
// =========================================================================
function getAccountsDashboard() {
  try {
    var thisPeriod = acc_period_(new Date());
    var todayStr = Utilities.formatDate(new Date(), ACC_CFG.TZ, "yyyy-MM-dd");
    var masterList = [];

    // Anything a row could not be read for. Returned to the client so a
    // broken cell is visible on the hub instead of quietly shrinking the
    // month's figures - the failure mode this whole function had.
    var readWarnings = [];

    // 1. Pull Expenses & Manual entries from Finance_Master_Ledger
    //
    // ONE BAD ROW USED TO COST EVERY ROW AFTER IT. The loop sat inside a
    // single `try { ... } catch (e) {}`, so the first cell that threw
    // abandoned the rest of the sheet and reported nothing at all - the hub
    // showed fewer expenses than the ledger held, with no way to tell.
    // The sheet read stays outside (a missing sheet really is fatal to this
    // block); each ROW is now its own failure boundary.
    try {
      var lData = acc_sheet_(ACC_CFG.LEDGER).getDataRange().getValues();
      for (var i = 1; i < lData.length; i++) {
        var r = lData[i];
        if (!r[0] || r[0] === "Txn_ID") continue;
        try {
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
            isToday: (acc_ymd_(ts) === todayStr)
          });
        } catch (rowErr) {
          readWarnings.push('Ledger row ' + (i + 1) + ' (' + acc_str_(r[0]) +
                            ') could not be read: ' + rowErr.message);
        }
      }
    } catch (e) {
      readWarnings.push('The expense ledger could not be read: ' + e.message);
    }

    // 2. Pull Clinical Income (Virtual Merge)
    //    Hospital_Invoices joins pharmacy, lab and OP as a fourth source: OP
    //    consultations, procedures and package bills are raised there now, and
    //    the ledger would not otherwise see any of them.
    var virtualIncome = acc_pharmaRows_()
      .concat(acc_labRows_())
      .concat(acc_opRows_())
      .concat(acc_hospitalRowsSafe_());
    virtualIncome.forEach(function(inc) {
      if (!inc.realized) return;
      try {
        var ts = inc.realizedDate;
        masterList.push({
          txnId: inc.billId,
          ts: ts,
          period: acc_period_(ts),
          category: inc.source,
          entity: inc.name || inc.patientId || "Walk-In",
          mode: inc.mode,
          // A part-paid bill realises only what was taken. Sources that settle
          // in full (pharmacy, lab) leave realizedAmount unset and keep using
          // the net, exactly as before.
          amtIn: (inc.realizedAmount === undefined) ? inc.net : inc.realizedAmount,
          amtOut: 0,
          locked: acc_isLocked_(acc_period_(ts)),
          isToday: (acc_ymd_(ts) === todayStr)
        });
      } catch (incErr) {
        readWarnings.push(acc_str_(inc.source) + ' bill ' + acc_str_(inc.billId) +
                          ' could not be read: ' + incErr.message);
      }
    });

    // 3. Sort chronologically (Newest first)
    // Newest first. A row whose timestamp could not be read sorts to the
    // bottom rather than throwing: ts is null for those now.
    masterList.sort(function (a, b) {
      return (b.ts ? b.ts.getTime() : 0) - (a.ts ? a.ts.getTime() : 0);
    });

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
          
          // acc_toDate_ returns null for an unreadable timestamp now, so an
          // open shift whose row lost its date no longer crashes the flag
          // sweep - it simply does not move the "earliest open" marker.
          var openTime = acc_toDate_(s['Timestamp']);
          if (openTime && openTime.getTime() < earliestOpen) {
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
      
      // OPTION A MATH: Add cash transactions that happened AFTER the shift opened.
      // row.ts is null for a row whose timestamp could not be read (acc_toDate_
      // returns null now rather than substituting today), so it is tested
      // before it is dereferenced. Such a row is left out of the drawer
      // reconciliation - counting it at an invented time is what produced the
      // mismatches the Cash Drawer tile reports.
      if (hasOpenShift && row.mode.toLowerCase() === 'cash' &&
          row.ts && row.ts.getTime() >= earliestOpen) {
        openDrawerCash += (row.amtIn - row.amtOut);
      }

      if (displayRows.length < 100) {
        displayRows.push({
          txnId: row.txnId,
          // formatDate(null, ...) throws. A row whose timestamp is unreadable
          // says so in the ledger rather than taking the whole screen down
          // with it, and says it where an accountant will see and fix it.
          ts: row.ts ? Utilities.formatDate(row.ts, ACC_CFG.TZ, "yyyy-MM-dd HH:mm")
                     : "(no date)",
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
        if (cd && Math.floor((now - cd) / 86400000) > ACC_CFG.TPA_DELAY_DAYS) {
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
      ledger: displayRows,
      // Capped: a sheet with a systematic date problem would otherwise send
      // one warning per row through google.script.run.
      warnings: readWarnings.slice(0, 10),
      warningCount: readWarnings.length
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// =========================================================================
// WRITE: Record a realized cash-flow entry (Expenses & Manual Incomes)
// =========================================================================
function recordLedgerEntry(obj, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    crescRequire_(sessionToken, 'accounts.write');
    if (!obj) return { success: false, message: "No data received." };
    var dir = acc_str_(obj.direction).toUpperCase();
    if (dir !== 'IN' && dir !== 'OUT') return { success: false, message: "direction must be IN or OUT." };
    var amount = acc_money_(obj.amount);
    if (amount <= 0) return { success: false, message: "Amount must be greater than 0." };

    var period = acc_period_(new Date());
    if (acc_isLocked_(period)) return { success: false, message: "Period " + period + " is locked. Entries are frozen." };

    var sh = acc_sheet_(ACC_CFG.LEDGER);
    var txnId = acc_newId_("TXN");
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
        acc_newId_("TAX"),
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
function lockFinancialPeriod(period, loggedBy, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'accounts.lock_period');
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