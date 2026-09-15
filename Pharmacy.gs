// =====================================================================
// CRESCENTIA HEALTHTECH — PHARMACY BACKEND  (Google Apps Script)
// =====================================================================
// SINGLE SOURCE OF TRUTH FOR STOCK: "Pharmacy_Inventory"
//   [0]Timestamp [1]Brand [2]Generic [3]Type [4]Qty [5]Unit [6]Batch
//   [7]Expiry [8]Rack [9]BuyPrice [10]MRP [11]GST% [12]Manufacturer [13]Supplier
//
// INVOICE STORE (relational, auto-created + auto-migrated):
//   "Pharmacy_Invoices"        -> one row per bill (header)
//   "Pharmacy_Invoice_Items"   -> one row per dispensed batch line
//
// PRESCRIPTION SOURCES (read-only; LATEST encounter/note only):
//   OPD ............. "Pharmacy_Queue_DB"  (status Pending, latest encounter)
//   IP Admission .... "IP_CaseSheets_DB"   (Prescription_JSON col 29, latest casesheet)
//   IP Ward Notes ... "IP_Pharmacy_Queue"  (Action_Flag ACTIVE, latest note/day)
// =====================================================================

var PH_SHEETS = {
  INVENTORY:    "Pharmacy_Inventory",
  INVOICES:     "Pharmacy_Invoices",
  INVOICE_ITEMS:"Pharmacy_Invoice_Items",
  PATIENTS:     "Patients",
  OPD_QUEUE:    "Pharmacy_Queue_DB",
  IP_CASESHEET: "IP_CaseSheets_DB",
  IP_QUEUE:     "IP_Pharmacy_Queue",
  DS_SUMMARIES: "DS_Summaries",
  DS_SNAPSHOTS: "DS_Snapshots",
  MASTER_LEDGER:"Pharmacy_Master",
  DISPOSALS:    "Pharmacy_Disposal_Log"
};

/**
 * The disposal register. One row per write-off, never edited, never deleted.
 *
 * Separate from Pharmacy_Master on purpose. The master ledger records stock
 * MOVEMENT — received, dispensed, adjusted. A write-off is not a movement, it
 * is a LOSS, and the questions asked of it are different ones: how much did we
 * throw away last quarter, which supplier's batches keep expiring, what did it
 * cost. Those are answered by a register you can total, not by filtering a
 * movement log for a category.
 *
 * Buy_Value is the number that matters for the loss and MRP_Value the number
 * that matters for the insurer, so both are stored at the moment of disposal
 * rather than recomputed later from a price that will have changed.
 */
var PH_DISPOSAL_HEADERS = [
  "Disposal_ID", "Timestamp", "Brand", "Generic", "Batch", "Expiry",
  "Qty_Discarded", "Unit", "Reason", "Note",
  "Buy_Value", "MRP_Value", "Stock_Before", "Stock_After",
  "Rack", "Supplier", "Manufacturer", "Discarded_By", "Witness"
];

/**
 * Why stock left the shelf without being sold.
 *
 * A free-text reason cannot be counted, and "what are we losing money to" is
 * the only question this register exists to answer. EXPIRED and DAMAGED are
 * ordinary wastage; RECALL is the supplier's problem and may be recoverable;
 * THEFT_LOSS is an incident, not wastage, and is deliberately separate so it
 * never hides inside a wastage total.
 */
var PH_DISPOSAL_REASONS = {
  EXPIRED:      "Expired",
  DAMAGED:      "Damaged in storage",
  BREAKAGE:     "Breakage / spillage",
  CONTAMINATED: "Contaminated or cold-chain break",
  RECALL:       "Manufacturer recall",
  RETURN_SUPP:  "Returned to supplier",
  THEFT_LOSS:   "Missing / unaccounted",
  OTHER:        "Other (see note)"
};

var PH_INVOICE_HEADERS = ["Invoice_No","Timestamp","Bill_Type","Patient_ID","Patient_Name",
  "Mobile","Age","Sex","Address","Doctor","Gross","Total_GST","Discount","Net",
  "Pay_Mode","Txn_ID","Pay_Status","Status","Item_Count","Created_By","Bill_UUID",
  "Settled_At","Settled_By"];


// =====================================================================
// SECTION A — INVENTORY (used by Ledger + Add screens)
// =====================================================================

function fetchPharmacyInventory(sessionToken) {
  try {
    crescRequire_(sessionToken, 'pharmacy.read');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.INVENTORY);
    if (!sheet) return { success: false, message: "Pharmacy_Inventory sheet not found." };
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };

    var out = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1]) continue;
      out.push({
        rowId: i + 1,
        brandName: String(row[1]),
        genericName: String(row[2] || ""),
        drugType: String(row[3] || ""),
        stock: parseInt(row[4], 10) || 0,
        unit: String(row[5] || ""),
        batch: String(row[6] || ""),
        expiry: row[7] instanceof Date
          ? Utilities.formatDate(row[7], Session.getScriptTimeZone(), "yyyy-MM") : String(row[7] || ""),
        rack: String(row[8] || ""),
        buyPrice: parseFloat(row[9]) || 0,
        mrp: parseFloat(row[10]) || 0,
        gst: parseFloat(row[11]) || 0,
        manufacturer: String(row[12] || ""),
        supplier: String(row[13] || "")
      });
    }
    return { success: true, data: out };
  } catch (error) {
    return { success: false, message: "Database Error: " + error.toString() };
  }
}

function savePharmacyInventory(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    crescRequire_((payload || {}).token, 'pharmacy.stock_add');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.INVENTORY);
    if (!sheet) {
      sheet = ss.insertSheet(PH_SHEETS.INVENTORY);
      sheet.appendRow(["Timestamp","Brand Name","Generic Name","Type","Qty","Unit",
        "Batch No","Expiry Date","Rack Location","Buy Price","MRP","GST %","Manufacturer","Supplier"]);
      sheet.getRange("A1:N1").setFontWeight("bold").setBackground("#d9d9d9");
    }
    sheet.appendRow([
      payload.timestamp ? new Date(payload.timestamp) : new Date(),
      String(payload.medicineName || ""), String(payload.genericName || ""),
      String(payload.drugType || ""), parseInt(payload.qty, 10) || 0, String(payload.unit || ""),
      String(payload.batchNo || ""), String(payload.expiryDate || ""), String(payload.rackLocation || ""),
      parseFloat(payload.buyPrice) || 0, parseFloat(payload.mrp) || 0, parseFloat(payload.gst) || 0,
      String(payload.manufacturer || ""), String(payload.supplier || "")
    ]);
    return { success: true, message: "Stock successfully added!" };
  } catch (error) {
    return { success: false, message: "Database Error: " + error.toString() };
  } finally { lock.releaseLock(); }
}

function updatePharmacyStock(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var actor = crescRequire_((payload || {}).token, 'pharmacy.stock_edit');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stockSheet = ss.getSheetByName(PH_SHEETS.INVENTORY);
    var masterSheet = ss.getSheetByName(PH_SHEETS.MASTER_LEDGER);
    var rowNum = parseInt(payload.rowId, 10);
    stockSheet.getRange(rowNum, 5).setValue(parseInt(payload.stock, 10) || 0);
    stockSheet.getRange(rowNum, 8).setValue(String(payload.expiry || ""));
    stockSheet.getRange(rowNum, 9).setValue(String(payload.rack || ""));
    stockSheet.getRange(rowNum, 11).setValue(parseFloat(payload.mrp) || 0);
    stockSheet.getRange(rowNum, 12).setValue(parseFloat(payload.gst) || 0);
    if (masterSheet) {
      masterSheet.appendRow([new Date(), "ADJ-" + Math.floor(1000 + Math.random() * 9000),
        "Manual Adjustment", String(payload.brandName || ""), String(payload.genericName || ""),
        "N/A", "N/A", parseInt(payload.stock, 10) || 0, String(payload.batch || ""),
        String(payload.expiry || ""), String(payload.rack || ""), 0, parseFloat(payload.mrp) || 0,
        parseFloat(payload.gst) || 0, String(payload.manufacturer || ""), String(payload.supplier || ""),
        // The actor column used to be the literal string "Admin", whoever was
        // signed in — so the master ledger recorded every manual adjustment
        // in this clinic's history as having been made by the same person.
        actor.displayName || actor.username]);
    }
    logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
              'PHARMACY_STOCK_ADJUST', 'Pharmacy_Inventory', String(payload.batch || ''),
              { brand: String(payload.brandName || ''), stock: parseInt(payload.stock, 10) || 0 });
    return { success: true, message: "Inventory updated securely." };
  } catch (error) {
    return { success: false, message: "Failed to update: " + error.toString() };
  } finally { lock.releaseLock(); }
}

// =====================================================================
// SECTION A2 — DISPOSAL (write-off)
// ---------------------------------------------------------------------
// WHY THIS EXISTS
//
// Live Inventory had exactly one row action: Edit. The dashboard counted
// expired batches on its own KPI card — and then offered no way to do
// anything about them. The only way to take expired stock off the shelf was
// to open the edit dialog and type 0 into "Live Stock Count".
//
// That loses everything worth keeping:
//
//   - HOW MANY went. The count is overwritten, not decremented, so the
//     quantity destroyed is gone the moment it is saved.
//   - WHY. Expired, broken, recalled and stolen all look identical
//     afterwards: a batch that used to have stock and now has none.
//   - WHAT IT COST. Nothing captures buy value, so "what did we lose to
//     expiry this quarter" cannot be answered at all.
//   - WHO. updatePharmacyStock writes the literal string "Admin" into the
//     master ledger's actor column, whoever was signed in.
//
// And the drug is still destroyed either way — this is a controlled
// substance register in every respect except the record. Schedule H and
// H1 stock in particular has to be accounted for on disposal.
//
// So: a first-class action, a reason code that can be totalled, the value at
// the moment it went, and a named actor taken from the session rather than
// from the client.
// =====================================================================

/** Creates the disposal register on first use. */
function _phDisposalSheet_(ss) {
  var sheet = ss.getSheetByName(PH_SHEETS.DISPOSALS);
  if (!sheet) {
    sheet = ss.insertSheet(PH_SHEETS.DISPOSALS);
    sheet.appendRow(PH_DISPOSAL_HEADERS);
    sheet.getRange(1, 1, 1, PH_DISPOSAL_HEADERS.length)
         .setFontWeight("bold").setBackground("#f4cccc");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** FRONTEND ENTRY. The reason codes, so the dialog and the server agree. */
function getDisposalReasons(sessionToken) {
  crescRequire_(sessionToken, 'pharmacy.read');
  return Object.keys(PH_DISPOSAL_REASONS).map(function (k) {
    return { code: k, label: PH_DISPOSAL_REASONS[k] };
  });
}

/**
 * FRONTEND ENTRY. Writes stock off.
 *
 * @param {{token:string, rowId:number, batch:string, brandName:string,
 *          qty:number, reason:string, note:string, witness:string}} payload
 * @return {{success:boolean, message:string, disposalId?:string, remaining?:number}}
 */
function discardPharmacyStock(payload) {
  var lock = LockService.getScriptLock();
  try {
    var p = payload || {};
    var actor = crescRequire_(p.token, 'pharmacy.stock_discard');

    var reason = String(p.reason || '').trim().toUpperCase();
    if (!PH_DISPOSAL_REASONS[reason]) {
      return { success: false, message: "Choose why this stock is being written off." };
    }
    var note = String(p.note || '').trim();
    // OTHER with no explanation is the same as no reason at all, and
    // THEFT_LOSS is an incident that has to say what happened.
    if ((reason === 'OTHER' || reason === 'THEFT_LOSS') && note.length < 5) {
      return { success: false,
               message: reason === 'OTHER'
                 ? "Say what the reason is — \"Other\" on its own cannot be accounted for."
                 : "Missing stock needs a note saying what is known about it." };
    }

    var qty = parseInt(p.qty, 10);
    if (!(qty > 0)) {
      return { success: false, message: "Enter how many units are being written off." };
    }

    lock.waitLock(15000);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.INVENTORY);
    if (!sheet) return { success: false, message: "Pharmacy_Inventory sheet not found." };

    var rowNum = parseInt(p.rowId, 10);
    if (!(rowNum > 1) || rowNum > sheet.getLastRow()) {
      return { success: false, message: "That batch is no longer in the inventory. Refresh and try again." };
    }

    // Re-read the row under the lock and CHECK IT IS STILL THE SAME BATCH.
    //
    // rowId is a position, not an identity. If anyone inserted or deleted a
    // row in the sheet between the screen loading and this call, position 47
    // is a different drug — and unlike an edit, which a person notices and
    // corrects, a write-off destroys the count silently. updatePharmacyStock
    // trusts the index blind; this one will not.
    var row = sheet.getRange(rowNum, 1, 1, 14).getValues()[0];
    var brand = String(row[1] || '');
    var batch = String(row[6] || '');
    var wantBrand = String(p.brandName || '').trim();
    var wantBatch = String(p.batch || '').trim();

    if ((wantBrand && wantBrand.toUpperCase() !== brand.trim().toUpperCase()) ||
        (wantBatch && wantBatch.toUpperCase() !== batch.trim().toUpperCase())) {
      return { success: false,
               message: "The inventory has changed since this screen was loaded — row " + rowNum +
                        " now holds " + (brand || "(blank)") + " batch " + (batch || "(blank)") +
                        ". Refresh and try again. Nothing was written off." };
    }

    var before = parseInt(row[4], 10) || 0;
    if (qty > before) {
      return { success: false,
               message: "Only " + before + " unit(s) of batch " + (batch || "—") +
                        " are in stock. You cannot write off " + qty + "." };
    }
    var after = before - qty;

    var buy = parseFloat(row[9]) || 0;
    var mrp = parseFloat(row[10]) || 0;
    var expiry = row[7] instanceof Date
      ? Utilities.formatDate(row[7], Session.getScriptTimeZone(), "yyyy-MM")
      : String(row[7] || "");

    var disposalId = 'DSP-' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' +
      Utilities.getUuid().substring(0, 6).toUpperCase();
    var who = actor.displayName || actor.username;

    // The register FIRST, then the stock.
    //
    // If the second write fails, the worst case is a disposal row with no
    // matching decrement — visible, reconcilable, and obvious on the next
    // stock check. The other order's worst case is stock silently destroyed
    // with no record of where it went, which is the failure this whole
    // function exists to prevent.
    _phDisposalSheet_(ss).appendRow([
      disposalId, new Date(), brand, String(row[2] || ''), batch, expiry,
      qty, String(row[5] || ''), reason, note,
      +(buy * qty).toFixed(2), +(mrp * qty).toFixed(2), before, after,
      String(row[8] || ''), String(row[13] || ''), String(row[12] || ''),
      who, String(p.witness || '').trim()
    ]);

    sheet.getRange(rowNum, 5).setValue(after);

    var master = ss.getSheetByName(PH_SHEETS.MASTER_LEDGER);
    if (master) {
      master.appendRow([new Date(), disposalId, "Disposal — " + PH_DISPOSAL_REASONS[reason],
        brand, String(row[2] || ''), "N/A", "N/A", -qty, batch, expiry,
        String(row[8] || ''), buy, mrp, parseFloat(row[11]) || 0,
        String(row[12] || ''), String(row[13] || ''), who]);
    }

    logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
              'PHARMACY_STOCK_DISCARD', 'Pharmacy_Inventory', disposalId,
              { brand: brand, batch: batch, qty: qty, reason: reason,
                note: note, buyValue: +(buy * qty).toFixed(2),
                stockBefore: before, stockAfter: after });

    return {
      success: true,
      disposalId: disposalId,
      remaining: after,
      message: qty + " unit(s) of " + brand + " (batch " + (batch || "—") + ") written off as " +
               PH_DISPOSAL_REASONS[reason].toLowerCase() + ". " +
               (after > 0 ? after + " left in this batch." : "This batch is now empty.") +
               " Recorded as " + disposalId + "."
    };
  } catch (error) {
    return { success: false, message: "Could not write off: " + error.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * FRONTEND ENTRY. The disposal register, newest first.
 *
 * The point of recording a reason code is being able to total it, so this
 * returns the breakdown alongside the rows — what expiry cost this period
 * versus what damage cost, which is the conversation a write-off register is
 * supposed to start.
 *
 * @param {{token:string, days?:number, limit?:number}} opts
 */
function getDisposalLog(opts) {
  try {
    var o = opts || {};
    crescRequire_(o.token, 'pharmacy.read');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.DISPOSALS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, rows: [], totals: {}, totalValue: 0, message: '' };
    }

    var days  = Math.min(Math.max(parseInt(o.days, 10) || 90, 1), 1095);
    var limit = Math.min(Math.max(parseInt(o.limit, 10) || 300, 1), 2000);
    var since = new Date(Date.now() - days * 86400000);

    var last = sheet.getLastRow();
    var cols = PH_DISPOSAL_HEADERS.length;
    // Read only the tail. The register is append-only and never pruned, so
    // reading it whole is a timeout waiting for a busy year.
    var from = Math.max(2, last - limit + 1);
    var data = sheet.getRange(from, 1, last - from + 1, cols).getValues();

    var rows = [], totals = {}, totalValue = 0;
    for (var i = data.length - 1; i >= 0; i--) {
      var r = data[i];
      var ts = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if (!(ts instanceof Date) || isNaN(ts.getTime()) || ts < since) continue;

      var reason = String(r[8] || 'OTHER');
      var value  = parseFloat(r[10]) || 0;
      totals[reason] = (totals[reason] || 0) + value;
      totalValue += value;

      rows.push({
        disposalId: String(r[0] || ''),
        at:      Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm'),
        brand:   String(r[2] || ''), generic: String(r[3] || ''),
        batch:   String(r[4] || ''), expiry:  String(r[5] || ''),
        qty:     parseInt(r[6], 10) || 0, unit: String(r[7] || ''),
        reason:  reason, reasonLabel: PH_DISPOSAL_REASONS[reason] || reason,
        note:    String(r[9] || ''),
        buyValue: value, mrpValue: parseFloat(r[11]) || 0,
        by:      String(r[17] || ''), witness: String(r[18] || '')
      });
    }

    return { success: true, rows: rows, totals: totals,
             totalValue: +totalValue.toFixed(2), windowDays: days, message: '' };
  } catch (err) {
    return { success: false, rows: [], totals: {}, totalValue: 0, message: err.message };
  }
}

// =====================================================================
// SECTION B — BILLING DATA READS
// =====================================================================

function fetchBillableStock(sessionToken) {
  try {
    crescRequire_(sessionToken, 'pharmacy.read');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.INVENTORY);
    if (!sheet) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };

    var batches = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var brand = row[1] ? String(row[1]).trim() : "";
      var qty = parseInt(row[4], 10) || 0;
      if (!brand || qty <= 0) continue;
      var exp = row[7] instanceof Date
        ? Utilities.formatDate(row[7], Session.getScriptTimeZone(), "yyyy-MM") : String(row[7] || "");
      batches.push({
        rowId: i + 1, brand: brand, generic: String(row[2] || ""), type: String(row[3] || ""),
        qty: qty, unit: String(row[5] || ""), batch: String(row[6] || ""),
        expiry: exp, expSort: _expiryToSortKey_(exp),
        mrp: parseFloat(row[10]) || 0, gst: parseFloat(row[11]) || 0, rack: String(row[8] || "")
      });
    }
    batches.sort(function (a, b) {
      if (a.brand.toLowerCase() < b.brand.toLowerCase()) return -1;
      if (a.brand.toLowerCase() > b.brand.toLowerCase()) return 1;
      return a.expSort - b.expSort;
    });
    return { success: true, data: batches };
  } catch (error) {
    return { success: false, message: "Stock load failed: " + error.toString() };
  }
}

function getPatientBillingContext(query, sessionToken) {
  try {
    crescRequire_(sessionToken, ['billing.read', 'pharmacy.read']);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var patient = _findPatient_(ss, query);
    var pid = patient ? String(patient.id) : String(query || "").trim();
    var rx = []
      .concat(_readOpdPrescriptions_(ss, pid))
      .concat(_readIpCasesheetPrescriptions_(ss, pid))
      .concat(_readIpWardPrescriptions_(ss, pid))
      // THE DISCHARGE SCRIPT. This was the missing fourth source: the
      // counter could see what the patient was given on the ward and what
      // the OP consult wrote, but not what the discharge summary actually
      // sends them home on - which is the one prescription they are
      // standing at the counter to collect. It was being read off a printed
      // sheet and typed in again, which is how a dose gets transcribed
      // wrong on the last transaction of an admission.
      .concat(_readDischargePrescriptions_(ss, pid));
    return { success: true, patient: patient, prescriptions: _dedupePrescriptions_(rx) };
  } catch (error) {
    return { success: false, message: "Lookup failed: " + error.toString() };
  }
}

function _findPatient_(ss, query) {
  var sheet = ss.getSheetByName(PH_SHEETS.PATIENTS);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var q = String(query || "").trim().toUpperCase();
  if (!q) return null;
  for (var i = 1; i < data.length; i++) {
    var id = data[i][0] ? String(data[i][0]).trim().toUpperCase() : "";
    var mob = data[i][6] ? String(data[i][6]).trim() : "";
    var alt = data[i][7] ? String(data[i][7]).trim() : "";
    if (id === q || mob === q || alt === q) {
      return { 
        id: String(data[i][0]), 
        name: String(data[i][2] || ""), 
        age: data[i][3] || "",
        sex: String(data[i][4] || ""), 
        mobile: mob, 
        whatsapp: String(data[i][7] || mob || ""), // Added for Dispatch Hub
        email: String(data[i][16] || ""),          // Added for Dispatch Hub (Assumes Col 17)
        address: String(data[i][8] || "") 
      };
    }
  }
  return null;
}

// --- OPD: only the LATEST encounter's pending meds --------------------------
function _readOpdPrescriptions_(ss, pid) {
  var out = [];
  var sheet = ss.getSheetByName(PH_SHEETS.OPD_QUEUE);
  if (!sheet || !pid) return out;
  var data = sheet.getDataRange().getValues();
  var latestEnc = null, latestKey = "";
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toUpperCase() !== pid.toUpperCase()) continue;
    if (String(data[i][7]).trim().toUpperCase() !== "PENDING") continue;
    var key = _dateKey_(data[i][3]) + "|" + String(data[i][2]);
    if (key >= latestKey) { latestKey = key; latestEnc = String(data[i][2]); }
  }
  if (latestEnc === null) return out;
  for (var j = 1; j < data.length; j++) {
    if (String(data[j][1]).trim().toUpperCase() !== pid.toUpperCase()) continue;
    if (String(data[j][7]).trim().toUpperCase() !== "PENDING") continue;
    if (String(data[j][2]) !== latestEnc) continue;
    var name = String(data[j][4] || "").trim();
    if (!name) continue;
    out.push({ source: "OPD", drugName: name, dose: String(data[j][5] || ""),
      orderedBy: String(data[j][6] || "Doctor"), date: _fmtDate_(data[j][3]) });
  }
  return out;
}

// --- IP Admission: latest casesheet only ------------------------------------
function _readIpCasesheetPrescriptions_(ss, pid) {
  var out = [];
  var sheet = ss.getSheetByName(PH_SHEETS.IP_CASESHEET);
  if (!sheet || !pid) return out;
  var data = sheet.getDataRange().getValues();
  var latestRow = null, latestTime = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim().toUpperCase() !== pid.toUpperCase()) continue;
    var t = (data[i][5] instanceof Date) ? data[i][5].getTime() : 0;
    if (t >= latestTime) { latestTime = t; latestRow = data[i]; }
  }
  if (!latestRow) return out;
  try {
    var arr = JSON.parse(latestRow[29] || "[]");
    for (var j = 0; j < arr.length; j++) {
      var it = arr[j] || {};
      var nm = String(it.name || "").trim();
      if (!nm) continue;
      out.push({ source: "IP-ADMISSION", drugName: nm, dose: String(it.sig || ""),
        orderedBy: "Doctor", date: _fmtDate_(latestRow[5]),
        instructions: [it.days ? it.days + " day(s)" : "", it.comment || ""].filter(String).join(" • ") });
    }
  } catch (e) {}
  return out;
}

// --- IP Ward: latest note (or latest day) of ACTIVE meds --------------------
function _readIpWardPrescriptions_(ss, pid) {
  var out = [];
  var sheet = ss.getSheetByName(PH_SHEETS.IP_QUEUE);
  if (!sheet || !pid) return out;
  var data = sheet.getDataRange().getValues();
  var bestTime = -1, bestNote = "", bestDay = "";
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim().toUpperCase() !== pid.toUpperCase()) continue;
    if (String(data[i][11]).trim().toUpperCase() !== "ACTIVE") continue;
    var t = (data[i][10] instanceof Date) ? data[i][10].getTime() : _dateKeyNum_(data[i][10]);
    if (t >= bestTime) { bestTime = t; bestNote = String(data[i][14] || ""); bestDay = _dateKey_(data[i][10]); }
  }
  if (bestTime < 0) return out;
  for (var j = 1; j < data.length; j++) {
    if (String(data[j][2]).trim().toUpperCase() !== pid.toUpperCase()) continue;
    if (String(data[j][11]).trim().toUpperCase() !== "ACTIVE") continue;
    var match = bestNote ? (String(data[j][14] || "") === bestNote) : (_dateKey_(data[j][10]) === bestDay);
    if (!match) continue;
    var nm = String(data[j][3] || "").trim();
    if (!nm) continue;
    out.push({ source: "IP-WARD", drugName: nm,
      dose: [data[j][4], data[j][5], data[j][6]].filter(String).join(" "),
      orderedBy: "Doctor", date: _fmtDate_(data[j][10]), instructions: String(data[j][7] || "") });
  }
  return out;
}

/**
 * The take-home medicines from a SIGNED discharge summary.
 *
 * Reads through the discharge engine's own accessors (DS_Data.gs) rather
 * than parsing DS_Snapshots here, so the payload's chunking, its schema
 * version and its column migration stay in one place. Those functions live
 * in a different .gs file, and Apps Script leaves a function undefined
 * rather than failing to load when a file was not copied across - so a
 * project without the discharge module simply contributes nothing here,
 * exactly as it does today.
 *
 * ONLY SIGNED summaries are read. A draft is a document still being argued
 * over; dispensing against one would hand the patient a prescription the
 * consultant had not agreed to.
 *
 * @param {Spreadsheet} ss
 * @param {string} pid
 * @return {Array<{source, drugName, dose, orderedBy, date, instructions}>}
 */
function _readDischargePrescriptions_(ss, pid) {
  var out = [];
  if (!pid) return out;
  if (typeof dsx_latestSnapshotOfType_ !== "function" ||
      typeof dsx_summaryIdFor_ !== "function") return out;
  if (!ss.getSheetByName(PH_SHEETS.DS_SUMMARIES)) return out;

  try {
    var sh = ss.getSheetByName(PH_SHEETS.DS_SUMMARIES);
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return out;

    var hdr = data[0].map(function (h) { return String(h || "").trim(); });
    var cPid = hdr.indexOf("Patient_ID"), cId = hdr.indexOf("Summary_ID"),
        cStatus = hdr.indexOf("Status"), cSigned = hdr.indexOf("Signed_At");
    if (cPid === -1 || cId === -1 || cStatus === -1) return out;

    // The most recently signed summary for this patient. A patient with two
    // admissions has two, and the one they are collecting against is the
    // latest - the earlier one was dispensed at the time.
    var bestId = "", bestMs = -1, bestSignedAt = "";
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][cPid] || "").trim().toUpperCase() !== pid.toUpperCase()) continue;
      var st = String(data[i][cStatus] || "").trim().toUpperCase();
      if (st !== "SIGNED" && st !== "AMENDMENT_IN_PROGRESS") continue;
      var ms = (cSigned > -1) ? _dateKeyNum_(data[i][cSigned]) : 0;
      if (ms >= bestMs) {
        bestMs = ms;
        bestId = String(data[i][cId] || "").trim();
        bestSignedAt = (cSigned > -1) ? data[i][cSigned] : "";
      }
    }
    if (!bestId) return out;

    var snap = dsx_latestSnapshotOfType_(bestId, "SIGNED");
    if (!snap || !snap.payload || !snap.payload.sections) return out;

    var sec = snap.payload.sections.DISCHARGE_MEDICATIONS;
    if (!sec || !sec.content || !sec.content.rows || !sec.content.columns) return out;

    // BY COLUMN NAME. The columns changed once already (Generic/Brand/... ->
    // the OP/IP five), and a signed summary keeps whichever set it was
    // signed with, so both shapes are read here.
    var cols = sec.content.columns.map(function (c) { return String(c || "").trim().toUpperCase(); });
    var at = function (names) {
      for (var n = 0; n < names.length; n++) {
        var k = cols.indexOf(names[n].toUpperCase());
        if (k > -1) return k;
      }
      return -1;
    };
    var iName  = at(["Medicine Name", "Brand", "Generic"]);
    var iType  = at(["Type"]);
    var iSig   = at(["Dosage / Sig", "Strength / Dose", "Dose"]);
    var iDays  = at(["Days", "Duration"]);
    var iNotes = at(["Notes / Timing", "Instructions", "Timing"]);
    if (iName === -1) return out;

    var when = _fmtDate_(bestSignedAt);

    sec.content.rows.forEach(function (r) {
      var name = String(r[iName] || "").trim();
      if (!name) return;
      // The medicine cell may carry "BRAND (GENERIC)". The pharmacy matches
      // on the brand it stocks, so the generic goes to the instructions
      // where a substitution decision is actually made.
      var generic = "";
      var paren = name.match(/^(.*?)\s*\(([^)]+)\)\s*$/);
      if (paren) { name = paren[1].trim(); generic = paren[2].trim(); }

      var type = (iType > -1) ? String(r[iType] || "").trim() : "";
      var days = (iDays > -1) ? String(r[iDays] || "").trim() : "";
      var notes = [generic ? "generic " + generic : "",
                   (iNotes > -1) ? String(r[iNotes] || "").trim() : "",
                   days ? days + (/^\d+$/.test(days) ? " day(s)" : "") : ""]
                  .filter(String).join(" · ");

      out.push({
        source: "DISCHARGE",
        drugName: type ? (type + " " + name) : name,
        dose: (iSig > -1) ? String(r[iSig] || "").trim() : "",
        orderedBy: "Discharge summary",
        date: when,
        instructions: notes
      });
    });
  } catch (e) { /* the other three sources must still reach the counter */ }
  return out;
}

function _dedupePrescriptions_(list) {
  var seen = {}, out = [];
  for (var i = 0; i < list.length; i++) {
    var k = (list[i].source + "|" + list[i].drugName + "|" + (list[i].dose || "")).toLowerCase();
    if (seen[k]) continue;
    seen[k] = true; out.push(list[i]);
  }
  return out;
}

// =====================================================================
// SECTION C — ATOMIC + IDEMPOTENT BILLING
// =====================================================================

function processPharmacyBill(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  crescRequire_(sessionToken, 'pharmacy.dispense');
  if (!lock.tryLock(10000)) return { success: false, message: "System busy, please retry." };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = ss.getSheetByName(PH_SHEETS.INVENTORY);
    if (!invSheet) throw new Error("Inventory sheet missing.");

    var headerSheet = _ensureInvoiceHeaderSheet_(ss);
    var itemsSheet = _ensureInvoiceItemsSheet_(ss);

    // 0. Idempotency
    if (payload.billUuid) {
      var prior = _findInvoiceByUuid_(headerSheet, payload.billUuid);
      if (prior) return { success: true, invoiceNo: prior, duplicate: true, message: "This bill was already saved." };
    }

    var items = payload.billedItems || [];
    if (items.length === 0) throw new Error("No billable items supplied.");

    // 1. Index inventory by Brand|Batch
    var data = invSheet.getDataRange().getValues();
    var idx = {};
    for (var r = 1; r < data.length; r++) {
      var key = (String(data[r][1]).trim() + "|" + String(data[r][6]).trim()).toUpperCase();
      idx[key] = { rowNum: r + 1, qty: parseInt(data[r][4], 10) || 0 };
    }

    // 2. Validate every line BEFORE writing
    var pending = {};
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var k = (String(it.drug).trim() + "|" + String(it.batch).trim()).toUpperCase();
      var ref = idx[k];
      if (!ref) throw new Error("Not in stock: " + it.drug + " (Batch " + it.batch + ").");
      var want = parseFloat(it.qty) || 0;
      if (want <= 0) throw new Error("Invalid quantity for " + it.drug + ".");
      pending[k] = (pending[k] || 0) + want;
      if (pending[k] > ref.qty)
        throw new Error("Insufficient stock for " + it.drug + " (Batch " + it.batch + "). Available " + ref.qty + ", requested " + pending[k] + ".");
    }
    var deductions = [];
    for (var key in pending) deductions.push({ rowNum: idx[key].rowNum, newQty: idx[key].qty - pending[key] });

    // 3. Invoice number + commit deductions
    var now = new Date();
    var invoiceNo = _nextInvoiceNo_(headerSheet, now);
    for (var d = 0; d < deductions.length; d++) invSheet.getRange(deductions[d].rowNum, 5).setValue(deductions[d].newQty);

    // 4. Recompute money server-side
    var gross = 0, totalGst = 0, itemRows = [];
    for (var m = 0; m < items.length; m++) {
      var li = items[m];
      var qty = parseFloat(li.qty) || 0, mrp = parseFloat(li.rate) || 0, gstPct = parseFloat(li.gst) || 0;
      var lineTotal = qty * mrp;
      var gstAmt = lineTotal - (lineTotal / (1 + gstPct / 100));
      gross += lineTotal; totalGst += gstAmt;
      itemRows.push([invoiceNo, now, String(payload.patientId || "WALK-IN"),
        String(li.drug || ""), String(li.generic || ""), String(li.batch || ""), String(li.expiry || ""),
        qty, String(li.unit || ""), round2_(mrp), gstPct, round2_(lineTotal - gstAmt),
        round2_(gstAmt), round2_(lineTotal), String(li.rowId || "")]);
    }
    var discount = parseFloat(payload.discount) || 0;
    if (discount > gross) discount = gross;
    var net = round2_(gross - discount);

    // 5. Write header + items
    var payStatus = (String(payload.payMode).toUpperCase() === "CREDIT") ? "PENDING" : "PAID";
    var billedBy = ""; try { billedBy = Session.getActiveUser().getEmail() || ""; } catch (e) {}
    headerSheet.appendRow([invoiceNo, now, String(payload.billType || "WALK-IN"),
      String(payload.patientId || "WALK-IN"), String(payload.patientName || ""), String(payload.mobile || ""),
      String(payload.age || ""), String(payload.sex || ""), String(payload.address || ""),
      String(payload.doctor || "Self / OTC"), round2_(gross), round2_(totalGst), round2_(discount), net,
      String(payload.payMode || "CASH"), String(payload.txnId || ""), payStatus, "ACTIVE",
      items.length, billedBy, String(payload.billUuid || ""), "", ""]);
    itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, itemRows.length, itemRows[0].length).setValues(itemRows);
    // Route IP credit bills to the running tab. Discharge owns recognition.
    if (String(payload.payMode).toUpperCase() === "CREDIT") {
      try {
        var ipNo = ipc_activeAdmissionByPatient_(String(payload.patientId || ""));
        if (ipNo) billChargeToIp({
          ipNumber: ipNo, source: 'PHARMACY', sourceRef: invoiceNo,
          amount: net, gst: round2_(totalGst),
          description: items.length + ' pharmacy item(s)', user: billedBy
        });
     } catch (e) {}
    }

    // 6. Return confirmed print payload
    return { success: true, invoiceNo: invoiceNo, print: {
      invoiceNo: invoiceNo,
      date: Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM-yyyy HH:mm"),
      billType: String(payload.billType || "WALK-IN"), patientId: String(payload.patientId || "WALK-IN"),
      patientName: String(payload.patientName || "Walk-in Patient"), mobile: String(payload.mobile || ""),
      age: String(payload.age || ""), sex: String(payload.sex || ""), address: String(payload.address || ""),
      doctor: String(payload.doctor || "Self / OTC"), payMode: String(payload.payMode || "CASH"),
      txnId: String(payload.txnId || ""), payStatus: payStatus,
      items: itemRows.map(function (row) { return { drug: row[3], generic: row[4], batch: row[5],
        expiry: row[6], qty: row[7], unit: row[8], mrp: row[9], gst: row[10],
        taxable: row[11], gstAmt: row[12], lineTotal: row[13] }; }),
      gross: round2_(gross), totalGst: round2_(totalGst), discount: round2_(discount), net: net } };
  } catch (error) {
    return { success: false, message: error.message || String(error) };
  } finally { lock.releaseLock(); }
}

// =====================================================================
// SECTION D — CREDIT SETTLEMENT (close IPD / staff credit bills)
// =====================================================================

function settleCreditBill(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  crescRequire_(sessionToken, 'billing.write');
  if (!lock.tryLock(10000)) return { success: false, message: "System busy, please retry." };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.INVOICES);
    if (!sheet) throw new Error("Invoice ledger not found.");
    var data = sheet.getDataRange().getValues();
    var target = String(payload.invoiceNo || "").trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== target) continue;
      if (String(data[i][16]).trim().toUpperCase() === "PAID") return { success: false, message: "Invoice " + target + " is already settled." };
      var rowNum = i + 1;
      var by = ""; try { by = Session.getActiveUser().getEmail() || ""; } catch (e) {}
      sheet.getRange(rowNum, 15).setValue(String(payload.payMode || "CASH")); 
      sheet.getRange(rowNum, 16).setValue(String(payload.txnId || ""));       
      sheet.getRange(rowNum, 17).setValue("PAID");                            
      sheet.getRange(rowNum, 22).setValue(new Date());                        
      sheet.getRange(rowNum, 23).setValue(by);                               
      return { success: true, message: "Invoice " + target + " settled." };
    }
    return { success: false, message: "Invoice " + target + " not found." };
  } catch (error) {
    return { success: false, message: error.message || String(error) };
  } finally { lock.releaseLock(); }
}

// =====================================================================
// SECTION E — INVOICE SHEET HELPERS + UTILITIES
// =====================================================================

function _ensureInvoiceHeaderSheet_(ss) {
  var sheet = ss.getSheetByName(PH_SHEETS.INVOICES);
  if (!sheet) {
    sheet = ss.insertSheet(PH_SHEETS.INVOICES);
    sheet.appendRow(PH_INVOICE_HEADERS);
    sheet.getRange(1, 1, 1, PH_INVOICE_HEADERS.length).setFontWeight("bold").setBackground("#d9ead3");
    sheet.setFrozenRows(1);
  } else {
    var width = Math.max(sheet.getLastColumn(), PH_INVOICE_HEADERS.length);
    var header = sheet.getRange(1, 1, 1, width).getValues()[0];
    for (var c = 0; c < PH_INVOICE_HEADERS.length; c++) {
      if (String(header[c] || "").trim() !== PH_INVOICE_HEADERS[c]) {
        sheet.getRange(1, c + 1).setValue(PH_INVOICE_HEADERS[c]).setFontWeight("bold");
      }
    }
  }
  return sheet;
}

function _ensureInvoiceItemsSheet_(ss) {
  var sheet = ss.getSheetByName(PH_SHEETS.INVOICE_ITEMS);
  if (!sheet) {
    sheet = ss.insertSheet(PH_SHEETS.INVOICE_ITEMS);
    sheet.appendRow(["Invoice_No","Timestamp","Patient_ID","Brand","Generic","Batch","Expiry",
      "Qty","Unit","MRP","GST_Pct","Taxable","GST_Amt","Line_Total","Inventory_RowId"]);
    sheet.getRange("A1:O1").setFontWeight("bold").setBackground("#d9ead3");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _findInvoiceByUuid_(headerSheet, uuid) {
  if (!uuid || headerSheet.getLastRow() <= 1) return null;
  var data = headerSheet.getDataRange().getValues();
  var col = data[0].indexOf("Bill_UUID");
  if (col < 0) return null;
  for (var i = 1; i < data.length; i++)
    if (String(data[i][col]).trim() === String(uuid).trim()) return String(data[i][0]);
  return null;
}

function _nextInvoiceNo_(headerSheet, now) {
  var yymm = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyMM");
  var maxSeq = 1000;
  if (headerSheet.getLastRow() > 1) {
    var nos = headerSheet.getRange(2, 1, headerSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < nos.length; i++) {
      var mm = String(nos[i][0]).match(/-(\d+)$/);
      if (mm) { var s = parseInt(mm[1], 10); if (s > maxSeq) maxSeq = s; }
    }
  }
  return "PH" + yymm + "-" + (maxSeq + 1);
}

function round2_(n) { return Math.round((parseFloat(n) || 0) * 100) / 100; }
function _fmtDate_(v) { return cresc_formatDate_(v, "dd-MMM-yyyy") || String(v || ""); }

/**
 * 'yyyy-MM-dd' for a prescription row.
 *
 * This used to recognise only a yyyy-MM-dd substring and return the raw text
 * for everything else. A ward order dated "13/09/2026" therefore keyed as
 * "13/09/2026", and _dateKeyNum_ turned that into 13092026 - a number that
 * sorts BELOW 20260913. Picking "the latest note" then picked the wrong one,
 * and the pharmacy dispensed against a superseded prescription.
 */
function _dateKey_(v) {
  var k = cresc_dayKey_(v);
  return k || String(v || "");
}
function _dateKeyNum_(v) { var n = parseInt(_dateKey_(v).replace(/[-\/]/g, ""), 10); return isNaN(n) ? 0 : n; }
function _expiryToSortKey_(exp) {
  if (!exp) return 999999;
  var m = String(exp).match(/(\d{4})[-\/](\d{1,2})/);
  return m ? parseInt(m[1], 10) * 12 + parseInt(m[2], 10) : 999999;
}

// =====================================================================
// SECTION F — DIGITAL DISPATCH ENGINE (PDF, WhatsApp, Email)
// =====================================================================

function generateAndStorePharmacyInvoicePDF(invoiceNo, htmlContent, sessionToken) {
  try {
    crescRequire_(sessionToken, 'billing.read');

    // ── DPDP s.6 / s.5: the patient's COMMUNICATION consent, checked here ──
    // The register has carried this purpose since it was built and nothing
    // read it. The WhatsApp dispatch now refuses unless the patient has agreed to be sent
    // documents this way, and the refusal says how to ask them. See
    // DPDP_Dispatch.gs.
    var __pid = dpdp_resolvePatientFor_('PHARMACY_INVOICE', invoiceNo);
    var __gate = dpdpRequireDispatchConsent_(__pid, 'WHATSAPP',
                   crescActor_(sessionToken), 'PHARMACY_INVOICE', { allowAnonymous: true });
    if (!__gate.ok) return { success: false, code: 'CONSENT_REQUIRED',
                             consentState: __gate.state, patientId: __pid,
                             notice: __gate.notice, message: __gate.message };
    const htmlBlob = Utilities.newBlob(htmlContent, 'text/html', 'invoice.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Crescentia_Pharmacy_Invoice_" + invoiceNo + ".pdf");
    
    const rootFolders = DriveApp.getFoldersByName("Crescentia_Pharmacy_Invoices");
    const rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder("Crescentia_Pharmacy_Invoices");
    
    const now = new Date();
    const yearStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
    const monthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");
    
    const yearFolder = rootFolder.getFoldersByName(yearStr).hasNext() ? rootFolder.getFoldersByName(yearStr).next() : rootFolder.createFolder(yearStr);
    const monthFolder = yearFolder.getFoldersByName(monthStr).hasNext() ? yearFolder.getFoldersByName(monthStr).next() : yearFolder.createFolder(monthStr);
    
    const file = monthFolder.createFile(pdfBlob);

    // Registered, not just shared. dpdpIssueDocumentLink_ publishes the
    // file to Drive so the patient gets a link WhatsApp can preview and
    // their phone can open, and writes it into Document_Grants with an
    // expiry — after which the daily sweep makes the file private again
    // and every forwarded copy of the link stops working. A bare
    // setSharing(ANYONE_WITH_LINK) here would publish this invoice
    // for ever with nothing able to take it back. See DPDP_Documents.gs.
    // __pid, not invoiceNo. The register's Patient_ID column was being given
    // an invoice number, so dpdpListDocumentLinks(patient) returned nothing
    // and the s.11(1)(b) answer to "what have you sent about me" was empty.
    var grant = dpdpIssueDocumentLink_(file, 'PHARMACY_INVOICE', __pid || invoiceNo,
                                       (crescActor_(sessionToken) || {}).username || '');
    if (!grant.success) return { success: false, message: grant.message };

    return { success: true, link: grant.url, expiresAt: grant.expiresAt,
             grantId: grant.grantId, message: grant.message };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function emailPharmacyInvoice(invoiceNo, htmlContent, patientEmail, sessionToken) {
  try {
    crescRequire_(sessionToken, 'billing.read');
    if (!patientEmail) throw new Error("No valid email address provided.");

    // ── DPDP s.6 / s.5: the patient's COMMUNICATION consent, checked here ──
    // The register has carried this purpose since it was built and nothing
    // read it. The email dispatch now refuses unless the patient has agreed to be sent
    // documents this way, and the refusal says how to ask them. See
    // DPDP_Dispatch.gs.
    var __pid = dpdp_resolvePatientFor_('PHARMACY_INVOICE', invoiceNo);
    var __gate = dpdpRequireDispatchConsent_(__pid, 'EMAIL',
                   crescActor_(sessionToken), 'PHARMACY_INVOICE', { allowAnonymous: true });
    if (!__gate.ok) return { success: false, code: 'CONSENT_REQUIRED',
                             consentState: __gate.state, patientId: __pid,
                             notice: __gate.notice, message: __gate.message };
    
    const htmlBlob = Utilities.newBlob(htmlContent, 'text/html', 'invoice.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Crescentia_Pharmacy_Invoice_" + invoiceNo + ".pdf");
    
    const subject = "Your Pharmacy Invoice - Crescentia HealthTech (" + invoiceNo + ")";
    const body = "Dear Patient,\n\nPlease find attached your pharmacy invoice " + invoiceNo + ".\n\nThank you for choosing Crescentia HealthTech.";
    
    GmailApp.sendEmail(patientEmail, subject, body, {
      attachments: [pdfBlob],
      name: "Crescentia Pharmacy"
    });
    
    return { success: true, message: "Invoice emailed successfully." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}