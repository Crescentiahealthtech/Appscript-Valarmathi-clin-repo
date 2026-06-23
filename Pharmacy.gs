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
  MASTER_LEDGER:"Pharmacy_Master"
};

var PH_INVOICE_HEADERS = ["Invoice_No","Timestamp","Bill_Type","Patient_ID","Patient_Name",
  "Mobile","Age","Sex","Address","Doctor","Gross","Total_GST","Discount","Net",
  "Pay_Mode","Txn_ID","Pay_Status","Status","Item_Count","Created_By","Bill_UUID",
  "Settled_At","Settled_By"];


// =====================================================================
// SECTION A — INVENTORY (used by Ledger + Add screens)
// =====================================================================

function fetchPharmacyInventory() {
  try {
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
        parseFloat(payload.gst) || 0, String(payload.manufacturer || ""), String(payload.supplier || ""), "Admin"]);
    }
    return { success: true, message: "Inventory updated securely." };
  } catch (error) {
    return { success: false, message: "Failed to update: " + error.toString() };
  } finally { lock.releaseLock(); }
}

// =====================================================================
// SECTION B — BILLING DATA READS
// =====================================================================

function fetchBillableStock() {
  try {
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

function getPatientBillingContext(query) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var patient = _findPatient_(ss, query);
    var pid = patient ? String(patient.id) : String(query || "").trim();
    var rx = []
      .concat(_readOpdPrescriptions_(ss, pid))
      .concat(_readIpCasesheetPrescriptions_(ss, pid))
      .concat(_readIpWardPrescriptions_(ss, pid));
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

function processPharmacyBill(payload) {
  var lock = LockService.getScriptLock();
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

function getPendingCreditBills() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.INVOICES);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][16]).trim().toUpperCase() !== "PENDING") continue;
      if (String(data[i][17]).trim().toUpperCase() !== "ACTIVE") continue;
      out.push({ invoiceNo: String(data[i][0]), date: _fmtDate_(data[i][1]),
        billType: String(data[i][2] || ""), patientId: String(data[i][3] || ""),
        patientName: String(data[i][4] || ""), net: parseFloat(data[i][13]) || 0 });
    }
    out.reverse();
    return { success: true, data: out };
  } catch (error) {
    return { success: false, message: "Could not load credits: " + error.toString() };
  }
}

function settleCreditBill(payload) {
  var lock = LockService.getScriptLock();
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
function _fmtDate_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), "dd-MMM-yyyy") : String(v || ""); }
function _dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var s = String(v || ""), m = s.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  return m ? m[1] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[3]).slice(-2) : s;
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

function generateAndStorePharmacyInvoicePDF(invoiceNo, htmlContent) {
  try {
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
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, link: file.getUrl() };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function emailPharmacyInvoice(invoiceNo, htmlContent, patientEmail) {
  try {
    if (!patientEmail) throw new Error("No valid email address provided.");
    
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