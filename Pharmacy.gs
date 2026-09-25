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

var PH_ITEM_HEADERS = ["Invoice_No","Timestamp","Patient_ID","Brand","Generic","Batch","Expiry",
  "Qty","Unit","MRP","GST_Pct","Taxable","GST_Amt","Line_Total","Inventory_RowId","Schedule"];

/** Modes a pharmacy bill may be paid by. CREDIT is settled later. */
var PH_PAY_MODES = ["CASH", "UPI", "CARD", "CREDIT"];

// ---------------------------------------------------------------------
// DRUG SCHEDULES (Drugs and Cosmetics Rules, 1945)
//
// Schedule H, H1 and X medicines may only be sold against a prescription,
// and the pharmacy has to be able to show — for any period an inspector
// names — who was given what, on whose prescription. That is the register
// getScheduleDrugRegister() prints.
//
// The schedule is a property of the MEDICINE, recorded on its inventory
// rows (column "Schedule", added on first use), and snapshotted onto every
// invoice line at the moment of sale, so a later reclassification never
// rewrites what the register says was sold.
// ---------------------------------------------------------------------
var PH_SCHEDULE_HEADER = "Schedule";
var PH_SCHEDULES = { H: "Schedule H", H1: "Schedule H1", X: "Schedule X" };

/** "Schedule H1", "sch h1", "h1" -> "H1"; anything unrecognised -> "". */
function _phSchedule_(v) {
  var k = String(v == null ? "" : v).toUpperCase().replace(/SCHEDULE|SCH\.?/g, "").replace(/[^A-Z0-9]/g, "");
  return PH_SCHEDULES[k] ? k : "";
}

/**
 * Makes sure the sheet's grid is at least `n` columns wide. getRange() past
 * the last column of the GRID throws, and a sheet whose unused columns were
 * deleted by hand is exactly as wide as its data.
 */
function _phEnsureWidth_(sheet, n) {
  try {
    var max = sheet.getMaxColumns();
    if (max < n) sheet.insertColumnsAfter(max, n - max);
  } catch (e) { /* a fake sheet in a test, or no permission: the write reports itself */ }
}

/** The zero-based Schedule column on Pharmacy_Inventory; -1 if absent and not ensured. */
function _phScheduleCol_(sheet, ensure) {
  var width = Math.max(1, sheet.getLastColumn());
  var head = sheet.getRange(1, 1, 1, width).getValues()[0]
    .map(function (h) { return String(h || "").trim().toLowerCase(); });
  var i = head.indexOf(PH_SCHEDULE_HEADER.toLowerCase());
  if (i !== -1 || !ensure) return i;
  // Column O at the earliest: the fourteen before it are read by position.
  var col = Math.max(width, 14) + 1;
  _phEnsureWidth_(sheet, col);
  sheet.getRange(1, col).setValue(PH_SCHEDULE_HEADER).setFontWeight("bold").setBackground("#f4cccc");
  return col - 1;
}

/** { "BRAND (lower-case)": "H" | "H1" | "X" } from every inventory row that carries one. */
function _phScheduleByBrand_(data, col) {
  var out = {};
  if (col < 0) return out;
  for (var i = 1; i < data.length; i++) {
    var sc = _phSchedule_(data[i][col]);
    var b = String(data[i][1] || "").trim().toLowerCase();
    if (sc && b && !out[b]) out[b] = sc;
  }
  return out;
}

/**
 * Whether the doctor on a bill names somebody. "Self / OTC" is the desk's
 * default and means nobody prescribed it.
 */
function _phNamesDoctor_(doctor) {
  var d = String(doctor || "").trim().toUpperCase().replace(/\s+/g, " ");
  return !!d && ["SELF / OTC", "SELF/OTC", "SELF", "OTC", "NONE", "NIL", "-", "N/A", "NA"].indexOf(d) === -1;
}


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
    var schCol = _phScheduleCol_(sheet, false);
    var schByBrand = _phScheduleByBrand_(data, schCol);

    var out = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1]) continue;
      out.push({
        rowId: i + 1,
        schedule: (schCol >= 0 && _phSchedule_(row[schCol])) ||
                  schByBrand[String(row[1]).trim().toLowerCase()] || "",
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
  try {
    var actor = crescRequire_((payload || {}).token, 'pharmacy.stock_add');
    payload = payload || {};

    // WHAT A STOCK ENTRY MUST CARRY. This used to append whatever arrived:
    // a quantity of -5, an MRP of 0, a batch with no number.
    var brand = String(payload.medicineName || "").trim();
    var batch = String(payload.batchNo || "").trim().toUpperCase();
    var qty = parseInt(payload.qty, 10);
    var mrp = parseFloat(payload.mrp);
    var buy = parseFloat(payload.buyPrice);
    var gst = parseFloat(payload.gst);
    if (!brand) return { success: false, message: "Enter the medicine (brand) name." };
    if (!batch) return { success: false, message: "Enter the batch number printed on the pack." };
    if (!(qty > 0)) return { success: false, message: "Enter how many units were received." };
    if (!(mrp > 0)) return { success: false, message: "Enter the MRP printed on the pack." };
    if (isNaN(buy) || buy < 0) return { success: false, message: "Enter the buying price." };
    if (buy > mrp) return { success: false, message: "The buying price (" + buy + ") is above the MRP (" + mrp + "). Check both." };
    if (isNaN(gst) || gst < 0 || gst > 28) return { success: false, message: "GST must be between 0 and 28%." };

    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.INVENTORY);
    if (!sheet) {
      sheet = ss.insertSheet(PH_SHEETS.INVENTORY);
      sheet.appendRow(["Timestamp","Brand Name","Generic Name","Type","Qty","Unit",
        "Batch No","Expiry Date","Rack Location","Buy Price","MRP","GST %","Manufacturer","Supplier",
        PH_SCHEDULE_HEADER]);
      sheet.getRange("A1:O1").setFontWeight("bold").setBackground("#d9d9d9");
    }
    var schCol = _phScheduleCol_(sheet, true);
    var data = sheet.getDataRange().getValues();

    // The schedule is the medicine's, so a new batch of a brand already on
    // the shelf inherits it unless the form says otherwise.
    var schedule = _phSchedule_(payload.schedule) ||
                   _phScheduleByBrand_(data, schCol)[brand.toLowerCase()] || "";

    // THE SAME BATCH, RECEIVED AGAIN. Appending a second row with the same
    // brand and batch split the stock in two, and billing — which looks a
    // batch up by brand and batch number — found only one of them: the other
    // row's units could never be sold, and a sale could decrement the wrong
    // one. A re-delivery of the same batch at the same MRP now tops up the
    // row that is already there.
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1] || "").trim().toUpperCase() !== brand.toUpperCase()) continue;
      if (String(data[i][6] || "").trim().toUpperCase() !== batch) continue;
      var oldMrp = parseFloat(data[i][10]) || 0;
      if (Math.abs(oldMrp - mrp) > 0.009) {
        return { success: false,
                 message: brand + " batch " + batch + " is already in stock at MRP " + oldMrp.toFixed(2) +
                          ". The same batch cannot carry two prices — check the pack, or adjust the " +
                          "existing batch from Live Inventory." };
      }
      var before = parseInt(data[i][4], 10) || 0;
      sheet.getRange(i + 1, 5).setValue(before + qty);
      if (schedule && schCol >= 0) sheet.getRange(i + 1, schCol + 1).setValue(schedule);
      logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
                'PHARMACY_STOCK_RECEIVED', 'Pharmacy_Inventory', brand + ' / ' + batch,
                { qty: qty, before: before, after: before + qty, toppedUp: true });
      return { success: true,
               message: qty + " unit(s) added to the existing " + brand + " batch " + batch +
                        " — " + (before + qty) + " now in stock." };
    }

    var row = [
      payload.timestamp ? (cresc_parseDate_(payload.timestamp) || new Date()) : new Date(),
      brand, String(payload.genericName || "").trim(),
      String(payload.drugType || ""), qty, String(payload.unit || ""),
      batch, String(payload.expiryDate || ""), String(payload.rackLocation || "").trim().toUpperCase(),
      buy, mrp, gst,
      String(payload.manufacturer || "").trim(), String(payload.supplier || "").trim()
    ];
    while (row.length < schCol) row.push("");
    row[schCol] = schedule;
    sheet.appendRow(row);
    logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
              'PHARMACY_STOCK_RECEIVED', 'Pharmacy_Inventory', brand + ' / ' + batch,
              { qty: qty, mrp: mrp, schedule: schedule });
    return { success: true,
             message: "Stock added: " + qty + " × " + brand + " (batch " + batch + ")" +
                      (schedule ? " — " + PH_SCHEDULES[schedule] + "." : ".") };
  } catch (error) {
    return { success: false, message: _phReason_(error) };
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

function updatePharmacyStock(payload) {
  var lock = LockService.getScriptLock();
  try {
    var actor = crescRequire_((payload || {}).token, 'pharmacy.stock_edit');
    payload = payload || {};
    var newStock = parseInt(payload.stock, 10);
    var mrp = parseFloat(payload.mrp), gst = parseFloat(payload.gst);
    if (isNaN(newStock) || newStock < 0) return { success: false, message: "The stock count cannot be negative." };
    if (!(mrp > 0)) return { success: false, message: "Enter the MRP." };
    if (isNaN(gst) || gst < 0 || gst > 28) return { success: false, message: "GST must be between 0 and 28%." };

    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stockSheet = ss.getSheetByName(PH_SHEETS.INVENTORY);
    var masterSheet = ss.getSheetByName(PH_SHEETS.MASTER_LEDGER);
    var rowNum = parseInt(payload.rowId, 10);
    if (!stockSheet || !(rowNum > 1) || rowNum > stockSheet.getLastRow()) {
      return { success: false, message: "That batch is no longer in the inventory. Refresh and try again." };
    }

    // rowId is a position, not an identity — the same check the write-off
    // makes. An edit that lands on whatever row now sits at that position
    // silently rewrites a different medicine's count and price.
    var current = stockSheet.getRange(rowNum, 1, 1, 14).getValues()[0];
    var brand = String(current[1] || "").trim();
    var batch = String(current[6] || "").trim();
    if ((payload.brandName && String(payload.brandName).trim().toUpperCase() !== brand.toUpperCase()) ||
        (payload.batch && String(payload.batch).trim().toUpperCase() !== batch.toUpperCase())) {
      return { success: false,
               message: "The inventory has changed since this screen was loaded — row " + rowNum +
                        " now holds " + (brand || "(blank)") + " batch " + (batch || "(blank)") +
                        ". Refresh and try again. Nothing was changed." };
    }

    var before = parseInt(current[4], 10) || 0;
    stockSheet.getRange(rowNum, 5).setValue(newStock);
    stockSheet.getRange(rowNum, 8).setValue(String(payload.expiry || ""));
    stockSheet.getRange(rowNum, 9).setValue(String(payload.rack || "").trim().toUpperCase());
    stockSheet.getRange(rowNum, 11).setValue(mrp);
    stockSheet.getRange(rowNum, 12).setValue(gst);

    // A schedule belongs to the medicine, so it is set on every batch of it:
    // an H1 drug whose older batch still said "none" would be sold from that
    // batch without the prescription check.
    var scheduleNote = "";
    if (payload.schedule !== undefined) {
      var sc = _phSchedule_(payload.schedule);
      var schCol = _phScheduleCol_(stockSheet, true);
      var all = stockSheet.getRange(1, 1, stockSheet.getLastRow(), schCol + 1).getValues();
      var changed = 0;
      for (var r = 1; r < all.length; r++) {
        if (String(all[r][1] || "").trim().toUpperCase() !== brand.toUpperCase()) continue;
        if (_phSchedule_(all[r][schCol]) === sc) continue;
        stockSheet.getRange(r + 1, schCol + 1).setValue(sc);
        changed++;
      }
      if (changed) scheduleNote = " " + brand + " is now " + (sc ? PH_SCHEDULES[sc] : "not scheduled") +
                                  " on " + changed + " batch row(s).";
    }

    if (masterSheet) {
      masterSheet.appendRow([new Date(), "ADJ-" + Utilities.getUuid().substring(0, 6).toUpperCase(),
        "Manual Adjustment", brand, String(current[2] || ""),
        "N/A", "N/A", newStock, batch,
        String(payload.expiry || ""), String(payload.rack || ""), 0, mrp,
        gst, String(current[12] || ""), String(current[13] || ""),
        // The actor column used to be the literal string "Admin", whoever was
        // signed in — so the master ledger recorded every manual adjustment
        // in this clinic's history as having been made by the same person.
        actor.displayName || actor.username]);
    }
    logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
              'PHARMACY_STOCK_ADJUST', 'Pharmacy_Inventory', batch,
              { brand: brand, stockBefore: before, stock: newStock, mrp: mrp, gst: gst });
    return { success: true, message: "Inventory updated." + scheduleNote };
  } catch (error) {
    return { success: false, message: "Failed to update: " + _phReason_(error) };
  } finally { try { lock.releaseLock(); } catch (e) {} }
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
/**
 * The write-off reasons, as { code, label } pairs in `data`.
 *
 * Guard inside the try, reply wrapped — see processPharmacyBill above. The
 * discard dialog's failure path set the reason list to empty and carried on
 * opening, so a refused call produced a dialog whose only mandatory field
 * had nothing in it and no explanation anywhere.
 */
function getDisposalReasons(sessionToken) {
  try {
    crescRequire_(sessionToken, 'pharmacy.read');
    return { success: true, message: '', data: Object.keys(PH_DISPOSAL_REASONS).map(function (k) {
      return { code: k, label: PH_DISPOSAL_REASONS[k] };
    }) };
  } catch (err) {
    return { success: false, data: [], message: _phReason_(err) };
  }
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
        batch:   String(r[4] || ''), expiry:  cresc_expiryText_(r[5]),
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

/**
 * FRONTEND ENTRY. REMOVES A BATCH ROW OUTRIGHT.
 *
 * DELETE IS NOT DISPOSAL, and the difference is the whole reason both exist.
 *
 *   DISPOSAL is stock that was real and has left the shelf — expired,
 *   damaged, lost. It keeps the row, zeroes the count and writes a permanent
 *   entry the clinic can total: this is what the wastage cost.
 *
 *   DELETE is a row that should never have existed. A duplicate entry, a
 *   typed batch number nobody can find, stock keyed against the wrong brand.
 *   Writing that off as wastage would inflate the disposal register with
 *   money the clinic never lost.
 *
 * Until now only disposal existed, so a mis-keyed row could be dealt with
 * only by writing it off as damaged — which is a false entry in the one
 * register whose value is that it is true.
 *
 * TWO REFUSALS keep the distinction honest:
 *
 *   A batch that has been BILLED is not a mistake. Something was dispensed
 *   against it, and an invoice line refers to it. It has to be written off,
 *   not removed, or the invoice points at a batch that no longer exists.
 *
 *   A reason is required, and the whole row is copied into the audit entry
 *   before it goes, so a deletion can be read back and undone by hand.
 *
 * @param {{rowId:number, reason:string, token:string}} payload
 */
function deletePharmacyBatch(payload) {
  var lock = LockService.getScriptLock();
  try {
    var p = payload || {};
    // The same permission a write-off needs: removing stock from the record
    // is the same kind of act whichever register it lands in.
    var actor = crescRequire_(p.token, 'pharmacy.stock_discard');

    var rowId = parseInt(p.rowId, 10);
    if (!rowId || rowId < 2) return { success: false, message: "Which row?" };

    var reason = String(p.reason || '').trim();
    if (reason.length < 5) {
      return { success: false,
               message: 'Say why this row is being removed. "Deleted" is not a ' +
                        'reason, and this is the only record that it happened.' };
    }

    if (!lock.tryLock(10000)) return { success: false, message: "System busy, please retry." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.INVENTORY);
    if (!sheet) return { success: false, message: "Pharmacy_Inventory not found." };
    if (rowId > sheet.getLastRow()) return { success: false, message: "That row no longer exists." };

    var row = sheet.getRange(rowId, 1, 1, sheet.getLastColumn()).getValues()[0];
    var brand = String(row[1] || '').trim();
    var batch = String(row[6] || '').trim();
    if (!brand) return { success: false, message: "That row is already empty." };

    // Has anything been dispensed against this batch?
    var billed = 0;
    try {
      var items = ss.getSheetByName(PH_SHEETS.INVOICE_ITEMS);
      if (items && items.getLastRow() > 1) {
        var idata = items.getDataRange().getValues();
        for (var i = 1; i < idata.length; i++) {
          if (String(idata[i][3] || '').trim().toLowerCase() !== brand.toLowerCase()) continue;
          if (String(idata[i][5] || '').trim().toLowerCase() !== batch.toLowerCase()) continue;
          billed++;
        }
      }
    } catch (e) { /* an unreadable ledger must not block the check below */ }

    if (billed > 0) {
      return { success: false, code: 'BILLED',
               message: brand + ' batch ' + (batch || '(no batch)') + ' appears on ' +
                        billed + ' invoice line' + (billed === 1 ? '' : 's') + ', so it ' +
                        'is stock that really existed. Write it off instead — ' +
                        'deleting it would leave those invoices pointing at a ' +
                        'batch that is not there.' };
    }

    // The whole row goes into the audit entry BEFORE it is removed, so the
    // deletion is reversible by somebody reading the log.
    var snapshot = {
      brand: brand,
      generic: String(row[2] || ''),
      type: String(row[3] || ''),
      qty: parseInt(row[4], 10) || 0,
      unit: String(row[5] || ''),
      batch: batch,
      expiry: cresc_expiryText_(row[7]),
      rack: String(row[8] || ''),
      buyPrice: parseFloat(row[9]) || 0,
      mrp: parseFloat(row[10]) || 0,
      gst: parseFloat(row[11]) || 0,
      manufacturer: String(row[12] || ''),
      supplier: String(row[13] || '')
    };

    sheet.deleteRow(rowId);
    SpreadsheetApp.flush();

    try {
      logAudit_({ username: actor.username, role: actor.role },
                'PHARMACY_BATCH_DELETED', 'Pharmacy_Inventory',
                brand + ' / ' + (batch || '-'),
                { reason: reason, row: snapshot });
    } catch (e) {}

    return { success: true,
             message: brand + (batch ? ' (batch ' + batch + ')' : '') +
                      ' removed. The row is in the audit log if it has to come back.' };
  } catch (error) {
    return { success: false, message: String(error.message || error).replace('FORBIDDEN: ', '') };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
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
    var doctors = _phClinicDoctors_();
    if (!sheet) return { success: true, data: [], doctors: doctors };
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [], doctors: doctors };
    var schCol = _phScheduleCol_(sheet, false);
    var schByBrand = _phScheduleByBrand_(data, schCol);

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
        mrp: parseFloat(row[10]) || 0, gst: parseFloat(row[11]) || 0, rack: String(row[8] || ""),
        schedule: (schCol >= 0 && _phSchedule_(row[schCol])) || schByBrand[brand.toLowerCase()] || ""
      });
    }
    batches.sort(function (a, b) {
      if (a.brand.toLowerCase() < b.brand.toLowerCase()) return -1;
      if (a.brand.toLowerCase() > b.brand.toLowerCase()) return 1;
      return a.expSort - b.expSort;
    });
    return { success: true, data: batches, doctors: doctors };
  } catch (error) {
    return { success: false, message: "Stock load failed: " + _phReason_(error) };
  }
}

/**
 * The clinic's own doctors, for the "Prescribing doctor" box. It was three
 * names typed into the page, two of whom are not on the Doctors sheet; an
 * outside prescriber is typed in free.
 */
function _phClinicDoctors_() {
  try {
    if (typeof getActiveDoctors_ !== 'function') return [];
    return (getActiveDoctors_() || []).map(function (d) { return String(d.name || '').trim(); })
      .filter(function (n) { return n && !/visiting/i.test(n); });
  } catch (e) { return []; }
}

function getPatientBillingContext(query, sessionToken) {
  try {
    crescRequire_(sessionToken, ['pharmacy.read', 'pharmacy.bill']);
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
  // THE GUARD IS INSIDE THE TRY, and it is not a style preference.
  //
  // crescRequire_ THROWS on an expired session or a role without
  // pharmacy.dispense. Thrown out of a frontend entry point, that reaches
  // google.script.run's FAILURE handler, not its success handler — and the
  // failure handler at this desk says "Network error: please retry." So a
  // pharmacist whose session had timed out pressed Save Bill, was told the
  // network was down, retried, and was told the same thing again. Nothing
  // on screen ever mentioned signing in.
  //
  // Returned as { success:false, message } it lands in the success handler,
  // where the real sentence is already displayed.
  var lock = null;
  try {
    var actor = crescRequire_(sessionToken, 'pharmacy.dispense');
    payload = payload || {};

    var payMode = String(payload.payMode || "CASH").trim().toUpperCase();
    if (PH_PAY_MODES.indexOf(payMode) === -1) {
      return { success: false, message: "Choose how the bill is paid: cash, UPI, card or credit." };
    }
    var discount = parseFloat(payload.discount) || 0;
    if (discount < 0) return { success: false, message: "The discount cannot be negative." };

    lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: "System busy, please retry." };
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

    // 1. Index inventory. BY ROW first: the desk sends the row each batch was
    //    picked from, and two rows can share a brand and batch number (a
    //    re-delivery keyed twice). Indexing by brand|batch alone kept only the
    //    LAST such row, so the other's stock could never be sold and a sale
    //    could come off the wrong row.
    var data = invSheet.getDataRange().getValues();
    var schCol = _phScheduleCol_(invSheet, false);
    var schByBrand = _phScheduleByBrand_(data, schCol);
    var byKey = {}, byRow = {};
    for (var r = 1; r < data.length; r++) {
      var key = (String(data[r][1]).trim() + "|" + String(data[r][6]).trim()).toUpperCase();
      var ref = {
        rowNum: r + 1, key: key, qty: parseInt(data[r][4], 10) || 0,
        brand: String(data[r][1] || "").trim(), generic: String(data[r][2] || ""),
        unit: String(data[r][5] || ""), batch: String(data[r][6] || "").trim(),
        expiry: data[r][7], mrp: parseFloat(data[r][10]) || 0, gst: parseFloat(data[r][11]) || 0,
        schedule: (schCol >= 0 && _phSchedule_(data[r][schCol])) ||
                  schByBrand[String(data[r][1] || "").trim().toLowerCase()] || ""
      };
      byRow[r + 1] = ref;
      (byKey[key] = byKey[key] || []).push(ref);
    }

    // 2. Validate every line BEFORE writing
    var pending = {}, lines = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i] || {};
      var k = (String(it.drug).trim() + "|" + String(it.batch).trim()).toUpperCase();
      var row = byRow[parseInt(it.rowId, 10)];
      if (!row || row.key !== k) {
        var cands = (byKey[k] || []).filter(function (c) { return c.qty - (pending[c.rowNum] || 0) > 0; });
        row = cands[0] || (byKey[k] || [])[0];
      }
      if (!row) throw new Error("Not in stock: " + it.drug + " (Batch " + it.batch + ").");
      var want = parseFloat(it.qty) || 0;
      if (want <= 0) throw new Error("Invalid quantity for " + it.drug + ".");
      pending[row.rowNum] = (pending[row.rowNum] || 0) + want;
      if (pending[row.rowNum] > row.qty)
        throw new Error("Insufficient stock for " + it.drug + " (Batch " + it.batch + "). Available " + row.qty + ", requested " + pending[row.rowNum] + ".");
      if (!(row.mrp > 0)) throw new Error(row.brand + " batch " + row.batch + " has no MRP on the inventory. Set it before selling.");
      lines.push({ ref: row, qty: want });
    }

    // 2b. THE PRESCRIPTION RULE. A Schedule H, H1 or X medicine is sold
    //     against a prescription, so the bill must name who prescribed it and
    //     who it is for — the two columns the register exists to show. The
    //     desk asks too; this is the check that holds when it does not.
    var scheduled = lines.filter(function (l) { return !!l.ref.schedule; });
    if (scheduled.length) {
      var names = scheduled.map(function (l) { return l.ref.brand + " (" + l.ref.schedule + ")"; })
        .filter(function (v, idx, a) { return a.indexOf(v) === idx; });
      if (!_phNamesDoctor_(payload.doctor)) {
        return { success: false, code: "SCHEDULE_DOCTOR_REQUIRED",
                 message: names.join(", ") + (names.length > 1 ? " are" : " is") +
                          " a scheduled drug and can only be sold on a prescription. Enter the " +
                          "prescribing doctor's name — \"Self / OTC\" is not accepted." };
      }
      if (!String(payload.patientName || "").trim()) {
        return { success: false, code: "SCHEDULE_PATIENT_REQUIRED",
                 message: "Enter the patient's name: " + names.join(", ") +
                          " must be entered in the Schedule H register against a named patient." };
      }
    }

    // 3. Invoice number + commit deductions
    var now = new Date();
    var invoiceNo = _nextInvoiceNo_(headerSheet, now);
    Object.keys(pending).forEach(function (rn) {
      invSheet.getRange(parseInt(rn, 10), 5).setValue(byRow[rn].qty - pending[rn]);
    });

    // 4. Money, from the INVENTORY. The browser's rate and GST used to be
    //    stored as sent, so a bill could be raised at any price the page was
    //    persuaded to post. The shelf's MRP is the price; a lower one is a
    //    discount, and there is a field for that.
    var gross = 0, totalGst = 0, itemRows = [];
    var tz = Session.getScriptTimeZone();
    for (var m = 0; m < lines.length; m++) {
      var ln = lines[m], rf = ln.ref;
      var qty = ln.qty, mrp = rf.mrp, gstPct = rf.gst;
      var lineTotal = qty * mrp;
      var gstAmt = lineTotal - (lineTotal / (1 + gstPct / 100));
      gross += lineTotal; totalGst += gstAmt;
      var expText = (rf.expiry instanceof Date) ? Utilities.formatDate(rf.expiry, tz, "yyyy-MM") : String(rf.expiry || "");
      itemRows.push([invoiceNo, now, String(payload.patientId || "WALK-IN"),
        rf.brand, rf.generic, rf.batch, expText,
        qty, rf.unit, round2_(mrp), gstPct, round2_(lineTotal - gstAmt),
        round2_(gstAmt), round2_(lineTotal), String(rf.rowNum), rf.schedule]);
    }
    if (discount > gross) discount = gross;
    var net = round2_(gross - discount);

    // 5. Write header + items. Created_By is the signed-in member of staff.
    //    It was Session.getActiveUser(), which for this web app is the
    //    account that DEPLOYED it — so every bill in the clinic's history was
    //    recorded as raised by the owner, whoever stood at the counter.
    var payStatus = (payMode === "CREDIT") ? "PENDING" : "PAID";
    var billedBy = actor.username;
    headerSheet.appendRow([invoiceNo, now, String(payload.billType || "WALK-IN"),
      String(payload.patientId || "WALK-IN"), String(payload.patientName || ""), String(payload.mobile || ""),
      String(payload.age || ""), String(payload.sex || ""), String(payload.address || ""),
      String(payload.doctor || "Self / OTC").trim(), round2_(gross), round2_(totalGst), round2_(discount), net,
      payMode, String(payload.txnId || ""), payStatus, "ACTIVE",
      itemRows.length, billedBy, String(payload.billUuid || ""), "", ""]);
    itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, itemRows.length, itemRows[0].length).setValues(itemRows);
    // Route IP credit bills to the running tab. Discharge owns recognition.
    if (payMode === "CREDIT") {
      try {
        var ipNo = ipc_activeAdmissionByPatient_(String(payload.patientId || ""));
        if (ipNo) billChargeToIp({
          ipNumber: ipNo, source: 'PHARMACY', sourceRef: invoiceNo,
          amount: net, gst: round2_(totalGst),
          description: itemRows.length + ' pharmacy item(s)'
        });
     } catch (e) {}
    }
    if (scheduled.length) {
      try {
        logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
                  'PHARMACY_SCHEDULED_SALE', 'Pharmacy_Invoice', invoiceNo,
                  { doctor: String(payload.doctor || ''), patientId: String(payload.patientId || ''),
                    items: scheduled.map(function (l) { return l.ref.brand + ' ×' + l.qty + ' (' + l.ref.schedule + ')'; }) });
      } catch (e) {}
    }

    // 6. Return confirmed print payload
    return { success: true, invoiceNo: invoiceNo, print: {
      invoiceNo: invoiceNo,
      date: Utilities.formatDate(now, tz, "dd-MMM-yyyy HH:mm"),
      billType: String(payload.billType || "WALK-IN"), patientId: String(payload.patientId || "WALK-IN"),
      patientName: String(payload.patientName || "Walk-in Patient"), mobile: String(payload.mobile || ""),
      age: String(payload.age || ""), sex: String(payload.sex || ""), address: String(payload.address || ""),
      doctor: String(payload.doctor || "Self / OTC"), payMode: payMode,
      txnId: String(payload.txnId || ""), payStatus: payStatus,
      // Expiry is rendered here, once, so the invoice printed from this reply
      // and the one reprinted later from the sheet read identically. A batch
      // expires at the end of its printed month, so "Jun 2028" is what the
      // pack says — not "01-Jun-2028", and certainly not the
      // "Thu Jun 01 2028 00:00:00 GMT+0530" that String() on a Sheets date
      // cell produces.
      items: itemRows.map(function (row) { return { drug: row[3], generic: row[4], batch: row[5],
        expiry: cresc_expiryText_(row[6]), qty: row[7], unit: row[8], mrp: row[9], gst: row[10],
        taxable: row[11], gstAmt: row[12], lineTotal: row[13], schedule: row[15] }; }),
      gross: round2_(gross), totalGst: round2_(totalGst), discount: round2_(discount), net: net } };
  } catch (error) {
    return { success: false, message: _phReason_(error) };
  } finally { if (lock) { try { lock.releaseLock(); } catch (e) {} } }
}

// =====================================================================
// SECTION D — CREDIT SETTLEMENT (close IPD / staff credit bills)
// =====================================================================

function settleCreditBill(payload, sessionToken) {
  // Guard inside the try — see processPharmacyBill above for why.
  var lock = null;
  try {
    var actor = crescRequire_(sessionToken, ['pharmacy.bill', 'accounts.settle']);
    payload = payload || {};
    var mode = String(payload.payMode || "CASH").trim().toUpperCase();
    if (["CASH", "UPI", "CARD", "BANK"].indexOf(mode) === -1) {
      return { success: false, message: "Choose how it was paid: cash, UPI, card or bank transfer." };
    }
    if (typeof acc_isLocked_ === 'function' && typeof acc_period_ === 'function' &&
        acc_isLocked_(acc_period_(new Date()))) {
      return { success: false, message: "This month is locked in the Finance Hub; settlements are frozen." };
    }
    lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { success: false, message: "System busy, please retry." };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.INVOICES);
    if (!sheet) throw new Error("Invoice ledger not found.");
    var data = sheet.getDataRange().getValues();
    var h = data[0].map(function (x) { return String(x || "").trim(); });
    var col = function (name, fallback) { var i = h.indexOf(name); return i === -1 ? fallback : i; };
    var cPay = col("Pay_Status", 16), cStat = col("Status", 17), cMode = col("Pay_Mode", 14),
        cTxn = col("Txn_ID", 15), cAt = col("Settled_At", 21), cBy = col("Settled_By", 22);
    var target = String(payload.invoiceNo || "").trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== target) continue;
      var pay = String(data[i][cPay]).trim().toUpperCase();
      var st = String(data[i][cStat]).trim().toUpperCase();
      if (st === "CANCELLED") return { success: false, message: "Invoice " + target + " was cancelled." };
      if (pay === "PAID") return { success: false, message: "Invoice " + target + " is already settled." };
      if (pay === "IP_SETTLED") return { success: false, message: "Invoice " + target + " was settled on the patient's discharge bill." };
      var rowNum = i + 1;
      var who = actor.displayName || actor.username;
      sheet.getRange(rowNum, cMode + 1).setValue(mode);
      sheet.getRange(rowNum, cTxn + 1).setValue(String(payload.txnId || ""));
      sheet.getRange(rowNum, cPay + 1).setValue("PAID");
      sheet.getRange(rowNum, cAt + 1).setValue(new Date());
      // Who took the money: the session, not Session.getActiveUser(), which
      // for this web app is always the account that deployed it.
      sheet.getRange(rowNum, cBy + 1).setValue(who);

      // A credit bill routed to an admission's running tab and then paid here
      // is taken off the tab, or the discharge bill charges it again.
      var offTab = false;
      try {
        if (typeof ipc_markChargePaidAtCounter_ === 'function') {
          offTab = ipc_markChargePaidAtCounter_('PHARMACY', target, mode, who);
        }
      } catch (e) {}
      try {
        logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
                  'PHARMACY_CREDIT_SETTLED', 'Pharmacy_Invoice', target,
                  { mode: mode, txnId: String(payload.txnId || ''), removedFromIpTab: offTab });
      } catch (e) {}
      SpreadsheetApp.flush();
      return { success: true, invoiceNo: target,
               message: "Invoice " + target + " settled by " + mode + "." +
                        (offTab ? " It has been taken off the patient's IP running bill." : "") };
    }
    return { success: false, message: "Invoice " + target + " not found." };
  } catch (error) {
    return { success: false, message: _phReason_(error) };
  } finally { if (lock) { try { lock.releaseLock(); } catch (e) {} } }
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
    sheet.appendRow(PH_ITEM_HEADERS);
    sheet.getRange(1, 1, 1, PH_ITEM_HEADERS.length).setFontWeight("bold").setBackground("#d9ead3");
    sheet.setFrozenRows(1);
    return sheet;
  }
  // The Schedule column (P) is new: lines are written sixteen wide, and a
  // column with no header would be invisible to every reader.
  _phEnsureWidth_(sheet, PH_ITEM_HEADERS.length);
  var width = Math.max(sheet.getLastColumn(), PH_ITEM_HEADERS.length);
  var head = sheet.getRange(1, 1, 1, width).getValues()[0];
  for (var c = 0; c < PH_ITEM_HEADERS.length; c++) {
    if (!String(head[c] || "").trim()) {
      sheet.getRange(1, c + 1).setValue(PH_ITEM_HEADERS[c]).setFontWeight("bold");
    }
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

/**
 * A caught error, said in a sentence the person at the counter can act on.
 *
 * "FORBIDDEN: your session has expired." is already the right words; the
 * prefix is not, and it is the prefix a pharmacist reads first. Everything
 * else is passed through unchanged — a stock shortfall or a missing sheet
 * already explains itself.
 */
function _phReason_(error) {
  // The same explainer every guarded endpoint now uses, kept under its
  // pharmacy name so this file's existing call sites read unchanged.
  return (typeof cresc_reason_ === 'function')
    ? cresc_reason_(error)
    : String((error && error.message) || error || '').replace(/^FORBIDDEN:\s*/, '');
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
    crescRequire_(sessionToken, ['pharmacy.bill', 'pharmacy.dispense']);

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
    crescRequire_(sessionToken, ['pharmacy.bill', 'pharmacy.dispense']);
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
// =====================================================================
// SECTION G — THE SCHEDULE H / H1 / X REGISTER
// ---------------------------------------------------------------------
// One row per scheduled medicine sold, for the period asked: S.No, date of
// issue, patient, prescribing doctor, the drug, the quantity — and a blank
// column the pharmacist signs on paper. Printed from Pharmacy → Schedule H
// Register; the same rows export to CSV.
//
// WHERE THE SCHEDULE COMES FROM. Every line sold since the Schedule column
// existed carries its own schedule, snapshotted at the sale. Lines sold
// before that carry none, and are classified by the medicine's CURRENT
// schedule on the inventory — marked `inferred`, so the printed register
// can say which is which. Cancelled bills are left out; a line with units
// returned says how many.
// =====================================================================

/**
 * FRONTEND ENTRY. The register rows for a date range.
 *
 * @param {{from:string, to:string, schedules?:Array<string>}} filter
 *        from/to as yyyy-MM-dd (inclusive); schedules default H, H1, X
 * @param {string} sessionToken
 */
function getScheduleDrugRegister(filter, sessionToken) {
  try {
    crescRequire_(sessionToken, 'pharmacy.register');
    filter = filter || {};
    var from = String(filter.from || '').trim(), to = String(filter.to || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) {
      return { success: false, message: 'Choose the first and last day of the period.' };
    }
    if (from > to) return { success: false, message: 'The period starts after it ends.' };
    var want = {};
    (filter.schedules && filter.schedules.length ? filter.schedules : Object.keys(PH_SCHEDULES))
      .forEach(function (k) { var sc = _phSchedule_(k); if (sc) want[sc] = true; });
    if (!Object.keys(want).length) return { success: false, message: 'Tick at least one schedule.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = Session.getScriptTimeZone();
    var itemsSheet = ss.getSheetByName(PH_SHEETS.INVOICE_ITEMS);
    var headSheet = ss.getSheetByName(PH_SHEETS.INVOICES);
    if (!itemsSheet || itemsSheet.getLastRow() < 2 || !headSheet) {
      return { success: true, rows: [], summary: {}, message: '' };
    }

    // The current classification, for lines sold before the snapshot existed.
    var current = {};
    var inv = ss.getSheetByName(PH_SHEETS.INVENTORY);
    if (inv && inv.getLastRow() > 1) {
      current = _phScheduleByBrand_(inv.getDataRange().getValues(), _phScheduleCol_(inv, false));
    }

    // Invoice headers: who, prescribed by whom, and whether it still stands.
    var hd = headSheet.getDataRange().getValues();
    var hh = hd[0].map(function (x) { return String(x || '').trim(); });
    var hc = function (n, f) { var i = hh.indexOf(n); return i === -1 ? f : i; };
    var H = { no: hc('Invoice_No', 0), pid: hc('Patient_ID', 3), name: hc('Patient_Name', 4),
              mob: hc('Mobile', 5), age: hc('Age', 6), sex: hc('Sex', 7), addr: hc('Address', 8),
              doc: hc('Doctor', 9), status: hc('Status', 17), by: hc('Created_By', 19) };
    var head = {};
    for (var i = 1; i < hd.length; i++) {
      var no = String(hd[i][H.no] || '').trim().toUpperCase();
      if (no) head[no] = hd[i];
    }

    // Units returned against each invoice line (brand|batch).
    var returned = {};
    try {
      var ri = ss.getSheetByName((typeof PH_RET !== 'undefined') ? PH_RET.ITEMS : 'Pharmacy_Return_Items');
      if (ri && ri.getLastRow() > 1) {
        ri.getDataRange().getValues().slice(1).forEach(function (r) {
          var k = [String(r[2] || '').trim(), String(r[3] || '').trim(), String(r[5] || '').trim()].join('|').toUpperCase();
          returned[k] = (returned[k] || 0) + (parseFloat(r[7]) || 0);
        });
      }
    } catch (e) {}

    var idata = itemsSheet.getDataRange().getValues();
    var ih = idata[0].map(function (x) { return String(x || '').trim(); });
    var cSch = ih.indexOf(PH_SCHEDULE_HEADER);
    var rows = [], summary = {};
    for (var j = 1; j < idata.length; j++) {
      var it = idata[j];
      var invNo = String(it[0] || '').trim();
      if (!invNo) continue;
      var when = cresc_parseDate_(it[1]);
      if (!when) continue;
      var day = Utilities.formatDate(when, tz, 'yyyy-MM-dd');
      if (day < from || day > to) continue;

      var brand = String(it[3] || '').trim();
      var own = (cSch >= 0) ? _phSchedule_(it[cSch]) : '';
      var sc = own || current[brand.toLowerCase()] || '';
      if (!sc || !want[sc]) continue;

      var h = head[invNo.toUpperCase()];
      if (h && String(h[H.status] || '').trim().toUpperCase() === 'CANCELLED') continue;

      var qty = parseFloat(it[7]) || 0;
      var back = returned[[invNo, brand, String(it[5] || '').trim()].join('|').toUpperCase()] || 0;
      var ageSex = h ? [String(h[H.age] || '').trim(), String(h[H.sex] || '').trim()].filter(String).join(' / ') : '';
      rows.push({
        sortMs: when.getTime(),
        date: Utilities.formatDate(when, tz, 'dd-MMM-yyyy'),
        time: Utilities.formatDate(when, tz, 'hh:mm a'),
        invoiceNo: invNo,
        patientName: h ? String(h[H.name] || '').trim() : '',
        patientId: h ? String(h[H.pid] || '').trim() : String(it[2] || '').trim(),
        ageSex: ageSex,
        address: h ? String(h[H.addr] || '').trim() : '',
        doctor: h ? String(h[H.doc] || '').trim() : '',
        drug: brand,
        generic: String(it[4] || '').trim(),
        batch: String(it[5] || '').trim(),
        qty: qty,
        unit: String(it[8] || '').trim(),
        returned: back,
        schedule: sc,
        inferred: !own,
        doctorMissing: !h || !_phNamesDoctor_(h[H.doc]),
        issuedBy: h ? String(h[H.by] || '').trim() : ''
      });
      summary[sc] = (summary[sc] || 0) + 1;
    }
    rows.sort(function (a, b) { return a.sortMs - b.sortMs; });
    rows.forEach(function (r, n) { r.sNo = n + 1; delete r.sortMs; });

    try {
      var actor = crescActor_(sessionToken);
      logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
                'SCHEDULE_REGISTER_READ', 'Pharmacy', from + '..' + to,
                { rows: rows.length, schedules: Object.keys(want) });
    } catch (e) {}

    return {
      success: true, rows: rows, summary: summary, from: from, to: to,
      schedules: Object.keys(want),
      withoutDoctor: rows.filter(function (r) { return r.doctorMissing; }).length,
      inferred: rows.filter(function (r) { return r.inferred; }).length,
      message: rows.length ? '' : 'No scheduled medicine was sold in this period.'
    };
  } catch (err) {
    return { success: false, rows: [], message: _phReason_(err) };
  }
}
