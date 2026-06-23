// =====================================================================
// CRESCENTIA HEALTHTECH — PHARMACY RETURNS / REFUND / GST ENGINE
// Add as a NEW .gs file in the SAME Apps Script project as the pharmacy
// backend. It reuses the existing globals: PH_SHEETS, round2_, _fmtDate_,
// _dateKey_  (do not redefine those here).
//
// DESIGN PRINCIPLES
//  1. The original invoice is NEVER financially edited. Only its Status
//     column flips: ACTIVE -> PARTIAL-RETURN -> RETURNED (or CANCELLED).
//     Gross / GST / Net stay as the original sale forever (audit trail).
//  2. Every return is its own record in Pharmacy_Returns (+_Items) that
//     references the original invoice and restocks the exact Brand|Batch.
//  3. All money is recomputed server-side from the STORED invoice lines.
//     The browser only says "which batch, how many" — never the price.
//  4. GST reverses on the returned gross (same basis the sale recorded it),
//     so a full return nets GST and revenue back to zero.
//  5. Refund respects any discount on the original bill (proportional),
//     so refunds can never exceed what was actually collected.
// =====================================================================

var PH_RET = { HEADER: "Pharmacy_Returns", ITEMS: "Pharmacy_Return_Items" };

var PH_RETURN_HEADERS = ["Return_No","Timestamp","Invoice_No","Patient_ID","Patient_Name",
  "Return_Type","Returned_Gross","Returned_GST","Refund_Amount","Refund_Mode","Refund_Ref",
  "Orig_PayMode","Orig_PayStatus","Reason","Status","Item_Count","Created_By","Return_UUID"];

var PH_RETURN_ITEM_HEADERS = ["Return_No","Timestamp","Invoice_No","Brand","Generic","Batch",
  "Expiry","Qty","Unit","MRP","GST_Pct","Taxable","GST_Amt","Line_Total","Refund_Line","Note"];


// ---------------------------------------------------------------------
// LOOKUP — load an invoice with per-line returnable quantities
// ---------------------------------------------------------------------
function getInvoiceForReturn(query) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hSheet = ss.getSheetByName(PH_SHEETS.INVOICES);
    var iSheet = ss.getSheetByName(PH_SHEETS.INVOICE_ITEMS);
    if (!hSheet || hSheet.getLastRow() <= 1) return { success: false, message: "No invoices found." };

    var target = String(query || "").trim().toUpperCase();
    if (!target) return { success: false, message: "Enter an invoice number." };

    var hdata = hSheet.getDataRange().getValues();
    var hrow = null;
    for (var i = 1; i < hdata.length; i++) {
      if (String(hdata[i][0]).trim().toUpperCase() === target) { hrow = hdata[i]; break; }
    }
    if (!hrow) return { success: false, message: "Invoice " + query + " not found." };

    var header = {
      invoiceNo: String(hrow[0]), date: _fmtDate_(hrow[1]), billType: String(hrow[2] || ""),
      patientId: String(hrow[3] || ""), patientName: String(hrow[4] || ""), mobile: String(hrow[5] || ""),
      doctor: String(hrow[9] || ""), gross: parseFloat(hrow[10]) || 0, totalGst: parseFloat(hrow[11]) || 0,
      discount: parseFloat(hrow[12]) || 0, net: parseFloat(hrow[13]) || 0,
      payMode: String(hrow[14] || ""), payStatus: String(hrow[16] || ""), status: String(hrow[17] || "ACTIVE")
    };
    header.discountFactor = header.gross > 0 ? (header.net / header.gross) : 1;

    var billed = {}, order = [];
    if (iSheet && iSheet.getLastRow() > 1) {
      var idata = iSheet.getDataRange().getValues();
      for (var j = 1; j < idata.length; j++) {
        if (String(idata[j][0]).trim().toUpperCase() !== target) continue;
        var key = _returnAggKey_(idata[j][3], idata[j][5]);
        if (!billed[key]) {
          billed[key] = { brand: String(idata[j][3] || ""), generic: String(idata[j][4] || ""),
            batch: String(idata[j][5] || ""), expiry: String(idata[j][6] || ""), unit: String(idata[j][8] || ""),
            mrp: parseFloat(idata[j][9]) || 0, gst: parseFloat(idata[j][10]) || 0, billedQty: 0 };
          order.push(key);
        }
        billed[key].billedQty += parseFloat(idata[j][7]) || 0;
      }
    }

    var prior = _returnedQtyByKey_(ss, target);
    var items = order.map(function (k) {
      var b = billed[k], rq = prior[k] || 0;
      return { brand: b.brand, generic: b.generic, batch: b.batch, expiry: b.expiry, unit: b.unit,
        mrp: b.mrp, gst: b.gst, billedQty: b.billedQty, returnedQty: rq,
        returnableQty: Math.max(0, b.billedQty - rq), lineTotal: round2_(b.billedQty * b.mrp) };
    });
    return { success: true, header: header, items: items };
  } catch (e) {
    return { success: false, message: "Lookup failed: " + e.toString() };
  }
}


// ---------------------------------------------------------------------
// CORE — process a return / partial or full cancellation + refund
// ---------------------------------------------------------------------
function processPharmacyReturn(payload) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: "System busy, please retry." };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hSheet = ss.getSheetByName(PH_SHEETS.INVOICES);
    var iSheet = ss.getSheetByName(PH_SHEETS.INVOICE_ITEMS);
    var invSheet = ss.getSheetByName(PH_SHEETS.INVENTORY);
    if (!hSheet || !invSheet) throw new Error("Required sheets are missing.");

    var rHeader = _ensureReturnsHeaderSheet_(ss);
    var rItems = _ensureReturnItemsSheet_(ss);

    // 0. Idempotency — a retried click never double-restocks or double-refunds.
    if (payload.returnUuid) {
      var dup = _findReturnByUuid_(rHeader, payload.returnUuid);
      if (dup) return { success: true, returnNo: dup, duplicate: true, message: "This return was already processed." };
    }

    var target = String(payload.invoiceNo || "").trim().toUpperCase();
    if (!target) throw new Error("Missing invoice number.");

    var hdata = hSheet.getDataRange().getValues();
    var hrowNum = -1, hrow = null;
    for (var i = 1; i < hdata.length; i++) {
      if (String(hdata[i][0]).trim().toUpperCase() === target) { hrowNum = i + 1; hrow = hdata[i]; break; }
    }
    if (!hrow) throw new Error("Invoice " + payload.invoiceNo + " not found.");

    var curStatus = String(hrow[17] || "ACTIVE").toUpperCase();
    if (curStatus === "RETURNED" || curStatus === "CANCELLED")
      throw new Error("Invoice " + payload.invoiceNo + " is already fully returned.");

    var gross = parseFloat(hrow[10]) || 0, net = parseFloat(hrow[13]) || 0;
    var factor = gross > 0 ? (net / gross) : 1;
    var origPayMode = String(hrow[14] || "CASH").toUpperCase();
    var origPayStatus = String(hrow[16] || "PAID").toUpperCase();

    // 1. Authoritative billed quantities + prices (from stored invoice items)
    var billed = {};
    if (iSheet && iSheet.getLastRow() > 1) {
      var idata = iSheet.getDataRange().getValues();
      for (var j = 1; j < idata.length; j++) {
        if (String(idata[j][0]).trim().toUpperCase() !== target) continue;
        var bk = _returnAggKey_(idata[j][3], idata[j][5]);
        if (!billed[bk]) billed[bk] = { brand: String(idata[j][3] || ""), generic: String(idata[j][4] || ""),
          batch: String(idata[j][5] || ""), expiry: String(idata[j][6] || ""), unit: String(idata[j][8] || ""),
          mrp: parseFloat(idata[j][9]) || 0, gst: parseFloat(idata[j][10]) || 0, billedQty: 0 };
        billed[bk].billedQty += parseFloat(idata[j][7]) || 0;
      }
    }
    var already = _returnedQtyByKey_(ss, target);

    // 2. Validate + compute every line against billed-minus-already-returned
    var lines = payload.lines || [];
    var clean = [], returnedGross = 0, returnedGst = 0, totalUnits = 0;
    for (var L = 0; L < lines.length; L++) {
      var qty = parseFloat(lines[L].qty) || 0;
      if (qty <= 0) continue;
      var key = _returnAggKey_(lines[L].brand, lines[L].batch);
      var b = billed[key];
      if (!b) throw new Error("Line not on this invoice: " + lines[L].brand + " (Batch " + lines[L].batch + ").");
      var cap = b.billedQty - (already[key] || 0);
      if (qty > cap) throw new Error("Cannot return " + qty + " of " + b.brand + " (Batch " + b.batch + "). Returnable: " + cap + ".");
      var lineTotal = round2_(qty * b.mrp);
      var gstAmt = b.gst > 0 ? round2_(lineTotal - (lineTotal / (1 + b.gst / 100))) : 0;
      var refundLine = round2_(lineTotal * factor);
      returnedGross += lineTotal; returnedGst += gstAmt; totalUnits += qty;
      clean.push({ brand: b.brand, generic: b.generic, batch: b.batch, expiry: b.expiry, unit: b.unit,
        mrp: b.mrp, gst: b.gst, qty: qty, lineTotal: lineTotal, gstAmt: gstAmt, refundLine: refundLine });
    }
    if (!clean.length) throw new Error("Nothing to return — all quantities were zero.");
    returnedGross = round2_(returnedGross);
    returnedGst = round2_(returnedGst);
    var refund = round2_(clean.reduce(function (s, c) { return s + c.refundLine; }, 0));

    // 3. Restock each returned batch (same lock). Never deletes — only adds back.
    var invData = invSheet.getDataRange().getValues();
    var invIdx = {};
    for (var r = 1; r < invData.length; r++) invIdx[_returnAggKey_(invData[r][1], invData[r][6])] = { rowNum: r + 1, qty: parseInt(invData[r][4], 10) || 0 };
    clean.forEach(function (c) {
      var k = _returnAggKey_(c.brand, c.batch), ref = invIdx[k];
      if (ref) { invSheet.getRange(ref.rowNum, 5).setValue(ref.qty + c.qty); ref.qty += c.qty; }
      else invSheet.appendRow([new Date(), c.brand, c.generic, "", c.qty, c.unit, c.batch, c.expiry, "RETURN", 0, c.mrp, c.gst, "", ""]); // batch row gone — re-add so stock isn't lost
    });

    // 4. New invoice status (financials untouched)
    var totalBilledUnits = 0, returnedBefore = 0;
    Object.keys(billed).forEach(function (k) { totalBilledUnits += billed[k].billedQty; returnedBefore += (already[k] || 0); });
    var fullNow = (returnedBefore + totalUnits) >= totalBilledUnits;
    var reqType = String(payload.returnType || "").toUpperCase();
    var newStatus = fullNow ? ((reqType === "CANCEL" && returnedBefore <= 0) ? "CANCELLED" : "RETURNED") : "PARTIAL-RETURN";

    // 5. Refund mode. An UNSETTLED credit bill is reduced, not paid out in cash.
    var refundMode = String(payload.refundMode || "").toUpperCase();
    var creditUnsettled = (origPayStatus === "PENDING");
    if (creditUnsettled) refundMode = "CREDIT-ADJUST";
    else if (!refundMode) refundMode = origPayMode || "CASH";

    var by = ""; try { by = Session.getActiveUser().getEmail() || ""; } catch (e) {}
    var now = new Date();
    var returnNo = _nextReturnNo_(rHeader, now);

    // 6. Write return items + header
    var itemRows = clean.map(function (c) {
      return [returnNo, now, String(payload.invoiceNo), c.brand, c.generic, c.batch, c.expiry,
        c.qty, c.unit, round2_(c.mrp), c.gst, round2_(c.lineTotal - c.gstAmt), c.gstAmt, c.lineTotal, c.refundLine, ""];
    });
    rItems.getRange(rItems.getLastRow() + 1, 1, itemRows.length, itemRows[0].length).setValues(itemRows);

    var headerType = fullNow ? (newStatus === "CANCELLED" ? "FULL-CANCEL" : "FULL") : "PARTIAL";
    rHeader.appendRow([returnNo, now, String(payload.invoiceNo), String(hrow[3] || ""), String(hrow[4] || ""),
      headerType, returnedGross, returnedGst, refund, refundMode, String(payload.refundRef || ""),
      origPayMode, origPayStatus, String(payload.reason || ""), "DONE", clean.length, by, String(payload.returnUuid || "")]);

    hSheet.getRange(hrowNum, 18).setValue(newStatus); // Status column only

    // 7. Mirror sale's IP posting with a negative charge for unsettled credit
    if (creditUnsettled) {
      try {
        var ipNo = (typeof ipc_activeAdmissionByPatient_ === "function") ? ipc_activeAdmissionByPatient_(String(hrow[3] || "")) : null;
        if (ipNo && typeof billChargeToIp === "function") billChargeToIp({
          ipNumber: ipNo, source: 'PHARMACY-RETURN', sourceRef: returnNo,
          amount: -refund, gst: -returnedGst,
          description: clean.length + ' item(s) returned (inv ' + payload.invoiceNo + ')', user: by
        });
      } catch (e) {}
    }

    return { success: true, returnNo: returnNo, newStatus: newStatus, refund: refund,
      returnedGross: returnedGross, returnedGst: returnedGst, refundMode: refundMode,
      print: {
        returnNo: returnNo, invoiceNo: String(payload.invoiceNo),
        date: Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM-yyyy HH:mm"),
        patientId: String(hrow[3] || ""), patientName: String(hrow[4] || "Walk-in Patient"),
        reason: String(payload.reason || ""), refundMode: refundMode, refundRef: String(payload.refundRef || ""),
        newStatus: newStatus, returnedGross: returnedGross, returnedGst: returnedGst, refund: refund,
        items: clean.map(function (c) { return { drug: c.brand, generic: c.generic, batch: c.batch, expiry: c.expiry,
          qty: c.qty, unit: c.unit, mrp: c.mrp, gst: c.gst, gstAmt: c.gstAmt, lineTotal: c.lineTotal, refundLine: c.refundLine }; })
      } };
  } catch (error) {
    return { success: false, message: error.message || String(error) };
  } finally { lock.releaseLock(); }
}


// ---------------------------------------------------------------------
// FINANCE — recompute sales, returns, refunds and net GST for a range
// ---------------------------------------------------------------------
function getPharmacyFinanceSummary(fromKey, toKey) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var from = String(fromKey || "0000-00-00"), to = String(toKey || "9999-99-99");

    var sales = { count: 0, gross: 0, discount: 0, net: 0, gst: 0, byMode: {} };
    var creditOutstanding = 0, creditCount = 0;
    var refByInv = _refundByInvoice_(ss);

    var hSheet = ss.getSheetByName(PH_SHEETS.INVOICES);
    if (hSheet && hSheet.getLastRow() > 1) {
      var h = hSheet.getDataRange().getValues();
      for (var i = 1; i < h.length; i++) {
        var dk = _dateKey_(h[i][1]); if (dk < from || dk > to) continue;
        var g = parseFloat(h[i][10]) || 0, disc = parseFloat(h[i][12]) || 0, n = parseFloat(h[i][13]) || 0, gst = parseFloat(h[i][11]) || 0;
        var mode = String(h[i][14] || "CASH").toUpperCase();
        sales.count++; sales.gross += g; sales.discount += disc; sales.net += n; sales.gst += gst;
        sales.byMode[mode] = (sales.byMode[mode] || 0) + n;
        var ps = String(h[i][16] || "").toUpperCase(), st = String(h[i][17] || "").toUpperCase();
        if (ps === "PENDING" && (st === "ACTIVE" || st === "PARTIAL-RETURN")) {
          var out = n - (refByInv[String(h[i][0]).trim().toUpperCase()] || 0);
          if (out > 0) { creditOutstanding += out; creditCount++; }
        }
      }
    }

    var ret = { count: 0, gross: 0, gst: 0, refund: 0 };
    var rSheet = ss.getSheetByName(PH_RET.HEADER);
    if (rSheet && rSheet.getLastRow() > 1) {
      var rd = rSheet.getDataRange().getValues();
      for (var k = 1; k < rd.length; k++) {
        var rdk = _dateKey_(rd[k][1]); if (rdk < from || rdk > to) continue;
        ret.count++; ret.gross += parseFloat(rd[k][6]) || 0; ret.gst += parseFloat(rd[k][7]) || 0; ret.refund += parseFloat(rd[k][8]) || 0;
      }
    }

    return { success: true,
      range: { from: from, to: to },
      sales: { count: sales.count, gross: round2_(sales.gross), discount: round2_(sales.discount), net: round2_(sales.net), gst: round2_(sales.gst), byMode: sales.byMode },
      returns: { count: ret.count, gross: round2_(ret.gross), gst: round2_(ret.gst), refund: round2_(ret.refund) },
      netRealised: round2_(sales.net - ret.refund),
      netGst: round2_(sales.gst - ret.gst),
      creditOutstanding: round2_(creditOutstanding), creditCount: creditCount };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}


// ---------------------------------------------------------------------
// REPLACE the existing getPendingCreditBills() with this version.
// It now counts partially-returned credit bills and shows the amount
// still collectible (net of refunds) instead of the original net.
// ---------------------------------------------------------------------
function getPendingCreditBills() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PH_SHEETS.INVOICES);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    var refByInv = (typeof _refundByInvoice_ === "function") ? _refundByInvoice_(ss) : {};
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][16]).trim().toUpperCase() !== "PENDING") continue;
      var st = String(data[i][17]).trim().toUpperCase();
      if (st !== "ACTIVE" && st !== "PARTIAL-RETURN") continue;
      var net = parseFloat(data[i][13]) || 0;
      var outstanding = round2_(net - (refByInv[String(data[i][0]).trim().toUpperCase()] || 0));
      if (outstanding <= 0) continue;
      out.push({ invoiceNo: String(data[i][0]), date: _fmtDate_(data[i][1]),
        billType: String(data[i][2] || ""), patientId: String(data[i][3] || ""),
        patientName: String(data[i][4] || ""), net: outstanding, original: net });
    }
    out.reverse();
    return { success: true, data: out };
  } catch (error) {
    return { success: false, message: "Could not load credits: " + error.toString() };
  }
}


// ---------------------------------------------------------------------
// HELPERS (new — safe to add; do not duplicate existing ones)
// ---------------------------------------------------------------------
function _returnAggKey_(brand, batch) {
  return (String(brand || "").trim() + "|" + String(batch || "").trim()).toUpperCase();
}

function _returnedQtyByKey_(ss, invoiceUpper) {
  var out = {};
  var ri = ss.getSheetByName(PH_RET.ITEMS);
  if (!ri || ri.getLastRow() <= 1) return out;
  var d = ri.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][2]).trim().toUpperCase() !== invoiceUpper) continue;
    var key = _returnAggKey_(d[i][3], d[i][5]);
    out[key] = (out[key] || 0) + (parseFloat(d[i][7]) || 0);
  }
  return out;
}

function _refundByInvoice_(ss) {
  var out = {};
  var rh = ss.getSheetByName(PH_RET.HEADER);
  if (!rh || rh.getLastRow() <= 1) return out;
  var d = rh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    var k = String(d[i][2]).trim().toUpperCase();
    out[k] = (out[k] || 0) + (parseFloat(d[i][8]) || 0);
  }
  return out;
}

function _findReturnByUuid_(sheet, uuid) {
  if (!uuid || sheet.getLastRow() <= 1) return null;
  var d = sheet.getDataRange().getValues();
  var col = d[0].indexOf("Return_UUID");
  if (col < 0) return null;
  for (var i = 1; i < d.length; i++) if (String(d[i][col]).trim() === String(uuid).trim()) return String(d[i][0]);
  return null;
}

function _nextReturnNo_(sheet, now) {
  var yymm = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyMM");
  var maxSeq = 1000;
  if (sheet.getLastRow() > 1) {
    var nos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < nos.length; i++) {
      var m = String(nos[i][0]).match(/-(\d+)$/);
      if (m) { var s = parseInt(m[1], 10); if (s > maxSeq) maxSeq = s; }
    }
  }
  return "RN" + yymm + "-" + (maxSeq + 1);
}

function _ensureReturnsHeaderSheet_(ss) {
  var sheet = ss.getSheetByName(PH_RET.HEADER);
  if (!sheet) {
    sheet = ss.insertSheet(PH_RET.HEADER);
    sheet.appendRow(PH_RETURN_HEADERS);
    sheet.getRange(1, 1, 1, PH_RETURN_HEADERS.length).setFontWeight("bold").setBackground("#fce5cd");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _ensureReturnItemsSheet_(ss) {
  var sheet = ss.getSheetByName(PH_RET.ITEMS);
  if (!sheet) {
    sheet = ss.insertSheet(PH_RET.ITEMS);
    sheet.appendRow(PH_RETURN_ITEM_HEADERS);
    sheet.getRange(1, 1, 1, PH_RETURN_ITEM_HEADERS.length).setFontWeight("bold").setBackground("#fce5cd");
    sheet.setFrozenRows(1);
  }
  return sheet;
}