// =====================================================================
// CRESCENTIA HEALTHTECH — PHARMACY DASHBOARD + BILL SEARCH + REPRINT
// New .gs file for the SAME project. Operational (counter) view only —
// the financial/tax recalculation lives in Accounts.
// Reuses globals: PH_SHEETS, round2_, _fmtDate_, _dateKey_.
// =====================================================================

// Per-invoice refund info (amount, returned gross, return count) from the returns ledger.
function _refundInfoByInvoice_(ss) {
  var out = {};
  var retName = (typeof PH_RET !== "undefined") ? PH_RET.HEADER : "Pharmacy_Returns";
  var rh = ss.getSheetByName(retName);
  if (!rh || rh.getLastRow() <= 1) return out;
  var d = rh.getDataRange().getValues();
  // Pharmacy_Returns: 2=Invoice_No, 6=Returned_Gross, 7=Returned_GST, 8=Refund_Amount
  for (var i = 1; i < d.length; i++) {
    var k = String(d[i][2]).trim().toUpperCase();
    if (!out[k]) out[k] = { refund: 0, gross: 0, count: 0 };
    out[k].refund += parseFloat(d[i][8]) || 0;
    out[k].gross += parseFloat(d[i][6]) || 0;
    out[k].count += 1;
  }
  return out;
}

// Operational snapshot for a date range: counts, mode tallies, per-bill refund, bill list.
function getPharmacyDashboard(fromKey, toKey) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var from = String(fromKey || "0000-00-00"), to = String(toKey || "9999-99-99");
    var tz = Session.getScriptTimeZone();
    var t = { bills: 0, refunds: 0, cash: 0, upi: 0, card: 0, credit: 0, net: 0, refundAmt: 0 };
    var bills = [];

    var refInfo = _refundInfoByInvoice_(ss);

    var h = ss.getSheetByName(PH_SHEETS.INVOICES);
    if (h && h.getLastRow() > 1) {
      var d = h.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        var dk = _dateKey_(d[i][1]); if (dk < from || dk > to) continue;
        var net = parseFloat(d[i][13]) || 0;
        var mode = String(d[i][14] || "CASH").toUpperCase();
        var invU = String(d[i][0] || "").trim().toUpperCase();
        var rf = refInfo[invU] || { refund: 0, gross: 0, count: 0 };
        t.bills++; t.net += net; t.refundAmt += rf.refund;
        if (mode === "CASH") t.cash += net;
        else if (mode === "UPI") t.upi += net;
        else if (mode === "CARD") t.card += net;
        else if (mode === "CREDIT") t.credit += net;
        var ts = d[i][1];
        bills.push({ invoiceNo: String(d[i][0]), date: _fmtDate_(ts),
          time: (ts instanceof Date) ? Utilities.formatDate(ts, tz, "HH:mm") : "",
          sortTs: (ts instanceof Date) ? ts.getTime() : 0,
          patientId: String(d[i][3] || ""), patientName: String(d[i][4] || ""), mobile: String(d[i][5] || ""),
          payMode: mode, payStatus: String(d[i][16] || ""), status: String(d[i][17] || "ACTIVE"),
          net: round2_(net),
          refunded: round2_(rf.refund), returnedGross: round2_(rf.gross), returnCount: rf.count,
          netRealised: round2_(net - rf.refund) });
      }
    }

    // refunds COUNT within range (by return date), independent of which bills are in range
    var retName = (typeof PH_RET !== "undefined") ? PH_RET.HEADER : "Pharmacy_Returns";
    var r = ss.getSheetByName(retName);
    if (r && r.getLastRow() > 1) {
      var rd = r.getDataRange().getValues();
      for (var k = 1; k < rd.length; k++) {
        var rdk = _dateKey_(rd[k][1]); if (rdk < from || rdk > to) continue;
        t.refunds++;
      }
    }

    bills.sort(function (a, b) { return b.sortTs - a.sortTs; });
    return { success: true, range: { from: from, to: to },
      totals: { bills: t.bills, refunds: t.refunds, cash: round2_(t.cash), upi: round2_(t.upi),
        card: round2_(t.card), credit: round2_(t.credit), net: round2_(t.net),
        refundAmt: round2_(t.refundAmt), netRealised: round2_(t.net - t.refundAmt) },
      bills: bills };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// Search past bills by Patient ID / Name / Mobile / Invoice No.
function searchPharmacyInvoices(query) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var q = String(query || "").trim().toLowerCase();
    if (!q) return { success: false, message: "Enter an ID, name, mobile, or invoice number." };
    var tz = Session.getScriptTimeZone();
    var refInfo = _refundInfoByInvoice_(ss);
    var h = ss.getSheetByName(PH_SHEETS.INVOICES);
    if (!h || h.getLastRow() <= 1) return { success: true, data: [] };
    var d = h.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < d.length; i++) {
      var inv = String(d[i][0] || "").toLowerCase();
      var pid = String(d[i][3] || "").toLowerCase();
      var name = String(d[i][4] || "").toLowerCase();
      var mob = String(d[i][5] || "").toLowerCase();
      if (inv.indexOf(q) > -1 || pid.indexOf(q) > -1 || name.indexOf(q) > -1 || mob.indexOf(q) > -1) {
        var ts = d[i][1];
        var rf = refInfo[String(d[i][0] || "").trim().toUpperCase()] || { refund: 0, count: 0 };
        out.push({ invoiceNo: String(d[i][0]), date: _fmtDate_(ts),
          time: (ts instanceof Date) ? Utilities.formatDate(ts, tz, "HH:mm") : "",
          sortTs: (ts instanceof Date) ? ts.getTime() : 0,
          patientId: String(d[i][3] || ""), patientName: String(d[i][4] || ""), mobile: String(d[i][5] || ""),
          payMode: String(d[i][14] || ""), payStatus: String(d[i][16] || ""), status: String(d[i][17] || "ACTIVE"),
          net: round2_(parseFloat(d[i][13]) || 0),
          refunded: round2_(rf.refund), returnCount: rf.count });
      }
    }
    out.sort(function (a, b) { return b.sortTs - a.sortTs; });
    return { success: true, data: out.slice(0, 200) };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// Rebuild the print payload for ANY past invoice (same shape buildPrintInvoiceRaw expects).
function getInvoiceForPrint(invoiceNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var target = String(invoiceNo || "").trim().toUpperCase();
    if (!target) return { success: false, message: "No invoice number." };
    var h = ss.getSheetByName(PH_SHEETS.INVOICES);
    var it = ss.getSheetByName(PH_SHEETS.INVOICE_ITEMS);
    if (!h) return { success: false, message: "Invoice ledger missing." };
    var hd = h.getDataRange().getValues();
    var row = null;
    for (var i = 1; i < hd.length; i++) if (String(hd[i][0]).trim().toUpperCase() === target) { row = hd[i]; break; }
    if (!row) return { success: false, message: "Invoice " + invoiceNo + " not found." };
    var tz = Session.getScriptTimeZone();
    var items = [];
    if (it && it.getLastRow() > 1) {
      var id = it.getDataRange().getValues();
      for (var j = 1; j < id.length; j++) {
        if (String(id[j][0]).trim().toUpperCase() !== target) continue;
        items.push({ drug: String(id[j][3] || ""), generic: String(id[j][4] || ""), batch: String(id[j][5] || ""),
          expiry: String(id[j][6] || ""), qty: parseFloat(id[j][7]) || 0, unit: String(id[j][8] || ""),
          mrp: parseFloat(id[j][9]) || 0, gst: parseFloat(id[j][10]) || 0, taxable: parseFloat(id[j][11]) || 0,
          gstAmt: parseFloat(id[j][12]) || 0, lineTotal: parseFloat(id[j][13]) || 0 });
      }
    }
    var ts = row[1];
    return { success: true, print: {
      invoiceNo: String(row[0]),
      date: (ts instanceof Date) ? Utilities.formatDate(ts, tz, "dd-MMM-yyyy HH:mm") : _fmtDate_(ts),
      billType: String(row[2] || "WALK-IN"), patientId: String(row[3] || "WALK-IN"),
      patientName: String(row[4] || "Walk-in Patient"), mobile: String(row[5] || ""),
      age: String(row[6] || ""), sex: String(row[7] || ""), address: String(row[8] || ""),
      doctor: String(row[9] || "Self / OTC"), payMode: String(row[14] || "CASH"),
      txnId: String(row[15] || ""), payStatus: String(row[16] || ""),
      items: items, gross: parseFloat(row[10]) || 0, totalGst: parseFloat(row[11]) || 0,
      discount: parseFloat(row[12]) || 0, net: parseFloat(row[13]) || 0 } };
  } catch (e) { return { success: false, message: e.toString() }; }
}