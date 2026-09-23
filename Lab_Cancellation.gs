// ============================================================================
// Lab_Cancellation.gs  —  Crescentia HealthTech
// Cancelling a lab order, and voiding a lab bill.
// ----------------------------------------------------------------------------
// WHAT WAS MISSING
//
// The lab billing desk could PROCESS a pending order and nothing else. Every
// other outcome — the patient left without paying, the clinician withdrew
// the request, the order was raised twice, the wrong patient was picked, a
// receipt was raised against the wrong bill — had no button. In practice
// those orders simply sat in the Pending queue for ever, which is why a
// count of "pending bills" stopped meaning anything: it was the real backlog
// plus every abandoned order since the system went in.
//
// TWO DIFFERENT ACTS, DELIBERATELY KEPT APART
//
//   CANCEL AN ORDER  — no money has moved. The request is withdrawn. The
//                      order leaves the queue and nothing is posted.
//
//   VOID A BILL      — money HAS been taken, or posted to an IP account.
//                      The document is reversed, not erased: the bill stays
//                      on the sheet marked CANCELLED with a reason, and the
//                      income drops out of the Finance Hub because
//                      acc_labRows_() counts only PaymentStatus = PAID.
//
// Conflating them is how a clinic loses the audit trail on a refund, so they
// are separate functions with separate permissions and separate messages.
//
// NOTHING IS EVER DELETED
//
// A row a patient was billed for is a financial record. Cancellation is a
// status and a reason, written beside the original figures, which stay
// exactly as they were. Deleting the row would leave the day's collection
// short with nothing to explain it.
//
// SETUP
//   Run labEnsureCancellationColumns() once after deploying this file.
// ============================================================================

/** Columns this file needs, appended to the two sheets that carry them. */
var LABX_CANCEL_COLS = {
  LAB_BILLING: ["CancelledAt", "CancelledBy", "CancelReason"],
  LAB_ORDERS:  ["CancelledAt", "CancelledBy", "CancelReason"]
};

/** '' for anything unusable. */
function labx_str_(v) { return (v === null || v === undefined) ? "" : String(v).trim(); }
function labx_up_(v) { return labx_str_(v).toUpperCase(); }

/**
 * ONE-OFF SETUP. Appends the cancellation columns to LAB_SCHEMA and to the
 * live sheets. Idempotent.
 *
 * The columns go on the END of both the schema array and the sheet, because
 * labHeaderMap_() resolves LAB_* sheets from LAB_SCHEMA by POSITION — every
 * existing index has to keep the value it has.
 */
function labEnsureCancellationColumns() {
  crescEditorOnly_('labEnsureCancellationColumns');
  var lock = LockService.getScriptLock();
  var report = [];
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    Object.keys(LABX_CANCEL_COLS).forEach(function (sheetName) {
      var wanted = LABX_CANCEL_COLS[sheetName];

      // a) the in-memory schema, so labHeaderMap_() knows the indices
      if (LAB_SCHEMA[sheetName]) {
        wanted.forEach(function (h) {
          if (LAB_SCHEMA[sheetName].indexOf(h) === -1) LAB_SCHEMA[sheetName].push(h);
        });
      }

      // b) the sheet itself
      var sh = ss.getSheetByName(sheetName);
      if (!sh) { report.push(sheetName + ": absent, skipped."); return; }

      var lastCol = sh.getLastColumn();
      var have = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(labx_str_) : [];
      var missing = wanted.filter(function (h) { return have.indexOf(h) === -1; });
      if (!missing.length) { report.push(sheetName + ": already has all three."); return; }

      sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing])
        .setFontWeight("bold").setBackground("#f4cccc");
      report.push(sheetName + ": added " + missing.join(", ") + ".");
    });

    SpreadsheetApp.flush();
    var msg = report.join("\n");
    Logger.log(msg);
    return msg;
  } catch (e) {
    return "labEnsureCancellationColumns failed: " + e.message;
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Sets a cell only when the sheet actually has that column. */
function labx_setIf_(sheet, rowNum, map, header, value) {
  if (map[header] === undefined) return;
  if (map[header] + 1 > sheet.getLastColumn()) return;
  sheet.getRange(rowNum, map[header] + 1).setValue(value);
}

/**
 * A cancellation reason has to say something.
 *
 * "cancel", "x", "test" and an empty box are all the same non-answer, and
 * this text is the whole of what a reviewer sees months later. Ten
 * characters is a low bar that a real reason clears without thinking.
 */
function labx_checkReason_(reason) {
  var r = labx_str_(reason);
  if (r.length < 10) {
    return "Give a reason for the cancellation — at least a few words. It is " +
           "the only explanation anyone reviewing this will have.";
  }
  if (!/[a-z]{3}/i.test(r)) return "The reason does not read as a reason.";
  return "";
}

// ---------------------------------------------------------------------------
// CANCEL AN ORDER THAT WAS NEVER BILLED
// ---------------------------------------------------------------------------

/**
 * Withdraws a pending lab order. No money is involved.
 *
 * REFUSED once the sample has been collected: at that point the lab has
 * consumed a tube and the patient has been bled, and the right record of
 * that is a rejected sample or an incomplete order, not a cancellation that
 * makes it look as though nothing happened. The message says so.
 *
 * @param {{orderId:string, reason:string}} payload
 * @param {string} sessionToken
 */
function labCancelPendingOrder(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};

    var actor = crescRequire_(sessionToken, "lab.order");

    var orderId = labx_str_(payload.orderId);
    if (!orderId) return { success: false, message: "No order was named." };

    var reasonProblem = labx_checkReason_(payload.reason);
    if (reasonProblem) return { success: false, message: reasonProblem };
    var reason = labx_str_(payload.reason);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(LAB.ORDERS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, message: "There are no lab orders." };
    }
    var map = labHeaderMap_(sheet);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    for (var i = 0; i < data.length; i++) {
      if (labx_str_(data[i][map["OrderID"]]) !== orderId) continue;
      var rowNum = i + 2;
      var status = labx_up_(data[i][map["OrderStatus"]]) || "PENDING";

      if (status === "CANCELLED") {
        return { success: false, message: "Order " + orderId + " is already cancelled." };
      }

      // Past collection, cancellation is the wrong record of what happened.
      var TOO_LATE = ["SAMPLE_COLLECTED", "IN_PROCESS", "RESULT_ENTERED",
                      "VERIFIED", "REPORT_DISPATCHED", "AMENDED"];
      if (TOO_LATE.indexOf(status) !== -1) {
        return { success: false,
                 message: "Order " + orderId + " is at " + status.replace(/_/g, " ").toLowerCase() +
                          ". The sample has already been taken, so this cannot be " +
                          "cancelled as though it never happened. Reject the sample " +
                          "or amend the report instead." };
      }

      // A billed order cannot be withdrawn on its own — the bill is the
      // thing that has to be voided, and that is a different permission.
      if (status === "BILLED" || labx_billFor_(orderId)) {
        return { success: false, code: "HAS_BILL",
                 message: "Order " + orderId + " has already been billed. Void the " +
                          "bill first — cancelling the order alone would leave the " +
                          "money on the books with nothing behind it." };
      }

      var now = new Date();
      labx_setIf_(sheet, rowNum, map, "OrderStatus", "CANCELLED");
      labx_setIf_(sheet, rowNum, map, "CancelledAt", now);
      labx_setIf_(sheet, rowNum, map, "CancelledBy", actor.username);
      labx_setIf_(sheet, rowNum, map, "CancelReason", reason);
      labx_setIf_(sheet, rowNum, map, "LastUpdatedAt",
        Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"));
      labx_setIf_(sheet, rowNum, map, "LastUpdatedBy", actor.username);

      // The order's individual tests go with it, or the worklist keeps them.
      labx_cancelOrderTests_(orderId, actor.username);

      labAudit_("ORDER_CANCELLED", "ORDER", orderId,
               { status: status }, { status: "CANCELLED", reason: reason, by: actor.username });

      SpreadsheetApp.flush();
      return { success: true,
               message: "Order " + orderId + " cancelled. It has left the pending queue." };
    }
    return { success: false, message: "No lab order with the id " + orderId + "." };

  } catch (err) {
    return { success: false, message: labx_friendly_(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Marks every test row of a cancelled order cancelled too. Best effort. */
function labx_cancelOrderTests_(orderId, username) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB.ORDER_TESTS);
    if (!sh || sh.getLastRow() < 2) return;
    var m = labHeaderMap_(sh);
    if (m["OrderID"] === undefined || m["TestStatus"] === undefined) return;

    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (labx_str_(data[i][m["OrderID"]]) !== labx_str_(orderId)) continue;
      var cur = labx_up_(data[i][m["TestStatus"]]);
      if (cur === "CANCELLED") continue;
      sh.getRange(i + 2, m["TestStatus"] + 1).setValue("CANCELLED");
    }
  } catch (e) { Logger.log("labx_cancelOrderTests_: " + e.message); }
}

/** The live (non-cancelled) bill id for an order, or ''. */
function labx_billFor_(orderId) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB.BILLING);
    if (!sh || sh.getLastRow() < 2) return "";
    var m = labHeaderMap_(sh);
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (labx_str_(data[i][m["OrderID"]]) !== labx_str_(orderId)) continue;
      if (labx_up_(data[i][m["PaymentStatus"]]) === "CANCELLED") continue;
      return labx_str_(data[i][m["BillID"]]);
    }
    return "";
  } catch (e) { return ""; }
}

// ---------------------------------------------------------------------------
// VOID A BILL THAT HAS ALREADY BEEN RAISED
// ---------------------------------------------------------------------------

/**
 * Reverses a lab bill. The row stays; its status becomes CANCELLED.
 *
 * WHAT THIS CHANGES ELSEWHERE, ON PURPOSE:
 *   • acc_labRows_() counts income only where PaymentStatus is PAID, so the
 *     amount leaves the Finance Hub, the day's collection and the drawer
 *     reconciliation the moment this is written. Nothing has to be told.
 *   • getLabBillingWorkspace() ignores cancelled bills when it decides which
 *     orders are already billed, so the order returns to the pending queue
 *     and can be billed again correctly — which is the usual reason for
 *     voiding one.
 *
 * IT DOES NOT HAND BACK CASH. Whoever voids the bill still has to take the
 * money out of the drawer, and the shift reconciliation will show it. The
 * message says so rather than letting anyone assume otherwise.
 *
 * @param {{billId:string, reason:string, returnToQueue:boolean}} payload
 * @param {string} sessionToken
 */
function labCancelBill(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    payload = payload || {};

    var actor = crescRequire_(sessionToken, "billing.write");

    var billId = labx_str_(payload.billId);
    if (!billId) return { success: false, message: "No bill was named." };

    var reasonProblem = labx_checkReason_(payload.reason);
    if (reasonProblem) return { success: false, message: reasonProblem };
    var reason = labx_str_(payload.reason);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(LAB.BILLING);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, message: "There are no lab bills." };
    }
    var map = labHeaderMap_(sheet);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    for (var i = 0; i < data.length; i++) {
      if (labx_str_(data[i][map["BillID"]]) !== billId) continue;
      var rowNum = i + 2;
      var status = labx_up_(data[i][map["PaymentStatus"]]);
      if (status === "CANCELLED") {
        return { success: false, message: "Bill " + billId + " is already cancelled." };
      }

      var orderId = labx_str_(data[i][map["OrderID"]]);
      var net = Number(data[i][map["NetAmount"]]) || 0;
      var mode = labx_str_(data[i][map["PaymentMode"]]);
      var admissionId = labx_str_(data[i][map["AdmissionID"]]);

      // A released report is a strong signal that this bill is not a
      // mis-keying but a real episode of care. It is still allowed — refunds
      // after a report happen — but it is named in the audit and reported
      // back so the desk knows what it has just done.
      var reportOut = labx_orderStatus_(orderId);
      var released = (reportOut === "VERIFIED" || reportOut === "REPORT_DISPATCHED" ||
                      reportOut === "AMENDED");

      var now = new Date();
      labx_setIf_(sheet, rowNum, map, "PaymentStatus", "CANCELLED");
      labx_setIf_(sheet, rowNum, map, "BalanceAmount", 0);
      labx_setIf_(sheet, rowNum, map, "PaidAmount", 0);
      labx_setIf_(sheet, rowNum, map, "CancelledAt", now);
      labx_setIf_(sheet, rowNum, map, "CancelledBy", actor.username);
      labx_setIf_(sheet, rowNum, map, "CancelReason", reason);

      labAudit_("BILL_CANCELLED", "BILL", billId,
               { status: status, net: net, mode: mode },
               { status: "CANCELLED", reason: reason, by: actor.username,
                 reportAlreadyReleased: released });

      // The order goes back to PENDING so it can be billed again, unless the
      // caller is voiding the whole episode.
      var orderNote = "";
      if (orderId) {
        if (payload.returnToQueue === false) {
          var c = labCancelPendingOrderInternal_(orderId, reason, actor.username);
          orderNote = c ? " Order " + orderId + " cancelled with it." : "";
        } else if (labx_orderStatus_(orderId) === "BILLED") {
          labx_setOrderStatus_(orderId, "PENDING", actor.username);
          orderNote = " Order " + orderId + " is back in the pending queue.";
        }
      }

      SpreadsheetApp.flush();

      var money = (labx_up_(mode) === "IP_ACCOUNT" || admissionId)
        ? " The ₹" + net.toFixed(2) + " posted to the IP account is reversed; check " +
          "the running tab before discharge."
        : " ₹" + net.toFixed(2) + " has NOT been taken out of the drawer — do that " +
          "by hand, or the shift will reconcile short.";

      return { success: true,
               returnedToQueue: orderNote.indexOf("pending queue") !== -1,
               reportAlreadyReleased: released,
               message: "Bill " + billId + " cancelled." + orderNote + money +
                        (released ? " Note: the report for this order has already been released."
                                  : "") };
    }
    return { success: false, message: "No lab bill with the id " + billId + "." };

  } catch (err) {
    return { success: false, message: labx_friendly_(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Cancels an order without re-checking permissions — internal to a void. */
function labCancelPendingOrderInternal_(orderId, reason, username) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB.ORDERS);
    if (!sh || sh.getLastRow() < 2) return false;
    var m = labHeaderMap_(sh);
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (labx_str_(data[i][m["OrderID"]]) !== labx_str_(orderId)) continue;
      var now = new Date();
      labx_setIf_(sh, i + 2, m, "OrderStatus", "CANCELLED");
      labx_setIf_(sh, i + 2, m, "CancelledAt", now);
      labx_setIf_(sh, i + 2, m, "CancelledBy", username);
      labx_setIf_(sh, i + 2, m, "CancelReason", reason);
      labx_cancelOrderTests_(orderId, username);
      labAudit_("ORDER_CANCELLED", "ORDER", orderId, null,
               { status: "CANCELLED", reason: reason, by: username, via: "BILL_VOID" });
      return true;
    }
    return false;
  } catch (e) { return false; }
}

/** An order's current status, upper-cased, or ''. */
function labx_orderStatus_(orderId) {
  try {
    if (!orderId) return "";
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB.ORDERS);
    if (!sh || sh.getLastRow() < 2) return "";
    var m = labHeaderMap_(sh);
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (labx_str_(data[i][m["OrderID"]]) === labx_str_(orderId)) {
        return labx_up_(data[i][m["OrderStatus"]]);
      }
    }
    return "";
  } catch (e) { return ""; }
}

/** Writes an order status directly — used to put a voided bill's order back. */
function labx_setOrderStatus_(orderId, status, username) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB.ORDERS);
    if (!sh || sh.getLastRow() < 2) return;
    var m = labHeaderMap_(sh);
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      if (labx_str_(data[i][m["OrderID"]]) !== labx_str_(orderId)) continue;
      labx_setIf_(sh, i + 2, m, "OrderStatus", status);
      labx_setIf_(sh, i + 2, m, "LastUpdatedAt",
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"));
      labx_setIf_(sh, i + 2, m, "LastUpdatedBy", username);
      return;
    }
  } catch (e) { Logger.log("labx_setOrderStatus_: " + e.message); }
}

/**
 * Recent cancellations, both kinds, newest first.
 * The queue counts are only trustworthy if what left them can be seen.
 */
function labListCancelled(sessionToken, limit) {
  try {
    crescRequire_(sessionToken, "lab.read");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var out = [];
    var cap = limit || 60;

    var bs = ss.getSheetByName(LAB.BILLING);
    if (bs && bs.getLastRow() > 1) {
      var bm = labHeaderMap_(bs);
      var bd = bs.getRange(2, 1, bs.getLastRow() - 1, bs.getLastColumn()).getValues();
      for (var i = bd.length - 1; i >= 0 && out.length < cap; i--) {
        if (labx_up_(bd[i][bm["PaymentStatus"]]) !== "CANCELLED") continue;
        out.push({
          kind: "BILL",
          id: labx_str_(bd[i][bm["BillID"]]),
          orderId: labx_str_(bd[i][bm["OrderID"]]),
          patientName: labx_str_(bd[i][bm["PatientName"]]),
          patientId: labx_str_(bd[i][bm["PatientID"]]),
          amount: Number(bd[i][bm["NetAmount"]]) || 0,
          at: labx_cell_(bd[i], bm, "CancelledAt"),
          by: labx_cell_(bd[i], bm, "CancelledBy"),
          reason: labx_cell_(bd[i], bm, "CancelReason")
        });
      }
    }

    var os = ss.getSheetByName(LAB.ORDERS);
    if (os && os.getLastRow() > 1) {
      var om = labHeaderMap_(os);
      var od = os.getRange(2, 1, os.getLastRow() - 1, os.getLastColumn()).getValues();
      for (var j = od.length - 1; j >= 0 && out.length < cap; j--) {
        if (labx_up_(od[j][om["OrderStatus"]]) !== "CANCELLED") continue;
        out.push({
          kind: "ORDER",
          id: labx_str_(od[j][om["OrderID"]]),
          orderId: labx_str_(od[j][om["OrderID"]]),
          patientName: labx_str_(od[j][om["PatientName"]]),
          patientId: labx_str_(od[j][om["PatientID"]]),
          amount: 0,
          at: labx_cell_(od[j], om, "CancelledAt"),
          by: labx_cell_(od[j], om, "CancelledBy"),
          reason: labx_cell_(od[j], om, "CancelReason")
        });
      }
    }

    out.sort(function (a, b) { return String(b.at).localeCompare(String(a.at)); });
    return { success: true, rows: out, message: "" };
  } catch (err) {
    return { success: false, rows: [], message: labx_friendly_(err) };
  }
}

/** A cell by header name, '' when the sheet predates the column. */
function labx_cell_(row, map, header) {
  if (map[header] === undefined) return "";
  var v = row[map[header]];
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd-MMM-yyyy HH:mm");
  }
  return labx_str_(v);
}

/**
 * A FORBIDDEN from crescRequire_ is already written for a human; anything
 * else is an internal fault and should not be shown raw.
 */
function labx_friendly_(err) {
  var m = (err && err.message) ? String(err.message) : String(err);
  if (m.indexOf("FORBIDDEN:") === 0) return m.replace("FORBIDDEN: ", "");
  return "Could not complete: " + m;
}
