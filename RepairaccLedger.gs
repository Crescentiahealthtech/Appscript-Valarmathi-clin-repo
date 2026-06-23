// =========================================================================
// 🛠️ ONE-TIME REPAIR — fix swapped Amount_In / Amount_Out
// Earlier builds of Payables / Insurance / IP posted In and Out reversed.
// This swaps ONLY the affected rows that still show the reversal signature,
// so it is safe to run and idempotent (running twice changes nothing more).
//
// HOW TO RUN: open the Apps Script editor, select repairLedgerInOut from the
// function dropdown, click Run. Read the toast/log for the count.
// =========================================================================

function repairLedgerInOut() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sh = acc_sheet_(ACC_CFG.LEDGER), d = sh.getDataRange().getValues();
    if (d.length < 2) return { success: true, message: "Ledger empty — nothing to repair." };
    var h = d[0].map(function (x) { return acc_str_(x).trim(); });
    var cV = h.indexOf('Voucher_Type'), cCat = h.indexOf('Category'), cIn = h.indexOf('Amount_In'), cOut = h.indexOf('Amount_Out');
    if (cIn < 0 || cOut < 0) return { success: false, message: "Amount_In / Amount_Out columns not found." };

    // rows that SHOULD be expenses (value belongs in Amount_Out)
    var EXPENSE_VOUCHERS = ['PAYABLE'];
    var OUT_CATS = ['IP_REFUND'];
    // rows that SHOULD be receipts (value belongs in Amount_In)
    var IN_CATS = ['INSURANCE_RECEIPT', 'IP_ADVANCE', 'IP_SETTLEMENT'];

    var fixedOut = 0, fixedIn = 0;
    for (var i = 1; i < d.length; i++) {
      var v = acc_str_(d[i][cV]).toUpperCase(), cat = acc_str_(d[i][cCat]).toUpperCase();
      var inn = acc_money_(d[i][cIn]), out = acc_money_(d[i][cOut]);

      var shouldBeOut = (EXPENSE_VOUCHERS.indexOf(v) !== -1) || (OUT_CATS.indexOf(cat) !== -1);
      var shouldBeIn = (IN_CATS.indexOf(cat) !== -1);

      // reversal signature: the value is in the wrong column and the right one is empty
      if (shouldBeOut && inn > 0 && out === 0) { d[i][cOut] = inn; d[i][cIn] = 0; fixedOut++; }
      else if (shouldBeIn && out > 0 && inn === 0) { d[i][cIn] = out; d[i][cOut] = 0; fixedIn++; }
    }

    if (fixedOut + fixedIn > 0) {
      sh.getRange(1, 1, d.length, d[0].length).setValues(d);
      acc_audit_('SYSTEM', 'LEDGER_REPAIR', 'Finance_Master_Ledger', '', '', fixedOut + fixedIn, fixedOut + ' expenses + ' + fixedIn + ' receipts re-aligned');
      SpreadsheetApp.flush();
    }
    var msg = "Repair done. Expenses re-aligned: " + fixedOut + " · Receipts re-aligned: " + fixedIn + ".";
    Logger.log(msg);
    try { SpreadsheetApp.getActive().toast(msg, 'Ledger Repair', 8); } catch (e) {}
    return { success: true, message: msg, fixedOut: fixedOut, fixedIn: fixedIn };
  } catch (e) { return { success: false, message: "Error: " + e.message }; }
  finally { lock.releaseLock(); }
}