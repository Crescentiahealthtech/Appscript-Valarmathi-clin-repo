// ============================================================================
// Billinginventory.gs  —  Crescentia HealthTech / CresRx
// LEGACY. Superseded by Hospital_Billing.gs.
// ----------------------------------------------------------------------------
// This file used to hold the whole of hospital billing: generateInvoice(),
// markInvoicePaid() and getAllInvoices(), all writing a flat `Billing_Ledger`
// sheet whose entire schema was
//
//     Invoice ID | Date | Appt ID | Patient ID | Consult Fee |
//     Pharmacy Fee | Lab Fee | Grand Total | Status
//
// — one row, three hard-coded fee buckets, no line items, no packages, no tax,
// no part payment, and nothing that Accounts ever read. Nothing called the
// writers: the billing screen's Save button was a placeholder alert, so the
// only rows that sheet ever received came from manual editing.
//
// Hospital_Billing.gs replaces all of it with Hospital_Invoices +
// Hospital_Invoice_Items, priced from Service_Master and Package_Master and
// merged into the Finance Hub by acc_hospitalRows_().
//
// The WRITERS are deliberately gone rather than deprecated. Leaving them in
// the project meant a second, invisible billing ledger was one stray call
// away — and a hospital with two ledgers has none.
//
// The reader below stays so the rows already in `Billing_Ledger` remain
// reachable. It is read-only and returns an empty list when the sheet was
// never created.
// ============================================================================

/**
 * Historical rows from the pre-2026 `Billing_Ledger` sheet, newest first.
 * Not part of any live workflow — for migration and reference only.
 */
function getLegacyBillingLedger() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Billing_Ledger');
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data = sheet.getDataRange().getValues();
    var out = [];
    for (var i = data.length - 1; i >= 1; i--) {
      if (!data[i][0]) continue;
      var d = data[i][1];
      out.push({
        invoiceId: String(data[i][0]),
        date: (d instanceof Date)
          ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy')
          : String(d || '').substring(0, 10),
        apptId: String(data[i][2] || ''),
        patientId: String(data[i][3] || ''),
        total: Number(data[i][7]) || 0,
        status: String(data[i][8] || '')
      });
    }
    return out;
  } catch (e) {
    return [];
  }
}