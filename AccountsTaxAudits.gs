// =========================================================================
// 🏛️ CRESCENTIA — TAX & AUDIT
// Tax: net GST liability for a month = Output GST (pharmacy PAID sales) −
//   Input GST (vendor invoices' GST_Input). Medical services are exempt, so
//   only pharmacy/consumables carry output GST. Computed live (no posting).
// Audit: reads Audit_Event_Ledger (newest first), optional module filter.
// Depends on Accounts.js: ACC_CFG, acc_readObjects_, acc_money_, acc_str_, acc_toDate_.
// =========================================================================

function tax_month_(v) { var d = acc_toDate_(v); return d ? Utilities.formatDate(d, ACC_CFG.TZ, 'yyyy-MM') : ''; }

function getTaxSummary(monthKey) {
  try {
    var month = acc_str_(monthKey) || Utilities.formatDate(new Date(), ACC_CFG.TZ, 'yyyy-MM');

    // OUTPUT GST — pharmacy sales actually realized (Pay_Status PAID), active bills
    var outputGST = 0, taxableSales = 0, salesCount = 0;
    acc_readObjects_('Pharmacy_Invoices').forEach(function (r) {
      if (acc_str_(r['Pay_Status']).toUpperCase() !== 'PAID') return;
      if (acc_str_(r['Status']).toUpperCase() === 'VOID') return;
      if (tax_month_(r['Timestamp']) !== month) return;
      var gst = acc_money_(r['Total_GST']);
      outputGST += gst; taxableSales += acc_money_(r['Net']); salesCount++;
    });

    // INPUT GST — vendor invoices with GST_Input in the month
    var inputGST = 0, inputBase = 0, vendorCount = 0;
    acc_readObjects_('Accounts_Payable').forEach(function (r) {
      var gst = acc_money_(r['GST_Input']);
      if (gst <= 0) return;
      if (tax_month_(r['Date']) !== month) return;
      inputGST += gst; inputBase += acc_money_(r['Total_Amount']); vendorCount++;
    });

    var net = acc_money_(outputGST - inputGST);
    return {
      success: true, month: month,
      outputGST: acc_money_(outputGST), inputGST: acc_money_(inputGST), netLiability: net,
      taxableSales: acc_money_(taxableSales), salesCount: salesCount, inputBase: acc_money_(inputBase), vendorCount: vendorCount,
      direction: net >= 0 ? 'PAYABLE' : 'CREDIT'
    };
  } catch (e) { return { success: false, message: e.message }; }
}

function getAuditLog(payload) {
  try {
    payload = payload || {};
    var limit = Math.min(Math.max(parseInt(payload.limit, 10) || 60, 1), 200);
    var mod = acc_str_(payload.module).toUpperCase();
    var q = acc_str_(payload.query).toLowerCase();
    var rows = acc_readObjects_(ACC_CFG.AUDIT || 'Audit_Event_Ledger'), out = [];
    for (var i = rows.length - 1; i >= 0 && out.length < limit; i--) {
      var r = rows[i];
      var module = acc_str_(r['Module']);
      if (mod && mod !== 'ALL' && module.toUpperCase().indexOf(mod) === -1) continue;
      var d = acc_toDate_(r['Timestamp']);
      var rec = {
        ts: d ? Utilities.formatDate(d, ACC_CFG.TZ, 'dd-MMM HH:mm') : acc_str_(r['Timestamp']),
        user: acc_str_(r['User_Name']), action: acc_str_(r['Action_Type']), module: module,
        ref: acc_str_(r['Reference_ID']), oldVal: acc_str_(r['Old_Value']), newVal: acc_str_(r['New_Value']), reason: acc_str_(r['Reason_Remarks'])
      };
      if (q) { var blob = (rec.user + rec.action + rec.module + rec.ref + rec.reason).toLowerCase(); if (blob.indexOf(q) === -1) continue; }
      out.push(rec);
    }
    return { success: true, events: out };
  } catch (e) { return { success: false, message: e.message }; }
}

// distinct modules for the audit filter dropdown
function getAuditModules() {
  try {
    var seen = {}, rows = acc_readObjects_(ACC_CFG.AUDIT || 'Audit_Event_Ledger');
    rows.forEach(function (r) { var m = acc_str_(r['Module']); if (m) seen[m] = true; });
    return { success: true, modules: Object.keys(seen).sort() };
  } catch (e) { return { success: false, message: e.message }; }
}