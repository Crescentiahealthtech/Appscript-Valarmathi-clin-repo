// ==========================================
// 🧾 BILLING INVENTORY MODULE
// ==========================================

function generateInvoice(apptId, patientId, consultFee, pharmacyFee, labFee) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Billing_Ledger');
  if(!sheet) return {success: false, message: "Create 'Billing_Ledger' sheet first."};

  let invoiceId = "INV-" + (sheet.getLastRow().toString().padStart(4, '0'));
  let total = parseInt(consultFee) + parseInt(pharmacyFee) + parseInt(labFee);
  let timestamp = new Date();

  // Schema: Invoice ID | Date | Appt ID | Patient ID | Consult Fee | Pharmacy Fee | Lab Fee | Grand Total | Status
  sheet.appendRow([invoiceId, timestamp, apptId, patientId, consultFee, pharmacyFee, labFee, total, "Unpaid"]);
  SpreadsheetApp.flush();

  return {success: true, invoiceId: invoiceId, total: total};
}

function markInvoicePaid(invoiceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Billing_Ledger');
  const data = sheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] == invoiceId) {
      sheet.getRange(i + 1, 9).setValue("Paid");
      SpreadsheetApp.flush();
      return "Payment recorded successfully.";
    }
  }
  return "Invoice not found.";
}

// 🚀 NEW: Retrieve entire billing ledger for the Records UI
function getAllInvoices() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Billing_Ledger');
    if(!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let invoices = [];
    
    // Iterate backwards so the newest invoices show up at the top of the UI
    for(let i = data.length - 1; i >= 1; i--) {
      if(data[i][0]) {
        let dateObj = data[i][1];
        let dateStr = (dateObj instanceof Date) 
            ? Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MMM-yyyy") 
            : dateObj.toString().substring(0,10);
        
        invoices.push({
          invoiceId: data[i][0],
          date: dateStr,
          apptId: data[i][2],
          patientId: data[i][3],
          total: data[i][7],
          status: data[i][8]
        });
      }
    }
    return invoices;
  } catch (e) {
    return [];
  }
}