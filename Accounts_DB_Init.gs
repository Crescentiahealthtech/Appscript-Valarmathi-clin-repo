// =========================================================================
// 🏛️ VALARMATHI CLINIC ERP - V2.0 ACCOUNTS DATABASE INITIALIZATION
// =========================================================================

function runEnterpriseAccountsSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define the 7 Enterprise Financial Pillars based on Master Blueprint
  const financialSheets = [
    {
      name: "Finance_Master_Ledger",
      headers: ["Txn_ID", "Timestamp", "Voucher_Type", "Category", "Department", "Reference_ID", "Entity_Name", "Payment_Mode", "Amount_In", "Amount_Out", "Logged_By", "Attachment_URL", "Is_Locked", "Notes"]
    },
    {
      name: "Accounts_Receivable", 
      headers: ["Due_ID", "Date", "Entity_Type", "Patient_Corp_Name", "Reference_ID", "Total_Bill_Amount", "Advance_Paid", "Pending_Balance", "Aging_Days", "Write_Off_Status", "Status", "Last_Updated"]
    },
    {
      name: "Accounts_Payable", 
      headers: ["Payable_ID", "Date", "Entity_Type", "Vendor_Doctor_Name", "Document_Type", "Invoice_Ref", "Total_Amount", "Amount_Paid", "Pending_Due", "Status"]
    },
    {
      name: "Insurance_Claims_Ledger", 
      headers: ["Claim_ID", "Date", "Patient_ID", "Patient_Name", "TPA_Company", "Policy_No", "PreAuth_Status", "Total_Billed", "Claimed_Amount", "Approved_Amount", "Deduction_Amount", "Patient_Liability_Log", "Settlement_Status"]
    },
    {
      name: "Shift_Registers", 
      headers: ["Shift_ID", "Date", "Shift_User", "Counter_Name", "Opening_Cash", "Cash_Collected", "Petty_Expenses", "Refunds_Issued", "Expected_Closing_Cash", "Actual_Physical_Cash", "Variance_Amount", "Status", "Timestamp"]
    },
    {
      name: "Tax_Ledger", 
      headers: ["Tax_Txn_ID", "Date", "Reference_Txn_ID", "Txn_Type", "GST_Category", "Base_Amount", "CGST", "SGST", "IGST", "Total_GST", "Filing_Status"]
    },
    {
      name: "Audit_Event_Ledger", 
      headers: ["Audit_ID", "Timestamp", "User_Name", "Action_Type", "Module", "Reference_ID", "Old_Value", "New_Value", "Reason_Remarks"]
    }
  ];

  let createdCount = 0;

  // Loop through and build the sheets safely
  financialSheets.forEach(sheetData => {
    let sheet = ss.getSheetByName(sheetData.name);
    
    // If sheet doesn't exist, create it
    if (!sheet) {
      sheet = ss.insertSheet(sheetData.name);
      
      // Set headers
      sheet.getRange(1, 1, 1, sheetData.headers.length).setValues([sheetData.headers]);
      
      // Style headers to look premium (SaaS Standard)
      sheet.getRange(1, 1, 1, sheetData.headers.length)
           .setFontWeight("bold")
           .setBackground("#1e293b") // Premium Slate Dark
           .setFontColor("#ffffff")
           .setWrap(true);
           
      // Freeze the top row so scrolling is easy
      sheet.setFrozenRows(1);
      
      createdCount++;
      Logger.log("✅ Created Enterprise Sheet: " + sheetData.name);
    } else {
      Logger.log("⚡ Sheet already exists, skipping: " + sheetData.name);
    }
  });
  
  if (createdCount > 0) {
    SpreadsheetApp.getUi().alert("✅ Success! " + createdCount + " Enterprise Accounts Sheets generated.");
  } else {
    SpreadsheetApp.getUi().alert("⚡ All 7 Accounts Sheets already exist.");
  }
}