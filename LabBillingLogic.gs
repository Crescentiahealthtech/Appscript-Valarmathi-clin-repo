// ==========================================
// 💰 LAB BILLING DESK BACKEND ENGINE
// ==========================================

function getLabBillingWorkspace() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const billingSheet = ss.getSheetByName("LAB_BILLING");
    const ordersSheet = ss.getSheetByName("LAB_ORDERS");
    
    if (!billingSheet || !ordersSheet) {
      return { success: false, message: "Database Sheets missing." };
    }

    // 1. Establish strict Midnight-to-Midnight Boundaries for "Today" (IST Timezone)
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    
    const startOfToday = new Date(now);
    startOfToday.setHours(0, 0, 0, 0);
    
    const endOfToday = new Date(now);
    endOfToday.setHours(23, 59, 59, 999);

    const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
    let d = new Date(now); d.setDate(d.getDate() - 1);
    const yesterdayStr = Utilities.formatDate(d, tz, "yyyy-MM-dd");

    // 2. Fetch & Parse Billing Data
    const bData = billingSheet.getDataRange().getValues();
    const bHeaders = bData.length > 0 ? bData.shift() : []; 
    
    let billedOrderIds = new Set();
    let billsList = [];
    let stats = { pendingCount: 0, todayOP: 0, todayIP: 0, todayCount: 0 };

    bData.forEach(row => {
      let b = {};
      bHeaders.forEach((h, i) => b[h] = row[i]);
      
      if (!b.BillID) return;
      if (b.OrderID) billedOrderIds.add(b.OrderID);

      let billDateObj = null;
      let billDateStr = "";
      
      if (b.BilledAt) {
        billDateObj = new Date(b.BilledAt);
        if (!isNaN(billDateObj.getTime())) {
          billDateStr = Utilities.formatDate(billDateObj, tz, "yyyy-MM-dd");
        }
      }
      
      let isStrictlyToday = false;
      if (billDateObj && billDateObj >= startOfToday && billDateObj <= endOfToday) {
        isStrictlyToday = true;
      } else if (billDateStr === todayStr) {
        isStrictlyToday = true;
      }

      if (isStrictlyToday || billDateStr === yesterdayStr) {
        
        let isIP = (b.PaymentMode === 'IP_ACCOUNT' || b.BillingCategory === 'IP_ACCOUNT' || b.AdmissionID);
        let tType = isIP ? 'IP' : 'PAID';
        
        billsList.push({
          tabType: tType,
          orderId: b.OrderID || '',
          billId: b.BillID,
          paymentMode: b.PaymentMode || 'CASH',
          patientName: b.PatientName || 'Unknown',
          patientId: b.PatientID || '',
          testNames: parseTestsForUI(b.TestsJSON) || "Lab Tests",
          receiptNumber: b.ReceiptNumber || b.BillID,
          billedAt: b.BilledAt ? new Date(b.BilledAt).toLocaleString('en-IN') : '',
          net: Number(b.NetAmount) || 0,
          discount: Number(b.DiscountAmount) || 0,
          isStrictlyToday: isStrictlyToday 
        });

        if (isStrictlyToday) {
          stats.todayCount++;
          if (isIP) {
            stats.todayIP += (Number(b.NetAmount) || 0);
          } else {
            stats.todayOP += (Number(b.NetAmount) || 0);
          }
        }
      }
    });

    // 3. Fetch PENDING Orders
    const oData = ordersSheet.getDataRange().getValues();
    const oHeaders = oData.length > 0 ? oData.shift() : [];

    oData.forEach(row => {
      let o = {};
      oHeaders.forEach((h, i) => o[h] = row[i]);
      
      if (!o.OrderID) return;
      
      if (!billedOrderIds.has(o.OrderID) && o.OrderStatus !== 'CANCELLED' && o.OrderStatus !== 'DELETE') {
        stats.pendingCount++;
        
        billsList.push({
          tabType: 'PENDING',
          orderId: o.OrderID,
          billId: '',
          paymentMode: '',
          patientName: o.PatientName || 'Unknown',
          patientId: o.PatientID || '',
          testNames: o.TestNames || "Pending Tests",
          receiptNumber: '',
          billedAt: o.CreatedAt ? new Date(o.CreatedAt).toLocaleString('en-IN') : '',
          net: 0, 
          discount: 0,
          isStrictlyToday: false
        });
      }
    });

    return { success: true, bills: billsList, stats: stats };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function parseTestsForUI(jsonStr) {
  try {
    let arr = JSON.parse(jsonStr);
    return arr.map(t => t.testName).join(", ");
  } catch (e) {
    return "Lab Tests";
  }
}

function getLabDailyCollection() {
  try {
    const ws = getLabBillingWorkspace();
    if (!ws.success) throw new Error(ws.message);

    let dayBills = ws.bills.filter(b => (b.tabType === 'PAID' || b.tabType === 'IP') && b.isStrictlyToday === true);
    dayBills.sort((a, b) => new Date(a.billedAt) - new Date(b.billedAt));

    return { success: true, bills: dayBills, totalNet: ws.stats.todayOP };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Generates the Physical HTML Cashier Receipt
 */
// ==========================================
// 🖨️ LAB RECEIPT PRINT ENGINE (BACKEND)
// ==========================================

/**
 * Generates the physical HTML for the Lab Receipt Pop-up
 * Matches the premium UI/UX of the Lab Integration Engine
 */
function getLabReceiptHtml(billId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("LAB_BILLING");
    if (!sheet) throw new Error("LAB_BILLING sheet not found.");

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let b = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf("BillID")] === billId) {
        b = {
          billId: data[i][headers.indexOf("BillID")] || '',
          orderId: data[i][headers.indexOf("OrderID")] || '',
          patientId: data[i][headers.indexOf("PatientID")] || '',
          patientName: data[i][headers.indexOf("PatientName")] || 'Unknown Patient',
          category: data[i][headers.indexOf("BillingCategory")] || '',
          itemsJSON: data[i][headers.indexOf("TestsJSON")] || '[]',
          gross: Number(data[i][headers.indexOf("GrossAmount")]) || 0,
          discPct: Number(data[i][headers.indexOf("DiscountPercent")]) || 0,
          discAmt: Number(data[i][headers.indexOf("DiscountAmount")]) || 0,
          net: Number(data[i][headers.indexOf("NetAmount")]) || 0,
          payMode: data[i][headers.indexOf("PaymentMode")] || '',
          paid: Number(data[i][headers.indexOf("PaidAmount")]) || 0,
          balance: Number(data[i][headers.indexOf("BalanceAmount")]) || 0,
          payStatus: data[i][headers.indexOf("PaymentStatus")] || '',
          receipt: data[i][headers.indexOf("ReceiptNumber")] || billId,
          billedAt: data[i][headers.indexOf("BilledAt")] ? new Date(data[i][headers.indexOf("BilledAt")]).toLocaleString('en-IN') : ""
        };
        break;
      }
    }
    
    if (!b) throw new Error("Bill not found in database.");

    // Parse itemized tests
    let items = [];
    try { if (b.itemsJSON) items = JSON.parse(b.itemsJSON); } catch (e) {}
    
    let itemRows = items.map((it, idx) => {
      return `
        <tr>
          <td style="padding:8px 10px; border-bottom:1px solid #e5e7eb;">${idx + 1}</td>
          <td style="padding:8px 10px; border-bottom:1px solid #e5e7eb;">
            <strong>${_esc(it.testName)}</strong>
            ${it.testId ? `<br><span style="font-size:10px;color:#6b7280;">${_esc(it.testId)}</span>` : ''}
          </td>
          <td style="padding:8px 10px; border-bottom:1px solid #e5e7eb; text-align:right;">₹${Number(it.price).toFixed(2)}</td>
        </tr>
      `;
    }).join('');

    // Branding Properties (Fallback if not set)
    const props = PropertiesService.getScriptProperties().getProperties();
    const clinicName    = props['CLINIC_NAME']    || 'Crescentia HealthTech';
    const clinicAddress = props['CLINIC_ADDRESS'] || 'Medical District, City';
    const clinicPhone   = props['CLINIC_PHONE']   || '+91 9876543210';
    const gstNumber     = props['CLINIC_GST']     || '';

    const isIp = (b.category === 'IP_ACCOUNT');
    
    const payLine = isIp
      ? `<div style="font-weight:700;color:#0369a1;">Posted to IP Account &bull; Settled at discharge</div>`
      : `<div>Payment: <strong>${_esc(b.payMode)}</strong> &bull; Status: <strong style="${b.payStatus === 'PAID' ? 'color:#10b981;' : 'color:#dc2626;'}">${_esc(b.payStatus)}</strong></div>`;

    // Construct Clean, Premium HTML
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>Lab Invoice - ${b.receipt}</title>
        <style>
          *{box-sizing:border-box;margin:0;padding:0;font-family:'Helvetica Neue',Arial,sans-serif;}
          body{background:#fff;color:#111827;}
          @media print{@page{margin:1cm;} .no-print{display:none!important;} body{background:#fff;} }
        </style>
      </head>
      <body>
        <div style="max-width:760px;margin:18px auto;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">
          
          <div style="padding:18px 24px;border-bottom:2px solid #0369a1;display:flex;justify-content:space-between;align-items:flex-start;">
            <div>
              <div style="font-size:20px;font-weight:800;color:#0369a1;text-transform:uppercase;">${_esc(clinicName)}</div>
              ${clinicAddress ? `<div style="font-size:12px;color:#6b7280;margin-top:4px;">${_esc(clinicAddress)}</div>` : ''}
              ${clinicPhone ? `<div style="font-size:12px;color:#6b7280;">Phone: ${_esc(clinicPhone)}</div>` : ''}
              ${gstNumber ? `<div style="font-size:11px;color:#6b7280;margin-top:2px;">GSTIN: ${_esc(gstNumber)}</div>` : ''}
            </div>
            <div style="text-align:right;">
              <div style="font-size:16px;font-weight:800;letter-spacing:1px;color:#111827;">LAB INVOICE</div>
              <div style="font-size:12px;color:#4b5563;margin-top:4px;">Inv: <strong>${_esc(b.receipt)}</strong></div>
              <div style="font-size:11px;color:#6b7280;">Date: ${_esc(b.billedAt)}</div>
            </div>
          </div>

          <div style="padding:12px 24px;background:#f9fafb;border-bottom:1px solid #e5e7eb;display:flex;justify-content:space-between;font-size:13px;">
            <div><span style="color:#6b7280;">Patient:</span> <strong>${_esc(b.patientName)}</strong> &bull; ${_esc(b.patientId)}</div>
            <div><span style="color:#6b7280;">Order ID:</span> ${_esc(b.orderId)}</div>
          </div>

          <div style="padding:8px 24px;">
            <table style="width:100%;border-collapse:collapse;font-size:13px;">
              <thead>
                <tr style="background:#f3f4f6;">
                  <th style="padding:8px 10px;text-align:left;color:#374151;font-size:11px;text-transform:uppercase;">#</th>
                  <th style="padding:8px 10px;text-align:left;color:#374151;font-size:11px;text-transform:uppercase;">Investigation</th>
                  <th style="padding:8px 10px;text-align:right;color:#374151;font-size:11px;text-transform:uppercase;">Amount</th>
                </tr>
              </thead>
              <tbody>
                ${itemRows}
              </tbody>
            </table>
          </div>

          <div style="padding:8px 24px 16px;display:flex;justify-content:flex-end;">
            <table style="font-size:13px;min-width:260px;">
              <tr><td style="padding:4px 10px;color:#6b7280;">Gross Amount</td><td style="padding:4px 10px;text-align:right;">₹${b.gross.toFixed(2)}</td></tr>
              ${b.discAmt > 0 ? `<tr><td style="padding:4px 10px;color:#10b981;">Discount (${b.discPct}%)</td><td style="padding:4px 10px;text-align:right;color:#10b981;">- ₹${b.discAmt.toFixed(2)}</td></tr>` : ''}
              <tr style="border-top:2px solid #111827;">
                <td style="padding:6px 10px;font-weight:800;font-size:15px;">Net Payable</td>
                <td style="padding:6px 10px;text-align:right;font-weight:800;font-size:16px;color:#0369a1;">₹${b.net.toFixed(2)}</td>
              </tr>
              ${!isIp ? `<tr><td style="padding:4px 10px;color:#6b7280;">Paid</td><td style="padding:4px 10px;text-align:right;">₹${b.paid.toFixed(2)}</td></tr>` : ''}
              ${!isIp && b.balance > 0 ? `<tr><td style="padding:4px 10px;color:#dc2626;">Balance</td><td style="padding:4px 10px;text-align:right;color:#dc2626;font-weight:700;">₹${b.balance.toFixed(2)}</td></tr>` : ''}
            </table>
          </div>

          <div style="padding:12px 24px;background:#f9fafb;border-top:1px solid #e5e7eb;font-size:12px;color:#374151;display:flex;justify-content:space-between;align-items:center;">
            ${payLine}
            <div style="color:#9ca3af;">This is a computer-generated invoice.</div>
          </div>

        </div>

        <div class="no-print" style="text-align:center;padding:14px;margin-bottom:20px;">
          <button onclick="window.print();" style="background:#0369a1;color:#fff;border:none;padding:12px 28px;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;box-shadow:0 4px 6px rgba(3,105,161,0.2);">
            🖨️ Print / Save PDF
          </button>
        </div>
      </body>
      </html>
    `;

    return { success: true, html: html };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function _esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"'`=\/]/g, function (s) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;', '/': '&#x2F;', '`': '&#x60;', '=': '&#x3D;' }[s];
  });
}

// ==========================================
// 🚀 LAB INVOICE COMMUNICATION ENGINE (WHATSAPP & EMAIL)
// ==========================================

/**
 * Server-Side function: Generates Invoice PDF, saves to Drive, and returns public link for WhatsApp.
 */
function generateAndStoreLabInvoicePDF(billId) {
  try {
    // 1. Generate HTML using existing engine
    const reportResponse = getLabReceiptHtml(billId); 
    if (!reportResponse.success) throw new Error("HTML Generation Failed: " + reportResponse.message);

    // 2. Convert to PDF Blob
    const htmlBlob = Utilities.newBlob(reportResponse.html, 'text/html', 'invoice.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Crescentia_Lab_Invoice_" + billId + ".pdf");

    // 3. Drive Folder Architecture specifically for Invoices
    const rootFolderName = "Crescentia_Lab_Invoices";
    let rootFolder;
    const rootFolders = DriveApp.getFoldersByName(rootFolderName);
    if (rootFolders.hasNext()) {
      rootFolder = rootFolders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }

    const now = new Date();
    const yearStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy");
    const monthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");

    let yearFolder;
    const yearFolders = rootFolder.getFoldersByName(yearStr);
    if (yearFolders.hasNext()) yearFolder = yearFolders.next();
    else yearFolder = rootFolder.createFolder(yearStr);

    let monthFolder;
    const monthFolders = yearFolder.getFoldersByName(monthStr);
    if (monthFolders.hasNext()) monthFolder = monthFolders.next();
    else monthFolder = yearFolder.createFolder(monthStr);

    // 4. Save File & Set Permissions
    const file = monthFolder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 5. Return Link
    return { success: true, link: file.getUrl() };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Server-Side function: Generates Invoice PDF and emails it directly via GMAIL API.
 */
function emailLabInvoicePDF(billId, patientEmail) {
  try {
    const reportResponse = getLabReceiptHtml(billId); 
    if (!reportResponse.success) throw new Error("HTML Generation Failed");

    const htmlBlob = Utilities.newBlob(reportResponse.html, 'text/html', 'invoice.html');
    const pdfBlob = htmlBlob.getAs('application/pdf');
    pdfBlob.setName("Crescentia_Lab_Invoice_" + billId + ".pdf");

    const subject = "Your Payment Receipt - Crescentia Clinic & Diagnostics";
    const plainBody = "Dear Patient, please find your lab invoice attached. Regards, Crescentia Clinic.";

    // Premium HTML Email Layout for GmailApp
    const richHtmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #0369a1; color: white; padding: 20px; text-align: center;">
          <h2 style="margin: 0; letter-spacing: 1px;">CRESCENTIA CLINIC & DIAGNOSTICS</h2>
        </div>
        <div style="padding: 30px;">
          <p style="font-size: 16px;">Dear Patient,</p>
          <p style="font-size: 15px; line-height: 1.5;">Thank you for choosing Crescentia HealthTech. We have received your payment for the recent laboratory investigations.</p>
          
          <div style="background-color: #f0fdf4; border-left: 4px solid #10b981; padding: 15px; margin: 25px 0;">
            <p style="margin: 0; color: #065f46;"><strong>Secure PDF Attached:</strong> Please find your official payment receipt / invoice attached to this email.</p>
          </div>
          
          <p style="font-size: 14px; color: #4b5563;">Wishing you the best of health,<br><br><strong>The Billing Team</strong><br>Crescentia Clinic & Diagnostics</p>
          <hr style="border: 0; border-top: 1px solid #e5e7eb; margin: 30px 0 15px 0;">
          <p style="font-size: 11px; color: #9ca3af; text-align: center; margin: 0;">This is an automatically generated dispatch. Please do not reply to this email.</p>
        </div>
      </div>
    `;

    GmailApp.sendEmail(patientEmail, subject, plainBody, {
      htmlBody: richHtmlBody,
      attachments: [pdfBlob],
      name: "Crescentia Billing Desk"
    });

    return { success: true, message: "Invoice emailed successfully." };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
