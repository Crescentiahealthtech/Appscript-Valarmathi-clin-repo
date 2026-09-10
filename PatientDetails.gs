// ==========================================
// 🧠 SYSTEM CORE & ROUTER (Code.gs)
// ==========================================

function getPatientNameById(patientId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Patients');
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) if(data[i][0].toString().toUpperCase() === patientId.toUpperCase()) return data[i][2];
  } catch(e) {} return null;
}

// 🚀 UI BUG FIX: Clean Date Extraction & Stacked Booking Engine
function getPatientDashboardStats(patientId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Check Appointments Ledger
    const apptSheet = ss.getSheetByName('Appointments');
    let lastVisit = "None Recorded";
    let upcomingBookings = [];
    
    if (apptSheet) {
      const apptData = apptSheet.getDataRange().getValues();
      let today = new Date();
      today.setHours(0,0,0,0);
      
      for(let i = 1; i < apptData.length; i++) {
        if(apptData[i][1] === patientId) {
          let rawDate = apptData[i][3];
          let apptDateObj = (rawDate instanceof Date) ? new Date(rawDate) : new Date(rawDate);
          apptDateObj.setHours(0,0,0,0);
          
          // Force strict string casting to kill the GMT+0530 bug
          let dateStr = (rawDate instanceof Date) ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd MMM yyyy") : String(rawDate).substring(0,10);
          
          // Use our safe time formatter
          let rawTime = apptData[i][4];
          let timeStr = "";
          if(rawTime instanceof Date) {
            timeStr = Utilities.formatDate(rawTime, Session.getScriptTimeZone(), "hh:mm a").toUpperCase();
          } else {
            timeStr = String(rawTime).replace(/([0-9])(AM|PM)/i, "$1 $2").toUpperCase();
          }

          let status = apptData[i][6];
          
          if(apptDateObj >= today && (status === 'Booked' || status === 'Arrived' || status === 'In-Progress')) {
            upcomingBookings.push({ dateVal: apptDateObj, display: `${dateStr} • ${timeStr}` });
          }
          if(apptDateObj < today && status === 'Completed') {
            lastVisit = dateStr; 
          }
        }
      }
      upcomingBookings.sort((a,b) => a.dateVal - b.dateVal);
    }
    
    // 2. Check EMR Records for Prescribed Dates
    const emrSheet = ss.getSheetByName('EMR_Records');
    let emrNextVisit = "Awaiting Doctor's Update";
    let emrNextLab = "Awaiting Doctor's Update";
    
    if(emrSheet) {
        const emrData = emrSheet.getDataRange().getValues();
        // Traverse backward to find the absolute latest EMR record for this patient
        for(let i = emrData.length - 1; i >= 1; i--) {
            if(emrData[i][2] === patientId) {
                let planText = String(emrData[i][8]).toLowerCase();
                let lines = String(emrData[i][8]).split('\n');
                
                lines.forEach(line => {
                    let l = line.toLowerCase();
                    if(l.includes('follow-up') || l.includes('next visit')) {
                        emrNextVisit = line.replace(/follow-up|next visit|:/gi, '').trim() || "See EMR Plan";
                    }
                    if(l.includes('lab') || l.includes('blood test') || l.includes('investigation')) {
                        emrNextLab = line.replace(/lab visit|lab date|next lab|investigations|:/gi, '').trim() || "See EMR Plan";
                    }
                });
                break; // Stop at the most recent record
            }
        }
    }

    return { 
      lastVisit: lastVisit, 
      upcomingBookings: upcomingBookings.map(b => b.display), 
      emrNextVisit: emrNextVisit, 
      emrNextLab: emrNextLab 
    };

  } catch (e) { 
    return { lastVisit: "Error", upcomingBookings: [], emrNextVisit: "Error", emrNextLab: "Error" }; 
  }
}

// 🚀 BATCH OPTIMIZED AVAILABILITY SAVER (Fixes the Speed Issue)
