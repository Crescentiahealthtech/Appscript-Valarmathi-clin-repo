/**
 * The ward's beds, as a JSON string in `data`.
 *
 * Guard inside the try, reply wrapped — see getAvailableTimeSlots in
 * Appointment.gs for why, and crescUnwrap in Shell_UX.html for the client
 * half. `data` stays the JSON string this has always returned, because
 * renderBedGrid parses it and its date fields are read as text.
 */
function getBedStatuses(sessionToken) {
  try {
    crescRequire_(sessionToken, 'ward.read');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = initializeBedsIfEmpty_(ss); // Auto-heals if empty

    const data = sheet.getDataRange().getValues();
    const results = [];

    for (let i = 1; i < data.length; i++) {
      results.push({
        bedId: data[i][0] ? data[i][0].toString() : "",
        ward: data[i][1] ? data[i][1].toString() : "",
        status: data[i][2] ? data[i][2].toString() : "Available",
        patientId: data[i][3] || "",
        patientName: data[i][4] || "",
        doa: data[i][5] || "",
        ipNumber: data[i][6] || ""
      });
    }
    return { success: true, data: JSON.stringify(results), message: '' };
  } catch (err) {
    return { success: false, data: JSON.stringify([]), message: cresc_reason_(err) };
  }
}

function mockBedDataGenerater() {
  crescEditorOnly_('mockBedDataGenerater');
  const beds = [];
  ['A101','A102','A103','A104','A105'].forEach(b => beds.push({ bedId: b, ward: 'A', status: 'Available' }));
  ['B201','B202','B203','B204','B205'].forEach(b => beds.push({ bedId: b, ward: 'B', status: 'Available' }));
  ['ICU-1','ICU-2','ICU-3'].forEach(b => beds.push({ bedId: b, ward: 'ICU', status: 'Available' }));
  
  beds[0].status = 'Occupied';
  beds[0].patientId = 'LMTVS0003';
  beds[0].patientName = 'THIYAGARAJAN N';
  beds[0].doa = '2026-05-18';
  beds[0].ipNumber = 'IP2605-001';

  beds[5].status = 'Cleaning';
  beds[10].status = 'Reserved';

  return JSON.stringify(beds);
}

/**
 * Sets one bed's status.
 *
 * Guard inside the try, reply wrapped — see getAvailableTimeSlots in
 * Appointment.gs.
 *
 * THE OLD SHAPE COULD NOT REPORT A FAILURE AT ALL. It returned a bare
 * boolean, and the only caller ignored it: bmUpdateStatus's success handler
 * was `function () { window.initBedMap(); }`, so `false` — the bed is not in
 * Master_Beds — repainted the map and said nothing, and the nurse watched
 * the bed they had just changed come back unchanged. A refusal did reach the
 * failure handler, and was reported as "Try again."
 */
function updateBedStatusInDB(bedId, newStatus, sessionToken) {
  try {
    crescRequire_(sessionToken, 'ward.admit');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Master_Beds");
    if (!sheet) {
      return { success: false, message: 'The Master_Beds sheet does not exist, so no ' +
                                        'bed status can be saved.' };
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === bedId.toString()) {
        sheet.getRange(i + 1, 3).setValue(newStatus);
        return { success: true, message: 'Bed ' + bedId + ' is now ' + newStatus + '.' };
      }
    }
    return { success: false, message: 'Bed ' + bedId + ' is not in Master_Beds, so its ' +
                                      'status was not changed.' };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  }
}
// SHARED INITIALIZER: Injects default schema if missing or empty
function initializeBedsIfEmpty_(ss) {
  let sheet = ss.getSheetByName("Master_Beds");
  if (!sheet) {
    sheet = ss.insertSheet("Master_Beds");
  }
  
  if (sheet.getLastRow() <= 1) {
    sheet.clear();
    sheet.appendRow(["Bed_ID", "Ward", "Status", "Patient_ID", "Patient_Name", "DOA", "IP_No"]);
    sheet.getRange("A1:G1").setFontWeight("bold");
    
    const defaultBeds = [
      ['A101', 'A', 'Available', '', '', '', ''],
      ['A102', 'A', 'Available', '', '', '', ''],
      ['A103', 'A', 'Available', '', '', '', ''],
      ['A104', 'A', 'Available', '', '', '', ''],
      ['A105', 'A', 'Available', '', '', '', ''],
      ['B201', 'B', 'Available', '', '', '', ''],
      ['B202', 'B', 'Available', '', '', '', ''],
      ['B203', 'B', 'Available', '', '', '', ''],
      ['B204', 'B', 'Available', '', '', '', ''],
      ['B205', 'B', 'Available', '', '', '', ''],
      ['ICU-1', 'ICU', 'Available', '', '', '', ''],
      ['ICU-2', 'ICU', 'Available', '', '', '', ''],
      ['ICU-3', 'ICU', 'Available', '', '', '', '']
    ];
    sheet.getRange(2, 1, defaultBeds.length, 7).setValues(defaultBeds);
  }
  return sheet;
}