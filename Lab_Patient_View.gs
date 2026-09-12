// =========================================================================
// 🔬 LAB RESULTS, WHERE THE CLINICIAN IS
// Crescentia HealthTech / CresRx
// -------------------------------------------------------------------------
// Reading a patient's results used to mean leaving whatever you were doing,
// opening the Lab module, searching for the patient by name or number, and
// finding your way back. In a consultation, on a ward round, and while
// writing a discharge summary, the results are part of the work in front of
// you — so they are fetched here, by patient, and shown over the top of
// wherever you already are (Lab_Results_Peek.html).
//
// Why not reuse searchLabRecords(): that one is a SEARCH. It matches the
// query against Patient_ID, name, mobile and WhatsApp with indexOf, which is
// right for a lab clerk hunting for "Kumar" and wrong at the bedside — a
// patient id that happens to appear inside another patient's phone number
// would put someone else's results on the screen. This matches Patient_ID
// exactly, and nothing else.
// =========================================================================

/** Order statuses whose results a clinician may read. */
var LPV_REPORTABLE = ['VERIFIED', 'REPORT_DISPATCHED', 'AMENDED'];

/**
 * FRONTEND ENTRY. Every reportable lab result for one patient.
 *
 * @param {string} patientId
 * @param {string} sessionToken
 * @param {Object} [opts]  { limit: number }  most recent N orders (default 40)
 * @return {{success:boolean, patientId:string, orders:Array, message:string}}
 *
 * Each order: { orderId, date, testNames, doctor, status, verifiedBy,
 *               abnormal:number, results:[{parameterName, value, unit, flag,
 *               refRangeText, critical:boolean}] }
 */
function getPatientLabResults(patientId, sessionToken, opts) {
  try {
    var sess = null;
    try { sess = dc_validateSession_(sessionToken); } catch (e) { sess = null; }
    if (!sess) {
      return { success: false, patientId: '', orders: [],
               message: 'Your session has expired. Please sign in again.' };
    }
    var role = String(sess.role || '').toLowerCase();
    if (role === 'patient') {
      return { success: false, patientId: '', orders: [], message: 'Not authorised.' };
    }

    var pid = String(patientId || '').trim().toUpperCase();
    if (!pid) {
      return { success: false, patientId: '', orders: [],
               message: 'A patient ID is required.' };
    }

    var limit = (opts && opts.limit) ? Number(opts.limit) : 40;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var oSheet = ss.getSheetByName(LAB.ORDERS);
    if (!oSheet || oSheet.getLastRow() < 2) {
      return { success: true, patientId: pid, orders: [], message: '' };
    }

    var oMap = labHeaderMap(oSheet);
    var oData = oSheet.getRange(2, 1, oSheet.getLastRow() - 1, oSheet.getLastColumn()).getValues();

    var byOrder = {}, ids = [];
    oData.forEach(function (r) {
      if (String(r[oMap['PatientID']] || '').trim().toUpperCase() !== pid) return;
      if (LPV_REPORTABLE.indexOf(String(r[oMap['OrderStatus']] || '')) === -1) return;
      var oid = String(r[oMap['OrderID']] || '');
      if (!oid || byOrder[oid]) return;
      byOrder[oid] = {
        orderId: oid,
        date: String(r[oMap['CreatedAt']] || ''),
        testNames: String(r[oMap['TestNames']] || ''),
        doctor: String(r[oMap['OrderingDoctorName']] || ''),
        source: String(r[oMap['SourceModule']] || ''),
        status: String(r[oMap['OrderStatus']] || ''),
        verifiedBy: '',
        verifiedAt: '',
        abnormal: 0,
        results: []
      };
      ids.push(oid);
    });

    if (!ids.length) return { success: true, patientId: pid, orders: [], message: '' };

    // Newest first, then trim, so the result sheet is only scanned for the
    // orders that will actually be returned.
    ids.sort(function (a, b) {
      return (byOrder[a].date < byOrder[b].date) ? 1 : -1;
    });
    var keep = {};
    ids.slice(0, limit).forEach(function (id) { keep[id] = true; });

    var rSheet = ss.getSheetByName(LAB.RESULTS);
    if (rSheet && rSheet.getLastRow() >= 2) {
      var rMap = labHeaderMap(rSheet);
      var rData = rSheet.getRange(2, 1, rSheet.getLastRow() - 1, rSheet.getLastColumn()).getValues();
      rData.forEach(function (r) {
        var oid = String(r[rMap['OrderID']] || '');
        if (!keep[oid]) return;
        var rec = byOrder[oid];

        var isLatest = (r[rMap['IsLatest']] === true ||
                        String(r[rMap['IsLatest']]).toUpperCase() === 'TRUE');
        var isDraft = (r[rMap['IsDraft']] === true ||
                       String(r[rMap['IsDraft']]).toUpperCase() === 'TRUE');
        // A draft is an unverified number. It is never shown to a clinician
        // here: acting on a value nobody has signed off is the harm this
        // whole screen exists to avoid.
        if (!isLatest || isDraft) return;

        var flag = String(r[rMap['Flag']] || '').toUpperCase();
        if (flag && flag !== 'N' && flag !== 'NORMAL') rec.abnormal++;

        rec.results.push({
          parameterName: String(r[rMap['ParameterName']] || ''),
          value: String(r[rMap['ResultValue']] || ''),
          unit: String(r[rMap['Unit']] || ''),
          flag: flag,
          critical: flag === 'C',
          refRangeText: String(r[rMap['RefRangeText']] || '')
        });
        if (!rec.verifiedBy) rec.verifiedBy = String(r[rMap['VerifiedBy']] || '');
        if (!rec.verifiedAt) rec.verifiedAt = String(r[rMap['VerifiedAt']] || '');
      });
    }

    var orders = ids.filter(function (id) { return keep[id]; })
                    .map(function (id) { return byOrder[id]; })
                    .filter(function (o) { return o.results.length > 0; });

    return { success: true, patientId: pid, orders: orders, message: '' };

  } catch (err) {
    return { success: false, patientId: '', orders: [],
             message: 'getPatientLabResults failed: ' + err.message };
  }
}
