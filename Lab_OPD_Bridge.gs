// ============================================================================
// Lab_OPD_Bridge.gs  —  Crescentia HealthTech
// Connects the OPD consultation to the real Lab Test Catalog.
// ----------------------------------------------------------------------------
// TWO CHANGES TO HOW LABS WORK
//
// 1. OPD orders by TestID, not by test name.
//    Previously OPD sent names, createLabRequest tried to match them, and on a
//    miss fell back to 'MANUAL_MAP' — a single pseudo-test priced at zero. That
//    order could not be billed correctly and had no parameters to result
//    against. Ordering by ID removes the guesswork entirely.
//
// 2. Billing moves from order time to collection time.
//    An OPD order is now created as PENDING and stays there. The lab bills when
//    the sample is actually collected. A patient advised tests who leaves
//    without giving a sample therefore carries no bill.
//
// REQUIRES: LabSetup.gs, LabIntegrationEngine.gs
// NOTE: delete getOPDOrderableTests() from Doctor_Setup_Debug.gs — this is the
//       final version and a duplicate name would silently shadow it.
// ============================================================================

/**
 * FRONTEND ENTRY. The complete orderable catalog for the OPD order modal.
 * Panels, individual tests and packages — names only.
 * Price, sample type and department are deliberately omitted: those are the
 * lab's concern, not the consulting doctor's, and showing them adds noise to
 * a screen that already has too much.
 *
 * @return {{success, panels:[], tests:[], packages:[], total}}
 *   each item = { testId, testCode, testName, requiresConsent }
 */
function getOPDOrderableTests() {
  try {
    var res = getOrderableTests();
    if (!res || !res.success) {
      return { success: false, message: (res && res.message) || 'Catalog unavailable.',
               panels: [], tests: [], packages: [], total: 0 };
    }

    var panels = [], tests = [], packages = [];

    (res.tests || []).forEach(function (t) {
      var item = {
        testId: String(t.testId),
        testCode: String(t.testCode || ''),
        testName: String(t.testName || ''),
        requiresConsent: !!t.requiresConsent
      };
      if (t.testType === 'PANEL') panels.push(item);
      else if (t.testType === 'PACKAGE') packages.push(item);
      else tests.push(item);
    });

    var byName = function (a, b) { return a.testName.localeCompare(b.testName); };
    panels.sort(byName); tests.sort(byName); packages.sort(byName);

    return {
      success: true,
      panels: panels,
      tests: tests,
      packages: packages,
      total: panels.length + tests.length + packages.length,
      message: (panels.length + tests.length + packages.length) ? '' :
        'No active tests in the catalog. Add them under Labs > Test Catalog.'
    };
  } catch (e) {
    return { success: false, message: 'getOPDOrderableTests: ' + e.message,
             panels: [], tests: [], packages: [], total: 0 };
  }
}

/**
 * Called from saveOPEncounterScoped. Creates one lab order carrying real
 * catalog IDs and the attributed consulting doctor.
 *
 * @param {object} p { patientId, encounterId, doctorId, doctorName,
 *                     testIds:[], priority, clinicalNote }
 */
function createOPDLabOrder(p) {
  try {
    var ids = (p.testIds || []).filter(Boolean);
    if (!ids.length) return { success: true, skipped: true };

    return createLabRequest({
      patientId: String(p.patientId || ''),
      sourceModule: 'OPD',
      visitId: String(p.encounterId || ''),
      testIds: ids,                                  // real catalog IDs
      orderingDoctorId: String(p.doctorId || ''),
      orderingDoctorName: String(p.doctorName || ''),
      priority: String(p.priority || 'ROUTINE').toUpperCase(),
      clinicalNote: String(p.clinicalNote || '')
    });
  } catch (e) {
    return { success: false, message: 'createOPDLabOrder: ' + e.message };
  }
}

/**
 * FRONTEND ENTRY (Lab module). Collect the sample and bill in one action.
 * This is the collect-then-bill entry point for OPD and walk-in orders.
 *
 * payload = { orderId, samples:[{sampleType}], paymentMode, paidAmount,
 *             discountPercent }
 *
 * IP orders skip payment entry and post to the IP account as before.
 */
function collectAndBillLabSample(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!payload || !payload.orderId) return { success: false, message: 'Order ID required.' };

    var od = getLabOrderDetail(payload.orderId);
    if (!od.success) return { success: false, message: od.message };
    var order = od.order;

    var isIp = (order.source === 'IP_CASESHEET' || order.source === 'IP_NOTES');
    var alreadyBilled = false;
    try {
      var existing = getLabBill(order.orderId);
      alreadyBilled = !!(existing && existing.success);
    } catch (e) { /* no bill yet */ }

    // ---- bill first, so a payment failure doesn't leave an unbilled sample
    if (!alreadyBilled) {
      var billRes = generateLabBill({
        orderId: order.orderId,
        paymentMode: payload.paymentMode,
        paidAmount: payload.paidAmount,
        discountPercent: payload.discountPercent
      });
      if (!billRes.success) return { success: false, message: billRes.message, stage: 'BILLING' };
    }

    // ---- then collect
    var colRes = collectLabSample({ orderId: order.orderId, samples: payload.samples });
    if (!colRes.success) {
      return { success: false, stage: 'COLLECTION',
               message: 'Billed, but collection failed: ' + colRes.message +
                        ' The bill stands — collect the sample from the Lab queue.' };
    }

    return {
      success: true,
      message: 'Billed and collected. ' + (colRes.sampleIds || []).length + ' tube(s).',
      sampleIds: colRes.sampleIds
    };
  } catch (e) {
    return { success: false, message: 'collectAndBillLabSample: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}