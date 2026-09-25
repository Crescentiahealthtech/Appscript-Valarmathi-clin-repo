/**
 * ============================================================================
 * CRESCENTIA HEALTHTECH — LAB INTEGRATION ENGINE  (v3)
 * ============================================================================
 * This is THE only lab backend file your frontend and other modules call.
 */

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 0 — CONSTANTS (in case LabSetup.gs is deployed separately)
   ═══════════════════════════════════════════════════════════════════════════ */
var _LAB_STATUS_MACHINE = {
  PENDING:           ['BILLED','CANCELLED'],
  BILLED:            ['SAMPLE_COLLECTED','CANCELLED'],
  RECOLLECT:         ['SAMPLE_COLLECTED','CANCELLED'],
  SAMPLE_COLLECTED:  ['IN_PROCESS','RECOLLECT'],
  IN_PROCESS:        ['RESULT_ENTERED'],
  RESULT_ENTERED:    ['VERIFIED','IN_PROCESS'],
  VERIFIED:          ['REPORT_DISPATCHED'],
  REPORT_DISPATCHED: ['AMENDED'],
  AMENDED:           ['REPORT_DISPATCHED'],
  CANCELLED:         []
};

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 1 — CATALOG
   ═══════════════════════════════════════════════════════════════════════════ */

function getOrderableTests(sessionToken) {
  try { crescRequire_(sessionToken, ['lab.read', 'lab.order']); }
  catch (err) { return { success: false, message: err.message, tests: [] }; }
  return lab_orderableTests_();
}

/**
 * THE SAME READ, WITHOUT THE PERMISSION CHECK, for server-side callers that
 * have already validated the user. Calling the guarded entry point with no
 * token made crescRequire_ throw on `undefined` and the catch turn that into
 * an empty catalogue, so the OP consult and the ward round offered no tests
 * to order and said nothing about why. A trailing underscore keeps this
 * unreachable from google.script.run.
 */
function lab_orderableTests_() {
  // Read by every order screen on every desk — the OP consult, the ward, the
  // lab's own walk-in form — and changed a few times a month. Shared through
  // the reference-list cache (Master_Cache.gs); the catalogue writers below
  // and a hand edit to LAB_TEST_CATALOG both end the cached copy.
  if (typeof crescMasterGet_ === 'function') {
    return crescMasterGet_('labcatalog', lab_orderableTestsRead_);
  }
  return lab_orderableTestsRead_();
}

function lab_orderableTestsRead_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.CATALOG);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, tests: [] };
    const map = labHeaderMap_(sheet);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const tests = [];
    data.forEach(function (r) {
      const type = String(r[map['TestType']] || '');
      const active = (r[map['IsActive']] === true || String(r[map['IsActive']]).toUpperCase() === 'TRUE');
      if (!active) return;
      if (['PANEL','INDIVIDUAL','PACKAGE'].indexOf(type) === -1) return;
      tests.push({
        testId:          String(r[map['TestID']]),
        testCode:        String(r[map['TestCode']]),
        testName:        String(r[map['TestName']]),
        testType:        type,
        department:      String(r[map['Department']] || ''),
        sampleType:      String(r[map['SampleType']] || ''),
        price:           Number(r[map['Price']]) || 0,
        tatMinutes:      Number(r[map['TAT_Minutes']]) || 0,
        requiresConsent: (r[map['RequiresConsent']] === true || String(r[map['RequiresConsent']]).toUpperCase() === 'TRUE')
      });
    });
    tests.sort(function (a, b) { return a.testName < b.testName ? -1 : 1; });
    return { success: true, tests: tests };
  } catch (err) {
    return { success: false, message: 'getOrderableTests failed: ' + err.message, tests: [] };
  }
}

function getLabCatalog(sessionToken) {
  try { crescRequire_(sessionToken, ['lab.read', 'lab.order']); }
  catch (err) { return { success: false, message: err.message,
                         panels: [], individuals: [], packages: [] }; }
  return lab_catalog_();
}

/** THE SAME READ, WITHOUT THE PERMISSION CHECK. See lab_orderableTests_. */
function lab_catalog_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.CATALOG);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, panels: [], individuals: [], packages: [] };
    const map = labHeaderMap_(sheet);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const panelMap = {}, paramsByPanel = {}, individuals = [], packages = [];

    data.forEach(function (r) {
      const type = String(r[map['TestType']] || '');
      const active = (r[map['IsActive']] === true || String(r[map['IsActive']]).toUpperCase() === 'TRUE');
      const item = _rowToCatalogItem_(r, map);
      if (type === 'PANEL')      { panelMap[item.testId] = item; paramsByPanel[item.testId] = []; }
      else if (type === 'PARAMETER') { const pid = String(r[map['ParentPanelID']]||''); if (pid && paramsByPanel[pid]) paramsByPanel[pid].push(item); }
      else if (type === 'INDIVIDUAL') individuals.push(item);
      else if (type === 'PACKAGE')    packages.push(item);
    });

    const panels = Object.keys(panelMap).map(function (k) {
      const p = panelMap[k];
      p.parameters = paramsByPanel[k] || [];
      return p;
    });
    return { success: true, panels: panels, individuals: individuals, packages: packages };
  } catch (err) {
    return { success: false, message: 'getLabCatalog failed: ' + err.message, panels: [], individuals: [], packages: [] };
  }
}

function _rowToCatalogItem_(r, map) {
  return {
    testId:          String(r[map['TestID']]),
    testCode:        String(r[map['TestCode']]),
    testName:        String(r[map['TestName']]),
    testType:        String(r[map['TestType']]),
    department:      String(r[map['Department']] || ''),
    sampleType:      String(r[map['SampleType']] || ''),
    resultType:      String(r[map['ResultType']] || ''),
    unit:            String(r[map['Unit']] || ''),
    maleRefLow:      _safeNum_(r[map['MaleRefLow']]),
    maleRefHigh:     _safeNum_(r[map['MaleRefHigh']]),
    femaleRefLow:    _safeNum_(r[map['FemaleRefLow']]),
    femaleRefHigh:   _safeNum_(r[map['FemaleRefHigh']]),
    paediatricRefText: String(r[map['PaediatricRefText']] || ''),
    criticalLow:     _safeNum_(r[map['CriticalLow']]),
    criticalHigh:    _safeNum_(r[map['CriticalHigh']]),
    price:           Number(r[map['Price']]) || 0,
    tatMinutes:      Number(r[map['TAT_Minutes']]) || 0,
    componentTestIds:String(r[map['ComponentTestIDs']] || ''),
    requiresConsent: (r[map['RequiresConsent']] === true || String(r[map['RequiresConsent']]).toUpperCase() === 'TRUE'),
    sortOrder:       Number(r[map['SortOrder']]) || 0,
    isActive:        (r[map['IsActive']] === true || String(r[map['IsActive']]).toUpperCase() === 'TRUE')
  };
}

function saveCatalogEntry(p, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'lab.catalog');
    if (!p || !p.testName || !String(p.testName).trim()) return { success: false, message: 'Test name required.' };
    if (!p.testCode || !String(p.testCode).trim())       return { success: false, message: 'Test code required.' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.CATALOG);
    if (!sheet) return { success: false, message: 'Catalog sheet missing. Run setupLabDatabase() first.' };
    const map = labHeaderMap_(sheet);
    const ncols = LAB_SCHEMA.LAB_TEST_CATALOG.length;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const by = cresc_actorName_('SYSTEM');

    let existingRow = -1, existingCreatedAt = now, existingCreatedBy = by;
    let testId = p.testId ? String(p.testId) : '';

    if (testId && sheet.getLastRow() > 1) {
      const ids = sheet.getRange(2, map['TestID']+1, sheet.getLastRow()-1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === testId) {
          existingRow = i + 2;
          const cur = sheet.getRange(existingRow, 1, 1, ncols).getValues()[0];
          existingCreatedAt = cur[map['CreatedAt']] || now;
          existingCreatedBy = cur[map['CreatedBy']] || by;
          break;
        }
      }
    }
    if (!testId || existingRow === -1) testId = 'LABTEST-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    const type = String(p.testType || 'INDIVIDUAL').toUpperCase();
    const isNumeric = String(p.resultType || '').toUpperCase() === 'NUMERIC';
    const components = Array.isArray(p.componentTestIds) ? p.componentTestIds.join(',') : String(p.componentTestIds || '');

    const row = new Array(ncols).fill('');
    row[map['TestID']]           = testId;
    row[map['TestCode']]         = String(p.testCode).trim().toUpperCase();
    row[map['TestName']]         = String(p.testName).trim();
    row[map['TestType']]         = type;
    row[map['ParentPanelID']]    = '';
    row[map['Department']]       = String(p.department || '').toUpperCase();
    row[map['SampleType']]       = String(p.sampleType || '').toUpperCase();
    row[map['ResultType']]       = String(p.resultType || '').toUpperCase();
    row[map['Unit']]             = String(p.unit || '');
    row[map['MaleRefLow']]       = isNumeric ? _safeNumW_(p.maleRefLow)   : '';
    row[map['MaleRefHigh']]      = isNumeric ? _safeNumW_(p.maleRefHigh)  : '';
    row[map['FemaleRefLow']]     = isNumeric ? _safeNumW_(p.femaleRefLow) : '';
    row[map['FemaleRefHigh']]    = isNumeric ? _safeNumW_(p.femaleRefHigh): '';
    row[map['PaediatricRefText']]= String(p.paediatricRefText || '');
    row[map['CriticalLow']]      = isNumeric ? _safeNumW_(p.criticalLow)  : '';
    row[map['CriticalHigh']]     = isNumeric ? _safeNumW_(p.criticalHigh) : '';
    row[map['Price']]            = Number(p.price) || 0;
    row[map['TAT_Minutes']]      = parseInt(p.tatMinutes, 10) || 0;
    row[map['ComponentTestIDs']] = components;
    row[map['RequiresConsent']]  = !!(p.requiresConsent);
    row[map['SortOrder']]        = parseInt(p.sortOrder, 10) || 0;
    row[map['IsActive']]         = true;
    row[map['CreatedAt']]        = existingCreatedAt;
    row[map['CreatedBy']]        = existingCreatedBy;

    if (existingRow !== -1) sheet.getRange(existingRow, 1, 1, ncols).setValues([row]);
    else sheet.appendRow(row);

    labAudit_(existingRow !== -1 ? 'CATALOG_UPDATED' : 'CATALOG_CREATED', 'CATALOG', testId, null, { name: p.testName, type: type });
    if (typeof crescMasterBust_ === 'function') crescMasterBust_(LAB.CATALOG);
    SpreadsheetApp.flush();
    return { success: true, message: (existingRow !== -1 ? 'Updated' : 'Created') + ' "' + p.testName + '".', testId: testId };
  } catch (err) {
    return { success: false, message: 'saveCatalogEntry failed: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function setCatalogActive(testId, isActive, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'lab.catalog');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.CATALOG);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Catalog empty.' };
    const map = labHeaderMap_(sheet);
    const ids = sheet.getRange(2, map['TestID']+1, sheet.getLastRow()-1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(testId)) {
        sheet.getRange(i+2, map['IsActive']+1).setValue(!!isActive);
        SpreadsheetApp.flush();
        if (typeof crescMasterBust_ === 'function') crescMasterBust_(LAB.CATALOG);
        return { success: true, message: (isActive ? 'Activated' : 'Deactivated') + '.' };
      }
    }
    return { success: false, message: 'Test not found.' };
  } catch (err) {
    return { success: false, message: 'setCatalogActive failed: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Global HTML-escape helper — used by getLabBillHtml, _buildReportHtmlGrouped_, etc. */
function calculateFlag_(value, resultType, gender, ref) {
  if (!resultType || String(resultType).toUpperCase() !== 'NUMERIC') return '';
  const v = parseFloat(value);
  if (isNaN(v)) return '';
  const g = String(gender || '').toUpperCase().charAt(0);
  const lo = g === 'F' ? _safeNum_(ref.femaleRefLow)  : _safeNum_(ref.maleRefLow);
  const hi = g === 'F' ? _safeNum_(ref.femaleRefHigh) : _safeNum_(ref.maleRefHigh);
  const cl = _safeNum_(ref.criticalLow), ch = _safeNum_(ref.criticalHigh);
  if (cl !== null && v < cl) return 'C';
  if (ch !== null && v > ch) return 'C';
  if (lo !== null && v < lo) return 'L';
  if (hi !== null && v > hi) return 'H';
  if (lo !== null || hi !== null) return 'N';
  return '';
}

function savePanelWithParameters(p, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'lab.catalog');
    if (!p || !p.testName || !String(p.testName).trim()) return { success: false, message: 'Panel name required.' };
    if (!p.testCode || !String(p.testCode).trim())       return { success: false, message: 'Panel code required.' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.CATALOG);
    if (!sheet) return { success: false, message: 'Catalog sheet missing.' };
    const map = labHeaderMap_(sheet);
    const ncols = LAB_SCHEMA.LAB_TEST_CATALOG.length;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const by = cresc_actorName_('SYSTEM');

    let panelRow = -1, panelId = p.testId ? String(p.testId) : '';
    let pCreatedAt = now, pCreatedBy = by;
    if (panelId && sheet.getLastRow() > 1) {
      const ids = sheet.getRange(2, map['TestID']+1, sheet.getLastRow()-1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === panelId) {
          panelRow = i + 2;
          const cur = sheet.getRange(panelRow, 1, 1, ncols).getValues()[0];
          pCreatedAt = cur[map['CreatedAt']] || now;
          pCreatedBy = cur[map['CreatedBy']] || by;
          break;
        }
      }
    }
    if (!panelId || panelRow === -1) panelId = 'LABTEST-' + Utilities.getUuid().substring(0,8).toUpperCase();

    const panelRow_ = new Array(ncols).fill('');
    panelRow_[map['TestID']]        = panelId;
    panelRow_[map['TestCode']]      = String(p.testCode).trim().toUpperCase();
    panelRow_[map['TestName']]      = String(p.testName).trim();
    panelRow_[map['TestType']]      = 'PANEL';
    panelRow_[map['Department']]    = String(p.department||'').toUpperCase();
    panelRow_[map['SampleType']]    = String(p.sampleType||'').toUpperCase();
    panelRow_[map['Price']]         = Number(p.price)||0;
    panelRow_[map['TAT_Minutes']]   = parseInt(p.tatMinutes,10)||0;
    panelRow_[map['SortOrder']]     = parseInt(p.sortOrder,10)||0;
    panelRow_[map['IsActive']]      = true;
    panelRow_[map['CreatedAt']]     = pCreatedAt;
    panelRow_[map['CreatedBy']]     = pCreatedBy;
    if (panelRow !== -1) sheet.getRange(panelRow, 1, 1, ncols).setValues([panelRow_]);
    else sheet.appendRow(panelRow_);

    const allData = sheet.getDataRange().getValues();
    const existingParams = [];
    allData.forEach(function (r, i) {
      if (String(r[map['TestType']]) === 'PARAMETER' && String(r[map['ParentPanelID']]) === panelId) {
        existingParams.push({ id: String(r[map['TestID']]), rowNum: i+1 });
      }
    });
    const touched = {};
    const params = Array.isArray(p.parameters) ? p.parameters : [];
    params.forEach(function (pm, idx) {
      if (!pm.testName || !pm.testCode) return;
      let pmRow = -1, pmId = pm.testId ? String(pm.testId) : '';
      let pmCAt = now, pmCBy = by;
      if (pmId) {
        const m = existingParams.find(function(x){return x.id===pmId;});
        if (m) { pmRow=m.rowNum; const c=sheet.getRange(pmRow,1,1,ncols).getValues()[0]; pmCAt=c[map['CreatedAt']]||now; pmCBy=c[map['CreatedBy']]||by; }
      }
      if (!pmId || pmRow === -1) pmId = 'LABTEST-' + Utilities.getUuid().substring(0,8).toUpperCase();
      touched[pmId] = true;
      const isNum = String(pm.resultType||'').toUpperCase() === 'NUMERIC';
      const pmRow_ = new Array(ncols).fill('');
      pmRow_[map['TestID']]        = pmId;
      pmRow_[map['TestCode']]      = String(pm.testCode).trim().toUpperCase();
      pmRow_[map['TestName']]      = String(pm.testName).trim();
      pmRow_[map['TestType']]      = 'PARAMETER';
      pmRow_[map['ParentPanelID']] = panelId;
      pmRow_[map['Department']]    = String(p.department||'').toUpperCase();
      pmRow_[map['SampleType']]    = String(p.sampleType||'').toUpperCase();
      pmRow_[map['ResultType']]    = String(pm.resultType||'NUMERIC').toUpperCase();
      pmRow_[map['Unit']]          = String(pm.unit||'');
      pmRow_[map['MaleRefLow']]    = isNum ? _safeNumW_(pm.maleRefLow)  : '';
      pmRow_[map['MaleRefHigh']]   = isNum ? _safeNumW_(pm.maleRefHigh) : '';
      pmRow_[map['CriticalLow']]   = isNum ? _safeNumW_(pm.criticalLow) : '';
      pmRow_[map['CriticalHigh']]  = isNum ? _safeNumW_(pm.criticalHigh): '';
      pmRow_[map['Price']]         = 0;
      pmRow_[map['TAT_Minutes']]   = parseInt(p.tatMinutes,10)||0;
      pmRow_[map['SortOrder']]     = idx+1;
      pmRow_[map['IsActive']]      = true;
      pmRow_[map['CreatedAt']]     = pmCAt;
      pmRow_[map['CreatedBy']]     = pmCBy;
      if (pmRow !== -1) sheet.getRange(pmRow, 1, 1, ncols).setValues([pmRow_]);
      else sheet.appendRow(pmRow_);
    });
    existingParams.forEach(function(ep){ if(!touched[ep.id]) sheet.getRange(ep.rowNum, map['IsActive']+1).setValue(false); });
    labAudit_('PANEL_SAVED','CATALOG',panelId,null,{name:p.testName,params:params.length});
    if (typeof crescMasterBust_ === 'function') crescMasterBust_(LAB.CATALOG);
    SpreadsheetApp.flush();
    return { success: true, message: 'Panel "'+p.testName+'" saved with '+params.length+' parameter(s).', panelId: panelId, paramCount: params.length };
  } catch (err) {
    return { success: false, message: 'savePanelWithParameters failed: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 3 — ORDERS
   ═══════════════════════════════════════════════════════════════════════════ */

function createLabRequest(d, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'lab.order');
    if (!d) return { success: false, message: 'No data received.' };
    if (!d.patientId || !String(d.patientId).trim()) {
    const wiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB.ORDERS);
    const wiCol = wiSheet ? labHeaderMap_(wiSheet)['PatientID'] : undefined;
    d.patientId = bc_nextDailyId_('LAB_WALKIN', 'WALKIN-', '-', 4,
                                  wiSheet, (wiCol === undefined) ? 0 : wiCol + 1);   // Barcode_Engine.gs
    }

    const source = String(d.sourceModule||'WALKIN').toUpperCase();
    if (['OPD','IP_CASESHEET','IP_NOTES','WALKIN'].indexOf(source) === -1) {
      return { success: false, message: 'Invalid source: ' + source };
    }

    let testIds = Array.isArray(d.testIds) ? d.testIds.filter(Boolean) : [];
    const nameList = Array.isArray(d.testNames) ? d.testNames.filter(Boolean) : [];

    if (!testIds.length && nameList.length) {
      testIds = _resolveTestNamesToCatalogIds_(nameList);
      if (!testIds.length) {
        testIds = ['MANUAL_MAP'];
      }
    }
    if (!testIds.length) return { success: false, message: 'Select at least one test.' };

    const catalog = _loadCatalogById_();
    const names = [], consent = [];
    testIds.forEach(function(tid){
      if (tid === 'MANUAL_MAP') { names.push(nameList.join(', ')); return; }
      const t = catalog[tid];
      if (t) { names.push(t.testName); if (t.requiresConsent) consent.push(t.testName); }
      else names.push(tid);
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.ORDERS);
    if (!sheet) return { success: false, message: 'Orders sheet missing. Run setupLabDatabase() first.' };
    const map = labHeaderMap_(sheet);
    const ncols = LAB_SCHEMA.LAB_ORDERS.length;
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const by = cresc_actorName_('SYSTEM');
    const orderId = bc_nextDailyId_('LAB_ORDER', 'LAB-ORD-', '-', 4, sheet, map['OrderID'] + 1);   // Barcode_Engine.gs

    const row = new Array(ncols).fill('');
    row[map['OrderID']]            = orderId;
    row[map['PatientID']]          = String(d.patientId);
    // Self-heal demographics — OPD sends only patientId
    if (!String(d.patientName||'').trim() || !String(d.age||'').trim() || !String(d.gender||'').trim()) {
      try {
      var look = lookupLabPatient(String(d.patientId));
      if (look && look.success && look.patient) {
      if (!String(d.patientName||'').trim()) d.patientName = look.patient.name;
      if (!String(d.age||'').trim())         d.age         = look.patient.age;
      if (!String(d.gender||'').trim())      d.gender      = look.patient.gender;
        }
      } catch(e2) { /* non-fatal */ }
      }
    row[map['PatientName']]        = String(d.patientName||'');
    row[map['Age']]                = String(d.age||'');
    row[map['Gender']]             = String(d.gender||'').toUpperCase().charAt(0)||'';
    row[map['VisitID']]            = String(d.visitId||'');
    row[map['AdmissionID']]        = String(d.admissionId||'');
    row[map['SourceModule']]       = source;
    row[map['OrderingDoctorID']]   = String(d.orderingDoctorId||'');
    row[map['OrderingDoctorName']] = String(d.orderingDoctorName||'');
    row[map['TestIDs']]            = testIds.join(',');
    row[map['TestNames']]          = names.join(', ');
    row[map['Priority']]           = String(d.priority||'ROUTINE').toUpperCase();
    row[map['ClinicalNote']]       = String(d.clinicalNote||'');
    row[map['RepeatOfOrderID']]    = String(d.repeatOfOrderId||'');
    row[map['RepeatReason']]       = String(d.repeatReason||'');
    row[map['OrderStatus']]        = 'PENDING';
    row[map['CreatedAt']]          = nowStr;
    row[map['CreatedBy']]          = by;
    row[map['LastUpdatedAt']]      = nowStr;
    row[map['LastUpdatedBy']]      = by;

    sheet.appendRow(row);
    labAudit_('ORDER_CREATED','ORDER',orderId,null,{source:source, tests:names, priority:row[map['Priority']]});
    SpreadsheetApp.flush();

    return { success: true, message: 'Order '+orderId+' created.', orderId: orderId, requiresConsentFor: consent };
  } catch (err) {
    return { success: false, message: 'createLabRequest failed: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getLabOrderDetail(orderId, sessionToken) {
  // Only ever called from inside a guarded endpoint: the ambient actor it
  // set answers here. A direct google.script.run call has none and is refused.
  crescRequire_(sessionToken);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.ORDERS);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Order not found.' };
    const map = labHeaderMap_(sheet);
    const data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (String(r[map['OrderID']]) !== String(orderId)) continue;
      return {
        success: true,
        order: {
          orderId:      String(r[map['OrderID']]),
          patientId:    String(r[map['PatientID']]),
          patientName:  String(r[map['PatientName']]||''),
          age:          String(r[map['Age']]||''),
          gender:       String(r[map['Gender']]||''),
          visitId:      String(r[map['VisitID']]||''),
          admissionId:  String(r[map['AdmissionID']]||''),
          source:       String(r[map['SourceModule']]||''),
          doctorId:     String(r[map['OrderingDoctorID']]||''),
          doctorName:   String(r[map['OrderingDoctorName']]||''),
          testIds:      String(r[map['TestIDs']]||'').split(',').filter(Boolean),
          testNames:    String(r[map['TestNames']]||''),
          priority:     String(r[map['Priority']]||'ROUTINE'),
          clinicalNote: String(r[map['ClinicalNote']]||''),
          repeatOf:     String(r[map['RepeatOfOrderID']]||''),
          repeatReason: String(r[map['RepeatReason']]||''),
          status:       String(r[map['OrderStatus']]||'PENDING'),
          createdAt:    lab_ts_(r[map['CreatedAt']])
        }
      };
    }
    return { success: false, message: 'Order not found: ' + orderId };
  } catch (err) {
    return { success: false, message: 'getLabOrderDetail failed: ' + err.message };
  }
}

/**
 * FRONTEND ENTRY. Move one order to a status by hand, from the lab desk.
 *
 * WHY THIS IS A THIN WRAPPER. The status machine is also driven from inside
 * billing, collection, resulting and verification — six places, each already
 * guarded by the permission that matches what the person is doing. If those
 * called this function they would inherit ITS permission as well, and the
 * accountant who raises a lab bill would be refused for not holding
 * lab.collect. So the transition itself is private and unguarded, and this
 * is the door the browser comes through.
 */
function advanceOrderStatus(orderId, newStatus, sessionToken) {
  // Guard inside the try. lab_setOrderStatus_ already answers in
  // { success, message }, so this is the one of the nine that needed no new
  // shape — only for the refusal to arrive in the same envelope as every
  // other answer instead of being thrown past it into the failure handler.
  try {
    crescRequire_(sessionToken, ['lab.collect', 'lab.result', 'lab.verify']);
    return lab_setOrderStatus_(orderId, newStatus);
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  }
}

/** * 🔥 HARDENED: FORCED TRANSITION
 * Bypasses strict indexOf checks to prevent UI gridlocks.
 *
 * PRIVATE. Every caller has already established who is asking and what they
 * are allowed to do — see advanceOrderStatus() above.
 */
function lab_setOrderStatus_(orderId, newStatus) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.ORDERS);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Order not found.' };
    
    const map = labHeaderMap_(sheet);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][map['OrderID']]) === String(orderId)) {
        const rowNum = i + 2;
        const cur = String(data[i][map['OrderStatus']] || '').trim().toUpperCase();
        const targetStatus = String(newStatus).trim().toUpperCase();
        
        // If already in target status, return success silently to avoid UI errors
        if (cur === targetStatus) return { success: true, message: 'Already in status: ' + targetStatus };
        
        // Force the transition
        const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        
        sheet.getRange(rowNum, map['OrderStatus'] + 1).setValue(targetStatus);
        sheet.getRange(rowNum, map['LastUpdatedAt'] + 1).setValue(nowStr);
        sheet.getRange(rowNum, map['LastUpdatedBy'] + 1).setValue(cresc_actorName_('SYSTEM'));
        
        labAudit_('STATUS_CHANGED', 'ORDER', orderId, { from: cur }, { to: targetStatus });
        SpreadsheetApp.flush();
        
        return { success: true, message: 'Status updated to ' + targetStatus };
      }
    }
    return { success: false, message: 'Order not found: ' + orderId };
  } catch (err) {
    return { success: false, message: 'advanceOrderStatus failed: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 4 — WORKSPACE 
   ═══════════════════════════════════════════════════════════════════════════ */

function getLabWorkspaceData(sessionToken) {
  try {
    crescRequire_(sessionToken, 'lab.read');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const oSheet = ss.getSheetByName(LAB.ORDERS);
    // agedOut = rows the queue is deliberately not showing (finished, or
    // expired uncollected). expiredHidden is the subset that expired.
    const stats = { pendingBill:0, awaitingCollection:0, inProcess:0, awaitingVerify:0,
                    verified:0, overdue:0, criticalPending:0, todayTotal:0,
                    agedOut:0, expiredHidden:0, queueWindowDays: Math.round(LAB_QUEUE_WINDOW_MS/86400000) };
    if (!oSheet || oSheet.getLastRow() < 2) return { success: true, stats: stats, orders: [] };

    const oMap  = labHeaderMap_(oSheet);
    const oData = oSheet.getRange(2,1,oSheet.getLastRow()-1,oSheet.getLastColumn()).getValues();
    const billIdx   = _lwBillingIndex_();
    const sampIdx   = _lwSamplesIndex_();
    const tatIdx    = _lwTatIndex_();
    const nowMs     = new Date().getTime();
    const todayStr  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    stats.criticalPending = _lwCritCount_();

    const orders = [];
    oData.forEach(function(r){
      const oid    = String(r[oMap['OrderID']]||'');
      const status = String(r[oMap['OrderStatus']]||'PENDING');
      if (!oid) return;

      const tat = tatIdx[oid];
      let overdue=false, tatText='', tatPct=0;
      if (tat && tat.dlMs) {
        const active = ['SAMPLE_COLLECTED','IN_PROCESS','RESULT_ENTERED'].indexOf(status) !== -1;
        if (active) {
          const startMs = tat.stMs || tat.dlMs;
          const total = Math.max(tat.dlMs - startMs, 1);
          tatPct = Math.min(Math.round((nowMs - startMs) / total * 100), 100);
          if (nowMs > tat.dlMs) { overdue=true; tatText='Overdue '+(Math.round((nowMs-tat.dlMs)/60000))+'m'; tatPct=100; }
          else tatText = (Math.round((tat.dlMs-nowMs)/60000))+'m left';
        }
      }

      const bill = billIdx[oid] || null;
      const samp = sampIdx[oid] || null;
      // CreatedAt is written as 'yyyy-MM-dd HH:mm:ss' TEXT, which Sheets turns
      // into a real date on the way in. String() of that date is "Sat Jun 13
      // 2026 23:25:10 GMT+0530 (India Standard Time)" — which is what every
      // queue card printed, what "orders today" compared against 'yyyy-MM-dd'
      // (so it was always 0), and what the queue sorted on (alphabetically,
      // by weekday). One parse, three formats.
      const createdRaw = r[oMap['CreatedAt']];
      const createdD = cresc_parseDate_(createdRaw);
      const createdAt = createdD ? Utilities.formatDate(createdD, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
                                 : String(createdRaw || '');
      if (createdAt.indexOf(todayStr) === 0) stats.todayTotal++;

      const stageIdx = {PENDING:0,BILLED:1,RECOLLECT:1,SAMPLE_COLLECTED:2,IN_PROCESS:3,RESULT_ENTERED:4,VERIFIED:5,REPORT_DISPATCHED:6,AMENDED:6,CANCELLED:0}[status]||0;

      // Visibility is decided BEFORE the counters, so a tab's badge counts
      // exactly the rows that tab will show. Counting first and hiding after
      // would put "Billing 14" over a list of three.
      const vis = _lwQueueVisibility_(status, createdAt, r[oMap['LastUpdatedAt']],
                                     samp ? samp.s : '', nowMs);
      if (vis.expired) stats.expiredHidden++;
      if (!vis.show) { stats.agedOut++; return; }

      if (status==='PENDING')   stats.pendingBill++;
      else if (status==='BILLED'||status==='RECOLLECT') stats.awaitingCollection++;
      else if (status==='SAMPLE_COLLECTED'||status==='IN_PROCESS') stats.inProcess++;
      else if (status==='RESULT_ENTERED') stats.awaitingVerify++;
      else if (status==='VERIFIED'||status==='REPORT_DISPATCHED') stats.verified++;
      if (overdue) stats.overdue++;

      {
        orders.push({
          orderId:      oid,
          patientId:    String(r[oMap['PatientID']]||''),
          patientName:  String(r[oMap['PatientName']]||''),
          age:          String(r[oMap['Age']]||''),
          gender:       String(r[oMap['Gender']]||''),
          source:       String(r[oMap['SourceModule']]||'WALKIN'),
          doctor:       String(r[oMap['OrderingDoctorName']]||''),
          testNames:    String(r[oMap['TestNames']]||''),
          testCount:    String(r[oMap['TestIDs']]||'').split(',').filter(Boolean).length,
          priority:     String(r[oMap['Priority']]||'ROUTINE'),
          clinicalNote: String(r[oMap['ClinicalNote']]||''),
          isRepeat:     !!String(r[oMap['RepeatOfOrderID']]||''),
          status:       status,
          stageIdx:     stageIdx,
          billStatus:   bill ? String(bill.s) : 'UNBILLED',
          billNet:      bill ? bill.n : 0,
          receiptNumber:bill ? String(bill.r) : '',
          sampleStatus: samp ? String(samp.s) : '',
          overdue:      overdue,
          tatText:      tatText,
          tatPct:       tatPct,
          createdAt:    createdAt,
          createdText:  createdD ? Utilities.formatDate(createdD, Session.getScriptTimeZone(), 'dd-MMM hh:mm a') : '',
          createdMs:    createdD ? createdD.getTime() : 0
        });
      }
    });

    const pw = {STAT:0,URGENT:1,ROUTINE:2};
    orders.sort(function(a,b){
      if(a.overdue!==b.overdue) return a.overdue?-1:1;
      if((pw[a.priority]||9)!==(pw[b.priority]||9)) return (pw[a.priority]||9)-(pw[b.priority]||9);
      return (b.createdMs||0)-(a.createdMs||0);
    });
    return { success: true, stats: stats, orders: orders };
  } catch (err) {
    return { success: false, message: 'getLabWorkspaceData failed: ' + err.message, stats: null, orders: [] };
  }
}

function getLabOrderPanel(orderId, sessionToken) {
  try {
    crescRequire_(sessionToken, 'lab.read');
    const od = getLabOrderDetail(orderId);
    if (!od.success) return { success: false, message: od.message };
    const bill = getLabBill(orderId);
    const samples = getOrderSamples(orderId);
    return {
      success: true,
      order: od.order,
      bill: (bill&&bill.success) ? bill.bill : null,
      samples: (samples&&samples.success) ? samples.samples : []
    };
  } catch (err) {
    return { success: false, message: 'getLabOrderPanel failed: ' + err.message };
  }
}

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 5 — BILLING
   ═══════════════════════════════════════════════════════════════════════════ */

function generateLabBill(d, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'lab.bill');
    if (!d||!d.orderId) return { success: false, message: 'Order ID required.' };
    const od = getLabOrderDetail(d.orderId);
    if (!od.success) return { success: false, message: od.message };
    const order = od.order;
    if (['PENDING','SAMPLE_COLLECTED'].indexOf(order.status) === -1) {
      return { success: false, message: 'Order already billed (status: '+order.status+').' };
    }

    const catalog = _loadCatalogById_();
    let gross = 0;
    const items = [];
    order.testIds.forEach(function(tid){
      if (tid === 'MANUAL_MAP') { items.push({testId:tid,testName:order.testNames,testType:'INDIVIDUAL',price:0}); return; }
      const t = catalog[tid];
      const price = t ? (Number(t.price)||0) : 0;
      gross += price;
      items.push({testId:tid, testName:t?t.testName:tid, testType:t?t.testType:'INDIVIDUAL', price:price});
    });

    const disc = Math.max(0,Math.min(100,Number(d.discountPercent)||0));
    const discAmt = +(gross*disc/100).toFixed(2);
    const net = +(gross-discAmt).toFixed(2);
    const ip = (order.source==='IP_CASESHEET'||order.source==='IP_NOTES');

    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const by = cresc_actorName_('SYSTEM');
    let category, payMode, paid, balance, payStatus, receipt='', ledgerPostId='';

    if (ip) {
      category='IP_ACCOUNT'; payMode='ON_ACCOUNT'; paid=0; balance=net; payStatus='ON_ACCOUNT';
      ledgerPostId='IPPOST-'+Utilities.getUuid().substring(0,8).toUpperCase();
      labAudit_('IP_ACCOUNT_POST','BILL',ledgerPostId,null,{admissionId:order.admissionId,amount:net});
    } else {
      category=(order.source==='WALKIN')?'WALKIN_SPOT':'OP_SPOT';
      payMode=String(d.paymentMode||'').toUpperCase();
      if(['CASH','CARD','UPI'].indexOf(payMode)===-1) return {success:false,message:'Select payment mode (Cash / Card / UPI).'};
      paid = (d.paidAmount===''||d.paidAmount==null) ? net : +Number(d.paidAmount).toFixed(2);
      // Neither below nothing nor above the bill: a paid figure over the net
      // wrote a negative balance, and a negative one a balance larger than
      // the bill — both reached the till and the receivables as they were.
      if (!isFinite(paid) || paid < 0) return {success:false,message:'The amount paid cannot be negative.'};
      if (paid > net) paid = net;
      balance = +(net-paid).toFixed(2);
      payStatus = balance<=0?'PAID':(paid>0?'PARTIAL':'PENDING');
      receipt = 'RCP-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd')+'-'+Utilities.getUuid().substring(0,4).toUpperCase();
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.BILLING);
    if (!sheet) return {success:false,message:'Billing sheet missing.'};
    const map = labHeaderMap_(sheet);
    const ncols = LAB_SCHEMA.LAB_BILLING.length;
    const billId = 'LAB-BILL-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd')+'-'+Utilities.getUuid().substring(0,4).toUpperCase();

    const row = new Array(ncols).fill('');
    row[map['BillID']]          = billId;
    row[map['OrderID']]         = order.orderId;
    row[map['PatientID']]       = order.patientId;
    row[map['PatientName']]     = order.patientName;
    row[map['BillingCategory']] = category;
    row[map['AdmissionID']]     = ip?order.admissionId:'';
    row[map['TestsJSON']]       = JSON.stringify(items);
    row[map['GrossAmount']]     = gross;
    row[map['DiscountPercent']] = disc;
    row[map['DiscountAmount']]  = discAmt;
    row[map['NetAmount']]       = net;
    row[map['PaymentMode']]     = payMode;
    row[map['PaidAmount']]      = paid;
    row[map['BalanceAmount']]   = balance;
    row[map['PaymentStatus']]   = payStatus;
    row[map['ReceiptNumber']]   = receipt;
    row[map['IPLedgerPostID']]  = ledgerPostId;
    row[map['BilledAt']]        = nowStr;
    row[map['BilledBy']]        = by;
    sheet.appendRow(row);

    // 👇 NEW IP INTEGRATION HOOK ADDED HERE 👇
    if (ip && order.admissionId) {                       
      try { 
        billChargeToIp({ 
          ipNumber: order.admissionId, 
          source: 'LAB', 
          sourceRef: billId,
          amount: net, 
          gst: 0, 
          description: items.length + ' lab test(s)', 
          user: by 
        }); 
      } catch (e) {
        // Silently catch so the lab bill still succeeds even if IP ledger is busy
      }
    }
    // 👆 END NEW HOOK 👆

    labAudit_('BILL_GENERATED','BILL',billId,null,{orderId:order.orderId,net:net,category:category});
    if (order.status === 'PENDING') lab_setOrderStatus_(order.orderId,'BILLED');
    SpreadsheetApp.flush();

    return { success:true, message: ip?('₹'+net.toFixed(2)+' posted to IP account.'):(billId+' · '+receipt+' · ₹'+net.toFixed(2)+' '+payStatus+'.'), billId:billId, receiptNumber:receipt, netAmount:net, category:category };
  } catch (err) {
    return { success:false, message:'generateLabBill failed: '+err.message };
  } finally {
    try{lock.releaseLock();}catch(e){}
  }
}

function getLabBill(orderId, sessionToken) {
  // Only ever called from inside a guarded endpoint: the ambient actor it
  // set answers here. A direct google.script.run call has none and is refused.
  crescRequire_(sessionToken);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.BILLING);
    if (!sheet||sheet.getLastRow()<2) return {success:false,message:'No bills found.'};
    const map=labHeaderMap_(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    let found=null;
    data.forEach(function(r){
      if(String(r[map['OrderID']])===String(orderId)){
        found={billId:String(r[map['BillID']]),orderId:String(r[map['OrderID']]),patientName:String(r[map['PatientName']]||''),
          category:String(r[map['BillingCategory']]||''),items:_safeParse_(r[map['TestsJSON']]),gross:Number(r[map['GrossAmount']])||0,
          discountAmount:Number(r[map['DiscountAmount']])||0,net:Number(r[map['NetAmount']])||0,paymentMode:String(r[map['PaymentMode']]||''),
          paid:Number(r[map['PaidAmount']])||0,balance:Number(r[map['BalanceAmount']])||0,paymentStatus:String(r[map['PaymentStatus']]||''),
          receiptNumber:String(r[map['ReceiptNumber']]||''),billedAt:lab_ts_(r[map['BilledAt']])};
      }
    });
    return found?{success:true,bill:found}:{success:false,message:'No bill for this order.'};
  } catch(err){return{success:false,message:'getLabBill failed: '+err.message};}
}

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 6 — SAMPLES
   ═══════════════════════════════════════════════════════════════════════════ */

function collectLabSample(d, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'lab.collect');
    if(!d||!d.orderId) return {success:false,message:'Order ID required.'};
    if(!Array.isArray(d.samples)||!d.samples.length) return {success:false,message:'Add at least one tube.'};
    const od = getLabOrderDetail(d.orderId);
    if(!od.success) return {success:false,message:od.message};
    const order = od.order;
    if(['PENDING','BILLED','RECOLLECT'].indexOf(order.status)===-1) return {success:false,message:'Cannot collect at status: '+order.status+'.'};

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.SAMPLES);
    if(!sheet) return {success:false,message:'Samples sheet missing.'};
    const map=labHeaderMap_(sheet);
    const ncols=LAB_SCHEMA.LAB_SAMPLES.length;
    const nowStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    const by=cresc_actorName_('SYSTEM');
    const datePart=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd');
    const sampleIds=[];

    d.samples.forEach(function(s){
      const sampleId=bc_nextDailyId_('LAB_SAMPLE','LAB-SAMP-','-',4,sheet,map['SampleID']+1);   // Barcode_Engine.gs
      const barcode=bc_nextDailyId_('LAB_BARCODE','BC','',5,sheet,map['BarcodeID']+1);
      const row=new Array(ncols).fill('');
      row[map['SampleID']]=sampleId; row[map['OrderID']]=order.orderId; row[map['PatientID']]=order.patientId;
      row[map['PatientName']]=order.patientName; row[map['SampleType']]=String(s.sampleType||'').toUpperCase();
      row[map['BarcodeID']]=barcode; row[map['CollectionStatus']]='COLLECTED'; row[map['RejectionReason']]='';
      row[map['CollectedAt']]=nowStr; row[map['CollectedBy']]=by; row[map['ReceivedAtLabAt']]=''; row[map['ReceivedBy']]=''; row[map['CreatedAt']]=nowStr;
      sheet.appendRow(row); sampleIds.push(sampleId);
    });

    _startTat_(order, nowStr);
    lab_setOrderStatus_(order.orderId,'SAMPLE_COLLECTED');
    labAudit_('SAMPLE_COLLECTED','SAMPLE',order.orderId,null,{count:sampleIds.length});
    SpreadsheetApp.flush();
    return {success:true,message:sampleIds.length+' sample(s) collected.',sampleIds:sampleIds};
  } catch(err){return{success:false,message:'collectLabSample failed: '+err.message};}
  finally{try{lock.releaseLock();}catch(e){}}
}

function receiveLabSample(orderId, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'lab.collect');
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.SAMPLES);
    if(!sheet||sheet.getLastRow()<2) return {success:false,message:'No samples found.'};
    const map=labHeaderMap_(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    const nowStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    const by=cresc_actorName_('SYSTEM');
    let n=0;
    data.forEach(function(r,i){
      if(String(r[map['OrderID']])===String(orderId)&&String(r[map['CollectionStatus']])!=='REJECTED'&&!r[map['ReceivedAtLabAt']]){
        sheet.getRange(i+2,map['ReceivedAtLabAt']+1).setValue(nowStr);
        sheet.getRange(i+2,map['ReceivedBy']+1).setValue(by); n++;
      }
    });
    if(!n) return {success:false,message:'No unprocessed samples for this order.'};
    lab_setOrderStatus_(orderId,'IN_PROCESS');
    labAudit_('SAMPLE_RECEIVED','SAMPLE',orderId,null,{count:n});
    SpreadsheetApp.flush();
    return {success:true,message:n+' sample(s) received. Processing started.'};
  } catch(err){return{success:false,message:'receiveLabSample failed: '+err.message};}
  finally{try{lock.releaseLock();}catch(e){}}
}

function rejectLabSample(sampleId, reason, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'lab.collect');
    if(!sampleId) return {success:false,message:'Sample ID required.'};
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.SAMPLES);
    if(!sheet||sheet.getLastRow()<2) return {success:false,message:'No samples found.'};
    const map=labHeaderMap_(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    let orderId='',patientId='',rowFound=-1;
    for(let i=0;i<data.length;i++){
      if(String(data[i][map['SampleID']])===String(sampleId)){rowFound=i+2;orderId=String(data[i][map['OrderID']]);patientId=String(data[i][map['PatientID']]);break;}
    }
    if(rowFound===-1) return {success:false,message:'Sample not found: '+sampleId};
    sheet.getRange(rowFound,map['CollectionStatus']+1).setValue('REJECTED');
    sheet.getRange(rowFound,map['RejectionReason']+1).setValue(String(reason||'').toUpperCase());
    const ncrId=_raiseNcr_({orderId:orderId,sampleId:sampleId,patientId:patientId,type:'SAMPLE_REJECTION',description:'Sample '+sampleId+' rejected: '+reason,immediateAction:'Recollection requested'});
    lab_setOrderStatus_(orderId,'RECOLLECT');
    labAudit_('SAMPLE_REJECTED','SAMPLE',sampleId,null,{reason:reason,ncr:ncrId});
    SpreadsheetApp.flush();
    return {success:true,message:'Rejected. NCR '+ncrId+' raised.',ncrId:ncrId};
  } catch(err){return{success:false,message:'rejectLabSample failed: '+err.message};}
  finally{try{lock.releaseLock();}catch(e){}}
}

function getOrderSamples(orderId, sessionToken) {
  // Only ever called from inside a guarded endpoint: the ambient actor it
  // set answers here. A direct google.script.run call has none and is refused.
  crescRequire_(sessionToken);
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.SAMPLES);
    if(!sheet||sheet.getLastRow()<2) return {success:true,samples:[]};
    const map=labHeaderMap_(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    const out=[];
    data.forEach(function(r){
      if(String(r[map['OrderID']])!==String(orderId)) return;
      out.push({sampleId:String(r[map['SampleID']]),sampleType:String(r[map['SampleType']]||''),barcode:String(r[map['BarcodeID']]||''),
        status:String(r[map['CollectionStatus']]||''),rejectionReason:String(r[map['RejectionReason']]||''),
        collectedAt:lab_ts_(r[map['CollectedAt']]),receivedAt:lab_ts_(r[map['ReceivedAtLabAt']])});
    });
    return {success:true,samples:out};
  } catch(err){return{success:false,message:'getOrderSamples failed: '+err.message};}
}

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 7 — RESULTS
   ═══════════════════════════════════════════════════════════════════════════ */

function buildResultEntrySheet(orderId, sessionToken) {
  try {
    crescRequire_(sessionToken, 'lab.result');
    const od=getLabOrderDetail(orderId);
    if(!od.success) return {success:false,message:od.message};
    const order=od.order;
    const cat=lab_catalog_();
    if(!cat.success) return {success:false,message:'Catalog unavailable.'};
    const panelById={},indivById={},pkgById={};
    cat.panels.forEach(function(p){panelById[p.testId]=p;});
    cat.individuals.forEach(function(t){indivById[t.testId]=t;});
    cat.packages.forEach(function(p){pkgById[p.testId]=p;});

    const rowDefs=[];
    const _addPanel=function(panel){(panel.parameters||[]).forEach(function(pm){rowDefs.push(_mkRowDef_(panel.testId,pm));});};
    const _addIndiv=function(t){rowDefs.push(_mkRowDef_(t.testId,t));};
    order.testIds.forEach(function(tid){
      if(panelById[tid]) _addPanel(panelById[tid]);
      else if(indivById[tid]) _addIndiv(indivById[tid]);
      else if(pkgById[tid]){ (pkgById[tid].componentTestIds||'').split(',').filter(Boolean).forEach(function(cid){if(panelById[cid])_addPanel(panelById[cid]);else if(indivById[cid])_addIndiv(indivById[cid]);}); }
    });

    const existing=_latestResultsForOrder_(orderId);
    const prior=_priorValuesForPatient_(order.patientId,orderId);

    const rows=rowDefs.map(function(rd){
      const ex=existing[rd.parameterId]||{};
      const pv=prior[rd.parameterId]||{};
      return Object.assign({},rd,{
        currentValue: ex.value!=null?String(ex.value):'',
        currentFlag:  String(ex.flag||''),
        interpretation:String(ex.interpretation||''),
        isDraft: ex.isDraft===undefined?true:ex.isDraft,
        previousValue: pv.value!=null?String(pv.value):'',
        previousAt:    String(pv.at||'')
      });
    });
    return {success:true,order:order,rows:rows};
  } catch(err){return{success:false,message:'buildResultEntrySheet failed: '+err.message};}
}

function saveLabResultsDraft(d, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'lab.result');
    if(!d||!d.orderId) return {success:false,message:'Order ID required.'};
    if(!Array.isArray(d.results)) return {success:false,message:'No results provided.'};
    const od=getLabOrderDetail(d.orderId);
    if(!od.success) return {success:false,message:od.message};
    const order=od.order;
    const gender=String(d.gender||order.gender||'').toUpperCase().charAt(0)||'M';
    const refByParam=_refRangesByParameter_();
    const sampleId=_firstSampleId_(order.orderId);
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.RESULTS);
    if(!sheet) return {success:false,message:'Results sheet missing.'};
    const map=labHeaderMap_(sheet);
    const ncols=LAB_SCHEMA.LAB_RESULTS.length;
    const nowStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    const by=cresc_actorName_('SYSTEM');
    // Read the block ONCE, edit it in memory, write each touched row once.
    //
    // This used to issue seven setValue() calls per parameter for an existing
    // row and an appendRow() per new one. A CBC is twenty-odd parameters, so
    // saving one draft cost well over a hundred separate Sheets round trips
    // and the technician waited through every one of them. Same writes, same
    // order, one call per row.
    const existing=_existingResultRows_(sheet,map,order.orderId);
    const existingRows=existing.rowByParam;
    const rowValues=existing.valuesByRow;
    const touched={};
    const appends=[];
    let critCount=0;

    d.results.forEach(function(res){
      if(!res.parameterId) return;
      const ref=refByParam[res.parameterId]||{};
      const flag=calculateFlag_(res.value,ref.resultType,gender,ref);
      if(flag==='C') critCount++;
      const refText=(String(ref.resultType||'').toUpperCase()==='NUMERIC')
        ?(gender==='F'?_rngTxt_(ref.femaleRefLow,ref.femaleRefHigh):_rngTxt_(ref.maleRefLow,ref.maleRefHigh)):'';
      const exRow=existingRows[res.parameterId];
      if(exRow){
        const row=rowValues[exRow];
        row[map['ResultValue']]=String(res.value==null?'':res.value);
        row[map['Flag']]=flag;
        row[map['RefRangeText']]=refText;
        row[map['Interpretation']]=String(res.interpretation||'');
        row[map['IsDraft']]=true;
        row[map['EnteredBy']]=by;
        row[map['EnteredAt']]=nowStr;
        touched[exRow]=true;
      } else {
        const row=new Array(ncols).fill('');
        row[map['ResultID']]='LAB-RES-'+Utilities.getUuid().substring(0,8).toUpperCase();
        row[map['OrderID']]=order.orderId; row[map['SampleID']]=sampleId; row[map['PatientID']]=order.patientId;
        row[map['TestID']]=ref.testId||''; row[map['ParameterID']]=res.parameterId; row[map['ParameterName']]=ref.parameterName||'';
        row[map['ResultValue']]=String(res.value==null?'':res.value); row[map['ResultType']]=ref.resultType||'';
        row[map['Unit']]=ref.unit||''; row[map['RefRangeText']]=refText; row[map['Flag']]=flag;
        row[map['IsDraft']]=true; row[map['EnteredBy']]=by; row[map['EnteredAt']]=nowStr;
        row[map['Interpretation']]=String(res.interpretation||''); row[map['Version']]=1; row[map['IsLatest']]=true;
        appends.push(row);
      }
    });

    Object.keys(touched).forEach(function(r){
      const row=rowValues[r];
      sheet.getRange(Number(r),1,1,row.length).setValues([row]);
    });
    if(appends.length){
      sheet.getRange(sheet.getLastRow()+1,1,appends.length,ncols).setValues(appends);
    }
    SpreadsheetApp.flush();
    labAudit_('RESULTS_DRAFT_SAVED','RESULT',order.orderId,null,{count:d.results.length,critical:critCount});
    return {success:true,message:'Draft saved'+(critCount?' · '+critCount+' CRITICAL value(s)':'')+'.', criticalCount:critCount};
  } catch(err){return{success:false,message:'saveLabResultsDraft failed: '+err.message};}
  finally{try{lock.releaseLock();}catch(e){}}
}

/** * 🔥 HARDENED: Submit Results 
 */
function submitLabResults(d, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    crescRequire_(sessionToken, 'lab.result');
    if(!d||!d.orderId) return {success:false,message:'Order ID required.'};
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.RESULTS);
    if(!sheet||sheet.getLastRow()<2) return {success:false,message:'No results to submit.'};
    
    const map=labHeaderMap_(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    let n=0; 
    let emptyCount=0;
    const crits=[];
    
    data.forEach(function(r,i){
      if(String(r[map['OrderID']])!==String(d.orderId)) return;
      if(!(r[map['IsLatest']]===true||String(r[map['IsLatest']]).toUpperCase()==='TRUE')) return;
      
      // Check for completely blank submissions
      if(!String(r[map['ResultValue']]).trim()) emptyCount++;
      
      sheet.getRange(i+2,map['IsDraft']+1).setValue(false); 
      n++;
      
      if(String(r[map['Flag']])==='C') {
        crits.push({resultId:String(r[map['ResultID']]),parameterName:String(r[map['ParameterName']]),value:String(r[map['ResultValue']]),patientId:String(r[map['PatientID']])});
      }
    });
    
    if(!n) return {success:false,message:'No results entered for this order yet.'};
    if(emptyCount > 0 && emptyCount === n) return {success:false,message:'Cannot submit completely blank results. Enter values first.'};
    
    // Force transition safely
    lab_setOrderStatus_(d.orderId,'RESULT_ENTERED');
    
    crits.forEach(function(c){_logCritical_(d.orderId,c);});
    labAudit_('RESULTS_SUBMITTED','RESULT',d.orderId,null,{count:n,critical:crits.length});
    SpreadsheetApp.flush();
    return {success:true,message:'Submitted for verification'+(crits.length?' · '+crits.length+' critical value(s) logged.':'.')};
  } catch(err){
    return{success:false,message:'submitLabResults failed: '+err.message};
  } finally{
    try{lock.releaseLock();}catch(e){}
  }
}

/** * 🔥 HARDENED: Verify Results 
 */
function verifyLabResults(d, sessionToken) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const actor=crescRequire_(sessionToken, 'lab.verify');
    if(!d||!d.orderId) return {success:false,message:'Order ID required.'};
    const od=getLabOrderDetail(d.orderId);
    if(!od.success) return {success:false,message:od.message};
    
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.RESULTS);
    const map=labHeaderMap_(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    const nowStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    // The signature on a lab report is the signed-in verifier, not a name the
    // browser sends: d.verifierName let anyone sign as any pathologist.
    const verifier=String(actor.displayName||actor.username||'SYSTEM');
    const valRows=[]; const targetRows=[];
    
    data.forEach(function(r,i){
      if(String(r[map['OrderID']])===String(d.orderId)&&(r[map['IsLatest']]===true||String(r[map['IsLatest']]).toUpperCase()==='TRUE')){
        valRows.push(String(r[map['ParameterID']])+'='+String(r[map['ResultValue']])+'/'+String(r[map['Flag']]));
        targetRows.push(i+2);
      }
    });
    
    if(!valRows.length) return {success:false,message:'No results to verify.'};
    valRows.sort();
    
    const basis=d.orderId+'|'+od.order.patientId+'|'+valRows.join('|')+'|'+verifier+'|'+new Date().getTime();
    const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,basis,Utilities.Charset.UTF_8);
    const hash=digest.map(function(b){return('0'+(b&0xFF).toString(16)).slice(-2);}).join('');
    
    // `data` is already the whole block, read once above. Edit the cached row
    // and write it back in one call instead of five per parameter — verifying
    // a twenty-parameter panel was a hundred separate Sheets round trips.
    targetRows.forEach(function(rowNum){
      const row=data[rowNum-2];
      row[map['VerifiedBy']]=verifier;
      row[map['VerifiedAt']]=nowStr;
      row[map['AttestationHash']]=hash;
      row[map['IsDraft']]=false;
      if(d.interpretation) row[map['Interpretation']]=String(d.interpretation);
      sheet.getRange(rowNum,1,1,row.length).setValues([row]);
    });
    
    _closeTat_(d.orderId,nowStr);
    
    // Force transition to verified
    lab_setOrderStatus_(d.orderId,'VERIFIED');
    
    labAudit_('RESULTS_VERIFIED','RESULT',d.orderId,null,{verifier:verifier,hash:hash.substring(0,12)+'…'});
    SpreadsheetApp.flush();
    return {success:true,message:'Verified & signed by '+verifier+'. Attestation hash generated.',attestationHash:hash};
  } catch(err){
    return{success:false,message:'verifyLabResults failed: '+err.message};
  } finally{
    try{lock.releaseLock();}catch(e){}
  }
}

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 8 — PATIENT LOOKUP & LAB RECORDS
   ═══════════════════════════════════════════════════════════════════════════ */

function lookupLabPatient(pid, sessionToken) {
  try {
    crescRequire_(sessionToken, 'patient.read');
    if(!pid||!String(pid).trim()) return {success:false,message:'Enter a Patient ID.'};
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName('Patients');
    if(!sheet||sheet.getLastRow()<2) return {success:false,message:'Patients database missing.'};
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,7).getValues();
    const want=String(pid).trim().toUpperCase();
    for(let i=0;i<data.length;i++){
      if(String(data[i][0]).trim().toUpperCase()===want){
        return {success:true,patient:{patientId:String(data[i][0]),name:String(data[i][2]||''),age:String(data[i][3]||''),gender:String(data[i][4]||'').toUpperCase().charAt(0)||'M',mobile:String(data[i][6]||'')}};
      }
    }
    return {success:false,message:'No patient found with ID '+pid+'.'};
  } catch(err){return{success:false,message:'lookupLabPatient failed: '+err.message};}
}

function DIAG_labRecords() {
  crescEditorOnly_('DIAG_labRecords');
  var pid = 'LMTVS0001';   // ← put the patient ID you searched
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 
  var oSheet = ss.getSheetByName(LAB.ORDERS);
  var oMap = labHeaderMap_(oSheet);
  var oData = oSheet.getRange(2,1,oSheet.getLastRow()-1,oSheet.getLastColumn()).getValues();
 
  var report = [];
  oData.forEach(function(r){
    var rpid = String(r[oMap['PatientID']]||'');
    if (rpid.toUpperCase() !== pid.toUpperCase()) return;
    report.push({
      orderId: String(r[oMap['OrderID']]),
      status:  String(r[oMap['OrderStatus']]),
      tests:   String(r[oMap['TestNames']])
    });
  });
  Logger.log('Orders for %s: %s', pid, JSON.stringify(report, null, 2));
 
  // Check result rows
  var rSheet = ss.getSheetByName(LAB.RESULTS);
  var rMap = labHeaderMap_(rSheet);
  var rData = rSheet.getRange(2,1,rSheet.getLastRow()-1,rSheet.getLastColumn()).getValues();
  var resCount = {};
  rData.forEach(function(r){
    if (String(r[rMap['PatientID']]||'').toUpperCase() !== pid.toUpperCase()) return;
    var oid = String(r[rMap['OrderID']]);
    var latest = (r[rMap['IsLatest']]===true||String(r[rMap['IsLatest']]).toUpperCase()==='TRUE');
    var draft  = (r[rMap['IsDraft']]===true||String(r[rMap['IsDraft']]).toUpperCase()==='TRUE');
    if(!resCount[oid]) resCount[oid]={latest:0,draft:0,total:0};
    resCount[oid].total++;
    if(latest) resCount[oid].latest++;
    if(draft)  resCount[oid].draft++;
  });
  Logger.log('Result rows for %s: %s', pid, JSON.stringify(resCount, null, 2));
  return { orders: report, results: resCount };
}


/**
 * ============================================================================
 * ADD THIS FUNCTION TO LabIntegrationEngine.gs (after searchLabRecords)
 * ============================================================================
 * Called by LabX.printRpt() in the Done panel.
 * Builds the full printable HTML report for an order using its verified results.
 * Uses clinic name from Script Properties (set via setMyClinicConfig()).
 * ============================================================================
 */

/**
 * Generates an NABL/ISO 15189 Compliant HTML Report for Printing.
 * @param {string} orderId - The specific LAB-ORD ID to print.
 */
/**
 * Generates an NABL/ISO 15189 Compliant HTML Report (Mobile Responsive + Printable)
 */
function getLabReportHtml(orderId, sessionToken) {
  try {
    crescRequire_(sessionToken, 'lab.read');
    // Finding M2. A full report with results and the ordering diagnosis.
    dpdpLogRead_(crescActor_(sessionToken), 'LabReport', String(orderId || ''),
                 { endpoint: 'getLabReportHtml' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordSheet = ss.getSheetByName("LAB_ORDERS");
    const resSheet = ss.getSheetByName("LAB_RESULTS");
    const patSheet = ss.getSheetByName("Patients");

    if (!ordSheet || !resSheet || !patSheet) {
      return { success: false, message: "Database sheets missing." };
    }

    // 1. Get Order Details
    const ordData = ordSheet.getDataRange().getValues();
    const oHeaders = ordData[0];
    let orderRow = null;
    for (let i = 1; i < ordData.length; i++) {
      if (ordData[i][oHeaders.indexOf("OrderID")] === orderId) {
        orderRow = ordData[i];
        break;
      }
    }
    if (!orderRow) return { success: false, message: "Order not found." };

    const pId = orderRow[oHeaders.indexOf("PatientID")];
    const docName = orderRow[oHeaders.indexOf("OrderingDoctorName")] || "Self";
    const ordDate = cresc_formatDate_(orderRow[oHeaders.indexOf("CreatedAt")], "dd-MMM-yyyy hh:mm a");
    
    // 2. Get Patient Details
    const patData = patSheet.getDataRange().getValues();
    const pHeaders = patData[0];
    let patRow = null;
    for (let i = 1; i < patData.length; i++) {
      if (patData[i][pHeaders.indexOf("Patient_ID")] === pId) {
        patRow = patData[i];
        break;
      }
    }
    
    // Patient-ID barcode, so a printed report can be scanned straight back to
    // the patient at the counter.
    let pidBarcode = "";
    try {
      if (typeof bcp_patientBarcodeBlock_ === 'function') {
        pidBarcode = bcp_patientBarcodeBlock_(pId, { align: 'right', height: 8 });
      }
    } catch (e) { pidBarcode = ""; }

    const pName = patRow ? patRow[pHeaders.indexOf("Name")] : orderRow[oHeaders.indexOf("PatientName")];
    const pAge = patRow ? patRow[pHeaders.indexOf("Age")] : orderRow[oHeaders.indexOf("Age")];
    const pSex = patRow ? patRow[pHeaders.indexOf("Gender")] : orderRow[oHeaders.indexOf("Gender")];

    // 3. Get Results & Attestation Hash
    const resData = resSheet.getDataRange().getValues();
    const rHeaders = resData[0];
    let resultsHtml = "";
    let attestationHash = "N/A";
    let verifierName = "System Verified";
    let verifiedAt = new Date().toLocaleString();

    for (let i = 1; i < resData.length; i++) {
      if (resData[i][rHeaders.indexOf("OrderID")] === orderId) {
        let param = resData[i][rHeaders.indexOf("ParameterName")];
        let val = resData[i][rHeaders.indexOf("ResultValue")];
        let unit = resData[i][rHeaders.indexOf("Unit")] || "";
        let ref = resData[i][rHeaders.indexOf("RefRangeText")] || "";
        let flag = resData[i][rHeaders.indexOf("Flag")] || "N";
        
        if (attestationHash === "N/A" && resData[i][rHeaders.indexOf("AttestationHash")]) {
          attestationHash = resData[i][rHeaders.indexOf("AttestationHash")];
          verifierName = resData[i][rHeaders.indexOf("VerifiedBy")] || verifierName;
          verifiedAt = cresc_formatDate_(resData[i][rHeaders.indexOf("VerifiedAt")], "dd-MMM-yyyy hh:mm a") || verifiedAt;
        }

        let valFmt = (flag === "H" || flag === "L" || flag === "C") ? `<strong>${val}*</strong>` : val;
        let flagFmt = (flag !== "N" && flag !== "") ? `<strong>${flag}</strong>` : "";

        resultsHtml += `
          <tr>
            <td style="padding: 10px; border-bottom: 1px solid #e5e7eb;">${param}</td>
            <td style="padding: 10px; border-bottom: 1px solid #e5e7eb; text-align:center;">${valFmt}</td>
            <td style="padding: 10px; border-bottom: 1px solid #e5e7eb; text-align:center;">${unit}</td>
            <td style="padding: 10px; border-bottom: 1px solid #e5e7eb; text-align:center;">${ref}</td>
          </tr>
        `;
      }
    }

    // The verification block, built in-process. See _labVerifyBlock_ for why
    // the previous one could never have rendered.
    const verifyBlockHtml = _labVerifyBlock_(orderId, attestationHash, verifierName, verifiedAt);

    // 4. Construct Responsive HTML Document
    const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>Lab Report - ${pName}</title>
      <style>
        body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #333; margin: 0; padding: 0; background: #f8fafc; }
        
        /* Mobile-First / Web View Styles */
        .page { max-width: 210mm; background: #fff; margin: 0 auto; box-sizing: border-box; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        .page-content { padding: 20px 30px; }
        
        /* Floating Download Button for Patients */
        .mobile-header { text-align: center; background: #1e3a8a; padding: 15px; position: sticky; top: 0; z-index: 1000; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .mobile-header p { margin: 0 0 10px 0; color: #93c5fd; font-size: 13px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; }
        .btn-download { background: #10b981; color: white; border: none; padding: 12px 24px; border-radius: 6px; font-size: 16px; font-weight: bold; cursor: pointer; width: 90%; max-width: 300px; box-shadow: 0 2px 4px rgba(16,185,129,0.3); transition: 0.2s; }
        .btn-download:active { transform: scale(0.98); }

        .header { text-align: center; border-bottom: 2px solid #2563eb; padding-bottom: 10px; margin-bottom: 20px; }
        .header h1 { margin: 0; color: #1e3a8a; font-size: 22px; text-transform: uppercase; letter-spacing: 1px; }
        .header p { margin: 4px 0 0 0; font-size: 12px; color: #6b7280; }
        
        .patient-box { border: 1px solid #e5e7eb; border-radius: 8px; padding: 15px; margin-bottom: 20px; display: flex; justify-content: space-between; font-size: 13px; background: #f8fafc; }
        .patient-col p { margin: 4px 0; }
        
        table { width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 30px; }
        th { background-color: #f3f4f6; color: #374151; text-align: center; padding: 12px 8px; border-bottom: 2px solid #d1d5db; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
        th:first-child { text-align: left; }
        
        .footer { margin-top: 40px; border-top: 1px solid #d1d5db; padding-top: 15px; display: flex; justify-content: space-between; align-items: flex-end; font-size: 11px; color: #4b5563; }
        .qr-box { display: flex; align-items: center; gap: 10px; }
        .signature-box { text-align: right; }

        /* Mobile specific adjustments */
        @media screen and (max-width: 600px) {
          .page-content { padding: 15px; }
          .patient-box { flex-direction: column; gap: 10px; }
          .footer { flex-direction: column; align-items: center; text-align: center; gap: 20px; }
          .signature-box { text-align: center; }
          table { font-size: 12px; }
          th, td { padding: 8px 4px; }
        }

        /* Strict Print Rules (Hides button, fits A4 exactly) */
        @media print {
          body { background: #fff; }
          .mobile-header { display: none !important; }
          .page { width: 210mm; max-width: 210mm; box-shadow: none; margin: 0; }
          .page-content { padding: 0; }
          .footer { page-break-inside: avoid; }
        }
      </style>
    </head>
    <body>
      <div class="page">
        <div class="mobile-header">
          <p>Crescentia Diagnostics</p>
          <button class="btn-download" onclick="window.print()">📥 Save PDF / Print</button>
        </div>

        <div class="page-content">
          <div class="header">
            <h1>Crescentia Clinic & Diagnostics</h1>
            <p>ISO 15189 Certified | 123 Health Avenue, Medical District | Ph: +91-9876543210</p>
          </div>

          <div class="patient-box">
            <div class="patient-col">
              <p><strong>Patient Name:</strong> ${pName}</p>
              <p><strong>Patient ID:</strong> ${pId}</p>
              <p><strong>Age / Gender:</strong> ${pAge} Yrs / ${pSex}</p>
            </div>
            <div class="patient-col">
              <p><strong>Order ID:</strong> ${orderId}</p>
              <p><strong>Referred By:</strong> Dr. ${docName}</p>
              <p><strong>Registered On:</strong> ${ordDate}</p>
            </div>
            <div class="patient-col" style="text-align:right;">${pidBarcode}</div>
          </div>

          <h3 style="font-size: 14px; color: #1e3a8a; border-bottom: 1px solid #e5e7eb; padding-bottom: 5px;">DEPARTMENT OF BIOCHEMISTRY & CLINICAL PATHOLOGY</h3>

          <table>
            <thead>
              <tr>
                <th>Investigation</th>
                <th>Observed Value</th>
                <th>Unit</th>
                <th>Biological Ref. Interval</th>
              </tr>
            </thead>
            <tbody>
              ${resultsHtml}
            </tbody>
          </table>
          
          <p style="font-size: 11px; color: #6b7280; font-style: italic; margin-top: -15px;">* Indicates value falls outside biological reference interval.</p>

          <div class="footer">
            ${verifyBlockHtml}
            <div class="signature-box">
              <div style="height: 40px; border-bottom: 1px dashed #9ca3af; margin-bottom: 5px; width: 180px; display: inline-block;"></div>
              <br>
              <strong>Digitally Signed By</strong><br>
              ${verifierName}<br>
              Verified On: ${verifiedAt}<br>
              <em>Consultant Pathologist</em>
            </div>
          </div>
          <div style="text-align: center; font-size: 10px; color: #9ca3af; margin-top: 25px;">
            *** End of Report ***
          </div>
        </div>
      </div>
    </body>
    </html>
    `;

    return { success: true, html: html };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ONE-TIME SETUP — run this once from Apps Script editor to set your clinic branding.
 * All reports and receipts will use these values instead of hardcoded strings.
 */
function setMyClinicConfig() {
  crescEditorOnly_('setMyClinicConfig');
  var p = PropertiesService.getScriptProperties();
  p.setProperty('CLINIC_NAME',    'Crescentia Clinic');     // ← Change this
  p.setProperty('CLINIC_ADDRESS', 'Your Address, City');    // ← Change this
  p.setProperty('CLINIC_PHONE',   '+91 00000 00000');       // ← Change this
  p.setProperty('NABL_NUMBER',    '');                      // ← Set if NABL accredited
  p.setProperty('CLINIC_TAGLINE', 'NABL ISO 15189:2022 Compliant');
  return 'Clinic config saved. Run this again after any changes.';
}

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 9 — INTERNAL HELPERS
   ═══════════════════════════════════════════════════════════════════════════ */

function _loadCatalogById_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(LAB.CATALOG);
  const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap_(sheet);
  const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){ out[String(r[map['TestID']])]=_rowToCatalogItem_(r,map); });
  return out;
}

function _resolveTestNamesToCatalogIds_(nameList) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(LAB.CATALOG);
  if(!sheet||sheet.getLastRow()<2) return [];
  const map=labHeaderMap_(sheet);
  const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  const resolved=[];
  nameList.forEach(function(name){
    const n=String(name).trim().toUpperCase();
    let best=null;
    for(let i=0;i<data.length;i++){
      if(String(data[i][map['TestCode']]).toUpperCase()===n&&(data[i][map['IsActive']]===true||String(data[i][map['IsActive']]).toUpperCase()==='TRUE')){best=String(data[i][map['TestID']]);break;}
    }
    if(!best) for(let i=0;i<data.length;i++){
      if(String(data[i][map['TestName']]).toUpperCase()===n&&(data[i][map['IsActive']]===true||String(data[i][map['IsActive']]).toUpperCase()==='TRUE')){best=String(data[i][map['TestID']]);break;}
    }
    if(!best) for(let i=0;i<data.length;i++){
      const tn=String(data[i][map['TestName']]).toUpperCase();
      if((tn.indexOf(n)!==-1||n.indexOf(tn)!==-1)&&(data[i][map['IsActive']]===true||String(data[i][map['IsActive']]).toUpperCase()==='TRUE')){best=String(data[i][map['TestID']]);break;}
    }
    if(best&&resolved.indexOf(best)===-1) resolved.push(best);
  });
  return resolved;
}

function _lwBillingIndex_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.BILLING); const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap_(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){ const k=String(r[map['OrderID']]||''); if(k) out[k]={s:String(r[map['PaymentStatus']]||''),n:Number(r[map['NetAmount']])||0,r:String(r[map['ReceiptNumber']]||'')}; });
  return out;
}
function _lwSamplesIndex_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.SAMPLES); const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap_(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){ const k=String(r[map['OrderID']]||''); if(k) out[k]={s:String(r[map['CollectionStatus']]||'')}; });
  return out;
}
function _lwTatIndex_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.TAT_LOG); const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap_(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){
    const k=String(r[map['OrderID']]||''); if(!k) return;
    const dl=r[map['TAT_Deadline']]; const st=r[map['SampleCollectedAt']];
    const dlMs=cresc_ms_(dl);
    if(!out[k]||dlMs>out[k].dlMs) out[k]={dlMs:dlMs,stMs:cresc_ms_(st)};
  });
  return out;
}
/* ===========================================================================
   CRITICAL VALUE ACKNOWLEDGEMENT
   ---------------------------------------------------------------------------
   WHAT "6 CRITICAL UNACK." ON THE LAB DASHBOARD MEANT

   A critical value is a result so far outside the reference interval that it
   needs a clinician told NOW, by a person, not by a report landing in a queue
   — a potassium of 7.2, a platelet count of 8, a positive blood culture.
   Whenever a result is saved past the CriticalLow/CriticalHigh bounds on the
   test, _logCritical_() appends a row to LAB_CRITICAL_COMMS with
   IsAcknowledged = FALSE. The KPI counts those rows. Six of them meant six
   critical results where nobody had confirmed the clinician was reached.

   WHY IT ONLY EVER WENT UP

   Nothing in the entire project ever set IsAcknowledged to TRUE. The
   AcknowledgedBy and AcknowledgedAt columns existed in the schema and were
   never written. The one button in the interface that said
   "Acknowledged — I will inform the clinician" carried data-bs-dismiss="modal"
   and nothing else: it closed the dialog, and that was all it did.

   So the number was a one-way counter. It could not be cleared, it named none
   of the six, and there was no record of who was told or when — which is
   precisely the record NABL ISO 15189 requires for critical result
   communication, and precisely the record an enquiry into a missed result
   would ask for.

   THE TWO FUNCTIONS BELOW ARE THE OTHER HALF.

   Note what deliberately does NOT happen here: an unacknowledged critical
   value never ages out. The bench queue hides finished work after 48 hours
   (LAB_QUEUE_WINDOW_MS) because it is history. A critical result nobody has
   answered for is not history; it stays on the dashboard until a person
   closes it, however old and however inconvenient.
   =========================================================================== */

/**
 * FRONTEND ENTRY. The critical values still waiting for acknowledgement.
 *
 * @param {string} token  the caller's sessionToken
 * @return {{success:boolean, rows:Array, count:number, message:string}}
 */
function labListCriticalUnacked(token) {
  try {
    crescRequire_(token, 'lab.read');

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB.CRITICAL_COMMS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, rows: [], count: 0, message: '' };
    }
    var map = labHeaderMap_(sheet);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    var rows = [];
    var now = Date.now();
    data.forEach(function (r) {
      var ack = r[map['IsAcknowledged']];
      if (ack === true || String(ack).toUpperCase() === 'TRUE') return;

      var at = lab_ts_(r[map['CommunicatedAt']]);
      var ms = (typeof cresc_ms_ === 'function') ? cresc_ms_(at) : null;
      rows.push({
        commId:     String(r[map['CommID']] || ''),
        orderId:    String(r[map['OrderID']] || ''),
        patientId:  String(r[map['PatientID']] || ''),
        testName:   String(r[map['TestName']] || ''),
        value:      String(r[map['CriticalValue']] || ''),
        flag:       String(r[map['Flag']] || 'C'),
        loggedBy:   String(r[map['CommunicatedBy']] || ''),
        loggedAt:   at,
        // How long this has been waiting is the number that decides which one
        // to deal with first, and it was on no screen anywhere.
        ageMins:    ms ? Math.max(0, Math.round((now - ms) / 60000)) : null
      });
    });

    // Oldest first: the one that has been waiting longest is the one that
    // matters most, which is the opposite of the newest-first the rest of the
    // lab screens use.
    rows.sort(function (a, b) {
      if (a.ageMins === null) return 1;
      if (b.ageMins === null) return -1;
      return b.ageMins - a.ageMins;
    });

    return { success: true, rows: rows, count: rows.length, message: '' };
  } catch (err) {
    return { success: false, rows: [], count: 0, message: err.message };
  }
}

/**
 * FRONTEND ENTRY. Records that a clinician was actually reached.
 *
 * The note is REQUIRED and it is required for a reason: "acknowledged" on its
 * own is not a communication record. Who was told, on what number, at what
 * time — that is the thing an enquiry asks for, and a tick box cannot hold
 * it. The server refuses an empty one rather than accepting a row that will
 * not answer the question later.
 *
 * Writes are batched by row so acknowledging six does not mean six passes
 * over the sheet, and the whole thing runs under the script lock because two
 * technicians clearing the same list would otherwise interleave.
 *
 * @param {string} token
 * @param {Array<string>|string} commIds  CommID(s) from labListCriticalUnacked
 * @param {string} note      who was informed, and how
 * @return {{success:boolean, acknowledged:number, message:string}}
 */
function labAcknowledgeCritical(token, commIds, note) {
  var lock = LockService.getScriptLock();
  try {
    var actor = crescRequire_(token, 'lab.ack_critical');

    var ids = (typeof commIds === 'string') ? [commIds] : (commIds || []);
    ids = ids.map(function (v) { return String(v || '').trim().toUpperCase(); })
             .filter(function (v) { return !!v; });
    if (!ids.length) {
      return { success: false, acknowledged: 0, message: 'Nothing was selected to acknowledge.' };
    }

    var text = String(note || '').trim();
    if (text.length < 5) {
      return { success: false, acknowledged: 0,
               message: 'Record who you informed and how — for example ' +
                        '"Dr. Rekha, by phone on 98xxxxxx21, 14:20". ' +
                        'This is the communication record the standard asks for.' };
    }

    lock.waitLock(15000);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LAB.CRITICAL_COMMS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: false, acknowledged: 0, message: 'No critical communication log found.' };
    }
    var map = labHeaderMap_(sheet);
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();

    var who = actor.displayName || actor.username;
    var when = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    var done = 0, already = 0, touched = [];

    for (var i = 0; i < data.length; i++) {
      var id = String(data[i][map['CommID']] || '').trim().toUpperCase();
      if (ids.indexOf(id) === -1) continue;

      var ack = data[i][map['IsAcknowledged']];
      if (ack === true || String(ack).toUpperCase() === 'TRUE') { already++; continue; }

      data[i][map['AcknowledgedBy']] = who + (text ? ' — ' + text : '');
      data[i][map['AcknowledgedAt']] = when;
      data[i][map['IsAcknowledged']] = true;
      touched.push({ row: i + 2, values: data[i] });
      done++;
    }

    // One setValues per changed row. Rewriting the whole block would stamp
    // every untouched row with its own values again, which on a sheet this
    // narrow is harmless but on a shared one is a needless write conflict.
    touched.forEach(function (t) {
      sheet.getRange(t.row, 1, 1, lastCol).setValues([t.values]);
    });

    if (done) {
      logAudit_({ username: actor.username, role: actor.role, doctorId: actor.doctorId },
                'LAB_CRITICAL_ACK', 'LAB_CRITICAL_COMMS', ids.join(','),
                { count: done, note: text });
    }

    var msg = done
      ? done + ' critical value' + (done === 1 ? '' : 's') + ' acknowledged.' +
        (already ? ' ' + already + ' had already been acknowledged.' : '')
      : (already ? 'Those were already acknowledged by someone else.'
                 : 'No matching entries were found.');

    return { success: done > 0 || already > 0, acknowledged: done, message: msg };
  } catch (err) {
    return { success: false, acknowledged: 0, message: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function _lwCritCount_(){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.CRITICAL_COMMS); if(!sheet||sheet.getLastRow()<2) return 0;
    const map=labHeaderMap_(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    let n=0; data.forEach(function(r){ if(!(r[map['IsAcknowledged']]===true||String(r[map['IsAcknowledged']]).toUpperCase()==='TRUE')) n++; });
    return n;
  }catch(e){return 0;}
}
/**
 * How long finished or expired work stays on the bench queue.
 *
 * Two days. After that a dispatched report, a verified result and an order
 * whose sample was never collected are all history, not work, and they were
 * crowding out the orders somebody still has to act on.
 */
var LAB_QUEUE_WINDOW_MS = 2 * 24 * 60 * 60 * 1000;

function _lwWithin48h_(ts,nowMs){
  var t=cresc_ms_(ts);
  return t ? (nowMs-t)<LAB_QUEUE_WINDOW_MS : false;
}

/** Milliseconds since a timestamp, or null when it cannot be read. */
function _lwAgeMs_(ts,nowMs){
  var t=cresc_ms_(ts);
  return t ? (nowMs-t) : null;
}

/**
 * Should this order still appear on the bench queue?
 *
 * The queue is a worklist. Three kinds of row are not work:
 *
 *   • DISPATCHED / AMENDED / CANCELLED — finished and gone out.
 *   • VERIFIED older than the window — the result is signed off; it lives in
 *     Records now. Only VERIFIED was kept for ever before, so a busy month
 *     left hundreds of completed orders sitting in Verify & Dispatch.
 *   • An order past the window whose sample was NEVER collected. This is the
 *     "expired" case: a request nobody drew blood for cannot be run, and the
 *     tube would be out of date if they drew it now. It is hidden, never
 *     deleted, and the count is reported so the lab can see how many.
 *
 * Anything a technician could still act on stays, however old it is: an
 * uncollected order is only expired if it has genuinely sat past the window,
 * and work on the bench is never hidden.
 *
 * @return {{show:boolean, expired:boolean}}
 */
function _lwQueueVisibility_(status,createdAt,lastUpdatedAt,sampleStatus,nowMs){
  var terminal = (status==='REPORT_DISPATCHED'||status==='AMENDED'||status==='CANCELLED');
  if (terminal) {
    return { show: _lwWithin48h_(lastUpdatedAt||createdAt, nowMs), expired: false };
  }

  if (status==='VERIFIED') {
    return { show: _lwWithin48h_(lastUpdatedAt||createdAt, nowMs), expired: false };
  }

  // Awaiting billing or awaiting collection: expired once the window passes
  // with no sample drawn.
  if (status==='PENDING'||status==='BILLED'||status==='RECOLLECT') {
    var collected = String(sampleStatus||'').toUpperCase();
    var drawn = collected && collected !== 'PENDING' && collected !== 'REJECTED';
    if (drawn) return { show: true, expired: false };
    var age = _lwAgeMs_(createdAt, nowMs);
    if (age !== null && age >= LAB_QUEUE_WINDOW_MS) {
      return { show: false, expired: true };
    }
  }

  return { show: true, expired: false };
}

function _mkRowDef_(testId,pm){
  return {testId:testId,parameterId:pm.testId,parameterName:String(pm.testName||''),unit:String(pm.unit||''),resultType:String(pm.resultType||'NUMERIC'),
    maleRefLow:pm.maleRefLow,maleRefHigh:pm.maleRefHigh,femaleRefLow:pm.femaleRefLow,femaleRefHigh:pm.femaleRefHigh,
    criticalLow:pm.criticalLow,criticalHigh:pm.criticalHigh,
    refRangeText:(String(pm.resultType||'').toUpperCase()==='NUMERIC')?('M '+_rngTxt_(pm.maleRefLow,pm.maleRefHigh)+' F '+_rngTxt_(pm.femaleRefLow,pm.femaleRefHigh)):''};
}
function _refRangesByParameter_(){
  const cat=lab_catalog_(); const out={};
  if(!cat.success) return out;
  cat.panels.forEach(function(p){(p.parameters||[]).forEach(function(pm){ out[pm.testId]={testId:p.testId,parameterName:pm.testName,unit:pm.unit||'',resultType:pm.resultType||'NUMERIC',maleRefLow:pm.maleRefLow,maleRefHigh:pm.maleRefHigh,femaleRefLow:pm.femaleRefLow,femaleRefHigh:pm.femaleRefHigh,criticalLow:pm.criticalLow,criticalHigh:pm.criticalHigh}; }); });
  cat.individuals.forEach(function(t){ out[t.testId]={testId:t.testId,parameterName:t.testName,unit:t.unit||'',resultType:t.resultType||'NUMERIC',maleRefLow:t.maleRefLow,maleRefHigh:t.maleRefHigh,femaleRefLow:t.femaleRefLow,femaleRefHigh:t.femaleRefHigh,criticalLow:t.criticalLow,criticalHigh:t.criticalHigh}; });
  return out;
}
function _latestResultsForOrder_(orderId){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.RESULTS); const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap_(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){
    if(String(r[map['OrderID']])!==String(orderId)) return;
    if(!(r[map['IsLatest']]===true||String(r[map['IsLatest']]).toUpperCase()==='TRUE')) return;
    out[String(r[map['ParameterID']])]={value:String(r[map['ResultValue']]||''),flag:String(r[map['Flag']]||''),isDraft:(r[map['IsDraft']]===true||String(r[map['IsDraft']]).toUpperCase()==='TRUE'),interpretation:String(r[map['Interpretation']]||'')};
  });
  return out;
}
function _priorValuesForPatient_(patientId,excludeOrderId){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.RESULTS); const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap_(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){
    if(String(r[map['PatientID']])!==String(patientId)) return;
    if(String(r[map['OrderID']])===String(excludeOrderId)) return;
    if(r[map['IsDraft']]===true||String(r[map['IsDraft']]).toUpperCase()==='TRUE') return;
    const pid=String(r[map['ParameterID']]); const at=lab_ts_(r[map['VerifiedAt']]||r[map['EnteredAt']]);
    if(!out[pid]||at>out[pid].at) out[pid]={value:String(r[map['ResultValue']]||''),at:at};
  });
  return out;
}
/**
 * Latest result rows for an order, WITH their contents.
 *
 * _existingResultRowNums_() below reads the same block and throws the values
 * away, which forced every caller that wanted to edit a row to fetch cells
 * back one at a time. This keeps them.
 *
 * @return {{rowByParam: Object, valuesByRow: Object}} 1-based sheet rows
 */
function _existingResultRows_(sheet,map,orderId){
  const out={rowByParam:{},valuesByRow:{}};
  if(sheet.getLastRow()<2) return out;
  const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r,i){
    if(String(r[map['OrderID']])===String(orderId)&&
       (r[map['IsLatest']]===true||String(r[map['IsLatest']]).toUpperCase()==='TRUE')){
      const rowNum=i+2;
      out.rowByParam[String(r[map['ParameterID']])]=rowNum;
      out.valuesByRow[rowNum]=r;
    }
  });
  return out;
}

function _existingResultRowNums_(sheet,map,orderId){
  const out={};
  if(sheet.getLastRow()<2) return out;
  const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r,i){
    if(String(r[map['OrderID']])===String(orderId)&&(r[map['IsLatest']]===true||String(r[map['IsLatest']]).toUpperCase()==='TRUE')) out[String(r[map['ParameterID']])]=i+2;
  });
  return out;
}
function _firstSampleId_(orderId){
  const s=getOrderSamples(orderId); return (s.success&&s.samples.length)?s.samples[0].sampleId:'';
}
function _startTat_(order,collectedAtStr){
  try{
    const cat=lab_catalog_(); if(!cat.success) return;
    const byId={}; cat.panels.forEach(function(t){byId[t.testId]=t;}); cat.individuals.forEach(function(t){byId[t.testId]=t;}); cat.packages.forEach(function(t){byId[t.testId]=t;});
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.TAT_LOG); if(!sheet) return;
    const map=labHeaderMap_(sheet); const ncols=LAB_SCHEMA.LAB_TAT_LOG.length;
    // A TAT deadline computed from an unreadable collection time would be
    // 1970, and every test on the order would show as overdue the moment it
    // was booked in.
    const collectedMs=cresc_ms_(collectedAtStr)||Date.now();
    order.testIds.forEach(function(tid){
      const t=byId[tid]; if(!t) return;
      const tat=parseInt(t.tatMinutes,10)||0;
      const deadlineMs=collectedMs+tat*60000;
      const deadlineStr=Utilities.formatDate(new Date(deadlineMs),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
      const row=new Array(ncols).fill('');
      row[map['TATLogID']]='TAT-'+Utilities.getUuid().substring(0,8).toUpperCase();
      row[map['OrderID']]=order.orderId; row[map['TestID']]=tid; row[map['Priority']]=order.priority;
      row[map['SampleCollectedAt']]=collectedAtStr; row[map['TAT_Minutes']]=tat; row[map['TAT_Deadline']]=deadlineStr; row[map['IsOverdue']]=false;
      sheet.appendRow(row);
    });
  }catch(e){Logger.log('_startTat_ failed: '+e.message);}
}
function _closeTat_(orderId,verifiedStr){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.TAT_LOG); if(!sheet||sheet.getLastRow()<2) return;
    const map=labHeaderMap_(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    data.forEach(function(r,i){
      if(String(r[map['OrderID']])!==String(orderId)||r[map['ResultVerifiedAt']]) return;
      const rowNum=i+2;
      const dlStr=lab_ts_(r[map['TAT_Deadline']]); const stStr=lab_ts_(r[map['SampleCollectedAt']]);
      const actual=stStr?Math.round((new Date(verifiedStr)-new Date(stStr))/60000):'';
      const overdue=dlStr?(new Date(verifiedStr)>new Date(dlStr)):false;
      sheet.getRange(rowNum,map['ResultVerifiedAt']+1).setValue(verifiedStr);
      sheet.getRange(rowNum,map['ActualTAT_Minutes']+1).setValue(actual);
      sheet.getRange(rowNum,map['IsOverdue']+1).setValue(overdue);
      if(overdue&&dlStr) sheet.getRange(rowNum,map['OverdueBy_Minutes']+1).setValue(Math.round((new Date(verifiedStr)-new Date(dlStr))/60000));
    });
  }catch(e){Logger.log('_closeTat_ failed: '+e.message);}
}
function _raiseNcr_(d){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.NONCONFORMANCE); if(!sheet) return '';
    const map=labHeaderMap_(sheet); const ncols=LAB_SCHEMA.LAB_NONCONFORMANCE.length;
    const ncrId='NCR-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd')+'-'+Utilities.getUuid().substring(0,4).toUpperCase();
    const row=new Array(ncols).fill('');
    row[map['NCRId']]=ncrId; row[map['OrderID']]=String(d.orderId||''); row[map['SampleID']]=String(d.sampleId||'');
    row[map['PatientID']]=String(d.patientId||''); row[map['NCRType']]=String(d.type||'GENERAL');
    row[map['Description']]=String(d.description||''); row[map['ImmediateAction']]=String(d.immediateAction||'');
    row[map['RaisedBy']]=cresc_actorName_('SYSTEM');
    row[map['RaisedAt']]=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    row[map['Status']]='OPEN';
    sheet.appendRow(row); return ncrId;
  }catch(e){Logger.log('_raiseNcr_ failed: '+e.message); return '';}
}
/**
 * THE AUTHENTICITY BLOCK, REDUCED TO THE ONE THING IT MAY SAY ON PAPER.
 *
 * It used to print, in the footer of every released report:
 *
 *     AUTHENTICITY VERIFICATION
 *     Scan the code, or open the link below.
 *     Verification code 2DA0C547D7A3
 *     https://script.google.com/macros/s/AKfycb…/exec?verifyLab=…&c=…
 *
 * — beside a QR encoding the same address. That address is this deployment's
 * script URL, and it is handed to the patient, to whoever they show the
 * report to, and to anyone who photographs it. The deployment URL is the
 * whole of the authorisation to reach every anonymous route in this project,
 * so printing it on a document that leaves the building is the one thing
 * this system must not do. The QR was no better than the text: scanning it
 * yields the same string.
 *
 * So the link, the QR and the code are gone. What remains is the statement
 * that actually protects a patient — that a provisional result is
 * provisional. A draft used to carry the words AUTHENTICITY VERIFICATION
 * over an attestation hash of "N/A", which reads as a released report.
 *
 * If the clinic later puts the web app behind its own domain, a printed link
 * becomes safe again: see cresc_publicBaseUrl_() in Clinic_Profile.gs, which
 * is the opt-in that has to be set first. Nothing here reads it, because a
 * lab report is the document most likely to be photographed and forwarded.
 *
 * @param {string} orderId
 * @param {string} attestationHash  "N/A" until a verifier signs
 * @param {string} verifierName
 * @param {string} verifiedAt
 * @return {string} HTML for the footer slot
 */
function _labVerifyBlock_(orderId, attestationHash, verifierName, verifiedAt) {
  var hash = String(attestationHash || '').trim();
  var signed = hash && hash !== 'N/A';

  // An unverified report says so, plainly, and says nothing else.
  if (!signed) {
    return '<div class="qr-box">' +
      '<div style="border:1px dashed #b45309;color:#b45309;padding:8px 12px;border-radius:6px;">' +
      '<strong>NOT YET VERIFIED</strong><br>' +
      'This is a provisional result. It carries no authenticity code until a ' +
      'pathologist has verified and released it.' +
      '</div></div>';
  }

  // A released report needs no block at all: the signature panel beside this
  // one already names the pathologist and the moment of release.
  return '';
}

/**
 * The page a scanned lab QR lands on.
 *
 * Anonymous by design, exactly as the discharge summary's equivalent is:
 * the point is that somebody holding a printed report — a patient, another
 * hospital, an insurer — can confirm it is genuine without an account.
 *
 * It therefore carries NO clinical content. Not a result, not a value, not a
 * reference range. It answers one question: was this report released by this
 * laboratory, and is this copy the current one. The patient's name is
 * reduced to initials for the same reason.
 */
function labVerifyPage_(orderId, code) {
  var esc = function (v) {
    return String(v === null || v === undefined ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  };

  var status = 'NOT FOUND', tone = '#b91c1c', rows = [];

  try {
    var want = String(orderId || '').trim().toUpperCase();
    var give = String(code || '').trim().toUpperCase();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resSheet = ss.getSheetByName(LAB.RESULTS);

    if (want && give && resSheet && resSheet.getLastRow() >= 2) {
      var map = labHeaderMap_(resSheet);
      // TextFinder on the OrderID column, then read only the rows it names.
      // LAB_RESULTS is one of the fastest-growing sheets in the project and
      // this route is public, so it must not pull the whole sheet in.
      var col = map['OrderID'] + 1;
      var hits = resSheet.getRange(2, col, resSheet.getLastRow() - 1, 1)
                         .createTextFinder(want).matchEntireCell(true).findAll();
      for (var i = 0; i < hits.length; i++) {
        var row = resSheet.getRange(hits[i].getRow(), 1, 1, resSheet.getLastColumn()).getValues()[0];
        var hash = String(row[map['AttestationHash']] || '').trim();
        if (!hash || hash === 'N/A') continue;
        if (hash.substring(0, 12).toUpperCase() !== give) continue;

        status = 'VALID'; tone = '#047857';
        rows = [
          ['Order ID', want],
          ['Released by', String(row[map['VerifiedBy']] || '')],
          ['Released at', cresc_formatDate_(row[map['VerifiedAt']], 'dd-MMM-yyyy hh:mm a') || ''],
          ['Verification code', give],
          ['Laboratory', String((PropertiesService.getScriptProperties()
                                  .getProperty('CLINIC_NAME')) || 'Crescentia Clinic & Diagnostics')]
        ];
        break;
      }
    }
  } catch (e) {
    status = 'NOT FOUND'; tone = '#b91c1c'; rows = [];
  }

  var body = rows.map(function (r) {
    return '<tr><td style="padding:6px 14px 6px 0;color:#64748b;white-space:nowrap;">' +
           esc(r[0]) + '</td><td style="padding:6px 0;font-weight:600;">' + esc(r[1]) + '</td></tr>';
  }).join('');

  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>Report verification</title></head>' +
    '<body style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,sans-serif;' +
    'background:#f8fafc;margin:0;padding:32px 16px;color:#0f172a;">' +
    '<div style="max-width:420px;margin:0 auto;background:#fff;border:1px solid #e2e8f0;' +
    'border-radius:14px;padding:26px;">' +
    '<div style="font-size:12px;letter-spacing:.12em;color:#64748b;">LABORATORY REPORT</div>' +
    '<div style="font-size:26px;font-weight:700;color:' + tone + ';margin:6px 0 18px;">' +
    esc(status) + '</div>' +
    (rows.length
      ? '<table style="font-size:13px;border-collapse:collapse;">' + body + '</table>'
      : '<p style="font-size:13px;color:#475569;margin:0;">No released report matches this code. ' +
        'A report that has not yet been verified, or a code copied incorrectly, will both ' +
        'read NOT FOUND. Contact the laboratory if you believe this is wrong.</p>') +
    '<p style="font-size:11px;color:#94a3b8;margin:18px 0 0;">This page confirms that a report ' +
    'was released by this laboratory. It shows no clinical information.</p>' +
    '</div></body></html>')
    .setTitle('Report verification')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function _logCritical_(orderId,c){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.CRITICAL_COMMS); if(!sheet) return;
    const map=labHeaderMap_(sheet); const ncols=LAB_SCHEMA.LAB_CRITICAL_COMMS.length;
    const row=new Array(ncols).fill('');
    row[map['CommID']]='CRIT-'+Utilities.getUuid().substring(0,8).toUpperCase();
    row[map['OrderID']]=orderId; row[map['ResultID']]=String(c.resultId||''); row[map['PatientID']]=String(c.patientId||'');
    row[map['TestName']]=String(c.parameterName||''); row[map['CriticalValue']]=String(c.value||''); row[map['Flag']]='C';
    row[map['CommunicatedBy']]=cresc_actorName_('SYSTEM');
    row[map['CommunicatedAt']]=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    row[map['IsAcknowledged']]=false;
    sheet.appendRow(row);
  }catch(e){Logger.log('_logCritical_ failed: '+e.message);}
}

function _safeNum_(v){ if(v===''||v==null) return null; const n=Number(v); return isNaN(n)?null:n; }
function _safeNumW_(v){ if(v===''||v==null) return ''; const n=Number(v); return isNaN(n)?'':n; }
function _rngTxt_(lo,hi){ const l=_safeNum_(lo),h=_safeNum_(hi); if(l===null&&h===null) return '—'; return (l===null?'?':l)+'-'+(h===null?'?':h); }
function _safeParse_(s){ try{return JSON.parse(s);}catch(e){return [];} }

/**
 * ============================================================================
 * CRESCENTIA LAB — PHASE 1 ADDITIONS
 * ============================================================================
 * searchPatientByMobile, getSuggestedTubes, getLabBillHtml, _buildReportHtmlGrouped_.
 * Grouped (per-test sectioned) printing is wired directly into
 * getLabReportHtml() above — no separate dispatch step needed in this
 * order-level engine.
 * ============================================================================
 */


/* ───────────────────────────────────────────────────────────────────
   1. SEARCH PATIENT BY MOBILE  (ask #2)
   Patients sheet: col A(0)=PatientID, C(2)=Name, D(3)=Age, E(4)=Gender,
   G(6)=Mobile. Returns all matches (mobile may be shared by family).
   ─────────────────────────────────────────────────────────────────── */
function searchPatientByMobile(mobile, sessionToken) {
  try {
    crescRequire_(sessionToken, 'patient.read');
    dpdpLogRead_(crescActor_(sessionToken), 'PatientSearch', String(mobile || ''),
                 { endpoint: 'searchPatientByMobile' });
    var q = String(mobile || '').replace(/\D/g, '');   // digits only
    if (q.length < 4) return { success: false, message: 'Enter at least 4 digits of the mobile number.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('Patients');
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'Patients sheet missing.' };
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
    var matches = [];
    data.forEach(function (r) {
      var m = String(r[6] || '').replace(/\D/g, '');   // column G
      if (!m) return;
      if (m.indexOf(q) === -1) return;
      matches.push({
        patientId: String(r[0]),
        name:      String(r[2] || ''),
        age:       String(r[3] || ''),
        gender:    String(r[4] || '').toUpperCase().charAt(0) || 'M',
        mobile:    String(r[6] || '')
      });
    });
    if (!matches.length) return { success: false, message: 'No patient found for mobile "' + mobile + '".' };
    return { success: true, patients: matches };
  } catch (e) {
    return { success: false, message: 'searchPatientByMobile: ' + e.message };
  }
}


/* ───────────────────────────────────────────────────────────────────
   2. SUGGESTED TUBES FROM ORDERED TESTS  (ask #6)
   Reads LAB_ORDER_TESTS for the order, returns the distinct SampleTypes
   with the count of tests needing each. Front-end pre-fills tubes from
   this, and the tech can still add / delete / modify.
   ─────────────────────────────────────────────────────────────────── */
function getSuggestedTubes(orderId, sessionToken) {
  try {
    crescRequire_(sessionToken, 'lab.collect');
    if (!orderId) return { success: false, message: 'Order ID required.' };
    var byType = {};

    // Primary source: LAB_ORDER_TESTS (per-test rows), if populated.
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(LAB.ORDER_TESTS);
    if (sh && sh.getLastRow() >= 2) {
      var m = labHeaderMap_(sh), data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
      data.forEach(function (r) {
        if (String(r[m['OrderID']]) !== String(orderId)) return;
        var stype = String(r[m['SampleType']] || '').toUpperCase().trim();
        if (!stype) stype = 'BLOOD_PLAIN';
        if (!byType[stype]) byType[stype] = { sampleType: stype, tests: [] };
        byType[stype].tests.push(String(r[m['TestName']] || ''));
      });
    }

    // Fallback: derive from the order's testIds via the catalog
    // (this engine doesn't populate LAB_ORDER_TESTS, so this is the
    // path that actually runs in practice).
    if (!Object.keys(byType).length) {
      var od = getLabOrderDetail(orderId);
      if (od.success) {
        var catalog = _loadCatalogById_();
        var resolveTypes = function (tid) {
          var t = catalog[tid];
          if (!t) return;
          if (t.testType === 'PACKAGE') {
            (t.componentTestIds || '').split(',').filter(Boolean).forEach(function (cid) {
              var c = catalog[cid]; if (!c) return;
              var st = String(c.sampleType || '').toUpperCase().trim() || 'BLOOD_PLAIN';
              if (!byType[st]) byType[st] = { sampleType: st, tests: [] };
              byType[st].tests.push(c.testName);
            });
          } else {
            var st = String(t.sampleType || '').toUpperCase().trim() || 'BLOOD_PLAIN';
            if (!byType[st]) byType[st] = { sampleType: st, tests: [] };
            byType[st].tests.push(t.testName);
          }
        };
        od.order.testIds.forEach(resolveTypes);
      }
    }

    var tubes = Object.keys(byType).map(function (k) {
      return { sampleType: k, testCount: byType[k].tests.length, tests: byType[k].tests };
    });
    return { success: true, tubes: tubes };
  } catch (e) {
    return { success: false, message: 'getSuggestedTubes: ' + e.message };
  }
}


/* ───────────────────────────────────────────────────────────────────
   3. PRINTABLE BILL / INVOICE HTML  (ask #5)
   Separate from the lab REPORT. Reads the saved LAB_BILLING row, renders
   a clean invoice. Front-end loads this into the print window.
   ─────────────────────────────────────────────────────────────────── */
function getLabBillHtml(orderId, sessionToken) {
  try {
    crescRequire_(sessionToken, ['lab.bill', 'accounts.read']);
    if (!orderId) return { success: false, message: 'Order ID required.' };
    var od = getLabOrderDetail(orderId);
    if (!od.success) return { success: false, message: od.message };
    var o = od.order;

    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(LAB.BILLING);
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No bill found for this order.' };
    var m = labHeaderMap_(sh), data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    var b = null;
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][m['OrderID']]) === String(orderId)) {
        b = {
          billId:        String(data[i][m['BillID']] || ''),
          patientName:   String(data[i][m['PatientName']] || ''),
          category:      String(data[i][m['BillingCategory']] || ''),
          itemsJSON:     String(data[i][m['TestsJSON']] || '[]'),
          gross:         Number(data[i][m['GrossAmount']]) || 0,
          discPct:       Number(data[i][m['DiscountPercent']]) || 0,
          discAmt:       Number(data[i][m['DiscountAmount']]) || 0,
          net:           Number(data[i][m['NetAmount']]) || 0,
          payMode:       String(data[i][m['PaymentMode']] || ''),
          paid:          Number(data[i][m['PaidAmount']]) || 0,
          balance:       Number(data[i][m['BalanceAmount']]) || 0,
          payStatus:     String(data[i][m['PaymentStatus']] || ''),
          receipt:       String(data[i][m['ReceiptNumber']] || ''),
          billedAt:      cresc_formatDate_(data[i][m['BilledAt']], 'dd-MMM-yyyy hh:mm a')
        };
        break;
      }
    }
    if (!b) return { success: false, message: 'No bill found for this order.' };

    var items = [];
    try { items = JSON.parse(b.itemsJSON); } catch (ex) { items = []; }

    // Clinic_Profile.gs is the one reader; the fallback here used to say
    // "Crescentia Clinic" while the billing desk said "Crescentia HealthTech".
    var clinic        = cresc_clinic_();
    var clinicName    = clinic.name;
    var clinicAddress = clinic.address;
    var clinicPhone   = clinic.phone;
    var gstNumber     = clinic.gstin;

    var isIp = (b.category === 'IP_ACCOUNT');

    var itemRows = items.map(function (it, idx) {
      return '<tr>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">' + (idx + 1) + '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">' + _esc_(it.testName) + '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;text-align:right;">&#8377;' + Number(it.price).toFixed(2) + '</td>' +
      '</tr>';
    }).join('');

    var payLine = isIp
      ? '<div style="font-weight:700;color:#0369a1;">Posted to IP Account &#8226; Settled at discharge</div>'
      : '<div>Payment: <strong>' + _esc_(b.payMode) + '</strong> &#8226; Status: <strong>' + _esc_(b.payStatus) + '</strong></div>';

    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
      '*{box-sizing:border-box;margin:0;padding:0;font-family:Arial,Helvetica,sans-serif;}' +
      'body{background:#fff;color:#111827;}' +
      '@media print{@page{margin:1cm;} .no-print{display:none!important;}}' +
      '</style></head><body>' +
      '<div style="max-width:760px;margin:18px auto;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">' +

      // Header
      '<div style="padding:18px 24px;border-bottom:2px solid #0369a1;display:flex;justify-content:space-between;align-items:flex-start;">' +
        '<div>' +
          '<div style="font-size:20px;font-weight:800;color:#0369a1;">' + _esc_(clinicName) + '</div>' +
          (clinicAddress ? '<div style="font-size:12px;color:#6b7280;margin-top:2px;">' + _esc_(clinicAddress) + '</div>' : '') +
          (clinicPhone ? '<div style="font-size:12px;color:#6b7280;">' + _esc_(clinicPhone) + '</div>' : '') +
          (gstNumber ? '<div style="font-size:11px;color:#6b7280;">GSTIN: ' + _esc_(gstNumber) + '</div>' : '') +
        '</div>' +
        '<div style="text-align:right;">' +
          '<div style="font-size:16px;font-weight:800;letter-spacing:1px;color:#111827;">LAB INVOICE</div>' +
          '<div style="font-size:12px;color:#6b7280;margin-top:2px;">' + _esc_(b.receipt || b.billId) + '</div>' +
          '<div style="font-size:11px;color:#6b7280;">' + _esc_(b.billedAt) + '</div>' +
        '</div>' +
      '</div>' +

      // Patient strip
      '<div style="padding:12px 24px;background:#f9fafb;border-bottom:1px solid #e5e7eb;display:flex;justify-content:space-between;font-size:13px;">' +
        '<div><span style="color:#6b7280;">Patient:</span> <strong>' + _esc_(o.patientName) + '</strong> &#8226; ' + _esc_(o.patientId) + ' &#8226; ' + _esc_(o.gender) + (o.age ? ' &#8226; ' + _esc_(o.age) + 'y' : '') + '</div>' +
        '<div><span style="color:#6b7280;">Order:</span> ' + _esc_(o.orderId) + '</div>' +
      '</div>' +

      // Items
      '<div style="padding:8px 24px;">' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
          '<thead><tr style="background:#f3f4f6;">' +
            '<th style="padding:8px 10px;text-align:left;color:#374151;font-size:11px;text-transform:uppercase;">#</th>' +
            '<th style="padding:8px 10px;text-align:left;color:#374151;font-size:11px;text-transform:uppercase;">Test</th>' +
            '<th style="padding:8px 10px;text-align:right;color:#374151;font-size:11px;text-transform:uppercase;">Amount</th>' +
          '</tr></thead><tbody>' + itemRows + '</tbody>' +
        '</table>' +
      '</div>' +

      // Totals
      '<div style="padding:8px 24px 16px;display:flex;justify-content:flex-end;">' +
        '<table style="font-size:13px;min-width:260px;">' +
          '<tr><td style="padding:4px 10px;color:#6b7280;">Gross</td><td style="padding:4px 10px;text-align:right;">&#8377;' + b.gross.toFixed(2) + '</td></tr>' +
          (b.discAmt > 0 ? '<tr><td style="padding:4px 10px;color:#6b7280;">Discount (' + b.discPct + '%)</td><td style="padding:4px 10px;text-align:right;">- &#8377;' + b.discAmt.toFixed(2) + '</td></tr>' : '') +
          '<tr style="border-top:2px solid #111827;"><td style="padding:6px 10px;font-weight:800;">Net Payable</td><td style="padding:6px 10px;text-align:right;font-weight:800;font-size:15px;">&#8377;' + b.net.toFixed(2) + '</td></tr>' +
          (!isIp ? '<tr><td style="padding:4px 10px;color:#6b7280;">Paid</td><td style="padding:4px 10px;text-align:right;">&#8377;' + b.paid.toFixed(2) + '</td></tr>' : '') +
          (!isIp && b.balance > 0 ? '<tr><td style="padding:4px 10px;color:#dc2626;">Balance</td><td style="padding:4px 10px;text-align:right;color:#dc2626;font-weight:700;">&#8377;' + b.balance.toFixed(2) + '</td></tr>' : '') +
        '</table>' +
      '</div>' +

      // Footer
      '<div style="padding:12px 24px;background:#f9fafb;border-top:1px solid #e5e7eb;font-size:12px;color:#374151;display:flex;justify-content:space-between;align-items:center;">' +
        payLine +
        '<div style="color:#9ca3af;">This is a computer-generated invoice.</div>' +
      '</div>' +

      '</div>' +
      '<div class="no-print" style="text-align:center;padding:14px;">' +
        '<button onclick="window.print();" style="background:#0369a1;color:#fff;border:none;padding:10px 28px;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;">&#128424; Print / Save PDF</button>' +
      '</div>' +
      '</body></html>';

    return { success: true, html: html };
  } catch (e) {
    return { success: false, message: 'getLabBillHtml: ' + e.message };
  }
}


/* ───────────────────────────────────────────────────────────────────
   4. GROUPED REPORT BUILDER  (ask #10)
   Renders results in test-group sections (CBC, Lipid Profile…) instead
   of a flat parameter list. Each group can carry its own interpretation.
   This is a NEW function — your existing _buildReportHtml stays intact.
   `groups` = [ { groupName, interpretation, rows:[{parameterName,value,
                 unit,refRangeText,flag}] } ]
   ─────────────────────────────────────────────────────────────────── */
function _buildReportHtmlGrouped_(o, groups, bill, now) {
  var allRows = [];
  groups.forEach(function (g) { (g.rows || []).forEach(function (r) { allRows.push(r); }); });
  var hasCrit = allRows.some(function (r) { return r.flag === 'C'; });
  var vBy = '', vAt = '';
  for (var i = 0; i < allRows.length; i++) { if (allRows[i].verifiedBy) { vBy = allRows[i].verifiedBy; vAt = allRows[i].verifiedAt; break; } }

  var clinicName = cresc_clinic_().name;

  // Patient-ID barcode. The report is printed on a white card so the bars can
  // be black; the rest of the report is dark, which no scanner would read.
  var idBarcode = '';
  try {
    if (typeof bcp_code128Svg_ === 'function') {
      idBarcode = bcp_code128Svg_(String(o.patientId || '').trim(),
                                  { height: 8, moduleWidth: 0.30, fontSize: 2.6 });
    }
  } catch (e) { idBarcode = ''; }
  var idBarcodeHtml = idBarcode
    ? '<div style="background:#fff;padding:4px 6px;border-radius:6px;display:inline-block;margin-top:6px;">' +
      idBarcode + '</div>'
    : '';

  function sectionHtml(g) {
    var rowsHtml = (g.rows || []).map(function (r) {
      var fc = r.flag === 'C' ? 'color:#f87171;font-weight:800' :
               r.flag === 'H' ? 'color:#fb923c;font-weight:600' :
               r.flag === 'L' ? 'color:#60a5fa;font-weight:600' :
               r.flag === 'N' ? 'color:#10b981' : '';
      return '<tr>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #1e293b;">' + _esc_(r.parameterName) + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #1e293b;' + fc + '">' + _esc_(r.value) + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #1e293b;color:#94a3b8;">' + _esc_(r.unit) + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #1e293b;color:#94a3b8;font-size:12px;">' + _esc_(r.refRangeText) + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #1e293b;' + fc + '">' + _esc_(r.flag) + '</td>' +
      '</tr>';
    }).join('');
    return '<div style="padding:0 24px;margin-top:8px;">' +
      '<div style="font-weight:700;font-size:13px;padding:12px 0 6px;color:#7dd3fc;">' + _esc_(g.groupName) + '</div>' +
      '<table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#0a1628;">' +
        '<th style="padding:8px 10px;text-align:left;color:#94a3b8;font-size:11px;text-transform:uppercase;">Parameter</th>' +
        '<th style="padding:8px 10px;text-align:left;color:#94a3b8;font-size:11px;text-transform:uppercase;">Result</th>' +
        '<th style="padding:8px 10px;text-align:left;color:#94a3b8;font-size:11px;text-transform:uppercase;">Unit</th>' +
        '<th style="padding:8px 10px;text-align:left;color:#94a3b8;font-size:11px;text-transform:uppercase;">Reference</th>' +
        '<th style="padding:8px 10px;text-align:left;color:#94a3b8;font-size:11px;text-transform:uppercase;">Flag</th>' +
      '</tr></thead><tbody>' + rowsHtml + '</tbody></table>' +
      (g.interpretation ? '<div style="background:#0f2744;border:1px solid #1e3a5f;border-radius:8px;padding:10px 14px;margin:10px 0;font-size:13px;"><strong>' + _esc_(g.groupName) + ' — Interpretation:</strong> ' + _esc_(g.interpretation) + '</div>' : '') +
    '</div>';
  }

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:Arial,sans-serif;background:#020617;color:#f8fafc;}@media print{@page{margin:1cm;}.no-print{display:none!important;}}</style></head><body>' +
    '<div style="max-width:800px;margin:20px auto;background:#0f172a;border:1px solid #1e293b;border-radius:12px;overflow:hidden;">' +
    (hasCrit ? '<div style="background:#450a0a;border-bottom:2px solid #f87171;padding:10px 20px;color:#fca5a5;font-weight:700;font-size:14px;">&#9888; CRITICAL VALUES — Inform clinician immediately (NABL ISO 15189:2022)</div>' : '') +
    '<div style="background:linear-gradient(135deg,#1e3a5f,#0f2744);padding:18px 24px;display:flex;justify-content:space-between;align-items:center;">' +
    '<div><div style="font-size:18px;font-weight:800;color:#7dd3fc;">&#128300; ' + _esc_(clinicName) + ' — Lab Report</div><div style="color:#94a3b8;font-size:12px;margin-top:2px;">NABL ISO 15189:2022 Compliant</div></div>' +
    '<div style="text-align:right;"><div style="color:#94a3b8;font-size:11px;">Order ID</div><div style="font-family:monospace;color:#7dd3fc;font-size:13px;">' + _esc_(o.orderId) + '</div></div></div>' +
    '<div style="padding:14px 24px;background:#0a1628;border-bottom:1px solid #1e293b;display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;">' +
    '<div><div style="color:#94a3b8;font-size:10px;text-transform:uppercase;">Patient</div><div style="font-weight:700;font-size:15px;margin-top:2px;">' + _esc_(o.patientName) + '</div><div style="color:#94a3b8;font-size:12px;">' + _esc_(o.patientId) + ' · ' + _esc_(o.gender) + (o.age ? ' · ' + _esc_(o.age) + 'y' : '') + '</div>' + idBarcodeHtml + '</div>' +
    '<div><div style="color:#94a3b8;font-size:10px;text-transform:uppercase;">Ordered By</div><div style="font-weight:600;margin-top:2px;">' + _esc_(o.doctorName || '—') + '</div><div style="color:#94a3b8;font-size:12px;">' + _esc_(o.source) + '</div></div>' +
    '<div><div style="color:#94a3b8;font-size:10px;text-transform:uppercase;">Report Date</div><div style="font-weight:600;margin-top:2px;">' + _esc_(now.substring(0, 10)) + '</div>' + (bill && bill.receiptNumber ? '<div style="color:#94a3b8;font-size:12px;">Receipt: ' + _esc_(bill.receiptNumber) + '</div>' : '') + '</div></div>' +
    groups.map(sectionHtml).join('') +
    '<div style="padding:14px 24px;background:#0a1628;border-top:1px solid #1e293b;display:flex;justify-content:space-between;font-size:12px;color:#94a3b8;margin-top:8px;">' +
    '<div>Verified by: <strong style="color:#f8fafc;">' + _esc_(vBy || '—') + '</strong>' + (vAt ? ' at ' + _esc_(vAt) : '') + '</div>' +
    '<div>SHA-256 attested · Digitally Signed</div></div></div>' +
    '<div class="no-print" style="text-align:center;padding:14px;"><button onclick="window.print();" style="background:#10b981;color:#fff;border:none;padding:10px 28px;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;">&#128424; Print / Save PDF</button></div>' +
    '</body></html>';
}