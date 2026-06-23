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

function getOrderableTests() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.CATALOG);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, tests: [] };
    const map = labHeaderMap(sheet);
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

function getLabCatalog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.CATALOG);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, panels: [], individuals: [], packages: [] };
    const map = labHeaderMap(sheet);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const panelMap = {}, paramsByPanel = {}, individuals = [], packages = [];

    data.forEach(function (r) {
      const type = String(r[map['TestType']] || '');
      const active = (r[map['IsActive']] === true || String(r[map['IsActive']]).toUpperCase() === 'TRUE');
      const item = _rowToCatalogItem(r, map);
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

function _rowToCatalogItem(r, map) {
  return {
    testId:          String(r[map['TestID']]),
    testCode:        String(r[map['TestCode']]),
    testName:        String(r[map['TestName']]),
    testType:        String(r[map['TestType']]),
    department:      String(r[map['Department']] || ''),
    sampleType:      String(r[map['SampleType']] || ''),
    resultType:      String(r[map['ResultType']] || ''),
    unit:            String(r[map['Unit']] || ''),
    maleRefLow:      _safeNum(r[map['MaleRefLow']]),
    maleRefHigh:     _safeNum(r[map['MaleRefHigh']]),
    femaleRefLow:    _safeNum(r[map['FemaleRefLow']]),
    femaleRefHigh:   _safeNum(r[map['FemaleRefHigh']]),
    paediatricRefText: String(r[map['PaediatricRefText']] || ''),
    criticalLow:     _safeNum(r[map['CriticalLow']]),
    criticalHigh:    _safeNum(r[map['CriticalHigh']]),
    price:           Number(r[map['Price']]) || 0,
    tatMinutes:      Number(r[map['TAT_Minutes']]) || 0,
    componentTestIds:String(r[map['ComponentTestIDs']] || ''),
    requiresConsent: (r[map['RequiresConsent']] === true || String(r[map['RequiresConsent']]).toUpperCase() === 'TRUE'),
    sortOrder:       Number(r[map['SortOrder']]) || 0,
    isActive:        (r[map['IsActive']] === true || String(r[map['IsActive']]).toUpperCase() === 'TRUE')
  };
}

function saveCatalogEntry(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!p || !p.testName || !String(p.testName).trim()) return { success: false, message: 'Test name required.' };
    if (!p.testCode || !String(p.testCode).trim())       return { success: false, message: 'Test code required.' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.CATALOG);
    if (!sheet) return { success: false, message: 'Catalog sheet missing. Run setupLabDatabase() first.' };
    const map = labHeaderMap(sheet);
    const ncols = LAB_SCHEMA.LAB_TEST_CATALOG.length;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const by = Session.getActiveUser().getEmail() || 'SYSTEM';

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
    row[map['MaleRefLow']]       = isNumeric ? _safeNumW(p.maleRefLow)   : '';
    row[map['MaleRefHigh']]      = isNumeric ? _safeNumW(p.maleRefHigh)  : '';
    row[map['FemaleRefLow']]     = isNumeric ? _safeNumW(p.femaleRefLow) : '';
    row[map['FemaleRefHigh']]    = isNumeric ? _safeNumW(p.femaleRefHigh): '';
    row[map['PaediatricRefText']]= String(p.paediatricRefText || '');
    row[map['CriticalLow']]      = isNumeric ? _safeNumW(p.criticalLow)  : '';
    row[map['CriticalHigh']]     = isNumeric ? _safeNumW(p.criticalHigh) : '';
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

    labAudit(existingRow !== -1 ? 'CATALOG_UPDATED' : 'CATALOG_CREATED', 'CATALOG', testId, null, { name: p.testName, type: type });
    SpreadsheetApp.flush();
    return { success: true, message: (existingRow !== -1 ? 'Updated' : 'Created') + ' "' + p.testName + '".', testId: testId };
  } catch (err) {
    return { success: false, message: 'saveCatalogEntry failed: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function setCatalogActive(testId, isActive) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.CATALOG);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Catalog empty.' };
    const map = labHeaderMap(sheet);
    const ids = sheet.getRange(2, map['TestID']+1, sheet.getLastRow()-1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(testId)) {
        sheet.getRange(i+2, map['IsActive']+1).setValue(!!isActive);
        SpreadsheetApp.flush();
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

/** Global HTML-escape helper — used by getLabBillHtml, _buildReportHtmlGrouped, etc. */
function _esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function calculateFlag(value, resultType, gender, ref) {
  if (!resultType || String(resultType).toUpperCase() !== 'NUMERIC') return '';
  const v = parseFloat(value);
  if (isNaN(v)) return '';
  const g = String(gender || '').toUpperCase().charAt(0);
  const lo = g === 'F' ? _safeNum(ref.femaleRefLow)  : _safeNum(ref.maleRefLow);
  const hi = g === 'F' ? _safeNum(ref.femaleRefHigh) : _safeNum(ref.maleRefHigh);
  const cl = _safeNum(ref.criticalLow), ch = _safeNum(ref.criticalHigh);
  if (cl !== null && v < cl) return 'C';
  if (ch !== null && v > ch) return 'C';
  if (lo !== null && v < lo) return 'L';
  if (hi !== null && v > hi) return 'H';
  if (lo !== null || hi !== null) return 'N';
  return '';
}

function savePanelWithParameters(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!p || !p.testName || !String(p.testName).trim()) return { success: false, message: 'Panel name required.' };
    if (!p.testCode || !String(p.testCode).trim())       return { success: false, message: 'Panel code required.' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.CATALOG);
    if (!sheet) return { success: false, message: 'Catalog sheet missing.' };
    const map = labHeaderMap(sheet);
    const ncols = LAB_SCHEMA.LAB_TEST_CATALOG.length;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const by = Session.getActiveUser().getEmail() || 'SYSTEM';

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
      pmRow_[map['MaleRefLow']]    = isNum ? _safeNumW(pm.maleRefLow)  : '';
      pmRow_[map['MaleRefHigh']]   = isNum ? _safeNumW(pm.maleRefHigh) : '';
      pmRow_[map['CriticalLow']]   = isNum ? _safeNumW(pm.criticalLow) : '';
      pmRow_[map['CriticalHigh']]  = isNum ? _safeNumW(pm.criticalHigh): '';
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
    labAudit('PANEL_SAVED','CATALOG',panelId,null,{name:p.testName,params:params.length});
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

function createLabRequest(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!d) return { success: false, message: 'No data received.' };
    if (!d.patientId || !String(d.patientId).trim()) {
    d.patientId = 'WALKIN-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + Utilities.getUuid().substring(0,4).toUpperCase();
    }

    const source = String(d.sourceModule||'WALKIN').toUpperCase();
    if (['OPD','IP_CASESHEET','IP_NOTES','WALKIN'].indexOf(source) === -1) {
      return { success: false, message: 'Invalid source: ' + source };
    }

    let testIds = Array.isArray(d.testIds) ? d.testIds.filter(Boolean) : [];
    const nameList = Array.isArray(d.testNames) ? d.testNames.filter(Boolean) : [];

    if (!testIds.length && nameList.length) {
      testIds = _resolveTestNamesToCatalogIds(nameList);
      if (!testIds.length) {
        testIds = ['MANUAL_MAP'];
      }
    }
    if (!testIds.length) return { success: false, message: 'Select at least one test.' };

    const catalog = _loadCatalogById();
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
    const map = labHeaderMap(sheet);
    const ncols = LAB_SCHEMA.LAB_ORDERS.length;
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const by = Session.getActiveUser().getEmail() || 'SYSTEM';
    const orderId = 'LAB-ORD-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')
                    + '-' + Utilities.getUuid().substring(0,4).toUpperCase();

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
    labAudit('ORDER_CREATED','ORDER',orderId,null,{source:source, tests:names, priority:row[map['Priority']]});
    SpreadsheetApp.flush();

    return { success: true, message: 'Order '+orderId+' created.', orderId: orderId, requiresConsentFor: consent };
  } catch (err) {
    return { success: false, message: 'createLabRequest failed: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getLabOrderDetail(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.ORDERS);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Order not found.' };
    const map = labHeaderMap(sheet);
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
          createdAt:    String(r[map['CreatedAt']]||'')
        }
      };
    }
    return { success: false, message: 'Order not found: ' + orderId };
  } catch (err) {
    return { success: false, message: 'getLabOrderDetail failed: ' + err.message };
  }
}

/** * 🔥 HARDENED: FORCED TRANSITION
 * Bypasses strict indexOf checks to prevent UI gridlocks.
 */
function advanceOrderStatus(orderId, newStatus) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.ORDERS);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Order not found.' };
    
    const map = labHeaderMap(sheet);
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
        sheet.getRange(rowNum, map['LastUpdatedBy'] + 1).setValue(Session.getActiveUser().getEmail() || 'SYSTEM');
        
        labAudit('STATUS_CHANGED', 'ORDER', orderId, { from: cur }, { to: targetStatus });
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

function getLabWorkspaceData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const oSheet = ss.getSheetByName(LAB.ORDERS);
    const stats = { pendingBill:0, awaitingCollection:0, inProcess:0, awaitingVerify:0, verified:0, overdue:0, criticalPending:0, todayTotal:0 };
    if (!oSheet || oSheet.getLastRow() < 2) return { success: true, stats: stats, orders: [] };

    const oMap  = labHeaderMap(oSheet);
    const oData = oSheet.getRange(2,1,oSheet.getLastRow()-1,oSheet.getLastColumn()).getValues();
    const billIdx   = _lwBillingIndex();
    const sampIdx   = _lwSamplesIndex();
    const tatIdx    = _lwTatIndex();
    const nowMs     = new Date().getTime();
    const todayStr  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    stats.criticalPending = _lwCritCount();

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
      const createdAt = String(r[oMap['CreatedAt']]||'');
      if (createdAt.indexOf(todayStr) === 0) stats.todayTotal++;

      const stageIdx = {PENDING:0,BILLED:1,RECOLLECT:1,SAMPLE_COLLECTED:2,IN_PROCESS:3,RESULT_ENTERED:4,VERIFIED:5,REPORT_DISPATCHED:6,AMENDED:6,CANCELLED:0}[status]||0;
      if (status==='PENDING')   stats.pendingBill++;
      else if (status==='BILLED'||status==='RECOLLECT') stats.awaitingCollection++;
      else if (status==='SAMPLE_COLLECTED'||status==='IN_PROCESS') stats.inProcess++;
      else if (status==='RESULT_ENTERED') stats.awaitingVerify++;
      else if (status==='VERIFIED'||status==='REPORT_DISPATCHED') stats.verified++;
      if (overdue) stats.overdue++;

      const terminal = (status==='REPORT_DISPATCHED'||status==='CANCELLED');
      if (!terminal || _lwWithin48h(r[oMap['LastUpdatedAt']], nowMs)) {
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
          createdAt:    createdAt
        });
      }
    });

    const pw = {STAT:0,URGENT:1,ROUTINE:2};
    orders.sort(function(a,b){
      if(a.overdue!==b.overdue) return a.overdue?-1:1;
      if((pw[a.priority]||9)!==(pw[b.priority]||9)) return (pw[a.priority]||9)-(pw[b.priority]||9);
      return a.createdAt<b.createdAt?1:-1;
    });
    return { success: true, stats: stats, orders: orders };
  } catch (err) {
    return { success: false, message: 'getLabWorkspaceData failed: ' + err.message, stats: null, orders: [] };
  }
}

function getLabOrderPanel(orderId) {
  try {
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

function generateLabBill(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!d||!d.orderId) return { success: false, message: 'Order ID required.' };
    const od = getLabOrderDetail(d.orderId);
    if (!od.success) return { success: false, message: od.message };
    const order = od.order;
    if (order.status !== 'PENDING') return { success: false, message: 'Order already billed (status: '+order.status+').' };

    const catalog = _loadCatalogById();
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
    const by = Session.getActiveUser().getEmail()||'SYSTEM';
    let category, payMode, paid, balance, payStatus, receipt='', ledgerPostId='';

    if (ip) {
      category='IP_ACCOUNT'; payMode='ON_ACCOUNT'; paid=0; balance=net; payStatus='ON_ACCOUNT';
      ledgerPostId='IPPOST-'+Utilities.getUuid().substring(0,8).toUpperCase();
      labAudit('IP_ACCOUNT_POST','BILL',ledgerPostId,null,{admissionId:order.admissionId,amount:net});
    } else {
      category=(order.source==='WALKIN')?'WALKIN_SPOT':'OP_SPOT';
      payMode=String(d.paymentMode||'').toUpperCase();
      if(['CASH','CARD','UPI'].indexOf(payMode)===-1) return {success:false,message:'Select payment mode (Cash / Card / UPI).'};
      paid = (d.paidAmount===''||d.paidAmount==null) ? net : +Number(d.paidAmount).toFixed(2);
      balance = +(net-paid).toFixed(2);
      payStatus = balance<=0?'PAID':(paid>0?'PARTIAL':'PENDING');
      receipt = 'RCP-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd')+'-'+Utilities.getUuid().substring(0,4).toUpperCase();
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.BILLING);
    if (!sheet) return {success:false,message:'Billing sheet missing.'};
    const map = labHeaderMap(sheet);
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

    labAudit('BILL_GENERATED','BILL',billId,null,{orderId:order.orderId,net:net,category:category});
    advanceOrderStatus(order.orderId,'BILLED');
    SpreadsheetApp.flush();

    return { success:true, message: ip?('₹'+net.toFixed(2)+' posted to IP account.'):(billId+' · '+receipt+' · ₹'+net.toFixed(2)+' '+payStatus+'.'), billId:billId, receiptNumber:receipt, netAmount:net, category:category };
  } catch (err) {
    return { success:false, message:'generateLabBill failed: '+err.message };
  } finally {
    try{lock.releaseLock();}catch(e){}
  }
}

function getLabBill(orderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.BILLING);
    if (!sheet||sheet.getLastRow()<2) return {success:false,message:'No bills found.'};
    const map=labHeaderMap(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    let found=null;
    data.forEach(function(r){
      if(String(r[map['OrderID']])===String(orderId)){
        found={billId:String(r[map['BillID']]),orderId:String(r[map['OrderID']]),patientName:String(r[map['PatientName']]||''),
          category:String(r[map['BillingCategory']]||''),items:_safeParse(r[map['TestsJSON']]),gross:Number(r[map['GrossAmount']])||0,
          discountAmount:Number(r[map['DiscountAmount']])||0,net:Number(r[map['NetAmount']])||0,paymentMode:String(r[map['PaymentMode']]||''),
          paid:Number(r[map['PaidAmount']])||0,balance:Number(r[map['BalanceAmount']])||0,paymentStatus:String(r[map['PaymentStatus']]||''),
          receiptNumber:String(r[map['ReceiptNumber']]||''),billedAt:String(r[map['BilledAt']]||'')};
      }
    });
    return found?{success:true,bill:found}:{success:false,message:'No bill for this order.'};
  } catch(err){return{success:false,message:'getLabBill failed: '+err.message};}
}

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 6 — SAMPLES
   ═══════════════════════════════════════════════════════════════════════════ */

function collectLabSample(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if(!d||!d.orderId) return {success:false,message:'Order ID required.'};
    if(!Array.isArray(d.samples)||!d.samples.length) return {success:false,message:'Add at least one tube.'};
    const od = getLabOrderDetail(d.orderId);
    if(!od.success) return {success:false,message:od.message};
    const order = od.order;
    if(['BILLED','RECOLLECT'].indexOf(order.status)===-1) return {success:false,message:'Order must be billed before collection (status: '+order.status+').'};

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LAB.SAMPLES);
    if(!sheet) return {success:false,message:'Samples sheet missing.'};
    const map=labHeaderMap(sheet);
    const ncols=LAB_SCHEMA.LAB_SAMPLES.length;
    const nowStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    const by=Session.getActiveUser().getEmail()||'SYSTEM';
    const datePart=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd');
    const sampleIds=[];

    d.samples.forEach(function(s){
      const sampleId='LAB-SAMP-'+datePart+'-'+Utilities.getUuid().substring(0,4).toUpperCase();
      const barcode='BC'+datePart+Utilities.getUuid().substring(0,5).toUpperCase();
      const row=new Array(ncols).fill('');
      row[map['SampleID']]=sampleId; row[map['OrderID']]=order.orderId; row[map['PatientID']]=order.patientId;
      row[map['PatientName']]=order.patientName; row[map['SampleType']]=String(s.sampleType||'').toUpperCase();
      row[map['BarcodeID']]=barcode; row[map['CollectionStatus']]='COLLECTED'; row[map['RejectionReason']]='';
      row[map['CollectedAt']]=nowStr; row[map['CollectedBy']]=by; row[map['ReceivedAtLabAt']]=''; row[map['ReceivedBy']]=''; row[map['CreatedAt']]=nowStr;
      sheet.appendRow(row); sampleIds.push(sampleId);
    });

    _startTat(order, nowStr);
    advanceOrderStatus(order.orderId,'SAMPLE_COLLECTED');
    labAudit('SAMPLE_COLLECTED','SAMPLE',order.orderId,null,{count:sampleIds.length});
    SpreadsheetApp.flush();
    return {success:true,message:sampleIds.length+' sample(s) collected.',sampleIds:sampleIds};
  } catch(err){return{success:false,message:'collectLabSample failed: '+err.message};}
  finally{try{lock.releaseLock();}catch(e){}}
}

function receiveLabSample(orderId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.SAMPLES);
    if(!sheet||sheet.getLastRow()<2) return {success:false,message:'No samples found.'};
    const map=labHeaderMap(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    const nowStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    const by=Session.getActiveUser().getEmail()||'SYSTEM';
    let n=0;
    data.forEach(function(r,i){
      if(String(r[map['OrderID']])===String(orderId)&&String(r[map['CollectionStatus']])!=='REJECTED'&&!r[map['ReceivedAtLabAt']]){
        sheet.getRange(i+2,map['ReceivedAtLabAt']+1).setValue(nowStr);
        sheet.getRange(i+2,map['ReceivedBy']+1).setValue(by); n++;
      }
    });
    if(!n) return {success:false,message:'No unprocessed samples for this order.'};
    advanceOrderStatus(orderId,'IN_PROCESS');
    labAudit('SAMPLE_RECEIVED','SAMPLE',orderId,null,{count:n});
    SpreadsheetApp.flush();
    return {success:true,message:n+' sample(s) received. Processing started.'};
  } catch(err){return{success:false,message:'receiveLabSample failed: '+err.message};}
  finally{try{lock.releaseLock();}catch(e){}}
}

function rejectLabSample(sampleId, reason) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if(!sampleId) return {success:false,message:'Sample ID required.'};
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.SAMPLES);
    if(!sheet||sheet.getLastRow()<2) return {success:false,message:'No samples found.'};
    const map=labHeaderMap(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    let orderId='',patientId='',rowFound=-1;
    for(let i=0;i<data.length;i++){
      if(String(data[i][map['SampleID']])===String(sampleId)){rowFound=i+2;orderId=String(data[i][map['OrderID']]);patientId=String(data[i][map['PatientID']]);break;}
    }
    if(rowFound===-1) return {success:false,message:'Sample not found: '+sampleId};
    sheet.getRange(rowFound,map['CollectionStatus']+1).setValue('REJECTED');
    sheet.getRange(rowFound,map['RejectionReason']+1).setValue(String(reason||'').toUpperCase());
    const ncrId=_raiseNcr({orderId:orderId,sampleId:sampleId,patientId:patientId,type:'SAMPLE_REJECTION',description:'Sample '+sampleId+' rejected: '+reason,immediateAction:'Recollection requested'});
    advanceOrderStatus(orderId,'RECOLLECT');
    labAudit('SAMPLE_REJECTED','SAMPLE',sampleId,null,{reason:reason,ncr:ncrId});
    SpreadsheetApp.flush();
    return {success:true,message:'Rejected. NCR '+ncrId+' raised.',ncrId:ncrId};
  } catch(err){return{success:false,message:'rejectLabSample failed: '+err.message};}
  finally{try{lock.releaseLock();}catch(e){}}
}

function getOrderSamples(orderId) {
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.SAMPLES);
    if(!sheet||sheet.getLastRow()<2) return {success:true,samples:[]};
    const map=labHeaderMap(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    const out=[];
    data.forEach(function(r){
      if(String(r[map['OrderID']])!==String(orderId)) return;
      out.push({sampleId:String(r[map['SampleID']]),sampleType:String(r[map['SampleType']]||''),barcode:String(r[map['BarcodeID']]||''),
        status:String(r[map['CollectionStatus']]||''),rejectionReason:String(r[map['RejectionReason']]||''),
        collectedAt:String(r[map['CollectedAt']]||''),receivedAt:String(r[map['ReceivedAtLabAt']]||'')});
    });
    return {success:true,samples:out};
  } catch(err){return{success:false,message:'getOrderSamples failed: '+err.message};}
}

/* ═══════════════════════════════════════════════════════════════════════════
   SECTION 7 — RESULTS
   ═══════════════════════════════════════════════════════════════════════════ */

function buildResultEntrySheet(orderId) {
  try {
    const od=getLabOrderDetail(orderId);
    if(!od.success) return {success:false,message:od.message};
    const order=od.order;
    const cat=getLabCatalog();
    if(!cat.success) return {success:false,message:'Catalog unavailable.'};
    const panelById={},indivById={},pkgById={};
    cat.panels.forEach(function(p){panelById[p.testId]=p;});
    cat.individuals.forEach(function(t){indivById[t.testId]=t;});
    cat.packages.forEach(function(p){pkgById[p.testId]=p;});

    const rowDefs=[];
    const _addPanel=function(panel){(panel.parameters||[]).forEach(function(pm){rowDefs.push(_mkRowDef(panel.testId,pm));});};
    const _addIndiv=function(t){rowDefs.push(_mkRowDef(t.testId,t));};
    order.testIds.forEach(function(tid){
      if(panelById[tid]) _addPanel(panelById[tid]);
      else if(indivById[tid]) _addIndiv(indivById[tid]);
      else if(pkgById[tid]){ (pkgById[tid].componentTestIds||'').split(',').filter(Boolean).forEach(function(cid){if(panelById[cid])_addPanel(panelById[cid]);else if(indivById[cid])_addIndiv(indivById[cid]);}); }
    });

    const existing=_latestResultsForOrder(orderId);
    const prior=_priorValuesForPatient(order.patientId,orderId);

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

function saveLabResultsDraft(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if(!d||!d.orderId) return {success:false,message:'Order ID required.'};
    if(!Array.isArray(d.results)) return {success:false,message:'No results provided.'};
    const od=getLabOrderDetail(d.orderId);
    if(!od.success) return {success:false,message:od.message};
    const order=od.order;
    const gender=String(d.gender||order.gender||'').toUpperCase().charAt(0)||'M';
    const refByParam=_refRangesByParameter();
    const sampleId=_firstSampleId(order.orderId);
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.RESULTS);
    if(!sheet) return {success:false,message:'Results sheet missing.'};
    const map=labHeaderMap(sheet);
    const ncols=LAB_SCHEMA.LAB_RESULTS.length;
    const nowStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    const by=Session.getActiveUser().getEmail()||'SYSTEM';
    const existingRows=_existingResultRowNums(sheet,map,order.orderId);
    let critCount=0;

    d.results.forEach(function(res){
      if(!res.parameterId) return;
      const ref=refByParam[res.parameterId]||{};
      const flag=calculateFlag(res.value,ref.resultType,gender,ref);
      if(flag==='C') critCount++;
      const refText=(String(ref.resultType||'').toUpperCase()==='NUMERIC')
        ?(gender==='F'?_rngTxt(ref.femaleRefLow,ref.femaleRefHigh):_rngTxt(ref.maleRefLow,ref.maleRefHigh)):'';
      const exRow=existingRows[res.parameterId];
      if(exRow){
        sheet.getRange(exRow,map['ResultValue']+1).setValue(String(res.value==null?'':res.value));
        sheet.getRange(exRow,map['Flag']+1).setValue(flag);
        sheet.getRange(exRow,map['RefRangeText']+1).setValue(refText);
        sheet.getRange(exRow,map['Interpretation']+1).setValue(String(res.interpretation||''));
        sheet.getRange(exRow,map['IsDraft']+1).setValue(true);
        sheet.getRange(exRow,map['EnteredBy']+1).setValue(by);
        sheet.getRange(exRow,map['EnteredAt']+1).setValue(nowStr);
      } else {
        const row=new Array(ncols).fill('');
        row[map['ResultID']]='LAB-RES-'+Utilities.getUuid().substring(0,8).toUpperCase();
        row[map['OrderID']]=order.orderId; row[map['SampleID']]=sampleId; row[map['PatientID']]=order.patientId;
        row[map['TestID']]=ref.testId||''; row[map['ParameterID']]=res.parameterId; row[map['ParameterName']]=ref.parameterName||'';
        row[map['ResultValue']]=String(res.value==null?'':res.value); row[map['ResultType']]=ref.resultType||'';
        row[map['Unit']]=ref.unit||''; row[map['RefRangeText']]=refText; row[map['Flag']]=flag;
        row[map['IsDraft']]=true; row[map['EnteredBy']]=by; row[map['EnteredAt']]=nowStr;
        row[map['Interpretation']]=String(res.interpretation||''); row[map['Version']]=1; row[map['IsLatest']]=true;
        sheet.appendRow(row);
      }
    });
    SpreadsheetApp.flush();
    labAudit('RESULTS_DRAFT_SAVED','RESULT',order.orderId,null,{count:d.results.length,critical:critCount});
    return {success:true,message:'Draft saved'+(critCount?' · '+critCount+' CRITICAL value(s)':'')+'.', criticalCount:critCount};
  } catch(err){return{success:false,message:'saveLabResultsDraft failed: '+err.message};}
  finally{try{lock.releaseLock();}catch(e){}}
}

/** * 🔥 HARDENED: Submit Results 
 */
function submitLabResults(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if(!d||!d.orderId) return {success:false,message:'Order ID required.'};
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.RESULTS);
    if(!sheet||sheet.getLastRow()<2) return {success:false,message:'No results to submit.'};
    
    const map=labHeaderMap(sheet);
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
    advanceOrderStatus(d.orderId,'RESULT_ENTERED');
    
    crits.forEach(function(c){_logCritical(d.orderId,c);});
    labAudit('RESULTS_SUBMITTED','RESULT',d.orderId,null,{count:n,critical:crits.length});
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
function verifyLabResults(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if(!d||!d.orderId) return {success:false,message:'Order ID required.'};
    const od=getLabOrderDetail(d.orderId);
    if(!od.success) return {success:false,message:od.message};
    
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=ss.getSheetByName(LAB.RESULTS);
    const map=labHeaderMap(sheet);
    const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    const nowStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    const verifier=String(d.verifierName||Session.getActiveUser().getEmail()||'SYSTEM');
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
    
    targetRows.forEach(function(rowNum){
      sheet.getRange(rowNum,map['VerifiedBy']+1).setValue(verifier);
      sheet.getRange(rowNum,map['VerifiedAt']+1).setValue(nowStr);
      sheet.getRange(rowNum,map['AttestationHash']+1).setValue(hash);
      sheet.getRange(rowNum,map['IsDraft']+1).setValue(false);
      if(d.interpretation) sheet.getRange(rowNum,map['Interpretation']+1).setValue(String(d.interpretation));
    });
    
    _closeTat(d.orderId,nowStr);
    
    // Force transition to verified
    advanceOrderStatus(d.orderId,'VERIFIED');
    
    labAudit('RESULTS_VERIFIED','RESULT',d.orderId,null,{verifier:verifier,hash:hash.substring(0,12)+'…'});
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

function lookupLabPatient(pid) {
  try {
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
  var pid = 'LMTVS0001';   // ← put the patient ID you searched
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 
  var oSheet = ss.getSheetByName(LAB.ORDERS);
  var oMap = labHeaderMap(oSheet);
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
  var rMap = labHeaderMap(rSheet);
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


function searchLabRecords(query) {
  try {
    if (!query || !String(query).trim()) return { success:false, message:'Enter a Patient ID or name.' };
    var q = String(query).trim().toUpperCase();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var oSheet = ss.getSheetByName(LAB.ORDERS);
    if (!oSheet || oSheet.getLastRow() < 2) return { success:true, records:[], patient:null };
 
    var oMap  = labHeaderMap(oSheet);
    var oData = oSheet.getRange(2,1,oSheet.getLastRow()-1,oSheet.getLastColumn()).getValues();
    var matched = {};
    var patHeader = null;
 
    var REPORTABLE = ['VERIFIED','REPORT_DISPATCHED','AMENDED'];
 
    oData.forEach(function(r){
      var pid    = String(r[oMap['PatientID']]||'').toUpperCase();
      var pname  = String(r[oMap['PatientName']]||'').toUpperCase();
      var status = String(r[oMap['OrderStatus']]||'');
      // OLD — fails on leading/trailing spaces or case differences
        if (pid !== q && pname.indexOf(q) === -1) return;

        // NEW — trim and uppercase both sides before comparing
      var pidClean   = String(r[oMap['PatientID']]  || '').trim().toUpperCase();
      var pnameClean = String(r[oMap['PatientName']]|| '').trim().toUpperCase();
      var qClean     = q.trim().toUpperCase();
      if (pidClean !== qClean && pnameClean.indexOf(qClean) === -1) return;
      if (REPORTABLE.indexOf(status) === -1) return;
      var oid = String(r[oMap['OrderID']]);
      matched[oid] = {
        orderId: oid,
        date:    String(r[oMap['CreatedAt']]||''),
        testNames: String(r[oMap['TestNames']]||''),
        source:  String(r[oMap['SourceModule']]||''),
        doctor:  String(r[oMap['OrderingDoctorName']]||''),
        status:  status,
        verifiedBy: '',
        results: []
      };
      if (!patHeader) patHeader = {
        patientId: String(r[oMap['PatientID']]),
        name:      String(r[oMap['PatientName']]||''),
        age:       String(r[oMap['Age']]||''),
        gender:    String(r[oMap['Gender']]||'')
      };
    });
 
    var oids = Object.keys(matched);
    if (!oids.length) return { success:true, records:[], patient:null };
 
    var rSheet = ss.getSheetByName(LAB.RESULTS);
    if (rSheet && rSheet.getLastRow() >= 2) {
      var rMap  = labHeaderMap(rSheet);
      var rData = rSheet.getRange(2,1,rSheet.getLastRow()-1,rSheet.getLastColumn()).getValues();
      rData.forEach(function(r){
        var oid = String(r[rMap['OrderID']]);
        var rec = matched[oid];
        if (!rec) return;
        var isLatest = (r[rMap['IsLatest']]===true || String(r[rMap['IsLatest']]).toUpperCase()==='TRUE');
        var isDraft  = (r[rMap['IsDraft']]===true  || String(r[rMap['IsDraft']]).toUpperCase()==='TRUE');
        if (!isLatest || isDraft) return;
        rec.results.push({
          parameterName: String(r[rMap['ParameterName']]||''),
          value:         String(r[rMap['ResultValue']]||''),
          unit:          String(r[rMap['Unit']]||''),
          flag:          String(r[rMap['Flag']]||''),
          refRangeText:  String(r[rMap['RefRangeText']]||'')
        });
        if (!rec.verifiedBy) rec.verifiedBy = String(r[rMap['VerifiedBy']]||'');
      });
    }
 
    var records = oids.map(function(k){ return matched[k]; })
                      .sort(function(a,b){ return a.date < b.date ? 1 : -1; });
    return { success:true, records:records, patient:patHeader };
  } catch(err) {
    return { success:false, message:'searchLabRecords failed: ' + err.message };
  }
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
function getLabReportHtml(orderId) {
  try {
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
    const ordDate = orderRow[oHeaders.indexOf("CreatedAt")] ? new Date(orderRow[oHeaders.indexOf("CreatedAt")]).toLocaleString() : "";
    
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
          verifiedAt = resData[i][rHeaders.indexOf("VerifiedAt")] ? new Date(resData[i][rHeaders.indexOf("VerifiedAt")]).toLocaleString() : verifiedAt;
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

    const verifyUrl = encodeURIComponent(`https://valarmathi.clinic/verify?id=${orderId}&hash=${attestationHash.substring(0, 10)}`);
    const qrCodeImg = `https://chart.googleapis.com/chart?chs=100x100&cht=qr&chl=${verifyUrl}`;

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
            <div class="qr-box">
              <img src="${qrCodeImg}" alt="Verification QR" width="70" height="70" />
              <div>
                <strong>AUTHENTICITY VERIFICATION</strong><br>
                Scan QR to verify report integrity.<br>
                <span style="font-family: monospace; font-size: 9px; word-break: break-all;">SHA-256: ${attestationHash}</span>
              </div>
            </div>
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

function _loadCatalogById(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(LAB.CATALOG);
  const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap(sheet);
  const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){ out[String(r[map['TestID']])]=_rowToCatalogItem(r,map); });
  return out;
}

function _resolveTestNamesToCatalogIds(nameList) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(LAB.CATALOG);
  if(!sheet||sheet.getLastRow()<2) return [];
  const map=labHeaderMap(sheet);
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

function _lwBillingIndex(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.BILLING); const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){ const k=String(r[map['OrderID']]||''); if(k) out[k]={s:String(r[map['PaymentStatus']]||''),n:Number(r[map['NetAmount']])||0,r:String(r[map['ReceiptNumber']]||'')}; });
  return out;
}
function _lwSamplesIndex(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.SAMPLES); const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){ const k=String(r[map['OrderID']]||''); if(k) out[k]={s:String(r[map['CollectionStatus']]||'')}; });
  return out;
}
function _lwTatIndex(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.TAT_LOG); const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){
    const k=String(r[map['OrderID']]||''); if(!k) return;
    const dl=r[map['TAT_Deadline']]; const st=r[map['SampleCollectedAt']];
    const dlMs=dl?new Date(dl).getTime():0;
    if(!out[k]||dlMs>out[k].dlMs) out[k]={dlMs:dlMs,stMs:st?new Date(st).getTime():0};
  });
  return out;
}
function _lwCritCount(){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.CRITICAL_COMMS); if(!sheet||sheet.getLastRow()<2) return 0;
    const map=labHeaderMap(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    let n=0; data.forEach(function(r){ if(!(r[map['IsAcknowledged']]===true||String(r[map['IsAcknowledged']]).toUpperCase()==='TRUE')) n++; });
    return n;
  }catch(e){return 0;}
}
function _lwWithin48h(ts,nowMs){
  if(!ts) return false;
  try{return (nowMs-new Date(ts).getTime())<172800000;}catch(e){return false;}
}

function _mkRowDef(testId,pm){
  return {testId:testId,parameterId:pm.testId,parameterName:String(pm.testName||''),unit:String(pm.unit||''),resultType:String(pm.resultType||'NUMERIC'),
    maleRefLow:pm.maleRefLow,maleRefHigh:pm.maleRefHigh,femaleRefLow:pm.femaleRefLow,femaleRefHigh:pm.femaleRefHigh,
    criticalLow:pm.criticalLow,criticalHigh:pm.criticalHigh,
    refRangeText:(String(pm.resultType||'').toUpperCase()==='NUMERIC')?('M '+_rngTxt(pm.maleRefLow,pm.maleRefHigh)+' F '+_rngTxt(pm.femaleRefLow,pm.femaleRefHigh)):''};
}
function _refRangesByParameter(){
  const cat=getLabCatalog(); const out={};
  if(!cat.success) return out;
  cat.panels.forEach(function(p){(p.parameters||[]).forEach(function(pm){ out[pm.testId]={testId:p.testId,parameterName:pm.testName,unit:pm.unit||'',resultType:pm.resultType||'NUMERIC',maleRefLow:pm.maleRefLow,maleRefHigh:pm.maleRefHigh,femaleRefLow:pm.femaleRefLow,femaleRefHigh:pm.femaleRefHigh,criticalLow:pm.criticalLow,criticalHigh:pm.criticalHigh}; }); });
  cat.individuals.forEach(function(t){ out[t.testId]={testId:t.testId,parameterName:t.testName,unit:t.unit||'',resultType:t.resultType||'NUMERIC',maleRefLow:t.maleRefLow,maleRefHigh:t.maleRefHigh,femaleRefLow:t.femaleRefLow,femaleRefHigh:t.femaleRefHigh,criticalLow:t.criticalLow,criticalHigh:t.criticalHigh}; });
  return out;
}
function _latestResultsForOrder(orderId){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.RESULTS); const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){
    if(String(r[map['OrderID']])!==String(orderId)) return;
    if(!(r[map['IsLatest']]===true||String(r[map['IsLatest']]).toUpperCase()==='TRUE')) return;
    out[String(r[map['ParameterID']])]={value:String(r[map['ResultValue']]||''),flag:String(r[map['Flag']]||''),isDraft:(r[map['IsDraft']]===true||String(r[map['IsDraft']]).toUpperCase()==='TRUE'),interpretation:String(r[map['Interpretation']]||'')};
  });
  return out;
}
function _priorValuesForPatient(patientId,excludeOrderId){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.RESULTS); const out={};
  if(!sheet||sheet.getLastRow()<2) return out;
  const map=labHeaderMap(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r){
    if(String(r[map['PatientID']])!==String(patientId)) return;
    if(String(r[map['OrderID']])===String(excludeOrderId)) return;
    if(r[map['IsDraft']]===true||String(r[map['IsDraft']]).toUpperCase()==='TRUE') return;
    const pid=String(r[map['ParameterID']]); const at=String(r[map['VerifiedAt']]||r[map['EnteredAt']]||'');
    if(!out[pid]||at>out[pid].at) out[pid]={value:String(r[map['ResultValue']]||''),at:at};
  });
  return out;
}
function _existingResultRowNums(sheet,map,orderId){
  const out={};
  if(sheet.getLastRow()<2) return out;
  const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  data.forEach(function(r,i){
    if(String(r[map['OrderID']])===String(orderId)&&(r[map['IsLatest']]===true||String(r[map['IsLatest']]).toUpperCase()==='TRUE')) out[String(r[map['ParameterID']])]=i+2;
  });
  return out;
}
function _firstSampleId(orderId){
  const s=getOrderSamples(orderId); return (s.success&&s.samples.length)?s.samples[0].sampleId:'';
}
function _startTat(order,collectedAtStr){
  try{
    const cat=getLabCatalog(); if(!cat.success) return;
    const byId={}; cat.panels.forEach(function(t){byId[t.testId]=t;}); cat.individuals.forEach(function(t){byId[t.testId]=t;}); cat.packages.forEach(function(t){byId[t.testId]=t;});
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.TAT_LOG); if(!sheet) return;
    const map=labHeaderMap(sheet); const ncols=LAB_SCHEMA.LAB_TAT_LOG.length;
    const collectedMs=new Date(collectedAtStr).getTime();
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
  }catch(e){Logger.log('_startTat failed: '+e.message);}
}
function _closeTat(orderId,verifiedStr){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.TAT_LOG); if(!sheet||sheet.getLastRow()<2) return;
    const map=labHeaderMap(sheet); const data=sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    data.forEach(function(r,i){
      if(String(r[map['OrderID']])!==String(orderId)||r[map['ResultVerifiedAt']]) return;
      const rowNum=i+2;
      const dlStr=String(r[map['TAT_Deadline']]||''); const stStr=String(r[map['SampleCollectedAt']]||'');
      const actual=stStr?Math.round((new Date(verifiedStr)-new Date(stStr))/60000):'';
      const overdue=dlStr?(new Date(verifiedStr)>new Date(dlStr)):false;
      sheet.getRange(rowNum,map['ResultVerifiedAt']+1).setValue(verifiedStr);
      sheet.getRange(rowNum,map['ActualTAT_Minutes']+1).setValue(actual);
      sheet.getRange(rowNum,map['IsOverdue']+1).setValue(overdue);
      if(overdue&&dlStr) sheet.getRange(rowNum,map['OverdueBy_Minutes']+1).setValue(Math.round((new Date(verifiedStr)-new Date(dlStr))/60000));
    });
  }catch(e){Logger.log('_closeTat failed: '+e.message);}
}
function _raiseNcr(d){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.NONCONFORMANCE); if(!sheet) return '';
    const map=labHeaderMap(sheet); const ncols=LAB_SCHEMA.LAB_NONCONFORMANCE.length;
    const ncrId='NCR-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd')+'-'+Utilities.getUuid().substring(0,4).toUpperCase();
    const row=new Array(ncols).fill('');
    row[map['NCRId']]=ncrId; row[map['OrderID']]=String(d.orderId||''); row[map['SampleID']]=String(d.sampleId||'');
    row[map['PatientID']]=String(d.patientId||''); row[map['NCRType']]=String(d.type||'GENERAL');
    row[map['Description']]=String(d.description||''); row[map['ImmediateAction']]=String(d.immediateAction||'');
    row[map['RaisedBy']]=Session.getActiveUser().getEmail()||'SYSTEM';
    row[map['RaisedAt']]=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    row[map['Status']]='OPEN';
    sheet.appendRow(row); return ncrId;
  }catch(e){Logger.log('_raiseNcr failed: '+e.message); return '';}
}
function _logCritical(orderId,c){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(LAB.CRITICAL_COMMS); if(!sheet) return;
    const map=labHeaderMap(sheet); const ncols=LAB_SCHEMA.LAB_CRITICAL_COMMS.length;
    const row=new Array(ncols).fill('');
    row[map['CommID']]='CRIT-'+Utilities.getUuid().substring(0,8).toUpperCase();
    row[map['OrderID']]=orderId; row[map['ResultID']]=String(c.resultId||''); row[map['PatientID']]=String(c.patientId||'');
    row[map['TestName']]=String(c.parameterName||''); row[map['CriticalValue']]=String(c.value||''); row[map['Flag']]='C';
    row[map['CommunicatedBy']]=Session.getActiveUser().getEmail()||'SYSTEM';
    row[map['CommunicatedAt']]=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
    row[map['IsAcknowledged']]=false;
    sheet.appendRow(row);
  }catch(e){Logger.log('_logCritical failed: '+e.message);}
}

function _safeNum(v){ if(v===''||v==null) return null; const n=Number(v); return isNaN(n)?null:n; }
function _safeNumW(v){ if(v===''||v==null) return ''; const n=Number(v); return isNaN(n)?'':n; }
function _rngTxt(lo,hi){ const l=_safeNum(lo),h=_safeNum(hi); if(l===null&&h===null) return '—'; return (l===null?'?':l)+'-'+(h===null?'?':h); }
function _safeParse(s){ try{return JSON.parse(s);}catch(e){return [];} }

/**
 * ============================================================================
 * CRESCENTIA LAB — PHASE 1 ADDITIONS
 * ============================================================================
 * searchPatientByMobile, getSuggestedTubes, getLabBillHtml, _buildReportHtmlGrouped.
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
function searchPatientByMobile(mobile) {
  try {
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
function getSuggestedTubes(orderId) {
  try {
    if (!orderId) return { success: false, message: 'Order ID required.' };
    var byType = {};

    // Primary source: LAB_ORDER_TESTS (per-test rows), if populated.
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(LAB.ORDER_TESTS);
    if (sh && sh.getLastRow() >= 2) {
      var m = labHeaderMap(sh), data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
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
        var catalog = _loadCatalogById();
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
function getLabBillHtml(orderId) {
  try {
    if (!orderId) return { success: false, message: 'Order ID required.' };
    var od = getLabOrderDetail(orderId);
    if (!od.success) return { success: false, message: od.message };
    var o = od.order;

    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(LAB.BILLING);
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No bill found for this order.' };
    var m = labHeaderMap(sh), data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
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
          billedAt:      String(data[i][m['BilledAt']] || '')
        };
        break;
      }
    }
    if (!b) return { success: false, message: 'No bill found for this order.' };

    var items = [];
    try { items = JSON.parse(b.itemsJSON); } catch (ex) { items = []; }

    var props = PropertiesService.getScriptProperties().getProperties();
    var clinicName    = props['CLINIC_NAME']    || 'Crescentia Clinic';
    var clinicAddress = props['CLINIC_ADDRESS'] || '';
    var clinicPhone   = props['CLINIC_PHONE']   || '';
    var gstNumber     = props['CLINIC_GST']     || '';

    var isIp = (b.category === 'IP_ACCOUNT');

    var itemRows = items.map(function (it, idx) {
      return '<tr>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">' + (idx + 1) + '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">' + _esc(it.testName) + '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;text-align:right;">&#8377;' + Number(it.price).toFixed(2) + '</td>' +
      '</tr>';
    }).join('');

    var payLine = isIp
      ? '<div style="font-weight:700;color:#0369a1;">Posted to IP Account &#8226; Settled at discharge</div>'
      : '<div>Payment: <strong>' + _esc(b.payMode) + '</strong> &#8226; Status: <strong>' + _esc(b.payStatus) + '</strong></div>';

    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
      '*{box-sizing:border-box;margin:0;padding:0;font-family:Arial,Helvetica,sans-serif;}' +
      'body{background:#fff;color:#111827;}' +
      '@media print{@page{margin:1cm;} .no-print{display:none!important;}}' +
      '</style></head><body>' +
      '<div style="max-width:760px;margin:18px auto;border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;">' +

      // Header
      '<div style="padding:18px 24px;border-bottom:2px solid #0369a1;display:flex;justify-content:space-between;align-items:flex-start;">' +
        '<div>' +
          '<div style="font-size:20px;font-weight:800;color:#0369a1;">' + _esc(clinicName) + '</div>' +
          (clinicAddress ? '<div style="font-size:12px;color:#6b7280;margin-top:2px;">' + _esc(clinicAddress) + '</div>' : '') +
          (clinicPhone ? '<div style="font-size:12px;color:#6b7280;">' + _esc(clinicPhone) + '</div>' : '') +
          (gstNumber ? '<div style="font-size:11px;color:#6b7280;">GSTIN: ' + _esc(gstNumber) + '</div>' : '') +
        '</div>' +
        '<div style="text-align:right;">' +
          '<div style="font-size:16px;font-weight:800;letter-spacing:1px;color:#111827;">LAB INVOICE</div>' +
          '<div style="font-size:12px;color:#6b7280;margin-top:2px;">' + _esc(b.receipt || b.billId) + '</div>' +
          '<div style="font-size:11px;color:#6b7280;">' + _esc(b.billedAt) + '</div>' +
        '</div>' +
      '</div>' +

      // Patient strip
      '<div style="padding:12px 24px;background:#f9fafb;border-bottom:1px solid #e5e7eb;display:flex;justify-content:space-between;font-size:13px;">' +
        '<div><span style="color:#6b7280;">Patient:</span> <strong>' + _esc(o.patientName) + '</strong> &#8226; ' + _esc(o.patientId) + ' &#8226; ' + _esc(o.gender) + (o.age ? ' &#8226; ' + _esc(o.age) + 'y' : '') + '</div>' +
        '<div><span style="color:#6b7280;">Order:</span> ' + _esc(o.orderId) + '</div>' +
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
function _buildReportHtmlGrouped(o, groups, bill, now) {
  var allRows = [];
  groups.forEach(function (g) { (g.rows || []).forEach(function (r) { allRows.push(r); }); });
  var hasCrit = allRows.some(function (r) { return r.flag === 'C'; });
  var vBy = '', vAt = '';
  for (var i = 0; i < allRows.length; i++) { if (allRows[i].verifiedBy) { vBy = allRows[i].verifiedBy; vAt = allRows[i].verifiedAt; break; } }

  var props = PropertiesService.getScriptProperties().getProperties();
  var clinicName = props['CLINIC_NAME'] || 'Crescentia Clinic';

  function sectionHtml(g) {
    var rowsHtml = (g.rows || []).map(function (r) {
      var fc = r.flag === 'C' ? 'color:#f87171;font-weight:800' :
               r.flag === 'H' ? 'color:#fb923c;font-weight:600' :
               r.flag === 'L' ? 'color:#60a5fa;font-weight:600' :
               r.flag === 'N' ? 'color:#10b981' : '';
      return '<tr>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #1e293b;">' + _esc(r.parameterName) + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #1e293b;' + fc + '">' + _esc(r.value) + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #1e293b;color:#94a3b8;">' + _esc(r.unit) + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #1e293b;color:#94a3b8;font-size:12px;">' + _esc(r.refRangeText) + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #1e293b;' + fc + '">' + _esc(r.flag) + '</td>' +
      '</tr>';
    }).join('');
    return '<div style="padding:0 24px;margin-top:8px;">' +
      '<div style="font-weight:700;font-size:13px;padding:12px 0 6px;color:#7dd3fc;">' + _esc(g.groupName) + '</div>' +
      '<table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#0a1628;">' +
        '<th style="padding:8px 10px;text-align:left;color:#94a3b8;font-size:11px;text-transform:uppercase;">Parameter</th>' +
        '<th style="padding:8px 10px;text-align:left;color:#94a3b8;font-size:11px;text-transform:uppercase;">Result</th>' +
        '<th style="padding:8px 10px;text-align:left;color:#94a3b8;font-size:11px;text-transform:uppercase;">Unit</th>' +
        '<th style="padding:8px 10px;text-align:left;color:#94a3b8;font-size:11px;text-transform:uppercase;">Reference</th>' +
        '<th style="padding:8px 10px;text-align:left;color:#94a3b8;font-size:11px;text-transform:uppercase;">Flag</th>' +
      '</tr></thead><tbody>' + rowsHtml + '</tbody></table>' +
      (g.interpretation ? '<div style="background:#0f2744;border:1px solid #1e3a5f;border-radius:8px;padding:10px 14px;margin:10px 0;font-size:13px;"><strong>' + _esc(g.groupName) + ' — Interpretation:</strong> ' + _esc(g.interpretation) + '</div>' : '') +
    '</div>';
  }

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:Arial,sans-serif;background:#020617;color:#f8fafc;}@media print{@page{margin:1cm;}.no-print{display:none!important;}}</style></head><body>' +
    '<div style="max-width:800px;margin:20px auto;background:#0f172a;border:1px solid #1e293b;border-radius:12px;overflow:hidden;">' +
    (hasCrit ? '<div style="background:#450a0a;border-bottom:2px solid #f87171;padding:10px 20px;color:#fca5a5;font-weight:700;font-size:14px;">&#9888; CRITICAL VALUES — Inform clinician immediately (NABL ISO 15189:2022)</div>' : '') +
    '<div style="background:linear-gradient(135deg,#1e3a5f,#0f2744);padding:18px 24px;display:flex;justify-content:space-between;align-items:center;">' +
    '<div><div style="font-size:18px;font-weight:800;color:#7dd3fc;">&#128300; ' + _esc(clinicName) + ' — Lab Report</div><div style="color:#94a3b8;font-size:12px;margin-top:2px;">NABL ISO 15189:2022 Compliant</div></div>' +
    '<div style="text-align:right;"><div style="color:#94a3b8;font-size:11px;">Order ID</div><div style="font-family:monospace;color:#7dd3fc;font-size:13px;">' + _esc(o.orderId) + '</div></div></div>' +
    '<div style="padding:14px 24px;background:#0a1628;border-bottom:1px solid #1e293b;display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;">' +
    '<div><div style="color:#94a3b8;font-size:10px;text-transform:uppercase;">Patient</div><div style="font-weight:700;font-size:15px;margin-top:2px;">' + _esc(o.patientName) + '</div><div style="color:#94a3b8;font-size:12px;">' + _esc(o.patientId) + ' · ' + _esc(o.gender) + (o.age ? ' · ' + _esc(o.age) + 'y' : '') + '</div></div>' +
    '<div><div style="color:#94a3b8;font-size:10px;text-transform:uppercase;">Ordered By</div><div style="font-weight:600;margin-top:2px;">' + _esc(o.doctorName || '—') + '</div><div style="color:#94a3b8;font-size:12px;">' + _esc(o.source) + '</div></div>' +
    '<div><div style="color:#94a3b8;font-size:10px;text-transform:uppercase;">Report Date</div><div style="font-weight:600;margin-top:2px;">' + _esc(now.substring(0, 10)) + '</div>' + (bill && bill.receiptNumber ? '<div style="color:#94a3b8;font-size:12px;">Receipt: ' + _esc(bill.receiptNumber) + '</div>' : '') + '</div></div>' +
    groups.map(sectionHtml).join('') +
    '<div style="padding:14px 24px;background:#0a1628;border-top:1px solid #1e293b;display:flex;justify-content:space-between;font-size:12px;color:#94a3b8;margin-top:8px;">' +
    '<div>Verified by: <strong style="color:#f8fafc;">' + _esc(vBy || '—') + '</strong>' + (vAt ? ' at ' + _esc(vAt) : '') + '</div>' +
    '<div>SHA-256 attested · Digitally Signed</div></div></div>' +
    '<div class="no-print" style="text-align:center;padding:14px;"><button onclick="window.print();" style="background:#10b981;color:#fff;border:none;padding:10px 28px;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;">&#128424; Print / Save PDF</button></div>' +
    '</body></html>';
}