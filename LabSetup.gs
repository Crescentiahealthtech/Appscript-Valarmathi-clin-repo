/**
 * ============================================================================
 * CRESCENTIA HEALTHTECH — LAB MODULE SETUP  (v5 — PRODUCTION)
 * ============================================================================
 * THIS FILE MUST BE FIRST IN THE APPS SCRIPT PROJECT FILE ORDER.
 * Defines: LAB constant, LAB_SCHEMA, labHeaderMap(), labAudit(),
 * LAB_ORDER_STATUS_OK, LAB_TEST_STATUS_OK
 * All consumed by LabIntegrationEngine.gs
 * ============================================================================
 */

// ── Canonical sheet name references ─────────────────────────────────────────
var LAB = {
  CATALOG:        'LAB_TEST_CATALOG',
  ORDERS:         'LAB_ORDERS',
  ORDER_TESTS:    'LAB_ORDER_TESTS',
  BILLING:        'LAB_BILLING',
  SAMPLES:        'LAB_SAMPLES',
  RESULTS:        'LAB_RESULTS',
  REPORTS:        'LAB_REPORTS',
  TAT_LOG:        'LAB_TAT_LOG',
  AUDIT_LOG:      'LAB_AUDIT_LOG',
  QC_LOG:         'LAB_QC_LOG',
  NONCONFORMANCE: 'LAB_NONCONFORMANCE',
  CRITICAL_COMMS: 'LAB_CRITICAL_COMMS',
  EQA_LOG:        'LAB_EQA_LOG'
};

// ── Column headers for every sheet ──────────────────────────────────────────
var LAB_SCHEMA = {
  LAB_TEST_CATALOG: [
    'TestID','TestCode','TestName','TestType','ParentPanelID','Department',
    'SampleType','ResultType','Unit','MaleRefLow','MaleRefHigh',
    'FemaleRefLow','FemaleRefHigh','PaediatricRefText','CriticalLow','CriticalHigh',
    'Price','TAT_Minutes','ComponentTestIDs','RequiresConsent','SortOrder',
    'IsActive','CreatedAt','CreatedBy'
  ],
  LAB_ORDERS: [
    'OrderID','PatientID','PatientName','Age','Gender','VisitID','AdmissionID',
    'SourceModule','OrderingDoctorID','OrderingDoctorName','TestIDs','TestNames',
    'Priority','ClinicalNote','RepeatOfOrderID','RepeatReason',
    'OrderStatus','CreatedAt','CreatedBy','LastUpdatedAt','LastUpdatedBy'
  ],
  LAB_ORDER_TESTS: [
    'RowID','OrderID','PatientID','TestID','TestCode','TestName','TestType',
    'SampleType','TAT_Minutes','TAT_Deadline','Priority','TestStatus',
    'SampleID','ResultEnteredAt','ResultEnteredBy',
    'VerifiedAt','VerifiedBy','AttestationHash','DispatchedAt','CreatedAt'
  ],
  LAB_BILLING: [
    'BillID','OrderID','PatientID','PatientName','BillingCategory','AdmissionID',
    'TestsJSON','GrossAmount','DiscountPercent','DiscountAmount','NetAmount',
    'PaymentMode','PaidAmount','BalanceAmount','PaymentStatus',
    'ReceiptNumber','IPLedgerPostID','BilledAt','BilledBy'
  ],
  LAB_SAMPLES: [
    'SampleID','OrderID','PatientID','PatientName','SampleType','BarcodeID',
    'CollectionStatus','RejectionReason','RejectedAt','RejectedBy',
    'CollectedAt','CollectedBy','ReceivedAtLabAt','ReceivedBy','CreatedAt'
  ],
  LAB_RESULTS: [
    'ResultID','OrderID','SampleID','PatientID','TestID','ParameterID',
    'ParameterName','ResultValue','ResultType','Unit','RefRangeText','Flag',
    'IsDraft','EnteredBy','EnteredAt','VerifiedBy','VerifiedAt',
    'AttestationHash','Interpretation','Version','IsLatest','AmendmentReason','AmendedBy','AmendedAt'
  ],
  LAB_REPORTS: [
    'ReportID','OrderID','TestID','PatientID','ReportStatus',
    'GeneratedAt','GeneratedBy','DispatchedAt','DispatchedTo','ReportHTML','AmendedFrom'
  ],
  LAB_TAT_LOG: [
    'TATLogID','OrderID','TestID','Priority','SampleCollectedAt',
    'TAT_Minutes','TAT_Deadline','ResultVerifiedAt',
    'ActualTAT_Minutes','IsOverdue','OverdueBy_Minutes'
  ],
  LAB_AUDIT_LOG: [
    'AuditID','Timestamp','Action','EntityType','EntityID',
    'OldValue','NewValue','PerformedBy'
  ],
  LAB_QC_LOG: [
    'QCLogID','Date','Shift','TestID','TestName','LotNumber',
    'MeanValue','SDValue','ObservedValue','ZScore',
    'WestgardRule','QCStatus','Comments','EnteredBy'
  ],
  LAB_NONCONFORMANCE: [
    'NCRId','OrderID','SampleID','PatientID','NCRType',
    'Description','ImmediateAction','CorrectiveAction',
    'RaisedBy','RaisedAt','ClosedBy','ClosedAt','Status'
  ],
  LAB_CRITICAL_COMMS: [
    'CommID','OrderID','ResultID','PatientID','TestName','CriticalValue','Flag',
    'CommunicatedBy','CommunicatedAt','AcknowledgedBy','AcknowledgedAt','IsAcknowledged'
  ],
  LAB_EQA_LOG: [
    'EQALogID','Program','Round','TestID','TestName','SentValue',
    'ExpectedValue','Bias','SDI','Status','Comments','LoggedAt','LoggedBy'
  ]
};

/** Returns {headerName: colIndex} map for any LAB_* sheet. */
function labHeaderMap(sheet) {
  var name = sheet.getName();
  if (LAB_SCHEMA[name]) {
    var m = {};
    LAB_SCHEMA[name].forEach(function(h, i) { m[h] = i; });
    return m;
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var m2 = {};
  headers.forEach(function(h, i) { if (h) m2[String(h)] = i; });
  return m2;
}

/** Appends one row to LAB_AUDIT_LOG. All strings. Never throws. */
function labAudit(action, entityType, entityId, oldVal, newVal) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getSheetByName(LAB.AUDIT_LOG); if (!s) return;
    var m = labHeaderMap(s); var nc = LAB_SCHEMA.LAB_AUDIT_LOG.length;
    var row = new Array(nc).fill('');
    row[m['AuditID']]     = 'AUD-' + Utilities.getUuid().substring(0,8).toUpperCase();
    row[m['Timestamp']]   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    row[m['Action']]      = String(action||'');
    row[m['EntityType']]  = String(entityType||'');
    row[m['EntityID']]    = String(entityId==null?'':entityId);
    row[m['OldValue']]    = oldVal==null?'':(typeof oldVal==='object'?JSON.stringify(oldVal):String(oldVal));
    row[m['NewValue']]    = newVal==null?'':(typeof newVal==='object'?JSON.stringify(newVal):String(newVal));
    row[m['PerformedBy']] = Session.getActiveUser().getEmail()||'SYSTEM';
    s.appendRow(row);
  } catch(e) { Logger.log('labAudit: '+e.message); }
}

/** Order-level status machine. */
var LAB_ORDER_STATUS_OK = {
  PENDING:            ['BILLED','SAMPLE_COLLECTED','CANCELLED'],
  BILLED:             ['SAMPLE_COLLECTED','CANCELLED'],
  RECOLLECT:          ['SAMPLE_COLLECTED','CANCELLED'],
  SAMPLE_COLLECTED:   ['IN_PROCESS','RECOLLECT'],
  IN_PROCESS:         ['RESULT_ENTERED','PARTIAL_VERIFIED'],
  RESULT_ENTERED:     ['VERIFIED','PARTIAL_VERIFIED','IN_PROCESS'],
  PARTIAL_VERIFIED:   ['VERIFIED','RESULT_ENTERED','IN_PROCESS'],
  VERIFIED:           ['PARTIAL_DISPATCHED','REPORT_DISPATCHED'],
  PARTIAL_DISPATCHED: ['REPORT_DISPATCHED'],
  REPORT_DISPATCHED:  ['AMENDED'],
  AMENDED:            ['REPORT_DISPATCHED'],
  CANCELLED:          []
};

/** Test-level status machine. */
var LAB_TEST_STATUS_OK = {
  PENDING:        ['IN_PROCESS','CANCELLED'],
  IN_PROCESS:     ['RESULT_ENTERED'],
  RESULT_ENTERED: ['VERIFIED','IN_PROCESS'],
  VERIFIED:       ['DISPATCHED'],
  DISPATCHED:     [],
  CANCELLED:      []
};

/** Creates all 13 LAB_* sheets. Idempotent — safe to re-run. */
function setupLabDatabase() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var created = [], existing = [];
    Object.keys(LAB_SCHEMA).forEach(function(name) {
      if (ss.getSheetByName(name)) { existing.push(name); return; }
      var sh = ss.insertSheet(name);
      var hdrs = LAB_SCHEMA[name];
      sh.getRange(1,1,1,hdrs.length).setValues([hdrs]);
      sh.getRange(1,1,1,hdrs.length).setBackground('#0f2744').setFontColor('#7dd3fc').setFontWeight('bold').setFontSize(9);
      sh.setFrozenRows(1);
      created.push(name);
    });
    SpreadsheetApp.flush();
    return { success:true, message:'Setup complete. Created: '+(created.length||'none (all existed)'), created:created, existing:existing };
  } catch(e) { return { success:false, message:'setupLabDatabase: '+e.message }; }
  finally { try { lock.releaseLock(); } catch(ex){} }
}

function seedStarterCatalog() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(LAB.CATALOG);
    if (!sh) return { success:false, message:'Run setupLabDatabase() first.' };
    if (sh.getLastRow() > 1) return { success:true, message:'Catalog already has data — skipped.' };
    var m = labHeaderMap(sh); var nc = LAB_SCHEMA.LAB_TEST_CATALOG.length;
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    function R(d) { return _catRow(m, nc, d, now); }
    var rows = [
      // CBC
      R({id:'LABTEST-CBC001',code:'CBC',  name:'Complete Blood Count',   type:'PANEL',     dept:'HAEMATOLOGY',  sample:'BLOOD_EDTA',    price:250,tat:60, sort:1}),
      R({id:'LABTEST-CBC002',code:'HB',   name:'Haemoglobin',            type:'PARAMETER', parent:'LABTEST-CBC001',rtype:'NUMERIC',unit:'g/dL',   mlo:13,  mhi:17,  flo:11.5,fhi:15.5,clo:7,   chi:20,  sort:1}),
      R({id:'LABTEST-CBC003',code:'RBC',  name:'RBC Count',              type:'PARAMETER', parent:'LABTEST-CBC001',rtype:'NUMERIC',unit:'mill/µL',mlo:4.5, mhi:5.9, flo:3.8, fhi:5.2, clo:2,   chi:8,   sort:2}),
      R({id:'LABTEST-CBC004',code:'WBC',  name:'Total WBC Count',        type:'PARAMETER', parent:'LABTEST-CBC001',rtype:'NUMERIC',unit:'cells/µL',mlo:4000,mhi:11000,flo:4000,fhi:11000,clo:2000,chi:30000,sort:3}),
      R({id:'LABTEST-CBC005',code:'PLT',  name:'Platelet Count',         type:'PARAMETER', parent:'LABTEST-CBC001',rtype:'NUMERIC',unit:'lakh/µL',mlo:1.5, mhi:4.5, flo:1.5, fhi:4.5, clo:0.5, chi:10,  sort:4}),
      R({id:'LABTEST-CBC006',code:'MCV',  name:'MCV',                    type:'PARAMETER', parent:'LABTEST-CBC001',rtype:'NUMERIC',unit:'fL',     mlo:80,  mhi:100, flo:80,  fhi:100, sort:5}),
      R({id:'LABTEST-CBC007',code:'MCH',  name:'MCH',                    type:'PARAMETER', parent:'LABTEST-CBC001',rtype:'NUMERIC',unit:'pg',     mlo:27,  mhi:33,  flo:27,  fhi:33,  sort:6}),
      R({id:'LABTEST-CBC008',code:'MCHC', name:'MCHC',                   type:'PARAMETER', parent:'LABTEST-CBC001',rtype:'NUMERIC',unit:'g/dL',   mlo:31.5,mhi:36,  flo:31.5,fhi:36,  sort:7}),
      R({id:'LABTEST-CBC009',code:'RDW',  name:'RDW-CV',                 type:'PARAMETER', parent:'LABTEST-CBC001',rtype:'NUMERIC',unit:'%',      mlo:11.5,mhi:14.5,flo:11.5,fhi:14.5,sort:8}),
      R({id:'LABTEST-CBC010',code:'PDW',  name:'PDW',                    type:'PARAMETER', parent:'LABTEST-CBC001',rtype:'NUMERIC',unit:'fL',     mlo:9,   mhi:17,  flo:9,   fhi:17,  sort:9}),
      // LFT
      R({id:'LABTEST-LFT001',code:'LFT',  name:'Liver Function Tests',   type:'PANEL',     dept:'BIOCHEMISTRY',  sample:'BLOOD_PLAIN',  price:350,tat:90, sort:2}),
      R({id:'LABTEST-LFT002',code:'TBIL', name:'Total Bilirubin',        type:'PARAMETER', parent:'LABTEST-LFT001',rtype:'NUMERIC',unit:'mg/dL',mlo:0.3,mhi:1.2, flo:0.3,fhi:1.2, chi:15,  sort:1}),
      R({id:'LABTEST-LFT003',code:'DBIL', name:'Direct Bilirubin',       type:'PARAMETER', parent:'LABTEST-LFT001',rtype:'NUMERIC',unit:'mg/dL',mlo:0,  mhi:0.3, flo:0,  fhi:0.3, sort:2}),
      R({id:'LABTEST-LFT004',code:'IBIL', name:'Indirect Bilirubin',     type:'PARAMETER', parent:'LABTEST-LFT001',rtype:'NUMERIC',unit:'mg/dL',mlo:0.2,mhi:0.9, flo:0.2,fhi:0.9, sort:3}),
      R({id:'LABTEST-LFT005',code:'SGOT', name:'SGOT / AST',             type:'PARAMETER', parent:'LABTEST-LFT001',rtype:'NUMERIC',unit:'U/L',  mlo:10, mhi:40,  flo:10, fhi:40,  chi:1000, sort:4}),
      R({id:'LABTEST-LFT006',code:'SGPT', name:'SGPT / ALT',             type:'PARAMETER', parent:'LABTEST-LFT001',rtype:'NUMERIC',unit:'U/L',  mlo:7,  mhi:56,  flo:7,  fhi:56,  chi:1000, sort:5}),
      R({id:'LABTEST-LFT007',code:'ALP',  name:'Alkaline Phosphatase',   type:'PARAMETER', parent:'LABTEST-LFT001',rtype:'NUMERIC',unit:'U/L',  mlo:44, mhi:147, flo:44, fhi:147, sort:6}),
      R({id:'LABTEST-LFT008',code:'TP',   name:'Total Protein',          type:'PARAMETER', parent:'LABTEST-LFT001',rtype:'NUMERIC',unit:'g/dL', mlo:6.3,mhi:8.2, flo:6.3,fhi:8.2, sort:7}),
      // RFT
      R({id:'LABTEST-RFT001',code:'RFT',  name:'Renal Function Tests',   type:'PANEL',     dept:'BIOCHEMISTRY',  sample:'BLOOD_PLAIN',  price:300,tat:90, sort:3}),
      R({id:'LABTEST-RFT002',code:'BUN',  name:'Blood Urea Nitrogen',    type:'PARAMETER', parent:'LABTEST-RFT001',rtype:'NUMERIC',unit:'mg/dL',mlo:7,  mhi:20,  flo:7,  fhi:20,  chi:100,  sort:1}),
      R({id:'LABTEST-RFT003',code:'CREAT',name:'Serum Creatinine',       type:'PARAMETER', parent:'LABTEST-RFT001',rtype:'NUMERIC',unit:'mg/dL',mlo:0.7,mhi:1.3, flo:0.5,fhi:1.1, chi:15,   sort:2}),
      R({id:'LABTEST-RFT004',code:'UA',   name:'Uric Acid',              type:'PARAMETER', parent:'LABTEST-RFT001',rtype:'NUMERIC',unit:'mg/dL',mlo:3.4,mhi:7.0, flo:2.4,fhi:6.0, sort:3}),
      R({id:'LABTEST-RFT005',code:'EGFR', name:'eGFR',                   type:'PARAMETER', parent:'LABTEST-RFT001',rtype:'NUMERIC',unit:'mL/min',mlo:90,mhi:null,flo:90,fhi:null, clo:15,   sort:4}),
      // Lipid
      R({id:'LABTEST-LIP001',code:'LIPID',name:'Lipid Profile',          type:'PANEL',     dept:'BIOCHEMISTRY',  sample:'BLOOD_PLAIN',  price:400,tat:90, sort:4}),
      R({id:'LABTEST-LIP002',code:'TCHOL',name:'Total Cholesterol',      type:'PARAMETER', parent:'LABTEST-LIP001',rtype:'NUMERIC',unit:'mg/dL',mhi:200, fhi:200, sort:1}),
      R({id:'LABTEST-LIP003',code:'TRIG', name:'Triglycerides',          type:'PARAMETER', parent:'LABTEST-LIP001',rtype:'NUMERIC',unit:'mg/dL',mhi:150, fhi:150, sort:2}),
      R({id:'LABTEST-LIP004',code:'HDL',  name:'HDL Cholesterol',        type:'PARAMETER', parent:'LABTEST-LIP001',rtype:'NUMERIC',unit:'mg/dL',mlo:40, flo:50,  sort:3}),
      R({id:'LABTEST-LIP005',code:'LDL',  name:'LDL Cholesterol',        type:'PARAMETER', parent:'LABTEST-LIP001',rtype:'NUMERIC',unit:'mg/dL',mhi:100, fhi:100, sort:4}),
      // Individuals
      R({id:'LABTEST-FBS001',code:'FBS',  name:'Fasting Blood Sugar',    type:'INDIVIDUAL',dept:'BIOCHEMISTRY',  sample:'BLOOD_FLUORIDE',rtype:'NUMERIC',unit:'mg/dL',mlo:70,mhi:99, flo:70,fhi:99, clo:40,chi:500, price:120,tat:60, sort:1}),
      R({id:'LABTEST-RBS001',code:'RBS',  name:'Random Blood Sugar',     type:'INDIVIDUAL',dept:'BIOCHEMISTRY',  sample:'BLOOD_FLUORIDE',rtype:'NUMERIC',unit:'mg/dL',mhi:140,fhi:140, clo:40,chi:500, price:100,tat:60, sort:2}),
      R({id:'LABTEST-HBA001',code:'HBA1C',name:'HbA1c',                  type:'INDIVIDUAL',dept:'BIOCHEMISTRY',  sample:'BLOOD_EDTA',    rtype:'NUMERIC',unit:'%',   mlo:4,mhi:5.7,flo:4,fhi:5.7, chi:15, price:450,tat:180,sort:3}),
      R({id:'LABTEST-TSH001',code:'TSH',  name:'TSH',                    type:'INDIVIDUAL',dept:'ENDOCRINOLOGY', sample:'BLOOD_PLAIN',   rtype:'NUMERIC',unit:'mIU/L',mlo:0.5,mhi:4.5,flo:0.5,fhi:4.5, chi:100, price:350,tat:120,sort:1}),
      R({id:'LABTEST-T3001', code:'T3',   name:'T3',                     type:'INDIVIDUAL',dept:'ENDOCRINOLOGY', sample:'BLOOD_PLAIN',   rtype:'NUMERIC',unit:'ng/dL',mlo:80,mhi:200,flo:80,fhi:200, price:200,tat:120,sort:2}),
      R({id:'LABTEST-T4001', code:'T4',   name:'T4',                     type:'INDIVIDUAL',dept:'ENDOCRINOLOGY', sample:'BLOOD_PLAIN',   rtype:'NUMERIC',unit:'ug/dL',mlo:5.1,mhi:14.1,flo:5.1,fhi:14.1, price:200,tat:120,sort:3}),
      R({id:'LABTEST-HIV001',code:'HIV',  name:'HIV 1 & 2 Antibodies',   type:'INDIVIDUAL',dept:'SEROLOGY',      sample:'BLOOD_PLAIN',   rtype:'QUALITATIVE', price:450,tat:60, sort:1,consent:true}),
      R({id:'LABTEST-HBS001',code:'HBSAG',name:'HBsAg',                  type:'INDIVIDUAL',dept:'SEROLOGY',      sample:'BLOOD_PLAIN',   rtype:'QUALITATIVE', price:300,tat:60, sort:2,consent:true}),
      R({id:'LABTEST-TRP001',code:'TROPI',name:'Troponin I (hsTnI)',     type:'INDIVIDUAL',dept:'BIOCHEMISTRY',  sample:'BLOOD_PLAIN',   rtype:'NUMERIC',unit:'ng/L',mhi:26,fhi:16, chi:2000, price:900,tat:60, sort:4}),
      R({id:'LABTEST-URN001',code:'URIN', name:'Urine Routine & Microscopy',type:'INDIVIDUAL',dept:'CLINICAL_PATH',sample:'URINE_RANDOM',  rtype:'DESCRIPTIVE', price:150,tat:60, sort:1}),
      // Package
      R({id:'LABTEST-PKG001',code:'DIAB', name:'Diabetes Panel',         type:'PACKAGE',   dept:'BIOCHEMISTRY',  components:'LABTEST-FBS001,LABTEST-HBA001', price:520,tat:180,sort:1})
    ];
    sh.getRange(2,1,rows.length,nc).setValues(rows);
    SpreadsheetApp.flush();
    return { success:true, message:rows.length+' entries seeded.' };
  } catch(e) { return { success:false, message:'seedStarterCatalog: '+e.message }; }
}

function _catRow(m, nc, d, now) {
  var row = new Array(nc).fill('');
  function n(v) { return (v===''||v==null)?'':Number(v); }
  row[m['TestID']]            = String(d.id);
  row[m['TestCode']]          = String(d.code);
  row[m['TestName']]          = String(d.name);
  row[m['TestType']]          = String(d.type);
  row[m['ParentPanelID']]     = String(d.parent||'');
  row[m['Department']]        = String(d.dept||'BIOCHEMISTRY');
  row[m['SampleType']]        = String(d.sample||'');
  row[m['ResultType']]        = String(d.rtype||'');
  row[m['Unit']]              = String(d.unit||'');
  row[m['MaleRefLow']]        = n(d.mlo);
  row[m['MaleRefHigh']]       = n(d.mhi);
  row[m['FemaleRefLow']]      = n(d.flo);
  row[m['FemaleRefHigh']]     = n(d.fhi);
  row[m['PaediatricRefText']] = String(d.paed||'');
  row[m['CriticalLow']]       = n(d.clo);
  row[m['CriticalHigh']]      = n(d.chi);
  row[m['Price']]             = n(d.price)||0;
  row[m['TAT_Minutes']]       = n(d.tat)||0;
  row[m['ComponentTestIDs']]  = String(d.components||'');
  row[m['RequiresConsent']]   = !!(d.consent);
  row[m['SortOrder']]         = n(d.sort)||0;
  row[m['IsActive']]          = true;
  row[m['CreatedAt']]         = now;
  row[m['CreatedBy']]         = 'SYSTEM_SEED';
  return row;
}