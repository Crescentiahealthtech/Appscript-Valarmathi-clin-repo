/* The pharmacy counter, exercised against a pretend spreadsheet.

   The bill is where stock and money move together, and where the loopholes
   were: a price the browser chose, a batch looked up by the wrong row, a
   Schedule H drug sold with no prescriber, every bill signed by the owner.
   This loads the real Pharmacy.gs, PharmacyDashboardLogic.gs,
   Billing_Ledger.gs and Shared_Dates.gs into one V8 context and walks the
   counter: stock in, a bill, the register, the ledger, a credit settled.

       node tools/pharmacy.js                 (part of ./tools/check.sh)  */
const fs = require('fs'), path = require('path'), vm = require('vm'), crypto = require('crypto');
const ROOT = process.env.REPO || path.resolve(__dirname, '..');
process.env.TZ = 'Asia/Kolkata';

function makeSheet(name, rows) {
  const data = rows.map(r => r.slice());
  let maxCols = 26;
  const width = () => Math.max(0, ...data.map(r => r.length));
  const cell = (r, c) => (data[r - 1] && data[r - 1][c - 1] !== undefined) ? data[r - 1][c - 1] : '';
  const set = (r, c, v) => {
    if (c > maxCols) throw new Error('Range outside the sheet: column ' + c + ' of ' + maxCols);
    while (data.length < r) data.push([]); const row = data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v;
  };
  const range = (r, c, nr, nc) => ({
    getValue: () => cell(r, c),
    getValues: () => Array.from({ length: nr || 1 }, (_, i) => Array.from({ length: nc || 1 }, (_, j) => cell(r + i, c + j))),
    setValue(v) { set(r, c, v); return this; },
    setValues(vs) { vs.forEach((row, i) => row.forEach((v, j) => set(r + i, c + j, v))); return this; },
    setNumberFormat() { return this; }, setFontWeight() { return this; }, setBackground() { return this; }
  });
  return {
    data, getName: () => name,
    getLastRow: () => data.length, getLastColumn: () => width(),
    getMaxColumns: () => maxCols, insertColumnsAfter(a, n) { maxCols += n; },
    getRange: (r, c, nr, nc) => range(r, c, nr, nc),
    getDataRange: () => ({ getValues: () => data.map(r => { const o = r.slice(); while (o.length < width()) o.push(''); return o; }) }),
    appendRow(r) { data.push(r.slice()); }, setFrozenRows() {}, deleteRow(n) { data.splice(n - 1, 1); },
    trim(n) { maxCols = n; }
  };
}

const INV_HEAD = ['Timestamp', 'Brand Name', 'Generic Name', 'Type', 'Qty', 'Unit', 'Batch No', 'Expiry Date',
                  'Rack Location', 'Buy Price', 'MRP', 'GST %', 'Manufacturer', 'Supplier'];
// Stable arrays: the sandbox holds these same objects for the whole run.
const audit = [], ipTab = [];
let SS;
function fresh() {
  audit.length = 0; ipTab.length = 0;
  const inv = makeSheet('Pharmacy_Inventory', [INV_HEAD,
    [new Date(2026, 0, 1), 'Dolo 650', 'Paracetamol', 'Tablet', 100, 'Strips', 'B1', new Date(2028, 5, 1), 'R1', 20, 30, 12, 'Micro', 'Sup'],
    [new Date(2026, 0, 1), 'Alprax 0.25', 'Alprazolam', 'Tablet', 5, 'Strips', 'AX1', '2028-03', 'R2', 40, 60, 12, 'Torrent', 'Sup'],
    [new Date(2026, 0, 1), 'Alprax 0.25', 'Alprazolam', 'Tablet', 50, 'Strips', 'AX1', '2028-03', 'R2', 40, 60, 12, 'Torrent', 'Sup']
  ]);
  inv.trim(14);                                   // a sheet whose spare columns were deleted
  const sheets = {
    Pharmacy_Inventory: inv,
    Patients: makeSheet('Patients', [['Patient_ID', 'Password', 'Name'], ['LMTVS0001', '', 'Meena']])
  };
  SS = { sheets, getSheetByName: n => sheets[n] || null,
         insertSheet: n => (sheets[n] = makeSheet(n, [])) };
}

const env = {
  console, Date, JSON, Math,
  Utilities: {
    getUuid: () => crypto.randomUUID(),
    formatDate: (d, tz, p) => {
      const o = new Intl.DateTimeFormat('en-GB', { timeZone: tz || 'Asia/Kolkata', year: 'numeric', month: '2-digit',
        day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).formatToParts(d)
        .reduce((a, x) => (a[x.type] = x.value, a), {});
      const M = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      const h = +o.hour % 24, h12 = ((h + 11) % 12) + 1;
      return p.replace('yyyy', o.year).replace('yy', o.year.slice(2)).replace('MMMM', M[+o.month - 1])
        .replace('MMM', M[+o.month - 1]).replace('MM', o.month).replace('dd', o.day)
        .replace('HH', String(h).padStart(2, '0')).replace('hh', String(h12).padStart(2, '0'))
        .replace('mm', o.minute).replace('ss', o.second).replace(' a', h < 12 ? ' AM' : ' PM');
    }
  },
  Session: { getScriptTimeZone: () => 'Asia/Kolkata' },
  LockService: { getScriptLock: () => ({ waitLock() {}, tryLock() { return true; }, releaseLock() {} }) },
  Logger: { log() {} },
  SpreadsheetApp: { getActiveSpreadsheet: () => SS, flush() {} }
};
vm.createContext(env);
vm.runInContext(`
  var __actors = {
    'T-PHARM': { username: 'pharm1', displayName: 'Pharm One', role: 'pharmacist',
                 permissions: ['pharmacy.read','pharmacy.dispense','pharmacy.bill','pharmacy.register','pharmacy.stock_add','pharmacy.stock_edit'] },
    'T-NURSE': { username: 'nurse1', displayName: 'Nurse', role: 'nurse', permissions: ['pharmacy.read'] }
  };
  var CRESC_CURRENT_ACTOR = null;
  function crescRequire_(token, perm) {
    var a = __actors[token] || (token === undefined ? CRESC_CURRENT_ACTOR : null);
    if (!a) throw new Error('FORBIDDEN: your session has expired. Please sign in again.');
    var need = perm ? (typeof perm === 'string' ? [perm] : perm) : [];
    if (need.length && !need.some(function (p) { return a.permissions.indexOf(p) !== -1; }))
      throw new Error('FORBIDDEN: your role (' + a.role + ') cannot do that.');
    CRESC_CURRENT_ACTOR = a;
    return a;
  }
  function crescActor_(t) { return __actors[t] || CRESC_CURRENT_ACTOR; }
  function cresc_reason_(e) { return String((e && e.message) || e).replace(/^FORBIDDEN:\\s*/, ''); }
  function cresc_actorName_(f) { return CRESC_CURRENT_ACTOR ? CRESC_CURRENT_ACTOR.username : (f || 'SYSTEM'); }
  function logAudit_(a, ev, t, id, d) { __audit.push(ev); }
  function getActiveDoctors_() { return [{ name: 'Dr. Valarmathi' }]; }
  function ipc_activeAdmissionByPatient_(pid) { return pid === 'LMTVS0001' ? 'IP2609-0001' : null; }
  function billChargeToIp(p) { __ip.push(p); return { success: true }; }
  function ipc_markChargePaidAtCounter_(src, ref) { __ip.push({ paidAtCounter: ref }); return true; }
`, Object.assign(env, { __audit: audit, __ip: ipTab }));
for (const f of ['Shared_Dates.gs', 'Pharmacy.gs', 'PharmacyDashboardLogic.gs', 'Billing_Ledger.gs']) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), env, { filename: f });
}
const run = code => { vm.runInContext('CRESC_CURRENT_ACTOR = null;', env); return vm.runInContext(code, env); };
const J = JSON.stringify;

let pass = 0; const fails = [];
function check(label, got, want) {
  if (J(got) === J(want)) pass++;
  else fails.push(label + ': got ' + J(got) + ', wanted ' + J(want) + (r && r.message ? '  (' + r.message + ')' : ''));
}
const inv = () => SS.sheets.Pharmacy_Inventory.data;

// ---- stock in ----------------------------------------------------------------
fresh();
let r = run(`updatePharmacyStock({ token: 'T-PHARM', rowId: 3, brandName: 'Alprax 0.25', batch: 'AX1',
             stock: 5, expiry: '2028-03', rack: 'R2', mrp: 60, gst: 12, schedule: 'H1' })`);
check('marking a schedule succeeds', r.success, true);
check('the schedule column is added even on a trimmed sheet', inv()[0][14], 'Schedule');
check('the schedule is set on every batch of the medicine', [inv()[2][14], inv()[3][14]], ['H1', 'H1']);
r = run(`updatePharmacyStock({ token: 'T-PHARM', rowId: 2, brandName: 'Alprax 0.25', batch: 'AX1', stock: 1, mrp: 60, gst: 12 })`);
check('an edit aimed at a row that now holds another medicine is refused', [r.success, inv()[1][4]], [false, 100]);
r = run(`savePharmacyInventory({ token: 'T-PHARM', medicineName: 'Dolo 650', batchNo: 'b1', qty: 20, mrp: 30, buyPrice: 20, gst: 12 })`);
check('the same batch received again tops up its row', [r.success, inv()[1][4], inv().length], [true, 120, 4]);
r = run(`savePharmacyInventory({ token: 'T-PHARM', medicineName: 'Dolo 650', batchNo: 'B1', qty: 5, mrp: 35, buyPrice: 20, gst: 12 })`);
check('...but not at a different MRP', r.success, false);
r = run(`savePharmacyInventory({ token: 'T-PHARM', medicineName: 'Zolfresh', batchNo: 'z9', qty: -3, mrp: 30, buyPrice: 20, gst: 12 })`);
check('a negative quantity is refused', r.success, false);

// ---- the bill ------------------------------------------------------------------
const bill = extra => run(`processPharmacyBill(${J(Object.assign({
  billUuid: 'U-' + crypto.randomUUID(), patientId: 'LMTVS0001', patientName: 'Meena', doctor: 'Self / OTC',
  payMode: 'CASH', discount: 0,
  billedItems: [{ drug: 'Alprax 0.25', batch: 'AX1', qty: 8, rate: 1, gst: 0, rowId: 4 }]
}, extra))}, 'T-PHARM')`);

r = bill({});
check('a Schedule H1 drug with "Self / OTC" is refused', [r.success, r.code], [false, 'SCHEDULE_DOCTOR_REQUIRED']);
check('...and no stock moves', [inv()[2][4], inv()[3][4]], [5, 50]);
r = bill({ doctor: 'Dr. Valarmathi', patientName: '' });
check('...and nor without a patient name', r.code, 'SCHEDULE_PATIENT_REQUIRED');
r = bill({ doctor: 'Dr. Valarmathi' });
check('with a named doctor the bill is saved', r.success, true);
check('the price is the shelf MRP, not the one the browser sent', r.print.net, 480);
check('stock comes off the row the batch was picked from', [inv()[2][4], inv()[3][4]], [5, 42]);
const items = SS.sheets.Pharmacy_Invoice_Items.data;
check('the invoice line records its schedule', items[items.length - 1][15], 'H1');
const heads = SS.sheets.Pharmacy_Invoices.data;
check('the bill is signed by the pharmacist, not the owner', heads[heads.length - 1][19], 'pharm1');
check('a scheduled sale is audited', audit.indexOf('PHARMACY_SCHEDULED_SALE') !== -1, true);

r = bill({ doctor: 'Dr. X', billedItems: [{ drug: 'Alprax 0.25', batch: 'AX1', qty: 50, rowId: 4 }] });
check('selling more than a row holds is refused', r.success, false);
r = run(`processPharmacyBill({ billUuid: 'U2', payMode: 'BITCOIN', billedItems: [{ drug: 'Dolo 650', batch: 'B1', qty: 1 }] }, 'T-PHARM')`);
check('an unknown payment mode is refused', r.success, false);
r = run(`processPharmacyBill({ billUuid: 'U3', billedItems: [{ drug: 'Dolo 650', batch: 'B1', qty: 1 }] }, 'T-NURSE')`);
check('a nurse cannot bill', r.success, false);
r = bill({ payMode: 'CREDIT', doctor: 'Dr. Y', billedItems: [{ drug: 'Dolo 650', batch: 'B1', qty: 2, rowId: 2 }] });
check('a credit bill for an admitted patient goes on the IP tab', [r.success, ipTab.length, ipTab[0] && ipTab[0].amount], [true, 1, 60]);
const creditNo = r.invoiceNo;

// ---- the register ---------------------------------------------------------------
const today = env.Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
r = run(`getScheduleDrugRegister({ from: '${today}', to: '${today}' }, 'T-PHARM')`);
check('the register lists the scheduled sale', r.rows.map(x => [x.sNo, x.patientName, x.doctor, x.drug, x.qty, x.schedule]),
      [[1, 'Meena', 'Dr. Valarmathi', 'Alprax 0.25', 8, 'H1']]);
r = run(`getScheduleDrugRegister({ from: '${today}', to: '${today}', schedules: ['X'] }, 'T-PHARM')`);
check('...and only the schedules asked for', r.rows.length, 0);
r = run(`getScheduleDrugRegister({ from: '${today}', to: '${today}' }, 'T-NURSE')`);
check('a nurse cannot read the register', r.success, false);

// ---- the ledger and the credit settled -------------------------------------------
r = run(`getPharmacyLedger({ status: 'UNSETTLED' }, 'T-PHARM')`);
check('the unsettled ledger shows the credit bill', r.rows.map(x => [x.id, x.status, x.balance]), [[creditNo, 'UNPAID', 60]]);
r = run(`settleCreditBill({ invoiceNo: '${creditNo}', payMode: 'UPI', txnId: 'U123' }, 'T-PHARM')`);
check('settling a credit bill takes it off the IP tab', [r.success, ipTab[1] && ipTab[1].paidAtCounter], [true, creditNo]);
check('...and records who took the money', heads[heads.length - 1][22], 'Pharm One');
r = run(`getPharmacyLedger({ status: 'ALL' }, 'T-PHARM')`);
check('the ledger totals what was collected, by mode', [r.totals.count, r.totals.paid, r.totals.byMode.CASH, r.totals.byMode.UPI], [2, 540, 480, 60]);

if (fails.length) {
  fails.forEach(f => console.log('FAIL  ' + f));
  console.log(pass + ' passed, ' + fails.length + ' failed.');
  process.exit(1);
}
console.log(pass + ' pharmacy checks passed (stock in, schedule marking, the prescription rule, shelf prices, batch rows, the register, the ledger, credit settlement).');
