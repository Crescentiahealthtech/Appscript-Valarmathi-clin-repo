// ============================================================
// 📊 ADMIN DASHBOARD — LIVE AGGREGATOR  (READ-ONLY)
// Crescentia HealthTech / CresRx
// ------------------------------------------------------------
// Single server round-trip for the whole dashboard.
// - Reads each source sheet ONCE (perf).
// - Role-gated: financial keys are never returned to non-finance roles.
// - CacheService (5 min) so repeat loads are instant.
// - No LockService: this module never writes.
// ============================================================

/**
 * Master entry point called by Admin_Dashboard.html
 * @param {string} role  - logged-in role (admin/doctor/nurse/...)
 * @param {boolean} bust - true to bypass cache (manual refresh)
 * @return {{success:boolean, data:Object, message:string}}
 */
function getDashboardData(role, bust) {
  try {
    role = (role || "").toString().trim().toLowerCase();

    // Only these roles may ever receive revenue figures (server-side gate).
    var isFinancial = (role === 'admin' || role === 'accounts' ||
                       role === 'receptionist' || role === 'reception');

    var cache = CacheService.getScriptCache();
    var cacheKey = 'DASH_' + (isFinancial ? 'FIN' : 'CLIN');

    if (!bust) {
      var hit = cache.get(cacheKey);
      if (hit) return { success: true, data: JSON.parse(hit), message: "cache" };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone() || "Asia/Kolkata";
    var now = new Date();
    var todayKey = Utilities.formatDate(now, tz, "yyyy-MM-dd");

    var data = {
      meta: {
        generatedAt: Utilities.formatDate(now, tz, "dd MMM yyyy • hh:mm a"),
        role: role,
        financial: isFinancial
      },
      census:    _dashCensus_(ss),
      footfall:  _dashFootfall_(ss, tz),
      operations:_dashOps_(ss, tz, todayKey),
      staffCount:_dashStaff_(ss)
    };

    if (isFinancial) {
      data.revenue = _dashRevenue_(ss, tz, todayKey);
    }

    cache.put(cacheKey, JSON.stringify(data), 300); // 5 minutes
    return { success: true, data: data, message: "fresh" };

  } catch (e) {
    return { success: false, message: "Dashboard load failed: " + e.message, data: null };
  }
}

/** Call this from any write-path module after data changes if you want instant freshness. */
function invalidateDashboardCache() {
  try { CacheService.getScriptCache().removeAll(['DASH_FIN', 'DASH_CLIN']); } catch (e) {}
}

// ------------------------------------------------------------
// CENSUS & BEDS  (headline KPI band)
// Master_Beds: Bed_ID[0] Ward[1] Status[2] Patient_ID[3] Patient_Name[4] DOA[5] IP_No[6]
// IP_Admissions: IP_No[0] PatientID[1] Name[2] AgeSex[3] DOA[4] ... Status[11] DOD[12]
// ------------------------------------------------------------
function _dashCensus_(ss) {
  var out = { totalBeds: 0, occupied: 0, available: 0, occupancyPct: 0,
              wards: [], currentInpatients: 0, hasData: false };

  var bedSheet = ss.getSheetByName("Master_Beds");
  if (bedSheet && bedSheet.getLastRow() > 1) {
    var d = bedSheet.getDataRange().getValues();
    var wardMap = {};
    for (var i = 1; i < d.length; i++) {
      if (!d[i][0]) continue;
      out.totalBeds++;
      var ward = (d[i][1] || "GEN").toString().trim();
      var status = (d[i][2] || "Available").toString().trim().toUpperCase();
      var occ = (status !== "AVAILABLE"); // Occupied / Cleaning / Blocked => not free
      if (occ) out.occupied++; else out.available++;
      if (!wardMap[ward]) wardMap[ward] = { ward: ward, total: 0, occupied: 0 };
      wardMap[ward].total++;
      if (occ) wardMap[ward].occupied++;
    }
    out.wards = Object.keys(wardMap).map(function (k) { return wardMap[k]; })
                      .sort(function (a, b) { return a.ward < b.ward ? -1 : 1; });
    out.occupancyPct = out.totalBeds ? Math.round((out.occupied / out.totalBeds) * 100) : 0;
    out.hasData = out.totalBeds > 0;
  }

  var ipSheet = ss.getSheetByName("IP_Admissions");
  if (ipSheet && ipSheet.getLastRow() > 1) {
    var ip = ipSheet.getDataRange().getValues();
    for (var j = 1; j < ip.length; j++) {
      if ((ip[j][11] || "").toString().trim().toUpperCase() === "ADMITTED") out.currentInpatients++;
    }
  }
  return out;
}

// ------------------------------------------------------------
// FOOTFALL MATRIX  (OP vs IP, by year → 12 months)  — the Phase-1 chart
// OP  = Appointments with real attendance status (not raw bookings)
// IP  = IP_Admissions by DOA
// Appointments: ApptID[0] PatID[1] Name[2] Date[3] Time[4] Purpose[5] Status[6] Fee[7]
// ------------------------------------------------------------
function _dashFootfall_(ss, tz) {
  var op = {}, ip = {}, years = {};

  var a = ss.getSheetByName("Appointments");
  if (a && a.getLastRow() > 1) {
    var ad = a.getDataRange().getValues();
    for (var i = 1; i < ad.length; i++) {
      var st = (ad[i][6] || "").toString();
      if (st !== "Arrived" && st !== "In-Progress" && st !== "Completed") continue;
      var dt = _dashToDate_(ad[i][3]); if (!dt) continue;
      var y = dt.getFullYear(), m = dt.getMonth();
      years[y] = true;
      if (!op[y]) op[y] = _zeros12_();
      op[y][m]++;
    }
  }

  var ipS = ss.getSheetByName("IP_Admissions");
  if (ipS && ipS.getLastRow() > 1) {
    var id = ipS.getDataRange().getValues();
    for (var k = 1; k < id.length; k++) {
      var dt2 = _dashToDate_(id[k][4]); if (!dt2) continue;
      var y2 = dt2.getFullYear(), m2 = dt2.getMonth();
      years[y2] = true;
      if (!ip[y2]) ip[y2] = _zeros12_();
      ip[y2][m2]++;
    }
  }

  var yrList = Object.keys(years).map(Number).sort(function (a, b) { return a - b; });
  return { years: yrList, op: op, ip: ip, hasData: yrList.length > 0 };
}

// ------------------------------------------------------------
// OPERATIONS / ALERTS (today's appointments, stock, lab queue)
// Pharmacy_Inventory: qty[4], expiry[7]  (per fetchBillableStock layout)
// LAB_ORDERS: status column resolved by header lookup (no hard-coded guess)
// ------------------------------------------------------------
function _dashOps_(ss, tz, todayKey) {
  var out = { apptToday: 0, apptPending: 0, apptCompleted: 0,
              lowStock: 0, expiringSoon: 0, labPending: 0 };

  var a = ss.getSheetByName("Appointments");
  if (a && a.getLastRow() > 1) {
    var ad = a.getDataRange().getValues();
    for (var i = 1; i < ad.length; i++) {
      var dt = _dashToDate_(ad[i][3]); if (!dt) continue;
      if (Utilities.formatDate(dt, tz, "yyyy-MM-dd") !== todayKey) continue;
      var st = (ad[i][6] || "").toString();
      if (st === "Blocked" || st === "Cancelled" || st === "DELETE") continue;
      out.apptToday++;
      if (st === "Completed") out.apptCompleted++;
      else if (st === "Booked" || st === "Arrived" || st === "In-Progress") out.apptPending++;
    }
  }

  var inv = ss.getSheetByName("Pharmacy_Inventory");
  if (inv && inv.getLastRow() > 1) {
    var d = inv.getDataRange().getValues();
    var nowYM = Utilities.formatDate(new Date(), tz, "yyyy-MM");
    for (var j = 1; j < d.length; j++) {
      var qty = parseInt(d[j][4], 10) || 0;
      if (qty <= 0) continue;
      if (qty <= 10) out.lowStock++;
      var exp = d[j][7];
      var expYM = (exp instanceof Date)
        ? Utilities.formatDate(exp, tz, "yyyy-MM")
        : String(exp || "").substring(0, 7);
      if (expYM && expYM <= nowYM) out.expiringSoon++;
    }
  }

  var lab = ss.getSheetByName("LAB_ORDERS");
  if (lab && lab.getLastRow() > 1) {
    var ld = lab.getDataRange().getValues();
    var sIdx = _dashHeaderIndex_(ld[0], ["status", "order_status", "order status"]);
    if (sIdx > -1) {
      for (var m = 1; m < ld.length; m++) {
        var ls = (ld[m][sIdx] || "").toString().trim().toUpperCase();
        if (ls && ls !== "VERIFIED" && ls !== "COMPLETED" && ls !== "REPORTED" && ls !== "CANCELLED") {
          out.labPending++;
        }
      }
    }
  }
  return out;
}

// ------------------------------------------------------------
// REVENUE (financial roles only) — today, three streams, no double-count
// Consult  = Appointments.Fee where today & Completed
// Pharmacy = Pharmacy_Invoices.Net[13] where Timestamp[1]=today & Pay_Status[16]=PAID
// Lab      = existing getLabDailyCollection().totalNet (today OP collection)
// ------------------------------------------------------------
function _dashRevenue_(ss, tz, todayKey) {
  var out = { consultToday: 0, pharmacyToday: 0, labToday: 0,
              totalToday: 0, pendingCredit: 0, hasData: true };

  var a = ss.getSheetByName("Appointments");
  if (a && a.getLastRow() > 1) {
    var ad = a.getDataRange().getValues();
    for (var i = 1; i < ad.length; i++) {
      var dt = _dashToDate_(ad[i][3]); if (!dt) continue;
      if (Utilities.formatDate(dt, tz, "yyyy-MM-dd") !== todayKey) continue;
      if ((ad[i][6] || "").toString() === "Completed") {
        out.consultToday += (parseFloat(ad[i][7]) || 0);
      }
    }
  }

  var pinv = ss.getSheetByName("Pharmacy_Invoices");
  if (pinv && pinv.getLastRow() > 1) {
    var pd = pinv.getDataRange().getValues();
    for (var j = 1; j < pd.length; j++) {
      var pdt = _dashToDate_(pd[j][1]); if (!pdt) continue;
      if (Utilities.formatDate(pdt, tz, "yyyy-MM-dd") !== todayKey) continue;
      var pay = (pd[j][16] || "").toString().trim().toUpperCase();
      var net = parseFloat(pd[j][13]) || 0;
      if (pay === "PAID") out.pharmacyToday += net;
      else if (pay === "PENDING") out.pendingCredit += net;
    }
  }

  try {
    if (typeof getLabDailyCollection === 'function') {
      var lc = getLabDailyCollection();
      if (lc && lc.success) out.labToday = parseFloat(lc.totalNet) || 0;
    }
  } catch (e) { /* lab module optional — degrade silently */ }

  out.consultToday  = Math.round(out.consultToday * 100) / 100;
  out.pharmacyToday = Math.round(out.pharmacyToday * 100) / 100;
  out.labToday      = Math.round(out.labToday * 100) / 100;
  out.pendingCredit = Math.round(out.pendingCredit * 100) / 100;
  out.totalToday    = Math.round((out.consultToday + out.pharmacyToday + out.labToday) * 100) / 100;
  return out;
}

// ------------------------------------------------------------
// STAFF COUNT (Users sheet)
// ------------------------------------------------------------
function _dashStaff_(ss) {
  var u = ss.getSheetByName("Users");
  if (!u || u.getLastRow() <= 1) return 0;
  var d = u.getDataRange().getValues(), n = 0;
  for (var i = 1; i < d.length; i++) { if (d[i][0]) n++; }
  return n;
}

// ============================================================
// HELPERS
// ============================================================
function _zeros12_() { return [0,0,0,0,0,0,0,0,0,0,0,0]; }

/** Robust date parse for Date objects, ISO, dd/MM/yyyy, MM/dd/yyyy timestamps. */
function _dashToDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v).trim(); if (!s) return null;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  var m = s.match(/(\d{4})-(\d{2})-(\d{2})/); // yyyy-MM-dd prefix fallback
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

/** Case-insensitive header lookup; returns column index or -1. */
function _dashHeaderIndex_(headerRow, candidates) {
  for (var c = 0; c < headerRow.length; c++) {
    var h = (headerRow[c] || "").toString().trim().toLowerCase();
    for (var k = 0; k < candidates.length; k++) {
      if (h === candidates[k]) return c;
    }
  }
  return -1;
}