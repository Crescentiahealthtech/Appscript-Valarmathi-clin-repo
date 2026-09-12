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
    // The version suffix moves whenever the payload shape changes, so a cached
    // copy from the previous deployment cannot reach a screen that now expects
    // the analytics block and find it missing.
    var cacheKey = 'DASH2_' + (isFinancial ? 'FIN' : 'CLIN');

    if (!bust) {
      var hit = cache.get(cacheKey);
      if (hit) return { success: true, data: JSON.parse(hit), message: "cache" };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone() || "Asia/Kolkata";
    var now = new Date();
    var todayKey = Utilities.formatDate(now, tz, "yyyy-MM-dd");

    // One read per sheet for the whole payload. _dashCensus_ and _dashFootfall_
    // both wanted IP_Admissions and both fetched it; the analytics block below
    // wants it a third time. The memo makes that one call.
    var memo = {};

    var data = {
      meta: {
        generatedAt: Utilities.formatDate(now, tz, "dd MMM yyyy • hh:mm a"),
        role: role,
        financial: isFinancial
      },
      census:    _dashCensus_(ss, memo),
      footfall:  _dashFootfall_(ss, tz, memo),
      operations:_dashOps_(ss, tz, todayKey, memo),
      staffCount:_dashStaff_(ss, memo),
      analytics: _dashAnalytics_(ss, tz, memo, isFinancial)
    };

    if (isFinancial) {
      data.revenue = _dashRevenue_(ss, tz, todayKey, memo);
    }

    cache.put(cacheKey, JSON.stringify(data), 300); // 5 minutes
    return { success: true, data: data, message: "fresh" };

  } catch (e) {
    return { success: false, message: "Dashboard load failed: " + e.message, data: null };
  }
}

/** Call this from any write-path module after data changes if you want instant freshness. */
function invalidateDashboardCache() {
  try { CacheService.getScriptCache().removeAll(['DASH2_FIN', 'DASH2_CLIN', 'DASH_FIN', 'DASH_CLIN']); } catch (e) {}
}

/**
 * One sheet, read once per dashboard build.
 *
 * Every block below used to call getDataRange() for itself, so a dashboard
 * touched IP_Admissions three times and Appointments four. The memo is per
 * call — it lives and dies inside getDashboardData().
 */
function _dashSheet_(ss, name, memo) {
  if (memo && memo[name] !== undefined) return memo[name];
  var sh = ss.getSheetByName(name);
  var values = (sh && sh.getLastRow() > 1) ? sh.getDataRange().getValues() : [];
  if (memo) memo[name] = values;
  return values;
}

// ------------------------------------------------------------
// CENSUS & BEDS  (headline KPI band)
// Master_Beds: Bed_ID[0] Ward[1] Status[2] Patient_ID[3] Patient_Name[4] DOA[5] IP_No[6]
// IP_Admissions: IP_No[0] PatientID[1] Name[2] AgeSex[3] DOA[4] ... Status[11] DOD[12]
// ------------------------------------------------------------
function _dashCensus_(ss, memo) {
  var out = { 
    totalBeds: 0, 
    occupied: 0, 
    reserved: 0, 
    cleaning: 0, 
    available: 0, 
    occupancyPct: 0,
    wards: [], 
    currentInpatients: 0, 
    hasData: false 
  };

  var d = _dashSheet_(ss, "Master_Beds", memo);
  if (d.length > 1) {
    var wardMap = {};
    
    for (var i = 1; i < d.length; i++) {
      if (!d[i][0]) continue;
      out.totalBeds++;
      
      var ward = (d[i][1] || "GEN").toString().trim();
      var status = (d[i][2] || "Available").toString().trim().toUpperCase();
      
      // Categorize the bed statuses exactly matching your UI
      if (status === "OCCUPIED" || status === "ISOLATION") {
        out.occupied++;
      } else if (status === "RESERVED") {
        out.reserved++;
      } else if (status === "CLEANING" || status === "MAINTENANCE") {
        out.cleaning++;
      } else {
        out.available++;
      }
      
      // Any bed that is not strictly "AVAILABLE" is considered utilized
      var isUtilized = (status !== "AVAILABLE"); 
      
      if (!wardMap[ward]) wardMap[ward] = { ward: ward, total: 0, utilized: 0 };
      wardMap[ward].total++;
      if (isUtilized) wardMap[ward].utilized++;
    }
    
    out.wards = Object.keys(wardMap).map(function (k) { return wardMap[k]; })
                      .sort(function (a, b) { return a.ward < b.ward ? -1 : 1; });
    
    var totalUnavailable = out.occupied + out.reserved + out.cleaning;
    out.occupancyPct = out.totalBeds ? Math.round((totalUnavailable / out.totalBeds) * 100) : 0;
    out.hasData = out.totalBeds > 0;
  }

  var ip = _dashSheet_(ss, "IP_Admissions", memo);
  if (ip.length > 1) {
    for (var j = 1; j < ip.length; j++) {
      // FIX: Check for both "ACTIVE" and "ADMITTED" to match the database
      var stat = (ip[j][11] || "").toString().trim().toUpperCase();
      if (stat === "ACTIVE" || stat === "ADMITTED") {
        out.currentInpatients++;
      }
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
function _dashFootfall_(ss, tz, memo) {
  var op = {}, ip = {}, years = {};

  var ad = _dashSheet_(ss, "Appointments", memo);
  if (ad.length > 1) {
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

  var id = _dashSheet_(ss, "IP_Admissions", memo);
  if (id.length > 1) {
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
function _dashOps_(ss, tz, todayKey, memo) {
  var out = { apptToday: 0, apptPending: 0, apptCompleted: 0,
              lowStock: 0, expiringSoon: 0, labPending: 0 };

  var ad = _dashSheet_(ss, "Appointments", memo);
  if (ad.length > 1) {
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

  var d = _dashSheet_(ss, "Pharmacy_Inventory", memo);
  if (d.length > 1) {
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

  var ld = _dashSheet_(ss, "LAB_ORDERS", memo);
  if (ld.length > 1) {
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
function _dashRevenue_(ss, tz, todayKey, memo) {
  var out = { consultToday: 0, hospitalToday: 0, pharmacyToday: 0, labToday: 0,
              totalToday: 0, pendingCredit: 0, hasData: true };

  // An appointment's Fee is the consultation TARIFF now, not money taken: the
  // receipt is a Hospital_Invoices row. Only a visit with no invoice still
  // counts its fee, which keeps historic days reading as they always did
  // without double-counting a modern one. Same rule as acc_opRows_().
  var billedAppts = {};
  try { if (typeof hb_billedApptIds_ === 'function') billedAppts = hb_billedApptIds_() || {}; }
  catch (e) { billedAppts = {}; }

  var ad = _dashSheet_(ss, "Appointments", memo);
  if (ad.length > 1) {
    for (var i = 1; i < ad.length; i++) {
      var dt = _dashToDate_(ad[i][3]); if (!dt) continue;
      if (Utilities.formatDate(dt, tz, "yyyy-MM-dd") !== todayKey) continue;
      if ((ad[i][6] || "").toString() !== "Completed") continue;
      if (billedAppts[String(ad[i][0] || "").toUpperCase()]) continue;
      out.consultToday += (parseFloat(ad[i][7]) || 0);
    }
  }

  // Hospital billing: consultations, procedures and packages. Only what was
  // actually received counts as today's collection; every unpaid balance,
  // whenever it was raised, is credit — the same definition the receivables
  // tab uses, so the two screens cannot disagree.
  var hinv = _dashSheet_(ss, "Hospital_Invoices", memo);
  if (hinv.length > 1) {
    var hIdx = _dashIndexMap_(hinv[0]);
    for (var h = 1; h < hinv.length; h++) {
      if (!hinv[h][0]) continue;
      if (String(_dashCell_(hinv[h], hIdx, "status") || "").toUpperCase() === "CANCELLED") continue;
      var paid = parseFloat(_dashCell_(hinv[h], hIdx, "paid")) || 0;
      var bal = parseFloat(_dashCell_(hinv[h], hIdx, "balance")) || 0;
      out.pendingCredit += bal;
      var hdt = _dashToDate_(_dashCell_(hinv[h], hIdx, "timestamp"));
      if (!hdt || Utilities.formatDate(hdt, tz, "yyyy-MM-dd") !== todayKey) continue;
      out.hospitalToday += paid;
    }
  }

  var pd = _dashSheet_(ss, "Pharmacy_Invoices", memo);
  if (pd.length > 1) {
    for (var j = 1; j < pd.length; j++) {
      var pdt = _dashToDate_(pd[j][1]); if (!pdt) continue;
      var pay = (pd[j][16] || "").toString().trim().toUpperCase();
      var net = parseFloat(pd[j][13]) || 0;
      // Credit is a balance, not an event: it is counted whenever it was
      // raised. Only the collection is scoped to today.
      if (pay === "PENDING" || pay === "CREDIT") out.pendingCredit += net;
      if (Utilities.formatDate(pdt, tz, "yyyy-MM-dd") !== todayKey) continue;
      if (pay === "PAID") out.pharmacyToday += net;
    }
  }

  // Lab dues were missing from Pending Credit altogether, so the tile could
  // read zero while the receivables tab listed thousands. Both now count the
  // same thing: every open balance, whenever it was raised.
  var lb = _dashSheet_(ss, "LAB_BILLING", memo);
  if (lb.length > 1) {
    var lIdx = _dashIndexMap_(lb[0]);
    for (var l = 1; l < lb.length; l++) {
      if (!lb[l][0]) continue;
      var lstat = String(_dashCell_(lb[l], lIdx, "paymentstatus") || "").toUpperCase();
      if (lstat === "PAID") continue;
      var lbal = parseFloat(_dashCell_(lb[l], lIdx, "balanceamount")) || 0;
      var lnet = parseFloat(_dashCell_(lb[l], lIdx, "netamount")) || 0;
      out.pendingCredit += (lbal || (lstat === "ON_ACCOUNT" ? lnet : 0));
    }
  }

  try {
    if (typeof getLabDailyCollection === 'function') {
      var lc = getLabDailyCollection();
      if (lc && lc.success) out.labToday = parseFloat(lc.totalNet) || 0;
    }
  } catch (e) { /* lab module optional — degrade silently */ }

  out.consultToday  = Math.round(out.consultToday * 100) / 100;
  out.hospitalToday = Math.round(out.hospitalToday * 100) / 100;
  out.pharmacyToday = Math.round(out.pharmacyToday * 100) / 100;
  out.labToday      = Math.round(out.labToday * 100) / 100;
  out.pendingCredit = Math.round(out.pendingCredit * 100) / 100;
  out.totalToday    = Math.round((out.consultToday + out.hospitalToday +
                                  out.pharmacyToday + out.labToday) * 100) / 100;
  return out;
}

// ------------------------------------------------------------
// STAFF COUNT (Users sheet)
// ------------------------------------------------------------
function _dashStaff_(ss, memo) {
  var d = _dashSheet_(ss, "Users", memo), n = 0;
  for (var i = 1; i < d.length; i++) { if (d[i][0]) n++; }
  return n;
}

// ============================================================
// 📈 ANALYTICS
// ------------------------------------------------------------
// What replaced the "Operational Hubs" grid. That grid was eight cards
// that each did what a sidebar link already did — navigation dressed as a
// dashboard, occupying the best real estate on the busiest screen.
//
// Everything below is computed from sheets the hospital already fills in.
// Nothing here needs a new column. Each block is guarded on its own, so a
// module that is not installed costs its own panel and not the dashboard.
// ============================================================
function _dashAnalytics_(ss, tz, memo, isFinancial) {
  var out = {
    window: { days: 30 },
    stay: { alosDays: 0, dischargesInWindow: 0, byWard: [] },
    appointments: { completed: 0, cancelled: 0, noShow: 0, upcoming: 0, completionPct: 0 },
    doctorLoad: [],
    diagnoses: [],
    registrations: { months: [], counts: [] },
    discharge: { pendingSignature: 0, returned: 0, inPreparation: 0,
                 signedToday: 0, medianTatHours: 0, available: false },
    revenueMix: [],
    collections: { days: [], collected: [], credit: [] },
    paymentMix: [],
    hasFinancial: !!isFinancial
  };

  var now = new Date();
  var dayMs = 86400000;
  var win30 = now.getTime() - 30 * dayMs;
  var win90 = now.getTime() - 90 * dayMs;
  var win180 = now.getTime() - 180 * dayMs;

  // ---- length of stay, by ward (90 days of completed stays) ---------------
  try {
    var ip = _dashSheet_(ss, "IP_Admissions", memo);
    var wards = {}, totalDays = 0, n = 0;
    for (var i = 1; i < ip.length; i++) {
      if (!ip[i][0]) continue;
      var doa = _dashToDate_(ip[i][4]);
      var dod = _dashToDate_(ip[i][12]);
      if (!doa || !dod || dod.getTime() < win90) continue;
      // A same-day discharge is one billed day, not zero.
      var days = Math.max(1, Math.round((dod.getTime() - doa.getTime()) / dayMs));
      var ward = String(ip[i][7] || ip[i][8] || "General").split("-")[0].trim() || "General";
      if (!wards[ward]) wards[ward] = { ward: ward, days: 0, count: 0 };
      wards[ward].days += days;
      wards[ward].count++;
      totalDays += days; n++;
    }
    out.stay.dischargesInWindow = n;
    out.stay.alosDays = n ? Math.round((totalDays / n) * 10) / 10 : 0;
    out.stay.byWard = Object.keys(wards).map(function (k) {
      return { ward: k, alos: Math.round((wards[k].days / wards[k].count) * 10) / 10,
               discharges: wards[k].count };
    }).sort(function (a, b) { return b.discharges - a.discharges; }).slice(0, 6);
  } catch (e) { /* the ward module is optional to the dashboard */ }

  // ---- appointment outcomes and doctor load (30 days) ---------------------
  try {
    var ad = _dashSheet_(ss, "Appointments", memo);
    var docIdx = _dashHeaderIndex_(ad.length ? ad[0] : [], ["doctor_name_snapshot"]);
    var load = {};
    for (var a = 1; a < ad.length; a++) {
      if (!ad[a][0]) continue;
      var when = _dashToDate_(ad[a][3]);
      if (!when || when.getTime() < win30) continue;
      var st = String(ad[a][6] || "").trim();
      if (st === "Blocked" || st === "DELETE") continue;

      if (st === "Completed") out.appointments.completed++;
      else if (st === "Cancelled") out.appointments.cancelled++;
      else if (when.getTime() < now.getTime() - dayMs) out.appointments.noShow++;
      else out.appointments.upcoming++;

      if (st !== "Cancelled") {
        var who = (docIdx > -1 ? String(ad[a][docIdx] || "") : "") || "Unassigned";
        load[who] = (load[who] || 0) + 1;
      }
    }
    var seen = out.appointments.completed + out.appointments.cancelled + out.appointments.noShow;
    out.appointments.completionPct = seen ? Math.round(out.appointments.completed * 100 / seen) : 0;
    out.doctorLoad = Object.keys(load).map(function (k) { return { name: k, count: load[k] }; })
      .sort(function (x, y) { return y.count - x.count; }).slice(0, 6);
  } catch (e) {}

  // ---- top diagnoses (180 days of admissions) ----------------------------
  try {
    var ipd = _dashSheet_(ss, "IP_Admissions", memo);
    var dx = {};
    for (var k = 1; k < ipd.length; k++) {
      if (!ipd[k][0]) continue;
      var kd = _dashToDate_(ipd[k][4]);
      if (!kd || kd.getTime() < win180) continue;
      // One admission can carry several comma-separated diagnoses; each is
      // counted, because "HYPOTHYROIDISM, Alopecia" is two findings.
      String(ipd[k][10] || "").split(/[,;]/).forEach(function (part) {
        var name = part.trim();
        if (!name || name.toLowerCase() === "pending") return;
        var key = name.toUpperCase();
        if (!dx[key]) dx[key] = { name: name, count: 0 };
        dx[key].count++;
      });
    }
    out.diagnoses = Object.keys(dx).map(function (key) { return dx[key]; })
      .sort(function (x, y) { return y.count - x.count; }).slice(0, 6);
  } catch (e) {}

  // ---- new registrations, last 6 months ----------------------------------
  try {
    var pt = _dashSheet_(ss, "Patients", memo);
    var regIdx = _dashHeaderIndex_(pt.length ? pt[0] : [], ["registration_date", "registration date"]);
    if (regIdx === -1) regIdx = 10;
    var buckets = {}, labels = [];
    for (var m = 5; m >= 0; m--) {
      var d0 = new Date(now.getFullYear(), now.getMonth() - m, 1);
      var key0 = Utilities.formatDate(d0, tz, "yyyy-MM");
      buckets[key0] = 0;
      labels.push({ key: key0, label: Utilities.formatDate(d0, tz, "MMM") });
    }
    for (var r = 1; r < pt.length; r++) {
      if (!pt[r][0]) continue;
      var rd = _dashToDate_(pt[r][regIdx]);
      if (!rd) continue;
      var rk = Utilities.formatDate(rd, tz, "yyyy-MM");
      if (buckets[rk] !== undefined) buckets[rk]++;
    }
    out.registrations.months = labels.map(function (l) { return l.label; });
    out.registrations.counts = labels.map(function (l) { return buckets[l.key]; });
  } catch (e) {}

  // ---- discharge-summary turnaround --------------------------------------
  try {
    var ds = _dashSheet_(ss, "DS_Summaries", memo);
    if (ds.length > 1) {
      out.discharge.available = true;
      var dIdx = _dashIndexMap_(ds[0]);
      var todayK = Utilities.formatDate(now, tz, "yyyy-MM-dd");
      var tats = [];
      for (var q = 1; q < ds.length; q++) {
        if (!ds[q][0]) continue;
        var status = String(_dashCell_(ds[q], dIdx, "status") || "").toUpperCase();
        if (status === "PENDING_SIGNATURE") out.discharge.pendingSignature++;
        else if (status === "RETURNED") out.discharge.returned++;
        else if (status === "IN_PREPARATION") out.discharge.inPreparation++;

        var sAt = _dashToDate_(_dashCell_(ds[q], dIdx, "signed_at"));
        var iAt = _dashToDate_(_dashCell_(ds[q], dIdx, "initiated_at"));
        if (!sAt || !iAt) continue;
        if (Utilities.formatDate(sAt, tz, "yyyy-MM-dd") === todayK) out.discharge.signedToday++;
        if (sAt.getTime() >= now.getTime() - 7 * dayMs) {
          tats.push((sAt.getTime() - iAt.getTime()) / 3600000);
        }
      }
      out.discharge.medianTatHours = _dashMedian_(tats);
    }
  } catch (e) {}

  if (!isFinancial) return out;

  // ---- revenue mix and daily collections (financial roles only) ----------
  try {
    var mix = { Consultation: 0, Pharmacy: 0, Laboratory: 0, Inpatient: 0 };
    var byDay = {}, creditByDay = {}, modes = {};
    var dayKeys = [];
    for (var b = 13; b >= 0; b--) {
      var db = new Date(now.getTime() - b * dayMs);
      dayKeys.push(Utilities.formatDate(db, tz, "yyyy-MM-dd"));
    }
    dayKeys.forEach(function (k) { byDay[k] = 0; creditByDay[k] = 0; });

    // `credit` means billed but not received. The mix and the payment-mode
    // split are both about money actually taken, so credit only ever lands in
    // the daily credit series — counting it as revenue was how the old
    // dashboard made a slow month look like a good one.
    var add = function (dt, amount, bucket, mode, credit) {
      if (!dt || !amount) return;
      var dk = Utilities.formatDate(dt, tz, "yyyy-MM-dd");
      if (!credit && dt.getTime() >= win30) {
        mix[bucket] += amount;
        if (mode) modes[mode] = (modes[mode] || 0) + amount;
      }
      if (byDay[dk] !== undefined) {
        if (credit) creditByDay[dk] += amount; else byDay[dk] += amount;
      }
    };

    // Hospital invoices — consultations, procedures, packages
    var hi = _dashSheet_(ss, "Hospital_Invoices", memo);
    if (hi.length > 1) {
      var hm = _dashIndexMap_(hi[0]);
      for (var x = 1; x < hi.length; x++) {
        if (!hi[x][0]) continue;
        if (String(_dashCell_(hi[x], hm, "status") || "").toUpperCase() === "CANCELLED") continue;
        var hdt = _dashToDate_(_dashCell_(hi[x], hm, "timestamp"));
        var hpaid = parseFloat(_dashCell_(hi[x], hm, "paid")) || 0;
        var hbal = parseFloat(_dashCell_(hi[x], hm, "balance")) || 0;
        var hmode = String(_dashCell_(hi[x], hm, "payment_mode") || "Cash");
        add(hdt, hpaid, "Consultation", hmode, false);
        add(hdt, hbal, "Consultation", null, true);
      }
    }

    // Pharmacy
    var pi = _dashSheet_(ss, "Pharmacy_Invoices", memo);
    for (var y = 1; y < pi.length; y++) {
      if (!pi[y][0]) continue;
      var pdt = _dashToDate_(pi[y][1]);
      var pnet = parseFloat(pi[y][13]) || 0;
      var pstat = String(pi[y][16] || "").trim().toUpperCase();
      if (pstat === "PAID") add(pdt, pnet, "Pharmacy", "Cash", false);
      else if (pstat === "PENDING") add(pdt, pnet, "Pharmacy", null, true);
    }

    // Lab
    var lb = _dashSheet_(ss, "LAB_BILLING", memo);
    if (lb.length > 1) {
      var lm = _dashIndexMap_(lb[0]);
      for (var z = 1; z < lb.length; z++) {
        if (!lb[z][0]) continue;
        var ldt = _dashToDate_(_dashCellAny_(lb[z], lm, ["billedat", "timestamp", "date"]));
        var lnet = parseFloat(_dashCell_(lb[z], lm, "netamount")) || 0;
        var lbal = parseFloat(_dashCell_(lb[z], lm, "balanceamount")) || 0;
        var lstat = String(_dashCell_(lb[z], lm, "paymentstatus") || "").toUpperCase();
        if (lstat === "PAID") add(ldt, lnet, "Laboratory", String(_dashCell_(lb[z], lm, "paymentmode") || "Cash"), false);
        else add(ldt, lbal || lnet, "Laboratory", null, true);
      }
    }

    // IP settlements
    var st = _dashSheet_(ss, "IP_Settlements", memo);
    if (st.length > 1) {
      var sm = _dashIndexMap_(st[0]);
      for (var w = 1; w < st.length; w++) {
        if (!st[w][0]) continue;
        var sdt = _dashToDate_(_dashCellAny_(st[w], sm, ["timestamp", "date"]) || st[w][1]);
        var spaid = parseFloat(_dashCellAny_(st[w], sm, ["patient_paid", "patientpaid", "patient paid"])) || 0;
        add(sdt, spaid, "Inpatient", String(_dashCellAny_(st[w], sm, ["payment_mode", "paymentmode"]) || "Cash"), false);
      }
    }

    out.revenueMix = Object.keys(mix)
      .map(function (k) { return { label: k, amount: Math.round(mix[k] * 100) / 100 }; })
      .filter(function (r) { return r.amount > 0; });

    out.collections.days = dayKeys.map(function (k) {
      return Utilities.formatDate(_dashToDate_(k), tz, "dd MMM");
    });
    out.collections.collected = dayKeys.map(function (k) { return Math.round(byDay[k]); });
    out.collections.credit = dayKeys.map(function (k) { return Math.round(creditByDay[k]); });

    out.paymentMix = Object.keys(modes)
      .map(function (k) { return { label: k, amount: Math.round(modes[k] * 100) / 100 }; })
      .filter(function (r) { return r.amount > 0; })
      .sort(function (a, b) { return b.amount - a.amount; }).slice(0, 6);
  } catch (e) { /* one bad money sheet must not cost the clinical panels */ }

  return out;
}

function _dashMedian_(arr) {
  if (!arr.length) return 0;
  var a = arr.slice().sort(function (x, y) { return x - y; });
  var mid = Math.floor(a.length / 2);
  var v = a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
  return Math.round(v * 10) / 10;
}

/**
 * {lowercased_header: index} for a header row.
 *
 * The money sheets in this project each spell their columns differently —
 * Timestamp vs BilledAt, Net vs NetAmount — so the analytics block looks them
 * up by name and tolerates a miss rather than counting the wrong column.
 */
function _dashIndexMap_(headerRow) {
  var map = {};
  (headerRow || []).forEach(function (h, i) {
    var key = String(h || "").trim().toLowerCase();
    if (key && map[key] === undefined) map[key] = i;
  });
  return map;
}

/** One cell by header name, or '' when the sheet has no such column. */
function _dashCell_(row, map, name) {
  var i = map[name];
  return (i === undefined || !row) ? '' : row[i];
}

/** The first of several spellings a column might use. */
function _dashCellAny_(row, map, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = map[names[i]];
    if (idx !== undefined && row) return row[idx];
  }
  return '';
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