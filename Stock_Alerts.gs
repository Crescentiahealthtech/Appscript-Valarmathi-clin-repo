// ============================================================================
// Stock_Alerts.gs — Crescentia HealthTech / CresRx
// What is running out and what is about to expire. How much to order is
// the pharmacist's decision; this only shows the facts behind it.
// ----------------------------------------------------------------------------
// Pharmacy_Inventory holds one row per batch with its live quantity (the sale
// path deducts in place, Pharmacy.gs). Nothing read it back to ask whether a
// medicine was running low or a batch was about to expire, so both were found
// at the counter: a patient turned away, or a strip sold two weeks before its
// date.
//
// LOW STOCK is judged by how long the stock will last, not by a fixed count:
// ten strips of a medicine sold once a month is plenty, ten of one sold
// twenty a day is half a day. Average daily sale comes from the last
// STK_CFG.USAGE_DAYS of Pharmacy_Invoice_Items. A medicine with no sales
// history falls back to a plain minimum (STK_CFG.MIN_UNITS, or a per-item
// level on the optional Pharmacy_Reorder_Levels sheet).
//
// NO ORDER QUANTITY IS SUGGESTED. Each low item shows what is left, how fast
// it sells, and the supplier and buying price of its latest batch — what the
// pharmacist needs to decide the order themselves.
// ============================================================================

var STK_CFG = {
  USAGE_DAYS: 60,        // sales window for the daily average
  LEAD_DAYS: 7,          // how long an order takes to arrive: low = under LEAD_DAYS + 7 days left
  MIN_UNITS: 10,         // floor for a medicine with no sales history
  NEAR_EXPIRY_DAYS: 90,  // flagged
  URGENT_EXPIRY_DAYS: 30,// flagged red: return to supplier or sell first
  LEVELS_SHEET: 'Pharmacy_Reorder_Levels'   // optional: Brand, Generic, Min_Qty
};

function stk_str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function stk_key_(brand, generic) { return stk_str_(brand).toUpperCase() + '|' + stk_str_(generic).toUpperCase(); }

/**
 * An expiry cell as the LAST day it may be sold. "2028-06" and "06/2028"
 * mean the end of June 2028, which is how expiry is printed on a strip.
 */
function stk_expiry_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = stk_str_(v);
  if (!s) return null;
  var m = /^(\d{4})[-\/](\d{1,2})$/.exec(s) || null;
  if (m) return new Date(+m[1], +m[2], 0);
  m = /^(\d{1,2})[-\/](\d{4})$/.exec(s);
  if (m) return new Date(+m[2], +m[1], 0);
  var d = (typeof cresc_parseDate_ === 'function') ? cresc_parseDate_(s) : new Date(s);
  return (d && !isNaN(d.getTime())) ? d : null;
}

/** Header index by any of several names, else the fallback position. */
function stk_col_(hdr, names, fallback) {
  for (var i = 0; i < names.length; i++) {
    var k = hdr.indexOf(names[i]);
    if (k !== -1) return k;
  }
  return fallback;
}

/** The analysis behind the alerts. No permission check: callers have one. */
function stk_analyse_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inv = ss.getSheetByName(typeof PH_SHEETS !== 'undefined' ? PH_SHEETS.INVENTORY : 'Pharmacy_Inventory');
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var DAY = 86400000;

  var items = {}, expiring = [], expired = [];
  if (inv && inv.getLastRow() > 1) {
    var data = inv.getDataRange().getValues();
    var h = data[0].map(stk_str_);
    var c = {
      ts: stk_col_(h, ['Timestamp'], 0), brand: stk_col_(h, ['Brand Name'], 1),
      generic: stk_col_(h, ['Generic Name'], 2), type: stk_col_(h, ['Type'], 3),
      qty: stk_col_(h, ['Qty'], 4), unit: stk_col_(h, ['Unit'], 5),
      batch: stk_col_(h, ['Batch No'], 6), exp: stk_col_(h, ['Expiry Date'], 7),
      rack: stk_col_(h, ['Rack Location'], 8), buy: stk_col_(h, ['Buy Price'], 9),
      mrp: stk_col_(h, ['MRP'], 10), supplier: stk_col_(h, ['Supplier'], 13)
    };
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var brand = stk_str_(r[c.brand]);
      if (!brand) continue;
      var key = stk_key_(brand, r[c.generic]);
      var qty = parseFloat(r[c.qty]) || 0;
      var it = items[key];
      if (!it) {
        it = items[key] = { brand: brand, generic: stk_str_(r[c.generic]), type: stk_str_(r[c.type]),
                            unit: stk_str_(r[c.unit]), stock: 0, sellable: 0, batches: 0,
                            supplier: '', buyPrice: 0, lastIn: 0 };
      }
      // The most recent batch is the supplier and price shown.
      var inAt = (r[c.ts] instanceof Date) ? r[c.ts].getTime() : i;
      if (inAt >= it.lastIn) {
        it.lastIn = inAt;
        if (stk_str_(r[c.supplier])) it.supplier = stk_str_(r[c.supplier]);
        if (parseFloat(r[c.buy]) > 0) it.buyPrice = parseFloat(r[c.buy]);
      }
      if (qty <= 0) continue;
      it.stock += qty;
      it.batches++;

      var exp = stk_expiry_(r[c.exp]);
      var days = exp ? Math.floor((exp.getTime() - today.getTime()) / DAY) : null;
      var row = { brand: brand, generic: it.generic, batch: stk_str_(r[c.batch]),
                  expiry: (typeof cresc_expiryText_ === 'function') ? cresc_expiryText_(r[c.exp]) : stk_str_(r[c.exp]),
                  daysLeft: days, qty: qty, unit: it.unit, rack: stk_str_(r[c.rack]),
                  supplier: stk_str_(r[c.supplier]),
                  value: Math.round(qty * (parseFloat(r[c.buy]) || 0) * 100) / 100,
                  mrpValue: Math.round(qty * (parseFloat(r[c.mrp]) || 0) * 100) / 100 };
      if (days !== null && days < 0) { expired.push(row); continue; }
      it.sellable += qty;
      if (days !== null && days <= STK_CFG.NEAR_EXPIRY_DAYS) {
        row.urgent = days <= STK_CFG.URGENT_EXPIRY_DAYS;
        expiring.push(row);
      }
    }
  }

  // Sales over the window -> units a day.
  var sold = {};
  var since = today.getTime() - STK_CFG.USAGE_DAYS * DAY;
  var ii = ss.getSheetByName(typeof PH_SHEETS !== 'undefined' ? PH_SHEETS.INVOICE_ITEMS : 'Pharmacy_Invoice_Items');
  if (ii && ii.getLastRow() > 1) {
    var sd = ii.getDataRange().getValues();
    var sh = sd[0].map(stk_str_);
    var sc = { ts: stk_col_(sh, ['Timestamp'], 1), brand: stk_col_(sh, ['Brand'], 3),
               generic: stk_col_(sh, ['Generic'], 4), qty: stk_col_(sh, ['Qty'], 7) };
    for (var j = 1; j < sd.length; j++) {
      var t = sd[j][sc.ts];
      var ms = (t instanceof Date) ? t.getTime()
             : ((typeof cresc_parseDate_ === 'function' && cresc_parseDate_(t)) ? cresc_parseDate_(t).getTime() : NaN);
      if (!(ms >= since)) continue;
      var k2 = stk_key_(sd[j][sc.brand], sd[j][sc.generic]);
      sold[k2] = (sold[k2] || 0) + (parseFloat(sd[j][sc.qty]) || 0);
    }
  }

  // Optional per-item levels.
  var levels = {};
  var lv = ss.getSheetByName(STK_CFG.LEVELS_SHEET);
  if (lv && lv.getLastRow() > 1) {
    lv.getDataRange().getValues().slice(1).forEach(function (r) {
      if (!stk_str_(r[0])) return;
      levels[stk_key_(r[0], r[1])] = { min: parseFloat(r[2]) || 0 };
    });
  }
  var minUnits = STK_CFG.MIN_UNITS;
  try {
    var p = PropertiesService.getScriptProperties().getProperty('STOCK_MIN_UNITS');
    if (p && !isNaN(parseFloat(p))) minUnits = parseFloat(p);
  } catch (e) {}

  var low = [];
  Object.keys(items).forEach(function (k) {
    var it = items[k];
    var perDay = (sold[k] || 0) / STK_CFG.USAGE_DAYS;
    var level = levels[k];
    var cover = perDay > 0 ? it.sellable / perDay : null;
    var reason = '';
    if (level && level.min > 0 && it.sellable <= level.min) {
      reason = 'at or below its reorder level (' + level.min + ')';
    } else if (perDay > 0 && cover < STK_CFG.LEAD_DAYS + 7) {
      reason = it.sellable <= 0 ? 'out of stock, sells ' + perDay.toFixed(1) + '/day'
             : 'about ' + Math.floor(cover) + ' day(s) left at ' + perDay.toFixed(1) + '/day';
    } else if (perDay === 0 && !level && it.sellable > 0 && it.sellable <= minUnits) {
      // No sales in the window: the plain floor. A medicine at zero with no
      // sales is not listed — that is a discontinued line, not a shortage.
      reason = 'only ' + it.sellable + ' left (no sales in ' + STK_CFG.USAGE_DAYS + ' days)';
    }
    if (!reason) return;
    low.push({ brand: it.brand, generic: it.generic, type: it.type, unit: it.unit,
               stock: it.sellable, perDay: Math.round(perDay * 10) / 10,
               coverDays: cover === null ? null : Math.floor(cover),
               reason: reason, supplier: it.supplier || 'Supplier not recorded',
               buyPrice: it.buyPrice });
  });
  low.sort(function (a, b) {
    var ca = a.coverDays === null ? 1e9 : a.coverDays, cb = b.coverDays === null ? 1e9 : b.coverDays;
    return (a.stock > 0) - (b.stock > 0) || ca - cb;
  });
  expiring.sort(function (a, b) { return a.daysLeft - b.daysLeft; });
  expired.sort(function (a, b) { return a.daysLeft - b.daysLeft; });

  var sum = function (list, k) { return Math.round(list.reduce(function (t, x) { return t + (x[k] || 0); }, 0) * 100) / 100; };
  return {
    low: low, expiring: expiring, expired: expired,
    counts: {
      low: low.length,
      outOfStock: low.filter(function (x) { return x.stock <= 0; }).length,
      expiring: expiring.length,
      urgent: expiring.filter(function (x) { return x.urgent; }).length,
      expired: expired.length,
      items: Object.keys(items).length
    },
    values: { expiring: sum(expiring, 'value'), expired: sum(expired, 'value') },
    cfg: { usageDays: STK_CFG.USAGE_DAYS, leadDays: STK_CFG.LEAD_DAYS,
           nearExpiryDays: STK_CFG.NEAR_EXPIRY_DAYS, urgentDays: STK_CFG.URGENT_EXPIRY_DAYS }
  };
}

/**
 * FRONTEND ENTRY. Low stock, near-expiry and expired batches, and a
 * supplier and last price of each low item. Pharmacy staff and administrators.
 */
function getStockAlerts(sessionToken) {
  try {
    crescRequire_(sessionToken, 'pharmacy.read');
    var out = stk_analyse_();
    out.success = true;
    out.generatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    return out;
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  }
}
