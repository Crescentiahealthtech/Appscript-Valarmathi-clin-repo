// ============================================================================
// IP_Print_Kit.gs  —  Crescentia HealthTech
// Shared print composition for every IP clinical document.
// ----------------------------------------------------------------------------
// WHY THIS FILE EXISTS
//   The casesheet printer and the progress-record printer each carried their
//   own inline CSS and their own idea of a letterhead. They drifted: different
//   fonts, different margins, different header markup, and both laid out
//   two-column blocks with `display:flex`, which print engines collapse or
//   overflow unpredictably. Labels and values were run together as inline
//   text, so nothing lined up down the page.
//
// LAYOUT RULES THIS KIT ENFORCES
//   1. Tables, never flexbox. `table-layout:fixed` is the only construct that
//      keeps a column at the same x-position on every page and in every
//      engine. Label columns are a fixed width, so every value in a document
//      starts on the same vertical line.
//   2. One outer <table class="page"> per document, so the patient-identity
//      strip repeats (thead) at the top of every printed page and the
//      provenance footer repeats (tfoot) at the bottom. A loose sheet from a
//      ward chart is worthless if it does not name its patient.
//   3. Long free text wraps rather than overflowing: overflow-wrap:anywhere.
//   4. Sections avoid breaking mid-block; headings never orphan from bodies.
//   5. print-color-adjust:exact, so shaded rules survive the printer driver.
// ============================================================================

var IPP_CLINIC = {
  name:    "Valarmathi Clinic",
  tagline: "Premium Healthcare Services",
  phone:   "+91 88387 23513"
};

/** Every value that reaches printed HTML goes through this. Clinical free
 *  text routinely contains "<", "&" and quotes. */
function ipp_esc_(v) {
  return String(v === null || v === undefined ? "" : v)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

/** Escape, then turn newlines into <br> so typed paragraphs keep their shape. */
function ipp_escMultiline_(v) {
  return ipp_esc_(v).replace(/\r\n|\r|\n/g, "<br>");
}

/** The single stylesheet every IP document shares. */
function ipp_css_() {
  return [
    '@page{size:A4;margin:13mm 12mm 14mm;}',
    '*{box-sizing:border-box;}',
    'html,body{margin:0;padding:0;}',
    'body{font-family:"Segoe UI",Roboto,Arial,Helvetica,sans-serif;font-size:10.5pt;',
      'line-height:1.45;color:#111;background:#fff;',
      '-webkit-print-color-adjust:exact;print-color-adjust:exact;}',

    // The outer page table: thead repeats per page, tfoot pins provenance.
    'table.page{width:100%;border-collapse:collapse;table-layout:fixed;}',
    'table.page>thead{display:table-header-group;}',
    'table.page>tfoot{display:table-footer-group;}',
    'table.page>thead>tr>td,table.page>tbody>tr>td,table.page>tfoot>tr>td{padding:0;}',

    // Letterhead (page 1 only) and the slim per-page identity strip.
    '.lh{text-align:center;border-bottom:2px solid #0369a1;padding-bottom:8px;margin-bottom:12px;}',
    '.lh h1{margin:0;font-size:17pt;letter-spacing:.06em;text-transform:uppercase;color:#0369a1;}',
    '.lh .sub{margin:2px 0 0;font-size:8.5pt;color:#555;}',
    '.lh h2{margin:7px 0 0;font-size:12pt;font-weight:700;color:#0f172a;}',
    '.runhead{font-size:8pt;color:#475569;border-bottom:.6pt solid #cbd5e1;',
      'padding-bottom:3px;margin-bottom:9px;}',
    '.runhead .r{float:right;}',
    '.runfoot{font-size:7.5pt;color:#64748b;border-top:.6pt solid #cbd5e1;',
      'padding-top:3px;margin-top:9px;}',
    '.runfoot .r{float:right;}',

    // Sections.
    '.sec{margin:0 0 13px;page-break-inside:avoid;break-inside:avoid;}',
    '.sec.loose{page-break-inside:auto;break-inside:auto;}',
    '.sec>h3{margin:0 0 6px;font-size:10pt;font-weight:700;color:#0369a1;',
      'text-transform:uppercase;letter-spacing:.04em;',
      'border-bottom:.8pt solid #cbd5e1;padding-bottom:3px;',
      'page-break-after:avoid;break-after:avoid;}',

    // Label/value grid — the fixed label column is what makes things line up.
    'table.kv{width:100%;border-collapse:collapse;table-layout:fixed;}',
    'table.kv th{width:38mm;text-align:left;vertical-align:top;font-weight:600;',
      'color:#334155;padding:2px 6px 2px 0;font-size:9.5pt;}',
    'table.kv td{vertical-align:top;padding:2px 0;overflow-wrap:anywhere;word-break:break-word;}',
    'table.kv.narrow th{width:31mm;}',
    'table.kv tr{page-break-inside:avoid;break-inside:avoid;}',

    // Two equal columns that hold their x-position across pages.
    'table.cols{width:100%;border-collapse:collapse;table-layout:fixed;}',
    'table.cols>tbody>tr>td{width:50%;vertical-align:top;padding:0 6mm 0 0;}',
    'table.cols>tbody>tr>td:last-child{padding:0 0 0 6mm;}',

    // Data tables.
    'table.dt{width:100%;border-collapse:collapse;table-layout:fixed;font-size:9.5pt;}',
    'table.dt th{background:#eef4f8;text-align:left;font-weight:600;color:#0f172a;',
      'border:.6pt solid #cbd5e1;padding:4px 6px;}',
    'table.dt td{border:.6pt solid #cbd5e1;padding:4px 6px;vertical-align:top;',
      'overflow-wrap:anywhere;word-break:break-word;}',
    'table.dt tr{page-break-inside:avoid;break-inside:avoid;}',
    'table.dt thead{display:table-header-group;}',
    'table.dt td.num,table.dt th.num{text-align:right;font-variant-numeric:tabular-nums;}',
    'table.dt td.ctr,table.dt th.ctr{text-align:center;}',

    // A single timeline entry.
    '.note{border:.6pt solid #cbd5e1;border-left:2.5pt solid #0369a1;border-radius:3px;',
      'padding:7px 9px;margin:0 0 8px;page-break-inside:avoid;break-inside:avoid;}',
    '.note>.nh{font-size:8.5pt;color:#475569;border-bottom:.5pt dotted #cbd5e1;',
      'padding-bottom:3px;margin-bottom:5px;}',
    '.note>.nh .r{float:right;}',
    '.note .tag{display:inline-block;font-size:7.5pt;font-weight:700;',
      'text-transform:uppercase;letter-spacing:.04em;color:#fff;background:#0369a1;',
      'border-radius:2px;padding:1px 5px;margin-right:5px;}',

    '.muted{color:#64748b;}',
    '.dx{font-size:11.5pt;font-weight:700;}',
    '.clearfix::after{content:"";display:block;clear:both;}',

    // Signature block.
    '.sign{margin-top:34px;text-align:right;page-break-inside:avoid;break-inside:avoid;}',
    '.sign .box{display:inline-block;width:62mm;text-align:center;',
      'border-top:.8pt solid #111;padding-top:4px;font-size:9.5pt;}',
    '.sign .box .cap{display:block;font-size:7.5pt;color:#64748b;margin-top:1px;}',

    '@media print{.noprint{display:none !important;}}'
  ].join("");
}

/**
 * Label/value block.
 * @param {Array<Array>} rows  [[label, valueHtml], ...] — value is ALREADY escaped
 * @param {Object} [opts]      {narrow:true, keepEmpty:true}
 * @return {string} "" when every row is empty, so callers can drop the section
 */
function ipp_kv_(rows, opts) {
  opts = opts || {};
  var body = (rows || []).filter(function (r) {
    return opts.keepEmpty || String(r[1] === null || r[1] === undefined ? "" : r[1]).trim() !== "";
  }).map(function (r) {
    var v = String(r[1] === null || r[1] === undefined ? "" : r[1]).trim();
    return '<tr><th>' + ipp_esc_(r[0]) + '</th><td>' + (v || '<span class="muted">--</span>') + '</td></tr>';
  }).join("");
  if (!body) return "";
  return '<table class="kv' + (opts.narrow ? ' narrow' : '') + '"><tbody>' + body + '</tbody></table>';
}

/** A titled section. Returns "" for empty content, so no blank headings print. */
function ipp_sec_(title, innerHtml, opts) {
  opts = opts || {};
  var inner = String(innerHtml || "").trim();
  if (!inner && !opts.keepEmpty) return "";
  return '<div class="sec' + (opts.loose ? ' loose' : '') + '">' +
           (title ? '<h3>' + ipp_esc_(title) + '</h3>' : '') +
           (inner || '<span class="muted">Not recorded.</span>') +
         '</div>';
}

/** Two columns that stay aligned across a page break. */
function ipp_cols_(leftHtml, rightHtml) {
  var l = String(leftHtml || "").trim(), r = String(rightHtml || "").trim();
  if (!l && !r) return "";
  if (!l || !r) return l || r;   // one column alone should use the full width
  return '<table class="cols"><tbody><tr><td>' + l + '</td><td>' + r + '</td></tr></tbody></table>';
}

/**
 * A bordered data table.
 * @param {Array} headers  strings, or {label, cls} for alignment
 * @param {Array<Array<string>>} rows  cell HTML, already escaped
 * @param {Array<string>} [colWidths]  CSS widths, positional
 */
function ipp_table_(headers, rows, colWidths) {
  if (!rows || !rows.length) return "";
  var cols = (colWidths || []).map(function (w) {
    return '<col style="width:' + w + '">';
  }).join("");
  var head = '<tr>' + headers.map(function (h) {
    var label = (h && h.label !== undefined) ? h.label : h;
    var cls   = (h && h.cls) ? ' class="' + h.cls + '"' : '';
    return '<th' + cls + '>' + ipp_esc_(label) + '</th>';
  }).join("") + '</tr>';
  var body = rows.map(function (r) {
    return '<tr>' + r.map(function (c, i) {
      var cls = (headers[i] && headers[i].cls) ? ' class="' + headers[i].cls + '"' : '';
      return '<td' + cls + '>' + (String(c === null || c === undefined ? "" : c) || '<span class="muted">--</span>') + '</td>';
    }).join("") + '</tr>';
  }).join("");
  return '<table class="dt">' + (cols ? '<colgroup>' + cols + '</colgroup>' : '') +
         '<thead>' + head + '</thead><tbody>' + body + '</tbody></table>';
}

/** The signature rule at the foot of a signed document. */
function ipp_sig_(name, caption) {
  return '<div class="sign"><div class="box"><strong>' + ipp_esc_(name || "Doctor's Signature") + '</strong>' +
         (caption ? '<span class="cap">' + ipp_esc_(caption) + '</span>' : '') +
         '</div></div>';
}

/** "dd-MMM-yyyy hh:mm a" in the script's timezone, tolerant of junk input. */
function ipp_when_(d, pattern) {
  try {
    // Shared parser (Date_Utils.gs): a printed document is the last place a
    // dd-mm-yyyy cell should come out blank or a month wrong.
    var dt = cresc_toDate_(d);
    if (!dt) return "";
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), pattern || "dd-MMM-yyyy hh:mm a");
  } catch (e) { return ""; }
}

/**
 * The patient-ID barcode for a letterhead, or "" when it cannot be drawn.
 *
 * Barcode_Print.gs is a separate file, and Apps Script leaves a function
 * undefined rather than failing to load when a file was not copied across.
 * A missing encoder must cost this document its barcode, never its printing.
 */
function ipp_barcodeCell_(patientId) {
  try {
    if (typeof bcp_patientBarcodeBlock_ !== 'function') return "";
    return bcp_patientBarcodeBlock_(patientId, { align: 'right', height: 8 });
  } catch (e) { return ""; }
}

/**
 * Wraps composed sections into a complete printable document.
 *
 * @param {Object} o
 *   docTitle   {string} e.g. "Inpatient Progress Record"
 *   patient    {name, pid, ipNumber, ageSex, wardBed, consultant, diagnosis}
 *   bannerHtml {string} optional, replaces the default identity panel
 *   bodyHtml   {string} the composed sections
 *   footNote   {string} optional extra provenance text
 * @return {string} full HTML document
 */
function ipp_doc_(o) {
  o = o || {};
  var p = o.patient || {};
  var printedAt = ipp_when_(new Date());

  // The slim strip that repeats at the top of every page.
  var strip =
    '<div class="runhead clearfix">' +
      '<span><strong>' + ipp_esc_(p.name || "--") + '</strong>' +
        (p.ageSex ? ' (' + ipp_esc_(p.ageSex) + ')' : '') +
        ' &nbsp;&middot;&nbsp; IP ' + ipp_esc_(p.ipNumber || "--") +
        ' &nbsp;&middot;&nbsp; PID ' + ipp_esc_(p.pid || "--") + '</span>' +
      '<span class="r">' + ipp_esc_(o.docTitle || "") + '</span>' +
    '</div>';

  var foot =
    '<div class="runfoot clearfix">' +
      '<span>' + ipp_esc_(IPP_CLINIC.name) + ' &middot; ' + ipp_esc_(p.name || "--") +
        ' &middot; IP ' + ipp_esc_(p.ipNumber || "--") + '</span>' +
      '<span class="r">Printed ' + ipp_esc_(printedAt) +
        (o.footNote ? ' &middot; ' + ipp_esc_(o.footNote) : '') + '</span>' +
    '</div>';

  // The patient-ID barcode rides on the letterhead of every IP document, so
  // any sheet the patient or the ward is holding can be scanned back to the
  // record. It sits in its own fixed-width table cell rather than floating:
  // floats are the one construct the PDF converter reflows unpredictably.
  var barcode = ipp_barcodeCell_(p.pid);

  var letterheadInner =
    '<div class="lh"' + (barcode ? ' style="border-bottom:0;margin-bottom:0;padding-bottom:0;"' : '') + '>' +
      '<h1>' + ipp_esc_(IPP_CLINIC.name) + '</h1>' +
      '<p class="sub">' + ipp_esc_(IPP_CLINIC.tagline) + ' &nbsp;|&nbsp; Ph: ' + ipp_esc_(IPP_CLINIC.phone) + '</p>' +
      '<h2>' + ipp_esc_(o.docTitle || "Clinical Record") + '</h2>' +
    '</div>';

  // A spacer cell mirrors the barcode cell so the clinic name stays centred on
  // the page rather than drifting left by the width of the barcode.
  // Spacing lives on the CELLS, not the table: under border-collapse a
  // table's own padding is ignored, which would have pulled the rule up
  // against the clinic name.
  var letterhead = barcode
    ? '<table style="width:100%;border-collapse:collapse;table-layout:fixed;' +
        'border-bottom:2px solid #0369a1;margin-bottom:12px;">' +
        '<tr>' +
          '<td style="width:48mm;padding:0 0 8px 0;"></td>' +
          '<td style="vertical-align:bottom;padding:0 0 8px 0;">' + letterheadInner + '</td>' +
          '<td style="width:48mm;vertical-align:bottom;padding:0 0 8px 0;">' + barcode + '</td>' +
        '</tr></table>'
    : letterheadInner;

  var banner = (o.bannerHtml !== undefined) ? o.bannerHtml : ipp_patientBanner_(p);

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<title>' + ipp_esc_((o.docTitle || "Record") + " - " + (p.ipNumber || "")) + '</title>' +
    '<style>' + ipp_css_() + '</style></head><body>' +
    '<table class="page">' +
      '<thead><tr><td>' + strip + '</td></tr></thead>' +
      '<tfoot><tr><td>' + foot + '</td></tr></tfoot>' +
      '<tbody><tr><td>' + letterhead + banner + (o.bodyHtml || "") + '</td></tr></tbody>' +
    '</table></body></html>';
}

/** The boxed patient identity panel under the letterhead, two aligned columns. */
function ipp_patientBanner_(p) {
  p = p || {};
  var left = ipp_kv_([
    ["Patient",   ipp_esc_(p.name)],
    ["Patient ID", ipp_esc_(p.pid)],
    ["Age / Sex", ipp_esc_(p.ageSex)]
  ], { narrow: true, keepEmpty: true });

  var right = ipp_kv_([
    ["IP Number",  ipp_esc_(p.ipNumber)],
    ["Ward / Bed", ipp_esc_(p.wardBed)],
    ["Consultant", ipp_esc_(p.consultant)]
  ], { narrow: true, keepEmpty: true });

  var dx = p.diagnosis
    ? ipp_kv_([["Working Dx", ipp_esc_(p.diagnosis)]], { narrow: true })
    : "";

  return '<div style="border:.8pt solid #cbd5e1;border-radius:4px;background:#f8fafc;' +
           'padding:8px 10px;margin-bottom:13px;page-break-inside:avoid;break-inside:avoid;">' +
           ipp_cols_(left, right) + dx +
         '</div>';
}

// ---------------------------------------------------------------------------
// VITALS TREND — an observation chart the ward can read at a glance
//
// Nursing notes already carry BP, pulse, SpO2 and temperature at every round.
// Printed one note at a time those numbers say nothing; printed as a series
// they show the patient's direction of travel, which is the single thing a
// consultant wants from a progress record. The sparkline is inline SVG, so it
// is vector on paper and needs no charting library inside Apps Script.
// ---------------------------------------------------------------------------

/**
 * @param {Array<{ts:Date|string, label:string, vitals:Object}>} series
 * @return {string} "" when there is nothing worth charting
 */
function ipp_vitalsTrend_(series) {
  var pts = (series || []).filter(function (s) { return s && s.vitals; });
  if (pts.length < 1) return "";

  var num = function (v) {
    var n = parseFloat(String(v === null || v === undefined ? "" : v).replace(/[^0-9.\-]/g, ""));
    return isFinite(n) ? n : null;
  };
  // "120/80" -> systolic. A single number is taken as the systolic reading.
  var sysOf = function (bp) {
    var s = String(bp || "");
    var m = s.match(/(\d{2,3})\s*\/\s*(\d{2,3})/);
    return m ? parseFloat(m[1]) : num(s);
  };

  var rows = pts.map(function (s) {
    var v = s.vitals || {};
    return {
      label: s.label || ipp_when_(s.ts, "dd-MMM HH:mm"),
      bp:    String(v.bp || "").trim(),
      sys:   sysOf(v.bp),
      pulse: num(v.pulse),
      spo2:  num(v.spo2),
      temp:  num(v.temp)
    };
  }).filter(function (r) {
    return r.bp || r.pulse !== null || r.spo2 !== null || r.temp !== null;
  });
  if (!rows.length) return "";

  var spark = function (vals, colour, lo, hi) {
    var real = vals.filter(function (v) { return v !== null; });
    if (real.length < 2) return '<span class="muted">--</span>';
    var min = Math.min.apply(null, real), max = Math.max.apply(null, real);
    if (lo !== undefined) min = Math.min(min, lo);
    if (hi !== undefined) max = Math.max(max, hi);
    if (max - min < 1) { max = min + 1; }

    var W = 150, H = 26, pad = 2;
    var step = (vals.length > 1) ? (W - pad * 2) / (vals.length - 1) : 0;
    var pathPts = [], dots = [];
    vals.forEach(function (v, i) {
      if (v === null) return;
      var x = pad + i * step;
      var y = H - pad - ((v - min) / (max - min)) * (H - pad * 2);
      pathPts.push((pathPts.length ? "L" : "M") + x.toFixed(1) + " " + y.toFixed(1));
      dots.push('<circle cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="1.4" fill="' + colour + '"/>');
    });
    return '<svg width="' + W + '" height="' + H + '" viewBox="0 0 ' + W + ' ' + H +
           '" preserveAspectRatio="none" style="display:block;">' +
           '<path d="' + pathPts.join(" ") + '" fill="none" stroke="' + colour +
           '" stroke-width="1.1" stroke-linejoin="round" stroke-linecap="round"/>' +
           dots.join("") + '</svg>';
  };

  var range = function (vals, unit) {
    var real = vals.filter(function (v) { return v !== null; });
    if (!real.length) return "";
    var min = Math.min.apply(null, real), max = Math.max.apply(null, real);
    var last = real[real.length - 1];
    return ipp_esc_(last + (unit || "")) +
           (min === max ? "" : '<span class="muted"> (' + ipp_esc_(min + "–" + max) + ')</span>');
  };

  // A parameter with no readings at all is left off the chart entirely rather
  // than printed as a row of dashes.
  var trendRows = [
    ["Systolic BP", rows.map(function (r) { return r.sys; }),   "#b91c1c", 90, 140, " mmHg"],
    ["Pulse",       rows.map(function (r) { return r.pulse; }), "#0369a1", 60, 100, " bpm"],
    ["SpO2",        rows.map(function (r) { return r.spo2; }),  "#047857", 92, 100, " %"],
    ["Temp",        rows.map(function (r) { return r.temp; }),  "#b45309", 97, 100, " °F"]
  ].filter(function (t) {
    return t[1].filter(function (v) { return v !== null; }).length > 0;
  }).map(function (t) {
    return [t[0], spark(t[1], t[2], t[3], t[4]), range(t[1], t[5])];
  });
  if (!trendRows.length) return "";

  var chart = ipp_table_(
    ["Parameter", { label: "Trend (" + rows.length + " readings)", cls: "ctr" }, { label: "Latest (range)", cls: "num" }],
    trendRows, ["30mm", "auto", "42mm"]
  );

  var log = ipp_table_(
    ["Recorded", "BP", { label: "Pulse", cls: "num" }, { label: "SpO2", cls: "num" }, { label: "Temp", cls: "num" }],
    rows.map(function (r) {
      return [
        ipp_esc_(r.label), ipp_esc_(r.bp),
        r.pulse !== null ? ipp_esc_(r.pulse) : "",
        r.spo2  !== null ? ipp_esc_(r.spo2)  : "",
        r.temp  !== null ? ipp_esc_(r.temp)  : ""
      ];
    }),
    ["auto", "26mm", "20mm", "20mm", "20mm"]
  );

  return chart + '<div style="height:8px;"></div>' + log;
}