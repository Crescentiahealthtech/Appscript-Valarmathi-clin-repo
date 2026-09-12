// ============================================================================
// Barcode_Print.gs  —  Crescentia HealthTech / CresRx
// SERVER-SIDE CODE 128 for every printed clinical document.
// ----------------------------------------------------------------------------
// WHY THIS EXISTS
//   A patient ID card is one more thing to lose. The reliable place for a
//   machine-readable patient ID is the paper the patient is already holding:
//   the prescription, the case sheet, the progress record, the discharge
//   summary, the lab report. Print the barcode on all of them and any of them
//   becomes the card — the front desk scans whatever the patient brought.
//
//   Barcode_Labels.html does this in the browser for sticker printing. PDFs,
//   however, are assembled on the server (Utilities.newBlob(...).getAs(pdf)),
//   which never runs page JavaScript, so the same encoder has to exist here.
//
// WHY SVG AND NOT AN IMAGE
//   Apps Script's HTML-to-PDF converter renders inline SVG as vector. A
//   base64 PNG would depend on the data-URI image path, which the discharge
//   engine's own print spike (DS_Print.gs) exists precisely because it does
//   not trust. Bars drawn as <rect> always print, and stay sharp at any
//   scale — which is what a scanner needs.
//
// FIRST-TIME PATIENTS
//   A patient ID is allocated at registration (bc_nextPatientId_), before any
//   document can be produced, so there is no window in which a printable
//   document lacks a scannable ID. bcp_patientBarcodeBlock_() still degrades
//   honestly: with no ID it prints nothing rather than an unscannable box, and
//   with an ID that Code 128-B cannot represent it prints the ID as plain
//   monospace text so a human can still key it in.
// ============================================================================

/**
 * Code 128 subset B pattern table (JsBarcode, MIT — lindell/JsBarcode).
 * Subset B covers ASCII 32-126, which is every character CresRx IDs use.
 * Subset C would pack digit pairs tighter, but a mis-encoded clinical
 * document is a patient-safety event and B is correct for every input here.
 */
var BCP_BARS = [
  '11011001100','11001101100','11001100110','10010011000','10010001100','10001001100','10011001000','10011000100','10001100100','11001001000',
  '11001000100','11000100100','10110011100','10011011100','10011001110','10111001100','10011101100','10011100110','11001110010','11001011100',
  '11001001110','11011100100','11001110100','11101101110','11101001100','11100101100','11100100110','11101100100','11100110100','11100110010',
  '11011011000','11011000110','11000110110','10100011000','10001011000','10001000110','10110001000','10001101000','10001100010','11010001000',
  '11000101000','11000100010','10110111000','10110001110','10001101110','10111011000','10111000110','10001110110','11101110110','11010001110',
  '11000101110','11011101000','11011100010','11011101110','11101011000','11101000110','11100010110','11101101000','11101100010','11100011010',
  '11101111010','11001000010','11110001010','10100110000','10100001100','10010110000','10010000110','10000101100','10000100110','10110010000',
  '10110000100','10011010000','10011000010','10000110100','10000110010','11000010010','11001010000','11110111010','11000010100','10001111010',
  '10100111100','10010111100','10010011110','10111100100','10011110100','10011110010','11110100100','11110010100','11110010010','11011011110',
  '11011110110','11110110110','10101111000','10100011110','10001011110','10111101000','10111100010','11110101000','11110100010','10111011110',
  '10111101110','11101011110','11110101110','11010000100','11010010000','11010011100','1100011101011'
];
var BCP_START_B = 104, BCP_STOP = 106, BCP_MODULO = 103;

/** Escape for attribute/text contexts. Never trust an ID to be clean. */
function bcp_esc_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * @param {string} text
 * @return {string} module string ("1" = bar, "0" = space); "" if unencodable.
 */
function bcp_code128_(text) {
  var s = String(text === null || text === undefined ? '' : text);
  if (!s) return '';

  var values = [BCP_START_B];
  for (var i = 0; i < s.length; i++) {
    var code = s.charCodeAt(i);
    if (code < 32 || code > 126) return '';        // outside subset B
    values.push(code - 32);
  }

  // Checksum: start value + sum(position * value), positions 1-based.
  var sum = BCP_START_B;
  for (var j = 1; j < values.length; j++) sum += j * values[j];
  values.push(sum % BCP_MODULO);
  values.push(BCP_STOP);

  var out = '';
  for (var k = 0; k < values.length; k++) out += BCP_BARS[values[k]];
  return out;
}

/**
 * Code 128 as inline SVG, sized in millimetres.
 *
 * @param {string} text
 * @param {{moduleWidth:number, height:number, quiet:number, text:boolean,
 *          fontSize:number}} [opts]
 *        moduleWidth mm per narrow module (default .30 — at 600 dpi that is
 *        ~7 dots, comfortably above the ~2-dot floor a scanner needs)
 *        height      bar height in mm (default 8)
 *        quiet       quiet-zone width in modules each side (default 10; the
 *                    spec's minimum is 10, and a barcode without it does not
 *                    scan however well the bars are printed)
 * @return {string} "" when the text cannot be encoded.
 */
function bcp_code128Svg_(text, opts) {
  var o = opts || {};
  var modules = bcp_code128_(text);
  if (!modules) return '';

  var mw    = o.moduleWidth || 0.30;
  var h     = o.height || 8;
  var quiet = (o.quiet === null || o.quiet === undefined) ? 10 : o.quiet;
  var showText = (o.text !== false);
  var fs    = o.fontSize || 2.6;
  var textH = showText ? (fs + 0.8) : 0;
  var w     = (modules.length + quiet * 2) * mw;

  var rects = '', run = 0;
  for (var i = 0; i <= modules.length; i++) {
    if (modules.charAt(i) === '1') { run++; continue; }
    if (run) {
      var x = (quiet + i - run) * mw;
      rects += '<rect x="' + x.toFixed(3) + '" y="0" width="' + (run * mw).toFixed(3) +
               '" height="' + h + '" fill="#000"/>';
      run = 0;
    }
  }

  var label = showText
    ? '<text x="' + (w / 2).toFixed(2) + '" y="' + (h + fs).toFixed(2) +
      '" font-family="monospace" font-size="' + fs + '" text-anchor="middle" ' +
      'fill="#000" letter-spacing="0.12">' + bcp_esc_(text) + '</text>'
    : '';

  // max-width:100% matters: an ID longer than the block it sits in would
  // otherwise be CLIPPED, and a clipped Code 128 is unscannable with no
  // outward sign of it. Scaling down costs module width — still well above a
  // scanner's floor at print resolution — and keeps the whole symbol on the
  // page, which is the only version of it worth printing.
  return '<svg xmlns="http://www.w3.org/2000/svg" width="' + w.toFixed(2) + 'mm" height="' +
         (h + textH).toFixed(2) + 'mm" viewBox="0 0 ' + w.toFixed(2) + ' ' +
         (h + textH).toFixed(2) + '" preserveAspectRatio="xMaxYMid meet" ' +
         'style="max-width:100%;height:auto;" shape-rendering="crispEdges" ' +
         'role="img" aria-label="Patient ID ' + bcp_esc_(text) + '">' + rects + label + '</svg>';
}

/**
 * The patient-ID barcode block that goes in a document letterhead.
 *
 * Degrades in two steps, because a clinical document must never be withheld
 * over a decoration:
 *   no ID at all  -> "" (nothing printed)
 *   unencodable   -> the ID as plain monospace text, still human-keyable
 *
 * @param {string} patientId
 * @param {{caption:string, height:number, moduleWidth:number,
 *          align:string}} [opts]
 */
function bcp_patientBarcodeBlock_(patientId, opts) {
  var o = opts || {};
  var id = String(patientId === null || patientId === undefined ? '' : patientId).trim();
  if (!id) return '';

  var align = o.align || 'right';
  var caption = (o.caption === undefined) ? 'PATIENT ID' : o.caption;
  var capHtml = caption
    ? '<div style="font-size:6.5pt;letter-spacing:.09em;color:#475569;' +
      'text-transform:uppercase;margin-bottom:1px;">' + bcp_esc_(caption) + '</div>'
    : '';

  var svg = bcp_code128Svg_(id, {
    moduleWidth: o.moduleWidth || 0.30,
    height: o.height || 8,
    fontSize: o.fontSize || 2.6
  });

  var body = svg
    ? svg
    : '<div style="font-family:monospace;font-size:11pt;font-weight:700;' +
      'letter-spacing:.12em;color:#111;">' + bcp_esc_(id) + '</div>';

  return '<div style="text-align:' + bcp_esc_(align) + ';line-height:1;">' +
         capHtml + body + '</div>';
}

/**
 * ADMIN / DIAGNOSTIC. Renders one barcode so the encoder can be eyeballed in
 * the Apps Script editor before a print run. Returns the SVG string.
 */
function previewPatientBarcode(patientId) {
  var id = String(patientId || 'LMTVS0001');
  var svg = bcp_code128Svg_(id, { height: 12, moduleWidth: 0.4, fontSize: 3.2 });
  Logger.log(svg ? ('Encoded ' + id + ' in ' + svg.length + ' bytes of SVG.')
                 : ('Could not encode "' + id + '" in Code 128-B.'));
  return svg;
}
