// =========================================================================
// 🗂️ DISCHARGE SUMMARY — PICKERS
// Crescentia HealthTech / CresRx
// -------------------------------------------------------------------------
// Everything the summary editor needs in order to offer a CHOICE rather than
// a blank box, fetched in ONE call and cached by the client for the editing
// session:
//
//   • the clinic's phrase library for complaints, history and advice — the
//     same three lists OP and IP prescribe from (OP_Templates_Engine), so a
//     condition worded one way in the consult is worded that way here;
//   • the drug formulary with live pharmacy stock, so the discharge script
//     autocompletes exactly as prescribing does in OP and IP;
//   • the field-label table, so a box's label on screen is the label that
//     prints beside it on paper.
//
// Read-only. Nothing here writes, and nothing here is patient-specific, so
// one fetch serves every summary the user opens.
// =========================================================================

/** Cap on the formulary sent to the browser. */
var DSP_DRUG_LIMIT = 1200;

/**
 * FRONTEND ENTRY. The editor's pick lists.
 *
 * @param {string} sessionToken
 * @return {{success:boolean, cc:Array, hx:Array, advice:Array,
 *           drugs:Array, fieldLabels:Object, message:string}}
 */
function ds_getPickers(sessionToken) {
  try {
    // Any role that may open the desk may see these lists; they are clinic
    // reference data, not a patient's record.
    var actor = dsx_requireRole_(sessionToken, []);

    var out = {
      success: true,
      cc: [], hx: [], advice: [],
      drugs: [],
      fieldLabels: DSX_FIELD_LABELS,
      message: ''
    };

    // ---- phrase library ---------------------------------------------------
    // Scoped exactly as the consult scopes it: the signed-in doctor's own
    // phrases plus the clinic's. A preparer with no doctor profile still gets
    // the clinic list, which is the point of having one.
    try {
      if (typeof getScopedTemplates === 'function') {
        var t = getScopedTemplates(actor.doctorId || '', sessionToken);
        if (t && t.success) {
          out.cc     = dsp_phrases_(t.CC);
          out.hx     = dsp_phrases_(t.HX);
          out.advice = dsp_phrases_(t.ADVICE);
        }
      }
    } catch (e) {
      out.message = 'The phrase library is unavailable (' + e.message + '); ' +
                    'the boxes still accept free text.';
    }

    // ---- formulary + stock ------------------------------------------------
    try {
      out.drugs = dsp_formulary_();
    } catch (e) {
      out.message = (out.message ? out.message + ' ' : '') +
                    'The drug list is unavailable (' + e.message + '); ' +
                    'drug names can still be typed.';
    }

    return out;
  } catch (err) {
    // FORBIDDEN from dsx_requireRole_ included: the editor degrades to plain
    // text boxes rather than showing an error over a document being written.
    return { success: false, cc: [], hx: [], advice: [], drugs: [],
             fieldLabels: DSX_FIELD_LABELS, message: err.message };
  }
}

/** {text, count} pairs, commonest first, de-duplicated case-insensitively. */
function dsp_phrases_(list) {
  var seen = {}, out = [];
  (list || []).forEach(function (p) {
    var text = dsx_str_(p && (p.text || p));
    if (!text) return;
    var k = text.toLowerCase();
    if (seen[k]) return;
    seen[k] = true;
    out.push({ text: text, count: dsx_int_(p && p.count), scope: dsx_str_(p && p.scope) });
  });
  out.sort(function (a, b) { return (b.count - a.count) || a.text.localeCompare(b.text); });
  return out;
}

/**
 * The prescribable formulary: pharmacy stock first, then the universal drug
 * master for anything the clinic does not carry.
 *
 * `stock` and `status` ride along so the picker can say "out of stock" at the
 * moment of prescribing — a discharge script for a drug the pharmacy cannot
 * dispense sends the patient away and back again.
 */
function dsp_formulary_() {
  var out = [], seen = {};

  var push = function (d, source) {
    var brand = dsx_str_(d.brand);
    if (!brand) return;
    var k = brand.toLowerCase();
    if (seen[k]) return;
    seen[k] = true;
    out.push({
      brand:   brand,
      generic: dsx_str_(d.generic),
      type:    dsx_str_(d.type) || 'Tab',
      dose:    dsx_str_(d.adultDose) || dsx_str_(d.refDose),
      stock:   (d.stock === null || d.stock === undefined) ? null : dsx_int_(d.stock),
      status:  dsx_str_(d.status) || source
    });
  };

  if (typeof fetchOPDrugMaster === 'function') {
    (fetchOPDrugMaster() || []).forEach(function (d) { push(d, 'internal'); });
  }
  if (out.length < DSP_DRUG_LIMIT && typeof fetchUniversalDrugs === 'function') {
    (fetchUniversalDrugs() || []).forEach(function (d) {
      if (out.length < DSP_DRUG_LIMIT) push(d, 'external');
    });
  }

  // In stock first, then alphabetical: what the pharmacy can hand over now
  // is what should be easiest to pick.
  out.sort(function (a, b) {
    var as = (a.stock > 0) ? 0 : 1, bs = (b.stock > 0) ? 0 : 1;
    return (as - bs) || a.brand.localeCompare(b.brand);
  });
  return out.slice(0, DSP_DRUG_LIMIT);
}
