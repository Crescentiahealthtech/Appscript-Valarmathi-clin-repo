// ============================================================================
// Clinic_Profile.gs  —  Crescentia HealthTech
// The letterhead: one source for the clinic's own details.
// ----------------------------------------------------------------------------
// The name, address, phone number and GSTIN printed at the top of an invoice
// were written in five places. Three read Script Properties with three
// different fallbacks ("Crescentia Clinic" in one file, "Crescentia
// HealthTech" in another), and two — the pharmacy invoice and the pharmacy
// returns note — had them typed into the markup as string literals, so a
// clinic that changed its phone number changed it on the lab receipt and
// nowhere else.
//
// A patient holding a pharmacy bill and a lab bill from the same visit can
// see both. They should not name two different clinics.
//
// This is the one reader. Server code calls cresc_clinic_(); the browser
// calls getClinicProfile(), which is unauthenticated ON PURPOSE — it returns
// what is already printed on every document that leaves the building, and
// the invoice renderers in the shell need it before any session exists.
// ============================================================================

/** Script Property key -> field name and default. */
var CRESC_CLINIC_KEYS = {
  CLINIC_NAME:     { field: 'name',       fallback: 'Crescentia HealthTech' },
  CLINIC_ADDRESS:  { field: 'address',    fallback: '' },
  CLINIC_PHONE:    { field: 'phone',      fallback: '' },
  CLINIC_EMAIL:    { field: 'email',      fallback: '' },
  CLINIC_GST:      { field: 'gstin',      fallback: '' },
  CLINIC_REG_NO:   { field: 'regNo',      fallback: '' },
  CLINIC_FOOTER:   { field: 'footer',     fallback: 'This is a computer-generated document.' }
};

/**
 * The clinic's own details. Never throws — a letterhead that cannot be read
 * must not stop a bill being printed.
 *
 * @return {{name,address,phone,email,gstin,regNo,footer}}
 */
function cresc_clinic_() {
  var out = {};
  Object.keys(CRESC_CLINIC_KEYS).forEach(function (k) {
    out[CRESC_CLINIC_KEYS[k].field] = CRESC_CLINIC_KEYS[k].fallback;
  });
  try {
    var props = PropertiesService.getScriptProperties().getProperties() || {};
    Object.keys(CRESC_CLINIC_KEYS).forEach(function (k) {
      var v = props[k];
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        out[CRESC_CLINIC_KEYS[k].field] = String(v).trim();
      }
    });
  } catch (e) { /* the fallbacks are a usable letterhead */ }
  return out;
}

/**
 * FRONTEND ENTRY. The same object, for the invoice renderer in the shell.
 *
 * No session check: this is the text already printed on every receipt the
 * clinic hands out, and the renderer runs before sign-in on a page that is
 * deployed anonymously. Nothing patient-identifying passes through here.
 */
function getClinicProfile() {
  try {
    return { success: true, clinic: cresc_clinic_(), message: '' };
  } catch (err) {
    return { success: false, clinic: null, message: err.message };
  }
}

/**
 * ONE-OFF SETUP / maintenance. Writes the clinic's details into Script
 * Properties, where every printed document reads them from.
 *
 * Edit the values and run it from the script editor. A blank string clears a
 * field; omitting a key leaves it alone.
 */
function setClinicProfile(values) {
  try {
    values = values || {};
    var props = PropertiesService.getScriptProperties();
    var written = [];
    Object.keys(CRESC_CLINIC_KEYS).forEach(function (k) {
      var field = CRESC_CLINIC_KEYS[k].field;
      if (!values.hasOwnProperty(field)) return;
      props.setProperty(k, String(values[field] == null ? '' : values[field]));
      written.push(k);
    });
    if (!written.length) {
      return 'Nothing was written. Pass an object like ' +
             '{ name: "…", address: "…", phone: "…", gstin: "…" }. ' +
             'Current values: ' + JSON.stringify(cresc_clinic_());
    }
    return 'Updated: ' + written.join(', ') + '. Now: ' + JSON.stringify(cresc_clinic_());
  } catch (err) {
    return 'setClinicProfile failed: ' + err.message;
  }
}
