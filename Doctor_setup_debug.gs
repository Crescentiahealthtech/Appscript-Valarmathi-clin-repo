// ============================================================================
// Doctor_Setup_Debug.gs  —  Crescentia HealthTech
// TEMPORARY. Run debugDoctorSetup() from the editor and paste the output.
// Diagnoses: empty doctor picker, unlinked login, lab catalog bridge.
// ============================================================================

function debugDoctorSetup() {
  var out = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  out.push("=== DOCTOR SETUP DIAGNOSTIC ===");
  out.push("Script tenant  : '" + getTenantId_() + "'");
  out.push("");

  // ---- 1. Doctors sheet, raw ---------------------------------------------
  var doc = ss.getSheetByName("Doctors");
  if (!doc) {
    out.push("FAIL: no sheet named 'Doctors'. Run setupDoctorsSheet().");
    Logger.log(out.join("\n")); return out.join("\n");
  }

  var hdr = doc.getRange(1, 1, 1, doc.getLastColumn()).getDisplayValues()[0];
  out.push("Doctors headers (in order):");
  hdr.forEach(function (h, i) { out.push("   [" + i + "] " + h); });
  out.push("");

  var EXPECTED = ["Doctor_ID", "Tenant_ID", "Display_Name", "Specialty",
                  "Reg_No", "Signature_Line", "Linked_Username", "Status"];
  var orderOk = true;
  EXPECTED.forEach(function (h, i) {
    if (String(hdr[i] || "").trim() !== h) {
      orderOk = false;
      out.push("   >>> MISMATCH at index " + i + ": expected '" + h +
               "', found '" + (hdr[i] || "(blank)") + "'");
    }
  });
  out.push(orderOk
    ? "Column order OK."
    : ">>> COLUMN ORDER IS WRONG. Doctors_Engine.gs reads by fixed index, so this alone empties the picker.");
  out.push("");

  // ---- 2. Row-by-row, showing why each row passes or fails ---------------
  var data = doc.getDataRange().getDisplayValues();
  if (data.length < 2) {
    out.push("FAIL: Doctors sheet has headers but no rows.");
    Logger.log(out.join("\n")); return out.join("\n");
  }

  var tenant = getTenantId_();
  out.push("Rows (" + (data.length - 1) + "):");
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0] || "").trim();
    var ten = String(data[i][1] || "").trim();
    var nm = String(data[i][2] || "").trim();
    var user = String(data[i][6] || "").trim();
    var st = String(data[i][7] || "").trim();

    var reasons = [];
    if (!id) reasons.push("blank Doctor_ID");
    if (ten !== tenant) reasons.push("Tenant_ID '" + ten + "' != '" + tenant + "'");
    if (st.toUpperCase() !== "ACTIVE") reasons.push("Status '" + st + "' is not exactly ACTIVE");
    if (!user) reasons.push("Linked_Username blank (login will not resolve to this doctor)");

    out.push("   " + id + " | " + nm + " | user='" + user + "' | status='" + st + "'" +
             (reasons.length ? "   >>> EXCLUDED: " + reasons.join("; ") : "   OK"));
  }
  out.push("");

  // ---- 3. What the picker actually returns -------------------------------
  try {
    var active = getActiveDoctors();
    out.push("getActiveDoctors() returned " + active.length + " doctor(s).");
    if (!active.length) {
      out.push(">>> THIS IS WHY THE DROPDOWN IS EMPTY. Fix the exclusions above.");
    } else {
      active.forEach(function (d) { out.push("   " + d.doctorId + " — " + d.name); });
    }
  } catch (e) {
    out.push("getActiveDoctors() THREW: " + e.message);
  }
  out.push("");

  // ---- 4. Users sheet linkage --------------------------------------------
  var users = ss.getSheetByName("Users");
  if (!users) out.push("No 'Users' sheet found — check the actual name.");
  else {
    var uh = users.getRange(1, 1, 1, users.getLastColumn()).getDisplayValues()[0];
    out.push("Users headers: " + uh.join(" | "));
    var ud = users.getDataRange().getDisplayValues();
    out.push("Usernames on file (first column):");
    for (var u = 1; u < Math.min(ud.length, 25); u++) {
      var uname = String(ud[u][0] || "").trim();
      if (!uname) continue;
      var linked = "";
      for (var d2 = 1; d2 < data.length; d2++) {
        if (String(data[d2][6] || "").trim().toLowerCase() === uname.toLowerCase()) {
          linked = " -> linked to " + data[d2][0];
        }
      }
      out.push("   '" + uname + "'" + (linked || "   >>> NOT linked to any Doctors row"));
    }
    out.push("");
    out.push("A doctor login only auto-selects itself when Users column A matches");
    out.push("Doctors.Linked_Username EXACTLY (case-insensitively, no trailing spaces).");
  }
  out.push("");

  // ---- 5. Lab catalog bridge --------------------------------------------
  try {
    var cat = getOrderableTests();
    out.push("getOrderableTests(): success=" + cat.success +
             ", tests=" + ((cat.tests || []).length));
    if (cat.tests && cat.tests.length) {
      out.push("   sample: " + cat.tests[0].testId + " / " + cat.tests[0].testName +
               " / " + cat.tests[0].sampleType + " / Rs." + cat.tests[0].price);
    } else {
      out.push("   >>> Lab catalog empty. Run setupLabDatabase() then seedStarterCatalog().");
    }
  } catch (e) {
    out.push("getOrderableTests() THREW: " + e.message +
             "  (LabSetup.gs may not be first in file order)");
  }

  Logger.log(out.join("\n"));
  return out.join("\n");
}

/**
 * FRONTEND ENTRY. Proper lab catalog for the OPD order modal.
 * Returns real catalog IDs so the OPD order routes by TestID rather than by
 * name-matching, which is what currently forces createLabRequest into its
 * 'MANUAL_MAP' fallback and loses the test identity.
 */
function getOPDOrderableTests() {
  try {
    var res = getOrderableTests();
    if (!res || !res.success) {
      return { success: false, message: (res && res.message) || 'Catalog unavailable.',
               groups: [] };
    }

    var byDept = {};
    (res.tests || []).forEach(function (t) {
      var d = String(t.department || 'OTHER');
      if (!byDept[d]) byDept[d] = [];
      byDept[d].push({
        testId: t.testId,
        testCode: t.testCode,
        testName: t.testName,
        testType: t.testType,
        sampleType: t.sampleType,
        price: t.price,
        tatMinutes: t.tatMinutes,
        requiresConsent: t.requiresConsent
      });
    });

    var groups = Object.keys(byDept).sort().map(function (d) {
      return {
        department: d,
        label: d.replace(/_/g, ' ').toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); }),
        tests: byDept[d]
      };
    });

    return { success: true, groups: groups, total: (res.tests || []).length };
  } catch (e) {
    return { success: false, message: 'getOPDOrderableTests: ' + e.message, groups: [] };
  }
}