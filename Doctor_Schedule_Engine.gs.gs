// ============================================================================
// Doctor_Schedule_Engine.gs  —  Crescentia HealthTech
// PHASE 2 : Per-doctor availability, generated slots, doctor-scoped booking
// ----------------------------------------------------------------------------
// REQUIRES Doctor_Core.gs (dc_* helpers, resolveScope_, resolveWriteDoctor_)
// REPLACES the stored-block model in saveEnterpriseAvailability().
// Slots are COMPUTED from a weekly template + sparse exceptions. Nothing is
// written to Appointments except real bookings.
// ============================================================================

var DS_WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday",
                   "Thursday", "Friday", "Saturday"];

function ds_scheduleSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Doctor_Schedules", [
    "Schedule_ID", "Tenant_ID", "Doctor_ID", "Weekday", "Session_Label",
    "Start_Time", "End_Time", "Slot_Minutes", "Max_Per_Slot", "Location", "Status"
  ]);
}

function ds_exceptionSheet_() {
  return dc_ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), "Doctor_Schedule_Exceptions", [
    "Exception_ID", "Tenant_ID", "Doctor_ID", "Date", "Type",
    "Start_Time", "End_Time", "Reason", "Created_By", "Created_At"
  ]);
}

/** "yyyy-MM-dd" -> local Date at midnight. Avoids UTC parsing drift. */
function ds_parseDateKey_(dateStr) {
  var s = dc_str_(dateStr);
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
}

// ============================================================================
// 1. WEEKLY TEMPLATE — READ / WRITE
// ============================================================================

/**
 * FRONTEND ENTRY. Returns the weekly recurring template for one doctor.
 * @return {{success, doctorId, doctorName, rows:[...] , message}}
 */
function getDoctorSchedule(doctorId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) return { success: false, message: scope.message, rows: [] };

    var target = dc_str_(doctorId) || scope.selfDoctorId;
    if (!target) return { success: false, message: "Select a doctor.", rows: [] };
    if (!dc_inScope_(scope, target)) {
      return { success: false, message: "You can only view your own schedule.", rows: [] };
    }

    var doc = dc_getDoctorById_(target);
    var sh = ds_scheduleSheet_();
    var data = dc_sheetValues_(sh);
    var tenant = getTenantId_();
    var rows = [];

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][2]) !== dc_upper_(target)) continue;
      if (dc_str_(data[i][1]) !== tenant) continue;
      if (dc_upper_(data[i][10]) !== "ACTIVE") continue;
      rows.push({
        scheduleId:  dc_str_(data[i][0]),
        weekday:     dc_int_(data[i][3]),
        weekdayName: DS_WEEKDAYS[dc_int_(data[i][3])] || "",
        sessionLabel: dc_str_(data[i][4]),
        startTime:   dc_to24_(data[i][5]),
        endTime:     dc_to24_(data[i][6]),
        slotMinutes: dc_int_(data[i][7]) || 15,
        maxPerSlot:  dc_int_(data[i][8]) || 1,
        location:    dc_str_(data[i][9])
      });
    }
    rows.sort(function (a, b) {
      return (a.weekday - b.weekday) || (dc_minutes_(a.startTime) - dc_minutes_(b.startTime));
    });

    return {
      success: true,
      doctorId: target,
      doctorName: doc ? doc.name : target,
      rows: rows,
      message: rows.length ? "" : "No recurring clinic hours set for this doctor yet."
    };
  } catch (e) {
    return { success: false, message: "Schedule unavailable: " + e.message, rows: [] };
  }
}

/**
 * Replaces the whole weekly template for one doctor in a single transaction.
 * payload = { doctorId, rows: [{weekday, sessionLabel, startTime, endTime,
 *                               slotMinutes, maxPerSlot, location}] }
 */
function saveDoctorSchedule(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var w = resolveWriteDoctor_(sessionToken, payload && payload.doctorId);
    if (!w.ok) return { success: false, message: w.message };

    var incoming = (payload && payload.rows) || [];
    var tenant = getTenantId_();

    // --- validate entirely before touching the sheet -----------------------
    var clean = [];
    for (var i = 0; i < incoming.length; i++) {
      var r = incoming[i];
      var wd = dc_int_(r.weekday);
      var st = dc_to24_(r.startTime);
      var en = dc_to24_(r.endTime);
      var mins = dc_int_(r.slotMinutes) || 15;

      if (wd < 0 || wd > 6) return { success: false, message: "Invalid weekday in row " + (i + 1) + "." };
      if (!st || !en)       return { success: false, message: "Start and end time are required in row " + (i + 1) + "." };
      if (dc_minutes_(en) <= dc_minutes_(st)) {
        return { success: false, message: "End time must be after start time (" + DS_WEEKDAYS[wd] + ")." };
      }
      if (mins < 5 || mins > 120) return { success: false, message: "Slot length must be between 5 and 120 minutes." };

      clean.push([
        String(dc_str_(r.scheduleId) || "SCH-" + Utilities.getUuid().substring(0, 8).toUpperCase()),
        String(tenant),
        String(w.doctorId),
        dc_int_(wd),
        String(dc_str_(r.sessionLabel) || "CLINIC"),
        String(st),
        String(en),
        dc_int_(mins),
        dc_int_(r.maxPerSlot) || 1,
        String(dc_str_(r.location)),
        "ACTIVE"
      ]);
    }

    // --- overlap check within the same weekday -----------------------------
    for (var a = 0; a < clean.length; a++) {
      for (var b = a + 1; b < clean.length; b++) {
        if (clean[a][3] !== clean[b][3]) continue;
        var aS = dc_minutes_(clean[a][5]), aE = dc_minutes_(clean[a][6]);
        var bS = dc_minutes_(clean[b][5]), bE = dc_minutes_(clean[b][6]);
        if (aS < bE && bS < aE) {
          return { success: false, message: "Overlapping sessions on " + DS_WEEKDAYS[clean[a][3]] + "." };
        }
      }
    }

    // --- swap: delete this doctor's rows, append the new set ---------------
    var sh = ds_scheduleSheet_();
    var data = dc_sheetValues_(sh);
    for (var d = data.length - 1; d >= 1; d--) {
      if (dc_upper_(data[d][2]) === dc_upper_(w.doctorId) && dc_str_(data[d][1]) === tenant) {
        sh.deleteRow(d + 1);
      }
    }
    if (clean.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, clean.length, clean[0].length).setValues(clean);
      dc_invalidate_(sh.getName());
    }

    SpreadsheetApp.flush();
    logAudit_(w.sess, "SCHEDULE_SAVE", "Doctor", w.doctorId, { sessions: clean.length });
    return { success: true, message: "Clinic hours saved for " + w.name + "." };

  } catch (e) {
    return { success: false, message: "Could not save schedule: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** Seeds the legacy 10:00–13:00 / 17:00–21:00 pattern, Mon–Sat, 15-min slots. */
function seedDefaultScheduleForDoctor(doctorId) {
  var rows = [];
  for (var wd = 1; wd <= 6; wd++) {
    rows.push({ weekday: wd, sessionLabel: "MORNING", startTime: "10:00",
                endTime: "13:00", slotMinutes: 15, maxPerSlot: 1, location: "OPD" });
    rows.push({ weekday: wd, sessionLabel: "EVENING", startTime: "17:00",
                endTime: "21:00", slotMinutes: 15, maxPerSlot: 1, location: "OPD" });
  }
  var token = issueSession_({ username: "SYSTEM_SEED", role: "admin",
                              doctorId: doctorId, name: "System" });
  return saveDoctorSchedule({ doctorId: doctorId, rows: rows }, token);
}

// ============================================================================
// 2. EXCEPTIONS — leave, one-off blocks, extra clinics
// ============================================================================

/**
 * payload = { doctorId, startDate, endDate, type:'LEAVE'|'BLOCK'|'EXTRA',
 *             startTime, endTime, reason }
 * LEAVE with no times = whole day off.
 */
function addScheduleException(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var w = resolveWriteDoctor_(sessionToken, payload && payload.doctorId);
    if (!w.ok) return { success: false, message: w.message };

    var type = dc_upper_(payload.type) || "BLOCK";
    if (["LEAVE", "BLOCK", "EXTRA"].indexOf(type) === -1) {
      return { success: false, message: "Invalid exception type." };
    }

    var start = ds_parseDateKey_(payload.startDate);
    var end   = ds_parseDateKey_(payload.endDate || payload.startDate);
    if (!start || !end) return { success: false, message: "Select a valid date." };
    if (end < start)    return { success: false, message: "End date cannot be before start date." };

    var st = dc_to24_(payload.startTime);
    var en = dc_to24_(payload.endTime);
    if ((st && !en) || (!st && en)) {
      return { success: false, message: "Provide both a start and an end time, or neither." };
    }
    if (st && en && dc_minutes_(en) <= dc_minutes_(st)) {
      return { success: false, message: "End time must be after start time." };
    }
    if (type === "EXTRA" && (!st || !en)) {
      return { success: false, message: "An extra clinic needs a start and end time." };
    }

    // Build the date list with a hard ceiling (no runaway loops).
    var dates = [], cursor = new Date(start.getTime()), guard = 0;
    while (cursor <= end && guard < 366) {
      dates.push(Utilities.formatDate(cursor, Session.getScriptTimeZone(), "yyyy-MM-dd"));
      cursor.setDate(cursor.getDate() + 1);
      guard++;
    }
    if (guard >= 366) return { success: false, message: "Range cannot exceed one year." };

    var now = new Date();
    var tenant = getTenantId_();
    var newRows = dates.map(function (dStr) {
      return [
        String("EXC-" + Utilities.getUuid().substring(0, 8).toUpperCase()),
        String(tenant),
        String(w.doctorId),
        String(dStr),
        String(type),
        String(st),
        String(en),
        String(dc_str_(payload.reason)),
        String(w.sess.username),
        now
      ];
    });

    var sh = ds_exceptionSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    dc_invalidate_(sh.getName());
    SpreadsheetApp.flush();

    logAudit_(w.sess, "SCHEDULE_EXCEPTION_ADD", "Doctor", w.doctorId,
              { type: type, from: dates[0], to: dates[dates.length - 1] });

    return { success: true,
             message: type + " recorded for " + w.name + " across " + dates.length + " day(s)." };

  } catch (e) {
    return { success: false, message: "Could not save exception: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/** Exceptions for one doctor between two dates (inclusive). */
function getScheduleExceptions(doctorId, fromDate, toDate, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) return { success: false, message: scope.message, rows: [] };

    var target = dc_str_(doctorId) || scope.selfDoctorId;
    if (!dc_inScope_(scope, target)) {
      return { success: false, message: "You can only view your own schedule.", rows: [] };
    }

    var from = dc_dateKey_(fromDate), to = dc_dateKey_(toDate || fromDate);
    var sh = ds_exceptionSheet_();
    var data = dc_sheetValues_(sh);
    var rows = [];

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][2]) !== dc_upper_(target)) continue;
      var dk = dc_dateKey_(data[i][3]);
      if (from && dk < from) continue;
      if (to && dk > to) continue;
      rows.push({
        exceptionId: dc_str_(data[i][0]),
        date: dk,
        type: dc_upper_(data[i][4]),
        startTime: dc_to24_(data[i][5]),
        endTime: dc_to24_(data[i][6]),
        reason: dc_str_(data[i][7])
      });
    }
    rows.sort(function (a, b) { return a.date < b.date ? -1 : 1; });
    return { success: true, rows: rows,
             message: rows.length ? "" : "No leave or blocks recorded for this period." };
  } catch (e) {
    return { success: false, message: "Could not load exceptions: " + e.message, rows: [] };
  }
}

function deleteScheduleException(exceptionId, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired." };

    var sh = ds_exceptionSheet_();
    var data = dc_sheetValues_(sh);
    var target = dc_upper_(exceptionId);

    for (var i = 1; i < data.length; i++) {
      if (dc_upper_(data[i][0]) !== target) continue;
      var owner = dc_str_(data[i][2]);
      var w = resolveWriteDoctor_(sessionToken, owner);
      if (!w.ok) return { success: false, message: w.message };
      sh.deleteRow(i + 1);
      SpreadsheetApp.flush();
      logAudit_(sess, "SCHEDULE_EXCEPTION_DELETE", "Doctor", owner, { exceptionId: exceptionId });
      return { success: true, message: "Entry removed. Slots restored." };
    }
    return { success: false, message: "Entry not found." };
  } catch (e) {
    return { success: false, message: "Could not remove entry: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 3. SLOT GENERATION  (the replacement for stored 'Blocked' rows)
// ============================================================================

/** Internal: computes raw slot times for a doctor on a date. No booking data. */
function ds_generateSlots_(doctorId, dateKey) {
  var dateObj = ds_parseDateKey_(dateKey);
  if (!dateObj) return [];
  var weekday = dateObj.getDay();
  var tenant = getTenantId_();
  var docU = dc_upper_(doctorId);

  // --- template windows for this weekday ---------------------------------
  var schedData = ds_scheduleSheet_().getDataRange().getDisplayValues();
  var windows = [];
  for (var i = 1; i < schedData.length; i++) {
    if (dc_upper_(schedData[i][2]) !== docU) continue;
    if (dc_str_(schedData[i][1]) !== tenant) continue;
    if (dc_int_(schedData[i][3]) !== weekday) continue;
    if (dc_upper_(schedData[i][10]) !== "ACTIVE") continue;
    windows.push({
      label: dc_str_(schedData[i][4]),
      start: dc_minutes_(schedData[i][5]),
      end:   dc_minutes_(schedData[i][6]),
      step:  dc_int_(schedData[i][7]) || 15,
      max:   dc_int_(schedData[i][8]) || 1
    });
  }

  // --- exceptions for this exact date ------------------------------------
  var excData = ds_exceptionSheet_().getDataRange().getDisplayValues();
  var blocks = [], fullDayOff = false;
  for (var e = 1; e < excData.length; e++) {
    if (dc_upper_(excData[e][2]) !== docU) continue;
    if (dc_dateKey_(excData[e][3]) !== dateKey) continue;
    var type = dc_upper_(excData[e][4]);
    var s = dc_minutes_(excData[e][5]);
    var n = dc_minutes_(excData[e][6]);

    if (type === "LEAVE" && s < 0) { fullDayOff = true; break; }
    if (type === "LEAVE" || type === "BLOCK") {
      blocks.push({ start: (s < 0 ? 0 : s), end: (n < 0 ? 1440 : n) });
    } else if (type === "EXTRA" && s >= 0 && n > s) {
      windows.push({ label: "EXTRA", start: s, end: n, step: 15, max: 1 });
    }
  }
  if (fullDayOff) return [];

  // --- expand windows into slots, skipping blocked ranges ----------------
  var seen = {}, slots = [];
  windows.forEach(function (win) {
    if (win.start < 0 || win.end <= win.start) return;
    for (var t = win.start; t + win.step <= win.end; t += win.step) {
      var blocked = blocks.some(function (b) { return t >= b.start && t < b.end; });
      if (blocked) continue;
      var key = dc_fromMinutes_(t);
      if (seen[key]) continue;
      seen[key] = true;
      slots.push({ time24: key, time12: dc_to12_(key),
                   sessionLabel: win.label, maxPerSlot: win.max });
    }
  });

  slots.sort(function (a, b) { return dc_minutes_(a.time24) - dc_minutes_(b.time24); });
  return slots;
}

/** Internal: bookings for a doctor on a date, keyed by "HH:mm". */
function ds_bookingsByTime_(doctorId, dateKey) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appointments");
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;

  var data = dc_sheetValues_(sh);
  var m = dc_headerMap_(sh);
  var docCol = (m["Doctor_ID"] === undefined) ? -1 : m["Doctor_ID"];
  var docU = dc_upper_(doctorId);

  for (var i = 1; i < data.length; i++) {
    // Cheap string checks FIRST. The legacy 'Blocked' rows are the bulk of this
    // sheet; parsing their dates before discarding them is what made this slow.
    var status = data[i][6];
    if (status === "Cancelled" || status === "DELETE" || status === "Blocked") continue;
    if (data[i][1] === "ADMIN") continue;

    if (dc_dateKey_(data[i][3]) !== dateKey) continue;
    status = dc_str_(status);

    // Legacy rows with no Doctor_ID belong to the original single doctor.
    var rowDoc = (docCol === -1) ? DC_DEFAULT_DOCTOR
                                 : (dc_str_(data[i][docCol]) || DC_DEFAULT_DOCTOR);
    if (dc_upper_(rowDoc) !== docU) continue;

    var key = dc_to24_(data[i][4]);
    if (!key) continue;
    if (!out[key]) out[key] = [];
    out[key].push({
      apptId: dc_str_(data[i][0]),
      patientId: dc_str_(data[i][1]),
      patientName: dc_str_(data[i][2]),
      purpose: dc_str_(data[i][5]),
      status: status
    });
  }
  return out;
}

/**
 * FRONTEND ENTRY. One doctor's slot grid for one date, with booking state.
 * Drives both the availability screen and the booking modal.
 */
function getDoctorSlotsForDate(doctorId, dateStr, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) return { success: false, message: scope.message, slots: [] };

    var target = dc_str_(doctorId) || scope.selfDoctorId;
    if (!target) return { success: false, message: "Select a doctor.", slots: [] };
    if (!dc_inScope_(scope, target)) {
      return { success: false, message: "You can only view your own schedule.", slots: [] };
    }

    var dateKey = dc_dateKey_(dateStr);
    if (!dateKey) return { success: false, message: "Select a valid date.", slots: [] };

    var doc = dc_getDoctorById_(target);
    var generated = ds_generateSlots_(target, dateKey);
    var booked = ds_bookingsByTime_(target, dateKey);

    var slots = generated.map(function (s) {
      var b = booked[s.time24] || [];
      return {
        time24: s.time24,
        time12: s.time12,
        sessionLabel: s.sessionLabel,
        available: b.length < s.maxPerSlot,
        bookedCount: b.length,
        maxPerSlot: s.maxPerSlot,
        bookings: b
      };
    });

    // Bookings that fall outside the current template (emergencies, template
    // changed after booking). Surfaced, never silently dropped.
    var orphans = [];
    var covered = {};
    generated.forEach(function (s) { covered[s.time24] = true; });
    Object.keys(booked).forEach(function (t) {
      if (covered[t]) return;
      booked[t].forEach(function (b) {
        orphans.push({ time24: t, time12: dc_to12_(t), booking: b });
      });
    });
    orphans.sort(function (a, b) { return dc_minutes_(a.time24) - dc_minutes_(b.time24); });

    return {
      success: true,
      doctorId: target,
      doctorName: doc ? doc.name : target,
      date: dateKey,
      weekdayName: DS_WEEKDAYS[ds_parseDateKey_(dateKey).getDay()],
      slots: slots,
      offSchedule: orphans,
      message: slots.length ? "" :
        (doc ? doc.name : target) + " has no clinic hours on this day."
    };
  } catch (e) {
    return { success: false, message: "Could not load slots: " + e.message, slots: [] };
  }
}

/** FRONTEND ENTRY. Day grid across every doctor in scope (reception view). */
function getClinicDayGrid(dateStr, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, null);
    if (!scope.ok) return { success: false, message: scope.message, columns: [] };

    var dateKey = dc_dateKey_(dateStr);
    if (!dateKey) return { success: false, message: "Select a valid date.", columns: [] };

    var all = getActiveDoctors();
    var visible = (scope.mode === "ALL") ? all : all.filter(function (d) {
      return scope.doctorIds.indexOf(d.doctorId) !== -1;
    });

    var columns = visible.map(function (d) {
      var res = getDoctorSlotsForDate(d.doctorId, dateKey, sessionToken);
      return {
        doctorId: d.doctorId,
        doctorName: d.name,
        specialty: d.specialty,
        slots: res.success ? res.slots : [],
        offSchedule: res.success ? res.offSchedule : []
      };
    }).filter(function (c) { return c.slots.length > 0 || c.offSchedule.length > 0; });

    return {
      success: true, date: dateKey, columns: columns,
      message: columns.length ? "" : "No doctor is scheduled on this date."
    };
  } catch (e) {
    return { success: false, message: "Could not load the day grid: " + e.message, columns: [] };
  }
}

// ============================================================================
// 4. DOCTOR-SCOPED BOOKING
// Uniqueness key is (Doctor_ID, Date, Time) — NOT (Date, Time).
// ============================================================================

/**
 * payload = { doctorId, patientId, patientName, date, time, purpose, fee, status }
 * `time` may be 12h or 24h; it is normalised before comparison.
 */
function bookAppointmentScoped(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sess = dc_validateSession_(sessionToken);
    if (!sess) return { success: false, message: "Your session has expired. Please sign in again." };

    var doc = dc_getDoctorById_(dc_str_(payload && payload.doctorId));
    if (!doc)  return { success: false, message: "Select a doctor before booking." };
    if (doc.status !== "ACTIVE") return { success: false, message: doc.name + " is not accepting bookings." };

    var dateKey = dc_dateKey_(payload.date);
    var time24  = dc_to24_(payload.time);
    var patientId = dc_upper_(payload.patientId);
    if (!dateKey)   return { success: false, message: "Select a valid date." };
    if (!time24)    return { success: false, message: "Select a valid time." };
    if (!patientId) return { success: false, message: "Patient ID is required." };

    // --- slot must exist in the doctor's generated grid -------------------
    var grid = ds_generateSlots_(doc.doctorId, dateKey);
    var slot = null;
    for (var g = 0; g < grid.length; g++) {
      if (grid[g].time24 === time24) { slot = grid[g]; break; }
    }
    var isOverride = (payload.forceOffSchedule === true);
    if (!slot && !isOverride) {
      return { success: false,
               message: doc.name + " is not available at " + dc_to12_(time24) + " on " + dateKey + "." };
    }
    var capacity = slot ? slot.maxPerSlot : 1;

    // --- capacity + duplicate checks --------------------------------------
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appointments");
    if (!sh) return { success: false, message: "Appointments sheet is missing." };
    var m = dc_headerMap_(sh);
    var docCol = (m["Doctor_ID"] === undefined) ? -1 : m["Doctor_ID"];
    if (docCol === -1) {
      return { success: false, message: "Run runMultiDoctorMigration() before booking." };
    }

    var data = dc_sheetValues_(sh);
    var liveStatuses = ["Booked", "Arrived", "In-Progress"];
    var sameSlot = 0;

    for (var i = 1; i < data.length; i++) {
      if (dc_dateKey_(data[i][3]) !== dateKey) continue;
      var status = dc_str_(data[i][6]);
      if (status === "Cancelled" || status === "DELETE" || status === "Blocked") continue;

      var rowDoc = dc_upper_(dc_str_(data[i][docCol]) || DC_DEFAULT_DOCTOR);
      var rowTime = dc_to24_(data[i][4]);

      if (rowDoc === dc_upper_(doc.doctorId) && rowTime === time24) sameSlot++;

      if (rowDoc === dc_upper_(doc.doctorId) &&
          dc_upper_(data[i][1]) === patientId &&
          liveStatuses.indexOf(status) !== -1) {
        return { success: false,
                 message: "This patient already has an active appointment with " +
                          doc.name + " on " + dateKey + "." };
      }
    }

    if (sameSlot >= capacity && !isOverride) {
      return { success: false,
               message: "That slot was just taken. Please pick another time." };
    }

    // --- write ------------------------------------------------------------
    var newId = "APT-" + Utilities.getUuid().substring(0, 6).toUpperCase();
    var row = new Array(sh.getLastColumn()).fill("");
    row[0] = String(newId);
    row[1] = String(patientId);
    row[2] = String(dc_str_(payload.patientName));
    row[3] = String(dateKey);
    row[4] = String(dc_to12_(time24));                    // legacy display format
    row[5] = String(dc_str_(payload.purpose) || "Consultation");
    row[6] = String(dc_str_(payload.status) || "Booked");
    row[7] = dc_money_(payload.fee !== undefined ? payload.fee : doc.consultFee);
    row[8] = new Date();
    row[docCol] = String(doc.doctorId);
    if (m["Doctor_Name_Snapshot"] !== undefined) row[m["Doctor_Name_Snapshot"]] = String(doc.name);
    if (m["Booked_By"] !== undefined)            row[m["Booked_By"]] = String(sess.username);
    if (m["Attribution_Source"] !== undefined)   row[m["Attribution_Source"]] = "RECORDED";

    sh.appendRow(row);
    dc_invalidate_(sh.getName());
    SpreadsheetApp.flush();

    logAudit_(sess, "APPOINTMENT_BOOK", "Appointment", newId,
              { doctorId: doc.doctorId, patientId: patientId,
                date: dateKey, time: time24, offSchedule: !slot });

    return { success: true, apptId: newId,
             message: "Booked with " + doc.name + " at " + dc_to12_(time24) + "." };

  } catch (e) {
    return { success: false, message: "Booking failed: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 5. SCOPED DAILY LEDGER  (drop-in replacement for fetchDailyLedger)
// ============================================================================

/**
 * FRONTEND ENTRY. Appointment ledger for a date, filtered to the caller's scope.
 * Pass doctorId = "" for "all doctors I'm allowed to see".
 */
function fetchDailyLedgerScoped(dateStr, doctorId, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) return { success: false, message: scope.message, ledger: [] };

    var dateKey = dc_dateKey_(dateStr);
    if (!dateKey) return { success: false, message: "Select a valid date.", ledger: [] };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var apptSheet = ss.getSheetByName("Appointments");
    var patSheet  = ss.getSheetByName("Patients");
    if (!apptSheet) return { success: false, message: "Appointments sheet is missing.", ledger: [] };

    var m = dc_headerMap_(apptSheet);
    var docCol  = (m["Doctor_ID"] === undefined) ? -1 : m["Doctor_ID"];
    var snapCol = (m["Doctor_Name_Snapshot"] === undefined) ? -1 : m["Doctor_Name_Snapshot"];

    var data = dc_sheetValues_(apptSheet);

    // Which patients actually appear on this date? Building the demographics
    // map for the WHOLE Patients sheet, to decorate the dozen rows on one
    // day's ledger, is most of what made this screen slow.
    var needed = {};
    for (var q = 1; q < data.length; q++) {
      if (dc_dateKey_(data[q][3]) !== dateKey) continue;
      var qp = dc_upper_(data[q][1]);
      if (qp && qp !== "ADMIN") needed[qp] = true;
    }

    var patMap = {};
    if (patSheet && patSheet.getLastRow() > 1) {
      var pd = dc_sheetValues_(patSheet);
      for (var p = 1; p < pd.length; p++) {
        var pk = dc_upper_(pd[p][0]);
        if (!needed[pk]) continue;
        patMap[pk] = { age: dc_str_(pd[p][3]), sex: dc_str_(pd[p][4]) };
      }
    }

    var ledger = [];

    for (var i = 1; i < data.length; i++) {
      if (dc_dateKey_(data[i][3]) !== dateKey) continue;

      var status = dc_str_(data[i][6]);
      if (status === "Blocked" || status === "DELETE") continue;

      var pid = dc_upper_(data[i][1]);
      if (pid === "ADMIN") continue;

      var rowDoc = (docCol === -1) ? DC_DEFAULT_DOCTOR
                                   : (dc_str_(data[i][docCol]) || DC_DEFAULT_DOCTOR);
      if (!dc_inScope_(scope, rowDoc)) continue;

      var docName = (snapCol !== -1) ? dc_str_(data[i][snapCol]) : "";
      if (!docName) {
        var dRec = dc_getDoctorById_(rowDoc);
        docName = dRec ? dRec.name : rowDoc;
      }

      ledger.push({
        apptId: dc_str_(data[i][0]),
        time12: dc_to12_(data[i][4]),
        time24: dc_to24_(data[i][4]),
        patientId: pid,
        patientName: dc_str_(data[i][2]),
        age: patMap[pid] ? patMap[pid].age : "",
        sex: patMap[pid] ? patMap[pid].sex : "",
        purpose: dc_str_(data[i][5]),
        status: status,
        fee: dc_money_(data[i][7]),
        doctorId: rowDoc,
        doctorName: docName
      });
    }

    ledger.sort(function (a, b) {
      return (dc_minutes_(a.time24) - dc_minutes_(b.time24)) ||
             (a.doctorName < b.doctorName ? -1 : 1);
    });

    return {
      success: true, date: dateKey, scopeMode: scope.mode, ledger: ledger,
      message: ledger.length ? "" : "No appointments recorded for this date."
    };
  } catch (e) {
    return { success: false, message: "Could not load the ledger: " + e.message, ledger: [] };
  }
}