// ============================================================================
// Doctor_Availability_Blocks.gs  —  Crescentia HealthTech
// Powers the classic slot-toggle availability grid on the per-doctor schedule
// engine. Toggling a slot OFF writes a BLOCK exception; toggling it back ON
// removes that exception. The weekly template is never touched here.
// REQUIRES: Doctor_Core.gs, Doctor_Schedule_Engine.gs
// ----------------------------------------------------------------------------
// PERFORMANCE NOTE
//   The previous version read Doctor_Schedules three times and
//   Doctor_Schedule_Exceptions twice per request, then scanned the whole
//   Appointments sheet on top. Combined with a Utilities.formatDate call per
//   row it exceeded the 6-minute execution limit, which the HTML service
//   surfaces as a call that never returns (an eternal spinner).
//   This version reads each sheet exactly once.
// ============================================================================

/**
 * FRONTEND ENTRY. Generated grid for a doctor/date, already grouped into hour
 * blocks and day periods so the UI renders without re-deriving anything.
 */
function getDoctorAvailabilityGrid(doctorId, dateStr, sessionToken) {
  try {
    var scope = resolveScope_(sessionToken, doctorId);
    if (!scope.ok) return { success: false, message: scope.message, periods: [] };

    var target = dc_str_(doctorId) || scope.selfDoctorId;
    if (!target) return { success: false, message: "Select a doctor.", periods: [] };
    if (!dc_inScope_(scope, target)) {
      return { success: false, message: "You can only manage your own schedule.", periods: [] };
    }

    var dateKey = dc_dateKey_(dateStr);
    if (!dateKey) return { success: false, message: "Select a valid date.", periods: [] };

    var dateObj = ds_parseDateKey_(dateKey);
    if (!dateObj) return { success: false, message: "Select a valid date.", periods: [] };

    var doc = dc_getDoctorById_(target);
    var docName = doc ? doc.name : target;
    var weekday = dateObj.getDay();
    var tenant  = getTenantId_();
    var docU    = dc_upper_(target);

    // ---- READ 1: weekly template (this doctor, this weekday) -------------
    var schedData = ds_scheduleSheet_().getDataRange().getDisplayValues();
    var windows = [];
    var step = 0;

    for (var i = 1; i < schedData.length; i++) {
      if (dc_upper_(schedData[i][2]) !== docU) continue;
      if (dc_str_(schedData[i][1]) !== tenant) continue;
      if (dc_int_(schedData[i][3]) !== weekday) continue;
      if (dc_upper_(schedData[i][10]) !== "ACTIVE") continue;

      var wStep = dc_int_(schedData[i][7]) || 15;
      if (!step || wStep < step) step = wStep;

      windows.push({
        label: dc_str_(schedData[i][4]),
        start: dc_minutes_(schedData[i][5]),
        end:   dc_minutes_(schedData[i][6]),
        step:  wStep,
        max:   dc_int_(schedData[i][8]) || 1
      });
    }
    if (!step) step = 15;

    // ---- READ 2: exceptions for this exact date --------------------------
    var excData = ds_exceptionSheet_().getDataRange().getDisplayValues();
    var blocks = [], fullDayOff = false, leaveReason = "";

    for (var e = 1; e < excData.length; e++) {
      if (dc_upper_(excData[e][2]) !== docU) continue;          // cheap first
      if (dc_dateKey_(excData[e][3]) !== dateKey) continue;

      var type = dc_upper_(excData[e][4]);
      var s = dc_minutes_(excData[e][5]);
      var n = dc_minutes_(excData[e][6]);

      if (type === "LEAVE" && s < 0) {
        fullDayOff = true;
        leaveReason = dc_str_(excData[e][7]);
        break;
      }
      if (type === "LEAVE" || type === "BLOCK") {
        blocks.push({ start: (s < 0 ? 0 : s), end: (n < 0 ? 1440 : n) });
      } else if (type === "EXTRA" && s >= 0 && n > s) {
        windows.push({ label: "EXTRA", start: s, end: n, step: step, max: 1 });
      }
    }

    if (fullDayOff) {
      return {
        success: true, doctorId: target, doctorName: docName, date: dateKey,
        weekdayName: DS_WEEKDAYS[weekday], slotMinutes: step,
        periods: [], offSchedule: [], fullDayOff: true,
        message: docName + " is on leave this day" +
                 (leaveReason ? " (" + leaveReason + ")" : "") + "."
      };
    }

    // ---- expand the template into every slot, blocked or not -------------
    // Blocked slots must still appear (toggled OFF), otherwise the user could
    // never un-block anything.
    var seen = {}, all = [];
    windows.forEach(function (win) {
      if (win.start < 0 || win.end <= win.start) return;
      for (var t = win.start; t + win.step <= win.end; t += win.step) {
        var key = dc_fromMinutes_(t);
        if (seen[key]) continue;
        seen[key] = true;

        var isBlocked = false;
        for (var b = 0; b < blocks.length; b++) {
          if (t >= blocks[b].start && t < blocks[b].end) { isBlocked = true; break; }
        }

        all.push({
          time24: key, time12: dc_to12_(key), minutes: t,
          available: !isBlocked, booked: false, bookedBy: '',
          maxPerSlot: win.max
        });
      }
    });

    all.sort(function (a, b) { return a.minutes - b.minutes; });

    var byTime = {};
    all.forEach(function (s) { byTime[s.time24] = s; });

    // ---- READ 3: bookings, one scan of Appointments ----------------------
    var booked = ds_bookingsByTime_(target, dateKey);
    var offSchedule = [];

    Object.keys(booked).forEach(function (t) {
      var list = booked[t];
      if (byTime[t]) {
        byTime[t].booked = true;
        byTime[t].bookedBy = list.map(function (b) {
          return b.patientName || b.patientId;
        }).join(', ');
      } else {
        list.forEach(function (b) {
          offSchedule.push({ time24: t, time12: dc_to12_(t), booking: b });
        });
      }
    });
    offSchedule.sort(function (a, b) {
      return dc_minutes_(a.time24) - dc_minutes_(b.time24);
    });

    // ---- group into hour blocks, then day periods ------------------------
    var hours = {}, hourOrder = [];
    all.forEach(function (s) {
      var h = Math.floor(s.minutes / 60);
      if (!hours[h]) { hours[h] = []; hourOrder.push(h); }
      hours[h].push({
        time24: s.time24, time12: s.time12,
        available: s.available, booked: s.booked, bookedBy: s.bookedBy
      });
    });
    hourOrder.sort(function (a, b) { return a - b; });

    var buckets = { Morning: [], Afternoon: [], Evening: [] };
    hourOrder.forEach(function (h) {
      var period = (h < 12) ? "Morning" : (h < 17 ? "Afternoon" : "Evening");
      buckets[period].push({
        blockId: "hr" + h,
        label: dc_to12_(dc_fromMinutes_(h * 60)) + " to " +
               dc_to12_(dc_fromMinutes_(((h + 1) % 24) * 60)),
        slots: hours[h]
      });
    });

    var periods = ["Morning", "Afternoon", "Evening"]
      .filter(function (p) { return buckets[p].length > 0; })
      .map(function (p) { return { title: p, blocks: buckets[p] }; });

    return {
      success: true,
      doctorId: target,
      doctorName: docName,
      date: dateKey,
      weekdayName: DS_WEEKDAYS[weekday],
      slotMinutes: step,
      periods: periods,
      offSchedule: offSchedule,
      fullDayOff: false,
      message: periods.length ? "" :
        docName + " has no clinic hours on " + DS_WEEKDAYS[weekday] +
        ". Set weekly hours first."
    };

  } catch (e) {
    return { success: false, message: "Could not load availability: " + e.message, periods: [] };
  }
}

/**
 * FRONTEND ENTRY. Replaces the BLOCK exceptions for a doctor across a date
 * range with exactly the set the user toggled off.
 *
 * payload = { doctorId, startDate, endDate, slotMinutes,
 *             blockedTimes: ["11:15","11:30"], overrideConflicts: bool }
 *
 * Returns { isConflict: true, conflicts: [...] } when a blocked slot already
 * holds a live booking and the caller has not confirmed the override.
 */
function setDoctorDayBlocks(payload, sessionToken) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var w = resolveWriteDoctor_(sessionToken, payload && payload.doctorId);
    if (!w.ok) return { success: false, message: w.message };

    var start = ds_parseDateKey_(dc_dateKey_(payload.startDate));
    var end   = ds_parseDateKey_(dc_dateKey_(payload.endDate || payload.startDate));
    if (!start || !end) return { success: false, message: "Select a valid date." };
    if (end < start)    return { success: false, message: "End date cannot be before start date." };

    var step = dc_int_(payload.slotMinutes) || 15;
    var times = (payload.blockedTimes || []).map(dc_to24_)
                  .filter(function (t) { return !!t; });

    // ---- date list, capped ----------------------------------------------
    var dates = [], cursor = new Date(start.getTime()), guard = 0;
    while (cursor <= end && guard < 366) {
      dates.push(dc_fmtDate_(cursor));       // local, no service call
      cursor.setDate(cursor.getDate() + 1);
      guard++;
    }
    if (guard >= 366) return { success: false, message: "Range cannot exceed one year." };

    // ---- conflict scan: never silently strand a booked patient ----------
    if (!payload.overrideConflicts && times.length > 0) {
      var conflicts = [];
      for (var d = 0; d < dates.length; d++) {
        var bookedMap = ds_bookingsByTime_(w.doctorId, dates[d]);
        for (var t = 0; t < times.length; t++) {
          var list = bookedMap[times[t]] || [];
          for (var k = 0; k < list.length; k++) {
            conflicts.push({ date: dates[d], time: dc_to12_(times[t]),
                             patient: list[k].patientName || list[k].patientId });
          }
        }
        if (conflicts.length > 50) break;    // enough to show the user
      }
      if (conflicts.length > 0) {
        return { success: false, isConflict: true, conflicts: conflicts.slice(0, 50),
                 message: conflicts.length + " booked slot(s) would be blocked." };
      }
    }

    // ---- swap the BLOCK rows for this doctor across the range ------------
    var sh = ds_exceptionSheet_();
    var data = sh.getDataRange().getDisplayValues();
    var inRange = {};
    dates.forEach(function (dk) { inRange[dk] = true; });

    // Collect rows to delete, then delete bottom-up in one pass.
    for (var i = data.length - 1; i >= 1; i--) {
      if (dc_upper_(data[i][4]) !== "BLOCK") continue;          // cheap first
      if (dc_upper_(data[i][2]) !== dc_upper_(w.doctorId)) continue;
      if (!inRange[dc_dateKey_(data[i][3])]) continue;
      sh.deleteRow(i + 1);                  // LEAVE / EXTRA untouched
    }

    if (times.length > 0) {
      var now = new Date();
      var tenant = getTenantId_();
      var rows = [];
      dates.forEach(function (dk) {
        times.forEach(function (t) {
          rows.push([
            String("EXC-" + Utilities.getUuid().substring(0, 8).toUpperCase()),
            String(tenant),
            String(w.doctorId),
            String(dk),
            "BLOCK",
            String(t),
            String(dc_fromMinutes_(dc_minutes_(t) + step)),
            String(dc_str_(payload.reason) || "Slot blocked"),
            String(w.sess.username),
            now
          ]);
        });
      });
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    SpreadsheetApp.flush();
    logAudit_(w.sess, "AVAILABILITY_SET", "Doctor", w.doctorId, {
      from: dates[0], to: dates[dates.length - 1],
      blocked: times.length, forced: payload.overrideConflicts === true
    });

    return {
      success: true,
      message: "Availability updated for " + w.name + " across " + dates.length + " day(s)."
    };

  } catch (e) {
    return { success: false, message: "Could not update availability: " + e.message };
  } finally {
    lock.releaseLock();
  }
}