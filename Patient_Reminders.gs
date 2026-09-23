// ============================================================================
// Patient_Reminders.gs — Crescentia HealthTech / CresRx
// Appointment, follow-up and vaccination reminders, by WhatsApp or SMS.
// ----------------------------------------------------------------------------
// WHO IS REMINDED OF WHAT (for a given day, normally tomorrow):
//   APPT     an appointment on that day that is still Booked
//   FOLLOWUP an OP consultation whose Next_Review_Date is that day, unless
//            the patient already has an appointment that day
//   VACCINE  a child under the immunisation card's age whose next dose
//            (Maternal_Child.gs) falls due that day
//
// CONSENT. A reminder is a message to the patient's phone through a third
// party (Meta for WhatsApp, the SMS gateway otherwise). It goes only to a
// patient whose COMMUNICATION consent is GIVEN in the DPDP register — the
// same gate as every other dispatch (DPDP_Dispatch.gs). Everyone else is
// listed with the reason, so the desk can ask at the next visit.
//
// HOW IT IS SENT. Two ways, chosen in Script Properties:
//
//   (no provider)   THE DESK SENDS. Admin Dashboard -> Operations ->
//                   Reminders lists tomorrow's, each with a WhatsApp and an
//                   SMS button that open the message ready to send from the
//                   clinic phone. Nothing leaves without a person clicking.
//
//   REMINDER_PROVIDER = WHATSAPP_CLOUD   WhatsApp Business Cloud API
//       WA_PHONE_NUMBER_ID, WA_ACCESS_TOKEN,
//       WA_TEMPLATE_NAME   an APPROVED template with four body parameters:
//                          {{1}} name  {{2}} what  {{3}} when  {{4}} clinic
//       WA_TEMPLATE_LANG   default "en"
//
//   REMINDER_PROVIDER = SMS_HTTP         any gateway with an HTTP GET API
//       SMS_URL_TEMPLATE   e.g. https://gateway.example/send?key=…&to={to}&text={text}
//                          {to} is 91XXXXXXXXXX, {text} is URL-encoded.
//       Indian SMS needs a DLT-registered template; the text below must be
//       registered as it is, or the gateway will drop it.
//
//   With a provider set, the daily job (remindersDaily, 6 pm) sends
//   tomorrow's on its own. Every attempt is written to Reminder_Log, and a
//   reminder already sent is never sent twice.
// ============================================================================

var REM_CFG = {
  LOG: 'Reminder_Log',
  HEADERS: ['Reminder_ID', 'Key', 'Kind', 'Patient_ID', 'Due_Date', 'Channel',
            'Status', 'Message', 'Logged_At', 'Logged_By', 'Error'],
  // Statuses that mean "this one is done": never send again.
  DONE: ['SENT', 'OPENED_BY_DESK']
};

var REM_KIND_LABEL = { APPT: 'Appointment', FOLLOWUP: 'Follow-up', VACCINE: 'Vaccination' };

function rem_str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function rem_tz_() { return Session.getScriptTimeZone() || 'Asia/Kolkata'; }
function rem_iso_(d) { return Utilities.formatDate(d, rem_tz_(), 'yyyy-MM-dd'); }
function rem_nice_(d) { return Utilities.formatDate(d, rem_tz_(), 'EEE, dd MMM'); }

function rem_date_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var d = (typeof cresc_parseDate_ === 'function') ? cresc_parseDate_(v) : new Date(v);
  return (d && !isNaN(d.getTime())) ? d : null;
}

/** 91XXXXXXXXXX, or '' if it is not an Indian mobile number. */
function rem_phone_(raw) {
  var d = rem_str_(raw).replace(/\D/g, '');
  if (d.length === 11 && d.charAt(0) === '0') d = d.slice(1);
  if (d.length === 10) d = '91' + d;
  return /^91[6-9]\d{9}$/.test(d) ? d : '';
}

/** Header index: exact name, else a header that starts with it. */
function rem_col_(hdr, names) {
  for (var i = 0; i < names.length; i++) {
    var k = hdr.indexOf(names[i]);
    if (k !== -1) return k;
  }
  for (var j = 0; j < names.length; j++) {
    for (var h = 0; h < hdr.length; h++) {
      if (hdr[h].indexOf(names[j]) === 0) return h;
    }
  }
  return -1;
}

function rem_values_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return null;
  var v = sh.getDataRange().getValues();
  return { hdr: v[0].map(rem_str_), rows: v.slice(1) };
}

function rem_logSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(REM_CFG.LOG);
  if (!sh) {
    sh = ss.insertSheet(REM_CFG.LOG);
    sh.appendRow(REM_CFG.HEADERS);
    sh.getRange(1, 1, 1, REM_CFG.HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Key -> last logged status. */
function rem_logged_() {
  var out = {};
  var v = rem_values_(REM_CFG.LOG);
  if (!v) return out;
  var k = v.hdr.indexOf('Key'), s = v.hdr.indexOf('Status'), c = v.hdr.indexOf('Channel'), t = v.hdr.indexOf('Logged_At');
  v.rows.forEach(function (r) {
    var key = rem_str_(r[k]);
    if (!key) return;
    var prev = out[key];
    // A DONE row is never overwritten by a later failure.
    if (prev && REM_CFG.DONE.indexOf(prev.status) !== -1) return;
    out[key] = { status: rem_str_(r[s]), channel: rem_str_(r[c]),
                 at: r[t] instanceof Date ? Utilities.formatDate(r[t], rem_tz_(), 'dd-MMM HH:mm') : rem_str_(r[t]) };
  });
  return out;
}

function rem_log_(item, channel, status, by, error) {
  try {
    rem_logSheet_().appendRow([
      'REM-' + Utilities.getUuid().slice(0, 8).toUpperCase(), item.key, item.kind, item.patientId,
      item.dueDate, channel, status, item.message, new Date(), by || 'SYSTEM', rem_str_(error).slice(0, 300)
    ]);
  } catch (e) { Logger.log('reminder log failed: ' + e.message); }
}

function rem_provider_() {
  var p = {};
  try { p = PropertiesService.getScriptProperties().getProperties() || {}; } catch (e) {}
  var kind = rem_str_(p.REMINDER_PROVIDER).toUpperCase();
  if (kind === 'WHATSAPP_CLOUD' && p.WA_PHONE_NUMBER_ID && p.WA_ACCESS_TOKEN && p.WA_TEMPLATE_NAME) {
    return { kind: kind, channel: 'WHATSAPP', label: 'WhatsApp Business (Meta)', p: p };
  }
  if (kind === 'SMS_HTTP' && p.SMS_URL_TEMPLATE) {
    return { kind: kind, channel: 'SMS', label: 'SMS gateway', p: p };
  }
  return null;
}

/**
 * The reminders due for one day. No permission check: callers have one.
 * @param {Date} day
 */
function rem_collect_(day) {
  var target = rem_iso_(day);
  var clinic = (typeof cresc_clinic_ === 'function') ? cresc_clinic_() : { name: 'the clinic', phone: '' };
  var clinicName = clinic.name || 'the clinic';
  var callLine = clinic.phone ? ' Call ' + clinic.phone + ' to change it.' : '';

  // Patients: id -> {name, phone, dob}
  var pts = {};
  var pv = rem_values_('Patients');
  if (pv) {
    var pc = { id: rem_col_(pv.hdr, ['Patient_ID']), name: rem_col_(pv.hdr, ['Name']),
               dob: rem_col_(pv.hdr, ['DOB']), mob: rem_col_(pv.hdr, ['Mobile']), wa: rem_col_(pv.hdr, ['WhatsApp']) };
    pv.rows.forEach(function (r) {
      var id = rem_str_(r[pc.id]).toUpperCase();
      if (!id) return;
      pts[id] = { name: rem_str_(r[pc.name]), dob: pc.dob >= 0 ? r[pc.dob] : '',
                  phone: rem_phone_(pc.wa >= 0 ? r[pc.wa] : '') || rem_phone_(pc.mob >= 0 ? r[pc.mob] : '') };
    });
  }

  var items = [];
  var bookedThatDay = {};
  var add = function (kind, pid, ref, name, what, when, message) {
    var p = pts[pid] || {};
    items.push({
      key: kind + '|' + ref + '|' + target,
      kind: kind, kindLabel: REM_KIND_LABEL[kind], patientId: pid,
      name: name || p.name || pid, phone: p.phone || '', dueDate: target,
      what: what, when: when, message: message
    });
  };

  // 1. Appointments.
  var av = rem_values_('Appointments');
  if (av) {
    var ac = { id: rem_col_(av.hdr, ['Appt ID', 'Appt_ID']), pid: rem_col_(av.hdr, ['Patient ID', 'Patient_ID']),
               name: rem_col_(av.hdr, ['Patient Name']), date: rem_col_(av.hdr, ['Date']),
               time: rem_col_(av.hdr, ['Time']), status: rem_col_(av.hdr, ['Status']),
               doc: rem_col_(av.hdr, ['Doctor_Name_Snapshot']) };
    av.rows.forEach(function (r) {
      var d = rem_date_(r[ac.date]);
      if (!d || rem_iso_(d) !== target) return;
      var pid = rem_str_(r[ac.pid]).toUpperCase();
      if (!pid) return;
      var st = rem_str_(r[ac.status]).toUpperCase();
      if (st && st !== 'BOOKED' && st !== 'SCHEDULED' && st !== 'CONFIRMED') return;
      bookedThatDay[pid] = true;
      var time = (typeof cresc_timeText_ === 'function') ? cresc_timeText_(r[ac.time]) : rem_str_(r[ac.time]);
      var doc = ac.doc >= 0 ? rem_str_(r[ac.doc]) : '';
      var name = rem_str_(r[ac.name]);
      var when = rem_nice_(d) + (time ? ' at ' + time : '');
      add('APPT', pid, rem_str_(r[ac.id]) || pid, name,
          'appointment' + (doc ? ' with ' + doc : ''), when,
          'Dear ' + (name || 'patient') + ', this is a reminder of your appointment' +
          (doc ? ' with ' + doc : '') + ' at ' + clinicName + ' on ' + when + '.' + callLine);
    });
  }

  // 2. Follow-ups advised at a consultation.
  var ov = rem_values_('OP_Encounters');
  if (ov) {
    var oc = { pid: rem_col_(ov.hdr, ['Patient_ID']), appt: rem_col_(ov.hdr, ['Appt_ID']),
               next: rem_col_(ov.hdr, ['Next_Review_Date']) };
    if (oc.next < 0 && ov.hdr.length > 25) oc.next = 25;     // column Z on older sheets
    var seen = {};
    if (oc.next >= 0) ov.rows.forEach(function (r) {
      var d = rem_date_(r[oc.next]);
      if (!d || rem_iso_(d) !== target) return;
      var pid = rem_str_(r[oc.pid]).toUpperCase();
      if (!pid || bookedThatDay[pid] || seen[pid]) return;
      seen[pid] = true;
      var p = pts[pid] || {};
      add('FOLLOWUP', pid, rem_str_(r[oc.appt]) || pid, p.name, 'follow-up review', rem_nice_(d),
          'Dear ' + (p.name || 'patient') + ', your doctor asked to see you again on ' + rem_nice_(d) +
          ' at ' + clinicName + '. Please book a time.' + (clinic.phone ? ' Call ' + clinic.phone + '.' : ''));
    });
  }

  // 3. Vaccinations for children on the card.
  if (typeof mc_immunisationView_ === 'function') {
    var maxDays = ((typeof MC !== 'undefined' && MC.IMMUN_MAX_AGE_YEARS) || 11) * 366;
    Object.keys(pts).forEach(function (pid) {
      var p = pts[pid];
      var dob = rem_date_(p.dob);
      if (!dob) return;
      var age = (day.getTime() - dob.getTime()) / 86400000;
      if (age < 0 || age > maxDays) return;
      var view = null;
      try { view = mc_immunisationView_(pid, dob); } catch (e) { view = null; }
      if (!view || !view.eligible || !view.nextDueGroup || !view.nextDueGroup.length) return;
      if (view.nextDueGroup[0].due !== target) return;
      var doses = view.nextDueGroup.map(function (x) { return x.label; }).join(', ');
      add('VACCINE', pid, pid + ':' + view.nextDueGroup[0].code, p.name, 'vaccination (' + doses + ')', rem_nice_(day),
          'Dear parent, ' + (p.name || 'your child') + '\'s vaccination (' + doses + ') is due on ' +
          rem_nice_(day) + ' at ' + clinicName + '. Please bring the immunisation card.' + callLine);
    });
  }

  // Consent and what has already gone.
  var logged = rem_logged_();
  var consentCache = {};
  items.forEach(function (it) {
    var c = consentCache[it.patientId];
    if (!c) {
      c = consentCache[it.patientId] = (typeof dpdp_consentState_ === 'function')
        ? dpdp_consentState_(it.patientId, 'COMMUNICATION') : { state: 'NOT_ASKED' };
    }
    it.consent = c.state;
    var l = logged[it.key];
    it.logged = l || null;
    if (l && REM_CFG.DONE.indexOf(l.status) !== -1) it.state = 'DONE';
    else if (c.state !== 'GIVEN') it.state = 'NO_CONSENT';
    else if (!it.phone) it.state = 'NO_PHONE';
    else it.state = 'READY';
    it.waLink = it.phone ? 'https://wa.me/' + it.phone + '?text=' + encodeURIComponent(it.message) : '';
    it.smsLink = it.phone ? 'sms:+' + it.phone + '?body=' + encodeURIComponent(it.message) : '';
  });
  var order = { APPT: 0, FOLLOWUP: 1, VACCINE: 2 };
  items.sort(function (a, b) { return order[a.kind] - order[b.kind] || a.name.localeCompare(b.name); });
  return items;
}

/** One message through the configured provider. Throws on failure. */
function rem_send_(provider, item) {
  if (provider.kind === 'WHATSAPP_CLOUD') {
    var p = provider.p;
    var res = UrlFetchApp.fetch('https://graph.facebook.com/v19.0/' + p.WA_PHONE_NUMBER_ID + '/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + p.WA_ACCESS_TOKEN },
      payload: JSON.stringify({
        messaging_product: 'whatsapp', to: item.phone, type: 'template',
        template: {
          name: p.WA_TEMPLATE_NAME, language: { code: p.WA_TEMPLATE_LANG || 'en' },
          components: [{ type: 'body', parameters: [item.name, item.what, item.when,
            ((typeof cresc_clinic_ === 'function' && cresc_clinic_().name) || 'the clinic')].map(function (t) {
              return { type: 'text', text: String(t).slice(0, 200) };
            }) }]
        }
      })
    });
    if (res.getResponseCode() >= 300) throw new Error('WhatsApp ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
    return;
  }
  if (provider.kind === 'SMS_HTTP') {
    var url = String(provider.p.SMS_URL_TEMPLATE)
      .replace('{to}', encodeURIComponent(item.phone))
      .replace('{text}', encodeURIComponent(item.message));
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (r.getResponseCode() >= 300) throw new Error('SMS ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
    return;
  }
  throw new Error('No reminder provider is configured.');
}

/** Sends every READY reminder for a day. @return {{sent, failed, skipped}} */
function rem_sendAll_(day, by) {
  var provider = rem_provider_();
  if (!provider) return { sent: 0, failed: 0, skipped: 0, provider: null };
  var out = { sent: 0, failed: 0, skipped: 0, provider: provider.label };
  rem_collect_(day).forEach(function (it) {
    if (it.state !== 'READY') { out.skipped++; return; }
    try {
      rem_send_(provider, it);
      rem_log_(it, provider.channel, 'SENT', by);
      out.sent++;
    } catch (e) {
      rem_log_(it, provider.channel, 'FAILED', by, e.message);
      out.failed++;
    }
  });
  return out;
}

function rem_dayFrom_(opt) {
  var d = new Date(); d.setHours(0, 0, 0, 0);
  var s = rem_str_(opt);
  if (!s || s === 'tomorrow') { d.setDate(d.getDate() + 1); return d; }
  if (s === 'today') return d;
  var p = rem_date_(s);
  if (!p) throw new Error('Unreadable date "' + s + '".');
  p.setHours(0, 0, 0, 0);
  return p;
}

/**
 * FRONTEND ENTRY. The reminders for a day (default tomorrow), with what has
 * been sent and why any cannot be.
 * @param {{date?:string}} opts   'today' | 'tomorrow' | yyyy-MM-dd
 */
function getReminderQueue(sessionToken, opts) {
  try {
    crescRequire_(sessionToken, 'appointment.read');
    var day = rem_dayFrom_((opts || {}).date);
    var items = rem_collect_(day);
    var provider = rem_provider_();
    var tally = { READY: 0, DONE: 0, NO_CONSENT: 0, NO_PHONE: 0 };
    items.forEach(function (x) { tally[x.state] = (tally[x.state] || 0) + 1; });
    return { success: true, date: rem_iso_(day), dateText: rem_nice_(day), items: items, tally: tally,
             provider: provider ? provider.label : '', auto: !!provider };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  }
}

/**
 * FRONTEND ENTRY. The desk opened this reminder in WhatsApp or SMS on the
 * clinic phone. Recorded so it is not sent twice; delivery cannot be seen.
 * The reminder is looked up again here — the browser only names it.
 */
function markReminderSent(sessionToken, key, channel) {
  try {
    var actor = crescRequire_(sessionToken, 'appointment.write');
    var date = rem_str_(key).split('|')[2] || '';
    var it = rem_collect_(rem_dayFrom_(date)).filter(function (x) { return x.key === key; })[0];
    if (!it) return { success: false, message: 'That reminder is no longer due.' };
    if (it.state === 'NO_CONSENT') return { success: false, message: 'This patient has not agreed to messages.' };
    rem_log_(it, /SMS/i.test(channel) ? 'SMS' : 'WHATSAPP', 'OPENED_BY_DESK', actor.username);
    return { success: true };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  }
}

/** FRONTEND ENTRY. Send a day's reminders now through the provider. */
function sendRemindersNow(sessionToken, opts) {
  var lock = LockService.getScriptLock();
  try {
    var actor = crescRequire_(sessionToken, 'appointment.write');
    if (!rem_provider_()) {
      return { success: false, message: 'No reminder provider is set up, so reminders are sent from the desk. ' +
                                        'See Patient_Reminders.gs for the WhatsApp Business and SMS settings.' };
    }
    lock.waitLock(20000);
    var r = rem_sendAll_(rem_dayFrom_((opts || {}).date), actor.username);
    return { success: true, sent: r.sent, failed: r.failed, skipped: r.skipped,
             message: r.sent + ' sent, ' + r.failed + ' failed, ' + r.skipped + ' not sendable.' };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** TIME-DRIVEN, 6 pm. Tomorrow's reminders, when a provider is set up. */
function remindersDaily(e) {
  crescTriggerOnly_(e, 'remindersDaily');
  var r = rem_sendAll_(rem_dayFrom_('tomorrow'), 'SYSTEM');
  var msg = r.provider ? ('Reminders: ' + r.sent + ' sent, ' + r.failed + ' failed, ' + r.skipped + ' skipped.')
                       : 'Reminders: no provider configured; the desk sends them from Operations -> Reminders.';
  Logger.log(msg);
  return msg;
}
