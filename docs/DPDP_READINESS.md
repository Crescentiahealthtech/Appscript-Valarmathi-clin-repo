# DPDP Act readiness — Crescentia HealthTech

**Digital Personal Data Protection Act, 2023** · assessed 14 September 2026
against the code in this repository.

This is an engineering assessment, not legal advice. It says what the
software does and does not do, and what to change. A lawyer has to decide
whether your clinic is a Significant Data Fiduciary and sign off your
Consent Notice.

---

## Where you stand

| | |
|---|---|
| **Overall** | Not ready. Four findings would each on their own be a reportable breach. |
| **Critical** | 3 |
| **High** | 6 |
| **Medium** | 5 |
| **What is already in place** | An append-only audit log, a server-side permission matrix (`RBAC.gs`), durable sessions, lockout on repeated failed sign-ins, and — as of this change — consent, rights and retention registers. |

Run `dpdpReadinessCheck()` from the Apps Script editor for the live version of
the technical half of this list. The organisational half is below and no
function can assert it.

---

## The findings

### C1 · The web app answers anyone, and answers as the owner — **CRITICAL**

`appsscript.json`:

```json
"webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }
```

Every `google.script.run` endpoint runs with the deploying account's **full**
access to the spreadsheet — not the caller's — for anybody who has the `/exec`
URL, signed in or not. Patient IDs are sequential (`LMTVS0001`, `LMTVS0002`, …)
and printed on the barcode on every patient's card, so the identifier needed to
walk the register is public.

Until this change, 66 client-callable functions carried no session check at
all. Two of them were as bad as it gets:

- `getAllPatients()` returned **the entire patient register** — name, age,
  sex, mobile, address — with no token and no check.
- `saveUserProfile(data)` was an **unauthenticated write** to any patient's
  name, date of birth, mobile and address, keyed on a `username` the caller
  supplied.

**Fixed in this change:** both are now behind `crescRequire_`, and their
callers pass the session token.

**Still to do — this is the fix that matters most:**

1. Redeploy with `"access": "ANYONE"` (Google sign-in required) or `"DOMAIN"`.
   `ANYONE_ANONYMOUS` means the only thing between the clinic's whole record
   and the internet is that nobody has guessed the URL.
2. Work `node tools/rbac.js` down to zero. 62 client-callable functions still
   have no guard, and it names every one of them. §C2 below ranks them.
3. Treat the current deployment as potentially breached. Under §8(6) you must
   notify the Data Protection Board and every affected Data Principal. Ask
   your lawyer whether the URL's distribution history makes that necessary.

---

### C2 · 62 client-callable endpoints still have no permission check — **CRITICAL**

Measured by `tools/rbac.js`, which is check 7 of `tools/check.sh`: **397
public `.gs` functions, 108 with a session or permission check, 289 with
none** — of which **62** are reachable from the browser today. Re-run it as
you work through the list; the number is the progress bar.

The worst, by what they return:

| Endpoint | What an unauthenticated caller gets |
|---|---|
| `getActiveIPWard`, `getWardBedBoard` | every inpatient: name, age/sex, bed, consultant, **diagnosis** |
| `getIPNotesPrintHtml`, `getIPRecordPrintHtml` | a complete ward-notes or admission document |
| `getLabReportHtml`, `getLabBillHtml` | a full lab report with results |
| `getActiveIPAdmissionsForNotes`, `getIPHistory` | the admission register |
| `getRunningTab`, `getOpenAdmissions` | live inpatient bills |
| `hb_getInvoice`, `hb_saveInvoice`, `hb_recordPayment` | read **and raise and settle** hospital bills |
| `getInsuranceClaims`, `settleClaim` | insurance claims, and settling them |
| `openShift`, `closeShift`, `lockFinancialPeriod` | open and reconcile cash drawers; lock the books |
| `getFinanceDashboard`, `getReceivables` | the clinic's financial position |
| `saveIPNote`, `saveNewAdmissionLedger` | **write** clinical records |
| `setupLabDatabase` | rebuild the lab database |

`node tools/rbac.js` prints the full list, by file.

**Fix.** One line at the top of each:

```js
function getActiveIPWard(sessionToken) {
  var actor = crescRequire_(sessionToken, 'ward.read');
  …
}
```

Do it in severity order: anything returning clinical text first, anything
writing second, anything financial third. `RBAC.gs` already has the vocabulary
and the matrix; nothing new has to be designed.

---

### C3 · Passwords are stored and compared in plain text — **CRITICAL**

`AuthLogin.gs`:

```js
if (storedPassword.toString().trim() === passwordInput) { … }
```

Staff passwords sit in column B of the `Users` sheet in clear. Anyone who can
open the spreadsheet — every staff member with edit access, every Google
account the file has ever been shared with, anyone who gets a copy — can read
all of them. People reuse passwords, so the blast radius is not this system.

Patient portal passwords are worse. `registerPatient()` generates them as
**the first three letters of the name plus the birth year** (`Mei2001`), stores
them in clear in column B, and prints both inputs on the registration
confirmation. Anybody holding a patient's prescription can derive their portal
password.

**Fix.**

1. Store a salted hash. Apps Script has `Utilities.computeDigest(SHA_256, …)`;
   generate a per-user random salt, store `salt` and `hash`, compare digests
   with a constant-time comparison.
2. Force every existing password to be reset when you deploy it. Migrating a
   plaintext column by hashing it in place keeps the compromise.
3. Generate patient portal passwords randomly and require a change at first
   sign-in. Never derive a credential from data printed on a document.
4. Reconsider whether the patient portal needs a password at all — an OTP to
   the registered mobile removes the stored credential entirely.

---

### H1 · Patient documents are published to the open web and never revoked — **HIGH**

Five functions do this:

```js
file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
```

— for lab reports, prescriptions, pharmacy invoices and lab invoices, then send
the link over WhatsApp. Each file carries the patient's name, ID, and in the
case of a report their **results and diagnosis**. The link never expires,
nothing recorded that the file existed, and WhatsApp messages are forwarded.

This is §8(7) (retain no longer than necessary) and §8(5) (reasonable security
safeguards) at the same time.

**Fixed in this change:** every one of the five now calls
`dpdpRegisterSharedFile()`, so the set is finite and listable.
`dpdpExpireSharedLinks()` revokes anything past a 30-day window.

**Still to do:** add a daily time-driven trigger for `dpdpExpireSharedLinks()`,
and run it once against the backlog. Longer term, serve documents through the
web app behind a session instead of publishing them to Drive at all.

---

### H2 · No consent was captured, ever — **HIGH**

§6 requires consent that is free, specific, informed, unambiguous, given by a
clear affirmative action, and **as easy to withdraw as to give**. §6(10) puts
the burden of proving it on you.

Before this change there was no consent record of any kind. Registration
collected twenty-two fields and asked nothing.

**Fixed in this change:** `Consent_Register`, append-only, one row per purpose
per decision, stamped with the notice version it was given against.
`recordConsent()`, `getConsentStatus()`, `withdrawConsent()`. Purposes are
separable — treatment, billing, insurance, communications, marketing, research
— because one "I agree" covering all six is not specific, and marketing is
precisely what a patient wants to refuse while still being treated.

Treatment and billing are recorded as **legitimate uses** under §7, not gated
behind a tick box: a patient who has approached a clinic for care should not be
blocked from being treated by a consent dialog.

**Still to do:**

1. Write the Consent Notice. The machinery is here; the words are not.
   `getDPDPNotice()` returns the structure to render it from.
2. Put consent capture in the registration flow, and a prompt at the next
   visit for the patients already on file.
3. §9 — a patient under 18 needs verifiable consent from a parent or guardian.
   `recordConsent()` refuses without a named guardian, but "verifiable" means
   you must also decide *how* you verify, and record that.

---

### H3 · No way to answer a data principal's request — **HIGH**

§11 (access), §12 (correction and erasure), §13 (grievance) each carry a
deadline. There was no register, no clock and no way to assemble what is held
about one patient — it is spread across eleven sheets.

**Fixed in this change:** `DPDP_Requests` with a `Due_By` on every row;
`raiseDPDPRequest()`, `listDPDPRequests()` (overdue first),
`closeDPDPRequest()` — which refuses an outcome shorter than a sentence,
because "Closed" is not an answer to a statutory request. `exportPatientData()`
assembles demographics, consent, appointments, consultations, admissions, case
sheets, ward notes, lab orders and bills, pharmacy and hospital bills,
referrals, insurance claims, and the list of documents shared — which is
§11(1)(b), the identities data has been disclosed to.

**Still to do:** name the grievance officer
(`dpdpSetGrievanceOfficer(…)`) — §13 requires you to publish one — and put the
request form somewhere a patient can reach it.

---

### H4 · No retention policy — **HIGH**

§8(7): erase when the purpose is served, unless the law requires otherwise.
Nothing in this system ever deletes anything. Sessions, audit rows, old
appointments and years-old consultations accumulate indefinitely.

**Fixed in this change:** `dpdpRetentionReport()` reports what is past its
period, per sheet, **and deletes nothing**. That is deliberate: a medico-legal
case or an insurance dispute can require a record years after the ordinary
period, and a sweep that deletes on its own is how a clinic loses the file it
is about to be asked for.

**Still to do:** decide the schedule and write it down. Starting points used in
`DPDP_CFG.RETENTION_DAYS`:

| Record | Period | Basis |
|---|---|---|
| Outpatient clinical record | 3 years from last entry | NMC Ethics Regulations 2002, reg. 1.3.1 |
| Inpatient / medico-legal | longer — take advice | state Clinical Establishments rules |
| Financial records | 8 years | Income Tax Act practice |
| Sign-in sessions | 30 days | no reason to keep them |
| Audit log | 3 years | matches the clinical period it documents |
| Marketing contact data | 1 year, or until withdrawal | no statutory basis to keep it |

---

### H5 · No breach detection or notification procedure — **HIGH**

§8(6) requires notification to the Data Protection Board **and to every
affected Data Principal** on a personal data breach. There is no procedure, no
template and no monitoring.

`Audit_Log` and `Auth_Audit` exist and are append-only, which is the hard part.
Nothing reads them for anomalies, and nobody is assigned to.

**Fix.**

1. Write the procedure: who decides it is a breach, who notifies, within what
   time, using what template. One page.
2. Add a weekly review of `Auth_Audit` for failed-login clusters and of
   `Audit_Log` for bulk reads. A trigger that emails the officer a summary is
   twenty lines.
3. Keep a breach register even for incidents you conclude are not notifiable.
   The record of having considered it is worth as much as the conclusion.

---

### H6 · No processor agreements — **HIGH**

§8(2): a Data Fiduciary remains responsible for processing done by a Data
Processor, and must have a contract with them.

This deployment processes personal data through:

| Processor | What reaches them |
|---|---|
| **Google** (Sheets, Apps Script, Drive, Gmail) | everything |
| **WhatsApp / Meta** | patient name and mobile in the message, plus a link to a document with their results |
| **Google speech services** | dictated clinical text, if voice typing is used |

Google Workspace's Data Processing Addendum covers the first if you are on
Workspace rather than a personal Gmail account. **Check which.** A clinic
running its records on a personal `@gmail.com` account has no DPA at all.

WhatsApp's consumer terms are not a processor agreement. Either move dispatch
to a business API with a contract, or record the patient's specific consent to
receive documents that way — which `recordConsent(… COMMUNICATION …)` now
supports.

---

### M1 · Data is collected that is never used — **MEDIUM**

§6(1) permits consent only for the purpose specified. Registration collects
`Education`, `Occupation` and `Marital_Status`; nothing in the codebase reads
them for any clinical or billing purpose.

**Fix.** Either drop the fields or name the purpose in the notice. Data
collected "because the form had a box" is the easiest finding to close and the
easiest to leave open.

---

### M2 · The audit log does not record reads — **MEDIUM**

`logAudit_()` is called on writes. Opening a patient's record, printing a
report and exporting a ledger leave no trace — so the question "who looked at
this patient's file" cannot be answered, and that is the question asked after
an incident.

**Fix.** Log reads of clinical documents at least: `getIPRecordPrintHtml`,
`getLabReportHtml`, `exportPatientData` (already logged), and the patient
search. Sample rather than log every row read if volume is a concern.

---

### M3 · Voice typing sends audio to a third party — **MEDIUM**

The dictation added in this change uses the browser's Web Speech API. In Chrome
and Edge that sends the audio to the browser vendor's speech service.

**Already handled:** the consent notice is shown before first use, in the words
of what actually happens, and the answer is stored per browser. It is opt-in
and the application records nothing itself.

**Still to do:** decide clinic policy and say so in the notice. A clinic that
does not want audio leaving the building simply never enables it.

---

### M4 · No nomination mechanism — **MEDIUM**

§14 gives a Data Principal the right to nominate someone to exercise their
rights if they die or become incapable. Nothing captures a nominee.

The `Relation_Type` / `Relation_Name` and emergency-contact fields on the
patient record are close but are not a nomination — they are a contact, and
were not collected for that purpose.

**Fix.** Add a nomination purpose to the consent register with the nominee's
name and relationship, captured with the same affirmative action.

---

### M5 · No verification that a requester is who they say — **MEDIUM**

`exportPatientData()` and `raiseDPDPRequest()` are guarded by role, so staff
raise requests on a patient's behalf. Nothing defines how the clinic satisfies
itself that the person at the desk is the patient.

**Fix.** Write the identity-verification step into the procedure — an OTP to
the registered mobile is the cheapest and is already possible with the number
on file — and record which method was used on the request row.

---

## What to do, in order

**This week**

1. Redeploy with `access: "ANYONE"` instead of `ANYONE_ANONYMOUS`. *(C1)*
2. Hash the passwords and force a reset. *(C3)*
3. Run `dpdpSetup()` and `dpdpSetGrievanceOfficer(…)`. *(H2, H3)*
4. Run `dpdpExpireSharedLinks()` against the backlog and add the daily trigger. *(H1)*
5. Check whether the spreadsheet is on Workspace or a personal Gmail account. *(H6)*

**This month**

6. Guard the clinical and write endpoints — 62 of them, worst first. *(C2)*
7. Write the Consent Notice and put capture into registration. *(H2)*
8. Decide and document the retention schedule. *(H4)*
9. Write the one-page breach procedure and name who decides. *(H5)*

**This quarter**

10. Move document delivery off public Drive links and behind the session. *(H1)*
11. Log clinical reads. *(M2)*
12. Drop the unused fields or justify them. *(M1)*
13. Add nomination and identity verification. *(M4, M5)*

---

## Checklist

Copy this into your compliance file and tick it as you go.

### Notice and consent
- [ ] Consent Notice written in plain language, and in every language the clinic serves
- [ ] Notice given **at or before** collection, not after
- [ ] Consent recorded per purpose, not as one blanket agreement
- [ ] Consent recorded against a notice **version**
- [ ] Withdrawal as easy as giving, and it works
- [ ] Guardian consent for patients under 18, with a stated verification method
- [ ] No tracking or targeted advertising directed at children (§9(3))

### Security (§8(5))
- [ ] Web app not deployed as `ANYONE_ANONYMOUS`
- [ ] Every client-callable endpoint carries `crescRequire_`
- [ ] Passwords hashed and salted; no plaintext anywhere
- [ ] Patient portal credentials not derivable from printed data
- [ ] No patient document published with a permanent public link
- [ ] Spreadsheet access limited to staff who need it, and reviewed
- [ ] Offboarding removes both the login and the spreadsheet access

### Retention (§8(7))
- [ ] Retention schedule written down, per record type, with its legal basis
- [ ] Sweep runs and is reviewed by a human before anything is deleted
- [ ] Sessions and logs cleared on their own schedule
- [ ] Marketing data erased on withdrawal

### Data principal rights (§§11–14)
- [ ] Access request answerable within the committed period
- [ ] Correction and completion possible — *(done: `updatePatientProfile`)*
- [ ] Erasure request assessed against the retention schedule, and the answer explained
- [ ] Grievance officer named and published
- [ ] Nomination capturable
- [ ] Requester identity verified, and the method recorded

### Accountability (§§8(2), 8(6), 10)
- [ ] Data Processor agreements in place — Google, WhatsApp, any other
- [ ] Breach procedure written, with who decides and who notifies
- [ ] Breach register kept, including incidents judged not notifiable
- [ ] Audit log covers reads as well as writes
- [ ] Whether you are a Significant Data Fiduciary assessed in writing
- [ ] If you are: DPO appointed, independent audit, DPIA done (§10)

### Records of processing
- [ ] Every category of personal data listed, with its purpose and legal basis
- [ ] Every place data leaves the system listed — WhatsApp, email, Drive, TPAs
- [ ] Cross-border transfers identified (Google's storage is not all in India)

---

## The machinery this repository now provides

| Function | What it does |
|---|---|
| `dpdpSetup()` | creates the three registers |
| `dpdpSetGrievanceOfficer(name, email, phone)` | §13 |
| `getDPDPNotice()` | the notice's structure, unauthenticated by design |
| `recordConsent(payload, token)` | §6, append-only, per purpose |
| `getConsentStatus(patientId, token)` | current state, `NOT_ASKED` distinguished from `REFUSED` |
| `withdrawConsent(patientId, purpose, reason, token)` | §6(6) |
| `raiseDPDPRequest(payload, token)` | §§11–13, with the clock |
| `listDPDPRequests(token, opts)` | overdue first |
| `closeDPDPRequest(id, outcome, token)` | refuses a non-answer |
| `exportPatientData(patientId, token)` | §11, across eleven sheets |
| `dpdpRegisterSharedFile(file, type, patientId, by)` | makes a shared link revocable |
| `dpdpExpireSharedLinks(dryRun)` | revokes the expired ones |
| `dpdpRetentionReport()` | §8(7) — reports, never deletes |
| `dpdpReadinessCheck()` | the technical half of this document, live |
