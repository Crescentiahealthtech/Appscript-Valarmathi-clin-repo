# DPDP Act readiness — Crescentia HealthTech

**Digital Personal Data Protection Act, 2023** · first assessed 14 September
2026 · **revised 15 September 2026, after the remediation described below**

This is an engineering assessment, not legal advice. It says what the software
does and does not do, and what to change. A lawyer has to decide whether your
clinic is a Significant Data Fiduciary and sign off your Consent Notice.

---

## Where you stand

| | |
|---|---|
| **Overall** | The code findings are closed. What is left is **deployment, decisions and documents** — and until the redeployment in step 1 below is done, the most serious finding is still live in production. |
| **Confirmed while doing this work** | The records are held in a spreadsheet owned by a **personal Gmail account**, so there is no Google Data Processing Addendum covering any of it (H6), and the patient passwords in it are the derived plain-text ones described in C3. |
| **Critical, was** | 3 — all three addressed in code; C1 needs a redeployment to take effect |
| **High, was** | 6 — five closed in code, H6 is a contract the clinic has to hold |
| **Medium, was** | 5 — all five closed |
| **New findings** | 1 — the old scanner undercounted: 136 endpoints were reachable without a session, not 62 |

Run `dpdpReadinessCheck()` from the Apps Script editor for the live version of
the technical half of this list, or open **Privacy → Posture** in the
application. The organisational half is below and no function can assert it.

### The four documents this assessment now has behind it

| | |
|---|---|
| `docs/CONSENT_NOTICE.md` | the words a patient is given, s.5 |
| `docs/RETENTION_SCHEDULE.md` | how long each record is kept and on what basis, s.8(7) |
| `docs/BREACH_PROCEDURE.md` | who decides, who notifies, by when, s.8(6) |
| `docs/PROCESSORS.md` | who else touches the data, and the record of processing, s.8(2) |

Each has blanks the clinic fills in. They are drafts of the clinic's documents,
not a substitute for the clinic adopting them.

---

## Do these five things this week

1. **Redeploy the web app.** `appsscript.json` now says `access: "ANYONE"`, but
   the manifest only takes effect on a **new version**: an `/exec` URL
   published before the change keeps `ANYONE_ANONYMOUS` until it is
   redeployed. Deploy → Manage deployments → edit → **New version**. Then open
   the URL in a private window: if it answers without asking you to sign in,
   the old deployment is still live. *(C1)*

   You no longer have to keep checking that by hand. `Deployment_Probe.gs`
   records, on every page load of the application itself, whether the caller
   arrived as an identified Google user — which is a direct consequence of the
   live access mode, and the only part of this that a script can observe. After
   a few days of use, `dpdpReadinessCheck()` reports the evidence rather than
   the manifest:

   * **CRITICAL** — "*12 of the last 340 page loads reached this application
     with no identified user, most recently 14-Sep*". The old deployment is
     still live; treat the period as a breach under s.8(6).
   * **LOW** — "*all 340 loads over 21 days were from an identified account*".
     The redeployment worked.

   `RUN_deploymentEvidence_()` prints the same answer on demand. The probe
   stores a count per day and nothing else — no email, no IP, no query string
   (which carries document grant keys).
2. **Hash the passwords.** `crescCredentialStatus()` to see where you are, then
   `crescMigrateCredentials()`. It issues a fresh random password per account,
   prints them **once**, and forces a change at first sign-in. Nobody can sign
   in with a plain-text credential from now on, so this is not optional — it is
   the step that lets people back in. *(C3)*
3. **Create the registers and name the officer.** `dpdpSetup()`, then
   `dpdpSetGrievanceOfficer("…", "…@…", "+91 …")` — or the Privacy console. *(H2, H3)*
4. **Install the scheduled jobs.** `dpdpInstallTriggers()`, then
   `dpdpTriggerStatus()` to confirm. Without them nothing expires, nobody reads
   the audit log, and no retention report is ever produced. *(H1, H4, H5)*
5. **Start the move to Google Workspace.** This has been checked: the
   spreadsheet is owned by `crescentiahealthtech@gmail.com`, a personal Google
   account, so there is no Data Processing Addendum and s.8(2) is met for none
   of the clinic's data — and there is no admin console to recover the records
   from if that account is ever lost. `docs/PROCESSORS.md` has the steps. *(H6)*

---

## Added since the 15 September revision

These close gaps the first assessment named and left open, and two that it did
not reach.

| | |
|---|---|
| **Consent is now read, not just recorded** | `DPDP_PURPOSES` has carried a `COMMUNICATION` purpose since the register was built, and **nothing consulted it**. Ten dispatch endpoints — pharmacy invoice, lab invoice, lab report, archived lab invoice, OP prescription, each by WhatsApp and by email — generated a PDF and sent it without looking. A patient who had refused reports by WhatsApp, or withdrawn that consent an hour earlier, got the message anyway. All ten now check it (`DPDP_Dispatch.gs`), and the notice they check against **names Meta**, because that is what happens to the data and s.5(1) wants the notice to describe what happens. A refusal at the counter becomes the question itself: the notice appears in the patient's words, either answer is written to the append-only register, and the send retries on a yes. |
| **Document links are Drive links with an expiry** | The private `?doc=&k=` route was the strongest link technically and the weakest in practice — a `script.google.com` URL with a 32-character key, arriving on WhatsApp with no preview, asking a patient to open their medical report. Files are published to Drive now and the patient gets the link Drive generates. The three protections that mattered are kept, because they were never the sharing mode: every file is a row in `Document_Grants` with an expiry, the daily sweep sets expired files back to `PRIVATE` so forwarded copies of the link die with them, and a revocation does it immediately. What is given up is the open counter, which Drive cannot report; the register says so rather than showing a `0` that reads as "nobody looked". **The daily trigger is what makes this safe** — without `dpdpInstallTriggers()` the clinic is back to permanent Drive links. |
| **The grant register can answer s.11(1)(b)** | Three of the five callers were passing an **invoice number** or a **bill id** into the `patientId` argument of `dpdpIssueDocumentLink_`, and a fourth passed a variable that was not in scope. So `Document_Grants` could not answer the one question s.11(1)(b) gives a patient the right to ask: what have you sent about me, and to whom. One resolver now finds the real patient for all of them. |
| **Consent for the existing patients** | `dpdpBackfillConsent()` in two modes. `LEGITIMATE` writes the two s.7 legitimate uses — care and billing — for every patient with no row, which is the accountability record the register was missing. It **refuses to invent** the four real consents: a row saying a patient agreed to something nobody asked them is a false record, and worse than the gap. `RECORD` writes consent the clinic actually holds on paper, but only with a stated source, which rides on every row. The Section 6 finding is now **two** findings, because "no record at all" is closeable by the clinic and "has not been asked" can only be closed by asking — the latter comes with a worklist (`dpdpConsentQueue()`, Privacy → Consent). |
| **Self-service password reset** | `crescRequestPasswordReset()` (`Auth_Reset.gs`) emails a temporary password to the address already on the record — never one supplied in the request. It gives **every outcome the same answer**, because patient ids are sequential and printed on every barcode label, so an endpoint that distinguishes "no such account" from "check your email" is a directory of who is a patient here. The difference is recorded in the audit log (`PASSWORD_RESET_REFUSED`). Rate-limited per account and per deployment: uncapped, it is a button that mails a working credential and invalidates the real one. |
| **The clinic's own web address on the notice** | `valarmathiclinic.crescentiahealthtech.com`, from `CLINIC_WEBSITE` in `Clinic_Profile.gs`, on the privacy page footer — s.5(1) identification a patient can retype. Printed verification links can use the same host once `CRESC_PUBLIC_BASE_URL` is set; it is **opt-in** because the QR beside the printed text is built from the same string, so an unconfigured host would break both. |
| **Reads by the patient are logged** | Every portal endpoint writes a `dpdpLogRead_` row against the patient, marked `self`. Finding M2 was about clinical reads leaving no trace; a patient reading their own record is still a read of it. |

---

## The findings

### C1 · The web app answered anyone, and answered as the owner — **CRITICAL**

`appsscript.json` said:

```json
"webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }
```

Every `google.script.run` endpoint ran with the deploying account's **full**
access to the spreadsheet — not the caller's — for anybody who had the `/exec`
URL, signed in or not. Patient IDs are sequential (`LMTVS0001`, `LMTVS0002`, …)
and printed on the barcode on every patient's card, so the identifier needed to
walk the register is public.

**Changed.** The manifest asks for `"access": "ANYONE"` — a Google sign-in is
required, which does not authorise anyone but does put a name and Google's own
rate limiting in front of the door. `executeAs` stays `USER_DEPLOYING` because
the script owns the spreadsheet; that is exactly why every endpoint now carries
its own check (C2).

**Still to do.**

1. **Redeploy** — see step 1 above. Until then this finding is live.
2. Treat the period the deployment was anonymous as potentially breached. Under
   s.8(6) you must notify the Board and every affected Data Principal. Ask your
   lawyer whether the URL's distribution history makes that necessary; record
   the decision either way in `Breach_Register` (`docs/BREACH_PROCEDURE.md`).

---

### C2 · 136 client-callable endpoints had no permission check — **CRITICAL — CLOSED**

The first assessment said 62. It was wrong, and the way it was wrong is worth
recording: `tools/rbac.js` read 700 characters after `google.script.run` and
took the last call in that window, so any endpoint sitting behind a long
success handler was never counted. `registerPatient` — an unauthenticated write
to the patient master — was one of the invisible ones. It also stopped at the
first `//` comment in a chain, which is how `setupLabDatabase` stayed hidden.
In the other direction it called every endpoint that delegates its check to a
gate helper unguarded.

The scanner now walks the chain properly and follows the call graph.

**The measurement today:**

```
428 public .gs functions
  281 carry a session or permission check
    8 are public by design (CRESC_PUBLIC_BY_DESIGN in RBAC.gs)
  139 carry none — 0 of those are reachable from the browser
```

**What was open**, before this change: the ward board with every inpatient's
name, bed, consultant and **diagnosis**; complete ward-notes and admission
documents; full lab reports with results; the admission register; live
inpatient bills; hospital invoices — read, raise **and settle**; insurance
claims and their settlement; the cash drawers and the financial period lock;
the finance dashboard; **writes** to clinical records; and the lab database
rebuild.

**Three things this needed beyond the one-line guards:**

- **An ambient actor** (`CRESC_CURRENT_ACTOR`, `RBAC.gs`). Server functions call
  each other — the case sheet raises a lab order, the discharge settlement
  settles the credit bill — and the inner call has no token to pass. A
  successful guard now records the actor for the length of **one execution**,
  so an inner guard checks its own permission against the same person. Apps
  Script gives every call a fresh script context, so it cannot leak between
  callers.
- **The lab status machine split in two.** Six internal callers drive it, each
  already guarded by what its own user is doing. Had they inherited the
  endpoint's permission, an accountant raising a lab bill would have been
  refused for not holding `lab.collect`.
- **An allowlist with reasons.** Eight endpoints answer without a session and
  say why in `CRESC_PUBLIC_BY_DESIGN`: the three sign-in steps, the password
  change, the clinic letterhead, the UI bundles, the DPDP notice, and the
  public request form. Anything not on that list and not guarded is a finding.

**Still to do.** 139 functions still carry no guard and are not called from any
`.html` here. Apps Script exposes every one of them to `google.script.run`
anyway, so they are a smaller hole rather than none. `node tools/rbac.js`
names them in its closing note; work them down when you touch those files.

---

### C3 · Passwords were stored and compared in plain text — **CRITICAL — CLOSED IN CODE**

`AuthLogin.gs` compared `storedPassword.toString().trim() === passwordInput`.
Staff passwords sat in column B of `Users` in clear: everyone with edit access
to the spreadsheet, every Google account it had ever been shared with and
anyone with a copy could read all of them. People reuse passwords, so the blast
radius was never this system.

Patient portal passwords were worse — they were **derived**.
`registerPatient()` generated `Mei2001`: the first three letters of the name
plus the birth year, stored in clear, with both inputs printed on the
registration slip, beside the barcode of the ID needed to use them.

**Changed** — `Auth_Credentials.gs`:

- A salted, iterated SHA-256 digest (`pbkdf2$sha256$iterations$salt$digest`),
  10,000 iterations by default, tunable with the `CRESC_PWD_ITERATIONS` script
  property. Apps Script has no native slow KDF and cannot load one; this is
  what the platform permits, and the file says so rather than implying more.
- Constant-time comparison.
- A plain-text cell is **refused** at sign-in and cannot be used to change
  itself either. Hashing a password that has already been readable keeps the
  compromise; letting it authorise its own replacement hands the account to
  whoever copied it. Only an administrator resets it.
- New portal passwords are random, shown once at the desk, and must be changed
  at first sign-in. `Auth.html` has the screen that happens on.
- `crescAdminResetPassword()` for one account, `crescMigrateCredentials()` for
  all of them.

**Still to do.**

1. Run `crescMigrateCredentials()` — step 2 above. Nobody can sign in until you
   do, which is the intended shape of a forced reset.
2. Hand the passwords out by a route that is not the one they sign in through,
   and do not paste the migration log anywhere.
3. Consider dropping the patient portal password entirely in favour of an OTP
   to the registered mobile. That removes the stored credential and gives you
   the identity check that s.11 requests need anyway (M5) — it needs an SMS
   gateway this deployment does not have.

---

### H1 · Patient documents were published to the open web and never revoked — **HIGH — CLOSED**

Five functions did `file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, …)` for
lab reports, prescriptions, pharmacy invoices and lab invoices, then sent the
Drive URL over WhatsApp. Each file carries the patient's name, ID, and for a
report their **results and diagnosis**. The link never expired, nothing recorded
that the file existed, and WhatsApp messages are forwarded, backed up and
restored onto devices nobody here will ever see.

**Changed** — `DPDP_Documents.gs`. The file stays **private**. The patient gets
a link back into the web app carrying a one-off random key:

- it expires (14 days by default),
- it counts its opens and closes itself after 25,
- it can be withdrawn in one call from the Privacy console,
- every open writes an audit row,
- the register stores only a **digest** of the key, so the spreadsheet does not
  hold the capability,
- `dpdpIssueDocumentLink_()` also un-shares any public sharing the file already
  had, so re-sending a document closes the old hole rather than adding to it.

`?viewReport=<orderId>` is **withdrawn rather than fixed**. It returned any
order's complete lab report as a PDF to anyone who asked, with no key and no
session, and order ids are sequential and printed on the patient's own
paperwork. Old messages now get a page saying the clinic will re-send it.

**Still to do.** Run `dpdpExpireSharedLinks()` once against the backlog of
files published before this change — `dpdpDailyMaintenance()` does it nightly,
but the first run over a long backlog is worth watching.

---

### H2 · No consent was captured, ever — **HIGH — CLOSED IN CODE**

s.6 requires consent that is free, specific, informed, unambiguous, given by a
clear affirmative action and **as easy to withdraw as to give**. s.6(10) puts
the burden of proving it on you. Registration collected twenty-two fields and
asked nothing.

**Changed.** `Consent_Register` is append-only, one row per purpose per
decision, stamped with the notice version. Registration now captures it in the
same call that creates the patient, and:

- **an unticked box is recorded as a REFUSAL** — "we never asked" and "they said
  no" are not the same fact, and only one of them needs asking again;
- treatment and billing are shown as **legitimate uses under s.7**, not tick
  boxes, because a consent dialog should not be able to stop somebody being
  treated;
- under-18 registration asks who consented and how they are related (s.9);
- the patient can give or withdraw any of it themselves, on their own portal
  page — which is what s.6(6) actually asks for;
- `docs/CONSENT_NOTICE.md` is the notice, in words.

**Still to do.** Adopt the notice: fill in its blanks, **translate it into
Tamil**, put it on the wall and print `?privacy` on the registration slip. Then
backfill — every patient already on file has no consent record, and
`dpdpReadinessCheck()` counts them until they do.

---

### H3 · No way to answer a data principal's request — **HIGH — CLOSED**

ss.11 (access), 12 (correction and erasure) and 13 (grievance) each carry a
deadline. There was no register, no clock and no way to assemble what is held
about one patient — it is spread across eleven sheets.

**Changed.** `DPDP_Requests` with a `Due_By` on every row; `raiseDPDPRequest()`,
`listDPDPRequests()` (overdue first), `closeDPDPRequest()` — which refuses an
outcome shorter than a sentence, because "Closed" is not an answer to a
statutory request. `exportPatientData()` assembles demographics, consent,
appointments, consultations, admissions, case sheets, ward notes, lab orders
and bills, pharmacy and hospital bills, referrals, insurance claims, the
nomination, and every document link issued — which is s.11(1)(b), the
identities data has been disclosed to.

And the two halves that were missing:

- **A form the patient can reach.** `?privacy` on the web app serves the notice
  and a request form with no staff login. A form only staff can reach is a way
  of receiving fewer requests, not of answering them.
- **A place to work the queue.** Privacy → Requests, on the rail.

**Still to do.** Name the officer (step 3), and publish the `?privacy` URL where
patients see it.

---

### H4 · No retention policy — **HIGH — CLOSED IN CODE**

s.8(7): erase when the purpose is served, unless the law requires otherwise.
Nothing in this system ever deleted anything.

**Changed.** `docs/RETENTION_SCHEDULE.md` is the schedule, per record type, with
its legal basis and **who decides**. `dpdpRetentionReport()` reports what is
past its period and **deletes nothing** — a medico-legal case or an insurance
dispute can require a record years after the ordinary period, and a sweep that
deletes on its own is how a clinic loses the file it is about to be asked for.
`dpdpMonthlyRetentionReport()` emails it. What *is* deleted automatically is
what nobody could ever need: expired document links, dead sessions, and
(once, by hand) the three columns nothing ever read.

**Still to do.** Have the schedule read and adopted — particularly the
inpatient and medico-legal periods, which depend on your state's Clinical
Establishments rules and are marked *take advice*.

---

### H5 · No breach detection or notification procedure — **HIGH — CLOSED**

s.8(6) requires notification to the Board **and to every affected Data
Principal**. There was no procedure, no template and no monitoring.

**Changed.** `DPDP_Breach.gs` and `docs/BREACH_PROCEDURE.md`:

- `Breach_Register`, which holds incidents judged **not** notifiable too, with
  the reason and the named person who decided. In an inquiry that register is
  the only thing that shows the clinic was looking.
- `dpdpAssessBreach()` refuses a decision without a reason — "not notifiable" is
  the one you will be asked to justify.
- `dpdpBreachNotice()` drafts both notifications from the row.
- `dpdpWeeklyReview()` reads the audit log every Monday for the four patterns
  that precede a disclosure here — repeated failures against one account,
  failures spread across many identifiers (which is what walking the patient ID
  range looks like), one person reading an unusual number of records in a day,
  and one shared document opened six times or more — and emails the officer.

**Still to do.** Put names in the procedure's table, and check the Board's
current intimation route **before** you need it.

---

### H6 · No processor agreements — **HIGH — OPEN, and not an engineering problem**

s.8(2): a Data Fiduciary remains responsible for processing done by a Data
Processor, and must have a contract with them. `docs/PROCESSORS.md` is the
register and the record of processing.

The one that matters has now been checked, and the answer is the bad one:
**the clinic's spreadsheet is owned by `crescentiahealthtech@gmail.com`, a
personal Google account.** There is therefore no Data Processing Addendum, and
s.8(2) is met for **none** of the clinic's patient data.

That also means there is no admin console: if that account is lost, recovered
by somebody else, or belongs to a person who leaves, the clinic's entire
medical record goes with it. Migrating to Google Workspace on the clinic's own
domain is the highest-value action on this whole list, and it is not something
code can do. `docs/PROCESSORS.md` has the steps.

WhatsApp's consumer terms are not a processor agreement either. Either move
dispatch to the Business API with a contract, or rely on the patient's specific
consent to receive documents that way — which `recordConsent(… COMMUNICATION …)`
records per patient — and say in the notice that the message goes through Meta.

---

### M1 · Data was collected that is never used — **MEDIUM — CLOSED**

Registration collected `Education`, `Occupation` and `Marital_Status`; nothing
in the codebase read them. s.6(1) permits collection for a specified purpose,
and a field that exists because the form had a box has no purpose to specify.

**Changed.** They are gone from registration, from the profile editor and from
the patient search. The columns stay — the sheet layout is positional and every
column index in the project depends on it — and `dpdpEraseUnusedFields(true)`
clears what was already collected. That function deletes where the retention
sweep deliberately does not, because there is no legal reason to keep a field
nothing reads and no upside to weigh against erasing it.

---

### M2 · The audit log did not record reads — **MEDIUM — CLOSED**

`logAudit_()` was called on writes. Opening a record, printing a report and
exporting a ledger left no trace, so "who looked at this patient's file" — the
question asked after every incident — could not be answered.

**Changed.** `dpdpLogRead_()` records the reads that matter: the whole patient
register, the master directory, a lab report, a prescription, a case sheet, an
admission record, a patient search, a longitudinal timeline, and every open of
a shared document link. Not every cell fetched to paint a list — that is volume
without information. `dpdpWeeklyReview()` then reads those rows for anomalies,
which is what makes logging them worth the space.

---

### M3 · Voice typing sends audio to a third party — **MEDIUM — CLOSED**

Dictation uses the browser's Web Speech API; in Chrome and Edge that sends the
audio to the browser vendor's speech service.

**Changed.** The per-browser consent dialog was already there and is the right
thing to show the person dictating — but it is the wrong place to make the
decision. One clinician clicking *Enable* on one laptop is not the clinic
deciding that recorded clinical speech may be sent to a third party. So the
clinic decides once (`dpdpSetVoicePolicy`, on the Privacy console), the browser
dialog can only ask within that decision, and the feature **fails closed** if it
cannot reach the policy. `dpdpReadinessCheck()` names an undecided policy until
somebody decides either way.

**Still to do.** Decide, and say which in the notice — the paragraph is drafted
both ways in `docs/CONSENT_NOTICE.md`.

---

### M4 · No nomination mechanism — **MEDIUM — CLOSED**

s.14 gives a Data Principal the right to nominate someone to exercise their
rights if they die or become incapable.

**Changed.** `Nomination_Register`, append-only, with `recordNomination()`,
`getNomination()` and `revokeNomination()`. A patient can do it themselves in
the portal.

Deliberately **not** the emergency contact already on the patient record: that
was collected so somebody could be reached in a hurry, it was never given for
this purpose, and s.14 requires the patient's own act. The register is
append-only because "who was nominated on the day they died" has exactly one
right answer, and it is not "whoever the row says now".

---

### M5 · No verification that a requester is who they say — **MEDIUM — CLOSED**

**Changed.** Every request row records **how** the clinic satisfied itself that
the person asking is the person the data is about, and `closeDPDPRequest()`
refuses to close a request that nobody verified. Answering an access request to
the wrong person is a disclosure dressed as compliance, and the statutory right
is exactly what an attacker would use to ask for it. A portal session counts as
verification by itself; everything else is recorded from a short list, of which
a call back to the number already on the record is the cheapest and the one
thing an impersonator cannot arrange.

---

## Checklist

Copy this into your compliance file and tick it as you go. Items already done
in code are ticked; the rest are yours.

### Notice and consent
- [x] Consent recorded per purpose, not as one blanket agreement
- [x] Consent recorded against a notice **version**
- [x] Withdrawal as easy as giving, and it works — the patient can do it themselves
- [x] Guardian consent for patients under 18, with the guardian named
- [x] No tracking or targeted advertising directed at children (§9(3)) — none exists
- [ ] Consent Notice adopted, in plain language, **and in Tamil**
- [ ] Notice given at or before collection — on the wall, and `?privacy` printed on the slip
- [ ] A stated method for verifying guardian consent, not just recording it
- [ ] Backfill: every patient already on file has no consent record

### Security (§8(5))
- [x] Every client-callable endpoint carries a check, or is on an allowlist with a reason
- [x] Passwords hashed and salted; plain-text ones refused
- [x] Patient portal credentials not derivable from printed data
- [x] No patient document published with a permanent public link
- [ ] Web app **redeployed** so `access: ANYONE` is actually in force
- [ ] Spreadsheet access limited to staff who need it, and reviewed
- [ ] Offboarding removes the login, the spreadsheet access **and** the Drive folders

### Retention (§8(7))
- [x] Retention schedule written down, per record type, with its legal basis
- [x] Sweep runs and is reviewed by a human before anything is deleted
- [x] Sessions and links cleared on their own schedule
- [x] Marketing data erased on withdrawal
- [ ] Inpatient and medico-legal periods confirmed against your state's rules
- [ ] A paper retention and shredding policy — nothing here covers print-outs

### Data principal rights (§§11–14)
- [x] Access request answerable — `exportPatientData()` across eleven sheets
- [x] Correction and completion possible — `updatePatientProfile()`
- [x] Erasure request assessed against the schedule, and the answer explained
- [x] Nomination capturable
- [x] Requester identity verified, and the method recorded
- [ ] Grievance officer named and published

### Accountability (§§8(2), 8(6), 10)
- [x] Audit log covers reads as well as writes
- [x] Breach register kept, including incidents judged not notifiable
- [x] Breach procedure written, with who decides and who notifies
- [ ] Names filled into that procedure
- [ ] Data Processor agreements in place — **Google first**
- [ ] Whether you are a Significant Data Fiduciary assessed in writing
- [ ] If you are: DPO appointed, independent audit, DPIA done (§10)

### Records of processing
- [x] Every category of personal data listed, with its purpose and legal basis
- [x] Every place data leaves the system listed
- [x] Cross-border transfer identified
- [ ] The blanks in `docs/PROCESSORS.md` filled in

---

## The machinery this repository provides

### Set-up, once
| Function | What it does |
|---|---|
| `dpdpSetup()` | creates every register |
| `dpdpSetGrievanceOfficer(name, email, phone)` | §13 |
| `dpdpInstallTriggers()` | the daily, weekly and monthly jobs |
| `crescCredentialStatus()` / `crescMigrateCredentials()` | §8(5) — see what is stored, then hash it |
| `dpdpEraseUnusedFields(true)` | §6(1) — clear the three columns nothing reads |

### Notice and consent
| Function | What it does |
|---|---|
| `getDPDPNotice()` | the notice's structure, unauthenticated by design |
| `recordConsent(payload, token)` | §6, append-only, per purpose |
| `getConsentStatus(patientId, token)` | current state, `NOT_ASKED` distinguished from `REFUSED` |
| `withdrawConsent(patientId, purpose, reason, token)` | §6(6) |
| `recordNomination` / `getNomination` / `revokeNomination` | §14 |

### Rights
| Function | What it does |
|---|---|
| `raiseDPDPRequest(payload, token)` | §§11–13, with the clock |
| `dpdpSubmitPublicRequest(payload)` | the same, from the public page, always unverified |
| `dpdpVerifyRequester(id, method, note, token)` | §11 — how you know it is them |
| `listDPDPRequests(token, opts)` | overdue first |
| `closeDPDPRequest(id, outcome, token)` | refuses a non-answer, and an unverified request |
| `exportPatientData(patientId, token)` | §11, across eleven sheets |

### Documents and retention
| Function | What it does |
|---|---|
| `dpdpIssueDocumentLink_(file, type, patientId, by)` | §8(5) — private, expiring, counted, revocable |
| `dpdpRevokeDocumentLink(grantId, token)` | withdraw one |
| `dpdpListDocumentLinks(patientId, token)` | §11(1)(b) — what was sent, and how often it was opened |
| `dpdpExpireDocumentGrants()` / `dpdpExpireSharedLinks(dryRun)` | the sweeps |
| `dpdpRetentionReport()` | §8(7) — reports, never deletes |

### Breach
| Function | What it does |
|---|---|
| `dpdpAnomalyScan(days, token)` | the four patterns, from the audit log |
| `dpdpRaiseBreach` / `dpdpAssessBreach` / `dpdpCloseBreach` | the register, and the decision with its reason |
| `dpdpBreachNotice(id, token)` | §8(6) — both notifications, drafted |
| `dpdpRecordBreachNotification(id, who, token)` | that you actually told them |

### Checking
| Function | What it does |
|---|---|
| `dpdpReadinessCheck()` | the technical half of this document, live |
| `dpdpConsoleSnapshot(token)` | the same, plus the queue, for the Privacy console |
| `dpdpTriggerStatus()` | whether the jobs are installed |
| `crescRbacCoverage()` / `node tools/rbac.js` | endpoints with no check |
