# IP Discharge Summary Engine — Phase 1 Discovery

**Status:** read-only survey. No production file was modified to produce this document.
**Evidence base:** the Apps Script project in this repository plus a full export of the live
bound spreadsheet (tenant `VALARMATHI`, exported 11 Sep 2026). Every schema below is the
*live header row*, not the header a setup function hopes to write — where the two disagree,
that is recorded as a finding.

---

## 1. Discharge code paths

There are **two** places an admission becomes `DISCHARGED`, and they do not know about each other.

| # | Entry point | File:line | Trigger | Writes |
|---|---|---|---|---|
| 1 | `processPatientDischarge(ipNumber, bedId)` | `IP_Admissions_Logic.gs:599` | IP Admissions live-ward "Discharge" action | `IP_Admissions.Status = 'DISCHARGED'` (`:620`), `IP_Admissions.DOD = new Date()` (`:621`), then `ipa_releaseBed_()` → `Master_Beds.Status` |
| 2 | `settleDischarge(payload)` | `AccountsIPChargesLogic.gs:299` | Accounts → IP Charges "grand settlement" | Posts ward charges, freezes `ON_TAB` → `SETTLED`, writes `IP_Settlements`, applies advances, creates insurance claims, ledger receipts, then `adsh.getRange(adm.row, 12).setValue('DISCHARGED')` and `.getRange(adm.row, 13)` = DOD (`:403–404`), bed → `Cleaning` (`:406`), deletes the `IP_Discharge_Drafts` row |

Notes that matter for the gate (Phase 8):

* Path 2 writes the status by **hard-coded column number 12/13**, not by header lookup. Any
  column insert into `IP_Admissions` silently corrupts discharges. (See §11.)
* Path 2 is where the money is, so it is the path that actually runs in production. Path 1 is
  the clinical shortcut. **Both must be gated** or the gate is decorative.
* Neither path takes a session token. `settleDischarge` takes `payload.user` — a *client-supplied
  string* — as the actor for audit. There is no server-side identity on either discharge path.
* `IP_Records_Logic.gs:354 ipr_discharge_()` already looks for a sheet named **`Discharge_Summary`**
  and falls back to the admission row when it is absent. That sheet **does not exist** in the live
  workbook. This is the natural attach point for the new module's record view (§10).

---

## 2. Sheet schemas (live header rows)

Sheet names come from `IPA_CFG` (`IP_Admissions_Logic.gs:15`), `LAB` (`LabSetup.gs:13`),
`IPC_*` (`IP_Clinical_Access.gs:28,41`) and `ipc_*_` in `AccountsIPChargesLogic.gs`.

### IP_Admissions — key `IP_Number`
Declared **13 columns, "FROZEN"** in the header comment at `IP_Admissions_Logic.gs:5-8` and in
`IPA_CFG.HEADERS` (`:19`). The **live sheet has 14**:

```
IP_Number | Patient_ID | Patient_Name | Age_Sex | DOA | TOA | Admission_Type |
Ward_Bed | Bed | Consultant | Diagnosis | Status | DOD | Primary_Doctor_ID
```

`IPA_COL` (`:25`) maps indices 0–12; `Primary_Doctor_ID` at index 13 is read by
`ipc_primaryDoctorId_()` (`IP_Clinical_Access.gs:244`) by header lookup, not by `IPA_COL`.
ID format: `IP` + `yyMM` + `-` + 4 digits (`ipa_nextIpNumber_`, `:79`), e.g. `IP2609-0002`.
Live active rows: `IP2609-0002`, `IP2609-0003`.

`Ward_Bed`/`Bed` are historically divergent — `ipa_resolveWardBed_()` (`:66`) treats the combined
`"B - B201"` form in `Ward_Bed` as authoritative. **Assembly must call that helper, never read
column 7 or 8 raw.**

### Patients — key `Patient_ID`
```
Patient_ID | Password | Name | Age | Gender | DOB | Mobile | WhatsApp | Address |
Conditions | Registration_Date | Salutation | Marital_Status | Blood_Group | Occupation |
Education | Email | Relation_Type | Relation_Name | Emergency_Contact_Name |
Emergency_Number | Referred_By | Allergies
```
ID format `LMTVS####`. **`Allergies` (last column) is populated for real patients**
(e.g. `LMTVS0006` = "Penicillin Allergy") — this is the auto-source for the ALLERGIES section.
No `ABHA` column exists. There is **no Tenant_ID column**. `Password` is stored in clear text
(§5). `ipa_pid_()` (`:36`) upper-cases patient IDs because lower-case rows exist historically.

### Doctors — key `Doctor_ID`  (the "Doctors master")
All six expected fields are present, plus five more:
```
Doctor_ID | Tenant_ID | Display_Name | Specialty | Reg_No | Signature_Line |
Linked_Username | Status | Can_View_All | Default_Consult_Fee | Colour_Tag
```
Live rows:

| Doctor_ID | Display_Name | Specialty | Reg_No | Signature_Line | Linked_Username |
|---|---|---|---|---|---|
| DOC001 | Dr. Meivasagam | General Medicine | `202332` | `Dr. Meivasagam T, MBBS` | doctor1 |
| DOC002 | Dr. Duty MO | Duty Medical Officer | *(empty)* | `Dr. Duty MO` | doctor2 |
| DOC003 | Dr. Nagamanikandan | Cardiothoracic surgeon | *(empty)* | `Dr. Logavignesh` | doctor3 |

* **`Qualification` does not exist as a column** — but `Signature_Line` already carries it
  (`"Dr. Meivasagam T, MBBS"`). Recommendation: **do not add a column**; print `Signature_Line`
  as the qualification line. This is also what every existing clinical note snapshots
  (`Author_Signature_Snapshot`).
* **There is no signature-image column.** Recommendation: no column, no Drive image in v1 —
  print the typed `Signature_Line` plus a wet-ink box. (Proposed, not applied — see §12 Q2.)
* **`Reg_No` is blank for DOC002 and DOC003.** A registration number is a statutory element of
  a discharge summary. This is a *data* gap, not a code gap: it must be filled before those
  doctors can sign. The signing path will refuse a signer with no `Reg_No`.
* `DOC003`'s `Signature_Line` reads "Dr. Logavignesh" while `Display_Name` is
  "Dr. Nagamanikandan". Recorded as a data defect (§11).

### Master_Beds — key `Bed_ID`
```
Bed_ID | Ward | Status | Patient_ID | Patient_Name | DOA | IP_No
```
Status vocabulary in use: `Available`, `Occupied`, `Cleaning` (mixed case — compare
case-insensitively). Bed release is by *index*: `setValue` on columns 3–7.

### IP_CaseSheets_DB — key `Encounter_ID`, foreign key `IP_Number`
45 columns, defined at `IP_Clinical_Access.gs:41 IPC_CASESHEET_HEADERS`; the live sheet matches
except that two legacy columns survive between `Patient Name` and `Sys_BP`
(`Legacy_Age_Sex`, `Legacy_Timestamp_2`) and `Age`/`Sex` sit *after* `Author_Username`.
**Therefore: header-map reads only. Never positional.**

Fields the discharge summary needs, with live examples from `IP-ENC-LMTVS0006-152410`:

| Column | Shape | Live value |
|---|---|---|
| `Chief_Complaints` | JSON array | `[{"duration":"2 Days","condition":"Cold"},{"condition":"Hairloss","duration":"3 Days Days"}]` |
| `History` | JSON array | `[{"duration":"","prefix":"K/C/O","condition":"Hypothyroidism"}]` |
| `Sys_BP`,`Dia_BP`,`PR`,`SpO2`,`Temp`,`Height`,`Weight` | scalars | `140`, `90` |
| `Pallor`…`Edema` | `Yes`/`No` | `No` |
| `CVS`,`RS`,`PA`,`CNS` | free text | `S1S2+, No Murmur` |
| `Primary Diagnosis` | free text (note the **space**, not underscore) | `HYPOTHYROIDISM, Alopecia` |
| `Prescription_JSON` | JSON array | `[{"strength":"Tab","duration":"5 Days","type":"Tab","source":"INTERNAL","comments":"","drugName":"Azithral 500","sig":"1-0-0"}]` |
| `Lab_Orders_JSON` | JSON array | `[]` |
| `Advice` | free text | `Drink plenty of warm fluids` |
| `Status`,`Version`,`Superseded_By`,`Amended_At` | amendment chain | blank on live rows |

**There is no dedicated allergy field on the casesheet** and no admission-medication list beyond
`Prescription_JSON` (which is the *admission* prescription). Allergies come from
`Patients.Allergies` only.

### IP_Notes / IP_Timeline_DB — one sheet, two names
`ipc_timelineSheet_()` (`IP_Clinical_Access.gs:36`) ensures **`IP_Timeline_DB`** with
`IPC_TIMELINE_HEADERS` (`:28`). This is the sheet every IP note is written to; "IP_Notes" is the
*module* name, not a sheet. Live header order differs from the declared order:

```
declared: Timestamp IP_Number Patient_ID Role_Type Note_Data_JSON Author Shift Flags
          Note_ID Author_Doctor_ID Author_Signature_Snapshot Author_Username
live    : Timestamp IP_Number Patient_ID Role_Type Note_Data_JSON Author Shift Flags
          Author_Doctor_ID Author_Signature_Snapshot Author_Username Note_ID
```
`Note_ID` is last on the live sheet, 9th in the constant. `dc_ensureSheet_` appends missing
headers rather than reordering, so both orders are legal. **Header-map reads only.**

Live data quality: of 31 rows, `Author_Doctor_ID`, `Author_Signature_Snapshot`,
`Author_Username` and `Note_ID` are **blank on every historic row**; `Author` sometimes contains
the literal string `"Loading..."` (a UI placeholder that was persisted). Attribution for the
hospital course must fall back to `Author` and tolerate `"Loading..."`.

### IP_Pharmacy_Queue — the queue fed by IP notes
```
Queue_ID | IP_Number | Patient_ID | Drug_Name | Dose | Frequency | Route | Instructions |
Status | Ordered_By | Ordered_At | Action_Flag | Modified_At | Modified_By | Encounter_Note_ID
```
Ensured by `_ensureIPPharmacyQueueSheet_` (`IP_Notes_Logic.gs:32`), written by
`_syncMedOrdersToPharmacyQueue_` (`:513`). `Queue_ID` = `IPQ` + `yyMMddHHmm` + `-` + 4 chars.
`Action_Flag` ∈ `ACTIVE | STOPPED | HOLD`; `Status` ∈ `Pending_Dispense | Dispensed_IP | …`.
`Encounter_Note_ID` is **blank on every live row** — the queue cannot be joined back to the note
that created it. **Consequence: the medication replay must be driven from the notes, and the
queue used only as a cross-check.**

### LAB_ORDERS / LAB_ORDER_TESTS / LAB_RESULTS
Canonical names in `LAB` (`LabSetup.gs:13`), headers in `LAB_SCHEMA` (`:29`). Live sheets match
the schema exactly. The admission link is **`LAB_ORDERS.AdmissionID`** (written at
`LabIntegrationEngine.gs:684` only when the order is IP). `LAB_ORDER_TESTS` and `LAB_RESULTS`
carry **no `AdmissionID`** — they join on `OrderID`.

Join path for the INVESTIGATIONS section:
`LAB_ORDERS` where `AdmissionID = <IP No>` → `OrderID`
→ `LAB_ORDER_TESTS` where `OrderID` matches → per-test status
→ `LAB_RESULTS` where `OrderID` matches → per-parameter values.

### Appointments
```
Appt ID | Patient ID | Patient Name | Date | Time (…) | Purpose | Status (…) | | |
Doctor_ID | Doctor_Name_Snapshot | Booked_By | Attribution_Source
```
Header names contain spaces, parenthetical prose, and **two unnamed columns**. Follow-up booking
(Phase 8) must go through the existing appointment function, never a direct write.

### Audit sheets — there are two, with different schemas
| Sheet | Written by | Headers |
|---|---|---|
| `Audit_Log` | `logAudit_(sess, event, entityType, entityId, details)` — `Doctors_Engine.gs:92` | `Audit_ID, Timestamp, Tenant_ID, Actor_Username, Actor_Role, Actor_Doctor_ID, Event, Entity_Type, Entity_ID, Details_JSON` |
| `Audit_Logs` | `acc_audit_(user, action, module, refId, oldVal, newVal, reason)` — `Accounts.gs:49` | `Audit_ID, Timestamp, User_Name, Action_Type, Module, Reference_ID, Old_Value, New_Value, Reason_Remarks` |

There is **no `Audit_Event_Ledger`** sheet. `Audit_Log` (65 live rows, clinical) is the correct
target for this module because it records the session-derived actor; `Audit_Logs` (50 live rows,
financial) takes a client-supplied user string. **Decision: the module writes `Audit_Log` via
`logAudit_`, and additionally writes `Audit_Logs` via `acc_audit_` for the billing-gate override
only**, so the accounts audit screen shows gate overrides where the accountant will look.

### DS_* sheets
`DS_Summaries`, `DS_Working`, `DS_Snapshots`, `DS_Workflow_Log`, `DS_Phrase_Library`:
**none exist.** `Discharge_Summary`: **does not exist** (read defensively at
`IP_Records_Logic.gs:368`). `IP_Discharge_Drafts` **does exist and is unrelated** — it is the
*billing* discount/package draft (`AccountsIPChargesLogic.gs:41`), 0 live rows. Do not reuse it.

---

## 3. Clinical data shapes (`Note_Data_JSON` by `Role_Type`)

Role types accepted: `DOCTOR | NURSE | CONSULTANT | PROCEDURE | HANDOVER | INVESTIGATION | QUICK`
(`IPC_ROLE_NOTE_TYPES` / `IPC_SECTION_RBAC`, `IP_Clinical_Access.gs:~75-120`).

**DOCTOR** (verbatim from the live sheet):
```json
{"subjectiveObjective":"Patient reviewed","adviceText":"",
 "medOrders":[{"drugName":"P-500 Tablets","dose":"500mg","route":"Oral","freq":"1-0-0",
               "action":"NEW","instructions":""}],
 "investigationOrders":[],
 "systemExam":{"rs":"NVBS bilateral","pa":"Soft, non-tender",
               "cns":"Conscious, oriented","cvs":"S1S2 heard, no murmurs"}}
```
A second live shape carries `{"drugName":"Neurobion Forte Tablets","action":"CONT","queueId":"IPQ…"}`
— i.e. a `CONT`/`STOP`/`HOLD` order may carry **only** `drugName` + `action` (+ sometimes `queueId`).

**PROCEDURE**:
```json
{"complications":"Nil","procedureName":"LUMBAR PUNCTURE","operator":"Dr. Meivasagam",
 "findings":"L4-L5 Palpated and marked. Sterile aseptic precautions done and Lumbar Puncture
             done. CSF collected and sent for lab investigations"}
```
No `anaesthesia` and no procedure `date` field — the note `Timestamp` is the date.

**Medication orders — the critical finding.**
* Fields: `drugName`, `dose`, `route`, `freq`, `action`, `instructions`, sometimes `queueId`,
  sometimes `newDrugName` (read at `IP_Notes_Logic.gs:1246`).
* `action` ∈ `NEW | CONT | STOP | HOLD | MODIFY` (`_syncMedOrdersToPharmacyQueue_`, `:566-577`).
* **There is no drug ID and no inventory ID.** The only identity a medication has is the free-text
  `drugName` (a *brand* string: "P-500 Tablets", "Zerodol-SP Tablets", "Neurobion Forte Tablets").
  There is no generic-name field anywhere in the IP note path. The pharmacy *inventory* sheet has
  `Brand Name` + `Generic Name`, and `_normDrug_()` (`IP_Notes_Logic.gs:67`) is the existing
  normaliser used to match an order back to a queue row.
* **Consequence:** the replay's drug key must be `_normDrug_(drugName)` — normalised **brand**,
  not generic. The DISCHARGE_MEDICATIONS "Generic (UPPER CASE)" column must be resolved by an
  optional lookup into pharmacy inventory `Brand Name → Generic Name`, and left **empty** (never
  guessed) when the brand is not in inventory. The keying strategy is reported in the payload so
  the UI can say so.
* There is **no IV-fluid structure** and **no separate vitals object** in DOCTOR notes; nursing
  vitals arrive under the canonical section `vitals` (alias `vitalSigns`, `IPC_SECTION_ALIASES`,
  `IP_Clinical_Access.gs:~152`).
* `systemExam` is canonicalised to **`sysExam`** on write (`ipc_canonicalSection_`) — a bug fixed
  earlier in the project's history. **Read both keys.**

**Casesheet JSON** — see §2 (`Chief_Complaints`, `History`, `Prescription_JSON`,
`Lab_Orders_JSON`). Provisional diagnosis = `IP_Admissions.Diagnosis` at admission and
`IP_CaseSheets_DB."Primary Diagnosis"`; `updateIPDiagnosis` (`IP_Clinical_Access.gs:793`) is the
structured-diagnosis writer and `getIPDiagnosis` (`:769`) the reader.

---

## 4. Lab results

* **Verified means `LAB_RESULTS.VerifiedAt` is non-empty and `IsDraft` is not TRUE.**
  `LAB_ORDER_TESTS.TestStatus` also carries a verified state and `VerifiedAt`/`VerifiedBy`.
  Belt and braces: require `VerifiedAt` on the result row.
* Amendment chain: `Version`, `IsLatest`, `AmendmentReason`, `AmendedBy`, `AmendedAt`.
  **Only `IsLatest` rows may print.**
* Reference ranges: `LAB_RESULTS.RefRangeText` (per-result snapshot) and, in the catalog,
  `MaleRefLow/High`, `FemaleRefLow/High`, `PaediatricRefText`, `CriticalLow/High`.
* Abnormal flag: `LAB_RESULTS.Flag`. **Use the stored flag; do not recompute** — the stored one is
  what the lab attested to.
* Admission link: `LAB_ORDERS.AdmissionID` only (§2).
* `AttestationHash` exists on both the test and the result row —
  `LabIntegrationEngine.gs:1046` computes it with `Utilities.computeDigest(SHA_256, …)`.
  That is the precedent for this module's hashing.

**Defect:** `getIPLabResults()` (`IP_Notes_Logic.gs:995`) — the function the IP Notes screen uses —
reads a legacy sheet **`Lab_Queue_DB`**, not `LAB_RESULTS`. It is positional
(`data[i][3]`, `[8]`, `[9]`) against a sheet that no longer exists in the live workbook, so it
returns an empty list. **The discharge assembly must not reuse it**; it reads `LAB_*` directly.

---

## 5. Auth

* **Token issue:** `issueSession_(obj)` (`Doctors_Engine.gs:70`) → `Utilities.getUuid()`, payload
  `{username, role, doctorId, name, tenantId}` into `CacheService` key `SESS_<token>`, TTL 21600 s.
* **Token validation:** `validateSession_(token)` (`:83`) is cache-only. **The durable replacement
  is `dc_validateSession_(token)`** (`Doctor_Session_Store.gs:54`) — cache fast path, then the
  `Sessions` sheet, sliding 8-hour expiry, **fails closed**. Every new function must use
  `dc_validateSession_`.
* `Sessions` sheet: `Token, Username, Role, Doctor_ID, Display_Name, Issued_At, Last_Seen,
  Expires_At, Status`.
* **Doctor identity:** `sess.doctorId`, or `resolveDoctorByUsername_()` (`Doctors_Engine.gs:35`)
  via `Doctors.Linked_Username`; profile by `dc_getDoctorById_()` (`Doctor_Core.gs:375`).
* **A reusable role guard does not exist.** The closest things are `resolveIPWrite_()`
  (`IP_Clinical_Access.gs:452`) — note-authorship specific — and `resolveIPRead_()` (`:513`).
  Neither is a general role guard. **This module ships its own, and it is the first statement of
  every public function.**
* **Salted/peppered password hashing (Slice A) is NOT live.** `Users` is
  `Username | Password | Role | Status | Email Address | MFA_Secret` and the live rows hold
  **clear-text passwords** (`admin/admin123`, `meivasagam/doctor123`, `nurse/nurse123`).
* **TOTP is live and reusable:** `verifyMFA(username, userCode)` (`AuthLogin.gs:202`) reads
  `MFA_Secret` from column F and delegates to `processTOTP(secretBase32, userToken)` (`:219`),
  which is a standard HMAC-SHA1 TOTP. `meivasagam` and `nurse` have secrets enrolled.
  **Consequence: TOTP is the only credential worth calling a signing re-authentication here.**
  Password re-entry is offered as a fallback but is, today, a clear-text comparison — it must be
  labelled as such to the customer and replaced when Slice A lands.

---

## 6. Facility identity

`IPP_CLINIC` (`IP_Print_Kit.gs:27`): `{name:"Valarmathi Clinic", tagline:"Premium Healthcare
Services", phone:"+91 88387 23513"}` — a hard-coded object, the single source every IP print uses.
`CLINIC_NAME` also exists in Script Properties, read only by `bc_clinicName_()`
(`Barcode_Engine.gs:440`) for barcode labels. `OP_Database_Engine.gs:546` hard-codes
"Valarmathi Clinic" again in its own letterhead.

There is **no stored address, no registration number, and no logo for print** (the app logo is a
Google user-content URL in `ScriptsMV.html:2`, unusable inside a PDF). A discharge summary needs a
facility address and registration line. **Design response:** read facility identity from Script
Properties with `IPP_CLINIC` as the fallback, so nothing changes until the keys are set. See §12 Q1.

---

## 7. Print and PDF pattern

* **Browser print:** the established pattern is a hidden iframe fed server-rendered HTML —
  `getIPNotesPrintHtml()` (`IP_Notes_Logic.gs:1055`), `getIPRecordPrintHtml()`
  (`IP_Records_Logic.gs:414`), both returning `{success, html}`.
* **HTML → PDF:** `doGet` (`CodeMV.gs:12-15`) does
  `Utilities.newBlob(html, 'text/html', 'report.html').getAs('application/pdf')` then
  `Utilities.base64Encode(blob.getBytes())` and hands the browser a `data:` URI. **This is the
  only conversion route in the project and it works today.**
* **CSS that survives that converter** is documented, with reasons, in `IP_Print_Kit.gs:1-25`:
  tables with `table-layout:fixed` (never flexbox), one outer `<table class="page">` with a
  repeating `<thead>` and pinned `<tfoot>`, `overflow-wrap:anywhere`, `page-break-inside:avoid`,
  `print-color-adjust:exact`, `@page{size:A4;margin:13mm 12mm 14mm}`.
  **`ipp_doc_()` (`:220`) already produces exactly the A4 document shell Prompt 07 specifies**,
  including the per-page patient strip. The discharge renderer composes into it rather than
  inventing a second stylesheet.
* Escaping helpers exist and are mandatory: `ipp_esc_()` (`:35`), `ipp_escMultiline_()` (`:42`).
  Layout helpers: `ipp_kv_`, `ipp_sec_`, `ipp_cols_`, `ipp_table_`, `ipp_sig_`, `ipp_when_`.
* **Data-URI images in the PDF converter: unproven.** Nothing in the project embeds one.
  Tamil text in the converter: unproven. **Both remain the Phase 7 spike.**
* **`SecurityTokens.js` does not exist.** There is no `getSecret_`, no `timingSafeEqual_`, no
  `bytesToB64Url_`. The available primitives are `Utilities.computeDigest(SHA_256, …)`
  (`LabIntegrationEngine.gs:1046`) and `Utilities.computeHmacSignature(HMAC_SHA_1, …)`
  (`AuthLogin.gs:243`). `Utilities.computeHmacSignature` supports `HMAC_SHA_256`.
  **This module must supply its own secret accessor, constant-time comparison and base64url
  encoder.** That is a deviation from the prompt pack and is recorded as such.
* A QR encoder **does exist in the front end**: `Barcode_QR_Lib.html` (57 KB, client-side).
  It is not callable server-side. Phase 7 still needs a server-side encoder for the stored PDF.

---

## 8. Frontend shell

* **Router:** `switchTab(tabId, clickedElement)` — `ScriptsMV.html:82`. It removes
  `.active-section` from every `.page-section` and adds it to `#tabId`, then runs a growing chain
  of per-tab hooks (`:93-140+`).
* **Both mechanisms exist.** `.page-section` / `.active-section` is the *router's* mechanism;
  `d-none` is used for sub-workspaces inside a tab (`:113`, `:181-188`). A new top-level module
  registers as a `.page-section`; workspaces inside it toggle `d-none`.
* **Includes:** `<?!= include('Name'); ?>` in `Index.html` — markup files inside
  `<div class="main-content">`, script files after `<?!= include('ScriptsMV'); ?>` at the end of
  `<body>`. The paired `X_UI.html` / `X_Scripts.html` convention is already established
  (`Accounts_Discharge_Script`, `AccountsInsuranceScripts`, …).
* **RBAC in the shell is cosmetic:** `applyRBAC(role)` (`ScriptsMV.html:19`) hides nav elements
  by ID. It is a UI convenience with no server counterpart. Note `nurse` is granted only
  `nav-tab-dash` and `nav-tab-emr` — the Discharge Desk entry must live somewhere a nurse can
  reach, i.e. inside the EMR workspace, not behind a new top-level nav item only admins see.
* **Theme:** `Styles.html` + the token helpers in `Index.html` (`cssVar`, `chartPalette`,
  `chartTheme`, `onThemeChange`); `light-theme` class on `<html>`/`<body>`, persisted in
  `localStorage['cresrx_theme']`. **Use variables only.**

---

## 9. Helpers available for reuse

| Need | Existing helper | File:line |
|---|---|---|
| Header map (cached) | `dc_headerMap_(sheet)` | `Doctor_Core.gs:49` |
| Bounded sheet values | `dc_sheetValues_(sheet)` | `:67` |
| Ensure sheet + append missing headers | `dc_ensureSheet_(ss, name, headers)` | `:100` |
| Ensure one column | `dc_ensureColumn_(sheet, header)` | `:83` |
| String / upper / int / money | `dc_str_`, `dc_upper_`, `dc_int_`, `dc_money_` | `:136-139` |
| Date formatting | `dc_fmtDate_`, `dc_dateKey_`, `ipp_when_` | `:183`, `:195`, `IP_Print_Kit.gs:201` |
| Session | `dc_validateSession_`, `dc_sessionName_` | `Doctor_Session_Store.gs:54,49` |
| Doctor profile | `dc_getDoctorById_` | `Doctor_Core.gs:375` |
| Care team | `dc_isOnCareTeam_`, `ipc_primaryDoctorId_` | `:751`, `IP_Clinical_Access.gs:244` |
| Clinical audit | `logAudit_(sess, event, type, id, details)` | `Doctors_Engine.gs:92` |
| Financial audit | `acc_audit_(user, action, module, ref, old, new, reason)` | `Accounts.gs:49` |
| Tenant | `getTenantId_()` → `"VALARMATHI"` | `Doctors_Engine.gs:7` |
| Drug-name normalisation | `_normDrug_(name)` | `IP_Notes_Logic.gs:67` |
| Ward/bed reconciliation | `ipa_resolveWardBed_` | `IP_Admissions_Logic.gs:66` |
| Print composition | the whole `ipp_*` kit | `IP_Print_Kit.gs` |

**ID generators:** each module rolls its own (`ipa_nextIpNumber_`, `ipc_id_('CHG')`,
`_generateNoteId_`, `Utilities.getUuid()`). No shared generator. **Config storage:** Script
Properties, ad hoc, with no accessor of its own — `ACC_CFG.LOCK_PROP` (`Accounts.gs:43`) and
`CLINIC_NAME` (`Barcode_Engine.gs:442`) are the only two patterns. **There is no config sheet.**

---

## 10. IP Records / IP Ledger attach points

* `getIPRecordsLedger(sessionToken)` — `IP_Records_Logic.gs:49` — the per-admission ledger list.
* `getIPRecordFile(encounterId, sessionToken)` — `:197` — the detail drawer, whose payload already
  contains a `discharge` block built by **`ipr_discharge_()` (`:354`)**. That function looks for a
  `Discharge_Summary` sheet and returns `{source, date, outcome, summary, diagnosis}`.
  **This is the single cleanest attach point in the codebase:** teach `ipr_discharge_()` to read
  `DS_Summaries` first and the drawer gains a discharge block with no UI change at all, then add
  a "Discharge Summary" card with View / Print / Download.
* `getIPRecordPrintHtml(encounterId, sessionToken, opts)` — `:414` — the full-file print already
  takes a `{casesheet, notes, trend, discharge}` opts object, so the signed summary can be
  appended to the complete printed file.
* `IP_Ledger_UI.html` — the status-pill ledger (`ALL / ACTIVE / DISCHARGED`, `:174`) — is where the
  per-admission discharge status badge belongs.
* `IP_Records_UI.html:281` already branches on `p.status === 'DISCHARGED'`.

---

## 11. Known defects that intersect this feature (listed, not fixed)

1. **`settleDischarge` writes `IP_Admissions` by hard-coded column 12/13**
   (`AccountsIPChargesLogic.gs:403-404`). Any column insertion silently corrupts discharge.
2. **`IPA_CFG.HEADERS` declares 13 columns; the live sheet has 14** (`Primary_Doctor_ID`).
   `IP_Admissions_Logic.gs:5-8` calls the 13-column schema "FROZEN".
3. **`getIPLabResults()` reads the dead `Lab_Queue_DB` sheet positionally**
   (`IP_Notes_Logic.gs:995-1030`) — the IP lab panel is empty for every patient.
4. **`IP_Pharmacy_Queue.Encounter_Note_ID` is never populated** — orders cannot be traced to the
   note that raised them.
5. **`Author_Doctor_ID` / `Author_Signature_Snapshot` / `Author_Username` / `Note_ID` are blank on
   every historic `IP_Timeline_DB` row**, and `Author` sometimes holds the literal `"Loading..."`.
6. **Clear-text passwords in `Users`** — Slice A hashing is not live (§5).
7. **No backend RBAC on the IP Admissions module** — `processPatientDischarge` takes no token.
8. **`Master_Beds` desync**: bed release is positional and the two discharge paths leave beds in
   different states (`Available` vs `Cleaning`).
9. **Lower-case patient IDs exist** — `ipa_pid_()` upper-cases on read; every new key comparison
   must do the same.
10. **`Doctors.DOC003`**: `Signature_Line` = "Dr. Logavignesh" but `Display_Name` =
    "Dr. Nagamanikandan". A signed document would print the wrong name.
11. **`Doctors.Reg_No` blank for DOC002 and DOC003.**
12. **`Appointments` has two unnamed header columns** and header names containing spaces and
    parentheses.

---

## 12. Design impact

| Finding | Consequence for the discharge module |
|---|---|
| `ds_` is **already in use** — `ds_sessionSheet_`, `ds_touchSession_`, `ds_slideExpiry_` (`Doctor_Session_Store.gs`), `ds_scheduleSheet_`, `ds_exceptionSheet_`, `ds_generateSlots_`, `ds_bookingsByTime_`, `ds_parseDateKey_` (`Doctor_Schedule_Engine.gs`), and constants `DS_SESSION_HOURS`, `DS_CACHE_SECONDS`, `DS_WEEKDAYS` | Apps Script is one flat global namespace and a duplicate `function` name is **silently overwritten**, last file wins. The public API keeps the contract names (`ds_getQueue`, `ds_sign`, … — each verified non-colliding), but **every private helper and constant in this module is prefixed `dsx_` / `DSX_`**. Sheet names stay `DS_*`. |
| No drug ID or generic name anywhere in the IP note path | Medication replay keys on normalised **brand** text via `_normDrug_()`. Generic is resolved by an optional pharmacy-inventory lookup and left **empty** when unknown — never guessed. The keying strategy is returned in the payload and shown in the UI. |
| `IP_Pharmacy_Queue.Encounter_Note_ID` empty | Replay is driven from `IP_Timeline_DB` DOCTOR notes; the queue is a cross-check that raises a SOFT warning on disagreement, not a source. |
| `getIPLabResults` reads a dead sheet | Assembly reads `LAB_ORDERS`(`AdmissionID`) → `LAB_ORDER_TESTS` → `LAB_RESULTS` directly, filtered to `IsLatest` ∧ `VerifiedAt` ∧ ¬`IsDraft`. |
| Two discharge paths, neither authenticated | The Phase 8 gate must patch **both** `processPatientDischarge` and `settleDischarge`. Because neither carries a token, the gate reads DS status by IP number and logs the override against the client-supplied user — a known weakness, recorded, not silently accepted. |
| `settleDischarge` writes by index | The gate patch must not add, remove or reorder any `IP_Admissions` column. The module adds **no** columns to existing sheets. |
| Casesheet/notes header order differs from the declared constants | **Every read is header-mapped.** No positional access anywhere in this module. |
| `systemExam` ⇄ `sysExam` alias | Assembly reads both keys for every system-exam lookup. |
| Clear-text passwords; TOTP live for 2 of 3 accounts | Signing prefers TOTP. Password fallback is offered but flagged; the sign lockout (5 failures / 15 min) matters more than usual because the password check is weak. |
| `Doctors.Reg_No` blank for 2 of 3 doctors | Signing **refuses** a signer whose `Reg_No` is empty, with a message naming the fix. This is deliberate: a discharge summary without a registration number is not a valid document. |
| No `Qualification` column, but `Signature_Line` carries it | No column is added. `Signature_Line` prints as the qualification line. |
| No signature image anywhere | v1 prints the typed signature line plus a wet-ink box. No Drive image, no public URL. |
| No facility address or registration line | Read from Script Properties with `IPP_CLINIC` as fallback; nothing changes until the keys are set. |
| `SecurityTokens.js` does not exist | The module supplies `dsx_secret_`, `dsx_hmacSha256_`, `dsx_timingSafeEqual_`, `dsx_b64url_` itself. |
| No `Audit_Event_Ledger` | Clinical mutations → `Audit_Log` via `logAudit_`. Billing-gate overrides additionally → `Audit_Logs` via `acc_audit_`. |
| `applyRBAC` gives `nurse` only Dashboard + EMR | The Discharge Desk entry lives inside the EMR workspace and on the IP Notes patient view, not behind a new admin-only nav item. |
| `ipr_discharge_()` already reads an absent `Discharge_Summary` sheet | Teach it `DS_Summaries` — the IP Records drawer then shows discharge state with no UI change. |
| `ipp_doc_()` already emits the A4 shell Prompt 07 describes | The discharge renderer composes into `IP_Print_Kit`; it does not carry its own stylesheet. |
| Single global script lock, 10 s waits across every module | Assembly, diffing and PDF generation run **outside** the lock. Only commits take it. |
| One tenant, `getTenantId_()` = `"VALARMATHI"` | `Tenant_ID` is written on every DS row now, so the Postgres migration does not need a backfill. |

---

## 13. Blocking questions — only what the code cannot answer

1. **Facility identity for the printed document.** A discharge summary should carry the clinic's
   postal address and registration/licence number. Neither exists anywhere in the workbook or in
   Script Properties. Please supply the exact address block and registration line to print, or
   confirm that the summary prints name + phone only (the `IPP_CLINIC` values) for now.
2. **Doctor signature images.** Confirm: v1 prints the typed `Signature_Line`
   (e.g. "Dr. Meivasagam T, MBBS") plus a wet-ink box, with **no** scanned signature image.
   If you want scanned signatures, that needs a new `Signature_Image_File_ID` column on `Doctors`
   and a Drive folder — proposed, not applied.
3. **`Reg_No` for DOC002 (Dr. Duty MO) and DOC003 (Dr. Nagamanikandan).** Signing will be refused
   until these are filled. Please confirm the numbers, and confirm whether "Dr. Duty MO" is a real
   registered person or a shared account — a shared login must not be allowed to sign.
4. **`DOC003` name mismatch**: `Display_Name` "Dr. Nagamanikandan" vs `Signature_Line`
   "Dr. Logavignesh". Which is correct? A signed document will print one of them.
5. **Who is "the consultant of record"?** `IP_Admissions` carries both `Consultant` (a display
   string, e.g. "Dr. Logavignesh") and `Primary_Doctor_ID` (e.g. `DOC001`), and they currently
   disagree on live row `IP2609-0002`. Confirm that **`Primary_Doctor_ID` is authoritative** and
   `Consultant` is display-only, or say otherwise.
6. **Password-based signing.** Passwords are clear text today. Options: (a) TOTP only — the
   doctor must enrol MFA before they can sign; (b) TOTP with password fallback, accepting the
   weakness until Slice A ships. Recommendation: **(b) with the fallback off by default per
   tenant**, so your clinic can start now and hospitals start strict.
7. **Retention of superseded snapshots.** Amendments keep every signed version forever.
   Confirm that is what you want (it is what a medico-legal record should do) versus any
   retention limit you have been asked for.
