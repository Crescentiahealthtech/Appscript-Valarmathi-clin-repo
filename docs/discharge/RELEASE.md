# IP Discharge Summary Engine — setup, operation, rollback

The module is inert until `setupDischargeSummaryModule()` is run. Nothing in it
adds a column to `IP_Admissions` or to any existing clinical sheet, and with
`DS_BILLING_GATE = OFF` both discharge paths behave exactly as they did before.

---

## 1. Setup, in order

| # | Step | How |
|---|---|---|
| 1 | Create the sheets | Apps Script editor → run `setupDischargeSummaryModule()`. Idempotent; safe to re-run after every upgrade. Read its `data.proposed` for data gaps in the Doctors master. |
| 2 | Check it | Run `ds_selfTest()` — read-only. Every line must say `pass`. |
| 3 | Check the state machine | Run `ds_testTransitions()` — in-memory, writes nothing. `0 failed`. |
| 4 | Fill doctor registration numbers | `Doctors` sheet → `Reg_No` for every doctor who will sign. **Signing is refused without one.** Today DOC002 and DOC003 are blank. |
| 5 | Check the signature lines | `Doctors.Signature_Line` is what prints under the signature and carries the qualification. (DOC003's was corrected at source on 11 Sep 2026.) |
| 6 | Set the Script Properties you want | See §2. All have working defaults; none is required to start. |
| 7 | Run the print spike | Run `ds_printSpike()` and read the log. It tells you whether the QR image and Tamil text survive Apps Script's PDF converter on *this* deployment. |
| 8 | Install the PDF trigger | Run `installDischargePdfTrigger()` once. It refuses to duplicate itself. |
| 9 | Dry-run on a real admission | Run `ds_dryRunAssembly("IP2609-0002")` — read-only, writes nothing. Read the section sizes, readiness list and warnings before letting staff near it. |
| 10 | Populate the phrase library | `DS_Phrase_Library` starts **empty and stays empty**. Clinic staff write the diet, activity, wound-care, red-flag and follow-up lines. **Tamil text is authored and reviewed by the clinic — the system never machine-translates it.** |
| 11 | Redeploy the web app | Menu → Deploy → Manage deployments → edit → new version. The `?verifyDS=` route and every new server function only exist in a fresh deployment. |

---

## 2. Configuration (Script Properties)

| Key | Default | What it does |
|---|---|---|
| `DS_PREPARER_ROLES` | `nurse,doctor` | Who may prepare and submit. A duty doctor (RMO) preparing is the `doctor` entry. |
| `DS_REQUIRE_PREPARER_REVIEW` | `true` | `false` turns on the **solo-clinic fast path**: a doctor may sign straight from GENERATED or IN_PREPARATION, logged as "signed without preparer review". Off by default; per tenant. |
| `DS_SIGNER_POLICY` | `CONSULTANT_OR_ANY_DOCTOR_WITH_REASON` | The consultant of record signs freely; any other doctor must give an on-behalf reason, which is logged and printed. Any other value restricts signing to the consultant. |
| `DS_BILLING_GATE` | `WARN` | `OFF` = no gate at all. `WARN` = refuse once with `DS_NOT_SIGNED`, proceed on a logged reason. `BLOCK` = refuse, except DEATH and ABSCONDED which always behave as WARN. |
| `DS_TAT_AMBER_MIN` / `DS_TAT_RED_MIN` | `120` / `240` | Desk timer thresholds, in minutes from Initiated_At. |
| `DS_SIGN_ALLOW_PASSWORD` | `true` | `false` requires an authenticator code to sign. A doctor who has enrolled MFA must use it either way. |
| `DS_FACILITY_ADDRESS` | `Crescentia Healthtech, Tamil Nadu, India` | Printed under the letterhead. The phone already prints from `IPP_CLINIC`, so it is not repeated here. Set this property to override per tenant. |
| `DS_FACILITY_REG_LINE` | *(empty)* | Facility registration / licence line. **Not yet supplied** — empty prints nothing rather than a placeholder. |
| `DS_AI_ENABLED` | `false` | Reserved. No AI code ships in this build. |
| `DS_AI_REQUIRE_CONSENT` | `true` | Reserved. |
| `DS_PUSH_DISCHARGE_RX` | `false` | Reserved: discharge prescription to pharmacy. |
| `DS_BOOK_FOLLOWUP` | `false` | Reserved: follow-up appointment from the sign modal. |
| `DS_VERIFY_SECRET` | *(auto)* | The HMAC secret behind the QR verification token. Created on first use. **Never change it** — every printed QR stops verifying. |
| `DS_WEBAPP_URL` | *(auto)* | Fallback for the verification URL when `ScriptApp.getService().getUrl()` is unavailable. |

---

## 3. What each role does

**Staff nurse (or duty doctor) — the preparer.**
Open **EMR → Discharge Desk**. Your tabs are Generated and Returned. Open a
summary and work down the left-hand list. Sections the machine filled are
marked with a wand; anything you change is marked with a pen and the doctor
will see exactly what you changed. Empty means the ward did not record it — fill
it, or leave it empty honestly. Click a source chip to read the note or result a
line came from. The red items on the right must all clear before **Submit for
signature** will work: final diagnosis, allergy status, medications ticked
Reviewed, condition at discharge, and a follow-up plan or "not required". Amber
items are warnings you should look at. Work saves every 30 seconds and on
Ctrl+S. If the doctor returns the summary, their comments are pinned to the
sections they apply to.

**Doctor.**
You sign without giving a reason when you are the consultant of record for that
admission — that means matching *either* `Primary_Doctor_ID` *or* the
`Consultant` name on the admission row. Any other doctor may still sign, but
must give an on-behalf reason, which is logged and printed on the summary.

Press **Initiate Discharge** when you decide the patient is going home, not when
they reach the billing counter — that is the whole point of the module. Your
default tab is Pending Signature. Before signing, use **Changes**: "Changes by
staff" shows what the ward altered from the machine draft, "Since last review"
shows what moved since you last looked. Signing asks you to tick each review
block, then a credential — your authenticator code, or your password if you have
not enrolled MFA. If you are not the consultant of record you must give a reason,
and it is printed on the summary. After signing the summary is locked and
hashed; to change anything, start an **amendment** with a reason. Earlier signed
versions stay viewable and print marked "Superseded".

**Billing desk / accounts.**
You see signed summaries and can print them. When you settle a discharge and the
summary is not signed, you are asked once and may proceed with a reason. That
reason is recorded in both the clinical and the accounts audit trails and shows
on the discharge dashboard, so please make it a real one.

---

## 4. Rollback

No data is deleted by any of these.

1. **Turn off the gate**: set `DS_BILLING_GATE = OFF`. Both discharge paths
   return to their previous behaviour immediately — no redeploy needed.
2. **Hide the module**: remove the Discharge Desk card from
   `EMR_Cockpit.html` (one `<div>`, marked with a comment) and redeploy.
3. **Remove it entirely**: delete the `DS_*.gs` files, the two
   `DS_Desk_*.html` includes from `Index.html`, the `verifyDS` branch at the top
   of `doGet` in `CodeMV.gs`, the gate blocks in `processPatientDischarge` and
   `settleDischarge`, and `ipr_dischargeSummaryBlock_` in `IP_Records_Logic.gs`.
   Every one of those call sites is guarded by `typeof ... === 'function'`, so
   the rest of the app keeps working with the files simply absent.
4. The `DS_*` sheets and the stored PDFs are left alone. Nothing is destroyed.

---

## 5. What is deliberately not in this build

* **No AI.** The engine is deterministic end to end. The optional AI draft of
  the hospital course is a later, flag-gated addition; the config keys are
  reserved so nothing has to be renamed later.
* **No discharge prescription push to pharmacy, and no follow-up booking.**
  Both are behind reserved flags and are the next integration.
* **No scanned signature images.** The typed `Signature_Line` prints, with a
  wet-ink box beside it. Adding images needs a new `Doctors` column and a Drive
  folder; it was proposed and not applied.
* **The password signing path is weak** and is labelled as such: passwords in
  the `Users` sheet are still clear text. Prefer TOTP, and set
  `DS_SIGN_ALLOW_PASSWORD = false` once every signer has enrolled.

---

## 6. Known risks carried into production

| Risk | Where it bites | What is in place |
|---|---|---|
| No drug or inventory ID in IP notes | Discharge medications key on normalised brand text | Generic resolved only from stocked inventory, never guessed; ambiguous keys raise a warning rather than merging; medications must be ticked Reviewed before signing |
| `settleDischarge` writes `IP_Admissions` by hard-coded column index | Any future column insert breaks discharge | This module adds no column to that sheet; the defect is recorded in DISCOVERY.md §11.1 |
| Apps Script PDF fidelity for images and Tamil | The stored PDF, not browser printing | `ds_printSpike()` measures it on the real deployment; the verification code and URL are plain text and always print, so the document verifies even if the QR image is dropped |
| Clear-text passwords | Signing strength | TOTP preferred and enforced for anyone enrolled; five failures in fifteen minutes locks signing and writes `DS_SIGN_LOCKOUT` |
| One global script lock | Contention at peak | Assembly, diffing, rendering and PDF generation all run outside the lock; only commits take it |
| `IP_Admissions.Consultant` may name somebody the Doctors master does not know | That doctor is asked for an on-behalf reason they should not need | Both consultant fields are accepted; when the display name resolves to nobody the signing log records `unresolvedConsultantName`, so it shows up as stale data rather than as a doctor behaving oddly |
| Sparse ward notes | Thin drafts | Empty sections stay visibly empty, the readiness engine blocks unsafe signing, and source chips make weak documentation obvious rather than hiding it |
